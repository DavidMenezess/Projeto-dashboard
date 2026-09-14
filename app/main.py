"""
main.py

Ponto de entrada da API. Define as rotas:

  POST /auth/login              -> troca Nome + Sobrenome + senha por um token JWT
  GET  /me                      -> dados do usuário logado (nome, cargo, tipo)
  GET  /usuarios                -> lista todos os usuários (só admin)
  POST /usuarios                -> cria um novo usuário (só admin)
  GET  /api/tarefas             -> dados da seção "Produção 2026" (protegida)
  GET  /api/processos-parados   -> dados da seção "Processos Parados" (protegida)
  GET  /health                  -> checagem simples de saúde do serviço

Rodar localmente:
  uvicorn app.main:app --reload

Rodar em produção: veja o Dockerfile (usa uvicorn sem --reload).
"""

import asyncio
import logging
from datetime import datetime, timezone, timedelta
from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy.orm import Session

from app.config import settings
from app.database import criar_tabelas, obter_sessao, Usuario
from app.auth import (
    verificar_senha, gerar_hash_senha, criar_token_acesso,
    obter_usuario_autenticado, exigir_administrador, gerar_chave_login,
)
from app.data_processor import processar_tarefas, processar_processos_parados, processar_tarefas_periodo
from app.cache import atualizar_cache, obter_cache, obter_ultima_sincronizacao

# Log estruturado simples. Em produção, isso ajuda a auditar acessos
# (exigência prática da LGPD: saber quem acessou o quê e quando).
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("dashboard-api")

app = FastAPI(
    title="API — Dashboard Directus Advocacia e Consultoria",
    description="Fornece os dados de produção e processos parados ao dashboard, com login por Nome e Sobrenome.",
    version="1.2.0",
)

# CORS: só os endereços listados em ORIGENS_PERMITIDAS (.env) podem chamar esta API
# a partir do navegador. Sem isso, qualquer site poderia tentar consumir os dados.
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.lista_origens_permitidas,
    allow_credentials=True,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


def _data_referencia_configurada() -> datetime | None:
    """
    Converte DATA_REFERENCIA_TESTE (se preenchida no .env) em um objeto datetime.
    Usado só para testes locais com planilhas antigas — veja comentário no config.py.
    """
    if not settings.DATA_REFERENCIA_TESTE:
        return None
    return datetime.strptime(settings.DATA_REFERENCIA_TESTE, "%d/%m/%Y")


@app.on_event("startup")
def ao_iniciar():
    """Roda uma vez quando o servidor sobe: cria as tabelas e faz a primeira carga de dados."""
    criar_tabelas()
    _sincronizar_dados_locais()
    asyncio.create_task(_loop_sincronizacao_automatica())
    logger.info("API iniciada e dados carregados.")


INTERVALO_SINCRONIZACAO_MINUTOS = 10


async def _loop_sincronizacao_automatica():
    """Reprocessa as planilhas periodicamente, sem precisar reiniciar o container."""
    while True:
        await asyncio.sleep(INTERVALO_SINCRONIZACAO_MINUTOS * 60)
        try:
            _sincronizar_dados_locais()
            logger.info("Re-sincronizacao automatica concluida.")
        except Exception:
            logger.exception("Falha na re-sincronizacao automatica - mantendo os ultimos dados validos em cache.")


def _sincronizar_dados_locais():
    """
    Carrega os dados a partir dos arquivos locais configurados no .env.

    Esta função é o único lugar que vai mudar quando a sincronização com o
    OneDrive for implementada: em vez de ler direto de CAMINHO_PLANILHA_*,
    vai primeiro baixar o arquivo mais recente do OneDrive para esse mesmo
    caminho, e então chamar as mesmas funções de processamento abaixo.
    """
    data_ref = _data_referencia_configurada()

    try:
        dados_tarefas = processar_tarefas(settings.CAMINHO_PLANILHA_TAREFAS, data_referencia=data_ref)
        atualizar_cache("tarefas", dados_tarefas)
        logger.info("Planilha de tarefas processada com sucesso.")
    except FileNotFoundError:
        logger.warning("Planilha de tarefas não encontrada em %s — aguardando arquivo.", settings.CAMINHO_PLANILHA_TAREFAS)

    try:
        dados_processos = processar_processos_parados(settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS, data_referencia=data_ref)
        atualizar_cache("processos_parados", dados_processos)
        logger.info("Planilha de processos parados processada com sucesso.")
    except FileNotFoundError:
        logger.warning("Planilha de processos parados não encontrada em %s — aguardando arquivo.", settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS)


@app.get("/health", tags=["Infraestrutura"])
def checagem_de_saude():
    """Usado por serviços de monitoramento (ex: Render) para saber se a API está no ar."""
    return {"status": "ok", "ultima_sincronizacao": obter_ultima_sincronizacao()}


class LoginRequest(BaseModel):
    """Login por Nome + Sobrenome — mais fácil de lembrar do que e-mail."""
    primeiro_nome: str
    sobrenome: str
    senha: str


@app.post("/auth/login", tags=["Autenticação"])
def login(dados: LoginRequest, sessao: Session = Depends(obter_sessao)):
    """
    Recebe Nome, Sobrenome e senha, confere no banco e devolve um token JWT.
    O dashboard deve enviar esse token no header 'Authorization: Bearer <token>'
    em todas as chamadas seguintes às rotas protegidas.
    """
    chave_login = gerar_chave_login(dados.primeiro_nome, dados.sobrenome)
    usuario = sessao.query(Usuario).filter(Usuario.chave_login == chave_login).first()

    credenciais_invalidas = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Nome, sobrenome ou senha incorretos.",
    )

    if usuario is None or not verificar_senha(dados.senha, usuario.senha_hash):
        # Mensagem genérica de propósito: não revela se o usuário existe ou não,
        # o que dificulta um ataque de enumeração de contas.
        logger.info("Tentativa de login malsucedida para '%s'.", chave_login)
        raise credenciais_invalidas

    if usuario.ativo != 1:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Usuário desativado.")

    usuario.ultimo_login = datetime.now(timezone.utc)
    sessao.commit()

    token = criar_token_acesso(usuario.chave_login)
    logger.info("Login bem-sucedido: %s.", usuario.chave_login)
    return {"access_token": token, "token_type": "bearer"}


@app.get("/me", tags=["Autenticação"])
def obter_meu_usuario(usuario: Usuario = Depends(obter_usuario_autenticado)):
    """Dados do usuário logado — o dashboard usa isso para saber o nome, cargo e se é administrador."""
    return {
        "primeiro_nome": usuario.primeiro_nome,
        "sobrenome": usuario.sobrenome,
        "nome": f"{usuario.primeiro_nome} {usuario.sobrenome}",
        "cargo": usuario.cargo,
        "tipo": usuario.tipo,
    }


class NovoUsuario(BaseModel):
    """Dados exigidos para um administrador cadastrar um novo usuário."""
    primeiro_nome: str
    sobrenome: str
    senha: str
    cargo: str | None = None
    tipo: str = "normal"  # "admin" ou "normal"


@app.get("/usuarios", tags=["Usuários"])
def listar_usuarios(
    admin: Usuario = Depends(exigir_administrador),
    sessao: Session = Depends(obter_sessao),
):
    """
    Lista todos os usuários cadastrados, com cargo, tipo de acesso e último
    login. Só administradores podem ver esta lista.

    O campo 'sessao_ativa' é uma aproximação: considera que a pessoa ainda
    está "conectada" se o último login foi dentro da janela de validade do
    token (ACCESS_TOKEN_EXPIRE_MINUTES). Não é uma contagem em tempo real
    (isso exigiria um mecanismo à parte, tipo WebSocket) — é uma estimativa
    razoável de quem esteve ativo recentemente.
    """
    limite_sessao_ativa = datetime.now(timezone.utc) - timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
    usuarios = sessao.query(Usuario).order_by(Usuario.primeiro_nome).all()

    resultado = []
    for u in usuarios:
        ultimo_login_utc = u.ultimo_login.replace(tzinfo=timezone.utc) if u.ultimo_login else None
        resultado.append({
            "id": u.id,
            "primeiro_nome": u.primeiro_nome,
            "sobrenome": u.sobrenome,
            "nome": f"{u.primeiro_nome} {u.sobrenome}",
            "chave_login": u.chave_login,
            "cargo": u.cargo,
            "tipo": u.tipo,
            "ativo": bool(u.ativo),
            "criado_em": u.criado_em.strftime("%d/%m/%Y") if u.criado_em else None,
            "ultimo_login": ultimo_login_utc.strftime("%d/%m/%Y %H:%M") if ultimo_login_utc else None,
            "sessao_ativa": bool(ultimo_login_utc and ultimo_login_utc > limite_sessao_ativa),
        })
    return {"usuarios": resultado}


@app.post("/usuarios", status_code=status.HTTP_201_CREATED, tags=["Usuários"])
def criar_usuario(
    dados: NovoUsuario,
    admin: Usuario = Depends(exigir_administrador),
    sessao: Session = Depends(obter_sessao),
):
    """Cria um novo usuário. Só administradores podem cadastrar outras pessoas."""
    if dados.tipo not in ("admin", "normal"):
        raise HTTPException(status_code=422, detail="Campo 'tipo' deve ser 'admin' ou 'normal'.")

    if len(dados.senha) < 8:
        raise HTTPException(status_code=422, detail="A senha deve ter pelo menos 8 caracteres.")

    chave_login = gerar_chave_login(dados.primeiro_nome, dados.sobrenome)
    ja_existe = sessao.query(Usuario).filter(Usuario.chave_login == chave_login).first()
    if ja_existe:
        raise HTTPException(
            status_code=409,
            detail="Já existe um usuário com este nome e sobrenome. Use um sobrenome mais completo para diferenciar.",
        )

    novo = Usuario(
        primeiro_nome=dados.primeiro_nome,
        sobrenome=dados.sobrenome,
        chave_login=chave_login,
        senha_hash=gerar_hash_senha(dados.senha),
        cargo=dados.cargo,
        tipo=dados.tipo,
        ativo=1,
    )
    sessao.add(novo)
    sessao.commit()

    logger.info("Usuário '%s' criado por '%s'.", novo.chave_login, admin.chave_login)
    return {
        "id": novo.id, "primeiro_nome": novo.primeiro_nome, "sobrenome": novo.sobrenome,
        "chave_login": novo.chave_login, "cargo": novo.cargo, "tipo": novo.tipo,
    }


@app.get("/api/tarefas", tags=["Dados"])
def obter_dados_tarefas(usuario: Usuario = Depends(obter_usuario_autenticado)):
    """
    Dados de tudo que vem da planilha de tarefas, organizados por ano:
    'anos_disponiveis' (lista de anos encontrados), 'por_ano' (Produção de
    cada ano), 'tipos_tarefa_por_ano' (Tipo de Tarefa de cada ano),
    'pendencias' (vencidas/a vencer, sempre o quadro atual) e
    'detalhamento' (todas as tarefas, de todos os anos, com o campo 'ano').
    Requer login.
    """
    dados = obter_cache("tarefas")
    if dados is None:
        raise HTTPException(status_code=503, detail="Dados ainda não sincronizados. Tente novamente em instantes.")
    return dados


@app.get("/api/tarefas/periodo", tags=["Dados"])
def obter_dados_tarefas_periodo(
    inicio: str,
    fim: str,
    usuario: Usuario = Depends(obter_usuario_autenticado),
):
    """
    Mesma estrutura de 'produção' / 'tipos_tarefa' / 'indicadores' do
    /api/tarefas, mas recortada por um período arbitrário (não só por ano
    inteiro) — usado pelo filtro de período do modo Apresentação, para o
    advogado mostrar só um intervalo específico (ex: 08/08 até 14/09).

    'inicio' e 'fim' no formato AAAA-MM-DD (o mesmo que o input type="date"
    do navegador já manda). Sempre lê a planilha do disco na hora (não usa o
    cache de 10 em 10 minutos), então já reflete a última sincronização.
    Requer login.
    """
    try:
        data_inicio = datetime.strptime(inicio, "%Y-%m-%d")
        data_fim = datetime.strptime(fim, "%Y-%m-%d")
    except ValueError:
        raise HTTPException(status_code=422, detail="As datas devem estar no formato AAAA-MM-DD.")

    if data_fim < data_inicio:
        raise HTTPException(status_code=422, detail="A data final não pode ser anterior à data inicial.")

    try:
        return processar_tarefas_periodo(settings.CAMINHO_PLANILHA_TAREFAS, data_inicio, data_fim)
    except FileNotFoundError:
        raise HTTPException(status_code=503, detail="Planilha de tarefas ainda não foi carregada no servidor.")


@app.get("/api/processos-parados", tags=["Dados"])
def obter_dados_processos_parados(usuario: Usuario = Depends(obter_usuario_autenticado)):
    """Dados da seção 'Processos Parados' do dashboard. Requer login."""
    dados = obter_cache("processos_parados")
    if dados is None:
        raise HTTPException(status_code=503, detail="Dados ainda não sincronizados. Tente novamente em instantes.")
    return dados
