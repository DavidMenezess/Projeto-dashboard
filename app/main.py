"""
main.py

Ponto de entrada da API. Define as rotas:

  POST /auth/login              -> troca Nome + Sobrenome + senha por um token JWT
                                    (bloqueia a conta por alguns minutos após
                                    várias senhas erradas seguidas — ver app/rate_limit.py)
  GET  /me                      -> dados do usuário logado (nome, cargo, tipo)
  GET  /usuarios                -> lista todos os usuários (só admin)
  POST /usuarios                -> cria um novo usuário (só admin)
  PATCH /usuarios/{usuario_id}  -> edita nome, cargo, tipo, status ou senha de um usuário (só admin)
  GET  /api/tarefas             -> dados da seção "Produção 2026" (protegida)
  GET  /api/processos-parados   -> dados da seção "Processos Parados" (protegida)
  GET  /api/processos           -> cadastro geral de processos (protegida)
  GET  /api/clientes            -> visão agregada por cliente, juntando as 3 planilhas (protegida)
  GET  /api/clientes/tarefas-periodo -> tarefas concluídas por cliente num período livre (protegida)
  POST /api/atualizar-planilha  -> substitui uma das 3 planilhas oficiais por um arquivo
                                    enviado pelo usuário (tela "Atualização", protegida)
  GET  /health                  -> checagem simples de saúde do serviço

Rodar localmente:
  uvicorn app.main:app --reload

Rodar em produção: veja o Dockerfile (usa uvicorn sem --reload).
"""

import asyncio
import logging
import shutil
from pathlib import Path
from datetime import datetime, timezone, timedelta
from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy.orm import Session

from app.config import settings
from app.database import criar_tabelas, obter_sessao, Usuario
from app.auth import (
    verificar_senha, gerar_hash_senha, criar_token_acesso,
    obter_usuario_autenticado, exigir_administrador, gerar_chave_login,
)
from app.rate_limit import verificar_bloqueio, registrar_falha, registrar_sucesso
from app.data_processor import (
    processar_tarefas, processar_processos_parados, processar_tarefas_periodo, processar_processos,
    processar_clientes, processar_tarefas_por_cliente_periodo,
    validar_planilha_atualizacao, CONFIG_PLANILHAS_ATUALIZACAO,
)
from app.cache import atualizar_cache, obter_cache, obter_ultima_sincronizacao

# Log estruturado simples. Em produção, isso ajuda a auditar acessos
# (exigência prática da LGPD: saber quem acessou o quê e quando).
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("dashboard-api")

app = FastAPI(
    title="API — Dashboard Directus Advocacia e Consultoria",
    description="Fornece os dados de produção e processos parados ao dashboard, com login por Nome e Sobrenome.",
    version="1.2.0",
    # Em produção (HABILITAR_DOCS=false no .env), desliga a documentação
    # interativa — sem isso, qualquer pessoa sem login consegue ver o mapa
    # completo das rotas da API em /docs. Continua ligada por padrão no
    # desenvolvimento local, onde é útil pra testar.
    docs_url="/docs" if settings.HABILITAR_DOCS else None,
    redoc_url="/redoc" if settings.HABILITAR_DOCS else None,
    openapi_url="/openapi.json" if settings.HABILITAR_DOCS else None,
)

# CORS: só os endereços listados em ORIGENS_PERMITIDAS (.env) podem chamar esta API
# a partir do navegador. Sem isso, qualquer site poderia tentar consumir os dados.
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.lista_origens_permitidas,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PATCH"],
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
        dados_processos_parados = processar_processos_parados(settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS, data_referencia=data_ref)
        atualizar_cache("processos_parados", dados_processos_parados)
        logger.info("Planilha de processos parados processada com sucesso.")
    except FileNotFoundError:
        logger.warning("Planilha de processos parados não encontrada em %s — aguardando arquivo.", settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS)

    try:
        dados_processos = processar_processos(settings.CAMINHO_PLANILHA_PROCESSOS)
        atualizar_cache("processos", dados_processos)
        logger.info("Planilha de processos processada com sucesso.")
    except FileNotFoundError:
        logger.warning("Planilha de processos não encontrada em %s — aguardando arquivo.", settings.CAMINHO_PLANILHA_PROCESSOS)

    _reprocessar_clientes()


def _reprocessar_clientes() -> None:
    """
    Recalcula a visão agregada "Clientes" (ver processar_clientes em
    data_processor.py), que junta as 3 planilhas oficiais pelo nome do
    cliente. Chamada tanto na sincronização periódica/inicial quanto logo
    após qualquer uma das 3 planilhas ser trocada pela tela de Atualização —
    como Clientes depende das 3 ao mesmo tempo, precisa ser reprocessada
    nos dois casos, não só quando "processos" muda.
    """
    try:
        dados_clientes = processar_clientes(
            settings.CAMINHO_PLANILHA_TAREFAS,
            settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS,
            settings.CAMINHO_PLANILHA_PROCESSOS,
        )
        atualizar_cache("clientes", dados_clientes)
        logger.info("Agregação de clientes processada com sucesso.")
    except FileNotFoundError:
        logger.warning("Não foi possível agregar clientes — alguma das 3 planilhas ainda não foi carregada.")


# Caminho oficial (configurado em .env) de cada planilha que a tela de
# Atualização pode substituir — usado tanto para saber onde salvar o arquivo
# enviado quanto, depois, para reprocessá-lo pelo mesmo caminho de sempre.
CAMINHO_POR_TIPO_ATUALIZACAO = {
    "tarefas": settings.CAMINHO_PLANILHA_TAREFAS,
    "processos_parados": settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS,
    "processos": settings.CAMINHO_PLANILHA_PROCESSOS,
}


def _reprocessar_e_atualizar_cache(tipo: str) -> None:
    """
    Reprocessa UM tipo de planilha (a partir do caminho já configurado em
    settings, que a esta altura já foi substituído pelo arquivo novo) e
    atualiza só o cache dela — chamado logo depois que a tela de Atualização
    troca o arquivo, para o dashboard já refletir os dados novos no próximo
    carregamento, sem esperar a sincronização automática de 10 em 10 minutos.
    """
    data_ref = _data_referencia_configurada()
    if tipo == "tarefas":
        atualizar_cache("tarefas", processar_tarefas(settings.CAMINHO_PLANILHA_TAREFAS, data_referencia=data_ref))
    elif tipo == "processos_parados":
        atualizar_cache("processos_parados", processar_processos_parados(settings.CAMINHO_PLANILHA_PROCESSOS_PARADOS, data_referencia=data_ref))
    elif tipo == "processos":
        atualizar_cache("processos", processar_processos(settings.CAMINHO_PLANILHA_PROCESSOS))
    else:
        raise ValueError(f"Tipo desconhecido: {tipo}")

    # "Clientes" agrega as 3 planilhas juntas — qualquer uma delas que for
    # atualizada aqui pode mudar os dados de cliente, então reprocessa
    # também (sem quebrar a atualização da planilha principal se as outras
    # 2 ainda não existirem no servidor — _reprocessar_clientes já trata isso).
    _reprocessar_clientes()


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

    Proteção contra força bruta: depois de MAX_TENTATIVAS senhas erradas
    seguidas para a mesma chave de login (nome.sobrenome), essa conta fica
    bloqueada por alguns minutos, mesmo que a senha certa seja digitada
    nesse meio tempo — ver app/rate_limit.py.
    """
    chave_login = gerar_chave_login(dados.primeiro_nome, dados.sobrenome)

    segundos_restantes = verificar_bloqueio(chave_login)
    if segundos_restantes is not None:
        minutos_restantes = max(1, -(-segundos_restantes // 60))  # arredonda pra cima
        logger.warning(
            "Tentativa de login bloqueada por excesso de tentativas: '%s' (faltam %d min).",
            chave_login, minutos_restantes,
        )
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail=f"Muitas tentativas erradas seguidas. Tente novamente em {minutos_restantes} minuto(s).",
        )

    usuario = sessao.query(Usuario).filter(Usuario.chave_login == chave_login).first()

    credenciais_invalidas = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Nome, sobrenome ou senha incorretos.",
    )

    if usuario is None or not verificar_senha(dados.senha, usuario.senha_hash):
        # Mensagem genérica de propósito: não revela se o usuário existe ou não,
        # o que dificulta um ataque de enumeração de contas.
        registrar_falha(chave_login)
        logger.info("Tentativa de login malsucedida para '%s'.", chave_login)
        raise credenciais_invalidas

    if usuario.ativo != 1:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Usuário desativado.")

    registrar_sucesso(chave_login)
    usuario.ultimo_login = datetime.now(timezone.utc)
    sessao.commit()

    token = criar_token_acesso(usuario.chave_login)
    logger.info("Login bem-sucedido: %s.", usuario.chave_login)
    return {"access_token": token, "token_type": "bearer"}


@app.get("/me", tags=["Autenticação"])
def obter_meu_usuario(usuario: Usuario = Depends(obter_usuario_autenticado)):
    """
    Dados do usuário logado — o dashboard usa isso para saber o nome, cargo
    e se é administrador. Inclui 'id' pra a tela de Usuários saber quando uma
    linha da tabela é a própria conta de quem está logado (usada pra impedir
    que um admin tire o próprio acesso de admin ou bloqueie a si mesmo sem
    querer — ver PATCH /usuarios/{usuario_id}).
    """
    return {
        "id": usuario.id,
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


class AtualizarUsuario(BaseModel):
    """
    Dados para editar um usuário já cadastrado — todos os campos são
    opcionais: manda só o que quer mudar, o resto fica como está.
    """
    primeiro_nome: str | None = None
    sobrenome: str | None = None
    senha: str | None = None  # deixe de fora (ou null) para manter a senha atual
    cargo: str | None = None
    tipo: str | None = None  # "admin" ou "normal"
    ativo: bool | None = None


@app.patch("/usuarios/{usuario_id}", tags=["Usuários"])
def editar_usuario(
    usuario_id: int,
    dados: AtualizarUsuario,
    admin: Usuario = Depends(exigir_administrador),
    sessao: Session = Depends(obter_sessao),
):
    """
    Edita um usuário já cadastrado: nome/sobrenome (o que também muda o
    login, já que ele é gerado a partir disso), cargo, tipo de acesso
    (admin/normal), status (ativo/bloqueado) e/ou senha. Só administradores
    podem editar.

    Por segurança, um administrador não pode mudar o próprio tipo de acesso
    nem o próprio status por aqui — isso evita que alguém se tire de
    administrador ou bloqueie a própria conta sem querer e fique sem
    conseguir desfazer. Se for realmente necessário, outro administrador
    precisa fazer essa alteração específica.
    """
    usuario = sessao.query(Usuario).filter(Usuario.id == usuario_id).first()
    if usuario is None:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")

    eh_a_propria_conta = usuario.id == admin.id

    if dados.tipo is not None:
        if dados.tipo not in ("admin", "normal"):
            raise HTTPException(status_code=422, detail="Campo 'tipo' deve ser 'admin' ou 'normal'.")
        if eh_a_propria_conta and dados.tipo != usuario.tipo:
            raise HTTPException(
                status_code=422,
                detail="Você não pode alterar seu próprio tipo de acesso. Peça para outro administrador fazer essa alteração.",
            )
        usuario.tipo = dados.tipo

    if dados.ativo is not None:
        if eh_a_propria_conta and dados.ativo != bool(usuario.ativo):
            raise HTTPException(
                status_code=422,
                detail="Você não pode ativar/bloquear a própria conta. Peça para outro administrador fazer essa alteração.",
            )
        usuario.ativo = 1 if dados.ativo else 0

    if dados.primeiro_nome is not None or dados.sobrenome is not None:
        novo_primeiro = dados.primeiro_nome if dados.primeiro_nome is not None else usuario.primeiro_nome
        novo_sobrenome = dados.sobrenome if dados.sobrenome is not None else usuario.sobrenome
        nova_chave = gerar_chave_login(novo_primeiro, novo_sobrenome)
        if nova_chave != usuario.chave_login:
            colisao = sessao.query(Usuario).filter(
                Usuario.chave_login == nova_chave, Usuario.id != usuario.id
            ).first()
            if colisao:
                raise HTTPException(
                    status_code=409,
                    detail="Já existe outro usuário com esse nome e sobrenome. Use um sobrenome mais completo para diferenciar.",
                )
        usuario.primeiro_nome = novo_primeiro
        usuario.sobrenome = novo_sobrenome
        usuario.chave_login = nova_chave

    if dados.cargo is not None:
        usuario.cargo = dados.cargo

    if dados.senha is not None and dados.senha != "":
        if len(dados.senha) < 8:
            raise HTTPException(status_code=422, detail="A senha deve ter pelo menos 8 caracteres.")
        usuario.senha_hash = gerar_hash_senha(dados.senha)

    sessao.commit()
    logger.info("Usuário '%s' editado por '%s'.", usuario.chave_login, admin.chave_login)

    return {
        "id": usuario.id,
        "primeiro_nome": usuario.primeiro_nome,
        "sobrenome": usuario.sobrenome,
        "chave_login": usuario.chave_login,
        "cargo": usuario.cargo,
        "tipo": usuario.tipo,
        "ativo": bool(usuario.ativo),
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


@app.get("/api/processos", tags=["Dados"])
def obter_dados_processos(usuario: Usuario = Depends(obter_usuario_autenticado)):
    """
    Cadastro geral de processos (planilha "Processos" do Projuris) — um
    registro por processo, com assunto, situação, justiça, instância, área,
    data de distribuição, valor da causa, data do último andamento,
    cliente(s), polo (autor x réu) e estado/cidade. Requer login.
    """
    dados = obter_cache("processos")
    if dados is None:
        raise HTTPException(status_code=503, detail="Dados ainda não sincronizados. Tente novamente em instantes.")
    return dados


@app.get("/api/clientes", tags=["Dados"])
def obter_dados_clientes(usuario: Usuario = Depends(obter_usuario_autenticado)):
    """
    Visão agregada POR CLIENTE, juntando as 3 planilhas oficiais (Processos,
    Processos Parados e Tarefas) pelo nome do cliente extraído das colunas de
    partes de cada uma — não existe um cadastro de cliente próprio no
    Projuris exportado, nem CPF/CNPJ. Requer login.
    """
    dados = obter_cache("clientes")
    if dados is None:
        raise HTTPException(status_code=503, detail="Dados ainda não sincronizados. Tente novamente em instantes.")
    return dados


@app.get("/api/clientes/tarefas-periodo", tags=["Dados"])
def obter_tarefas_por_cliente_periodo(
    inicio: str,
    fim: str,
    usuario: Usuario = Depends(obter_usuario_autenticado),
):
    """
    Quantas tarefas foram CONCLUÍDAS (pela "Data da conclusão") para cada
    cliente, dentro de um período escolhido livremente — 'inicio' e 'fim' no
    formato AAAA-MM-DD (o mesmo que o input type="date" do navegador manda).
    Sempre lê a planilha de Tarefas do disco na hora (não usa o cache de 10
    em 10 minutos), então já reflete a última sincronização. Requer login.
    """
    try:
        data_inicio = datetime.strptime(inicio, "%Y-%m-%d")
        data_fim = datetime.strptime(fim, "%Y-%m-%d")
    except ValueError:
        raise HTTPException(status_code=422, detail="As datas devem estar no formato AAAA-MM-DD.")

    if data_fim < data_inicio:
        raise HTTPException(status_code=422, detail="A data final não pode ser anterior à data inicial.")

    try:
        return processar_tarefas_por_cliente_periodo(settings.CAMINHO_PLANILHA_TAREFAS, data_inicio, data_fim)
    except FileNotFoundError:
        raise HTTPException(status_code=503, detail="Planilha de tarefas ainda não foi carregada no servidor.")


@app.post("/api/atualizar-planilha", tags=["Dados"])
async def atualizar_planilha(
    tipo: str = Form(...),
    arquivo: UploadFile = File(...),
    usuario: Usuario = Depends(obter_usuario_autenticado),
):
    """
    Tela "Atualização": permite que a própria pessoa do escritório troque uma
    das 3 planilhas oficiais (Tarefas, Processos Parados ou Processos)
    enviando um arquivo novo, sem precisar mexer no servidor manualmente.

    'tipo' deve ser uma das chaves de CONFIG_PLANILHAS_ATUALIZACAO
    ("tarefas", "processos_parados" ou "processos") — a mesma planilha,
    independente do nome do arquivo enviado.

    O arquivo é validado (ver validar_planilha_atualizacao em
    data_processor.py) ANTES de substituir a planilha oficial: confere se
    todas as colunas que o dashboard usa estão presentes, tolerando colunas a
    mais, linhas a mais/a menos e a ordem das colunas — só rejeita quando
    falta alguma coluna necessária (planilha de tipo errado, ou exportada com
    campos removidos). Se for inválida, a planilha atual é mantida intacta e
    o motivo é devolvido para o usuário corrigir (HTTP 422).

    Se for válida: a planilha anterior é guardada com sufixo
    ".bak-<timestamp>" (só por segurança — nunca é lida pelo dashboard), o
    arquivo novo assume o lugar da oficial, os dados são reprocessados na
    hora e o cache é atualizado — o dashboard já reflete a mudança no
    próximo carregamento. Requer login (qualquer usuário, não só admin).
    """
    if tipo not in CONFIG_PLANILHAS_ATUALIZACAO:
        raise HTTPException(
            status_code=422,
            detail=f"Tipo de planilha desconhecido: '{tipo}'. Use um de: {', '.join(CONFIG_PLANILHAS_ATUALIZACAO)}.",
        )

    nome_exibicao = CONFIG_PLANILHAS_ATUALIZACAO[tipo]["nome_exibicao"]
    extensao = Path(arquivo.filename or "").suffix or ".xlsx"

    caminho_destino = Path(CAMINHO_POR_TIPO_ATUALIZACAO[tipo])
    caminho_destino.parent.mkdir(parents=True, exist_ok=True)
    caminho_temporario = caminho_destino.with_name(caminho_destino.stem + ".upload_tmp" + extensao)

    try:
        with open(caminho_temporario, "wb") as f:
            shutil.copyfileobj(arquivo.file, f)
    finally:
        await arquivo.close()

    resultado_validacao = validar_planilha_atualizacao(tipo, str(caminho_temporario))
    if not resultado_validacao["valido"]:
        caminho_temporario.unlink(missing_ok=True)
        logger.warning(
            "Upload de planilha '%s' rejeitado (usuário '%s'): %s",
            tipo, usuario.chave_login, resultado_validacao["motivo"],
        )
        raise HTTPException(status_code=422, detail=resultado_validacao["motivo"])

    # Planilha válida: guarda a anterior como backup (se já existir alguma) e promove a nova.
    if caminho_destino.exists():
        backup = caminho_destino.with_name(
            caminho_destino.stem + f".bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}" + caminho_destino.suffix
        )
        shutil.move(str(caminho_destino), str(backup))
    shutil.move(str(caminho_temporario), str(caminho_destino))

    try:
        _reprocessar_e_atualizar_cache(tipo)
    except Exception as exc:
        # Não deveria acontecer (as colunas já foram validadas acima), mas se
        # algo inesperado quebrar no processamento, é melhor avisar na hora
        # do que deixar o cache desatualizado silenciosamente.
        logger.exception("Falha ao reprocessar a planilha '%s' logo após o upload.", tipo)
        raise HTTPException(
            status_code=500,
            detail=f"O arquivo foi salvo, mas houve um erro ao processá-lo: {exc}. Tente novamente ou avise o suporte.",
        )

    logger.info(
        "Planilha '%s' (%s) atualizada por '%s' — %d linha(s), %d coluna(s) extra(s) ignorada(s).",
        tipo, nome_exibicao, usuario.chave_login,
        resultado_validacao["total_linhas"], len(resultado_validacao["colunas_extras"]),
    )

    return {
        "sucesso": True,
        "tipo": tipo,
        "nome_exibicao": nome_exibicao,
        "total_linhas": resultado_validacao["total_linhas"],
        "colunas_extras_ignoradas": resultado_validacao["colunas_extras"],
    }
