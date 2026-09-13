"""
auth.py

Tudo relacionado a "provar quem é o usuário":
  1. Transformar senha em hash (e conferir senha digitada x hash guardado)
  2. Gerar a "chave de login" a partir de Nome + Sobrenome
  3. Gerar um token JWT quando o login dá certo
  4. Conferir esse token nas rotas protegidas

Nenhuma senha em texto puro passa perto do banco de dados ou dos logs.

Nota sobre a escolha de bcrypt "puro" em vez de passlib: a combinação
passlib + bcrypt tem um problema de compatibilidade conhecido nas versões
mais recentes do pacote bcrypt (passlib tenta ler bcrypt.__about__, que foi
removido). Usar a biblioteca bcrypt diretamente evita essa armadilha e tem
menos uma camada de abstração — mais simples de entender e manter.
"""

import re
import unicodedata
from datetime import datetime, timedelta, timezone
import bcrypt
from jose import jwt, JWTError
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from sqlalchemy.orm import Session

from app.config import settings
from app.database import obter_sessao, Usuario

# Mantido só para o Swagger (/docs) saber onde fica o endpoint de login —
# a rota em si não usa mais o formulário padrão OAuth2 (usuário/senha),
# e sim um corpo JSON com primeiro_nome/sobrenome/senha (veja main.py).
esquema_oauth2 = OAuth2PasswordBearer(tokenUrl="/auth/login")


def normalizar_texto(texto: str) -> str:
    """
    Remove acentos, espaços e caracteres especiais, deixa tudo minúsculo.
    'José' -> 'jose', 'Silva Neto' -> 'silvaneto'.
    """
    texto = texto.strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", texto)


def gerar_chave_login(primeiro_nome: str, sobrenome: str) -> str:
    """
    Constrói a chave única de login a partir de Nome + Sobrenome.
    Ex: 'David' + 'Menezes' -> 'david.menezes'. Duas pessoas com o mesmo
    nome e sobrenome vão colidir aqui — nesse caso, o administrador precisa
    cadastrar com um sobrenome mais completo para diferenciá-las.
    """
    return f"{normalizar_texto(primeiro_nome)}.{normalizar_texto(sobrenome)}"


def gerar_hash_senha(senha_texto_puro: str) -> str:
    """Transforma uma senha digitada em um hash seguro (bcrypt) para salvar no banco."""
    senha_bytes = senha_texto_puro.encode("utf-8")
    hash_bytes = bcrypt.hashpw(senha_bytes, bcrypt.gensalt())
    return hash_bytes.decode("utf-8")


def verificar_senha(senha_texto_puro: str, senha_hash: str) -> bool:
    """Confere se a senha digitada no login bate com o hash salvo no banco."""
    return bcrypt.checkpw(senha_texto_puro.encode("utf-8"), senha_hash.encode("utf-8"))


def criar_token_acesso(chave_login: str) -> str:
    """
    Gera um token JWT contendo a chave de login do usuário e uma data de
    expiração. O token é a "credencial temporária" que o dashboard usa em
    cada requisição depois do login, para não precisar mandar usuário/senha
    toda hora.
    """
    expira_em = datetime.now(timezone.utc) + timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
    dados_do_token = {"sub": chave_login, "exp": expira_em}
    return jwt.encode(dados_do_token, settings.SECRET_KEY, algorithm="HS256")


def obter_usuario_autenticado(
    token: str = Depends(esquema_oauth2),
    sessao: Session = Depends(obter_sessao),
) -> Usuario:
    """
    Dependência usada em toda rota protegida: pega o token enviado pelo
    cliente, valida a assinatura e a validade, e retorna o usuário do banco.
    Se qualquer coisa falhar, devolve 401 (não autorizado) — a rota nem
    chega a executar.
    """
    erro_credenciais = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Não foi possível validar as credenciais. Faça login novamente.",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, settings.SECRET_KEY, algorithms=["HS256"])
        chave_login: str | None = payload.get("sub")
        if chave_login is None:
            raise erro_credenciais
    except JWTError:
        raise erro_credenciais

    usuario = sessao.query(Usuario).filter(Usuario.chave_login == chave_login).first()
    if usuario is None or usuario.ativo != 1:
        raise erro_credenciais

    return usuario


def exigir_administrador(usuario: Usuario = Depends(obter_usuario_autenticado)) -> Usuario:
    """
    Dependência usada nas rotas que só administradores podem acessar
    (criar e listar usuários). Quem não é admin recebe 403 (proibido).
    """
    if usuario.tipo != "admin":
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Apenas administradores podem acessar este recurso.",
        )
    return usuario
