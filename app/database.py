"""
database.py

Banco de dados de USUÁRIOS apenas (quem pode logar no dashboard).

Importante: os dados do dashboard (tarefas, processos parados) NÃO ficam
aqui — eles são recalculados a partir da planilha a cada sincronização e
mantidos em memória (veja data_processor.py). Isso evita guardar dado
processual sensível em um banco permanente sem necessidade, o que ajuda
na conformidade com a LGPD (minimização de dados).
"""

from sqlalchemy import create_engine, Column, Integer, String, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime, timezone

from app.config import settings

# echo=False -> não imprime cada SQL executado no log (deixe True só se for depurar)
engine = create_engine(
    settings.DATABASE_URL,
    connect_args={"check_same_thread": False} if "sqlite" in settings.DATABASE_URL else {},
    echo=False,
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Usuario(Base):
    """
    Um usuário que pode logar no dashboard.

    O login é feito com Nome + Sobrenome (não e-mail) — mais fácil de
    lembrar e digitar no dia a dia. Internamente guardamos uma "chave de
    login" normalizada (sem acento, sem espaço, tudo minúsculo), gerada a
    partir de primeiro_nome + sobrenome, para garantir que cada combinação
    seja única mesmo se a pessoa digitar com acento/maiúsculas diferentes.

    A senha NUNCA é guardada em texto puro — apenas o hash (senha_hash),
    gerado com bcrypt em auth.py. Nem o CONTRATADO nem ninguém com acesso
    ao banco consegue "ver" a senha original.

    Campos de controle de acesso:
      tipo   -> 'admin' (pode criar/listar outros usuários) ou 'normal'
      cargo  -> texto livre, o cargo da pessoa no escritório (ex: "Advogado",
                "Estagiário", "Sócio") — só para exibição, não afeta permissões.
    """
    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True, index=True)
    primeiro_nome = Column(String, nullable=False)
    sobrenome = Column(String, nullable=False)
    chave_login = Column(String, unique=True, index=True, nullable=False)
    senha_hash = Column(String, nullable=False)
    cargo = Column(String, nullable=True)
    tipo = Column(String, nullable=False, default="normal")  # "admin" ou "normal"
    ativo = Column(Integer, default=1)  # 1 = pode logar, 0 = bloqueado
    criado_em = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    ultimo_login = Column(DateTime, nullable=True)


def criar_tabelas():
    """Cria as tabelas no banco se ainda não existirem. Chamado na inicialização da API."""
    Base.metadata.create_all(bind=engine)


def obter_sessao():
    """
    Fornece uma sessão de banco de dados para uma requisição e garante que
    ela seja fechada no final, mesmo se der erro. Usado como dependência do FastAPI.
    """
    sessao = SessionLocal()
    try:
        yield sessao
    finally:
        sessao.close()
