"""
config.py

Ponto único de leitura das variáveis de ambiente (arquivo .env).

Por que isso existe separado:
Assim nenhuma outra parte do código lê o .env diretamente — todo mundo importa
o objeto `settings` daqui. Se um dia mudar de onde vem a configuração (ex: de
.env para um serviço de segredos na nuvem), só este arquivo muda.
"""

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    # --- Segurança / login ---
    SECRET_KEY: str
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 60

    # --- Banco de dados de usuários ---
    DATABASE_URL: str = "sqlite:///./data/usuarios.db"

    # --- Arquivos de dados (planilhas) ---
    CAMINHO_PLANILHA_TAREFAS: str = "./data/tarefas.xlsx"
    CAMINHO_PLANILHA_PROCESSOS_PARADOS: str = "./data/processos_parados.xlsx"

    # --- CORS ---
    ORIGENS_PERMITIDAS: str = "http://localhost:3000"

    # --- Apenas para testes locais com planilhas antigas ---
    # Se preenchida (formato DD/MM/AAAA), a API usa essa data como "hoje" ao
    # calcular prazos e o filtro de "90 a 120 dias parado", em vez da data
    # real do sistema. Deixe em branco em produção — sem isso, a API sempre
    # usa a data real, que é o comportamento correto quando a planilha está
    # sendo atualizada de verdade.
    DATA_REFERENCIA_TESTE: str = ""

    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8")

    @property
    def lista_origens_permitidas(self) -> list[str]:
        """Transforma 'a,b,c' do .env em ['a', 'b', 'c'] para o FastAPI."""
        return [origem.strip() for origem in self.ORIGENS_PERMITIDAS.split(",") if origem.strip()]


# Instância única, importada por todo o resto do projeto.
settings = Settings()
