"""
cache.py

Guarda em memória o resultado mais recente do processamento das planilhas.

Por que em memória e não em banco de dados:
1) Os dados já vêm de uma fonte de verdade (o Excel no OneDrive) — não faz
   sentido duplicar isso em outro lugar permanente.
2) Menos um lugar guardando dado processual sensível = mais fácil de
   justificar a conformidade com a LGPD (minimização de dados).
3) Simplicidade: reinicia o container, os dados são recarregados da
   planilha na próxima sincronização. Nada se perde de fato.
"""

from datetime import datetime, timezone
from threading import Lock

# Lock evita que o job de sincronização e uma requisição HTTP leiam/escrevam
# o cache ao mesmo tempo e causem inconsistência.
_trava = Lock()

_cache: dict = {
    "tarefas": None,
    "processos_parados": None,
    "processos": None,
    "ultima_sincronizacao": None,
}


def atualizar_cache(chave: str, dados: dict) -> None:
    """Chamado pelo job de sincronização depois de processar uma planilha."""
    with _trava:
        _cache[chave] = dados
        _cache["ultima_sincronizacao"] = datetime.now(timezone.utc).isoformat()


def obter_cache(chave: str) -> dict | None:
    """Chamado pelas rotas da API para servir os dados já processados."""
    with _trava:
        return _cache.get(chave)


def obter_ultima_sincronizacao() -> str | None:
    with _trava:
        return _cache.get("ultima_sincronizacao")
