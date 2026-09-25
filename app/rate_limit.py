"""
rate_limit.py

Proteção simples contra força bruta no login: limita quantas tentativas
erradas de senha uma mesma chave de login (nome.sobrenome) pode ter num
intervalo de tempo, e bloqueia temporariamente depois disso.

Por que em memória e não em banco de dados/Redis:
- Só um container roda a API (não é um cluster com várias instâncias ao
  mesmo tempo), então não precisa de um armazenamento compartilhado entre
  processos.
- Se o container reiniciar, os contadores zeram — o pior caso é alguém
  ganhar mais algumas tentativas de volta, não é uma falha de segurança
  grave, só um detalhe operacional.

Por que por chave_login e não por IP:
- O IP de quem acessa pode mudar (4G, redes de escritório com NAT) e um
  único IP pode representar várias pessoas legítimas ao mesmo tempo — travar
  por IP arrisca bloquear gente de verdade. Travar pela conta que está
  sendo alvo do ataque é mais direto: é exatamente essa conta que precisa de
  proteção contra tentativa e erro de senha.
"""

from datetime import datetime, timezone, timedelta
from threading import Lock

MAX_TENTATIVAS = 5
JANELA_MINUTOS = 15
BLOQUEIO_MINUTOS = 15

_trava = Lock()
_tentativas: dict[str, list[datetime]] = {}
_bloqueados: dict[str, datetime] = {}


def verificar_bloqueio(chave_login: str) -> int | None:
    """
    Retorna quantos segundos faltam para essa chave de login poder tentar
    de novo, ou None se não está bloqueada agora.
    """
    with _trava:
        expira_em = _bloqueados.get(chave_login)
        if expira_em is None:
            return None
        agora = datetime.now(timezone.utc)
        if agora >= expira_em:
            # Bloqueio já venceu — libera e limpa o histórico, começando do zero.
            _bloqueados.pop(chave_login, None)
            _tentativas.pop(chave_login, None)
            return None
        return int((expira_em - agora).total_seconds())


def registrar_falha(chave_login: str) -> None:
    """
    Registra uma tentativa de login mal-sucedida. Se essa chave de login
    acumular MAX_TENTATIVAS erradas dentro da janela de tempo, bloqueia por
    BLOQUEIO_MINUTOS.
    """
    with _trava:
        agora = datetime.now(timezone.utc)
        limite_janela = agora - timedelta(minutes=JANELA_MINUTOS)
        lista = [t for t in _tentativas.get(chave_login, []) if t > limite_janela]
        lista.append(agora)
        _tentativas[chave_login] = lista
        if len(lista) >= MAX_TENTATIVAS:
            _bloqueados[chave_login] = agora + timedelta(minutes=BLOQUEIO_MINUTOS)


def registrar_sucesso(chave_login: str) -> None:
    """Login certo — zera qualquer histórico de tentativas erradas dessa chave."""
    with _trava:
        _tentativas.pop(chave_login, None)
        _bloqueados.pop(chave_login, None)
