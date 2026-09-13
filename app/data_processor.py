"""
data_processor.py

Transforma as planilhas (Excel) nos dados que o dashboard consome:
KPIs, gráficos e tabelas — a MESMA lógica que já foi validada na versão
estática do dashboard, só reorganizada em funções reutilizáveis.

Este módulo não sabe de onde o arquivo Excel veio (disco local hoje,
OneDrive amanhã) — ele só recebe um caminho de arquivo e devolve dados
prontos. Quando a sincronização com o OneDrive for implementada, ela só
precisa salvar o arquivo baixado em CAMINHO_PLANILHA_* e chamar estas
mesmas funções. Nada aqui precisa mudar.

Duas funções públicas, uma por planilha:
  processar_tarefas(caminho)             -> dados de "Produção" (por ano),
                                             "Tipo de Tarefa" (por ano),
                                             "Indicadores" (por ano),
                                             pendências (atual, todos os anos)
                                             e o detalhamento por tarefa.
  processar_processos_parados(caminho)   -> dados da seção "Processos Parados",
                                             separados em "recentes" (90-120
                                             dias) e "cronicos" (mais de 120).

Sobre os campos "Indicadores": foram adicionados depois de uma checagem de
quanto cada coluna da planilha está realmente preenchida — só entrou o que
tem dado suficiente para ser confiável (ver conversa com o cliente). Campos
como "Total de horas do timesheet" (0% preenchido), "Instância", "Fase" e
"Vara" (10-15% preenchido) ficaram de fora por enquanto: o Projuris ainda
não está sendo alimentado com esses dados de forma consistente. Quando
passar a ser, essas funções são o lugar certo para adicioná-los.
"""

import re
import pandas as pd
from datetime import datetime


# ---------------------------------------------------------------------------
# Classificação de polo processual (autor x réu). Usada tanto nos processos
# parados quanto (com uma fonte de dado diferente) na produção geral — ver
# _classifica_polo_producao mais abaixo.
# ---------------------------------------------------------------------------
PAPEIS_POLO_ATIVO = {
    "autor", "requerente", "exequente", "reclamante", "apelante",
    "agravante", "recorrente", "embargante", "impugnante", "impetrante", "excipiente",
}
PAPEIS_POLO_PASSIVO = {
    "réu", "reu", "requerido", "executado", "reclamado", "apelado", "agravado",
    "recorrido", "embargado", "impugnado", "denunciado", "acusado", "investigado",
    "sujeito passivo", "demandado",
}

MESES_NOME = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
              7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}


def _extrai_papeis(texto) -> list[str]:
    """Extrai os papéis entre parênteses: 'Fulano (Autor), Beltrano (Autor)' -> ['autor', 'autor']."""
    if pd.isna(texto):
        return []
    return [p.strip().lower() for p in re.findall(r"\(([^)]+)\)", str(texto))]


def _classifica_polo(papeis: list[str]) -> str:
    """
    Usado na planilha de Processos Parados, onde a coluna 'Envolvidos cliente'
    já vem filtrada só com a parte que o escritório representa.
    """
    tem_ativo = any(p in PAPEIS_POLO_ATIVO for p in papeis)
    tem_passivo = any(p in PAPEIS_POLO_PASSIVO for p in papeis)
    if tem_ativo and tem_passivo:
        return "Misto"
    if tem_ativo:
        return "Polo Ativo (Autor)"
    if tem_passivo:
        return "Polo Passivo (Réu)"
    return "Não identificado"


def _lado_tem_cliente(texto) -> bool:
    """
    Usado na planilha de Tarefas, onde não existe uma coluna só com o
    envolvido-cliente — em vez disso, as colunas 'Envolvidos do processo
    (partes ativas)' e '(partes passivas)' listam todo mundo, e o nome do
    cliente do escritório vem marcado com o sufixo '- Cliente' dentro do
    parênteses (ex: 'Fulano (Autor - Cliente)'). Essa função só confere se
    ESSE lado específico (ativo OU passivo) tem alguém marcado como cliente.
    """
    if pd.isna(texto):
        return False
    return "cliente" in str(texto).lower()


def _extrai_nomes_clientes(texto) -> list[str]:
    """
    Extrai o(s) NOME(S) de quem está marcado como cliente do escritório
    numa das colunas de partes do processo (ex: 'Fulano (Autor - Cliente) |
    Beltrano (Autor)' -> ['Fulano']). Usada para contar clientes distintos
    de um jeito mais completo do que a coluna 'Envolvidos do atendimento
    (clientes)', que só é preenchida em tarefas administrativas (cerca de
    9% da planilha) — via partes ativas/passivas, a cobertura sobe para
    mais de 80% das tarefas.
    """
    if pd.isna(texto):
        return []
    nomes = []
    for parte in str(texto).split("|"):
        m = re.match(r"^(.*?)\s*\(([^)]*)\)\s*$", parte.strip())
        if m and "cliente" in m.group(2).lower():
            nomes.append(m.group(1).strip())
    return nomes


def _classifica_polo_producao(tem_cliente_ativo: bool, tem_cliente_passivo: bool) -> str:
    """Mesma ideia de _classifica_polo, mas a partir das duas colunas de partes da planilha de Tarefas."""
    if tem_cliente_ativo and tem_cliente_passivo:
        return "Misto"
    if tem_cliente_ativo:
        return "Polo Ativo (Autor)"
    if tem_cliente_passivo:
        return "Polo Passivo (Réu)"
    return "Não identificado"


def _extrai_area_direito(assunto) -> str:
    """
    Reduz o campo 'Assunto' (que vem como uma classificação jurídica às
    vezes bem longa, ex: 'DIREITO CIVIL (899) - Coisas (10432) - Propriedade
    (10448)...') para só a primeira categoria, mais curta e legível em um
    gráfico. Nem todo valor tem essa estrutura hierárquica — quando não tem
    (ex: 'Acidente de Trânsito'), o valor original é mantido como está.
    """
    if pd.isna(assunto):
        return "Não informado"
    texto = str(assunto)
    primeira_parte = re.split(r"\s[-/]\s", texto, maxsplit=1)[0]
    primeira_parte = re.sub(r"\s*\(\d+\)\s*$", "", primeira_parte).strip()
    if len(primeira_parte) > 45:
        primeira_parte = primeira_parte[:45] + "..."
    return primeira_parte or "Não informado"


def _texto_seguro(valor, tamanho_max: int | None = None) -> str:
    """Converte um valor de célula em string segura para JSON (nunca NaN, nunca None cru)."""
    if pd.isna(valor):
        return ""
    texto = str(valor)
    if tamanho_max and len(texto) > tamanho_max:
        texto = texto[:tamanho_max] + "..."
    return texto


def _categoriza_tarefa(modulo) -> str:
    """Processual (vinculada a um processo), Atendimento (administrativa) ou Outras (sem vínculo)."""
    if modulo == "Processo":
        return "Processual"
    if modulo == "Atendimento":
        return "Atendimento/Administrativa"
    return "Outras (sem vínculo)"


# ---------------------------------------------------------------------------
# Planilha de TAREFAS
# ---------------------------------------------------------------------------

def processar_tarefas(caminho_arquivo: str, data_referencia: datetime | None = None) -> dict:
    """
    Lê a planilha de tarefas e devolve:
      anos_disponiveis      -> lista de anos encontrados na planilha (ex: [2026, 2027])
      por_ano               -> {"2026": {...produção...}, "2027": {...}, ...}
      tipos_tarefa_por_ano  -> {"2026": {...}, "2027": {...}, ...}
      indicadores_por_ano   -> {"2026": {...ciclo, polo, área do direito, delegação...}, ...}
      pendencias            -> vencidas/a vencer, sempre o quadro atual (não muda por ano)
      detalhamento          -> lista de todas as tarefas, de todos os anos, com o campo "ano"
    """
    hoje = pd.Timestamp((data_referencia or datetime.now()).date())

    df = pd.read_excel(caminho_arquivo, sheet_name="Tarefas", header=2)

    for coluna in ["Data prevista", "Data fatal", "Data da conclusão", "Data de criação"]:
        df[coluna + "_dt"] = pd.to_datetime(df[coluna], format="%d/%m/%Y", errors="coerce")

    df["categoria"] = df["Módulo"].apply(_categoriza_tarefa)
    df["ano"] = df["Data prevista_dt"].dt.year

    anos_disponiveis = sorted(int(a) for a in df["ano"].dropna().unique())

    por_ano = {}
    tipos_tarefa_por_ano = {}
    indicadores_por_ano = {}
    for ano in anos_disponiveis:
        df_ano = df[df["ano"] == ano]
        por_ano[str(ano)] = _producao_do_ano(df_ano)
        tipos_tarefa_por_ano[str(ano)] = _tipos_tarefa(df_ano)
        indicadores_por_ano[str(ano)] = _indicadores_do_ano(df_ano)

    return {
        "atualizado_em": hoje.strftime("%d/%m/%Y"),
        "anos_disponiveis": anos_disponiveis,
        "por_ano": por_ano,
        "tipos_tarefa_por_ano": tipos_tarefa_por_ano,
        "indicadores_por_ano": indicadores_por_ano,
        "pendencias": _calcular_pendencias(df, hoje),
        "detalhamento": _detalhamento_tarefas(df),
    }


def _producao_do_ano(df: pd.DataFrame) -> dict:
    """KPIs, evolução mensal, responsáveis e tipos de tarefa — tudo restrito a um único ano."""
    situacao_counts = df["Situação"].value_counts().to_dict()

    # --- Evolução mensal — sempre os 12 meses, mesmo que o ano ainda não tenha
    # terminado (ex: ano corrente) ou que algum mês não tenha nenhuma tarefa.
    df = df.copy()
    df["mes_num"] = df["Data prevista_dt"].dt.month
    mensal = df.groupby("mes_num").agg(
        total=("Identificador da tarefa", "count"),
        concluidas=("Situação", lambda s: (s == "Concluída com sucesso").sum()),
        pendentes=("Situação", lambda s: s.isin(["Pendente", "Em execução"]).sum()),
        canceladas=("Situação", lambda s: (s == "Cancelado").sum()),
    )
    mensal = mensal.reindex(range(1, 13), fill_value=0)
    mensal["mes_nome"] = [MESES_NOME[m] for m in mensal.index]

    # --- Por responsável (todos, não só um "top N") ---
    responsaveis = df["Responsáveis da tarefa"].dropna().astype(str).str.split("|").explode().str.strip()
    resp_counts = responsaveis.value_counts().to_dict()

    resp_situacao = {}
    for responsavel in resp_counts:
        filtro = df["Responsáveis da tarefa"].astype(str).str.contains(responsavel, na=False, regex=False)
        subconjunto = df[filtro]
        resp_situacao[responsavel] = {
            "concluidas": int((subconjunto["Situação"] == "Concluída com sucesso").sum()),
            "pendentes": int(subconjunto["Situação"].isin(["Pendente", "Em execução"]).sum()),
            "canceladas": int((subconjunto["Situação"] == "Cancelado").sum()),
            "total": int(len(subconjunto)),
        }

    tipo_counts = df["Tipo de tarefa"].value_counts().head(10).to_dict()
    modulo_counts = df["Módulo"].value_counts().to_dict()
    sitproc_counts = df["Situação do processo"].value_counts().to_dict()

    grupos = df["Grupos de trabalho"].dropna().astype(str).str.split("|").explode().str.strip()
    grupo_counts = grupos.value_counts().head(8).to_dict()

    # Conta clientes distintos a partir de QUEM está marcado como cliente nas
    # partes do processo (ativas + passivas) — dá uma cobertura bem maior do
    # que a coluna "Envolvidos do atendimento (clientes)", que só existe em
    # tarefas administrativas (uma fração pequena da planilha).
    nomes_clientes = set()
    for col in ["Envolvidos do processo (partes ativas)", "Envolvidos do processo (partes passivas)"]:
        for lista in df[col].apply(_extrai_nomes_clientes):
            nomes_clientes.update(lista)

    return {
        "total": int(len(df)),
        "situacao_counts": {k: int(v) for k, v in situacao_counts.items()},
        "monthly": mensal[["mes_nome", "total", "concluidas", "pendentes", "canceladas"]].to_dict("records"),
        "resp_counts": {k: int(v) for k, v in resp_counts.items()},
        "resp_situacao": resp_situacao,
        "tipo_counts": {k: int(v) for k, v in tipo_counts.items()},
        "modulo_counts": {k: int(v) for k, v in modulo_counts.items()},
        "sitproc_counts": {k: int(v) for k, v in sitproc_counts.items()},
        "grupo_counts": {k: int(v) for k, v in grupo_counts.items()},
        "processos_distintos": int(df["Identificador do módulo"].nunique()),
        "clientes_distintos": len(nomes_clientes),
    }


def _calcular_pendencias(df: pd.DataFrame, hoje: pd.Timestamp) -> dict:
    """
    Pendências vencidas e a vencer — sempre relativas a HOJE, olhando a
    planilha inteira (todos os anos). Uma tarefa vencida de 2026 continua
    aparecendo aqui em 2027 enquanto não for resolvida; não faz sentido
    "escondê-la" só porque o ano mudou.
    """
    pendentes = df[df["Situação"].isin(["Pendente", "Em execução"])].copy()

    vencidas = pendentes[pendentes["Data fatal_dt"] < hoje].copy()
    vencidas["dias_atraso"] = (hoje - vencidas["Data fatal_dt"]).dt.days
    vencidas = vencidas.sort_values("dias_atraso", ascending=False)

    a_vencer = pendentes[pendentes["Data fatal_dt"] >= hoje].copy()
    a_vencer["dias_restantes"] = (a_vencer["Data fatal_dt"] - hoje).dt.days
    a_vencer = a_vencer.sort_values("dias_restantes").head(60)

    def _linha_pendencia(row, campo_dias):
        return {
            "id": _texto_seguro(row["Identificador da tarefa"]),
            "tipo": _texto_seguro(row["Tipo de tarefa"]),
            "titulo": _texto_seguro(row["Título"]),
            "responsavel": _texto_seguro(row["Responsáveis da tarefa"]),
            "data_fatal": _texto_seguro(row["Data fatal"]),
            "situacao": _texto_seguro(row["Situação"]),
            campo_dias: int(row[campo_dias]),
            "processo": _texto_seguro(row["Número do processo"]),
            "orgao": _texto_seguro(row["Órgão"]),
        }

    return {
        "overdue_list": [_linha_pendencia(r, "dias_atraso") for _, r in vencidas.iterrows()],
        "upcoming_list": [_linha_pendencia(r, "dias_restantes") for _, r in a_vencer.iterrows()],
        "total_atrasadas": int(len(vencidas)),
        "total_pendentes": int(len(pendentes)),
    }


def _tipos_tarefa(df: pd.DataFrame) -> dict:
    """Separação processual x administrativa x outras, geral e por responsável — restrito a um ano."""
    cat_counts = df["categoria"].value_counts().to_dict()

    responsaveis = df["Responsáveis da tarefa"].dropna().astype(str).str.split("|").explode().str.strip().unique()
    resp_categoria = {}
    for responsavel in responsaveis:
        filtro = df["Responsáveis da tarefa"].astype(str).str.contains(responsavel, na=False, regex=False)
        subconjunto = df[filtro]
        resp_categoria[responsavel] = {
            "processual": int((subconjunto["categoria"] == "Processual").sum()),
            "atendimento": int((subconjunto["categoria"] == "Atendimento/Administrativa").sum()),
            "outras": int((subconjunto["categoria"] == "Outras (sem vínculo)").sum()),
        }

    return {
        "cat_counts": {k: int(v) for k, v in cat_counts.items()},
        "resp_categoria": resp_categoria,
    }


def _indicadores_do_ano(df: pd.DataFrame) -> dict:
    """
    Indicadores mais analíticos, restritos a um único ano:
      - tempo de ciclo (dias entre a criação da tarefa e a conclusão dela)
      - polo no processo (autor x réu), aplicado à produção geral, não só
        aos processos parados
      - área do direito mais frequente (a partir do campo "Assunto")
      - quem mais cria tarefas, e quanto disso é delegado para outra pessoa
        em vez de ficar com quem criou
    """
    df = df.copy()

    # --- Tempo de ciclo: só dá pra calcular em tarefas já concluídas ---
    concluidas = df[df["Data da conclusão_dt"].notna() & df["Data de criação_dt"].notna()].copy()
    concluidas["dias_ciclo"] = (concluidas["Data da conclusão_dt"] - concluidas["Data de criação_dt"]).dt.days
    concluidas = concluidas[concluidas["dias_ciclo"] >= 0]  # descarta datas inconsistentes, se houver

    ciclo_medio_geral = float(concluidas["dias_ciclo"].mean()) if len(concluidas) else None

    ciclo_por_responsavel = {}
    if len(concluidas):
        responsaveis_ciclo = concluidas["Responsáveis da tarefa"].dropna().astype(str).str.split("|").explode().str.strip().unique()
        for responsavel in responsaveis_ciclo:
            filtro = concluidas["Responsáveis da tarefa"].astype(str).str.contains(responsavel, na=False, regex=False)
            media = concluidas.loc[filtro, "dias_ciclo"].mean()
            if pd.notna(media):
                ciclo_por_responsavel[responsavel] = round(float(media), 1)
    # Mantém só os 10 com mais tarefas concluídas consideradas, pra não poluir o gráfico
    ciclo_por_responsavel = dict(sorted(ciclo_por_responsavel.items(), key=lambda kv: kv[1])[:15])

    # --- Polo no processo (autor x réu) na produção geral ---
    tem_cliente_ativo = df["Envolvidos do processo (partes ativas)"].apply(_lado_tem_cliente)
    tem_cliente_passivo = df["Envolvidos do processo (partes passivas)"].apply(_lado_tem_cliente)
    polo = [
        _classifica_polo_producao(a, p)
        for a, p in zip(tem_cliente_ativo, tem_cliente_passivo)
    ]
    polo_counts = pd.Series(polo).value_counts().to_dict()

    # --- Área do direito (a partir do "Assunto") ---
    area_direito = df["Assunto"].apply(_extrai_area_direito)
    area_direito_counts = area_direito[area_direito != "Não informado"].value_counts().head(10).to_dict()

    # --- Quem mais cria tarefas, e o quanto disso é delegado ---
    criador_counts = df["Criada por"].dropna().value_counts().head(10).to_dict()

    def _e_autoatribuida(row) -> bool:
        criador = row["Criada por"]
        responsaveis = row["Responsáveis da tarefa"]
        if pd.isna(criador) or pd.isna(responsaveis):
            return False
        lista_responsaveis = [r.strip() for r in str(responsaveis).split("|")]
        return criador.strip() in lista_responsaveis

    com_criador_e_responsavel = df[df["Criada por"].notna() & df["Responsáveis da tarefa"].notna()]
    autoatribuidas = int(com_criador_e_responsavel.apply(_e_autoatribuida, axis=1).sum())
    delegadas = int(len(com_criador_e_responsavel)) - autoatribuidas

    return {
        "ciclo_medio_geral": round(ciclo_medio_geral, 1) if ciclo_medio_geral is not None else None,
        "ciclo_por_responsavel": ciclo_por_responsavel,
        "polo_counts": {k: int(v) for k, v in polo_counts.items()},
        "area_direito_counts": {k: int(v) for k, v in area_direito_counts.items()},
        "criador_counts": {k: int(v) for k, v in criador_counts.items()},
        "autoatribuidas": autoatribuidas,
        "delegadas": delegadas,
    }


def _detalhamento_tarefas(df: pd.DataFrame) -> list[dict]:
    """
    Uma linha por tarefa, de TODOS os anos, para a tabela pesquisável
    'qual tarefa cada responsável fez'. O campo "ano" permite ao dashboard
    filtrar essa tabela por ano, mesmo ela vindo completa da API.
    """
    linhas = []
    for _, row in df.iterrows():
        linhas.append({
            "responsavel": _texto_seguro(row["Responsáveis da tarefa"]) or "Não atribuído",
            "categoria": row["categoria"],
            "tipo": _texto_seguro(row["Tipo de tarefa"]),
            "titulo": _texto_seguro(row["Título"], tamanho_max=70),
            "processo": _texto_seguro(row["Número do processo"]),
            "situacao": _texto_seguro(row["Situação"]),
            "data_prevista": _texto_seguro(row["Data prevista"]),
            "data_conclusao": _texto_seguro(row["Data da conclusão"]),
            "ano": int(row["ano"]) if pd.notna(row["ano"]) else None,
        })
    return linhas


# ---------------------------------------------------------------------------
# Planilha de PROCESSOS PARADOS
# ---------------------------------------------------------------------------

def processar_processos_parados(caminho_arquivo: str, data_referencia: datetime | None = None) -> dict:
    """
    Lê a planilha "processos parados" exportada do Projuris e devolve os
    dados da seção "Processos Parados", divididos em dois grupos:

      recentes -> parados entre 90 e 120 dias (os que acabaram de cruzar o limite)
      cronicos -> parados há MAIS de 120 dias

    Antes só existia o grupo "recentes" — um processo parado há 300 dias
    simplesmente não aparecia em lugar nenhum do dashboard. Isso foi
    corrigido para não esconder justamente os casos mais graves.
    """
    hoje = pd.Timestamp((data_referencia or datetime.now()).date())

    df = pd.read_excel(caminho_arquivo, header=18)
    df = df[df["Situação processo"] != "Usuário emissor:"].copy()

    df["data_mov_dt"] = pd.to_datetime(df["Data último movimento"], errors="coerce")
    df["dias_parado"] = (hoje - df["data_mov_dt"]).dt.days
    df["data_distribuicao_dt"] = pd.to_datetime(df["Data distribuição"], errors="coerce")
    df["idade_processo_dias"] = (hoje - df["data_distribuicao_dt"]).dt.days

    df["papeis"] = df["Envolvidos cliente"].apply(_extrai_papeis)
    df["polo"] = df["papeis"].apply(_classifica_polo)

    df_parados = df[df["dias_parado"] >= 90].copy()
    df_recentes = df_parados[df_parados["dias_parado"] <= 120].copy()
    df_cronicos = df_parados[df_parados["dias_parado"] > 120].copy()

    return {
        "atualizado_em": hoje.strftime("%d/%m/%Y"),
        "recentes": _construir_bloco_parados(df_recentes),
        "cronicos": _construir_bloco_parados(df_cronicos),
    }


def _construir_bloco_parados(df: pd.DataFrame) -> dict:
    """Monta o dicionário de saída (KPIs, gráficos e lista) para um grupo de processos parados."""
    # Prioriza polo ativo primeiro (ver Cláusula 2ª do contrato / decisão de negócio
    # de que processos em que o escritório representa o autor precisam de atenção
    # mais urgente do que aqueles em que representa o réu).
    ordem_polo = {"Polo Ativo (Autor)": 0, "Misto": 1, "Polo Passivo (Réu)": 2, "Não identificado": 3}
    df = df.copy()
    df["polo_rank"] = df["polo"].map(ordem_polo)
    df = df.sort_values(["polo_rank", "dias_parado"], ascending=[True, False])

    area_counts = df["Área"].fillna("Não informado").value_counts().head(10).to_dict()
    orgao_counts = df["Orgão"].fillna("Não informado").value_counts().head(8).to_dict()

    responsaveis = df["Usuários responsáveis"].dropna().astype(str).str.split(",").explode().str.strip()
    resp_counts = responsaveis.value_counts().to_dict()

    idades_validas = df["idade_processo_dias"].dropna()
    idade_media_anos = round(float(idades_validas.mean()) / 365, 1) if len(idades_validas) else None
    idade_cobertura_pct = round(len(idades_validas) / len(df) * 100) if len(df) else 0

    lista = []
    for _, row in df.iterrows():
        lista.append({
            "processo": _texto_seguro(row["Numero processo"]),
            "assunto": _texto_seguro(row["Assunto"], tamanho_max=80),
            "orgao": _texto_seguro(row["Orgão"]),
            "area": _texto_seguro(row["Área"]),
            "situacao": _texto_seguro(row["Situação processo"]),
            "responsavel": _texto_seguro(row["Usuários responsáveis"]),
            "data_mov": _texto_seguro(row["Data último movimento"]),
            "dias_parado": int(row["dias_parado"]) if pd.notna(row["dias_parado"]) else None,
            "polo": row["polo"],
            "representando": _texto_seguro(row["Envolvidos cliente"], tamanho_max=55),
            "idade_processo_anos": round(row["idade_processo_dias"] / 365, 1) if pd.notna(row["idade_processo_dias"]) else None,
        })

    return {
        "total": int(len(df)),
        "situacao_counts": {k: int(v) for k, v in df["Situação processo"].value_counts().to_dict().items()},
        "area_counts": {k: int(v) for k, v in area_counts.items()},
        "orgao_counts": {k: int(v) for k, v in orgao_counts.items()},
        "resp_counts": {k: int(v) for k, v in resp_counts.items()},
        "polo_counts": {k: int(v) for k, v in df["polo"].value_counts().to_dict().items()},
        "idade_media_anos": idade_media_anos,
        "idade_cobertura_pct": idade_cobertura_pct,
        "lista": lista,
    }
