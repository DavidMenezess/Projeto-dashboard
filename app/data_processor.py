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
import unicodedata
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

# Trava de sanidade para "Valor ação" (valor da causa) na planilha de
# Processos: já apareceu um valor de R$ 41 quatrilhões numa exportação real
# (claramente um erro de digitação/exportação lá no Projuris — o maior valor
# "de verdade" na mesma planilha era ~R$ 108 milhões). Qualquer valor acima
# deste limite é tratado como erro de digitação: continua aparecendo no
# registro do processo (pra dar pra achar e corrigir na fonte), mas é
# excluído do TOTAL somado, pra um erro de digitação não inflar o KPI a
# ponto de ficar sem sentido nenhum pro cliente.
LIMITE_SANIDADE_VALOR_CAUSA = 1_000_000_000  # R$ 1 bilhão


# ---------------------------------------------------------------------------
# Leitura de Excel "à prova de extensão errada" — usada por todo o resto do
# módulo. Motivo: a tela de Atualização (ver validar_planilha_atualizacao,
# mais abaixo) deixa o cliente enviar uma planilha nova para SUBSTITUIR uma
# das 3 planilhas oficiais, e o arquivo que ele exporta do Projuris pode vir
# num formato (.xlsx, que é um ZIP por dentro) diferente do que o caminho
# configurado no servidor sugere pelo nome (ex: "processos_parados.xls",
# formato antigo OLE2) — ou vice-versa. Sem isso, o pandas escolhe o motor de
# leitura só pela extensão do NOME do arquivo, e pode tentar ler um ZIP com o
# motor de arquivo antigo (xlrd) ou um arquivo antigo com o motor novo
# (openpyxl), quebrando mesmo com uma planilha perfeitamente válida.
# ---------------------------------------------------------------------------

def _detectar_engine_excel(caminho_arquivo: str) -> str | None:
    """Olha os primeiros bytes do arquivo para descobrir o formato de verdade, em vez de confiar na extensão do nome. Devolve None (deixa o pandas decidir pela extensão, como antes) quando a assinatura não é reconhecida."""
    try:
        with open(caminho_arquivo, "rb") as f:
            assinatura = f.read(8)
    except OSError:
        return None
    if assinatura.startswith(b"PK\x03\x04"):
        return "openpyxl"  # .xlsx / .xlsm (é um ZIP por dentro)
    if assinatura.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "xlrd"  # .xls antigo (formato OLE2)
    return None


def _ler_excel(caminho_arquivo: str, **kwargs) -> pd.DataFrame:
    """Wrapper de pd.read_excel que escolhe o motor pelo CONTEÚDO real do arquivo (ver _detectar_engine_excel) em vez da extensão do nome — todo o resto do módulo lê arquivos Excel por aqui, nunca direto com pd.read_excel."""
    engine = _detectar_engine_excel(caminho_arquivo)
    if engine:
        kwargs.setdefault("engine", engine)
    return pd.read_excel(caminho_arquivo, **kwargs)


def _localizar_linha_cabecalho(caminho_arquivo: str, coluna_marcador: str, sheet_name: str | None = None, max_linhas: int = 40) -> int | None:
    """
    Acha em qual linha (0-indexed) está o cabeçalho de verdade de uma
    planilha, procurando a primeira linha que contém `coluna_marcador` — em
    vez de confiar num número de linha fixo, já que o Projuris varia a
    quantidade de linhas de metadados no topo do arquivo ("Filtros
    utilizados:", "Data início:" etc.) conforme os filtros usados na
    exportação.

    Função genérica por trás de _localizar_linha_cabecalho_parados,
    _localizar_linha_cabecalho_processos e _localizar_linha_cabecalho_tarefas
    logo abaixo, e também usada por validar_planilha_atualizacao (no fim do
    arquivo) para conferir se um arquivo enviado pelo usuário na tela de
    Atualização tem o formato esperado antes de aceitá-lo.

    Devolve None se não encontrar a coluna nas primeiras `max_linhas` linhas
    — quem chama decide o que fazer: um valor de compatibilidade, no
    processamento normal de todo dia, ou reportar "planilha inválida", na
    validação de um arquivo novo.
    """
    kwargs = {"header": None, "nrows": max_linhas}
    if sheet_name:
        kwargs["sheet_name"] = sheet_name
    bruto = _ler_excel(caminho_arquivo, **kwargs)
    for i in range(len(bruto)):
        if coluna_marcador in bruto.iloc[i].astype(str).values:
            return i
    return None


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


def _extrai_nomes_papeis(texto) -> list[tuple[str, str]]:
    """
    Extrai pares (nome, papel) de um texto no formato 'Fulano (Autor), Empresa
    X (Réu)' — usado nas colunas 'Cliente' (planilha Processos) e 'Envolvidos
    cliente' (Processos Parados), que já vêm só com o(s) envolvido(s) que o
    escritório representa. Não depende de um separador fixo entre as pessoas
    (vírgula, '|', etc.): cada nome é capturado como o texto entre o fim do
    parêntese anterior (ou o início da string) e o parêntese seguinte, com
    pontuação de separação (vírgula, ponto e vírgula, barra) removida das
    pontas. Usada para agregação por CLIENTE (ver processar_clientes).
    """
    if pd.isna(texto):
        return []
    pares = []
    for m in re.finditer(r"([^()]+)\(([^)]+)\)", str(texto)):
        nome = m.group(1).strip(" ,;/|\t\n")
        papel = m.group(2).strip()
        if nome:
            pares.append((nome, papel))
    return pares


def _normalizar_nome_cliente(nome: str) -> str:
    """
    Chave de agrupamento por cliente: sem acento, minúsculo, espaços
    colapsados — para 'João Silva' e 'joão  silva' (ou variações de
    maiúsculas/acentuação entre exportações) caírem no mesmo cliente em vez
    de virarem duas linhas diferentes na aba Clientes.
    """
    sem_acento = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", sem_acento).strip().lower()


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

def _localizar_linha_cabecalho_tarefas(caminho_arquivo: str, max_linhas: int = 40) -> int:
    """
    Ver _localizar_linha_cabecalho. Procura a linha com "Identificador da
    tarefa" (coluna que sempre existe nesta planilha) em vez de confiar num
    número de linha fixo — antes esta planilha sempre usava "header=2" fixo;
    mantém esse mesmo valor como último recurso caso a coluna não seja
    encontrada.
    """
    linha = _localizar_linha_cabecalho(caminho_arquivo, "Identificador da tarefa", sheet_name="Tarefas", max_linhas=max_linhas)
    return linha if linha is not None else 2


def _carregar_df_tarefas(caminho_arquivo: str) -> pd.DataFrame:
    """
    Lê a planilha de tarefas do disco e prepara as colunas usadas em todo o
    resto do módulo (datas convertidas, categoria e ano). Extraído de
    processar_tarefas() para poder ser reaproveitado por
    processar_tarefas_periodo() (filtro de período arbitrário, usado pelo
    modo Apresentação) sem duplicar a leitura/preparação da planilha.
    """
    linha_cabecalho = _localizar_linha_cabecalho_tarefas(caminho_arquivo)
    df = _ler_excel(caminho_arquivo, sheet_name="Tarefas", header=linha_cabecalho)

    for coluna in ["Data prevista", "Data fatal", "Data da conclusão", "Data de criação"]:
        df[coluna + "_dt"] = pd.to_datetime(df[coluna], format="%d/%m/%Y", errors="coerce")

    df["categoria"] = df["Módulo"].apply(_categoriza_tarefa)
    df["ano"] = df["Data prevista_dt"].dt.year
    return df


def processar_tarefas(caminho_arquivo: str, data_referencia: datetime | None = None) -> dict:
    """
    Lê a planilha de tarefas e devolve:
      anos_disponiveis      -> lista de anos encontrados na planilha (ex: [2026, 2027])
      por_ano               -> {"2026": {...produção...}, "2027": {...}, ...}
      tipos_tarefa_por_ano  -> {"2026": {...}, "2027": {...}, ...}
      indicadores_por_ano   -> {"2026": {...ciclo, polo, área do direito, delegação...}, ...}
      pendencias            -> vencidas/a vencer, sempre o quadro atual (não muda por ano)
      detalhamento          -> lista de todas as tarefas, de todos os anos, com o campo "ano"
      periodo_disponivel    -> {"min", "max"} (AAAA-MM-DD) com a menor e maior "Data prevista"
                                de toda a planilha — usado pelo seletor de período do
                                modo Apresentação para limitar as datas escolhíveis.
    """
    hoje = pd.Timestamp((data_referencia or datetime.now()).date())

    df = _carregar_df_tarefas(caminho_arquivo)

    anos_disponiveis = sorted(int(a) for a in df["ano"].dropna().unique())

    por_ano = {}
    tipos_tarefa_por_ano = {}
    indicadores_por_ano = {}
    for ano in anos_disponiveis:
        df_ano = df[df["ano"] == ano]
        por_ano[str(ano)] = _producao_do_ano(df_ano)
        tipos_tarefa_por_ano[str(ano)] = _tipos_tarefa(df_ano)
        indicadores_por_ano[str(ano)] = _indicadores_do_ano(df_ano)

    datas_previstas_validas = df["Data prevista_dt"].dropna()
    periodo_disponivel = {
        "min": datas_previstas_validas.min().strftime("%Y-%m-%d") if len(datas_previstas_validas) else None,
        "max": datas_previstas_validas.max().strftime("%Y-%m-%d") if len(datas_previstas_validas) else None,
    }

    return {
        "atualizado_em": hoje.strftime("%d/%m/%Y"),
        "anos_disponiveis": anos_disponiveis,
        "por_ano": por_ano,
        "tipos_tarefa_por_ano": tipos_tarefa_por_ano,
        "indicadores_por_ano": indicadores_por_ano,
        "pendencias": _calcular_pendencias(df, hoje),
        "detalhamento": _detalhamento_tarefas(df),
        "periodo_disponivel": periodo_disponivel,
    }


def processar_tarefas_periodo(caminho_arquivo: str, data_inicio: datetime, data_fim: datetime) -> dict:
    """
    Mesma lógica de produção/tipos de tarefa/indicadores de processar_tarefas,
    mas recortada por um período arbitrário (data_inicio até data_fim,
    inclusive) em vez de por ano inteiro — usado pelo filtro de período do
    modo Apresentação, quando o advogado quer mostrar, por exemplo, só de
    08/08 até 14/09. Compara pela "Data prevista", a mesma referência já
    usada para agrupar por ano em processar_tarefas — mantém os dois modos
    (ano inteiro x período customizado) consistentes entre si.
    """
    df = _carregar_df_tarefas(caminho_arquivo)

    inicio = pd.Timestamp(data_inicio)
    fim = pd.Timestamp(data_fim)
    df_periodo = df[
        df["Data prevista_dt"].notna()
        & (df["Data prevista_dt"] >= inicio)
        & (df["Data prevista_dt"] <= fim)
    ]

    return {
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y"),
        "total_no_periodo": int(len(df_periodo)),
        "producao": _producao_do_ano(df_periodo),
        "tipos_tarefa": _tipos_tarefa(df_periodo),
        "indicadores": _indicadores_do_ano(df_periodo),
    }


def _producao_do_ano(df: pd.DataFrame) -> dict:
    """KPIs, evolução mensal, responsáveis e tipos de tarefa — tudo restrito a um único ano."""
    situacao_counts = df["Situação"].value_counts().to_dict()

    # --- Evolução mensal — sempre os 12 meses, mesmo que o ano ainda não tenha
    # terminado (ex: ano corrente) ou que algum mês não tenha nenhuma tarefa.
    df = df.copy()
    df["mes_num"] = df["Data prevista_dt"].dt.month
    df["_is_audiencia"] = df["Tipo de tarefa"].astype(str).str.contains("audi", case=False, na=False)
    df["_is_atendimento"] = df["categoria"] == "Atendimento/Administrativa"
    mensal = df.groupby("mes_num").agg(
        total=("Identificador da tarefa", "count"),
        concluidas=("Situação", lambda s: (s == "Concluída com sucesso").sum()),
        # "Pendente" e "Em execução" são situações diferentes e NÃO podem ser
        # somadas aqui — antes eram, o que fazia o total de "pendentes" do
        # gráfico não bater com o KPI "Pendentes" da aba Produção (que conta
        # só Situação=="Pendente"). "Em execução" tem sua própria contagem.
        pendentes=("Situação", lambda s: (s == "Pendente").sum()),
        em_execucao=("Situação", lambda s: (s == "Em execução").sum()),
        canceladas=("Situação", lambda s: (s == "Cancelado").sum()),
        audiencias=("_is_audiencia", "sum"),
        atendimentos=("_is_atendimento", "sum"),
    )
    # Força tipo numérico antes do reindex: quando o período filtrado não tem
    # NENHUMA tarefa (ex: um intervalo de datas sem dados), o groupby fica
    # vazio e o pandas às vezes infere as colunas somadas via lambda como
    # texto em vez de número — aí o reindex com fill_value=0 quebra
    # ("Invalid value '0' for dtype 'str'"). Forçar int64 aqui evita o erro
    # nesse caso raro, sem mudar nada no caso normal (com dados).
    mensal = mensal.astype({coluna: "int64" for coluna in mensal.columns})
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
            # "Pendente" e "Em execução" contadas separadas (ver comentário em
            # _producao_do_ano) — não podem ser somadas numa "pendentes" só.
            "pendentes": int((subconjunto["Situação"] == "Pendente").sum()),
            "em_execucao": int((subconjunto["Situação"] == "Em execução").sum()),
            "canceladas": int((subconjunto["Situação"] == "Cancelado").sum()),
            "total": int(len(subconjunto)),
        }

    tipo_counts = df["Tipo de tarefa"].value_counts().to_dict()
    # Mesma contagem de tipo_counts, mas separada por categoria (Processual x
    # Atendimento/Administrativa x Outras) — usada no dashboard para dividir a
    # lista de "tipos de tarefa mais frequentes" em abas, em vez de uma lista
    # única gigante quando não há mais limite de Top N.
    tipo_counts_processual = df[df["categoria"] == "Processual"]["Tipo de tarefa"].value_counts().to_dict()
    tipo_counts_atendimento = df[df["categoria"] == "Atendimento/Administrativa"]["Tipo de tarefa"].value_counts().to_dict()
    tipo_counts_outras = df[df["categoria"] == "Outras (sem vínculo)"]["Tipo de tarefa"].value_counts().to_dict()
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
        "monthly": mensal[["mes_nome", "total", "concluidas", "pendentes", "em_execucao", "canceladas", "audiencias", "atendimentos"]].to_dict("records"),
        "resp_counts": {k: int(v) for k, v in resp_counts.items()},
        "resp_situacao": resp_situacao,
        "tipo_counts": {k: int(v) for k, v in tipo_counts.items()},
        "tipo_counts_processual": {k: int(v) for k, v in tipo_counts_processual.items()},
        "tipo_counts_atendimento": {k: int(v) for k, v in tipo_counts_atendimento.items()},
        "tipo_counts_outras": {k: int(v) for k, v in tipo_counts_outras.items()},
        "modulo_counts": {k: int(v) for k, v in modulo_counts.items()},
        "sitproc_counts": {k: int(v) for k, v in sitproc_counts.items()},
        "grupo_counts": {k: int(v) for k, v in grupo_counts.items()},
        "processos_distintos": int(df["Identificador do módulo"].nunique()),
        "clientes_distintos": len(nomes_clientes),
    }


def _linha_pendencia(row, campo_dias: str) -> dict:
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


def _bloco_pendencias(df_situacao: pd.DataFrame, hoje: pd.Timestamp) -> dict:
    """
    Vencidas/a vencer de UMA situação só (Pendente OU Em execução) —
    nunca as duas misturadas na mesma lista (ver comentário em
    _calcular_pendencias, que chama esta função duas vezes, uma pra cada
    situação).
    """
    vencidas = df_situacao[df_situacao["Data fatal_dt"] < hoje].copy()
    vencidas["dias_atraso"] = (hoje - vencidas["Data fatal_dt"]).dt.days
    vencidas = vencidas.sort_values("dias_atraso", ascending=False)

    a_vencer = df_situacao[df_situacao["Data fatal_dt"] >= hoje].copy()
    a_vencer["dias_restantes"] = (a_vencer["Data fatal_dt"] - hoje).dt.days
    a_vencer = a_vencer.sort_values("dias_restantes").head(60)

    return {
        "overdue_list": [_linha_pendencia(r, "dias_atraso") for _, r in vencidas.iterrows()],
        "upcoming_list": [_linha_pendencia(r, "dias_restantes") for _, r in a_vencer.iterrows()],
        "total_atrasadas": int(len(vencidas)),
        "total_pendentes": int(len(df_situacao)),
    }


def _calcular_pendencias(df: pd.DataFrame, hoje: pd.Timestamp) -> dict:
    """
    Pendências vencidas e a vencer — sempre relativas a HOJE, olhando a
    planilha inteira (todos os anos). Uma tarefa vencida de 2026 continua
    aparecendo aqui em 2027 enquanto não for resolvida; não faz sentido
    "escondê-la" só porque o ano mudou.

    "Pendente" e "Em execução" são tratadas SEPARADAMENTE, nunca somadas
    numa lista só — antes eram, e isso fazia o total dessas listas não
    bater com o KPI "Pendentes" da aba Produção (que conta só
    Situação=="Pendente"):
      - as chaves de topo (overdue_list, upcoming_list, total_atrasadas,
        total_pendentes) cobrem só Situação=="Pendente" — mantidas com
        esses nomes por compatibilidade com o resto do dashboard, que já
        lê essas chaves assim.
      - "em_execucao" tem a mesma estrutura, para Situação=="Em execução"
        — uma tarefa já em andamento mas com o prazo fatal vencido (ou
        perto de vencer) não pode ficar invisível só por já estar "em
        execução".
    """
    pendentes_df = df[df["Situação"] == "Pendente"].copy()
    em_execucao_df = df[df["Situação"] == "Em execução"].copy()

    resultado = _bloco_pendencias(pendentes_df, hoje)
    resultado["em_execucao"] = _bloco_pendencias(em_execucao_df, hoje)
    return resultado


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

    # --- Tempo de ciclo por tipo de tarefa (ex: "Petição Inicial") ---
    ciclo_por_tipo_tarefa = {}
    if len(concluidas):
        tipos_ciclo = concluidas["Tipo de tarefa"].dropna().astype(str).str.strip().unique()
        for tipo in tipos_ciclo:
            filtro = concluidas["Tipo de tarefa"].astype(str).str.strip() == tipo
            media = concluidas.loc[filtro, "dias_ciclo"].mean()
            if pd.notna(media):
                ciclo_por_tipo_tarefa[tipo] = round(float(media), 1)

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
        "ciclo_por_tipo_tarefa": ciclo_por_tipo_tarefa,
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

def _localizar_linha_cabecalho_parados(caminho_arquivo: str, max_linhas: int = 40) -> int:
    """
    Ver _localizar_linha_cabecalho. Procura a linha com "Situação processo"
    (coluna que sempre existe neste relatório) em vez de confiar num número
    de linha fixo — antes isso era fixo em "header=18", mas o número de
    linhas de metadados que o Projuris coloca no topo do arquivo ("Filtros
    utilizados:", "Data início:", etc.) pode variar de uma exportação para
    outra; mantém esse mesmo valor como último recurso caso a coluna não seja
    encontrada.
    """
    linha = _localizar_linha_cabecalho(caminho_arquivo, "Situação processo", max_linhas=max_linhas)
    return linha if linha is not None else 18


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

    linha_cabecalho = _localizar_linha_cabecalho_parados(caminho_arquivo)
    df = _ler_excel(caminho_arquivo, header=linha_cabecalho)
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


# ---------------------------------------------------------------------------
# Planilha de PROCESSOS (cadastro geral — um por processo, não por tarefa)
# ---------------------------------------------------------------------------

def _localizar_linha_cabecalho_processos(caminho_arquivo: str, max_linhas: int = 40) -> int:
    """
    Ver _localizar_linha_cabecalho. Procura a linha com "Identificador" (o
    código do processo, ex: "PRO.0000001", que sempre existe nesse relatório)
    em vez de confiar num número de linha fixo — essa planilha também vem com
    linhas de metadados do Projuris no topo antes do cabeçalho de verdade, e
    o número de linhas pode variar de uma exportação para outra; mantém o
    mesmo valor observado na primeira exportação (linha 2) como último
    recurso caso a coluna não seja encontrada.
    """
    linha = _localizar_linha_cabecalho(caminho_arquivo, "Identificador", sheet_name="Processos", max_linhas=max_linhas)
    return linha if linha is not None else 2


def processar_processos(caminho_arquivo: str) -> dict:
    """
    Lê a planilha "Processos" exportada do Projuris (aba "Processos" — a aba
    "Filtros" só documenta os filtros usados na exportação, não tem dado
    nenhum) e devolve o cadastro de cada processo com os campos pedidos pelo
    cliente: assunto, situação, justiça (Federal/Estadual/Trabalhista/etc.),
    instância, área do direito, data de distribuição, valor da causa, data
    do último andamento, cliente(s) do escritório naquele processo, o polo
    em que atuamos (autor x réu) e o estado/cidade do processo.

    Diferente da planilha de Tarefas, aqui cada LINHA já é um processo (não
    uma tarefa) — não precisa nenhuma agregação por "Número do processo".

    O polo é calculado a partir da coluna "Cliente", que já vem só com o(s)
    envolvido(s) que o escritório representa, com o papel entre parênteses
    (ex: "Fulano (Autor)", "Empresa X (Executado)") — mesmo formato usado na
    coluna "Envolvidos cliente" da planilha de Processos Parados, por isso a
    reutilização de _extrai_papeis/_classifica_polo.
    """
    linha_cabecalho = _localizar_linha_cabecalho_processos(caminho_arquivo)
    df = _ler_excel(caminho_arquivo, sheet_name="Processos", header=linha_cabecalho)
    df = df[df["Identificador"].notna()].copy()

    df["papeis"] = df["Cliente"].apply(_extrai_papeis)
    df["polo"] = df["papeis"].apply(_classifica_polo)

    lista = []
    valor_total_causa = 0.0
    qtd_valores_suspeitos = 0
    for _, row in df.iterrows():
        data_distribuicao = row["Data distribuição"]
        valor_causa = float(row["Valor ação"]) if pd.notna(row["Valor ação"]) else None
        valor_suspeito = valor_causa is not None and valor_causa > LIMITE_SANIDADE_VALOR_CAUSA
        if valor_suspeito:
            qtd_valores_suspeitos += 1
        elif valor_causa is not None:
            valor_total_causa += valor_causa

        lista.append({
            "identificador": _texto_seguro(row["Identificador"]),
            "assunto": _texto_seguro(row["Assunto"], tamanho_max=160),
            "situacao": _texto_seguro(row["Situação"]),
            "justica": _texto_seguro(row["Justiça"]),
            "instancia": _texto_seguro(row["Instância"]),
            "area": _texto_seguro(row["Área"]),
            "data_distribuicao": data_distribuicao.strftime("%d/%m/%Y") if pd.notna(data_distribuicao) else None,
            "valor_causa": valor_causa,
            "valor_causa_suspeito": valor_suspeito,
            "data_ultimo_andamento": _texto_seguro(row["Data dos últimos andamentos"]),
            "cliente": _texto_seguro(row["Cliente"], tamanho_max=160),
            "polo": row["polo"],
            "estado": _texto_seguro(row["Estado"]),
            "cidade": _texto_seguro(row["Cidade"]),
        })

    return {
        "total": int(len(df)),
        "situacao_counts": {k: int(v) for k, v in df["Situação"].fillna("Não informado").value_counts().to_dict().items()},
        "justica_counts": {k: int(v) for k, v in df["Justiça"].fillna("Não informado").value_counts().to_dict().items()},
        "instancia_counts": {k: int(v) for k, v in df["Instância"].fillna("Não informado").value_counts().to_dict().items()},
        "area_counts": {k: int(v) for k, v in df["Área"].fillna("Não informado").value_counts().to_dict().items()},
        "estado_counts": {k: int(v) for k, v in df["Estado"].fillna("Não informado").value_counts().to_dict().items()},
        "polo_counts": {k: int(v) for k, v in df["polo"].value_counts().to_dict().items()},
        # Soma já excluindo valores acima de LIMITE_SANIDADE_VALOR_CAUSA (ver
        # comentário na constante) — "qtd_valores_suspeitos" avisa quando algum
        # processo foi deixado de fora dessa soma, pra não sumir em silêncio.
        "valor_total_causa": round(valor_total_causa, 2),
        "qtd_valores_suspeitos": qtd_valores_suspeitos,
        "lista": lista,
    }


# ---------------------------------------------------------------------------
# Visão "Clientes" — não existe uma planilha de cadastro de clientes própria
# no Projuris exportado, então esta função AGREGA, por cliente, o que já está
# espalhado nas 3 planilhas oficiais (Processos, Processos Parados e
# Tarefas), agrupando pelo NOME extraído das colunas de partes de cada uma.
# Sem CPF/CNPJ nem qualquer outro identificador único — a chave de
# agrupamento é o nome normalizado (ver _normalizar_nome_cliente), a pedido
# do cliente do escritório. Também não inclui valor da causa.
# ---------------------------------------------------------------------------

def processar_clientes(caminho_tarefas: str, caminho_processos_parados: str, caminho_processos: str) -> dict:
    """
    Lê as 3 planilhas oficiais e devolve, por cliente:
      - da planilha Processos: cada processo em que ele é parte (assunto,
        situação, justiça, instância, área, data de distribuição, data do
        último andamento, estado, cidade, polo)
      - da planilha Processos Parados: cada processo parado em que ele é
        parte (assunto, órgão, área, situação, data do último movimento, polo)
      - da planilha Tarefas: quantidade de tarefas vinculadas a ele

    Devolve também agregados gerais (total de clientes, quantos têm processo
    parado, distribuição por área do direito e por polo) para os KPIs e
    gráficos do topo da aba.
    """
    linha_cab_processos = _localizar_linha_cabecalho_processos(caminho_processos)
    df_proc = _ler_excel(caminho_processos, sheet_name="Processos", header=linha_cab_processos)
    df_proc = df_proc[df_proc["Identificador"].notna()].copy()

    linha_cab_parados = _localizar_linha_cabecalho_parados(caminho_processos_parados)
    df_par = _ler_excel(caminho_processos_parados, header=linha_cab_parados)
    df_par = df_par[df_par["Situação processo"] != "Usuário emissor:"].copy()

    df_tar = _carregar_df_tarefas(caminho_tarefas)

    clientes: dict[str, dict] = {}

    def _obter_cliente(chave: str, nome_exibicao: str) -> dict:
        if chave not in clientes:
            clientes[chave] = {
                "nome": nome_exibicao,
                "processos": [],
                "processos_parados": [],
                "total_tarefas": 0,
            }
        return clientes[chave]

    # --- Processos ---
    for _, row in df_proc.iterrows():
        for nome, papel in _extrai_nomes_papeis(row["Cliente"]):
            chave = _normalizar_nome_cliente(nome)
            if not chave:
                continue
            cliente = _obter_cliente(chave, nome)
            data_distribuicao = row["Data distribuição"]
            cliente["processos"].append({
                "identificador": _texto_seguro(row["Identificador"]),
                "assunto": _texto_seguro(row["Assunto"], tamanho_max=160),
                "situacao": _texto_seguro(row["Situação"]),
                "justica": _texto_seguro(row["Justiça"]),
                "instancia": _texto_seguro(row["Instância"]),
                "area": _texto_seguro(row["Área"]),
                "data_distribuicao": data_distribuicao.strftime("%d/%m/%Y") if pd.notna(data_distribuicao) else None,
                "data_ultimo_andamento": _texto_seguro(row["Data dos últimos andamentos"]),
                "estado": _texto_seguro(row["Estado"]),
                "cidade": _texto_seguro(row["Cidade"]),
                "polo": _classifica_polo([papel.strip().lower()]),
            })

    # --- Processos Parados ---
    for _, row in df_par.iterrows():
        for nome, papel in _extrai_nomes_papeis(row["Envolvidos cliente"]):
            chave = _normalizar_nome_cliente(nome)
            if not chave:
                continue
            cliente = _obter_cliente(chave, nome)
            cliente["processos_parados"].append({
                "processo": _texto_seguro(row["Numero processo"]),
                "assunto": _texto_seguro(row["Assunto"], tamanho_max=160),
                "orgao": _texto_seguro(row["Orgão"]),
                "area": _texto_seguro(row["Área"]),
                "situacao": _texto_seguro(row["Situação processo"]),
                "data_mov": _texto_seguro(row["Data último movimento"]),
                "polo": _classifica_polo([papel.strip().lower()]),
            })

    # --- Tarefas: só conta (o nome vem marcado com o sufixo "- Cliente") ---
    for coluna in ["Envolvidos do processo (partes ativas)", "Envolvidos do processo (partes passivas)"]:
        for lista_nomes in df_tar[coluna].apply(_extrai_nomes_clientes):
            for nome in lista_nomes:
                chave = _normalizar_nome_cliente(nome)
                if not chave:
                    continue
                cliente = _obter_cliente(chave, nome)
                cliente["total_tarefas"] += 1

    # ---- Consolida cada cliente: totais e polo/área/situação predominantes ----
    lista_clientes = []
    for cliente in clientes.values():
        todos_polos = [p["polo"] for p in cliente["processos"]] + [p["polo"] for p in cliente["processos_parados"]]
        polo_counts_cliente = pd.Series(todos_polos).value_counts().to_dict() if todos_polos else {}
        polo_predominante = max(polo_counts_cliente, key=polo_counts_cliente.get) if polo_counts_cliente else "Não identificado"

        areas = [p["area"] for p in cliente["processos"] if p["area"]] or [p["area"] for p in cliente["processos_parados"] if p["area"]]
        area_counts_cliente = pd.Series(areas).value_counts().to_dict() if areas else {}
        area_principal = max(area_counts_cliente, key=area_counts_cliente.get) if area_counts_cliente else "Não informado"

        situacoes = [p["situacao"] for p in cliente["processos"] if p["situacao"]]
        situacao_counts_cliente = pd.Series(situacoes).value_counts().to_dict() if situacoes else {}
        situacao_predominante = max(situacao_counts_cliente, key=situacao_counts_cliente.get) if situacao_counts_cliente else "Não informado"

        estados = sorted({p["estado"] for p in cliente["processos"] if p["estado"]})

        datas_andamento = [p["data_ultimo_andamento"] for p in cliente["processos"] if p["data_ultimo_andamento"]]
        datas_dt = pd.to_datetime(pd.Series(datas_andamento, dtype="object"), format="%d/%m/%Y", errors="coerce").dropna()
        ultimo_andamento = datas_dt.max().strftime("%d/%m/%Y") if len(datas_dt) else None

        lista_clientes.append({
            "nome": cliente["nome"],
            "total_processos": len(cliente["processos"]),
            "total_processos_parados": len(cliente["processos_parados"]),
            "total_tarefas": cliente["total_tarefas"],
            "polo": polo_predominante,
            "area_principal": area_principal,
            "situacao_predominante": situacao_predominante,
            "estados": estados,
            "ultimo_andamento": ultimo_andamento,
            "processos": cliente["processos"],
            "processos_parados": cliente["processos_parados"],
        })

    lista_clientes.sort(key=lambda c: c["total_processos"], reverse=True)

    todas_areas = [p["area"] for c in lista_clientes for p in c["processos"] if p["area"]]
    area_counts = pd.Series(todas_areas).value_counts().head(10).to_dict() if todas_areas else {}

    todos_polos_geral = [p["polo"] for c in lista_clientes for p in c["processos"]]
    polo_counts = pd.Series(todos_polos_geral).value_counts().to_dict() if todos_polos_geral else {}

    clientes_com_parado = sum(1 for c in lista_clientes if c["total_processos_parados"] > 0)

    return {
        "total_clientes": len(lista_clientes),
        "clientes_com_processo_parado": clientes_com_parado,
        "area_counts": {k: int(v) for k, v in area_counts.items()},
        "polo_counts": {k: int(v) for k, v in polo_counts.items()},
        "lista": lista_clientes,
    }


def processar_tarefas_por_cliente_periodo(caminho_arquivo: str, data_inicio: datetime, data_fim: datetime) -> dict:
    """
    Conta quantas tarefas foram CONCLUÍDAS para cada cliente dentro de um
    período escolhido livremente (um dia, um mês, um ano — qualquer
    intervalo), usando a "Data da conclusão" de cada tarefa — não é o mesmo
    filtro de "Data prevista" usado no resto da Apresentação, porque aqui o
    que importa é o trabalho de fato REALIZADO no período ("esse mês eu fiz
    X tarefas pra tal cliente"), não o que estava agendado. Por isso, tarefa
    sem "Data da conclusão" preenchida (ainda pendente ou em execução) não
    entra na contagem.

    Usa a mesma extração de nome de cliente das tarefas que processar_clientes
    (colunas "Envolvidos do processo (partes ativas/passivas)", onde o nome
    do cliente do escritório vem marcado com o sufixo "- Cliente").
    """
    df = _carregar_df_tarefas(caminho_arquivo)

    inicio = pd.Timestamp(data_inicio)
    fim = pd.Timestamp(data_fim)
    df_periodo = df[
        df["Data da conclusão_dt"].notna()
        & (df["Data da conclusão_dt"] >= inicio)
        & (df["Data da conclusão_dt"] <= fim)
    ]

    # Por linha (tarefa) — não por coluna — pra não contar a mesma tarefa duas
    # vezes caso o nome do cliente apareça tanto em "partes ativas" quanto em
    # "partes passivas" (não deveria acontecer, mas evita duplicidade se
    # acontecer). Guarda também os dados da própria tarefa, pra dar pra
    # mostrar quais tarefas foram essas, não só a contagem.
    contagem: dict[str, dict] = {}
    for _, row in df_periodo.iterrows():
        nomes_da_linha: dict[str, str] = {}
        for coluna in ["Envolvidos do processo (partes ativas)", "Envolvidos do processo (partes passivas)"]:
            for nome in _extrai_nomes_clientes(row[coluna]):
                chave = _normalizar_nome_cliente(nome)
                if chave:
                    nomes_da_linha[chave] = nome

        if not nomes_da_linha:
            continue

        info_tarefa = {
            "id": _texto_seguro(row["Identificador da tarefa"]),
            "tipo": _texto_seguro(row["Tipo de tarefa"]),
            "titulo": _texto_seguro(row["Título"]),
            "responsavel": _texto_seguro(row["Responsáveis da tarefa"]),
            "data_conclusao": _texto_seguro(row["Data da conclusão"]),
            "processo": _texto_seguro(row["Número do processo"]),
            "orgao": _texto_seguro(row["Órgão"]),
        }

        for chave, nome in nomes_da_linha.items():
            if chave not in contagem:
                contagem[chave] = {"nome": nome, "tarefas": 0, "tarefas_lista": []}
            contagem[chave]["tarefas"] += 1
            contagem[chave]["tarefas_lista"].append(info_tarefa)

    lista = sorted(contagem.values(), key=lambda c: c["tarefas"], reverse=True)

    return {
        "inicio": inicio.strftime("%d/%m/%Y"),
        "fim": fim.strftime("%d/%m/%Y"),
        "total_tarefas_concluidas": int(len(df_periodo)),
        "total_clientes_atendidos": len(lista),
        "lista": lista,
    }


# ---------------------------------------------------------------------------
# Tela de "Atualização" — permite que a própria pessoa do escritório troque
# uma das 3 planilhas oficiais (Tarefas, Processos Parados ou Processos) sem
# precisar de mim/TI para enviar o arquivo manualmente no servidor.
#
# A ideia: antes de aceitar o arquivo enviado e SUBSTITUIR a planilha oficial
# com ele, confere se é realmente uma planilha do tipo escolhido — mesmas
# colunas que o processamento (as funções acima) efetivamente usa. Não
# importa a ORDEM das colunas, não importa se tem linhas a mais ou a menos, e
# não importa se tem colunas A MAIS do que o necessário (essas simplesmente
# são ignoradas — o processamento sempre lê pelo NOME da coluna, nunca pela
# posição, então uma coluna nova que o Projuris passe a exportar não quebra
# nada e não precisa de nenhum tratamento especial). Só é considerada
# inválida a planilha em que FALTA alguma coluna que o dashboard precisa.
# ---------------------------------------------------------------------------

CONFIG_PLANILHAS_ATUALIZACAO = {
    "tarefas": {
        "nome_exibicao": "Tarefas",
        "sheet_name": "Tarefas",
        "coluna_marcador": "Identificador da tarefa",
        "colunas_obrigatorias": [
            "Identificador da tarefa", "Situação", "Módulo", "Tipo de tarefa", "Título",
            "Data prevista", "Data fatal", "Data da conclusão", "Data de criação",
            "Responsáveis da tarefa", "Criada por", "Número do processo", "Órgão",
            "Identificador do módulo", "Situação do processo", "Grupos de trabalho",
            "Envolvidos do processo (partes ativas)", "Envolvidos do processo (partes passivas)",
            "Assunto",
        ],
    },
    "processos_parados": {
        "nome_exibicao": "Processos Parados",
        "sheet_name": None,
        "coluna_marcador": "Situação processo",
        "colunas_obrigatorias": [
            "Situação processo", "Data último movimento", "Data distribuição",
            "Envolvidos cliente", "Numero processo", "Assunto", "Orgão", "Área",
            "Usuários responsáveis",
        ],
    },
    "processos": {
        "nome_exibicao": "Processos",
        "sheet_name": "Processos",
        "coluna_marcador": "Identificador",
        "colunas_obrigatorias": [
            "Identificador", "Assunto", "Situação", "Justiça", "Instância", "Área",
            "Data distribuição", "Valor ação", "Data dos últimos andamentos", "Cliente",
            "Estado", "Cidade",
        ],
    },
}


def validar_planilha_atualizacao(tipo: str, caminho_arquivo: str) -> dict:
    """
    Confere se `caminho_arquivo` é uma planilha válida do `tipo` escolhido
    (uma chave de CONFIG_PLANILHAS_ATUALIZACAO), ANTES de ela substituir a
    planilha oficial correspondente.

    Devolve sempre um dict com "valido": True/False.
      - Inválida:  {"valido": False, "motivo": "texto explicando o problema, para mostrar ao usuário"}
      - Válida:    {"valido": True, "total_linhas": N, "colunas_extras": [...]}
        "colunas_extras" são colunas que a planilha enviada tem e a oficial
        não exige — informativo (ex: pra mostrar "detectamos N colunas
        novas, foram ignoradas"), não impede a atualização.
    """
    if tipo not in CONFIG_PLANILHAS_ATUALIZACAO:
        return {"valido": False, "motivo": f"Tipo de planilha desconhecido: '{tipo}'."}

    config = CONFIG_PLANILHAS_ATUALIZACAO[tipo]

    try:
        linha_cabecalho = _localizar_linha_cabecalho(
            caminho_arquivo, config["coluna_marcador"], sheet_name=config["sheet_name"]
        )
    except Exception:
        motivo = "Não foi possível abrir o arquivo como planilha Excel. Confira se é um arquivo .xlsx ou .xls válido"
        if config["sheet_name"]:
            motivo += f' e se tem uma aba chamada "{config["sheet_name"]}".'
        else:
            motivo += "."
        return {"valido": False, "motivo": motivo}

    if linha_cabecalho is None:
        return {
            "valido": False,
            "motivo": (
                f'Não encontramos a coluna "{config["coluna_marcador"]}" nas primeiras linhas do arquivo — '
                f'toda planilha de "{config["nome_exibicao"]}" precisa ter essa coluna. '
                f"Confira se você selecionou o arquivo certo."
            ),
        }

    kwargs = {"header": linha_cabecalho}
    if config["sheet_name"]:
        kwargs["sheet_name"] = config["sheet_name"]
    try:
        df = _ler_excel(caminho_arquivo, **kwargs)
    except Exception as exc:
        return {"valido": False, "motivo": f"Não foi possível ler a planilha: {exc}"}

    colunas_arquivo = {str(c) for c in df.columns}
    colunas_faltando = [c for c in config["colunas_obrigatorias"] if c not in colunas_arquivo]
    if colunas_faltando:
        return {
            "valido": False,
            "motivo": (
                f'Esta planilha não bate com o formato de "{config["nome_exibicao"]}" — faltam as colunas: '
                + ", ".join(colunas_faltando)
                + ". Envie a planilha exportada do Projuris para esse mesmo tipo de dado, sem remover colunas."
            ),
        }

    colunas_extras = sorted(colunas_arquivo - set(config["colunas_obrigatorias"]))
    return {"valido": True, "total_linhas": int(len(df)), "colunas_extras": colunas_extras}
