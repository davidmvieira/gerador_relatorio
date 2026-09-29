#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_relatorio_glpi.py
========================
Gera o relatório mensal da Equipe de Sistemas a partir de um export do GLPI.

REGRA DE NEGÓCIO (v2)
---------------------
A base pode conter chamados de VÁRIOS meses. O relatório sempre olha para um
MÊS DE REFERÊNCIA (default: mês corrente) e considera como "relevante para o
período" todo chamado que:

    1. Foi ABERTO no mês de referência, OU
    2. Foi RESOLVIDO/FECHADO no mês de referência (mesmo se aberto antes), OU
    3. AINDA ESTÁ ABERTO hoje (backlog ativo).

Isso garante que:
  - Projetos longos abertos em meses anteriores apareçam enquanto ativos.
  - Chamados herdados e entregues no mês contem como trabalho entregue.
  - O backlog real (chamados travados) não desapareça do relatório.

O painel passa a mostrar o FLUXO: Backlog inicial → Novos → Entregues → Backlog
final, em vez de um "total" ambíguo.
"""

import os
import re
import sys
import json
import shutil
import logging
import argparse
from datetime import datetime
from calendar import monthrange
from typing import Optional

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# =============================================================================
# LOGGING
# =============================================================================
logger = logging.getLogger("relatorio_glpi")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(_h)

# =============================================================================
# CONFIG
# =============================================================================
CONFIG = {
    "arquivo_entrada": "base_de_dados.xlsx",
    "aba_principal": "_WITH_RECURSIVE_entidades_AS_SE",
    "aba_detalhe": "Página1",

    "entidade_sustentacao": "Sistemas",
    "nome_frente_sustentacao": "Sustentação",
    "nome_frente_projetos": "Projetos",

    "titulo_relatorio": "Equipe de Sistemas — Sustentação & Operações",
    "subtitulo": "Relatório Mensal · Entidade Sistemas (14) e sub-entidades",
    "autor": "David Vieira",

    "max_categorias_detalhadas": 10,

    "sla_alvo_horas": {
        "Crítica": 4.0, "Muito Alta": 8.0, "Alta": 24.0,
        "Média": 48.0, "Baixa": 72.0, "Muito Baixa": 120.0,
    },

    "outlier_iqr_mult_atencao": 1.5,
    "outlier_iqr_mult_critico": 3.0,

    "premissas_equipe": {
        "analistas": 6,
        "horas_por_dia": 8,
        "dias_uteis_mes": 20,
        "dias_ferias_no_mes": 11,
    },

    "pasta_saida": os.path.join("relatorios_glpi", "execucoes"),
    "nome_docx": "Relatorio_Sistemas.docx",
    "pasta_historico": os.path.join("relatorios_glpi", "historico"),
}

COR = {
    "navy": "1F3864", "navy_rgb": (0x1F, 0x38, 0x64),
    "orange": "C55A11", "orange_rgb": (0xC5, 0x5A, 0x11),
    "red": "C00000", "red_rgb": (0xC0, 0x00, 0x00),
    "green": "375623", "green_rgb": (0x37, 0x56, 0x23),
    "green_bg": "E2EFDA", "amber_bg": "FCE4D6", "gray_bg": "F2F2F2",
    "banner_blue_bg": "DCE6F1", "banner_orange_bg": "FBE5D6",
    "lightgray": "#D9D9D9",
}

STATUS_MAP = {1: "Novo", 2: "Processando", 3: "Pendente",
              4: "Planejado", 5: "Solucionado", 6: "Fechado"}
PRIORIDADE_MAP = {1: "Muito Baixa", 2: "Baixa", 3: "Média",
                  4: "Alta", 5: "Muito Alta", 6: "Crítica"}
_MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
             "agosto", "setembro", "outubro", "novembro", "dezembro"]


# =============================================================================
# HELPERS
# =============================================================================
def _resolver_mes_referencia(mes_ref: Optional[str]) -> tuple[int, int]:
    if mes_ref is None:
        hoje = datetime.now()
        return hoje.year, hoje.month
    try:
        dt = datetime.strptime(mes_ref, "%Y-%m")
        return dt.year, dt.month
    except ValueError as exc:
        raise ValueError(f"mes_referencia inválido: {mes_ref!r}. Use 'YYYY-MM'.") from exc


def _validar_schema(df: pd.DataFrame, obrigatorias: set, contexto: str):
    faltando = obrigatorias - set(df.columns)
    if faltando:
        raise ValueError(
            f"[{contexto}] Colunas ausentes: {sorted(faltando)}. "
            f"Encontradas: {sorted(df.columns)}"
        )


def _numero(valor, casas: Optional[int] = None):
    if valor is None:
        return None
    try:
        if pd.isna(valor) or np.isinf(valor):
            return None
    except (TypeError, ValueError):
        return None
    resultado = float(valor)
    return round(resultado, casas) if casas is not None else resultado


def _fmt_num(valor, unidade: str = "", casas: int = 1) -> str:
    if valor is None:
        return "—"
    if isinstance(valor, (float, np.floating)):
        return f"{valor:.{casas}f}{unidade}"
    return f"{valor}{unidade}"


def _fmt_variacao(valor, unidade: str = "", casas: int = 1) -> str:
    if valor is None:
        return "—"
    sinal = "+" if valor > 0 else ""
    if isinstance(valor, (float, np.floating)):
        return f"{sinal}{valor:.{casas}f}{unidade}"
    return f"{sinal}{valor}{unidade}"


# =============================================================================
# 1. CARGA E LIMPEZA
# =============================================================================
_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}")


def _ler_entrada(caminho_entrada: str):
    extensao = os.path.splitext(caminho_entrada)[1].lower()
    if extensao == ".csv":
        try:
            return pd.read_csv(caminho_entrada, sep=None, engine="python",
                               encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(caminho_entrada, sep=None, engine="python",
                               encoding="latin-1")
    if extensao in (".xlsx", ".xlsm"):
        return pd.ExcelFile(caminho_entrada)
    raise ValueError("Formato não suportado. Use .xlsx, .xlsm ou .csv.")


def _reparar_linhas_deslocadas(m: pd.DataFrame) -> pd.DataFrame:
    problema = (pd.to_datetime(m["data_abertura"], errors="coerce").isna()
                & m["data_abertura"].notna())
    if not problema.any():
        return m

    correcoes_ab, correcoes_sol = {}, {}
    ids_problematicos = []
    for idx in m.index[problema]:
        linha = m.loc[idx]
        candidatos = sorted({_DATE_RE.search(v).group(0)
                             for v in linha
                             if isinstance(v, str) and _DATE_RE.search(v)})
        if not candidatos:
            continue
        abertura_real = pd.to_datetime(linha["date"], errors="coerce")
        if pd.isna(abertura_real):
            continue
        cand_dt = [pd.to_datetime(c) for c in candidatos]
        posteriores = [c for c in cand_dt if c > abertura_real]
        solucao_real = min(posteriores) if posteriores else max(cand_dt)
        correcoes_ab[idx] = abertura_real
        status_norm = str(linha.get("status_atual", "")).strip().lower()
        if status_norm in ("solucionado", "fechado") or linha.get("status") in (5, 6):
            correcoes_sol[idx] = solucao_real
        ids_problematicos.append(int(linha["id"]))

    if correcoes_ab:
        m.loc[list(correcoes_ab), "data_abertura"] = pd.Series(correcoes_ab)
    if correcoes_sol:
        m.loc[list(correcoes_sol), "data_solucao"] = pd.Series(correcoes_sol)
    if ids_problematicos:
        logger.warning("Linhas com colunas deslocadas reparadas: %s", ids_problematicos)
    return m


def carregar_dados(caminho_entrada: str) -> pd.DataFrame:
    entrada = _ler_entrada(caminho_entrada)
    eh_csv = isinstance(entrada, pd.DataFrame)
    abas = [None] if eh_csv else entrada.sheet_names

    if len(abas) == 1:
        df = entrada if eh_csv else pd.read_excel(caminho_entrada, sheet_name=abas[0])
        _validar_schema(df,
                        {"id_chamado", "titulo", "data_abertura", "status_atual",
                         "prioridade", "tipo_chamado", "categoria", "entidade"},
                        "aba única")
        df = df[df["id_chamado"].notna()].copy()
        df = df.rename(columns={"id_chamado": "id", "titulo": "name",
                                "data_abertura": "date", "status_atual": "status",
                                "prioridade": "priority"})
        df["status"] = df["status"].map(lambda v: {"novo": 1, "processando": 2,
                                                    "pendente": 3, "planejado": 4,
                                                    "solucionado": 5, "fechado": 6}
                                          .get(str(v).strip().lower(), v))
        df["priority"] = df["priority"].map(lambda v: {"muito baixa": 1, "baixa": 2,
                                                        "média": 3, "media": 3,
                                                        "alta": 4, "muito alta": 5,
                                                        "crítica": 6, "critica": 6}
                                             .get(str(v).strip().lower(), v))
        df["type"] = df["tipo_chamado"].map(lambda v: {"incidente": 1,
                                                        "requisição": 2,
                                                        "requisicao": 2}
                                             .get(str(v).strip().lower(), v))
        if "categoria.1" in df.columns:
            df["categoria"] = df["categoria"].fillna(df["categoria.1"])
        df["categoria"] = df["categoria"].fillna("Sem categoria")
        m = df
    else:
        df1 = pd.read_excel(caminho_entrada, sheet_name=CONFIG["aba_principal"])
        df2 = pd.read_excel(caminho_entrada, sheet_name=CONFIG["aba_detalhe"])
        _validar_schema(df1, {"id", "name", "date", "categoria", "status",
                              "priority", "type", "entidade"},
                        CONFIG["aba_principal"])
        _validar_schema(df2, {"id_chamado", "data_abertura", "data_solucao"},
                        CONFIG["aba_detalhe"])
        df2 = df2[df2["id_chamado"].notna()].copy()
        df2["id_chamado"] = df2["id_chamado"].astype(int)
        m = df1.merge(df2, left_on="id", right_on="id_chamado",
                      how="left", suffixes=("", "_s2"))

    m["date"] = pd.to_datetime(m["date"], errors="coerce")
    if "data_abertura" not in m.columns:
        m["data_abertura"] = m["date"]
    m["data_abertura"] = pd.to_datetime(m["data_abertura"], errors="coerce")
    if "data_solucao" not in m.columns:
        m["data_solucao"] = pd.NaT
    m["data_solucao"] = pd.to_datetime(m["data_solucao"], errors="coerce")

    m = _reparar_linhas_deslocadas(m)
    m["data_abertura"] = pd.to_datetime(m["data_abertura"], errors="coerce")
    m["data_solucao"] = pd.to_datetime(m["data_solucao"], errors="coerce")
    return m


def _consolidar_categorias(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_cat_key"] = df["categoria"].astype(str).str.strip().str.lower()
    contagem = df.groupby(["_cat_key", "categoria"]).size().reset_index(name="n")
    contagem = contagem.sort_values(by=["_cat_key", "n", "categoria"],
                                    ascending=[True, False, True])
    melhor = contagem.drop_duplicates("_cat_key")
    mapa = dict(zip(melhor["_cat_key"], melhor["categoria"]))
    df["categoria_final"] = df["_cat_key"].map(mapa).fillna("Sem categoria")
    return df.drop(columns=["_cat_key"])


def preparar_dataframe(m: pd.DataFrame, agora: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    if agora is None:
        agora = pd.Timestamp(datetime.now())
    m = m.copy()
    m = _consolidar_categorias(m)
    m["status_nome"] = m["status"].map(STATUS_MAP)
    m["prioridade_nome"] = m["priority"].map(PRIORIDADE_MAP)
    m["frente"] = np.where(m["entidade"] == CONFIG["entidade_sustentacao"],
                           CONFIG["nome_frente_sustentacao"],
                           CONFIG["nome_frente_projetos"])

    concluido = m["status"].isin([5, 6])
    m["concluido"] = concluido

    tem_coluna_tempo_util = "tempo_resolucao_horas" in m.columns
    if tem_coluna_tempo_util:
        tempo_util = pd.to_numeric(m["tempo_resolucao_horas"], errors="coerce")
    else:
        tempo_util = pd.Series(np.nan, index=m.index)

    tempo_corrido_concluido = (m["data_solucao"] - m["data_abertura"]).dt.total_seconds() / 3600
    tempo_corrido_aberto = (agora - m["data_abertura"]).dt.total_seconds() / 3600

    m["tempo_h"] = np.where(concluido,
                            tempo_util.fillna(tempo_corrido_concluido), np.nan)
    m["horas_acumuladas"] = np.where(~concluido, tempo_corrido_aberto, np.nan)
    m["idade_dias"] = ((agora - m["data_abertura"]).dt.total_seconds() / 86400).round(0)
    m.attrs["tempo_util_disponivel"] = bool(tem_coluna_tempo_util and tempo_util.notna().any())
    return m


# ---------------------------------------------------------------------------
# FILTRO DE PERÍODO — CORAÇÃO DA MUDANÇA
# ---------------------------------------------------------------------------
def filtrar_periodo_referencia(m: pd.DataFrame, ano: int, mes: int) -> pd.DataFrame:
    """
    Retorna o subconjunto de chamados RELEVANTES para o mês de referência:
      1. Abertos no mês (novos do período)
      2. Resolvidos/fechados no mês (entregues, mesmo se herdados)
      3. Ainda abertos hoje (backlog ativo — projetos longos, chamados travados)

    Adiciona flags:
      novo_no_mes       : aberto dentro do mês de referência
      entregue_no_mes   : resolvido/fechado dentro do mês de referência
      herdado_aberto    : aberto antes do mês, ainda não concluído
      concluido_no_mes  : alias de entregue_no_mes (usado em métricas)
    """
    inicio = pd.Timestamp(year=ano, month=mes, day=1)
    fim = inicio + pd.offsets.MonthEnd(1) + pd.Timedelta(days=1)

    abertos_no_mes = (m["date"] >= inicio) & (m["date"] < fim)
    resolvidos_no_mes = (m["data_solucao"] >= inicio) & (m["data_solucao"] < fim)
    ainda_abertos = ~m["concluido"]
    relevante = abertos_no_mes | resolvidos_no_mes | ainda_abertos

    sub = m[relevante].copy()
    sub["novo_no_mes"] = (sub["date"] >= inicio) & (sub["date"] < fim)
    sub["entregue_no_mes"] = (sub["data_solucao"] >= inicio) & (sub["data_solucao"] < fim)
    sub["herdado_aberto"] = (~sub["novo_no_mes"]) & (~sub["concluido"])

    sub.attrs["periodo_inicio"] = inicio
    sub.attrs["periodo_fim"] = fim
    logger.info(
        "Filtro do período %04d-%02d: %d relevantes "
        "(%d novos, %d entregues no mês, %d herdados ainda abertos)",
        ano, mes, len(sub),
        int(sub["novo_no_mes"].sum()),
        int(sub["entregue_no_mes"].sum()),
        int(sub["herdado_aberto"].sum()),
    )
    return sub


# =============================================================================
# 2. MÉTRICAS
# =============================================================================
def periodo_texto(ano: int, mes: int) -> str:
    ultimo = monthrange(ano, mes)[1]
    return f"01 a {ultimo:02d} de {_MESES_PT[mes - 1].capitalize()} de {ano}"


def metricas_painel(m: pd.DataFrame, frente: Optional[str] = None) -> dict:
    """
    Painel baseado em FLUXO:
        backlog_inicial = backlog_final - novos + entregues   (fluxo de estoque)
        carga_total     = backlog_inicial + novos
        taxa_resolucao  = entregues / carga_total
        tmr_h           = média do tempo dos ENTREGUES no mês
    """
    sub = m if frente is None else m[m["frente"] == frente]

    novos = int(sub["novo_no_mes"].sum())
    entregues_df = sub[sub["entregue_no_mes"]]
    entregues = int(len(entregues_df))
    backlog_final = int((~sub["concluido"]).sum())
    backlog_inicial = max(backlog_final - novos + entregues, 0)
    carga_total = backlog_inicial + novos

    taxa = round(entregues / carga_total * 100, 1) if carga_total else 0.0
    tmr = round(entregues_df["tempo_h"].mean(), 1) if entregues else None

    tipo_counts = sub["type"].value_counts()
    req = int(tipo_counts.get(2, 0))
    inc = int(tipo_counts.get(1, 0))
    outros = len(sub) - req - inc

    def pct(x, base):
        return round(x / base * 100, 1) if base else 0.0

    return {
        "novos": novos,
        "entregues": entregues,
        "backlog_inicial": backlog_inicial,
        "backlog_final": backlog_final,
        "carga_total": carga_total,
        "taxa_resolucao": taxa,
        "tmr_h": tmr,
        "req": req, "inc": inc, "outros": max(outros, 0),
        "pct_req": pct(req, len(sub)),
        "pct_inc": pct(inc, len(sub)),
        "pct_outros": pct(max(outros, 0), len(sub)),
        "total_relevante": len(sub),
    }


def _status_breakdown(m: pd.DataFrame, frente: Optional[str] = None) -> dict:
    sub = m if frente is None else m[m["frente"] == frente]
    return {nome: int((sub["status"] == cod).sum())
            for cod, nome in STATUS_MAP.items()}


def metricas_categoria(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    sub = m[m["frente"] == frente]
    qtd = sub.groupby("categoria_final").size().rename("qtd")
    entregues = sub[sub["entregue_no_mes"]]
    tempo = entregues.groupby("categoria_final")["tempo_h"].sum().round(1).rename("tempo_h")
    full = pd.concat([qtd, tempo], axis=1).fillna(0)
    full["qtd"] = full["qtd"].astype(int)
    total_qtd = full["qtd"].sum()
    total_tempo = full["tempo_h"].sum()
    full["pct_qtd"] = (full["qtd"] / total_qtd * 100).round(1) if total_qtd else 0
    full["pct_tempo"] = (full["tempo_h"] / total_tempo * 100).round(1) if total_tempo else 0
    return full.sort_values("qtd", ascending=False)


def compactar_categorias(cat_df: pd.DataFrame, top_n: int) -> pd.DataFrame:
    if len(cat_df) <= top_n:
        return cat_df
    top = cat_df.iloc[:top_n].copy()
    resto = cat_df.iloc[top_n:]
    outros = pd.DataFrame({
        "qtd": [int(resto["qtd"].sum())],
        "tempo_h": [round(resto["tempo_h"].sum(), 1)],
        "pct_qtd": [round(resto["pct_qtd"].sum(), 1)],
        "pct_tempo": [round(resto["pct_tempo"].sum(), 1)],
    }, index=[f"Outros ({len(resto)} categorias)"])
    return pd.concat([top, outros])


def metricas_sla_prioridade(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    """SLA sobre os ENTREGUES no mês (independente de quando foram abertos)."""
    sub = m[(m["frente"] == frente) & m["entregue_no_mes"]].copy()
    if sub.empty:
        return pd.DataFrame(columns=["count", "mean", "min", "max",
                                     "meta_h", "pct_dentro_sla"])
    alvo = CONFIG["sla_alvo_horas"]
    sub["meta_h"] = sub["prioridade_nome"].map(alvo)
    sub["dentro_sla"] = sub["tempo_h"] <= sub["meta_h"]

    ordem = ["Crítica", "Muito Alta", "Alta", "Média", "Baixa", "Muito Baixa"]
    g = sub.groupby("prioridade_nome").agg(
        count=("tempo_h", "count"),
        mean=("tempo_h", "mean"),
        min=("tempo_h", "min"),
        max=("tempo_h", "max"),
        meta_h=("meta_h", "first"),
        pct_dentro_sla=("dentro_sla", lambda s: round(s.mean() * 100, 1)),
    ).round(1)
    return g.reindex([o for o in ordem if o in g.index])


def detectar_outliers(m: pd.DataFrame, frente: str) -> dict:
    vazio = pd.DataFrame(columns=m.columns)
    sub = m[(m["frente"] == frente) & m["entregue_no_mes"]].copy()
    if len(sub) < 4:
        return {"atencao": vazio, "critico": vazio,
                "limite_atencao": None, "limite_critico": None}
    q1, q3 = sub["tempo_h"].quantile([0.25, 0.75])
    iqr = q3 - q1
    lim_at = q3 + CONFIG["outlier_iqr_mult_atencao"] * iqr
    lim_cr = q3 + CONFIG["outlier_iqr_mult_critico"] * iqr
    atencao = sub[sub["tempo_h"] > lim_at].sort_values("tempo_h", ascending=False)
    critico = sub[sub["tempo_h"] > lim_cr].sort_values("tempo_h", ascending=False)
    return {"atencao": atencao, "critico": critico,
            "limite_atencao": round(float(lim_at), 1),
            "limite_critico": round(float(lim_cr), 1)}


def metricas_capacidade(m: pd.DataFrame, frente: str) -> dict:
    p = CONFIG["premissas_equipe"]
    hmm = p["analistas"] * p["dias_uteis_mes"] * p["horas_por_dia"]
    he = hmm - p["dias_ferias_no_mes"] * p["horas_por_dia"]

    sub = m[(m["frente"] == frente) & m["entregue_no_mes"]]
    hpc = round(float(sub["tempo_h"].sum()), 1)
    hha = round(hpc / p["horas_por_dia"], 1)
    ad_pct = round(hha / he * 100, 1) if he else 0
    return {"hmm": hmm, "he": he, "hpc": hpc, "hha": hha, "ad_pct": ad_pct}


def limpar_titulo(titulo: str, max_len: int = 70) -> str:
    if not isinstance(titulo, str):
        return ""
    t = titulo.strip()
    padrao = r"\s*-\s*(?:[A-ZÀ-Ú][a-zà-ú]+\s+){1,}[A-ZÀ-Ú][a-zà-ú]+\s*-\s*\d+\s*-?\s*$"
    t = re.sub(padrao, "", t).strip(" -\t")
    if len(t) > max_len:
        t = t[:max_len - 1].rstrip() + "…"
    return t


def tabela_frente_secundaria(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    """Projetos: TODOS os relevantes (novos, entregues no mês ou ainda abertos)."""
    sub = m[m["frente"] == frente].copy()
    sub["titulo_limpo"] = sub["name"].apply(limpar_titulo)
    sub["origem"] = np.where(sub["novo_no_mes"], "Novo",
                             np.where(sub["entregue_no_mes"], "Entregue no mês",
                                      "Em andamento"))
    # Ordena: em andamento mais antigos primeiro, depois entregues, depois novos
    sub["_ordem"] = np.where(sub["herdado_aberto"], 0,
                             np.where(sub["entregue_no_mes"], 1, 2))
    sub = sub.sort_values(["_ordem", "date"])
    cols = ["id", "titulo_limpo", "origem", "status_nome", "date",
            "data_solucao", "tempo_h", "horas_acumuladas", "idade_dias"]
    return sub[cols]


# =============================================================================
# 3. HISTÓRICO
# =============================================================================
def _metricas_historicas(periodo: str, periodo_txt: str,
                         painel_total, painel_sust, painel_proj,
                         cap_sust, cap_proj, status_total,
                         outliers, cat_sust_full, cat_proj_full) -> dict:
    def painel_resumido(p):
        return {
            "novos": p["novos"], "entregues": p["entregues"],
            "backlog_inicial": p["backlog_inicial"],
            "backlog_final": p["backlog_final"],
            "carga_total": p["carga_total"],
            "taxa_resolucao": _numero(p["taxa_resolucao"], 1),
            "tmr_h": _numero(p["tmr_h"], 1),
            "solicitacoes": p["req"], "incidentes": p["inc"],
            "outros": p["outros"],
        }

    def capacidade_resumida(c):
        return {k: _numero(v, 1) for k, v in c.items()}

    def categorias_resumidas(cat_df):
        return [{
            "categoria": str(cat), "qtd": int(linha["qtd"]),
            "tempo_h": _numero(linha["tempo_h"], 1),
            "pct_qtd": _numero(linha["pct_qtd"], 1),
            "pct_tempo": _numero(linha["pct_tempo"], 1),
        } for cat, linha in cat_df.head(CONFIG["max_categorias_detalhadas"]).iterrows()]

    return {
        "periodo": periodo, "periodo_texto": periodo_txt,
        "total": painel_resumido(painel_total),
        "sustentacao": painel_resumido(painel_sust),
        "projetos": painel_resumido(painel_proj),
        "status_total": status_total,
        "capacidade": {"sustentacao": capacidade_resumida(cap_sust),
                       "projetos": capacidade_resumida(cap_proj)},
        "aderencia_pct": _numero(cap_sust["ad_pct"], 1),
        "outliers": {
            "qtd_atencao": int(len(outliers["atencao"])),
            "qtd_critico": int(len(outliers["critico"])),
            "maior_tempo_h": _numero(outliers["critico"]["tempo_h"].max(), 1)
            if len(outliers["critico"]) else None,
        },
        "categorias": {"sustentacao": categorias_resumidas(cat_sust_full),
                       "projetos": categorias_resumidas(cat_proj_full)},
    }


def carregar_historico() -> list:
    historico = []
    pasta = CONFIG["pasta_historico"]
    if not os.path.isdir(pasta):
        return historico
    padrao = re.compile(r"\d{4}-\d{2}\.json$")
    for nome in sorted(os.listdir(pasta)):
        if not padrao.fullmatch(nome):
            continue
        try:
            with open(os.path.join(pasta, nome), "r", encoding="utf-8") as f:
                reg = json.load(f)
            reg["periodo"] = nome[:-5]
            historico.append(reg)
        except (OSError, json.JSONDecodeError, TypeError) as erro:
            logger.warning("Histórico ignorado (%s): %s", nome, erro)
    return sorted(historico, key=lambda x: x.get("periodo", ""))


def salvar_historico(registro: dict):
    pasta = CONFIG["pasta_historico"]
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, f'{registro["periodo"]}.json')
    if os.path.exists(destino):
        shutil.copy2(destino, destino + ".bak")
        logger.info("Backup do histórico: %s.bak", destino)
    with open(destino, "w", encoding="utf-8") as f:
        json.dump(registro, f, ensure_ascii=False, indent=2)


# =============================================================================
# 4. COMPARAÇÕES E PROJEÇÕES
# =============================================================================
def _safe_get(reg, *chaves, default=None):
    """Acesso seguro a chaves aninhadas, compatível com JSONs antigos."""
    atual = reg
    for k in chaves:
        if not isinstance(atual, dict) or k not in atual:
            return default
        atual = atual[k]
    return atual if atual is not None else default


def comparar_metricas(atual: dict, anterior: dict) -> list:
    indicadores = [
        ("Novos", ("total", "novos")),
        ("Entregues", ("total", "entregues")),
        ("Carga total", ("total", "carga_total")),
        ("Backlog final", ("total", "backlog_final")),
        ("Taxa de resolução", ("total", "taxa_resolucao")),
        ("TMR", ("total", "tmr_h")),
        ("Sustentação (novos)", ("sustentacao", "novos")),
        ("Projetos (novos)", ("projetos", "novos")),
        ("Aderência", (None, "aderencia_pct")),
    ]
    comparacoes = []
    for nome, caminho in indicadores:
        if caminho[0] is None:
            va = _safe_get(atual, caminho[1])
            vb = _safe_get(anterior, caminho[1])
        else:
            va = _safe_get(atual, *caminho)
            vb = _safe_get(anterior, *caminho)
        if va is None or vb is None:
            continue
        dif = va - vb
        var = dif / vb * 100 if vb else None
        unidade = "%" if nome in ("Taxa de resolução", "Aderência") else \
                  "h" if nome == "TMR" else ""
        comparacoes.append({"indicador": nome, "atual": va, "anterior": vb,
                            "diferenca": dif, "variacao_pct": var, "unidade": unidade})
    return comparacoes


def preparar_visao_historica(historico: list, atual: dict) -> dict:
    serie = sorted(historico + [atual], key=lambda x: x["periodo"])
    anterior = next((item for item in reversed(historico)
                     if item["periodo"] < atual["periodo"]), None)

    projecao = {}
    if len(serie) >= 3:
        ultimos = serie[-3:]
        campos = [
            ("Novos", ("total", "novos")),
            ("Entregues", ("total", "entregues")),
            ("Carga total", ("total", "carga_total")),
            ("Taxa de resolução", ("total", "taxa_resolucao")),
            ("TMR", ("total", "tmr_h")),
            ("Sustentação (novos)", ("sustentacao", "novos")),
            ("Projetos (novos)", ("projetos", "novos")),
            ("Aderência", (None, "aderencia_pct")),
        ]
        for nome, (grupo, campo) in campos:
            valores = []
            for item in ultimos:
                v = _safe_get(item, grupo, campo) if grupo else _safe_get(item, campo)
                if v is not None:
                    valores.append(v)
            if valores:
                projecao[nome] = round(sum(valores) / len(valores), 1)
    return {"anterior": anterior,
            "comparacoes": comparar_metricas(atual, anterior) if anterior else [],
            "serie": serie, "projecao": projecao}


# =============================================================================
# 5. GRÁFICOS
# =============================================================================
def _nova_pasta_execucao() -> str:
    base = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(CONFIG["pasta_saida"], base)
    contador = 1
    while os.path.exists(caminho):
        caminho = os.path.join(CONFIG["pasta_saida"], f"{base}_{contador:02d}")
        contador += 1
    os.makedirs(caminho, exist_ok=True)
    return caminho


def _aplicar_estilo_eixo(ax):
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.3)


def grafico_donut_natureza(painel: dict, caminho: str):
    sizes, labels, colors = [], [], []
    if painel["req"] > 0:
        sizes.append(painel["req"]); labels.append("Requisições"); colors.append("#548235")
    if painel["inc"] > 0:
        sizes.append(painel["inc"]); labels.append("Incidentes"); colors.append("#" + COR["red"])
    if painel["outros"] > 0:
        sizes.append(painel["outros"]); labels.append("Outros"); colors.append("#" + COR["navy"])
    if not sizes:
        return
    total = sum(sizes)
    fig, ax = plt.subplots(figsize=(6, 6), dpi=150)
    _, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors, startangle=90, pctdistance=0.78,
        autopct=lambda pct: f"{pct:.1f}%\n({int(round(pct * total / 100))})",
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=3),
        textprops={"fontsize": 13, "fontweight": "bold"})
    for t in texts:
        t.set_fontsize(14); t.set_fontweight("bold"); t.set_color("#333333")
    for t in autotexts:
        t.set_color("white")
    ax.set_title("Natureza da Demanda", fontsize=15, fontweight="bold",
                 color="#" + COR["navy"], pad=20)
    ax.text(0, 0, f"{total}\nrelevantes", ha="center", va="center",
            fontsize=13, fontweight="bold", color="#555555")
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_barras_categoria(cat_df, titulo, caminho, destaque_top_n=2):
    if cat_df.empty:
        return
    cats = cat_df.index.tolist()[::-1]
    vals = cat_df["qtd"].tolist()[::-1]
    total = sum(vals) or 1
    destaque = set(cat_df.index[:destaque_top_n])
    colors = ["#" + COR["navy"] if c in destaque else COR["lightgray"] for c in cats]
    fig, ax = plt.subplots(figsize=(9, max(3.5, 0.55 * len(cats))), dpi=150)
    bars = ax.barh(cats, vals, color=colors, edgecolor="white", height=0.65)
    xmax = max(vals) if vals else 1
    for bar, v in zip(bars, vals):
        ax.text(bar.get_width() + xmax * 0.015,
                bar.get_y() + bar.get_height() / 2,
                f"{v}  ({v/total*100:.1f}%)", va="center", fontsize=10.5,
                fontweight="bold", color="#333333")
    ax.set_xlim(0, xmax * 1.2)
    ax.set_title(titulo, fontsize=15, fontweight="bold", color="#" + COR["navy"], pad=15)
    ax.set_xlabel("Quantidade", fontsize=10, color="#555555")
    _aplicar_estilo_eixo(ax)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_fluxo_historico(serie: list, caminho: str):
    """Backlog final x Novos x Entregues — visão de fluxo por mês."""
    periodos = [it.get("periodo", "?") for it in serie]
    x = list(range(len(periodos)))
    novos = np.array([_safe_get(it, "total", "novos", default=np.nan) for it in serie], dtype=float)
    entregues = np.array([_safe_get(it, "total", "entregues", default=np.nan) for it in serie], dtype=float)
    backlog = np.array([_safe_get(it, "total", "backlog_final", default=np.nan) for it in serie], dtype=float)

    if np.all(np.isnan(novos)) and np.all(np.isnan(entregues)):
        return
    fig, ax = plt.subplots(figsize=(9.5, 5), dpi=150)
    ax.plot(x, novos, marker="o", linewidth=2.4, color="#" + COR["navy"], label="Novos")
    ax.plot(x, entregues, marker="s", linewidth=2.4, color="#" + COR["green"], label="Entregues")
    ax.plot(x, backlog, marker="^", linewidth=2.4, color="#" + COR["red"], label="Backlog final")
    for xi, yi in zip(x, novos):
        if not np.isnan(yi):
            ax.annotate(f"{yi:.0f}", (xi, yi), textcoords="offset points",
                        xytext=(0, 9), ha="center", fontsize=8.5,
                        fontweight="bold", color="#" + COR["navy"])
    for xi, yi in zip(x, entregues):
        if not np.isnan(yi):
            ax.annotate(f"{yi:.0f}", (xi, yi), textcoords="offset points",
                        xytext=(0, -14), ha="center", fontsize=8.5,
                        fontweight="bold", color="#" + COR["green"])
    ax.set_xticks(x); ax.set_xticklabels(periodos, fontsize=10)
    ax.set_ylabel("Chamados", fontsize=10, color="#555555")
    ax.set_title("Fluxo Mensal — Novos, Entregues e Backlog final",
                 fontsize=14, fontweight="bold", color="#" + COR["navy"], pad=15)
    ax.legend(loc="best", fontsize=9, frameon=False)
    _aplicar_estilo_eixo(ax)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_evolucao_historica(serie, grupo, campo, titulo, caminho,
                               unidade="", cor=None, cor_nome="navy"):
    cor = cor or ("#" + COR[cor_nome])
    pontos = [(_safe_get(it, grupo, campo) if grupo else _safe_get(it, campo))
              for it in serie]
    periodos = [it.get("periodo", "?") for it in serie]
    if all(v is None for v in pontos):
        return
    x = list(range(len(periodos)))
    val = np.array([np.nan if v is None else v for v in pontos], dtype=float)
    fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=150)
    ax.plot(x, val, marker="o", linewidth=2.4, color=cor)
    for xi, yi in zip(x, val):
        if not np.isnan(yi):
            ax.annotate(f"{yi:.1f}{unidade}" if isinstance(yi, float) else f"{yi}{unidade}",
                        (xi, yi), textcoords="offset points", xytext=(0, 9),
                        ha="center", fontsize=8.5, fontweight="bold", color=cor)
    ax.set_xticks(x); ax.set_xticklabels(periodos, fontsize=10)
    ax.set_ylabel(unidade.strip() or "Valor", fontsize=10, color="#555555")
    ax.set_title(titulo, fontsize=13, fontweight="bold", color="#" + COR["navy"], pad=12)
    _aplicar_estilo_eixo(ax)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_backlog_frente(serie: list, caminho: str):
    """Sustentação x Projetos — backlog final ao longo do tempo."""
    periodos = [it.get("periodo", "?") for it in serie]
    x = list(range(len(periodos)))
    s = np.array([_safe_get(it, "sustentacao", "backlog_final", default=np.nan) for it in serie], dtype=float)
    p = np.array([_safe_get(it, "projetos", "backlog_final", default=np.nan) for it in serie], dtype=float)
    if np.all(np.isnan(s)) and np.all(np.isnan(p)):
        return
    largura = 0.38
    fig, ax = plt.subplots(figsize=(9.5, 4.5), dpi=150)
    ax.bar([xi - largura/2 for xi in x], np.nan_to_num(s), largura,
           color="#" + COR["navy"], label="Sustentação")
    ax.bar([xi + largura/2 for xi in x], np.nan_to_num(p), largura,
           color="#" + COR["orange"], label="Projetos")
    ax.set_xticks(x); ax.set_xticklabels(periodos, fontsize=10)
    ax.set_ylabel("Backlog final (chamados)", fontsize=10, color="#555555")
    ax.set_title("Backlog Final por Frente", fontsize=14, fontweight="bold",
                 color="#" + COR["navy"], pad=15)
    ax.legend(loc="best", fontsize=9, frameon=False)
    _aplicar_estilo_eixo(ax)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


# =============================================================================
# 6. DOCX
# =============================================================================
def _set_cell_background(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_text(cell, text, *, bold=False, color=None, size=10, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    run = p.add_run(str(text))
    run.font.size = Pt(size); run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def adicionar_tabela(doc, headers, rows, widths_cm, header_bg=None):
    header_bg = header_bg or COR["navy"]
    tabela = doc.add_table(rows=1, cols=len(headers))
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.autofit = False
    hdr = tabela.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].width = Cm(widths_cm[i])
        _set_cell_text(hdr[i], h, bold=True, color=(255, 255, 255), size=10)
        _set_cell_background(hdr[i], header_bg)
    trPr = tabela.rows[0]._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)
    for ri, row in enumerate(rows):
        cells = tabela.add_row().cells
        alt = (ri % 2 == 1)
        for i, val in enumerate(row):
            cells[i].width = Cm(widths_cm[i])
            _set_cell_text(cells[i], val, size=9.5)
            if alt:
                _set_cell_background(cells[i], COR["gray_bg"])
    return tabela


def adicionar_titulo(doc, texto, nivel=1):
    h = doc.add_heading(level=nivel)
    run = h.add_run(texto)
    run.font.color.rgb = RGBColor(*COR["navy_rgb"])
    run.font.size = Pt(16 if nivel == 1 else 13)
    return h


def adicionar_nota(doc, texto):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.italic = True; run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    return p


def adicionar_paragrafo(doc, texto, *, italic=False, bold=False, size=10.5):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.italic = italic; run.bold = bold; run.font.size = Pt(size)
    return p


def adicionar_bullet(doc, texto):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(texto); run.font.size = Pt(10.5)
    return p


def adicionar_banner(doc, texto, cor_fundo, cor_texto):
    tabela = doc.add_table(rows=1, cols=1)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tabela.rows[0].cells[0]
    _set_cell_background(cell, cor_fundo)
    _set_cell_text(cell, texto, bold=True, color=cor_texto, size=11, align="left")
    doc.add_paragraph()
    return tabela


def adicionar_imagem(doc, caminho, largura_cm):
    if not os.path.exists(caminho):
        adicionar_nota(doc, f"[gráfico indisponível: {os.path.basename(caminho)}]")
        return
    doc.add_picture(caminho, width=Cm(largura_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _descricao_periodo_anterior(atual: str, anterior: str) -> str:
    dt_a = datetime.strptime(atual, "%Y-%m")
    dt_b = datetime.strptime(anterior, "%Y-%m")
    distancia = (dt_a.year - dt_b.year) * 12 + dt_a.month - dt_b.month
    if distancia == 1:
        return f"comparação com {anterior}"
    return f"comparação com {anterior} (lacuna de {distancia - 1} mês(es))"


def montar_documento(dados: dict, caminho_saida: str):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2)
    section.top_margin = section.bottom_margin = Cm(1.8)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"; style.font.size = Pt(10.5)

    # Cabeçalho
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(CONFIG["titulo_relatorio"])
    r.bold = True; r.font.size = Pt(20); r.font.color.rgb = RGBColor(*COR["navy_rgb"])
    p = doc.add_paragraph(CONFIG["subtitulo"]); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = Pt(11)
    p = doc.add_paragraph(f'Período de referência: {dados["periodo_texto"]}')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.runs[0].italic = True
    p = doc.add_paragraph(
        f'Redigido por {CONFIG["autor"]} · Gerado em '
        f'{datetime.now().strftime("%d/%m/%Y %H:%M")}')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = Pt(9); p.runs[0].font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    doc.add_paragraph()

    F_SUST = CONFIG["nome_frente_sustentacao"]
    F_PROJ = CONFIG["nome_frente_projetos"]
    visao = dados["visao_historica"]
    atual = dados["historico_atual"]

    # ===================== Escopo do Período =====================
    adicionar_titulo(doc, "Escopo do Período")
    adicionar_paragrafo(
        doc,
        f'Este relatório considera todo chamado RELEVANTE para {dados["periodo_texto"]}:'
    )
    for b in [
        "Chamados ABERTOS no período;",
        "Chamados RESOLVIDOS/FECHADOS no período, mesmo que abertos em meses anteriores;",
        "Chamados ainda ABERTOS hoje (backlog ativo, incluindo projetos em andamento).",
    ]:
        adicionar_bullet(doc, b)
    adicionar_nota(
        doc,
        "O painel abaixo é apresentado em FLUXO (Backlog inicial → Novos → "
        "Entregues → Backlog final), que reflete o que a equipe efetivamente "
        "trabalhou no mês — não apenas o que foi aberto nele.")

    # ===================== Visão Histórica =====================
    adicionar_titulo(doc, "Visão Histórica")
    adicionar_paragrafo(
        doc,
        f'Período atual: {atual["periodo"]} — '
        f'{atual["total"]["novos"]} novos, '
        f'{atual["total"]["entregues"]} entregues, '
        f'backlog final de {atual["total"]["backlog_final"]}.')
    if visao["anterior"]:
        adicionar_paragrafo(
            doc,
            _descricao_periodo_anterior(atual["periodo"], visao["anterior"]["periodo"])
            + ". Os valores não são classificados como favoráveis ou desfavoráveis.",
            italic=True)
        rows = [[c["indicador"],
                 _fmt_num(c["atual"], c["unidade"]),
                 _fmt_num(c["anterior"], c["unidade"]),
                 _fmt_variacao(c["diferenca"], c["unidade"]),
                 _fmt_variacao(c["variacao_pct"], "%")]
                for c in visao["comparacoes"]]
        adicionar_tabela(doc, ["Indicador", "Atual", "Anterior", "Diferença", "Variação"],
                         rows, [4.8, 2.6, 2.6, 2.9, 2.9])
    else:
        adicionar_nota(doc, "Sem período anterior para comparação.")

    if visao["projecao"]:
        adicionar_titulo(doc, "Projeção do Próximo Mês", nivel=2)
        adicionar_paragrafo(
            doc, "ESTIMATIVA via média móvel simples dos 3 últimos meses. "
                 "Não é uma previsão garantida.", italic=True)
        rows = [[nome,
                 _fmt_num(valor,
                          "%" if nome in ("Taxa de resolução", "Aderência")
                          else "h" if nome == "TMR" else "")]
                for nome, valor in visao["projecao"].items()]
        adicionar_tabela(doc, ["Indicador", "Estimativa"], rows, [8, 5])

    # ===================== 1. Painel Geral (fluxo) =====================
    adicionar_titulo(doc, "1. Painel Geral — Fluxo do Período")
    adicionar_paragrafo(
        doc,
        "Leitura: **Backlog inicial** é o que estava em aberto no dia 01; "
        "**Novos** são os abertos no mês; **Carga total** é o que a equipe "
        "precisava endereçar; **Entregues** são os resolvidos/fechados no mês; "
        "**Backlog final** é o que permanece em aberto.")
    pg, pp, pt = dados["painel_sust"], dados["painel_proj"], dados["painel_total"]
    adicionar_tabela(
        doc, ["Indicador", F_SUST, F_PROJ, "Total"],
        [
            ["Backlog inicial (est.)", pg["backlog_inicial"], pp["backlog_inicial"], pt["backlog_inicial"]],
            ["Novos no período", pg["novos"], pp["novos"], pt["novos"]],
            ["Carga total", pg["carga_total"], pp["carga_total"], pt["carga_total"]],
            ["Entregues no período", pg["entregues"], pp["entregues"], pt["entregues"]],
            ["Backlog final", pg["backlog_final"], pp["backlog_final"], pt["backlog_final"]],
            ["Taxa de resolução", f'{pg["taxa_resolucao"]}%',
             f'{pp["taxa_resolucao"]}%', f'{pt["taxa_resolucao"]}%'],
            ["TMR (entregues)",
             f'{pg["tmr_h"]} h' if pg["tmr_h"] else "—",
             f'{pp["tmr_h"]} h' if pp["tmr_h"] else "—",
             f'{pt["tmr_h"]} h' if pt["tmr_h"] else "—"],
            ["Natureza (base relevante)",
             f'{pg["pct_req"]}% Req / {pg["pct_inc"]}% Inc',
             f'{pp["pct_req"]}% Req / {pp["pct_inc"]}% Inc',
             f'{pt["pct_req"]}% Req / {pt["pct_inc"]}% Inc'],
        ],
        [5.2, 4.1, 4.1, 3.1])
    adicionar_nota(
        doc,
        "O backlog inicial é estimado pelo fluxo (Backlog final − Novos + Entregues). "
        "Sem snapshot do estado do GLPI em 01, é a melhor aproximação disponível.")
    doc.add_paragraph()
    adicionar_imagem(doc, dados["img_donut"], 9)
    adicionar_nota(doc, "Natureza da Demanda — base relevante do período.")

    # ===================== 2. Categorias =====================
    adicionar_titulo(doc, "2. Distribuição por Categoria (base relevante)")
    for frente, palette in ((F_SUST, "navy"), (F_PROJ, "orange")):
        bg = COR["banner_blue_bg"] if palette == "navy" else COR["banner_orange_bg"]
        tx = COR["navy_rgb"] if palette == "navy" else COR["orange_rgb"]
        total_frente = dados["painel_sust"]["total_relevante"] if frente == F_SUST \
            else dados["painel_proj"]["total_relevante"]
        adicionar_banner(doc, f"FRENTE {frente.upper()} — {total_frente} chamados relevantes",
                         bg, tx)
        cat_df = dados["cat_sust"] if frente == F_SUST else dados["cat_proj"]
        rows = [[idx, int(r["qtd"]), f'{r["pct_qtd"]}%',
                 f'{r["tempo_h"]:.1f}', f'{r["pct_tempo"]}%']
                for idx, r in cat_df.iterrows()]
        adicionar_tabela(doc, ["Categoria", "Qtd.", "% Qtd.",
                               "Tempo Total (h) dos entregues", "% Tempo"],
                         rows, [5.8, 2, 2.2, 3.5, 2.5])
        doc.add_paragraph()
    adicionar_imagem(doc, dados["img_barras_sust"], 14)

    # ===================== 3. Projetos =====================
    adicionar_titulo(doc, f"3. Frente {F_PROJ} — Detalhamento Completo")
    adicionar_paragrafo(
        doc,
        f'Lista todos os chamados de {F_PROJ} relevantes para o período: '
        f'novos no mês, entregues no mês (mesmo se abertos antes) e ainda em '
        f'andamento. Projetos longos permanecem visíveis enquanto ativos.',
        italic=True)
    tab_proj = dados["tabela_projetos"]
    rows, total_concluido, total_aberto = [], 0.0, 0.0
    for _, r in tab_proj.iterrows():
        origem = str(r.get("origem", "")).strip().lower()
        eh_entregue_no_mes = origem == "entregue no mês"

        if pd.notna(r["tempo_h"]) and eh_entregue_no_mes:
            horas_txt = f'{r["tempo_h"]:.1f}'
            total_concluido += float(r["tempo_h"])
        elif pd.notna(r["horas_acumuladas"]):
            horas_txt = f'{r["horas_acumuladas"]:.1f} (corrido)'
            total_aberto += float(r["horas_acumuladas"])
        else:
            horas_txt = "—"
        rows.append([
            int(r["id"]), r["titulo_limpo"], r["origem"], r["status_nome"],
            r["date"].strftime("%d/%m/%y"),
            r["data_solucao"].strftime("%d/%m/%y") if pd.notna(r["data_solucao"]) else "—",
            f'{int(r["idade_dias"])}d',
            horas_txt,
        ])
    rows.append(["", "", "", "", "", "", "TOTAL CONCLUÍDO",
                 f'{total_concluido:,.1f}'.replace(",", ".")])
    rows.append(["", "", "", "", "", "", "TOTAL EM ABERTO (corrido)",
                 f'{total_aberto:,.1f}'.replace(",", ".")])
    adicionar_tabela(doc, ["ID", "Título", "Origem", "Status", "Abertura",
                           "Solução", "Idade", "Horas"],
                     rows, [1.2, 5.0, 2.0, 2.0, 1.7, 1.7, 1.3, 1.9])

    # ===================== 4. SLA por Prioridade (só tabela) =====================
    adicionar_titulo(doc, f"4. SLA por Prioridade — Frente {F_SUST}")
    adicionar_paragrafo(
        doc,
        "Considera apenas chamados ENTREGUES no período (resolvidos/fechados "
        "no mês), independente de quando foram abertos.", italic=True)
    sla = dados["sla_sust"]
    if sla.empty:
        adicionar_nota(doc, "Sem chamados entregues na frente Sustentação no período.")
    else:
        rows = [[idx, int(r["count"]), f'{r["mean"]:.1f} h',
                 f'{r["min"]:.1f} h – {r["max"]:.1f} h',
                 _fmt_num(r["meta_h"], " h"), f'{r["pct_dentro_sla"]}%']
                for idx, r in sla.iterrows()]
        adicionar_tabela(
            doc,
            ["Prioridade", "Entregues", "Tempo Médio", "Faixa (mín–máx)",
             "Meta SLA", "% Dentro do SLA"],
            rows, [3.2, 2.2, 2.6, 3.2, 2, 2.8])
        adicionar_nota(doc, "Meta SLA configurada em CONFIG['sla_alvo_horas'].")

    outliers = dados["outliers_sust"]
    if len(outliers["critico"]):
        top = outliers["critico"].iloc[0]
        adicionar_nota(
            doc, f'Outlier crítico (Q3 + 3·IQR = {outliers["limite_critico"]} h): '
                 f'chamado {int(top["id"])} ({top["categoria_final"]}), '
                 f'{top["tempo_h"]:.1f} h — considere tratá-lo à parte.')
    elif len(outliers["atencao"]):
        adicionar_nota(
            doc, f'{len(outliers["atencao"])} chamado(s) acima do limiar de atenção '
                 f'(Q3 + 1.5·IQR = {outliers["limite_atencao"]} h).')

    # ===================== 5. Capacidade (só tabela) =====================
    adicionar_titulo(doc, f"5. Capacidade da Equipe — Frente {F_SUST}")
    p_cfg = CONFIG["premissas_equipe"]
    adicionar_paragrafo(
        doc,
        f'Premissas: {p_cfg["analistas"]} analistas; {p_cfg["horas_por_dia"]}h/dia; '
        f'{p_cfg["dias_uteis_mes"]} dias úteis/mês; desconto de '
        f'{p_cfg["dias_ferias_no_mes"]} dias úteis de férias/ausências.',
        italic=True)
    cap = dados["cap_sust"]
    adicionar_tabela(
        doc, ["Sigla", "Descrição / Fórmula", "Valor"],
        [
            ["HMM", "Horas totais projetadas", f'{cap["hmm"]} h'],
            ["HE", "Horas efetivas (descontando ausências)", f'{cap["he"]} h'],
            ["HPC", f"Tempo registrado nos ENTREGUES do mês ({F_SUST})", f'{cap["hpc"]} h'],
            ["HHA", "Horas úteis alocadas (HPC ÷ horas/dia)", f'{cap["hha"]} h'],
            ["AD%", "Aderência (HHA ÷ HE)", f'{cap["ad_pct"]}%'],
        ],
        [2, 9.5, 3])
    if not dados.get("tempo_util_disponivel", True):
        adicionar_nota(
            doc, "⚠ A base não trouxe 'tempo_resolucao_horas'. O HPC usa tempo "
                 "CORRIDO (inclui noites/fins de semana), então a Aderência está "
                 "superestimada e deve ser lida com cautela.")
    adicionar_nota(
        doc, "Limitação conhecida: sem log de atividade do GLPI, todo o tempo "
             "de um chamado herdado é atribuído ao mês de solução. O HPC fica "
             "superestimado em meses com muitas entregas de chamados antigos.")

    # ===================== 6. Série Histórica =====================
    adicionar_titulo(doc, "6. Série Histórica Gerencial")
    serie = visao["serie"]
    if len(serie) < 2:
        adicionar_nota(doc, "A série histórica aparecerá quando houver ≥ 2 meses consolidados.")
    else:
        adicionar_paragrafo(
            doc, "Série consolidada a partir dos JSONs mensais. O mês corrente "
                 "é reprocessado a cada execução.", italic=True)
        adicionar_imagem(doc, dados["img_hist_fluxo"], 15)
        adicionar_imagem(doc, dados["img_hist_backlog_frente"], 15)
        adicionar_imagem(doc, dados["img_hist_tmr"], 15)
        adicionar_imagem(doc, dados["img_hist_taxa"], 15)

    # ===================== 7. Observações =====================
    adicionar_titulo(doc, "7. Observações e Plano de Ação")
    adicionar_paragrafo(doc, "Espaço para o autor completar com as ações do mês.",
                        italic=True)
    for item in ["Ação 1: ", "Ação 2: ", "Ação 3: "]:
        adicionar_bullet(doc, item)

    doc.save(caminho_saida)


# =============================================================================
# 7. ORQUESTRAÇÃO
# =============================================================================
def gerar_relatorio_dataframe(m: pd.DataFrame, ano: int, mes: int):
    pasta = _nova_pasta_execucao()
    logger.info("[1/5] Preparando dados do período %04d-%02d ...", ano, mes)
    txt_periodo = periodo_texto(ano, mes)
    F_SUST = CONFIG["nome_frente_sustentacao"]
    F_PROJ = CONFIG["nome_frente_projetos"]

    logger.info("[2/5] Calculando métricas ...")
    painel_total = metricas_painel(m)
    painel_sust = metricas_painel(m, F_SUST)
    painel_proj = metricas_painel(m, F_PROJ)

    cat_sust_full = metricas_categoria(m, F_SUST)
    cat_proj_full = metricas_categoria(m, F_PROJ)
    cat_sust = compactar_categorias(cat_sust_full, CONFIG["max_categorias_detalhadas"])
    cat_proj = compactar_categorias(cat_proj_full, CONFIG["max_categorias_detalhadas"])

    sla_sust = metricas_sla_prioridade(m, F_SUST)
    outliers_sust = detectar_outliers(m, F_SUST)
    cap_sust = metricas_capacidade(m, F_SUST)
    cap_proj = metricas_capacidade(m, F_PROJ)
    status_total = _status_breakdown(m)
    tabela_projetos = tabela_frente_secundaria(m, F_PROJ)

    periodo = f"{ano:04d}-{mes:02d}"
    historico_atual = _metricas_historicas(
        periodo, txt_periodo, painel_total, painel_sust, painel_proj,
        cap_sust, cap_proj, status_total, outliers_sust,
        cat_sust_full, cat_proj_full)
    historico_anterior = [r for r in carregar_historico()
                          if r.get("periodo") != periodo]
    visao = preparar_visao_historica(historico_anterior, historico_atual)

    logger.info("[3/5] Gerando gráficos ...")
    img_donut = os.path.join(pasta, "01_donut_natureza.png")
    img_barras_sust = os.path.join(pasta, "02_barras_categoria_sustentacao.png")
    img_hist_fluxo = os.path.join(pasta, "06_hist_fluxo.png")
    img_hist_backlog_frente = os.path.join(pasta, "07_hist_backlog_frente.png")
    img_hist_tmr = os.path.join(pasta, "08_hist_tmr.png")
    img_hist_taxa = os.path.join(pasta, "09_hist_taxa.png")

    grafico_donut_natureza(painel_total, img_donut)
    grafico_barras_categoria(
        cat_sust,
        f"{F_SUST} — Volume por Categoria ({painel_sust['total_relevante']} chamados)",
        img_barras_sust)

    if len(visao["serie"]) >= 2:
        grafico_fluxo_historico(visao["serie"], img_hist_fluxo)
        grafico_backlog_frente(visao["serie"], img_hist_backlog_frente)
        grafico_evolucao_historica(
            visao["serie"], "total", "tmr_h",
            "Evolução — Tempo Médio de Resolução (entregues)",
            img_hist_tmr, "h", cor_nome="navy")
        grafico_evolucao_historica(
            visao["serie"], "total", "taxa_resolucao",
            "Evolução — Taxa de Resolução",
            img_hist_taxa, "%", cor_nome="green")

    logger.info("[4/5] Montando .docx ...")
    dados = dict(
        periodo_texto=txt_periodo,
        painel_total=painel_total, painel_sust=painel_sust, painel_proj=painel_proj,
        cat_sust=cat_sust, cat_proj=cat_proj,
        sla_sust=sla_sust, outliers_sust=outliers_sust,
        cap_sust=cap_sust, cap_proj=cap_proj,
        tabela_projetos=tabela_projetos,
        historico_atual=historico_atual, visao_historica=visao,
        tempo_util_disponivel=m.attrs.get("tempo_util_disponivel", False),
        img_donut=img_donut, img_barras_sust=img_barras_sust,
        img_hist_fluxo=img_hist_fluxo,
        img_hist_backlog_frente=img_hist_backlog_frente,
        img_hist_tmr=img_hist_tmr, img_hist_taxa=img_hist_taxa,
    )
    caminho_docx = os.path.join(pasta, CONFIG["nome_docx"])
    montar_documento(dados, caminho_docx)
    salvar_historico(historico_atual)

    logger.info("[5/5] Relatório salvo em: %s", caminho_docx)
    return caminho_docx


def gerar_relatorio(caminho_entrada: str, mes_referencia: Optional[str] = None):
    ano, mes = _resolver_mes_referencia(mes_referencia)
    bruto = carregar_dados(caminho_entrada)
    m = preparar_dataframe(bruto)
    m_filtrado = filtrar_periodo_referencia(m, ano, mes)
    if m_filtrado.empty:
        raise RuntimeError(
            f"Nenhum chamado relevante para {ano}-{mes:02d} em {caminho_entrada}.")
    return gerar_relatorio_dataframe(m_filtrado, ano, mes)


# =============================================================================
# ENTRYPOINT
# =============================================================================
def _parse_args(argv):
    parser = argparse.ArgumentParser(description="Gera relatório mensal do GLPI.")
    parser.add_argument("entrada", nargs="?", default=CONFIG["arquivo_entrada"],
                        help="Caminho do .xlsx/.xlsm/.csv de entrada.")
    parser.add_argument("--mes", dest="mes", default=None,
                        help="Mês de referência YYYY-MM (default: mês corrente).")
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = _parse_args(sys.argv[1:])
    if not os.path.exists(args.entrada):
        logger.error("Arquivo não encontrado: %s", args.entrada)
        sys.exit(1)
    try:
        gerar_relatorio(args.entrada, mes_referencia=args.mes)
    except Exception as exc:
        logger.exception("Falha ao gerar relatório: %s", exc)
        sys.exit(2)