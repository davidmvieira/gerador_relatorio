#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_relatorio_glpi.py
========================
Gera o relatório mensal da Equipe de Sistemas a partir de um export do GLPI.

REGRA DE NEGÓCIO (v3 — point-in-time)
-------------------------------------
A base pode conter chamados de VÁRIOS meses. O relatório olha para um MÊS DE
REFERÊNCIA e reconstrói tudo pelas DATAS (nunca pelo status do momento do
export) — ver ``metricas.py``, a fonte única dos cálculos:

    - Entram os chamados em aberto no dia 01 e os abertos dentro do mês;
      chamados abertos depois do mês nunca entram.
    - Backlog inicial/final = em aberto no início/fim do mês.
    - Absorção = solucionados ÷ (backlog inicial + abertos).
    - Tempo de resolução em horas corridas, com média e mediana.
    - Sustentação e Projetos são sempre calculadas separadamente.
"""

import os
import re
import sys
import io
import json
import shutil
import logging
import threading
import argparse
import copy
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

import metricas

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

    # SLA institucional. Prioridade do GLPI (1 = Muito Baixa … 6 = Crítica)
    # → classe P1 (mais crítica) … P4. Mapeamento provisório: validar com a gestão.
    "sla": {
        "mapa_prioridade": {6: "P1", 5: "P1", 4: "P2", 3: "P3", 2: "P4", 1: "P4"},
        "metas": {
            "P1": {"valor": 4, "unidade": "horas"},
            "P2": {"valor": 8, "unidade": "horas"},
            "P3": {"valor": 3, "unidade": "dias_uteis"},
            "P4": {"valor": 5, "unidade": "dias_uteis"},
        },
    },

    # Dias úteis: feriados nacionais automáticos + datas extras (estaduais,
    # municipais, pontos facultativos), no formato AAAA-MM-DD.
    "calendario": {"feriados_nacionais": True, "feriados_extras": []},

    "outlier_iqr_mult_atencao": 1.5,
    "outlier_iqr_mult_critico": 3.0,
    # limiar para a "média sem outliers" do TMR (valor validado no projeto)
    "outlier_iqr_mult_tmr": metricas.OUTLIER_IQR_MULT_PADRAO,

    "premissas_equipe": {
        "analistas": 6,
        "horas_por_dia": 8,
        "dias_uteis_mes": None,  # None = calculado pelo calendário do mês
        "dias_ferias_no_mes": 11,
    },

    "pasta_saida": os.path.join("relatorios_glpi", "execucoes"),
    "nome_docx": "Relatorio_Sistemas.docx",
    "pasta_historico": os.path.join("relatorios_glpi", "historico"),
}

# Padrões de fábrica: a configuração salva pelo usuário é aplicada por cima.
CONFIG_PADRAO = copy.deepcopy(CONFIG)

# Chaves que o usuário pode alterar (tela de configuração / config.json).
CHAVES_EDITAVEIS = ("titulo_relatorio", "subtitulo", "autor", "entidade_sustentacao",
                    "max_categorias_detalhadas", "outlier_iqr_mult_tmr",
                    "premissas_equipe", "sla", "calendario")
CLASSES_SLA = ("P1", "P2", "P3", "P4")
UNIDADES_SLA = ("horas", "dias_uteis")


def caminho_config() -> str:
    """Arquivo local da configuração (fora do git): instance/config.json."""
    return os.environ.get("GLPI_CONFIG") or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "instance", "config.json")


def _normalizar_config(valores: dict) -> dict:
    """Converte tipos vindos de JSON/formulário (chaves int no mapa, números)."""
    cfg = copy.deepcopy(valores)
    sla = cfg.get("sla", {})
    if "mapa_prioridade" in sla:
        sla["mapa_prioridade"] = {int(k): str(v) for k, v in sla["mapa_prioridade"].items()}
    for meta in sla.get("metas", {}).values():
        meta["valor"] = float(meta["valor"])
    premissas = cfg.get("premissas_equipe", {})
    for chave in ("analistas", "horas_por_dia", "dias_ferias_no_mes"):
        if chave in premissas:
            premissas[chave] = float(premissas[chave])
            if premissas[chave].is_integer():
                premissas[chave] = int(premissas[chave])
    if premissas.get("dias_uteis_mes") in ("", 0, "0"):
        premissas["dias_uteis_mes"] = None
    elif premissas.get("dias_uteis_mes") is not None:
        premissas["dias_uteis_mes"] = int(premissas["dias_uteis_mes"])
    if "max_categorias_detalhadas" in cfg:
        cfg["max_categorias_detalhadas"] = int(cfg["max_categorias_detalhadas"])
    if "outlier_iqr_mult_tmr" in cfg:
        cfg["outlier_iqr_mult_tmr"] = float(cfg["outlier_iqr_mult_tmr"])
    cal = cfg.get("calendario", {})
    if "feriados_extras" in cal:
        cal["feriados_extras"] = sorted({str(d).strip() for d in cal["feriados_extras"]
                                         if str(d).strip()})
    return cfg


def validar_config(valores: dict) -> list:
    """Lista de erros (vazia = válida). Não altera nada."""
    erros = []
    try:
        cfg = _normalizar_config(valores)
    except (TypeError, ValueError, AttributeError, KeyError) as erro:
        return [f"Valor inválido: {erro}"]
    for chave in ("titulo_relatorio", "entidade_sustentacao"):
        if chave in cfg and not str(cfg[chave]).strip():
            erros.append(f"'{chave}' não pode ficar vazio.")
    p = cfg.get("premissas_equipe", {})
    if p.get("analistas", 1) <= 0 or p.get("horas_por_dia", 1) <= 0:
        erros.append("Analistas e horas por dia devem ser maiores que zero.")
    if p.get("dias_ferias_no_mes", 0) < 0:
        erros.append("Dias de férias/ausência não pode ser negativo.")
    sla = cfg.get("sla", {})
    for prioridade, classe in sla.get("mapa_prioridade", {}).items():
        if prioridade not in range(1, 7) or classe not in CLASSES_SLA:
            erros.append(f"Mapeamento inválido: prioridade {prioridade} → {classe}.")
    for classe, meta in sla.get("metas", {}).items():
        if classe not in CLASSES_SLA or meta.get("unidade") not in UNIDADES_SLA \
                or meta["valor"] <= 0:
            erros.append(f"Meta de SLA inválida para {classe}.")
    for data in cfg.get("calendario", {}).get("feriados_extras", []):
        if pd.isna(pd.to_datetime(data, format="%Y-%m-%d", errors="coerce")):
            erros.append(f"Feriado extra inválido: '{data}' (use AAAA-MM-DD).")
    if cfg.get("max_categorias_detalhadas", 1) < 1:
        erros.append("Máximo de categorias deve ser ao menos 1.")
    if cfg.get("outlier_iqr_mult_tmr", 1) <= 0:
        erros.append("Multiplicador de outlier deve ser maior que zero.")
    return erros


def aplicar_config(valores: dict) -> None:
    """Aplica (em memória) os valores editáveis sobre os padrões de fábrica."""
    novo = copy.deepcopy(CONFIG_PADRAO)
    novo.update({k: v for k, v in _normalizar_config(valores).items()
                 if k in CHAVES_EDITAVEIS})
    preservar = {k: CONFIG[k] for k in ("pasta_saida", "pasta_historico") if k in CONFIG}
    CONFIG.clear()
    CONFIG.update(novo)
    CONFIG.update(preservar)


def carregar_config(caminho: Optional[str] = None) -> list:
    """Lê o config.json local (se existir) e aplica. Devolve erros encontrados."""
    caminho = caminho or caminho_config()
    if not os.path.isfile(caminho):
        return []
    try:
        with open(caminho, "r", encoding="utf-8") as arquivo:
            valores = json.load(arquivo)
    except (OSError, json.JSONDecodeError) as erro:
        logger.warning("Configuração ignorada (%s): %s", caminho, erro)
        return [str(erro)]
    erros = validar_config(valores)
    if erros:
        logger.warning("Configuração ignorada (%s): %s", caminho, "; ".join(erros))
        return erros
    aplicar_config(valores)
    return []


def salvar_config(valores: dict, caminho: Optional[str] = None) -> list:
    """Valida, grava o config.json local e aplica. Devolve erros (vazio = salvo)."""
    valores = {k: v for k, v in valores.items() if k in CHAVES_EDITAVEIS}
    erros = validar_config(valores)
    if erros:
        return erros
    cfg = _normalizar_config(valores)
    caminho = caminho or caminho_config()
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    serializavel = copy.deepcopy(cfg)
    if "sla" in serializavel:
        serializavel["sla"]["mapa_prioridade"] = {
            str(k): v for k, v in serializavel["sla"]["mapa_prioridade"].items()}
    temporario = caminho + ".tmp"
    with open(temporario, "w", encoding="utf-8") as arquivo:
        json.dump(serializavel, arquivo, ensure_ascii=False, indent=2)
    os.replace(temporario, caminho)
    aplicar_config(cfg)
    return []


def config_editavel() -> dict:
    """Valores atuais das chaves editáveis (para a tela)."""
    return {k: copy.deepcopy(CONFIG[k]) for k in CHAVES_EDITAVEIS}


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


def _ler_entrada(entrada, nome: Optional[str] = None):
    """Lê CSV/Excel de um caminho ou de um arquivo em memória (``nome`` dá a extensão)."""
    nome = nome or str(entrada)
    extensao = os.path.splitext(nome)[1].lower()
    if not isinstance(entrada, (str, os.PathLike)):
        entrada = io.BytesIO(entrada.read() if hasattr(entrada, "read") else entrada)
    if extensao == ".csv":
        try:
            return pd.read_csv(entrada, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError:
            if hasattr(entrada, "seek"):
                entrada.seek(0)
            return pd.read_csv(entrada, sep=None, engine="python", encoding="latin-1")
    if extensao in (".xlsx", ".xlsm"):
        return pd.ExcelFile(entrada)
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


def carregar_dados(caminho_entrada, nome: Optional[str] = None) -> pd.DataFrame:
    """``caminho_entrada`` pode ser um caminho ou bytes/arquivo em memória (com ``nome``)."""
    entrada = _ler_entrada(caminho_entrada, nome)
    eh_csv = isinstance(entrada, pd.DataFrame)
    abas = [None] if eh_csv else entrada.sheet_names

    if len(abas) == 1:
        df = entrada if eh_csv else pd.read_excel(entrada, sheet_name=abas[0])
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
        df1 = pd.read_excel(entrada, sheet_name=CONFIG["aba_principal"])
        df2 = pd.read_excel(entrada, sheet_name=CONFIG["aba_detalhe"])
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

    # Tempo de resolução SEMPRE em horas corridas (solução − abertura). A coluna
    # 'tempo_resolucao_horas' do export também é tempo corrido (validado contra
    # as datas); só é usada quando as datas não permitem o cálculo.
    solucao = metricas.solucao_efetiva(m)
    tempo_corrido_concluido = (solucao - m["data_abertura"]).dt.total_seconds() / 3600
    if "tempo_resolucao_horas" in m.columns:
        tempo_corrido_concluido = tempo_corrido_concluido.fillna(
            pd.to_numeric(m["tempo_resolucao_horas"], errors="coerce"))
    tempo_corrido_aberto = (agora - m["data_abertura"]).dt.total_seconds() / 3600

    m["tempo_h"] = np.where(concluido, tempo_corrido_concluido, np.nan)
    m["horas_acumuladas"] = np.where(~concluido, tempo_corrido_aberto, np.nan)
    m["idade_dias"] = ((agora - m["data_abertura"]).dt.total_seconds() / 86400).round(0)
    return m


# ---------------------------------------------------------------------------
# FILTRO DE PERÍODO — CORAÇÃO DA MUDANÇA
# ---------------------------------------------------------------------------
def filtrar_periodo_referencia(m: pd.DataFrame, ano: int, mes: int,
                               agora: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    """
    Retorna os chamados que TOCARAM o mês de referência, reconstruídos pelas
    datas (point-in-time — ver ``metricas.marcar_periodo``):
      1. Em aberto no dia 01 (backlog inicial), OU
      2. Abertos dentro do mês.
    Todo chamado solucionado no mês cai em um dos dois. Chamados abertos
    DEPOIS do mês nunca entram, mesmo que ainda estejam abertos hoje.

    Flags (nomes mantidos por compatibilidade com o app e o .docx):
      novo_no_mes      : aberto dentro do mês
      entregue_no_mes  : solucionado dentro do mês
      herdado_aberto   : aberto antes do mês e ainda em aberto no fim do mês
      backlog_inicio / backlog_fim : em aberto no início / fim do mês
    """
    marcado = metricas.marcar_periodo(m, ano, mes, agora=agora)
    sub = marcado[marcado["relevante"]].copy()
    sub["novo_no_mes"] = sub["aberto_no_mes"]
    sub["entregue_no_mes"] = sub["solucionado_no_mes"]
    sub["herdado_aberto"] = ~sub["novo_no_mes"] & sub["backlog_fim"]
    sub["tempo_h"] = sub["tempo_resolucao_h"]
    sub["horas_acumuladas"] = sub["horas_em_aberto_h"]

    sub.attrs.update(marcado.attrs)
    logger.info(
        "Período %04d-%02d: %d chamados no mês (backlog inicial %d, %d abertos, "
        "%d solucionados, backlog final %d)",
        ano, mes, len(sub),
        int(sub["backlog_inicio"].sum()), int(sub["novo_no_mes"].sum()),
        int(sub["entregue_no_mes"].sum()), int(sub["backlog_fim"].sum()),
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
    Painel de FLUXO do mês, contado pelas datas (``metricas.indicadores``):
        backlog_inicial  = em aberto no dia 01
        backlog_final    = em aberto no fim do mês
        backlog_final    = backlog_inicial + novos − entregues  (identidade)
        taxa_resolucao   = absorção = entregues ÷ (backlog_inicial + novos)
        tmr_h / tmr_mediana_h = média / mediana (horas corridas) dos entregues
    Recebe o resultado de ``filtrar_periodo_referencia``.
    """
    sub = m if frente is None else m[m["frente"] == frente]
    ind = metricas.indicadores(sub, iqr_mult=CONFIG["outlier_iqr_mult_tmr"])
    if not ind["fluxo_confere"]:
        logger.warning("Fluxo do backlog não fecha para %s — verifique datas inválidas.",
                       frente or "todas as frentes")

    def pct(x, base):
        return round(x / base * 100, 1) if base else 0.0

    return {
        **ind,
        # nomes legados usados pelo app, histórico e .docx
        "novos": ind["abertos"],
        "entregues": ind["solucionados"],
        "taxa_resolucao": ind["absorcao_pct"] if ind["absorcao_pct"] is not None else 0.0,
        "tmr_h": ind["tmr_media_h"],
        "req": ind["requisicoes"], "inc": ind["incidentes"],
        "outros": max(ind["outros_tipos"], 0),
        "pct_req": pct(ind["requisicoes"], ind["abertos"]),
        "pct_inc": pct(ind["incidentes"], ind["abertos"]),
        "pct_outros": pct(max(ind["outros_tipos"], 0), ind["abertos"]),
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


def feriados_configurados(anos) -> list:
    cal = CONFIG["calendario"]
    return metricas.calendario_feriados(anos, cal.get("feriados_extras", []),
                                        cal.get("feriados_nacionais", True))


def _anos_do_recorte(m: pd.DataFrame) -> range:
    """Anos cobertos pelas aberturas e pelo corte, mais o seguinte (prazos que viram o ano)."""
    datas = pd.to_datetime(m.get("data_abertura", m.get("date")), errors="coerce").dropna()
    corte = m.attrs.get("corte", pd.Timestamp.now())
    inicio = int(datas.min().year) if len(datas) else int(corte.year)
    return range(inicio, int(corte.year) + 2)


def metricas_sla(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    """SLA institucional (P1–P4) da frente: solucionados no mês e vencidos em aberto."""
    sla = CONFIG["sla"]
    sub = m[m["frente"] == frente]
    avaliado = metricas.avaliar_sla(sub, sla["mapa_prioridade"], sla["metas"],
                                    feriados_configurados(_anos_do_recorte(m)))
    return metricas.resumo_sla(avaliado, sla["metas"])


def pct_no_prazo(resumo: pd.DataFrame):
    """% dos solucionados no mês que cumpriram o prazo, todas as classes juntas."""
    if resumo.empty or not resumo["solucionados"].sum():
        return None
    return round(resumo["no_prazo"].sum() / resumo["solucionados"].sum() * 100, 1)


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


def dias_uteis_do_mes(m: pd.DataFrame) -> tuple[int, bool]:
    """(dias úteis, calculado automaticamente?) do mês do recorte."""
    informado = CONFIG["premissas_equipe"].get("dias_uteis_mes")
    if informado:
        return int(informado), False
    inicio = m.attrs.get("periodo_inicio")
    if inicio is None:
        return 0, True
    return metricas.dias_uteis_no_mes(inicio.year, inicio.month,
                                      feriados_configurados([inicio.year])), True


def metricas_capacidade(m: pd.DataFrame, frente: str) -> dict:
    p = CONFIG["premissas_equipe"]
    dias_uteis, automatico = dias_uteis_do_mes(m)
    hmm = p["analistas"] * dias_uteis * p["horas_por_dia"]
    he = hmm - p["dias_ferias_no_mes"] * p["horas_por_dia"]

    sub = m[(m["frente"] == frente) & m["entregue_no_mes"]]
    hpc = round(float(sub["tempo_h"].sum()), 1)
    hha = round(hpc / p["horas_por_dia"], 1)
    ad_pct = round(hha / he * 100, 1) if he else 0
    return {"hmm": hmm, "he": he, "hpc": hpc, "hha": hha, "ad_pct": ad_pct,
            "dias_uteis": dias_uteis, "dias_uteis_automatico": automatico}


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
    sub["origem"] = np.where(sub["entregue_no_mes"], "Entregue no mês",
                             np.where(sub["novo_no_mes"], "Novo", "Em andamento"))
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
# Versão da regra de cálculo gravada em cada JSON mensal. Meses calculados com
# outra regra NÃO são comparados (números com definições diferentes).
REGRA_METRICAS = "point-in-time-v1"

# (rótulo, (frente, campo), unidade) — sempre por frente, nunca agregado.
INDICADORES_HISTORICOS = [
    ("Sustentação — abertos", ("sustentacao", "novos"), ""),
    ("Sustentação — solucionados", ("sustentacao", "entregues"), ""),
    ("Sustentação — backlog final", ("sustentacao", "backlog_final"), ""),
    ("Sustentação — absorção", ("sustentacao", "taxa_resolucao"), "%"),
    ("Sustentação — TMR mediana", ("sustentacao", "tmr_mediana_h"), "h"),
    ("Sustentação — TMR média", ("sustentacao", "tmr_h"), "h"),
    ("Sustentação — no prazo (SLA)", ("sustentacao", "sla_pct_no_prazo"), "%"),
    ("Projetos — abertos", ("projetos", "novos"), ""),
    ("Projetos — solucionados", ("projetos", "entregues"), ""),
    ("Projetos — backlog final", ("projetos", "backlog_final"), ""),
]


def _metricas_historicas(periodo: str, periodo_txt: str,
                         painel_sust, painel_proj,
                         cap_sust, cap_proj, status_total,
                         outliers, cat_sust_full, cat_proj_full) -> dict:
    def painel_resumido(p):
        return {
            "novos": p["novos"], "entregues": p["entregues"],
            "backlog_inicial": p["backlog_inicial"],
            "backlog_final": p["backlog_final"],
            "saldo": p["saldo"],
            "carga_total": p["carga_total"],
            "taxa_resolucao": _numero(p["taxa_resolucao"], 1),
            "tmr_h": _numero(p["tmr_h"], 1),
            "tmr_mediana_h": _numero(p["tmr_mediana_h"], 1),
            "tmr_media_sem_outliers_h": _numero(p["tmr_media_sem_outliers_h"], 1),
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
        "regra_metricas": REGRA_METRICAS,
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


def carregar_historico() -> tuple[list, list]:
    """Retorna (registros compatíveis, períodos ignorados por regra antiga)."""
    historico, ignorados = [], []
    pasta = CONFIG["pasta_historico"]
    if not os.path.isdir(pasta):
        return historico, ignorados
    padrao = re.compile(r"\d{4}-\d{2}\.json$")
    for nome in sorted(os.listdir(pasta)):
        if not padrao.fullmatch(nome):
            continue
        try:
            with open(os.path.join(pasta, nome), "r", encoding="utf-8") as f:
                reg = json.load(f)
        except (OSError, json.JSONDecodeError, TypeError) as erro:
            logger.warning("Histórico ignorado (%s): %s", nome, erro)
            continue
        reg["periodo"] = nome[:-5]
        if reg.get("regra_metricas") != REGRA_METRICAS:
            logger.warning("Histórico %s calculado com regra antiga — fora da comparação.",
                           reg["periodo"])
            ignorados.append(reg["periodo"])
            continue
        historico.append(reg)
    return sorted(historico, key=lambda x: x.get("periodo", "")), ignorados


def salvar_historico(registro: dict):
    pasta = CONFIG["pasta_historico"]
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, f'{registro["periodo"]}.json')
    if os.path.exists(destino):
        shutil.copy2(destino, destino + ".bak")
        logger.info("Backup do histórico: %s.bak", destino)
    registro = {k: v for k, v in registro.items() if k != "fonte"}
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
    comparacoes = []
    for nome, caminho, unidade in INDICADORES_HISTORICOS:
        va = _safe_get(atual, *caminho)
        vb = _safe_get(anterior, *caminho)
        if va is None or vb is None:
            continue
        dif = va - vb
        var = dif / vb * 100 if vb else None
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
        for nome, caminho, unidade in INDICADORES_HISTORICOS:
            valores = [v for v in (_safe_get(item, *caminho) for item in ultimos)
                       if v is not None]
            if valores:
                projecao[nome] = (round(sum(valores) / len(valores), 1), unidade)
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
    ax.set_title("Sustentação — Natureza dos Chamados Abertos", fontsize=14, fontweight="bold",
                 color="#" + COR["navy"], pad=20)
    ax.text(0, 0, f"{total}\nabertos", ha="center", va="center",
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
    """Sustentação: abertos x solucionados x backlog final por mês."""
    periodos = [it.get("periodo", "?") for it in serie]
    x = list(range(len(periodos)))
    novos = np.array([_safe_get(it, "sustentacao", "novos", default=np.nan) for it in serie], dtype=float)
    entregues = np.array([_safe_get(it, "sustentacao", "entregues", default=np.nan) for it in serie], dtype=float)
    backlog = np.array([_safe_get(it, "sustentacao", "backlog_final", default=np.nan) for it in serie], dtype=float)

    if np.all(np.isnan(novos)) and np.all(np.isnan(entregues)):
        return
    fig, ax = plt.subplots(figsize=(9.5, 5), dpi=150)
    ax.plot(x, novos, marker="o", linewidth=2.4, color="#" + COR["navy"], label="Abertos")
    ax.plot(x, entregues, marker="s", linewidth=2.4, color="#" + COR["green"], label="Solucionados")
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
    ax.set_title("Sustentação — Abertos, Solucionados e Backlog final",
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


def adicionar_imagem(doc, imagem: Optional[bytes], largura_cm):
    """Insere um PNG em memória (bytes). ``None`` é ignorado (gráfico sem dados)."""
    if not imagem:
        return
    doc.add_picture(io.BytesIO(imagem), width=Cm(largura_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _descricao_periodo_anterior(atual: str, anterior: str) -> str:
    dt_a = datetime.strptime(atual, "%Y-%m")
    dt_b = datetime.strptime(anterior, "%Y-%m")
    distancia = (dt_a.year - dt_b.year) * 12 + dt_a.month - dt_b.month
    if distancia == 1:
        return f"comparação com {anterior}"
    return f"comparação com {anterior} (lacuna de {distancia - 1} mês(es))"


# =============================================================================
# 6b. TABELAS E TEXTOS COMPARTILHADOS (prévia HTML e .docx usam os mesmos)
# =============================================================================
# Cada função devolve (cabeçalhos, linhas) já formatados como texto. A prévia
# da tela e o documento Word montam suas tabelas a partir daqui, então o que
# aparece na tela é exatamente o que vai impresso.

def _vazio(valor) -> bool:
    return valor is None or (isinstance(valor, (float, np.floating)) and np.isnan(valor))


def _h(valor) -> str:
    return "—" if _vazio(valor) else f"{valor} h"


def _p(valor) -> str:
    return "—" if _vazio(valor) else f"{valor}%"


def _saldo(valor) -> str:
    return f"+{valor}" if valor > 0 else str(valor)


def tabela_painel(dados: dict) -> tuple[list, list]:
    pg, pp = dados["painel_sust"], dados["painel_proj"]
    linhas = [
        ["Backlog inicial (dia 01)", pg["backlog_inicial"], pp["backlog_inicial"]],
        ["Abertos no período", pg["novos"], pp["novos"]],
        ["Solucionados no período", pg["entregues"], pp["entregues"]],
        ["Saldo (abertos − solucionados)", _saldo(pg["saldo"]), _saldo(pp["saldo"])],
        ["Backlog final", pg["backlog_final"], pp["backlog_final"]],
        ["Taxa de absorção", _p(pg["absorcao_pct"]), _p(pp["absorcao_pct"])],
        ["Tempo de resolução — mediana", _h(pg["tmr_mediana_h"]), _h(pp["tmr_mediana_h"])],
        ["Tempo de resolução — média", _h(pg["tmr_h"]), _h(pp["tmr_h"])],
        ["Natureza dos abertos",
         f'{pg["pct_req"]}% Req / {pg["pct_inc"]}% Inc',
         f'{pp["pct_req"]}% Req / {pp["pct_inc"]}% Inc'],
    ]
    return ["Indicador", dados["F_SUST"], dados["F_PROJ"]], linhas


def tabela_comparacao(dados: dict) -> tuple[list, list]:
    linhas = [[c["indicador"],
               _fmt_num(c["atual"], c["unidade"]),
               _fmt_num(c["anterior"], c["unidade"]),
               _fmt_variacao(c["diferenca"],
                             " p.p." if c["unidade"] == "%" else c["unidade"]),
               _fmt_variacao(c["variacao_pct"], "%")]
              for c in dados["visao_historica"]["comparacoes"]]
    return ["Indicador", "Atual", "Anterior", "Diferença", "Variação"], linhas


def tabela_projecao(dados: dict) -> tuple[list, list]:
    linhas = [[nome, _fmt_num(valor, unidade)]
              for nome, (valor, unidade) in dados["visao_historica"]["projecao"].items()]
    return ["Indicador", "Estimativa"], linhas


def tabela_serie(dados: dict) -> tuple[list, list]:
    """Série mensal da Sustentação, com a origem de cada mês."""
    linhas = []
    for reg in dados["visao_historica"]["serie"]:
        s = reg.get("sustentacao", {})
        periodo = reg["periodo"] + (" *" if reg.get("backlog_inicial_incompleto") else "")
        if reg["periodo"] == dados["mes_referencia"]:
            fonte = "este relatório"
        else:
            fonte = reg.get("fonte", FONTE_REGISTRADO)
        linhas.append([periodo, _fmt_num(s.get("novos")), _fmt_num(s.get("entregues")),
                       _fmt_num(s.get("backlog_final")), _p(s.get("taxa_resolucao")),
                       _h(s.get("tmr_mediana_h")), fonte])
    return ["Mês", "Abertos", "Solucionados", "Backlog final", "Absorção",
            "TMR mediana", "Origem"], linhas


TEXTO_SERIE = (
    "Meses anteriores presentes na planilha são calculados na hora, com a mesma "
    "regra deste relatório. Quando um mês foi registrado no histórico, valem os "
    "números registrados (os já apresentados). * = a planilha não cobre os meses "
    "anteriores a ele, então o backlog inicial desse mês pode estar subestimado.")


TEXTO_SEM_ANTERIOR = ("Sem mês anterior para comparação: a planilha não contém o mês "
                      "anterior e ele não foi registrado no histórico.")
TEXTO_SEM_SERIE = ("A série aparece quando houver ao menos 2 meses: na própria planilha "
                   "ou registrados no histórico.")


def tabela_categorias(cat_df: pd.DataFrame) -> tuple[list, list]:
    linhas = [[idx, int(r["qtd"]), f'{r["pct_qtd"]}%',
               f'{r["tempo_h"]:.1f}', f'{r["pct_tempo"]}%']
              for idx, r in cat_df.iterrows()]
    return ["Categoria", "Qtd.", "% Qtd.", "Tempo Total (h) dos solucionados", "% Tempo"], linhas


def tabela_projetos(tab_proj: pd.DataFrame) -> tuple[list, list]:
    linhas, total_concluido, total_aberto = [], 0.0, 0.0
    for _, r in tab_proj.iterrows():
        entregue = str(r.get("origem", "")).strip().lower() == "entregue no mês"
        if entregue and pd.notna(r["tempo_h"]):
            horas_txt = f'{r["tempo_h"]:.1f}'
            total_concluido += float(r["tempo_h"])
        elif pd.notna(r["horas_acumuladas"]):
            horas_txt = f'{r["horas_acumuladas"]:.1f} (corrido)'
            total_aberto += float(r["horas_acumuladas"])
        else:
            horas_txt = "—"
        solucao = r.get("data_solucao_efetiva", r["data_solucao"])
        linhas.append([
            int(r["id"]), r["titulo_limpo"], r["origem"], r["status_nome"],
            r["date"].strftime("%d/%m/%y"),
            solucao.strftime("%d/%m/%y") if entregue and pd.notna(solucao) else "—",
            f'{int(r["idade_dias"])}d' if pd.notna(r["idade_dias"]) else "—",
            horas_txt,
        ])
    linhas.append(["", "", "", "", "", "", "TOTAL CONCLUÍDO",
                   f'{total_concluido:,.1f}'.replace(",", ".")])
    linhas.append(["", "", "", "", "", "", "TOTAL EM ABERTO (corrido)",
                   f'{total_aberto:,.1f}'.replace(",", ".")])
    return ["ID", "Título", "Origem", "Status", "Abertura", "Solução", "Idade", "Horas"], linhas


_NOMES_PRIORIDADE = {6: "Crítica", 5: "Muito Alta", 4: "Alta", 3: "Média",
                     2: "Baixa", 1: "Muito Baixa"}


def texto_meta_sla(meta: dict) -> str:
    valor = _fmt_num(int(meta["valor"]) if float(meta["valor"]).is_integer() else meta["valor"])
    return f"{valor} h corridas" if meta["unidade"] == "horas" else f"{valor} dias úteis"


def tabela_sla(sla: pd.DataFrame) -> tuple[list, list]:
    metas = CONFIG["sla"]["metas"]
    mapa = CONFIG["sla"]["mapa_prioridade"]
    linhas = []
    for _, r in sla.iterrows():
        prioridades = ", ".join(_NOMES_PRIORIDADE[p] for p in sorted(mapa, reverse=True)
                                if mapa[p] == r["classe"]) or "—"
        linhas.append([r["classe"], prioridades, texto_meta_sla(metas[r["classe"]]),
                       r["solucionados"], r["no_prazo"], _p(r["pct_no_prazo"]),
                       _h(r["mediana_h"]), r["vencidos_em_aberto"]])
    total_sol = int(sla["solucionados"].sum()) if len(sla) else 0
    linhas.append(["Total", "", "", total_sol,
                   int(sla["no_prazo"].sum()) if len(sla) else 0,
                   _p(pct_no_prazo(sla)), "",
                   int(sla["vencidos_em_aberto"].sum()) if len(sla) else 0])
    return ["Classe", "Prioridades GLPI", "Prazo de solução", "Solucionados",
            "No prazo", "% no prazo", "Mediana", "Vencidos em aberto"], linhas


TEXTO_SLA = ("Prazo de solução por classe de criticidade (SLA institucional). Considera os "
             "chamados SOLUCIONADOS no período, independente de quando foram abertos, e os "
             "que terminaram o mês em aberto com o prazo vencido.")


def notas_sla(dados: dict) -> list:
    notas = [
        "Prazo em horas: contado em horas corridas desde a abertura (atendimento 24×7). "
        "Prazo em dias úteis: vence no fim do N-ésimo dia útil seguinte à abertura, "
        "sem fins de semana e feriados configurados.",
        "Vencidos em aberto: chamados ainda sem solução no fim do mês cujo prazo já passou.",
        "SLA de resposta (primeiro atendimento) não é medido: o export do GLPI não traz "
        "a data do primeiro atendimento.",
    ]
    sem_classe = dados["sla_sust"].attrs.get("sem_classe", 0)
    if sem_classe:
        notas.append(f"⚠ {sem_classe} chamado(s) sem prioridade mapeada para P1–P4 "
                     "ficaram fora do SLA.")
    return notas


def tabela_capacidade(cap: dict, frente: str) -> tuple[list, list]:
    return ["Sigla", "Descrição / Fórmula", "Valor"], [
        ["HMM", "Horas totais projetadas", f'{cap["hmm"]} h'],
        ["HE", "Horas efetivas (descontando ausências)", f'{cap["he"]} h'],
        ["HPC", f"Tempo registrado nos SOLUCIONADOS do mês ({frente})", f'{cap["hpc"]} h'],
        ["HHA", "Horas úteis alocadas (HPC ÷ horas/dia)", f'{cap["hha"]} h'],
        ["AD%", "Aderência (HHA ÷ HE)", f'{cap["ad_pct"]}%'],
    ]


TEXTO_ESCOPO = [
    "Chamados EM ABERTO no dia 01 (backlog inicial), mesmo que abertos em meses anteriores;",
    "Chamados ABERTOS dentro do período.",
]

TEXTO_LEITURA_PAINEL = (
    "Leitura: Backlog inicial é o que estava em aberto no dia 01; Abertos são "
    "as novas demandas do mês; Solucionados são os concluídos no mês (mesmo que "
    "abertos antes); Backlog final é o que permanece em aberto no último dia. "
    "Saldo = Abertos − Solucionados (positivo: a fila cresceu).")

TEXTO_DEFINICOES_PAINEL = (
    "Taxa de absorção = solucionados ÷ (backlog inicial + abertos): quanto da "
    "carga do mês foi resolvida; nunca passa de 100%. Tempos em horas CORRIDAS "
    "(abertura → solução, incluindo noites e fins de semana). A mediana mostra "
    "o caso típico; a média é puxada por chamados muito longos.")


def texto_escopo_nota(dados: dict) -> str:
    return ("Todos os números são reconstruídos pelas datas de abertura e solução, "
            "e não pelo status atual do chamado: o mesmo arquivo gera sempre os mesmos "
            "números para o mesmo mês, e chamados abertos depois do período não entram. "
            f'As frentes {dados["F_SUST"]} e {dados["F_PROJ"]} são apresentadas '
            "separadamente e nunca somadas em taxas ou médias.")


def avisos_painel(dados: dict) -> list:
    """Avisos que mudam a leitura dos números do mês (mostrados na tela e no .docx)."""
    avisos = []
    if dados.get("mes_parcial") and dados.get("corte") is not None:
        avisos.append(
            f'⚠ Mês em andamento: dados considerados até '
            f'{pd.Timestamp(dados["corte"]).strftime("%d/%m/%Y %H:%M")}. '
            "O backlog final é o desse momento, não o do último dia do mês.")
    if dados.get("backlog_inicial_incompleto"):
        avisos.append(
            "⚠ A planilha não contém chamados abertos antes deste mês: o backlog "
            "inicial reflete apenas o que está no export e pode estar subestimado. "
            "Para um valor fiel, exporte também os chamados abertos em meses "
            "anteriores que ainda estavam pendentes no dia 01.")
    pg = dados["painel_sust"]
    if pg["n_outliers"]:
        avisos.append(
            f'{dados["F_SUST"]}: {pg["n_outliers"]} chamado(s) acima de '
            f'{pg["limite_outlier_h"]} h (Q3 + {CONFIG["outlier_iqr_mult_tmr"]:g}·IQR). '
            f'Sem eles, a média seria {_h(pg["tmr_media_sem_outliers_h"])}.')
    return avisos


def texto_historico_atual(dados: dict) -> str:
    atual = dados["historico_atual"]
    return (f'{dados["F_SUST"]} em {atual["periodo"]}: '
            f'{atual["sustentacao"]["novos"]} abertos, '
            f'{atual["sustentacao"]["entregues"]} solucionados, '
            f'backlog final de {atual["sustentacao"]["backlog_final"]}.')


def texto_historico_ignorado(dados: dict) -> Optional[str]:
    if not dados.get("historico_ignorado"):
        return None
    return ("Meses fora da comparação por terem sido calculados com a regra "
            f'anterior: {", ".join(dados["historico_ignorado"])}. Registre-os novamente '
            "a partir do export para incluí-los.")


def texto_outlier_sla(dados: dict) -> Optional[str]:
    outliers = dados["outliers_sust"]
    if len(outliers["critico"]):
        top = outliers["critico"].iloc[0]
        return (f'Outlier crítico (Q3 + 3·IQR = {outliers["limite_critico"]} h): '
                f'chamado {int(top["id"])} ({top["categoria_final"]}), '
                f'{top["tempo_h"]:.1f} h — considere tratá-lo à parte.')
    if len(outliers["atencao"]):
        return (f'{len(outliers["atencao"])} chamado(s) acima do limiar de atenção '
                f'(Q3 + 1.5·IQR = {outliers["limite_atencao"]} h).')
    return None


def texto_premissas(cap: dict) -> str:
    p_cfg = CONFIG["premissas_equipe"]
    origem = "pelo calendário do mês" if cap.get("dias_uteis_automatico") else "informados"
    return (f'Premissas: {p_cfg["analistas"]} analistas; {p_cfg["horas_por_dia"]}h/dia; '
            f'{cap.get("dias_uteis")} dias úteis no mês ({origem}); desconto de '
            f'{p_cfg["dias_ferias_no_mes"]} dias úteis de férias/ausências.')


NOTAS_CAPACIDADE = [
    "⚠ O HPC soma tempo CORRIDO (inclui noites e fins de semana), não "
    "esforço registrado. A Aderência é indicativa e não deve ser lida "
    "como ocupação real da equipe.",
    "Limitação conhecida: sem log de atividade do GLPI, todo o tempo "
    "de um chamado herdado é atribuído ao mês de solução. O HPC fica "
    "superestimado em meses com muitas entregas de chamados antigos.",
]


# =============================================================================
# 6c. MONTAGEM DO .DOCX (em memória)
# =============================================================================
def montar_documento(dados: dict, imagens: dict, destino):
    """Monta o .docx. ``destino`` pode ser um caminho ou um arquivo em memória."""
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

    F_SUST, F_PROJ = dados["F_SUST"], dados["F_PROJ"]
    visao = dados["visao_historica"]

    # ===================== Escopo do Período =====================
    adicionar_titulo(doc, "Escopo do Período")
    adicionar_paragrafo(
        doc, f'Este relatório considera todo chamado que passou por {dados["periodo_texto"]}:')
    for b in TEXTO_ESCOPO:
        adicionar_bullet(doc, b)
    adicionar_nota(doc, texto_escopo_nota(dados))

    # ===================== Visão Histórica =====================
    adicionar_titulo(doc, "Visão Histórica")
    adicionar_paragrafo(doc, texto_historico_atual(dados))
    if texto_historico_ignorado(dados):
        adicionar_nota(doc, texto_historico_ignorado(dados))
    if visao["anterior"]:
        adicionar_paragrafo(
            doc,
            _descricao_periodo_anterior(dados["mes_referencia"], visao["anterior"]["periodo"])
            + ". Os valores não são classificados como favoráveis ou desfavoráveis.",
            italic=True)
        cab, linhas = tabela_comparacao(dados)
        adicionar_tabela(doc, cab, linhas, [4.8, 2.6, 2.6, 2.9, 2.9])
    else:
        adicionar_nota(doc, TEXTO_SEM_ANTERIOR)

    if visao["projecao"]:
        adicionar_titulo(doc, "Projeção do Próximo Mês", nivel=2)
        adicionar_paragrafo(
            doc, "ESTIMATIVA via média móvel simples dos 3 últimos meses. "
                 "Não é uma previsão garantida.", italic=True)
        cab, linhas = tabela_projecao(dados)
        adicionar_tabela(doc, cab, linhas, [8, 5])

    # ===================== 1. Painel Geral (fluxo) =====================
    adicionar_titulo(doc, "1. Painel Geral — Fluxo do Período")
    adicionar_paragrafo(doc, TEXTO_LEITURA_PAINEL)
    cab, linhas = tabela_painel(dados)
    adicionar_tabela(doc, cab, linhas, [6.2, 4.6, 4.6])
    adicionar_nota(doc, TEXTO_DEFINICOES_PAINEL)
    for aviso in avisos_painel(dados):
        adicionar_nota(doc, aviso)
    doc.add_paragraph()
    adicionar_imagem(doc, imagens.get("donut"), 9)
    adicionar_nota(doc, f"Natureza dos chamados abertos no período — frente {F_SUST}.")

    # ===================== 2. Categorias =====================
    adicionar_titulo(doc, "2. Distribuição por Categoria (base relevante)")
    for frente, palette, chave_painel, chave_cat in (
            (F_SUST, "navy", "painel_sust", "cat_sust"),
            (F_PROJ, "orange", "painel_proj", "cat_proj")):
        bg = COR["banner_blue_bg"] if palette == "navy" else COR["banner_orange_bg"]
        tx = COR["navy_rgb"] if palette == "navy" else COR["orange_rgb"]
        adicionar_banner(
            doc, f'FRENTE {frente.upper()} — {dados[chave_painel]["total_relevante"]} '
                 'chamados no mês', bg, tx)
        cab, linhas = tabela_categorias(dados[chave_cat])
        adicionar_tabela(doc, cab, linhas, [5.8, 2, 2.2, 3.5, 2.5])
        doc.add_paragraph()
    adicionar_imagem(doc, imagens.get("barras_sust"), 14)

    # ===================== 3. Projetos =====================
    adicionar_titulo(doc, f"3. Frente {F_PROJ} — Detalhamento Completo")
    adicionar_paragrafo(
        doc,
        f'Lista todos os chamados de {F_PROJ} do período: em aberto no dia 01 ou '
        f'abertos no mês. Projetos longos permanecem visíveis enquanto ativos.',
        italic=True)
    cab, linhas = tabela_projetos(dados["tabela_projetos"])
    adicionar_tabela(doc, cab, linhas, [1.2, 5.0, 2.0, 2.0, 1.7, 1.7, 1.3, 1.9])

    # ===================== 4. SLA Institucional =====================
    adicionar_titulo(doc, f"4. SLA Institucional — Frente {F_SUST}")
    adicionar_paragrafo(doc, TEXTO_SLA, italic=True)
    cab, linhas = tabela_sla(dados["sla_sust"])
    adicionar_tabela(doc, cab, linhas, [1.4, 3.0, 2.4, 2.0, 1.6, 1.8, 1.8, 2.0])
    for nota in notas_sla(dados):
        adicionar_nota(doc, nota)
    if texto_outlier_sla(dados):
        adicionar_nota(doc, texto_outlier_sla(dados))

    # ===================== 5. Capacidade =====================
    adicionar_titulo(doc, f"5. Capacidade da Equipe — Frente {F_SUST}")
    adicionar_paragrafo(doc, texto_premissas(dados["cap_sust"]), italic=True)
    cab, linhas = tabela_capacidade(dados["cap_sust"], F_SUST)
    adicionar_tabela(doc, cab, linhas, [2, 9.5, 3])
    for nota in NOTAS_CAPACIDADE:
        adicionar_nota(doc, nota)

    # ===================== 6. Série Histórica =====================
    adicionar_titulo(doc, "6. Série Histórica Gerencial")
    if len(visao["serie"]) < 2:
        adicionar_nota(doc, TEXTO_SEM_SERIE)
    else:
        adicionar_paragrafo(doc, f"Série mensal da frente {F_SUST}.", italic=True)
        cab, linhas = tabela_serie(dados)
        adicionar_tabela(doc, cab, linhas, [1.9, 1.9, 2.4, 2.3, 2.0, 2.4, 3.6])
        adicionar_nota(doc, TEXTO_SERIE)
        for chave in ("hist_fluxo", "hist_backlog_frente", "hist_tmr", "hist_taxa"):
            adicionar_imagem(doc, imagens.get(chave), 15)

    # ===================== 7. Observações =====================
    adicionar_titulo(doc, "7. Observações e Plano de Ação")
    adicionar_paragrafo(doc, "Espaço para o autor completar com as ações do mês.",
                        italic=True)
    for item in ["Ação 1: ", "Ação 2: ", "Ação 3: "]:
        adicionar_bullet(doc, item)

    doc.save(destino)


# =============================================================================
# 7. ORQUESTRAÇÃO
# =============================================================================
_LOCK_GRAFICOS = threading.Lock()  # pyplot não é thread-safe (Flask atende em threads)


FONTE_PLANILHA = "planilha"
FONTE_REGISTRADO = "histórico registrado"


def _blocos_do_mes(m: pd.DataFrame, ano: int, mes: int) -> dict:
    """Métricas de um mês já recortado (base do relatório e da série)."""
    F_SUST = CONFIG["nome_frente_sustentacao"]
    F_PROJ = CONFIG["nome_frente_projetos"]
    b = dict(
        painel_sust=metricas_painel(m, F_SUST),
        painel_proj=metricas_painel(m, F_PROJ),
        cat_sust_full=metricas_categoria(m, F_SUST),
        cat_proj_full=metricas_categoria(m, F_PROJ),
        outliers_sust=detectar_outliers(m, F_SUST),
        cap_sust=metricas_capacidade(m, F_SUST),
        cap_proj=metricas_capacidade(m, F_PROJ),
        sla_sust=metricas_sla(m, F_SUST),
    )
    b["registro"] = _metricas_historicas(
        f"{ano:04d}-{mes:02d}", periodo_texto(ano, mes),
        b["painel_sust"], b["painel_proj"], b["cap_sust"], b["cap_proj"],
        _status_breakdown(m), b["outliers_sust"], b["cat_sust_full"], b["cat_proj_full"])
    b["registro"]["sustentacao"]["sla_pct_no_prazo"] = pct_no_prazo(b["sla_sust"])
    b["registro"]["backlog_inicial_incompleto"] = bool(
        m.attrs.get("backlog_inicial_incompleto", False))
    return b


def serie_da_planilha(base: pd.DataFrame, ate_periodo: str,
                      agora: Optional[pd.Timestamp] = None) -> list:
    """Registro mensal de cada mês da planilha ANTERIOR a ``ate_periodo``.

    Cada mês é recortado da mesma base e calculado com a mesma regra do
    relatório, então a comparação mês a mês não depende de nada gravado.
    """
    serie = []
    for periodo in metricas.meses_disponiveis(base)["mes"]:
        if periodo >= ate_periodo:
            continue
        ano, mes = (int(p) for p in periodo.split("-"))
        sub = filtrar_periodo_referencia(base, ano, mes, agora=agora)
        if sub.empty:
            continue
        registro = _blocos_do_mes(sub, ano, mes)["registro"]
        registro["fonte"] = FONTE_PLANILHA
        serie.append(registro)
    return serie


def calcular_relatorio(m: pd.DataFrame, ano: int, mes: int,
                       historico: Optional[tuple] = None,
                       base: Optional[pd.DataFrame] = None) -> dict:
    """Calcula TUDO o que o relatório mostra, sem gravar nada em disco.

    ``m`` é o resultado de ``filtrar_periodo_referencia``. ``base`` é a
    planilha inteira já preparada: os meses anteriores presentes nela entram
    na série/comparação automaticamente. ``historico`` é o par (registros,
    ignorados) de ``carregar_historico``; se omitido, é LIDO da pasta
    configurada. Para um mesmo mês, o registrado no histórico prevalece
    sobre o calculado da planilha (são os números já apresentados).
    """
    if {"novo_no_mes", "entregue_no_mes"} - set(m.columns):
        raise ValueError("DataFrame sem flags do período. Execute "
                         "filtrar_periodo_referencia() antes de calcular o relatório.")
    if m.empty:
        raise ValueError(f"Nenhum chamado relevante para {ano:04d}-{mes:02d}.")

    F_SUST = CONFIG["nome_frente_sustentacao"]
    F_PROJ = CONFIG["nome_frente_projetos"]
    periodo = f"{ano:04d}-{mes:02d}"
    txt_periodo = periodo_texto(ano, mes)

    blocos = _blocos_do_mes(m, ano, mes)
    painel_sust, painel_proj = blocos["painel_sust"], blocos["painel_proj"]
    cat_sust_full, cat_proj_full = blocos["cat_sust_full"], blocos["cat_proj_full"]
    outliers_sust = blocos["outliers_sust"]
    cap_sust, cap_proj = blocos["cap_sust"], blocos["cap_proj"]
    historico_atual = blocos["registro"]
    historico_atual["fonte"] = FONTE_PLANILHA

    registros, ignorados = historico if historico is not None else carregar_historico()
    da_planilha = ({r["periodo"]: r for r in serie_da_planilha(base, periodo,
                                                               agora=m.attrs.get("corte"))}
                   if base is not None else {})
    registrados = {r["periodo"]: {**r, "fonte": FONTE_REGISTRADO}
                   for r in registros if r.get("periodo", "") < periodo}
    anteriores = sorted({**da_planilha, **registrados}.values(),
                        key=lambda r: r["periodo"])
    visao = preparar_visao_historica(anteriores, historico_atual)

    corte, fim = m.attrs.get("corte"), m.attrs.get("periodo_fim")
    return dict(
        F_SUST=F_SUST, F_PROJ=F_PROJ,
        mes_referencia=periodo, periodo_texto=txt_periodo,
        painel_sust=painel_sust, painel_proj=painel_proj,
        cat_sust=compactar_categorias(cat_sust_full, CONFIG["max_categorias_detalhadas"]),
        cat_proj=compactar_categorias(cat_proj_full, CONFIG["max_categorias_detalhadas"]),
        sla_sust=blocos["sla_sust"], outliers_sust=outliers_sust,
        cap_sust=cap_sust, cap_proj=cap_proj,
        tabela_projetos=tabela_frente_secundaria(m, F_PROJ),
        historico_atual=historico_atual, visao_historica=visao,
        historico_ignorado=[p for p in ignorados
                            if p < periodo and p not in da_planilha],
        backlog_inicial_incompleto=m.attrs.get("backlog_inicial_incompleto", False),
        mes_parcial=bool(corte is not None and fim is not None and corte < fim),
        corte=corte,
    )


def _png(desenhar) -> Optional[bytes]:
    """Executa uma função de gráfico apontando para um buffer em memória."""
    buffer = io.BytesIO()
    desenhar(buffer)
    plt.close("all")
    return buffer.getvalue() or None


def gerar_graficos(dados: dict) -> dict:
    """Gera os PNGs do relatório em memória: {nome: bytes | None}."""
    F_SUST = dados["F_SUST"]
    serie = dados["visao_historica"]["serie"]
    with _LOCK_GRAFICOS:
        imagens = {
            "donut": _png(lambda b: grafico_donut_natureza(dados["painel_sust"], b)),
            "barras_sust": _png(lambda b: grafico_barras_categoria(
                dados["cat_sust"],
                f"{F_SUST} — Volume por Categoria "
                f"({dados['painel_sust']['total_relevante']} chamados)", b)),
        }
        if len(serie) >= 2:
            imagens.update({
                "hist_fluxo": _png(lambda b: grafico_fluxo_historico(serie, b)),
                "hist_backlog_frente": _png(lambda b: grafico_backlog_frente(serie, b)),
                "hist_tmr": _png(lambda b: grafico_evolucao_historica(
                    serie, "sustentacao", "tmr_mediana_h",
                    "Sustentação — Mediana do Tempo de Resolução (horas corridas)",
                    b, "h", cor_nome="navy")),
                "hist_taxa": _png(lambda b: grafico_evolucao_historica(
                    serie, "sustentacao", "taxa_resolucao",
                    "Sustentação — Taxa de Absorção", b, "%", cor_nome="green")),
            })
    return imagens


def gerar_docx_bytes(dados: dict, imagens: Optional[dict] = None) -> bytes:
    """Devolve o .docx pronto em memória (nada é gravado em disco)."""
    if imagens is None:
        imagens = gerar_graficos(dados)
    buffer = io.BytesIO()
    montar_documento(dados, imagens, buffer)
    return buffer.getvalue()


def gerar_relatorio_dataframe(m: pd.DataFrame, ano: int, mes: int,
                              base: Optional[pd.DataFrame] = None):
    """Uso em linha de comando: grava o .docx e registra o mês no histórico."""
    logger.info("[1/3] Calculando métricas de %04d-%02d ...", ano, mes)
    dados = calcular_relatorio(m, ano, mes, base=base)
    logger.info("[2/3] Gerando gráficos e documento em memória ...")
    conteudo = gerar_docx_bytes(dados)
    pasta = _nova_pasta_execucao()
    caminho_docx = os.path.join(pasta, CONFIG["nome_docx"])
    with open(caminho_docx, "wb") as arquivo:
        arquivo.write(conteudo)
    salvar_historico(dados["historico_atual"])
    logger.info("[3/3] Relatório salvo em: %s", caminho_docx)
    return caminho_docx


def gerar_relatorio(caminho_entrada: str, mes_referencia: Optional[str] = None):
    ano, mes = _resolver_mes_referencia(mes_referencia)
    bruto = carregar_dados(caminho_entrada)
    m = preparar_dataframe(bruto)
    m_filtrado = filtrar_periodo_referencia(m, ano, mes)
    if m_filtrado.empty:
        raise RuntimeError(
            f"Nenhum chamado relevante para {ano}-{mes:02d} em {caminho_entrada}.")
    return gerar_relatorio_dataframe(m_filtrado, ano, mes, base=m)


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
    carregar_config()
    if not os.path.exists(args.entrada):
        logger.error("Arquivo não encontrado: %s", args.entrada)
        sys.exit(1)
    try:
        gerar_relatorio(args.entrada, mes_referencia=args.mes)
    except Exception as exc:
        logger.exception("Falha ao gerar relatório: %s", exc)
        sys.exit(2)