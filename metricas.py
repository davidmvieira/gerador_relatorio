"""Núcleo de métricas do relatório GLPI — cálculo point-in-time por mês.

Regras (validadas com a liderança, ver base de conhecimento do projeto):

* Toda métrica de um mês é reconstruída pelas DATAS, nunca pelo status atual
  do chamado no momento do export. Assim, o mesmo arquivo gera sempre os
  mesmos números para o mesmo mês, e um export com vários meses pode ser
  fatiado mês a mês sem um mês "vazar" para o outro.
* Backlog em uma data D = aberto antes de D e (sem solução ou solução >= D).
* Absorção = solucionados ÷ (backlog inicial + abertos). Nunca passa de 100%.
* Tempo de resolução = solução − abertura, em horas CORRIDAS (calendário).
* As frentes (Sustentação × Projetos) são calculadas separadamente; este
  módulo não oferece agregado que misture as duas taxas.

O módulo é puro (só pandas/numpy): não lê arquivos, não grava nada e não
depende do Word nem do Flask — é a fonte única de verdade para a prévia e
para o documento exportado.
"""

from __future__ import annotations

from typing import Optional

import numpy as np
import pandas as pd

STATUS_CONCLUIDO = (5, 6)  # 5 = Solucionado, 6 = Fechado
TIPO_INCIDENTE = 1
TIPO_REQUISICAO = 2
OUTLIER_IQR_MULT_PADRAO = 8.0  # calibrado empiricamente para este dataset


# ---------------------------------------------------------------------------
# Datas e normalização
# ---------------------------------------------------------------------------
def limites_mes(ano: int, mes: int) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Retorna [início, fim) do mês: fim é o 1º instante do mês seguinte."""
    inicio = pd.Timestamp(year=ano, month=mes, day=1)
    return inicio, inicio + pd.offsets.MonthBegin(1)


def _abertura(df: pd.DataFrame) -> pd.Series:
    col = df["data_abertura"] if "data_abertura" in df.columns else df["date"]
    abertura = pd.to_datetime(col, errors="coerce")
    if "date" in df.columns:
        abertura = abertura.fillna(pd.to_datetime(df["date"], errors="coerce"))
    return abertura


def solucao_efetiva(df: pd.DataFrame) -> pd.Series:
    """Data de solução considerada nas métricas.

    * Chamado com status em aberto (1–4) não conta como solucionado, mesmo que
      a planilha traga uma data de solução (típico de chamado reaberto).
    * Chamado concluído (5/6) sem data de solução usa a data de fechamento,
      quando a coluna existir.
    * Status desconhecido/ausente: confia na data informada.
    """
    solucao = pd.to_datetime(df.get("data_solucao"), errors="coerce")
    if solucao is None or not isinstance(solucao, pd.Series):
        solucao = pd.Series(pd.NaT, index=df.index)
    if "data_fechamento" in df.columns:
        fechamento = pd.to_datetime(df["data_fechamento"], errors="coerce")
    else:
        fechamento = pd.Series(pd.NaT, index=df.index)

    status = pd.to_numeric(df["status"], errors="coerce") if "status" in df.columns \
        else pd.Series(np.nan, index=df.index)
    concluido = status.isin(STATUS_CONCLUIDO)
    em_aberto = status.notna() & ~concluido

    efetiva = solucao.where(~concluido | solucao.notna(), fechamento)
    return efetiva.where(~em_aberto, pd.NaT)


def inconsistencias(df: pd.DataFrame) -> pd.DataFrame:
    """Lista chamados cujos dados impedem um cálculo 100% fiel.

    Colunas: id, problema. Não altera nada — serve para a tela de revisão.
    """
    abertura = _abertura(df)
    solucao = pd.to_datetime(df.get("data_solucao"), errors="coerce")
    if not isinstance(solucao, pd.Series):
        solucao = pd.Series(pd.NaT, index=df.index)
    efetiva = solucao_efetiva(df)
    status = pd.to_numeric(df.get("status"), errors="coerce") if "status" in df.columns \
        else pd.Series(np.nan, index=df.index)
    ids = df["id"] if "id" in df.columns else pd.Series(df.index, index=df.index)

    regras = [
        (abertura.isna(), "data de abertura ausente ou inválida"),
        (efetiva.notna() & abertura.notna() & (efetiva < abertura),
         "data de solução anterior à abertura"),
        (status.isin(STATUS_CONCLUIDO) & efetiva.isna(),
         "concluído sem data de solução/fechamento (fica no backlog)"),
        (status.notna() & ~status.isin(STATUS_CONCLUIDO) & solucao.notna(),
         "status em aberto com data de solução (reaberto?) — tratado como aberto"),
    ]
    linhas = []
    for mascara, problema in regras:
        for valor in ids[mascara.fillna(False)]:
            linhas.append({"id": valor, "problema": problema})
    return pd.DataFrame(linhas, columns=["id", "problema"])


# ---------------------------------------------------------------------------
# Meses presentes na planilha
# ---------------------------------------------------------------------------
def meses_disponiveis(df: pd.DataFrame) -> pd.DataFrame:
    """Meses cobertos pela base, com abertos e solucionados em cada um.

    Retorna DataFrame ordenado com colunas: mes ('YYYY-MM'), abertos,
    solucionados.
    """
    abertura = _abertura(df).dropna().dt.to_period("M").value_counts()
    solucao = solucao_efetiva(df).dropna().dt.to_period("M").value_counts()
    periodos = sorted(set(abertura.index) | set(solucao.index))
    return pd.DataFrame({
        "mes": [str(p) for p in periodos],
        "abertos": [int(abertura.get(p, 0)) for p in periodos],
        "solucionados": [int(solucao.get(p, 0)) for p in periodos],
    }, columns=["mes", "abertos", "solucionados"])


# ---------------------------------------------------------------------------
# Recorte do mês
# ---------------------------------------------------------------------------
def marcar_periodo(df: pd.DataFrame, ano: int, mes: int,
                   agora: Optional[pd.Timestamp] = None) -> pd.DataFrame:
    """Adiciona as flags point-in-time do mês a TODAS as linhas de ``df``.

    Flags/colunas adicionadas:
      aberto_no_mes        aberto dentro de [início, fim)
      solucionado_no_mes   solução efetiva dentro de [início, fim)
      backlog_inicio       em aberto no instante 'início' (dia 01, 00:00)
      backlog_fim          em aberto no instante 'fim' (1º dia do mês seguinte)
      relevante            tocou o mês: aberto_no_mes | backlog_inicio
                           (todo solucionado no mês satisfaz um dos dois)
      tempo_resolucao_h    horas corridas, só para solucionados no mês
      horas_em_aberto_h    horas corridas até o corte, só para backlog_fim
      idade_dias           dias corridos até o corte (solução, se houve no mês)

    O corte é o fim do mês ou ``agora``, o que vier antes (mês corrente).
    """
    inicio, fim = limites_mes(ano, mes)
    agora = pd.Timestamp.now() if agora is None else pd.Timestamp(agora)
    corte = min(fim, agora)

    m = df.copy()
    abertura = _abertura(m)
    solucao = solucao_efetiva(m)

    m["data_solucao_efetiva"] = solucao
    m["aberto_no_mes"] = (abertura >= inicio) & (abertura < fim)
    m["solucionado_no_mes"] = (solucao >= inicio) & (solucao < fim)
    m["backlog_inicio"] = (abertura < inicio) & (solucao.isna() | (solucao >= inicio))
    m["backlog_fim"] = (abertura < fim) & (solucao.isna() | (solucao >= fim))
    m["relevante"] = m["aberto_no_mes"] | m["backlog_inicio"]

    horas_resolucao = (solucao - abertura).dt.total_seconds() / 3600
    horas_abertas = (corte - abertura).dt.total_seconds() / 3600
    m["tempo_resolucao_h"] = horas_resolucao.where(m["solucionado_no_mes"])
    m["horas_em_aberto_h"] = horas_abertas.where(m["backlog_fim"])
    referencia = solucao.where(m["solucionado_no_mes"], corte)
    m["idade_dias"] = ((referencia - abertura).dt.total_seconds() / 86400).round(0)

    m.attrs["periodo_inicio"] = inicio
    m.attrs["periodo_fim"] = fim
    m.attrs["corte"] = corte
    # Se a base não tem nenhum chamado aberto antes do mês, o backlog inicial
    # só pode ser confiável se o export realmente incluir chamados antigos.
    primeira = abertura.min()
    m.attrs["backlog_inicial_incompleto"] = bool(pd.isna(primeira) or primeira >= inicio)
    return m


# ---------------------------------------------------------------------------
# Indicadores
# ---------------------------------------------------------------------------
def estatisticas_tempo(horas: pd.Series,
                       iqr_mult: float = OUTLIER_IQR_MULT_PADRAO) -> dict:
    """Média, mediana e média sem outliers (regra Q3 + iqr_mult·IQR)."""
    h = pd.to_numeric(horas, errors="coerce").dropna()
    if h.empty:
        return {"n": 0, "media_h": None, "mediana_h": None,
                "media_sem_outliers_h": None, "n_outliers": 0, "limite_outlier_h": None}
    limite = None
    sem_outliers = h
    if len(h) >= 4:
        q1, q3 = h.quantile([0.25, 0.75])
        limite = float(q3 + iqr_mult * (q3 - q1))
        sem_outliers = h[h <= limite]
    return {
        "n": int(len(h)),
        "media_h": round(float(h.mean()), 1),
        "mediana_h": round(float(h.median()), 1),
        "media_sem_outliers_h": round(float(sem_outliers.mean()), 1),
        "n_outliers": int(len(h) - len(sem_outliers)),
        "limite_outlier_h": round(limite, 1) if limite is not None else None,
    }


def _pct(parte, base) -> Optional[float]:
    return round(parte / base * 100, 1) if base else None


def indicadores(m: pd.DataFrame, frente: Optional[str] = None,
                iqr_mult: float = OUTLIER_IQR_MULT_PADRAO) -> dict:
    """Indicadores do mês para uma frente (``m`` já passou por marcar_periodo).

    ``frente=None`` usa todas as linhas recebidas — o chamador é responsável
    por não misturar frentes quando isso distorcer a leitura.
    """
    sub = m if frente is None else m[m["frente"] == frente]

    backlog_inicial = int(sub["backlog_inicio"].sum())
    abertos = int(sub["aberto_no_mes"].sum())
    solucionados = int(sub["solucionado_no_mes"].sum())
    backlog_final = int(sub["backlog_fim"].sum())
    carga = backlog_inicial + abertos

    tempos = estatisticas_tempo(sub.loc[sub["solucionado_no_mes"], "tempo_resolucao_h"],
                                iqr_mult)

    tipos = pd.to_numeric(sub.loc[sub["aberto_no_mes"], "type"], errors="coerce") \
        if "type" in sub.columns else pd.Series(dtype=float)
    req = int((tipos == TIPO_REQUISICAO).sum())
    inc = int((tipos == TIPO_INCIDENTE).sum())

    return {
        "backlog_inicial": backlog_inicial,
        "abertos": abertos,
        "solucionados": solucionados,
        "backlog_final": backlog_final,
        "saldo": abertos - solucionados,
        "carga_total": carga,
        "absorcao_pct": _pct(solucionados, carga),
        "variacao_backlog_pct": _pct(backlog_final - backlog_inicial, backlog_inicial),
        "fluxo_confere": backlog_inicial + abertos - solucionados == backlog_final,
        "tmr_media_h": tempos["media_h"],
        "tmr_mediana_h": tempos["mediana_h"],
        "tmr_media_sem_outliers_h": tempos["media_sem_outliers_h"],
        "n_outliers": tempos["n_outliers"],
        "limite_outlier_h": tempos["limite_outlier_h"],
        "requisicoes": req,
        "incidentes": inc,
        "outros_tipos": abertos - req - inc,
        "pct_requisicoes": _pct(req, abertos),
        "pct_incidentes": _pct(inc, abertos),
    }
