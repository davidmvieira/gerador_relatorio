#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_relatorio_glpi.py
========================
Gera, de forma autônoma, o relatório mensal da Equipe de Sistemas a partir de
um export do GLPI em Excel (uma ou duas abas), separando as demandas em duas frentes
(Sustentação x Projetos) e produzindo um .docx pronto para leitura da liderança.

COMO USAR
---------
1. Instale as dependências:
       pip install pandas openpyxl matplotlib python-docx --break-system-packages

2. Rode o script apontando para o export do mês:
    python3 gerar_relatorio_glpi.py caminho/para/base_de_dados.xlsx
    python3 gerar_relatorio_glpi.py caminho/para/base_de_dados.csv

   Se nenhum caminho for passado, o script procura por "base_de_dados.xlsx"
   na pasta atual.

3. O relatório final e os gráficos ficam em ./saida_relatorio/

FORMATO ESPERADO DO EXCEL (duas abas, mesmo padrão usado até hoje)
-------------------------------------------------------------------
Aba 1 (dados "limpos", fonte da verdade para id/categoria/status/prioridade):
    id, name, date, itilcategories_id, categoria, demanda, status,
    priority, type, entities_id, entidade

Aba 2 (detalhe com datas de solução/fechamento e tempo de resolução):
    id_chamado, titulo, entities_id, entidade, categoria, tipo_chamado,
    prioridade, status_atual, data_abertura, data_solucao, data_fechamento,
    tempo_resolucao_horas

O nome exato das abas é configurável em CONFIG['aba_principal'] / ['aba_detalhe'].

O QUE O SCRIPT FAZ
------------------
1. Lê uma ou duas abas e cruza as duas fontes pelo ID do chamado quando necessário.
2. Detecta e tenta corrigir automaticamente linhas com colunas deslocadas
   (comum quando o título do chamado tem um caractere especial embutido).
3. Separa os chamados em duas frentes:
       - Sustentação: entidade configurada em CONFIG['entidade_sustentacao']
       - Projetos: todas as demais entidades (Projetos, Informação, etc.)
4. Consolida categorias com grafias diferentes (ex.: "Intranet"/"intranet").
5. Calcula: painel geral, distribuição por categoria, SLA por prioridade,
   capacidade/aderência da equipe, outliers de tempo de resolução,
   incidência de termos-chave (ex.: problemas recorrentes de um sistema)
   e a tabela detalhada da frente Projetos.
6. Consolida métricas mensais em `relatorios_glpi/historico/YYYY-MM.json`.
7. Compara o mês atual com o histórico, mostra tendências e calcula projeção
    por média móvel simples quando há dados suficientes.
8. Gera os gráficos (matplotlib) e monta o .docx final (python-docx).

Ajuste a seção CONFIG abaixo a cada mês / a cada mudança de contexto.
"""

import os
import re
import sys
import json
import unicodedata
from datetime import datetime, timedelta
from calendar import monthrange

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================================
# CONFIG — ajuste aqui a cada rodada
# ============================================================================
CONFIG = {
    # arquivo de entrada (pode ser sobrescrito via argumento de linha de comando)
    "arquivo_entrada": "base_de_dados.xlsx",
    "aba_principal": "_WITH_RECURSIVE_entidades_AS_SE",
    "aba_detalhe": "Página1",

    # como dividir as frentes
    "entidade_sustentacao": "Sistemas",
    "nome_frente_sustentacao": "Sustentação",
    "nome_frente_projetos": "Projetos",

    # cabeçalho do relatório
    "titulo_relatorio": "Equipe de Sistemas — Sustentação & Operações",
    "subtitulo": "Relatório Mensal · Entidade Sistemas (14) e sub-entidades",
    "autor": "David Vieira",

    # categorias que ficam destacadas na tabela (as demais viram "Outros")
    "max_categorias_detalhadas": 10,

    # termos-chave para detectar um problema recorrente pelo título do chamado
    # (edite a cada mês conforme o que estiver mais quente operacionalmente)
    "termos_chave_recorrencia": [r"censo", r"leito", r"alta manti", r"pool d[eo] pacient", r"fora do censo"],
    "nome_recorrencia": "Censo/Pool de Pacientes (AGHU)",

    # premissas de capacidade da equipe (ajuste todo mês)
    "premissas_equipe": {
        "analistas": 6,
        "horas_por_dia": 8,
        "dias_uteis_mes": 20,
        "dias_ferias_no_mes": 11,   # dias úteis perdidos por férias de 1 analista
    },

    # limiar de outlier: chamados com tempo de resolução acima de
    # Q3 + OUTLIER_IQR_MULT * IQR são reportados individualmente
    # (valor alto de propósito: a ideia é pegar só os casos realmente
    # extremos, tipo "ficou 27 dias parado", não qualquer chamado acima da média)
    "outlier_iqr_mult": 8.0,

    # pasta e nome de saída
    "pasta_saida": os.path.join("relatorios_glpi", "execucoes"),
    "nome_docx": "Relatorio_Sistemas.docx",
    "pasta_historico": os.path.join("relatorios_glpi", "historico"),
}

# Paleta de cores (hex, sem #) usada em tabelas e gráficos
COR = {
    "navy": "1F3864",
    "navy_rgb": (0x1F, 0x38, 0x64),
    "orange": "C55A11",
    "orange_rgb": (0xC5, 0x5A, 0x11),
    "red": "C00000",
    "red_rgb": (0xC0, 0x00, 0x00),
    "green": "375623",
    "green_rgb": (0x37, 0x56, 0x23),
    "green_bg": "E2EFDA",
    "amber_bg": "FCE4D6",
    "gray_bg": "F2F2F2",
    "banner_blue_bg": "DCE6F1",
    "banner_orange_bg": "FBE5D6",
    "lightgray": "#D9D9D9",
}


# ============================================================================
# 1. CARGA E LIMPEZA DOS DADOS
# ============================================================================

def _ler_entrada(caminho_entrada: str):
    extensao = os.path.splitext(caminho_entrada)[1].lower()
    if extensao == ".csv":
        try:
            return pd.read_csv(caminho_entrada, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(caminho_entrada, sep=None, engine="python", encoding="latin-1")
    if extensao in (".xlsx", ".xlsm"):
        return pd.ExcelFile(caminho_entrada)
    raise ValueError("Formato não suportado. Use um arquivo .xlsx, .xlsm ou .csv.")


def carregar_dados(caminho_entrada: str) -> pd.DataFrame:
    """Lê CSV ou Excel em uma ou duas abas e devolve dados limpos."""
    entrada = _ler_entrada(caminho_entrada)
    eh_csv = isinstance(entrada, pd.DataFrame)
    abas = [None] if eh_csv else entrada.sheet_names

    if len(abas) == 1:
        df = entrada if eh_csv else pd.read_excel(caminho_entrada, sheet_name=abas[0])
        df = df[df["id_chamado"].notna()].copy()
        colunas_renomeadas = {
            "id_chamado": "id",
            "titulo": "name",
            "data_abertura": "date",
            "status_atual": "status",
            "prioridade": "priority",
        }
        df = df.rename(columns=colunas_renomeadas)

        status_map = {
            "novo": 1,
            "processando": 2,
            "pendente": 3,
            "planejado": 4,
            "solucionado": 5,
            "fechado": 6,
        }
        prioridade_map = {
            "muito baixa": 1,
            "baixa": 2,
            "média": 3,
            "media": 3,
            "alta": 4,
            "muito alta": 5,
            "crítica": 6,
            "critica": 6,
        }
        df["status"] = df["status"].map(
            lambda valor: status_map.get(str(valor).strip().lower(), valor)
        )
        df["priority"] = df["priority"].map(
            lambda valor: prioridade_map.get(str(valor).strip().lower(), valor)
        )
        if "categoria.1" in df.columns:
            df["categoria"] = df["categoria"].fillna(df["categoria.1"])
        df["categoria"] = df["categoria"].fillna("Sem categoria")
        tipo_map = {"incidente": 1, "requisição": 2, "requisicao": 2}
        df["type"] = df["tipo_chamado"].map(
            lambda valor: tipo_map.get(str(valor).strip().lower(), valor)
        )
        df["data_abertura"] = df["date"]
        m = df
    else:
        df1 = pd.read_excel(caminho_entrada, sheet_name=CONFIG["aba_principal"])
        df2 = pd.read_excel(caminho_entrada, sheet_name=CONFIG["aba_detalhe"])

        # remove linha de rodapé/resumo (sem id) que alguns exports trazem no final
        df2 = df2[df2["id_chamado"].notna()].copy()
        df2["id_chamado"] = df2["id_chamado"].astype(int)

        m = df1.merge(df2, left_on="id", right_on="id_chamado", how="left", suffixes=("", "_s2"))
    m["date"] = pd.to_datetime(m["date"], errors="coerce")

    m = _reparar_linhas_deslocadas(m)

    m["data_abertura"] = pd.to_datetime(m["data_abertura"], errors="coerce")
    m["data_solucao"] = pd.to_datetime(m["data_solucao"], errors="coerce")

    return m


_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}")


def _reparar_linhas_deslocadas(m: pd.DataFrame) -> pd.DataFrame:
    """
    Detecta linhas cuja coluna 'data_abertura' não é uma data válida
    (sintoma de deslocamento de colunas, geralmente causado por um caractere
    especial embutido no título do chamado) e tenta reconstruir as datas
    varrendo todas as colunas textuais da linha em busca de timestamps.

    A data de abertura confiável é sempre a coluna 'date' da aba principal
    (aba 1), que não sofre esse tipo de deslocamento.
    """
    problema = pd.to_datetime(m["data_abertura"], errors="coerce").isna() & m["data_abertura"].notna()
    ids_com_problema = m.loc[problema, "id"].tolist()

    # garante dtype 'object' para podermos atribuir Timestamps sem erro de tipo
    m["data_abertura"] = m["data_abertura"].astype(object)
    m["data_solucao"] = m["data_solucao"].astype(object)

    for idx in m.index[problema]:
        linha = m.loc[idx]
        candidatos = []
        for col in m.columns:
            val = linha[col]
            if isinstance(val, str):
                achado = _DATE_RE.search(val)
                if achado:
                    candidatos.append(achado.group(0))
        candidatos = sorted(set(candidatos))
        if len(candidatos) >= 1:
            abertura_real = linha["date"]
            # entre os candidatos textuais, pega o mais próximo/posterior à abertura
            # como possível data de solução (o que faz sentido para o modelo de dados)
            candidatos_dt = [pd.to_datetime(c) for c in candidatos]
            posteriores = [c for c in candidatos_dt if c > abertura_real]
            solucao_real = min(posteriores) if posteriores else max(candidatos_dt)
            m.at[idx, "data_abertura"] = abertura_real
            # só marca como solucionado se o status_atual/status numérico indicar isso
            if str(linha.get("status_atual", "")).strip().lower() in ("solucionado", "fechado") or linha.get("status") in (5, 6):
                m.at[idx, "data_solucao"] = solucao_real

    if ids_com_problema:
        print(f"[aviso] {len(ids_com_problema)} linha(s) com colunas deslocadas foram "
              f"reparadas automaticamente: {ids_com_problema}")

    return m


def canonicalizar_categoria(cat: str) -> str:
    """Normaliza grafias diferentes da mesma categoria (case/espaço)."""
    return cat.strip()


def consolidar_categorias(df: pd.DataFrame) -> pd.DataFrame:
    """Agrupa categorias que só diferem por maiúsculas/minúsculas/espaços,
    usando a grafia mais frequente como rótulo final."""
    chave = df["categoria"].astype(str).str.strip().str.lower()
    moda_por_chave = df.groupby(chave)["categoria"].agg(lambda s: s.value_counts().idxmax())
    df = df.copy()
    df["categoria_final"] = chave.map(moda_por_chave)
    return df


def preparar_dataframe_df(m: pd.DataFrame) -> pd.DataFrame:
    m = m.copy()
    m = consolidar_categorias(m)

    status_map = {1: "Novo", 2: "Processando", 3: "Pendente", 4: "Planejado", 5: "Solucionado", 6: "Fechado"}
    prioridade_map = {1: "Muito Baixa", 2: "Baixa", 3: "Média", 4: "Alta", 5: "Muito Alta", 6: "Crítica"}
    m["status_nome"] = m["status"].map(status_map)
    m["prioridade_nome"] = m["priority"].map(prioridade_map)

    m["frente"] = np.where(
        m["entidade"] == CONFIG["entidade_sustentacao"],
        CONFIG["nome_frente_sustentacao"],
        CONFIG["nome_frente_projetos"],
    )

    concluido = m["status"].isin([5, 6])
    agora = pd.Timestamp(datetime.now())

    tempo_concluido = (m["data_solucao"] - m["data_abertura"]).dt.total_seconds() / 3600
    tempo_acumulado = (agora - m["data_abertura"]).dt.total_seconds() / 3600

    m["tempo_h"] = np.where(concluido, tempo_concluido, np.nan)
    m["horas_acumuladas"] = np.where(~concluido, tempo_acumulado, np.nan)
    m["concluido"] = concluido

    return m


def preparar_dataframe(caminho_entrada: str) -> pd.DataFrame:
    return preparar_dataframe_df(carregar_dados(caminho_entrada))


def _numero(valor, casas=None):
    """Converte escalares pandas/NumPy para tipos simples serializáveis."""
    if valor is None or pd.isna(valor):
        return None
    resultado = float(valor)
    return round(resultado, casas) if casas is not None else resultado


def _metricas_historicas(periodo_texto, painel_total, painel_sust, painel_proj,
                         cap_sust, cap_proj, cat_sust_full, cat_proj_full,
                         outliers_sust, recorrencia):
    def painel_resumido(painel):
        return {
            "total": int(painel["total"]),
            "resolvidos": int(painel["entregues"]),
            "taxa_resolucao": _numero(painel["taxa_resolucao"], 1),
            "tmr_h": _numero(painel["tmr_h"], 1),
            "solicitacoes": int(painel["req"]),
            "incidentes": int(painel["inc"]),
        }

    def capacidade_resumida(capacidade):
        return {
            "hmm": _numero(capacidade["hmm"], 1),
            "he": _numero(capacidade["he"], 1),
            "hpc": _numero(capacidade["hpc"], 1),
            "hha": _numero(capacidade["hha"], 1),
            "aderencia_pct": _numero(capacidade["ad_pct"], 1),
        }

    def categorias_resumidas(categorias):
        resultado = []
        for categoria, linha in categorias.head(CONFIG["max_categorias_detalhadas"]).iterrows():
            resultado.append({
                "categoria": str(categoria),
                "qtd": int(linha["qtd"]),
                "tempo_h": _numero(linha["tempo_h"], 1),
                "pct_qtd": _numero(linha["pct_qtd"], 1),
                "pct_tempo": _numero(linha["pct_tempo"], 1),
            })
        return resultado

    return {
        "periodo": None,
        "periodo_texto": periodo_texto,
        "total": painel_resumido(painel_total),
        "sustentacao": painel_resumido(painel_sust),
        "projetos": painel_resumido(painel_proj),
        "capacidade": {
            "sustentacao": capacidade_resumida(cap_sust),
            "projetos": capacidade_resumida(cap_proj),
        },
        "aderencia_pct": _numero(cap_sust["ad_pct"], 1),
        "outliers": {
            "quantidade": int(len(outliers_sust)),
            "maior_tempo_h": _numero(outliers_sust["tempo_h"].max(), 1)
            if len(outliers_sust) else None,
        },
        "recorrencias": {
            "nome": CONFIG["nome_recorrencia"],
            "quantidade": int(len(recorrencia)),
        },
        "categorias": {
            "sustentacao": categorias_resumidas(cat_sust_full),
            "projetos": categorias_resumidas(cat_proj_full),
        },
    }


def carregar_historico():
    historico = []
    pasta = CONFIG["pasta_historico"]
    if not os.path.isdir(pasta):
        return historico
    for nome in os.listdir(pasta):
        if not re.fullmatch(r"\d{4}-\d{2}\.json", nome):
            continue
        caminho = os.path.join(pasta, nome)
        try:
            with open(caminho, "r", encoding="utf-8") as arquivo:
                registro = json.load(arquivo)
            registro["periodo"] = nome[:-5]
            historico.append(registro)
        except (OSError, json.JSONDecodeError, TypeError) as erro:
            print(f"[aviso] Histórico ignorado ({nome}): {erro}")
    return sorted(historico, key=lambda item: item.get("periodo", ""))


def salvar_historico(registro):
    pasta = CONFIG["pasta_historico"]
    os.makedirs(pasta, exist_ok=True)
    caminho = os.path.join(pasta, f'{registro["periodo"]}.json')
    with open(caminho, "w", encoding="utf-8") as arquivo:
        json.dump(registro, arquivo, ensure_ascii=False, indent=2)


def comparar_metricas(atual, anterior):
    comparacoes = []
    indicadores = [
        ("Chamados", atual["total"]["total"], anterior["total"]["total"], ""),
        ("Resolvidos", atual["total"]["resolvidos"], anterior["total"]["resolvidos"], ""),
        ("Taxa de resolução", atual["total"]["taxa_resolucao"], anterior["total"]["taxa_resolucao"], "%"),
        ("TMR", atual["total"]["tmr_h"], anterior["total"]["tmr_h"], "h"),
        ("Sustentação", atual["sustentacao"]["total"], anterior["sustentacao"]["total"], ""),
        ("Projetos", atual["projetos"]["total"], anterior["projetos"]["total"], ""),
        ("Aderência", atual["aderencia_pct"], anterior["aderencia_pct"], "%"),
    ]
    for nome, valor_atual, valor_anterior, unidade in indicadores:
        if valor_atual is None or valor_anterior is None:
            continue
        diferenca = valor_atual - valor_anterior
        variacao = diferenca / valor_anterior * 100 if valor_anterior else None
        comparacoes.append({
            "indicador": nome,
            "atual": valor_atual,
            "anterior": valor_anterior,
            "diferenca": diferenca,
            "variacao_pct": variacao,
            "unidade": unidade,
        })
    return comparacoes


def preparar_visao_historica(historico, atual):
    serie = sorted(historico + [atual], key=lambda item: item["periodo"])
    anterior = next((item for item in reversed(historico)
                     if item["periodo"] < atual["periodo"]), None)
    tendencia = serie if len(serie) >= 3 else []
    projecao = {}
    if len(serie) >= 3:
        ultimos = serie[-3:]
        campos = [
            ("Chamados", ("total", "total")),
            ("Resolvidos", ("total", "resolvidos")),
            ("Taxa de resolução", ("total", "taxa_resolucao")),
            ("TMR", ("total", "tmr_h")),
            ("Sustentação", ("sustentacao", "total")),
            ("Projetos", ("projetos", "total")),
            ("Aderência", (None, "aderencia_pct")),
        ]
        for nome, (grupo, campo) in campos:
            valores = [item[campo] if grupo is None else item[grupo][campo]
                       for item in ultimos]
            valores = [valor for valor in valores if valor is not None]
            if valores:
                projecao[nome] = sum(valores) / len(valores)
    return {
        "anterior": anterior,
        "comparacoes": comparar_metricas(atual, anterior) if anterior else [],
        "tendencia": tendencia,
        "projecao": projecao,
    }


# ============================================================================
# 2. MÉTRICAS
# ============================================================================

def periodo_do_relatorio(m: pd.DataFrame):
    """Deduz o texto do período (mês/ano) a partir das datas de abertura."""
    dt_min, dt_max = m["date"].min(), m["date"].max()
    ano, mes = dt_min.year, dt_min.month
    ultimo_dia = monthrange(ano, mes)[1]
    meses_pt = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
                "agosto", "setembro", "outubro", "novembro", "dezembro"]
    texto = f"01 a {ultimo_dia:02d} de {meses_pt[mes-1].capitalize()} de {ano}"
    return texto, ano, mes, ultimo_dia


def buckets_semanais(m: pd.DataFrame, ano: int, mes: int, ultimo_dia: int):
    """Divide o mês em ~4 blocos semanais (o último pode ter >7 dias)."""
    inicio = pd.Timestamp(year=ano, month=mes, day=1)
    fim = pd.Timestamp(year=ano, month=mes, day=ultimo_dia) + pd.Timedelta(days=1)
    cortes = pd.date_range(inicio, fim, periods=5)
    labels = [f"Sem {i+1}\n({cortes[i].day:02d}–{(cortes[i+1]-pd.Timedelta(days=1)).day:02d}/{mes:02d})"
              for i in range(4)]
    bucket = pd.cut(m["date"], bins=cortes, labels=labels, right=False, include_lowest=True)
    return bucket, labels


def metricas_painel(m: pd.DataFrame, frente: str = None) -> dict:
    sub = m if frente is None else m[m["frente"] == frente]
    total = len(sub)
    entregues = sub[sub["concluido"]]
    taxa = round(len(entregues) / total * 100, 1) if total else 0.0
    tmr = round(entregues["tempo_h"].mean(), 1) if len(entregues) else None
    tipo_counts = sub["type"].value_counts()
    req = int(tipo_counts.get(2, 0))
    inc = int(tipo_counts.get(1, 0))
    pct_req = round(req / total * 100, 1) if total else 0
    pct_inc = round(inc / total * 100, 1) if total else 0
    return {
        "total": total, "entregues": len(entregues), "taxa_resolucao": taxa,
        "tmr_h": tmr, "req": req, "inc": inc, "pct_req": pct_req, "pct_inc": pct_inc,
    }


def metricas_categoria(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    sub = m[m["frente"] == frente]
    qtd = sub.groupby("categoria_final").size().rename("qtd")
    entregues = sub[sub["concluido"]]
    tempo = entregues.groupby("categoria_final")["tempo_h"].sum().round(1).rename("tempo_h")
    full = pd.concat([qtd, tempo], axis=1).fillna(0)
    full["qtd"] = full["qtd"].astype(int)
    total_qtd = full["qtd"].sum()
    total_tempo = full["tempo_h"].sum()
    full["pct_qtd"] = (full["qtd"] / total_qtd * 100).round(1) if total_qtd else 0
    full["pct_tempo"] = (full["tempo_h"] / total_tempo * 100).round(1) if total_tempo else 0
    return full.sort_values("qtd", ascending=False)


def compactar_categorias(cat_df: pd.DataFrame, top_n: int) -> pd.DataFrame:
    """Mantém as top_n categorias e agrupa o resto em 'Outros (...)'."""
    if len(cat_df) <= top_n:
        return cat_df
    top = cat_df.iloc[:top_n].copy()
    resto = cat_df.iloc[top_n:]
    outros = pd.DataFrame({
        "qtd": [resto["qtd"].sum()],
        "tempo_h": [resto["tempo_h"].sum()],
        "pct_qtd": [round(resto["pct_qtd"].sum(), 1)],
        "pct_tempo": [round(resto["pct_tempo"].sum(), 1)],
    }, index=[f"Outros ({len(resto)} categorias)"])
    return pd.concat([top, outros])


def metricas_sla_prioridade(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    sub = m[(m["frente"] == frente) & (m["concluido"])]
    ordem = ["Crítica", "Muito Alta", "Alta", "Média", "Baixa", "Muito Baixa"]
    g = sub.groupby("prioridade_nome")["tempo_h"].agg(["count", "mean", "min", "max"]).round(1)
    return g.reindex([o for o in ordem if o in g.index])


def detectar_outliers(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    """Outliers via regra do IQR sobre o tempo de resolução dos concluídos."""
    sub = m[(m["frente"] == frente) & (m["concluido"])].copy()
    if len(sub) < 4:
        return sub.iloc[0:0]
    q1, q3 = sub["tempo_h"].quantile([0.25, 0.75])
    iqr = q3 - q1
    limite = q3 + CONFIG["outlier_iqr_mult"] * iqr
    return sub[sub["tempo_h"] > limite].sort_values("tempo_h", ascending=False)


def metricas_capacidade(m: pd.DataFrame, frente: str) -> dict:
    p = CONFIG["premissas_equipe"]
    hmm = p["analistas"] * p["dias_uteis_mes"] * p["horas_por_dia"]
    # dias_ferias_no_mes = dias úteis PERDIDOS por férias/ausências (já em dias úteis)
    he = hmm - p["dias_ferias_no_mes"] * p["horas_por_dia"]
    sub = m[(m["frente"] == frente) & (m["concluido"])]
    hpc = round(sub["tempo_h"].sum(), 1)
    hha = round(hpc / p["horas_por_dia"], 1)
    ad_pct = round(hha / he * 100, 1) if he else 0
    return {"hmm": hmm, "he": he, "hpc": hpc, "hha": hha, "ad_pct": ad_pct}


def detectar_recorrencia(m: pd.DataFrame, termos: list) -> pd.DataFrame:
    padrao = "|".join(termos)
    kw = m["name"].str.lower().str.contains(padrao, na=False, regex=True)
    return m[kw].copy()


def limpar_titulo(titulo: str, max_len: int = 70) -> str:
    """Remove o padrão '- Nome Sobrenome - 1234 -' do fim do título (quando existir)
    para não expor nome de solicitante/técnico em tabelas gerenciais, e corta o
    tamanho para caber bem na tabela."""
    if not isinstance(titulo, str):
        return ""
    t = titulo.strip()
    padrao_nome = r"\s*-\s*[A-ZÀ-Ú][a-zà-úA-ZÀ-Ú\s]+-\s*\d+\s*-?\s*$"
    t = re.sub(padrao_nome, "", t).strip(" -\t")
    if len(t) > max_len:
        t = t[:max_len - 1].rstrip() + "…"
    return t


def tabela_frente_secundaria(m: pd.DataFrame, frente: str) -> pd.DataFrame:
    """Lista item a item os chamados de uma frente (pensado para a frente
    Projetos, tipicamente pequena o suficiente para não precisar agregação)."""
    sub = m[m["frente"] == frente].copy()
    sub["titulo_limpo"] = sub["name"].apply(limpar_titulo)
    sub = sub.sort_values("date")
    cols = ["id", "titulo_limpo", "status_nome", "date", "data_solucao", "tempo_h", "horas_acumuladas"]
    return sub[cols]


# ============================================================================
# 3. GRÁFICOS
# ============================================================================

def _preparar_pasta_saida(caminho_saida):
    os.makedirs(caminho_saida, exist_ok=True)


def _nova_pasta_execucao():
    base = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(CONFIG["pasta_saida"], base)
    contador = 1
    while os.path.exists(caminho):
        caminho = os.path.join(CONFIG["pasta_saida"], f"{base}_{contador:02d}")
        contador += 1
    return caminho


def grafico_donut_natureza(painel: dict, caminho: str):
    fig, ax = plt.subplots(figsize=(6, 6), dpi=150)
    sizes = [painel["req"], painel["inc"]]
    labels = ["Requisições", "Incidentes"]
    colors = ["#548235", "#" + COR["red"]]
    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors, startangle=90, pctdistance=0.78,
        autopct=lambda pct: f"{pct:.1f}%\n({int(round(pct*sum(sizes)/100))})",
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=3),
        textprops={"fontsize": 13, "fontweight": "bold"},
    )
    for t in texts:
        t.set_fontsize(14); t.set_fontweight("bold"); t.set_color("#333333")
    for t in autotexts:
        t.set_color("white")
    ax.set_title("Natureza da Demanda", fontsize=15, fontweight="bold", color="#" + COR["navy"], pad=20)
    ax.text(0, 0, f'{painel["total"]}\nchamados', ha="center", va="center",
            fontsize=13, fontweight="bold", color="#555555")
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_barras_categoria(cat_df: pd.DataFrame, titulo: str, caminho: str, destaque_top_n: int = 2):
    cats = cat_df.index.tolist()[::-1]
    vals = cat_df["qtd"].tolist()[::-1]
    total = sum(vals)
    destaque = set(cat_df.index[:destaque_top_n])
    colors = ["#" + COR["navy"] if c in destaque else COR["lightgray"] for c in cats]

    fig, ax = plt.subplots(figsize=(9, max(3.5, 0.55 * len(cats))), dpi=150)
    bars = ax.barh(cats, vals, color=colors, edgecolor="white", height=0.65)
    for bar, v in zip(bars, vals):
        pct = v / total * 100 if total else 0
        ax.text(bar.get_width() + max(vals) * 0.015, bar.get_y() + bar.get_height() / 2,
                f"{v}  ({pct:.1f}%)", va="center", fontsize=10.5, fontweight="bold", color="#333333")
    ax.set_xlim(0, max(vals) * 1.2)
    ax.set_title(titulo, fontsize=15, fontweight="bold", color="#" + COR["navy"], pad=15)
    ax.set_xlabel("Quantidade de chamados", fontsize=10, color="#555555")
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_scatter_prioridade(m: pd.DataFrame, frente: str, outliers: pd.DataFrame, caminho: str):
    sub = m[(m["frente"] == frente) & (m["concluido"])].copy()
    ordem = ["Muito Baixa", "Baixa", "Média", "Alta", "Muito Alta", "Crítica"]
    ordem_labels = [o.replace(" ", "\n") for o in ordem]
    prio_x = {o: i for i, o in enumerate(ordem)}
    sub["x"] = sub["prioridade_nome"].map(prio_x)
    rng = np.random.RandomState(42)
    sub["jitter"] = rng.uniform(-0.18, 0.18, len(sub))
    outlier_ids = set(outliers["id"]) if len(outliers) else set()
    sub["is_outlier"] = sub["id"].isin(outlier_ids)

    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=150)
    normal = sub[~sub["is_outlier"]]
    out = sub[sub["is_outlier"]]
    ax.scatter(normal["x"] + normal["jitter"], normal["tempo_h"], s=55, color="#" + COR["navy"],
               alpha=0.55, edgecolor="white", linewidth=0.5, zorder=2)
    if len(out):
        ax.scatter(out["x"] + out["jitter"], out["tempo_h"], s=180, color="#" + COR["red"],
                   edgecolor="black", linewidth=1.2, zorder=3, marker="D")
        top = out.iloc[0]
        ax.annotate(f'Chamado {int(top["id"])}\n({top["tempo_h"]:.0f} h — {top["categoria_final"]})',
                    xy=(top["x"], top["tempo_h"]), xytext=(0.4, sub["tempo_h"].max() * 0.85),
                    fontsize=10, fontweight="bold", color="#" + COR["red"],
                    arrowprops=dict(arrowstyle="->", color="#" + COR["red"], lw=1.5))
    ax.set_xticks(range(len(ordem)))
    ax.set_xticklabels(ordem_labels, fontsize=10)
    ax.set_ylabel("Tempo de Resolução (horas)", fontsize=10, color="#555555")
    ax.set_title(f"{frente} — Tempo de Resolução vs. Prioridade ({len(sub)} chamados)",
                 fontsize=14, fontweight="bold", color="#" + COR["navy"], pad=15)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_linha_recorrencia(rec_df: pd.DataFrame, bucket_labels: list, bucket_series, titulo: str, caminho: str):
    contagem = rec_df.groupby(bucket_series.loc[rec_df.index]).size().reindex(bucket_labels, fill_value=0)
    fig, ax = plt.subplots(figsize=(8, 5), dpi=150)
    x = range(len(contagem))
    ax.plot(x, contagem.values, marker="o", markersize=9, linewidth=3, color="#" + COR["red"], zorder=3)
    ax.fill_between(x, contagem.values, color="#" + COR["red"], alpha=0.08)
    for xi, yi in zip(x, contagem.values):
        ax.annotate(str(yi), (xi, yi), textcoords="offset points", xytext=(0, 12),
                    ha="center", fontsize=12, fontweight="bold", color="#" + COR["red"])
    ax.set_xticks(list(x))
    ax.set_xticklabels(bucket_labels, fontsize=10)
    ax.set_ylim(0, max(contagem.values.max() + 2, 3))
    ax.set_ylabel("Nº de incidentes", fontsize=10, color="#555555")
    ax.set_title(titulo, fontsize=15, fontweight="bold", color="#" + COR["navy"], pad=15)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_pizza_horas_frente(cap_sust: dict, cap_proj: dict, caminho: str):
    labels = [CONFIG["nome_frente_sustentacao"], CONFIG["nome_frente_projetos"]]
    sizes = [cap_sust["hpc"], cap_proj["hpc"]]
    colors = ["#" + COR["navy"], "#" + COR["orange"]]
    fig, ax = plt.subplots(figsize=(6.5, 6.5), dpi=150)
    ax.pie(sizes, labels=labels, colors=colors, startangle=90,
           autopct=lambda p: f"{p:.1f}%\n({p*sum(sizes)/100:,.0f} h)".replace(",", "."),
           pctdistance=0.62, textprops={"fontsize": 13, "fontweight": "bold"},
           wedgeprops=dict(edgecolor="white", linewidth=3))
    ax.set_title("Horas Concluídas por Frente", fontsize=16, fontweight="bold", color="#" + COR["navy"], pad=20)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


def grafico_evolucao_historica(serie, grupo, campo, titulo, caminho, unidade=""):
    pontos = [(item["periodo"], item[campo] if grupo is None else item[grupo][campo])
              for item in serie]
    pontos = [(periodo, valor) for periodo, valor in pontos if valor is not None]
    if not pontos:
        return
    periodos, valores = zip(*pontos)
    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=150)
    ax.plot(periodos, valores, marker="o", linewidth=2.5, color="#" + COR["navy"])
    for periodo, valor in zip(periodos, valores):
        ax.annotate(f"{valor:.1f}{unidade}" if isinstance(valor, float) else f"{valor}{unidade}",
                    (periodo, valor), textcoords="offset points", xytext=(0, 9),
                    ha="center", fontsize=9, fontweight="bold")
    ax.set_title(titulo, fontsize=14, fontweight="bold", color="#" + COR["navy"], pad=15)
    ax.set_ylabel(unidade.strip() or "Valor", fontsize=10, color="#555555")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(caminho, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()


# ============================================================================
# 4. GERAÇÃO DO .DOCX (python-docx)
# ============================================================================

def _set_cell_background(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_text(cell, text, *, bold=False, color=None, size=10, align="center"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def adicionar_tabela(doc, headers, rows, widths_cm, header_bg=COR["navy"]):
    """rows: lista de listas de strings. Cada linha pode ter os atributos
    especiais opcionais como dict em rows_meta (mesma posição) para cor/negrito
    de células específicas — ver uso em `montar_documento`."""
    tabela = doc.add_table(rows=1, cols=len(headers))
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.autofit = False

    hdr_cells = tabela.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].width = Cm(widths_cm[i])
        _set_cell_text(hdr_cells[i], h, bold=True, color=(255, 255, 255), size=10)
        _set_cell_background(hdr_cells[i], header_bg)
    # marca a linha de cabeçalho para repetir em cada página (evita cabeçalho
    # "fantasma" vazio quando a tabela quebra entre páginas)
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
    run.italic = True
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    return p


def adicionar_paragrafo(doc, texto, *, italic=False, bold=False, size=10.5):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.italic = italic
    run.bold = bold
    run.font.size = Pt(size)
    return p


def adicionar_bullet(doc, texto):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(texto)
    run.font.size = Pt(10.5)
    return p


def adicionar_banner(doc, texto, cor_fundo, cor_texto):
    p = doc.add_paragraph()
    p_fmt = p.paragraph_format
    p_fmt.space_before = Pt(6)
    p_fmt.space_after = Pt(10)
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(*cor_texto)
    # fundo do parágrafo (via shading na borda do parágrafo)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), cor_fundo)
    pPr.append(shd)
    return p


def _formatar_valor_historico(valor, unidade=""):
    if valor is None:
        return "—"
    if isinstance(valor, float):
        texto = f"{valor:.1f}"
    else:
        texto = str(valor)
    return f"{texto}{unidade}"


def _formatar_variacao(valor, unidade=""):
    if valor is None:
        return "—"
    sinal = "+" if valor > 0 else ""
    if isinstance(valor, float):
        return f"{sinal}{valor:.1f}{unidade}"
    return f"{sinal}{valor}{unidade}"


def _descricao_periodo_anterior(periodo_atual, periodo_anterior):
    atual = datetime.strptime(periodo_atual, "%Y-%m")
    anterior = datetime.strptime(periodo_anterior, "%Y-%m")
    distancia = (atual.year - anterior.year) * 12 + atual.month - anterior.month
    if distancia == 1:
        return f"comparação com {periodo_anterior}"
    return f"comparação com {periodo_anterior} (lacuna de {distancia - 1} mês(es))"


def adicionar_imagem(doc, caminho, largura_cm):
    doc.add_picture(caminho, width=Cm(largura_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def montar_documento(dados: dict, caminho_saida: str):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = Cm(2)
    section.top_margin = section.bottom_margin = Cm(1.8)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    # ---- Cabeçalho ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(CONFIG["titulo_relatorio"])
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(*COR["navy_rgb"])

    p = doc.add_paragraph(CONFIG["subtitulo"])
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = Pt(11)

    p = doc.add_paragraph(f'Período: {dados["periodo_texto"]}')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True

    p = doc.add_paragraph(f'Redigido por {CONFIG["autor"]} · Gerado automaticamente em '
                           f'{datetime.now().strftime("%d/%m/%Y %H:%M")}')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = Pt(9)
    p.runs[0].font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    doc.add_paragraph()

    F_SUST = CONFIG["nome_frente_sustentacao"]
    F_PROJ = CONFIG["nome_frente_projetos"]

    # ---- Visão Histórica ----
    visao = dados["visao_historica"]
    atual_historico = dados["historico_atual"]
    adicionar_titulo(doc, "Visão Histórica")
    adicionar_paragrafo(
        doc,
        f'Período atual: {atual_historico["periodo"]} — '
        f'{atual_historico["total"]["total"]} chamados, '
        f'{atual_historico["total"]["resolvidos"]} resolvidos e '
        f'taxa de resolução de {atual_historico["total"]["taxa_resolucao"]}%.'
    )
    if visao["anterior"]:
        periodo_anterior = visao["anterior"]["periodo"]
        adicionar_paragrafo(
            doc,
            _descricao_periodo_anterior(atual_historico["periodo"], periodo_anterior) + ". "
            "Os valores não são classificados como favoráveis ou desfavoráveis.",
            italic=True,
        )
        rows = []
        unidades = {"Chamados": "", "Resolvidos": "", "Taxa de resolução": "%",
                    "TMR": "h", "Sustentação": "", "Projetos": "", "Aderência": "%"}
        for comparacao in visao["comparacoes"]:
            unidade = unidades[comparacao["indicador"]]
            rows.append([
                comparacao["indicador"],
                _formatar_valor_historico(comparacao["atual"], unidade),
                _formatar_valor_historico(comparacao["anterior"], unidade),
                _formatar_variacao(comparacao["diferenca"], unidade),
                _formatar_variacao(comparacao["variacao_pct"], "%"),
            ])
        adicionar_tabela(doc, ["Indicador", "Atual", "Anterior", "Diferença", "Variação"],
                          rows, [4.5, 2.7, 2.7, 3, 3])
    else:
        adicionar_nota(doc, "Ainda não há outro período histórico disponível para comparação.")

    if visao["tendencia"]:
        adicionar_titulo(doc, "Tendência Histórica", nivel=2)
        adicionar_paragrafo(doc, "Série dos meses já consolidados, sem reprocessar os Excels anteriores.")
        for caminho in dados["img_historico"]:
            adicionar_imagem(doc, caminho, 13)
    else:
        adicionar_nota(doc, "A tendência será exibida quando houver pelo menos três meses históricos.")

    if visao["projecao"]:
        adicionar_titulo(doc, "Projeção do Próximo Mês", nivel=2)
        adicionar_paragrafo(doc, "ESTIMATIVA baseada na média móvel simples dos três últimos meses; não é uma previsão garantida.", italic=True)
        rows = [[nome, _formatar_valor_historico(valor, "%" if nome in ("Taxa de resolução", "Aderência") else "h" if nome == "TMR" else "")]
                for nome, valor in visao["projecao"].items()]
        adicionar_tabela(doc, ["Indicador", "Estimativa"], rows, [8, 5])
    else:
        adicionar_nota(doc, "A projeção será exibida quando houver pelo menos três meses históricos.")

    # ---- 1. Painel Geral ----
    adicionar_titulo(doc, "1. Painel Geral")
    adicionar_paragrafo(doc, f'No período, foram registrados {dados["painel_total"]["total"]} chamados. '
                              f'Consolidando as duas frentes:')
    pg, pp, pt = dados["painel_sust"], dados["painel_proj"], dados["painel_total"]
    adicionar_tabela(
        doc,
        ["Indicador", F_SUST, F_PROJ, "Total"],
        [
            ["Total de Demandas", pg["total"], pp["total"], pt["total"]],
            ["Entregues", pg["entregues"], pp["entregues"], pt["entregues"]],
            ["Taxa de Resolução", f'{pg["taxa_resolucao"]}%', f'{pp["taxa_resolucao"]}%', f'{pt["taxa_resolucao"]}%'],
            ["TMR (concluídos)", f'{pg["tmr_h"]} h' if pg["tmr_h"] else "—",
             f'{pp["tmr_h"]} h' if pp["tmr_h"] else "—", f'{pt["tmr_h"]} h' if pt["tmr_h"] else "—"],
            ["Natureza da Demanda", f'{pg["pct_req"]}% Req / {pg["pct_inc"]}% Inc',
             f'{pp["pct_req"]}% Req / {pp["pct_inc"]}% Inc', f'{pt["pct_req"]}% Req / {pt["pct_inc"]}% Inc'],
        ],
        [5.5, 4, 4, 3],
    )
    doc.add_paragraph()
    adicionar_imagem(doc, dados["img_donut"], 9)
    adicionar_nota(doc, "Natureza da Demanda — visão combinada das duas frentes.")

    # ---- 2. Categorias ----
    adicionar_titulo(doc, "2. Distribuição por Categoria")
    for frente, cor_banner in [(F_SUST, "navy"), (F_PROJ, "orange")]:
        bg = COR["banner_blue_bg"] if cor_banner == "navy" else COR["banner_orange_bg"]
        tx = COR["navy_rgb"] if cor_banner == "navy" else COR["orange_rgb"]
        total_frente = dados["painel_sust"]["total"] if frente == F_SUST else dados["painel_proj"]["total"]
        adicionar_banner(doc, f"FRENTE {frente.upper()} — {total_frente} chamados", bg, tx)
        cat_df = dados["cat_sust"] if frente == F_SUST else dados["cat_proj"]
        rows = []
        for idx, r in cat_df.iterrows():
            rows.append([idx, int(r["qtd"]), f'{r["pct_qtd"]}%', f'{r["tempo_h"]:.1f}', f'{r["pct_tempo"]}%'])
        adicionar_tabela(doc, ["Categoria", "Qtd.", "% Qtd.", "Tempo Total (h)", "% Tempo"],
                          rows, [6, 2, 2.2, 3.3, 2.5])
        doc.add_paragraph()

    adicionar_imagem(doc, dados["img_barras_sust"], 14)

    # ---- 3. Detalhamento da frente Projetos ----
    adicionar_titulo(doc, f"3. Frente {F_PROJ} — Detalhamento Completo")
    adicionar_paragrafo(doc, f'Chamados fora da entidade {CONFIG["entidade_sustentacao"]}, com status, datas e horas '
                              f'decorridas. Para os que ainda estão em aberto, a coluna de horas é o acumulado desde '
                              f'a abertura até agora — não é tempo de trabalho efetivo, apenas tempo corrido.', italic=True)
    tab_proj = dados["tabela_projetos"]
    rows = []
    total_h = 0.0
    for _, r in tab_proj.iterrows():
        horas = r["tempo_h"] if pd.notna(r["tempo_h"]) else r["horas_acumuladas"]
        total_h += horas if pd.notna(horas) else 0
        rows.append([
            int(r["id"]), r["titulo_limpo"], r["status_nome"],
            r["date"].strftime("%d/%m %H:%M"),
            r["data_solucao"].strftime("%d/%m %H:%M") if pd.notna(r["data_solucao"]) else "—",
            f"{horas:.1f}" if pd.notna(horas) else "—",
        ])
    rows.append(["", "", "", "", "TOTAL ACUMULADO", f"{total_h:,.1f}".replace(",", ".")])
    adicionar_tabela(doc, ["ID", "Título", "Status", "Abertura", "Solução", "Horas"],
                      rows, [1.3, 6.5, 2.3, 2.3, 2.3, 2])

    # ---- 4. SLA por prioridade (frente Sustentação) ----
    adicionar_titulo(doc, f"4. SLA por Prioridade — Frente {F_SUST}")
    sla = dados["sla_sust"]
    rows = []
    for idx, r in sla.iterrows():
        rows.append([idx, int(r["count"]), f'{r["mean"]:.1f} h', f'{r["min"]:.1f} h – {r["max"]:.1f} h'])
    adicionar_tabela(doc, ["Prioridade", "Resolvidos", "Tempo Médio", "Faixa (mín–máx)"],
                      rows, [4.5, 3, 3, 4])
    if len(dados["outliers_sust"]):
        top = dados["outliers_sust"].iloc[0]
        adicionar_nota(doc, f'Outlier identificado automaticamente (regra do IQR): chamado '
                             f'{int(top["id"])} ({top["categoria_final"]}), {top["tempo_h"]:.1f}h — muito acima do '
                             f'restante da amostra. Considere excluí-lo de médias de SLA e tratá-lo à parte.')
    adicionar_imagem(doc, dados["img_scatter_sust"], 14)

    # ---- 5. Capacidade ----
    adicionar_titulo(doc, f"5. Capacidade da Equipe — Frente {F_SUST}")
    p_cfg = CONFIG["premissas_equipe"]
    adicionar_paragrafo(
        doc,
        f'Premissas: {p_cfg["analistas"]} analistas; jornada de {p_cfg["horas_por_dia"]}h/dia; '
        f'{p_cfg["dias_uteis_mes"]} dias úteis no mês; desconto de {p_cfg["dias_ferias_no_mes"]} dias úteis '
        f'de férias/ausências no período.', italic=True)
    cap = dados["cap_sust"]
    adicionar_tabela(
        doc, ["Sigla", "Descrição / Fórmula", "Valor"],
        [
            ["HMM", "Horas totais projetadas", f'{cap["hmm"]} h'],
            ["HE", "Horas efetivas (descontando ausências)", f'{cap["he"]} h'],
            ["HPC", f"Tempo total decorrido registrado ({F_SUST}, concluídos)", f'{cap["hpc"]} h'],
            ["HHA", "Horas úteis alocadas (HPC ÷ horas/dia)", f'{cap["hha"]} h'],
            ["AD%", "Aderência (HHA ÷ HE)", f'{cap["ad_pct"]}%'],
        ],
        [2, 9.5, 3],
    )
    adicionar_imagem(doc, dados["img_pizza"], 8)

    # ---- 6. Recorrência detectada ----
    if len(dados["recorrencia"]):
        adicionar_titulo(doc, f'6. Padrão Recorrente Detectado — {CONFIG["nome_recorrencia"]}')
        n_sust = int((dados["recorrencia"]["frente"] == F_SUST).sum())
        n_proj = int((dados["recorrencia"]["frente"] == F_PROJ).sum())
        adicionar_paragrafo(doc, f'{len(dados["recorrencia"])} chamados no período mencionam os termos configurados '
                                 f'({", ".join(CONFIG["termos_chave_recorrencia"])}) — {n_sust} na frente {F_SUST} '
                                 f'e {n_proj} na frente {F_PROJ}.')
        adicionar_imagem(doc, dados["img_linha_recorrencia"], 12)

    # ---- 7. Plano de ação (esqueleto para o autor preencher) ----
    adicionar_titulo(doc, "7. Observações e Plano de Ação")
    adicionar_paragrafo(doc, "Espaço para o autor do relatório completar com as ações do mês "
                             "(este bloco é só um roteiro — edite livremente após gerar o .docx):", italic=True)
    for item in [
        "Ação 1: ",
        "Ação 2: ",
        "Ação 3: ",
    ]:
        adicionar_bullet(doc, item)

    doc.save(caminho_saida)


# ============================================================================
# 5. ORQUESTRAÇÃO
# ============================================================================

def gerar_relatorio_dataframe(m: pd.DataFrame):
    pasta = _nova_pasta_execucao()
    _preparar_pasta_saida(pasta)

    print("[1/5] Usando dados revisados e normalizados ...")
    periodo_texto, ano, mes, ultimo_dia = periodo_do_relatorio(m)
    bucket_series, bucket_labels = buckets_semanais(m, ano, mes, ultimo_dia)

    F_SUST = CONFIG["nome_frente_sustentacao"]
    F_PROJ = CONFIG["nome_frente_projetos"]

    print("[2/5] Calculando métricas (painel, categorias, SLA, capacidade, outliers) ...")
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

    tabela_projetos = tabela_frente_secundaria(m, F_PROJ)
    recorrencia = detectar_recorrencia(m, CONFIG["termos_chave_recorrencia"])

    periodo = f"{ano:04d}-{mes:02d}"
    historico_atual = _metricas_historicas(
        periodo_texto, painel_total, painel_sust, painel_proj,
        cap_sust, cap_proj, cat_sust_full, cat_proj_full,
        outliers_sust, recorrencia,
    )
    historico_atual["periodo"] = periodo
    historico_anterior = [registro for registro in carregar_historico()
                          if registro.get("periodo") != periodo]
    visao_historica = preparar_visao_historica(historico_anterior, historico_atual)

    print("[3/5] Gerando gráficos ...")
    img_donut = os.path.join(pasta, "01_donut_natureza.png")
    img_barras_sust = os.path.join(pasta, "02_barras_categoria_sustentacao.png")
    img_scatter_sust = os.path.join(pasta, "03_scatter_sustentacao.png")
    img_pizza = os.path.join(pasta, "04_pizza_horas_frente.png")
    img_linha_recorrencia = os.path.join(pasta, "05_linha_recorrencia.png")
    img_historico = []
    serie_historica = sorted(historico_anterior + [historico_atual],
                             key=lambda item: item["periodo"])
    if len(serie_historica) >= 2:
        img_volume = os.path.join(pasta, "06_evolucao_volume.png")
        img_tmr = os.path.join(pasta, "07_evolucao_tmr.png")
        img_taxa = os.path.join(pasta, "08_evolucao_taxa_resolucao.png")
        grafico_evolucao_historica(serie_historica, "total", "total",
                                   "Evolução Mensal — Volume de Chamados", img_volume)
        grafico_evolucao_historica(serie_historica, "total", "tmr_h",
                                   "Evolução Mensal — TMR", img_tmr, "h")
        grafico_evolucao_historica(serie_historica, "total", "taxa_resolucao",
                                   "Evolução Mensal — Taxa de Resolução", img_taxa, "%")
        img_historico = [img_volume, img_tmr, img_taxa]

    grafico_donut_natureza(painel_total, img_donut)
    grafico_barras_categoria(cat_sust, f"{F_SUST} — Volume por Categoria ({painel_sust['total']} chamados)",
                              img_barras_sust)
    grafico_scatter_prioridade(m, F_SUST, outliers_sust, img_scatter_sust)
    grafico_pizza_horas_frente(cap_sust, cap_proj, img_pizza)
    if len(recorrencia):
        grafico_linha_recorrencia(recorrencia, bucket_labels, bucket_series,
                                   f'Volume Semanal — {CONFIG["nome_recorrencia"]}', img_linha_recorrencia)

    print("[4/5] Montando o documento .docx ...")
    dados = dict(
        periodo_texto=periodo_texto,
        painel_total=painel_total, painel_sust=painel_sust, painel_proj=painel_proj,
        cat_sust=cat_sust, cat_proj=cat_proj,
        sla_sust=sla_sust, outliers_sust=outliers_sust,
        cap_sust=cap_sust, cap_proj=cap_proj,
        tabela_projetos=tabela_projetos,
        recorrencia=recorrencia,
        historico_atual=historico_atual,
        visao_historica=visao_historica,
        img_historico=img_historico,
        img_donut=img_donut, img_barras_sust=img_barras_sust,
        img_scatter_sust=img_scatter_sust, img_pizza=img_pizza,
        img_linha_recorrencia=img_linha_recorrencia,
    )
    caminho_docx = os.path.join(pasta, CONFIG["nome_docx"])
    montar_documento(dados, caminho_docx)
    salvar_historico(historico_atual)

    print(f"[5/5] Concluído! Relatório salvo em: {caminho_docx}")
    return caminho_docx


def gerar_relatorio(caminho_entrada: str):
    return gerar_relatorio_dataframe(preparar_dataframe(caminho_entrada))


if __name__ == "__main__":
    entrada = sys.argv[1] if len(sys.argv) > 1 else CONFIG["arquivo_entrada"]
    if not os.path.exists(entrada):
        print(f"Arquivo não encontrado: {entrada}")
        print("Uso: python3 gerar_relatorio_glpi.py caminho/para/base_de_dados.xlsx|csv")
        sys.exit(1)
    gerar_relatorio(entrada)
