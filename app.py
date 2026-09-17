import os
import re
import uuid
from datetime import datetime

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

import gerar_relatorio_glpi as gerador


app = Flask(__name__)
app.secret_key = os.environ.get("GLPI_REPORT_SECRET", "glpi-local-review")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

SESSOES = {}
EXTENSOES_PERMITIDAS = {".csv", ".xlsx", ".xlsm"}
PAGE_SIZE = 20

STATUS = {
    1: "Novo", 2: "Processando", 3: "Pendente", 4: "Planejado",
    5: "Solucionado", 6: "Fechado",
}
PRIORIDADES = {
    1: "Muito Baixa", 2: "Baixa", 3: "Média", 4: "Alta",
    5: "Muito Alta", 6: "Crítica",
}
TIPOS = {1: "Incidente", 2: "Requisição"}


def _sessao(id_sessao):
    return SESSOES.get(id_sessao)


def _periodo(df):
    texto, ano, mes, _ = gerador.periodo_do_relatorio(df)
    return {"texto": texto, "ano": ano, "mes": mes, "rotulo": f"{mes:02d}/{ano}"}


def _indicadores(df):
    total = gerador.metricas_painel(df)
    sust = gerador.metricas_painel(df, gerador.CONFIG["nome_frente_sustentacao"])
    proj = gerador.metricas_painel(df, gerador.CONFIG["nome_frente_projetos"])
    return {
        "total": total["total"],
        "resolvidos": total["entregues"],
        "taxa": total["taxa_resolucao"],
        "tmr": total["tmr_h"],
        "sustentacao": sust["total"],
        "projetos": proj["total"],
    }


def _valor_data(valor):
    if pd.isna(valor):
        return ""
    return pd.Timestamp(valor).strftime("%Y-%m-%dT%H:%M")


def _texto(valor):
    return "" if pd.isna(valor) else str(valor)


def _validar(df):
    erros = []
    avisos = []
    obrigatorias = ["id", "name", "entidade", "categoria", "status", "priority", "type", "date"]
    faltantes = [coluna for coluna in obrigatorias if coluna not in df.columns]
    if faltantes:
        return [f"Colunas obrigatórias ausentes: {', '.join(faltantes)}"], []

    datas_abertura = pd.to_datetime(df["data_abertura"], errors="coerce")
    datas_solucao = pd.to_datetime(df["data_solucao"], errors="coerce")
    for indice, linha in df.iterrows():
        chamado = _texto(linha["id"])
        if pd.isna(linha["date"]) or pd.isna(datas_abertura[indice]):
            erros.append(f"Chamado {chamado}: data de abertura inválida")
        if pd.notna(datas_solucao[indice]) and pd.notna(datas_abertura[indice]) and datas_solucao[indice] < datas_abertura[indice]:
            erros.append(f"Chamado {chamado}: data de solução anterior à abertura")
        if not _texto(linha["categoria"]).strip():
            erros.append(f"Chamado {chamado}: categoria não identificada")
        if linha["status"] not in STATUS:
            erros.append(f"Chamado {chamado}: status inválido")
        if linha["priority"] not in PRIORIDADES:
            erros.append(f"Chamado {chamado}: prioridade inválida")
        if linha["type"] not in TIPOS:
            erros.append(f"Chamado {chamado}: tipo inválido")
        if not _texto(linha["entidade"]).strip():
            erros.append(f"Chamado {chamado}: entidade não identificada")

    if df["id"].duplicated().any():
        erros.append("Existem identificadores de chamados duplicados")
    if df["name"].isna().sum():
        avisos.append(f"{int(df['name'].isna().sum())} registro(s) sem título")
    return erros, avisos


def _atualizar_por_formulario(df):
    campos = ["categoria", "entidade", "status", "priority", "type", "data_abertura", "data_solucao", "grupo_tecnico", "observacoes"]
    df = df.copy()
    for indice in df.index:
        for campo in campos:
            chave = f"{campo}_{indice}"
            if chave not in request.form:
                continue
            valor = request.form[chave].strip()
            if campo in ("status", "priority", "type"):
                df.at[indice, campo] = int(valor) if valor else None
            elif campo in ("data_abertura", "data_solucao"):
                df.at[indice, campo] = pd.to_datetime(valor, errors="coerce") if valor else pd.NaT
                if campo == "data_abertura":
                    df.at[indice, "date"] = df.at[indice, campo]
            else:
                df.at[indice, campo] = valor
    return gerador.preparar_dataframe_df(df)


def _colunas_tabela(df):
    return [
        coluna for coluna in
        ["id", "name", "categoria", "priority", "status", "type", "entidade", "data_abertura", "data_solucao"]
        if coluna in df.columns
    ]


def _garantir_campos_livres(df):
    df = df.copy()
    for coluna in ("grupo_tecnico", "observacoes"):
        if coluna not in df.columns:
            df[coluna] = ""
    return df


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        if not arquivo or not arquivo.filename:
            flash("Selecione um arquivo CSV ou Excel.", "error")
            return redirect(url_for("index"))
        extensao = os.path.splitext(arquivo.filename)[1].lower()
        if extensao not in EXTENSOES_PERMITIDAS:
            flash("Formato inválido. Use .csv, .xlsx ou .xlsm.", "error")
            return redirect(url_for("index"))

        nome = secure_filename(arquivo.filename)
        caminho = os.path.join(app.instance_path, "uploads", f"{uuid.uuid4().hex}_{nome}")
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        arquivo.save(caminho)
        try:
            df = _garantir_campos_livres(gerador.preparar_dataframe(caminho))
            if df.empty:
                raise ValueError("O arquivo não contém registros com identificador.")
            erros, avisos = _validar(df)
        except (KeyError, ValueError, pd.errors.ParserError) as erro:
            flash(f"Arquivo inválido: {erro}", "error")
            return redirect(url_for("index"))
        finally:
            if os.path.exists(caminho):
                os.remove(caminho)

        id_sessao = uuid.uuid4().hex
        SESSOES[id_sessao] = {"df": df.reset_index(drop=True), "avisos": avisos, "erros": erros}
        return redirect(url_for("revisao", id_sessao=id_sessao))
    return render_template("index.html")


@app.route("/revisao/<id_sessao>")
def revisao(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada. Envie o arquivo novamente.", "error")
        return redirect(url_for("index"))
    df = sessao["df"]
    pagina = max(1, request.args.get("pagina", 1, type=int))
    total_paginas = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
    pagina = min(pagina, total_paginas)
    inicio = (pagina - 1) * PAGE_SIZE
    fim = inicio + PAGE_SIZE
    registros = []
    for indice, linha in df.iloc[inicio:fim].iterrows():
        registros.append({
            "indice": indice,
            "id": _texto(linha["id"]),
            "name": _texto(linha["name"]),
            "categoria": _texto(linha["categoria"]),
            "entidade": _texto(linha["entidade"]),
            "status": int(linha["status"]) if pd.notna(linha["status"]) else "",
            "priority": int(linha["priority"]) if pd.notna(linha["priority"]) else "",
            "type": int(linha["type"]) if pd.notna(linha["type"]) else "",
            "data_abertura": _valor_data(linha["data_abertura"]),
            "data_solucao": _valor_data(linha["data_solucao"]),
            "grupo_tecnico": _texto(linha["grupo_tecnico"]),
            "observacoes": _texto(linha["observacoes"]),
        })
    return render_template(
        "review.html", sessao=id_sessao, registros=registros, pagina=pagina,
        total_paginas=total_paginas, total=len(df), indicadores=_indicadores(df),
        erros=sessao["erros"], avisos=sessao["avisos"], status=STATUS,
        prioridades=PRIORIDADES, tipos=TIPOS,
    )


@app.route("/revisao/<id_sessao>/salvar", methods=["POST"])
def salvar_revisao(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada.", "error")
        return redirect(url_for("index"))
    sessao["df"] = _atualizar_por_formulario(sessao["df"])
    sessao["erros"], sessao["avisos"] = _validar(sessao["df"])
    pagina = request.form.get("pagina", "1")
    flash("Alterações salvas e indicadores recalculados.", "success")
    return redirect(url_for("revisao", id_sessao=id_sessao, pagina=pagina))


@app.route("/gerar/<id_sessao>", methods=["POST"])
def gerar(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada.", "error")
        return redirect(url_for("index"))
    sessao["df"] = _atualizar_por_formulario(sessao["df"])
    sessao["erros"], sessao["avisos"] = _validar(sessao["df"])
    if sessao["erros"]:
        flash("Corrija os erros impeditivos antes de gerar o relatório.", "error")
        return redirect(url_for("revisao", id_sessao=id_sessao))
    try:
        caminho = gerador.gerar_relatorio_dataframe(sessao["df"])
    except Exception as erro:
        flash(f"Falha ao gerar o relatório: {erro}", "error")
        return redirect(url_for("revisao", id_sessao=id_sessao))
    periodo = _periodo(sessao["df"])
    execucao = os.path.basename(os.path.dirname(caminho))
    sessao["resultado"] = {"caminho": caminho, "periodo": periodo, "execucao": execucao}
    return redirect(url_for("resultado", id_sessao=id_sessao))


@app.route("/resultado/<id_sessao>")
def resultado(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao or "resultado" not in sessao:
        return redirect(url_for("index"))
    resultado = sessao["resultado"]
    pasta = os.path.dirname(resultado["caminho"])
    arquivos = sorted(os.listdir(pasta))
    return render_template("result.html", sessao=id_sessao, resultado=resultado, arquivos=arquivos)


@app.route("/resultado/<id_sessao>/arquivo/<nome>")
def arquivo_resultado(id_sessao, nome):
    sessao = _sessao(id_sessao)
    if not sessao or "resultado" not in sessao:
        return redirect(url_for("index"))
    pasta = os.path.dirname(sessao["resultado"]["caminho"])
    caminho = os.path.join(pasta, nome)
    if os.path.commonpath([os.path.abspath(pasta), os.path.abspath(caminho)]) != os.path.abspath(pasta):
        return "Arquivo inválido", 400
    if not os.path.isfile(caminho):
        return "Arquivo não encontrado", 404
    return send_file(caminho, as_attachment=nome.lower().endswith(".docx"))


if __name__ == "__main__":
    app.run(debug=True)
