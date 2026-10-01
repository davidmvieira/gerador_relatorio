"""Interface Flask para revisão e geração de relatórios GLPI.

O estado das sessões é persistido em JSON sob ``instance/sessoes`` por
``sessoes_store``; o dict legado existe apenas para migração durante reload.
"""

import os
import uuid
from datetime import datetime

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

import gerar_relatorio_glpi as gerador
import sessoes_store


app = Flask(__name__)
app.secret_key = os.environ.get("GLPI_REPORT_SECRET", "glpi-local-review")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
sessoes_store.configurar_pasta(os.path.join(app.instance_path, "sessoes"))
app.config["SESSION_TTL_DAYS"] = int(os.environ.get("GLPI_SESSION_TTL_DAYS", "7"))
sessoes_store.limpar_sessoes_expiradas(app.config["SESSION_TTL_DAYS"])

# Ponte temporária para sessões existentes durante reload/deploy no mesmo processo.
SESSOES = globals().get("SESSOES", {})
EXTENSOES_PERMITIDAS = {".csv", ".xlsx", ".xlsm"}
PAGE_SIZE = 20

STATUS = {1: "Novo", 2: "Processando", 3: "Pendente", 4: "Planejado",
          5: "Solucionado", 6: "Fechado"}
PRIORIDADES = {1: "Muito Baixa", 2: "Baixa", 3: "Média", 4: "Alta",
               5: "Muito Alta", 6: "Crítica"}
TIPOS = {1: "Incidente", 2: "Requisição"}

# Fases do fluxo de Projetos (ordem importa — usada nos selects e ordenação)
FASES = {
    "backlog": "Backlog",
    "planejamento": "Planejamento",
    "fila": "Fila",
    "desenvolvimento": "Desenvolvimento",
    "bloqueado": "Bloqueado",
    "testes": "Testes",
    "finalizados": "Finalizados",
}
FASE_PADRAO = "backlog"
FASE_ORDEM = {chave: i for i, chave in enumerate(FASES)}

# Mapeamento entre chave da aba e valor da coluna 'frente' do DataFrame
FRENTES = {
    "sustentacao": gerador.CONFIG["nome_frente_sustentacao"],
    "projetos": gerador.CONFIG["nome_frente_projetos"],
}
FRENTE_PADRAO = "sustentacao"


# -----------------------------------------------------------------------------
# Helpers de sessão
# -----------------------------------------------------------------------------
def _sessao(id_sessao):
    try:
        sessao = sessoes_store.carregar_sessao(id_sessao)
    except ValueError:
        return None
    if sessao is None:
        legado = SESSOES.get(id_sessao)
        if not legado:
            return None
        df_legado = legado.get("df")
        if not isinstance(df_legado, pd.DataFrame):
            return None
        df_legado = _garantir_campos_livres(df_legado)
        mes = legado.get("mes_referencia") or _mes_referencia_padrao(df_legado)
        sessoes_store.salvar_sessao({
            "id_sessao": id_sessao,
            "criado_em": datetime.now().isoformat(timespec="seconds"),
            "nome_arquivo_origem": "sessao-legada",
            "mes_referencia": mes,
            "erros": legado.get("erros", []),
            "avisos": legado.get("avisos", []),
            "resultado": legado.get("resultado"),
            "df": df_legado,
        })
        SESSOES.pop(id_sessao, None)
        sessao = sessoes_store.carregar_sessao(id_sessao)
    if sessao is not None:
        sessao["df"] = sessoes_store.df_da_sessao(sessao)
    return sessao


def _mes_referencia_padrao(df):
    datas = pd.to_datetime(df["date"], errors="coerce").dropna()
    if datas.empty:
        return datetime.now().strftime("%Y-%m")
    return pd.Timestamp(datas.median()).strftime("%Y-%m")


def _normalizar_mes_referencia(sessao, valor=None):
    if valor:
        sessao["mes_referencia"] = valor
    if not sessao.get("mes_referencia"):
        sessao["mes_referencia"] = _mes_referencia_padrao(sessao["df"])
    return sessao["mes_referencia"]


def _periodo(df, mes_referencia=None):
    if mes_referencia is None:
        mes_referencia = _mes_referencia_padrao(df)
    try:
        ano, mes = [int(p) for p in mes_referencia.split("-")]
        texto = gerador.periodo_texto(ano, mes)
    except (AttributeError, ValueError):
        ano = datetime.now().year
        mes = datetime.now().month
        texto = f"{mes:02d}/{ano}"
    return {"texto": texto, "ano": ano, "mes": mes, "rotulo": f"{mes:02d}/{ano}"}


def _df_com_frente(sessao):
    """Garante que o df da sessão tem as colunas derivadas (frente, fase, incluir)."""
    df = sessao["df"]
    if "frente" not in df.columns or "fase" not in df.columns:
        df = gerador.preparar_dataframe(df)
        df = _garantir_campos_livres(df)
        sessao["df"] = df
    return df


# -----------------------------------------------------------------------------
# Filtros / ordenação
# -----------------------------------------------------------------------------
def _aplicar_filtros(df, filtros):
    df = df.copy()

    # Frente (aba) — filtro principal
    frente_key = (filtros.get("frente") or "").strip().lower()
    if frente_key in FRENTES:
        df = df[df["frente"] == FRENTES[frente_key]]

    # Fase (só existe para projetos, mas o filtro é idempotente)
    fase = (filtros.get("fase") or "").strip()
    if fase and "fase" in df.columns:
        df = df[df["fase"].astype(str) == fase]

    busca = (filtros.get("q") or "").strip().lower()
    if busca:
        mascara = (
            df["id"].astype(str).str.lower().str.contains(busca, na=False)
            | df["name"].fillna("").astype(str).str.lower().str.contains(busca, na=False)
            | df["grupo_tecnico"].fillna("").astype(str).str.lower().str.contains(busca, na=False)
        )
        df = df[mascara].copy()

    entidade = (filtros.get("entidade") or "").strip()
    categoria = (filtros.get("categoria") or "").strip()
    status = (filtros.get("status") or "").strip()
    priority = (filtros.get("priority") or "").strip()
    tipo = (filtros.get("type") or "").strip()
    data_de = filtros.get("data_de")
    data_ate = filtros.get("data_ate")

    if entidade:
        df = df[df["entidade"].fillna("").astype(str).str.contains(entidade, case=False, na=False)]
    if categoria:
        df = df[df["categoria"].fillna("").astype(str).str.contains(categoria, case=False, na=False)]
    if status:
        df = df[df["status"].astype(str) == status]
    if priority:
        df = df[df["priority"].astype(str) == priority]
    if tipo:
        df = df[df["type"].astype(str) == tipo]
    if data_de:
        df = df[pd.to_datetime(df["data_abertura"], errors="coerce") >= pd.Timestamp(data_de)]
    if data_ate:
        df = df[pd.to_datetime(df["data_abertura"], errors="coerce") <= pd.Timestamp(data_ate)]
    return df


def _ordenar(df, sort, direcao):
    ordem_asc = (direcao or "asc").lower() == "asc"
    sort = sort or "data_abertura"
    campos_validos = {"id", "data_abertura", "priority", "status", "fase"}
    if sort not in campos_validos:
        sort = "data_abertura"

    df = df.copy()
    if sort == "data_abertura":
        df["_ordem"] = pd.to_datetime(df["data_abertura"], errors="coerce")
    elif sort == "fase":
        df["_ordem"] = df["fase"].map(FASE_ORDEM).fillna(99)
    else:
        df["_ordem"] = df[sort]
    df = df.sort_values(by="_ordem", ascending=ordem_asc, na_position="last")
    return df.drop(columns=["_ordem"], errors="ignore")


def _indicadores(df, mes_referencia=None, frente_key=None):
    base = df.copy()
    if {"frente", "concluido", "status", "priority", "type", "date",
        "data_abertura", "data_solucao"} - set(base.columns):
        base = gerador.preparar_dataframe(base)
    mes_referencia = mes_referencia or _mes_referencia_padrao(base)
    try:
        ano, mes = [int(p) for p in mes_referencia.split("-")]
        base = gerador.filtrar_periodo_referencia(base, ano, mes)
    except Exception:
        pass

    if frente_key and frente_key in FRENTES:
        base = base[base["frente"] == FRENTES[frente_key]]

    total = gerador.metricas_painel(base)
    return {
        "total": total.get("total_relevante", len(base)),
        "resolvidos": total["entregues"],
        "taxa": total["taxa_resolucao"],
        "tmr": total["tmr_h"],
    }


# -----------------------------------------------------------------------------
# Validação / normalização
# -----------------------------------------------------------------------------
def _valor_data(valor):
    if pd.isna(valor):
        return ""
    return pd.Timestamp(valor).strftime("%Y-%m-%dT%H:%M")


def _texto(valor):
    return "" if pd.isna(valor) else str(valor)


def _validar(df):
    erros, avisos = [], []
    obrigatorias = ["id", "name", "entidade", "categoria", "status",
                    "priority", "type", "date"]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        return [f"Colunas obrigatórias ausentes: {', '.join(faltantes)}"], []

    datas_abertura = pd.to_datetime(df["data_abertura"], errors="coerce")
    datas_solucao = pd.to_datetime(df["data_solucao"], errors="coerce")
    for indice, linha in df.iterrows():
        chamado = _texto(linha["id"])
        if pd.isna(linha["date"]) or pd.isna(datas_abertura[indice]):
            erros.append(f"Chamado {chamado}: data de abertura inválida")
        if pd.notna(datas_solucao[indice]) and pd.notna(datas_abertura[indice]) \
                and datas_solucao[indice] < datas_abertura[indice]:
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
        if "fase" in df.columns and linha.get("frente") == FRENTES["projetos"]:
            if str(linha.get("fase", "")).strip() not in FASES:
                erros.append(f"Chamado {chamado}: fase inválida ({linha.get('fase')!r})")

    if df["id"].duplicated().any():
        erros.append("Existem identificadores de chamados duplicados")
    if df["name"].isna().sum():
        avisos.append(f"{int(df['name'].isna().sum())} registro(s) sem título")
    return erros, avisos


def _atualizar_por_formulario(df):
    campos = ["categoria", "entidade", "status", "priority", "type", "fase",
              "data_abertura", "data_solucao", "grupo_tecnico", "observacoes"]
    df = df.copy()
    for indice in df.index:
        for campo in campos:
            chave = f"{campo}_{indice}"
            if chave not in request.form:
                continue
            valor = request.form[chave].strip()
            if campo in ("status", "priority", "type"):
                df.at[indice, campo] = int(valor) if valor else None
            elif campo == "fase":
                df.at[indice, campo] = valor if valor in FASES else FASE_PADRAO
            elif campo in ("data_abertura", "data_solucao"):
                df.at[indice, campo] = pd.to_datetime(valor, errors="coerce") if valor else pd.NaT
                if campo == "data_abertura":
                    df.at[indice, "date"] = df.at[indice, campo]
            else:
                df.at[indice, campo] = valor
    return gerador.preparar_dataframe(df)


def _garantir_campos_livres(df):
    df = df.copy()
    for coluna in ("grupo_tecnico", "observacoes"):
        if coluna not in df.columns:
            df[coluna] = ""
    if "fase" not in df.columns:
        df["fase"] = FASE_PADRAO
    df["fase"] = df["fase"].fillna(FASE_PADRAO).replace("", FASE_PADRAO)
    if "incluir" not in df.columns:
        df["incluir"] = True
    df["incluir"] = df["incluir"].fillna(True).astype(bool)
    return df


# -----------------------------------------------------------------------------
# Rotas
# -----------------------------------------------------------------------------
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
            base = gerador.carregar_dados(caminho)
            df = _garantir_campos_livres(gerador.preparar_dataframe(base))
            if df.empty:
                raise ValueError("O arquivo não contém registros com identificador.")
            erros, avisos = _validar(df)
        except (KeyError, ValueError, pd.errors.ParserError) as erro:
            flash(f"Arquivo inválido: {erro}", "error")
            return redirect(url_for("index"))
        finally:
            if os.path.exists(caminho):
                os.remove(caminho)

        id_sessao = sessoes_store.criar_sessao(
            df.reset_index(drop=True), arquivo.filename, erros, avisos)
        return redirect(url_for("revisao", id_sessao=id_sessao))
    return render_template("index.html")


@app.route("/revisao/<id_sessao>")
def revisao(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada. Envie o arquivo novamente.", "error")
        return redirect(url_for("index"))

    frente_key = (request.args.get("frente") or FRENTE_PADRAO).strip().lower()
    if frente_key not in FRENTES:
        frente_key = FRENTE_PADRAO

    filtros = {
        "frente": frente_key,
        "fase": request.args.get("fase", "").strip(),
        "entidade": request.args.get("entidade", "").strip(),
        "categoria": request.args.get("categoria", "").strip(),
        "status": request.args.get("status", "").strip(),
        "priority": request.args.get("priority", "").strip(),
        "type": request.args.get("type", "").strip(),
        "data_de": request.args.get("data_de", "").strip(),
        "data_ate": request.args.get("data_ate", "").strip(),
        "q": request.args.get("q", "").strip(),
    }
    sort = request.args.get("sort", "data_abertura")
    direcao = request.args.get("dir", "desc")

    mes_solicitado = request.args.get("mes_referencia")
    if mes_solicitado and mes_solicitado != sessao.get("mes_referencia"):
        sessao = sessoes_store.atualizar_sessao(
            id_sessao, lambda atual: atual.update({"mes_referencia": mes_solicitado}))
        sessao["df"] = sessoes_store.df_da_sessao(sessao)
    _normalizar_mes_referencia(sessao, mes_solicitado or sessao.get("mes_referencia"))
    df = _df_com_frente(sessao).copy()

    # Contadores globais por aba (antes de aplicar filtros locais)
    total_sust = int((df["frente"] == FRENTES["sustentacao"]).sum())
    total_proj = int((df["frente"] == FRENTES["projetos"]).sum())

    df = _aplicar_filtros(df, filtros)
    df = _ordenar(df, sort, direcao)

    total_geral = len(sessao["df"])
    total_filtrado = len(df)
    pagina = max(1, request.args.get("pagina", 1, type=int))
    total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)
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
            "fase": _texto(linha.get("fase", FASE_PADRAO)),
            "status": int(linha["status"]) if pd.notna(linha["status"]) else "",
            "priority": int(linha["priority"]) if pd.notna(linha["priority"]) else "",
            "type": int(linha["type"]) if pd.notna(linha["type"]) else "",
            "data_abertura": _valor_data(linha["data_abertura"]),
            "data_solucao": _valor_data(linha["data_solucao"]),
            "grupo_tecnico": _texto(linha["grupo_tecnico"]),
            "observacoes": _texto(linha["observacoes"]),
            "incluir": bool(linha.get("incluir", True)),
        })

    return render_template(
        "review.html",
        sessao=id_sessao,
        registros=registros,
        pagina=pagina,
        total_paginas=total_paginas,
        total=len(df),
        total_geral=total_geral,
        total_sust=total_sust,
        total_proj=total_proj,
        indicadores=_indicadores(sessao["df"], sessao["mes_referencia"], frente_key),
        filtros=filtros,
        sort=sort,
        direcao=direcao,
        mes_referencia=sessao["mes_referencia"],
        erros=sessao["erros"],
        avisos=sessao["avisos"],
        status=STATUS,
        prioridades=PRIORIDADES,
        tipos=TIPOS,
        fases=FASES,
        frente=frente_key,
        frente_nome=FRENTES[frente_key],
    )


@app.route("/revisao/<id_sessao>/salvar", methods=["POST"])
def salvar_revisao(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada.", "error")
        return redirect(url_for("index"))

    def aplicar_edicoes(atual):
        df = _atualizar_por_formulario(sessoes_store.df_da_sessao(atual))
        df = _garantir_campos_livres(df)
        for indice in df.index:
            if f"categoria_{indice}" in request.form:
                df.at[indice, "incluir"] = f"incluir_{indice}" in request.form
        atual["erros"], atual["avisos"] = _validar(df)
        atual["df"] = df
        atual.pop("registros", None)

    sessoes_store.atualizar_sessao(id_sessao, aplicar_edicoes)

    frente = request.form.get("frente", FRENTE_PADRAO)
    pagina = request.form.get("pagina", "1")
    flash("Alterações salvas e indicadores recalculados.", "success")
    return redirect(url_for("revisao", id_sessao=id_sessao,
                            frente=frente, pagina=pagina))


@app.route("/revisao/<id_sessao>/bulk", methods=["POST"])
def bulk_revisao(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada.", "error")
        return redirect(url_for("index"))

    frente = request.form.get("frente", FRENTE_PADRAO)
    try:
        selecionados = [int(item) for item in request.form.getlist("selecionados")]
    except ValueError:
        selecionados = []
    if not selecionados:
        flash("Selecione pelo menos um registro para aplicar a ação.", "warning")
        return redirect(url_for("revisao", id_sessao=id_sessao, frente=frente))

    acao = request.form.get("acao")
    def aplicar_acao_em_lote(atual):
        df = sessoes_store.df_da_sessao(atual)
        indices_validos = [i for i in selecionados if i in df.index]

        if acao == "excluir":
            df.loc[indices_validos, "incluir"] = False
        elif acao == "incluir":
            df.loc[indices_validos, "incluir"] = True
        elif acao == "bulk_edit":
            if request.form.get("bulk_status"):
                df.loc[indices_validos, "status"] = int(request.form["bulk_status"])
            if request.form.get("bulk_priority"):
                df.loc[indices_validos, "priority"] = int(request.form["bulk_priority"])
            if request.form.get("bulk_type"):
                df.loc[indices_validos, "type"] = int(request.form["bulk_type"])
            if request.form.get("bulk_categoria"):
                df.loc[indices_validos, "categoria"] = request.form["bulk_categoria"]
            if request.form.get("bulk_entidade"):
                df.loc[indices_validos, "entidade"] = request.form["bulk_entidade"]
            if request.form.get("bulk_fase") in FASES:
                df.loc[indices_validos, "fase"] = request.form["bulk_fase"]

        df = _garantir_campos_livres(gerador.preparar_dataframe(df))
        atual["erros"], atual["avisos"] = _validar(df)
        atual["df"] = df
        atual.pop("registros", None)

    sessoes_store.atualizar_sessao(id_sessao, aplicar_acao_em_lote)
    flash("Ação em massa aplicada com sucesso.", "success")
    return redirect(url_for("revisao", id_sessao=id_sessao, frente=frente))


@app.route("/preview/<id_sessao>")
def preview(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada.", "error")
        return redirect(url_for("index"))

    mes_referencia = _normalizar_mes_referencia(
        sessao, request.args.get("mes_referencia") or sessao.get("mes_referencia"))
    ano, mes = [int(p) for p in mes_referencia.split("-")]
    df = _df_com_frente(sessao).copy()
    df = df[df.get("incluir", True)].copy()
    sub = gerador.filtrar_periodo_referencia(df, ano, mes)

    F_SUST = gerador.CONFIG["nome_frente_sustentacao"]
    F_PROJ = gerador.CONFIG["nome_frente_projetos"]

    painel_total = gerador.metricas_painel(sub)
    painel_sust = gerador.metricas_painel(sub, F_SUST)
    painel_proj = gerador.metricas_painel(sub, F_PROJ)

    def _cat_records(frente):
        d = gerador.metricas_categoria(sub, frente).head(6).reset_index()
        d = d.rename(columns={"index": "categoria"})
        return d.to_dict("records")

    categorias_sust = _cat_records(F_SUST)
    categorias_proj = _cat_records(F_PROJ)
    sla = gerador.metricas_sla_prioridade(sub, F_SUST).reset_index() \
        .rename(columns={"index": "prioridade"}).to_dict("records")
    capacidade_sust = gerador.metricas_capacidade(sub, F_SUST)
    capacidade_proj = gerador.metricas_capacidade(sub, F_PROJ)

    # Projetos relevantes com fase (substitui o campo 'origem' derivado)
    proj_df = sub[sub["frente"] == F_PROJ].copy()
    if "titulo_limpo" not in proj_df.columns:
        proj_df["titulo_limpo"] = proj_df["name"].apply(gerador.limpar_titulo)
    proj_df["fase_nome"] = proj_df["fase"].map(FASES).fillna(FASES[FASE_PADRAO])
    proj_df["data_abertura_fmt"] = pd.to_datetime(proj_df["date"], errors="coerce") \
        .dt.strftime("%d/%m/%y")
    projetos = proj_df.sort_values(["fase", "date"]).head(50).to_dict("records")

    return render_template(
        "preview.html",
        sessao=id_sessao,
        periodo=_periodo(sub, mes_referencia),
        mes_referencia=mes_referencia,
        painel_total=painel_total,
        painel_sust=painel_sust,
        painel_proj=painel_proj,
        categorias_sust=categorias_sust,
        categorias_proj=categorias_proj,
        sla=sla,
        capacidade_sust=capacidade_sust,
        capacidade_proj=capacidade_proj,
        projetos=projetos,
        status=STATUS,
        prioridades=PRIORIDADES,
        fases=FASES,
    )


@app.route("/gerar/<id_sessao>", methods=["POST"])
def gerar(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao:
        flash("Sessão de revisão não encontrada.", "error")
        return redirect(url_for("index"))

    def salvar_edicoes_para_geracao(atual):
        df = _atualizar_por_formulario(sessoes_store.df_da_sessao(atual))
        df = _garantir_campos_livres(df)
        for indice in df.index:
            if f"categoria_{indice}" in request.form:
                df.at[indice, "incluir"] = f"incluir_{indice}" in request.form
        atual["erros"], atual["avisos"] = _validar(df)
        atual["df"] = df
        atual.pop("registros", None)
        mes_form = request.form.get("mes_referencia")
        if mes_form:
            atual["mes_referencia"] = mes_form

    sessao = sessoes_store.atualizar_sessao(id_sessao, salvar_edicoes_para_geracao)
    sessao["df"] = sessoes_store.df_da_sessao(sessao)
    if sessao["erros"]:
        flash("Corrija os erros impeditivos antes de gerar o relatório.", "error")
        return redirect(url_for("revisao", id_sessao=id_sessao))

    dados_relevantes = sessao["df"][sessao["df"].get("incluir", True)].copy()
    if dados_relevantes.empty:
        flash("Nenhum registro está incluído para geração.", "warning")
        return redirect(url_for("revisao", id_sessao=id_sessao))

    mes_referencia = _normalizar_mes_referencia(
        sessao, request.form.get("mes_referencia") or sessao.get("mes_referencia"))
    ano, mes = [int(p) for p in mes_referencia.split("-")]
    df_filtrado = gerador.filtrar_periodo_referencia(dados_relevantes, ano, mes)
    if df_filtrado.empty:
        flash(f"Nenhum chamado relevante para {mes_referencia}.", "warning")
        return redirect(url_for("revisao", id_sessao=id_sessao))
    try:
        caminho = gerador.gerar_relatorio_dataframe(df_filtrado, ano, mes)
        if not os.path.isfile(caminho):
            raise RuntimeError("O arquivo DOCX não foi criado.")
    except Exception as erro:
        flash(f"Falha ao gerar o relatório: {erro}", "error")
        return redirect(url_for("revisao", id_sessao=id_sessao))

    periodo = _periodo(sessao["df"], mes_referencia)
    execucao = os.path.basename(os.path.dirname(caminho))
    sessoes_store.atualizar_sessao(
        id_sessao,
        lambda atual: atual.update({"resultado": {
            "caminho": caminho, "periodo": periodo, "execucao": execucao}}))
    return redirect(url_for("resultado", id_sessao=id_sessao))


@app.route("/resultado/<id_sessao>")
def resultado(id_sessao):
    sessao = _sessao(id_sessao)
    if not sessao or "resultado" not in sessao:
        return redirect(url_for("index"))
    resultado = sessao["resultado"]
    pasta = os.path.dirname(resultado["caminho"])
    arquivos = sorted(os.listdir(pasta))
    return render_template("result.html", sessao=id_sessao,
                           resultado=resultado, arquivos=arquivos)


@app.route("/resultado/<id_sessao>/arquivo/<nome>")
def arquivo_resultado(id_sessao, nome):
    sessao = _sessao(id_sessao)
    if not sessao or "resultado" not in sessao:
        return redirect(url_for("index"))
    pasta = os.path.dirname(sessao["resultado"]["caminho"])
    caminho = os.path.join(pasta, nome)
    if os.path.commonpath([os.path.abspath(pasta), os.path.abspath(caminho)]) \
            != os.path.abspath(pasta):
        return "Arquivo inválido", 400
    if not os.path.isfile(caminho):
        return "Arquivo não encontrado", 404
    return send_file(caminho, as_attachment=nome.lower().endswith(".docx"))


if __name__ == "__main__":
    app.run(debug=True)