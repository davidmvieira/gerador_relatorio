"""Persistência JSON das sessões de revisão do relatório GLPI.

Cada arquivo de sessão é a fonte da verdade. As escritas usam arquivo
temporário, substituição atômica, backup da versão anterior e lock por sessão.
"""

import errno
import json
import os
import re
import shutil
import time
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Callable, Optional

import pandas as pd

import gerar_relatorio_glpi as gerador


_PASTA_SESSOES = os.path.join(os.path.dirname(__file__), "instance", "sessoes")
_ID_RE = re.compile(r"^[0-9a-f]{32}$")
_CAMPOS = (
    "id", "name", "categoria", "entidade", "status", "priority", "type", "date",
    "data_abertura", "data_solucao", "tempo_resolucao_horas", "fase",
    "grupo_tecnico", "observacoes", "incluir",
)


def configurar_pasta(caminho: str) -> None:
    """Define o diretório de persistência, normalmente Flask.instance_path."""
    global _PASTA_SESSOES
    _PASTA_SESSOES = os.path.abspath(caminho)


def _validar_id(id_sessao: str) -> str:
    valor = str(id_sessao or "")
    if not _ID_RE.fullmatch(valor):
        raise ValueError("Identificador de sessão inválido.")
    return valor


def _caminhos(id_sessao: str):
    nome = _validar_id(id_sessao)
    os.makedirs(_PASTA_SESSOES, exist_ok=True)
    pasta_locks = os.path.join(_PASTA_SESSOES, ".locks")
    os.makedirs(pasta_locks, exist_ok=True)
    destino = os.path.join(_PASTA_SESSOES, f"{nome}.json")
    return destino, os.path.join(pasta_locks, f"{nome}.lock")


@contextmanager
def _lock_sessao(id_sessao: str):
    _, caminho_lock = _caminhos(id_sessao)
    with open(caminho_lock, "a+b") as arquivo:
        arquivo.seek(0, os.SEEK_END)
        if arquivo.tell() == 0:
            arquivo.write(b"\0")
            arquivo.flush()
        arquivo.seek(0)
        if os.name == "nt":
            import msvcrt

            while True:
                try:
                    msvcrt.locking(arquivo.fileno(), msvcrt.LK_NBLCK, 1)
                    break
                except OSError as erro:
                    if erro.errno not in (errno.EACCES, errno.EDEADLK, errno.EAGAIN):
                        raise
                    time.sleep(0.05)
        else:
            import fcntl

            fcntl.flock(arquivo.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            if os.name == "nt":
                arquivo.seek(0)
                msvcrt.locking(arquivo.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(arquivo.fileno(), fcntl.LOCK_UN)


def _nulo(valor) -> bool:
    if valor is None or valor is pd.NaT:
        return True
    try:
        resultado = pd.isna(valor)
        return bool(resultado) if not hasattr(resultado, "__len__") else False
    except (TypeError, ValueError):
        return False


def _data_iso(valor) -> Optional[str]:
    if _nulo(valor):
        return None
    return pd.Timestamp(valor).isoformat()


def _inteiro_ou_valor(valor):
    if _nulo(valor):
        return None
    if isinstance(valor, str):
        try:
            numero = float(valor)
        except ValueError:
            return valor
    else:
        try:
            numero = float(valor)
        except (TypeError, ValueError):
            return valor
    return int(numero) if numero.is_integer() else numero


def _json_escalar(valor):
    if _nulo(valor):
        return None
    if hasattr(valor, "item"):
        valor = valor.item()
    if isinstance(valor, (str, int, float, bool)):
        return valor
    return str(valor)


def _registros_dataframe(df: pd.DataFrame) -> list:
    registros = []
    for _, linha in df.iterrows():
        abertura = linha.get("data_abertura", linha.get("date"))
        if _nulo(abertura):
            abertura = linha.get("date")
        registro = {
            "id": _inteiro_ou_valor(linha.get("id")),
            "name": _json_escalar(linha.get("name")),
            "categoria": _json_escalar(linha.get("categoria")),
            "entidade": _json_escalar(linha.get("entidade")),
            "status": _inteiro_ou_valor(linha.get("status")),
            "priority": _inteiro_ou_valor(linha.get("priority")),
            "type": _inteiro_ou_valor(linha.get("type")),
            "date": _data_iso(linha.get("date", abertura)),
            "data_abertura": _data_iso(abertura),
            "data_solucao": _data_iso(linha.get("data_solucao")),
            "fase": _json_escalar(linha.get("fase")) or "backlog",
            "grupo_tecnico": _json_escalar(linha.get("grupo_tecnico")) or "",
            "observacoes": _json_escalar(linha.get("observacoes")) or "",
            "incluir": bool(linha.get("incluir", True)) if not _nulo(
                linha.get("incluir", True)) else True,
        }
        if "tempo_resolucao_horas" in df.columns:
            registro["tempo_resolucao_horas"] = _json_escalar(
                linha.get("tempo_resolucao_horas"))
        registros.append(registro)
    return registros


def _caminho_base(id_sessao: str) -> str:
    destino, _ = _caminhos(id_sessao)
    return destino[:-len(".json")] + ".base.json"


def _normalizar_registros(lista: list) -> list:
    registros = []
    for original in lista:
        registro = {campo: original.get(campo) for campo in _CAMPOS if campo in original}
        if "data_abertura" not in registro:
            registro["data_abertura"] = _data_iso(original.get("date"))
        for campo in ("date", "data_abertura", "data_solucao"):
            if campo in registro:
                registro[campo] = _data_iso(registro[campo])
        for campo in ("id", "status", "priority", "type"):
            if campo in registro:
                registro[campo] = _inteiro_ou_valor(registro[campo])
        registro.setdefault("fase", "backlog")
        registro.setdefault("grupo_tecnico", "")
        registro.setdefault("observacoes", "")
        registro.setdefault("incluir", True)
        registro["incluir"] = bool(registro["incluir"])
        if "tempo_resolucao_horas" in registro:
            registro["tempo_resolucao_horas"] = _json_escalar(
                registro["tempo_resolucao_horas"])
        registros.append(registro)
    return registros


def _gravar_json_atomico(destino: str, objeto, backup: bool) -> None:
    temporario = destino + ".tmp"
    backup_tmp = destino + ".bak.tmp"
    try:
        with open(temporario, "w", encoding="utf-8", newline="\n") as arquivo:
            json.dump(objeto, arquivo, ensure_ascii=False, indent=None if not backup else 2,
                      allow_nan=False, separators=None if backup else (",", ":"))
            arquivo.flush()
            os.fsync(arquivo.fileno())
        if backup and os.path.isfile(destino):
            shutil.copy2(destino, backup_tmp)
            os.replace(backup_tmp, destino + ".bak")
        os.replace(temporario, destino)
    finally:
        for caminho in (temporario, backup_tmp):
            if os.path.exists(caminho):
                os.remove(caminho)


# Cache da base por sessão: (mtime, registros). A base só é gravada na criação.
_CACHE_BASE: dict = {}


def _carregar_base(id_sessao: str) -> Optional[list]:
    caminho = _caminho_base(id_sessao)
    if not os.path.isfile(caminho):
        return None
    mtime = os.path.getmtime(caminho)
    em_cache = _CACHE_BASE.get(id_sessao)
    if em_cache and em_cache[0] == mtime:
        return em_cache[1]
    with open(caminho, "r", encoding="utf-8") as arquivo:
        registros = json.load(arquivo)
    _CACHE_BASE[id_sessao] = (mtime, registros)
    return registros


def _gravar_base(id_sessao: str, registros: list) -> None:
    _gravar_json_atomico(_caminho_base(id_sessao), registros, backup=False)
    _CACHE_BASE.pop(id_sessao, None)


def _diferencas(base: list, registros: list) -> dict:
    """Só o que mudou em relação à base: {posição: {campo: valor}}."""
    edicoes = {}
    for posicao, (original, atual) in enumerate(zip(base, registros)):
        mudou = {campo: atual.get(campo) for campo in _CAMPOS
                 if campo != "id" and atual.get(campo) != original.get(campo)}
        if mudou:
            edicoes[str(posicao)] = mudou
    return edicoes


def _registros_da_sessao(sessao: dict) -> list:
    """Base gravada na criação + edições. Aceita o formato antigo (registros inline)."""
    if "registros" in sessao:
        return sessao["registros"]
    base = _carregar_base(sessao["id_sessao"]) or []
    registros = [dict(registro) for registro in base]
    for posicao, campos in (sessao.get("edicoes") or {}).items():
        indice = int(posicao)
        if 0 <= indice < len(registros):
            registros[indice].update(campos)
    return registros


def _normalizar_sessao(sessao: dict) -> dict:
    """Prepara o arquivo pequeno da sessão (metadados + edições).

    Se vierem registros (ou um DataFrame), eles viram diferenças em relação à
    base; a base é gravada só quando ainda não existe (criação ou migração do
    formato antigo, que guardava a planilha inteira a cada edição).
    """
    agora = datetime.now().isoformat(timespec="seconds")
    payload = dict(sessao)
    if isinstance(payload.get("df"), pd.DataFrame):
        payload["registros"] = _registros_dataframe(payload["df"])
    payload.pop("df", None)
    payload.setdefault("id_sessao", uuid.uuid4().hex)
    payload["id_sessao"] = _validar_id(payload["id_sessao"])
    payload.setdefault("criado_em", agora)
    payload["atualizado_em"] = agora
    payload.setdefault("nome_arquivo_origem", "")
    payload.setdefault("mes_referencia", datetime.now().strftime("%Y-%m"))
    payload.setdefault("erros", [])
    payload.setdefault("avisos", [])
    payload.setdefault("edicoes", {})
    payload.pop("resultado", None)  # página de resultado não existe mais

    if "registros" in payload:
        registros = _normalizar_registros(payload.pop("registros"))
        base = _carregar_base(payload["id_sessao"])
        if base is None or len(base) != len(registros):
            _gravar_base(payload["id_sessao"], registros)
            payload["edicoes"] = {}
        else:
            payload["edicoes"] = _diferencas(base, registros)

    payload["erros"] = [str(item) for item in payload["erros"]]
    payload["avisos"] = [str(item) for item in payload["avisos"]]
    return payload


def _gravar_sem_lock(sessao: dict) -> dict:
    payload = _normalizar_sessao(sessao)
    destino, _ = _caminhos(payload["id_sessao"])
    _gravar_json_atomico(destino, payload, backup=True)
    return payload


def criar_sessao(df: pd.DataFrame, nome_arquivo: str, erros: list,
                 avisos: list) -> str:
    mes = gerador.metricas.mes_padrao(df)
    id_sessao = uuid.uuid4().hex
    salvar_sessao({
        "id_sessao": id_sessao,
        "criado_em": datetime.now().isoformat(timespec="seconds"),
        "nome_arquivo_origem": os.path.basename(nome_arquivo),
        "mes_referencia": mes,
        "erros": erros,
        "avisos": avisos,
        "registros": _registros_dataframe(df.reset_index(drop=True)),
    })
    return id_sessao


def carregar_sessao(id_sessao: str) -> Optional[dict]:
    destino, _ = _caminhos(id_sessao)
    if not os.path.isfile(destino):
        return None
    with open(destino, "r", encoding="utf-8") as arquivo:
        sessao = json.load(arquivo)
    if not isinstance(sessao, dict) or sessao.get("id_sessao") != id_sessao:
        raise ValueError("Arquivo de sessão inválido.")
    return sessao


def salvar_sessao(sessao: dict) -> None:
    id_sessao = _validar_id(sessao.get("id_sessao") or uuid.uuid4().hex)
    sessao = {**sessao, "id_sessao": id_sessao}
    with _lock_sessao(id_sessao):
        _gravar_sem_lock(sessao)


def atualizar_sessao(id_sessao: str,
                     mutador: Callable[[dict], None]) -> dict:
    _validar_id(id_sessao)
    with _lock_sessao(id_sessao):
        sessao = carregar_sessao(id_sessao)
        if sessao is None:
            raise FileNotFoundError(f"Sessão não encontrada: {id_sessao}")
        mutador(sessao)
        return _gravar_sem_lock(sessao)


def df_da_sessao(sessao: dict) -> pd.DataFrame:
    df = pd.DataFrame(_registros_da_sessao(sessao))
    if "data_abertura" not in df.columns:
        df["data_abertura"] = pd.NaT
    for coluna in ("date", "data_abertura", "data_solucao"):
        if coluna not in df.columns:
            df[coluna] = df.get("data_abertura", pd.NaT) if coluna == "date" else pd.NaT
        df[coluna] = pd.to_datetime(df[coluna], errors="coerce")

    for coluna in ("status", "priority", "type"):
        if coluna not in df.columns:
            df[coluna] = pd.NA
        valores = pd.to_numeric(df[coluna], errors="coerce")
        df[coluna] = valores.astype("Int64") if valores.isna().any() else valores.astype("int64")
    if "tempo_resolucao_horas" not in df.columns:
        df["tempo_resolucao_horas"] = float("nan")
    else:
        df["tempo_resolucao_horas"] = pd.to_numeric(
            df["tempo_resolucao_horas"], errors="coerce")

    df = gerador.preparar_dataframe(df)
    for coluna, padrao in (("fase", "backlog"), ("grupo_tecnico", ""),
                           ("observacoes", "")):
        if coluna not in df.columns:
            df[coluna] = padrao
        df[coluna] = df[coluna].fillna(padrao)
    if "incluir" not in df.columns:
        df["incluir"] = True
    else:
        df["incluir"] = df["incluir"].map(
            lambda valor: valor if isinstance(valor, bool)
            else str(valor).strip().lower() not in ("false", "0", "nao", "não", ""))
    return df


def limpar_sessoes_expiradas(dias: int = 7) -> int:
    os.makedirs(_PASTA_SESSOES, exist_ok=True)
    limite = datetime.now() - timedelta(days=max(0, int(dias)))
    removidas = 0
    for nome in os.listdir(_PASTA_SESSOES):
        if not nome.endswith(".json"):
            continue
        id_sessao = nome[:-5]
        if not _ID_RE.fullmatch(id_sessao):
            continue
        destino, _ = _caminhos(id_sessao)
        with _lock_sessao(id_sessao):
            sessao = carregar_sessao(id_sessao)
            if sessao is None:
                continue
            try:
                atualizado = datetime.fromisoformat(sessao["atualizado_em"])
                if atualizado.tzinfo is not None:
                    atualizado = atualizado.replace(tzinfo=None)
            except (KeyError, TypeError, ValueError):
                continue
            if atualizado >= limite:
                continue
            for caminho in (destino, destino + ".bak", destino + ".tmp",
                            destino + ".bak.tmp", _caminho_base(id_sessao)):
                if os.path.exists(caminho):
                    os.remove(caminho)
            _CACHE_BASE.pop(id_sessao, None)
            removidas += 1
    return removidas