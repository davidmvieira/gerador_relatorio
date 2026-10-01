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


def _normalizar_sessao(sessao: dict) -> dict:
    agora = datetime.now().isoformat(timespec="seconds")
    payload = dict(sessao)
    if "registros" not in payload and isinstance(payload.get("df"), pd.DataFrame):
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
    payload.setdefault("resultado", None)

    registros = []
    for original in payload.get("registros", []):
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
    payload["registros"] = registros
    payload["erros"] = [str(item) for item in payload["erros"]]
    payload["avisos"] = [str(item) for item in payload["avisos"]]
    resultado = payload.get("resultado")
    if resultado is not None:
        payload["resultado"] = {
            "caminho": str(resultado.get("caminho", "")),
            "periodo": resultado.get("periodo", {}),
            "execucao": str(resultado.get("execucao", "")),
        }
    return payload


def _gravar_sem_lock(sessao: dict) -> dict:
    payload = _normalizar_sessao(sessao)
    destino, _ = _caminhos(payload["id_sessao"])
    temporario = destino + ".tmp"
    backup = destino + ".bak"
    backup_tmp = backup + ".tmp"
    try:
        with open(temporario, "w", encoding="utf-8", newline="\n") as arquivo:
            json.dump(payload, arquivo, ensure_ascii=False, indent=2, allow_nan=False)
            arquivo.flush()
            os.fsync(arquivo.fileno())
        if os.path.isfile(destino):
            shutil.copy2(destino, backup_tmp)
            os.replace(backup_tmp, backup)
        os.replace(temporario, destino)
    finally:
        for caminho in (temporario, backup_tmp):
            if os.path.exists(caminho):
                os.remove(caminho)
    return payload


def criar_sessao(df: pd.DataFrame, nome_arquivo: str, erros: list,
                 avisos: list) -> str:
    datas = pd.to_datetime(df.get("date"), errors="coerce").dropna()
    mes = (pd.Timestamp(datas.median()).strftime("%Y-%m") if not datas.empty
           else datetime.now().strftime("%Y-%m"))
    id_sessao = uuid.uuid4().hex
    salvar_sessao({
        "id_sessao": id_sessao,
        "criado_em": datetime.now().isoformat(timespec="seconds"),
        "nome_arquivo_origem": os.path.basename(nome_arquivo),
        "mes_referencia": mes,
        "erros": erros,
        "avisos": avisos,
        "resultado": None,
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
    payload = _normalizar_sessao(sessao)
    with _lock_sessao(payload["id_sessao"]):
        _gravar_sem_lock(payload)


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
    df = pd.DataFrame(sessao.get("registros", []))
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
                            destino + ".bak.tmp"):
                if os.path.exists(caminho):
                    os.remove(caminho)
            removidas += 1
    return removidas