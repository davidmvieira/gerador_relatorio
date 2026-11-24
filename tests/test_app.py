"""Testes do fluxo web: prévia fiel ao .docx e nenhuma gravação desnecessária.

Rodar na raiz do projeto:
    python -m unittest discover -s tests -v
"""

import io
import logging
import os
import re
import sys
import tempfile
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
# O app limpa sessões antigas ao ser importado: com TTL alto, nada é apagado.
os.environ["GLPI_SESSION_TTL_DAYS"] = "36500"

from docx import Document  # noqa: E402

import app as aplicacao  # noqa: E402
import gerar_relatorio_glpi as gerador  # noqa: E402
import sessoes_store  # noqa: E402

logging.getLogger("relatorio_glpi").setLevel(logging.ERROR)
ARQUIVO_REAL = os.path.join(RAIZ, "dados_202609301724.csv")


def _arquivos(pasta):
    return sorted(os.path.relpath(os.path.join(raiz, nome), pasta)
                  for raiz, _, nomes in os.walk(pasta) for nome in nomes)


def _texto_celulas_html(html, titulo_secao):
    """Linhas da primeira tabela após o título da seção, como listas de texto."""
    trecho = html.split(titulo_secao, 1)[1].split("</table>", 1)[0]
    linhas = re.findall(r"<tr>(.*?)</tr>", trecho, flags=re.S)
    return [[re.sub(r"\s+", " ", c).strip()
             for c in re.findall(r"<t[dh]>(.*?)</t[dh]>", linha, flags=re.S)]
            for linha in linhas]


@unittest.skipUnless(os.path.exists(ARQUIVO_REAL), "export real ausente")
class TestFluxoWeb(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.saidas = os.path.join(self.tmp.name, "saidas")
        os.makedirs(self.saidas)
        self.config_original = dict(gerador.CONFIG)
        gerador.CONFIG["pasta_saida"] = os.path.join(self.saidas, "execucoes")
        gerador.CONFIG["pasta_historico"] = os.path.join(self.saidas, "historico")
        sessoes_store.configurar_pasta(os.path.join(self.tmp.name, "sessoes"))
        aplicacao.app.config["TESTING"] = True
        self.cliente = aplicacao.app.test_client()
        with open(ARQUIVO_REAL, "rb") as arquivo:
            resposta = self.cliente.post(
                "/", data={"arquivo": (io.BytesIO(arquivo.read()), "dados.csv")},
                content_type="multipart/form-data")
        self.assertEqual(resposta.status_code, 302)
        self.sessao = resposta.headers["Location"].rsplit("/", 1)[-1]

    def tearDown(self):
        gerador.CONFIG.clear()
        gerador.CONFIG.update(self.config_original)
        self.tmp.cleanup()

    def test_mes_padrao_e_o_mais_recente(self):
        sessao = sessoes_store.carregar_sessao(self.sessao)
        self.assertEqual(sessao["mes_referencia"], "2026-09")

    def test_previa_completa_com_seletor_de_meses(self):
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-08") \
            .get_data(as_text=True)
        for esperado in ("ago/2026", "set/2026", "1. Painel Geral", "2. Distribuição",
                         "3. Frente Projetos", "4. SLA", "5. Capacidade",
                         "data:image/png;base64,", "Baixar DOCX de 2026-08"):
            self.assertIn(esperado, html)

    def test_previa_e_docx_mostram_os_mesmos_numeros(self):
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-08") \
            .get_data(as_text=True)
        painel_html = _texto_celulas_html(html, "1. Painel Geral")
        resposta = self.cliente.post(f"/gerar/{self.sessao}",
                                     data={"mes_referencia": "2026-08"})
        doc = Document(io.BytesIO(resposta.data))
        tabela = next(t for t in doc.tables
                      if t.rows[0].cells[0].text == "Indicador"
                      and t.rows[1].cells[0].text.startswith("Backlog inicial"))
        painel_docx = [[c.text for c in linha.cells] for linha in tabela.rows]
        self.assertEqual(painel_html, painel_docx)
        self.assertEqual(painel_docx[5][1:], ["8", "12"])  # backlog final de agosto

    def test_download_nao_grava_arquivos(self):
        resposta = self.cliente.post(f"/gerar/{self.sessao}",
                                     data={"mes_referencia": "2026-08"})
        self.assertEqual(resposta.status_code, 200)
        self.assertIn("Relatorio_Sistemas_2026-08.docx",
                      resposta.headers["Content-Disposition"])
        self.assertTrue(resposta.data.startswith(b"PK"))
        self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-09")
        self.assertEqual(_arquivos(self.saidas), [])

    def test_registrar_historico_grava_so_o_json_do_mes(self):
        resposta = self.cliente.post(f"/historico/{self.sessao}",
                                     data={"mes_referencia": "2026-08"})
        self.assertEqual(resposta.status_code, 302)
        self.assertEqual(_arquivos(self.saidas), [os.path.join("historico", "2026-08.json")])
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-09") \
            .get_data(as_text=True)
        self.assertIn("Comparação com 2026-08", html)

    def test_mes_sem_chamados(self):
        resposta = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2020-01")
        self.assertEqual(resposta.status_code, 200)
        self.assertIn("Nenhum chamado da planilha passou por 2020-01",
                      resposta.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()
