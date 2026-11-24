"""Testes do fluxo web: prévia fiel ao .docx e nenhuma gravação desnecessária.

Rodar na raiz do projeto:
    python -m unittest discover -s tests -v
"""

import io
import json
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
# Configuração de teste isolada (nunca lê nem grava o instance/config.json real).
_CONFIG_TMP = tempfile.TemporaryDirectory()
os.environ["GLPI_CONFIG"] = os.path.join(_CONFIG_TMP.name, "config.json")

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

    # ---------------- Fase 3: série da própria planilha ----------------
    def test_comparacao_com_mes_anterior_da_planilha_sem_registrar(self):
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-09") \
            .get_data(as_text=True)
        self.assertIn("Comparação com 2026-08", html)
        serie = _texto_celulas_html(html, "6. Série Histórica")
        self.assertEqual(serie[1], ["2026-08 *", "127", "119", "8", "93.7%", "17.8 h",
                                    "planilha"])
        self.assertEqual(serie[2][0], "2026-09")
        self.assertEqual(serie[2][-1], "este relatório")
        self.assertEqual(_arquivos(self.saidas), [])

    def test_mes_registrado_prevalece_sobre_a_planilha(self):
        self.cliente.post(f"/historico/{self.sessao}", data={"mes_referencia": "2026-08"})
        caminho = os.path.join(gerador.CONFIG["pasta_historico"], "2026-08.json")
        with open(caminho, encoding="utf-8") as arquivo:
            registro = json.load(arquivo)
        self.assertNotIn("fonte", registro)
        registro["sustentacao"]["novos"] = 999
        with open(caminho, "w", encoding="utf-8") as arquivo:
            json.dump(registro, arquivo)
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-09") \
            .get_data(as_text=True)
        serie = _texto_celulas_html(html, "6. Série Histórica")
        self.assertEqual(serie[1][1], "999")
        self.assertEqual(serie[1][-1], "histórico registrado")

    def test_serie_nao_mostra_meses_posteriores(self):
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-08") \
            .get_data(as_text=True)
        self.assertNotIn("2026-09", html.split("6. Série Histórica", 1)[1])

    # ---------------- Fase 3: sessão leve ----------------
    def test_sessao_guarda_base_uma_vez_e_so_edicoes(self):
        pasta = os.path.join(self.tmp.name, "sessoes")
        meta = os.path.join(pasta, f"{self.sessao}.json")
        base = os.path.join(pasta, f"{self.sessao}.base.json")
        mtime_base = os.path.getmtime(base)
        self.assertLess(os.path.getsize(meta), 3_000)
        self.assertGreater(os.path.getsize(base), 20_000)

        resposta = self.cliente.post(f"/revisao/{self.sessao}/salvar", data={
            "categoria_0": "Categoria Editada", "incluir_0": "on",
            "frente": "sustentacao", "pagina": "1"})
        self.assertEqual(resposta.status_code, 302)
        self.cliente.get(f"/revisao/{self.sessao}?mes_referencia=2026-08")

        self.assertEqual(os.path.getmtime(base), mtime_base)
        self.assertLess(os.path.getsize(meta), 3_000)
        sessao = sessoes_store.carregar_sessao(self.sessao)
        self.assertNotIn("registros", sessao)
        self.assertEqual(sessao["edicoes"]["0"], {"categoria": "Categoria Editada"})
        self.assertEqual(sessao["mes_referencia"], "2026-08")
        df = sessoes_store.df_da_sessao(sessao)
        self.assertEqual(df.loc[0, "categoria"], "Categoria Editada")

    def test_upload_nao_grava_o_arquivo_enviado(self):
        uploads = os.path.join(aplicacao.app.instance_path, "uploads")
        antes = _arquivos(uploads) if os.path.isdir(uploads) else []
        with open(ARQUIVO_REAL, "rb") as arquivo:
            self.cliente.post("/", data={"arquivo": (io.BytesIO(arquivo.read()), "x.csv")},
                              content_type="multipart/form-data")
        depois = _arquivos(uploads) if os.path.isdir(uploads) else []
        self.assertEqual(antes, depois)

    # ---------------- Fase 4: SLA institucional e configuração ----------------
    def test_previa_mostra_sla_institucional(self):
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-08") \
            .get_data(as_text=True)
        sla = _texto_celulas_html(html, "4. SLA Institucional")
        self.assertEqual(sla[0][:3], ["Classe", "Prioridades GLPI", "Prazo de solução"])
        self.assertEqual(sla[3][:6], ["P3", "Média", "3 dias úteis", "107", "103", "96.3%"])
        self.assertEqual(sla[-1][0], "Total")
        self.assertIn("21 dias úteis no mês (pelo calendário do mês)", html)

    def test_configuracao_salva_e_afeta_o_relatorio(self):
        self.assertEqual(self.cliente.get("/configuracao").status_code, 200)
        formulario = {
            "titulo_relatorio": "Relatório de Teste", "subtitulo": "", "autor": "Equipe",
            "entidade_sustentacao": "Sistemas", "max_categorias_detalhadas": "10",
            "outlier_iqr_mult_tmr": "8", "analistas": "6", "horas_por_dia": "8",
            "dias_uteis_mes": "", "dias_ferias_no_mes": "0", "feriados_nacionais": "on",
            "feriados_extras": "2026-08-14",
            **{f"prioridade_{p}": c for p, c in
               {6: "P1", 5: "P1", 4: "P2", 3: "P3", 2: "P4", 1: "P4"}.items()},
            **{f"meta_{c}_valor": v for c, v in
               {"P1": "4", "P2": "8", "P3": "1", "P4": "5"}.items()},
            **{f"meta_{c}_unidade": u for c, u in
               {"P1": "horas", "P2": "horas", "P3": "dias_uteis",
                "P4": "dias_uteis"}.items()},
            "voltar": "//site-externo.com",
        }
        resposta = self.cliente.post("/configuracao", data=formulario)
        self.assertEqual(resposta.status_code, 302)
        self.assertEqual(resposta.headers["Location"], "/configuracao")  # sem redirect externo
        self.assertTrue(os.path.exists(os.environ["GLPI_CONFIG"]))
        html = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2026-08") \
            .get_data(as_text=True)
        self.assertIn("1 dias úteis", html)
        self.assertIn("20 dias úteis no mês", html)  # 14/08 virou feriado extra

    def test_configuracao_invalida_nao_grava(self):
        caminho = os.environ["GLPI_CONFIG"]
        if os.path.exists(caminho):
            os.remove(caminho)
        resposta = self.cliente.post("/configuracao", data={
            "titulo_relatorio": "", "entidade_sustentacao": "Sistemas",
            "feriados_extras": "31/12/2026"})
        self.assertEqual(resposta.status_code, 200)
        self.assertIn("Não foi possível salvar", resposta.get_data(as_text=True))
        self.assertFalse(os.path.exists(caminho))

    def test_mes_sem_chamados(self):
        resposta = self.cliente.get(f"/preview/{self.sessao}?mes_referencia=2020-01")
        self.assertEqual(resposta.status_code, 200)
        self.assertIn("Nenhum chamado da planilha passou por 2020-01",
                      resposta.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()
