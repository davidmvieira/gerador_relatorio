"""Testes do SLA institucional, calendário de dias úteis e configuração.

Rodar na raiz do projeto:
    python -m unittest discover -s tests -v
"""

import copy
import datetime as dt
import json
import logging
import os
import sys
import tempfile
import unittest

import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import gerar_relatorio_glpi as gerador  # noqa: E402
import metricas  # noqa: E402

logging.getLogger("relatorio_glpi").setLevel(logging.ERROR)

METAS = {"P1": {"valor": 4, "unidade": "horas"},
         "P3": {"valor": 3, "unidade": "dias_uteis"}}
FERIADOS_2026 = metricas.calendario_feriados([2026])


class TestCalendario(unittest.TestCase):
    def test_feriados_nacionais_2026(self):
        feriados = metricas.feriados_nacionais(2026)
        self.assertIn(dt.date(2026, 4, 3), feriados)    # Sexta-feira Santa (Páscoa 05/04)
        self.assertIn(dt.date(2026, 11, 20), feriados)  # Consciência Negra
        self.assertIn(dt.date(2026, 9, 7), feriados)
        self.assertNotIn(dt.date(2026, 2, 17), feriados)  # Carnaval é ponto facultativo

    def test_dias_uteis_no_mes(self):
        self.assertEqual(metricas.dias_uteis_no_mes(2026, 8, FERIADOS_2026), 21)
        self.assertEqual(metricas.dias_uteis_no_mes(2026, 11, FERIADOS_2026), 19)

    def test_feriados_extras(self):
        cal = metricas.calendario_feriados([2026], extras=["2026-02-17", "invalida"])
        self.assertIn(dt.date(2026, 2, 17), cal)


class TestPrazoSla(unittest.TestCase):
    def _prazo(self, abertura, classe):
        return metricas.prazo_sla(pd.Series([pd.Timestamp(abertura)]), pd.Series([classe]),
                                  METAS, FERIADOS_2026).iloc[0]

    def test_horas_corridas(self):
        self.assertEqual(self._prazo("2026-08-07 22:00", "P1"), pd.Timestamp("2026-08-08 02:00"))

    def test_dias_uteis_a_partir_de_sexta(self):
        self.assertEqual(self._prazo("2026-08-07 10:00", "P3"),
                         pd.Timestamp("2026-08-12 23:59:59"))

    def test_aberto_no_sabado_conta_a_partir_de_segunda(self):
        self.assertEqual(self._prazo("2026-08-08 10:00", "P3"),
                         pd.Timestamp("2026-08-12 23:59:59"))

    def test_pula_feriado(self):  # 07/09/2026 é segunda-feira e feriado
        self.assertEqual(self._prazo("2026-09-04 18:00", "P3"),
                         pd.Timestamp("2026-09-10 23:59:59"))

    def test_classe_sem_meta_fica_sem_prazo(self):
        self.assertTrue(pd.isna(self._prazo("2026-08-07", "P9")))


class TestAvaliarSla(unittest.TestCase):
    def test_no_prazo_e_vencidos_em_aberto(self):
        df = pd.DataFrame([
            # P3 solucionado em 2 dias úteis: no prazo
            {"id": 1, "date": "2026-08-03 09:00", "data_solucao": "2026-08-05 09:00",
             "status": 5, "priority": 3},
            # P3 solucionado depois do prazo (prazo 06/08)
            {"id": 2, "date": "2026-08-03 09:00", "data_solucao": "2026-08-10 09:00",
             "status": 5, "priority": 3},
            # P1 ainda aberto no fim do mês, prazo de 4 h há muito vencido
            {"id": 3, "date": "2026-08-20 09:00", "data_solucao": None,
             "status": 2, "priority": 6},
            # sem prioridade mapeada
            {"id": 4, "date": "2026-08-20 09:00", "data_solucao": None,
             "status": 2, "priority": None},
        ])
        df["data_abertura"] = pd.to_datetime(df["date"])
        df["date"] = df["data_abertura"]
        df["data_solucao"] = pd.to_datetime(df["data_solucao"])
        m = metricas.marcar_periodo(df, 2026, 8, agora=pd.Timestamp("2026-09-30"))
        mapa = {6: "P1", 3: "P3"}
        avaliado = metricas.avaliar_sla(m, mapa, METAS, FERIADOS_2026)
        resumo = metricas.resumo_sla(avaliado, METAS).set_index("classe")
        self.assertEqual(resumo.loc["P3", "solucionados"], 2)
        self.assertEqual(resumo.loc["P3", "no_prazo"], 1)
        self.assertEqual(resumo.loc["P3", "pct_no_prazo"], 50.0)
        self.assertEqual(resumo.loc["P1", "vencidos_em_aberto"], 1)
        self.assertEqual(resumo.attrs["sem_classe"], 1)


class TestConfiguracao(unittest.TestCase):
    def setUp(self):
        self.original = copy.deepcopy(gerador.CONFIG)
        self.tmp = tempfile.TemporaryDirectory()
        self.caminho = os.path.join(self.tmp.name, "config.json")

    def tearDown(self):
        gerador.CONFIG.clear()
        gerador.CONFIG.update(self.original)
        self.tmp.cleanup()

    def _valida(self, **mudancas):
        cfg = gerador.config_editavel()
        cfg.update(mudancas)
        return cfg

    def test_padrao_e_valido(self):
        self.assertEqual(gerador.validar_config(gerador.config_editavel()), [])

    def test_rejeita_valores_invalidos(self):
        cfg = self._valida()
        cfg["sla"]["mapa_prioridade"][3] = "P7"
        cfg["sla"]["metas"]["P2"]["unidade"] = "semanas"
        cfg["calendario"]["feriados_extras"] = ["31/12/2026"]
        cfg["premissas_equipe"]["analistas"] = 0
        erros = gerador.validar_config(cfg)
        self.assertEqual(len(erros), 4, erros)
        self.assertEqual(gerador.salvar_config(cfg, self.caminho), erros)
        self.assertFalse(os.path.exists(self.caminho))

    def test_salvar_e_carregar(self):
        cfg = self._valida(autor="Equipe de Sistemas")
        cfg["sla"]["metas"]["P3"] = {"valor": 1, "unidade": "dias_uteis"}
        cfg["premissas_equipe"]["dias_uteis_mes"] = ""
        self.assertEqual(gerador.salvar_config(cfg, self.caminho), [])
        with open(self.caminho, encoding="utf-8") as arquivo:
            gravado = json.load(arquivo)
        self.assertEqual(gravado["sla"]["mapa_prioridade"]["3"], "P3")
        self.assertNotIn("pasta_saida", gravado)

        gerador.CONFIG.clear()
        gerador.CONFIG.update(copy.deepcopy(self.original))
        self.assertEqual(gerador.carregar_config(self.caminho), [])
        self.assertEqual(gerador.CONFIG["autor"], "Equipe de Sistemas")
        self.assertEqual(gerador.CONFIG["sla"]["metas"]["P3"]["valor"], 1.0)
        self.assertEqual(gerador.CONFIG["sla"]["mapa_prioridade"][3], "P3")
        self.assertIsNone(gerador.CONFIG["premissas_equipe"]["dias_uteis_mes"])

    def test_arquivo_invalido_e_ignorado(self):
        with open(self.caminho, "w", encoding="utf-8") as arquivo:
            arquivo.write("{quebrado")
        self.assertTrue(gerador.carregar_config(self.caminho))
        self.assertEqual(gerador.CONFIG["autor"], self.original["autor"])

    def test_meta_mais_rigida_reduz_cumprimento(self):
        df = pd.DataFrame([{"id": 1, "date": "2026-08-03 09:00",
                            "data_solucao": "2026-08-05 09:00", "status": 5, "priority": 3,
                            "entidade": "Sistemas", "name": "t", "categoria": "c", "type": 1}])
        df["data_solucao"] = pd.to_datetime(df["data_solucao"])
        df["date"] = pd.to_datetime(df["date"])
        df["data_abertura"] = df["date"]
        m = gerador.preparar_dataframe(df, agora=pd.Timestamp("2026-09-30"))
        sub = gerador.filtrar_periodo_referencia(m, 2026, 8, agora=pd.Timestamp("2026-09-30"))
        frente = gerador.CONFIG["nome_frente_sustentacao"]
        self.assertEqual(gerador.pct_no_prazo(gerador.metricas_sla(sub, frente)), 100.0)
        cfg = self._valida()
        cfg["sla"]["metas"]["P3"] = {"valor": 1, "unidade": "dias_uteis"}
        gerador.aplicar_config(cfg)
        self.assertEqual(gerador.pct_no_prazo(gerador.metricas_sla(sub, frente)), 0.0)


if __name__ == "__main__":
    unittest.main()
