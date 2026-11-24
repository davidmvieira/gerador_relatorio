"""Testes do núcleo de métricas (point-in-time).

Rodar na raiz do projeto:
    python -m unittest discover -s tests -v
"""

import logging
import os
import sys
import unittest

import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import gerar_relatorio_glpi as gerador  # noqa: E402
import metricas  # noqa: E402

logging.getLogger("relatorio_glpi").setLevel(logging.ERROR)

AGORA = pd.Timestamp("2026-09-30 23:59")
SUST = gerador.CONFIG["nome_frente_sustentacao"]
PROJ = gerador.CONFIG["nome_frente_projetos"]


def _ticket(id_, abertura, solucao=None, status=2, tipo=1, frente=SUST, fechamento=None):
    return {
        "id": id_, "date": pd.Timestamp(abertura), "data_abertura": pd.Timestamp(abertura),
        "data_solucao": pd.Timestamp(solucao) if solucao else pd.NaT,
        "data_fechamento": pd.Timestamp(fechamento) if fechamento else pd.NaT,
        "status": status, "type": tipo, "frente": frente,
    }


def base_sintetica():
    """Cenário de agosto/2026 com todos os casos de borda relevantes."""
    return pd.DataFrame([
        # A: herdado de julho, solucionado em agosto (384 h)
        _ticket("A", "2026-07-20", "2026-08-05", status=6),
        # B: aberto em agosto, ainda aberto
        _ticket("B", "2026-08-03", tipo=2),
        # C: aberto e solucionado em agosto (12 h)
        _ticket("C", "2026-08-10 08:00", "2026-08-10 20:00", status=5, tipo=2),
        # D: aberto em SETEMBRO e ainda aberto — não pode entrar em agosto
        _ticket("D", "2026-09-02"),
        # E: aberto em julho, solucionado só em setembro — backlog no início e no fim
        _ticket("E", "2026-07-01", "2026-09-10", status=6),
        # F: aberto e solucionado em junho — irrelevante
        _ticket("F", "2026-06-01", "2026-06-05", status=6),
        # G: reaberto (status em aberto, mas com data de solução) — conta como aberto
        _ticket("G", "2026-08-15", "2026-08-20", status=2),
        # H: solucionado sem data de solução, mas com fechamento (120 h)
        _ticket("H", "2026-08-20", None, status=5, fechamento="2026-08-25"),
    ])


class TestMarcarPeriodo(unittest.TestCase):
    def setUp(self):
        self.m = metricas.marcar_periodo(base_sintetica(), 2026, 8, agora=AGORA)
        self.flags = self.m.set_index("id")

    def test_chamado_aberto_depois_do_mes_nao_entra(self):
        self.assertFalse(self.flags.loc["D", "relevante"])

    def test_chamado_de_mes_anterior_ja_resolvido_nao_entra(self):
        self.assertFalse(self.flags.loc["F", "relevante"])

    def test_herdado_resolvido_depois_segue_no_backlog(self):
        self.assertTrue(self.flags.loc["E", "backlog_inicio"])
        self.assertTrue(self.flags.loc["E", "backlog_fim"])
        self.assertFalse(self.flags.loc["E", "solucionado_no_mes"])

    def test_reaberto_e_tratado_como_aberto(self):
        self.assertFalse(self.flags.loc["G", "solucionado_no_mes"])
        self.assertTrue(self.flags.loc["G", "backlog_fim"])

    def test_aviso_de_backlog_inicial_incompleto(self):
        self.assertFalse(self.m.attrs["backlog_inicial_incompleto"])  # base tem junho
        so_agosto = base_sintetica().iloc[[1, 2]]
        m = metricas.marcar_periodo(so_agosto, 2026, 8, agora=AGORA)
        self.assertTrue(m.attrs["backlog_inicial_incompleto"])

    def test_fechamento_supre_solucao_ausente(self):
        self.assertTrue(self.flags.loc["H", "solucionado_no_mes"])
        self.assertAlmostEqual(self.flags.loc["H", "tempo_resolucao_h"], 120.0)


class TestIndicadores(unittest.TestCase):
    def setUp(self):
        m = metricas.marcar_periodo(base_sintetica(), 2026, 8, agora=AGORA)
        self.ind = metricas.indicadores(m[m["relevante"]])

    def test_fluxo_point_in_time(self):
        self.assertEqual(self.ind["backlog_inicial"], 2)  # A, E
        self.assertEqual(self.ind["abertos"], 4)          # B, C, G, H
        self.assertEqual(self.ind["solucionados"], 3)     # A, C, H
        self.assertEqual(self.ind["backlog_final"], 3)    # B, E, G
        self.assertEqual(self.ind["saldo"], 1)
        self.assertTrue(self.ind["fluxo_confere"])

    def test_absorcao_usa_backlog_mais_abertos(self):
        self.assertEqual(self.ind["absorcao_pct"], 50.0)  # 3 / (2 + 4)

    def test_tempos_media_e_mediana(self):
        self.assertEqual(self.ind["tmr_mediana_h"], 120.0)  # 12, 120, 384
        self.assertEqual(self.ind["tmr_media_h"], 172.0)

    def test_natureza_considera_so_abertos_no_mes(self):
        self.assertEqual(self.ind["requisicoes"], 2)  # B, C
        self.assertEqual(self.ind["incidentes"], 2)   # G, H

    def test_frentes_calculadas_separadamente(self):
        df = pd.concat([base_sintetica(), pd.DataFrame([
            _ticket("P1", "2026-08-01", frente=PROJ),
            _ticket("P2", "2026-08-02", frente=PROJ)])], ignore_index=True)
        m = metricas.marcar_periodo(df, 2026, 8, agora=AGORA)
        proj = metricas.indicadores(m[m["relevante"]], PROJ)
        sust = metricas.indicadores(m[m["relevante"]], SUST)
        self.assertEqual(proj["abertos"], 2)
        self.assertEqual(proj["absorcao_pct"], 0.0)
        self.assertEqual(sust["absorcao_pct"], 50.0)


class TestOutliers(unittest.TestCase):
    def test_media_sem_outliers(self):
        horas = pd.Series([10, 11, 12, 13, 14, 10_000])
        est = metricas.estatisticas_tempo(horas, iqr_mult=8.0)
        self.assertEqual(est["n_outliers"], 1)
        self.assertEqual(est["media_sem_outliers_h"], 12.0)
        self.assertEqual(est["mediana_h"], 12.5)

    def test_amostra_pequena_nao_remove_nada(self):
        est = metricas.estatisticas_tempo(pd.Series([1, 1000]))
        self.assertEqual(est["n_outliers"], 0)

    def test_sem_dados(self):
        self.assertIsNone(metricas.estatisticas_tempo(pd.Series([], dtype=float))["media_h"])


class TestMesesEInconsistencias(unittest.TestCase):
    def test_meses_disponiveis(self):
        meses = metricas.meses_disponiveis(base_sintetica()).set_index("mes")
        self.assertEqual(list(meses.index), ["2026-06", "2026-07", "2026-08", "2026-09"])
        self.assertEqual(meses.loc["2026-08", "abertos"], 4)
        self.assertEqual(meses.loc["2026-08", "solucionados"], 3)
        self.assertEqual(meses.loc["2026-09", "abertos"], 1)

    def test_inconsistencias(self):
        df = pd.concat([base_sintetica(), pd.DataFrame([
            _ticket("X", "2026-08-10", "2026-08-01", status=5),
            _ticket("Y", "2026-08-10", None, status=6)])], ignore_index=True)
        problemas = metricas.inconsistencias(df).groupby("id")["problema"].first()
        self.assertIn("anterior à abertura", problemas["X"])
        self.assertIn("sem data de solução", problemas["Y"])
        self.assertIn("reaberto", problemas["G"])
        self.assertNotIn("H", problemas.index)  # fechamento resolve


class TestIntegracaoGerador(unittest.TestCase):
    def test_painel_do_gerador_usa_o_nucleo(self):
        df = base_sintetica().assign(entidade="Sistemas", name="t", categoria="c",
                                     priority=3)
        df = gerador.preparar_dataframe(df, agora=AGORA)
        sub = gerador.filtrar_periodo_referencia(df, 2026, 8, agora=AGORA)
        self.assertNotIn("D", set(sub["id"]))
        painel = gerador.metricas_painel(sub, SUST)
        self.assertEqual((painel["backlog_inicial"], painel["novos"],
                          painel["entregues"], painel["backlog_final"]), (2, 4, 3, 3))
        self.assertEqual(painel["taxa_resolucao"], 50.0)


ARQUIVO_REAL = os.path.join(RAIZ, "dados_202609301724.csv")


@unittest.skipUnless(os.path.exists(ARQUIVO_REAL), "export real ausente")
class TestExportReal(unittest.TestCase):
    """Export com agosto e setembro/2026 na mesma planilha."""

    @classmethod
    def setUpClass(cls):
        bruto = gerador.carregar_dados(ARQUIVO_REAL)
        cls.m = gerador.preparar_dataframe(bruto, agora=AGORA)
        cls.ago = gerador.filtrar_periodo_referencia(cls.m, 2026, 8, agora=AGORA)
        cls.set = gerador.filtrar_periodo_referencia(cls.m, 2026, 9, agora=AGORA)

    def test_setembro_nao_vaza_para_agosto(self):
        self.assertEqual(int((self.ago["date"] >= "2026-09-01").sum()), 0)

    def test_backlog_final_de_agosto_point_in_time(self):
        # Antes da correção o painel mostrava 26 (status no dia do export).
        self.assertEqual(int(self.ago["backlog_fim"].sum()), 20)

    def test_continuidade_entre_meses(self):
        for frente in (SUST, PROJ):
            ago = gerador.metricas_painel(self.ago, frente)
            set_ = gerador.metricas_painel(self.set, frente)
            self.assertEqual(ago["backlog_final"], set_["backlog_inicial"], frente)
            self.assertTrue(ago["fluxo_confere"] and set_["fluxo_confere"], frente)

    def test_absorcao_nunca_passa_de_100(self):
        for sub in (self.ago, self.set):
            for frente in (SUST, PROJ):
                absorcao = gerador.metricas_painel(sub, frente)["absorcao_pct"]
                if absorcao is not None:
                    self.assertLessEqual(absorcao, 100.0)


if __name__ == "__main__":
    unittest.main()
