# Gerador de Relatório GLPI — Sustentação x Projetos

Gera o relatório gerencial mensal (`.docx`) da Equipe de Sistemas a partir de um
export do GLPI (CSV ou Excel). As métricas de cada mês são reconstruídas pelas
datas de abertura e solução (point-in-time), separando sempre as frentes
Sustentação e Projetos.

## Instalação

```bash
python -m pip install -r requirements.txt
```

## Uso pelo app (recomendado)

```bash
python app.py
```

Abra `http://127.0.0.1:5000` e siga o fluxo:

1. **Upload** do export (`.csv`, `.xlsx` ou `.xlsm`). O arquivo é lido em
   memória e não é gravado em disco.
2. **Revisão**: corrija categoria, status, prioridade, datas e fase dos
   projetos, inclua ou exclua chamados (individualmente ou em lote).
3. **Prévia** do relatório completo para o mês escolhido, idêntica ao DOCX.
4. **Download** do `.docx`, montado em memória.

Opcionalmente, **registre o mês no histórico** pela prévia
(`relatorios_glpi/historico/AAAA-MM.json`, só métricas agregadas). Meses
anteriores presentes na própria planilha já entram na comparação e na série
sem precisar de registro.

As sessões de revisão ficam em `instance/sessoes/` (base gravada uma vez +
edições) e são apagadas na inicialização após 7 dias sem uso. Ajuste o prazo
com a variável `GLPI_SESSION_TTL_DAYS`.

## Uso pela linha de comando

```bash
python gerar_relatorio_glpi.py caminho/do/export.csv --mes 2026-09
```

Sem `--mes`, usa o mês corrente. O DOCX é gravado em
`relatorios_glpi/execucoes/<timestamp>/` e o mês é registrado no histórico.
Para Excel com duas abas, os nomes das abas estão em `CONFIG["aba_principal"]`
e `CONFIG["aba_detalhe"]`.

## Configuração

Os parâmetros ficam em `instance/config.json` e são editados na tela
`/configuracao` do app (a linha de comando usa o mesmo arquivo). Sem esse
arquivo, valem os padrões do `CONFIG` em `gerar_relatorio_glpi.py`.

- **Cabeçalho e escopo**: título, subtítulo, autor e a entidade da frente
  Sustentação. As demais entidades formam a frente Projetos.
- **SLA institucional**: classe (P1–P4) de cada prioridade do GLPI e o prazo de
  cada classe, em horas corridas ou em dias úteis. O mapeamento padrão é
  provisório e deve ser validado com a gestão. O SLA de resposta não é medido,
  porque o export não traz a data do primeiro atendimento.
- **Calendário**: feriados nacionais calculados automaticamente; datas extras
  informadas uma por linha (AAAA-MM-DD).
- **Capacidade**: analistas, horas por dia e dias de férias no mês.
- **Outlier**: multiplicador do IQR para o TMR (padrão 8).

## Dados e privacidade (LGPD)

Títulos de chamados podem conter nomes de pessoas. O `.gitignore` mantém fora do
repositório os exports (`dados_*.csv`, `*.xlsx`), o estado local do app
(`instance/`) e tudo o que é gerado em `relatorios_glpi/`. Nunca versione esses
arquivos.

## Testes

```bash
python -m unittest discover -s tests
```

Os testes que usam o export real (`dados_202609301724.csv`, na raiz) só rodam
quando esse arquivo existe na máquina; os demais usam dados sintéticos.

## Arquivos

- `app.py` — app Flask (upload, revisão, prévia, download, configuração).
- `metricas.py` — fonte única dos cálculos por data.
- `gerar_relatorio_glpi.py` — leitura do export, tabelas e textos do relatório,
  gráficos, DOCX e CLI. A prévia e o DOCX usam as mesmas funções.
- `sessoes_store.py` — persistência das sessões de revisão.
- `templates/`, `static/app.css` — interface.
