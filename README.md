# Gerador Autônomo de Relatório GLPI — Sustentação x Projetos

Script Python único que lê o export mensal do GLPI (Excel, uma ou duas abas) e gera
o relatório `.docx` completo, já com:

- Painel geral (Sustentação x Projetos x Total)
- Distribuição por categoria (por frente, com normalização de grafia)
- Detalhamento item a item da frente Projetos (título, status, datas, horas)
- SLA por prioridade (frente Sustentação)
- Detecção automática de outliers de tempo de resolução (regra do IQR)
- Capacidade da equipe / aderência (HMM, HE, HPC, HHA, AD%)
- Detecção de um padrão recorrente por palavra-chave no título (ex.: um bug
  que aparece repetidas vezes) com gráfico de volume semanal
- Histórico mensal incremental com comparação, tendência e projeção estatística
- Todos os gráficos (matplotlib) embutidos no próprio .docx

## Instalação

```bash
pip install -r requirements.txt --break-system-packages
```

(remova `--break-system-packages` se estiver usando um virtualenv)

## Uso

```bash
python3 gerar_relatorio_glpi.py caminho/para/base_de_dados.xlsx
# ou
python3 gerar_relatorio_glpi.py caminho/para/base_de_dados.csv
```

Se você não passar nenhum argumento, ele procura por `base_de_dados.xlsx` na
pasta atual.

Para usar a interface web:

```bash
python -m pip install -r requirements.txt
python app.py
```

Abra `http://127.0.0.1:5000`, envie o CSV/Excel, revise os dados, salve as
alterações e confirme a geração. As execuções ficam em
`relatorios_glpi/execucoes/YYYYMMDD_HHMMSS/`.

O resultado fica em uma pasta própria dentro de
`./relatorios_glpi/execucoes/`, nomeada com o timestamp da execução, por
exemplo: `relatorios_glpi/execucoes/2026-09-16_21-55-03_123456/`.

- `Relatorio_Sistemas.docx` — o relatório final
- `01_donut_natureza.png`, `02_...png` etc. — os gráficos gerados (também já
  embutidos no `.docx`, ficam na pasta da execução caso você queira reusar em
  outro lugar)

A estrutura principal fica assim:

```text
relatorios_glpi/
├── historico/   # métricas mensais consolidadas
└── execucoes/   # DOCX e gráficos de cada geração
```

As métricas consolidadas de cada mês ficam em
`./relatorios_glpi/historico/YYYY-MM.json`. Esses arquivos não armazenam
chamados individuais. O mesmo mês é substituído quando o relatório é gerado
novamente.

Com pelo menos dois meses, o relatório mostra a comparação e os gráficos de
evolução. Com pelo menos três meses, mostra tendência e uma estimativa do mês
seguinte pela média móvel simples dos três últimos meses. Lacunas entre meses
são mantidas e identificadas na comparação.

## Formato esperado do Excel

O script aceita o padrão antigo de duas abas, um arquivo Excel com uma única
aba de detalhes ou um arquivo CSV.

**Aba 1** (fonte confiável de id/categoria/status/prioridade/entidade):
`id, name, date, itilcategories_id, categoria, demanda, status, priority, type, entities_id, entidade`

**Aba 2** (detalhe com datas de solução e tempo de resolução):
`id_chamado, titulo, entities_id, entidade, categoria, tipo_chamado, prioridade, status_atual, data_abertura, data_solucao, data_fechamento, tempo_resolucao_horas`

Quando houver apenas uma aba, ela deve conter as colunas da Aba 2 acima. O
script converte automaticamente os nomes e os valores de status, prioridade e
tipo de chamado para o formato interno do relatório.

O CSV deve conter as mesmas colunas da Aba 2. O separador (vírgula ou ponto e
vírgula) é detectado automaticamente, assim como as codificações UTF-8 e
Latin-1.

O arquivo [base_ficticia_agosto_2026.csv](base_ficticia_agosto_2026.csv) pode
ser usado para validar a execução sem dados reais:

```bash
python gerar_relatorio_glpi.py base_ficticia_agosto_2026.csv
```

Se os nomes das abas no seu export forem diferentes, ajuste
`CONFIG["aba_principal"]` e `CONFIG["aba_detalhe"]` no topo do script.

## O que ajustar todo mês (seção CONFIG no topo do arquivo)

- `entidade_sustentacao`: nome exato da entidade GLPI que representa a
  sustentação corrente (hoje: `"Sistemas"`). Tudo que não for essa entidade
  vira automaticamente a frente "Projetos".
- `termos_chave_recorrencia` / `nome_recorrencia`: troque pelas palavras-chave
  do problema mais quente do mês (ex.: nome de um sistema com bug recorrente).
  Se não houver nada relevante, deixe a lista vazia (`[]`) que a seção 6 some
  sozinha do relatório.
- `premissas_equipe`: número de analistas, jornada, dias úteis do mês e dias
  de férias/ausência — isso muda todo mês.
- `outlier_iqr_mult`: quão "extremo" um chamado precisa ser para aparecer
  destacado como outlier. Quanto maior, mais exigente (só pega casos muito
  fora da curva). O padrão (8.0) foi calibrado para isolar só casos realmente
  extremos, não qualquer chamado acima da média.
- `pasta_historico`: diretório dos arquivos históricos mensais.

## Limitações conhecidas (e por quê)

- **Reparo automático de linhas com colunas deslocadas**: alguns exports do
  GLPI trazem títulos de chamado com caracteres especiais (tabulação, etc.)
  que empurram as colunas seguintes. O script detecta isso (data de abertura
  que não vira uma data válida) e tenta reconstruir varrendo a linha em busca
  de timestamps. Funciona bem para o padrão visto até agora, mas **sempre
  confira o aviso impresso no terminal** (`[aviso] N linha(s) reparadas...`)
  e valide manualmente os IDs listados antes de bater o olho só no relatório.
- **Limpeza de título**: o script tenta remover o padrão
  "- Nome Sobrenome - 1234 -" do fim dos títulos (para não expor nome de
  quem abriu o chamado numa tabela gerencial) e corta títulos longos. Isso é
  "melhor esforço" — nem todo formato de título é coberto, então revise a
  tabela da Seção 3 antes de enviar para a liderança.
- **Tempo de chamados em aberto**: para chamados não concluídos, "horas
  acumuladas" é tempo corrido (agora − abertura), não tempo de trabalho
  efetivo — a base do GLPI não registra esforço real, só tempo corrido (SLA).
- A Seção 7 (Plano de Ação) sai como um roteiro vazio de propósito — é o tipo
  de conteúdo que só faz sentido escrito por quem conhece o contexto do mês.

## Estrutura do código (caso queira estender)

- `carregar_dados` / `_reparar_linhas_deslocadas` / `preparar_dataframe`:
  ingestão e limpeza.
- `metricas_*`: cada função calcula um bloco do relatório (painel, categoria,
  SLA, capacidade, outliers, recorrência).
- `grafico_*`: geração dos PNGs com matplotlib.
- `montar_documento`: monta o .docx com python-docx (tabelas, banners,
  imagens).
- `gerar_relatorio`: orquestra tudo — é a função chamada pelo `if __name__`.

Cada bloco do relatório é gerado por uma função separada, então dá para
comentar/remover uma seção inteira do `montar_documento` sem afetar o resto,
ou adicionar uma seção nova seguindo o mesmo padrão.
