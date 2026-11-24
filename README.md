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

## Persistência das sessões web

Cada upload recebe um identificador UUID e cria um arquivo JSON em
`instance/sessoes/<id_sessao>.json`. Esse arquivo é a fonte da verdade da
revisão e contém os registros editáveis, mês de referência, nome de origem,
validações, timestamps e resultado da geração. Datas são armazenadas em ISO
8601; colunas derivadas são reconstruídas por `preparar_dataframe` ao carregar.

Toda edição, alteração em lote, seleção de mês ou resultado é gravada com
substituição atômica. Antes de uma atualização, a versão válida anterior é
preservada em `instance/sessoes/<id_sessao>.json.bak`. Arquivos temporários e
locks internos não são fontes de estado. O upload em `instance/uploads/` é
removido após a leitura do arquivo.

Por padrão, sessões sem atualização por mais de 7 dias e seus backups são
removidos durante a inicialização do app. Altere esse prazo com a variável de
ambiente `GLPI_SESSION_TTL_DAYS` (dias inteiros). O DOCX continua em
`relatorios_glpi/execucoes/`; a sessão JSON apenas mantém o caminho e os dados
necessários para consultar e baixar o resultado.

Em reload do módulo no mesmo processo, uma sessão antiga ainda presente no
dict legado `SESSOES` é convertida para JSON no primeiro acesso e removida do
dict. Um reinício completo de um processo da versão antiga não pode recuperar
sessões que nunca foram gravadas em disco; antes de encerrar essa versão,
conclua ou exporte qualquer revisão ainda ativa.

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

## Configuração (tela "Configuração" do app)

Os parâmetros ficam em `instance/config.json`, local e fora do git, e são
editados pela tela `/configuracao`. A linha de comando usa o mesmo arquivo.
Sem esse arquivo, valem os padrões do `CONFIG` no topo do script.

- **Cabeçalho e escopo**: título, subtítulo, autor e a entidade da frente
  Sustentação. As demais entidades formam a frente Projetos.
- **SLA institucional**: classe (P1–P4) de cada prioridade do GLPI e o prazo
  de solução de cada classe, em horas corridas ou em dias úteis.
  - Horas corridas: contadas desde a abertura, em regime 24×7.
  - Dias úteis: vencem no fim do N-ésimo dia útil seguinte à abertura.
  - O mapeamento padrão (Crítica e Muito Alta = P1, Alta = P2, Média = P3,
    Baixa e Muito Baixa = P4) é provisório e deve ser validado com a gestão.
  - O SLA de resposta não é medido, porque o export não traz a data do
    primeiro atendimento.
- **Calendário**: os feriados nacionais são calculados automaticamente,
  inclusive a Sexta-feira Santa. Datas extras (estaduais, municipais, pontos
  facultativos) são informadas uma por linha, no formato AAAA-MM-DD.
- **Capacidade**: analistas, horas por dia e dias de férias ou ausência.
  Revise as férias a cada mês. Os dias úteis do mês são calculados pelo
  calendário, a menos que você informe outro valor.
- **Outlier**: o multiplicador do IQR tem padrão 8, valor validado para este
  dataset.

## Dados e privacidade (LGPD)

Os exports do GLPI trazem títulos de chamados que podem conter nomes de
pessoas. Por isso o `.gitignore` mantém fora do repositório:

- os exports (`dados_*.csv`, `*.xlsx`);
- as sessões de revisão e a configuração local (`instance/`);
- os relatórios gerados (`relatorios_glpi/execucoes/`).

Os testes que usam o export real (`dados_202609301724.csv`) só rodam quando
esse arquivo existe na máquina. Os demais testes usam dados sintéticos.

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
