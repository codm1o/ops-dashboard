# Operational Analytics Dashboard

Dashboard analítico executado no navegador, desenvolvido para monitorar desempenho de equipes e conformidade de prazos em operações jurídicas. O projeto demonstra integração de pipeline de dados, renderização dinâmica com Chart.js e um sistema completo de temas claro/escuro, tudo em um único arquivo HTML autocontido, sem necessidade de ferramentas de build.

## Visão Geral

O dashboard lê um arquivo JSON local e renderiza um conjunto de painéis que cobrem:

- Cards de KPI mensais: total de atividades concluídas, tarefas, prazos e taxa de cumprimento
- Distribuição de prazos: breakdown D0 / D-1 / D-2 / D-3+ / D+ com gráfico de rosca
- Ranking por profissional e gráficos de barras empilhadas
- Distribuição de tipos de workflow com painel de detalhamento
- Breakdown por área jurídica (Cível, Trabalhista, Tributário)
- Painel de auditoria: atividades concluídas sem justificativa, agrupadas por cargo e área

O script Python (`gerar_dados.py`) lê uma planilha Excel estruturada, processa os registros de atividades e escreve as métricas calculadas em `dados.json`, que o dashboard carrega em tempo de execução.


## Estrutura de Arquivos

```
ops-dashboard/
├── index.html        Dashboard principal — HTML, CSS e JS em um único arquivo
├── dados.json        Arquivo de dados carregado pelo dashboard em tempo de execução
├── gerar_dados.py    Script Python que processa a planilha Excel e atualiza dados.json
└── README.md         Este arquivo
```


## Como Executar

### 1. Visualizar o dashboard

Abra `index.html` no navegador. Como o dashboard carrega `dados.json` via `fetch()`, é necessário servir os arquivos por um servidor HTTP local — não é possível abrir o HTML diretamente como URL `file://`.

Com Python (recomendado):

```bash
cd ops-dashboard
python -m http.server 8000
```

Em seguida, acesse `http://localhost:8000` no navegador.

Com Node.js (`npx serve`):

```bash
cd ops-dashboard
npx serve .
```

### 2. Gerar dados a partir de uma planilha real

Instale as dependências Python necessárias:

```bash
pip install pandas openpyxl
```

Execute o script:

```bash
python gerar_dados.py dados_abril.xlsx abril "Abril 2026"
```

O script adiciona ou substitui o mês informado em `dados.json`. Pode ser executado múltiplas vezes para meses diferentes e os dados existentes são preservados.

---

## Formato da Planilha Excel

A planilha de entrada deve conter as seguintes colunas (os nomes devem ser exatamente iguais):

| Coluna | Descrição |
|---|---|
| Data fatal | Data limite do prazo (formato DD/MM/AAAA) |
| Data conclusão | Data em que a atividade foi concluída (formato DD/MM/AAAA) |
| Prazo | Inteiro indicando dias de antecedência (0, -1, -2, etc.) ou positivo se após o prazo |
| Evento | "Principal" para atividades principais |
| Tipo | "Prazo" para atividades de prazo; qualquer outro valor é tratado como tarefa |
| Área | Cível, Trabalhista ou Tributário |
| Responsável efetivo | Nome completo do profissional responsável |
| Cargo | Advogado, Coordenador ou Estagiário |
| Workflow | Tipo de atividade jurídica (ex: Manifestação, Contestação) |
| Auditoria | Preenchido com "Sem justificativa" quando nenhuma justificativa foi registrada |

---

## Classificação de Prazos

O campo `Prazo` é um inteiro que indica quantos dias antes da data fatal a atividade foi concluída. Valores positivos indicam que a conclusão ocorreu após o prazo.

| Valor | Classificação | Significado |
|---|---|---|
| 0 | D0 | Concluído no dia do prazo |
| -1 | D-1 | Concluído 1 dia antes |
| -2 | D-2 | Concluído 2 dias antes |
| -3 ou menos | D-3+ | Concluído 3 ou mais dias antes |
| 1 ou mais | D+ | Concluído após a data fatal |

---

## Formato do Arquivo de Dados (dados.json)

O arquivo JSON é um dicionário com chave por mês. Cada entrada contém as métricas calculadas e um array com os dados por profissional. Exemplo de estrutura:

```json
{
  "abril": {
    "label": "Abril 2026",
    "available": true,
    "totalConcluidos": 1903,
    "tarefas": 1417,
    "prazos": 414,
    "etapas": 72,
    "d0": 83,
    "dm1": 48,
    "dm2": 143,
    "dm3": 140,
    "dp": 0,
    "areaTrabalhista": 174,
    "areaCivel": 197,
    "areaTributario": 43,
    "workflowDist": { "Análise de Publicação": 1313, "...": 0 },
    "sjAdv": { "total": 1, "civel": 1, "trabalhista": 0, "tributario": 0 },
    "sjEst": { "total": 4, "civel": 3, "trabalhista": 1, "tributario": 0 },
    "profData": [
      {
        "name": "Amanda Ferreira",
        "cargo": "Advogado",
        "area": "Cível",
        "tarefas": 213,
        "prazos": 42,
        "etapas": 0,
        "total": 255,
        "d0": 10, "dm1": 6, "dm2": 15, "dm3": 11, "dp": 0,
        "sem_justificativa": 0,
        "workflow_list": ["Análise de Publicação", "..."]
      }
    ]
  }
}
```

Para adicionar um mês futuro sem dados ainda, inclua apenas:

```json
{
  "maio": {
    "label": "Maio 2026",
    "available": false
  }
}
```

---

## Adicionando um Novo Mês ao Dashboard

1. Execute `gerar_dados.py` com a nova planilha para atualizar `dados.json`.
2. Em `index.html`, adicione um novo `<button>` dentro da seção `.month-nav`:

```html
<button class="month-tab" data-month="maio">Maio</button>
```

Remova a classe `upcoming` ou `unavailable` assim que os dados estiverem disponíveis.

---

## Sistema de Temas

O dashboard inclui temas claro e escuro controlados por CSS custom properties. O tema selecionado é salvo em `localStorage` sob a chave `sl-theme` e persiste entre sessões. O botão de alternância no cabeçalho superior direito alterna entre os modos.

---

## Dependências

- Chart.js 4.4.1 (carregado via cdnjs, sem instalação necessária)
- Python 3.8 ou superior (apenas para o script de pipeline de dados)
- pandas e openpyxl (pacotes Python, instaláveis via pip)

Nenhuma etapa de build, bundler ou gerenciador de pacotes de frontend é necessário para rodar o dashboard.

---

## Licença

MIT License. Livre para usar, modificar e distribuir.
