# 📊 Controle Financeiro Pessoal - 12 Meses (Excel Dashboard)

Planilha automatizada em Microsoft Excel com Dashboard executivo, KPIs dinâmicos, validação de dados, gráficos nativos e matriz anual para acompanhamento financeiro de 12 meses.

---

## 🚀 Funcionalidades

- **📊 Dashboard Executivo**:
  - **7 Cards de KPI**: Salário Líquido, Rendas Extras, Total de Receitas, Total de Despesas, Disponível no Mês, Reserva Financeira (com % da meta de 20%) e Gastos com Estudos.
  - **Filtro de Mês Dinâmico (Célula N1)**: Alterne entre `Todos` ou qualquer mês específico (`Jan`, `Fev`, `Mar`...) e veja todos os cálculos e gráficos atualizarem instantaneamente.
  - **Gastos por Categoria**: Tabela analítica com cálculo dinâmico de valor e percentual (`% do Total`).
  - **Gráfico de Pizza Integrado**: Distribuição visual proporcional das despesas.
  - **Visão Anual (12 Meses)**: Matriz consolidada de Janeiro a Dezembro somando Receitas, Despesas, Saldo do Mês e Reserva.
- **📝 Base de Lançamentos**:
  - Filtros automáticos no cabeçalho.
  - Validação de Dados (listas suspensas) para `Tipo`, `Categoria` e `Mês`.
  - Formatação monetária (`R$ #,##0.00`) e datas (`DD/MM/AAAA`).
  - Cores temáticas por categoria.
- **⚙️ Configurações**:
  - Aba de apoio com listas mestras e configuração de metas.

---

## 📁 Estrutura do Projeto

```
.
├── Controle_Financeiro_2025.xlsx  # Planilha final pronta para uso
├── build_financial_sheet.py       # Script Python em openpyxl para gerar/customizar
├── requirements.txt               # Dependências do projeto
├── .gitignore                     # Arquivos ignorados pelo Git
└── README.md                      # Documentação do projeto
```

---

## 🛠️ Como Executar ou Regenerar a Planilha

Caso deseje recriar a planilha via Python:

1. Crie e ative um ambiente virtual:
   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # Linux/Mac:
   source .venv/bin/activate
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Execute o script gerador:
   ```bash
   python build_financial_sheet.py
   ```
