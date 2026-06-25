# 📊 Dashboard Financeiro Maré Educação — Raízes

Dashboard interativo para acompanhamento financeiro dos produtos **Raízes** e **Raízes - R** da Maré Educação. Lê exports da Hotmart, calcula projeção de faturamento, inadimplência por canal/lançamento/vendedor e gera um dashboard visual em HTML.

🌐 **Dashboard ao vivo:** [Acesse aqui](#) <!-- Será preenchido após ativar GitHub Pages -->

---

## 🚀 O que esse pacote faz

A partir de relatórios exportados da Hotmart (`sales_history_*.xls`), o script:

- Combina histórico completo (deduplica transações repetidas)
- Identifica **parcelas pagas, atrasadas, futuras e desconhecidas** por venda
- Detecta cancelamento de assinatura (sem atividade >45 dias)
- Calcula **projeção de 12 meses** (cenário otimista e ajustado pela inadimplência histórica)
- Gera um **dashboard HTML** com:
  - KPIs financeiros
  - Projeção mensal (com filtro de período)
  - Histórico de recebimento
  - Inadimplência por tipo de cobrança, forma de pagamento, vendedor e lançamento
  - Distribuição de devedores por nº de parcelas
  - Análise por vendedor (Código SRC)

---

## 📦 Conteúdo do repositório

| Arquivo | Descrição |
|---|---|
| `gerar_base_mare.py` | Script Python que processa os dados e gera xlsx + dashboard |
| `index.html` | Dashboard renderizado (versão pública sem dados sensíveis) |
| `dashboard_mare_compartilhar.html` | Cópia do dashboard sem dados sensíveis |
| `ATUALIZAR_DASHBOARD.bat` | Runner Windows — dois cliques pra atualizar tudo |
| `ATUALIZAR_DASHBOARD_AGENDADO.bat` | Versão silenciosa pro Agendador de Tarefas |

---

## 🛠️ Como usar localmente

### Pré-requisitos
- Python 3.10+
- Bibliotecas: `pandas`, `openpyxl`, `python-dateutil`

```bash
pip install pandas openpyxl python-dateutil
```

### Rodar
1. Exporte o relatório de vendas da Hotmart (com Atrasados incluídos)
2. Salve em `C:\Users\<seu_usuario>\Downloads\` — qualquer arquivo `sales_history_*.xls` ou `venda*.xlsx` é detectado automaticamente
3. Execute:

```bash
python gerar_base_mare.py
```

ou clique duas vezes no `ATUALIZAR_DASHBOARD.bat`.

O script gera:
- `base_mare_parcelas_projecao.xlsx` (planilha completa com 5 abas)
- `dashboard_mare.html` (dashboard interativo)

---

## 🔒 Privacidade

Este repositório contém **apenas o código** e a **versão pública do dashboard** (sem nomes, e-mails ou telefones de clientes). Dados sensíveis são tratados localmente e **nunca são commitados** (veja `.gitignore`).

Para abrir mão da privacidade (uso interno completo com nomes), use a versão `dashboard_mare.html` gerada localmente — ela inclui a tabela de "Top 12 maiores devedores".

---

## 📝 Licença

Uso interno Maré Educação. Não distribuir externamente.
