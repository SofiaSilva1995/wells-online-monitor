# Pierre Fabre · Wells Online Monitor

Bot automático de monitorização de stock para as marcas Pierre Fabre em [Wells.pt](https://www.wells.pt).  
Corre todos os dias às **15:00 (Hora de Lisboa — Verão)** via GitHub Actions.

---

## 📦 Marcas Monitorizadas

| Marca | Loja |
|-------|------|
| Avène | wells.pt |
| Ducray | wells.pt |
| Klorane | wells.pt |
| René Furterer | wells.pt |
| A-Derma | wells.pt |

---

## 🚀 Como Funciona

1. O bot percorre todos os produtos de cada marca no Wells.pt
2. Deteta quais os produtos **Out of Stock**
3. Identifica os produtos **novos** Out of Stock desde a última execução
4. Gera um **dashboard HTML** interativo com gráficos e filtros
5. Exporta um ficheiro **Excel** consolidado com todos os dados
6. Envia tudo por **email** automaticamente

---

## 📊 Funcionalidades do Dashboard

- KPIs — Total, Out of Stock, In Stock, Novos OOS, Sem Desconto
- Cards por marca com setas de tendência (↑↓ vs dia anterior)
- Gráfico de evolução de OOS por marca e por dia
- Tabela de produtos com filtros e ordenação
- Contagem de dias em Out of Stock por produto
- Detecção de desconto por variante
- Exportação para CSV

---

## ⚙️ Workflows

| Workflow | Quando corre | Descrição |
|----------|-------------|-----------|
| `Wells Online Monitor - Daily Run` | Todos os dias às 15:00 | Execução completa — scraping, dashboard, email |
| `Wells - Enviar Dashboard por Email` | Manual | Execução a pedido com envio do dashboard por email |

---

## 🗂️ Estrutura do Repositório

```
wells-online-monitor/
├── .github/
│   └── workflows/
│       ├── wells_daily.yml           # Execução diária automática
│       └── wells_send_dashboard.yml  # Execução manual a pedido
├── wells_oos_monitor_ALL_BRANDS_AUTO_V4.py  # Bot principal
├── .gitignore
└── README.md
```

> **Nota:** O ficheiro `config_email.py` **nunca é feito commit** — as credenciais ficam nos GitHub Secrets.

---

## 🔐 GitHub Secrets Necessários

Vai a **Settings → Secrets and variables → Actions** e adiciona:

| Secret | Descrição |
|--------|-----------|
| `SMTP_USER` | Endereço Gmail usado para enviar os emails |
| `SMTP_PASS` | App Password do Gmail (16 caracteres) |
| `EMAIL_TO` | Endereço de email destinatário |

---

## 📁 Ficheiros Persistentes (commit automático pelo bot)

Após cada execução, o bot faz commit automático destes ficheiros para o repositório, garantindo que o histórico é preservado entre execuções:

- `logs/historico.csv` — histórico diário completo de todos os produtos
- `known_oos_*.json` — último estado OOS conhecido por marca (usado para detetar *novos* OOS)

---

## 🛠️ Dependências

```
requests
beautifulsoup4
openpyxl
playwright
```

---

## 📬 Artefactos Gerados

Após cada execução, os seguintes ficheiros ficam disponíveis em **Actions → [execução] → Artifacts**:

- `dashboard_[RUN_ID].html` — dashboard interativo
- `todas_marcas_[RUN_ID].xlsx` — Excel consolidado com todas as marcas + separador de variantes

Os artefactos são mantidos durante **30 dias**.

---

*Desenvolvido para Pierre Fabre Portugal — Uso interno.*
