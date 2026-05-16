# 📊 KPI Master — Gestão de Estoque

> Gerador automatizado de relatórios KPI para controle de estoque, EPI/EPC e gestão do RM.
> Desenvolvido para a **UHE Estrela**.

---

## 🚀 Sobre o Projeto

O **KPI Master** é uma aplicação web que transforma planilhas brutas de estoque em relatórios profissionais e analíticos. Ele foi criado para automatizar a geração de grades de controle de estoque, acompanhamento de EPI (Equipamentos de Proteção Individual) e análise completa do estoque do sistema RM.

Basta enviar a planilha, configurar os parâmetros e baixar o relatório formatado — tudo em segundos.

---

## ✨ Funcionalidades

### 📦 KPI Saída (Controle de Estoque e EPI)

Gera relatórios completos a partir das abas **ESTOQUE**, **ENTRADA** e **SAÍDA**:

| Aba do Relatório | Descrição |
|---|---|
| **Resumo Geral** | Visão consolidada de todos os meses com total de saídas, itens distintos e grupo líder |
| **KPI Mensal** | Top 20 materiais mais saídos + saída por grupo com cards de destaque |
| **Classificação ABC** | Curva ABC completa com classes A, B e C coloridas |
| **Valor por Categoria** | Valor total do estoque agrupado por categoria/grupo |
| **Estoque Morto** | Itens sem movimentação nos últimos N meses com valor parado |
| **⚠ Alerta Estoque** | Classificação por nível crítico: **Ruptura**, **Crítico** e **Normal** com consumo médio mensal |

### 📊 KPI do RM (Estoque do Sistema RM)

Gera 5 abas analíticas a partir da planilha de estoque do RM:

| Aba do Relatório | Descrição |
|---|---|
| **Resumo Executivo** | Dashboard com cards de KPIs, curva ABC e Top 10 Grupos |
| **Estoque Completo** | Todo o estoque com classificação ABC, filtros e freeze panes |
| **Análise por Grupo** | Estatísticas por grupo: saldo, custo médio/máx/mín, valor total |
| **Top 50 Itens** | Itens com maior valor financeiro (ouro/prata/bronze) |
| **Estatísticas** | Indicadores analíticos: média, mediana, desvio padrão, concentração |

---

## 🛠️ Tecnologias

- **Python 3.11+**
- **Flask** — framework web
- **Pandas** — processamento de dados
- **OpenPyXL** — geração de planilhas Excel formatadas
- **Gunicorn** — servidor WSGI para produção

---

## 📋 Requisitos

### KPI Saída
A planilha de entrada deve conter as abas:
- **ESTOQUE** — com colunas: COD, GRUPO, DESCRIÇÃO, SALDO, VALOR UNIT, VALOR TOTAL
- **SAÍDA** — com colunas: DATA, COD, GRUPO, DESCRIÇÃO, QUANT, UN
- **ENTRADA** — (opcional, para análise complementar)

### KPI do RM
A planilha deve conter as colunas:
- `LOCESTOQUE` — Local de estoque
- `GRUPO` — Grupo do material
- `CODIGOPRD` — Código do produto
- `PRODUTO` — Descrição do produto
- `SALDO` — Quantidade em estoque
- `CUSTOMEDIO` — Custo médio
- `VALORFINANCEIRO` — Valor financeiro total

---

## 🚦 Como Usar

### Instalação Local

```bash
# Clonar o repositório
git clone https://github.com/eullon1234-creator/gerador-de-kpi.git
cd gerador-de-kpi

# Instalar dependências
pip install -r requirements.txt

# Rodar a aplicação
python app.py
```

Acesse: **http://127.0.0.1:5000**

### Deploy na Vercel

O projeto já está configurado para deploy na Vercel com o arquivo `vercel.json`.

---

## ⚙️ Parâmetros Configuráveis

| Parâmetro | Padrão | Descrição |
|---|---|---|
| **Período** | Livre | Filtro de data início/fim para análise |
| **Estoque Morto** | 3 meses | Itens sem saída nos últimos N meses |
| **Curva A** | 80% | Limite superior da classe A |
| **Curva B** | 95% | Limite superior da classe B |
| **Top Grupos (RM)** | 10 | Quantidade de grupos no ranking |
| **Top Itens (RM)** | 50 | Quantidade de itens no ranking |

---

## 📁 Estrutura do Projeto

```
gerador-de-kpi/
├── app.py                 # Aplicação Flask (rotas e lógica web)
├── kpi_generator.py       # Gerador de KPI Saída (estoque/EPI)
├── kpi_rm_generator.py    # Gerador de KPI do RM
├── requirements.txt       # Dependências Python
├── runtime.txt            # Versão do Python para deploy
├── vercel.json            # Configuração de deploy Vercel
└── templates/
    └── index.html         # Interface web com UI animada
```

---

## 🎨 Interface

A aplicação possui uma interface moderna com:
- **Background animado** com orbs flutuantes e partículas
- **Glassmorphism** com blur e transparência
- **Tabs** para alternar entre KPI Saída e KPI RM
- **Drag & Drop** para upload de planilhas
- **Animações** em todos os elementos interativos
- **Design responsivo** para mobile e desktop

---

## 📄 Licença

Projeto desenvolvido para uso interno da **UHE Estrela**.

---

## 👨‍💻 Autor

Desenvolvido por **Eullon** — [GitHub](https://github.com/eullon1234-creator)
