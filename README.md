# Curva de Juros DI — Estrutura a Termo (ETTJ Prefixada)

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![xlwings](https://img.shields.io/badge/xlwings-0.30%2B-green)
![ANBIMA](https://img.shields.io/badge/Dados-ANBIMA-orange)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

---

## PT-BR

-- Caso não queira rodar o código completo main.py, você pode ver uma prévia do resultado fazendo o download do arquivo Amostra_do_projeto.xlsx

				Espero que aprecie meu projeto, dediquei-me por um bom tempo para faze-lo :)

### Introdução

A curva de juros prefixada é a espinha dorsal do mercado de renda fixa brasileiro. Ela mostra o preço do dinheiro no tempo — quanto o mercado exige de retorno para emprestar por 1 mês, 6 meses, 2 anos, 5 anos. Tudo o que envolve taxa de juros no Brasil passa por ela.

**Para uma trading desk**, a curva é ferramenta de trabalho diária: precificação de títulos, marcação a mercado do book, análise de movimentos (a curva abriu? fechou? inclinou?). Saber se o mercado está fazendo steepening ou flattening determina decisões de duration e hedge.

**Para o investidor pessoa física**, entender a curva responde perguntas práticas: vale mais comprar um CDB de 1 ano ou 3 anos agora? A curva está invertida — o mercado está precificando queda de juros no futuro. Essa leitura muda a estratégia de alocação.

Este projeto busca os dados diretamente da ANBIMA, interpola a curva completa e entrega um dashboard Excel com gráfico, consulta de taxas, precificação de LTN e análise de movimentos — tudo com um único comando.

```bash
pip install -r requirements.txt
python main.py
```

> **Requisitos:** Windows · Python 3.10+ · Microsoft Excel instalado

---

### Metodologia

#### Fonte de dados

Os dados vêm da ANBIMA via biblioteca `pyettj`. A ANBIMA publica diariamente 15 vértices da ETTJ Prefixada (de 21 DU a 1.260 DU). O projeto tenta buscar os dados do dia atual; se ainda não publicados, busca automaticamente o dia útil anterior via `bizdays` (calendário ANBIMA) e avisa no terminal.

#### Interpolação — Flat Forward

Entre os vértices publicados, o projeto interpola usando **Flat Forward**, padrão do mercado brasileiro (Manual de Curvas B3, seção 1.4.2). A interpolação ocorre sobre os **fatores de capitalização**, não sobre as taxas diretamente — isso elimina oportunidades de arbitragem teórica entre prazos.

```
r_i = (1 + r_a) × [(1 + r_p) / (1 + r_a)] ^ [(du_i − du_a) / (du_p − du_a)] − 1
```

#### Curva Forward Mês a Mês

A taxa forward implícita entre dois meses consecutivos revela o que o mercado precifica para cada período futuro isoladamente — diferente da taxa spot, que é acumulada desde hoje.

```
fator_fwd = (1 + r_B)^(du_B/252) / (1 + r_A)^(du_A/252)
taxa_fwd  = fator_fwd^(252 / (du_B − du_A)) − 1
```

#### Convenções brasileiras
- Base: **252 dias úteis** (não 365 dias corridos)
- Fator de desconto: `1 / (1 + r)^(du/252)`

#### Pipeline end-to-end

```
pyettj (ANBIMA) → 15 vértices reais
      ↓
buscador_dados.py → valida, salva histórico local (xlsx)
      ↓
interpolacao.py → Flat Forward → 58 pontos mensais (DU 21 a 1.218)
                              → Forward mês a mês
      ↓
analise.py → formato da curva (Normal/Invertida/Flat)
           → variação em bps vs 30 DU atrás
           → Shift / Steepening / Flattening
      ↓
precificacao.py → PU · Duration · DV01 · Stress +50bps (LTN)
      ↓
construtor_excel.py (xlwings) → Dashboard · Vértices · Movimentos
```

---

### Conclusão

O projeto replica, em Python, o fluxo de análise de curva que uma mesa de renda fixa executa diariamente. Os dados são sempre reais e diretos da ANBIMA. A lógica de fallback garante que o usuário nunca receba taxas fictícias — se o dia atual não foi publicado, busca o dia anterior na própria fonte.

Faz par com o [Treasury P&L Dashboard](https://github.com/murilomendesz/Treasury-P-L): um modela o book de títulos, este modela a curva que precifica esse book.

---

## EN

### Introduction

The Brazilian pre-fixed yield curve sets the price of money over time — what the market demands to lend for 1 month, 6 months, 2 years, 5 years. Every fixed-income instrument in Brazil is priced against it.

**On a trading desk**, it drives daily mark-to-market, duration positioning, and hedge decisions. Identifying whether the curve is shifting, steepening, or flattening directly informs book management.

**For retail investors**, reading the curve answers practical questions: is it better to lock in a 1-year or 3-year rate today? An inverted curve signals the market is pricing in future rate cuts — that changes allocation strategy.

This project fetches live ANBIMA data, interpolates the full curve, and delivers an Excel dashboard with chart, rate lookup, LTN pricing, and movement analysis — all from a single command.

```bash
pip install -r requirements.txt
python main.py
```

> **Requirements:** Windows · Python 3.10+ · Microsoft Excel installed

---

### Methodology

**Data** — fetched from ANBIMA via `pyettj` (15 vertices, 21 to 1,260 business days). If today's data hasn't been published yet, the previous business day is fetched automatically from the same source.

**Flat Forward interpolation** — market standard in Brazil (B3 Curves Manual, §1.4.2). Interpolates over compounding factors, not rates, to prevent theoretical arbitrage between tenors.

```
r_i = (1 + r_a) × [(1 + r_p) / (1 + r_a)] ^ [(du_i − du_a) / (du_p − du_a)] − 1
```

**Month-by-month forward** — isolates the rate the market implies for each future period, separate from the accumulated spot rate.

**Day count convention** — 252 business days (Brazilian standard), ANBIMA calendar via `bizdays`.

**Pipeline** — `pyettj` → validation & local cache → Flat Forward interpolation (58 monthly points) → curve shape classification → bps movement analysis → xlwings Excel population.

---

### Conclusion

The project replicates in Python the yield curve workflow of a fixed-income desk. Data is always real and sourced directly from ANBIMA. Pairs with the [Treasury P&L Dashboard](https://github.com/murilomendesz/Treasury-P-L): one models the bond book, this one models the curve that prices it.

---

**Murilo Mendes** — Economics student (UNIP), CEA-ANBIMA & AAI certified · [murilomendesz](https://github.com/murilomendesz)
