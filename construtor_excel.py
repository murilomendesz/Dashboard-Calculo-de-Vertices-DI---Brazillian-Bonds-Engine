"""
construtor_excel.py — Popula o template Excel com dados reais via xlwings

Estratégia:
  - Dashboard: escreve dados bimestrais nas células de gráfico (rows 10-38).
    CONSULTA e PRECIFICAÇÃO ficam nas posições originais do template (rows 32-51).
  - Aux: 58 linhas mensais para o dropdown (DU 21 a 1218).
  - Vértices: 58 pontos mensais com formatação idêntica ao template.
  - Movimentos: vértices ANBIMA com formatação condicional (verde/vermelho).
"""
import shutil
import os
import time
import xlwings as xw
import pandas as pd
from datetime import date

# Cores do template
COR_HEADER_BG  = "FF243356"   # azul escuro — cabeçalhos
COR_DADOS_BG   = "FF162240"   # azul escuro mais escuro — linhas de dados
COR_VERDE_BG   = "FF1F4E2C"   # verde escuro — variação positiva
COR_VERMELHO_BG = "FF4E1F1F"  # vermelho escuro — variação negativa


# ──────────────────────────────────────────────────────────────────────────
# Helpers de label
# ──────────────────────────────────────────────────────────────────────────

def _mes_pt(mes: int) -> str:
    return ["Jan","Fev","Mar","Abr","Mai","Jun",
            "Jul","Ago","Set","Out","Nov","Dez"][mes - 1]


def _label_vencimento(du: int, data_ref: date, step: int = 21) -> str:
    meses = round(du / step)
    ano, mes = data_ref.year, data_ref.month + meses
    while mes > 12:
        mes -= 12; ano += 1
    return f"{_mes_pt(mes)}/{ano} ({du} DU)"


def _label_ltn(du: int, data_ref: date, step: int = 21) -> str:
    meses = round(du / step)
    ano, mes = data_ref.year, data_ref.month + meses
    while mes > 12:
        mes -= 12; ano += 1
    return f"LTN {_mes_pt(mes)}/{ano}"


# ──────────────────────────────────────────────────────────────────────────
# Aba Aux  (tabela de lookup — 58 linhas mensais)
# ──────────────────────────────────────────────────────────────────────────

def _popular_aux(ws_aux, curva_mensal: pd.DataFrame, data_ref: date, data_comp: date):
    """
    Substitui os dados mock da aba Aux por dados reais mensais.
    Colunas A:E = lookup dropdown (vencimento, DU, taxa_hoje, taxa_comp, forward)
    Colunas F:G = lista de LTNs para dropdown de precificação
    """
    # Cabeçalhos com datas reais para evitar ambiguidade
    ws_aux.range("A1").value = "Vencimento"
    ws_aux.range("B1").value = "Prazo (DU)"
    ws_aux.range("C1").value = f"Spot {data_ref.strftime('%d/%m/%Y')}"
    ws_aux.range("D1").value = f"Spot {data_comp.strftime('%d/%m/%Y')}"
    ws_aux.range("E1").value = "Forward Mes a Mes"
    ws_aux.range("F1").value = "LTN"
    ws_aux.range("G1").value = "Prazo (DU)"

    # Limpar apenas dados
    ws_aux.range("A2:G200").clear_contents()

    n = len(curva_mensal)

    # A:E — lookup mensal
    lookup = []
    for i, row in curva_mensal.iterrows():
        du = int(row["dias_uteis"])
        label = _label_vencimento(du, data_ref)
        taxa_hoje = float(row["taxa_spot"])
        taxa_30d  = float(row["taxa_30d"]) if pd.notna(row.get("taxa_30d")) else None
        forward   = float(row["forward"])
        lookup.append([label, du, taxa_hoje, taxa_30d or "", forward])

    ws_aux.range("A2").value = lookup

    # F:G — LTNs mensais (a partir de DU=42, ~2 meses à frente)
    ltns = []
    for i, row in curva_mensal.iterrows():
        du = int(row["dias_uteis"])
        if du < 42:
            continue
        ltns.append([_label_ltn(du, data_ref), du])

    ws_aux.range("F2").value = ltns

    return n, len(ltns)


# ──────────────────────────────────────────────────────────────────────────
# Dashboard — cards de resumo
# ──────────────────────────────────────────────────────────────────────────

def _popular_dashboard_cards(ws_dash, curva_mensal: pd.DataFrame,
                              data_ref: date, data_comp: date,
                              spread_bps: float, formato: str):
    taxa_curta = float(curva_mensal.iloc[0]["taxa_spot"])
    taxa_longa = float(curva_mensal.iloc[-1]["taxa_spot"])
    sinal = "+" if spread_bps >= 0 else ""

    ws_dash.range("B6").value = data_ref.strftime("%d/%m/%Y")
    ws_dash.range("E6").value = f"{taxa_curta*100:.2f}% a.a."
    ws_dash.range("H6").value = f"{taxa_longa*100:.2f}% a.a."
    ws_dash.range("K6").value = f"{sinal}{spread_bps:.0f} bps | {formato}"

    labels = {"Normal": "Taxa longa > Taxa curta",
               "Invertida": "Taxa curta > Taxa longa",
               "Flat": "Curva sem inclinação relevante"}
    ws_dash.range("K7").value = labels.get(formato, "")

    # Nomes das séries do gráfico (C9, D9, E9) — escritos dinamicamente
    # para que o label deixe claro o que cada série representa
    ws_dash.range("C9").value = f"Spot {data_ref.strftime('%d/%m/%Y')} (hoje)"
    ws_dash.range("D9").value = f"Spot {data_comp.strftime('%d/%m/%Y')} (30 DU atras)"
    ws_dash.range("E9").value = "Forward Mes a Mes"


# ──────────────────────────────────────────────────────────────────────────
# Dashboard — dados do gráfico (bimestral, colunas B/C/D/E rows 10-38)
# ──────────────────────────────────────────────────────────────────────────

def _popular_dashboard_chart(ws_dash, curva_mensal: pd.DataFrame):
    """
    Atualiza as células de dados do gráfico (B10:E38) com valores bimestrais.

    Layout original do template (preservado):
      rows 10-31 : 22 pontos bimestrais (DU=42 a 924)
      row  32    : cabeçalho CONSULTA (mesclado B32:F32) — não tocar
      rows 33-38 : pontos DU 966-1218 interleaved com form fields
        row 33: dados, row 34: form, row 35: dados,
        row 36: form, row 37: form, row 38: dados

    A coluna E permanece com fórmulas =Aux!E{n}*100 do template (já resolvem
    via Aux atualizado).  Só atualizamos B, C, D.
    """
    # Selecionar pontos bimestrais (DU múltiplos de 42)
    bim = curva_mensal[curva_mensal["dias_uteis"] % 42 == 0].reset_index(drop=True)

    # Construir mapa du → (taxa_spot, taxa_30d)
    def get_rates(du):
        rows = bim[bim["dias_uteis"] == du]
        if rows.empty:
            return None, None
        r = rows.iloc[0]
        spot = float(r["taxa_spot"]) * 100
        t30  = float(r["taxa_30d"])  * 100 if pd.notna(r.get("taxa_30d")) else ""
        return spot, t30

    # Rows 10-31: DU 42, 84, ..., 924  (22 pontos) — escreve B:D em bulk
    bcd_block = []
    for idx in range(min(22, len(bim))):
        du = int(bim.loc[idx, "dias_uteis"])
        spot, t30 = get_rates(du)
        bcd_block.append([du, spot if spot is not None else "", t30 if spot is not None else ""])
    if bcd_block:
        ws_dash.range("B10").value = bcd_block

    # Rows 33, 35, 38: pontos reais DU 1008, 1092, 1218 — B+C+D
    # (rows 34, 36, 37 são linhas de formulário — não tocar: chart será corrigido
    #  para pular essas linhas via _corrigir_series_chart)
    data_rows_extra = {33: 23, 35: 25, 38: 28}   # row → bim_idx correto
    for row, bim_idx in data_rows_extra.items():
        if bim_idx >= len(bim):
            continue
        du = int(bim.loc[bim_idx, "dias_uteis"])
        spot, t30 = get_rates(du)
        if spot is None:
            continue
        ws_dash.range(f"B{row}").value = du
        ws_dash.range(f"C{row}").value = spot
        ws_dash.range(f"D{row}").value = t30

    # Atualizar fórmulas de forward (coluna E) para apontar ao Aux correto
    def aux_row_for_du(du):
        matches = curva_mensal[curva_mensal["dias_uteis"] == du]
        if matches.empty:
            return None
        return int(matches.index[0]) + 2   # +2 pois Aux começa na linha 2

    for idx in range(min(22, len(bim))):
        row = 10 + idx
        du = int(bim.loc[idx, "dias_uteis"])
        ar = aux_row_for_du(du)
        if ar:
            ws_dash.range(f"E{row}").formula = f"=Aux!E{ar}*100"

    for row, bim_idx in data_rows_extra.items():
        if bim_idx >= len(bim):
            continue
        du = int(bim.loc[bim_idx, "dias_uteis"])
        ar = aux_row_for_du(du)
        if ar:
            ws_dash.range(f"E{row}").formula = f"=Aux!E{ar}*100"


# ──────────────────────────────────────────────────────────────────────────
# Dashboard — atualizar dropdowns e fórmula de taxa do título
# ──────────────────────────────────────────────────────────────────────────

def _atualizar_chart_eixo_y(ws_dash, curva_mensal: pd.DataFrame):
    """Define escala do eixo Y com base no range real dos dados + margem."""
    try:
        spot_vals  = curva_mensal["taxa_spot"].dropna() * 100
        t30_vals   = curva_mensal["taxa_30d"].dropna()  * 100
        fwd_vals   = curva_mensal["forward"].dropna()   * 100

        todas = pd.concat([spot_vals, t30_vals, fwd_vals], ignore_index=True)
        data_min = todas.min()
        data_max = todas.max()

        # Margem de 1 pp, arredondado para inteiro mais próximo
        y_min = max(0.0, round(data_min - 1.0))
        y_max = round(data_max + 1.0)

        chart = ws_dash.charts[0]
        ax = chart.api[1].Axes(2)   # 2 = xlValue (eixo Y)
        ax.MinimumScaleIsAuto = False
        ax.MaximumScaleIsAuto = False
        ax.MinimumScale = y_min
        ax.MaximumScale = y_max
    except Exception as e:
        print(f"  [aviso] chart axis: {e}")


def _corrigir_series_chart(ws_dash):
    """
    Reescreve as SERIES do gráfico usando ranges não contíguos que excluem as
    linhas de formulário (34, 36, 37).  Linhas de dados reais: 10-31, 33, 35, 38.
    Isso elimina os rótulos "Selecione o vencimento" / "Taxa interpolada" / "Dias úteis"
    do eixo X e corrige a linha pontilhada da curva 30d.
    """
    try:
        chart = ws_dash.charts[0]
        sc = chart.api[1].SeriesCollection()

        x = "(Dashboard!$B$10:$B$31,Dashboard!$B$33,Dashboard!$B$35,Dashboard!$B$38)"
        y_by_order = {1: "C", 2: "D", 3: "E"}

        for i in range(1, sc.Count + 1):
            s = sc.Item(i)
            if i not in y_by_order:
                continue
            col = y_by_order[i]
            y = (f"(Dashboard!${col}$10:${col}$31,"
                 f"Dashboard!${col}$33,Dashboard!${col}$35,Dashboard!${col}$38)")
            name_ref = f"Dashboard!${col}$9"
            s.Formula = f"=SERIES({name_ref},{x},{y},{i})"
    except Exception as e:
        print(f"  [aviso] corrigir series chart: {e}")


def _add_dropdown(ws, cell_addr: str, formula1: str):
    try:
        cell = ws.range(cell_addr).api
        cell.Validation.Delete()
        cell.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula1)
    except Exception as e:
        print(f"  [aviso] dropdown {cell_addr}: {e}")


def _atualizar_dashboard_extras(ws_dash, n_aux: int, n_ltn: int):
    """
    - Atualiza dropdown D34 (consulta) → Aux!$A$2:$A${n_aux+1}
    - Atualiza dropdown D42 (LTN)      → Aux!$F$2:$F${n_ltn+1}
    - Define I43 como fórmula de lookup de taxa da curva
    """
    _add_dropdown(ws_dash, "D34", f"=Aux!$A$2:$A${n_aux + 1}")
    _add_dropdown(ws_dash, "D42", f"=Aux!$F$2:$F${n_ltn + 1}")

    # I43: taxa da curva = lookup do DU do título no Aux
    ws_dash.range("I43").formula = "=INDEX(Aux!C:C,MATCH(I42,Aux!B:B,0))"


# ──────────────────────────────────────────────────────────────────────────
# Aba Vértices — 58 linhas com formatação idêntica ao template
# ──────────────────────────────────────────────────────────────────────────

def _popular_vertice(ws_vert, curva_mensal: pd.DataFrame, data_ref: date, data_comp: date):
    # Subtítulo
    ws_vert.range("B3").value = (
        f"Data de referência: {data_ref.strftime('%d/%m/%Y')} | "
        f"Curva: PRE | Fonte: ANBIMA"
    )

    # Fundo azul escuro limitado à área útil da planilha (#0F1B33)
    ws_vert.range("A1:J65").api.Interior.Color = 0x331B0F
    ws_vert.range("A66:J67").api.Interior.Color = 0x331B0F
    # Restaurar cor do cabeçalho (linha 5) — azul mais escuro do template
    ws_vert.range("B5:G5").api.Interior.Color = 0x563324
    # Cabeçalhos com labels descritivos e datas reais
    ws_vert.range("B5").value = "Prazo (DU)"
    ws_vert.range("C5").value = f"Spot {data_ref.strftime('%d/%m/%Y')} (% a.a.)"
    ws_vert.range("D5").value = "Forward Mes a Mes (% a.a.)"
    ws_vert.range("E5").value = "Fator de Desconto"
    ws_vert.range("F5").value = "Tipo"
    ws_vert.range("G5").value = f"Spot {data_comp.strftime('%d/%m/%Y')} (% a.a.)"

    # Largura das colunas ajustada aos novos nomes
    ws_vert.range("B:B").api.ColumnWidth = 12
    ws_vert.range("C:C").api.ColumnWidth = 24
    ws_vert.range("D:D").api.ColumnWidth = 26
    ws_vert.range("E:E").api.ColumnWidth = 20
    ws_vert.range("F:F").api.ColumnWidth = 14
    ws_vert.range("G:G").api.ColumnWidth = 24

    # Legenda H3 — explica a coluna Spot 30 DU atrás para leigos
    legenda = (
        f"O que é a coluna 'Spot {data_comp.strftime('%d/%m/%Y')}'?\n"
        f"Ela mostra a taxa de juros para cada prazo conforme publicada pela ANBIMA "
        f"há 30 dias úteis (em {data_comp.strftime('%d/%m/%Y')}). "
        f"Não é uma taxa nova, é a mesma curva de hoje, só que fotografada no passado. "
        f"Compare com a coluna 'Spot hoje' para ver se os juros subiram ou caíram no período."
    )
    h3 = ws_vert.range("H3")
    h3.value = legenda
    h3.api.WrapText = True
    h3.api.Font.Color = 0xD9C5B8
    h3.api.Font.Size = 9
    ws_vert.range("H:H").api.ColumnWidth = 52

    # Desmerge + limpar em bulk
    n = len(curva_mensal)
    last_row = 5 + n
    try:
        ws_vert.range(f"B6:G{last_row}").api.UnMerge()
    except Exception:
        pass
    try:
        ws_vert.range(f"B6:G{last_row}").clear_contents()
    except Exception:
        for r in range(6, last_row + 1):
            try:
                ws_vert.range(f"B{r}:G{r}").clear_contents()
            except Exception:
                pass

    # Montar matriz de dados e escrever em UM único call
    data = []
    for _, row in curva_mensal.iterrows():
        du        = int(row["dias_uteis"])
        taxa_spot = float(row["taxa_spot"])
        forward   = float(row["forward"])
        fator     = 1 / ((1 + taxa_spot) ** (du / 252))
        tipo      = str(row.get("tipo", "Interpolado"))
        taxa_30d  = float(row["taxa_30d"]) if pd.notna(row.get("taxa_30d")) else ""
        data.append([du, taxa_spot, forward, fator, tipo, taxa_30d])

    ws_vert.range("B6").value = data   # escreve 58×6 de uma vez

    # Formatos numéricos por coluna inteira
    ws_vert.range(f"C6:C{last_row}").number_format = "0,0000%"
    ws_vert.range(f"D6:D{last_row}").number_format = "0,0000%"
    ws_vert.range(f"E6:E{last_row}").number_format = "0,000000"
    ws_vert.range(f"G6:G{last_row}").number_format = "0,0000%"

    # Background + alinhamento + bordas no range inteiro
    data_range = ws_vert.range(f"B6:G{last_row}").api
    data_range.Interior.Color = 0x402216   # #162240 BGR
    data_range.HorizontalAlignment = -4108  # xlCenter
    data_range.Borders.LineStyle = 1        # xlContinuous
    data_range.Borders.Weight = 2           # xlThin
    data_range.Borders.Color = 0x6A3F2A    # #2A3F6A em BGR

    # ── Cores de fonte por coluna (extraídas do template) ──────────────────
    # Col B (DU): branco
    ws_vert.range(f"B6:B{last_row}").api.Font.Color = 0xFFFFFF
    # Col C (Spot): default = azul-acinzentado (interpolado); Original = branco
    ws_vert.range(f"C6:C{last_row}").api.Font.Color = 0xD9C5B8
    # Col D (Forward): laranja em todas as linhas
    ws_vert.range(f"D6:D{last_row}").api.Font.Color = 0x129CF3   # BGR→ RGB(243,156,18)
    # Col E (Fator): azul-acinzentado
    ws_vert.range(f"E6:E{last_row}").api.Font.Color = 0xD9C5B8
    # Col F (Tipo): default = cor muted (interpolado)
    ws_vert.range(f"F6:F{last_row}").api.Font.Color = 0xA88B7A
    # Col G (Taxa 30d): azul-acinzentado
    ws_vert.range(f"G6:G{last_row}").api.Font.Color = 0xD9C5B8

    # Override linhas "Original": C=branco, F=verde
    COR_VERDE  = 0x71CC2E   # BGR → RGB(46,204,113)
    for i, row_data in enumerate(data):
        if row_data[4] == "Original":   # índice 4 = tipo
            r = 6 + i
            ws_vert.range(f"C{r}").api.Font.Color = 0xFFFFFF
            ws_vert.range(f"F{r}").api.Font.Color = COR_VERDE


# ──────────────────────────────────────────────────────────────────────────
# Aba Movimentos — variação em bps com formatação condicional
# ──────────────────────────────────────────────────────────────────────────

def _popular_movimentos(ws_mov, movimentos: pd.DataFrame,
                         data_ref: date, data_comp: date, analise: dict):
    # Subtítulo
    ws_mov.range("B3").value = (
        f"Comparação: {data_ref.strftime('%d/%m/%Y')} vs. "
        f"{data_comp.strftime('%d/%m/%Y')} (30 dias úteis)"
    )

    # Limpar dados antigos em bulk
    try:
        ws_mov.range("B6:H30").clear_contents()
    except Exception:
        for r in range(6, 30):
            try:
                ws_mov.range(f"B{r}:H{r}").clear_contents()
            except Exception:
                pass

    n = len(movimentos)
    ultima = 5 + n   # última linha de dados

    # Montar matrizes e escrever em bulk
    mov = movimentos.reset_index(drop=True)
    b_data = [[int(r["dias_uteis"])] for _, r in mov.iterrows()]
    cd_data = [[float(r["taxa_hoje"]), float(r["taxa_30d"])] for _, r in mov.iterrows()]
    f_data = [[str(r["direcao"])] for _, r in mov.iterrows()]

    ws_mov.range("B6").value = b_data          # coluna B
    ws_mov.range("C6").value = cd_data         # colunas C:D
    ws_mov.range("F6").value = f_data          # coluna F

    # Coluna E: fórmulas de variação — precisa ser célula a célula
    # (xlwings não suporta array de fórmulas com referências relativas via .value)
    for i in range(n):
        r = 6 + i
        ws_mov.range(f"E{r}").formula = f"=(C{r}-D{r})*10000"

    # Formatação em range inteiro (4 COM calls)
    data_range = ws_mov.range(f"B6:F{ultima}").api
    data_range.Interior.Color = 0x402216
    data_range.Font.Color = 0xFFFFFF

    ws_mov.range(f"C6:C{ultima}").number_format = "0,00%"
    ws_mov.range(f"D6:D{ultima}").number_format = "0,00%"
    ws_mov.range(f"E6:E{ultima}").number_format = "+0;-0;0"

    # Formatação condicional na coluna E: verde=positivo, vermelho=negativo
    _aplicar_cond_format_movimentos(ws_mov, first_row=6, last_row=ultima)

    # Cards resumo
    ws_mov.range("H5").value = "SHIFT PARALELO (MÉDIA)"
    ws_mov.range("H6").formula = f"=AVERAGE(E6:E{ultima})"
    ws_mov.range("H7").value = "bps (média todos vértices)"
    ws_mov.range("H8").value = "TIPO DE MOVIMENTO"
    ws_mov.range("H9").formula = (
        f'=IF(ABS(E{ultima}-E6)>10,IF(E{ultima}>E6,"Steepening","Flattening"),"Paralelo")'
    )
    ws_mov.range("H10").formula = (
        f'=CONCATENATE("Longo: ",TEXT(E{ultima},"+0;-0")," bps | Curto: ",TEXT(E6,"+0;-0")," bps")'
    )

    # Legenda H14 — explica a coluna Spot 30 DU atrás para leigos
    legenda_mov = (
        f"O que é a coluna 'Spot {data_comp.strftime('%d/%m/%Y')}'?\n"
        f"Ela mostra a taxa de juros para cada prazo conforme publicada pela ANBIMA "
        f"há 30 dias úteis (em {data_comp.strftime('%d/%m/%Y')}). "
        f"Não é uma taxa nova, é a mesma curva de hoje, só que fotografada no passado. "
        f"Compare com a coluna 'Spot hoje' para ver se os juros subiram ou caíram no período."
    )
    h14 = ws_mov.range("H14")
    h14.value = legenda_mov
    h14.api.WrapText = True
    h14.api.Font.Color = 0xD9C5B8
    h14.api.Font.Size = 9
    ws_mov.range("H:H").api.ColumnWidth = 52

    linha_media = ultima + 2
    ws_mov.range(f"B{linha_media}").value = "MÉDIA DA VARIAÇÃO"
    ws_mov.range(f"E{linha_media}").formula = f"=AVERAGE(E6:E{ultima})"
    ws_mov.range(f"E{linha_media}").number_format = "+0,0;-0,0;0"

    # Azul claro na linha da média (B:E) — #2E75B6 BGR=0xB6752E (2 COM calls)
    media_range = ws_mov.range(f"B{linha_media}:E{linha_media}").api
    media_range.Interior.Color = 0xB6752E
    media_range.Font.Color = 0xFFFFFF
    media_range.Font.Bold = True


def _aplicar_cond_format_movimentos(ws_mov, first_row: int, last_row: int):
    """
    Aplica formatação condicional à coluna E (variação bps):
      Verde (#1F4E2C interior) se > 0
      Vermelho (#4E1F1F interior) se < 0
    """
    try:
        from xlwings.utils import col_name
        range_e = ws_mov.range(f"E{first_row}:E{last_row}").api

        # Remover formatação condicional existente na coluna E
        range_e.FormatConditions.Delete()

        # Condição 1: valor > 0 → verde
        fc_pos = range_e.FormatConditions.Add(
            Type=1,       # xlCellValue
            Operator=5,   # xlGreater
            Formula1="0"
        )
        fc_pos.Interior.Color = 0x2C4E1F   # #1F4E2C em BGR: R=1F,G=4E,B=2C → 0x2C4E1F

        # Condição 2: valor < 0 → vermelho
        fc_neg = range_e.FormatConditions.Add(
            Type=1,
            Operator=6,   # xlLess
            Formula1="0"
        )
        fc_neg.Interior.Color = 0x1F1F4E   # #4E1F1F em BGR: R=4E,G=1F,B=1F → 0x1F1F4E

    except Exception as e:
        print(f"  [aviso] conditional format: {e}")


# ──────────────────────────────────────────────────────────────────────────
# Função principal
# ──────────────────────────────────────────────────────────────────────────

def construir_dashboard(
    curva_mensal: pd.DataFrame,
    movimentos: pd.DataFrame,
    analise: dict,
    formato_curva: str,
    data_ref: date,
    data_comp: date,
    template_path: str,
    output_path: str,
):
    """
    Copia o template, abre com xlwings e substitui dados mock por dados reais.
    Não recria formatação, gráficos ou estrutura do template.
    """
    # 1. Copiar template (fechar output se estiver aberto no Excel)
    if os.path.exists(output_path):
        try:
            # Tentar fechar instância aberta no Excel via COM
            for app_inst in xw.apps:
                for wb_open in app_inst.books:
                    if os.path.abspath(wb_open.fullname) == os.path.abspath(output_path):
                        wb_open.close()
        except Exception:
            pass
        try:
            os.remove(output_path)
        except PermissionError:
            raise PermissionError(
                f"\nFeche o arquivo '{os.path.basename(output_path)}' no Excel antes de rodar."
            )
    shutil.copy(template_path, output_path)
    print(f"  Template copiado -> {output_path}")

    # 2. Abrir com xlwings
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(output_path)
    wb.app.calculation = 'manual'   # recalcula só no save

    try:
        ws_dash = wb.sheets["Dashboard"]
        ws_aux  = wb.sheets["Aux"]
        ws_vert = wb.sheets["Vértices"]
        ws_mov  = wb.sheets["Movimentos"]

        taxa_curta = float(curva_mensal.iloc[0]["taxa_spot"])
        taxa_longa = float(curva_mensal.iloc[-1]["taxa_spot"])
        spread_bps = (taxa_longa - taxa_curta) * 10000

        t0 = time.time()

        # 3. Aux (lookup table mensal)
        print("  Populando Aux...", end=" ", flush=True)
        n_aux, n_ltn = _popular_aux(ws_aux, curva_mensal, data_ref, data_comp)
        print(f"{time.time()-t0:.1f}s")

        # 4. Dashboard — cards
        print("  Populando cards...", end=" ", flush=True)
        _popular_dashboard_cards(ws_dash, curva_mensal, data_ref, data_comp, spread_bps, formato_curva)
        print(f"{time.time()-t0:.1f}s")

        # 5. Dashboard — dados do gráfico (bimestral, rows 10-38)
        print("  Populando dados do grafico...", end=" ", flush=True)
        _popular_dashboard_chart(ws_dash, curva_mensal)
        print(f"{time.time()-t0:.1f}s")

        # 6. Atualizar dropdowns + I43
        print("  Atualizando dropdowns...", end=" ", flush=True)
        _atualizar_dashboard_extras(ws_dash, n_aux, n_ltn)
        print(f"{time.time()-t0:.1f}s")

        # 6b. Escala do eixo Y + corrigir series (excluir linhas de formulário)
        print("  Ajustando grafico...", end=" ", flush=True)
        _atualizar_chart_eixo_y(ws_dash, curva_mensal)
        _corrigir_series_chart(ws_dash)
        print(f"{time.time()-t0:.1f}s")

        # 7. Vértices
        print("  Populando Vertices...", end=" ", flush=True)
        _popular_vertice(ws_vert, curva_mensal, data_ref, data_comp)
        print(f"{time.time()-t0:.1f}s")

        # 8. Movimentos
        print("  Populando Movimentos...", end=" ", flush=True)
        _popular_movimentos(ws_mov, movimentos, data_ref, data_comp, analise)
        print(f"{time.time()-t0:.1f}s")

        # 9. Salvar (reativar cálculo antes de salvar)
        wb.app.calculation = 'automatic'
        app.screen_updating = True
        wb.save()
        print(f"  Salvo: {output_path}")

    finally:
        wb.close()
        app.quit()
