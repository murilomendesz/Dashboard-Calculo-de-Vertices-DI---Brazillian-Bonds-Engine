"""
buscador_dados.py — Coleta de dados ANBIMA via pyettj e histórico XLSX
"""
import os
import pandas as pd
from datetime import date, timedelta
import logging

logger = logging.getLogger(__name__)

# ── Vértices padrão DU ───────────────────────────────────────────────────
# 15 pontos: curto (21-84), semestral (126-1260).
# Reduz o erro de interpolação de ~2.7 bps (10 vértices) para ~0.7 bps RMSE.
VERTICES_PADRAO_DU = [
    21, 42, 63, 84, 126, 189, 252, 378,
    504, 630, 756, 882, 1008, 1134, 1260,
]

# 9 vértices "de mercado" usados na aba Movimentos (comparação hoje vs 30d)
VERTICES_MOVIMENTOS = {21, 42, 63, 126, 252, 504, 756, 1008, 1260}

# ── Dados estáticos de fallback ────────────────────────────────────────────
FALLBACK_VERTICES = [
    (21,   0.1495), (42,  0.1492), (63,  0.1489), (84,  0.1484),
    (126,  0.1478), (189, 0.1470), (252, 0.1462), (378, 0.1454),
    (504,  0.1446), (630, 0.1440), (756, 0.1435), (882, 0.1431),
    (1008, 0.1427), (1134, 0.1426), (1260, 0.1425),
]
FALLBACK_VERTICES_30D = [
    (21,   0.1460), (42,  0.1458), (63,  0.1455), (84,  0.1450),
    (126,  0.1442), (189, 0.1436), (252, 0.1430), (378, 0.1422),
    (504,  0.1415), (630, 0.1410), (756, 0.1405), (882, 0.1401),
    (1008, 0.1398), (1134, 0.1396), (1260, 0.1395),
]

NOME_SHEET    = "Historico"
COLUNAS_XLSX  = ["Data Referencia", "Prazo (DU)", "Taxa Spot (% a.a.)"]


# ── Helpers ────────────────────────────────────────────────────────────────

def _data_str(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _dia_util_anterior(d: date, n: int) -> date:
    """Retorna o dia útil n dias úteis atrás usando calendário ANBIMA (bizdays)."""
    try:
        import bizdays
        cal = bizdays.Calendar.load("ANBIMA")
        result = cal.offset(d, -n)
        return result.date() if hasattr(result, "date") else result
    except Exception:
        logger.warning("bizdays indisponível — usando contagem simples de dias úteis.")
        current = d
        count = 0
        while count < n:
            current -= timedelta(days=1)
            if current.weekday() < 5:
                count += 1
        return current


# ── Histórico XLSX ─────────────────────────────────────────────────────────

def _ler_historico_xlsx(historico_path: str) -> pd.DataFrame | None:
    """Lê o histórico xlsx. Retorna DataFrame com [data_referencia, dias_uteis, taxa] ou None."""
    if not os.path.exists(historico_path):
        return None
    try:
        df = pd.read_excel(historico_path, sheet_name=NOME_SHEET, engine="openpyxl")
        df.columns = ["data_referencia", "dias_uteis", "taxa"]
        df["data_referencia"] = pd.to_datetime(df["data_referencia"]).dt.date
        df["dias_uteis"] = pd.to_numeric(df["dias_uteis"], errors="coerce").astype("Int64")
        df["taxa"]       = pd.to_numeric(df["taxa"], errors="coerce")
        return df.dropna().reset_index(drop=True)
    except Exception as e:
        logger.warning(f"Erro lendo histórico xlsx: {e}")
        return None


def _escrever_historico_xlsx(df_total: pd.DataFrame, historico_path: str):
    """Reescreve o histórico xlsx completo com formatação dark navy."""
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    os.makedirs(os.path.dirname(historico_path), exist_ok=True)

    COR_HEADER  = "243356"
    CORES_DADOS = ["162240", "1C2B4A"]   # alternância por data

    borda_lado = Side(style="thin", color="2A3F6A")
    borda  = Border(left=borda_lado, right=borda_lado,
                    top=borda_lado, bottom=borda_lado)
    centro = Alignment(horizontal="center", vertical="center")

    header_fill = PatternFill("solid", fgColor=COR_HEADER)
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    data_font   = Font(color="FFFFFF", name="Calibri", size=10)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = NOME_SHEET

    # Cabeçalhos
    for col, nome in enumerate(COLUNAS_XLSX, 1):
        c = ws.cell(row=1, column=col, value=nome)
        c.fill      = header_fill
        c.font      = header_font
        c.alignment = centro
        c.border    = borda

    # Índice de cor por data (alternância)
    datas_unicas = sorted(df_total["data_referencia"].unique())
    cor_por_data = {d: i % 2 for i, d in enumerate(datas_unicas)}

    # Dados
    for row_idx, row in enumerate(df_total.itertuples(index=False), start=2):
        fill = PatternFill("solid", fgColor=CORES_DADOS[cor_por_data[row.data_referencia]])

        c1 = ws.cell(row=row_idx, column=1, value=str(row.data_referencia))
        c2 = ws.cell(row=row_idx, column=2, value=int(row.dias_uteis))
        c3 = ws.cell(row=row_idx, column=3, value=float(row.taxa))
        c3.number_format = "0.00%"

        for c in [c1, c2, c3]:
            c.fill      = fill
            c.font      = data_font
            c.alignment = centro
            c.border    = borda

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 22
    ws.freeze_panes = "A2"
    ws.sheet_properties.tabColor = "243356"

    wb.save(historico_path)


def _migrar_csv_para_xlsx(csv_path: str, xlsx_path: str):
    """Migra o CSV legado para xlsx formatado."""
    try:
        df = pd.read_csv(csv_path)
        df.columns = ["data_referencia", "dias_uteis", "taxa"]
        df["data_referencia"] = pd.to_datetime(df["data_referencia"]).dt.date
        df = df.sort_values(["data_referencia", "dias_uteis"]).reset_index(drop=True)
        _escrever_historico_xlsx(df, xlsx_path)
        logger.info(f"CSV migrado para xlsx: {len(df)} registros. "
                    f"O arquivo '{csv_path}' pode ser deletado manualmente.")
    except Exception as e:
        logger.warning(f"Migração CSV→xlsx falhou: {e}")


def salvar_historico(data: date, curva: pd.DataFrame, historico_path: str):
    """Salva curva no histórico xlsx (sem duplicar). Migra CSV legado se necessário."""
    # Migrar CSV legado se xlsx ainda não existe
    if not os.path.exists(historico_path):
        csv_path = historico_path.replace(".xlsx", ".csv")
        if os.path.exists(csv_path):
            _migrar_csv_para_xlsx(csv_path, historico_path)

    # Verificar duplicata
    df_existente = _ler_historico_xlsx(historico_path)
    if df_existente is not None and data in df_existente["data_referencia"].values:
        return

    # Montar DataFrame total e reescrever
    df_novo = curva[["dias_uteis", "taxa"]].copy()
    df_novo.insert(0, "data_referencia", data)

    if df_existente is not None:
        df_total = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_total = df_novo

    df_total = df_total.sort_values(["data_referencia", "dias_uteis"]).reset_index(drop=True)
    _escrever_historico_xlsx(df_total, historico_path)
    logger.info(f"Histórico xlsx atualizado: {data} ({len(curva)} vértices)")


def _buscar_historico(data: date, historico_path: str) -> pd.DataFrame | None:
    """Busca curva no histórico local. Retorna None se entrada tiver poucos vértices."""
    df = _ler_historico_xlsx(historico_path)
    if df is None:
        return None
    subset = df[df["data_referencia"] == data][["dias_uteis", "taxa"]].reset_index(drop=True)
    min_vertices = max(5, len(VERTICES_PADRAO_DU) - 2)
    if len(subset) >= min_vertices:
        logger.info(f"[historico] {data}: {len(subset)} vértices")
        return subset
    if len(subset) > 0:
        logger.info(f"[historico] {data}: {len(subset)} vértices (insuficiente, re-buscando)")
    return None


def _buscar_data_mais_proxima(data_alvo: date, historico_path: str) -> tuple[pd.DataFrame | None, date | None]:
    """Busca a data mais próxima disponível no histórico."""
    df = _ler_historico_xlsx(historico_path)
    if df is None:
        return None, None
    datas = sorted(df["data_referencia"].unique())
    if not datas:
        return None, None
    anteriores = [d for d in datas if d <= data_alvo]
    data_escolhida = max(anteriores) if anteriores else min(datas)
    subset = df[df["data_referencia"] == data_escolhida][["dias_uteis", "taxa"]].reset_index(drop=True)
    if len(subset) >= 5:
        return subset, data_escolhida
    return None, None


def _buscar_pyettj(data: date) -> pd.DataFrame | None:
    """Busca via pyettj — extrai taxas nos vértices padrão DU."""
    try:
        from pyettj import ettj
        df_raw = ettj.get_ettj(_data_str(data), curva="PRE")
        if df_raw is None or len(df_raw) == 0:
            return None
        col_dc   = df_raw.columns[0]
        col_taxa = df_raw.columns[1]
        df_raw = df_raw[[col_dc, col_taxa]].dropna().copy()
        df_raw.columns = ["dc", "taxa_pct"]
        df_raw["dc"]      = pd.to_numeric(df_raw["dc"],      errors="coerce")
        df_raw["taxa_pct"] = pd.to_numeric(df_raw["taxa_pct"], errors="coerce")
        df_raw = df_raw.dropna()
        if len(df_raw) < len(VERTICES_PADRAO_DU):
            return None
        # Calcular DC correto para cada DU usando calendário ANBIMA (bizdays)
        # Fallback: aproximação 365/252 se bizdays indisponível
        try:
            import bizdays
            cal = bizdays.Calendar.load("ANBIMA")
            def _du_para_dc(ref: date, du: int) -> int:
                venc = cal.offset(ref, du)
                venc = venc.date() if hasattr(venc, "date") else venc
                return (venc - ref).days
        except Exception:
            logger.warning("bizdays indisponível — usando 365/252 para DC.")
            def _du_para_dc(ref: date, du: int) -> int:
                return round(du * 365 / 252)

        rows = []
        for du in VERTICES_PADRAO_DU:
            dc_alvo = _du_para_dc(data, du)
            idx  = (df_raw["dc"] - dc_alvo).abs().idxmin()
            taxa = float(df_raw.loc[idx, "taxa_pct"]) / 100.0
            rows.append({"dias_uteis": du, "taxa": taxa})
        result = pd.DataFrame(rows)
        logger.info(f"[pyettj] {data}: {len(result)} vértices")
        return result
    except Exception as e:
        logger.warning(f"pyettj falhou para {data}: {e}")
        return None



def _fallback_estatico(tipo: str = "hoje") -> pd.DataFrame:
    """Dados estáticos como último recurso."""
    logger.warning("Usando dados estáticos de fallback.")
    vertices = FALLBACK_VERTICES if tipo == "hoje" else FALLBACK_VERTICES_30D
    return pd.DataFrame(vertices, columns=["dias_uteis", "taxa"])


# ── Funções públicas ───────────────────────────────────────────────────────

def buscar_curva(data: date, historico_path: str) -> pd.DataFrame:
    """Busca curva ETTJ. Ordem: pyettj hoje → pyettj dia anterior → histórico local → fallback."""
    df = _buscar_pyettj(data)
    if df is not None:
        salvar_historico(data, df, historico_path)
        return df
    # ANBIMA ainda não publicou hoje — busca o dia útil anterior direto no pyettj
    data_anterior = _dia_util_anterior(data, 1)
    df = _buscar_pyettj(data_anterior)
    if df is not None:
        salvar_historico(data_anterior, df, historico_path)
        logger.warning(
            f"ANBIMA ainda não publicou dados para {data.strftime('%d/%m/%Y')} — "
            f"usando dados do dia anterior ({data_anterior.strftime('%d/%m/%Y')})."
        )
        return df
    df = _buscar_historico(data, historico_path)
    if df is not None:
        return df
    return _fallback_estatico("hoje")


def buscar_curva_comparacao(data_ref: date, dias_atras: int, historico_path: str) -> tuple[pd.DataFrame, date]:
    """Busca curva de comparação (N dias úteis atrás). Retorna (curva, data_efetiva)."""
    data_alvo = _dia_util_anterior(data_ref, dias_atras)
    logger.info(f"Data comparação: {data_alvo}")

    df = _buscar_historico(data_alvo, historico_path)
    if df is not None:
        return df, data_alvo

    df = _buscar_pyettj(data_alvo)
    if df is not None:
        salvar_historico(data_alvo, df, historico_path)
        return df, data_alvo

    df, data_real = _buscar_data_mais_proxima(data_alvo, historico_path)
    if df is not None:
        logger.info(f"Usando data mais próxima: {data_real}")
        return df, data_real

    return _fallback_estatico("30d"), data_alvo
