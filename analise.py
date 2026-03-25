"""
analise.py — Análise de movimentos da curva (bps, Shift/Steepening, formato)
"""
import pandas as pd
from datetime import date


def classificar_formato(taxa_curta: float, taxa_longa: float, limiar_bps: int = 30) -> str:
    """
    Classifica o formato da curva baseado no spread curto-longo.
    Retorna: 'Normal', 'Invertida' ou 'Flat'
    """
    spread_bps = (taxa_longa - taxa_curta) * 10000
    if spread_bps > limiar_bps:
        return "Normal"
    elif spread_bps < -limiar_bps:
        return "Invertida"
    else:
        return "Flat"


def calcular_spread_bps(taxa_curta: float, taxa_longa: float) -> float:
    """Retorna spread em basis points (longa - curta)."""
    return (taxa_longa - taxa_curta) * 10000


def calcular_movimentos(curva_hoje: pd.DataFrame, curva_anterior: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula variação em bps por vértice entre duas curvas.

    Parâmetros:
        curva_hoje     : DataFrame ['dias_uteis', 'taxa']
        curva_anterior : DataFrame ['dias_uteis', 'taxa']

    Retorna DataFrame com:
        dias_uteis, taxa_hoje, taxa_30d, variacao_bps, direcao
    """
    df = pd.merge(
        curva_hoje.rename(columns={"taxa": "taxa_hoje"}),
        curva_anterior.rename(columns={"taxa": "taxa_30d"}),
        on="dias_uteis",
        how="inner",
    )
    df["variacao_bps"] = (df["taxa_hoje"] - df["taxa_30d"]) * 10000
    df["direcao"] = df["variacao_bps"].apply(
        lambda x: "+ Abriu" if x > 0 else ("- Fechou" if x < 0 else "= Estavel")
    )
    return df.sort_values("dias_uteis").reset_index(drop=True)


def analisar_shift_steepening(movimentos: pd.DataFrame) -> dict:
    """
    Decompõe o movimento da curva em Shift paralelo e Steepening/Flattening.

    Parâmetros:
        movimentos : DataFrame retornado por calcular_movimentos()

    Retorna dict com shift_bps, inclinacao_bps, tipo_movimento
    """
    if len(movimentos) == 0:
        return {"shift_bps": 0.0, "inclinacao_bps": 0.0, "tipo_movimento": "N/A",
                "var_curto_bps": 0.0, "var_longo_bps": 0.0}

    shift_bps = float(movimentos["variacao_bps"].mean())
    var_curto = float(movimentos.loc[movimentos["dias_uteis"].idxmin(), "variacao_bps"])
    var_longo = float(movimentos.loc[movimentos["dias_uteis"].idxmax(), "variacao_bps"])
    inclinacao_bps = var_longo - var_curto

    if inclinacao_bps > 10:
        tipo = "Steepening"
    elif inclinacao_bps < -10:
        tipo = "Flattening"
    else:
        tipo = "Paralelo"

    return {
        "shift_bps":       round(shift_bps, 1),
        "inclinacao_bps":  round(inclinacao_bps, 1),
        "tipo_movimento":  tipo,
        "var_curto_bps":   round(var_curto, 1),
        "var_longo_bps":   round(var_longo, 1),
    }
