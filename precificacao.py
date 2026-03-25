"""
precificacao.py — Precificação de título prefixado (LTN), Duration e DV01
Convenção brasileira: base 252 DU, composição anual
"""
from datetime import date
import pandas as pd
from interpolacao import interpolar_flat_forward


def dias_uteis_entre(data_ini: date, data_fim: date) -> int:
    """
    Conta dias úteis entre duas datas usando bizdays (calendário ANBIMA).
    Fallback para contagem simples (252/365) se bizdays não disponível.
    """
    try:
        import bizdays
        cal = bizdays.Calendar.load("ANBIMA")
        return cal.bizdays(data_ini, data_fim)
    except Exception:
        dc = (data_fim - data_ini).days
        return max(1, round(dc * 252 / 365))


def precificar_ltn(taxa: float, du: int, valor_face: float = 1000.0) -> dict:
    """
    Precifica uma LTN (zero-cupom prefixado) dados taxa e DU até vencimento.

    Parâmetros:
        taxa        : taxa spot (decimal, ex: 0.1455)
        du          : dias úteis até vencimento
        valor_face  : valor de face (padrão R$1.000,00)

    Retorna dict com: pu, duration, dv01, pu_stress, impacto_rs, impacto_pct
    """
    if du <= 0:
        du = 1

    # PU base
    pu = valor_face / ((1 + taxa) ** (du / 252))

    # Duration Macaulay (zero-cupom = prazo até vencimento em anos)
    duration = du / 252

    # DV01 via perturbação numérica
    pu_up   = valor_face / ((1 + taxa + 0.0001) ** (du / 252))
    pu_down = valor_face / ((1 + taxa - 0.0001) ** (du / 252))
    dv01 = (pu_down - pu_up) / 2

    # Stress Test +50 bps
    pu_stress  = valor_face / ((1 + taxa + 0.005) ** (du / 252))
    impacto_rs  = pu_stress - pu
    impacto_pct = impacto_rs / pu

    return {
        "pu":          round(pu, 4),
        "duration":    round(duration, 4),
        "dv01":        round(dv01, 4),
        "pu_stress":   round(pu_stress, 4),
        "impacto_rs":  round(impacto_rs, 4),
        "impacto_pct": round(impacto_pct, 6),
        "taxa":        taxa,
        "du":          du,
    }


def precificar_por_vencimento(
    vencimento: date,
    data_ref: date,
    vertices: pd.DataFrame,
    valor_face: float = 1000.0,
) -> dict:
    """Precifica LTN calculando DU até vencimento e interpolando taxa da curva."""
    du       = dias_uteis_entre(data_ref, vencimento)
    taxa     = interpolar_flat_forward(vertices, du)
    resultado = precificar_ltn(taxa, du, valor_face)
    resultado["vencimento"] = vencimento
    resultado["data_ref"]   = data_ref
    return resultado
