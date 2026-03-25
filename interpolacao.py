"""
interpolacao.py — Interpolação Flat Forward (método padrão B3/ANBIMA)
Base 252 dias úteis (convenção brasileira)
"""
import pandas as pd
import numpy as np
from datetime import date


def interpolar_flat_forward(vertices: pd.DataFrame, du_alvo: int) -> float:
    """
    Interpola a taxa spot para du_alvo usando Flat Forward entre os vértices adjacentes.

    Parâmetros:
        vertices : DataFrame com colunas ['dias_uteis', 'taxa'] (taxa em decimal)
        du_alvo  : dias úteis do prazo desejado

    Retorna:
        taxa interpolada (decimal, ex: 0.1489)
    """
    verts = vertices.sort_values("dias_uteis").reset_index(drop=True)
    dus   = verts["dias_uteis"].values
    taxas = verts["taxa"].values

    # Extrapolação: fora do range usa o extremo mais próximo
    if du_alvo <= dus[0]:
        return float(taxas[0])
    if du_alvo >= dus[-1]:
        return float(taxas[-1])

    # Achar índice do vértice anterior
    idx = np.searchsorted(dus, du_alvo, side="right") - 1
    du_a, r_a = int(dus[idx]),     float(taxas[idx])
    du_p, r_p = int(dus[idx + 1]), float(taxas[idx + 1])

    # Flat Forward (padrão ANBIMA/B3 — interpolação sobre fatores de acumulação):
    # (1+r_i)^(du_i/252) = (1+r_a)^(du_a/252) * [(1+r_p)^(du_p/252)/(1+r_a)^(du_a/252)]^α
    # onde α = (du_i - du_a) / (du_p - du_a)
    alpha   = (du_alvo - du_a) / (du_p - du_a)
    fator_a = (1 + r_a) ** (du_a / 252)
    fator_p = (1 + r_p) ** (du_p / 252)
    fator_i = fator_a * (fator_p / fator_a) ** alpha
    r_i     = fator_i ** (252 / du_alvo) - 1
    return float(r_i)


def calcular_forward_par(r_a: float, du_a: int, r_b: float, du_b: int) -> float:
    """
    Calcula a taxa forward implícita entre dois vértices A e B (du_B > du_A).
    Taxa a termo no período [du_A, du_B], base 252 DU.
    """
    if du_b <= du_a:
        raise ValueError("du_b deve ser maior que du_a")
    fator_a   = (1 + r_a) ** (du_a / 252)
    fator_b   = (1 + r_b) ** (du_b / 252)
    fator_fwd = fator_b / fator_a
    taxa_fwd  = fator_fwd ** (252 / (du_b - du_a)) - 1
    return float(taxa_fwd)


def construir_curva_completa(
    vertices: pd.DataFrame,
    du_min: int,
    du_max: int,
    step: int = 21,
) -> pd.DataFrame:
    """
    Constrói curva interpolada mensalmente de du_min a du_max.

    Retorna DataFrame com colunas:
        dias_uteis, taxa_spot, tipo ('Original' | 'Interpolado'), forward
    """
    verts_sorted  = vertices.sort_values("dias_uteis").reset_index(drop=True)
    dus_originais = set(verts_sorted["dias_uteis"].values.tolist())

    rows = []
    for du in range(du_min, du_max + 1, step):
        taxa = interpolar_flat_forward(verts_sorted, du)
        tipo = "Original" if du in dus_originais else "Interpolado"
        rows.append({"dias_uteis": du, "taxa_spot": taxa, "tipo": tipo})

    curva = pd.DataFrame(rows)

    # Calcular forward mês a mês
    forwards = [None]
    for i in range(1, len(curva)):
        fwd = calcular_forward_par(
            curva.loc[i - 1, "taxa_spot"], curva.loc[i - 1, "dias_uteis"],
            curva.loc[i,     "taxa_spot"], curva.loc[i,     "dias_uteis"],
        )
        forwards.append(fwd)

    curva["forward"] = forwards
    curva.loc[0, "forward"] = curva.loc[0, "taxa_spot"]   # primeiro mês sem período anterior

    return curva


def du_para_vencimento_label(du: int, data_ref: date, step: int = 21) -> str:
    """Converte dias úteis em label legível de vencimento (ex: 'Abr/2026 (21 DU)')."""
    meses    = round(du / step)
    ano, mes = data_ref.year, data_ref.month + meses
    while mes > 12:
        mes -= 12
        ano += 1
    meses_pt = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
                 "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    return f"{meses_pt[mes - 1]}/{ano} ({du} DU)"
