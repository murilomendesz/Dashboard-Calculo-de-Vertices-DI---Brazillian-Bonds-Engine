"""
configuracao.py — Configurações globais do projeto Curva DI
"""
from datetime import date
import os

# Datas
DATA_REFERENCIA  = date.today()
DIAS_COMPARACAO  = 30   # dias úteis para buscar curva histórica

# Threshold para classificação de formato da curva
LIMIAR_FLAT_BPS = 30

# Caminhos
BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH  = os.path.join(BASE_DIR, "template", "curva_di_template.xlsx")
OUTPUT_PATH    = os.path.join(BASE_DIR, "curva_di_dashboard.xlsx")
HISTORICO_PATH = os.path.join(BASE_DIR, "historico", "curvas_historicas.xlsx")

# Parâmetros de interpolação mensal
DU_STEP = 21    # ~1 mês em dias úteis
DU_MIN  = 21    # 1 mês à frente
DU_MAX  = 1218  # 58 meses à frente
