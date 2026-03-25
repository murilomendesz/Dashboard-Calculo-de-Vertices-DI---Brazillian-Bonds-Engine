"""
main.py — Ponto de entrada único do projeto Curva DI

Uso:
    python main.py

Fluxo:
    1. Busca curva ETTJ de hoje via ANBIMA (pyettj)
    2. Busca curva de ~30 dias úteis atrás (histórico ou pyettj)
    3. Interpola mensalmente (Flat Forward, 21 DU/ponto)
    4. Calcula curva forward mês a mês
    5. Classifica formato da curva
    6. Calcula movimentos (variação em bps)
    7. Popula dashboard Excel e abre o arquivo
"""
import time
import logging
import subprocess

# ── Setup logging ─────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── Imports do projeto ────────────────────────────────────────────────────
import configuracao
from buscador_dados   import buscar_curva, buscar_curva_comparacao, VERTICES_MOVIMENTOS
from interpolacao     import construir_curva_completa
from analise          import classificar_formato, calcular_movimentos, analisar_shift_steepening
from construtor_excel import construir_dashboard


def main():
    print("=" * 60)
    print("  CURVA DE JUROS DI — Estrutura a Termo (ETTJ Pre)")
    print("  Dados: ANBIMA | Metodo: Flat Forward 252")
    print("=" * 60)

    t_inicio = time.time()

    # ── 1. Datas ──────────────────────────────────────────────────────────
    data_ref = configuracao.DATA_REFERENCIA
    print(f"\n[1/7] Data referencia: {data_ref.strftime('%d/%m/%Y')}")

    # ── 2. Buscar curvas ──────────────────────────────────────────────────
    print("\n[2/7] Buscando curva de hoje...")
    curva_hoje_raw = buscar_curva(data_ref, configuracao.HISTORICO_PATH)
    print(f"      {len(curva_hoje_raw)} vertices carregados")

    print("\n[3/7] Buscando curva de comparacao (~30 DU atras)...")
    curva_30d_raw, data_comp = buscar_curva_comparacao(
        data_ref, configuracao.DIAS_COMPARACAO, configuracao.HISTORICO_PATH
    )
    print(f"      Data comparacao: {data_comp.strftime('%d/%m/%Y')} | {len(curva_30d_raw)} vertices")

    # ── 3. Interpolar mensalmente ─────────────────────────────────────────
    print("\n[4/7] Interpolando curvas mensalmente (Flat Forward)...")
    curva_mensal = construir_curva_completa(
        curva_hoje_raw,
        du_min=configuracao.DU_MIN,
        du_max=configuracao.DU_MAX,
        step=configuracao.DU_STEP,
    )
    curva_30d_mensal = construir_curva_completa(
        curva_30d_raw,
        du_min=configuracao.DU_MIN,
        du_max=configuracao.DU_MAX,
        step=configuracao.DU_STEP,
    )
    curva_mensal = curva_mensal.merge(
        curva_30d_mensal[["dias_uteis", "taxa_spot"]].rename(columns={"taxa_spot": "taxa_30d"}),
        on="dias_uteis",
        how="left",
    )
    print(f"      {len(curva_mensal)} pontos mensais gerados (DU {configuracao.DU_MIN} a {configuracao.DU_MAX})")

    # ── 4. Análise da curva ───────────────────────────────────────────────
    print("\n[5/7] Analisando curva...")
    taxa_curta = float(curva_mensal.iloc[0]["taxa_spot"])
    taxa_longa = float(curva_mensal.iloc[-1]["taxa_spot"])
    spread_bps = (taxa_longa - taxa_curta) * 10000
    formato_curva = classificar_formato(taxa_curta, taxa_longa, configuracao.LIMIAR_FLAT_BPS)

    print(f"      Taxa curta:  {taxa_curta*100:.2f}% a.a.")
    print(f"      Taxa longa:  {taxa_longa*100:.2f}% a.a.")
    print(f"      Spread:      {spread_bps:+.0f} bps")
    print(f"      Formato:     {formato_curva}")

    # ── 5. Movimentos (variação em bps por vértice ANBIMA) ────────────────
    print("\n[6/7] Calculando movimentos...")
    curva_hoje_mov = curva_hoje_raw[curva_hoje_raw["dias_uteis"].isin(VERTICES_MOVIMENTOS)].reset_index(drop=True)
    curva_30d_mov  = curva_30d_raw[curva_30d_raw["dias_uteis"].isin(VERTICES_MOVIMENTOS)].reset_index(drop=True)
    movimentos = calcular_movimentos(curva_hoje_mov, curva_30d_mov)
    analise    = analisar_shift_steepening(movimentos)

    print(f"      Shift paralelo: {analise['shift_bps']:+.1f} bps")
    print(f"      Tipo:           {analise['tipo_movimento']}")
    print(f"      Longo: {analise['var_longo_bps']:+.1f} bps | Curto: {analise['var_curto_bps']:+.1f} bps")

    # ── 6. Construir dashboard Excel ──────────────────────────────────────
    print("\n[7/7] Construindo dashboard Excel...")
    construir_dashboard(
        curva_mensal   = curva_mensal,
        movimentos     = movimentos,
        analise        = analise,
        formato_curva  = formato_curva,
        data_ref       = data_ref,
        data_comp      = data_comp,
        template_path  = configuracao.TEMPLATE_PATH,
        output_path    = configuracao.OUTPUT_PATH,
    )

    # ── 7. Abrir o arquivo Excel ──────────────────────────────────────────
    print("\n" + "=" * 60)
    print(f"  Dashboard gerado: {configuracao.OUTPUT_PATH}")
    print("=" * 60)

    try:
        time.sleep(4)
        subprocess.Popen(f'start "" "{configuracao.OUTPUT_PATH}"', shell=True)
        print("  Excel aberto automaticamente.")
    except Exception as e:
        print(f"  Erro ao abrir Excel ({e}). Abra manualmente: {configuracao.OUTPUT_PATH}")

    print(f"\nConcluido em {time.time() - t_inicio:.1f}s.")


if __name__ == "__main__":
    main()
