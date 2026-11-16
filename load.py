from config import DATA_EXCEL_PATH, INT_SEMANA, MERCADO_INDICADORES
from utils_excel import refresh_workbook, run_macros, write_cells


MACROS_MERCADO_INDICADORES = ["GenerarPPT_DosVeces", "Test_ProcesarTodoMetricas_Ambos"]


def update(semana=INT_SEMANA):
    """
    Updates the Excel chain and generates the outputs:
    1. Writes the week to evaluate in Tabla1!O2 of MERCADO_INDICADORES.
    2. Refreshes DATA_EXCEL_PATH (Power Query over ouput/DATA.csv).
    3. Refreshes MERCADO_INDICADORES (Power Query over DATA_EXCEL_PATH).
    4. Runs the MERCADO_INDICADORES macros.

    Args:
        semana (int): Week to evaluate (SEMANA_EVALUAR).
    """
    write_cells(MERCADO_INDICADORES, "Tabla1", {"O2": semana})
    refresh_workbook(DATA_EXCEL_PATH)
    refresh_workbook(MERCADO_INDICADORES)
    run_macros(MERCADO_INDICADORES, MACROS_MERCADO_INDICADORES)
