import os 
import sys

import streamlit as st


# Añadir el directorio padre al path para importar correctamente
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.read_and_prepare_data import leer_tablas
import pandas as pd
import matplotlib.pyplot as plt

def sacar_variables_informe(
    categorias_df: pd.DataFrame,
    categorias_anterior_df: pd.DataFrame,
) -> tuple:
    total_recibidas = categorias_df[categorias_df["Categoría"] == "Total"]["Recibidas"].astype(int).iloc[0]
    total_no_atendidas = total_recibidas - categorias_df[categorias_df["Categoría"] == "Total"]["Atendidas_num"].astype(int).iloc[0]
    desbordadas = categorias_df[categorias_df["Categoría"] == "Total"]["Desborde_cantidad"].astype(int).iloc[0]
    abandonadas = categorias_df[categorias_df["Categoría"] == "Total"]["Abandonadas"].astype(int).iloc[0]

    porcentaje_mes = categorias_df[categorias_df["Categoría"] == "Total"]["Atendidas_%"].astype(float).iloc[0]
    porcentaje_mes_anterior = categorias_anterior_df[categorias_anterior_df["Categoría"] == "Total"]["Atendidas_%"].astype(float).iloc[0]

    cumple_objetivo = porcentaje_mes >= 85.0

    categorias_conflictivas = list(
        categorias_df.iloc[:-1][
            (categorias_df["Atendidas_%"] > 0) & (categorias_df["Atendidas_%"] < 85)
        ][["Categoría", "Atendidas_%"]]
        .itertuples(index=False, name=None)
    )


    return (
        porcentaje_mes, porcentaje_mes_anterior, cumple_objetivo,
        total_recibidas, total_no_atendidas, desbordadas, abandonadas,
        categorias_conflictivas
    )

def generar_grafico_meses_barplot_dinamico(
    tablas: dict[str, pd.DataFrame],
    meses: list[str]
) -> plt.Figure:
    """
    Genera un gráfico de barras para todos los meses disponibles en la lista `meses`,
    comparando el porcentaje de atendidas.
    Devuelve la figura (sin mostrarla).
    """
    valores = [
        tablas[mes][tablas[mes]["Categoría"] == "Total"]["Atendidas_%"].astype(float).iloc[0]
        for mes in meses
    ]
    colores = ['green' if v >= 85 else 'red' for v in valores]

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(meses, valores, color=colores)
    ax.set_title("Estadísticas Año 2025")
    ax.set_ylim(0, 100)
    ax.grid(axis='y', linestyle='--', alpha=0.7)

    for bar, val in zip(bars, valores):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                f"{val:.2f}%", ha='center', va='bottom', fontsize=10)

    return fig

def generar_variables_informe(
    path_paquete: str,
    meses: list[str],
    mes_actual: str
) -> tuple:
    """
    Genera las variables necesarias para el informe a partir de las tablas de categorías.
    """

    st.write("📌 Paso 1: empezando a leer tablas...")

    # ← Esto puede lanzar el error
    tablas, tablas_mes, excel = leer_tablas(path_paquete, meses)
    st.write("✅ Paso 2: tablas leídas correctamente")

    categorias_df = tablas['categorias']
    st.write("✅ Paso 3: df de categorías leído", categorias_df.head())

    categorias_anterior_df = tablas_mes[meses[-2]]
    st.write("✅ Paso 4: df del mes anterior leído", categorias_anterior_df.head())

    variables_informe = sacar_variables_informe(
        categorias_df, 
        categorias_anterior_df
    )
    st.write("✅ Paso 5: variables del informe calculadas")

    tablas_mes[mes_actual] = categorias_df
    fig = generar_grafico_meses_barplot_dinamico(tablas_mes, meses)
    st.write("✅ Paso 6: gráfico generado")

    return variables_informe, fig, tablas, excel
if __name__ == "__main__":
    MESES = [ "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                       "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" ]
    path_excel = "data/INFORME DE LLAMADAS AUTOCLIMA ABRIL 2025.xlsx"
    nombre_hoja = "ABRIL"
    meses_hasta_ahora = MESES[:MESES.index(nombre_hoja) + 1]

    variables_informe, fig, _  = generar_variables_informe(path_excel, meses_hasta_ahora)
    print(variables_informe)