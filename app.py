import streamlit as st
from io import BytesIO
from scripts.generate_word_pdf import main  # importa tu función
import pandas as pd

st.set_page_config(page_title="Generador de Informes", layout="centered")
st.title("📊 Generador de Informes Automáticos")

st.markdown("Sube los tres CSVs, el Excel y escribe el mes correspondiente para generar tu informe.")

csv1 = st.file_uploader("categorias", type=["csv"])
df1 = pd.read_csv(csv1)
st.write("Categorías CSV preview", df1.head())
csv2 = st.file_uploader("dias", type=["csv"])
df2 = pd.read_csv(csv2)
st.write("Días CSV preview", df2.head())
csv3 = st.file_uploader("franjas", type=["csv"])
df3 = pd.read_csv(csv3)
st.write("Franjas CSV preview", df3.head())
excel = st.file_uploader("historico", type=["xlsx", "xls"])
mes = st.text_input("Mes (ej. 'ENERO')")

if st.button("Generar informe") and all([csv1, csv2, csv3, excel, mes]):
    with st.spinner("Generando informe..."):
        # Llama a tu función adaptada para aceptar archivos en memoria
        try:
            word_file, excel_file = main(mes, csv1, csv2, csv3, excel)
            print("Generado correctamente:", word_file, excel_file)
        except Exception as e:
            st.error(f"Error al generar el informe: {e}")

        st.success("¡Informe generado!")
        st.download_button("📥 Descargar Word", word_file, file_name="informe.docx")
        st.download_button(
        "📥 Descargar Excel actualizado",
        data=excel_file,
        file_name="informe_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
