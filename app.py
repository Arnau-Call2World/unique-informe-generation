import streamlit as st
from io import BytesIO
from scripts.generate_word_pdf import main  # importa tu función

st.set_page_config(page_title="Generador de Informes", layout="centered")
st.title("📊 Generador de Informes Automáticos")

st.markdown("Sube los tres CSVs, el Excel y escribe el mes correspondiente para generar tu informe.")

csv1 = st.file_uploader("CSV 1", type=["csv"])
csv2 = st.file_uploader("CSV 2", type=["csv"])
csv3 = st.file_uploader("CSV 3", type=["csv"])
excel = st.file_uploader("Excel", type=["xlsx", "xls"])
mes = st.text_input("Mes (ej. 'enero')")

if st.button("Generar informe") and all([csv1, csv2, csv3, excel, mes]):
    with st.spinner("Generando informe..."):
        # Llama a tu función adaptada para aceptar archivos en memoria
        word_file, pdf_file = main(
            mes, csv1, csv2, csv3, excel
        )

        st.success("¡Informe generado!")
        st.download_button("📥 Descargar Word", word_file, file_name="informe.docx")
        st.download_button("📥 Descargar PDF", pdf_file, file_name="informe.pdf")
