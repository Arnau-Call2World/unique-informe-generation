import streamlit as st
from io import BytesIO
from scripts.generate_word_pdf import main  # importa tu función

st.set_page_config(page_title="Generador de Informes", layout="centered")
st.title("📊 Generador de Informes Automáticos")

st.markdown("Sube los tres CSVs, el Excel y escribe el mes correspondiente para generar tu informe.")

csv1 = st.file_uploader("categorias", type=["csv"])
csv2 = st.file_uploader("dias", type=["csv"])
csv3 = st.file_uploader("franjas", type=["csv"])
excel = st.file_uploader("historico", type=["xlsx", "xls"])
mes = st.text_input("Mes (ej. 'ENERO')")

if st.button("Generar informe") and all([csv1, csv2, csv3, excel, mes]):
    with st.spinner("Generando informe..."):
        # Llama a tu función adaptada para aceptar archivos en memoria
        word_file, excel_file = main(
            mes, csv1, csv2, csv3, excel
        )

        st.success("¡Informe generado!")
        st.download_button("📥 Descargar Word", word_file, file_name="informe.docx")
        st.download_button(
        "📥 Descargar Excel actualizado",
        data=excel_file,
        file_name="informe_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
