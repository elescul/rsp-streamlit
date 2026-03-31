from matplotlib.pyplot import rc
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from num2words import num2words
from io import BytesIO
import zipfile
import os

# 👉 BASE PATH (CLAVE PARA CLOUD)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def procesar(archivo_excel, per, parametro, dias_num):
    zip_buffer = BytesIO()
    dias_letra = num2words(int(dias_num), lang='es').upper()
    
    year = per[:4]
    peri = per[-1]
    if peri == '1':
        fecha_per = f'31.03.{year}'
    elif peri == '2':
        fecha_per = f'30.06.{year}'
    elif peri == '3':
        fecha_per = f'30.09.{year}'
    elif peri == '4':
        fecha_per = f'31.12.{year}'

    try:
        den = pd.read_excel(os.path.join(BASE_DIR, 'data', 'dataset_cias.xlsx'), sheet_name='Activas')
        auditores = pd.read_excel(os.path.join(BASE_DIR, 'data', 'dataset_auditores.xlsx'), sheet_name='cias')
    except Exception as e:
        st.error(f"Error cargando datasets auxiliares: {e}")
        return None

    df_resumen = pd.read_excel(archivo_excel, sheet_name="Resumen", header=3, usecols="B:AM")
    cuadros_auto = pd.read_excel(archivo_excel, sheet_name="Cuadros estado automotor", header=1, usecols="B:M")
    cuadros_moto = pd.read_excel(archivo_excel, sheet_name="Cuadros estado motovehicular", header=1, usecols="B:M")
    cuadros_rc = pd.read_excel(archivo_excel, sheet_name="Cuadros estado RC", header=1, usecols="B:M")
    cuadros_resto_ramas = pd.read_excel(archivo_excel, sheet_name="Cuadros estado Resto de ramas", header=1, usecols="B:M")
    cuadros_indet = pd.read_excel(archivo_excel, sheet_name="Cuadro Estado Demandas Indeterm")
    cuadros_mediaciones = pd.read_excel(archivo_excel, sheet_name="Cuadro Estado Mediaciones")

    cias_caen = df_resumen[df_resumen['% de consumo de superavit'] > parametro]['Cod. Cía'].unique()

    def format_mil(num):
        try:
            return f"{float(num):,.0f}".replace(",", ".")
        except:
            return "0"

    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for cia in cias_caen:
            try:
                cod_cia = cia
                nom_cia = den[den['cod_cia'] == cod_cia]['des_cia'].iloc[0]
                corta_cia = den[den['cod_cia'] == cod_cia]['denominacion_corta'].iloc[0]
                nombre_auditor = auditores[auditores['Cod'] == cod_cia]['AUDITOR'].iloc[0]
                numero_registro = int(auditores[auditores['Cod'] == cod_cia]['MATRICULA'].iloc[0])

                df_cia_res = df_resumen[df_resumen['Cod. Cía'] == cia]

                auto_flag = df_cia_res['Diferencia a reservar por EP 1 casos subvaluados AUTOS'].iloc[0] > 0
                moto_flag = df_cia_res['Diferencia a reservar por EP 1 casos subvaluados MOTOS'].iloc[0] > 0
                rc_flag = df_cia_res['Diferencia a reservar por EP 1 casos subvaluados RC'].iloc[0] > 0
                resto_ramas_flag = df_cia_res['Diferencia a reservar por EP 1 casos subvaluados Resto de ramas'].iloc[0] > 0
                indeterminados_flag = df_cia_res['Diferencia a reservar por casos indeterminados subvaluados'].iloc[0] > 0
                mediaciones_flag = df_cia_res['Diferencia a reservar mediaciones subvaluadas'].iloc[0] > 0

                def get_vals(df, col_casos, col_reserva, col_mm, col_diff):
                    row = df[df['Cod. Cía'] == cod_cia].iloc[0]
                    return (
                        format_mil(row[col_casos]),
                        format_mil(row[col_reserva]),
                        format_mil(row[col_mm]),
                        format_mil(row[col_diff])
                    )

                if auto_flag:
                    auto_casos, auto_reserva, auto_casos_mm, auto_diff = get_vals(
                        cuadros_auto,
                        'Casos Estado Procesal 1',
                        'Reserva Automotor \n(1)',
                        'Casos reservados por debajo del mínimo\n(4)',
                        'Diferencia\n(9)=(7)-(6)'
                    )

                if moto_flag:
                    moto_casos, moto_reserva, moto_casos_mm, moto_diff = get_vals(
                        cuadros_moto,
                        'Casos Estado Procesal 1',
                        'Reserva Motovehículos \n(1)',
                        'Casos reservados por debajo del mínimo\n(4)',
                        'Diferencia\n(9)=(7)-(6)'
                    )

                if rc_flag:
                    rc_casos, rc_reserva, rc_casos_mm, rc_diff = get_vals(
                        cuadros_rc,
                        'Casos Estado Procesal 1',
                        'Reserva RC \n(1)',
                        'Casos reservados por debajo del mínimo\n(4)',
                        'Diferencia\n(9)=(7)-(6)'
                    )

                if resto_ramas_flag:
                    rr_casos, rr_reserva, rr_casos_mm, rr_diff = get_vals(
                        cuadros_resto_ramas,
                        'Casos Estado Procesal 1',
                        'Reserva Resto de ramas \n(1)',
                        'Casos reservados por debajo del mínimo\n(4)',
                        'Diferencia\n(9)=(7)-(6)'
                    )

                if indeterminados_flag:
                    indeterminados_casos, indeterminados_reserva, indeterminados_casos_mm, indeterminados_diff = get_vals(
                        cuadros_indet,
                        'Casos Demanda Indeterminada',
                        'Reserva Demandas Indeterminadas',
                        'Casos reservados por debajo del mínimo',
                        'Monto Subvaluado'
                    )

                if mediaciones_flag:
                    mediaciones_casos, mediaciones_reserva, mediaciones_casos_mm, mediaciones_diff = get_vals(
                        cuadros_mediaciones,
                        'Casos Estado Procesal 1',
                        'Reserva Mediaciones',
                        'Casos reservados por debajo del mínimo',
                        'Monto Subvaluado'
                    )

                context = {
                    "nom_cia": nom_cia,
                    "cod_cia": cod_cia,
                    "periodo": per,
                    "fecha_per": fecha_per,
                    "auto_flag": auto_flag,
                    "auto_casos": auto_casos if auto_flag else 0,
                    "auto_reserva": auto_reserva if auto_flag else 0,
                    "auto_casos_mm": auto_casos_mm if auto_flag else 0,
                    "auto_diff": auto_diff if auto_flag else 0,
                    "moto_flag": moto_flag,
                    "moto_casos": moto_casos if moto_flag else 0,
                    "moto_reserva": moto_reserva if moto_flag else 0,
                    "moto_casos_mm": moto_casos_mm if moto_flag else 0,
                    "moto_diff": moto_diff if moto_flag else 0,
                    "rc_flag": rc_flag,
                    "rc_casos": rc_casos if rc_flag else 0,
                    "rc_reserva": rc_reserva if rc_flag else 0,
                    "rc_casos_mm": rc_casos_mm if rc_flag else 0,
                    "rc_diff": rc_diff if rc_flag else 0,
                    "resto_ramas_flag": resto_ramas_flag,
                    "rr_casos": rr_casos if resto_ramas_flag else 0,
                    "rr_reserva": rr_reserva if resto_ramas_flag else 0,
                    "rr_casos_mm": rr_casos_mm if resto_ramas_flag else 0,
                    "rr_diff": rr_diff if resto_ramas_flag else 0,
                    "indeterminados_flag": indeterminados_flag,
                    "indeterminados_casos": indeterminados_casos if indeterminados_flag else 0,
                    "indeterminados_reserva": indeterminados_reserva if indeterminados_flag else 0,
                    "indeterminados_casos_mm": indeterminados_casos_mm if indeterminados_flag else 0,
                    "indeterminados_diff": indeterminados_diff if indeterminados_flag else 0,
                    "mediaciones_flag": mediaciones_flag,
                    "mediaciones_casos": mediaciones_casos if mediaciones_flag else 0,
                    "mediaciones_reserva": mediaciones_reserva if mediaciones_flag else 0,
                    "mediaciones_casos_mm": mediaciones_casos_mm if mediaciones_flag else 0,
                    "mediaciones_diff": mediaciones_diff if mediaciones_flag else 0,
                    "nombre_auditor": nombre_auditor,
                    "numero_registro": numero_registro,
                    "dias_letra": dias_letra,
                    "dias_num": dias_num
                }

                for plant, name in [
                    (os.path.join(BASE_DIR, 'templates', 'modelo_informe_GE.docx'), f"(Cod {cod_cia}) {nom_cia}_RSP juicios.docx"),
                    (os.path.join(BASE_DIR, 'templates', 'modelo_informe_auditor.docx'), f"(Cod {cod_cia}) {nom_cia}_RSP juicios_{nombre_auditor}.docx")
                ]:
                    doc = DocxTemplate(plant)
                    doc.render(context)
                    buf = BytesIO()
                    doc.save(buf)
                    zf.writestr(name, buf.getvalue())

            except Exception as e:
                st.warning(f"Error en compañía {cia}: {e}")
                continue

    zip_buffer.seek(0)
    return zip_buffer

# ================= UI =================

st.set_page_config(page_title="RSP - SSN", page_icon="⚖️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #003366; }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { color: #ffffff !important; }
[data-testid="stSidebar"] input { color: #000000 !important; }
.block-container { padding-top: 2rem; }
h1 { text-transform: uppercase; text-align: center; font-size: 2.2rem; font-weight: 800; color: #003366; }
.stButton>button { border-radius: 8px; font-weight: bold; }
.stDownloadButton>button { background-color: #28a745 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("<h2 style='text-align: center; color: white;'>⚙️ CONFIGURACIÓN</h2>", unsafe_allow_html=True)
    st.markdown("<hr style='border: 0.5px solid white;'>", unsafe_allow_html=True)
    periodo = st.text_input("📅 PERIODO", value="2025-4")
    parametro = st.number_input("📊 UMBRAL %", value=0.39, format="%.2f")
    dias = st.number_input("⏳ DÍAS PLAZO", value=10)

st.markdown("<h1>Generador de Providencias RSP</h1>", unsafe_allow_html=True)
st.markdown("---")

col_izq, col_cen, col_der = st.columns([1, 2, 1])

with col_cen:
    st.subheader("📂 CARGA DE DATOS")
    archivo_principal = st.file_uploader("Subir archivo Excel", type=["xlsx"], label_visibility="visible")
    
    if st.button("🚀 INICIAR PROCESAMIENTO", use_container_width=True):
        if archivo_principal:
            with st.spinner("Procesando..."):
                res = procesar(archivo_principal, periodo, parametro, dias)
                if res:
                    st.success("✅ ¡PROCESO FINALIZADO!")
                    st.download_button(
                        label="📥 DESCARGAR RESULTADOS (ZIP)", 
                        data=res, 
                        file_name=f"PROVIDENCIAS_RSP_{periodo}.zip", 
                        use_container_width=True
                    )
        else:
            st.error("⚠️ DEBE CARGAR EL ARCHIVO EXCEL.")

st.markdown("<div style='text-align: center; color: grey; font-size: 0.7em; margin-top: 50px;'>SSN 2026</div>", unsafe_allow_html=True)