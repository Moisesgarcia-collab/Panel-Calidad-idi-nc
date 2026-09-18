import streamlit as st
import pandas as pd
import plotly.express as px
import io
import os
import tempfile
import datetime
from fpdf import FPDF
import gspread 
from oauth2client.service_account import ServiceAccountCredentials
import json

# --- 1. ANCLAJE ABSOLUTO ---
DIRECTORIO_ACTUAL = os.path.dirname(os.path.abspath(__file__))
RUTA_CREDENCIALES = os.path.join(DIRECTORIO_ACTUAL, "credenciales.json")

# 2. Configuración general de la página
st.set_page_config(page_title="Panel de Control de Calidad - Oil & Gas", layout="wide")

# Escudo Anti-traducción y CSS Futurista
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)
st.markdown("""
<style>
    div[data-testid="metric-container"] {
        background-color: #1E1E2E; 
        border: 1px solid #4A90E2; 
        padding: 5% 5% 5% 10%;
        border-radius: 10px; 
        box-shadow: 0px 4px 15px rgba(74, 144, 226, 0.3); 
        transition: transform 0.2s; 
    }
    div[data-testid="metric-container"]:hover {
        transform: scale(1.02);
        box-shadow: 0px 6px 20px rgba(74, 144, 226, 0.6);
    }
</style>
""", unsafe_allow_html=True)

st.title("Panel de seguimiento de Calidad ☁️")
st.markdown("Sistema conectado a la Nube (Google Sheets)")

# --- INICIALIZAR MEMORIA DE SESIÓN (SEGURIDAD) ---
if 'es_admin' not in st.session_state:
    st.session_state['es_admin'] = False

# 3. CONEXIÓN A GOOGLE SHEETS
@st.cache_resource
def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "GOOGLE_JSON" in st.secrets:
            datos_credenciales = json.loads(st.secrets["GOOGLE_JSON"])
            with open("credenciales_temp.json", "w") as f:
                json.dump(datos_credenciales, f)
            creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales_temp.json", scope)
        else:
            creds = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
    except Exception:
        creds = ServiceAccountCredentials.from_json_keyfile_name(RUTA_CREDENCIALES, scope)
        
    cliente = gspread.authorize(creds)
    return cliente.open("Base_HSEQ")

db = conectar_google_sheets()
hoja_idi = db.worksheet("IDI")
hoja_ncs = db.worksheet("NCs")

# 4. Funciones de extracción de datos
@st.cache_data(ttl=60)
def cargar_datos_idi():
    datos = hoja_idi.get_all_records()
    df = pd.DataFrame(datos)
    if not df.empty and 'Fecha' in df.columns:
        df['Filtro_Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.date
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').fillna("Sin fecha")
        columnas_texto = [c for c in df.columns if c not in ['Fecha', 'Filtro_Fecha']]
        df[columnas_texto] = df[columnas_texto].fillna("N/A").astype(str)
    return df

@st.cache_data(ttl=60)
def cargar_datos_ncs():
    datos = hoja_ncs.get_all_records()
    df = pd.DataFrame(datos)
    if not df.empty and 'Fecha' in df.columns:
        df['Filtro_Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.date
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').fillna("Sin fecha")
        columnas_texto = [c for c in df.columns if c not in ['Fecha', 'Filtro_Fecha']]
        df[columnas_texto] = df[columnas_texto].fillna("N/A").astype(str)
    return df

with st.sidebar:
    st.write("### 🔐 Acceso de Administrador")
    clave_ingresada = st.text_input("Contraseña para Editar", type="password")
    
    if clave_ingresada == "HSEQ2026":
        st.session_state['es_admin'] = True
        st.success("Modo Edición Activado")
    elif clave_ingresada != "":
        st.session_state['es_admin'] = False
        st.error("Contraseña incorrecta")
    else:
        st.session_state['es_admin'] = False
        st.info("Modo Solo Lectura")

    st.divider()
    st.write("### Control de Nube ☁️")
    if st.button("🔄 Sincronizar Datos Ahora"):
        st.cache_data.clear() 
        st.rerun() 

df_inspecciones = cargar_datos_idi()
df_ncs = cargar_datos_ncs()

# 5. Funciones de Exportación (Excel y PDF)
def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

def generar_pdf_reporte(df, titulo_reporte, tipo_reporte, figuras=None):
    pdf = FPDF(orientation='L') 
    
    if figuras and len(figuras) > 0:
        for index, fig in enumerate(figuras):
            pdf.add_page()
            if index == 0:
                pdf.set_font('Arial', 'B', 16)
                pdf.cell(0, 10, titulo_reporte + " - Resumen Visual", ln=True, align='C')
                pdf.ln(5)
                
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                fig.write_image(tmpfile.name, format="png", width=900, height=450)
                pdf.image(tmpfile.name, x=15, w=260)
            os.remove(tmpfile.name)
            
    pdf.add_page()
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 10, titulo_reporte + " - Datos Detallados", ln=True, align='C')
    pdf.ln(5)

    if tipo_reporte == "IDI":
        columnas_pdf = ['Fecha', 'Disciplina', 'Contratista', 'Estado']
    else:
        columnas_pdf = ['ID', 'Criticidad', 'Responsable', 'Estado', 'Dias_Abiertas']
        
    cols_existentes = [col for col in columnas_pdf if col in df.columns]
    ancho_col = 270 / max(len(cols_existentes), 1)

    pdf.set_font('Arial', 'B', 10)
    for col in cols_existentes:
        pdf.cell(ancho_col, 10, str(col), border=1, align='C')
    pdf.ln()

    pdf.set_font('Arial', '', 8)
    for _, row in df.head(100).iterrows():
        for col in cols_existentes:
            valor = row[col]
            texto = str(valor) if pd.notna(valor) else "N/A"
            texto = texto.encode('latin-1', 'replace').decode('latin-1')
            pdf.cell(ancho_col, 8, texto[:35], border=1)
        pdf.ln()

    return bytes(pdf.output())

# 6. Creación de pestañas
tab1, tab2 = st.tabs(["📋 Informes de Inspección (IDI)", "⚠️ Productos No Conformes (NCs)"])

# --- PESTAÑA 1: INSPECCIONES (Tablero IDI) ---
with tab1:
    if st.session_state['es_admin']:
        with st.expander("➕ Cargar nuevo Informe de Inspección (IDI)"):
            with st.form("form_nueva_idi", clear_on_submit=True):
                st.write("**Complete todos los campos de la nueva inspección:**")
                f1_c1, f1_c2, f1_c3, f1_c4 = st.columns(4)
                input_fecha = f1_c1.date_input("Fecha", format="DD/MM/YYYY")
                input_confecciono = f1_c2.text_input("Confeccionó")
                input_inf_n2 = f1_c3.text_input("Informe N°2")
                input_proyecto = f1_c4.selectbox("Proyecto", ["WELLPADS", "LMT", "TRONCALES", "TRAMPAS", "DUCTOS", "OTROS"])
                
                f2_c1, f2_c2, f2_c3, f2_c4 = st.columns(4)
                input_contra = f2_c1.selectbox("Contratista", ["OPS", "EDVSA"])
                input_disc = f2_c2.selectbox("Disciplina", ["MECANICO", "PIPING", "CIVIL", "E&I", "SOLDADURA", "CALIDAD"])
                input_aspecto = f2_c3.selectbox("Aspecto", ["NEGATIVO", "POSITIVO"])
                input_crit = f2_c4.selectbox("Criticidad", ["BAJO", "MODERADO", "ALTO", "CRÍTICO"])
                
                f3_c1, f3_c2 = st.columns(2)
                input_desc = f3_c1.text_input("Descripción (breve)")
                input_informe_de = f3_c2.selectbox("Informe de:", ["DESVÍO", "AVANCE", "INSPECCIÓN", "OTROS"])
                
                f4_c1, f4_c2, f4_c3, f4_c4 = st.columns(4)
                input_tipo = f4_c1.selectbox("Tipo", ["PRODUCTO", "DOCUMENTO", "PROCESO"])
                input_estado = f4_c2.selectbox("Estado", ["ABIERTO", "CERRADO"])
                input_zonas = f4_c3.text_input("Zonas")
                input_uso = f4_c4.selectbox("Uso", ["CONTRATISTA", "INTERNO"])
                
                input_link = st.text_input("Enlace del Informe (URL de SharePoint)")
                
                submit_btn = st.form_submit_button("☁️ Guardar en Google Sheets")
                
                if submit_btn:
                    try:
                        fecha_str = input_fecha.strftime('%d/%m/%Y')
                        nueva_fila = [
                            input_confecciono, input_inf_n2, fecha_str, input_proyecto, 
                            input_contra, input_disc, input_aspecto, input_crit, 
                            input_desc, input_informe_de, input_tipo, input_estado, input_zonas, 
                            "", input_uso, input_link, "", "N/A"
                        ]
                        hoja_idi.append_row(nueva_fila)
                        st.cache_data.clear() 
                        st.toast("✅ ¡Inspección guardada exitosamente en la nube!", icon="☁️")
                    except Exception as e:
                        st.error(f"Ocurrió un error al enviar a la nube: {e}")

        with st.expander("🔒 Cerrar Inspección (IDI)"):
            if 'Estado' in df_inspecciones.columns and 'Informe N°2' in df_inspecciones.columns:
                abiertos_idi = df_inspecciones[df_inspecciones['Estado'].astype(str).str.upper().str.contains("ABIERTO", na=False)]
                lista_informes_abiertos = abiertos_idi['Informe N°2'].dropna().unique().tolist()
                lista_informes_abiertos = [inf for inf in lista_informes_abiertos if str(inf).strip() != ""]
                
                if len(lista_informes_abiertos) > 0:
                    with st.form("form_cerrar_idi", clear_on_submit=True):
                        st.write("**Seleccione el Informe que desea marcar como CERRADO:**")
                        idi_a_cerrar = st.selectbox("Informe N°2 a cerrar", lista_informes_abiertos)
                        btn_cerrar_idi = st.form_submit_button("🔒 Confirmar Cierre en la Nube")
                        
                        if btn_cerrar_idi:
                            try:
                                celda_encontrada = hoja_idi.find(idi_a_cerrar)
                                if celda_encontrada:
                                    hoja_idi.update_cell(celda_encontrada.row, 12, "CERRADO")
                                    st.cache_data.clear()
                                    st.toast(f"✅ ¡El informe {idi_a_cerrar} ha sido actualizado a CERRADO!", icon="🔒")
                                else:
                                    st.error("No se encontró el informe en la nube.")
                            except Exception as e:
                                st.error(f"Error al actualizar la base de datos: {e}")
                else:
                    st.info("🎉 ¡Excelente trabajo! No hay informes ABIERTOS en este momento.")

    st.divider() 

    col_filtros1, col_filtros2, col_filtros3, col_filtros4 = st.columns(4)
    
    if 'Disciplina' in df_inspecciones.columns:
        disciplinas = ["Todas"] + list(df_inspecciones['Disciplina'].astype(str).unique())
        filtro_disc = col_filtros1.selectbox("Disciplina:", disciplinas)
    else:
        filtro_disc = "Todas"
        
    filtro_contra_idi = col_filtros2.selectbox("Contratista (IDI):", ["Todas", "EDVSA", "OPS"])

    if 'Filtro_Fecha' in df_inspecciones.columns:
        fechas_validas_idi = df_inspecciones['Filtro_Fecha'].dropna()
        if not fechas_validas_idi.empty:
            fecha_min_idi = fechas_validas_idi.min()
            fecha_max_idi = fechas_validas_idi.max()
        else:
            fecha_min_idi = datetime.date(2026, 1, 1)
            fecha_max_idi = datetime.date(2026, 12, 31)
            
        fecha_inicio_idi = col_filtros3.date_input("Desde (IDI):", value=fecha_min_idi, format="DD/MM/YYYY")
        fecha_fin_idi = col_filtros4.date_input("Hasta (IDI):", value=fecha_max_idi, format="DD/MM/YYYY")
    else:
        fecha_inicio_idi = None
        fecha_fin_idi = None

    df_idi_filtrado = df_inspecciones.copy()
    
    if filtro_disc != "Todas" and 'Disciplina' in df_idi_filtrado.columns:
        df_idi_filtrado = df_idi_filtrado[df_idi_filtrado['Disciplina'] == filtro_disc]
        
    if filtro_contra_idi != "Todas" and 'Contratista' in df_idi_filtrado.columns:
        if filtro_contra_idi == "OPS":
            df_idi_filtrado = df_idi_filtrado[df_idi_filtrado['Contratista'].astype(str).str.upper().str.contains("OPS|OPERACION", na=False)]
        elif filtro_contra_idi == "EDVSA":
            df_idi_filtrado = df_idi_filtrado[df_idi_filtrado['Contratista'].astype(str).str.upper().str.contains("EDVSA", na=False)]

    if fecha_inicio_idi and fecha_fin_idi and 'Filtro_Fecha' in df_idi_filtrado.columns:
        df_idi_filtrado = df_idi_filtrado[
            df_idi_filtrado['Filtro_Fecha'].apply(
                lambda x: True if pd.isna(x) else (fecha_inicio_idi <= x <= fecha_fin_idi)
            )
        ]

    columna_estado = 'Estado'
    figuras_idi_export = []
    
    if columna_estado in df_idi_filtrado.columns:
        total_idi = len(df_idi_filtrado)
        abiertos_idi = len(df_idi_filtrado[df_idi_filtrado[columna_estado].astype(str).str.upper().str.contains("ABIERTO", na=False)])
        cerrados_idi = len(df_idi_filtrado[df_idi_filtrado[columna_estado].astype(str).str.upper().str.contains("CERRADO", na=False)])
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Pendientes (Abiertos)", abiertos_idi, "- Tareas por resolver", delta_color="inverse")
        m2.metric("Completados (Cerrados)", cerrados_idi, "+ Tareas resueltas")
        m3.metric("Total Informes", total_idi)
        
        if total_idi > 0:
            progreso_idi = int((cerrados_idi / total_idi) * 100)
            st.progress(progreso_idi, text=f"Productividad de Cierre de Inspecciones: {progreso_idi}%")
            
        st.divider()

        # ==========================================
        # GRÁFICOS GERENCIALES IDI
        # ==========================================
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            if total_idi > 0:
                mapa_colores_idi = {"ABIERTO": "#FF4B4B", "CERRADO": "#28A745"}
                fig_dona_idi = px.pie(
                    df_idi_filtrado, names=columna_estado, hole=0.6, 
                    title="ESTADO DE INSPECCIONES (IDI)",
                    color=columna_estado,
                    color_discrete_map=mapa_colores_idi
                )
                
                fig_dona_idi.update_traces(
                    textinfo='percent+label', 
                    textfont_size=14,
                    textfont_color='white',
                    marker=dict(line=dict(color='#000000', width=1))
                )
                fig_dona_idi.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                    showlegend=False, 
                    title_font=dict(size=18, color='#FFFFFF', family="Arial")
                )
                st.plotly_chart(fig_dona_idi, use_container_width=True)
                figuras_idi_export.append(fig_dona_idi)
            
        with col_g2:
            if 'Disciplina' in df_idi_filtrado.columns and total_idi > 0:
                conteo_disc = df_idi_filtrado['Disciplina'].value_counts().reset_index()
                conteo_disc.columns = ['Disciplina', 'Cantidad']
                fig_bar_idi = px.bar(
                    conteo_disc, x='Cantidad', y='Disciplina', orientation='h', 
                    title="INFORMES POR DISCIPLINA", text='Cantidad'
                )
                
                fig_bar_idi.update_traces(
                    marker_color='#008FFB', 
                    textposition='inside',  
                    textfont_color='white',
                    textfont_size=14
                )
                fig_bar_idi.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(showgrid=False, visible=False), 
                    yaxis=dict(showgrid=False, title=None, tickfont=dict(color='white')),
                    title_font=dict(size=18, color='#FFFFFF', family="Arial"),
                    margin=dict(l=0, r=0, t=40, b=0)
                )
                st.plotly_chart(fig_bar_idi, use_container_width=True)
                figuras_idi_export.append(fig_bar_idi)

    st.write("**Opciones de Exportación de Informes IDI**")
    col_btn1, col_btn2 = st.columns(2)
    
    df_idi_mostrar = df_idi_filtrado.drop(columns=['Filtro_Fecha']) if 'Filtro_Fecha' in df_idi_filtrado.columns else df_idi_filtrado
    excel_idi = convertir_df_a_excel(df_idi_mostrar)
    col_btn1.download_button(label="📥 Descargar a Excel", data=excel_idi, file_name="Reporte_IDI.xlsx")
    
    with col_btn2:
        if st.button("⚙️ Preparar PDF con Gráficos (IDI)"):
            with st.spinner("Procesando imágenes y construyendo PDF..."):
                pdf_idi = generar_pdf_reporte(df_idi_mostrar, "Reporte de Inspecciones IDI", "IDI", figuras_idi_export)
                st.download_button(label="📄 Haz clic aquí para descargar tu PDF", data=pdf_idi, file_name="Reporte_IDI.pdf", mime="application/pdf")
    
    st.write("### Base de datos IDI en tiempo real")
    st.dataframe(
        df_idi_mostrar, 
        use_container_width=True,
        column_config={
            "Informe": st.column_config.LinkColumn(
                "Documento de Respaldo", 
                display_text="Abrir PDF 🔗"
            )
        }
    )
# --- PESTAÑA 2: NO CONFORMIDADES (MasterNCs) ---
with tab2:
    if st.session_state['es_admin']:
        with st.expander("➕ Cargar nueva No Conformidad (NC)"):
            with st.form("form_nueva_nc", clear_on_submit=True):
                st.write("**Complete los datos del nuevo hallazgo / NC:**")
                
                f1_nc1, f1_nc2, f1_nc3, f1_nc4 = st.columns(4)
                nc_id = f1_nc1.text_input("ID de la NC")
                nc_fecha = f1_nc2.date_input("Fecha", format="DD/MM/YYYY")
                nc_contra = f1_nc3.selectbox("Contratista", ["OPS", "EDVSA"])
                nc_contrato = f1_nc4.text_input("Contrato")
                
                f2_nc1, f2_nc2, f2_nc3, f2_nc4 = st.columns(4)
                nc_origen = f2_nc1.selectbox("Origen", ["Inspección", "Auditoría", "Reclamo", "Otro"])
                nc_crit = f2_nc2.selectbox("Criticidad", ["Leve", "Menor", "Mayor", "Crítica"])
                nc_estado = f2_nc3.selectbox("Estado", ["Abierta", "En Gestión", "Cerrada"])
                nc_resp = f2_nc4.text_input("Responsable")
                
                nc_desc = st.text_input("Descripción del Desvío")
                nc_evidencia = st.text_input("Evidencia")
                
                f3_nc1, f3_nc2, f3_nc3 = st.columns(3)
                nc_causa = f3_nc1.text_input("Causa Raíz")
                nc_accion = f3_nc2.text_input("Acción Correctiva")
                nc_disposiciones = f3_nc3.text_input("Disposiciones a tomar")
                
                nc_fecha_limite = st.date_input("Fecha Límite", format="DD/MM/YYYY")
                
                submit_nc = st.form_submit_button("☁️ Guardar NC en Google Sheets")
                
                if submit_nc:
                    try:
                        fecha_nc_str = nc_fecha.strftime('%d/%m/%Y')
                        fecha_limite_str = nc_fecha_limite.strftime('%d/%m/%Y')
                        
                        nueva_fila_nc = [
                            nc_id, fecha_nc_str, nc_contra, nc_contrato, nc_origen, 
                            nc_crit, nc_desc, nc_evidencia, nc_causa, nc_accion, 
                            nc_disposiciones, nc_resp, fecha_limite_str, nc_estado, 0
                        ]
                        hoja_ncs.append_row(nueva_fila_nc)
                        st.cache_data.clear()
                        st.toast("✅ ¡No Conformidad guardada exitosamente en la nube!", icon="☁️")
                                
                    except Exception as e:
                        st.error(f"Ocurrió un error al enviar a la nube: {e}")

        with st.expander("🔒 Cerrar No Conformidad (NC)"):
            if 'Estado' in df_ncs.columns and 'ID' in df_ncs.columns:
                abiertas_nc = df_ncs[~df_ncs['Estado'].astype(str).str.upper().str.contains("CERRADA", na=False)]
                lista_ncs_abiertas = abiertas_nc['ID'].dropna().unique().tolist()
                lista_ncs_abiertas = [nc for nc in lista_ncs_abiertas if str(nc).strip() != ""]
                
                if len(lista_ncs_abiertas) > 0:
                    with st.form("form_cerrar_nc", clear_on_submit=True):
                        st.write("**Seleccione la NC que desea marcar como CERRADA:**")
                        nc_a_cerrar = st.selectbox("ID de NC a cerrar", lista_ncs_abiertas)
                        btn_cerrar_nc = st.form_submit_button("🔒 Confirmar Cierre en la Nube")
                        
                        if btn_cerrar_nc:
                            try:
                                celda_encontrada = hoja_ncs.find(nc_a_cerrar)
                                if celda_encontrada:
                                    hoja_ncs.update_cell(celda_encontrada.row, 14, "Cerrada")
                                    st.cache_data.clear()
                                    st.toast(f"✅ ¡La NC {nc_a_cerrar} ha sido actualizada a CERRADA!", icon="🔒")
                                else:
                                    st.error("No se encontró la NC en la nube.")
                            except Exception as e:
                                st.error(f"Error al actualizar la base de datos: {e}")
                else:
                    st.info("🎉 ¡Excelente trabajo! No hay No Conformidades pendientes de cierre.")

    st.divider()

    col_nc_f1, col_nc_f2, col_nc_f3 = st.columns(3)
    
    filtro_contra_nc = col_nc_f1.selectbox("Contratista (NCs):", ["Todas", "EDVSA", "OPS"])

    if 'Filtro_Fecha' in df_ncs.columns:
        fechas_validas_nc = df_ncs['Filtro_Fecha'].dropna()
        if not fechas_validas_nc.empty:
            fecha_min_nc = fechas_validas_nc.min()
            fecha_max_nc = fechas_validas_nc.max()
        else:
            fecha_min_nc = datetime.date(2026, 1, 1)
            fecha_max_nc = datetime.date(2026, 12, 31)
            
        fecha_inicio_nc = col_nc_f2.date_input("Desde (NCs):", value=fecha_min_nc, format="DD/MM/YYYY")
        fecha_fin_nc = col_nc_f3.date_input("Hasta (NCs):", value=fecha_max_nc, format="DD/MM/YYYY")
    else:
        fecha_inicio_nc = None
        fecha_fin_nc = None

    df_ncs_filtrado = df_ncs.copy()
    
    if filtro_contra_nc != "Todas" and 'Contratista' in df_ncs_filtrado.columns:
        if filtro_contra_nc == "OPS":
            df_ncs_filtrado = df_ncs_filtrado[df_ncs_filtrado['Contratista'].astype(str).str.upper().str.contains("OPS|OPERACION", na=False)]
        elif filtro_contra_nc == "EDVSA":
            df_ncs_filtrado = df_ncs_filtrado[df_ncs_filtrado['Contratista'].astype(str).str.upper().str.contains("EDVSA", na=False)]

    if fecha_inicio_nc and fecha_fin_nc and 'Filtro_Fecha' in df_ncs_filtrado.columns:
        df_ncs_filtrado = df_ncs_filtrado[
            df_ncs_filtrado['Filtro_Fecha'].apply(
                lambda x: True if pd.isna(x) else (fecha_inicio_nc <= x <= fecha_fin_nc)
            )
        ]

    figuras_nc_export = [] 

    if 'Estado' in df_ncs_filtrado.columns and 'Criticidad' in df_ncs_filtrado.columns:
        total_nc = len(df_ncs_filtrado)
        total_criticos = len(df_ncs_filtrado[df_ncs_filtrado['Criticidad'].astype(str).str.upper().str.contains("CRITIC|CRÍTICA", na=False)])
        total_cerradas = len(df_ncs_filtrado[df_ncs_filtrado['Estado'].astype(str).str.upper().str.contains("CERRAD", na=False)])
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total NC", total_nc)
        m2.metric("Total Críticos", total_criticos, "- Requiere atención prioritaria", delta_color="inverse")
        
        if 'Dias_Abiertas' in df_ncs_filtrado.columns:
            df_ncs_filtrado.loc[:, 'Dias_Abiertos_Num'] = pd.to_numeric(df_ncs_filtrado['Dias_Abiertas'], errors='coerce')
            promedio_dias = df_ncs_filtrado['Dias_Abiertos_Num'].mean()
            m3.metric("Promedio días de atraso", f"{promedio_dias:.0f}" if pd.notnull(promedio_dias) else "0")
        else:
            m3.metric("Promedio días de atraso", "0")
            
        m4.metric("Cerradas totales", total_cerradas, "+ Gestión exitosa")
        
        if total_nc > 0:
            progreso_nc = int((total_cerradas / total_nc) * 100)
            st.progress(progreso_nc, text=f"Productividad de Cierre de No Conformidades: {progreso_nc}%")
        
        st.divider()

        # ==========================================
        # GRÁFICOS GERENCIALES NC
        # ==========================================
        col_nc1, col_nc2, col_nc3 = st.columns(3)
        
        with col_nc1:
            if total_nc > 0:
                conteo_crit = df_ncs_filtrado['Criticidad'].value_counts().reset_index()
                conteo_crit.columns = ['Criticidad', 'Cantidad']
                
                mapa_colores = {"Leve": "#28A745", "Menor": "#008FFB", "Mayor": "#FEB019", "Crítica": "#FF4B4B", "Crítico": "#FF4B4B"}
                fig_riesgo = px.bar(
                    conteo_crit, x='Criticidad', y='Cantidad', 
                    title="PERFIL DE RIESGO", text='Cantidad',
                    color='Criticidad', color_discrete_map=mapa_colores
                )
                
                fig_riesgo.update_traces(textposition='inside', textfont_color='white', textfont_size=16)
                fig_riesgo.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    showlegend=False,
                    xaxis=dict(showgrid=False, title=None, tickfont=dict(color='white')),
                    yaxis=dict(showgrid=False, visible=False),
                    title_font=dict(size=18, color='#FFFFFF', family="Arial")
                )
                st.plotly_chart(fig_riesgo, use_container_width=True)
                figuras_nc_export.append(fig_riesgo)
            
        with col_nc2:
            if 'Contratista' in df_ncs_filtrado.columns and total_nc > 0:
                df_ncs_grafico = df_ncs_filtrado.copy()
                df_ncs_grafico.loc[df_ncs_grafico['Contratista'].astype(str).str.upper().str.contains("OPS|OPERACION", na=False), 'Contratista'] = 'OPS'
                
                col_limite_encontrada = None
                for col in df_ncs_grafico.columns:
                    if col.strip().lower() in ['fecha límite', 'fecha limite', 'fechalimite']:
                        col_limite_encontrada = col
                        break
                
                if col_limite_encontrada:
                    hoy = pd.to_datetime('today')
                    df_ncs_grafico['Fecha_Limite_DT'] = pd.to_datetime(df_ncs_grafico[col_limite_encontrada], format='%d/%m/%Y', errors='coerce')
                    df_ncs_grafico['Dias_Restantes'] = (df_ncs_grafico['Fecha_Limite_DT'] - hoy).dt.days
                    
                    def asignar_color_alerta(row):
                        if 'CERRAD' in str(row['Estado']).upper():
                            return 'Cerrada'
                        if pd.isna(row['Dias_Restantes']):
                            return 'Sin Fecha'
                        if row['Dias_Restantes'] < 0:
                            return 'Vencida (Rojo)'
                        elif 0 <= row['Dias_Restantes'] <= 7:
                            return 'Alerta (Amarillo)'
                        else:
                            return 'A tiempo (Verde)'
                            
                    df_ncs_grafico['Semaforo'] = df_ncs_grafico.apply(asignar_color_alerta, axis=1)
                    conteo_semaforo = df_ncs_grafico.groupby(['Contratista', 'Semaforo']).size().reset_index(name='Cantidad')
                    
                    # --- CORRECCIÓN DE COLOR PARA "CERRADA" EN MODO OSCURO ---
                    colores_alerta = {
                        'Vencida (Rojo)': '#FF4B4B',
                        'Alerta (Amarillo)': '#FEB019',
                        'A tiempo (Verde)': '#00CC96',
                        'Cerrada': '#9E9E9E', # Color Gris Plata Luminoso para destacar en el fondo oscuro
                        'Sin Fecha': '#808080'
                    }
                    
                    fig_atraso = px.bar(
                        conteo_semaforo, x='Cantidad', y='Contratista', color='Semaforo', orientation='h', 
                        title="ESTADO DE ALERTAS POR CONTRATISTA", text='Cantidad',
                        color_discrete_map=colores_alerta
                    )
                    
                    fig_atraso.update_traces(textposition='inside', textfont_color='white', textfont_size=14)
                    
                    fig_atraso.update_layout(
                        barmode='stack', 
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        showlegend=True,
                        legend_title_text=None,
                        legend=dict(
                            orientation="h", 
                            yanchor="top", 
                            y=-0.15, 
                            xanchor="center", 
                            x=0.5, 
                            font=dict(size=12, color="white")
                        ),
                        xaxis=dict(showgrid=False, visible=False),
                        yaxis=dict(showgrid=False, title=None, tickfont=dict(color='white')),
                        title_font=dict(size=16, color='#FFFFFF', family="Arial"),
                        margin=dict(l=0, r=0, t=50, b=50) 
                    )
                    st.plotly_chart(fig_atraso, use_container_width=True)
                    figuras_nc_export.append(fig_atraso)
                else:
                    atrasos = df_ncs_grafico.groupby('Contratista')['Dias_Abiertos_Num'].sum().reset_index()
                    fig_atraso = px.bar(
                        atrasos, x='Dias_Abiertos_Num', y='Contratista', orientation='h', 
                        title="RANKING DE ATRASOS (DÍAS)", text='Dias_Abiertos_Num'
                    )
                    fig_atraso.update_traces(marker_color='#FEB019', textposition='inside', textfont_color='black', textfont_size=14)
                    fig_atraso.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False,
                        xaxis=dict(showgrid=False, visible=False),
                        yaxis=dict(showgrid=False, title=None, tickfont=dict(color='white')),
                        title_font=dict(size=16, color='#FFFFFF', family="Arial"), margin=dict(l=0, r=0, t=40, b=0)
                    )
                    st.plotly_chart(fig_atraso, use_container_width=True)
                    figuras_nc_export.append(fig_atraso)
                
        with col_nc3:
            if total_nc > 0:
                mapa_colores_estatus = {
                    "Abierta": "#FF4B4B",     
                    "En Gestión": "#FEB019",  
                    "Cerrada": "#28A745"      
                }
                
                fig_estatus = px.pie(
                    df_ncs_filtrado, names='Estado', hole=0.6, 
                    title="ESTATUS OPERATIVO",
                    color='Estado',
                    color_discrete_map=mapa_colores_estatus
                )
                
                fig_estatus.update_traces(
                    textinfo='percent+label',
                    textfont_size=14,
                    textfont_color='white',
                    marker=dict(line=dict(color='#000000', width=1))
                )
                fig_estatus.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    showlegend=False,
                    title_font=dict(size=18, color='#FFFFFF', family="Arial")
                )
                st.plotly_chart(fig_estatus, use_container_width=True)
                figuras_nc_export.append(fig_estatus)

        if 'Filtro_Fecha' in df_ncs_filtrado.columns and total_nc > 0:
            df_ncs_filtrado['Mes'] = pd.to_datetime(df_ncs_filtrado['Filtro_Fecha']).dt.month_name()
            evolucion = df_ncs_filtrado.groupby('Mes').size().reset_index(name='Hallazgos')
            if not evolucion.empty:
                fig_linea = px.line(
                    evolucion, x='Mes', y='Hallazgos', 
                    title="EVOLUCIÓN MENSUAL DE HALLAZGOS", markers=True
                )
                
                fig_linea.update_traces(line_color='#008FFB', marker=dict(size=10, color='#FF4B4B'))
                fig_linea.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(showgrid=False, title=None, tickfont=dict(color='white')),
                    yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.1)', title=None, tickfont=dict(color='white')),
                    title_font=dict(size=18, color='#FFFFFF', family="Arial")
                )
                st.plotly_chart(fig_linea, use_container_width=True)
                figuras_nc_export.append(fig_linea)

        columnas_visibles = ['ID', 'Fecha', 'Criticidad', 'Descripcion_Desvio', 'Responsable', 'Estado', 'Dias_Abiertas']
        columnas_finales = [col for col in columnas_visibles if col in df_ncs_filtrado.columns]
        
        st.write("**Opciones de Exportación de No Conformidades**")
        col_btn_nc1, col_btn_nc2 = st.columns(2)
        
        df_ncs_mostrar = df_ncs_filtrado.drop(columns=['Filtro_Fecha']) if 'Filtro_Fecha' in df_ncs_filtrado.columns else df_ncs_filtrado
        
        excel_nc = convertir_df_a_excel(df_ncs_mostrar)
        col_btn_nc1.download_button(label="📥 Descargar a Excel", data=excel_nc, file_name="Reporte_NC.xlsx")
        
        with col_btn_nc2:
            if st.button("⚙️ Preparar PDF con Gráficos (NC)"):
                with st.spinner("Procesando imágenes y construyendo PDF..."):
                    pdf_nc = generar_pdf_reporte(df_ncs_mostrar, "Reporte de No Conformidades", "NC", figuras_nc_export)
                    st.download_button(label="📄 Haz clic aquí para descargar tu PDF", data=pdf_nc, file_name="Reporte_NC.pdf", mime="application/pdf")
        
        st.write("### Base de datos NCs en tiempo real")
        st.dataframe(
            df_ncs_mostrar[columnas_finales], 
            use_container_width=True,
            column_config={
                "Evidencia": st.column_config.LinkColumn(
                    "Evidencia Adjunta", 
                    display_text="Ver Archivo 🔗"
                )
            }
        )
        
        with st.expander("Ver base de datos completa de NCs"):
            st.dataframe(
                df_ncs_mostrar, 
                use_container_width=True,
                column_config={
                    "Evidencia": st.column_config.LinkColumn(
                        "Evidencia Adjunta", 
                        display_text="Ver Archivo 🔗"
                    )
                }
            )

# --- PIE DE PÁGINA CORPORATIVO ---
st.divider()
st.markdown("""
    <div style="text-align: center; color: gray; font-size: 12px;">
        Sistema de Gestión de Calidad | Desarrollado por Moises Garcia para Facilities Oil & Gas <br>
        Versión 1.0 - Producción y desarrollo continuo
    </div>
""", unsafe_allow_html=True)