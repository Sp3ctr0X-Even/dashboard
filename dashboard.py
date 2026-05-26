import streamlit as st
import pandas as pd
import sqlite3
import io
#import win32com.client 
import pythoncom # NUEVO: Librería obligatoria para que Streamlit hable con SAP

st.set_page_config(page_title="Dashboard de Validaciones", layout="wide", initial_sidebar_state="collapsed")
st.title("📊 Panel Central de Validaciones y Robot SAP")
st.markdown("---")

ruta_exacta = 'D:/PROY PERCEP/transformar/base_de_datos_maestra.db'

# --- 1. OBTENER CANALES ---
@st.cache_data
def obtener_canales():
    try:
        conexion = sqlite3.connect(ruta_exacta, timeout=20)
        query = "SELECT DISTINCT CDis FROM tabla_unificada WHERE CDis IS NOT NULL ORDER BY CDis"
        df_canales = pd.read_sql(query, conexion)
        conexion.close()
        return ["Todos"] + df_canales['CDis'].astype(str).tolist()
    except Exception:
        return ["Todos"]

canales_disponibles = obtener_canales()

# --- 2. CONTROLES DE BÚSQUEDA ---
st.subheader("🔍 Controles de Búsqueda")
col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    filtro_estado = st.selectbox("Filtrar por Estado:", ("Todos", "correcto", "incorrecto"))
with col2:
    filtro_canal = st.selectbox("Filtrar por Canal (CDis):", canales_disponibles)
with col3:
    limite_filas = st.selectbox("Cantidad de registros a mostrar:", (5000, 10000, 50000))

st.markdown("---")

# --- 3. CARGAR DATOS ---
@st.cache_data
def cargar_datos(estado, canal, limite):
    conexion = sqlite3.connect(ruta_exacta, timeout=20)
    query = '''
    SELECT 
        t1.ID_Concatenado, t1.OrgVt, t1.CDis, 
        CASE WHEN t1.Se IN ('1', 1) THEN '01' ELSE t1.Se END AS Se, 
        t1.Deudor, t1.[Nombre 1], t1.TpImp, t1.ClfFi, 
        v.[Tipo Impuesto], v.Clasificación,
        CASE WHEN t1.ClfFi = v.Clasificación THEN 'correcto' ELSE 'incorrecto' END AS Estado_Validacion
    FROM tabla_unificada t1
    LEFT JOIN Validar v ON t1.ID_Concatenado = v.Concat
    WHERE 1=1
    '''
    if estado != "Todos":
        query += f" AND Estado_Validacion = '{estado}'"
    if canal != "Todos":
        query += f" AND t1.CDis = '{canal}'"
        
    query += f" LIMIT {limite}"
    df = pd.read_sql(query, conexion)
    conexion.close()
    return df

datos = cargar_datos(filtro_estado, filtro_canal, limite_filas)

# --- 4. VISUALIZACIÓN INTERACTIVA CON MÉTRICAS ---
if not datos.empty:
    st.subheader("📋 Selecciona los registros a gestionar")
    espacio_metricas = st.container()

    configuracion_columnas = {"ID_Concatenado": None}
    evento = st.dataframe(
        datos, 
        height=400, 
        use_container_width="stretch", 
        column_config=configuracion_columnas,
        selection_mode="multi-row",
        on_select="rerun"
    )

    filas_seleccionadas = evento.selection.rows
    cantidad_seleccionada = len(filas_seleccionadas)

    with espacio_metricas:
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        with m_col1:
            st.metric(label="Registros filtrados", value=len(datos))
        with m_col2:
            st.metric(label="📌 Seleccionados", value=cantidad_seleccionada, delta="Listos para procesar" if cantidad_seleccionada > 0 else None)
        with m_col3:
            st.metric(label="Estado activo", value=filtro_estado.capitalize())
        with m_col4:
            st.metric(label="Canal activo", value=filtro_canal)

    # --- 5. PANEL DE ACCIONES (DB Y SAP) ---
    if cantidad_seleccionada > 0:
        datos_seleccionados = datos.iloc[filas_seleccionadas]
        st.markdown("### ⚙️ Acciones a ejecutar")
        
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            if st.button(f"💾 Actualizar {cantidad_seleccionada} registros en BD Local", use_container_width=True):
                with st.spinner('Actualizando base de datos local...'):
                    try:
                        ids_a_actualizar = datos_seleccionados['ID_Concatenado'].tolist()
                        conexion = sqlite3.connect(ruta_exacta, timeout=30)
                        cursor = conexion.cursor()
                        format_strings = ','.join(['?'] * len(ids_a_actualizar))
                        query_update = f'''
                        UPDATE tabla_unificada SET ClfFi = (
                            SELECT Clasificación FROM Validar WHERE Concat = tabla_unificada.ID_Concatenado
                        ) WHERE ID_Concatenado IN ({format_strings})
                        '''
                        cursor.execute(query_update, ids_a_actualizar)
                        conexion.commit()
                        conexion.close()
                        st.cache_data.clear()
                        st.success("✅ BD Local actualizada.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

        with btn_col2:
            if st.button(f"🤖 Ejecutar Robot SAP para {cantidad_seleccionada} registros", type="primary", use_container_width=True):
                st.info("⚠️ Robot en ejecución: Por favor, **no toques tu teclado ni tu mouse** hasta que termine.")
                
                try:
                    # NUEVO: Inicializamos el hilo para evitar el error de Sintaxis
                    pythoncom.CoInitialize()
                    
                    # NUEVO: Conexión idéntica a tu VBScript asegurada para Python
                    SapGuiAuto = win32com.client.GetObject("SAPGUI")
                    application = SapGuiAuto.GetScriptingEngine
                    connection = application.Children(0)
                    session = connection.Children(0)
                    
                    barra_progreso = st.progress(0)
                    texto_progreso = st.empty()
                    errores = 0
                    total = len(datos_seleccionados)
                    
                    for i, (index, row) in enumerate(datos_seleccionados.iterrows()):
                        deudor = str(row['Deudor']).split('.')[0] 
                        orgvt = str(row['OrgVt']).split('.')[0]
                        cdis = str(row['CDis']).split('.')[0]
                        se = str(row['Se']).split('.')[0]
                        tpimp = str(row['TpImp']).strip()
                        clasificacion = str(row['Clasificación']).split('.')[0]
                        
                        texto_progreso.text(f"Procesando Cliente {deudor} ({i+1} de {total})...")
                        
                        try:
                            session.findById("wnd[0]").resizeWorkingPane(86, 28, False)
                            session.findById("wnd[0]/tbar[0]/okcd").text = "/NXD02"
                            session.findById("wnd[0]").sendVKey(0) 
                            
                            session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = deudor
                            session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = orgvt
                            session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = cdis
                            session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = se
                            session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus()
                            session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
                            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            
                            session.findById("wnd[0]/tbar[1]/btn[27]").press()
                            session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select()
                            
                            if tpimp == "MWST":
                                id_posicion = "[4,0]"
                            elif tpimp == "Z1AP":
                                id_posicion = "[4,1]"
                            elif tpimp == "ZISC":
                                id_posicion = "[4,2]"
                            else:
                                id_posicion = "[4,0]" 
                                
                            ruta_campo = f"wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA4:SAPMF02D:7323/subSUB_STEUER:SAPMF02D:7350/tblSAPMF02DTCTRL_STEUERN/ctxtKNVI-TAXKD{id_posicion}"
                            
                            campo_impuesto = session.findById(ruta_campo)
                            campo_impuesto.text = clasificacion
                            campo_impuesto.setFocus()
                            campo_impuesto.caretPosition = 1
                            session.findById("wnd[0]").sendVKey(0) 
                            
                            session.findById("wnd[0]/tbar[1]/btn[25]").press()
                            session.findById("wnd[0]/tbar[0]/btn[11]").press()
                            session.findById("wnd[0]").sendVKey(0)
                            
                        except Exception as error_fila:
                            errores += 1
                            st.toast(f"⚠️ Error al procesar deudor {deudor}: puede estar bloqueado.", icon="⚠️")
                            
                        barra_progreso.progress((i + 1) / total)
                        
                    texto_progreso.empty() 
                    
                    if errores == 0:
                        st.success(f"✅ ¡Robot finalizado! Se actualizaron {total} clientes exitosamente en SAP.")
                    else:
                        st.warning(f"⚠️ El robot terminó su recorrido, pero hubo {errores} clientes que no se pudieron actualizar.")
                        
                except Exception as e:
                    st.error(f"❌ Error de conexión con SAP: {e}")
                finally:
                    # NUEVO: Limpiamos la conexión COM al terminar para no saturar la memoria
                    pythoncom.CoUninitialize()

    # --- 6. DESCARGA EXCEL ---
    st.markdown("---")
    def convertir_excel(df):
        output = io.BytesIO()
        df_export = df.drop(columns=['ID_Concatenado'])
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Reporte')
        return output.getvalue()

    st.download_button(
        label="📥 Descargar vista en Excel",
        data=convertir_excel(datos),
        file_name=f'Reporte_Validaciones.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
