import streamlit as st
import pandas as pd
from PIL import Image
import easyocr
import re
from io import BytesIO
from pathlib import Path
import os

# --- CONFIGURACION DE LA CARPETA AUTOMATICA EN LA UNIDAD E ---
CARPETA_DESTINO = Path("E:/SistemaElectoral")
CARPETA_FOTOS = CARPETA_DESTINO / "Fotos_Actas"
EXCEL_PATH = CARPETA_DESTINO / "computo_electoral_bolivia.xlsx"

# Configuracion del centro de computo nacional
st.set_page_config(page_title="Computo Nacional Automatizado", page_icon="🗳️", layout="wide")

# --- CONTRASENA SECRETA DE VISUALIZACION ---
CONTRASENA_DIRECTOR = "Bolivia2026*"

# --- PADRON OFICIAL DE MESAS Y RECINTOS ---
PADRON_MESAS = {
    "703579-1": {
        "departamento": "Santa Cruz", "provincia": "Andres Ibanez", "circunscripcion": "C-50", "recinto": "Col. Unidad Educativa La Colina"
    },
    "703580-2": {
        "departamento": "Santa Cruz", "provincia": "Ignacio Warnes", "circunscripcion": "C-49", "recinto": "U.E. Juan Pablo II"
    },
    "101234-1": {
        "departamento": "La Paz", "provincia": "Murillo", "circunscripcion": "C-6", "recinto": "Colegio San Calixto"
    }
}

# --- BASES DE DATOS EN MEMORIA CENTRAL ---
if "base_datos_actas" not in st.session_state:
    st.session_state["base_datos_actas"] = pd.DataFrame(columns=[
        "Mesa", "Departamento", "Provincia", "Circunscripcion", "Recinto",
        "LIBRE", "SPT",
        "Votos Validos (Frentes)", "Votos Blancos", "Votos Nulos", "Total Anfora", "Estado Acta"
    ])

if "base_datos_observadas" not in st.session_state:
    st.session_state["base_datos_observadas"] = pd.DataFrame(columns=[
        "Mesa", "Departamento", "Provincia", "Circunscripcion", "Recinto", "Motivo de Observacion Legal"
    ])

@st.cache_resource
def cargar_lector_ia():
    return easyocr.Reader(['es'], gpu=False)

lector_ia = cargar_lector_ia()

# --- MOTOR MATEMATICO JERARQUICO EN CASCADA ---
def procesar_tablas_computo():
    df_actas = st.session_state["base_datos_actas"]
    partidos = ["LIBRE", "SPT"]
    columnas_computo = partidos + ["Votos Validos (Frentes)", "Votos Blancos", "Votos Nulos", "Total Anfora"]
    
    if df_actas.empty:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state["base_datos_actas"].to_excel(writer, sheet_name="1. Detalle por Actas", index=False)
            st.session_state["base_datos_observadas"].to_excel(writer, sheet_name="7. Historial de Observadas", index=False)
        return output.getvalue()
        
    for col in columnas_computo:
        df_actas[col] = pd.to_numeric(df_actas[col], errors='coerce').fillna(0)

    # Computacion piramidal automatica por niveles geograficos
    df_recintos = df_actas.groupby(["Departamento", "Circunscripcion", "Provincia", "Recinto"])[columnas_computo].sum().reset_index()
    df_circunscripciones = df_actas.groupby(["Departamento", "Circunscripcion"])[columnas_computo].sum().reset_index()
    df_provincias = df_actas.groupby(["Departamento", "Provincia"])[columnas_computo].sum().reset_index()
    df_departamentos = df_actas.groupby(["Departamento"])[columnas_computo].sum().reset_index()
    
    total_nacional_valores = df_departamentos[columnas_computo].sum().to_dict()
    total_nacional_valores["Ambito"] = "TOTAL NACIONAL"
    df_nacional = pd.DataFrame([total_nacional_valores])
    df_nacional = df_nacional[["Ambito"] + columnas_computo]
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_actas.to_excel(writer, sheet_name="1. Detalle por Actas", index=False)
        df_recintos.to_excel(writer, sheet_name="2. Suma por Recinto", index=False)
        df_circunscripciones.to_excel(writer, sheet_name="3. Suma por Circunscripcion", index=False)
        df_provincias.to_excel(writer, sheet_name="4. Suma por Provincia", index=False)
        df_departamentos.to_excel(writer, sheet_name="5. Computo Departamental", index=False)
        df_nacional.to_excel(writer, sheet_name="6. Computo Total Nacional", index=False)
        st.session_state["base_datos_observadas"].to_excel(writer, sheet_name="7. Historial de Observadas", index=False)
    
    return output.getvalue()

# --- FUNCIÓN DE ALMACENAMIENTO AUTOMÁTICO EN LA UNIDAD E ---
def guardar_localmente_en_unidad_e(acta_datos, archivo_foto):
    try:
        CARPETA_DESTINO.mkdir(parents=True, exist_ok=True)
        CARPETA_FOTOS.mkdir(parents=True, exist_ok=True)
        
        ext = Path(archivo_foto.name).suffix if archivo_foto else ".png"
        ruta_copia_foto = CARPETA_FOTOS / f"Mesa_{acta_datos['Mesa']}{ext}"
        
        if archivo_foto is not None:
            with open(ruta_copia_foto, "wb") as f:
                f.write(archivo_foto.getbuffer())
        return True
    except Exception:
        return False

# --- NAVEGACION ---
menu_lateral = st.sidebar.radio("Navegacion del Sistema", ["📥 Procesar Actas (Operadores)", "🔐 Resultados Oficiales (Privado)"])

# ==================== MODULO PARA OPERADORES ====================
if menu_lateral == "📥 Procesar Actas (Operadores)":
    st.title("🗳️ Panel de Transcripcion Oficial (LIBRE vs SPT)")
    st.write("Sube la foto del acta. La IA detectara el codigo preimpreso y ubicara el recinto automaticamente.")
    
    foto_subida = st.file_uploader("Sube el acta aqui (JPG, PNG)", type=["jpg", "jpeg", "png"])

    if foto_subida is not None:
        col1, col2 = st.columns(2)
        imagen = Image.open(foto_subida)
        with col1:
            st.image(imagen, caption="Imagen del acta cargada", use_container_width=True)

        with col2:
            st.subheader("⚙️ Procesamiento de la IA")
            with st.spinner("Buscando codigo preimpreso en el acta..."):
                texto_detectado = " ".join(lector_ia.readtext(foto_subida.getvalue(), detail=0))
                mesas_encontradas = re.findall(r'\b\d{6}-\d\b', texto_detectado)
                codigo_mesa_automatico = mesas_encontradas if mesas_encontradas else "703579-1"

            if codigo_mesa_automatico in PADRON_MESAS:
                info_geo = PADRON_MESAS[codigo_mesa_automatico]
                st.success(f"🆔 **Mesa Detectada:** {codigo_mesa_automatico} | 🏫 **Recinto:** {info_geo['recinto']}")
                
                with st.form("formulario_operador"):
                    # Bloqueo geografico automatico (Solo lectura)
                    st.write("#### 📍 Datos Tecnicos de Ubicacion (Fijos e Inalterables)")
                    g1, g2 = st.columns(2)
                    with g1:
                        st.text_input("Departamento", value=info_geo["departamento"], disabled=True)
                        st.text_input("Provincia", value=info_geo["provincia"], disabled=True)
                    with g2:
                        st.text_input("Circunscripcion", value=info_geo["circunscripcion"], disabled=True)
                        st.text_input("Recinto Electoral", value=info_geo["recinto"], disabled=True)
                    
                    st.write("---")
                    st.write("#### 🔢 Votos por Frente Politico")
                    v_lib = st.number_input("Votos LIBRE", min_value=0, value=94)
                    v_spt = st.number_input("Votos SPT (Santa Cruz Para Todos)", min_value=0, value=66)
                    
                    st.write("#### 📋 Casillas de Control")
                    cx, cy, cz = st.columns(3)
                    v_val_dec = cx.number_input("Casilla total VOTOS VALIDOS", min_value=0, value=160)
                    v_bla = cy.number_input("Casilla VOTOS BLANCOS", min_value=0, value=1)
                    v_nul = cz.number_input("Casilla VOTOS NULOS", min_value=0, value=10)
                    v_tot_anf = st.number_input("Total en ANFORA", min_value=0, value=171)
                    
                    st.write("#### ⚖️ Auditoria de Incidencias")
                    tiene_corre_vale_simple = st.checkbox("⚠️ Tiene 'CORRE Y VALE' por errores de suma")
                    confundieron_partidos = st.checkbox("🚨 'CORRE Y VALE' indica confusion o cruce de votos entre ambos partidos")
                    
                    if st.form_submit_button("🚀 Transmitir Computo"):
                        suma_manual_partidos = v_lib + v_spt
                        suma_manual_anfora = v_val_dec + v_bla + v_nul
                        
                        matematica_cuadra_sola = (suma_manual_partidos == v_val_dec) and (suma_manual_anfora == v_tot_anf)
                        
                        datos_acta_dinamicos = {
                            "Mesa": codigo_mesa_automatico, "Departamento": info_geo["departamento"], "Provincia": info_geo["provincia"],
                            "Circunscripcion": info_geo["circunscripcion"], "Recinto": info_geo["recinto"]
                        }
                        
                        if matematica_cuadra_sola or tiene_corre_vale_simple or confundieron_partidos:
                            if confundieron_partidos:
                                estado = "Computada (Observada por Cruce de Partidos)"
                            elif tiene_corre_vale_simple:
                                estado = "Computada por 'CORRE Y VALE' (Suma)"
                            else:
                                estado = "Computada Normal"
                            
                            nueva_fila = pd.DataFrame([{
                                "Mesa": codigo_mesa_automatico, "Departamento": info_geo["departamento"], "Provincia": info_geo["provincia"],
                                "Circunscripcion": info_geo["circunscripcion"], "Recinto": info_geo["recinto"],
                                "LIBRE": v_lib, "SPT": v_spt,
                                "Votos Validos (Frentes)": v_val_dec, "Votos Blancos": v_bla, "Votos Nulos": v_nul, "Total Anfora": v_tot_anf, "Estado Acta": estado
                            }])
                            
                            if confundieron_partidos:
                                fila_obs = pd.DataFrame([{"Mesa": codigo_mesa_automatico, "Departamento": info_geo["departamento"], "Provincia": info_geo["provincia"], "Circunscripcion": info_geo["circunscripcion"], "Recinto": info_geo["recinto"], "Motivo de Observacion Legal": "Los jurados intercambiaron las casillas de LIBRE y SPT"}])
                                st.session_state["base_datos_observadas"] = st.session_state["base_datos_observadas"][st.session_state["base_datos_observadas"]["Mesa"] != codigo_mesa_automatico]
                                st.session_state["base_datos_observadas"] = pd.concat([st.session_state["base_datos_observadas"], fila_obs], ignore_index=True)
                            else:
                                if not st.session_state["base_datos_observadas"].empty:
                                    st.session_state["base_datos_observadas"] = st.session_state["base_datos_observadas"][st.session_state["base_datos_observadas"]["Mesa"] != codigo_mesa_automatico]

                            st.session_state["base_datos_actas"] = st.session_state["base_datos_actas"][st.session_state["base_datos_actas"]["Mesa"] != codigo_mesa_automatico]
                            st.session_state["base_datos_actas"] = pd.concat([st.session_state["base_datos_actas"], nueva_fila], ignore_index=True)
                            
                            guardar_localmente_en_unidad_e(datos_acta_dinamicos, foto_subida)
                            st.balloons()
                            st.success(f"📊 ¡Exito! Los votos se sumaron inmediatamente al recinto.")
                        else:
                            st.error("❌ RECHAZO DIGITAL: El acta esta descuadrada y no registra fe de erratas fisica.")
            else:
                st.error(f"❌ MESA NO REGISTRADA: El codigo '{codigo_mesa_automatico}' no figura en el Padron Oficial.")

# ==================== MODULO PRIVADO EXCLUSIVO ====================
elif menu_lateral == "🔐 Resultados Oficiales (Privado)":
    st.title("🔐 Panel de Control y Resultados Consolidados")
    password_introduced = st.text_input("Introduce la clave de acceso autorizada:", type="password")
    
    if password_introduced == CONTRASENA_DIRECTOR:
        st.success("🔓 Acceso concedido.")
        
        st.write("---")
        st.subheader("⚖️ Modulo de Saneamiento Jurisdiccional (Ordenes del Tribunal)")
        mesas_observadas_lista = ["Seleccionar..."] if st.session_state["base_datos_observadas"].empty else ["Seleccionar..."] + list(st.session_state["base_datos_observadas"]["Mesa"].values)
        mesa_a_sanear = st.selectbox("Elija la mesa observada a modificar segun Resolucion del Tribunal:", mesas_observadas_lista)
        
        if mesa_a_sanear != "Seleccionar...":
            acta_original = st.session_state["base_datos_actas"][st.session_state["base_datos_actas"]["Mesa"] == mesa_a_sanear].iloc
            with st.form("formulario_saneamiento_tribunal"):
                st.write(f"### Ingrese la Votacion Rectificada para la Mesa: {mesa_a_sanear}")
                t_lib = st.number_input("LIBRE Oficial", min_value=0, value=int(acta_original["LIBRE"]))
                t_spt = st.number_input("SPT Oficial", min_value=0, value=int(acta_original["SPT"]))
                st.write("---")
                c_val = st.number_input("Votos Validos Oficiales", min_value=0, value=int(acta_original["Votos Validos (Frentes)"]))
                c_bla = st.number_input("Votos Blancos Oficiales", min_value=0, value=int(acta_original["Votos Blancos"]))
                c_nul = st.number_input("Votos Nulos Oficiales", min_value=0, value=int(acta_original["Votos Nulos"]))
                c_anf = st.number_input("Total Anfora Oficial", min_value=0, value=int(acta_original["Total Anfora"]))
                
                if st.form_submit_button("⚖️ Aplicar Sentencia Electoral"):
                    fila_saneada = pd.DataFrame([{
                        "Mesa": mesa_a_sanear, "Departamento": acta_original["Departamento"], "Provincia": acta_original["Provincia"],
                        "Circunscripcion": acta_original["Circunscripcion"], "Recinto": acta_original["Recinto"],
                        "LIBRE": t_lib, "SPT": t_spt,
                        "Votos Validos (Frentes)": c_val, "Votos Blancos": c_bla, "Votos Nulos": c_nul, "Total Anfora": c_anf, "Estado Acta": "SANEADA POR EL TRIBUNAL ELECTORAL"
                    }])
                    st.session_state["base_datos_observadas"] = st.session_state["base_datos_observadas"][st.session_state["base_datos_observadas"]["Mesa"] != mesa_a_sanear]
                    st.session_state["base_datos_actas"] = st.session_state["base_datos_actas"][st.session_state["base_datos_actas"]["Mesa"] != mesa_a_sanear]
                    st.session_state["base_datos_actas"] = pd.concat([st.session_state["base_datos_actas"], fila_saneada], ignore_index=True)
                    procesar_tablas_computo()
                    st.success(f"⚖️ ¡Mesa {mesa_a_sanear} Modificada por el Tribunal! Computo recalculado.")
                    st.rerun()
st.write("---")
total_validas = len(st.session_state["base_datos_actas"])
total_observadas = len(st.session_state["base_datos_observadas"])

c1, c2 = st.columns(2)
with c1: st.metric(label="🗳️ Total de Actas Sumadas/Computadas (Nacional)", value=total_validas)
with c2: st.metric(label="⚠️ Acumulado Paralelo de Actas Observadas (Pendientes)", value=total_observadas, delta=total_observadas, delta_color="inverse")

    if total_validas > 0:
        st.write("---")
        st.subheader("📊 Tendencia Electoral Nacional (LIBRE vs SPT)")
        
        partidos = ["LIBRE", "SPT"]
        votos_totales_partidos = st.session_state["base_datos_actas"][partidos].sum()
        
        df_grafico = pd.DataFrame({
            "Frente Politico": ["LIBRE", "Santa Cruz Para Todos (SPT)"],
            "Votos Totales": [votos_totales_partidos["LIBRE"], votos_totales_partidos["SPT"]]
        })
        
        st.pie_chart(data=df_grafico, names="Frente Politico", values="Votos Totales", use_container_width=True)
        
        excel_binario = procesar_tablas_computo()
        st.download_button(label="🟩 DESCARGAR LIBRO ELECTORAL CENTRAL (Excel con 7 Pestanas)", data=excel_binario, file_name="computo_electoral_bolivia.xlsx")
        st.write("### 📊 Detalle del Computo General")
        st.dataframe(st.session_state["base_datos_actas"], use_container_width=True)
    else:
        st.info("💡 Esperando la transmision de datos desde las circunscripciones.")
elif password_introduced != "":
    st.error("❌ Clave incorrecta.")