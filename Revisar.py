"""Aplicacion para revisar la informacion de ordenanzas municipales y auditar carpetas."""

import os
from io import BytesIO

import pandas as pd
import streamlit as st


DEFAULT_REQUIRED_COLUMNS = "Municipio, Numero de ordenanza, Fecha, Descripcion"


def validate_ordinances(dataframe: pd.DataFrame, required_columns: list[str]) -> pd.DataFrame:
    """Valida que las columnas existan y no estén vacías."""
    result = dataframe.copy()
    missing_columns = [
        column for column in required_columns if column not in result.columns
    ]

    if missing_columns:
        raise ValueError(
            "No se encontraron estas columnas en el Excel: "
            + ", ".join(missing_columns)
        )

    empty_fields = result[required_columns].isna() | result[required_columns].astype(str).apply(
        lambda column: column.str.strip().eq("")
    )
    result["Detalle de validacion"] = empty_fields.apply(
        lambda row: "Falta: " + ", ".join(row.index[row].tolist())
        if row.any()
        else "Informacion completa",
        axis=1,
    )
    result["Estado Datos"] = empty_fields.any(axis=1).map({True: "NO OK", False: "OK"})
    return result


def validate_folders(dataframe: pd.DataFrame, base_path: str) -> pd.DataFrame:
    """Verifica si los archivos físicos coinciden con lo que dice el Excel."""
    result = dataframe.copy()
    
    # Buscamos las columnas sin importar si están en mayúsculas o minúsculas
    col_municipio = next((c for c in result.columns if c.strip().lower() == "municipio"), None)
    col_check = next((c for c in result.columns if c.strip().lower() == "check"), None)

    if not col_municipio or not col_check:
        st.warning("⚠️ Para auditar las carpetas, el Excel DEBE tener una columna llamada 'Municipio' y otra llamada 'Check'.")
        return result

    estados_carpeta = []

    # Iteramos fila por fila aplicando EXACTAMENTE tu lógica original
    for index, row in result.iterrows():
        municipio = str(row[col_municipio]).strip().lower() # Convertimos a minúscula (ej. "laplata")
        valor_check = str(row[col_check]).strip().lower()
        
        ruta_municipio = os.path.join(base_path, municipio)
        
        # 1) Verificamos si la carpeta física existe
        if not os.path.exists(ruta_municipio):
            estados_carpeta.append(f"ERROR: No existe la carpeta '{municipio}'")
            continue
            
        # 2) Leemos el contenido de la carpeta
        try:
            contenido = os.listdir(ruta_municipio)
        except Exception as e:
            estados_carpeta.append(f"Error al leer carpeta: {e}")
            continue

        # 3) Aplicamos tu lógica de los "ok" y "no hay ordenanza"
        if len(contenido) > 0:
            # Hay archivo en la carpeta
            if valor_check in ["ok", "no hay informacion"]:
                estados_carpeta.append("CORRECTO (Tiene archivo)")
            else:
                estados_carpeta.append(f"INCONSISTENCIA: Excel dice '{valor_check}' | Carpeta: Tiene archivo")
        else:
            # La carpeta está vacía
            if valor_check == "no hay ordenanza":
                estados_carpeta.append("CORRECTO (Carpeta vacía)")
            else:
                estados_carpeta.append(f"INCONSISTENCIA: Excel dice '{valor_check}' | Carpeta: Vacía")

    # Agregamos los resultados como una nueva columna en el Excel final
    result["Estado Carpeta Física"] = estados_carpeta
    return result


def read_excel(uploaded_file, header_index: int) -> dict[str, pd.DataFrame]:
    """Lee el archivo Excel usando la fila seleccionada como títulos."""
    return pd.read_excel(BytesIO(uploaded_file.getvalue()), sheet_name=None, header=header_index)


# --- INTERFAZ VISUAL ---
st.set_page_config(page_title="Auditoría de Ordenanzas TISH", page_icon="📁", layout="wide")
st.title("Control Integral de Ordenanzas Municipales")
st.write("Verifica la información del Excel y audita que los archivos físicos existan en tus carpetas.")

st.divider() # Línea separadora

# Sección 1: Carga de datos
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("1. Selecciona el archivo Excel", type=["xlsx", "xls"])
with col2:
    fila_titulos = st.number_input("Fila de los títulos", min_value=1, value=7, step=1)

required_text = st.text_input(
    "2. Columnas obligatorias de texto (separadas por coma)",
    value=DEFAULT_REQUIRED_COLUMNS,
)

# Sección 2: Auditoría de carpetas físicas
st.markdown("### 3. Auditoría Física de Archivos (Opcional)")
verificar_carpetas = st.checkbox("Habilitar revisión de carpetas locales")
carpeta_base = ""
if verificar_carpetas:
    carpeta_base = st.text_input(
        "Ruta de la carpeta madre (Ej: C:\\Users\\loren\\Desktop\\Pyton\\Proyecto 1\\SIMULACION)",
        help="La carpeta que contiene todas las subcarpetas de los municipios."
    )

st.divider()

if uploaded_file is None:
    st.info("Sube un archivo para comenzar.")
else:
    try:
        header_index_python = fila_titulos - 1
        sheets = read_excel(uploaded_file, header_index_python)
        
        sheet_name = st.selectbox("Hoja a revisar", list(sheets))
        dataframe = sheets[sheet_name]
        required_columns = [column.strip() for column in required_text.split(",") if column.strip()]

        if not required_columns:
            st.warning("Indica al menos una columna obligatoria.")
        else:
            # 1. Validamos las celdas vacías del Excel
            checked_data = validate_ordinances(dataframe, required_columns)
            
            # 2. Validamos las carpetas si el usuario lo activó
            if verificar_carpetas and carpeta_base:
                if os.path.exists(carpeta_base):
                    checked_data = validate_folders(checked_data, carpeta_base)
                else:
                    st.error("❌ La ruta de la carpeta base no existe o está mal escrita. Revisa la ruta.")

            # Mostrar métricas
            ok_datos = int((checked_data["Estado Datos"] == "OK").sum())
            
            col_metricas1, col_metricas2 = st.columns(2)
            col_metricas1.metric("Registros revisados", len(checked_data))
            col_metricas2.metric("Filas con Datos OK", ok_datos)
            
            # Mostrar tabla
            st.dataframe(checked_data, use_container_width=True, hide_index=True)

            # Botón de descarga
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                checked_data.to_excel(writer, index=False, sheet_name="Auditoria")
            
            st.download_button(
                "Descargar Reporte Final",
                data=output.getvalue(),
                file_name="reporte_auditoria_ordenanzas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" # Hace que el botón sea rojo/color principal
            )
    except ValueError as error:
        st.error(str(error))
    except Exception as error:
        st.error(f"Error inesperado: {error}")