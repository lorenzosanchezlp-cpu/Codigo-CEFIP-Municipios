"""Aplicacion para revisar la informacion de ordenanzas municipales."""
print ("hola")

from io import BytesIO

import pandas as pd
import streamlit as st


DEFAULT_REQUIRED_COLUMNS = "Año, Codigo municipal, Municipio, Check, Comentario"


def validate_ordinances(dataframe: pd.DataFrame, required_columns: list[str]) -> pd.DataFrame:
	"""Return a copy with a status and explanation for every ordinance row."""
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
	result["Estado"] = empty_fields.any(axis=1).map({True: "NO OK", False: "OK"})
	return result


def read_excel(uploaded_file) -> dict[str, pd.DataFrame]:
	"""Read all worksheets from the uploaded Excel file."""
	return pd.read_excel(BytesIO(uploaded_file.getvalue()), sheet_name=None, header=6)


st.set_page_config(page_title="Control de ordenanzas", page_icon="OK", layout="wide")
st.title("Control de ordenanzas municipales")
st.write("Carga un Excel y verifica si cada registro tiene la informacion obligatoria.")

uploaded_file = st.file_uploader("Selecciona el archivo Excel", type=["xlsx", "xls"])
required_text = st.text_input(
	"Columnas obligatorias (separadas por coma)",
	value=DEFAULT_REQUIRED_COLUMNS,
	help="Escribe los nombres exactamente como aparecen en la primera fila del Excel.",
)

if uploaded_file is None:
	st.info("Selecciona un archivo para comenzar.")
else:
	try:
		sheets = read_excel(uploaded_file)
		sheet_name = st.selectbox("Hoja a revisar", list(sheets))
		dataframe = sheets[sheet_name]
		required_columns = [column.strip() for column in required_text.split(",") if column.strip()]

		if not required_columns:
			st.warning("Indica al menos una columna obligatoria.")
		else:
			checked_data = validate_ordinances(dataframe, required_columns)
			ok_count = int((checked_data["Estado"] == "OK").sum())
			not_ok_count = len(checked_data) - ok_count

			first_column, second_column, third_column = st.columns(3)
			first_column.metric("Registros revisados", len(checked_data))
			second_column.metric("OK", ok_count)
			third_column.metric("NO OK", not_ok_count)
			st.dataframe(checked_data, use_container_width=True, hide_index=True)

			output = BytesIO()
			with pd.ExcelWriter(output, engine="openpyxl") as writer:
				checked_data.to_excel(writer, index=False, sheet_name="Resultado")
			st.download_button(
				"Descargar resultado",
				data=output.getvalue(),
				file_name="resultado_ordenanzas.xlsx",
				mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
			)
	except ValueError as error:
		st.error(str(error))
	except Exception as error:
		st.error(f"No se pudo leer el archivo: {error}")