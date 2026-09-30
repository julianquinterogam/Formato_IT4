"""
App IT4 - Interrupciones del servicio (SUI)
Parte 1: registros desde el IT2 final (maniobras de mantenimiento, causal 5)
Parte 2: PQR_Dispower cruzado con BASE por NUI (servicio no disponible, causal 7)
Salida: union de ambas partes (primero IT2, luego PQR) -> IT4 editable
IT4 final: cruce con IT1 por serial + filtro de fechas del mes reportado
"""
import calendar
import hashlib
import io
import re
import unicodedata

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------------
MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

COLUMNAS_IT4 = [
    "Codigo de Localidad",
    "Fecha y hora inicio",
    "Fecha y hora fin",
    "Causal de no prestacion",
    "Tipo elemento afectado",
    "Serial del elemento afectado",
    "Observacion",
]

COLUMNAS_IT2 = [
    "Codigo Localidad",
    "Tipo de Elemento",
    "Serial del Elemento",
    "Fecha y hora de inicio",
    "Fecha y hora fin",
]

CAUSAL_MANTENIMIENTO = 5
OBSERVACION_MANTENIMIENTO = "Ejecucion de mantenimientos"
CAUSAL_PQR = 7

COLUMNAS_BASE = ["NIU", "cod_localidad", "TIPO_UC", "SERIAL_INTERNO"]
COLUMNAS_PQR = ["NUI", "FechaCreacion", "FechaCierre", "concatenado formato it4"]

# Homologacion TIPO_UC (BASE) -> Tipo elemento afectado (IT4)
HOMOLOGACION_TIPO_UC = {1: 5, 2: 6, 3: 7, 4: 10, 5: 13, 6: 21, 7: 21, 8: 21, 9: 21, 10: 21, 11: 19}

# Horario laboral para las horas generadas (minutos desde medianoche)
HORA_MIN = 8 * 60    # 08:00
HORA_MAX = 17 * 60   # 17:00
FORMATO_FECHA = "%d-%m-%Y %H:%M"
MAX_OBSERVACION = 400


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def limpiar_texto(texto, max_len=MAX_OBSERVACION):
    """Quita tildes, deja solo letras, numeros y espacios, y corta a max_len sin partir palabras."""
    if pd.isna(texto):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Za-z0-9 ]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    if len(texto) <= max_len:
        return texto
    corte = texto[:max_len + 1].rsplit(" ", 1)[0]  # no cortar palabras a la mitad
    return corte.strip()


@st.cache_data(show_spinner="Leyendo archivo...")
def _leer_excel(contenido):
    return pd.read_excel(io.BytesIO(contenido), sheet_name=0)


def corregir_tipo_14(df, columna):
    """El poste debe ir con tipo 13, no 14 (el 14 genera error en el SUI).
    Devuelve (df corregido, filas corregidas)."""
    df = df.copy()
    tipo = pd.to_numeric(df[columna], errors="coerce")
    corregidos = int((tipo == 14).sum())
    df[columna] = tipo.where(tipo != 14, 13).astype("Int64")
    return df, corregidos


def leer_primera_hoja(archivo):
    return _leer_excel(archivo.getvalue()).copy()


def validar_columnas(df, requeridas, nombre):
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        st.error(f"Al archivo {nombre} le faltan las columnas: {', '.join(faltantes)}")
        st.stop()


# ----------------------------------------------------------------------------
# Parte 1: IT2 -> IT4
# ----------------------------------------------------------------------------
def procesar_it2(df_it2, mes, anio):
    df = df_it2.copy()
    df["Fecha y hora de inicio"] = pd.to_datetime(df["Fecha y hora de inicio"], errors="coerce")
    df["Fecha y hora fin"] = pd.to_datetime(df["Fecha y hora fin"], errors="coerce")

    df, corregidos_14 = corregir_tipo_14(df, "Tipo de Elemento")

    en_mes = (df["Fecha y hora de inicio"].dt.month == mes) & (df["Fecha y hora de inicio"].dt.year == anio)
    descartados = df[~en_mes]
    df = df[en_mes]

    salida = pd.DataFrame({
        "Codigo de Localidad": df["Codigo Localidad"],
        "Fecha y hora inicio": df["Fecha y hora de inicio"].dt.strftime(FORMATO_FECHA),
        "Fecha y hora fin": df["Fecha y hora fin"].dt.strftime(FORMATO_FECHA),
        "Causal de no prestacion": CAUSAL_MANTENIMIENTO,
        "Tipo elemento afectado": df["Tipo de Elemento"],
        "Serial del elemento afectado": df["Serial del Elemento"],
        "Observacion": OBSERVACION_MANTENIMIENTO,
    })[COLUMNAS_IT4]
    return salida.reset_index(drop=True), descartados, corregidos_14


# ----------------------------------------------------------------------------
# Parte 2: PQR_Dispower + BASE -> IT4
# ----------------------------------------------------------------------------
def _numero_fijo(*partes):
    """Numero 'aleatorio' pero fijo: siempre el mismo para los mismos datos."""
    semilla = "|".join(str(p) for p in partes)
    return int(hashlib.md5(semilla.encode()).hexdigest()[:8], 16)


def asignar_horas(nui, creacion, cierre):
    """Devuelve (inicio, fin) con horas fijas entre 08:00 y 17:00. Fin siempre > inicio."""
    dia_ini = creacion.normalize()
    min_ini = HORA_MIN + _numero_fijo(nui, dia_ini.date(), "ini") % (HORA_MAX - HORA_MIN)  # 08:00 a 16:59
    inicio = dia_ini + pd.Timedelta(minutes=min_ini)

    if pd.isna(cierre):
        return inicio, pd.NaT

    dia_fin = cierre.normalize()
    if dia_fin == dia_ini:
        min_fin = min_ini + 1 + _numero_fijo(nui, dia_fin.date(), "fin") % (HORA_MAX - min_ini)
    else:
        min_fin = HORA_MIN + _numero_fijo(nui, dia_fin.date(), "fin") % (HORA_MAX - HORA_MIN + 1)
    return inicio, dia_fin + pd.Timedelta(minutes=min_fin)


def procesar_pqr(df_pqr, df_base, mes, anio):
    pqr = df_pqr.copy()
    pqr["FechaCreacion"] = pd.to_datetime(pqr["FechaCreacion"], errors="coerce")
    pqr["FechaCierre"] = pd.to_datetime(pqr["FechaCierre"], errors="coerce")

    # PQR activas en el mes: creadas antes de que termine el mes y
    # cerradas despues de que empiece (o todavia abiertas)
    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1], 23, 59, 59)
    activas = (pqr["FechaCreacion"] <= fin_mes) & (
        pqr["FechaCierre"].isna() | (pqr["FechaCierre"] >= inicio_mes)
    )
    pqr = pqr[activas]

    # Cruce con BASE: todos los elementos del NUI; los NUI que no esten se eliminan
    base = df_base[COLUMNAS_BASE].copy()
    base = base[base["TIPO_UC"].isin(HOMOLOGACION_TIPO_UC)]
    sin_base = pqr[~pqr["NUI"].isin(base["NIU"])]
    pqr = pqr[pqr["NUI"].isin(base["NIU"])].reset_index(drop=True)

    horas = [asignar_horas(n, c, f) for n, c, f in zip(pqr["NUI"], pqr["FechaCreacion"], pqr["FechaCierre"])]
    pqr["_inicio"] = [h[0] for h in horas]
    pqr["_fin"] = [h[1] for h in horas]
    pqr["_orden"] = range(len(pqr))

    # Solo las columnas necesarias del PQR (el PQR tambien trae cod_localidad; manda el de BASE)
    pqr_min = pqr[COLUMNAS_PQR + ["_inicio", "_fin", "_orden"]]
    cruce = pqr_min.merge(base, left_on="NUI", right_on="NIU", how="inner")
    cruce = cruce.sort_values(["_inicio", "NUI", "_orden"], kind="stable")

    salida = pd.DataFrame({
        "Codigo de Localidad": cruce["cod_localidad"],
        "Fecha y hora inicio": cruce["_inicio"].dt.strftime(FORMATO_FECHA),
        "Fecha y hora fin": cruce["_fin"].dt.strftime(FORMATO_FECHA).fillna(""),
        "Causal de no prestacion": CAUSAL_PQR,
        "Tipo elemento afectado": cruce["TIPO_UC"].map(HOMOLOGACION_TIPO_UC),
        "Serial del elemento afectado": cruce["SERIAL_INTERNO"],
        "Observacion": cruce["concatenado formato it4"].apply(limpiar_texto),
    })[COLUMNAS_IT4]

    resumen = {
        "pqr_mes": int(activas.sum()),
        "pqr_sin_base": len(sin_base),
        "pqr_reportadas": len(pqr),
        "pqr_abiertas": int(pqr["FechaCierre"].isna().sum()),
    }
    return salida.reset_index(drop=True), sin_base, resumen


# ----------------------------------------------------------------------------
# IT4 final: cruce con IT1 + filtro de fechas
# ----------------------------------------------------------------------------
COLUMNAS_IT1 = ["serial", "cod_localidad"]


def cruzar_localidad_it1(df_it4, df_it1):
    """Reemplaza el Codigo de Localidad por el del IT1 (cruce por serial).
    Si el serial no aparece en el IT1, se deja el codigo que ya traia."""
    it1 = df_it1[COLUMNAS_IT1].dropna(subset=["serial"]).copy()
    it1["serial"] = it1["serial"].astype(str).str.strip()
    it1 = it1.drop_duplicates("serial")
    mapa = dict(zip(it1["serial"], it1["cod_localidad"]))

    df = df_it4.copy()
    seriales = df["Serial del elemento afectado"].astype(str).str.strip()
    nuevo = seriales.map(mapa)
    sin_cruce = seriales[nuevo.isna()]
    df["Codigo de Localidad"] = nuevo.fillna(df["Codigo de Localidad"])
    df["Codigo de Localidad"] = pd.to_numeric(df["Codigo de Localidad"], errors="coerce").astype("Int64")
    return df, sin_cruce


def filtrar_fechas_mes(df_it4, mes, anio):
    """Deja solo los registros con inicio del mes reportado o anterior
    y fin dentro del mes reportado. Devuelve (final, excluidos con motivo)."""
    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1], 23, 59, 59)

    ini = pd.to_datetime(df_it4["Fecha y hora inicio"], format=FORMATO_FECHA, errors="coerce")
    fin = pd.to_datetime(df_it4["Fecha y hora fin"], format=FORMATO_FECHA, errors="coerce")

    motivo = pd.Series("", index=df_it4.index)
    motivo[ini > fin_mes] = "Inicio posterior al mes"
    motivo[fin < inicio_mes] = "Fin anterior al mes"
    motivo[fin > fin_mes] = "Fin posterior al mes"
    motivo[fin.isna()] = "Sin fecha fin"
    motivo[ini.isna()] = "Sin fecha inicio"

    queda = motivo == ""
    excluidos = df_it4[~queda].assign(Motivo=motivo[~queda])
    return df_it4[queda].reset_index(drop=True), excluidos


def a_csv(df):
    """CSV delimitado por comas, UTF-8, con encabezados."""
    return df.to_csv(index=False, sep=",").encode("utf-8")


def a_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="IT4")
        ws = writer.sheets["IT4"]
        anchos = [20, 18, 18, 22, 22, 28, 80]
        for i, ancho in enumerate(anchos):
            ws.column_dimensions[chr(65 + i)].width = ancho
        for celda in ws["A"][1:]:
            celda.number_format = "0"
    return buffer.getvalue()


# ----------------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------------
def tab_generar(mes, anio):
    it4_mtto = it4_pqr = None

    st.subheader("1. Interrupciones por mantenimiento (IT2) · causal 5")
    archivo_it2 = st.file_uploader("IT2 final (9 columnas)", type=["xlsx"], key="it2")

    if archivo_it2:
        df_it2 = leer_primera_hoja(archivo_it2)
        validar_columnas(df_it2, COLUMNAS_IT2, "IT2")
        it4_mtto, descartados, corregidos_14 = procesar_it2(df_it2, mes, anio)
        if corregidos_14:
            st.info(f"{corregidos_14} filas del IT2 tenian Tipo de Elemento = 14; se corrigieron a 13.")

        c1, c2, c3 = st.columns(3)
        c1.metric("Filas en IT2", len(df_it2))
        c2.metric(f"Filas de {MESES[mes]} {anio}", len(it4_mtto))
        c3.metric("Fuera del periodo", len(descartados))
        if len(descartados):
            with st.expander("Ver filas del IT2 fuera del periodo"):
                st.dataframe(descartados, use_container_width=True)

    st.subheader("2. Servicio no disponible (PQR) · causal 7")
    col_a, col_b = st.columns(2)
    archivo_base = col_a.file_uploader("BASE", type=["xlsx"], key="base")
    archivo_pqr = col_b.file_uploader("PQR_Dispower", type=["xlsx"], key="pqr")

    if archivo_base and archivo_pqr:
        df_base = leer_primera_hoja(archivo_base)
        df_pqr = leer_primera_hoja(archivo_pqr)
        validar_columnas(df_base, COLUMNAS_BASE, "BASE")
        validar_columnas(df_pqr, COLUMNAS_PQR, "PQR_Dispower")
        it4_pqr, sin_base, r = procesar_pqr(df_pqr, df_base, mes, anio)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(f"PQR activas en {MESES[mes]}", r["pqr_mes"])
        c2.metric("Eliminadas (NUI sin BASE)", r["pqr_sin_base"])
        c3.metric("PQR reportadas", r["pqr_reportadas"], help=f"{r['pqr_abiertas']} siguen abiertas (sin FechaCierre)")
        c4.metric("Filas generadas", len(it4_pqr))
        if len(sin_base):
            with st.expander("Ver PQR eliminadas por NUI sin BASE"):
                st.dataframe(sin_base[["NUI"] + [c for c in df_pqr.columns if c != "NUI"]], use_container_width=True)

    st.subheader("3. IT4 editable")
    partes = [p for p in (it4_mtto, it4_pqr) if p is not None]
    if not partes:
        st.session_state.pop("it4_editable", None)
        st.info("Sube el IT2 y/o los archivos BASE y PQR_Dispower para generar el IT4.")
        return

    if it4_mtto is None:
        st.warning("Falta el IT2: el editable solo tiene los registros de PQR.")
    if it4_pqr is None:
        st.warning("Faltan BASE y/o PQR_Dispower: el editable solo tiene los registros del IT2.")

    it4 = pd.concat(partes, ignore_index=True)
    st.session_state["it4_editable"] = {"df": it4, "mes": mes, "anio": anio}

    st.metric("Filas IT4 editable", len(it4))
    st.dataframe(it4, use_container_width=True, height=450)
    st.download_button(
        "Descargar IT4 editable",
        data=a_excel(it4),
        file_name=f"IT4_{mes:02d}_{anio}_editable.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        help="Todos los registros, incluidos los que no tienen fecha fin o cierran despues del mes",
    )
    st.caption("Para el archivo de envio, ve a la pestaña **IT4 final (con IT1)**.")


def tab_final(mes, anio):
    st.subheader("Origen del IT4")
    generado = st.session_state.get("it4_editable")
    usar_generado = False
    if generado and generado["mes"] == mes and generado["anio"] == anio:
        usar_generado = st.checkbox(
            f"Usar el IT4 generado en la otra pestaña ({len(generado['df']):,} filas)", value=True
        )

    if usar_generado:
        origen = generado["df"]
    else:
        archivo = st.file_uploader("IT4 editable", type=["xlsx"], key="it4_ed")
        if not archivo:
            st.info("Genera el IT4 en la otra pestaña o sube un IT4 editable.")
            return
        origen = leer_primera_hoja(archivo)
        validar_columnas(origen, COLUMNAS_IT4, "IT4 editable")
        origen = origen[COLUMNAS_IT4].copy()
        for col in ("Fecha y hora inicio", "Fecha y hora fin"):
            origen[col] = origen[col].fillna("").astype(str).str.strip()
        st.caption(f"Verifica que el archivo corresponda a {MESES[mes]} {anio} (periodo seleccionado en la barra lateral).")

    origen, corregidos_14 = corregir_tipo_14(origen, "Tipo elemento afectado")
    if corregidos_14:
        st.info(f"{corregidos_14} filas del IT4 tenian Tipo elemento afectado = 14; se corrigieron a 13.")

    archivo_it1 = st.file_uploader("IT1", type=["xlsx"], key="it1")
    if not archivo_it1:
        st.info("Sube el IT1 para cruzar el codigo de localidad.")
        return
    df_it1 = leer_primera_hoja(archivo_it1)
    validar_columnas(df_it1, COLUMNAS_IT1, "IT1")

    cruzado, sin_cruce = cruzar_localidad_it1(origen, df_it1)
    final, excluidos = filtrar_fechas_mes(cruzado, mes, anio)

    c1, c2, c3 = st.columns(3)
    c1.metric("Filas IT4 editable", len(origen))
    c2.metric("Excluidas por fechas", len(excluidos))
    c3.metric("Filas IT4 final", len(final))

    if len(sin_cruce):
        st.warning(f"{len(sin_cruce)} filas con un serial que no aparece en el IT1; se dejo el Codigo de Localidad que ya traian.")
        with st.expander("Ver esos seriales"):
            st.dataframe(sin_cruce.rename("Serial del elemento afectado").drop_duplicates(), use_container_width=True)

    if len(excluidos):
        with st.expander("Ver registros excluidos por fechas"):
            resumen = excluidos.groupby(["Causal de no prestacion", "Motivo"]).size().rename("Filas").reset_index()
            st.dataframe(resumen, use_container_width=True, hide_index=True)
            st.dataframe(excluidos, use_container_width=True)
        n_it2 = int((excluidos["Causal de no prestacion"] == CAUSAL_MANTENIMIENTO).sum())
        if n_it2:
            st.warning(f"{n_it2} filas del IT2 (causal 5) quedaron fuera por fechas. Revisalas en el detalle.")

    st.dataframe(final, use_container_width=True, height=450)

    nombre = f"IT4_{mes:02d}_{anio}"
    d1, d2 = st.columns(2)
    d1.download_button(
        f"Descargar {nombre}.xlsx",
        data=a_excel(final),
        file_name=f"{nombre}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    d2.download_button(
        f"Descargar {nombre}.csv",
        data=a_csv(final),
        file_name=f"{nombre}.csv",
        mime="text/csv",
    )


st.set_page_config(page_title="IT4 Interrupciones", page_icon="⚡", layout="wide")
st.title("Formato IT4 · Interrupciones del servicio")

with st.sidebar:
    st.header("Periodo a reportar")
    mes_sel = st.selectbox("Mes", list(MESES.keys()), format_func=lambda m: MESES[m])
    anio_sel = int(st.number_input("Año", min_value=2020, max_value=2100, value=2024, step=1))

tab1, tab2 = st.tabs(["Generar IT4", "IT4 final (con IT1)"])
with tab1:
    tab_generar(mes_sel, anio_sel)
with tab2:
    tab_final(mes_sel, anio_sel)
