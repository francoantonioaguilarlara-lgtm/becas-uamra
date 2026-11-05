"""
Programa de Detección de Alumnos con Necesidad de Beca - UAMRA
Aplicación web en Python (Streamlit).

Funciona de dos maneras:
  - EN LÍNEA (Streamlit Community Cloud): los alumnos se registran desde cualquier celular
    y los datos se guardan en la nube (Hoja de Google, ver nube/Codigo.gs).
  - EN LA LAPTOP (Windows + Excel): guarda en Programa_Becas_UAMRA.xlsm y tiene un botón
    para traer a Excel los alumnos que se registraron en línea.

Ejecutar:   streamlit run app.py
"""

import sys

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import excel_db as db
import nube

st.set_page_config(page_title="Becas UAMRA", page_icon="🎓", layout="wide")


def config_nube():
    try:
        cfg = dict(st.secrets["nube"])
    except Exception:
        return None
    return cfg if cfg.get("url") and cfg.get("token") else None


NUBE = config_nube()
# En el servidor (Linux) siempre es "nube"; en Windows se puede forzar con modo = "nube" en secrets.
EN_LINEA = NUBE is not None and (sys.platform != "win32" or NUBE.get("modo") == "nube")

# ---------------------------------------------------------------------------
# Colores (paleta validada para daltonismo; prioridad usa colores de estado)
# ---------------------------------------------------------------------------
SERIE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
COLOR_PRIORIDAD = {"ALTA PRIORIDAD": "#d03b3b", "MEDIA PRIORIDAD": "#fab219", "BAJA PRIORIDAD": "#0ca30c"}
ORDEN_PRIORIDAD = list(COLOR_PRIORIDAD)
COLOR_RESIDENCIA = {"Local": SERIE[0], "Foráneo": SERIE[1]}
ORDEN_BECAS = ["Manutención Integral", "Beca Transporte", "Beca Alimentaria", "Sin Asignación Urgente"]

px.defaults.template = "plotly_white"
px.defaults.color_discrete_sequence = SERIE


def estilo(fig, alto=380):
    fig.update_layout(
        height=alto, margin=dict(l=10, r=10, t=50, b=10),
        font=dict(family="Segoe UI, system-ui, sans-serif", size=13),
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="left", x=0, title=None),
        hoverlabel=dict(font_size=13),
    )
    fig.update_xaxes(showgrid=False, linecolor="#c3c2b7")
    fig.update_yaxes(gridcolor="#e1e0d9", zeroline=False)
    return fig


def dinero(x):
    return f"${x:,.0f}"


# ---------------------------------------------------------------------------
# Datos
# ---------------------------------------------------------------------------
st.sidebar.title("🎓 Becas UAMRA")
pagina = st.sidebar.radio(
    "Ir a:",
    ["📝 Registro de alumnos", "📊 Gráficas y estadística", "🏅 Padrón de becas", "🗂️ Base de datos"],
)
st.sidebar.divider()
modo_demo = st.sidebar.toggle(
    "Modo demostración (datos ficticios)",
    help="Muestra 80 alumnos inventados para probar las gráficas. No modifica los datos reales.",
)


@st.cache_data(ttl=30, show_spinner="Leyendo los datos de la nube…")
def datos_nube(url, token):
    return nube.descargar(url, token)


if EN_LINEA:
    RUTA = None
    params = db.leer_parametros(db.EXCEL_REAL if db.EXCEL_REAL.exists() else db.PLANTILLA)
    st.sidebar.caption("☁️ Conectado a la base de datos en la nube")
    if modo_demo:
        crudos = pd.DataFrame(db.alumnos_ficticios(80, params))
    else:
        try:
            crudos = datos_nube(NUBE["url"], NUBE["token"])
        except nube.ErrorNube as e:
            st.error(f"No se pudieron leer los datos de la nube. {e}")
            crudos = pd.DataFrame(columns=db.COLUMNAS_CAPTURA)
    datos = db.calcular(crudos, params)
else:
    if modo_demo and not db.EXCEL_DEMO.exists():
        with st.spinner("Creando el archivo de demostración en Excel…"):
            db.crear_demo()
    RUTA = db.EXCEL_DEMO if modo_demo else db.EXCEL_REAL
    st.sidebar.caption(f"Archivo en uso: `datos/{RUTA.name}`")
    if modo_demo and st.sidebar.button("Regenerar datos ficticios"):
        with st.spinner("Generando nuevos datos ficticios en Excel…"):
            db.crear_demo(semilla=np.random.randint(1, 10_000))
        st.rerun()
    if NUBE and not modo_demo:
        if st.sidebar.button("☁️ Traer datos de la nube a Excel", width="stretch"):
            try:
                with st.spinner("Descargando de la nube y guardando en Excel…"):
                    res = nube.sincronizar_a_excel(RUTA, NUBE["url"], NUBE["token"])
                st.sidebar.success(f"Nuevos en Excel: {res['agregados']} · ya estaban: {res['ya_estaban']} · "
                                   f"total en la nube: {res['en_nube']}")
                for err in res["errores"]:
                    st.sidebar.warning(f"No se agregó {err}")
            except (nube.ErrorNube, db.ErrorRegistro) as e:
                st.sidebar.error(str(e))
    try:
        params = db.leer_parametros(RUTA)
        datos = db.calcular(db.leer_registros(RUTA), params)
    except PermissionError:
        st.error("No se pudo leer el Excel. Si está abierto con bloqueo, ciérrelo y recargue la página.")
        st.stop()


# ===========================================================================
# 1. REGISTRO
# ===========================================================================
def pagina_registro():
    st.title("📝 Formulario de captura estudiantil")
    destino = "en la **base de datos en línea**" if EN_LINEA else "directamente en la hoja **Base_Datos** del Excel"
    st.caption(f"Programa de detección de alumnos con necesidad de beca · UASLP · UAMRA · "
               f"Los datos se guardan {destino}.")
    if EN_LINEA and modo_demo:
        st.warning("Está en modo demostración: desactívelo en el menú para poder registrarse.")

    if msg := st.session_state.pop("registro_ok", None):
        st.success(msg)
        diag = st.session_state.pop("ultimo_diag", None)
        if diag:
            mostrar_diagnostico(diag)

    # El formulario solo se vacía cuando el registro se guardó (con otra "llave"); si hay un error,
    # el alumno conserva lo que escribió.
    with st.form(f"registro_{st.session_state.get('form_n', 0)}"):
        consentimiento = st.checkbox(
            "Acepto participar de forma voluntaria. Mis respuestas se asocian a mi clave "
            "universitaria, se tratarán de forma confidencial y se usarán solo con fines de "
            "investigación académica."
        )
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Datos del alumno")
            clave = st.text_input("Clave universitaria (6 dígitos)", max_chars=6, placeholder="399809")
            carrera = st.selectbox("Carrera", params["carreras"])
            semestre = st.selectbox("Semestre", params["semestres"])
            residencia = st.radio("Estatus de residencia", params["residencias"], horizontal=True)
            procedencia = st.selectbox("Municipio de procedencia", list(params["geo"]))
            fuente = st.selectbox("Fuente principal de ingresos", list(params["fuente"]))
        with c2:
            st.subheader("Ingresos y gastos semanales ($ MXN)")
            ingreso = st.number_input("Ingreso semanal", min_value=0.0, step=50.0, format="%.0f")
            transp = st.number_input("Gasto semanal en transporte", min_value=0.0, step=10.0, format="%.0f")
            alim = st.number_input("Gasto semanal en alimentación", min_value=0.0, step=10.0, format="%.0f")
            otros = st.number_input("Otros gastos académicos semanales", min_value=0.0, step=10.0, format="%.0f")
        enviado = st.form_submit_button("💾 Enviar registro" if EN_LINEA else "💾 Guardar registro en Excel",
                                        type="primary", width="stretch",
                                        disabled=EN_LINEA and modo_demo)

    if enviado:
        if not consentimiento:
            st.error("Debe aceptar el consentimiento para registrar.")
            return
        r = {
            "Clave_Alumno": clave.strip(), "Carrera": carrera, "Semestre": semestre,
            "Estatus_Residencia": residencia, "Procedencia": procedencia, "Fuente_Ingreso": fuente,
            "Ingreso_Semanal": ingreso, "Gasto_Transporte": transp,
            "Gasto_Alimentacion": alim, "Otros_Gastos_Acad": otros,
        }
        if EN_LINEA:
            errores = db.validar(r, params, set(datos["Clave_Alumno"]) if not datos.empty else set())
            if errores:
                for linea in errores:
                    st.error(linea)
                return
            try:
                with st.spinner("Enviando a la nube…"):
                    nube.enviar_registro(r, NUBE["url"], NUBE["token"])
            except nube.ErrorNube as e:
                st.error(str(e))
                return
            datos_nube.clear()
            st.session_state["registro_ok"] = f"✅ ¡Gracias! El registro {r['Clave_Alumno']} se guardó en línea."
        else:
            try:
                with st.spinner("Guardando en Excel…"):
                    fila = db.agregar_registro(r, RUTA)
            except db.ErrorRegistro as e:
                for linea in str(e).splitlines():
                    st.error(linea)
                return
            st.session_state["registro_ok"] = f"✅ Alumno {r['Clave_Alumno']} guardado en la fila {fila} de Base_Datos."
        st.session_state["ultimo_diag"] = {**r, **db.diagnosticar(r, params)}
        st.session_state["form_n"] = st.session_state.get("form_n", 0) + 1
        st.rerun()

    st.divider()
    st.subheader("Monitoreo global de la muestra")
    kpis(datos)


def mostrar_diagnostico(d):
    st.markdown("#### Diagnóstico en tiempo real del alumno")
    icono = {"ALTA PRIORIDAD": "🔴", "MEDIA PRIORIDAD": "🟡", "BAJA PRIORIDAD": "🟢"}[d["Nivel_Vulnerabilidad"]]
    c = st.columns(4)
    c[0].metric("Gasto total semanal", dinero(d["Gasto_Total_Semanal"]))
    c[1].metric("Saldo / ahorro semanal", dinero(d["Saldo_Ahorro_Semanal"]))
    c[2].metric("Índice de carga fija (ICF)", f"{d['Indice_Carga_Fija_%']:.1%}")
    c[3].metric("Puntaje de necesidad", f"{d['Puntaje_Total']:.1f} / 100")
    c = st.columns(3)
    c[0].metric("Nivel de vulnerabilidad", f"{icono} {d['Nivel_Vulnerabilidad']}")
    c[1].metric("Beca sugerida", d["Beca_Sugerida"])
    c[2].metric("¿Candidato a beca?", "SÍ - CANDIDATO" if d["Requiere_Beca"] == "SÍ" else "NO")


def kpis(df):
    if df.empty:
        st.info("Aún no hay alumnos registrados.")
        return
    n = len(df)
    cand = (df["Requiere_Beca"] == "SÍ").sum()
    c = st.columns(5)
    c[0].metric("Encuestas registradas", n)
    c[1].metric("Locales / Foráneos",
                f"{(df['Estatus_Residencia'] == 'Local').sum()} / {(df['Estatus_Residencia'] == 'Foráneo').sum()}")
    c[2].metric("Candidatos a beca", cand)
    c[3].metric("% que requiere beca", f"{cand / n:.1%}")
    c[4].metric("Gasto semanal promedio", dinero(df["Gasto_Total_Semanal"].mean()))


# ===========================================================================
# 2. GRÁFICAS
# ===========================================================================
def filtros(df):
    c = st.columns(3)
    carreras = c[0].multiselect("Carrera", sorted(df["Carrera"].unique()), placeholder="Todas")
    resid = c[1].multiselect("Residencia", sorted(df["Estatus_Residencia"].unique()), placeholder="Todas")
    sem = c[2].multiselect("Semestre", sorted(df["Semestre"].unique()), placeholder="Todos")
    if carreras:
        df = df[df["Carrera"].isin(carreras)]
    if resid:
        df = df[df["Estatus_Residencia"].isin(resid)]
    if sem:
        df = df[df["Semestre"].isin(sem)]
    return df


def pagina_graficas():
    st.title("📊 Gráficas y estadística descriptiva")
    if datos.empty:
        st.info("No hay datos todavía. Registre alumnos o active el **Modo demostración** en la barra lateral.")
        return
    df = filtros(datos)
    if df.empty:
        st.warning("Ningún alumno cumple los filtros.")
        return
    kpis(df)

    t1, t2, t3, t4, t5 = st.tabs([
        "🎯 Becas y prioridad", "📈 Distribuciones (Sturges)", "⚖️ Local vs Foráneo",
        "👥 Perfil de la muestra", "🧮 Estadística descriptiva",
    ])

    # ---------------- Becas y prioridad
    with t1:
        c1, c2 = st.columns(2)
        conteo = df["Nivel_Vulnerabilidad"].value_counts().reindex(ORDEN_PRIORIDAD, fill_value=0).reset_index()
        conteo.columns = ["Nivel", "Alumnos"]
        fig = px.pie(conteo, names="Nivel", values="Alumnos", hole=0.55, color="Nivel",
                     color_discrete_map=COLOR_PRIORIDAD, title="Nivel de vulnerabilidad financiera")
        fig.update_traces(textinfo="percent+label", textposition="outside",
                          marker_line_color="white", marker_line_width=2, sort=False)
        c1.plotly_chart(estilo(fig).update_layout(showlegend=False), width="stretch")

        cand = df[df["Requiere_Beca"] == "SÍ"]
        mod = cand["Beca_Sugerida"].value_counts().reindex(ORDEN_BECAS, fill_value=0).reset_index()
        mod.columns = ["Modalidad", "Candidatos"]
        fig = px.bar(mod, y="Modalidad", x="Candidatos", orientation="h", text="Candidatos",
                     title="Demanda por modalidad de beca (solo candidatos)")
        fig.update_traces(marker_color=SERIE[0], textposition="outside")
        fig.update_yaxes(categoryorder="array", categoryarray=ORDEN_BECAS[::-1], title=None)
        c2.plotly_chart(estilo(fig), width="stretch")

        c1, c2 = st.columns(2)
        fig = px.histogram(df, x="Puntaje_Total", nbins=20, color="Requiere_Beca",
                           color_discrete_map={"SÍ": SERIE[1], "NO": SERIE[0]},
                           category_orders={"Requiere_Beca": ["SÍ", "NO"]},
                           title="Distribución del puntaje de necesidad (0-100)",
                           labels={"Puntaje_Total": "Puntaje", "Requiere_Beca": "¿Candidato?"})
        fig.add_vline(x=params["puntaje_minimo"], line_dash="dash", line_color="#52514e",
                      annotation_text=f"Mínimo {params['puntaje_minimo']:.0f} pts")
        fig.update_traces(marker_line_color="white", marker_line_width=1)
        c1.plotly_chart(estilo(fig).update_layout(bargap=0.05, yaxis_title="Alumnos"), width="stretch")

        fig = px.scatter(df, x="Ingreso_Semanal", y="Gasto_Total_Semanal", color="Nivel_Vulnerabilidad",
                         color_discrete_map=COLOR_PRIORIDAD, category_orders={"Nivel_Vulnerabilidad": ORDEN_PRIORIDAD},
                         hover_data=["Clave_Alumno", "Carrera", "Procedencia"],
                         title="Ingreso vs gasto semanal (arriba de la línea = déficit)",
                         labels={"Ingreso_Semanal": "Ingreso semanal ($)", "Gasto_Total_Semanal": "Gasto total ($)",
                                 "Nivel_Vulnerabilidad": "Prioridad"})
        lim = max(df["Ingreso_Semanal"].max(), df["Gasto_Total_Semanal"].max()) * 1.05
        fig.add_trace(go.Scatter(x=[0, lim], y=[0, lim], mode="lines", name="Ingreso = Gasto",
                                 line=dict(color="#898781", dash="dot", width=2)))
        fig.update_traces(marker=dict(size=10, line=dict(color="white", width=1.5)), selector=dict(mode="markers"))
        c2.plotly_chart(estilo(fig), width="stretch")

        st.markdown("##### Puntaje promedio por componente del tabulador")
        comp = df.groupby("Estatus_Residencia")[["Pts_ICF", "Pts_Geografia", "Pts_Saldo", "Pts_Fuente"]].mean()
        comp = comp.rename(columns={"Pts_ICF": "Gastos fijos (ICF)", "Pts_Geografia": "Geografía",
                                    "Pts_Saldo": "Saldo", "Pts_Fuente": "Fuente de ingreso"})
        largo = comp.reset_index().melt(id_vars="Estatus_Residencia", var_name="Componente", value_name="Puntos")
        fig = px.bar(largo, x="Puntos", y="Estatus_Residencia", color="Componente", orientation="h",
                     text_auto=".1f", labels={"Estatus_Residencia": ""})
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        st.plotly_chart(estilo(fig, 260), width="stretch")

    # ---------------- Distribuciones
    with t2:
        variables = {
            "Gasto_Total_Semanal": "Gasto total semanal ($)", "Ingreso_Semanal": "Ingreso semanal ($)",
            "Saldo_Ahorro_Semanal": "Saldo / ahorro semanal ($)", "Gasto_Transporte": "Gasto en transporte ($)",
            "Gasto_Alimentacion": "Gasto en alimentación ($)", "Puntaje_Total": "Puntaje de necesidad",
        }
        var = st.selectbox("Variable", list(variables), format_func=variables.get)
        tabla = db.tabla_sturges(df[var])
        n = len(df)
        st.caption(f"n = {n} · k = 1 + 3.322·log₁₀({n}) = {1 + 3.322 * np.log10(n):.2f} → "
                   f"**{len(tabla)} clases** · amplitud = {tabla['Lím. sup.'].iloc[0] - tabla['Lím. inf.'].iloc[0]:,.0f}")
        c1, c2 = st.columns(2)
        fig = go.Figure()
        fig.add_bar(x=tabla["Marca de clase"], y=tabla["fi"], width=(tabla["Lím. sup."] - tabla["Lím. inf."]) * 0.97,
                    marker_color=SERIE[0], name="Frecuencia", text=tabla["fi"], textposition="outside",
                    customdata=tabla["Intervalo"], hovertemplate="%{customdata}<br>fi = %{y}<extra></extra>")
        fig.add_scatter(x=tabla["Marca de clase"], y=tabla["fi"], mode="lines+markers", name="Polígono de frecuencias",
                        line=dict(color=SERIE[1], width=2), marker=dict(size=8))
        fig.update_layout(title="Histograma y polígono de frecuencias", xaxis_title=variables[var], yaxis_title="fi")
        c1.plotly_chart(estilo(fig), width="stretch")

        fig = go.Figure()
        fig.add_scatter(x=[tabla["Lím. inf."].iloc[0], *tabla["Lím. sup."]], y=[0, *(tabla["Hi"] * 100)],
                        mode="lines+markers", line=dict(color=SERIE[2], width=2), marker=dict(size=8),
                        name="Ojiva", hovertemplate="≤ %{x:,.0f}: %{y:.1f}%<extra></extra>")
        fig.update_layout(title="Ojiva (frecuencia relativa acumulada)", xaxis_title=variables[var],
                          yaxis_title="% acumulado", yaxis_range=[0, 105])
        c2.plotly_chart(estilo(fig), width="stretch")

        st.markdown("##### Tabla de distribución de frecuencias (Regla de Sturges)")
        st.dataframe(
            tabla.drop(columns=["Lím. inf.", "Lím. sup."]).style.format(
                {"Marca de clase": "{:,.1f}", "hi": "{:.3f}", "Hi": "{:.1%}"}),
            hide_index=True, width="stretch",
        )

        c1, c2 = st.columns(2)
        fig = px.box(df, y=var, x="Estatus_Residencia", color="Estatus_Residencia", points="all",
                     color_discrete_map=COLOR_RESIDENCIA, title=f"Diagrama de caja: {variables[var]}",
                     labels={var: variables[var], "Estatus_Residencia": ""})
        c1.plotly_chart(estilo(fig).update_layout(showlegend=False), width="stretch")
        fig = px.violin(df, y=var, x="Carrera", color="Carrera", box=True,
                        title=f"Distribución por carrera: {variables[var]}", labels={var: variables[var], "Carrera": ""})
        c2.plotly_chart(estilo(fig).update_layout(showlegend=False), width="stretch")

    # ---------------- Local vs foráneo
    with t3:
        rubros = {"Gasto_Transporte": "Transporte", "Gasto_Alimentacion": "Alimentación",
                  "Otros_Gastos_Acad": "Otros gastos", "Gasto_Total_Semanal": "Gasto total",
                  "Ingreso_Semanal": "Ingreso", "Saldo_Ahorro_Semanal": "Saldo / ahorro"}
        prom = df.groupby("Estatus_Residencia")[list(rubros)].mean().rename(columns=rubros)
        largo = prom.reset_index().melt(id_vars="Estatus_Residencia", var_name="Rubro", value_name="Promedio")
        fig = px.bar(largo, x="Rubro", y="Promedio", color="Estatus_Residencia", barmode="group",
                     color_discrete_map=COLOR_RESIDENCIA, text_auto="$,.0f",
                     title="Promedios semanales: Locales vs Foráneos", labels={"Promedio": "$ MXN", "Rubro": ""})
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        st.plotly_chart(estilo(fig, 420), width="stretch")

        c1, c2 = st.columns(2)
        estructura = df.groupby("Estatus_Residencia")[["Gasto_Transporte", "Gasto_Alimentacion", "Otros_Gastos_Acad"]].sum()
        estructura = estructura.div(estructura.sum(axis=1), axis=0) * 100
        estructura.columns = ["Transporte", "Alimentación", "Otros"]
        largo = estructura.reset_index().melt(id_vars="Estatus_Residencia", var_name="Rubro", value_name="%")
        fig = px.bar(largo, x="Estatus_Residencia", y="%", color="Rubro", text_auto=".1f",
                     title="Estructura del gasto (% del gasto total)", labels={"Estatus_Residencia": ""})
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        c1.plotly_chart(estilo(fig), width="stretch")

        tasa = df.groupby("Estatus_Residencia")["Requiere_Beca"].apply(lambda s: (s == "SÍ").mean() * 100).reset_index()
        tasa.columns = ["Residencia", "% candidatos"]
        fig = px.bar(tasa, x="Residencia", y="% candidatos", color="Residencia", text_auto=".1f",
                     color_discrete_map=COLOR_RESIDENCIA, title="Tasa de candidatos a beca por grupo")
        c2.plotly_chart(estilo(fig).update_layout(showlegend=False), width="stretch")

        st.markdown("##### Tabla de brechas")
        brecha = prom.T
        if {"Local", "Foráneo"} <= set(brecha.columns):
            brecha["Diferencia"] = brecha["Foráneo"] - brecha["Local"]
            brecha["Brecha %"] = brecha["Diferencia"] / brecha["Local"].replace(0, np.nan)
        st.dataframe(brecha.style.format("${:,.0f}").format({"Brecha %": "{:+.1%}"}, subset=[c for c in ["Brecha %"] if c in brecha]),
                     width="stretch")

    # ---------------- Perfil
    with t4:
        c1, c2 = st.columns(2)
        car = df.groupby("Carrera").agg(Encuestados=("Clave_Alumno", "count"),
                                        Candidatos=("Requiere_Beca", lambda s: (s == "SÍ").sum())).reset_index()
        largo = car.melt(id_vars="Carrera", var_name="Grupo", value_name="Alumnos")
        fig = px.bar(largo, x="Carrera", y="Alumnos", color="Grupo", barmode="group", text_auto=True,
                     title="Alumnos encuestados y candidatos por carrera", labels={"Carrera": ""})
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        c1.plotly_chart(estilo(fig), width="stretch")

        proc = df["Procedencia"].value_counts().reset_index()
        proc.columns = ["Municipio", "Alumnos"]
        fig = px.bar(proc, y="Municipio", x="Alumnos", orientation="h", text="Alumnos",
                     title="Alumnos por municipio de procedencia")
        fig.update_traces(marker_color=SERIE[0], textposition="outside")
        fig.update_yaxes(categoryorder="total ascending", title=None)
        c2.plotly_chart(estilo(fig), width="stretch")

        c1, c2 = st.columns(2)
        fue = df.groupby(["Fuente_Ingreso", "Nivel_Vulnerabilidad"]).size().reset_index(name="Alumnos")
        fig = px.bar(fue, y="Fuente_Ingreso", x="Alumnos", color="Nivel_Vulnerabilidad", orientation="h",
                     color_discrete_map=COLOR_PRIORIDAD, category_orders={"Nivel_Vulnerabilidad": ORDEN_PRIORIDAD},
                     title="Fuente de ingresos y nivel de prioridad",
                     labels={"Fuente_Ingreso": "", "Nivel_Vulnerabilidad": "Prioridad"})
        fig.update_traces(marker_line_color="white", marker_line_width=2)
        fig.update_yaxes(categoryorder="total ascending")
        c1.plotly_chart(estilo(fig), width="stretch")

        sem = df.groupby("Semestre").agg(Alumnos=("Clave_Alumno", "count"),
                                         Gasto=("Gasto_Total_Semanal", "mean")).reset_index()
        fig = px.bar(sem, x="Semestre", y="Alumnos", text="Alumnos", title="Alumnos por semestre",
                     hover_data={"Gasto": ":$,.0f"})
        fig.update_traces(marker_color=SERIE[0], textposition="outside")
        fig.update_xaxes(dtick=1)
        c2.plotly_chart(estilo(fig), width="stretch")

    # ---------------- Estadística
    with t5:
        columnas = {
            "Ingreso_Semanal": "Ingreso semanal", "Gasto_Transporte": "Gasto transporte",
            "Gasto_Alimentacion": "Gasto alimentación", "Otros_Gastos_Acad": "Otros gastos",
            "Gasto_Total_Semanal": "Gasto total semanal", "Saldo_Ahorro_Semanal": "Saldo / ahorro",
            "Indice_Carga_Fija_%": "ICF", "Puntaje_Total": "Puntaje de beca",
        }
        st.markdown("##### Medidas de tendencia central, posición y dispersión")
        resumen = db.resumen_descriptivo(df, columnas)
        st.dataframe(resumen.style.format("{:,.2f}").format("{:,.0f}", subset=pd.IndexSlice["n", :]),
                     width="stretch")
        lectura = {
            col: ("Sesgo a la derecha" if r["Media"] > r["Mediana"] * 1.05 else
                  "Sesgo a la izquierda" if r["Media"] < r["Mediana"] * 0.95 else "Aprox. simétrica")
            for col, r in resumen.items()
        }
        st.caption("Forma de la distribución (media vs mediana): " +
                   " · ".join(f"**{k}**: {v}" for k, v in lectura.items()))

        st.markdown("##### Correlación entre variables (Pearson)")
        corr = df[list(columnas)].rename(columns=columnas).corr()
        fig = px.imshow(corr, text_auto=".2f", zmin=-1, zmax=1, aspect="auto",
                        color_continuous_scale=["#2a78d6", "#f0efec", "#e34948"])
        st.plotly_chart(estilo(fig, 520), width="stretch")


# ===========================================================================
# 3. PADRÓN
# ===========================================================================
def pagina_padron():
    st.title("🏅 Padrón de alumnos con necesidad de beca")
    st.caption(f"Ordenado de mayor a menor puntaje · Puntaje mínimo vigente: {params['puntaje_minimo']:.0f} pts "
               "(se cambia en la hoja Parametros del Excel)")
    if datos.empty:
        st.info("No hay alumnos registrados.")
        return
    padron = datos[datos["Requiere_Beca"] == "SÍ"].sort_values("Puntaje_Total", ascending=False)
    st.metric("Candidatos detectados", len(padron))
    vista = padron[["Clave_Alumno", "Carrera", "Semestre", "Estatus_Residencia", "Procedencia", "Fuente_Ingreso",
                    "Ingreso_Semanal", "Saldo_Ahorro_Semanal", "Indice_Carga_Fija_%", "Puntaje_Total",
                    "Nivel_Vulnerabilidad", "Beca_Sugerida"]].reset_index(drop=True)
    vista.index += 1
    st.dataframe(
        vista, width="stretch", height=560,
        column_config={
            "Ingreso_Semanal": st.column_config.NumberColumn("Ingreso", format="$%.0f"),
            "Saldo_Ahorro_Semanal": st.column_config.NumberColumn("Saldo", format="$%.0f"),
            "Indice_Carga_Fija_%": st.column_config.NumberColumn("ICF", format="percent"),
            "Puntaje_Total": st.column_config.ProgressColumn("Puntaje", min_value=0, max_value=100, format="%.1f"),
        },
    )


# ===========================================================================
# 4. BASE DE DATOS
# ===========================================================================
def pagina_base():
    st.title("🗂️ Base de datos (hoja Base_Datos)")
    if datos.empty:
        st.info("No hay alumnos registrados.")
        return
    st.dataframe(datos, width="stretch", height=480, hide_index=True)
    st.download_button("⬇️ Descargar como CSV", datos.to_csv(index=False).encode("utf-8-sig"),
                       "base_datos_becas.csv", "text/csv")
    if EN_LINEA:
        st.caption("Para corregir o borrar un registro en línea, edítelo en la Hoja de Google (pestaña Registros).")
        return
    with st.expander("🗑️ Eliminar un registro (por error de captura)"):
        clave = st.selectbox("Clave del alumno", datos["Clave_Alumno"])
        if st.button("Eliminar definitivamente", type="secondary"):
            try:
                with st.spinner("Eliminando en Excel…"):
                    db.eliminar_registro(clave, RUTA)
                st.success(f"Registro {clave} eliminado.")
                st.rerun()
            except db.ErrorRegistro as e:
                st.error(str(e))


def acceso_permitido() -> bool:
    """En línea, las páginas con datos de los alumnos piden la contraseña del equipo."""
    clave = (NUBE or {}).get("admin_password")
    if not EN_LINEA or not clave or modo_demo or st.session_state.get("admin_ok"):
        return True
    st.title("🔒 Acceso del equipo")
    st.write("Esta sección muestra datos de los alumnos. Escriba la contraseña del equipo.")
    intento = st.text_input("Contraseña", type="password")
    if intento:
        if intento == clave:
            st.session_state["admin_ok"] = True
            st.rerun()
        st.error("Contraseña incorrecta.")
    return False


paginas = {
    "📝 Registro de alumnos": pagina_registro,
    "📊 Gráficas y estadística": pagina_graficas,
    "🏅 Padrón de becas": pagina_padron,
    "🗂️ Base de datos": pagina_base,
}
if pagina == "📝 Registro de alumnos" or acceso_permitido():
    paginas[pagina]()
