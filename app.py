import streamlit as st
import pandas as pd
import datetime
import io

pd.set_option("styler.render.max_elements", 5_000_000)

st.set_page_config(
    page_title="Kardex Viewer",
    page_icon="assets/icono.ico",
    layout="wide"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e24a8, #343552);
        padding: 1.2rem 2rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #3439c9;
        margin-top: 1.2rem;
        margin-bottom: 0.5rem;
        border-bottom: 2px solid #3439c9;
        padding-bottom: 4px;
    }
    div[data-testid="metric-container"] {
        background: #fff8e1;
        border-left: 4px solid #f5a623;
        border-radius: 8px;
        padding: 0.6rem 1rem;
    }
    .error-box {
        background: #fff3f3;
        border-left: 4px solid #e53935;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        color: #b71c1c;
        margin-bottom: 1rem;
    }
    .alert-box {
        background: #fff8e1;
        border-left: 4px solid #f5a623;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        color: #e65100;
        margin-bottom: 1rem;
    }
    .ok-box {
        background: #f1f8e9;
        border-left: 4px solid #43a047;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        color: #2e7d32;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h2 style='margin:0;'>📊 Kardex Viewer — Inventario</h2>
    <p style='margin:4px 0 0 0; font-size:0.9rem; opacity:0.85;'>Visualizador de movimientos · Soporta múltiples archivos</p>
</div>
""", unsafe_allow_html=True)


# ── Parser ───────────────────────────────────────────────────────────────────
COLUMNAS_REQUERIDAS = 15

@st.cache_data
def load_kardex(file_bytes, filename):
    try:
        df = pd.read_excel(file_bytes, header=1, dtype={0: str})

        if df.shape[1] < COLUMNAS_REQUERIDAS:
            return None, f"❌ '{filename}' no tiene el formato esperado ({df.shape[1]} columnas encontradas, se esperan {COLUMNAS_REQUERIDAS})."

        df.columns = [
            "Codigo",
            "Fecha", "Tipo", "Serie", "Numero", "Tipo_Operacion",
            "Ent_Cantidad", "Ent_Costo_Unit", "Ent_Costo_Total",
            "Sal_Cantidad", "Sal_Costo_Unit", "Sal_Costo_Total",
            "Saldo_Cantidad", "Saldo_Costo_Unit", "Saldo_Costo_Total"
        ] + list(df.columns[COLUMNAS_REQUERIDAS:])

        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")

        cols_numericas = [
            "Ent_Cantidad","Ent_Costo_Unit","Ent_Costo_Total",
            "Sal_Cantidad","Sal_Costo_Unit","Sal_Costo_Total",
            "Saldo_Cantidad","Saldo_Costo_Unit","Saldo_Costo_Total"
        ]
        for col in cols_numericas:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round(10)

        df["Codigo"] = df["Codigo"].ffill().astype(str).str.strip()
        df = df[df["Fecha"].notna()].reset_index(drop=True)

        if len(df) == 0:
            return None, f"❌ '{filename}' no contiene fechas válidas."

        df["Tipo"] = pd.to_numeric(df["Tipo"], errors="coerce").fillna(0).astype(int)

        return df, None

    except Exception as e:
        return None, f"❌ Error al leer '{filename}': {str(e)}"


# ── Verificación de Saldo Costo Total ────────────────────────────────────────
def verificar_saldo_costo_total(df, tolerancia=0.01):
    """
    Verifica fila por fila comparando con la fila anterior:
        Saldo[i] == Saldo[i-1] + Entrada[i] - Salida[i]
    Un error en una fila no contamina las siguientes.
    La fila de 'Saldo Anterior' se acepta como válida y sirve de base.
    """
    df = df.copy()
    df["Saldo_Calculado"] = 0.0
    df["Diferencia"]      = 0.0
    df["Alterado"]        = False

    for codigo in df["Codigo"].unique():
        mask    = df["Codigo"] == codigo
        indices = df[mask].index
        saldo_anterior = None

        for idx in indices:
            op = str(df.at[idx, "Tipo_Operacion"]).strip().lower()
            saldo_excel = df.at[idx, "Saldo_Costo_Total"]

            # Saldo Anterior: aceptar directamente como base
            if "saldo anterior" in op or saldo_anterior is None:
                df.at[idx, "Saldo_Calculado"] = saldo_excel
                df.at[idx, "Diferencia"]      = 0.0
                df.at[idx, "Alterado"]        = False
                saldo_anterior = saldo_excel
                continue

            # Calcular lo esperado usando el saldo de la fila anterior del Excel
            esperado = round(
                saldo_anterior + df.at[idx, "Ent_Costo_Total"] - df.at[idx, "Sal_Costo_Total"], 10
            )
            diff = round(abs(esperado - saldo_excel), 10)

            df.at[idx, "Saldo_Calculado"] = esperado
            df.at[idx, "Diferencia"]      = diff
            df.at[idx, "Alterado"]        = diff > tolerancia

            # Siempre usar el valor del Excel como base para la siguiente fila
            saldo_anterior = saldo_excel

    return df


def render_metricas(dff):
    saldo_final_cant  = dff["Saldo_Cantidad"].iloc[-1]    if len(dff) else 0
    saldo_final_valor = dff["Saldo_Costo_Total"].iloc[-1] if len(dff) else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Cantidad ENTRADA total", f"{dff['Ent_Cantidad'].sum():,.3f} kg")
    m2.metric("Costo ENTRADA total",    f"S/ {dff['Ent_Costo_Total'].sum():,.2f}")
    m3.metric("Cantidad SALIDA total",  f"{dff['Sal_Cantidad'].sum():,.3f} kg")

    m4, m5, m6 = st.columns(3)
    m4.metric("Costo SALIDA total",   f"S/ {dff['Sal_Costo_Total'].sum():,.2f}")
    m5.metric("Cantidad SALDO total", f"{saldo_final_cant:,.3f} kg")
    m6.metric("Costo SALDO total",    f"S/ {saldo_final_valor:,.2f}")


def render_tabla(dff, mostrar_verificacion=False):
    display = dff.copy()
    display["Fecha"] = display["Fecha"].dt.strftime("%d/%m/%Y").fillna("")

    cols_rename = {
        "Codigo":"Código",
        "Fecha":"Fecha", "Tipo":"Tipo", "Serie":"Serie",
        "Numero":"Número", "Tipo_Operacion":"Operación",
        "Ent_Cantidad":"Ent. Cant.", "Ent_Costo_Unit":"Ent. C.Unit",
        "Ent_Costo_Total":"Ent. C.Total", "Sal_Cantidad":"Sal. Cant.",
        "Sal_Costo_Unit":"Sal. C.Unit", "Sal_Costo_Total":"Sal. C.Total",
        "Saldo_Cantidad":"Saldo Cant.", "Saldo_Costo_Unit":"Saldo C.Unit",
        "Saldo_Costo_Total":"Saldo C.Total",
    }

    fmt = {
        "Ent. Cant.":"{:,.3f}", "Ent. C.Unit":"{:,.5f}", "Ent. C.Total":"{:,.2f}",
        "Sal. Cant.":"{:,.3f}", "Sal. C.Unit":"{:,.5f}", "Sal. C.Total":"{:,.2f}",
        "Saldo Cant.":"{:,.3f}", "Saldo C.Unit":"{:,.5f}", "Saldo C.Total":"{:,.2f}",
    }

    if mostrar_verificacion and "Alterado" in display.columns:
        # Añadir columnas de verificación visibles
        display["Saldo Calculado"] = display["Saldo_Calculado"]
        display["Diferencia"]      = display["Diferencia"]
        alterado_mask              = display["Alterado"].values
        display = display.drop(columns=["Saldo_Calculado", "Alterado"])
        display = display.rename(columns=cols_rename)
        fmt["Saldo Calculado"] = "{:,.2f}"
        fmt["Diferencia"]      = "{:,.4f}"

        def highlight_alterado(row):
            idx = display.index.get_loc(row.name)
            if alterado_mask[idx]:
                return ["background-color: #ffebee; color: #b71c1c"] * len(row)
            return [""] * len(row)

        st.dataframe(
            display.style.format(fmt).apply(highlight_alterado, axis=1),
            use_container_width=True, height=600
        )
    else:
        display = display.drop(columns=["Saldo_Calculado", "Diferencia", "Alterado"],
                               errors="ignore")
        display = display.rename(columns=cols_rename)
        st.dataframe(display.style.format(fmt), use_container_width=True, height=600)


def exportar_excel(df):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Kardex"

    headers_grupo = [
        ("Código",            1, 1),
        ("COMPROBANTE",       2, 5),
        ("Tipo de Operación", 6, 6),
        ("ENTRADAS",          7, 9),
        ("SALIDAS",           10, 12),
        ("SALDO FINAL",       13, 15),
    ]

    sin_fill = PatternFill(fill_type=None)
    bold     = Font(bold=True)
    center   = Alignment(horizontal="center", vertical="center")
    thin     = Side(style="thin")
    borde    = Border(left=thin, right=thin, top=thin, bottom=thin)

    for (titulo, col_ini, col_fin) in headers_grupo:
        cell = ws.cell(row=1, column=col_ini, value=titulo)
        cell.font      = bold
        cell.alignment = center
        cell.fill      = sin_fill
        cell.border    = borde
        if col_ini != col_fin:
            ws.merge_cells(start_row=1, start_column=col_ini, end_row=1, end_column=col_fin)
        else:
            ws.merge_cells(start_row=1, start_column=col_ini, end_row=2, end_column=col_fin)
            ws.cell(row=2, column=col_ini).border = borde

    subheaders = [
        "Código",
        "Fecha", "Tipo", "Serie", "Número", "Tipo de Operación",
        "Cantidad", "Costo Unitario", "Costo Total",
        "Cantidad", "Costo Unitario", "Costo Total",
        "Cantidad", "Costo Unitario", "Costo Total",
    ]

    for col, sh in enumerate(subheaders, start=1):
        if col in (1, 6):
            continue
        cell = ws.cell(row=2, column=col, value=sh)
        cell.font      = bold
        cell.alignment = center
        cell.fill      = sin_fill
        cell.border    = borde

    cols_centradas = {2, 3, 4, 5, 6}

    for row_idx, row in enumerate(df.itertuples(index=False), start=3):
        datos = [
            (str(row.Codigo), "@"),
            (row.Fecha.strftime("%d/%m/%Y") if pd.notna(row.Fecha) else "", "@"),
            (row.Tipo, "00"),
            (str(row.Serie), "@"),
            (int(row.Numero) if str(row.Numero).strip() not in ("", "nan") else 0, r'[$-408]00000000'),
            (row.Tipo_Operacion,    "@"),
            (row.Ent_Cantidad,      "#,##0.000"),
            (row.Ent_Costo_Unit,    "#,##0.0000"),
            (row.Ent_Costo_Total,   "#,##0.000"),
            (row.Sal_Cantidad,      "#,##0.000"),
            (row.Sal_Costo_Unit,    "#,##0.0000"),
            (row.Sal_Costo_Total,   "#,##0.000"),
            (row.Saldo_Cantidad,    "#,##0.000"),
            (row.Saldo_Costo_Unit,  "#,##0.0000"),
            (row.Saldo_Costo_Total, "#,##0.000"),
        ]
        for col, (val, fmt_num) in enumerate(datos, start=1):
            cell = ws.cell(row=row_idx, column=col)
            if col == 1:
                cell.value        = str(val)
                cell.number_format = "@"
                cell.quotePrefix  = True
            else:
                cell.value         = val
                cell.number_format = fmt_num
            cell.border    = borde
            cell.alignment = Alignment(
                horizontal="center" if col in cols_centradas else "general",
                vertical="center"
            )

    anchos = [10, 12, 6, 8, 14, 18, 12, 14, 16, 12, 14, 16, 12, 14, 16]
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 20

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Archivos Excel")
    uploaded_files = st.file_uploader(
        "Sube uno o más archivos (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
    )
    if uploaded_files:
        st.markdown("---")
        st.markdown("**Archivos cargados:**")
        for f in uploaded_files:
            st.markdown(f"✅ `{f.name}`")

if not uploaded_files:
    st.info("👈 Carga uno o más archivos Excel desde el panel izquierdo para comenzar.")
    st.stop()

# ── Cargar y validar archivos ─────────────────────────────────────────────────
frames = {}
errores = []

for f in uploaded_files:
    df_temp, error = load_kardex(f.read(), f.name)
    if error:
        errores.append(error)
    else:
        frames[f.name] = df_temp

if errores:
    for err in errores:
        st.markdown(f'<div class="error-box">{err}</div>', unsafe_allow_html=True)

if not frames:
    st.warning("No se pudo cargar ningún archivo válido. Verifica el formato de tus archivos.")
    st.stop()

df_all = pd.concat(frames.values(), ignore_index=True)
df_all = df_all.sort_values(["Codigo", "Fecha"], na_position="first").reset_index(drop=True)

# ── Verificación: se corre por archivo individual para respetar el orden original
df_verificados = []
for nombre, df_ind in frames.items():
    df_ind["_orden_op"] = df_ind["Tipo_Operacion"].apply(lambda x: 0 if "compra" in str(x).lower() else 1)
    df_ind = df_ind.sort_values(["Fecha", "_orden_op"], na_position="first").reset_index(drop=True)
    df_ind = df_ind.drop(columns=["_orden_op"])
    df_ind = verificar_saldo_costo_total(df_ind)
    df_verificados.append(df_ind)

df_all = pd.concat(df_verificados, ignore_index=True)

# Ordenar: por Codigo, Fecha, y dentro del mismo día compras antes que ventas
df_all["_orden_op"] = df_all["Tipo_Operacion"].apply(lambda x: 0 if "compra" in str(x).lower() else 1)
df_all = df_all.sort_values(["Codigo", "Fecha", "_orden_op"], na_position="first").reset_index(drop=True)
df_all = df_all.drop(columns=["_orden_op"])

# ── Filtro por código ─────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📦 Filtro por Código</div>', unsafe_allow_html=True)

codigos_disponibles = sorted(df_all["Codigo"].dropna().unique().tolist())

col_id1, col_id2 = st.columns([3, 1])
with col_id1:
    id_buscado = st.text_input(
        "Ingresa el código del producto",
        placeholder="Ej: 021007",
        help="Escribe el código exacto para filtrar"
    )
with col_id2:
    st.markdown("<br>", unsafe_allow_html=True)
    st.button("🔍 Buscar", use_container_width=True)

if id_buscado.strip():
    buscado_norm = id_buscado.strip().zfill(6)
    coincidencias = [c for c in codigos_disponibles if c.zfill(6) == buscado_norm]
    if coincidencias:
        df_all = df_all[df_all["Codigo"].isin(coincidencias)]
    else:
        st.warning(f"⚠️ No se encontró ningún producto con el código '{id_buscado.strip()}'.")

# ── Filtro de fecha ───────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📅 Filtro por Fecha</div>', unsafe_allow_html=True)

fechas_validas = df_all["Fecha"].dropna()
f_min = fechas_validas.min().date()
f_max = fechas_validas.max().date()
años_disponibles = sorted(fechas_validas.dt.year.unique().tolist())
meses_nombres = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}

tipo_filtro = st.radio(
    "Modo de filtro",
    ["Por Año / Mes", "Por fecha exacta", "Por rango de fechas"],
    horizontal=True
)

if tipo_filtro == "Por Año / Mes":
    col_a, col_m = st.columns(2)
    with col_a:
        año_sel = st.selectbox("Año", ["Todos"] + años_disponibles)
    with col_m:
        if año_sel == "Todos":
            st.selectbox("Mes", ["Todos"], disabled=True)
            mes_sel = "Todos"
        else:
            meses_en_año = sorted(
                fechas_validas[fechas_validas.dt.year == año_sel].dt.month.unique().tolist()
            )
            opciones_mes = ["Todos"] + [meses_nombres[m] for m in meses_en_año]
            mes_label = st.selectbox("Mes", opciones_mes)
            mes_sel = next((k for k, v in meses_nombres.items() if v == mes_label), "Todos")

    if año_sel == "Todos":
        pass
    elif mes_sel == "Todos":
        df_all = df_all[(df_all["Fecha"].isna()) | (df_all["Fecha"].dt.year == año_sel)]
    else:
        df_all = df_all[
            (df_all["Fecha"].isna()) |
            ((df_all["Fecha"].dt.year == año_sel) & (df_all["Fecha"].dt.month == mes_sel))
        ]

elif tipo_filtro == "Por fecha exacta":
    fecha_exacta = st.date_input("Selecciona una fecha", value=f_max, min_value=f_min, max_value=f_max)
    df_all = df_all[
        (df_all["Fecha"].isna()) |
        (df_all["Fecha"].dt.date == fecha_exacta)
    ]

elif tipo_filtro == "Por rango de fechas":
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        fecha_desde = st.date_input("Desde", value=f_min, min_value=f_min, max_value=f_max)
    with col_r2:
        fecha_hasta = st.date_input("Hasta", value=f_max, min_value=f_min, max_value=f_max)
    if fecha_desde <= fecha_hasta:
        df_all = df_all[
            (df_all["Fecha"].isna()) |
            ((df_all["Fecha"].dt.date >= fecha_desde) & (df_all["Fecha"].dt.date <= fecha_hasta))
        ]
    else:
        st.warning("⚠️ La fecha de inicio no puede ser mayor a la fecha final.")

# ── Badges ────────────────────────────────────────────────────────────────────
codigos_visibles = df_all["Codigo"].drop_duplicates().tolist()
badges = " ".join([
    f'<span style="display:inline-block;background:#e4e6f2;border:1px solid #1e24a8;border-radius:20px;padding:0.2rem 0.9rem;font-size:0.85rem;color:#1e24a8;font-weight:600;margin-right:0.4rem;">📦 {c}</span>'
    for c in codigos_visibles
])
st.markdown(f'<div style="margin-bottom:1rem;">{badges}</div>', unsafe_allow_html=True)

# ── Alerta de integridad ──────────────────────────────────────────────────────
n_alterados = df_all["Alterado"].sum()
if n_alterados > 0:
    st.markdown(
        f'<div class="alert-box">⚠️ Se detectaron <strong>{n_alterados} fila(s)</strong> donde el '
        f'Costo Total del Saldo Final en el Excel <strong>no coincide</strong> con el valor recalculado por el programa. '
        f'Esto puede deberse a un error en el sistema de origen o a una posible alteración. '
        f'Activa <strong>"Mostrar verificación"</strong> para ver el detalle fila por fila.</div>',
        unsafe_allow_html=True
    )
else:
    st.markdown(
        '<div class="ok-box">✅ Todos los valores de Costo Total Saldo Final coinciden con el cálculo esperado.</div>',
        unsafe_allow_html=True
    )

# ── Métricas y tabla ──────────────────────────────────────────────────────────
st.markdown(
    f'<div class="section-title">📋 Movimientos '
    f'<span style="font-weight:400;font-size:0.85rem;color:#888;">({len(df_all)} registros)</span></div>',
    unsafe_allow_html=True
)
render_metricas(df_all)

mostrar_ver = st.toggle("🔍 Mostrar verificación de integridad", value=False)
render_tabla(df_all, mostrar_verificacion=mostrar_ver)

# ── Descargar Excel ───────────────────────────────────────────────────────────
st.markdown('<div class="section-title">💾 Descargar datos</div>', unsafe_allow_html=True)

# Exportar con Saldo_Costo_Total corregido donde haya diferencia
df_export = df_all.copy()
if "Saldo_Calculado" in df_export.columns:
    df_export["Saldo_Costo_Total"] = df_export["Saldo_Calculado"]
df_export = df_export.drop(columns=["Saldo_Calculado", "Diferencia", "Alterado"], errors="ignore")
buffer = exportar_excel(df_export)

nombre_personalizado = st.text_input(
    "Nombre del archivo (opcional)",
    placeholder=f"kardex_{datetime.date.today().strftime('%Y%m%d')}",
    help="Si lo dejas vacío se usará el nombre por defecto"
)

nombre_base = nombre_personalizado.strip() if nombre_personalizado.strip() else f"kardex_{datetime.date.today().strftime('%Y%m%d')}"
if not nombre_base.endswith(".xlsx"):
    nombre_base += ".xlsx"

st.download_button(
    label="⬇️ Descargar Excel",
    data=buffer,
    file_name=nombre_base,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)