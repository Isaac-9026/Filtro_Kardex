import streamlit as st
import pandas as pd
import datetime

pd.set_option("styler.render.max_elements", 5_000_000)

st.set_page_config(
    page_title="Kardex Viewer",
    page_icon="🌽",
    layout="wide"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #f5a623, #f0c040);
        padding: 1.2rem 2rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #e65100;
        margin-top: 1.2rem;
        margin-bottom: 0.5rem;
        border-bottom: 2px solid #ffe0b2;
        padding-bottom: 4px;
    }
    div[data-testid="metric-container"] {
        background: #fff8e1;
        border-left: 4px solid #f5a623;
        border-radius: 8px;
        padding: 0.6rem 1rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h2 style='margin:0;'>🌽 Kardex Viewer — Inventario</h2>
    <p style='margin:4px 0 0 0; font-size:0.9rem; opacity:0.85;'>Visualizador de movimientos · Soporta múltiples archivos</p>
</div>
""", unsafe_allow_html=True)


# ── Parser ───────────────────────────────────────────────────────────────────
@st.cache_data
def load_kardex(file_bytes):
    df = pd.read_excel(file_bytes, header=1)
    df.columns = [
        "Codigo", "Descripcion",
        "Fecha", "Tipo", "Serie", "Numero", "Tipo_Operacion",
        "Ent_Cantidad", "Ent_Costo_Unit", "Ent_Costo_Total",
        "Sal_Cantidad", "Sal_Costo_Unit", "Sal_Costo_Total",
        "Saldo_Cantidad", "Saldo_Costo_Unit", "Saldo_Costo_Total"
    ]
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    for col in ["Ent_Cantidad","Ent_Costo_Unit","Ent_Costo_Total",
                "Sal_Cantidad","Sal_Costo_Unit","Sal_Costo_Total",
                "Saldo_Cantidad","Saldo_Costo_Unit","Saldo_Costo_Total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).round(10)
    df["Codigo"]      = df["Codigo"].ffill().astype(str).str.strip()
    df["Descripcion"] = df["Descripcion"].ffill().astype(str).str.strip()
    return df


def render_metricas(dff):
    saldo_rows        = dff[dff["Saldo_Cantidad"] > 0]
    saldo_final_cant  = saldo_rows["Saldo_Cantidad"].iloc[-1]    if len(saldo_rows) else 0
    saldo_final_valor = saldo_rows["Saldo_Costo_Total"].iloc[-1] if len(saldo_rows) else 0
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("📥 Ent. Cantidad",  f"{dff['Ent_Cantidad'].sum():,.3f} kg")
    m2.metric("📥 Ent. Valor",     f"S/ {dff['Ent_Costo_Total'].sum():,.2f}")
    m3.metric("📤 Sal. Cantidad",  f"{dff['Sal_Cantidad'].sum():,.3f} kg")
    m4.metric("📤 Sal. Valor",     f"S/ {dff['Sal_Costo_Total'].sum():,.2f}")
    m5.metric("📦 Saldo Final Kg", f"{saldo_final_cant:,.3f} kg")
    m6.metric("💰 Saldo Final S/", f"S/ {saldo_final_valor:,.10f}")


def render_tabla(dff):
    display = dff.copy()
    display["Fecha"] = display["Fecha"].dt.strftime("%d/%m/%Y").fillna("")
    display = display.rename(columns={
        "Codigo":"Código", "Descripcion":"Descripción",
        "Fecha":"Fecha", "Tipo":"Tipo", "Serie":"Serie",
        "Numero":"Número", "Tipo_Operacion":"Operación",
        "Ent_Cantidad":"Ent. Cant.", "Ent_Costo_Unit":"Ent. C.Unit",
        "Ent_Costo_Total":"Ent. C.Total", "Sal_Cantidad":"Sal. Cant.",
        "Sal_Costo_Unit":"Sal. C.Unit", "Sal_Costo_Total":"Sal. C.Total",
        "Saldo_Cantidad":"Saldo Cant.", "Saldo_Costo_Unit":"Saldo C.Unit",
        "Saldo_Costo_Total":"Saldo C.Total",
    })
    fmt = {
        "Ent. Cant.":"{:,.3f}", "Ent. C.Unit":"{:,.5f}", "Ent. C.Total":"{:,.10f}",
        "Sal. Cant.":"{:,.3f}", "Sal. C.Unit":"{:,.5f}", "Sal. C.Total":"{:,.10f}",
        "Saldo Cant.":"{:,.3f}", "Saldo C.Unit":"{:,.5f}", "Saldo C.Total":"{:,.10f}",
    }
    st.dataframe(display.style.format(fmt), use_container_width=True, height=600)


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

# ── Cargar y unir todos los archivos ─────────────────────────────────────────
frames = {f.name: load_kardex(f.read()) for f in uploaded_files}
df_all = pd.concat(frames.values(), ignore_index=True)

# ── Filtro de fecha ───────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📅 Filtro por Fecha</div>', unsafe_allow_html=True)

fechas_validas = df_all["Fecha"].dropna()
años_disponibles  = sorted(fechas_validas.dt.year.unique().tolist())
meses_nombres = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}

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

#Aplicar filtro
if año_sel == "Todos":
    pass  # sin filtro
elif mes_sel == "Todos":
    df_all = df_all[(df_all["Fecha"].isna()) | (df_all["Fecha"].dt.year == año_sel)]
else:
    df_all = df_all[
        (df_all["Fecha"].isna()) |
        ((df_all["Fecha"].dt.year == año_sel) & (df_all["Fecha"].dt.month == mes_sel))
    ]

# ── Una sola tabla unificada ──────────────────────────────────────────────────
productos = df_all[["Codigo","Descripcion"]].drop_duplicates()
badges = " ".join([
    f'<span style="display:inline-block;background:#fff3e0;border:1px solid #f5a623;border-radius:20px;padding:0.2rem 0.9rem;font-size:0.85rem;color:#e65100;font-weight:600;margin-right:0.4rem;">📦 {r.Codigo} — {r.Descripcion}</span>'
    for r in productos.itertuples()
])
st.markdown(f'<div style="margin-bottom:1rem;">{badges}</div>', unsafe_allow_html=True)

st.markdown(
    f'<div class="section-title">📋 Todos los Movimientos '
    f'<span style="font-weight:400;font-size:0.85rem;color:#888;">({len(df_all)} registros)</span></div>',
    unsafe_allow_html=True
)
render_metricas(df_all)
render_tabla(df_all)

# ── Descargar el Excel filtrado ──────────────────────────────────────────────────
st.markdown('<div class="section-title">💾 Descargar datos filtrados</div>', unsafe_allow_html=True)

def exportar_excel(df):
    export = df.copy()

    # Restaurar estructura original: fila 0 = headers grupo, fila 1 = subheaders
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Kardex"

    # ── Fila 1: headers de grupo ──────────────────────────────────────────────
    headers_grupo = [
        ("Código",            1, 1),
        ("Descripción",       2, 2),
        ("COMPROBANTE",       3, 6),
        ("Tipo de Operación", 7, 7),
        ("ENTRADAS",          8, 10),
        ("SALIDAS",           11, 13),
        ("SALDO FINAL",       14, 16),
    ]

    sin_fill  = PatternFill(fill_type=None)
    bold      = Font(bold=True)
    center    = Alignment(horizontal="center", vertical="center")
    thin      = Side(style="thin")
    borde     = Border(left=thin, right=thin, top=thin, bottom=thin)

    for (titulo, col_ini, col_fin), in [(h,) for h in headers_grupo]:
        cell = ws.cell(row=1, column=col_ini, value=titulo)
        cell.font      = bold
        cell.alignment = center
        cell.fill      = sin_fill
        cell.border    = borde
        if col_ini != col_fin:
            # Fusionar horizontalmente (COMPROBANTE, ENTRADAS, etc.)
            ws.merge_cells(start_row=1, start_column=col_ini,
                           end_row=1,   end_column=col_fin)
        else:
            # Código (col1), Descripción (col2), Tipo de Operación (col7)
            # → fusionar fila 1 y fila 2 verticalmente
            ws.merge_cells(start_row=1, start_column=col_ini,
                           end_row=2,   end_column=col_fin)
            # Aplicar borde también a la celda de fila 2 fusionada
            ws.cell(row=2, column=col_ini).border = borde

    # ── Fila 2: subheaders ────────────────────────────────────────────────────
    subheaders = [
        "Código", "Descripción",
        "Fecha", "Tipo", "Serie", "Número", "Tipo de Operación",
        "Cantidad", "Costo Unitario", "Costo Total",
        "Cantidad", "Costo Unitario", "Costo Total",
        "Cantidad", "Costo Unitario", "Costo Total",
    ]

    for col, sh in enumerate(subheaders, start=1):
        if col in (1, 2, 7):  # ya fusionadas verticalmente en fila 1+2
            continue
        cell = ws.cell(row=2, column=col, value=sh)
        cell.font      = bold
        cell.alignment = center
        cell.fill      = sin_fill
        cell.border    = borde

    # Columnas que van centradas en datos: 3=Fecha,4=Tipo,5=Serie,6=Número,7=TipoOp
    cols_centradas = {3, 4, 5, 6, 7}

    # ── Filas de datos ────────────────────────────────────────────────────────
    for row_idx, row in enumerate(export.itertuples(index=False), start=3):
        datos = [
            (str(row.Codigo).zfill(6), "@"),
            (row.Descripcion,          "@"),
            (row.Fecha.strftime("%d/%m/%Y") if pd.notna(row.Fecha) else "", "@"),
            (row.Tipo,                 "00"),
            (str(row.Serie),           "@"),
            (int(row.Numero) if str(row.Numero).strip() not in ("", "nan") else 0, r'[$-408]00000000'),
            (row.Tipo_Operacion,       "@"),
            (row.Ent_Cantidad,         "#,##0.000"),
            (row.Ent_Costo_Unit,       "#,##0.0000"),
            (row.Ent_Costo_Total,      "#,##0.000"),
            (row.Sal_Cantidad,         "#,##0.000"),
            (row.Sal_Costo_Unit,       "#,##0.0000"),
            (row.Sal_Costo_Total,      "#,##0.000"),
            (row.Saldo_Cantidad,       "#,##0.000"),
            (row.Saldo_Costo_Unit,     "#,##0.0000"),
            (row.Saldo_Costo_Total,    "#,##0.000"),
        ]
        for col, (val, fmt_num) in enumerate(datos, start=1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            if col == 1:
                cell.data_type = "s"
            cell.number_format = fmt_num
            cell.border        = borde
            cell.alignment     = Alignment(
                horizontal="center" if col in cols_centradas else "general",
                vertical="center"
            )

    # ── Anchos de columna ─────────────────────────────────────────────────────
    anchos = [10, 22, 12, 6, 8, 14, 18, 12, 14, 16, 12, 14, 16, 12, 14, 16]
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 20

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

buffer = exportar_excel(df_all)

st.markdown('<div class="section-title">💾 Descargar datos</div>', unsafe_allow_html=True)

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