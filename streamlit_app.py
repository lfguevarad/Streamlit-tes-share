"""
ALM · TES -- dashboard de operaciones y curva de tasas.
Los datos los sube el administrador desde la barra lateral (con contraseña);
cualquier persona con acceso a la app ve el dashboard ya cargado, sin subir
nada.
"""
import os
import pickle
import datetime

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

from parsing import (
    read_raw_rows, parse_date_from_filename, parse_workbook_trades,
    parse_boletin_sen, dedupe_trades, dedupe_boletines, classify,
    parse_maturity_from_titulo,
)

st.set_page_config(page_title="ALM · TES", layout="wide", page_icon="📈")

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)
TRADES_PATH = os.path.join(DATA_DIR, "trades.pkl")
BOLETIN_PATH = os.path.join(DATA_DIR, "boletines.pkl")

CATS = ["Tasa Fija", "UVR", "Otro"]

# ---------------------------------------------------------------- estilos --
st.markdown("""
<style>
  .stApp{background:#0A0D12;}
  [data-testid="stMetricValue"]{font-family:'JetBrains Mono', monospace; color:#E6EAF2;}
  [data-testid="stMetricLabel"]{color:#8892A6;}
  .titulo-badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;}
  .badge-tf{background:rgba(79,180,216,0.15);color:#4FB4D8;}
  .badge-uvr{background:rgba(185,139,224,0.15);color:#B98BE0;}
  .badge-otro{background:rgba(136,146,166,0.15);color:#8892A6;}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------- estado --
def load_data():
    trades = pickle.load(open(TRADES_PATH, "rb")) if os.path.exists(TRADES_PATH) else []
    boletines = pickle.load(open(BOLETIN_PATH, "rb")) if os.path.exists(BOLETIN_PATH) else []
    return trades, boletines


def save_trades(trades):
    pickle.dump(trades, open(TRADES_PATH, "wb"))


def save_boletines(regs):
    pickle.dump(regs, open(BOLETIN_PATH, "wb"))


if "trades" not in st.session_state or "boletines" not in st.session_state:
    st.session_state.trades, st.session_state.boletines = load_data()


# ---------------------------------------------------------------- sidebar admin --
def admin_section():
    st.sidebar.markdown("### 🔒 Actualizar datos")
    with st.sidebar.expander("Panel de administrador", expanded=False):
        pw = st.text_input("Contraseña", type="password", key="admin_pw")
        real_pw = st.secrets.get("admin_password", None)
        if real_pw is None:
            st.warning("No hay contraseña configurada (admin_password en Secrets). Cualquiera puede subir datos.")
            unlocked = True
        else:
            unlocked = pw == real_pw
            if pw and not unlocked:
                st.error("Contraseña incorrecta.")

        if unlocked and (pw or real_pw is None):
            st.markdown("**Operaciones** (OperDia o histórico mensual)")
            op_files = st.file_uploader(
                "Sube uno o varios archivos", type=["xlsx", "xls"],
                accept_multiple_files=True, key="op_uploader",
            )
            if op_files and st.button("Guardar operaciones", key="save_trades_btn"):
                new_trades, skipped = [], []
                for f in op_files:
                    try:
                        rows = read_raw_rows(f.read(), f.name)
                        file_date = parse_date_from_filename(f.name)
                        parsed = parse_workbook_trades(rows, file_date)
                        if not parsed:
                            skipped.append(f.name)
                        else:
                            new_trades.extend(parsed)
                    except Exception as e:
                        skipped.append(f"{f.name} ({e})")
                st.session_state.trades = dedupe_trades(st.session_state.trades + new_trades)
                save_trades(st.session_state.trades)
                st.success(f"{len(new_trades)} operaciones procesadas de {len(op_files)} archivo(s).")
                if skipped:
                    st.warning("Sin datos válidos: " + ", ".join(skipped))
                st.rerun()

            st.markdown("**Boletín de cierre (SEN)**")
            bo_files = st.file_uploader(
                "Sube uno o varios archivos", type=["xlsx", "xls"],
                accept_multiple_files=True, key="bo_uploader",
            )
            if bo_files and st.button("Guardar boletín", key="save_boletin_btn"):
                new_regs, skipped = [], []
                for f in bo_files:
                    try:
                        rows = read_raw_rows(f.read(), f.name)
                        parsed = parse_boletin_sen(rows, f.name)
                        if not parsed:
                            skipped.append(f.name)
                        else:
                            new_regs.extend(parsed)
                    except Exception as e:
                        skipped.append(f"{f.name} ({e})")
                st.session_state.boletines = dedupe_boletines(st.session_state.boletines + new_regs)
                save_boletines(st.session_state.boletines)
                st.success(f"{len(new_regs)} registros procesados de {len(bo_files)} archivo(s).")
                if skipped:
                    st.warning("Sin datos válidos: " + ", ".join(skipped))
                st.rerun()

            st.markdown("---")
            if st.button("🗑️ Borrar todos los datos", key="clear_btn"):
                st.session_state.trades, st.session_state.boletines = [], []
                save_trades([]); save_boletines([])
                st.success("Datos borrados.")
                st.rerun()

    n_t = len(st.session_state.trades)
    n_b = len(st.session_state.boletines)
    st.sidebar.caption(f"{n_t} operaciones · {n_b} registros de boletín cargados en el servidor.")


admin_section()

trades = st.session_state.trades
boletines = st.session_state.boletines
df_trades = pd.DataFrame(trades) if trades else pd.DataFrame(
    columns=["date", "tsec", "tipo", "titulo", "precio", "tasa", "nominal", "monto"])
df_bol = pd.DataFrame(boletines) if boletines else pd.DataFrame(
    columns=["date", "titulo", "nominal_trans", "cant_cierres", "precio_apertura", "tasa_apertura",
             "precio_min", "tasa_max", "precio_medio", "tasa_medio", "precio_max", "tasa_min",
             "precio_cierre", "tasa_cierre"])

# ---------------------------------------------------------------- header --
st.markdown("<div style='font-family:JetBrains Mono, monospace;font-size:11px;letter-spacing:.14em;color:#E8A33D;font-weight:700;'>ALM</div>", unsafe_allow_html=True)
st.title("TES · Operaciones por título")

if df_trades.empty and df_bol.empty:
    st.info(
        "Todavía no hay datos cargados. Si eres el administrador, abre el panel "
        "'🔒 Actualizar datos' en la barra lateral izquierda y sube operaciones "
        "y/o el boletín de cierre SEN."
    )
    st.stop()

titulo_cat = {}
for t in set(df_trades["titulo"]) | set(df_bol["titulo"]):
    titulo_cat[t] = classify(t)


# ---------------------------------------------------------------- helpers de formato --
def fmt_pct(v, dec=3):
    if v is None or pd.isna(v):
        return "—"
    return f"{v:.{dec}f}%"


def fmt_mm(v):
    if v is None or pd.isna(v):
        return "—"
    return f"{v/1e6:,.1f} MM"


def fmt_price(v, dec=None):
    if v is None or pd.isna(v):
        return "—"
    if dec is None:
        dec = 3 if abs(v % 1) > 0.001 else 0
    return f"{v:,.{dec}f}"


def fmt_time(tsec):
    if tsec is None or pd.isna(tsec):
        return "—"
    tsec = int(tsec)
    return f"{tsec//3600:02d}:{(tsec%3600)//60:02d}:{tsec%60:02d}"


def badge(cat):
    cls = {"Tasa Fija": "badge-tf", "UVR": "badge-uvr", "Otro": "badge-otro"}.get(cat, "badge-otro")
    return f'<span class="titulo-badge {cls}">{cat}</span>'


# ---------------------------------------------------------------- Resumen --
def render_resumen():
    st.caption("Cierre oficial del boletín SEN, título por título.")
    if df_bol.empty:
        st.info("Sube un boletín de cierre (SEN) para ver el resumen.")
        return
    latest = df_bol.sort_values("date").groupby("titulo", as_index=False).tail(1).copy()
    latest["Tipo"] = latest["titulo"].map(titulo_cat)
    latest = latest.sort_values("titulo")
    show = pd.DataFrame({
        "Título": latest["titulo"],
        "Tipo": latest["Tipo"],
        "Fecha cierre": latest["date"].map(lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "—"),
        "Apertura": latest["tasa_apertura"].map(fmt_pct),
        "Mínimo": latest["tasa_min"].map(fmt_pct),
        "Promedio": latest["tasa_medio"].map(fmt_pct),
        "Máximo": latest["tasa_max"].map(fmt_pct),
        "Cierre": latest["tasa_cierre"].map(fmt_pct),
    })
    st.dataframe(show, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------- distribución --
def build_distribution_chart(sub_df, col, label, color):
    n_bins = max(3, min(14, len(sub_df)))
    vals = sub_df[col].astype(float)
    mn, mx = vals.min(), vals.max()
    bin_size = max(0.001, (mx - mn) / n_bins) if mx > mn else 1.0
    bins = {}
    for _, row in sub_df.iterrows():
        b = mn + round((row[col] - mn) / bin_size) * bin_size
        b = round(b, 3)
        bins[b] = bins.get(b, 0) + row["nominal"]
    keys = sorted(bins.keys(), reverse=True)
    fig = go.Figure(go.Bar(
        x=[bins[k] / 1e6 for k in keys],
        y=[f"{k:.3f}" for k in keys],
        orientation="h",
        marker_color=color,
    ))
    fig.update_layout(
        template="plotly_dark", height=340,
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis_title="Nominal (COP MM)",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# ---------------------------------------------------------------- detalle de título --
def render_titulo_detail(titulo):
    t_trades = df_trades[df_trades["titulo"] == titulo].copy() if not df_trades.empty else df_trades
    t_bol = df_bol[df_bol["titulo"] == titulo].copy() if not df_bol.empty else df_bol
    t_bol = t_bol.sort_values("date") if not t_bol.empty else t_bol

    st.markdown(f"#### Cierre oficial · Boletín SEN" + (f" · {t_bol.iloc[-1]['date'].strftime('%d/%m/%Y')}" if not t_bol.empty else ""))
    if not t_bol.empty:
        last = t_bol.iloc[-1]
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Apertura", fmt_pct(last["tasa_apertura"]))
        c2.metric("Mínimo", fmt_pct(last["tasa_min"]))
        c3.metric("Promedio", fmt_pct(last["tasa_medio"]))
        c4.metric("Máximo", fmt_pct(last["tasa_max"]))
        c5.metric("Cierre", fmt_pct(last["tasa_cierre"]))
        if len(t_bol) >= 2:
            with st.expander("Histórico de cierres SEN cargados"):
                hist = t_bol.sort_values("date", ascending=False)
                show = pd.DataFrame({
                    "Fecha": hist["date"].map(lambda d: d.strftime("%d/%m/%Y")),
                    "Apertura": hist["tasa_apertura"].map(fmt_pct),
                    "Mínimo": hist["tasa_min"].map(fmt_pct),
                    "Promedio": hist["tasa_medio"].map(fmt_pct),
                    "Máximo": hist["tasa_max"].map(fmt_pct),
                    "Cierre": hist["tasa_cierre"].map(fmt_pct),
                })
                st.dataframe(show, use_container_width=True, hide_index=True)
    else:
        st.caption("Sin boletín SEN cargado para este título.")

    st.markdown("---")
    st.markdown("#### Operaciones registradas")
    if t_trades.empty:
        st.caption("Sin operaciones registradas para este título — solo cierre oficial SEN disponible.")
        return

    has_dates = t_trades["date"].notna().any()
    if has_dates:
        period = st.radio(
            "Período", ["Semana", "Mes", "2 meses", "Todo"], horizontal=True,
            index=3, key=f"period_{titulo}",
        )
        n_map = {"Semana": 5, "Mes": 21, "2 meses": 42, "Todo": None}
        n = n_map[period]
        if n:
            unique_dates = sorted(t_trades.loc[t_trades["date"].notna(), "date"].unique())
            cutoff = set(unique_dates[-n:])
            t_trades = t_trades[t_trades["date"].isna() | t_trades["date"].isin(cutoff)]

    total_nominal = t_trades["nominal"].sum()
    n_ops = len(t_trades)
    has_tipo = t_trades["tipo"].isin(["Compra", "Venta"]).any()
    compra = t_trades.loc[t_trades["tipo"] == "Compra", "nominal"].sum() if has_tipo else None
    venta = t_trades.loc[t_trades["tipo"] == "Venta", "nominal"].sum() if has_tipo else None

    c1, c2, c3 = st.columns(3)
    c1.metric("Operaciones", n_ops)
    c2.metric("Nominal negociado", fmt_mm(total_nominal))
    c3.metric("Compra / Venta", f"{fmt_mm(compra)} / {fmt_mm(venta)}" if has_tipo else "N/D")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Distribución de precio**")
        st.plotly_chart(build_distribution_chart(t_trades, "precio", "Precio", "rgba(79,180,216,0.7)"),
                         use_container_width=True)
    with col2:
        t_tasa = t_trades[t_trades["tasa"].notna()]
        st.markdown("**Distribución de tasa**")
        if not t_tasa.empty:
            st.plotly_chart(build_distribution_chart(t_tasa, "tasa", "Tasa", "rgba(79,180,216,0.7)"),
                             use_container_width=True)
        else:
            st.caption("Sin datos de tasa en las operaciones cargadas.")

    if has_dates:
        rows = []
        for d, g in t_trades.dropna(subset=["date"]).groupby("date"):
            vwap_p = (g["precio"] * g["nominal"]).sum() / g["nominal"].sum()
            gt = g.dropna(subset=["tasa"])
            vwap_t = (gt["tasa"] * gt["nominal"]).sum() / gt["nominal"].sum() if not gt.empty else None
            rows.append({"date": d, "vwap_precio": vwap_p, "vwap_tasa": vwap_t})
        serie = pd.DataFrame(rows).sort_values("date")
        if len(serie) >= 2:
            st.markdown("#### Serie diaria — precio y tasa promedio ponderados")
            c1, c2 = st.columns(2)
            with c1:
                fig = go.Figure(go.Scatter(x=serie["date"], y=serie["vwap_precio"], mode="lines+markers",
                                            line=dict(color="#33C285", width=2)))
                fig.update_layout(template="plotly_dark", height=220, margin=dict(l=10, r=10, t=10, b=10),
                                   paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = go.Figure(go.Scatter(x=serie["date"], y=serie["vwap_tasa"], mode="lines+markers",
                                            line=dict(color="#4FB4D8", width=2)))
                fig.update_layout(template="plotly_dark", height=220, margin=dict(l=10, r=10, t=10, b=10),
                                   paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)

    with st.expander(f"Operaciones del título ({len(t_trades)})"):
        show_trades = t_trades.sort_values(["date", "tsec"], ascending=False).copy()
        show = pd.DataFrame({
            "Fecha": show_trades["date"].map(lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "—"),
            "Hora": show_trades["tsec"].map(fmt_time),
            "Tipo": show_trades["tipo"].fillna("—"),
            "Precio": show_trades["precio"].map(fmt_price),
            "Tasa": show_trades["tasa"].map(fmt_pct),
            "Nominal": show_trades["nominal"].map(fmt_mm),
            "Monto": show_trades["monto"].map(fmt_mm),
        })
        st.dataframe(show, use_container_width=True, hide_index=True)


def render_categoria(cat_name):
    titulos_cat = sorted([t for t, c in titulo_cat.items() if c == cat_name])
    if not titulos_cat:
        st.info(f"No hay títulos cargados en {cat_name}.")
        return
    sel = st.selectbox("Título", titulos_cat, key=f"sel_{cat_name}")
    render_titulo_detail(sel)


# ---------------------------------------------------------------- Curva TF --
def render_curva():
    st.caption("Compara la curva de tasas y el nominal negociado entre dos fechas, título por título.")
    if df_bol.empty:
        st.info("Sube boletines con títulos TFIT para ver la curva.")
        return
    tf = df_bol[df_bol["titulo"].str.startswith("TFIT")].copy()
    if tf.empty:
        st.info("Sube boletines con títulos TFIT para ver la curva.")
        return

    fechas = sorted(tf["date"].dropna().unique())
    if not fechas:
        st.info("Los boletines cargados no tienen fecha reconocible.")
        return

    c1, c2 = st.columns(2)
    with c1:
        f1 = st.selectbox("Fecha 1", fechas, index=0, format_func=lambda d: d.strftime("%d/%m/%Y"), key="curva_f1")
    with c2:
        f2 = st.selectbox("Fecha 2", fechas, index=len(fechas) - 1, format_func=lambda d: d.strftime("%d/%m/%Y"), key="curva_f2")

    def nodes_for(f):
        sub = tf[tf["date"] == f].copy()
        sub["maturity"] = sub["titulo"].map(parse_maturity_from_titulo)
        return sub.dropna(subset=["maturity"])

    n1, n2 = nodes_for(f1), nodes_for(f2)
    merged = pd.merge(
        n1[["titulo", "maturity", "tasa_cierre", "nominal_trans"]],
        n2[["titulo", "maturity", "tasa_cierre", "nominal_trans"]],
        on="titulo", how="outer", suffixes=("_1", "_2"),
    )
    merged["maturity"] = merged["maturity_1"].combine_first(merged["maturity_2"])
    merged = merged.dropna(subset=["maturity"]).sort_values("maturity")

    st.markdown(
        f"🔵 **Fecha 1** — {f1.strftime('%d/%m/%Y')}"
        f"&nbsp;&nbsp;&nbsp; 🟢 **Fecha 2** — {f2.strftime('%d/%m/%Y')}"
    )

    years = [m.year for m in merged["maturity"]]
    fig = go.Figure()
    fig.add_bar(x=years, y=merged["nominal_trans_1"], name=f"Nominal {f1.strftime('%d/%m/%Y')}",
                marker_color="rgba(79,180,216,0.30)", yaxis="y2")
    fig.add_bar(x=years, y=merged["nominal_trans_2"], name=f"Nominal {f2.strftime('%d/%m/%Y')}",
                marker_color="rgba(51,194,133,0.30)", yaxis="y2")
    fig.add_trace(go.Scatter(x=years, y=merged["tasa_cierre_1"], name=f"Tasa {f1.strftime('%d/%m/%Y')}",
                              mode="lines+markers", line=dict(color="#4FB4D8", width=3), marker=dict(size=7)))
    fig.add_trace(go.Scatter(x=years, y=merged["tasa_cierre_2"], name=f"Tasa {f2.strftime('%d/%m/%Y')}",
                              mode="lines+markers", line=dict(color="#33C285", width=3), marker=dict(size=7)))
    fig.update_layout(
        template="plotly_dark", height=440,
        yaxis=dict(title="Tasa de cierre (%)"),
        yaxis2=dict(title="Nominal (COP MM)", overlaying="y", side="right", showgrid=False),
        xaxis=dict(title="Vencimiento", type="category"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Cambio por nodo (Fecha 2 − Fecha 1)")
    st.caption("Diferencia en puntos básicos entre las dos fechas. Positivo: subió la tasa. Negativo: bajó.")
    diff = merged.dropna(subset=["tasa_cierre_1", "tasa_cierre_2"]).copy()
    if diff.empty:
        st.info("Ningún título tiene boletín en ambas fechas seleccionadas.")
    else:
        diff["diff_bp"] = (diff["tasa_cierre_2"] - diff["tasa_cierre_1"]) * 100
        show = pd.DataFrame({
            "Título": diff["titulo"],
            "Vencimiento": diff["maturity"].map(lambda d: d.strftime("%d/%m/%Y")),
            "Tasa Fecha 1": diff["tasa_cierre_1"].map(fmt_pct),
            "Tasa Fecha 2": diff["tasa_cierre_2"].map(fmt_pct),
            "Δ (pb)": diff["diff_bp"].map(lambda v: f"{'+' if v>=0 else ''}{v:.1f} pb"),
        })
        st.dataframe(show, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------- tabs --
tab_names = ["Resumen", "Tasa Fija", "UVR", "Otro"]
if not df_bol.empty and df_bol["titulo"].str.startswith("TFIT").any():
    tab_names.append("Curva TF")

tabs = st.tabs(tab_names)
with tabs[0]:
    render_resumen()
with tabs[1]:
    render_categoria("Tasa Fija")
with tabs[2]:
    render_categoria("UVR")
with tabs[3]:
    render_categoria("Otro")
if len(tabs) > 4:
    with tabs[4]:
        render_curva()

st.caption(
    f"{len(titulo_cat)} títulos distintos · {len(df_trades)} operaciones totales · "
    f"{len(df_bol)} registros de boletín."
)
