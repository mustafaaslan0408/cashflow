import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, date

st.set_page_config(
    page_title="İnpharmus Nakit Akış Raporu",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Stil ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #0f1117; }
    .block-container { padding-top: 1rem; }
    .kpi-card {
        background: #1e2130;
        border-radius: 12px;
        padding: 20px 24px;
        border-left: 4px solid;
        margin-bottom: 8px;
    }
    .kpi-label { font-size: 12px; color: #8b9099; text-transform: uppercase; letter-spacing: 1px; }
    .kpi-value { font-size: 28px; font-weight: 700; margin-top: 4px; }
    .kpi-sub   { font-size: 12px; color: #8b9099; margin-top: 2px; }
    .section-title { font-size: 16px; font-weight: 600; color: #e0e4ef; margin: 16px 0 8px; }
    div[data-testid="stSidebar"] { background-color: #161824; }
    .stDataFrame { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── Sabitler ──────────────────────────────────────────────────────────────────
DAHIL_FIS = {
    "Banka İşlem Fişi",
    "Gelen Havale /Eft",
    "Gönderilen Havale /Eft",
    "Banka Virman Fişi",
    "Çek Çıkış (Banka Tahsil)",
}
HARIC_KAT = {"İştirak Gelen", "İştirak Giden"}

GELIR_KATLAR = [
    "Yurtiçi Mal Satışı", "İhracat", "Hizmet Geliri", "Faiz Geliri",
    "Kira Geliri", "Diğer Gelirler",
]
GIDER_KATLAR = [
    "Hammadde / Malzeme Alımı", "Personel Ödemeleri", "SGK Ödemeleri",
    "Vergi Ödemeleri", "Kira Gideri", "Enerji Gideri", "Lojistik",
    "Pazarlama", "Faiz Gideri", "Diğer Giderler",
]

KAT_NORM = {
    "yurt içi mal satışı": "Yurtiçi Mal Satışı",
    "yurtiçi mal satisi": "Yurtiçi Mal Satışı",
    "ihracat": "İhracat",
    "hizmet geliri": "Hizmet Geliri",
    "faiz geliri": "Faiz Geliri",
    "kira geliri": "Kira Geliri",
    "hammadde": "Hammadde / Malzeme Alımı",
    "malzeme": "Hammadde / Malzeme Alımı",
    "personel": "Personel Ödemeleri",
    "maaş": "Personel Ödemeleri",
    "sgk": "SGK Ödemeleri",
    "vergi": "Vergi Ödemeleri",
    "kdv": "Vergi Ödemeleri",
    "kira gider": "Kira Gideri",
    "enerji": "Enerji Gideri",
    "elektrik": "Enerji Gideri",
    "lojistik": "Lojistik",
    "kargo": "Lojistik",
    "pazarlama": "Pazarlama",
    "reklam": "Pazarlama",
    "faiz gider": "Faiz Gideri",
}

def normalize_kat(k):
    if not isinstance(k, str):
        return "Diğer"
    kl = k.lower().strip()
    for anahtar, deger in KAT_NORM.items():
        if anahtar in kl:
            return deger
    return k.strip() if k.strip() else "Diğer"

def fmt_tl(val):
    if pd.isna(val):
        return "—"
    sign = "-" if val < 0 else ""
    return f"{sign}₺{abs(val):,.0f}".replace(",", ".")

# ── Veri yükleme ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Veri okunuyor...")
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name="CHIP_BANKA_DOKUMU",
        header=0,
        dtype=str,
    )
    col_map = {
        df.columns[0]:  "Firma",
        df.columns[1]:  "BankaAdi",
        df.columns[5]:  "IslemTuru",
        df.columns[6]:  "HareketTuru",
        df.columns[9]:  "Tarih",
        df.columns[10]: "CariKod",
        df.columns[11]: "CariUnvan",
        df.columns[12]: "Aciklama",
        df.columns[13]: "HareketOzelKod",
        df.columns[14]: "HareketOzelAciklama",
        df.columns[15]: "DovizTuru",
        df.columns[16]: "DovizKuru",
        df.columns[17]: "DovizTutari",
        df.columns[19]: "GiderAdi",
    }
    df = df.rename(columns=col_map)
    keep = list(col_map.values())
    df = df[[c for c in keep if c in df.columns]]

    # Tarih
    df["Tarih"] = pd.to_datetime(df["Tarih"], errors="coerce", dayfirst=True)

    # Sayısal
    for col in ["DovizKuru", "DovizTutari"]:
        df[col] = pd.to_numeric(df[col].str.replace(",", "."), errors="coerce").fillna(0)

    # TL Tutar
    df["TL_Tutar"] = df.apply(
        lambda r: r["DovizTutari"] if r["DovizTuru"] == "TL"
                  else r["DovizKuru"] * r["DovizTutari"],
        axis=1
    )

    # Kategori
    df["Kategori"] = df.apply(
        lambda r: r["GiderAdi"] if r["IslemTuru"] == "Banka İşlem Fişi"
                  else r["HareketOzelAciklama"],
        axis=1
    )
    df["Kategori"] = df["Kategori"].apply(normalize_kat)

    # CariLabel: boşsa GiderAdi
    df["CariLabel"] = df.apply(
        lambda r: r["GiderAdi"] if (not isinstance(r["CariUnvan"], str) or r["CariUnvan"].strip() == "")
                  else r["CariUnvan"],
        axis=1
    )

    # Filtrele
    df = df[df["IslemTuru"].isin(DAHIL_FIS)]
    df = df[~df["Kategori"].isin(HARIC_KAT)]
    df = df.dropna(subset=["Tarih"])
    return df

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💊 İnpharmus")
    st.markdown("### Nakit Akış Raporu")
    st.divider()

    uploaded = st.file_uploader("Excel Dosyası", type=["xlsx"], label_visibility="collapsed")

    if uploaded:
        import tempfile, os
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name
        df_raw = load_data(tmp_path)
        os.unlink(tmp_path)
    else:
        st.info("Lütfen BANKADOKUM.xlsx dosyasını yükleyin.")
        st.stop()

    st.markdown("#### 📅 Tarih Aralığı")
    min_d = df_raw["Tarih"].min().date()
    max_d = df_raw["Tarih"].max().date()
    bas_tarih = st.date_input("Başlangıç", value=date(max_d.year, 1, 1), min_value=min_d, max_value=max_d)
    bit_tarih = st.date_input("Bitiş",     value=max_d,                  min_value=min_d, max_value=max_d)

    st.markdown("#### 🏢 Firma")
    firmalar = ["Tümü"] + sorted(df_raw["Firma"].dropna().unique().tolist())
    secili_firma = st.selectbox("Firma", firmalar)

    st.markdown("#### 🏦 Banka")
    bankalar = ["Tümü"] + sorted(df_raw["BankaAdi"].dropna().unique().tolist())
    secili_banka = st.selectbox("Banka", bankalar)

    st.markdown("#### 📊 Periyot")
    periyot = st.radio("Görünüm", ["Aylık", "Haftalık"], horizontal=True)

# ── Filtreleme ────────────────────────────────────────────────────────────────
df = df_raw.copy()
df = df[(df["Tarih"].dt.date >= bas_tarih) & (df["Tarih"].dt.date <= bit_tarih)]
if secili_firma != "Tümü":
    df = df[df["Firma"] == secili_firma]
if secili_banka != "Tümü":
    df = df[df["BankaAdi"] == secili_banka]

# ── KPI hesaplama ─────────────────────────────────────────────────────────────
GELIR_df = df[df["Kategori"].isin(GELIR_KATLAR)]
GIDER_df = df[df["Kategori"].isin(GIDER_KATLAR)]
diger_df  = df[~df["Kategori"].isin(GELIR_KATLAR + GIDER_KATLAR)]

toplam_giris = GELIR_df["TL_Tutar"].sum()
toplam_cikis = abs(GIDER_df["TL_Tutar"].sum())
net          = toplam_giris - toplam_cikis
virman_top   = df[df["IslemTuru"] == "Banka Virman Fişi"]["TL_Tutar"].abs().sum()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# İnpharmus İlaç — Nakit Akış Raporu")
st.caption(f"{bas_tarih.strftime('%d.%m.%Y')} – {bit_tarih.strftime('%d.%m.%Y')}  •  {len(df):,} işlem")
st.divider()

# ── KPI Kartları ──────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"""<div class="kpi-card" style="border-color:#22c55e">
        <div class="kpi-label">Toplam Giriş</div>
        <div class="kpi-value" style="color:#22c55e">{fmt_tl(toplam_giris)}</div>
        <div class="kpi-sub">{len(GELIR_df):,} işlem</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="kpi-card" style="border-color:#ef4444">
        <div class="kpi-label">Toplam Çıkış</div>
        <div class="kpi-value" style="color:#ef4444">-₺{toplam_cikis:,.0f}".replace(",", ".")</div>
        <div class="kpi-sub">{len(GIDER_df):,} işlem</div>
    </div>""", unsafe_allow_html=True)
with c3:
    net_renk = "#22c55e" if net >= 0 else "#ef4444"
    st.markdown(f"""<div class="kpi-card" style="border-color:{net_renk}">
        <div class="kpi-label">Net Nakit Akışı</div>
        <div class="kpi-value" style="color:{net_renk}">{fmt_tl(net)}</div>
        <div class="kpi-sub">Giriş − Çıkış</div>
    </div>""", unsafe_allow_html=True)
with c4:
    st.markdown(f"""<div class="kpi-card" style="border-color:#a78bfa">
        <div class="kpi-label">Virman Toplamı</div>
        <div class="kpi-value" style="color:#a78bfa">{fmt_tl(virman_top)}</div>
        <div class="kpi-sub">İştirak transferleri</div>
    </div>""", unsafe_allow_html=True)

st.divider()

# ── Trend Grafiği ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📈 Nakit Akışı Trendi</div>', unsafe_allow_html=True)

if periyot == "Aylık":
    df["Periyot"] = df["Tarih"].dt.to_period("M").astype(str)
else:
    df["Periyot"] = df["Tarih"].dt.to_period("W").apply(lambda r: r.start_time.strftime("%Y-W%V"))

trend_giris = df[df["Kategori"].isin(GELIR_KATLAR)].groupby("Periyot")["TL_Tutar"].sum()
trend_cikis = df[df["Kategori"].isin(GIDER_KATLAR)].groupby("Periyot")["TL_Tutar"].sum().abs()
all_periods  = sorted(set(trend_giris.index) | set(trend_cikis.index))

fig_trend = go.Figure()
fig_trend.add_trace(go.Scatter(
    x=all_periods, y=[trend_giris.get(p, 0) for p in all_periods],
    name="Giriş", line=dict(color="#22c55e", width=2.5),
    fill="tozeroy", fillcolor="rgba(34,197,94,0.08)",
    mode="lines+markers", marker=dict(size=5)
))
fig_trend.add_trace(go.Scatter(
    x=all_periods, y=[-trend_cikis.get(p, 0) for p in all_periods],
    name="Çıkış", line=dict(color="#ef4444", width=2.5),
    fill="tozeroy", fillcolor="rgba(239,68,68,0.08)",
    mode="lines+markers", marker=dict(size=5)
))
fig_trend.update_layout(
    paper_bgcolor="#1e2130", plot_bgcolor="#1e2130",
    font=dict(color="#c9d1d9"), height=320,
    legend=dict(orientation="h", y=1.08),
    margin=dict(l=0, r=0, t=10, b=0),
    yaxis=dict(gridcolor="#2d3146", tickformat=",.0f"),
    xaxis=dict(gridcolor="#2d3146"),
    hovermode="x unified"
)
st.plotly_chart(fig_trend, use_container_width=True)

# ── Kategori Bar + Tablo ──────────────────────────────────────────────────────
col_bar, col_tbl = st.columns([1, 1])

with col_bar:
    st.markdown('<div class="section-title">📊 Kategori Dağılımı</div>', unsafe_allow_html=True)
    kat_top = (
        df.groupby("Kategori")["TL_Tutar"].sum()
        .sort_values(ascending=True)
        .tail(15)
    )
    renkler = ["#22c55e" if v >= 0 else "#ef4444" for v in kat_top.values]
    fig_bar = go.Figure(go.Bar(
        x=kat_top.values,
        y=kat_top.index,
        orientation="h",
        marker_color=renkler,
        text=[fmt_tl(v) for v in kat_top.values],
        textposition="outside",
    ))
    fig_bar.update_layout(
        paper_bgcolor="#1e2130", plot_bgcolor="#1e2130",
        font=dict(color="#c9d1d9", size=11), height=420,
        margin=dict(l=0, r=80, t=10, b=0),
        xaxis=dict(gridcolor="#2d3146", showticklabels=False),
        yaxis=dict(gridcolor="rgba(0,0,0,0)"),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col_tbl:
    st.markdown('<div class="section-title">🗂 Nakit Akış Tablosu</div>', unsafe_allow_html=True)
    rows = []
    for kat in GELIR_KATLAR:
        kat_df = df[df["Kategori"] == kat]
        total  = kat_df["TL_Tutar"].sum()
        if total == 0:
            continue
        rows.append({"Kategori / Cari": f"▶ {kat}", "Tutar (TL)": fmt_tl(total), "_sort": total, "_level": 0})
        top_caris = (
            kat_df.groupby("CariLabel")["TL_Tutar"].sum()
            .sort_values(ascending=False)
            .head(8)
        )
        for cari, val in top_caris.items():
            rows.append({"Kategori / Cari": f"    {cari}", "Tutar (TL)": fmt_tl(val), "_sort": val, "_level": 1})

    rows.append({"Kategori / Cari": "━━━ GİDERLER ━━━", "Tutar (TL)": "", "_sort": 0, "_level": -1})

    for kat in GIDER_KATLAR:
        kat_df = df[df["Kategori"] == kat]
        total  = kat_df["TL_Tutar"].sum()
        if total == 0:
            continue
        rows.append({"Kategori / Cari": f"▶ {kat}", "Tutar (TL)": fmt_tl(total), "_sort": total, "_level": 0})
        top_caris = (
            kat_df.groupby("CariLabel")["TL_Tutar"].sum()
            .sort_values(ascending=True)
            .head(8)
        )
        for cari, val in top_caris.items():
            rows.append({"Kategori / Cari": f"    {cari}", "Tutar (TL)": fmt_tl(val), "_sort": val, "_level": 1})

    tbl_df = pd.DataFrame(rows)[["Kategori / Cari", "Tutar (TL)"]]
    st.dataframe(tbl_df, use_container_width=True, height=420, hide_index=True)

# ── Virman Tablosu ────────────────────────────────────────────────────────────
virman_df = df[df["IslemTuru"] == "Banka Virman Fişi"]
if not virman_df.empty:
    st.divider()
    st.markdown('<div class="section-title">🔄 İştirakler Arası Virmanlar</div>', unsafe_allow_html=True)
    virman_tbl = (
        virman_df.groupby(["Firma", "BankaAdi", "CariLabel"])["TL_Tutar"]
        .sum()
        .reset_index()
        .rename(columns={"Firma": "Firma", "BankaAdi": "Banka", "CariLabel": "Karşı Taraf", "TL_Tutar": "Tutar (TL)"})
        .sort_values("Tutar (TL)", ascending=False)
    )
    virman_tbl["Tutar (TL)"] = virman_tbl["Tutar (TL)"].apply(fmt_tl)
    st.dataframe(virman_tbl, use_container_width=True, hide_index=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("İnpharmus İlaç Sanayi Tic. A.Ş. — Hazine & Finans  |  Veriler CHIP ERP'den alınmıştır.")
