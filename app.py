"""
SIKABI — Sistem Intelijen Karier Bank Indonesia
Transformasi Digital Manajemen Karier Pegawai Bank Indonesia

Jalankan lokal:
    streamlit run app.py
"""

import io
import os

import pandas as pd
import plotly.express as px
import streamlit as st

from scoring import ADMIN_CHECK_LABELS, KPP_CHECK_LABELS, compute_all

BASE_DIR = os.path.dirname(__file__)
DATA_PATH = os.path.join(BASE_DIR, "data", "sikabi_data.xlsx")
LOGO_PATH = os.path.join(BASE_DIR, "data", "Sikabi_header.png")
LOGO_FALLBACK_PATH = os.path.join(BASE_DIR, "data", "Sikabi.png")
REF_SHEETS = [
    "Quant_Weights", "Qual_Weights", "Rank_Weights",
    "Education_Score", "Certification_Score", "K3_Score", "Thresholds",
]

# Palet utama dibuat lebih restrained: navy BI, blue accent, slate neutrals.
NAVY = "#123A63"
BLUE = "#1677C8"
BLUE_SOFT = "#EAF4FC"
INK = "#172B4D"
SLATE = "#5E6C84"
BORDER = "#DCE3EA"
SURFACE = "#FFFFFF"
BG = "#F6F8FB"
KUADRAN_COLOR = {
    "I": "#1677C8",
    "II": "#3B82A0",
    "III": "#7A8E9E",
    "IV": "#A1ABB5",
}
KUADRAN_INFO = {
    "I": ("High QScore · High MDP", "QScore > mean KPP dan MDP > mean KPP"),
    "II": ("High QScore · Low MDP", "QScore > mean KPP dan MDP ≤ mean KPP"),
    "III": ("Low QScore · High MDP", "QScore ≤ mean KPP dan MDP > mean KPP"),
    "IV": ("Low QScore · Low MDP", "QScore ≤ mean KPP dan MDP ≤ mean KPP"),
}

st.set_page_config(
    page_title="SIKABI — Sistem Intelijen Karier",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "S",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    :root {{
        --sikabi-navy: #123A63; --sikabi-blue: #1677C8; --sikabi-blue-soft: #EAF4FC;
        --sikabi-border: #DCE3EA; --sikabi-bg: #F6F8FB; --sikabi-surface: #FFFFFF;
    }}
    .stApp {{ background: var(--background-color, {BG}); color: var(--text-color, {INK}); }}
    .block-container {{ padding-top: 1.2rem; padding-bottom: 2.5rem; max-width: 1500px; }}
    [data-testid="stHeader"] {{ background: var(--background-color, {BG}); }}
    [data-testid="stSidebar"] {{ display: none; }}
    .sikabi-header {{
        background: var(--secondary-background-color, #FFFFFF);
        border: 1px solid var(--sikabi-border); border-radius: 18px; padding: 18px 22px;
        margin-bottom: 18px; box-shadow: 0 6px 24px rgba(18,58,99,.06);
    }}
    .header-actions {{ padding-top: 2px; }}
    .eyebrow {{ color: var(--sikabi-blue); font-size: .70rem; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; margin-bottom: 3px; }}
    .page-title {{ color: var(--text-color, {NAVY}); font-size: 1.52rem; font-weight: 700; margin: 0; }}
    .page-subtitle {{ color: var(--secondary-text-color, {SLATE}); font-size: .86rem; margin-top: 4px; }}
    .section-title {{ color: var(--text-color, {NAVY}); font-size: 1.05rem; font-weight: 700; margin: .2rem 0 .65rem; }}
    .muted {{ color: var(--secondary-text-color, {SLATE}); font-size: .82rem; }}
    .metric-card {{ background: var(--secondary-background-color, #FFFFFF); border:1px solid var(--sikabi-border); border-radius:14px; padding:17px 18px; min-height:112px; box-shadow:0 3px 14px rgba(18,58,99,.035); }}
    .metric-label {{ color: var(--secondary-text-color, {SLATE}); font-size:.78rem; font-weight:600; }}
    .metric-value {{ color: var(--text-color, {NAVY}); font-size:1.72rem; font-weight:700; margin-top:7px; }}
    .metric-note {{ color: var(--secondary-text-color, #7C8A9A); font-size:.72rem; margin-top:2px; }}
    .quad-card {{ background: var(--secondary-background-color, #FFFFFF); border:1px solid var(--sikabi-border); border-radius:14px; padding:16px; height:100%; box-shadow:0 3px 14px rgba(18,58,99,.035); transition:.15s ease; }}
    .quad-card:hover {{ border-color:#7DA8CC; box-shadow:0 8px 22px rgba(18,58,99,.10); transform:translateY(-1px); }}
    .quad-top {{ display:flex; align-items:center; justify-content:space-between; gap:10px; }}
    .quad-name {{ font-weight:700; color: var(--text-color, {NAVY}); font-size:.94rem; }}
    .quad-count {{ font-size:1.45rem; font-weight:700; color: var(--text-color, {INK}); }}
    .quad-desc {{ color: var(--secondary-text-color, {SLATE}); font-size:.74rem; line-height:1.45; margin-top:4px; }}
    .quad-dot {{ width:9px; height:9px; border-radius:50%; display:inline-block; margin-right:7px; }}
    .gate-subsection {{ background: var(--secondary-background-color, #FFFFFF); border:1px solid var(--sikabi-border); border-radius:14px; padding:16px 18px; margin:10px 0 14px; }}
    .gate-heading {{ color: var(--text-color, {NAVY}); font-weight:700; font-size:1rem; margin-bottom:3px; }}
    div[data-testid="stMetric"] {{ background: var(--secondary-background-color, #FFFFFF); border:1px solid var(--sikabi-border); padding:14px 16px; border-radius:12px; }}
    .stButton > button, .stDownloadButton > button {{ border-radius:9px; border:1px solid var(--sikabi-border); background: var(--secondary-background-color, #FFFFFF); color: var(--text-color, {NAVY}); font-weight:600; min-height:38px; }}
    .stButton > button:hover, .stDownloadButton > button:hover {{ border-color:#7DA8CC; color:var(--sikabi-blue); background:var(--sikabi-blue-soft); }}
    .stButton > button[kind="primary"] {{ background:var(--sikabi-navy); border-color:var(--sikabi-navy); color:#fff; }}
    div[data-baseweb="select"] > div {{ border-radius:9px; border-color:var(--sikabi-border); background: var(--secondary-background-color, #FFFFFF); }}
    div[data-baseweb="select"] [data-baseweb="tag"] {{ background: var(--sikabi-blue); color:#fff; border-radius:7px; }}
    div[data-baseweb="select"] [data-baseweb="tag"] svg {{ color:#fff; }}
    div[data-testid="stDataFrame"] {{ border:1px solid var(--sikabi-border); border-radius:12px; overflow:hidden; }}
    .stTabs [data-baseweb="tab-list"] {{ gap:8px; border-bottom:1px solid var(--sikabi-border); }}
    .stTabs [data-baseweb="tab"] {{ padding:10px 14px; color:var(--secondary-text-color, {SLATE}); font-weight:600; }}
    .stTabs [aria-selected="true"] {{ color:var(--sikabi-blue) !important; }}
    .stTabs [data-baseweb="tab-highlight"] {{ background-color:var(--sikabi-blue) !important; }}
    hr {{ border-color:var(--sikabi-border); }}
    @media (prefers-color-scheme: dark) {{
        .sikabi-header {{ box-shadow:0 6px 24px rgba(0,0,0,.18); }}
        .stButton > button:hover, .stDownloadButton > button:hover {{ background:rgba(22,119,200,.14); }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def load_workbook(path):
    return pd.read_excel(path, sheet_name=None)


def init_state():
    if "data_loaded" not in st.session_state:
        sheets = load_workbook(DATA_PATH)
        st.session_state.employees = sheets["Employees"]
        for name in REF_SHEETS:
            st.session_state[f"ref_{name}"] = sheets[name]
        st.session_state.data_loaded = True


def get_refs():
    return {name: st.session_state[f"ref_{name}"] for name in REF_SHEETS}


def recompute():
    return compute_all(st.session_state.employees, get_refs())


def save_to_disk():
    os.makedirs(os.path.dirname(DATA_PATH), exist_ok=True)
    with pd.ExcelWriter(DATA_PATH, engine="openpyxl") as writer:
        st.session_state.employees.to_excel(writer, sheet_name="Employees", index=False)
        for name in REF_SHEETS:
            st.session_state[f"ref_{name}"].to_excel(writer, sheet_name=name, index=False)


def workbook_bytes():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        st.session_state.employees.to_excel(writer, sheet_name="Employees", index=False)
        for name in REF_SHEETS:
            st.session_state[f"ref_{name}"].to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()


def xlsx_bytes(frame):
    buf = io.BytesIO()
    frame.to_excel(buf, index=False)
    return buf.getvalue()


def safe_contains(series, query):
    return series.astype(str).str.contains(query, case=False, na=False, regex=False)


def apply_multiselect_filter(frame, column, values):
    if not values:
        return frame.iloc[0:0]
    return frame[frame[column].isin(values)]


init_state()
df = recompute()
PANGKAT_OPTS = sorted(df["Pangkat"].dropna().unique().tolist())
SATKER_OPTS = sorted(df["Satker"].dropna().unique().tolist())
STATUS_OPTS = ["Proses KPP", "General Talent"]

# Header: logo aplikasi + identitas produk + utility actions.
with st.container(border=True):
    head_logo, head_text, head_actions = st.columns([1.15, 3.0, 1.85], vertical_alignment="center")
    with head_logo:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=220)
        elif os.path.exists(LOGO_FALLBACK_PATH):
            st.image(LOGO_FALLBACK_PATH, width=220)
    with head_text:
        st.markdown('<div class="eyebrow">BANK INDONESIA · TALENT INTELLIGENCE</div>', unsafe_allow_html=True)
        st.markdown('<div class="page-title">Sistem Intelijen Karier Bank Indonesia</div>', unsafe_allow_html=True)
        st.markdown('<div class="page-subtitle">Transformasi digital manajemen karier berbasis scoring, decision gate, dan talent mapping.</div>', unsafe_allow_html=True)
    with head_actions:
        st.download_button("Unduh data (.xlsx)", data=workbook_bytes(), file_name="sikabi_data.xlsx", use_container_width=True, key="download_master_workbook")
        st.button("Sinkronisasi HRIS & KATALIS", disabled=True, use_container_width=True, key="sync_hris_katalis")
        st.caption("Integrasi eksternal belum terhubung.")

# Top navigation — tabs menggantikan radio button/sidebar navigation.
tab_dashboard, tab_pegawai, tab_gate, tab_master = st.tabs([
    "Dashboard",
    "Data Pegawai",
    "Gate Keputusan",
    "Master Data",
])


def metric_card(label, value, note=""):
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-note">{note}</div>
    </div>
    """


def filter_bar_dashboard():
    st.markdown('<div class="section-title">Filter analisis</div>', unsafe_allow_html=True)
    a, b, c = st.columns([1, 1.35, 1.35])
    with a:
        fp = st.multiselect("Pangkat", PANGKAT_OPTS, default=PANGKAT_OPTS, key="dash_pangkat")
    with b:
        fs = st.multiselect("Satuan kerja", SATKER_OPTS, default=SATKER_OPTS, key="dash_satker")
    with c:
        fq = st.multiselect("Kuadran", ["I", "II", "III", "IV"], default=["I", "II", "III", "IV"], key="dash_kuadran")
    return fp, fs, fq


def page_dashboard(data):
    st.markdown('<div class="sikabi-header"><div class="eyebrow">EXECUTIVE OVERVIEW</div><div class="page-title">Dashboard Kuadran & Laporan</div><div class="page-subtitle">Pemetaan kandidat berdasarkan QScore dan Masa Dinas Pangkat (MDP).</div></div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(metric_card("Total populasi", f"{len(data):,}", "Seluruh data pegawai"), unsafe_allow_html=True)
    c2.markdown(metric_card("Lolos administrasi", f"{int(data['Lolos_Administrasi'].sum()):,}", "Gate administrasi"), unsafe_allow_html=True)
    c3.markdown(metric_card("Lolos kriteria KPP", f"{int(data['Lolos_KPP'].sum()):,}", "Gate KPP"), unsafe_allow_html=True)
    c4.markdown(metric_card("Proses KPP", f"{int(data['Masuk_Proses_KPP'].sum()):,}", "Populasi yang dipetakan"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    f_pangkat, f_satker, f_kuadran = filter_bar_dashboard()

    kpp_pop = data[data["Masuk_Proses_KPP"]].copy()
    scoped = apply_multiselect_filter(kpp_pop, "Pangkat", f_pangkat)
    scoped = apply_multiselect_filter(scoped, "Satker", f_satker)
    visible = apply_multiselect_filter(scoped, "Kuadran", f_kuadran)

    mean_q = float(kpp_pop["Mean_QScore_Populasi_KPP"].dropna().iloc[0]) if len(kpp_pop) else None
    mean_mdp = float(kpp_pop["Mean_MDP_Populasi_KPP"].dropna().iloc[0]) if len(kpp_pop) else None

    left, right = st.columns([1.8, 1], gap="large")
    with left:
        st.markdown('<div class="section-title">Peta kuadran</div>', unsafe_allow_html=True)
        st.caption(f"{len(visible):,} pegawai terlihat · {len(kpp_pop):,} pegawai Proses KPP · garis referensi menggunakan mean populasi KPP")
        if len(visible):
            fig = px.scatter(
                visible,
                x="MDP_Tahun", y="QScore", color="Kuadran",
                color_discrete_map=KUADRAN_COLOR,
                hover_name="Nama",
                hover_data={
                    "Satker": True, "Pangkat": True, "Sublevel": True,
                    "QScore": ":.1f", "MDP_Tahun": ":.2f", "Kuadran": True,
                },
                labels={"MDP_Tahun": "Masa Dinas Pangkat (tahun)", "QScore": "QScore"},
            )
            if mean_q is not None:
                fig.add_hline(y=mean_q, line_dash="dash", line_color="#7A8795", annotation_text=f"Mean QScore KPP {mean_q:.1f}", annotation_position="top left")
            if mean_mdp is not None:
                fig.add_vline(x=mean_mdp, line_dash="dash", line_color="#7A8795", annotation_text=f"Mean MDP KPP {mean_mdp:.1f}", annotation_position="top right")
            fig.update_traces(marker=dict(size=9, line=dict(width=0.7, color="white")), opacity=0.86)
            fig.update_layout(
                height=470,
                margin=dict(l=20, r=20, t=20, b=20),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                legend_title_text="",
                font=dict(family="Inter", color="var(--text-color)"),
                xaxis=dict(gridcolor="rgba(127,127,127,0.18)"), yaxis=dict(gridcolor="rgba(127,127,127,0.18)"),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("Tidak ada data pada kombinasi filter ini.")

    with right:
        st.markdown('<div class="section-title">Ringkasan kuadran</div>', unsafe_allow_html=True)
        st.caption("Arahkan kursor ke kartu untuk melihat nama pegawai. Tombol unduh mengikuti filter pangkat & satker yang aktif.")
        for k in ["I", "II", "III", "IV"]:
            qdf = scoped[scoped["Kuadran"] == k].copy()
            names = qdf["Nama"].astype(str).tolist()
            tooltip = "\n".join(names) if names else "Tidak ada pegawai"
            title, desc = KUADRAN_INFO[k]
            opacity = 1 if k in f_kuadran else 0.42
            col = KUADRAN_COLOR[k]
            st.markdown(
                f"""
                <div class="quad-card" title="{tooltip}" style="opacity:{opacity}; border-left:4px solid {col}; margin-bottom:10px;">
                    <div class="quad-top">
                        <div><span class="quad-dot" style="background:{col}"></span><span class="quad-name">Kuadran {k}</span></div>
                        <div class="quad-count">{len(qdf)}</div>
                    </div>
                    <div class="quad-desc"><b>{title}</b><br>{desc}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.download_button(
                f"Unduh Kuadran {k}",
                data=xlsx_bytes(qdf[["NIP", "Nama", "Satker", "Pangkat", "Sublevel", "QScore", "MDP_Tahun", "Kuadran", "Readiness"]]) if len(qdf) else b"",
                file_name=f"sikabi_kuadran_{k}.xlsx",
                disabled=not len(qdf),
                use_container_width=True,
                key=f"download_q_{k}",
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Daftar prioritas promosi</div>', unsafe_allow_html=True)
    show_cols = ["NIP", "Nama", "Satker", "Pangkat", "Sublevel", "QScore", "MDP_Tahun", "Kuadran", "Readiness"]
    sorted_visible = visible[show_cols].sort_values(["Kuadran", "QScore"], ascending=[True, False])
    st.dataframe(sorted_visible, use_container_width=True, hide_index=True, height=340)
    st.download_button(
        "Unduh laporan terfilter (.xlsx)",
        data=xlsx_bytes(sorted_visible),
        file_name="sikabi_laporan_kuadran_terfilter.xlsx",
        use_container_width=False,
        key="download_filtered_report",
    )


def page_pegawai(data):
    st.markdown('<div class="sikabi-header"><div class="eyebrow">EMPLOYEE INTELLIGENCE</div><div class="page-title">Data Pegawai</div><div class="page-subtitle">Eksplorasi data pegawai dan rincian pembentukan skor secara transparan.</div></div>', unsafe_allow_html=True)

    a, b, c, d = st.columns([2, 1.15, 1.45, 1.25])
    q = a.text_input("Cari NIP / nama", placeholder="Ketik NIP atau nama...")
    fp = b.multiselect("Pangkat", PANGKAT_OPTS, default=PANGKAT_OPTS, key="emp_pangkat")
    fs = c.multiselect("Satuan kerja", SATKER_OPTS, default=SATKER_OPTS, key="emp_satker")
    fst = d.multiselect("Status akhir", STATUS_OPTS, default=STATUS_OPTS, key="emp_status")

    view = data.copy()
    if q:
        view = view[safe_contains(view["Nama"], q) | safe_contains(view["NIP"], q)]
    view = apply_multiselect_filter(view, "Pangkat", fp)
    view = apply_multiselect_filter(view, "Satker", fs)
    view = apply_multiselect_filter(view, "Status_Akhir", fst)

    st.caption(f"Menampilkan {len(view):,} dari {len(data):,} pegawai")
    tab_ringkas, tab_lengkap = st.tabs(["Ringkas", "Lengkap"])

    ringkas_cols = ["NIP", "Nama", "Satker", "Pangkat", "Sublevel", "Quantitative_Score", "Qualitative_Score", "QScore", "Status_Akhir", "Kuadran"]
    lengkap_cols = ringkas_cols[:5] + [
        "NK_1", "NK_2", "NK_3", "NK_4", "NK_5", "NK_Mean", "Skor_NK",
        "MDG_Tahun", "MDP_Tahun", "Skor_MDP", "Pendidikan", "Skor_Pendidikan",
        "Sertifikasi", "Skor_Sertifikasi", "Quantitative_Score", "Exposure", "Potensi",
        "K3", "Skor_K3", "Qualitative_Score", "QScore", "Status_Akhir", "Kuadran",
    ]
    with tab_ringkas:
        st.dataframe(view[ringkas_cols].sort_values("QScore", ascending=False), use_container_width=True, hide_index=True, height=430)
    with tab_lengkap:
        st.dataframe(view[lengkap_cols].sort_values("QScore", ascending=False), use_container_width=True, hide_index=True, height=430)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Detail & breakdown skor</div>', unsafe_allow_html=True)
    if len(view):
        pilihan = view.copy()
        pilihan["_label"] = pilihan["NIP"].astype(str) + " — " + pilihan["Nama"]
        label = st.selectbox("Pilih pegawai", pilihan["_label"].tolist(), key="employee_detail")
        row = pilihan[pilihan["_label"] == label].iloc[0]

        d1, d2, d3 = st.columns(3)
        with d1:
            st.markdown("**Data dasar**")
            st.write(f"NIP: {row['NIP']}")
            st.write(f"Satker: {row['Satker']}")
            st.write(f"Pangkat: {row['Pangkat']} ({row['Sublevel']})")
            st.write(f"Masa dinas grade: {row['MDG_Tahun']} tahun")
            st.write(f"Sisa masa dinas: {row['Remaining_Service']} tahun")
        with d2:
            st.markdown("**Quantitative**")
            st.write(f"NK rata-rata 5 th: {row['NK_Mean']} → skor {row['Skor_NK']}")
            st.write(f"MDP: {row['MDP_Tahun']} th → skor {row['Skor_MDP']}")
            st.write(f"Pendidikan: {row['Pendidikan']} → skor {row['Skor_Pendidikan']}")
            st.write(f"Sertifikasi: {row['Sertifikasi']} → skor {row['Skor_Sertifikasi']}")
            st.write(f"**Quantitative Score: {row['Quantitative_Score']}**")
        with d3:
            st.markdown("**Qualitative**")
            st.write(f"Exposure: {row['Exposure']}")
            st.write(f"Potensi: {row['Potensi']}")
            st.write(f"K3: {row['K3']} → skor {row['Skor_K3']}")
            st.write(f"**Qualitative Score: {row['Qualitative_Score']}**")
            st.write(f"**QScore final: {row['QScore']}**")

        st.info(
            f"Gate Administrasi: {'Lolos' if row['Lolos_Administrasi'] else 'Tidak lolos'} · "
            f"Gate KPP: {'Lolos' if row['Lolos_KPP'] else 'Tidak lolos'} · "
            f"Status akhir: **{row['Status_Akhir']}**" +
            (f" · Kuadran **{row['Kuadran']}**" if row['Kuadran'] else "")
        )
    else:
        st.info("Tidak ada pegawai yang sesuai dengan filter.")


def page_gate(data):
    st.markdown('<div class="sikabi-header"><div class="eyebrow">DECISION GATE</div><div class="page-title">Gate Keputusan</div><div class="page-subtitle">Tahapan penyaringan kandidat sebelum masuk proses KPP.</div></div>', unsafe_allow_html=True)

    a, b, c = st.columns([1.2, 1.2, 2])
    fp = a.multiselect("Pangkat", PANGKAT_OPTS, default=PANGKAT_OPTS, key="gate_pangkat")
    fs = b.multiselect("Status akhir", STATUS_OPTS, default=STATUS_OPTS, key="gate_status")
    st.caption("Filter berlaku ke seluruh sub-section gate.")

    view = apply_multiselect_filter(data.copy(), "Pangkat", fp)
    view = apply_multiselect_filter(view, "Status_Akhir", fs)

    # Sub-section tanpa nomor, sesuai permintaan.
    st.markdown('<div class="gate-subsection"><div class="gate-heading">Syarat Administrasi</div><div class="muted">Penyaringan persyaratan dasar sebelum evaluasi KPP.</div></div>', unsafe_allow_html=True)
    cols = ["Nama", "Pangkat"] + list(ADMIN_CHECK_LABELS.keys()) + ["Lolos_Administrasi"]
    show = view[cols].rename(columns={**ADMIN_CHECK_LABELS, "Lolos_Administrasi": "LOLOS ADMINISTRASI"})
    st.dataframe(show, use_container_width=True, hide_index=True, height=300)

    st.markdown('<div class="gate-subsection"><div class="gate-heading">Kriteria KPP</div><div class="muted">Pembanding quantitative, qualitative, dan QScore terhadap passing grade per pangkat.</div></div>', unsafe_allow_html=True)
    cols = ["Nama", "Pangkat", "Passing_Grade_QScore", "QScore"] + list(KPP_CHECK_LABELS.keys()) + ["Lolos_KPP"]
    show = view[cols].rename(columns={**KPP_CHECK_LABELS, "Lolos_KPP": "LOLOS KPP"})
    st.dataframe(show, use_container_width=True, hide_index=True, height=300)

    st.markdown('<div class="gate-subsection"><div class="gate-heading">Syarat KPP Grade Senior / MDG</div><div class="muted">Routing akhir berdasarkan status Senior atau terpenuhinya threshold MDG.</div></div>', unsafe_allow_html=True)
    cols = ["Nama", "Pangkat", "Sublevel", "Is_Senior", "MDG_Tahun", "Chk_MDG_Terpenuhi", "Status_Akhir"]
    show = view[cols].rename(columns={"Is_Senior": "Grade Senior?", "Chk_MDG_Terpenuhi": "MDG Terpenuhi?"})
    st.dataframe(show, use_container_width=True, hide_index=True, height=300)


def crud_section(sheet, key_col, value_cols, value_step=0.01, value_format="%.4f"):
    df_ref = st.session_state[f"ref_{sheet}"]
    st.markdown(f"**{sheet}**")
    edited = st.data_editor(
        df_ref,
        num_rows="fixed",
        disabled=[key_col],
        use_container_width=True,
        hide_index=True,
        key=f"editor_{sheet}",
    )
    if st.button("Simpan perubahan", key=f"save_{sheet}"):
        st.session_state[f"ref_{sheet}"] = edited
        save_to_disk()
        st.success(f"Perubahan {sheet} disimpan.")
        st.rerun()

    add_col, del_col = st.columns(2, gap="large")
    with add_col:
        st.markdown("**Tambah baris**")
        with st.form(f"add_{sheet}", clear_on_submit=True):
            new_key = st.text_input(key_col)
            new_values = {}
            for vc in value_cols:
                new_values[vc] = st.number_input(vc, step=value_step, format=value_format, key=f"newval_{sheet}_{vc}")
            submitted = st.form_submit_button("Tambah baris")
            if submitted:
                if not new_key.strip():
                    st.error(f"{key_col} tidak boleh kosong.")
                elif new_key in df_ref[key_col].astype(str).values:
                    st.error(f"'{new_key}' sudah ada di {key_col}.")
                else:
                    new_row = {key_col: new_key, **new_values}
                    st.session_state[f"ref_{sheet}"] = pd.concat([df_ref, pd.DataFrame([new_row])], ignore_index=True)
                    save_to_disk()
                    st.success(f"Baris '{new_key}' ditambahkan.")
                    st.rerun()

    with del_col:
        st.markdown("**Hapus baris**")
        to_delete = st.multiselect(f"Pilih {key_col} yang mau dihapus", df_ref[key_col].astype(str).tolist(), key=f"del_{sheet}")
        if st.button("Hapus baris terpilih", key=f"delbtn_{sheet}", disabled=len(to_delete) == 0):
            st.session_state[f"ref_{sheet}"] = df_ref[~df_ref[key_col].astype(str).isin(to_delete)].reset_index(drop=True)
            save_to_disk()
            st.success(f"{len(to_delete)} baris dihapus.")
            st.rerun()


def page_master():
    st.markdown('<div class="sikabi-header"><div class="eyebrow">CONFIGURATION</div><div class="page-title">Master Data</div><div class="page-subtitle">Bobot dan tabel referensi yang digunakan mesin scoring.</div></div>', unsafe_allow_html=True)
    st.caption("Perubahan pada master data langsung memengaruhi hasil scoring dan gate setelah disimpan.")
    tabs = st.tabs([
        "Bobot Quantitative", "Bobot Qualitative", "Bobot per Pangkat",
        "Skor Pendidikan", "Skor Sertifikasi", "Skor K3", "Threshold",
    ])
    with tabs[0]:
        crud_section("Quant_Weights", "Parameter", ["Value"])
    with tabs[1]:
        crud_section("Qual_Weights", "Parameter", ["Value"])
    with tabs[2]:
        crud_section("Rank_Weights", "Pangkat", ["Quantitative", "Qualitative"])
    with tabs[3]:
        crud_section("Education_Score", "Kategori", ["Nilai"], value_step=1.0, value_format="%.0f")
    with tabs[4]:
        crud_section("Certification_Score", "Kategori", ["Nilai"], value_step=1.0, value_format="%.0f")
    with tabs[5]:
        crud_section("K3_Score", "Kategori", ["Nilai"], value_step=1.0, value_format="%.0f")
    with tabs[6]:
        crud_section("Thresholds", "Parameter", ["Value"])


with tab_dashboard:
    page_dashboard(df)
with tab_pegawai:
    page_pegawai(df)
with tab_gate:
    page_gate(df)
with tab_master:
    page_master()
