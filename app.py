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

from scoring import (
    ADMIN_CHECK_LABELS,
    KPP_CHECK_LABELS,
    compute_all,
)

DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "sikabi_data.xlsx")
REF_SHEETS = [
    "Quant_Weights", "Qual_Weights", "Rank_Weights",
    "Education_Score", "Certification_Score", "K3_Score", "Thresholds",
]

KUADRAN_COLOR = {"I": "#3E7C74", "II": "#7B9E3F", "III": "#C08A2E", "IV": "#B0574A"}

st.set_page_config(page_title="SIKABI", page_icon="🏦", layout="wide")


# ----------------------------------------------------------------------------
# State & data loading
# ----------------------------------------------------------------------------
def load_workbook(path):
    sheets = pd.read_excel(path, sheet_name=None)
    return sheets


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


init_state()
df = recompute()

PANGKAT_OPTS = sorted(df["Pangkat"].dropna().unique().tolist())
SATKER_OPTS = sorted(df["Satker"].dropna().unique().tolist())


# ----------------------------------------------------------------------------
# Sidebar navigation
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🏦 SIKABI")
    st.caption("Sistem Intelijen Karier Bank Indonesia")
    page = st.radio(
        "Menu",
        ["Dashboard (Kuadran & Laporan)", "Data Pegawai", "Gate Keputusan", "Master Data"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption(f"Total populasi: **{len(df)}** pegawai")
    st.caption(f"Lolos administrasi: **{int(df['Lolos_Administrasi'].sum())}**")
    st.caption(f"Proses KPP: **{int(df['Masuk_Proses_KPP'].sum())}**")
    st.divider()
    st.download_button(
        "⬇️ Unduh data (.xlsx)",
        data=workbook_bytes(),
        file_name="sikabi_data.xlsx",
        use_container_width=True,
    )


# ----------------------------------------------------------------------------
# Page: Dashboard (Kuadran & Laporan)
# ----------------------------------------------------------------------------
def page_dashboard(df):
    st.subheader("Dashboard — Kuadran & Laporan")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total populasi", len(df))
    c2.metric("Lolos administrasi", int(df["Lolos_Administrasi"].sum()))
    c3.metric("Lolos kriteria KPP", int(df["Lolos_KPP"].sum()))
    c4.metric("General talent", int((df["Status_Akhir"] == "General Talent").sum()))

    st.divider()

    fcol1, fcol2, fcol3 = st.columns([1, 1, 2])
    f_pangkat = fcol1.multiselect("Pangkat", PANGKAT_OPTS, default=PANGKAT_OPTS)
    f_satker = fcol2.multiselect("Satuan kerja", SATKER_OPTS, default=SATKER_OPTS)
    f_kuadran = fcol3.multiselect("Kuadran", ["I", "II", "III", "IV"], default=["I", "II", "III", "IV"])

    kpp_pop = df[df["Masuk_Proses_KPP"]]
    scoped = kpp_pop[kpp_pop["Pangkat"].isin(f_pangkat) & kpp_pop["Satker"].isin(f_satker)]
    visible = scoped[scoped["Kuadran"].isin(f_kuadran)]

    left, right = st.columns([3, 2])
    with left:
        st.markdown(f"**Peta Kuadran** — {len(visible)} dari {len(kpp_pop)} pegawai Proses KPP")
        if len(visible) > 0:
            fig = px.scatter(
                visible,
                x="MDP_Tahun", y="QScore", color="Kuadran",
                color_discrete_map=KUADRAN_COLOR,
                hover_data=["Nama", "Satker", "Pangkat", "Sublevel"],
                labels={"MDP_Tahun": "Masa Dinas Pangkat (tahun)", "QScore": "QScore"},
            )
            mean_q = scoped["QScore"].mean()
            mean_m = scoped["MDP_Tahun"].mean()
            fig.add_hline(y=mean_q, line_dash="dash", line_color="gray")
            fig.add_vline(x=mean_m, line_dash="dash", line_color="gray")
            fig.update_layout(height=420, legend_title_text="Kuadran")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Tidak ada data pada kombinasi filter ini.")

    with right:
        st.markdown("**Ringkasan per kuadran**")
        for k in ["I", "II", "III", "IV"]:
            count = scoped[scoped["Kuadran"] == k].shape[0]
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;padding:8px 12px;"
                f"border:1px solid #E4E7EA;border-radius:8px;margin-bottom:6px;"
                f"opacity:{1 if k in f_kuadran else 0.4}'>"
                f"<span><span style='color:{KUADRAN_COLOR[k]}'>⬤</span> Kuadran {k}</span>"
                f"<b>{count}</b></div>",
                unsafe_allow_html=True,
            )

    st.divider()
    st.markdown(f"**Daftar prioritas promosi** ({len(visible)} pegawai)")
    show_cols = ["Nama", "Satker", "Pangkat", "Sublevel", "QScore", "MDP_Tahun", "Kuadran", "Readiness"]
    sorted_visible = visible[show_cols].sort_values("QScore", ascending=False)
    st.dataframe(sorted_visible, use_container_width=True, hide_index=True, height=320)

    buf = io.BytesIO()
    sorted_visible.to_excel(buf, index=False)
    st.download_button("⬇️ Export laporan (.xlsx)", data=buf.getvalue(), file_name="laporan_kuadran.xlsx")


# ----------------------------------------------------------------------------
# Page: Data Pegawai (+ scoring breakdown, merged)
# ----------------------------------------------------------------------------
def page_pegawai(df):
    st.subheader("Data Pegawai & Rincian Penilaian")
    st.caption("Setiap kolom skor di tabel ini bisa ditelusuri dari kolom mentah di sebelah kanannya — supaya transparan.")

    fcol1, fcol2, fcol3, fcol4 = st.columns([2, 1, 1, 1])
    q = fcol1.text_input("Cari NIP / nama")
    f_pangkat = fcol2.selectbox("Pangkat", ["Semua"] + PANGKAT_OPTS)
    f_satker = fcol3.selectbox("Satuan kerja", ["Semua"] + SATKER_OPTS)
    f_status = fcol4.selectbox("Status akhir", ["Semua", "Proses KPP", "General Talent"])

    view = df.copy()
    if q:
        view = view[view["Nama"].str.contains(q, case=False) | view["NIP"].astype(str).str.contains(q)]
    if f_pangkat != "Semua":
        view = view[view["Pangkat"] == f_pangkat]
    if f_satker != "Semua":
        view = view[view["Satker"] == f_satker]
    if f_status != "Semua":
        view = view[view["Status_Akhir"] == f_status]

    tab_ringkas, tab_lengkap = st.tabs(["Tampilan ringkas", "Tampilan lengkap (semua kolom skor)"])

    ringkas_cols = [
        "NIP", "Nama", "Satker", "Pangkat", "Sublevel",
        "Quantitative_Score", "Qualitative_Score", "QScore", "Status_Akhir", "Kuadran",
    ]
    lengkap_cols = ringkas_cols[:5] + [
        "NK_1", "NK_2", "NK_3", "NK_4", "NK_5", "NK_Mean", "Skor_NK",
        "MDG_Tahun", "Skor_MDP",
        "Pendidikan", "Skor_Pendidikan",
        "Sertifikasi", "Skor_Sertifikasi",
        "Quantitative_Score",
        "Exposure", "Potensi", "K3", "Skor_K3",
        "Qualitative_Score", "QScore", "Status_Akhir", "Kuadran",
    ]

    with tab_ringkas:
        st.dataframe(
            view[ringkas_cols].sort_values("QScore", ascending=False),
            use_container_width=True, hide_index=True, height=420,
        )
    with tab_lengkap:
        st.dataframe(
            view[lengkap_cols].sort_values("QScore", ascending=False),
            use_container_width=True, hide_index=True, height=420,
        )

    st.caption(f"Menampilkan {len(view)} dari {len(df)} pegawai")

    st.divider()
    st.markdown("#### Detail & breakdown skor per pegawai")
    view = view.copy()
    view["_label"] = view["NIP"].astype(str) + " — " + view["Nama"]
    label_pilihan = st.selectbox("Pilih pegawai", view["_label"].tolist() if len(view) else [])
    if label_pilihan:
        row = view[view["_label"] == label_pilihan].iloc[0]
        d1, d2, d3 = st.columns(3)
        with d1:
            st.markdown("**Data dasar**")
            st.write(f"NIP: {row['NIP']}")
            st.write(f"Satker: {row['Satker']}")
            st.write(f"Pangkat: {row['Pangkat']} ({row['Sublevel']})")
            st.write(f"Masa dinas grade: {row['MDG_Tahun']} tahun")
            st.write(f"Sisa masa dinas: {row['Remaining_Service']} tahun")
        with d2:
            st.markdown("**Komponen Quantitative** (bobot pangkat berlaku)")
            st.write(f"NK rata-rata 5 th: {row['NK_Mean']} → skor {row['Skor_NK']}")
            st.write(f"MDP: {row['MDP_Tahun']} th → skor {row['Skor_MDP']}")
            st.write(f"Pendidikan: {row['Pendidikan']} → skor {row['Skor_Pendidikan']}")
            st.write(f"Sertifikasi: {row['Sertifikasi']} → skor {row['Skor_Sertifikasi']}")
            st.write(f"**Quantitative Score: {row['Quantitative_Score']}**")
        with d3:
            st.markdown("**Komponen Qualitative**")
            st.write(f"Exposure: {row['Exposure']}")
            st.write(f"Potensi: {row['Potensi']}")
            st.write(f"K3: {row['K3']} → skor {row['Skor_K3']}")
            st.write(f"**Qualitative Score: {row['Qualitative_Score']}**")
            st.write(f"**QScore final: {row['QScore']}**")

        badge_admin = "✅ Lolos" if row["Lolos_Administrasi"] else "❌ Tidak lolos"
        badge_kpp = "✅ Lolos" if row["Lolos_KPP"] else "❌ Tidak lolos"
        st.info(f"Gate Administrasi: {badge_admin} · Gate KPP: {badge_kpp} · Status akhir: **{row['Status_Akhir']}**" + (f" · Kuadran **{row['Kuadran']}**" if row["Kuadran"] else ""))


# ----------------------------------------------------------------------------
# Page: Gate Keputusan
# ----------------------------------------------------------------------------
def page_gate(df):
    st.subheader("Gate Keputusan BI Wide")

    tab1, tab2, tab3 = st.tabs(["1. Syarat Administrasi", "2. Kriteria KPP", "3. Grade Senior / MDG"])

    fcol1, fcol2 = st.columns(2)
    f_pangkat = fcol1.selectbox("Filter pangkat", ["Semua"] + PANGKAT_OPTS, key="gate_pangkat")
    f_status = fcol2.selectbox("Filter status akhir", ["Semua", "Proses KPP", "General Talent"], key="gate_status")

    view = df.copy()
    if f_pangkat != "Semua":
        view = view[view["Pangkat"] == f_pangkat]
    if f_status != "Semua":
        view = view[view["Status_Akhir"] == f_status]

    with tab1:
        cols = ["Nama", "Pangkat"] + list(ADMIN_CHECK_LABELS.keys()) + ["Lolos_Administrasi"]
        show = view[cols].rename(columns={**ADMIN_CHECK_LABELS, "Lolos_Administrasi": "LOLOS ADMINISTRASI"})
        st.dataframe(show, use_container_width=True, hide_index=True, height=420)

    with tab2:
        cols = ["Nama", "Pangkat", "Passing_Grade_QScore", "QScore"] + list(KPP_CHECK_LABELS.keys()) + ["Lolos_KPP"]
        show = view[cols].rename(columns={**KPP_CHECK_LABELS, "Lolos_KPP": "LOLOS KPP"})
        st.dataframe(show, use_container_width=True, hide_index=True, height=420)

    with tab3:
        cols = ["Nama", "Pangkat", "Sublevel", "Is_Senior", "MDG_Tahun", "Chk_MDG_Terpenuhi", "Status_Akhir"]
        show = view[cols].rename(columns={
            "Is_Senior": "Grade Senior?",
            "Chk_MDG_Terpenuhi": "MDG Terpenuhi?",
        })
        st.dataframe(show, use_container_width=True, hide_index=True, height=420)


# ----------------------------------------------------------------------------
# Page: Master Data (CRUD)
# ----------------------------------------------------------------------------
def page_master():
    st.subheader("Master Data")
    st.caption("Ubah, tambah, atau hapus baris langsung di tabel. Klik **Simpan Perubahan** untuk menerapkan ke seluruh perhitungan skor.")

    tabs = st.tabs([
        "Bobot Quantitative", "Bobot Qualitative", "Bobot per Pangkat",
        "Skor Pendidikan", "Skor Sertifikasi", "Skor K3", "Threshold Administrasi",
    ])
    sheet_names = [
        "Quant_Weights", "Qual_Weights", "Rank_Weights",
        "Education_Score", "Certification_Score", "K3_Score", "Thresholds",
    ]

    for tab, sheet in zip(tabs, sheet_names):
        with tab:
            key = f"editor_{sheet}"
            edited = st.data_editor(
                st.session_state[f"ref_{sheet}"],
                num_rows="dynamic",
                use_container_width=True,
                key=key,
            )
            colA, colB = st.columns([1, 4])
            if colA.button("💾 Simpan perubahan", key=f"save_{sheet}"):
                st.session_state[f"ref_{sheet}"] = edited
                save_to_disk()
                st.success(f"Perubahan pada {sheet} disimpan. Skor akan dihitung ulang.")
                st.rerun()

    st.divider()
    with st.expander("⚙️ Kelola data pegawai (tambah / edit / hapus)"):
        edited_emp = st.data_editor(
            st.session_state.employees,
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            key="editor_employees",
        )
        if st.button("💾 Simpan data pegawai"):
            st.session_state.employees = edited_emp
            save_to_disk()
            st.success("Data pegawai disimpan. Skor akan dihitung ulang.")
            st.rerun()


# ----------------------------------------------------------------------------
# Router
# ----------------------------------------------------------------------------
if page == "Dashboard (Kuadran & Laporan)":
    page_dashboard(df)
elif page == "Data Pegawai":
    page_pegawai(df)
elif page == "Gate Keputusan":
    page_gate(df)
elif page == "Master Data":
    page_master()
