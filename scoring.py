"""
SIKABI - Scoring Engine

Modul ini murni fungsi (tanpa Streamlit) supaya mudah ditest terpisah.
Semua perhitungan skor & gate keputusan mengikuti dokumen System Requirements
SIKABI: Quantitative Score, Qualitative Score, QScore, Gate Administrasi,
Gate Kriteria KPP, Grade Senior/MDG, dan Kuadran.
"""

from datetime import datetime

import numpy as np
import pandas as pd

TODAY = datetime(2026, 9, 14)


def years_since(date_val, ref=TODAY):
    if pd.isna(date_val):
        return np.nan
    return (ref - pd.to_datetime(date_val)).days / 365.25


def zscore_band(value, mean, std):
    """Mengembalikan skor 100/60/20 berdasarkan posisi value terhadap mean +/- 1 stdev."""
    if std == 0 or pd.isna(std):
        return 60
    if value > mean + std:
        return 100
    if value < mean - std:
        return 20
    return 60


def load_reference_tables(xls_path_or_buffer):
    """Membaca semua sheet referensi dari file Excel menjadi dict of DataFrame."""
    sheets = pd.read_excel(xls_path_or_buffer, sheet_name=None)
    return sheets


def compute_all(employees: pd.DataFrame, refs: dict) -> pd.DataFrame:
    """
    Input:
      employees : DataFrame mentah dari sheet Employees
      refs      : dict berisi DataFrame Quant_Weights, Qual_Weights, Rank_Weights,
                  Education_Score, Certification_Score, K3_Score, Thresholds
    Output:
      DataFrame employees + seluruh kolom turunan (skor komponen, gate, kuadran)
    """
    df = employees.copy()

    quant_w = dict(zip(refs["Quant_Weights"]["Parameter"], refs["Quant_Weights"]["Value"]))
    qual_w = dict(zip(refs["Qual_Weights"]["Parameter"], refs["Qual_Weights"]["Value"]))
    rank_w = refs["Rank_Weights"].set_index("Pangkat")
    edu_map = dict(zip(refs["Education_Score"]["Kategori"], refs["Education_Score"]["Nilai"]))
    cert_map = dict(zip(refs["Certification_Score"]["Kategori"], refs["Certification_Score"]["Nilai"]))
    k3_map = dict(zip(refs["K3_Score"]["Kategori"], refs["K3_Score"]["Nilai"]))
    thr = dict(zip(refs["Thresholds"]["Parameter"], refs["Thresholds"]["Value"]))

    # ---------- Turunan dasar ----------
    nk_cols = [c for c in df.columns if c.startswith("NK_")]
    df["NK_Mean"] = df[nk_cols].mean(axis=1).round(2)
    df["MDG_Tahun"] = df["Tanggal_Grade"].apply(years_since).round(2)
    df["MDP_Tahun"] = df["MDG_Tahun"]  # satu-satunya tanggal acuan yang tersedia

    # ---------- Skor komponen Quantitative ----------
    nk_mean_pop, nk_std_pop = df["NK_Mean"].mean(), df["NK_Mean"].std()
    mdp_mean_pop, mdp_std_pop = df["MDP_Tahun"].mean(), df["MDP_Tahun"].std()

    df["Skor_NK"] = df["NK_Mean"].apply(lambda v: zscore_band(v, nk_mean_pop, nk_std_pop))
    df["Skor_MDP"] = df["MDP_Tahun"].apply(lambda v: zscore_band(v, mdp_mean_pop, mdp_std_pop))
    df["Skor_Pendidikan"] = df["Pendidikan"].map(edu_map).fillna(0)
    df["Skor_Sertifikasi"] = df["Sertifikasi"].map(cert_map).fillna(0)

    df["Quantitative_Score"] = (
        df["Skor_NK"] * quant_w.get("NK", 0)
        + df["Skor_MDP"] * quant_w.get("MDP", 0)
        + df["Skor_Pendidikan"] * quant_w.get("Pendidikan", 0)
        + df["Skor_Sertifikasi"] * quant_w.get("Sertifikasi", 0)
    ).round(1)

    # ---------- Skor komponen Qualitative ----------
    df["Skor_K3"] = df["K3"].map(k3_map).fillna(0)
    df["Qualitative_Score"] = (
        df["Exposure"] * qual_w.get("Exposure", 0)
        + df["Potensi"] * qual_w.get("Potensi", 0)
        + df["Skor_K3"] * qual_w.get("K3", 0)
    ).round(1)

    # ---------- QScore final (bobot per pangkat) ----------
    def qscore_row(row):
        w = rank_w.loc[row["Pangkat"]] if row["Pangkat"] in rank_w.index else None
        if w is None:
            return np.nan
        return round(row["Quantitative_Score"] * w["Quantitative"] + row["Qualitative_Score"] * w["Qualitative"], 1)

    df["QScore"] = df.apply(qscore_row, axis=1)

    # ---------- Gate 1: Syarat Administrasi ----------
    min_nk = thr.get("Minimum NK", 3.0)
    min_sisa = thr.get("Minimum Remaining Service (tahun)", 0.5)
    mdg_threshold = thr.get("MDG Threshold (tahun)", 2)

    def pendidikan_ok(row):
        if "Officer" not in row["Pendidikan"] and "Non-Officer" not in row["Pendidikan"]:
            return True  # S2/S3 dianggap selalu lolos syarat minimal
        if "Non-Officer" in row["Pendidikan"]:
            return True  # sudah minimal D3
        return True  # S1 Officer juga lolos (>= S1)

    df["Chk_NK"] = df["NK_Mean"] >= min_nk
    df["Chk_SisaDinas"] = df["Remaining_Service"] > min_sisa
    df["Chk_Pendidikan"] = df.apply(pendidikan_ok, axis=1)
    df["Chk_Rekomendasi"] = df["Satker_Recommendation"] == "Direkomendasikan"
    df["Chk_StatusAktif"] = df["Status"] == "Aktif"
    df["Chk_PTB_S2"] = ~df["PTB_S2"].astype(bool)
    df["Chk_PromosiAward"] = ~df["Promotion_Award"].astype(bool)

    admin_checks = [
        "Chk_NK", "Chk_SisaDinas", "Chk_Pendidikan", "Chk_Rekomendasi",
        "Chk_StatusAktif", "Chk_PTB_S2", "Chk_PromosiAward",
    ]
    df["Lolos_Administrasi"] = df[admin_checks].all(axis=1)

    # ---------- Gate 2: Kriteria KPP (passing grade = mean QScore per pangkat, dari populasi lolos administrasi) ----------
    lolos_admin_df = df[df["Lolos_Administrasi"]]
    passing_grade_map = lolos_admin_df.groupby("Pangkat")["QScore"].mean().round(1).to_dict()
    quant_pg_map = lolos_admin_df.groupby("Pangkat")["Quantitative_Score"].mean().round(1).to_dict()
    qual_pg_map = lolos_admin_df.groupby("Pangkat")["Qualitative_Score"].mean().round(1).to_dict()

    df["Passing_Grade_QScore"] = df["Pangkat"].map(passing_grade_map)
    df["Passing_Grade_Quant"] = df["Pangkat"].map(quant_pg_map)
    df["Passing_Grade_Qual"] = df["Pangkat"].map(qual_pg_map)

    df["Chk_Kinerja"] = df["NK_Mean"] >= min_nk
    df["Chk_Quant_PG"] = df["Quantitative_Score"] >= df["Passing_Grade_Quant"]
    df["Chk_Qual_PG"] = df["Qualitative_Score"] >= df["Passing_Grade_Qual"]
    df["Chk_QScore_PG"] = df["QScore"] >= df["Passing_Grade_QScore"]

    kpp_checks = ["Chk_Kinerja", "Chk_Quant_PG", "Chk_Qual_PG", "Chk_QScore_PG"]
    df["Lolos_KPP"] = df["Lolos_Administrasi"] & df[kpp_checks].all(axis=1)

    # ---------- Gate 3: Grade Senior / MDG ----------
    df["Is_Senior"] = df["Sublevel"] == "Senior"
    df["Chk_MDG_Terpenuhi"] = df["MDG_Tahun"] >= mdg_threshold

    def routing(row):
        if not row["Lolos_KPP"]:
            return "General Talent"
        if row["Is_Senior"]:
            return "Proses KPP"
        if row["Chk_MDG_Terpenuhi"]:
            return "Proses KPP"
        return "General Talent"

    df["Status_Akhir"] = df.apply(routing, axis=1)
    df["Masuk_Proses_KPP"] = df["Status_Akhir"] == "Proses KPP"

    # ---------- Kuadran (hanya untuk populasi Proses KPP, mean dihitung per pangkat) ----------
    kpp_pop = df[df["Masuk_Proses_KPP"]]
    mean_q_map = kpp_pop.groupby("Pangkat")["QScore"].mean().to_dict()
    mean_m_map = kpp_pop.groupby("Pangkat")["MDP_Tahun"].mean().to_dict()

    def kuadran_row(row):
        if not row["Masuk_Proses_KPP"]:
            return None
        mq = mean_q_map.get(row["Pangkat"], row["QScore"])
        mm = mean_m_map.get(row["Pangkat"], row["MDP_Tahun"])
        if row["QScore"] > mq and row["MDP_Tahun"] > mm:
            return "I"
        if row["QScore"] > mq and row["MDP_Tahun"] <= mm:
            return "II"
        if row["QScore"] <= mq and row["MDP_Tahun"] > mm:
            return "III"
        return "IV"

    df["Mean_QScore_Pangkat_KPP"] = df["Pangkat"].map(mean_q_map).round(1)
    df["Mean_MDP_Pangkat_KPP"] = df["Pangkat"].map(mean_m_map).round(1)
    df["Kuadran"] = df.apply(kuadran_row, axis=1)

    return df


ADMIN_CHECK_LABELS = {
    "Chk_NK": "Nilai Kinerja rata-rata >= 3.00",
    "Chk_SisaDinas": "Sisa masa dinas > 6 bulan",
    "Chk_Pendidikan": "Pendidikan minimal sesuai kategori jabatan",
    "Chk_Rekomendasi": "Direkomendasikan oleh Satuan Kerja",
    "Chk_StatusAktif": "Status kepegawaian aktif",
    "Chk_PTB_S2": "Tidak sedang PTB S2",
    "Chk_PromosiAward": "Tidak pernah promosi pangkat penghargaan",
}

KPP_CHECK_LABELS = {
    "Chk_Kinerja": "Kinerja: NK rata-rata >= 3.00",
    "Chk_Quant_PG": "Quantitative Score >= Passing Grade",
    "Chk_Qual_PG": "Qualitative Score >= Passing Grade",
    "Chk_QScore_PG": "QScore >= Passing Grade (Mean QScore per pangkat)",
}
