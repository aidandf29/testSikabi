"""
SIKABI - Data Generator
Membuat file data/sikabi_data.xlsx berisi 500 data dummy pegawai
beserta seluruh tabel referensi (master data) yang dipakai mesin skoring.

Jalankan: python data_gen.py
"""

import random
from datetime import datetime, timedelta

import pandas as pd

random.seed(42)

N_EMPLOYEES = 500

FIRST_NAMES = [
    "Andi", "Bima", "Citra", "Dimas", "Eka", "Farhan", "Gita", "Hana", "Irfan", "Joko",
    "Kirana", "Luthfi", "Maya", "Nadia", "Oscar", "Putri", "Raka", "Sinta", "Taufik", "Vina",
    "Wulan", "Yusuf", "Zahra", "Bagas", "Cahyo", "Dewi", "Erlangga", "Fitri", "Galih", "Hesti",
    "Ilham", "Julia", "Krisna", "Lestari", "Miko", "Nurul", "Omar", "Prita", "Rian", "Sari",
]
LAST_NAMES = [
    "Pratama", "Wijaya", "Lestari", "Saputra", "Putri", "Akbar", "Maharani", "Salsabila",
    "Ramadhan", "Santoso", "Ayu", "Hakim", "Anindita", "Permata", "Amelia", "Nugraha",
    "Kartika", "Hidayat", "Firmansyah", "Utami", "Gunawan", "Rahayu", "Handayani", "Setiawan",
]

SATKER_UNIT = {
    "DMST": "Direktorat DMST",
    "DSDMM": "Direktorat DSDMM",
    "DKMP": "Direktorat DKMP",
    "DEIH": "Direktorat DEIH",
    "DMR": "Direktorat DMR",
    "DKEM": "Direktorat DKEM",
    "DPSP": "Direktorat DPSP",
    "DKOM": "Direktorat DKOM",
}
SATKERS = list(SATKER_UNIT.keys())

PANGKAT_ORDER = ["DD", "AD", "M", "AM", "S-A"]

PENDIDIKAN_OPTS = ["S3", "S2 A/PTB", "S2 lainnya", "S1 Officer", "D3 Non-Officer"]
PENDIDIKAN_WEIGHTS = [0.08, 0.16, 0.18, 0.40, 0.18]

SERTIFIKASI_OPTS = ["Tier 1", "Tier 2", "PMK", "Kum Mengajar", "ELP", "None"]
SERTIFIKASI_WEIGHTS = [0.14, 0.18, 0.14, 0.12, 0.12, 0.30]

K3_OPTS = ["SB", "B", "CB", "KB"]
K3_WEIGHTS = [0.28, 0.36, 0.24, 0.12]

READINESS_OPTS = ["Ready Now", "Ready 1-2 Tahun", "Belum Siap"]


def random_date_grade(pangkat):
    """Semakin tinggi pangkat, cenderung tanggal grade lebih lama (semakin lama menjabat)."""
    base_years_ago = {"DD": 4, "AD": 4, "M": 3, "AM": 3, "S-A": 6}[pangkat]
    days_ago = int(random.uniform(0.2, base_years_ago) * 365)
    return datetime(2026, 9, 14) - timedelta(days=days_ago)


def gen_employee(i):
    pangkat = PANGKAT_ORDER[i % len(PANGKAT_ORDER)]
    sublevel = "Reguler"
    if pangkat != "S-A" and random.random() < 0.28:
        sublevel = "Senior"

    satker = random.choice(SATKERS)
    nama = f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"
    nip = f"19{80 + i % 20}{(i % 12) + 1:02d}{(i % 28) + 1:02d}{1000 + i}"

    nk_values = [round(random.uniform(2.5, 4.0), 2) for _ in range(5)]

    pendidikan = random.choices(PENDIDIKAN_OPTS, weights=PENDIDIKAN_WEIGHTS)[0]
    sertifikasi = random.choices(SERTIFIKASI_OPTS, weights=SERTIFIKASI_WEIGHTS)[0]
    exposure = random.randint(35, 100)
    potensi = random.randint(35, 100)
    k3 = random.choices(K3_OPTS, weights=K3_WEIGHTS)[0]

    status = random.choices(
        ["Aktif", "Cuti Sakit", "Sanksi", "CLTB", "Pemberhentian Sementara"],
        weights=[0.90, 0.03, 0.03, 0.02, 0.02],
    )[0]
    ptb_s2 = random.random() < 0.05
    promotion_award = random.random() < 0.08
    satker_rec = random.choices(["Direkomendasikan", "Tidak Direkomendasikan"], weights=[0.85, 0.15])[0]
    remaining_service = round(random.uniform(0.2, 20), 1)
    readiness = random.choices(READINESS_OPTS, weights=[0.30, 0.40, 0.30])[0]

    return {
        "NIP": nip,
        "Nama": nama,
        "Unit": SATKER_UNIT[satker],
        "Satker": satker,
        "Pangkat": pangkat,
        "Sublevel": sublevel,
        "Tanggal_Grade": random_date_grade(pangkat),
        "NK_1": nk_values[0],
        "NK_2": nk_values[1],
        "NK_3": nk_values[2],
        "NK_4": nk_values[3],
        "NK_5": nk_values[4],
        "Pendidikan": pendidikan,
        "Sertifikasi": sertifikasi,
        "Exposure": exposure,
        "Potensi": potensi,
        "K3": k3,
        "Status": status,
        "PTB_S2": ptb_s2,
        "Promotion_Award": promotion_award,
        "Satker_Recommendation": satker_rec,
        "Remaining_Service": remaining_service,
        "Readiness": readiness,
    }


def build_workbook(path):
    employees = pd.DataFrame([gen_employee(i) for i in range(N_EMPLOYEES)])

    quant_weights = pd.DataFrame(
        {"Parameter": ["NK", "MDP", "Pendidikan", "Sertifikasi"], "Value": [0.35, 0.25, 0.30, 0.10]}
    )
    qual_weights = pd.DataFrame(
        {"Parameter": ["Exposure", "Potensi", "K3"], "Value": [0.25, 0.20, 0.55]}
    )
    rank_weights = pd.DataFrame(
        {
            "Pangkat": PANGKAT_ORDER,
            "Quantitative": [0.40, 0.45, 0.55, 0.60, 0.70],
            "Qualitative": [0.60, 0.55, 0.45, 0.40, 0.30],
        }
    )
    education_score = pd.DataFrame(
        {"Kategori": PENDIDIKAN_OPTS, "Nilai": [100, 85, 70, 55, 40]}
    )
    certification_score = pd.DataFrame(
        {"Kategori": SERTIFIKASI_OPTS, "Nilai": [40, 25, 15, 10, 10, 0]}
    )
    # Tabel tambahan: konversi kategori K3 ke skor numerik (tidak ada di file asli,
    # ditambahkan supaya K3 bisa masuk hitungan Qualitative Score secara transparan)
    k3_score = pd.DataFrame({"Kategori": K3_OPTS, "Nilai": [100, 75, 50, 25]})
    thresholds = pd.DataFrame(
        {
            "Parameter": [
                "MDG Threshold (tahun)",
                "Minimum NK",
                "Minimum Remaining Service (tahun)",
            ],
            "Value": [2, 3, 0.5],
        }
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        employees.to_excel(writer, sheet_name="Employees", index=False)
        quant_weights.to_excel(writer, sheet_name="Quant_Weights", index=False)
        qual_weights.to_excel(writer, sheet_name="Qual_Weights", index=False)
        rank_weights.to_excel(writer, sheet_name="Rank_Weights", index=False)
        education_score.to_excel(writer, sheet_name="Education_Score", index=False)
        certification_score.to_excel(writer, sheet_name="Certification_Score", index=False)
        k3_score.to_excel(writer, sheet_name="K3_Score", index=False)
        thresholds.to_excel(writer, sheet_name="Thresholds", index=False)

    print(f"OK: {path} dibuat dengan {len(employees)} pegawai.")


if __name__ == "__main__":
    build_workbook("data/sikabi_data.xlsx")
