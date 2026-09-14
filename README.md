# SIKABI — Sistem Intelijen Karier Bank Indonesia

Prototipe aplikasi internal untuk staf DSDM mengolah data penilaian kelayakan
promosi pegawai (Quantitative/Qualitative Score → Gate Administrasi → Gate KPP
→ Grade Senior/MDG → Kuadran Prioritas Promosi).

## Struktur proyek

```
sikabi/
├── app.py              # Aplikasi Streamlit (entry point)
├── scoring.py          # Mesin perhitungan skor & gate keputusan (murni Python, testable)
├── data_gen.py         # Generator data dummy (500 pegawai + tabel referensi)
├── data/
│   └── sikabi_data.xlsx  # Sumber data (Employees + 7 tabel master/referensi)
├── requirements.txt
└── README.md
```

## Menjalankan secara lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

Buka `http://localhost:8501` di browser.

## Membuat ulang data dummy

Kalau ingin regenerasi 500 data pegawai dari awal (misalnya untuk reset demo):

```bash
python data_gen.py
```

Ini akan menimpa `data/sikabi_data.xlsx`.

## Struktur data (`data/sikabi_data.xlsx`)

| Sheet | Isi |
|---|---|
| `Employees` | Data mentah pegawai: NIP, Nama, Unit, Satker, Pangkat, Sublevel, Tanggal_Grade, NK_1..NK_5, Pendidikan, Sertifikasi, Exposure, Potensi, K3, Status, PTB_S2, Promotion_Award, Satker_Recommendation, Remaining_Service, Readiness |
| `Quant_Weights` | Bobot komponen Quantitative Score (NK, MDP, Pendidikan, Sertifikasi) |
| `Qual_Weights` | Bobot komponen Qualitative Score (Exposure, Potensi, K3) |
| `Rank_Weights` | Bobot Quantitative vs Qualitative per pangkat (DD, AD, M, AM, S-A) |
| `Education_Score` | Konversi kategori pendidikan → skor numerik |
| `Certification_Score` | Konversi kategori sertifikasi → skor numerik |
| `K3_Score` | Konversi kategori K3 (SB/B/CB/KB) → skor numerik — **sheet tambahan**, tidak ada di file asli, dibuat supaya K3 bisa dihitung transparan sebagai bagian Qualitative Score |
| `Thresholds` | Ambang batas syarat administrasi (minimum NK, minimum sisa dinas, MDG threshold) |

Semua tabel di atas bisa diedit langsung (tambah/ubah/hapus baris) lewat halaman
**Master Data** di aplikasi — perubahan langsung memengaruhi seluruh hasil
perhitungan skor dan gate keputusan.

## Catatan penting soal penyimpanan data

Tombol **Simpan Perubahan** di halaman Master Data menulis ulang file
`data/sikabi_data.xlsx` di server tempat aplikasi berjalan. Ini bekerja normal
saat dijalankan **lokal**. Kalau di-deploy ke **Streamlit Community Cloud**,
penyimpanan filesystem bersifat sementara (ephemeral) — perubahan bisa hilang
saat aplikasi restart/redeploy. Untuk pemakaian produksi jangka panjang,
sebaiknya ganti sumber data dari file Excel lokal ke database (mis. Google
Sheets API, PostgreSQL, atau Supabase) — struktur `scoring.py` sudah dipisah
dari logika baca/tulis data sehingga penggantian ini tidak perlu mengubah
rumus perhitungan.

Tombol **Unduh data (.xlsx)** di sidebar selalu tersedia sebagai cara aman
untuk mengambil salinan data terbaru kapan saja.

## Deploy ke Streamlit Community Cloud

1. Push folder ini ke repository GitHub.
2. Buka [share.streamlit.io](https://share.streamlit.io), sambungkan ke repo tersebut.
3. Set **Main file path** ke `app.py`.
4. Deploy.

## Asumsi & catatan implementasi

- **MDP (Masa Dinas Pangkat)** dan **MDG (Masa Dinas Grade)** dihitung dari
  kolom `Tanggal_Grade` yang sama, karena hanya satu tanggal acuan yang
  tersedia di data. Jika BI memiliki tanggal terpisah untuk mulai pangkat vs
  mulai grade, tambahkan kolom terpisah dan sesuaikan `scoring.py`.
- **Mean & Standar Deviasi** untuk skor NK dan MDP dihitung dari seluruh
  populasi pegawai (BI Wide), bukan per pangkat — sesuaikan fungsi
  `compute_all()` di `scoring.py` jika perhitungan per pangkat diperlukan.
- **Passing Grade** Gate KPP dihitung sebagai rata-rata QScore/Quantitative/
  Qualitative dari pegawai yang sudah lolos Gate Administrasi, dikelompokkan
  per pangkat.
- **Readiness KPP** saat ini diambil langsung dari kolom `Readiness` di data
  mentah (bukan dihitung otomatis), karena formulanya belum didefinisikan
  secara pasti di system requirements.
