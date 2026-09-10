# MCPS KPK - Formulir Pendataan Aset & Inventaris Pegawai

Aplikasi formulir web modern dan responsif untuk mengumpulkan data inventarisasi pegawai (Kendaraan Roda 2, Kendaraan Roda 4, Rumah Dinas, dan Peralatan Kantor) beserta unggah bukti foto fisik secara langsung ke Google Drive dan Google Sheets.

## Fitur Utama
1. **Frontend Modern**: Menggunakan Tailwind CSS, responsif untuk layar HP maupun Desktop.
2. **Kompresi Gambar Otomatis**: Foto dikompresi di sisi klien sebelum diunggah sehingga pengiriman data sangat cepat dan hemat kuota.
3. **Penyimpanan Terintegrasi Google Cloud**:
   - **Google Sheets**: Seluruh data teks, tanggal/waktu pencatatan, dan link foto tersimpan rapi di spreadsheet tujuan.
   - **Google Drive**: Foto fisik aset otomatis tersimpan di folder Uploads_MCPS_KPK dengan link pratinjau publik.
4. **Validasi & Integritas Data**: Dilengkapi checkbox pernyataan kebenaran data dan validasi NIP/identitas.

---

## Tautan Terkait
- **Spreadsheet Data**: [Google Sheets MCPS KPK](https://docs.google.com/spreadsheets/d/1kucE7kAQ3wJ_XAQ9JefRaDaRE4W_uBe2ICOXuEwt1kA/edit?usp=sharing)
- **Script Editor (Apps Script)**: [Google Apps Script Project](https://script.google.com/d/1m9IGdpAemdKmEzQD3avtuaftj1wtelklvGEi6rNIzimOHTySbsL-B1EA/edit)
- **Live Web Form (GitHub Pages)**: https://koesmamr.github.io/mcps-kpk/

---

## Struktur Proyek
\\\
├── index.html          # Halaman formulir web (Frontend)
├── gas/
│   ├── appsscript.json # Manifest Apps Script
│   └── Code.js         # Backend Google Apps Script (doPost & Google Drive/Sheet Handler)
├── .gitignore
└── README.md
\\\

---

## Petunjuk Konfigurasi & Otorisasi Google Apps Script

Karena script mengakses Google Drive dan Google Spreadsheet Anda untuk pertama kali, Google mewajibkan otorisasi sekali saja:
1. Buka [Editor Apps Script Proyek Ini](https://script.google.com/d/1m9IGdpAemdKmEzQD3avtuaftj1wtelklvGEi6rNIzimOHTySbsL-B1EA/edit).
2. Di toolbar atas, pilih fungsi doGet lalu klik tombol **Jalankan (Run)**.
3. Google akan menampilkan popup **"Review Permissions" / "Tinjau Izin"**.
4. Pilih akun Google Anda (koesmarapat@gmail.com), klik **Advanced / Lanjutan**, lalu klik **Go to Backend MCPS KPK (Buka)** dan pilih **Allow / Izinkan**.
5. Klik menu **Terapkan (Deploy)** -> **Kelola Penerapan (Manage Deployments)**:
   - Pastikan Versi Web App memiliki setelan:
     - **Jalankan sebagai (Execute as)**: *Saya (Me / koesmarapat@gmail.com)*
     - **Siapa yang memiliki akses (Who has access)**: *Siapa saja (Anyone)*
6. Formulir siap digunakan dan dapat diisi oleh siapa saja tanpa perlu login!
