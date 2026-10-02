# Jadwal Produksi Bengkel — versi final (frontend + backend + kode Dinda)

## Isi folder
```
backend/final_colab_scheduler.py   kode penjadwalan Dinda (isi asli, 1 baris diubah, lihat bawah)
backend/main.py                    server FastAPI: API + menyajikan halaman web
backend/storage.py                 penyimpanan riwayat (SQLite)
backend/requirements.txt           daftar library Python
frontend/index.html                halaman web (Gantt harian/mingguan/tahunan, rentang tanggal, export)
data/                              Excel contoh (full 694 operasi, dan 1 SO saja untuk uji cepat)
Dockerfile, railway.json, render.yaml   konfigurasi deploy
run_local.bat / run_local.sh       jalankan di laptop
```

## Cara kerjanya
Kode Dinda **bukan server**. Itu skrip Python biasa (fungsi + command line). `backend/main.py` yang membungkusnya menjadi server web.

    Browser (index.html) --upload Excel + pilih metode--> POST /api/schedule
                                                              |
                                          main.py memanggil final_colab_scheduler.py
                                                              |
    Browser <--JSON jadwal (tanggal absolut, pakai kalender Dinda)--+

Satu service = satu deploy. `index.html` dibuka dari alamat server yang sama, jadi tidak ada urusan CORS.

## 1. Coba di laptop dulu (wajib sebelum deploy)
Butuh Python 3.10+.

    cd backend
    pip install -r requirements.txt
    uvicorn main:app --port 8000

Buka http://localhost:8000, unggah `data/Input_Penjadwalan_Python_TA.xlsx`, pilih metode, jalankan.
(Windows: klik dua kali `run_local.bat`.) Jangan buka `index.html` dengan klik dua kali; harus lewat alamat server.

## 2. Deploy ke Railway (paling mudah)
1. Buat repo GitHub (private), unggah isi folder ini. Folder `data/` boleh tidak ikut (berisi data perusahaan); sudah di-ignore oleh `.dockerignore` untuk build.
2. railway.com → New Project → Deploy from GitHub repo → pilih repo. Railway otomatis memakai `Dockerfile`.
3. Tambah **Volume** supaya riwayat tidak hilang tiap deploy ulang: klik service → Settings → Volumes (atau tombol + New → Volume) → mount path `/data`. Lalu di Variables tambah `DATA_DIR` = `/data`.
4. Tab Variables, tambah juga:
   - `APP_USER` = nama login yang kamu mau
   - `APP_PASSWORD` = password panjang
   (ini mengunci seluruh situs dengan login; sangat disarankan karena data produksi bersifat rahasia)
5. Tab Settings → Networking → Generate Domain. Buka alamatnya, login, pakai.
5. Pastikan memori service minimal 1 GB (numba + pandas + GA). Paket gratis/trial sering terlalu kecil.

Alternatif: Render.com → New → Blueprint → pilih repo (membaca `render.yaml`). Isi `APP_USER`/`APP_PASSWORD` saat diminta. Paket Starter ke atas, jangan Free (tidur setelah 15 menit dan memori 512 MB).

## 2b. Mengatur metode penjadwalan
Dinda minta di app **tidak ada pilihan metode** (metode final masih diputuskan). Metode yang dipakai diatur di server:

    SCHEDULE_METHOD = EDD     (pilihan: EDD, FCFS, SPT, CR, GA; default EDD)

Setelah Dinda memutuskan, ubah nilainya di Variables Railway (deploy ulang otomatis). Untuk menguji semua metode, buka alamat dengan `?uji=1` (misal `https://.../?uji=1`) atau set `ALLOW_METHOD_CHOICE=1`.

## 3. Memakai aplikasinya
1. Unggah Excel 7 sheet (SO, Operations, Eligible_Resources, Resources, Human_Capacity, Calendar, Lead_Time).
2. Pilih EDD / FCFS / SPT / CR / Genetic Algorithm, klik Jalankan. Heuristik ±1–2 detik, GA ±10–60 detik.
3. Gantt: tombol Harian / Mingguan / Tahunan, rentang "Dari – Sampai tanggal" lalu Terapkan. Klik SO di legenda untuk menyorot.
4. Lembar jadwal (tabel per operasi) mengikuti rentang tanggal yang sama.
5. Setiap hasil otomatis tersimpan di bagian **Riwayat penjadwalan** (tanggal, nama/catatan, metode). Klik Buka untuk melihat lagi, atau unduh Excel input yang dulu dipakai. Tersimpan di SQLite di `DATA_DIR` (tanpa Volume, riwayat hilang saat deploy ulang). Isi nama/catatan sebelum menjalankan supaya mudah dicari.
6. Info pekerjaan tambahan: kalau sheet `SO` punya kolom selain SO_ID/Release_Date/Due_Date (misal Jenis_Unit, Diameter_mm, Panjang_mm, Lingkup_Pekerjaan, Tgl_Masuk_Workshop), kolom itu otomatis tampil di tabel keterlambatan (maks 8 kolom). Kode Dinda tidak memakainya untuk menjadwalkan.
7. Export Excel (3 sheet: Jadwal, Keterlambatan, Utilisasi) dan Export PDF (dialog cetak browser, landscape). Untuk PDF yang enak dibaca pilih skala Mingguan atau batasi rentang tanggal dulu.

## Hal yang perlu kamu tahu
- **Perubahan di kode Dinda**: hanya 1 baris di `load_problem`, kunci `EXPECTED_OPERATION_COUNT = 694` dibuat opsional (`is not None`). Nilai 694 tetap ada, jadi menjalankan skrip lewat command line untuk skripsi berperilaku sama persis. `main.py` mematikan kunci itu supaya data lain (misal 1 SO) bisa diproses. Sudah dicoba dengan `data/contoh_1_SO_1074-01.xlsx` (86 operasi, semua metode jalan).
- **Kunci lain yang masih ada di kode Dinda**: kalender harus tepat 400 menit produktif per hari (`productive_min_day != 400` akan error). Kalau shift/jam istirahat diubah sehingga totalnya bukan 400, kode Dinda perlu disesuaikan, bukan hanya Excel-nya.
- **Perubahan shift/pekerja/mesin**: dilakukan lewat Excel (sheet Calendar, Human_Capacity, Resources), bukan lewat form di web seperti versi lama.
- **GA di web = 1 seed** (seed 101, 70 generasi, parameter dasar). Hasil skripsi (20 seed final + tuning) tetap dijalankan terpisah:
  `python final_colab_scheduler.py --input "Input.xlsx" --workdir hasil --stage all`
- **Utilisasi mesin** dihitung `main.py` (menit proses ÷ unit × menit efektif hari kerja selama jadwal berjalan). Itu angka tampilan aplikasi, bukan bagian dari kode Dinda.
- Penghapusan riwayat tidak bisa dibatalkan. Siapa pun yang punya login bisa melihat dan menghapus semua riwayat (belum ada akun per pengguna).
- Satu penjadwalan diproses satu per satu (antrean) supaya server kecil tidak kehabisan memori.
