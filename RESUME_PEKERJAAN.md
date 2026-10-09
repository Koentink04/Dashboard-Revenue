# RESUME PEKERJAAN & HANDOVER DOKUMENTASI
**Proyek:** Executive Dashboard Revenue Monitoring & Collection — PT PLN Icon Plus (Danantara Indonesia)  
**Repository GitHub:** [https://github.com/Koentink04/Dashboard-Revenue](https://github.com/Koentink04/Dashboard-Revenue)  
**Branch Utama:** `main`  
**Waktu Pembaruan Terakhir:** 09 Oktober 2026

---

## 1. Status Proyek Terkini
- Seluruh kode sumber, aset, dataset (15.013 faktur), dan visualisasi telah berhasil dikomit dan di-push ke repositori GitHub: `https://github.com/Koentink04/Dashboard-Revenue` pada branch `main`.
- Fitur terbaru yang selesai diimplementasikan dan telah melalui uji visual serta uji interaksi (*browser live verification*) adalah **Interactive Drag-and-Drop Reordering & Bidirectional Synchronization untuk Sub Tab & Sub Menu Modul Collection**.

---

## 2. Rincian Fitur Terakhir yang Diimplementasikan
### A. Drag-and-Drop Reordering Sub Menu Collection
1. **Dua Mode Interaksi (Horizontal & Vertikal):**
   - **Horizontal Toolbar (`#col-subtab-container`)**: 8 tombol sub tab (*Dashboard Collection, Invoice PLN v1, Invoice PLN v2, Invoice STI PLN v1, Invoice STI PLN v2, Utang Usaha, Cash In, COP*) dapat digeser secara horizontal dengan indikator drop guideline kiri/kanan (`.drag-over-left`, `.drag-over-right`).
   - **Vertical Sidebar (`#nav-sub-collection`)**: 8 item submenu vertikal dapat digeser naik/turun dengan indikator drop guideline atas/bawah (`.drag-over-top`, `.drag-over-bottom`).
2. **Sinkronisasi Dua Arah (*Bidirectional Sync*):**
   - Mengubah urutan pada tab toolbar atas otomatis mengubah urutan item submenu di sidebar.
   - Mengubah urutan pada sidebar otomatis memperbarui susunan tombol tab di toolbar atas.
3. **Penyimpanan Permanen (*Persistence*):**
   - Urutan baru tersimpan di browser via `localStorage` dengan key `pln_collection_subtab_order`.
   - Urutan kustom otomatis diterapkan kembali saat halaman dimuat ulang (*reload*).
   - Terdapat validasi proteksi jika di masa depan terdapat penambahan tab baru agar tidak hilang.
4. **Tombol "Reset Urutan":**
   - Tombol `#col-subtab-reset-btn` disematkan di ujung kanan toolbar horizontal.
   - Mengembalikan susunan tab toolbar dan sidebar ke urutan *default* awal, menghapus data kustom dari `localStorage`, serta menampilkan notifikasi *toast*.
   - Terisolasi dari aksi drag dan otomatis disembunyikan saat cetak dokumen (`pln-no-print`).

---

## 3. Modul & Komponen yang Tersedia dalam Proyek
1. **Dashboard Revenue:**
   - Monitoring capaian 1 tahun (2026), breakdown bulanan Kumulatif & Stand Alone, serta filter SBU (14 SBU).
2. **Prognosa Revenue:**
   - Analisis target RKAP, Carry Over (CO), New Revenue (NR), waterfall chart AS IS vs Most Likely, dan tabel *Top 10 Exception Data*.
3. **Monitoring PA Open:**
   - Tracking Project Authorization (PA) Open, metrik realisasi & estimasi aktivasi revenue.
4. **Collection & Monitoring Invoice (List Invoice 15.013 Data):**
   - Tren penagihan vs Cash In (2025–2026), distribusi umur piutang (*aging categories*), rekapitulasi SBU, dan watchlist Top 10 debitur.
   - Subview khusus: *Invoice PLN v1*, *Invoice PLN v2*, *Invoice STI PLN v1*, *Invoice STI PLN v2*, *Penyelesaian Utang Usaha*, *Cash In*, dan *Collection of Period (COP)*.
5. **Opportunity & PO:**
   - Rekap 36 kolom & detail angsuran 37 kolom.
6. **Authentication & Security:**
   - Enterprise authentication gate modal dengan fallback 3D logo (Three.js).
   - Akun default: `admin` / `admin!@#`.
   - Parameter URL untuk bypass sesi dev: `?auth=1`.

---

## 4. Panduan Menjalankan di Perangkat Baru

1. **Clone Repositori:**
   ```bash
   git clone https://github.com/Koentink04/Dashboard-Revenue.git
   cd Dashboard-Revenue
   ```

2. **Jalankan Web Server Lokal:**
   - **Opsi PowerShell (bawaan Windows):**
     ```powershell
     powershell -ExecutionPolicy Bypass -File server.ps1
     ```
   - **Atau Python:**
     ```bash
     python -m http.server 8080
     ```
   - **Atau VS Code / Cursor:**
     Gunakan ekstensi *Live Server* dan buka `index.html`.

3. **Buka di Browser:**
   - Masuk ke: `http://localhost:8080/?auth=1#collection` (langsung masuk ke modul Collection) atau `http://localhost:8080` (login: `admin` / `admin!@#`).

---

## 5. File Utama Proyek
- `index.html`: File inti aplikasi (tampilan UI, CSS tokens, tata letak antarmuka, dan seluruh logika interaksi JavaScript).
- `collection_data.js`: Data agregasi dan fungsi analitik transaksi faktur piutang.
- `server.ps1`: Script server HTTP lokal sederhana untuk menjalankan aplikasi pada port 8080.
- `.gitignore`: Konfigurasi file yang diabaikan (backup, temporary, dan log).
