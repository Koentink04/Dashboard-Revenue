# 📊 Revenue Assurance & Monitoring Dashboard - PT PLN Icon Plus

A comprehensive interactive web application for monitoring 1-year revenue performance, monthly breakdowns, and opportunity/contract pipeline management (CR/JR) for **PT PLN Icon Plus**.

## 🚀 Repository & Deployment
- **GitHub:** [https://github.com/bagussundaru/Revenue_Monitoring](https://github.com/bagussundaru/Revenue_Monitoring)
- **Deployment:** Hosted on Vercel

## ✨ Key Features
- **Dashboard Revenue Tahunan:** 
  - Real-time monthly revenue breakdown, target achievement, and revenue trend analysis.
  - Interactive 2.5D Isometric Monthly Bar Chart with glassmorphism inspection card and click-outside dismiss.
- **Opportunity & PO Pipeline Management:**
  - Complete tracking for Sheet REKAP CR/JR (36 Columns) and DETIL ANGSURAN (37 Columns).
  - 6 Key Performance Metric Cards:
    1. `TOTAL OPPORTUNITY` (Nilai Full)
    2. `TOTAL NILAI WBS` (Anggaran WBS yang disetujui)
    3. `PEKERJAAN TANPA WBS` (Proyek tanpa nomor WBS dengan click-to-filter)
    4. `TIDAK ADA BA UAT` (Proyek belum memiliki BAUT / UAT)
    5. `TIDAK ADA BA DEPLOY` (Proyek belum Go-Live / Deploy)
    6. `BELUM CAIR` (Termin angsuran belum lunas)
- **Interactive Multi-Filter & Search:**
  - Real-time search by description / STI letter number.
  - Granular filters by Layanan (AP2T, dll), Tipe (CR/JR), Skema Pembayaran (Bertahap/Sekaligus), Status WBS, Status UAT, dan Status Deploy.
- **Data Management:**
  - Form input modal lengkap untuk Sheet REKAP dan DETIL ANGSURAN.
  - Excel file import (.xlsx) support untuk `Sampling_data_rev3.xlsx`.
  - LocalStorage persistence.

## 🛠️ Tech Stack
- Pure HTML5 / CSS3 (Responsive Grid & Flexbox)
- Vanilla JavaScript (ES6+)
- SheetJS (xlsx parser)
- Vercel Static Hosting

---
© 2026 PT PLN Icon Plus - Core Process Suite
