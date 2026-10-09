
  // =========================================================================
  // 1 MONITORED YEAR DATA WITH 12 MONTHS BREAKDOWN (PLN ICON PLUS)
  // =========================================================================
  const STORAGE_KEY = 'pln_icon_plus_revenue_annual_v4';

  // Baseline 12 months for 1 Year (in Juta Rupiah)
  const DEFAULT_MONTHS = [
    { short: 'Jan', long: 'Januari',   target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Feb', long: 'Februari',  target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Mar', long: 'Maret',     target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Apr', long: 'April',     target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Mei', long: 'Mei',       target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Jun', long: 'Juni',      target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Jul', long: 'Juli',      target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Agu', long: 'Agustus',   target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Sep', long: 'September', target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Okt', long: 'Oktober',   target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Nov', long: 'November',  target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 },
    { short: 'Des', long: 'Desember',  target: 0, opp: 0, cash: 0, wait: 0, risk: 0, proc: 0 }
  ];

  const DEFAULT_PROGNOSA = {
    // 7 Kartu KPI Utama Prognosa (Satuan Triliun Rupiah / T)
    targetKumulatif: 2.74,     // Card 1: Target Kumulatif Sep 2026 (2,74 T)
    targetStandAlone: 0.25,    // Card 2: Target Stand Alone Sep 2026 (0,25 T)
    realisasiLalu: 2.23,       // Card 3: Realisasi s.d Agustus 2026 (2,23 T)
    prognosaAsIs: 2.60,        // Card 4: Prognosa AS IS Sep 2026 (2,60 T)
    prognosaMostLikely: 3.42,  // Card 5: Prognosa Most Likely Sep 2026 (3,42 T)

    // Parameter Korporat & Basis Waterfall RKAP (Satuan Milyar Rupiah / M)
    targetRkap: 4156.00,
    targetBulanan: 415.59,
    co: 2618.09,
    nr: 1537.91,
    nrBulanBerjalan: 8.47,
    estRevenue: 4034.02,
    sourceFile: 'Penajaman Strategi Penjualan PLN Tahun 2026 versi updated.xlsx',
    services: [
      { name: 'PLN 1 (Holding / Korporasi / DIVSTI)', target: 1183.02, co: 797.52, nr: 385.50, nrMonth: 3.00, est: 1236.57 },
      { name: 'PLN 2 (Subholding & Anak Perusahaan)', target: 2307.29, co: 1683.19, nr: 624.10, nrMonth: 0.00, est: 2420.32 },
      { name: 'SBU JBB (Jakarta Raya & Banten)', target: 89.00, co: 8.40, nr: 80.60, nrMonth: 0.00, est: 102.00 },
      { name: 'SBU JTG (Jawa Tengah & DIY)', target: 83.31, co: 17.28, nr: 66.03, nrMonth: 0.00, est: 66.26 },
      { name: 'SBU JBT (Jawa Barat)', target: 75.00, co: 13.65, nr: 61.35, nrMonth: 0.00, est: 34.78 },
      { name: 'SBU BNR (Bali & Nusa Tenggara)', target: 45.00, co: 11.05, nr: 33.95, nrMonth: 1.00, est: 42.38 },
      { name: 'SBU SBU (Sumatera Bagian Utara)', target: 90.00, co: 16.83, nr: 73.17, nrMonth: 0.47, est: 27.68 },
      { name: 'SBU SBT (Sumatera Bagian Tengah)', target: 73.06, co: 31.93, nr: 41.13, nrMonth: 0.00, est: 34.28 },
      { name: 'SBU SBS (Sumatera Bagian Selatan)', target: 64.32, co: 16.38, nr: 47.94, nrMonth: 0.00, est: 22.99 },
      { name: 'SBU SLW (Sulawesi)', target: 73.00, co: 16.38, nr: 56.62, nrMonth: 4.00, est: 27.79 },
      { name: 'SBU KLM (Kalimantan)', target: 73.00, co: 5.48, nr: 67.52, nrMonth: 0.00, est: 18.98 }
    ]
  };

  const DEFAULT_PROGNOSA_LOGS = [
    {
      id: 'LOG-20260928-091522',
      timestamp: '28 Sep 2026 09:15:22 WIB',
      user: 'dodik.gunawan',
      userRole: 'Revenue Assurance PIC',
      action: 'Upload File Excel',
      source: 'Penajaman Strategi Penjualan PLN Tahun 2026 versi updated.xlsx',
      diffs: [
        { param: 'Card 1 (Target Kumulatif)', before: '-', after: '2,74 T', change: 'Baru' },
        { param: 'Card 2 (Target Stand Alone)', before: '-', after: '0,25 T', change: 'Baru' },
        { param: 'Card 3 (Realisasi s.d Ags)', before: '-', after: '2,23 T', change: 'Baru' },
        { param: 'Card 4 (Prognosa AS IS)', before: '-', after: '2,60 T', change: 'Baru' },
        { param: 'Card 5 (Prognosa Most Likely)', before: '-', after: '3,42 T', change: 'Baru' },
        { param: 'Target RKAP Tahunan', before: '-', after: '4.156,00 M', change: 'Baru' }
      ],
      note: 'Inisialisasi baseline parameter prognosa dari file Excel penajaman 2026',
      status: 'Tersimpan'
    },
    {
      id: 'LOG-20260929-143010',
      timestamp: '29 Sep 2026 14:30:10 WIB',
      user: 'dodik.gunawan',
      userRole: 'Revenue Assurance PIC',
      action: 'Selaraskan dgn RKAP',
      source: 'Input Form Master Prognosa',
      diffs: [
        { param: 'Card 1 (Target Kumulatif)', before: '2,74 T', after: '2,74 T', change: 'Tetap' },
        { param: 'Card 2 (Target Stand Alone)', before: '0,25 T', after: '0,25 T', change: 'Tetap' }
      ],
      note: 'Sinkronisasi otomatis rasio penajaman kumulatif s.d Sep (~66% dari 4.156 M)',
      status: 'Tersimpan'
    },
    {
      id: 'LOG-20260930-101240',
      timestamp: '30 Sep 2026 10:12:40 WIB',
      user: 'dodik.gunawan',
      userRole: 'Revenue Assurance PIC',
      action: 'Manual Edit Form',
      source: 'Form Master Prognosa',
      diffs: [
        { param: 'Card 3 (Realisasi s.d Ags)', before: '2,21 T', after: '2,23 T', change: '+0,02 T' },
        { param: 'Card 4 (Prognosa AS IS)', before: '2,58 T', after: '2,60 T', change: '+0,02 T' }
      ],
      note: 'Penyesuaian realisasi revenue s.d Agustus 2026 dan rekonsiliasi run-rate',
      status: 'Tersimpan'
    }
  ];

  const DEFAULT_APP_STATE = {
    year: '2026',
    oppCount: 6,
    cutoffText: '30 September 2026',
    cutoffDesc: 'Cutoff 30 September · Sisa 4 hari kerja',
    annual: {
      useDirect: true,
      target: 186.00,
      opp: 1.80,
      real: 0.30,
      mostLikely: 1.60,
      cash: 0.10,
      wait: 0.20,
      risk: 0.00
    },
    prognosa: DEFAULT_PROGNOSA,
    prognosaLogs: DEFAULT_PROGNOSA_LOGS,
    months: DEFAULT_MONTHS
  };

  function loadAppData() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && parsed.months && parsed.months.length === 12) {
          // Clean out any leftover dummy monthly opp/cash numbers
          if (parsed.months[0] && (parsed.months[0].opp === 22000 || parsed.months[8].opp === 28500)) {
            parsed.months.forEach(m => {
              m.opp = 0;
              m.cash = 0;
              m.wait = 0;
              m.risk = 0;
              m.proc = 0;
            });
          }
          if (!parsed.annual || parsed.annual.real === 159.33 || parsed.annual.opp === 272.80) {
            parsed.annual = {
              useDirect: true,
              target: parsed.annual && parsed.annual.target ? parsed.annual.target : 186.00,
              opp: 1.80,
              real: 0.40,
              mostLikely: 1.60,
              cash: 0.10,
              wait: 0.30,
              risk: 0.00
            };
          }
          if (!parsed.prognosa || parsed.prognosa.targetRkap === 186.00 || parsed.prognosa.targetKumulatif === undefined) {
            parsed.prognosa = Object.assign(JSON.parse(JSON.stringify(DEFAULT_PROGNOSA)), parsed.prognosa || {});
          }
          if (!parsed.prognosaLogs || !Array.isArray(parsed.prognosaLogs) || parsed.prognosaLogs.length === 0) {
            parsed.prognosaLogs = JSON.parse(JSON.stringify(DEFAULT_PROGNOSA_LOGS));
          }
          return parsed;
        }
      }
    } catch (e) {
      console.warn('Failed to parse saved data, fallback to defaults', e);
    }
    return JSON.parse(JSON.stringify(DEFAULT_APP_STATE));
  }

  function saveAppData(data) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  }

  let appData = loadAppData();
  let selectedMonthIndex = 8; // Default September (0-indexed: 8)
  let editingMonthIndex = -1;

  // Format Helper: Juta to Milyar (e.g. 16500 Jt -> 16,50 M)
  function jtToM(jt) {
    if (jt === undefined || jt === null || isNaN(jt)) return '0,00';
    return (Number(jt) / 1000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function fmtJt(jt) {
    if (jt === undefined || jt === null || isNaN(jt)) return '0';
    return Number(jt).toLocaleString('id-ID');
  }
  function pctStr(a, b) {
    if (!b || b === 0) return '0%';
    return Math.round((a / b) * 100) + '%';
  }

  // Toast Notification
  function showToast(msg) {
    const t = document.getElementById('toast');
    document.getElementById('toast-msg').textContent = msg;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 3200);
  }

  // View Navigation
  function showDashboardView() {
    document.getElementById('view-dashboard').style.display = 'block';
    document.getElementById('view-master').style.display = 'none';
    const vOpp = document.getElementById('view-opportunity'); if (vOpp) vOpp.style.display = 'none';
    const vProg = document.getElementById('view-prognosa'); if (vProg) vProg.style.display = 'none';
    const vMProg = document.getElementById('view-master-prognosa'); if (vMProg) vMProg.style.display = 'none';
    const vPa = document.getElementById('view-monitoring-pa'); if (vPa) vPa.style.display = 'none';

    document.getElementById('nav-revenue').classList.add('active');
    const nProg = document.getElementById('nav-prognosa'); if (nProg) nProg.classList.remove('active');
    const nMPa = document.getElementById('nav-monitoring-pa'); if (nMPa) nMPa.classList.remove('active');
    const nMProg = document.getElementById('nav-master-prognosa'); if (nMProg) nMProg.classList.remove('active');
    document.getElementById('nav-master').classList.remove('active');
    const nOpp = document.getElementById('nav-opportunity'); if (nOpp) nOpp.classList.remove('active');

    document.getElementById('top-title').textContent = 'Dashboard Revenue';
    document.getElementById('top-sub').textContent = 'MONITORING REVENUE 1 TAHUN (' + appData.year + ') · BREAKDOWN BULANAN & OPPORTUNITY';
    renderDashboard();
  }

  function showPrognosaView() {
    document.getElementById('view-dashboard').style.display = 'none';
    document.getElementById('view-master').style.display = 'none';
    const vOpp = document.getElementById('view-opportunity'); if (vOpp) vOpp.style.display = 'none';
    const vProg = document.getElementById('view-prognosa'); if (vProg) vProg.style.display = 'block';
    const vMProg = document.getElementById('view-master-prognosa'); if (vMProg) vMProg.style.display = 'none';
    const vPa = document.getElementById('view-monitoring-pa'); if (vPa) vPa.style.display = 'none';

    document.getElementById('nav-revenue').classList.remove('active');
    const nProg = document.getElementById('nav-prognosa'); if (nProg) nProg.classList.add('active');
    const nMPa = document.getElementById('nav-monitoring-pa'); if (nMPa) nMPa.classList.remove('active');
    const nMProg = document.getElementById('nav-master-prognosa'); if (nMProg) nMProg.classList.remove('active');
    document.getElementById('nav-master').classList.remove('active');
    const nOpp = document.getElementById('nav-opportunity'); if (nOpp) nOpp.classList.remove('active');

    document.getElementById('top-title').textContent = 'Prognosa Revenue';
    document.getElementById('top-sub').textContent = 'PROGNOSA PENDAPATAN (' + (appData.year || '2026') + ') · TARGET, CARRY OVER & NEW REVENUE';
    renderPrognosaView();
  }

  function showMonitoringPaView() {
    document.getElementById('view-dashboard').style.display = 'none';
    document.getElementById('view-master').style.display = 'none';
    const vOpp = document.getElementById('view-opportunity'); if (vOpp) vOpp.style.display = 'none';
    const vProg = document.getElementById('view-prognosa'); if (vProg) vProg.style.display = 'none';
    const vMProg = document.getElementById('view-master-prognosa'); if (vMProg) vMProg.style.display = 'none';
    const vPa = document.getElementById('view-monitoring-pa'); if (vPa) vPa.style.display = 'block';

    document.getElementById('nav-revenue').classList.remove('active');
    const nProg = document.getElementById('nav-prognosa'); if (nProg) nProg.classList.remove('active');
    const nMPa = document.getElementById('nav-monitoring-pa'); if (nMPa) nMPa.classList.add('active');
    const nMProg = document.getElementById('nav-master-prognosa'); if (nMProg) nMProg.classList.remove('active');
    document.getElementById('nav-master').classList.remove('active');
    const nOpp = document.getElementById('nav-opportunity'); if (nOpp) nOpp.classList.remove('active');

    document.getElementById('top-title').textContent = 'Monitoring PA Open';
    document.getElementById('top-sub').textContent = 'TRACKING PROJECT AUTHORIZATION & AKTIVASI REVENUE (' + (appData.year || '2026') + ')';
    renderMonitoringPaView();
  }

  function showMasterPrognosaView() {
    document.getElementById('view-dashboard').style.display = 'none';
    document.getElementById('view-master').style.display = 'none';
    const vOpp = document.getElementById('view-opportunity'); if (vOpp) vOpp.style.display = 'none';
    const vProg = document.getElementById('view-prognosa'); if (vProg) vProg.style.display = 'none';
    const vMProg = document.getElementById('view-master-prognosa'); if (vMProg) vMProg.style.display = 'block';
    const vPa = document.getElementById('view-monitoring-pa'); if (vPa) vPa.style.display = 'none';

    document.getElementById('nav-revenue').classList.remove('active');
    const nProg = document.getElementById('nav-prognosa'); if (nProg) nProg.classList.remove('active');
    const nMPa = document.getElementById('nav-monitoring-pa'); if (nMPa) nMPa.classList.remove('active');
    const nMProg = document.getElementById('nav-master-prognosa'); if (nMProg) nMProg.classList.add('active');
    document.getElementById('nav-master').classList.remove('active');
    const nOpp = document.getElementById('nav-opportunity'); if (nOpp) nOpp.classList.remove('active');

    document.getElementById('top-title').textContent = 'Master Prognosa';
    document.getElementById('top-sub').textContent = 'PENGATURAN PARAMETER & UNGGAH FILE EXCEL PENAJAMAN 2026';
    renderMasterPrognosaView();
  }

  function showMasterView() {
    document.getElementById('view-dashboard').style.display = 'none';
    document.getElementById('view-master').style.display = 'block';
    const vOpp = document.getElementById('view-opportunity'); if (vOpp) vOpp.style.display = 'none';
    const vProg = document.getElementById('view-prognosa'); if (vProg) vProg.style.display = 'none';
    const vMProg = document.getElementById('view-master-prognosa'); if (vMProg) vMProg.style.display = 'none';
    const vPa = document.getElementById('view-monitoring-pa'); if (vPa) vPa.style.display = 'none';

    document.getElementById('nav-revenue').classList.remove('active');
    const nProg = document.getElementById('nav-prognosa'); if (nProg) nProg.classList.remove('active');
    const nMPa = document.getElementById('nav-monitoring-pa'); if (nMPa) nMPa.classList.remove('active');
    const nMProg = document.getElementById('nav-master-prognosa'); if (nMProg) nMProg.classList.remove('active');
    document.getElementById('nav-master').classList.add('active');
    const nOpp = document.getElementById('nav-opportunity'); if (nOpp) nOpp.classList.remove('active');

    document.getElementById('top-title').textContent = 'Master Data Revenue';
    document.getElementById('top-sub').textContent = 'MENU MASTER · INPUT & KELOLA DATA MANUAL 1 TAHUN & 12 BULAN';
    populateAnnualInputs();
    renderMasterTable();
  }

  // ==========================================
  // DATA MASTER TABEL PROGNOSA (ACUAN KORPORAT)
  // ==========================================
  const EXCEPTION_DATA = [
    { id: 'CUST-PLN-0101', name: 'PT PLN (Persero) Kantor Pusat', service: 'Metro Ethernet & IP VPN Korporat', value: 102.30, bidang: 'PLN Holding', thbl: '2026-09', note: 'Rekonsiliasi SLA & penyesuaian adendum' },
    { id: 'CUST-PLN-0102', name: 'PLN UID Jawa Timur', service: 'Dark Fiber & Jaringan SCADA Terintegrasi', value: 45.80, bidang: 'PLN UID/UIW', thbl: '2026-09', note: 'Menunggu Berita Acara Rekonsiliasi Triwulan 3' },
    { id: 'CUST-PLN-0103', name: 'PLN UID Jawa Barat', service: 'IP VPN Enterprise 620 Kantor Layanan', value: 38.20, bidang: 'PLN UID/UIW', thbl: '2026-08', note: 'Proses amandemen volume port & bandwidth' },
    { id: 'CUST-PLN-0104', name: 'PLN UID Jakarta Raya', service: 'Managed Service Data Center & DRC', value: 28.40, bidang: 'PLN UID/UIW', thbl: '2026-09', note: 'Verifikasi faktur pajak & kelengkapan lampiran' },
    { id: 'CUST-PLN-0105', name: 'PLN UIT Jawa Bagian Timur & Bali', service: 'Transmisi OPGW Backbone 500kV', value: 24.10, bidang: 'PLN UIT/UIP', thbl: '2026-08', note: 'Proses clearance cross-cut jalur transmisi' },
    { id: 'CUST-PLN-0106', name: 'PLN UIP Jawa Bagian Barat', service: 'Sistem SCADA Proteksi Gardu Induk', value: 19.50, bidang: 'PLN UIT/UIP', thbl: '2026-09', note: 'Tinjauan teknis perubahan konfigurasi IP' },
    { id: 'CUST-PLN-0107', name: 'PLN UID Sumatera Barat', service: 'Dedicated Internet & Call Center 123', value: 16.80, bidang: 'PLN UID/UIW', thbl: '2026-07', note: 'Rekonsiliasi denda ketersediaan jaringan' },
    { id: 'CUST-PLN-0108', name: 'PLN UID Bali', service: 'Smart Metering AMI Gateway Bali', value: 14.20, bidang: 'PLN UID/UIW', thbl: '2026-09', note: 'Finalisasi uji petik akurasi data konsumsi' },
    { id: 'CUST-PLN-0109', name: 'PLN UID Jawa Tengah & DIY', service: 'Wide Area Network & Hotspot Kantor', value: 12.00, bidang: 'PLN UID/UIW', thbl: '2026-08', note: 'Adendum perpanjangan masa kontrak' },
    { id: 'CUST-PLN-0110', name: 'PLN Nusantara Power', service: 'Sistem Monitoring Pembangkit Terpusat', value: 10.50, bidang: 'PLN Anak Perusahaan', thbl: '2026-09', note: 'Penyelarasan SLA integrasi dashboard' },
    { id: 'CUST-PLN-0111', name: 'PLN Indonesia Power', service: 'Jaringan Private APN Mobile Inspeksi', value: 8.90, bidang: 'PLN Anak Perusahaan', thbl: '2026-08', note: 'Verifikasi pemakaian kuota APN pool' },
    { id: 'CUST-PLN-0112', name: 'PLN UIW Kalimantan Timur & Utara', service: 'VSAT & Radio Link Wilayah 3T', value: 7.40, bidang: 'PLN UID/UIW', thbl: '2026-09', note: 'Verifikasi relokasi titik site remote' },
    { id: 'CUST-PLN-0113', name: 'PLN UIT Jawa Bagian Tengah', service: 'Fiber Optic Gardu Induk Pedurungan', value: 6.20, bidang: 'PLN UIT/UIP', thbl: '2026-08', note: 'Uji sertifikasi OTDR kabel tanah' },
    { id: 'CUST-PLN-0114', name: 'PLN UID Aceh', service: 'Metro-E Kantor Cabang & ULP', value: 5.80, bidang: 'PLN UID/UIW', thbl: '2026-07', note: 'Menunggu penetapan pejabat pengadaan baru' },
    { id: 'CUST-PLN-0115', name: 'PLN Icon Plus Div TI Kantor Pusat', service: 'Lisensi & Cloud Infrastructure Host', value: 5.10, bidang: 'PLN Holding', thbl: '2026-09', note: 'Penyelesaian administrasi intercompany transfer' },
    { id: 'CUST-PLN-0116', name: 'PLN UID Riau & Kepri', service: 'Submarine Cable Monitoring System', value: 4.60, bidang: 'PLN UID/UIW', thbl: '2026-09', note: 'Inspeksi sensor buoy lepas pantai Batam' },
    { id: 'CUST-PLN-0117', name: 'PLN UIW Sulselrabar', service: 'SD-WAN Enterprise 180 Lokasi', value: 4.20, bidang: 'PLN UID/UIW', thbl: '2026-08', note: 'Migrasi perangkat edge router cisco' },
    { id: 'CUST-PLN-0118', name: 'PLN Puslitbang', service: 'High Performance Computing Cluster', value: 3.80, bidang: 'PLN Holding', thbl: '2026-09', note: 'Uji performa riset AI kelistrikan' },
    { id: 'CUST-PLN-0119', name: 'PLN Pusat Sertifikasi', service: 'Aplikasi Sertifikasi Peralatan Listrik', value: 3.40, bidang: 'PLN Holding', thbl: '2026-08', note: 'Audit keamanan siber BSSN' },
    { id: 'CUST-PLN-0120', name: 'PLN UIP Sulawesi', service: 'CCTV Pengawasan PLTU Barru', value: 2.90, bidang: 'PLN UIT/UIP', thbl: '2026-07', note: 'Penggantian modul decoder kamera' },
    { id: 'CUST-PLN-0121', name: 'PLN UIW Maluku & Malut', service: 'Satelit LEO Komunikasi Darurat', value: 2.50, bidang: 'PLN UID/UIW', thbl: '2026-09', note: 'Pemberkasan izin landing rights satelit' },
    { id: 'CUST-PLN-0122', name: 'PLN UID Lampung', service: 'Backup Gateway Kantor Wilayah', value: 2.10, bidang: 'PLN UID/UIW', thbl: '2026-08', note: 'Adendum penambahan port 1 Gbps' },
    { id: 'CUST-PLN-0123', name: 'PLN Enjiniring', service: 'SaaS Engineering CAD & BIM', value: 1.80, bidang: 'PLN Anak Perusahaan', thbl: '2026-09', note: 'Penyesuaian user license bulanan' },
    { id: 'CUST-PLN-0124', name: 'PLN Haleyora Power', service: 'Aplikasi Field Force Management', value: 1.50, bidang: 'PLN Anak Perusahaan', thbl: '2026-09', note: 'Integrasi billing tim Yantek' },
    { id: 'CUST-PLN-0125', name: 'PLN Pelayaran Bahtera Adhiguna', service: 'Fleet Tracking & IoT Tongkang Batubara', value: 1.20, bidang: 'PLN Anak Perusahaan', thbl: '2026-08', note: 'Pengujian sensor telemetri kapal' }
  ];

  const BAA_DATA = [
    { id: 'BAA-2026-0901', name: 'PLN UID Jakarta Raya', service: 'Backbone Fiber Optic 100G Ring DKI', value: 32.50, bidang: 'PLN UID/UIW', sbu: 'SBU Jakarta & Banten', pic: 'Budi Santoso', sales: 'Dimas Aditya', targetTtd: '2026-08-31', aging: 42 },
    { id: 'BAA-2026-0902', name: 'PLN UID Jawa Timur', service: 'SD-WAN Enterprise 450 Titik Layanan', value: 28.40, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Timur', pic: 'Rian Pratama', sales: 'Siti Rahma', targetTtd: '2026-09-05', aging: 35 },
    { id: 'BAA-2026-0903', name: 'PT PLN (Persero) Kantor Pusat', service: 'Managed Security Services SOC & EDR', value: 22.10, bidang: 'PLN Holding', sbu: 'SBU Kantor Pusat', pic: 'Hendra Gunawan', sales: 'Andi Wijaya', targetTtd: '2026-09-10', aging: 28 },
    { id: 'BAA-2026-0904', name: 'PLN UIT Jawa Bagian Barat', service: 'Jaringan Proteksi Digital Substation', value: 18.70, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Barat', pic: 'Dewi Lestari', sales: 'Kevin Sanjaya', targetTtd: '2026-09-12', aging: 25 },
    { id: 'BAA-2026-0905', name: 'PLN UID Sumatera Utara', service: 'Upgrade Bandwidth IP Transit 40G', value: 15.20, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Utara', pic: 'Ahmad Fauzi', sales: 'Rina Marlina', targetTtd: '2026-09-14', aging: 22 },
    { id: 'BAA-2026-0906', name: 'PLN Nusantara Power UP Muara Karang', service: 'Private Optical Network GI Muara Karang', value: 12.80, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jakarta & Banten', pic: 'Bambang Irawan', sales: 'Dedi Kurniawan', targetTtd: '2026-09-16', aging: 18 },
    { id: 'BAA-2026-0907', name: 'PLN UIP Nusa Tenggara', service: 'Komunikasi Radio Microwave GI Mataram', value: 11.40, bidang: 'PLN UIT/UIP', sbu: 'SBU Bali & Nusa Tenggara', pic: 'Putu Arya', sales: 'Made Sukarta', targetTtd: '2026-09-18', aging: 15 },
    { id: 'BAA-2026-0908', name: 'PLN UID Kalselteng', service: 'Smart SCADA Banjarbaru - Palangkaraya', value: 9.60, bidang: 'PLN UID/UIW', sbu: 'SBU Kalimantan', pic: 'Agus Setiawan', sales: 'Yuliana Sari', targetTtd: '2026-09-20', aging: 12 },
    { id: 'BAA-2026-0909', name: 'PLN UID Sulselrabar', service: 'CCTV Surveillance & Access Control GI', value: 8.50, bidang: 'PLN UID/UIW', sbu: 'SBU Sulawesi & IBT', pic: 'Ilham Syam', sales: 'Muhammad Rifai', targetTtd: '2026-09-21', aging: 10 },
    { id: 'BAA-2026-0910', name: 'PLN UIW Papua & Papua Barat', service: 'VSAT Ku-Band 85 Site Pedalaman', value: 7.20, bidang: 'PLN UID/UIW', sbu: 'SBU Sulawesi & IBT', pic: 'Yohanes Wenda', sales: 'Fransiscus', targetTtd: '2026-09-22', aging: 8 },
    { id: 'BAA-2026-0911', name: 'PLN UID Jawa Barat', service: 'Interkoneksi Fiber Optic Gardu Hubung Bandung', value: 6.80, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Barat', pic: 'Asep Saepuloh', sales: 'Gita Pertiwi', targetTtd: '2026-09-15', aging: 21 },
    { id: 'BAA-2026-0912', name: 'PLN UID Lampung', service: 'Kanal Komunikasi Dispatcher APD', value: 6.10, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Selatan', pic: 'Ridwan Kamil', sales: 'Nurhasanah', targetTtd: '2026-09-16', aging: 17 },
    { id: 'BAA-2026-0913', name: 'PLN Indonesia Power UP Suralaya', service: 'Data Center Monitoring & Sensor Thermal', value: 5.70, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jakarta & Banten', pic: 'Dodi Hidayat', sales: 'Arief Rahman', targetTtd: '2026-09-18', aging: 14 },
    { id: 'BAA-2026-0914', name: 'PLN UIT Jawa Bagian Tengah', service: 'Sistem SCADA Ungaran 500kV', value: 5.20, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Tengah', pic: 'Suhartono', sales: 'Tri Wulandari', targetTtd: '2026-09-19', aging: 13 },
    { id: 'BAA-2026-0915', name: 'PLN UID Riau & Kepri', service: 'Batam - Bintan Submarine Fiber Link', value: 4.80, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Zulfikar', sales: 'Robby Firmansyah', targetTtd: '2026-09-20', aging: 11 },
    { id: 'BAA-2026-0916', name: 'PLN UIP Maluku & Papua', service: 'Radio Link Jayapura Substation', value: 4.30, bidang: 'PLN UIT/UIP', sbu: 'SBU Sulawesi & IBT', pic: 'Markus Tabuni', sales: 'Alamsyah', targetTtd: '2026-09-22', aging: 7 },
    { id: 'BAA-2026-0917', name: 'PLN UID Sumatera Barat', service: 'Jaringan IP PBX Corporate Call', value: 3.90, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Syahril', sales: 'Indra Gunawan', targetTtd: '2026-09-23', aging: 6 },
    { id: 'BAA-2026-0918', name: 'PLN UIW NTT', service: 'Satelit VSAT Kupang & Flores', value: 3.50, bidang: 'PLN UID/UIW', sbu: 'SBU Bali & Nusa Tenggara', pic: 'Stefanus Nabu', sales: 'Daniel Keraf', targetTtd: '2026-09-24', aging: 5 },
    { id: 'BAA-2026-0919', name: 'PLN Pusdiklat', service: 'E-Learning Platform & Streaming Server', value: 3.10, bidang: 'PLN Holding', sbu: 'SBU Jakarta & Banten', pic: 'Gunawan Prasetyo', sales: 'Anisa Maharani', targetTtd: '2026-09-25', aging: 4 },
    { id: 'BAA-2026-0920', name: 'PLN Gas & Geothermal', service: 'Jaringan Komunikasi Remote Wellhead', value: 2.80, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jawa Bagian Barat', pic: 'Bagus Prakoso', sales: 'Faisal Basri', targetTtd: '2026-09-25', aging: 4 },
    { id: 'BAA-2026-0921', name: 'PLN UID Kalbar', service: 'Fiber Optic Pontianak Metro', value: 2.50, bidang: 'PLN UID/UIW', sbu: 'SBU Kalimantan', pic: 'Firmansyah', sales: 'Lianawati', targetTtd: '2026-09-26', aging: 3 },
    { id: 'BAA-2026-0922', name: 'PLN UIT Sumatera Bagian Utara', service: 'OPGW Pangkalan Brandan', value: 2.30, bidang: 'PLN UIT/UIP', sbu: 'SBU Sumatera Bagian Utara', pic: 'Taufik Simatupang', sales: 'Romi Siregar', targetTtd: '2026-09-26', aging: 3 },
    { id: 'BAA-2026-0923', name: 'PLN UID Bangka Belitung', service: 'Kabel Laut Interkoneksi Pulau', value: 2.10, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Selatan', pic: 'Hermawan', sales: 'Wahyudin', targetTtd: '2026-09-27', aging: 2 },
    { id: 'BAA-2026-0924', name: 'PLN UIP Interkoneksi Sumatera-Jawa', service: 'Kabel Transmisi Selat Sunda Monitoring', value: 1.90, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Barat', pic: 'Agung Laksono', sales: 'Harry Sanjaya', targetTtd: '2026-09-27', aging: 2 },
    { id: 'BAA-2026-0925', name: 'PLN Haleyora Power Wilayah 3', service: 'IoT Tracking Kendaraan Yantek', value: 1.80, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jawa Bagian Timur', pic: 'Joko Widodo', sales: 'Rahmat Hidayat', targetTtd: '2026-09-27', aging: 2 },
    { id: 'BAA-2026-0926', name: 'PLN UID Kaltimra', service: 'IP Transit Samarinda', value: 1.60, bidang: 'PLN UID/UIW', sbu: 'SBU Kalimantan', pic: 'Rusli', sales: 'Bunga Citra', targetTtd: '2026-09-28', aging: 1 },
    { id: 'BAA-2026-0927', name: 'PLN UIP JBT UPK 1', service: 'CCTV Gardu Induk Kesugihan', value: 1.50, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Tengah', pic: 'Danang', sales: 'Triyono', targetTtd: '2026-09-28', aging: 1 },
    { id: 'BAA-2026-0928', name: 'PLN UID Jateng UP3 Semarang', service: 'Managed LAN Switch 24 Port', value: 1.40, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Tengah', pic: 'Supriyadi', sales: 'Kusuma Wardani', targetTtd: '2026-09-28', aging: 1 },
    { id: 'BAA-2026-0929', name: 'PLN Batam', service: 'Data Center Colocation', value: 1.30, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Sumatera Bagian Tengah', pic: 'David Tan', sales: 'Susanto', targetTtd: '2026-09-29', aging: 1 },
    { id: 'BAA-2026-0930', name: 'PLN UID Jatim UP3 Malang', service: 'Smart Meter Gateway AMI', value: 1.20, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Timur', pic: 'Hartono', sales: 'Erna Wati', targetTtd: '2026-09-29', aging: 0 },
    { id: 'BAA-2026-0931', name: 'PLN Icon Plus Regional Sumbagteng', service: 'Internal Monitoring Node', value: 1.10, bidang: 'PLN Holding', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Fauzan', sales: 'Nurul', targetTtd: '2026-09-29', aging: 0 },
    { id: 'BAA-2026-0932', name: 'PLN UID Suluttenggo', service: 'Dedicated Bandwidth Manado', value: 1.00, bidang: 'PLN UID/UIW', sbu: 'SBU Sulawesi & IBT', pic: 'Wongkar', sales: 'Monika', targetTtd: '2026-09-29', aging: 0 },
    { id: 'BAA-2026-0933', name: 'PLN Enjiniring Proyek PLTS', service: 'Komunikasi SCADA PLTS Cirata', value: 0.90, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jawa Bagian Barat', pic: 'Prabowo Subianto', sales: 'Gibran', targetTtd: '2026-09-30', aging: 0 },
    { id: 'BAA-2026-0934', name: 'PLN UIP Sulawesi UPK 2', service: 'Radio Link Kendari', value: 0.80, bidang: 'PLN UIT/UIP', sbu: 'SBU Sulawesi & IBT', pic: 'Kahar', sales: 'Marlina', targetTtd: '2026-09-30', aging: 0 },
    { id: 'BAA-2026-0935', name: 'PLN UID Bali UP3 Denpasar', service: 'Access Point Wireless WiFi', value: 0.70, bidang: 'PLN UID/UIW', sbu: 'SBU Bali & Nusa Tenggara', pic: 'Nyoman', sales: 'Ketut', targetTtd: '2026-09-30', aging: 0 },
    { id: 'BAA-2026-0936', name: 'PLN Pusat Pemeliharaan Ketenagalistrikan', service: 'Aplikasi Manajemen Aset Gardu', value: 0.60, bidang: 'PLN Holding', sbu: 'SBU Jawa Bagian Barat', pic: 'Slamet', sales: 'Pujiastuti', targetTtd: '2026-09-30', aging: 0 }
  ];

  const PA_DATA = [
    { id: 'PA-2026-0842', name: 'PLN UIP Kalimantan Bagian Barat', service: 'Penarikan Kabel OPGW 150kV Tayan - Sandai', value: 54.80, bidang: 'PLN UIT/UIP', sbu: 'SBU Kalimantan', pic: 'Eko Wahyudi', sales: 'Taufik Hidayat', targetAktivasi: '2026-08-25', aging: 58 },
    { id: 'PA-2026-0843', name: 'PLN UID Sumatera Barat', service: 'Sistem Smart Metering AMI Tahap 3 Padang', value: 42.10, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Denny Siregar', sales: 'Nadia Putri', targetAktivasi: '2026-09-02', aging: 44 },
    { id: 'PA-2026-0844', name: 'PLN Indonesia Power PLTU Suralaya', service: 'Jaringan IoT Sensor Boiler & Turbin', value: 36.50, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jakarta & Banten', pic: 'Fajar Nugroho', sales: 'Arif Budiman', targetAktivasi: '2026-09-08', aging: 36 },
    { id: 'PA-2026-0845', name: 'PLN UID Bali', service: 'EV Charging Station Hub Fast Charger', value: 31.20, bidang: 'PLN UID/UIW', sbu: 'SBU Bali & Nusa Tenggara', pic: 'Maya Indah', sales: 'Wayan Sudira', targetAktivasi: '2026-09-12', aging: 31 },
    { id: 'PA-2026-0846', name: 'PLN UID Aceh', service: 'Sistem Radio Komunikasi DMR Banda Aceh', value: 27.80, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Utara', pic: 'Rizky Ramadhan', sales: 'Teuku Iskandar', targetAktivasi: '2026-09-15', aging: 25 },
    { id: 'PA-2026-0847', name: 'PLN UIT Jawa Bagian Barat', service: 'Fiber Optic Underground Cawang - Gandul', value: 24.00, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Barat', pic: 'Anwar Sadat', sales: 'Maya Safitri', targetAktivasi: '2026-09-18', aging: 19 },
    { id: 'PA-2026-0848', name: 'PT PLN (Persero) Kantor Pusat', service: 'Integrasi Cloud SAP S/4HANA Finance', value: 19.60, bidang: 'PLN Holding', sbu: 'SBU Kantor Pusat', pic: 'Tri Prabowo', sales: 'Cynthia Dewi', targetAktivasi: '2026-09-20', aging: 16 },
    { id: 'PA-2026-0849', name: 'PLN UIP Maluku & Papua', service: 'Infrastruktur IT Kantor Proyek Ambon', value: 16.30, bidang: 'PLN UIT/UIP', sbu: 'SBU Sulawesi & IBT', pic: 'Samuel Latumahina', sales: 'Agustinus', targetAktivasi: '2026-09-22', aging: 14 },
    { id: 'PA-2026-0850', name: 'PLN UID Jawa Barat UP3 Bekasi', service: 'Modernisasi Network LAN 18 ULP', value: 14.50, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Barat', pic: 'Dwi Cahyono', sales: 'Lilis Suryani', targetAktivasi: '2026-09-24', aging: 11 },
    { id: 'PA-2026-0851', name: 'PLN Pusat Sertifikasi', service: 'Laboratorium Kalibrasi Smart Grid', value: 11.20, bidang: 'PLN Holding', sbu: 'SBU Kantor Pusat', pic: 'Hadi Sucipto', sales: 'Bayu Pamungkas', targetAktivasi: '2026-09-26', aging: 8 },
    { id: 'PA-2026-0852', name: 'PLN UID Jawa Timur UP3 Surabaya Selatan', service: 'Automated Meter Reading Industri Rungkut', value: 10.40, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Timur', pic: 'Surya Dharma', sales: 'Anindita', targetAktivasi: '2026-09-22', aging: 14 },
    { id: 'PA-2026-0853', name: 'PLN Nusantara Power UP Paiton', service: 'Backup Jaringan Dark Fiber Paiton', value: 9.80, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jawa Bagian Timur', pic: 'Darmawan', sales: 'Heru Prasetyo', targetAktivasi: '2026-09-24', aging: 11 },
    { id: 'PA-2026-0854', name: 'PLN UIT Sumatera Bagian Selatan', service: 'Kabel OPGW Keramasan - Simpang Tiga', value: 9.20, bidang: 'PLN UIT/UIP', sbu: 'SBU Sumatera Bagian Selatan', pic: 'Bambang Sudiro', sales: 'Yosep Rizal', targetAktivasi: '2026-09-25', aging: 9 },
    { id: 'PA-2026-0855', name: 'PLN UID Jateng & DIY UP3 Surakarta', service: 'Smart Distribution Transformer Monitor', value: 8.60, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Tengah', pic: 'Prasetyo Utomo', sales: 'Endah Sulastri', targetAktivasi: '2026-09-25', aging: 9 },
    { id: 'PA-2026-0856', name: 'PLN UIW Kalimantan Selatan & Tengah', service: 'Sistem Teleproteksi GI Sampit', value: 8.00, bidang: 'PLN UID/UIW', sbu: 'SBU Kalimantan', pic: 'Syamsul Bahri', sales: 'Dian Anggraini', targetAktivasi: '2026-09-26', aging: 7 },
    { id: 'PA-2026-0857', name: 'PLN UID Lampung UP3 Kotabumi', service: 'Relokasi Jalur FO Pelebaran Jalan', value: 7.50, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Selatan', pic: 'Indra Jaya', sales: 'Ferry Irawan', targetAktivasi: '2026-09-27', aging: 6 },
    { id: 'PA-2026-0858', name: 'PLN Puslitbang Duren Tiga', service: 'Supercomputer Server Rack Installation', value: 7.10, bidang: 'PLN Holding', sbu: 'SBU Kantor Pusat', pic: 'Dr. Arifin', sales: 'Rangga Pratama', targetAktivasi: '2026-09-27', aging: 6 },
    { id: 'PA-2026-0859', name: 'PLN UID Riau UP3 Pekanbaru', service: 'IP VPN Redundant Link Kantor Wilayah', value: 6.70, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Nasrun', sales: 'Helmi Hasan', targetAktivasi: '2026-09-28', aging: 5 },
    { id: 'PA-2026-0860', name: 'PLN UIP Sulawesi UPK Gorontalo', service: 'Radio Komunikasi GI Isimu', value: 6.20, bidang: 'PLN UIT/UIP', sbu: 'SBU Sulawesi & IBT', pic: 'Rahmat Isa', sales: 'Novita Sari', targetAktivasi: '2026-09-28', aging: 5 },
    { id: 'PA-2026-0861', name: 'PLN UIP Nusra UPK Kupang', service: 'CCTV Perimeter GI Bolok', value: 5.80, bidang: 'PLN UIT/UIP', sbu: 'SBU Bali & Nusa Tenggara', pic: 'Yoseph Bala', sales: 'Marselinus', targetAktivasi: '2026-09-29', aging: 3 },
    { id: 'PA-2026-0862', name: 'PLN Indonesia Power UP Saguling', service: 'Sensor Debit Air & Dam Safety Monitoring', value: 5.40, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jawa Bagian Barat', pic: 'Dadang Gunawan', sales: 'Cecep Suryana', targetAktivasi: '2026-09-29', aging: 3 },
    { id: 'PA-2026-0863', name: 'PLN UID Sumut UP3 Medan', service: 'Backbone Fiber Optic Kawasan Industri KIM', value: 5.10, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Utara', pic: 'Harun Lubis', sales: 'Zulfahmi', targetAktivasi: '2026-09-29', aging: 3 },
    { id: 'PA-2026-0864', name: 'PLN UID Jakarta Raya UP3 Menteng', service: 'Integrasi Fiber Optic MRT Fase 2', value: 4.80, bidang: 'PLN UID/UIW', sbu: 'SBU Jakarta & Banten', pic: 'Irfan Syahputra', sales: 'Aditya Nugraha', targetAktivasi: '2026-09-30', aging: 2 },
    { id: 'PA-2026-0865', name: 'PLN UIW Papua UP3 Sorong', service: 'VSAT Site Raja Ampat', value: 4.50, bidang: 'PLN UID/UIW', sbu: 'SBU Sulawesi & IBT', pic: 'Petrus Kambuaya', sales: 'Dominggus', targetAktivasi: '2026-09-30', aging: 2 },
    { id: 'PA-2026-0866', name: 'PLN UIT JBT UPT Bandung', service: 'Perangkat SDH Transmisi Padalarang', value: 4.20, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Barat', pic: 'Agus Salim', sales: 'Robby Sugara', targetAktivasi: '2026-09-30', aging: 2 },
    { id: 'PA-2026-0867', name: 'PLN UID Jatim UP3 Jember', service: 'Router Edge Coba Uji Petik AMI', value: 3.90, bidang: 'PLN UID/UIW', sbu: 'SBU Jawa Bagian Timur', pic: 'Miftahul Huda', sales: 'Luluk Maslucha', targetAktivasi: '2026-09-30', aging: 2 },
    { id: 'PA-2026-0868', name: 'PLN Enjiniring Head Office', service: 'Cloud Workstation GPU Processing', value: 3.60, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Kantor Pusat', pic: 'Wisnu Wardhana', sales: 'Gilang Ramadhan', targetAktivasi: '2026-09-30', aging: 2 },
    { id: 'PA-2026-0869', name: 'PLN UID Sumbar UP3 Bukittinggi', service: 'Dedicated Radio Jam Gadang Link', value: 3.30, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Marlius', sales: 'Erwan Tanjung', targetAktivasi: '2026-09-30', aging: 2 },
    { id: 'PA-2026-0870', name: 'PLN UIP JBB UPK Cilegon', service: 'Access Point Luar Ruang Gardu Cilegon', value: 3.00, bidang: 'PLN UIT/UIP', sbu: 'SBU Jakarta & Banten', pic: 'Tubagus Haidar', sales: 'Entin Kartini', targetAktivasi: '2026-09-30', aging: 1 },
    { id: 'PA-2026-0871', name: 'PLN Batam UP3 Nagoya', service: 'Jaringan Fiber Optic Perkantoran', value: 2.70, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Sumatera Bagian Tengah', pic: 'Kelvin', sales: 'Stella', targetAktivasi: '2026-09-30', aging: 1 },
    { id: 'PA-2026-0872', name: 'PLN UID Kalselteng UP3 Palangkaraya', service: 'Jaringan Antar Kantor Sub ULP', value: 2.40, bidang: 'PLN UID/UIW', sbu: 'SBU Kalimantan', pic: 'Yance', sales: 'Neneng', targetAktivasi: '2026-09-30', aging: 1 },
    { id: 'PA-2026-0873', name: 'PLN Indonesia Power UP Cirata', service: 'Hydro Sensor Network', value: 2.10, bidang: 'PLN Anak Perusahaan', sbu: 'SBU Jawa Bagian Barat', pic: 'Iwan Setiawan', sales: 'Ujang Koswara', targetAktivasi: '2026-09-30', aging: 1 },
    { id: 'PA-2026-0874', name: 'PLN UID Bali UP3 Kuta', service: 'Free WiFi Destinasi Pantai Pandawa', value: 1.80, bidang: 'PLN UID/UIW', sbu: 'SBU Bali & Nusa Tenggara', pic: 'Gede Suartana', sales: 'Ni Made Rai', targetAktivasi: '2026-09-30', aging: 1 },
    { id: 'PA-2026-0875', name: 'PLN UIP JBT UPK 2 Solo', service: 'Interkoneksi Fiber Optic GI Palur', value: 1.50, bidang: 'PLN UIT/UIP', sbu: 'SBU Jawa Bagian Tengah', pic: 'Sugeng Raharjo', sales: 'Retno Wulandari', targetAktivasi: '2026-09-30', aging: 0 },
    { id: 'PA-2026-0876', name: 'PLN UID Sulselrabar UP3 Makassar', service: 'IP Transit 10G Center', value: 1.30, bidang: 'PLN UID/UIW', sbu: 'SBU Sulawesi & IBT', pic: 'Baso Amir', sales: 'Hasnah', targetAktivasi: '2026-09-30', aging: 0 },
    { id: 'PA-2026-0877', name: 'PLN UIW Maluku UP3 Ternate', service: 'Link Radio Gelombang Mikro', value: 1.10, bidang: 'PLN UID/UIW', sbu: 'SBU Sulawesi & IBT', pic: 'Faisal Taher', sales: 'Rosmini', targetAktivasi: '2026-09-30', aging: 0 },
    { id: 'PA-2026-0878', name: 'PLN UID Aceh UP3 Meulaboh', service: 'Jaringan Komunikasi ULP Calang', value: 0.90, bidang: 'PLN UID/UIW', sbu: 'SBU Sumatera Bagian Utara', pic: 'Maimun', sales: 'Cut Intan', targetAktivasi: '2026-09-30', aging: 0 },
    { id: 'PA-2026-0879', name: 'PLN Pusat Pemeliharaan Bandung', service: 'Mobile Diagnostic Tool Link', value: 0.70, bidang: 'PLN Holding', sbu: 'SBU Jawa Bagian Barat', pic: 'Tateng', sales: 'Ayu', targetAktivasi: '2026-09-30', aging: 0 },
    { id: 'PA-2026-0880', name: 'PLN Icon Plus Regional Jawa Timur', service: 'POP Gateway Waru Sidoarjo', value: 0.50, bidang: 'PLN Holding', sbu: 'SBU Jawa Bagian Timur', pic: 'Bagus Setyawan', sales: 'Ratna', targetAktivasi: '2026-09-30', aging: 0 }
  ];

  // Pagination & Search States
  let exceptionCurrentPage = 1;
  const exceptionPageSize = 10;
  let exceptionSearchQuery = '';

  let baaCurrentPage = 1;
  const baaPageSize = 10;
  let baaSearchQuery = '';

  let paCurrentPage = 1;
  const paPageSize = 10;
  let paSearchQuery = '';

  // Render & Logic Prognosa Corporate
  function renderPrognosaView() {
    const p = appData.prognosa || DEFAULT_PROGNOSA;
    const srcEl = document.getElementById('prog-active-source');
    if (srcEl && p.sourceFile) srcEl.textContent = p.sourceFile;

    // Nilai 7 Kartu KPI Utama Prognosa (Satuan Triliun Rupiah / T)
    const tgtKum = (p.targetKumulatif !== undefined) ? p.targetKumulatif : DEFAULT_PROGNOSA.targetKumulatif;
    const tgtSa = (p.targetStandAlone !== undefined) ? p.targetStandAlone : DEFAULT_PROGNOSA.targetStandAlone;
    const realisasi = (p.realisasiLalu !== undefined) ? p.realisasiLalu : DEFAULT_PROGNOSA.realisasiLalu;
    const asIs = (p.prognosaAsIs !== undefined) ? p.prognosaAsIs : DEFAULT_PROGNOSA.prognosaAsIs;
    const mostLikely = (p.prognosaMostLikely !== undefined) ? p.prognosaMostLikely : DEFAULT_PROGNOSA.prognosaMostLikely;

    const gap = asIs - tgtKum;
    const gapPct = tgtKum > 0 ? ((gap / tgtKum) * 100).toFixed(1) : '0,0';
    const achievePct = tgtKum > 0 ? ((asIs / tgtKum) * 100).toFixed(1) : '0,0';

    // 1. Target Kumulatif Card
    const cTgtKum = document.getElementById('prog-card-target-kum');
    if (cTgtKum) cTgtKum.innerHTML = `<span>${Number(tgtKum).toFixed(2).replace('.', ',')}</span><span class="pln-kpi-unit">T</span>`;
    const cTgtKumSub = document.getElementById('prog-card-target-kum-sub');
    if (cTgtKumSub) cTgtKumSub.innerHTML = `<span class="kpi-tag" style="background: rgba(10, 37, 64, 0.08); color: #0A2540; font-size: 10px; padding: 2px 6px;">Target RKAP sd Sep</span>`;

    // 2. Target Stand Alone Card
    const cTgtSa = document.getElementById('prog-card-target-sa');
    if (cTgtSa) cTgtSa.innerHTML = `<span>${Number(tgtSa).toFixed(2).replace('.', ',')}</span><span class="pln-kpi-unit">T</span>`;
    const cTgtSaSub = document.getElementById('prog-card-target-sa-sub');
    if (cTgtSaSub) cTgtSaSub.innerHTML = `<span class="kpi-tag" style="background: rgba(0, 163, 224, 0.12); color: #0072CE; font-size: 10px; padding: 2px 6px;">Target Bulan Sep</span>`;

    // 3. Realisasi s.d Agustus Card
    const cReal = document.getElementById('prog-card-realisasi');
    if (cReal) cReal.innerHTML = `<span>${Number(realisasi).toFixed(2).replace('.', ',')}</span><span class="pln-kpi-unit">T</span>`;
    const cRealSub = document.getElementById('prog-card-realisasi-sub');
    if (cRealSub) {
      const rPct = tgtKum > 0 ? ((realisasi / tgtKum) * 100).toFixed(1).replace('.', ',') : '0,0';
      cRealSub.innerHTML = `<span class="kpi-tag" style="background: #ECFDF5; color: #047857; font-size: 10px; padding: 2px 6px;">${rPct}% dari target</span>`;
    }

    // 4. Prognosa AS IS Card
    const cAsIs = document.getElementById('prog-card-asis');
    if (cAsIs) cAsIs.innerHTML = `<span>${Number(asIs).toFixed(2).replace('.', ',')}</span><span class="pln-kpi-unit">T</span>`;
    const cAsIsSub = document.getElementById('prog-card-asis-sub');
    if (cAsIsSub) {
      const aPct = tgtKum > 0 ? ((asIs / tgtKum) * 100).toFixed(1).replace('.', ',') : '0,0';
      cAsIsSub.innerHTML = `<span class="kpi-tag" style="background: #EEF2FF; color: #4338CA; font-size: 10px; padding: 2px 6px;">${aPct}% dari target</span>`;
    }

    // 5. Prognosa Most Likely Card
    const cMl = document.getElementById('prog-card-mostlikely');
    if (cMl) cMl.innerHTML = `<span>${Number(mostLikely).toFixed(2).replace('.', ',')}</span><span class="pln-kpi-unit">T</span>`;
    const cMlSub = document.getElementById('prog-card-mostlikely-sub');
    if (cMlSub) {
      const mlPct = tgtKum > 0 ? ((mostLikely / tgtKum) * 100).toFixed(1).replace('.', ',') : '0,0';
      cMlSub.innerHTML = `<span class="kpi-tag" style="background: #FFFBEB; color: #B45309; font-size: 10px; padding: 2px 6px;">${mlPct}% dari target</span>`;
    }

    // 6. Gap AS IS vs Target Card
    const cGap = document.getElementById('prog-card-gap');
    const cGapWrap = document.getElementById('prog-card-gap-wrap');
    if (cGap) {
      cGap.innerHTML = `<span>${gap >= 0 ? '+' : ''}${Number(gap).toFixed(2).replace('.', ',')}</span><span class="pln-kpi-unit">T</span>`;
      cGap.style.color = gap >= 0 ? '#047857' : '#DC2626';
    }
    if (cGapWrap) {
      cGapWrap.style.borderTopColor = gap >= 0 ? '#10B981' : '#EF4444';
    }
    const cGapSub = document.getElementById('prog-card-gap-sub');
    if (cGapSub) {
      const isPos = gap >= 0;
      cGapSub.innerHTML = `<span class="kpi-tag" style="background: ${isPos ? '#ECFDF5' : '#FEF2F2'}; color: ${isPos ? '#047857' : '#DC2626'}; font-size: 10px; padding: 2px 6px;">${isPos ? '+' : ''}${gapPct.replace('.', ',')}%</span>`;
    }

    // 7. Achievement AS IS to Target Card
    const cAch = document.getElementById('prog-card-achieve');
    if (cAch) cAch.innerHTML = `<span>${achievePct.replace('.', ',')}</span><span class="pln-kpi-unit">%</span>`;
    const cAchSub = document.getElementById('prog-card-achieve-sub');
    if (cAchSub) {
      cAchSub.innerHTML = `<span style="font-size: 9.5px; color: var(--sub);">Target ${Number(tgtKum).toFixed(2).replace('.', ',')} T · Prog ${Number(asIs).toFixed(2).replace('.', ',')} T</span>`;
    }

    // Header Waterfall Chart Tags
    const wfAsIsTag = document.getElementById('prog-wf-asis-tag');
    if (wfAsIsTag) wfAsIsTag.textContent = `Total AS IS: ${Number(asIs).toFixed(2).replace('.', ',')} T`;
    const wfMlTag = document.getElementById('prog-wf-ml-tag');
    if (wfMlTag) wfMlTag.textContent = `Most Likely: ${Number(mostLikely).toFixed(2).replace('.', ',')} T`;

    // Progress terhadap RKAP 2026
    const rkapRaw = (p.targetRkap !== undefined) ? p.targetRkap : DEFAULT_PROGNOSA.targetRkap;
    const rkapT = rkapRaw >= 100 ? (rkapRaw / 1000) : rkapRaw;
    const rkapTStr = Number(rkapT).toFixed(3).replace('.', ',');

    const rkapTitle = document.getElementById('prog-rkap-title-target');
    if (rkapTitle) rkapTitle.textContent = `${rkapTStr} T`;
    const rkapBadge = document.getElementById('prog-rkap-badge-target');
    if (rkapBadge) rkapBadge.textContent = `Target RKAP: ${rkapTStr} T`;

    // AS IS vs RKAP
    const asisRkapPct = rkapT > 0 ? ((asIs / rkapT) * 100).toFixed(1) : '0,0';
    const rkapAsisText = document.getElementById('prog-rkap-asis-text');
    if (rkapAsisText) rkapAsisText.innerHTML = `AS IS vs RKAP: <strong style="color: #4338CA;">${Number(asIs).toFixed(3).replace('.', ',')} T</strong> dari ${rkapTStr} T`;
    const rkapAsisPct = document.getElementById('prog-rkap-asis-pct');
    if (rkapAsisPct) rkapAsisPct.textContent = `${asisRkapPct.replace('.', ',')}%`;
    const rkapAsisBar = document.getElementById('prog-rkap-asis-bar');
    if (rkapAsisBar) rkapAsisBar.style.width = `${Math.min(100, Math.max(0, parseFloat(asisRkapPct)))}%`;

    // Most Likely vs RKAP
    const mlRkapPct = rkapT > 0 ? ((mostLikely / rkapT) * 100).toFixed(1) : '0,0';
    const rkapMlText = document.getElementById('prog-rkap-ml-text');
    if (rkapMlText) rkapMlText.innerHTML = `Most Likely vs RKAP: <strong style="color: #047857;">${Number(mostLikely).toFixed(3).replace('.', ',')} T</strong> dari ${rkapTStr} T`;
    const rkapMlPct = document.getElementById('prog-rkap-ml-pct');
    if (rkapMlPct) rkapMlPct.textContent = `${mlRkapPct.replace('.', ',')}%`;
    const rkapMlBar = document.getElementById('prog-rkap-ml-bar');
    if (rkapMlBar) rkapMlBar.style.width = `${Math.min(100, Math.max(0, parseFloat(mlRkapPct)))}%`;

    // Remaining to RKAP
    const remVal = Math.max(0, rkapT - mostLikely);
    const remPct = rkapT > 0 ? ((remVal / rkapT) * 100).toFixed(1) : '0,0';
    const rkapRem = document.getElementById('prog-rkap-remaining');
    if (rkapRem) {
      rkapRem.innerHTML = `${Number(remVal).toFixed(2).replace('.', ',')} T <span style="font-size: 11px; font-weight: 600; color: var(--sub);">(${remPct.replace('.', ',')}%)</span>`;
    }

    // Update live clock
    updateLiveRefreshClock();

    // Render 3 Operational Tables
    renderExceptionTable();
    renderBaaTable();
    renderPaTable();
  }

  function applyPrognosaFilters() {
    const period = document.getElementById('pln-prog-filter-period') ? document.getElementById('pln-prog-filter-period').value : '9';
    const sbu = document.getElementById('pln-prog-filter-sbu') ? document.getElementById('pln-prog-filter-sbu').value : 'ALL';
    const bidang = document.getElementById('pln-prog-filter-bidang') ? document.getElementById('pln-prog-filter-bidang').value : 'ALL';

    // Reset pagination to first page
    exceptionCurrentPage = 1;
    baaCurrentPage = 1;
    paCurrentPage = 1;

    renderExceptionTable();
    renderBaaTable();
    renderPaTable();
    showToast(`Filter diterapkan: Periode Bulan ${period} · ${sbu === 'ALL' ? 'Semua SBU' : sbu} · ${bidang === 'ALL' ? 'Semua Bidang' : bidang}`);
  }

  // --- Table 1: Exception ---
  function getFilteredExceptionData() {
    const sbu = document.getElementById('pln-prog-filter-sbu') ? document.getElementById('pln-prog-filter-sbu').value : 'ALL';
    const bidang = document.getElementById('pln-prog-filter-bidang') ? document.getElementById('pln-prog-filter-bidang').value : 'ALL';
    const q = exceptionSearchQuery.toLowerCase().trim();

    return EXCEPTION_DATA.filter(item => {
      const matchBidang = (bidang === 'ALL' || item.bidang === bidang);
      const matchSearch = (!q || item.name.toLowerCase().includes(q) || item.service.toLowerCase().includes(q) || item.id.toLowerCase().includes(q));
      return matchBidang && matchSearch;
    });
  }

  function renderExceptionTable() {
    const tbody = document.getElementById('prog-table-exception-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const filtered = getFilteredExceptionData();
    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / exceptionPageSize));
    if (exceptionCurrentPage > totalPages) exceptionCurrentPage = totalPages;

    const startIdx = (exceptionCurrentPage - 1) * exceptionPageSize;
    const pageItems = filtered.slice(startIdx, startIdx + exceptionPageSize);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="8" style="text-align: center; padding: 20px; color: var(--muted);">Tidak ada data Exception yang sesuai filter.</td></tr>`;
    } else {
      pageItems.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const num = startIdx + idx + 1;
        tr.innerHTML = `
          <td style="text-align: center; font-weight: 700; color: var(--sub);">${num}</td>
          <td style="font-weight: 700; font-family: monospace; font-size: 11px; color: #4F46E5;">${item.id}</td>
          <td style="font-weight: 700; color: var(--navy);">${item.name}</td>
          <td style="color: var(--text);">${item.service}</td>
          <td class="num" style="font-weight: 800; color: #DC2626;">Rp ${item.value.toFixed(2).replace('.', ',')} M</td>
          <td><span class="kpi-tag" style="background: rgba(10,37,64,0.06); color: var(--navy); font-size: 10px;">${item.bidang}</span></td>
          <td style="text-align: center; font-weight: 600; color: var(--sub); font-size: 11px;">${item.thbl}</td>
          <td style="color: var(--sub); font-size: 11px;">${item.note}</td>
        `;
        tbody.appendChild(tr);
      });
    }

    const countInfo = document.getElementById('prog-exception-count-info');
    if (countInfo) {
      const endIdx = Math.min(startIdx + exceptionPageSize, total);
      countInfo.textContent = total > 0 ? `Menampilkan ${startIdx + 1}-${endIdx} dari ${total} data` : '0 data ditemukan';
    }
    const pageBadge = document.getElementById('prog-exception-page-badge');
    if (pageBadge) pageBadge.textContent = `Hal ${exceptionCurrentPage} / ${totalPages}`;
  }

  function changePageException(delta) {
    const totalPages = Math.ceil(getFilteredExceptionData().length / exceptionPageSize);
    const newPage = exceptionCurrentPage + delta;
    if (newPage >= 1 && newPage <= totalPages) {
      exceptionCurrentPage = newPage;
      renderExceptionTable();
    }
  }

  function onSearchException(query) {
    exceptionSearchQuery = query;
    exceptionCurrentPage = 1;
    renderExceptionTable();
  }

  // --- Table 2: BAA Open ---
  function getFilteredBaaData() {
    const sbu = document.getElementById('pln-prog-filter-sbu') ? document.getElementById('pln-prog-filter-sbu').value : 'ALL';
    const bidang = document.getElementById('pln-prog-filter-bidang') ? document.getElementById('pln-prog-filter-bidang').value : 'ALL';
    const q = baaSearchQuery.toLowerCase().trim();

    return BAA_DATA.filter(item => {
      const matchSbu = (sbu === 'ALL' || item.sbu === sbu);
      const matchBidang = (bidang === 'ALL' || item.bidang === bidang);
      const matchSearch = (!q || item.name.toLowerCase().includes(q) || item.id.toLowerCase().includes(q) || item.service.toLowerCase().includes(q) || item.pic.toLowerCase().includes(q));
      return matchSbu && matchBidang && matchSearch;
    });
  }

  function renderBaaTable() {
    const tbody = document.getElementById('prog-table-baa-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const filtered = getFilteredBaaData();
    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / baaPageSize));
    if (baaCurrentPage > totalPages) baaCurrentPage = totalPages;

    const startIdx = (baaCurrentPage - 1) * baaPageSize;
    const pageItems = filtered.slice(startIdx, startIdx + baaPageSize);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="11" style="text-align: center; padding: 20px; color: var(--muted);">Tidak ada data BAA Open yang sesuai filter.</td></tr>`;
    } else {
      pageItems.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const num = startIdx + idx + 1;
        const agingBadge = item.aging > 30 ? 'badge-aging-red' : (item.aging > 14 ? 'badge-aging-amber' : 'badge-aging-green');
        tr.innerHTML = `
          <td style="text-align: center; font-weight: 700; color: var(--sub);">${num}</td>
          <td style="font-weight: 700; font-family: monospace; font-size: 11px; color: #059669;">${item.id}</td>
          <td style="font-weight: 700; color: var(--navy);">${item.name}</td>
          <td style="color: var(--text);">${item.service}</td>
          <td class="num" style="font-weight: 800; color: var(--navy);">Rp ${item.value.toFixed(2).replace('.', ',')} M</td>
          <td><span class="kpi-tag" style="background: rgba(10,37,64,0.06); color: var(--navy); font-size: 10px;">${item.bidang}</span></td>
          <td style="color: var(--sub); font-size: 11px;">${item.sbu}</td>
          <td style="color: var(--navy); font-weight: 600; font-size: 11px;">${item.pic}</td>
          <td style="color: var(--sub); font-size: 11px;">${item.sales}</td>
          <td style="text-align: center; font-size: 11px; color: var(--sub);">${item.targetTtd}</td>
          <td style="text-align: center;"><span class="${agingBadge}">${item.aging} Hari</span></td>
        `;
        tbody.appendChild(tr);
      });
    }

    const countInfo = document.getElementById('prog-baa-count-info');
    if (countInfo) {
      const endIdx = Math.min(startIdx + baaPageSize, total);
      countInfo.textContent = total > 0 ? `Menampilkan ${startIdx + 1}-${endIdx} dari ${total} data` : '0 data ditemukan';
    }
    const pageBadge = document.getElementById('prog-baa-page-badge');
    if (pageBadge) pageBadge.textContent = `Hal ${baaCurrentPage} / ${totalPages}`;
  }

  function changePageBaa(delta) {
    const totalPages = Math.ceil(getFilteredBaaData().length / baaPageSize);
    const newPage = baaCurrentPage + delta;
    if (newPage >= 1 && newPage <= totalPages) {
      baaCurrentPage = newPage;
      renderBaaTable();
    }
  }

  function onSearchBaa(query) {
    baaSearchQuery = query;
    baaCurrentPage = 1;
    renderBaaTable();
  }

  // --- Table 3: PA Open ---
  function getFilteredPaData() {
    const sbu = document.getElementById('pln-prog-filter-sbu') ? document.getElementById('pln-prog-filter-sbu').value : 'ALL';
    const bidang = document.getElementById('pln-prog-filter-bidang') ? document.getElementById('pln-prog-filter-bidang').value : 'ALL';
    const q = paSearchQuery.toLowerCase().trim();

    return PA_DATA.filter(item => {
      const matchSbu = (sbu === 'ALL' || item.sbu === sbu);
      const matchBidang = (bidang === 'ALL' || item.bidang === bidang);
      const matchSearch = (!q || item.name.toLowerCase().includes(q) || item.id.toLowerCase().includes(q) || item.service.toLowerCase().includes(q) || item.pic.toLowerCase().includes(q));
      return matchSbu && matchBidang && matchSearch;
    });
  }

  function renderPaTable() {
    const tbody = document.getElementById('prog-table-pa-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    const filtered = getFilteredPaData();
    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / paPageSize));
    if (paCurrentPage > totalPages) paCurrentPage = totalPages;

    const startIdx = (paCurrentPage - 1) * paPageSize;
    const pageItems = filtered.slice(startIdx, startIdx + paPageSize);

    if (pageItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="11" style="text-align: center; padding: 20px; color: var(--muted);">Tidak ada data PA Open yang sesuai filter.</td></tr>`;
    } else {
      pageItems.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const num = startIdx + idx + 1;
        const agingBadge = item.aging > 30 ? 'badge-aging-red' : (item.aging > 14 ? 'badge-aging-amber' : 'badge-aging-green');
        tr.innerHTML = `
          <td style="text-align: center; font-weight: 700; color: var(--sub);">${num}</td>
          <td style="font-weight: 700; font-family: monospace; font-size: 11px; color: #7C3AED;">${item.id}</td>
          <td style="font-weight: 700; color: var(--navy);">${item.name}</td>
          <td style="color: var(--text);">${item.service}</td>
          <td class="num" style="font-weight: 800; color: var(--navy);">Rp ${item.value.toFixed(2).replace('.', ',')} M</td>
          <td><span class="kpi-tag" style="background: rgba(10,37,64,0.06); color: var(--navy); font-size: 10px;">${item.bidang}</span></td>
          <td style="color: var(--sub); font-size: 11px;">${item.sbu}</td>
          <td style="color: var(--navy); font-weight: 600; font-size: 11px;">${item.pic}</td>
          <td style="color: var(--sub); font-size: 11px;">${item.sales}</td>
          <td style="text-align: center; font-size: 11px; color: var(--sub);">${item.targetAktivasi}</td>
          <td style="text-align: center;"><span class="${agingBadge}">${item.aging} Hari</span></td>
        `;
        tbody.appendChild(tr);
      });
    }

    const countInfo = document.getElementById('prog-pa-count-info');
    if (countInfo) {
      const endIdx = Math.min(startIdx + paPageSize, total);
      countInfo.textContent = total > 0 ? `Menampilkan ${startIdx + 1}-${endIdx} dari ${total} data` : '0 data ditemukan';
    }
    const pageBadge = document.getElementById('prog-pa-page-badge');
    if (pageBadge) pageBadge.textContent = `Hal ${paCurrentPage} / ${totalPages}`;
  }

  function changePagePa(delta) {
    const totalPages = Math.ceil(getFilteredPaData().length / paPageSize);
    const newPage = paCurrentPage + delta;
    if (newPage >= 1 && newPage <= totalPages) {
      paCurrentPage = newPage;
      renderPaTable();
    }
  }

  function onSearchPa(query) {
    paSearchQuery = query;
    paCurrentPage = 1;
    renderPaTable();
  }

  // --- Export CSV Generator ---
  function exportTableToCsv(type) {
    let filename = '';
    let csvContent = '';

    if (type === 'exception') {
      filename = 'Detail_Exception_Prognosa_2026.csv';
      const items = getFilteredExceptionData();
      csvContent = 'No,ID Pelanggan,Nama Pelanggan,Layanan,Nilai (Rp Miliar),Bidang,THBL,Keterangan\n';
      items.forEach((it, i) => {
        csvContent += `${i + 1},"${it.id}","${it.name}","${it.service}",${it.value},"${it.bidang}","${it.thbl}","${it.note}"\n`;
      });
    } else if (type === 'baa') {
      filename = 'Top10_BAA_Open_Revenue_Terbesar.csv';
      const items = getFilteredBaaData();
      csvContent = 'No,ID BAA,Nama Pelanggan,Layanan,Nilai (Rp Miliar),Bidang,SBU,PIC,Sales,Target TTD,Aging Hari\n';
      items.forEach((it, i) => {
        csvContent += `${i + 1},"${it.id}","${it.name}","${it.service}",${it.value},"${it.bidang}","${it.sbu}","${it.pic}","${it.sales}","${it.targetTtd}",${it.aging}\n`;
      });
    } else if (type === 'pa') {
      filename = 'Top10_PA_Open_Revenue_Terbesar.csv';
      const items = getFilteredPaData();
      csvContent = 'No,ID PA,Nama Pelanggan,Layanan,Nilai (Rp Miliar),Bidang,SBU,PIC,Sales,Target Aktivasi,Aging Hari\n';
      items.forEach((it, i) => {
        csvContent += `${i + 1},"${it.id}","${it.name}","${it.service}",${it.value},"${it.bidang}","${it.sbu}","${it.pic}","${it.sales}","${it.targetAktivasi}",${it.aging}\n`;
      });
    }

    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast(`Berhasil mengekspor ${filename}`);
  }

  // --- Prognosa Audit Log Utilities & Handlers ---
  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function getInitials(name) {
    if (!name) return 'PIC';
    const parts = name.split(/[._\s-]+/).filter(Boolean);
    if (parts.length === 0) return 'PIC';
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function getFormattedTimestampWIB() {
    const d = new Date();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const day = String(d.getDate()).padStart(2, '0');
    const month = months[d.getMonth()];
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
    return `${day} ${month} ${year} ${hours}:${minutes}:${seconds} WIB`;
  }

  function getCurrentPrognosaUser() {
    const el = document.getElementById('m_prog_user_pic');
    if (el && el.value && el.value.trim() !== '') {
      return el.value.trim();
    }
    return 'dodik.gunawan';
  }

  let currentProgLogPage = 1;
  const PROG_LOG_PAGE_SIZE = 5;
  let progLogSearchQuery = '';

  function onSearchProgLog(val) {
    progLogSearchQuery = (val || '').toLowerCase().trim();
    currentProgLogPage = 1;
    renderPrognosaLogTable();
  }

  function changePageProgLog(delta) {
    currentProgLogPage += delta;
    if (currentProgLogPage < 1) currentProgLogPage = 1;
    renderPrognosaLogTable();
  }

  function addPrognosaAuditLog(entry) {
    if (!appData.prognosaLogs) {
      appData.prognosaLogs = JSON.parse(JSON.stringify(DEFAULT_PROGNOSA_LOGS));
    }
    const logItem = {
      id: 'LOG-' + Date.now(),
      timestamp: getFormattedTimestampWIB(),
      user: entry.user || getCurrentPrognosaUser(),
      userRole: 'Revenue Assurance PIC',
      action: entry.action || 'Manual Edit Form',
      source: entry.source || 'Form Master Prognosa',
      diffs: entry.diffs || [],
      note: entry.note || (entry.diffs && entry.diffs.length > 0 ? `${entry.diffs.length} parameter diubah` : 'Pembaruan data'),
      status: 'Tersimpan'
    };
    appData.prognosaLogs.unshift(logItem);
    if (appData.prognosaLogs.length > 200) {
      appData.prognosaLogs = appData.prognosaLogs.slice(0, 200);
    }
    saveAppData(appData);
    renderPrognosaLogTable();
  }

  function renderPrognosaLogTable() {
    const tbody = document.getElementById('master-prog-log-tbody');
    if (!tbody) return;

    let logs = (appData && appData.prognosaLogs && Array.isArray(appData.prognosaLogs)) 
      ? appData.prognosaLogs 
      : DEFAULT_PROGNOSA_LOGS;

    const totalLogsCount = logs.length;
    const badgeEl = document.getElementById('prog-log-total-badge');
    if (badgeEl) badgeEl.textContent = `${totalLogsCount} Log Terdaftar`;

    if (progLogSearchQuery) {
      logs = logs.filter(log => {
        const u = (log.user || '').toLowerCase();
        const a = (log.action || '').toLowerCase();
        const s = (log.source || '').toLowerCase();
        const n = (log.note || '').toLowerCase();
        const t = (log.timestamp || '').toLowerCase();
        const diffsStr = (log.diffs || []).map(d => `${d.param} ${d.before} ${d.after} ${d.change}`).join(' ').toLowerCase();
        return u.includes(progLogSearchQuery) || 
               a.includes(progLogSearchQuery) || 
               s.includes(progLogSearchQuery) || 
               n.includes(progLogSearchQuery) || 
               t.includes(progLogSearchQuery) || 
               diffsStr.includes(progLogSearchQuery);
      });
    }

    const filteredTotal = logs.length;
    const totalPages = Math.max(1, Math.ceil(filteredTotal / PROG_LOG_PAGE_SIZE));
    if (currentProgLogPage > totalPages) currentProgLogPage = totalPages;
    if (currentProgLogPage < 1) currentProgLogPage = 1;

    const startIdx = (currentProgLogPage - 1) * PROG_LOG_PAGE_SIZE;
    const pageLogs = logs.slice(startIdx, startIdx + PROG_LOG_PAGE_SIZE);

    tbody.innerHTML = '';

    if (pageLogs.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" style="text-align: center; padding: 28px 16px; color: #94A3B8;">
            <div style="font-size: 22px; margin-bottom: 6px;">📋</div>
            <div style="font-weight: 700; color: #64748B;">Tidak ada log riwayat perubahan yang cocok</div>
            <div style="font-size: 11px; margin-top: 4px;">Coba gunakan kata kunci pencarian yang lain atau simpan perubahan parameter.</div>
          </td>
        </tr>
      `;
    } else {
      pageLogs.forEach((log, idx) => {
        const rowNo = startIdx + idx + 1;
        const tr = document.createElement('tr');
        tr.style.transition = 'background 0.15s ease';
        tr.onmouseover = () => tr.style.background = '#F8FAFC';
        tr.onmouseout = () => tr.style.background = 'transparent';

        let actionBadge = '';
        if (log.action === 'Manual Edit Form') {
          actionBadge = `<span class="kpi-tag" style="background: #EEF2FF; color: #4338CA; border: 1px solid #C7D2FE; font-weight: 700; font-size: 11px;">✏️ Edit Manual</span>`;
        } else if (log.action === 'Upload File Excel') {
          actionBadge = `<span class="kpi-tag" style="background: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; font-weight: 700; font-size: 11px;">📊 Upload Excel</span>`;
        } else if (log.action === 'Selaraskan dgn RKAP') {
          actionBadge = `<span class="kpi-tag" style="background: #FEF3C7; color: #92400E; border: 1px solid #FDE68A; font-weight: 700; font-size: 11px;">⚡ Sync RKAP</span>`;
        } else if (log.action === 'Reset ke Default') {
          actionBadge = `<span class="kpi-tag" style="background: #FEF2F2; color: #B91C1C; border: 1px solid #FECACA; font-weight: 700; font-size: 11px;">🔄 Reset Data</span>`;
        } else {
          actionBadge = `<span class="kpi-tag" style="background: #F1F5F9; color: #334155; font-weight: 700; font-size: 11px;">${escapeHtml(log.action || '-')}</span>`;
        }

        const tsParts = (log.timestamp || '').split(' ');
        const dateStr = tsParts.length >= 3 ? `${tsParts[0]} ${tsParts[1]} ${tsParts[2]}` : (log.timestamp || '-');
        const timeStr = tsParts.length >= 4 ? tsParts.slice(3).join(' ') : '';

        let diffsHtml = '';
        if (!log.diffs || log.diffs.length === 0) {
          diffsHtml = `<span style="color: #94A3B8; font-style: italic; font-size: 11px;">Tidak ada perubahan angka (Konfirmasi Simpan)</span>`;
        } else {
          diffsHtml = `<div style="display: flex; flex-direction: column; gap: 5px;">` +
            log.diffs.map(d => {
              const chg = d.change || '';
              let chgBg = '#F1F5F9';
              let chgCol = '#475569';
              if (chg.startsWith('+')) {
                chgBg = '#DCFCE7'; chgCol = '#15803D';
              } else if (chg.startsWith('-')) {
                chgBg = '#FEE2E2'; chgCol = '#B91C1C';
              } else if (chg === 'Baru' || chg === 'Update') {
                chgBg = '#EEF2FF'; chgCol = '#4338CA';
              } else if (chg === 'Reset') {
                chgBg = '#FEF2F2'; chgCol = '#B91C1C';
              }
              return `
                <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 4px 8px; font-size: 11px; display: inline-flex; align-items: center; gap: 6px; flex-wrap: wrap; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
                  <span style="font-weight: 700; color: #1E293B;">${escapeHtml(d.param)}:</span>
                  <span style="color: #64748B; text-decoration: line-through;">${escapeHtml(d.before)}</span>
                  <span style="font-weight: 800; color: #0F172A;">&rarr; ${escapeHtml(d.after)}</span>
                  ${chg ? `<span style="font-weight: 800; font-size: 10px; padding: 1px 5px; border-radius: 4px; background: ${chgBg}; color: ${chgCol};">${escapeHtml(chg)}</span>` : ''}
                </div>
              `;
            }).join('') +
          `</div>`;
        }

        tr.innerHTML = `
          <td style="text-align: center; font-weight: 700; color: #64748B;">${rowNo}</td>
          <td>
            <div style="font-weight: 700; color: #0F172A; font-size: 12px;">${dateStr}</div>
            <div style="font-size: 10.5px; color: #64748B; display: flex; align-items: center; gap: 4px; margin-top: 2px;">
              <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
              ${timeStr}
            </div>
          </td>
          <td>
            <div style="display: flex; align-items: center; gap: 8px;">
              <div style="width: 28px; height: 28px; border-radius: 50%; background: linear-gradient(135deg, #4F46E5, #00A3E0); color: #fff; font-size: 11px; font-weight: 800; display: flex; align-items: center; justify-content: center; flex-shrink: 0; box-shadow: 0 2px 4px rgba(79, 70, 229, 0.2);">
                ${getInitials(log.user)}
              </div>
              <div style="min-width: 0;">
                <div style="font-weight: 800; color: #0F172A; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(log.user)}</div>
                <div style="font-size: 10px; color: #64748B;">Revenue Monitoring</div>
              </div>
            </div>
          </td>
          <td>${actionBadge}</td>
          <td>${diffsHtml}</td>
          <td>
            <div style="font-weight: 600; color: #334155; font-size: 11px; line-height: 1.3;">${escapeHtml(log.note || '-')}</div>
            <div style="font-size: 10px; color: #64748B; margin-top: 3px; max-width: 150px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${escapeHtml(log.source || '')}">
              📁 ${escapeHtml(log.source || '-')}
            </div>
          </td>
          <td style="text-align: center;">
            <span class="kpi-tag" style="background: #DCFCE7; color: #15803D; font-weight: 800; font-size: 10.5px; border: 1px solid #BBF7D0;">
              ✓ ${escapeHtml(log.status || 'Tersimpan')}
            </span>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    const infoEl = document.getElementById('prog-log-pagination-info');
    if (infoEl) {
      if (filteredTotal === 0) {
        infoEl.textContent = 'Menampilkan 0 log riwayat';
      } else {
        const from = startIdx + 1;
        const to = Math.min(startIdx + PROG_LOG_PAGE_SIZE, filteredTotal);
        infoEl.textContent = `Menampilkan ${from}-${to} dari total ${filteredTotal} log riwayat`;
      }
    }

    const curPageEl = document.getElementById('prog-log-current-page');
    if (curPageEl) curPageEl.textContent = `${currentProgLogPage} / ${totalPages}`;

    const prevBtn = document.getElementById('btn-prog-log-prev');
    if (prevBtn) prevBtn.disabled = (currentProgLogPage <= 1);

    const nextBtn = document.getElementById('btn-prog-log-next');
    if (nextBtn) nextBtn.disabled = (currentProgLogPage >= totalPages);
  }

  function exportPrognosaLogsCSV() {
    const logs = (appData.prognosaLogs && appData.prognosaLogs.length > 0) ? appData.prognosaLogs : DEFAULT_PROGNOSA_LOGS;
    if (!logs || logs.length === 0) {
      alert('Tidak ada data log riwayat untuk diekspor.');
      return;
    }

    const headers = ['No', 'ID Log', 'Waktu & Tanggal (WIB)', 'Pengguna / PIC', 'Jenis Aksi', 'Sumber / File', 'Rincian Perubahan (Diffs)', 'Keterangan', 'Status'];
    const rows = logs.map((log, idx) => {
      const diffText = (log.diffs && log.diffs.length > 0)
        ? log.diffs.map(d => `${d.param}: ${d.before} -> ${d.after} (${d.change || ''})`).join('; ')
        : 'Tidak ada perubahan nilai';
      return [
        idx + 1,
        `"${(log.id || '').replace(/"/g, '""')}"`,
        `"${(log.timestamp || '').replace(/"/g, '""')}"`,
        `"${(log.user || '').replace(/"/g, '""')}"`,
        `"${(log.action || '').replace(/"/g, '""')}"`,
        `"${(log.source || '').replace(/"/g, '""')}"`,
        `"${diffText.replace(/"/g, '""')}"`,
        `"${(log.note || '').replace(/"/g, '""')}"`,
        `"${(log.status || 'Tersimpan').replace(/"/g, '""')}"`
      ].join(',');
    });

    const csvContent = '\uFEFF' + [headers.join(','), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Log_Perubahan_Parameter_Prognosa_${Date.now()}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Log audit berhasil diekspor ke file CSV!');
  }

  function clearPrognosaLogs() {
    if (confirm('Apakah Anda yakin ingin membersihkan riwayat log perubahan parameter Prognosa? Log akan di-reset ke data baseline penajaman.')) {
      appData.prognosaLogs = JSON.parse(JSON.stringify(DEFAULT_PROGNOSA_LOGS));
      saveAppData(appData);
      renderPrognosaLogTable();
      showToast('Riwayat log berhasil dibersihkan.');
    }
  }

  // ==========================================
  // MONITORING PA OPEN VIEW & HANDLERS
  // ==========================================
  let monPaCurrentPage = 1;
  const monPaPageSize = 10;
  // DATA DARI KERTAS KERJA PA OPEN 1.XLSX (Sheet: MonitoringPotentialRevenueActiv)
  const MON_PA_KERTAS_KERJA_DATA = {
    "top10Tahun":  [
                       {
                           "idPa":  "A112104000042",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "DISTRIBUSI/TRANSMISI/PERENCANAAN-PLN",
                           "nilai":  119.41,
                           "nilaiFmt":  "Rp 119,41 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "Zalfaa Azhaar Octaviana",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "31/12/2026",
                           "badgeText":  "92 hari"
                       },
                       {
                           "idPa":  "A112105001253",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "KORPORASI - PLN",
                           "nilai":  60.72,
                           "nilaiFmt":  "Rp 60,72 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN II",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "FEBRI NASIKUL HIMAM",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "31/12/2026",
                           "badgeText":  "92 hari"
                       },
                       {
                           "idPa":  "A112402002492",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "ONE TIME SERVICE",
                           "nilai":  23.14,
                           "nilaiFmt":  "Rp 23,14 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN II",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "Raya Maulana Asdam",
                           "ptl":  "PUTRI NABILAH SYIFA",
                           "targetAktivasi":  "31/12/2026",
                           "badgeText":  "92 hari"
                       },
                       {
                           "idPa":  "A111501000226",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "COLOCATION DATA CENTER",
                           "nilai":  21.64,
                           "nilaiFmt":  "Rp 21,64 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN II",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "AGUS PURNAMA",
                           "ptl":  "ERMAWAN WICAKSONO",
                           "targetAktivasi":  "01/10/2026",
                           "badgeText":  "1 hari"
                       },
                       {
                           "idPa":  "A772105000008",
                           "pelanggan":  "PT. PLN INDONESIA POWER",
                           "layanan":  "KORPORASI - PLN",
                           "nilai":  16.9,
                           "nilaiFmt":  "Rp 16,90 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "NOVAL ALFANSYAH",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A112104000056",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "DISTRIBUSI/TRANSMISI/PERENCANAAN-PLN",
                           "nilai":  16.6,
                           "nilaiFmt":  "Rp 16,60 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "Zalfaa Azhaar Octaviana",
                           "ptl":  "OMAN ROHMAN",
                           "targetAktivasi":  "31/12/2026",
                           "badgeText":  "92 hari"
                       },
                       {
                           "idPa":  "A772105000006",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "KORPORASI - PLN",
                           "nilai":  14.16,
                           "nilaiFmt":  "Rp 14,16 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN II",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "IYGO PRATAMA PUTRA",
                           "ptl":  "RISWULANDARI",
                           "targetAktivasi":  "31/12/2026",
                           "badgeText":  "92 hari"
                       },
                       {
                           "idPa":  "A142204001256",
                           "pelanggan":  "PT. PLN NUSANTARA POWER",
                           "layanan":  "PRODUK DIGITAL",
                           "nilai":  13.6,
                           "nilaiFmt":  "Rp 13,60 M",
                           "bidang":  "JAWA BAGIAN TIMUR",
                           "sbu":  "JAWA BAGIAN TIMUR",
                           "sales":  "JOHN MARI ERVIAN SAPUTRA",
                           "ptl":  "RISKA MAULIDA RAHAYU",
                           "targetAktivasi":  "09/11/2026",
                           "badgeText":  "40 hari"
                       },
                       {
                           "idPa":  "A135101000120",
                           "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK DISTRIBUSI JATENG \u0026 D.I.YOGYAKARTA",
                           "layanan":  "PAKET",
                           "nilai":  12.42,
                           "nilaiFmt":  "Rp 12,42 M",
                           "bidang":  "JAWA BAGIAN TENGAH",
                           "sbu":  "JAWA BAGIAN TENGAH",
                           "sales":  "HERO NURSETO MARANDITA",
                           "ptl":  "KEVIN DAVI SAMUEL",
                           "targetAktivasi":  "22/05/2026",
                           "badgeText":  "131 hari overdue"
                       },
                       {
                           "idPa":  "A112105001224",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "KORPORASI - PLN",
                           "nilai":  11.28,
                           "nilaiFmt":  "Rp 11,28 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN II",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "FEBRI NASIKUL HIMAM",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "31/12/2026",
                           "badgeText":  "92 hari"
                       }
                   ],
    "top10Overdue":  [
                         {
                             "idPa":  "A135101000120",
                             "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK DISTRIBUSI JATENG \u0026 D.I.YOGYAKARTA",
                             "layanan":  "PAKET",
                             "nilai":  12.42,
                             "nilaiFmt":  "Rp 12,42 M",
                             "bidang":  "JAWA BAGIAN TENGAH",
                             "sbu":  "JAWA BAGIAN TENGAH",
                             "sales":  "HERO NURSETO MARANDITA",
                             "ptl":  "KEVIN DAVI SAMUEL",
                             "targetAktivasi":  "22/05/2026",
                             "badgeText":  "131 hari"
                         },
                         {
                             "idPa":  "A112103000067",
                             "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                             "layanan":  "PEMBANGKITAN-PLN",
                             "nilai":  7.98,
                             "nilaiFmt":  "Rp 7,98 M",
                             "bidang":  "MANAJEMEN PELAYANAN PLN I",
                             "sbu":  "JAKARTA \u0026 BANTEN",
                             "sales":  "Muhammad Aby Wicaksono",
                             "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                             "targetAktivasi":  "31/08/2026",
                             "badgeText":  "30 hari"
                         },
                         {
                             "idPa":  "A412402000681",
                             "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK WILAYAH PAPUA DAN PAPUA BARAT",
                             "layanan":  "ONE TIME SERVICE",
                             "nilai":  5.04,
                             "nilaiFmt":  "Rp 5,04 M",
                             "bidang":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sbu":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sales":  "EKA DARMA BUDI PRAKASA",
                             "ptl":  "ARGA DIAN UTAMA SULTAN",
                             "targetAktivasi":  "24/12/2025",
                             "badgeText":  "280 hari"
                         },
                         {
                             "idPa":  "A111801001311",
                             "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                             "layanan":  "I-VIP / VIDEO CONFERENCE",
                             "nilai":  4.57,
                             "nilaiFmt":  "Rp 4,57 M",
                             "bidang":  "MANAJEMEN PELAYANAN PLN II",
                             "sbu":  "JAKARTA \u0026 BANTEN",
                             "sales":  "M. GHUFRON ALI HANIF MANURUNG",
                             "ptl":  "ERMAWAN WICAKSONO",
                             "targetAktivasi":  "31/08/2026",
                             "badgeText":  "30 hari"
                         },
                         {
                             "idPa":  "A111601007059",
                             "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                             "layanan":  "MANAGE SERVICE ROUTER ",
                             "nilai":  3.81,
                             "nilaiFmt":  "Rp 3,81 M",
                             "bidang":  "MANAJEMEN PELAYANAN PLN II",
                             "sbu":  "JAKARTA \u0026 BANTEN",
                             "sales":  "Hary Alfajri",
                             "ptl":  "ERMAWAN WICAKSONO",
                             "targetAktivasi":  "31/08/2026",
                             "badgeText":  "30 hari"
                         },
                         {
                             "idPa":  "A412402001022",
                             "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK DISTRIBUSI SULAWESI SELATAN SULAWESI TENGGARA DAN SULAWESI BARAT",
                             "layanan":  "ONE TIME SERVICE",
                             "nilai":  3.35,
                             "nilaiFmt":  "Rp 3,35 M",
                             "bidang":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sbu":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sales":  "FIRMAN HARRANG",
                             "ptl":  "FADHIL ILHAM MUHAMMAD",
                             "targetAktivasi":  "24/04/2026",
                             "badgeText":  "159 hari"
                         },
                         {
                             "idPa":  "A412402001023",
                             "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK DISTRIBUSI SULAWESI SELATAN SULAWESI TENGGARA DAN SULAWESI BARAT",
                             "layanan":  "ONE TIME SERVICE",
                             "nilai":  2.51,
                             "nilaiFmt":  "Rp 2,51 M",
                             "bidang":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sbu":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sales":  "FIRMAN HARRANG",
                             "ptl":  "FADHIL ILHAM MUHAMMAD",
                             "targetAktivasi":  "24/04/2026",
                             "badgeText":  "159 hari"
                         },
                         {
                             "idPa":  "A412402000679",
                             "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK WILAYAH PAPUA DAN PAPUA BARAT",
                             "layanan":  "ONE TIME SERVICE",
                             "nilai":  2.41,
                             "nilaiFmt":  "Rp 2,41 M",
                             "bidang":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sbu":  "SULAWESI \u0026 INDONESIA TIMUR",
                             "sales":  "EKA DARMA BUDI PRAKASA",
                             "ptl":  "ARGA DIAN UTAMA SULTAN",
                             "targetAktivasi":  "24/12/2025",
                             "badgeText":  "280 hari"
                         },
                         {
                             "idPa":  "A125101000039",
                             "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK DISTRIBUSI JAWA BARAT",
                             "layanan":  "PAKET",
                             "nilai":  2.18,
                             "nilaiFmt":  "Rp 2,18 M",
                             "bidang":  "JAWA BAGIAN BARAT",
                             "sbu":  "JAWA BAGIAN BARAT",
                             "sales":  "ISTI DARMA ASTUTI",
                             "ptl":  "RIZKI NOVITRI SUSANTI SETIA PUTRI",
                             "targetAktivasi":  "08/09/2026",
                             "badgeText":  "22 hari"
                         },
                         {
                             "idPa":  "A112104000047",
                             "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                             "layanan":  "DISTRIBUSI/TRANSMISI/PERENCANAAN-PLN",
                             "nilai":  1.74,
                             "nilaiFmt":  "Rp 1,74 M",
                             "bidang":  "MANAJEMEN PELAYANAN PLN I",
                             "sbu":  "JAKARTA \u0026 BANTEN",
                             "sales":  "Muhammad Aby Wicaksono",
                             "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                             "targetAktivasi":  "01/09/2026",
                             "badgeText":  "29 hari"
                         }
                     ],
    "top10Bulan":  [
                       {
                           "idPa":  "A772105000008",
                           "pelanggan":  "PT. PLN INDONESIA POWER",
                           "layanan":  "KORPORASI - PLN",
                           "nilai":  16.9,
                           "nilaiFmt":  "Rp 16,90 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "NOVAL ALFANSYAH",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A142402003116",
                           "pelanggan":  "PT PLN (PERSERO) UNIT PELAKSANA PELAYANAN PELANGGAN (UP3) MADURA",
                           "layanan":  "ONE TIME SERVICE",
                           "nilai":  6.96,
                           "nilaiFmt":  "Rp 6,96 M",
                           "bidang":  "JAWA BAGIAN TIMUR",
                           "sbu":  "JAWA BAGIAN TIMUR",
                           "sales":  "ALYA FADHIYAN",
                           "ptl":  "BAGAS BRAWIJAYA",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A125101000039",
                           "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK DISTRIBUSI JAWA BARAT",
                           "layanan":  "PAKET",
                           "nilai":  2.18,
                           "nilaiFmt":  "Rp 2,18 M",
                           "bidang":  "JAWA BAGIAN BARAT",
                           "sbu":  "JAWA BAGIAN BARAT",
                           "sales":  "ISTI DARMA ASTUTI",
                           "ptl":  "RIZKI NOVITRI SUSANTI SETIA PUTRI",
                           "targetAktivasi":  "08/09/2026",
                           "badgeText":  "22 hari overdue"
                       },
                       {
                           "idPa":  "A112104000047",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "DISTRIBUSI/TRANSMISI/PERENCANAAN-PLN",
                           "nilai":  1.74,
                           "nilaiFmt":  "Rp 1,74 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "Muhammad Aby Wicaksono",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "01/09/2026",
                           "badgeText":  "29 hari overdue"
                       },
                       {
                           "idPa":  "A412402001385",
                           "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK WILAYAH PAPUA DAN PAPUA BARAT",
                           "layanan":  "ONE TIME SERVICE",
                           "nilai":  1.52,
                           "nilaiFmt":  "Rp 1,52 M",
                           "bidang":  "SULAWESI \u0026 INDONESIA TIMUR",
                           "sbu":  "SULAWESI \u0026 INDONESIA TIMUR",
                           "sales":  "JURINE JESIKA RUNTUKAHU",
                           "ptl":  "ARGA DIAN UTAMA SULTAN",
                           "targetAktivasi":  "21/09/2026",
                           "badgeText":  "9 hari overdue"
                       },
                       {
                           "idPa":  "A212402001093",
                           "pelanggan":  "PT. PLN (PERSERO) UNIT INDUK WILAYAH BANGKA BELITUNG",
                           "layanan":  "ONE TIME SERVICE",
                           "nilai":  1.44,
                           "nilaiFmt":  "Rp 1,44 M",
                           "bidang":  "SUMATERA BAGIAN SELATAN",
                           "sbu":  "SUMATERA BAGIAN SELATAN",
                           "sales":  "RICKY JUVIANTO",
                           "ptl":  "SUPANDI",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A832204000840",
                           "pelanggan":  "PT INDONESIA POWER UJP BANTEN 3 LONTAR (PLTU LONTAR)",
                           "layanan":  "PRODUK DIGITAL",
                           "nilai":  1.3,
                           "nilaiFmt":  "Rp 1,30 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "DEWANTI NUR SARI",
                           "ptl":  "MARLINA CHRISTIAJI",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A112105001177",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "KORPORASI - PLN",
                           "nilai":  0.93,
                           "nilaiFmt":  "Rp 0,93 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN II",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "IYGO PRATAMA PUTRA",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A112104000069",
                           "pelanggan":  "PT. PLN (PERSERO) KANTOR PUSAT",
                           "layanan":  "DISTRIBUSI/TRANSMISI/PERENCANAAN-PLN",
                           "nilai":  0.66,
                           "nilaiFmt":  "Rp 0,66 M",
                           "bidang":  "MANAJEMEN PELAYANAN PLN I",
                           "sbu":  "JAKARTA \u0026 BANTEN",
                           "sales":  "Zalfaa Azhaar Octaviana",
                           "ptl":  "Tim Delivery JAKARTA \u0026 BANTEN",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       },
                       {
                           "idPa":  "A142402003189",
                           "pelanggan":  "PT. PLN (PERSERO) UNIT PELAKSANA PENGATUR DISTRIBUSI JAWA TIMUR",
                           "layanan":  "ONE TIME SERVICE",
                           "nilai":  0.61,
                           "nilaiFmt":  "Rp 0,61 M",
                           "bidang":  "JAWA BAGIAN TIMUR",
                           "sbu":  "JAWA BAGIAN TIMUR",
                           "sales":  "ALYA FADHIYAN",
                           "ptl":  "RISKA MAULIDA RAHAYU",
                           "targetAktivasi":  "30/09/2026",
                           "badgeText":  "0 hari"
                       }
                   ]
}
;

  const MON_PA_SUMMARY_METRICS = {
    totalPa: 7958,
    potensiRev: 801.81,
    buckets: {
      overdue: { pa: 4814, rev: 75.94 },
      d7: { pa: 377, rev: 58.84 },
      d30: { pa: 263, rev: 5.24 },
      d60: { pa: 58, rev: 21.83 },
      d90: { pa: 258, rev: 29.71 },
      far: { pa: 2188, rev: 610.25 }
    },
    monthPerf: {
      targetPa: 4230,
      realPa: 3141,
      achPa: 74.2,
      devPa: 1089,
      targetRev: 425.00,
      realRev: 326.00,
      achRev: 76.7,
      devRev: 99.00
    },
    yearPerf: {
      targetPa: 7958,
      realPa: 3141,
      achPa: 39.5,
      devPa: 4817,
      targetRev: 801.81,
      realRev: 326.00,
      achRev: 40.7,
      devRev: 475.81
    }
  };

  function scrollToTable(id) {
    const el = document.getElementById(id);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  function renderMonitoringPaTemplateTable(tbodyId, list, type) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!list || list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 24px; color: #94A3B8;">Tidak ada data pada kategori ini</td></tr>';
      return;
    }

    list.forEach(item => {
      const tr = document.createElement('tr');
      tr.style.transition = 'background 0.15s ease';
      tr.onmouseover = () => tr.style.background = '#F8FAFC';
      tr.onmouseout = () => tr.style.background = 'transparent';

      let badgeHtml = '';
      if (type === 'overdue') {
        badgeHtml = <span style="display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 800; background: #FEE2E2; color: #DC2626;"></span>;
      } else if (type === 'bulan') {
        badgeHtml = <span style="display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 800; background: #FEF3C7; color: #D97706;"></span>;
      } else {
        const isOvd = item.badgeText && item.badgeText.includes('overdue');
        const bg = isOvd ? '#FEE2E2' : '#DCFCE7';
        const col = isOvd ? '#DC2626' : '#15803D';
        badgeHtml = <span style="display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 800; background: ; color: ;"></span>;
      }

      tr.innerHTML = 
        <td style="font-family: monospace; font-size: 11px; font-weight: 800; color: #4338CA;">
          
        </td>
        <td style="font-weight: 700; color: #0F172A; max-width: 250px; line-height: 1.35;">
          
        </td>
        <td style="color: #334155; max-width: 200px;">
          
        </td>
        <td class="num" style="font-weight: 800; color: #0F172A; white-space: nowrap;">
          
        </td>
        <td style="font-size: 11px; color: #475569; white-space: nowrap;">
          
        </td>
        <td style="font-size: 11px; color: #475569; white-space: nowrap;">
          
        </td>
        <td style="font-weight: 600; color: #1E293B; white-space: nowrap;">
          
        </td>
        <td style="font-size: 11px; color: #64748B; white-space: nowrap;">
          
        </td>
        <td style="text-align: center; font-size: 11px; font-weight: 600; color: #334155; white-space: nowrap;">
          
        </td>
        <td style="text-align: center; white-space: nowrap;">
          
        </td>
      ;
      tbody.appendChild(tr);
    });
  }

  function onMonPaTemplateFilterChange() {
    const sbuVal = (document.getElementById('mon-pa-filter-sbu') || {}).value || 'ALL';
    const bidVal = (document.getElementById('mon-pa-filter-bidang') || {}).value || 'ALL';
    const salesVal = (document.getElementById('mon-pa-filter-sales') || {}).value || 'ALL';

    function filterList(list) {
      if (!list) return [];
      return list.filter(it => {
        const matchSbu = (sbuVal === 'ALL' || it.sbu === sbuVal);
        const matchBid = (bidVal === 'ALL' || it.bidang === bidVal);
        const matchSales = (salesVal === 'ALL' || it.sales === salesVal);
        return matchSbu && matchBid && matchSales;
      });
    }

    renderMonitoringPaTemplateTable('mon-pa-tbody-overdue', filterList(MON_PA_KERTAS_KERJA_DATA.top10Overdue), 'overdue');
    renderMonitoringPaTemplateTable('mon-pa-tbody-bulan', filterList(MON_PA_KERTAS_KERJA_DATA.top10Bulan), 'bulan');
    renderMonitoringPaTemplateTable('mon-pa-tbody-tahun', filterList(MON_PA_KERTAS_KERJA_DATA.top10Tahun), 'tahun');
  }

  function renderMonitoringPaView() {
    // Populate Top 10 Tables
    renderMonitoringPaTemplateTable('mon-pa-tbody-overdue', MON_PA_KERTAS_KERJA_DATA.top10Overdue, 'overdue');
    renderMonitoringPaTemplateTable('mon-pa-tbody-bulan', MON_PA_KERTAS_KERJA_DATA.top10Bulan, 'bulan');
    renderMonitoringPaTemplateTable('mon-pa-tbody-tahun', MON_PA_KERTAS_KERJA_DATA.top10Tahun, 'tahun');

    // Populate Metrics
    const m = MON_PA_SUMMARY_METRICS;
    const elTotPa = document.getElementById('tpl-metric-total-pa');
    if (elTotPa) elTotPa.textContent = Number(m.totalPa).toLocaleString('id-ID');
    const elPotRev = document.getElementById('tpl-metric-potensi-rev');
    if (elPotRev) elPotRev.textContent = 'Rp ' + Number(m.potensiRev).toFixed(2).replace('.', ',');

    const elOvPa = document.getElementById('tpl-card-val-overdue');
    if (elOvPa) elOvPa.textContent = Number(m.buckets.overdue.pa).toLocaleString('id-ID');
    const elOvRev = document.getElementById('tpl-card-rev-overdue');
    if (elOvRev) elOvRev.textContent = 'Rp ' + Number(m.buckets.overdue.rev).toFixed(2).replace('.', ',') + ' M';

    const elD7Pa = document.getElementById('tpl-card-val-d7');
    if (elD7Pa) elD7Pa.textContent = Number(m.buckets.d7.pa).toLocaleString('id-ID');
    const elD7Rev = document.getElementById('tpl-card-rev-d7');
    if (elD7Rev) elD7Rev.textContent = 'Rp ' + Number(m.buckets.d7.rev).toFixed(2).replace('.', ',') + ' M';

    const elD30Pa = document.getElementById('tpl-card-val-d30');
    if (elD30Pa) elD30Pa.textContent = Number(m.buckets.d30.pa).toLocaleString('id-ID');
    const elD30Rev = document.getElementById('tpl-card-rev-d30');
    if (elD30Rev) elD30Rev.textContent = 'Rp ' + Number(m.buckets.d30.rev).toFixed(2).replace('.', ',') + ' M';

    const elD60Pa = document.getElementById('tpl-card-val-d60');
    if (elD60Pa) elD60Pa.textContent = Number(m.buckets.d60.pa).toLocaleString('id-ID');
    const elD60Rev = document.getElementById('tpl-card-rev-d60');
    if (elD60Rev) elD60Rev.textContent = 'Rp ' + Number(m.buckets.d60.rev).toFixed(2).replace('.', ',') + ' M';

    const elD90Pa = document.getElementById('tpl-card-val-d90');
    if (elD90Pa) elD90Pa.textContent = Number(m.buckets.d90.pa).toLocaleString('id-ID');
    const elD90Rev = document.getElementById('tpl-card-rev-d90');
    if (elD90Rev) elD90Rev.textContent = 'Rp ' + Number(m.buckets.d90.rev).toFixed(2).replace('.', ',') + ' M';

    const elFarPa = document.getElementById('tpl-card-val-far');
    if (elFarPa) elFarPa.textContent = Number(m.buckets.far.pa).toLocaleString('id-ID');
    const elFarRev = document.getElementById('tpl-card-rev-far');
    if (elFarRev) elFarRev.textContent = 'Rp ' + Number(m.buckets.far.rev).toFixed(2).replace('.', ',') + ' M';
  }

  // Master Prognosa View & Handlers
  function renderMasterPrognosaView() {
    const p = appData.prognosa || DEFAULT_PROGNOSA;

    const picInput = document.getElementById('m_prog_user_pic');
    if (picInput && !picInput.value) picInput.value = 'dodik.gunawan';

    // 7 Kartu KPI Parameters (Triliun Rp)
    const tgtKum = (p.targetKumulatif !== undefined) ? p.targetKumulatif : DEFAULT_PROGNOSA.targetKumulatif;
    const tgtSa = (p.targetStandAlone !== undefined) ? p.targetStandAlone : DEFAULT_PROGNOSA.targetStandAlone;
    const realisasi = (p.realisasiLalu !== undefined) ? p.realisasiLalu : DEFAULT_PROGNOSA.realisasiLalu;
    const asIs = (p.prognosaAsIs !== undefined) ? p.prognosaAsIs : DEFAULT_PROGNOSA.prognosaAsIs;
    const mostLikely = (p.prognosaMostLikely !== undefined) ? p.prognosaMostLikely : DEFAULT_PROGNOSA.prognosaMostLikely;

    const mTgtKum = document.getElementById('m_prog_target_kum'); if (mTgtKum) mTgtKum.value = tgtKum;
    const mTgtSa = document.getElementById('m_prog_target_sa'); if (mTgtSa) mTgtSa.value = tgtSa;
    const mReal = document.getElementById('m_prog_realisasi'); if (mReal) mReal.value = realisasi;
    const mAsIs = document.getElementById('m_prog_asis'); if (mAsIs) mAsIs.value = asIs;
    const mMl = document.getElementById('m_prog_mostlikely'); if (mMl) mMl.value = mostLikely;

    // Update live previews on Master Prognosa
    masterAutoCalcPrognosaCards();

    // Basis RKAP Parameters (Milyar Rp)
    const targetRkap = (p && p.targetRkap !== undefined) ? p.targetRkap : DEFAULT_PROGNOSA.targetRkap;
    const targetBulan = (p && p.targetBulanan !== undefined) ? p.targetBulanan : DEFAULT_PROGNOSA.targetBulanan;
    const co = (p && p.co !== undefined) ? p.co : DEFAULT_PROGNOSA.co;
    const nr = (p && p.nr !== undefined) ? p.nr : DEFAULT_PROGNOSA.nr;
    const nrMonth = (p && p.nrBulanBerjalan !== undefined) ? p.nrBulanBerjalan : DEFAULT_PROGNOSA.nrBulanBerjalan;
    const est = (p && p.estRevenue !== undefined) ? p.estRevenue : DEFAULT_PROGNOSA.estRevenue;

    const mTgt = document.getElementById('m_prog_target_rkap'); if (mTgt) mTgt.value = targetRkap;
    const mTgtBln = document.getElementById('m_prog_target_bulan'); if (mTgtBln) mTgtBln.value = targetBulan;
    const mCo = document.getElementById('m_prog_co'); if (mCo) mCo.value = co;
    const mNr = document.getElementById('m_prog_nr'); if (mNr) mNr.value = nr;
    const mNrM = document.getElementById('m_prog_nr_bulan'); if (mNrM) mNrM.value = nrMonth;
    const mEst = document.getElementById('m_prog_est'); if (mEst) mEst.value = est;

    const srcBadge = document.getElementById('master-prog-source-badge');
    if (srcBadge && p.sourceFile) srcBadge.textContent = p.sourceFile;

    renderMasterPrognosaTable(p.services || DEFAULT_PROGNOSA.services, est);
    renderPrognosaLogTable();
  }

  function masterAutoCalcPrognosaCards() {
    const tgtKum = parseFloat(document.getElementById('m_prog_target_kum') ? document.getElementById('m_prog_target_kum').value : 2.74) || 0;
    const asIs = parseFloat(document.getElementById('m_prog_asis') ? document.getElementById('m_prog_asis').value : 2.60) || 0;
    
    // Gap = AS IS - Target
    const gap = asIs - tgtKum;
    const gapPct = tgtKum > 0 ? ((gap / tgtKum) * 100).toFixed(1) : '0,0';
    const gapEl = document.getElementById('m_prog_preview_gap_val');
    const gapPctEl = document.getElementById('m_prog_preview_gap_pct');
    if (gapEl) {
      gapEl.textContent = (gap >= 0 ? '+' : '') + gap.toFixed(2).replace('.', ',') + ' T';
      gapEl.style.color = gap >= 0 ? '#047857' : '#DC2626';
    }
    if (gapPctEl) {
      gapPctEl.textContent = (gap >= 0 ? '+' : '') + gapPct.replace('.', ',') + '%';
      gapPctEl.style.background = gap >= 0 ? '#DCFCE7' : '#FEE2E2';
      gapPctEl.style.color = gap >= 0 ? '#047857' : '#DC2626';
    }

    // Achievement = (AS IS / Target) * 100
    const achievePct = tgtKum > 0 ? ((asIs / tgtKum) * 100).toFixed(1) : '0,0';
    const achEl = document.getElementById('m_prog_preview_achieve_val');
    const achSubEl = document.getElementById('m_prog_preview_achieve_sub');
    if (achEl) {
      achEl.textContent = achievePct.replace('.', ',') + '%';
      achEl.style.color = parseFloat(achievePct) >= 100 ? '#047857' : '#4338CA';
    }
    if (achSubEl) {
      achSubEl.textContent = `Target ${tgtKum.toFixed(2).replace('.', ',')} · Prog ${asIs.toFixed(2).replace('.', ',')}`;
    }
  }

  function masterSyncCardsWithRkap() {
    const current = appData.prognosa || DEFAULT_PROGNOSA;
    const targetRkap = parseFloat(document.getElementById('m_prog_target_rkap') ? document.getElementById('m_prog_target_rkap').value : 4156) || 4156.00;
    const targetBulan = parseFloat(document.getElementById('m_prog_target_bulan') ? document.getElementById('m_prog_target_bulan').value : 415.59) || 415.59;
    
    // Standar rasio penajaman kumulatif s.d Sep: 2.742,95 / 4.156 = ~66%
    const estKumulatif = Number(((targetRkap * 0.66) / 1000).toFixed(2));
    const estStandAlone = Number((targetBulan / 1000).toFixed(2));

    const oldKum = current.targetKumulatif || 2.74;
    const oldSa = current.targetStandAlone || 0.25;

    const mTgtKum = document.getElementById('m_prog_target_kum');
    if (mTgtKum) mTgtKum.value = estKumulatif;

    const mTgtSa = document.getElementById('m_prog_target_sa');
    if (mTgtSa) mTgtSa.value = estStandAlone;

    masterAutoCalcPrognosaCards();

    const diffs = [];
    if (Math.abs(oldKum - estKumulatif) > 0.001) {
      const d = estKumulatif - oldKum;
      diffs.push({
        param: 'Card 1 (Target Kumulatif)',
        before: oldKum.toFixed(2).replace('.', ',') + ' T',
        after: estKumulatif.toFixed(2).replace('.', ',') + ' T',
        change: (d >= 0 ? '+' : '') + d.toFixed(2).replace('.', ',') + ' T'
      });
    }
    if (Math.abs(oldSa - estStandAlone) > 0.001) {
      const d = estStandAlone - oldSa;
      diffs.push({
        param: 'Card 2 (Target Stand Alone)',
        before: oldSa.toFixed(2).replace('.', ',') + ' T',
        after: estStandAlone.toFixed(2).replace('.', ',') + ' T',
        change: (d >= 0 ? '+' : '') + d.toFixed(2).replace('.', ',') + ' T'
      });
    }

    if (diffs.length > 0) {
      addPrognosaAuditLog({
        user: getCurrentPrognosaUser(),
        action: 'Selaraskan dgn RKAP',
        source: 'Form Master Prognosa',
        diffs: diffs,
        note: `Penyelarasan rasio target kumulatif (~66%) & bulanan dari RKAP ${targetRkap.toFixed(2).replace('.', ',')} M`
      });
    }

    showToast('Nilai kartu diselaraskan dengan Target RKAP & Bulanan!');
  }

  function masterAutoCalcEst() {
    const co = parseFloat(document.getElementById('m_prog_co').value) || 0;
    const nr = parseFloat(document.getElementById('m_prog_nr').value) || 0;
    document.getElementById('m_prog_est').value = (co + nr).toFixed(2);
  }

  function saveMasterPrognosaData() {
    const current = appData.prognosa || DEFAULT_PROGNOSA;

    // 7 KPI Cards
    const targetKum = parseFloat(document.getElementById('m_prog_target_kum') ? document.getElementById('m_prog_target_kum').value : 2.74) || 2.74;
    const targetSa = parseFloat(document.getElementById('m_prog_target_sa') ? document.getElementById('m_prog_target_sa').value : 0.25) || 0.25;
    const realisasi = parseFloat(document.getElementById('m_prog_realisasi') ? document.getElementById('m_prog_realisasi').value : 2.23) || 2.23;
    const asIs = parseFloat(document.getElementById('m_prog_asis') ? document.getElementById('m_prog_asis').value : 2.60) || 2.60;
    const mostLikely = parseFloat(document.getElementById('m_prog_mostlikely') ? document.getElementById('m_prog_mostlikely').value : 3.42) || 3.42;

    // Basis RKAP Parameters
    const targetRkap = parseFloat(document.getElementById('m_prog_target_rkap').value) || 4156.00;
    const targetBulan = parseFloat(document.getElementById('m_prog_target_bulan').value) || 415.59;
    const co = parseFloat(document.getElementById('m_prog_co').value) || 2618.09;
    const nr = parseFloat(document.getElementById('m_prog_nr').value) || 1537.91;
    const nrMonth = parseFloat(document.getElementById('m_prog_nr_bulan').value) || 8.47;
    const est = parseFloat(document.getElementById('m_prog_est').value) || (co + nr);

    // Diff engine
    const diffs = [];
    function checkDiffT(paramName, oldVal, newVal) {
      const o = (oldVal !== undefined) ? Number(oldVal) : 0;
      const n = Number(newVal);
      if (Math.abs(o - n) > 0.0001) {
        const delta = n - o;
        diffs.push({
          param: paramName,
          before: o.toFixed(2).replace('.', ',') + ' T',
          after: n.toFixed(2).replace('.', ',') + ' T',
          change: (delta >= 0 ? '+' : '') + delta.toFixed(2).replace('.', ',') + ' T'
        });
      }
    }
    function checkDiffM(paramName, oldVal, newVal) {
      const o = (oldVal !== undefined) ? Number(oldVal) : 0;
      const n = Number(newVal);
      if (Math.abs(o - n) > 0.01) {
        const delta = n - o;
        diffs.push({
          param: paramName,
          before: o.toFixed(2).replace('.', ',') + ' M',
          after: n.toFixed(2).replace('.', ',') + ' M',
          change: (delta >= 0 ? '+' : '') + delta.toFixed(2).replace('.', ',') + ' M'
        });
      }
    }

    checkDiffT('Card 1 (Target Kumulatif)', current.targetKumulatif, targetKum);
    checkDiffT('Card 2 (Target Stand Alone)', current.targetStandAlone, targetSa);
    checkDiffT('Card 3 (Realisasi s.d Ags)', current.realisasiLalu, realisasi);
    checkDiffT('Card 4 (Prognosa AS IS)', current.prognosaAsIs, asIs);
    checkDiffT('Card 5 (Prognosa Most Likely)', current.prognosaMostLikely, mostLikely);

    checkDiffM('Target RKAP Korporat', current.targetRkap, targetRkap);
    checkDiffM('Target Bulanan', current.targetBulanan, targetBulan);
    checkDiffM('Carry Over (CO)', current.co, co);
    checkDiffM('New Revenue (NR)', current.nr, nr);
    checkDiffM('NR Bulan Berjalan', current.nrBulanBerjalan, nrMonth);
    checkDiffM('Est Revenue Akhir Tahun', current.estRevenue, est);

    if (!appData.prognosa) appData.prognosa = JSON.parse(JSON.stringify(DEFAULT_PROGNOSA));
    appData.prognosa.targetKumulatif = targetKum;
    appData.prognosa.targetStandAlone = targetSa;
    appData.prognosa.realisasiLalu = realisasi;
    appData.prognosa.prognosaAsIs = asIs;
    appData.prognosa.prognosaMostLikely = mostLikely;

    appData.prognosa.targetRkap = targetRkap;
    appData.prognosa.targetBulanan = targetBulan;
    appData.prognosa.co = co;
    appData.prognosa.nr = nr;
    appData.prognosa.nrBulanBerjalan = nrMonth;
    appData.prognosa.estRevenue = est;

    const user = getCurrentPrognosaUser();
    if (diffs.length > 0) {
      addPrognosaAuditLog({
        user: user,
        action: 'Manual Edit Form',
        source: 'Form Master Prognosa',
        diffs: diffs,
        note: `${diffs.length} parameter diubah secara manual via form master`
      });
    } else {
      addPrognosaAuditLog({
        user: user,
        action: 'Manual Edit Form',
        source: 'Form Master Prognosa',
        diffs: [],
        note: 'Konfirmasi simpan parameter (tidak ada perubahan angka)'
      });
    }

    saveAppData(appData);
    renderMasterPrognosaView();
    renderPrognosaView();
    showToast('Parameter 7 Kartu KPI & Master Prognosa berhasil disimpan! Log audit tercatat.');
  }

  function renderMasterPrognosaTable(services, totalEst) {
    const tbody = document.getElementById('master-prog-table-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    services.forEach((s) => {
      const tr = document.createElement('tr');
      const tgt = s.target || 0;
      const pctCapai = tgt > 0 ? ((s.est / tgt) * 100).toFixed(1) : (totalEst > 0 ? Math.round((s.est / totalEst) * 100) : 0);
      const isOver = parseFloat(pctCapai) >= 100;
      tr.innerHTML = `
        <td style="font-weight: 700; color: var(--navy); display: flex; align-items: center; gap: 8px;">
          <span style="width: 7px; height: 7px; border-radius: 50%; background: ${isOver ? '#10B981' : '#6366F1'};"></span>
          ${s.name}
        </td>
        <td class="num" style="font-weight: 700; color: var(--navy);">${tgt > 0 ? Number(tgt).toFixed(2).replace('.', ',') : '-'}</td>
        <td class="num" style="color: #047857; font-weight: 600;">${Number(s.co).toFixed(2).replace('.', ',')}</td>
        <td class="num" style="color: #6D28D9; font-weight: 600;">${Number(s.nr).toFixed(2).replace('.', ',')}</td>
        <td class="num" style="color: #B45309; font-weight: 600;">${Number(s.nrMonth || 0).toFixed(2).replace('.', ',')}</td>
        <td class="num" style="color: #4338CA; font-weight: 800;">${Number(s.est).toFixed(2).replace('.', ',')}</td>
        <td style="text-align: center;">
          <span class="kpi-tag" style="background: ${isOver ? '#ECFDF5' : 'rgba(79, 70, 229, 0.1)'}; color: ${isOver ? '#047857' : '#4338CA'}; font-weight: 700;">
            ${pctCapai}%
          </span>
        </td>
      `;
      tbody.appendChild(tr);
    });

    let totTgt = 0, totCo = 0, totNr = 0, totMonth = 0, totEstRow = 0;
    services.forEach(s => {
      totTgt += (s.target || 0);
      totCo += s.co;
      totNr += s.nr;
      totMonth += (s.nrMonth || 0);
      totEstRow += s.est;
    });
    const totPct = totTgt > 0 ? ((totEstRow / totTgt) * 100).toFixed(1) : '100';
    const trTot = document.createElement('tr');
    trTot.style.background = '#F8FAFC';
    trTot.style.fontWeight = '800';
    trTot.innerHTML = `
      <td style="font-weight: 800; color: var(--navy); text-transform: uppercase;">Total Seluruh Segmen / SBU</td>
      <td class="num" style="color: var(--navy); font-size: 13px;">${totTgt > 0 ? totTgt.toFixed(2).replace('.', ',') : '-'} M</td>
      <td class="num" style="color: #047857; font-size: 13px;">${totCo.toFixed(2).replace('.', ',')} M</td>
      <td class="num" style="color: #6D28D9; font-size: 13px;">${totNr.toFixed(2).replace('.', ',')} M</td>
      <td class="num" style="color: #B45309; font-size: 13px;">${totMonth.toFixed(2).replace('.', ',')} M</td>
      <td class="num" style="color: #4338CA; font-size: 14px;">${totEstRow.toFixed(2).replace('.', ',')} M</td>
      <td style="text-align: center;"><span class="kpi-tag" style="background: #4338CA; color: #fff; font-weight: 800;">${totPct}%</span></td>
    `;
    tbody.appendChild(trTot);
  }

  function openPrognosaModal() {
    const p = appData.prognosa || DEFAULT_PROGNOSA;
    document.getElementById('in_prog_target_rkap').value = (p.targetRkap || 4156.00);
    document.getElementById('in_prog_target_bulan').value = p.targetBulanan || 415.59;
    document.getElementById('in_prog_co').value = p.co || 2618.09;
    document.getElementById('in_prog_nr').value = p.nr || 1537.91;
    document.getElementById('in_prog_nr_bulan').value = p.nrBulanBerjalan || 8.47;
    document.getElementById('in_prog_est').value = p.estRevenue || 4034.02;
    document.getElementById('modalPrognosa').classList.add('open');
  }

  function closePrognosaModal() {
    document.getElementById('modalPrognosa').classList.remove('open');
  }

  function autoCalcEst() {
    const co = parseFloat(document.getElementById('in_prog_co').value) || 0;
    const nr = parseFloat(document.getElementById('in_prog_nr').value) || 0;
    document.getElementById('in_prog_est').value = (co + nr).toFixed(2);
  }

  function savePrognosaData() {
    const current = appData.prognosa || DEFAULT_PROGNOSA;
    const targetRkap = parseFloat(document.getElementById('in_prog_target_rkap').value) || 4156.00;
    const targetBulan = parseFloat(document.getElementById('in_prog_target_bulan').value) || 415.59;
    const co = parseFloat(document.getElementById('in_prog_co').value) || 2618.09;
    const nr = parseFloat(document.getElementById('in_prog_nr').value) || 1537.91;
    const nrMonth = parseFloat(document.getElementById('in_prog_nr_bulan').value) || 8.47;
    const est = parseFloat(document.getElementById('in_prog_est').value) || (co + nr);

    const diffs = [];
    function checkDiffM(paramName, oldVal, newVal) {
      const o = (oldVal !== undefined) ? Number(oldVal) : 0;
      const n = Number(newVal);
      if (Math.abs(o - n) > 0.01) {
        const delta = n - o;
        diffs.push({
          param: paramName,
          before: o.toFixed(2).replace('.', ',') + ' M',
          after: n.toFixed(2).replace('.', ',') + ' M',
          change: (delta >= 0 ? '+' : '') + delta.toFixed(2).replace('.', ',') + ' M'
        });
      }
    }
    checkDiffM('Target RKAP Korporat', current.targetRkap, targetRkap);
    checkDiffM('Target Bulanan', current.targetBulanan, targetBulan);
    checkDiffM('Carry Over (CO)', current.co, co);
    checkDiffM('New Revenue (NR)', current.nr, nr);
    checkDiffM('NR Bulan Berjalan', current.nrBulanBerjalan, nrMonth);
    checkDiffM('Est Revenue Akhir Tahun', current.estRevenue, est);

    if (!appData.prognosa) appData.prognosa = JSON.parse(JSON.stringify(DEFAULT_PROGNOSA));
    appData.prognosa.targetRkap = targetRkap;
    appData.prognosa.targetBulanan = targetBulan;
    appData.prognosa.co = co;
    appData.prognosa.nr = nr;
    appData.prognosa.nrBulanBerjalan = nrMonth;
    appData.prognosa.estRevenue = est;

    if (diffs.length > 0) {
      addPrognosaAuditLog({
        user: getCurrentPrognosaUser(),
        action: 'Manual Edit Form',
        source: 'Modal Edit Prognosa',
        diffs: diffs,
        note: `${diffs.length} parameter korporat diubah via modal popup`
      });
    }

    saveAppData(appData);
    closePrognosaModal();
    renderPrognosaView();
    renderMasterPrognosaView();
    showToast('Parameter Prognosa Revenue berhasil diperbarui! Log audit tercatat.');
  }

  function resetPrognosaData() {
    if (confirm('Kembalikan nilai Prognosa Revenue ke data Penajaman Strategi Penjualan 2026?')) {
      const user = getCurrentPrognosaUser();
      const current = appData.prognosa || DEFAULT_PROGNOSA;
      
      const diffs = [
        {
          param: 'Seluruh Parameter Prognosa',
          before: `Target ${Number(current.targetKumulatif || 2.74).toFixed(2).replace('.', ',')} T / AS IS ${Number(current.prognosaAsIs || 2.60).toFixed(2).replace('.', ',')} T`,
          after: 'Nilai Standar Penajaman 2026',
          change: 'Reset'
        }
      ];

      addPrognosaAuditLog({
        user: user,
        action: 'Reset ke Default',
        source: 'Penajaman Strategi Penjualan PLN Tahun 2026.xlsx',
        diffs: diffs,
        note: 'Pengembalian seluruh parameter ke baseline standar RKAP Penajaman 2026'
      });

      appData.prognosa = JSON.parse(JSON.stringify(DEFAULT_PROGNOSA));
      saveAppData(appData);
      renderPrognosaView();
      renderMasterPrognosaView();
      showToast('Data Prognosa Revenue telah di-reset ke data Penajaman 2026. Log audit tercatat.');
    }
  }

  function handlePrognosaExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (typeof XLSX === 'undefined') {
      alert('Library SheetJS (XLSX) sedang dimuat. Silakan coba sesaat lagi.');
      return;
    }

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // 7 Kartu KPI Utama Prognosa (Satuan Triliun Rupiah / T)
        let targetKumulatif = 2.74;
        let targetStandAlone = 0.25;
        let realisasiLalu = 2.23;
        let prognosaAsIs = 2.60;
        let prognosaMostLikely = 3.42;

        // Basis RKAP Parameters (Satuan Milyar Rupiah / M)
        let targetRkap = 4156.00;
        let targetBulan = 415.59;
        let co = 2618.09;
        let nr = 1537.91;
        let nrMonth = 8.47;
        let est = 4034.02;
        let services = [];

        // 1. Ekstrak data 7 Kartu KPI dari Sheet CHECK POINT / CHECKPOINT
        const checkSheet = workbook.Sheets['CHECK POINT'] || workbook.Sheets['CHECKPOINT'];
        if (checkSheet) {
          // Target Kumulatif Sep: AO20 (2.742,95 M) or L20 (2.742,96 M) or AN16
          let cTgtKumVal = null;
          if (checkSheet['AO20'] && !isNaN(parseFloat(checkSheet['AO20'].v))) cTgtKumVal = parseFloat(checkSheet['AO20'].v);
          else if (checkSheet['L20'] && !isNaN(parseFloat(checkSheet['L20'].v))) cTgtKumVal = parseFloat(checkSheet['L20'].v);
          else if (checkSheet['AN16'] && !isNaN(parseFloat(checkSheet['AN16'].v))) cTgtKumVal = parseFloat(checkSheet['AN16'].v);
          if (cTgtKumVal !== null && cTgtKumVal > 0) {
            targetKumulatif = Number((cTgtKumVal >= 100 ? (cTgtKumVal / 1000) : cTgtKumVal).toFixed(2));
          }

          // Realisasi s.d Agustus: AO21 (2.213 M) or AL16 (2.118,27 M) or J23 (1.895,99 M)
          let cRealVal = null;
          if (checkSheet['AO21'] && !isNaN(parseFloat(checkSheet['AO21'].v))) cRealVal = parseFloat(checkSheet['AO21'].v);
          else if (checkSheet['AL16'] && !isNaN(parseFloat(checkSheet['AL16'].v))) cRealVal = parseFloat(checkSheet['AL16'].v);
          if (cRealVal !== null && cRealVal > 0) {
            realisasiLalu = Number((cRealVal >= 100 ? (cRealVal / 1000) : cRealVal).toFixed(2));
          }

          // Target Stand Alone: jika tertera di AO24/AO25 atau kalkulasi dari bulanan
          if (checkSheet['AO24'] && !isNaN(parseFloat(checkSheet['AO24'].v))) {
            const coRecurring = parseFloat(checkSheet['AO24'].v) / 1000;
            // Rekonsiliasi AS IS: realisasi + CO recurring + NR
            if (realisasiLalu > 0) {
              prognosaAsIs = Number((realisasiLalu + coRecurring + 0.18 - 0.02).toFixed(2));
            }
          }
        }

        // 2. Ekstrak data Sheet TARGET
        if (workbook.SheetNames.includes('TARGET')) {
          const sheet = workbook.Sheets['TARGET'];
          if (sheet['D16']) targetRkap = parseFloat(sheet['D16'].v) || targetRkap;
          if (sheet['E16']) co = parseFloat(sheet['E16'].v) || co;
          if (sheet['F16']) nr = parseFloat(sheet['F16'].v) || nr;
          if (sheet['O17']) {
            targetBulan = parseFloat(sheet['O17'].v) || targetBulan;
            if (targetBulan < 300) {
              targetStandAlone = Number((targetBulan / 1000).toFixed(2));
            }
          }
        }

        // 3. Ekstrak data Sheet PROGNOSA to SPKK
        if (workbook.SheetNames.includes('PROGNOSA to SPKK')) {
          const spkk = workbook.Sheets['PROGNOSA to SPKK'];
          if (spkk['O12']) nrMonth = parseFloat(spkk['O12'].v) || nrMonth;
        }

        // 4. Ekstrak data Sheet SUMMARY 2026
        if (workbook.SheetNames.includes('SUMMARY 2026')) {
          const sumSheet = workbook.Sheets['SUMMARY 2026'];
          if (sumSheet['R15']) est = parseFloat(sumSheet['R15'].v) || est;

          const sbuRows = [
            { r: 4, name: 'PLN 1 (Holding / Korporasi / DIVSTI)', target: 1183.02 },
            { r: 5, name: 'PLN 2 (Subholding & Anak Perusahaan)', target: 2307.29 },
            { r: 7, name: 'SBU JBB (Jakarta Raya & Banten)', target: 89.00 },
            { r: 9, name: 'SBU JTG (Jawa Tengah & DIY)', target: 83.31 },
            { r: 8, name: 'SBU JBT (Jawa Barat)', target: 75.00 },
            { r: 6, name: 'SBU BNR (Bali & Nusa Tenggara)', target: 45.00 },
            { r: 14, name: 'SBU SBU (Sumatera Bagian Utara)', target: 90.00 },
            { r: 12, name: 'SBU SBT (Sumatera Bagian Tengah)', target: 73.06 },
            { r: 11, name: 'SBU SBS (Sumatera Bagian Selatan)', target: 64.32 },
            { r: 13, name: 'SBU SLW (Sulawesi)', target: 73.00 },
            { r: 10, name: 'SBU KLM (Kalimantan)', target: 73.00 }
          ];

          services = sbuRows.map(rowDef => {
            const rNum = rowDef.r;
            const sCo = sumSheet['D' + rNum] ? parseFloat(sumSheet['D' + rNum].v) : 0;
            const sNr = sumSheet['Q' + rNum] ? parseFloat(sumSheet['Q' + rNum].v) : 0;
            const sNrM = sumSheet['M' + rNum] ? parseFloat(sumSheet['M' + rNum].v) : 0;
            const sEst = sumSheet['R' + rNum] ? parseFloat(sumSheet['R' + rNum].v) : (sCo + sNr);
            return {
              name: rowDef.name,
              target: rowDef.target,
              co: Number(sCo.toFixed(2)),
              nr: Number(sNr.toFixed(2)),
              nrMonth: Number(sNrM.toFixed(2)),
              est: Number(sEst.toFixed(2))
            };
          });
        }

        const current = appData.prognosa || DEFAULT_PROGNOSA;
        const diffs = [];
        function checkDiffT(paramName, oldVal, newVal) {
          const o = (oldVal !== undefined) ? Number(oldVal) : 0;
          const n = Number(newVal);
          if (Math.abs(o - n) > 0.0001) {
            const delta = n - o;
            diffs.push({
              param: paramName,
              before: o.toFixed(2).replace('.', ',') + ' T',
              after: n.toFixed(2).replace('.', ',') + ' T',
              change: (delta >= 0 ? '+' : '') + delta.toFixed(2).replace('.', ',') + ' T'
            });
          }
        }
        function checkDiffM(paramName, oldVal, newVal) {
          const o = (oldVal !== undefined) ? Number(oldVal) : 0;
          const n = Number(newVal);
          if (Math.abs(o - n) > 0.01) {
            const delta = n - o;
            diffs.push({
              param: paramName,
              before: o.toFixed(2).replace('.', ',') + ' M',
              after: n.toFixed(2).replace('.', ',') + ' M',
              change: (delta >= 0 ? '+' : '') + delta.toFixed(2).replace('.', ',') + ' M'
            });
          }
        }

        checkDiffT('Card 1 (Target Kumulatif)', current.targetKumulatif, targetKumulatif);
        checkDiffT('Card 2 (Target Stand Alone)', current.targetStandAlone, targetStandAlone);
        checkDiffT('Card 3 (Realisasi s.d Ags)', current.realisasiLalu, realisasiLalu);
        checkDiffT('Card 4 (Prognosa AS IS)', current.prognosaAsIs, prognosaAsIs);
        checkDiffT('Card 5 (Prognosa Most Likely)', current.prognosaMostLikely, prognosaMostLikely);
        checkDiffM('Target RKAP', current.targetRkap, targetRkap);
        checkDiffM('Target Bulanan', current.targetBulanan, targetBulan);
        checkDiffM('Carry Over (CO)', current.co, co);
        checkDiffM('New Revenue (NR)', current.nr, nr);
        checkDiffM('NR Bulan Berjalan', current.nrBulanBerjalan, nrMonth);
        checkDiffM('Est Revenue', current.estRevenue, est);

        if (!appData.prognosa) appData.prognosa = {};
        appData.prognosa.targetKumulatif = targetKumulatif;
        appData.prognosa.targetStandAlone = targetStandAlone;
        appData.prognosa.realisasiLalu = realisasiLalu;
        appData.prognosa.prognosaAsIs = prognosaAsIs;
        appData.prognosa.prognosaMostLikely = prognosaMostLikely;

        appData.prognosa.targetRkap = targetRkap;
        appData.prognosa.targetBulanan = targetBulan;
        appData.prognosa.co = co;
        appData.prognosa.nr = nr;
        appData.prognosa.nrBulanBerjalan = nrMonth;
        appData.prognosa.estRevenue = est;
        if (services.length > 0) appData.prognosa.services = services;
        appData.prognosa.sourceFile = file.name;

        addPrognosaAuditLog({
          user: getCurrentPrognosaUser(),
          action: 'Upload File Excel',
          source: file.name,
          diffs: diffs.length > 0 ? diffs : [
            { param: 'File Excel Diunggah', before: current.sourceFile || '-', after: file.name, change: 'Update' }
          ],
          note: `Impor data dari spreadsheet Excel '${file.name}' (${services.length > 0 ? services.length + ' SBU' : '7 Kartu & RKAP'})`
        });

        saveAppData(appData);
        renderPrognosaView();
        renderMasterPrognosaView();
        showToast('File Excel ' + file.name + ' berhasil diimpor! Parameter Prognosa diperbarui dan log audit tercatat.');
      } catch (err) {
        console.error(err);
        alert('Gagal membaca file Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  // Render Month Chips (Jan - Des)
  function renderMonthChips() {
    const wrap = document.getElementById('month-chips-wrap');
    if (!wrap) return;
    wrap.innerHTML = '<span class="period-label">PERIODE</span>';
    appData.months.forEach((m, idx) => {
      const btn = document.createElement('button');
      btn.className = 'month-chip' + (idx === selectedMonthIndex ? ' active' : '');
      btn.textContent = m.short;
      btn.onclick = () => {
        selectedMonthIndex = idx;
        renderDashboard();
      };
      wrap.appendChild(btn);
    });
  }

  // =========================================================================
  // RENDER DASHBOARD: 1 TAHUN REVENUE & OPPORTUNITY + BREAKDOWN BULANAN
  // =========================================================================
  function renderDashboard() {
    renderMonthChips();

    // 1. Calculate / Read Full Year Totals directly from Rekap & Detil Angsuran
    const rekap = appData.rekap || DEFAULT_REKAP || [];
    const detil = appData.detilAngsuran || DEFAULT_DETIL || [];

    // Sum Full, Tahap, & Nilai WBS from Rekap (in Juta)
    let sumFullJt = 0;
    let sumTahapJt = 0;
    let sumWbsJt = 0;
    rekap.forEach(r => {
      sumFullJt += Math.round((Number(r.NILAI_FULL) || 0) / 1e6);
      sumTahapJt += Math.round((Number(r.NILAI_TAHAP) || 0) / 1e6);
      sumWbsJt += Math.round((Number(r.NILAI_WBS) || 0) / 1e6);
    });

    // Sum Cash In & Belum Cair from Detil Angsuran (in Juta)
    let sumCashJt = 0;
    let sumWaitJt = 0;
    detil.forEach(d => {
      const valJt = Math.round((Number(d.NILAI_TAHAP) || Number(d.NILAI_ANGSURAN) || 0) / 1e6);
      if ((d.STATUS_LUNAS || '').toUpperCase() === 'LUNAS') {
        sumCashJt += valJt;
      } else {
        sumWaitJt += valJt;
      }
    });

    const yearTargetJt = Math.round((appData.annual && appData.annual.target ? appData.annual.target : 186.00) * 1000);
    const yearOppJt = sumFullJt;
    const yearCashJt = sumCashJt;
    const yearWaitJt = sumWaitJt;
    const yearRealJt = (yearCashJt + yearWaitJt) > 0 ? (yearCashJt + yearWaitJt) : sumTahapJt;
    const yearMostLikelyJt = (appData.annual && appData.annual.mostLikely !== undefined)
      ? Math.round(appData.annual.mostLikely * 1000)
      : (sumWbsJt > 0 ? sumWbsJt : Math.round(yearOppJt * 0.85));

    // Keep appData.annual in sync
    if (!appData.annual) appData.annual = {};
    appData.annual.opp = parseFloat((yearOppJt / 1000).toFixed(2));
    appData.annual.real = parseFloat((yearRealJt / 1000).toFixed(2));
    appData.annual.mostLikely = parseFloat((yearMostLikelyJt / 1000).toFixed(2));
    appData.annual.cash = parseFloat((yearCashJt / 1000).toFixed(2));
    appData.annual.wait = parseFloat((yearWaitJt / 1000).toFixed(2));
    appData.annual.risk = 0.00;

    // Update Top 6 KPI Cards with 1 YEAR Totals
    const elTargetLbl = document.getElementById('kpi-target-label');
    if (elTargetLbl) elTargetLbl.textContent = 'TARGET 1 TAHUN (' + (appData.year || '2026') + ')';
    const elTargetVal = document.getElementById('kpi-target-val');
    if (elTargetVal) elTargetVal.textContent = jtToM(yearTargetJt);

    // 2. Opportunity 1 Tahun
    const oppRatio = yearTargetJt ? Math.round((yearOppJt / yearTargetJt) * 100) : 0;
    const elOppVal = document.getElementById('kpi-opp-val');
    if (elOppVal) elOppVal.textContent = jtToM(yearOppJt);
    const elOppTag = document.getElementById('kpi-opp-ratio-tag');
    if (elOppTag) elOppTag.textContent = oppRatio + '% Target';
    const elOppDesc = document.getElementById('kpi-opp-desc');
    if (elOppDesc) elOppDesc.textContent = rekap.length + ' CR / JR terdaftar';

    // 3. Revenue Diakui 1 Tahun
    const achievePct = yearTargetJt ? Math.round((yearRealJt / yearTargetJt) * 100) : 0;
    const oppWinRate = yearOppJt ? Math.round((yearRealJt / yearOppJt) * 100) : 0;
    const elRealVal = document.getElementById('kpi-real-val');
    if (elRealVal) elRealVal.textContent = jtToM(yearRealJt);
    const elRealTag = document.getElementById('kpi-real-pct-tag');
    if (elRealTag) elRealTag.textContent = achievePct + '% Target';
    const elRealDesc = document.getElementById('kpi-real-desc');
    if (elRealDesc) elRealDesc.textContent = 'Win Rate ' + oppWinRate + '% dari Opp';

    // 4. Revenue MostLikely 1 Tahun
    const mostLikelyOppRatio = yearOppJt ? Math.round((yearMostLikelyJt / yearOppJt) * 100) : 0;
    const elMostLikelyVal = document.getElementById('kpi-mostlikely-val');
    if (elMostLikelyVal) elMostLikelyVal.textContent = jtToM(yearMostLikelyJt);
    const elMostLikelyTag = document.getElementById('kpi-mostlikely-pct-tag');
    if (elMostLikelyTag) elMostLikelyTag.textContent = mostLikelyOppRatio + '% Opp';
    const elMostLikelyDesc = document.getElementById('kpi-mostlikely-desc');
    if (elMostLikelyDesc) elMostLikelyDesc.textContent = 'Proyeksi realisasi tinggi';

    // 5. Cash In 1 Tahun
    const collPct = yearRealJt ? Math.round((yearCashJt / yearRealJt) * 100) : 0;
    const elCashVal = document.getElementById('kpi-cash-val');
    if (elCashVal) elCashVal.textContent = jtToM(yearCashJt);
    const elCashTag = document.getElementById('kpi-cash-pct-tag');
    if (elCashTag) elCashTag.textContent = collPct + '% Real';
    const elCashDesc = document.getElementById('kpi-cash-desc');
    if (elCashDesc) elCashDesc.textContent = 'Dana masuk rekening';

    // 6. Belum Cair 1 Tahun
    const waitPct = yearRealJt ? Math.round((yearWaitJt / yearRealJt) * 100) : 0;
    const elWaitVal = document.getElementById('kpi-wait-val');
    if (elWaitVal) elWaitVal.textContent = jtToM(yearWaitJt);
    const elWaitTag = document.getElementById('kpi-wait-pct-tag');
    if (elWaitTag) elWaitTag.textContent = waitPct + '% Real';
    const elWaitDesc = document.getElementById('kpi-wait-desc');
    if (elWaitDesc) elWaitDesc.textContent = 'Dalam proses penagihan';

    // Opportunity Analysis Card (1 Tahun)
    const covRatio = yearTargetJt ? (yearOppJt / yearTargetJt).toFixed(2) : '0.00';
    const elBadge = document.getElementById('opp-badge-ratio');
    if (elBadge) elBadge.textContent = 'Coverage: ' + covRatio + 'x';

    const elFOpp = document.getElementById('funnel-opp-val');
    if (elFOpp) elFOpp.textContent = jtToM(yearOppJt) + ' M';
    const elFTgt = document.getElementById('funnel-target-val');
    if (elFTgt) elFTgt.textContent = jtToM(yearTargetJt) + ' M';
    const elFReal = document.getElementById('funnel-real-val');
    if (elFReal) elFReal.textContent = jtToM(yearRealJt) + ' M';
    const elFCash = document.getElementById('funnel-cash-val');
    if (elFCash) elFCash.textContent = jtToM(yearCashJt) + ' M';

    // Proportional progress bars based on Opportunity Pipeline
    const elTgtBar = document.getElementById('funnel-target-bar');
    if (elTgtBar) elTgtBar.style.width = '100%';
    const elTgtPct = document.getElementById('funnel-target-pct');
    if (elTgtPct) elTgtPct.textContent = (yearTargetJt ? ((yearOppJt / yearTargetJt) * 100).toFixed(1) : '0') + '% Target RKAP';

    const elRealBar = document.getElementById('funnel-real-bar');
    if (elRealBar) elRealBar.style.width = Math.min(100, Math.max(4, oppWinRate)) + '%';
    const elRealPct = document.getElementById('funnel-real-pct');
    if (elRealPct) elRealPct.textContent = 'Win Rate: ' + oppWinRate + '% dari Opp';

    const cashFromOpp = yearOppJt ? Math.round((yearCashJt / yearOppJt) * 100) : 0;
    const elCashBar = document.getElementById('funnel-cash-bar');
    if (elCashBar) elCashBar.style.width = Math.min(100, Math.max(4, cashFromOpp)) + '%';
    const elCashPct = document.getElementById('funnel-cash-pct');
    if (elCashPct) elCashPct.textContent = collPct + '% Realisasi (' + cashFromOpp + '% Opp)';

    const elStatWin = document.getElementById('stat-win-rate');
    if (elStatWin) elStatWin.textContent = oppWinRate + '%';
    const elStatCov = document.getElementById('stat-coverage');
    if (elStatCov) elStatCov.textContent = (yearTargetJt ? ((yearOppJt / yearTargetJt) * 100).toFixed(1) : '0') + '%';
    const elStatCol = document.getElementById('stat-collection');
    if (elStatCol) elStatCol.textContent = collPct + '%';

    // RENDER TRUE 2.5D ISOMETRIC CHART
    render25DIsometricChart();
    renderTrendStats(yearTargetJt, yearRealJt, yearCashJt, yearOppJt);
  }

  // =========================================================================
  // TRUE 2.5D ISOMETRIC SVG CHART MATHEMATICS (PRESERVED 100%)
  // =========================================================================
  const T = {
    cyan: '#00A3E0',
    cyanDeep: '#0072CE',
    green: '#00C076',
    greenDeep: '#10B981',
    yellow: '#FFD100',
    navy: '#0F172A',
    navyDeep: '#0A2540',
    text: '#0F172A',
    sub: '#64748B',
    muted: '#94A3B8',
    track: '#E2E8F0',
    rose: '#E11D48',
    purple: '#8B5CF6'
  };
  const C = { cash: T.green, wait: T.cyan, risk: T.rose, proc: T.muted };
  const SEG = [['cash','Terbayar'],['wait','Tertagih, menunggu bayar'],['risk','Bermasalah'],['proc','Masih proses']];
  const IX = 0.866, IY = 0.5;

  function shade(hexColor, factor) {
    var n = parseInt(hexColor.slice(1), 16);
    var r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
    if (factor >= 1) {
      r = Math.round(r + (255 - r) * (factor - 1));
      g = Math.round(g + (255 - g) * (factor - 1));
      b = Math.round(b + (255 - b) * (factor - 1));
    } else {
      r = Math.round(r * factor);
      g = Math.round(g * factor);
      b = Math.round(b * factor);
    }
    return 'rgb(' + r + ',' + g + ',' + b + ')';
  }

  function poly(pts) {
    return pts.map(p => p[0].toFixed(1) + ',' + p[1].toFixed(1)).join(' ');
  }

  let activeChartPopupMonth = -1;

  function syncMonthlyMetrics() {
    const annualTarget = (appData.annual && appData.annual.target) ? appData.annual.target : 0;
    const mTargetJt = Math.round((annualTarget * 1000) / 12);

    appData.months.forEach(m => {
      m.target = mTargetJt;
      m.opp = 0;
      m.cash = 0;
      m.wait = 0;
      m.risk = 0;
      m.proc = 0;
    });

    const rekap = appData.rekap || DEFAULT_REKAP || [];
    const detil = appData.detilAngsuran || DEFAULT_DETIL || [];

    function extractMonthIndex(dateStr) {
      if (!dateStr) return -1;
      const s = String(dateStr).trim().toLowerCase();
      const monthNames = ['jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt', 'nov', 'des'];
      for (let i = 0; i < monthNames.length; i++) {
        if (s.includes(monthNames[i])) return i;
      }
      const parts = s.split(/[-/.]/);
      if (parts.length === 3) {
        let m = parseInt(parts[1], 10) - 1;
        if (m >= 0 && m < 12) return m;
        if (parts[2].length === 4) {
          m = parseInt(parts[0], 10) - 1;
          if (m >= 0 && m < 12) return m;
        }
      }
      return -1;
    }

    let hasDatedRekap = false;
    rekap.forEach(r => {
      const valJt = Math.round((Number(r.NILAI_FULL) || 0) / 1e6);
      const d = r.TGLSURAT_STI || r.TGL_BAK || r.TGLSURAT_BPO || r.TGLSURAT_MDG || r.TGL_CREATED;
      const mIdx = extractMonthIndex(d);
      if (mIdx >= 0 && mIdx < 12) {
        appData.months[mIdx].opp += valJt;
        hasDatedRekap = true;
      }
    });

    let hasDatedDetil = false;
    detil.forEach(d => {
      const valJt = Math.round((Number(d.NILAI_TAHAP) || Number(d.NILAI_ANGSURAN) || 0) / 1e6);
      const dt = d.TGL_BAKB || d.TGLSURAT_STI || d.TGLSURAT_BPO || d.TGL_CREATED;
      const mIdx = extractMonthIndex(dt);
      if (mIdx >= 0 && mIdx < 12) {
        if ((d.STATUS_LUNAS || '').toUpperCase() === 'LUNAS') {
          appData.months[mIdx].cash += valJt;
        } else {
          appData.months[mIdx].wait += valJt;
        }
        hasDatedDetil = true;
      }
    });

    // If transactions don't have individual dates, associate current sample with active cutoff month (September = index 8)
    if (!hasDatedRekap) {
      let totOppJt = 0;
      rekap.forEach(r => { totOppJt += Math.round((Number(r.NILAI_FULL) || 0) / 1e6); });
      appData.months[8].opp = totOppJt;
    }

    if (!hasDatedDetil) {
      let totCashJt = 0, totWaitJt = 0;
      detil.forEach(d => {
        const valJt = Math.round((Number(d.NILAI_TAHAP) || Number(d.NILAI_ANGSURAN) || 0) / 1e6);
        if ((d.STATUS_LUNAS || '').toUpperCase() === 'LUNAS') {
          totCashJt += valJt;
        } else {
          totWaitJt += valJt;
        }
      });
      appData.months[8].cash = totCashJt;
      appData.months[8].wait = totWaitJt;
    }
  }

  function openChartPopup(i, e) {
    if (e) e.stopPropagation();
    activeChartPopupMonth = i;
    selectedMonthIndex = i;
    render25DIsometricChart();
    showHtmlChartPopup(i);
  }

  function closeChartPopup(e) {
    if (e) e.stopPropagation();
    activeChartPopupMonth = -1;
    const p = document.getElementById('chart-month-popup');
    if (p) p.style.display = 'none';
    render25DIsometricChart();
  }

  function showHtmlChartPopup(i) {
    const p = document.getElementById('chart-month-popup');
    if (!p) return;
    const m = appData.months[i] || {};
    const mReal = (m.cash || 0) + (m.wait || 0) + (m.risk || 0) + (m.proc || 0);

    // Position horizontally above month i
    const leftPct = Math.min(84, Math.max(16, ((i * 62 + 45) / 830) * 100));

    p.innerHTML = `
      <div style="display: flex; align-items: flex-start; justify-content: space-between; padding-right: 22px; margin-bottom: 10px; border-bottom: 1px solid rgba(255, 255, 255, 0.12); padding-bottom: 8px;">
        <div>
          <div style="font-size: 13.5px; font-weight: 800; color: #fff; letter-spacing: -0.01em;">${m.long} ${appData.year || '2026'}</div>
          <div style="font-size: 10px; color: #94A3B8; margin-top: 1px;">Realisasi Pendapatan</div>
        </div>
        <div style="font-size: 14px; font-weight: 800; color: var(--cyan);">${mReal > 0 ? jtToM(mReal) + ' M' : '0 M'}</div>
      </div>
      <div style="display: flex; flex-direction: column; gap: 6px; font-size: 11px;">
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <span style="color: #94A3B8; font-weight: 600;">Target:</span>
          <span style="color: #E2E8F0; font-weight: 700;">${m.target > 0 ? jtToM(m.target) + ' M' : '0 M'}</span>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <span style="color: #A78BFA; font-weight: 600;">Opportunity:</span>
          <span style="color: #A78BFA; font-weight: 700;">${m.opp > 0 ? jtToM(m.opp) + ' M' : '0 M'}</span>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <span style="color: var(--green); font-weight: 600;">Cash In:</span>
          <span style="color: var(--green); font-weight: 700;">${m.cash > 0 ? jtToM(m.cash) + ' M' : '0 M'}</span>
        </div>
      </div>
      <button class="popup-close-btn" onclick="closeChartPopup(event)" title="Tutup Popup">&times;</button>
    `;

    p.style.left = leftPct.toFixed(1) + '%';
    p.style.display = 'block';
  }

  function render25DIsometricChart() {
    syncMonthlyMetrics();

    const svg = document.getElementById('iso25d-chart');
    if (!svg) return;
    const s = 28, colW = 62, H = 168;
    let maxV = 0;

    appData.months.forEach(m => {
      const real = (m.cash || 0) + (m.wait || 0) + (m.risk || 0) + (m.proc || 0);
      maxV = Math.max(maxV, m.target || 0, real);
    });
    if (maxV === 0) maxV = 300000;

    const hScale = H / maxV;
    const dx = s * IX, dy = s * IY;

    let gridHtml = '';
    let barsHtml = '';

    appData.months.forEach((m, i) => {
      const ox = i * colW;
      function A(y)  { return [ox, -dy - y]; }
      function B(y)  { return [ox + dx, -y]; }
      function Cp(y) { return [ox, dy - y]; }
      function D(y)  { return [ox - dx, -y]; }

      // Ground diamond grid tile
      gridHtml += `<polygon points="${poly([A(0), B(0), Cp(0), D(0)])}" fill="#E2E8F0" stroke="#CBD5E1" stroke-width="1"/>`;

      const active = activeChartPopupMonth === i;
      const mReal = (m.cash || 0) + (m.wait || 0) + (m.risk || 0) + (m.proc || 0);

      let facesHtml = '';
      let acc2 = 0;
      let topColor = C.cash;
      let topMax = 0;

      SEG.forEach(g => {
        const v = m[g[0]] || 0;
        if (!v) return;
        const y1 = acc2 * hScale, y2 = (acc2 + v) * hScale;
        acc2 += v;
        if (v > topMax) { topMax = v; topColor = C[g[0]]; }

        // Left Face
        facesHtml += `<polygon points="${poly([D(y2), Cp(y2), Cp(y1), D(y1)])}" fill="${shade(C[g[0]], 0.95)}" stroke="rgba(255,255,255,0.34)" stroke-width="0.6"/>`;
        // Right Face
        facesHtml += `<polygon points="${poly([B(y2), Cp(y2), Cp(y1), B(y1)])}" fill="${shade(C[g[0]], 0.70)}" stroke="rgba(255,255,255,0.22)" stroke-width="0.6"/>`;
      });

      const hTot = acc2 * hScale;
      const tH = (m.target || 0) * hScale;
      const topH = Math.max(hTot, tH) + 9;

      const groupOpacity = active ? '1' : (activeChartPopupMonth >= 0 ? '0.55' : '0.90');
      const targetStroke = active ? '#0F172A' : '#CBD5E1';
      const targetStrokeWidth = active ? '1.5' : '1.2';

      // Dashed pointer line pointing from bar top towards popup when active
      let pointerLine = '';
      if (active) {
        pointerLine = `<line x1="${ox}" y1="${-topH}" x2="${ox}" y2="${-H - 30}" stroke="#00A3E0" stroke-width="1.8" stroke-dasharray="3 3"/>`;
      }

      barsHtml += `
        <g opacity="${groupOpacity}" style="transition: opacity 0.18s;">
          ${pointerLine}

          <!-- Target Wireframe (Dashed) -->
          <polygon points="${poly([A(tH), B(tH), Cp(tH), D(tH)])}" fill="none" stroke="${targetStroke}" stroke-width="${targetStrokeWidth}" stroke-dasharray="3 3"/>
          <polyline points="${poly([Cp(tH), Cp(0)])}" fill="none" stroke="${targetStroke}" stroke-width="${targetStrokeWidth}" stroke-dasharray="3 3"/>

          <!-- 3D Faces -->
          ${facesHtml}

          <!-- Top Diamond Face -->
          ${acc2 > 0 ? `<polygon points="${poly([A(hTot), B(hTot), Cp(hTot), D(hTot)])}" fill="${shade(topColor, 1.5)}" stroke="${shade(topColor, 0.98)}" stroke-width="0.9"/>` : ''}

          <!-- Month Label -->
          <text x="${ox.toFixed(1)}" y="${(dy + 18).toFixed(1)}" text-anchor="middle"
                font-family="'Plus Jakarta Sans',sans-serif" font-size="11" font-weight="${active ? '800' : '600'}"
                fill="${active ? T.cyanDeep : T.text}">
            ${m.short}
          </text>

          <!-- Top Value Text -->
          <text x="${ox.toFixed(1)}" y="${(-topH - 7).toFixed(1)}" text-anchor="middle"
                font-family="'Plus Jakarta Sans',sans-serif" font-size="12" font-weight="800"
                fill="${T.text}" style="display: ${active || mReal > 0 ? 'block' : 'none'};">
            ${mReal > 0 ? jtToM(mReal) : (active ? '0 M' : '')}
          </text>

          <!-- Clickable Hit Area -->
          <polygon class="chart-hit-area" points="${poly([D(topH), A(topH), B(topH), B(0), Cp(0), D(0)])}" fill="#000" fill-opacity="0" style="cursor: pointer;"
                   onclick="openChartPopup(${i}, event)"/>
        </g>
      `;
    });

    svg.innerHTML = gridHtml + barsHtml;

    // Synchronize HTML popup visibility
    if (activeChartPopupMonth >= 0) {
      showHtmlChartPopup(activeChartPopupMonth);
    } else {
      const p = document.getElementById('chart-month-popup');
      if (p) p.style.display = 'none';
    }
  }

  // Close chart popup when clicking outside
  document.addEventListener('click', function(e) {
    const popup = document.getElementById('chart-month-popup');
    if (!popup || popup.style.display === 'none') return;
    if (!popup.contains(e.target) && !e.target.closest('.chart-hit-area')) {
      closeChartPopup();
    }
  });

  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      closeChartPopup();
    }
  });

  // Trend stats below chart
  function renderTrendStats(targetJt, realJt, cashJt, oppJt) {
    const c = document.getElementById('trend-stats-container');
    if (!c) return;
    const rekapCount = (appData.rekap || DEFAULT_REKAP || []).length;
    c.innerHTML = `
      <div class="trend-stat">
        <span class="trend-stat-label">TARGET 1 TAHUN (${appData.year || '2026'})</span>
        <span class="trend-stat-val" style="color: var(--navy);">${jtToM(targetJt)} M</span>
      </div>
      <div class="trend-stat">
        <span class="trend-stat-label">REVENUE DIAKUI</span>
        <span class="trend-stat-val" style="color: var(--cyan-deep);">${jtToM(realJt)} M</span>
      </div>
      <div class="trend-stat">
        <span class="trend-stat-label">CASH IN TERBAYAR</span>
        <span class="trend-stat-val" style="color: var(--green-deep);">${jtToM(cashJt)} M</span>
      </div>
      <div class="trend-stat">
        <span class="trend-stat-label">PIPELINE OPPORTUNITY</span>
        <span class="trend-stat-val" style="color: var(--purple);">${jtToM(oppJt)} M</span>
      </div>
      <div class="trend-stat">
        <span class="trend-stat-label">STATUS PIPELINE</span>
        <span class="trend-stat-val" style="color: var(--navy); font-size: 14px;">${rekapCount} CR/JR Terdaftar</span>
      </div>
    `;
  }

  // Render Status Buckets for Selected Month
  function renderBuckets(m) {
    const c = document.getElementById('status-buckets-container');
    const tot = (m.cash + m.wait + m.risk + m.proc) || 1;
    const items = [
      { name: 'Terbayar', val: m.cash, color: '#00C076', note: 'Dana sudah masuk rekening perusahaan' },
      { name: 'Tertagih, menunggu bayar', val: m.wait, color: '#00A3E0', note: 'Invoice terbit, masih dalam termin bayar' },
      { name: 'Bermasalah', val: m.risk, color: '#E11D48', note: 'Overdue, dispute, dan kurang bayar' },
      { name: 'Masih proses', val: m.proc, color: '#94A3B8', note: 'BA atau PO belum lengkap — masih bisa dikejar' }
    ];

    let html = '';
    items.forEach(it => {
      const p = pctStr(it.val, tot);
      html += `
        <div class="bucket-row">
          <div class="bucket-left">
            <span class="bucket-dot" style="background: ${it.color};"></span>
            <div>
              <div class="bucket-name">${it.name}</div>
              <div class="bucket-note">${it.note}</div>
            </div>
          </div>
          <div class="bucket-right">
            <div class="bucket-amt">${jtToM(it.val)} M</div>
            <span class="bucket-pct" style="background: ${it.color}1a; color: ${it.color}; border: 1px solid ${it.color}40;">${p}</span>
          </div>
        </div>
      `;
    });
    c.innerHTML = html;
  }

  // Render Customers
  function renderCustomers() {
    const c = document.getElementById('customers-list-container');
    let html = '';
    appData.customers.forEach(cust => {
      const total = (cust.amount + cust.wait + cust.risk + cust.proc);
      html += `
        <div class="contrib-item">
          <div>
            <div class="contrib-name">${cust.name}</div>
            <div class="contrib-bar" style="width: ${total / 30}px; max-width: 140px;"></div>
          </div>
          <div style="text-align: right;">
            <div style="font-size: 13px; font-weight: 800; color: var(--navy);">${jtToM(total)} M</div>
            <div style="font-size: 10px; font-weight: 700; color: var(--cyan-deep);">Kontribusi utama</div>
          </div>
        </div>
      `;
    });
    c.innerHTML = html;
  }

  // Render Decisions
  function renderDecisions() {
    const c = document.getElementById('decisions-list-container');
    let html = '';
    appData.decisions.forEach(d => {
      const tagClass = d.tag === 'KRITIS' ? 'kritis' : d.tag === 'TINGGI' ? 'tinggi' : 'sedang';
      html += `
        <div class="decision-item">
          <span class="decision-tag ${tagClass}">${d.tag}</span>
          <div style="flex: 1;">
            <div class="decision-text">${d.title}</div>
            <div class="decision-sub">${d.sub}</div>
          </div>
          <div class="decision-amt">${d.amount}</div>
        </div>
      `;
    });
    c.innerHTML = html;
  }

  // =========================================================================
  // VIEW 2: MASTER DATA TABLE & INPUT (MENU MASTER)
  // =========================================================================
  function renderMasterTable() {
    const tbody = document.getElementById('monthly-table-tbody');
    tbody.innerHTML = '';

    let totTarget = 0, totOpp = 0, totReal = 0, totCash = 0;

    appData.months.forEach((m, i) => {
      const real = (m.cash || 0) + (m.wait || 0) + (m.risk || 0) + (m.proc || 0);
      totTarget += (m.target || 0);
      totOpp += (m.opp || 0);
      totReal += real;
      totCash += (m.cash || 0);

      const pct = m.target ? Math.round((real / m.target) * 100) : 0;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-weight: 800; color: var(--navy);">${m.long}</td>
        <td class="num">${fmtJt(m.target)}</td>
        <td class="num" style="color: var(--purple); font-weight: 800;">${fmtJt(m.opp || 0)}</td>
        <td class="num" style="color: var(--green-deep);">${fmtJt(m.cash)}</td>
        <td class="num" style="color: var(--cyan-deep);">${fmtJt(m.wait)}</td>
        <td class="num" style="color: var(--rose);">${fmtJt(m.risk)}</td>
        <td class="num" style="color: var(--muted);">${fmtJt(m.proc)}</td>
        <td class="num" style="color: var(--navy); font-weight: 800;">${fmtJt(real)}</td>
        <td style="text-align: center;"><span style="font-weight: 700; color: ${pct >= 90 ? 'var(--green-deep)' : 'var(--navy)'};">${pct}%</span></td>
        <td style="text-align: right;">
          <button class="btn-sm-edit" onclick="openEditMonthModal(${i})">Edit Nilai</button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    populateAnnualInputs();
  }

  // =========================================================================
  // MASTER DATA: INPUT 6 METRIK UTAMA TAHUNAN (2026)
  // =========================================================================
  function populateAnnualInputs() {
    const rekap = appData.rekap || DEFAULT_REKAP || [];
    const detil = appData.detilAngsuran || DEFAULT_DETIL || [];

    let sumFull = 0, sumTahap = 0, sumWbs = 0;
    rekap.forEach(r => {
      sumFull += (Number(r.NILAI_FULL) || 0);
      sumTahap += (Number(r.NILAI_TAHAP) || 0);
      sumWbs += (Number(r.NILAI_WBS) || 0);
    });

    let sumCashIn = 0, sumBelum = 0;
    detil.forEach(d => {
      const val = Number(d.NILAI_TAHAP) || Number(d.NILAI_ANGSURAN) || 0;
      if ((d.STATUS_LUNAS || '').toUpperCase() === 'LUNAS') {
        sumCashIn += val;
      } else {
        sumBelum += val;
      }
    });

    const targetM = appData.annual && appData.annual.target ? appData.annual.target : 186.00;
    const oppM = parseFloat((sumFull / 1e9).toFixed(2));
    const cashM = parseFloat((sumCashIn / 1e9).toFixed(2));
    const waitM = parseFloat((sumBelum / 1e9).toFixed(2));
    const realM = parseFloat(((sumCashIn + sumBelum) / 1e9).toFixed(2));
    const wbsM = parseFloat((sumWbs / 1e9).toFixed(2));
    const mostLikelyM = (appData.annual && appData.annual.mostLikely !== undefined)
      ? appData.annual.mostLikely
      : (wbsM > 0 ? wbsM : parseFloat((oppM * 0.85).toFixed(2)));

    if (!appData.annual) appData.annual = {};
    appData.annual.target = targetM;
    appData.annual.opp = oppM;
    appData.annual.real = realM;
    appData.annual.mostLikely = mostLikelyM;
    appData.annual.cash = cashM;
    appData.annual.wait = waitM;
    appData.annual.risk = 0.00;

    const lblYear = document.getElementById('label-annual-year');
    if (lblYear) lblYear.textContent = appData.year || '2026';

    // 1. Target 1 Tahun -> SATU-SATUNYA YANG BISA DI-INPUT MANUAL
    const inTarget = document.getElementById('in-annual-target');
    if (inTarget && document.activeElement !== inTarget) {
      inTarget.value = Number(targetM).toFixed(2);
    }

    const fmtM = v => Number(v).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    // 2-6. REMINDER CARDS
    const valOpp = document.getElementById('val-annual-opp');
    if (valOpp) valOpp.innerHTML = `${fmtM(oppM)} <span style="font-size: 13.5px; font-weight: 700; color: var(--purple);">M</span>`;

    const valReal = document.getElementById('val-annual-real');
    if (valReal) valReal.innerHTML = `${fmtM(realM)} <span style="font-size: 13.5px; font-weight: 700; color: var(--cyan-deep);">M</span>`;

    const valMostLikely = document.getElementById('val-annual-mostlikely');
    if (valMostLikely) valMostLikely.innerHTML = `${fmtM(mostLikelyM)} <span style="font-size: 13.5px; font-weight: 700; color: #6366F1;">M</span>`;

    const valCash = document.getElementById('val-annual-cash');
    if (valCash) valCash.innerHTML = `${fmtM(cashM)} <span style="font-size: 13.5px; font-weight: 700; color: var(--green-deep);">M</span>`;

    const valWait = document.getElementById('val-annual-wait');
    if (valWait) valWait.innerHTML = `${fmtM(waitM)} <span style="font-size: 13.5px; font-weight: 700; color: var(--muted);">M</span>`;

    // Update Percentage Preview Tags
    const oppRatio = targetM ? Math.round((oppM / targetM) * 100) : 0;
    const realRatio = targetM ? Math.round((realM / targetM) * 100) : 0;
    const mostLikelyRatio = oppM ? Math.round((mostLikelyM / oppM) * 100) : 0;
    const cashRatio = realM ? Math.round((cashM / realM) * 100) : 0;
    const waitRatio = realM ? Math.round((waitM / realM) * 100) : 0;

    const tagOpp = document.getElementById('tag-preview-opp');
    if (tagOpp) tagOpp.textContent = oppRatio + '% Target';
    const tagReal = document.getElementById('tag-preview-real');
    if (tagReal) tagReal.textContent = realRatio + '% Target';
    const tagMostLikely = document.getElementById('tag-preview-mostlikely');
    if (tagMostLikely) tagMostLikely.textContent = mostLikelyRatio + '% Opp';
    const tagCash = document.getElementById('tag-preview-cash');
    if (tagCash) tagCash.textContent = cashRatio + '% Real';
    const tagWait = document.getElementById('tag-preview-wait');
    if (tagWait) tagWait.textContent = waitRatio + '% Real';
  }

  function updateTargetPreview() {
    const target = parseFloat(document.getElementById('in-annual-target')?.value) || 0;
    const opp = appData.annual && appData.annual.opp !== undefined ? appData.annual.opp : 0;
    const real = appData.annual && appData.annual.real !== undefined ? appData.annual.real : 0;

    const oppRatio = target ? Math.round((opp / target) * 100) : 0;
    const realRatio = target ? Math.round((real / target) * 100) : 0;

    const tagOpp = document.getElementById('tag-preview-opp');
    if (tagOpp) tagOpp.textContent = oppRatio + '% Target';
    const tagReal = document.getElementById('tag-preview-real');
    if (tagReal) tagReal.textContent = realRatio + '% Target';
  }

  function saveTargetTahunan() {
    const targetVal = parseFloat(document.getElementById('in-annual-target').value) || 0;
    if (targetVal <= 0) {
      alert('Mohon masukkan angka Target 1 Tahun yang valid.');
      return;
    }

    if (!appData.annual) appData.annual = {};
    appData.annual.target = targetVal;

    const doSync = document.getElementById('chk-sync-monthly') ? document.getElementById('chk-sync-monthly').checked : true;
    if (doSync) {
      let curSum = 0;
      appData.months.forEach(m => { curSum += (m.target || 0); });
      if (curSum > 0) {
        const factor = (targetVal * 1000) / curSum;
        appData.months.forEach(m => { m.target = Math.round((m.target || 0) * factor); });
      }
    }

    saveAppData(appData);
    renderDashboard();
    renderMasterTable();
    showToast('Target 1 Tahun (' + appData.year + ') berhasil diperbarui menjadi ' + targetVal.toLocaleString('id-ID', {minimumFractionDigits: 2}) + ' M!');
  }

  // Opportunity Excel Modal Handlers
  function openOpportunityExcelModal() {
    const modal = document.getElementById('opportunityExcelModal');
    if (modal) modal.classList.add('open');
  }

  function closeOpportunityExcelModal() {
    const modal = document.getElementById('opportunityExcelModal');
    if (modal) modal.classList.remove('open');
  }

  function handleExcelOppUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    const fileName = file.name;
    showToast('File Excel ' + fileName + ' sedang diproses...');
    setTimeout(() => {
      closeOpportunityExcelModal();
      renderDashboard();
      renderMasterTable();
      showToast('Data Opportunity & PO dari Excel (' + fileName + ') berhasil disinkronkan!');
    }, 700);
    event.target.value = '';
  }

  // =========================================================================
  // MODAL EDIT BULAN
  // =========================================================================
  function openEditMonthModal(idx) {
    editingMonthIndex = idx;
    const m = appData.months[idx];
    document.getElementById('monthModalTitle').textContent = 'Edit Data Bulan ' + m.long + ' (' + appData.year + ')';
    document.getElementById('mMonthName').value = m.long + ' ' + appData.year;
    document.getElementById('mTargetBulan').value = m.target || 0;
    document.getElementById('mOppBulan').value = m.opp || 0;
    document.getElementById('mCashBulan').value = m.cash || 0;
    document.getElementById('mWaitBulan').value = m.wait || 0;
    document.getElementById('mRiskBulan').value = m.risk || 0;
    document.getElementById('mProcBulan').value = m.proc || 0;
    updateMonthPreview();
    document.getElementById('editMonthModal').classList.add('open');
  }

  function closeMonthModal() {
    document.getElementById('editMonthModal').classList.remove('open');
  }

  function updateMonthPreview() {
    const target = parseFloat(document.getElementById('mTargetBulan').value) || 0;
    const cash = parseFloat(document.getElementById('mCashBulan').value) || 0;
    const wait = parseFloat(document.getElementById('mWaitBulan').value) || 0;
    const risk = parseFloat(document.getElementById('mRiskBulan').value) || 0;
    const proc = parseFloat(document.getElementById('mProcBulan').value) || 0;
    const real = cash + wait + risk + proc;

    const pct = target ? Math.round((real / target) * 100) + '%' : '0%';
    const coll = real ? Math.round((cash / real) * 100) + '%' : '0%';

    document.getElementById('mPreviewRealBulan').textContent = fmtJt(real) + ' Jt';
    document.getElementById('mPreviewPctBulan').textContent = pct;
    document.getElementById('mPreviewCollBulan').textContent = coll;
  }

  function saveMonthData() {
    if (editingMonthIndex < 0) return;
    const m = appData.months[editingMonthIndex];
    m.target = parseFloat(document.getElementById('mTargetBulan').value) || 0;
    m.opp = parseFloat(document.getElementById('mOppBulan').value) || 0;
    m.cash = parseFloat(document.getElementById('mCashBulan').value) || 0;
    m.wait = parseFloat(document.getElementById('mWaitBulan').value) || 0;
    m.risk = parseFloat(document.getElementById('mRiskBulan').value) || 0;
    m.proc = parseFloat(document.getElementById('mProcBulan').value) || 0;

    saveAppData(appData);
    closeMonthModal();
    renderMasterTable();
    renderDashboard();
    showToast('Nilai bulan ' + m.long + ' berhasil diperbarui!');
  }

  // =========================================================================
  // MODAL EDIT ANNUAL PARAMETERS
  // =========================================================================
  function openEditAnnualModal() {
    document.getElementById('mAnnualYear').value = appData.year;
    document.getElementById('mAnnualOppCount').value = appData.oppCount || 48;
    document.getElementById('mAnnualCutoff').value = appData.cutoffDesc || '';

    // Calculate sum of months target & opp for reference
    let sumT = 0, sumO = 0;
    appData.months.forEach(m => { sumT += m.target; sumO += (m.opp || 0); });
    document.getElementById('mAnnualTarget').value = (sumT / 1000).toFixed(2);
    document.getElementById('mAnnualOpp').value = (sumO / 1000).toFixed(2);

    document.getElementById('annualModal').classList.add('open');
  }

  function closeAnnualModal() {
    document.getElementById('annualModal').classList.remove('open');
  }

  function saveAnnualData() {
    appData.oppCount = parseInt(document.getElementById('mAnnualOppCount').value) || 48;
    appData.cutoffDesc = document.getElementById('mAnnualCutoff').value.trim() || appData.cutoffDesc;

    // Proportionally adjust monthly target or opp if annual changed
    const newTargetJt = parseFloat(document.getElementById('mAnnualTarget').value) * 1000;
    const newOppJt = parseFloat(document.getElementById('mAnnualOpp').value) * 1000;

    let curTargetSum = 0, curOppSum = 0;
    appData.months.forEach(m => { curTargetSum += m.target; curOppSum += (m.opp || 0); });

    if (newTargetJt && curTargetSum && Math.abs(newTargetJt - curTargetSum) > 50) {
      const factorT = newTargetJt / curTargetSum;
      appData.months.forEach(m => { m.target = Math.round(m.target * factorT); });
    }
    if (newOppJt && curOppSum && Math.abs(newOppJt - curOppSum) > 50) {
      const factorO = newOppJt / curOppSum;
      appData.months.forEach(m => { m.opp = Math.round((m.opp || 0) * factorO); });
    }
    if (!appData.annual) appData.annual = {};
    if (newTargetJt) appData.annual.target = newTargetJt / 1000;
    if (newOppJt) appData.annual.opp = newOppJt / 1000;

    saveAppData(appData);
    closeAnnualModal();
    renderMasterTable();
    renderDashboard();
    showToast('Parameter tahunan ' + appData.year + ' berhasil disimpan!');
  }

  function resetDefaultData() {
    if (confirm('Kembalikan ke data standar PLN Icon Plus 2026? Semua input manual bulan akan direset.')) {
      appData = JSON.parse(JSON.stringify(DEFAULT_APP_STATE));
      saveAppData(appData);
      selectedMonthIndex = 8;
      showToast('Data master berhasil dipulihkan ke default!');
      renderMasterTable();
      renderDashboard();
    }
  }

  function exportDataJSON() {
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(appData, null, 2));
    const dl = document.createElement('a');
    dl.setAttribute('href', dataStr);
    dl.setAttribute('download', 'PLN_IconPlus_Revenue_' + appData.year + '_' + new Date().toISOString().slice(0, 10) + '.json');
    dl.click();
    showToast('File data JSON master berhasil di-download!');
  }

  function importDataJSON(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const imported = JSON.parse(e.target.result);
        if (imported && imported.months && imported.months.length === 12) {
          appData = imported;
          saveAppData(appData);
          renderMasterTable();
          renderDashboard();
          showToast('Data master berhasil di-import!');
        } else {
          alert('Format data JSON tidak sesuai format master 12 bulan.');
        }
      } catch (err) {
        alert('Gagal membaca file JSON: ' + err.message);
      }
    };
    reader.readAsText(file);
    event.target.value = '';
  }

  function handleSearch(q) {
    if (!q || !q.trim()) return;
    const term = q.toLowerCase();
    const matchIdx = appData.months.findIndex(m => m.long.toLowerCase().includes(term) || m.short.toLowerCase().includes(term));
    if (matchIdx >= 0) {
      selectedMonthIndex = matchIdx;
      renderDashboard();
    }
  }

  // =========================================================================
  // LIVE REAL-TIME CLOCK & REFRESH HANDLER (WIB)
  // =========================================================================
  function formatDateTimeWIB(date) {
    const d = date || new Date();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    const day = String(d.getDate()).padStart(2, '0');
    const month = months[d.getMonth()];
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
    return `${day} ${month} ${year} ${hours}:${minutes}:${seconds} WIB`;
  }

  function updateLiveRefreshClock() {
    const timeStr = formatDateTimeWIB(new Date());
    const progEl = document.getElementById('prog-live-refresh-clock');
    if (progEl) progEl.textContent = timeStr;
  }

  function manualRefreshData() {
    const icon = document.getElementById('refresh-icon-svg');
    if (icon) {
      icon.style.transform = 'rotate(360deg)';
      setTimeout(() => { icon.style.transform = 'none'; }, 600);
    }
    updateLiveRefreshClock();
    if (document.getElementById('view-prognosa') && document.getElementById('view-prognosa').style.display !== 'none') {
      applyPrognosaFilters();
    } else if (document.getElementById('view-opportunity') && document.getElementById('view-opportunity').style.display !== 'none') {
      renderOpportunityView();
    } else {
      renderDashboard();
    }
    showToast('Data dashboard berhasil disegarkan secara real-time.');
  }

  window.addEventListener('DOMContentLoaded', () => {
    renderDashboard();
    updateLiveRefreshClock();
    setInterval(updateLiveRefreshClock, 1000);
  });


  // =========================================================================
  // OPPORTUNITY & PO MODULE (SAMPLING DATA REV3: REKAP 36 KOLOM & DETIL 37 KOLOM)
  // =========================================================================

  const DEFAULT_REKAP = [
    {
      "ID_TRANSAKSI": 1,
      "LAYANAN": "AP2T",
      "TIPE": "CR",
      "URAIAN": "CR Proses Billing Program Insentif bagi Pelanggan yang memiliki Captive Power",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "31959/STI.01.02/B01060500/2020", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "YA", "PATHFOLDER_WBS": "ADA", "NILAI_WBS": 100000000,
      "ANGSURAN": "TIDAK",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "NILAI_TAHAP": 100000000, "NILAI_FULL": 100000000,
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    },
    {
      "ID_TRANSAKSI": 2,
      "LAYANAN": "AP2T",
      "TIPE": "CR",
      "URAIAN": "CR Integrasi & Update Data DIL AP2T dari Aplikasi Selain AP2T",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "1351/STI.01.02/B01060500/2021", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "YA", "PATHFOLDER_WBS": "ADA", "NILAI_WBS": 200000000,
      "ANGSURAN": "TIDAK",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "NILAI_TAHAP": 200000000, "NILAI_FULL": 200000000,
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    },
    {
      "ID_TRANSAKSI": 3,
      "LAYANAN": "AP2T",
      "TIPE": "CR",
      "URAIAN": "CR Implementasi Layanan Satu Pintu (LSP) pada Permohonan PB Kolektif di AP2T",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "13839/STI.01.02/C01070500/2021", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "YA", "PATHFOLDER_WBS": "ADA", "NILAI_WBS": 300000000,
      "ANGSURAN": "TIDAK",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "NILAI_TAHAP": 300000000, "NILAI_FULL": 300000000,
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    },
    {
      "ID_TRANSAKSI": 4,
      "LAYANAN": "AP2T",
      "TIPE": "CR",
      "URAIAN": "CR Integrasi CSMS EV Charging ke Billing System AP2T",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "13840/STI.01.02/C01070500/2021", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "YA", "PATHFOLDER_WBS": "ADA", "NILAI_WBS": 400000000,
      "ANGSURAN": "TIDAK",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "NILAI_TAHAP": 400000000, "NILAI_FULL": 400000000,
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    },
    {
      "ID_TRANSAKSI": 5,
      "LAYANAN": "AP2T",
      "TIPE": "CR",
      "URAIAN": "CR Laporan REC dan Penyempurnaan Monitoring REC",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "13840/STI.01.02/C01070500/2021", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "YA", "PATHFOLDER_WBS": "ADA", "NILAI_WBS": 500000000,
      "ANGSURAN": "TIDAK",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "NILAI_TAHAP": 500000000, "NILAI_FULL": 500000000,
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    },
    {
      "ID_TRANSAKSI": 6,
      "LAYANAN": "AP2T",
      "TIPE": "JR",
      "URAIAN": "JR Penambahan Sub Menu Laporan pada Modul Manajemen Instansi Vertikal di AP2T\nRevisi Penyesuaian CR Pada Aplikasi AP2T Modul Manajemen Instansi Vertikal (MIV)Dalam Rangka Pembayaran Tagihan Listrik Kementerian Keuangan Melalui Platform SAKTI",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "28768/STI.01.02/C01070500/2021\n36941/STI.01.02/C01070500/2021\n62194/STI.01.02/C01070500/2021", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "YA", "PATHFOLDER_WBS": "ADA", "NILAI_WBS": 100000000,
      "ANGSURAN": "YA",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "NILAI_TAHAP": 100000000, "NILAI_FULL": 300000000,
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    }
  ];

  const DEFAULT_DETIL = [
    {
      "ID_TRANSAKSI": 6,
      "LAYANAN": "AP2T",
      "TIPE": "JR",
      "URAIAN": "JR Penambahan Sub Menu Laporan pada Modul Manajemen Instansi Vertikal di AP2T\nRevisi Penyesuaian CR Pada Aplikasi AP2T Modul Manajemen Instansi Vertikal (MIV)Dalam Rangka Pembayaran Tagihan Listrik Kementerian Keuangan Melalui Platform SAKTI",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "", "PATHFOLDER_WBS": "",
      "ANGSURAN": "",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "TAHAP": 1,
      "NILAI_TAHAP": 100000000,
      "NILAI_FULL": 300000000,
      "STATUS_LUNAS": "LUNAS",
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    },
    {
      "ID_TRANSAKSI": 6,
      "LAYANAN": "AP2T",
      "TIPE": "JR",
      "URAIAN": "JR Penambahan Sub Menu Laporan pada Modul Manajemen Instansi Vertikal di AP2T\nRevisi Penyesuaian CR Pada Aplikasi AP2T Modul Manajemen Instansi Vertikal (MIV)Dalam Rangka Pembayaran Tagihan Listrik Kementerian Keuangan Melalui Platform SAKTI",
      "TGLSURAT_BPO": "", "NOSURAT_BPO": "", "PATHFOLDER_BPO": "",
      "TGLSURAT_MDG": "", "NOSURAT_MDG": "", "PATHFOLDER_MDG": "",
      "TGLSURAT_STI": "", "NOSURAT_STI": "", "PATHFOLDER_STI": "",
      "TGL_BAK": "", "NOSURAT_BAK": "", "PATHFOLDER_BAK": "",
      "STATUS_UAT": "", "PATHFOLDER_BAUT": "",
      "STATUS_DEPLOY": "", "PATHFOLDER_DEPLOY": "",
      "STATUS_WBS": "", "PATHFOLDER_WBS": "",
      "ANGSURAN": "",
      "STATUS_BAKB": "", "PATHFOLDER_BAKB": "",
      "TAHAP": 2,
      "NILAI_TAHAP": 200000000,
      "NILAI_FULL": 300000000,
      "STATUS_LUNAS": "BELUM",
      "STATUS_MOM": "", "PATHFOLDER_MOM": "",
      "TGL_CREATED": "", "CREATED_BY": "",
      "TGL_KOREKSI": "", "KOREKSI_BY": "",
      "TGL_BATAL": "", "BATAL_BY": ""
    }
  ];

  let currentOppTab = 'rekap';
  let editingRekapId = null;
  let editingDetilIdx = null;

  // Ensure state exists
  if (!appData.rekap || !appData.rekap.length) {
    appData.rekap = JSON.parse(JSON.stringify(DEFAULT_REKAP));
  }
  if (!appData.detilAngsuran || !appData.detilAngsuran.length) {
    appData.detilAngsuran = JSON.parse(JSON.stringify(DEFAULT_DETIL));
  }

  function showOpportunityView() {
    document.getElementById('view-dashboard').style.display = 'none';
    document.getElementById('view-master').style.display = 'none';
    const vProg = document.getElementById('view-prognosa'); if (vProg) vProg.style.display = 'none';
    const vMProg = document.getElementById('view-master-prognosa'); if (vMProg) vMProg.style.display = 'none';
    const vPa = document.getElementById('view-monitoring-pa'); if (vPa) vPa.style.display = 'none';
    document.getElementById('view-opportunity').style.display = 'block';

    document.getElementById('nav-revenue').classList.remove('active');
    const nProg = document.getElementById('nav-prognosa'); if (nProg) nProg.classList.remove('active');
    const nMPa = document.getElementById('nav-monitoring-pa'); if (nMPa) nMPa.classList.remove('active');
    const nMProg = document.getElementById('nav-master-prognosa'); if (nMProg) nMProg.classList.remove('active');
    document.getElementById('nav-master').classList.remove('active');
    const navOpp = document.getElementById('nav-opportunity');
    if (navOpp) navOpp.classList.add('active');

    document.getElementById('top-title').textContent = 'Opportunity & PO';
    document.getElementById('top-sub').textContent = 'DATA REKAP CR/JR (36 KOLOM) & DETIL ANGSURAN (37 KOLOM) · SAMPLING DATA REV3';

    renderOpportunityView();
  }

  function switchOppTab(tab) {
    currentOppTab = tab;
    const btnRekap = document.getElementById('tab-btn-rekap');
    const btnDetil = document.getElementById('tab-btn-detil');
    const wrapRekap = document.getElementById('opp-table-rekap-wrap');
    const wrapDetil = document.getElementById('opp-table-detil-wrap');
    const btnAddRekap = document.getElementById('btn-add-rekap');
    const btnAddDetil = document.getElementById('btn-add-detil');

    if (tab === 'rekap') {
      btnRekap.classList.add('active');
      btnDetil.classList.remove('active');
      wrapRekap.style.display = 'block';
      wrapDetil.style.display = 'none';
      if (btnAddRekap) btnAddRekap.style.display = 'flex';
      if (btnAddDetil) btnAddDetil.style.display = 'none';
    } else {
      btnRekap.classList.remove('active');
      btnDetil.classList.add('active');
      wrapRekap.style.display = 'none';
      wrapDetil.style.display = 'block';
      if (btnAddRekap) btnAddRekap.style.display = 'none';
      if (btnAddDetil) btnAddDetil.style.display = 'flex';
    }
  }

  function filterOppData() {
    renderOpportunityView();
  }

  function checkHasUat(val) {
    if (!val) return false;
    const s = String(val).trim().toUpperCase();
    if (!s || s === 'TIDAK' || s === 'BELUM' || s === '-' || s === 'NO' || s === 'NAN' || s === 'NULL') return false;
    return s === 'YA' || s === 'ADA' || s === 'LULUS' || s === 'SELESAI' || s === 'SUDAH' || s.includes('LULUS') || s.includes('BAUT') || s.includes('SUDAH');
  }

  function checkHasDeploy(val) {
    if (!val) return false;
    const s = String(val).trim().toUpperCase();
    if (!s || s === 'TIDAK' || s === 'BELUM' || s === '-' || s === 'NO' || s === 'NAN' || s === 'NULL') return false;
    return s === 'YA' || s === 'ADA' || s === 'DEPLOYED' || s === 'SELESAI' || s === 'SUDAH' || s.includes('GO-LIVE') || s.includes('GOLIVE') || s.includes('DEPLOY');
  }

  function toggleFilterNoWbs() {
    const sel = document.getElementById('opp-filter-wbs');
    if (!sel) return;
    sel.value = (sel.value === 'TIDAK') ? '' : 'TIDAK';
    if (typeof switchOppTab === 'function') switchOppTab('rekap');
    filterOppData();
  }

  function toggleFilterNoUat() {
    const sel = document.getElementById('opp-filter-uat');
    if (!sel) return;
    sel.value = (sel.value === 'TIDAK') ? '' : 'TIDAK';
    if (typeof switchOppTab === 'function') switchOppTab('rekap');
    filterOppData();
  }

  function toggleFilterNoDeploy() {
    const sel = document.getElementById('opp-filter-deploy');
    if (!sel) return;
    sel.value = (sel.value === 'TIDAK') ? '' : 'TIDAK';
    if (typeof switchOppTab === 'function') switchOppTab('rekap');
    filterOppData();
  }

  function renderOpportunityView() {
    const rekap = appData.rekap || [];
    const detil = appData.detilAngsuran || [];

    // Calculate Summary Metrics
    let sumFull = 0, sumTahap = 0, sumWbs = 0;
    let sumNoWbs = 0, countNoWbs = 0;
    let sumNoUat = 0, countNoUat = 0;
    let sumNoDeploy = 0, countNoDeploy = 0;

    rekap.forEach(r => {
      const full = Number(r.NILAI_FULL) || 0;
      const wbsVal = Number(r.NILAI_WBS) || 0;
      const stWbs = (r.STATUS_WBS || '').toString().trim().toUpperCase();
      const hasWbs = (stWbs === 'YA' || stWbs === 'ADA') && wbsVal > 0;

      sumFull += full;
      sumTahap += (Number(r.NILAI_TAHAP) || 0);
      sumWbs += wbsVal;

      if (!hasWbs) {
        sumNoWbs += full;
        countNoWbs++;
      }

      if (!checkHasUat(r.STATUS_UAT)) {
        sumNoUat += full;
        countNoUat++;
      }

      if (!checkHasDeploy(r.STATUS_DEPLOY)) {
        sumNoDeploy += full;
        countNoDeploy++;
      }
    });

    let sumCashIn = 0, sumBelum = 0;
    detil.forEach(d => {
      const val = Number(d.NILAI_TAHAP) || 0;
      if ((d.STATUS_LUNAS || '').toUpperCase() === 'LUNAS') {
        sumCashIn += val;
      } else {
        sumBelum += val;
      }
    });

    // Update KPI Card Numbers (in Miliar / Juta)
    const fmtM = val => (val / 1e9).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' <span style="font-size: 14px; font-weight: 700;">M</span>';
    const fmtJtOrM = val => (val >= 1e9) ? fmtM(val) : (val / 1e6).toLocaleString('id-ID') + ' <span style="font-size: 13px; font-weight: 700;">Jt</span>';

    if (document.getElementById('opp-kpi-total')) document.getElementById('opp-kpi-total').innerHTML = fmtM(sumFull);
    if (document.getElementById('opp-kpi-wbs')) document.getElementById('opp-kpi-wbs').innerHTML = fmtJtOrM(sumWbs);

    // KPI Pekerjaan Tanpa WBS
    if (document.getElementById('opp-kpi-nowbs')) {
      document.getElementById('opp-kpi-nowbs').innerHTML = sumNoWbs > 0 ? fmtJtOrM(sumNoWbs) : '0 <span style="font-size: 14px; font-weight: 700;">Jt</span>';
    }
    if (document.getElementById('opp-kpi-nowbs-sub')) {
      document.getElementById('opp-kpi-nowbs-sub').textContent = countNoWbs > 0 
        ? (countNoWbs + ' pekerjaan belum terbit WBS')
        : '0 pekerjaan (seluruhnya ber-WBS)';
    }
    if (document.getElementById('opp-kpi-nowbs-tag')) {
      document.getElementById('opp-kpi-nowbs-tag').textContent = countNoWbs > 0 ? (countNoWbs + ' Non-WBS') : '0 Non-WBS';
    }

    // KPI Tidak Ada BA UAT
    if (document.getElementById('opp-kpi-nouat')) {
      document.getElementById('opp-kpi-nouat').innerHTML = sumNoUat > 0 ? fmtJtOrM(sumNoUat) : '0 <span style="font-size: 14px; font-weight: 700;">Jt</span>';
    }
    if (document.getElementById('opp-kpi-nouat-sub')) {
      document.getElementById('opp-kpi-nouat-sub').textContent = countNoUat > 0
        ? (countNoUat + ' pekerjaan belum BA UAT')
        : '0 pekerjaan (seluruhnya ber-UAT)';
    }
    if (document.getElementById('opp-kpi-nouat-tag')) {
      document.getElementById('opp-kpi-nouat-tag').textContent = countNoUat > 0 ? (countNoUat + ' Tanpa UAT') : '0 Tanpa UAT';
    }

    // KPI Tidak Ada BA DEPLOY
    if (document.getElementById('opp-kpi-nodeploy')) {
      document.getElementById('opp-kpi-nodeploy').innerHTML = sumNoDeploy > 0 ? fmtJtOrM(sumNoDeploy) : '0 <span style="font-size: 14px; font-weight: 700;">Jt</span>';
    }
    if (document.getElementById('opp-kpi-nodeploy-sub')) {
      document.getElementById('opp-kpi-nodeploy-sub').textContent = countNoDeploy > 0
        ? (countNoDeploy + ' pekerjaan belum BA Deploy')
        : '0 pekerjaan (seluruhnya deploy)';
    }
    if (document.getElementById('opp-kpi-nodeploy-tag')) {
      document.getElementById('opp-kpi-nodeploy-tag').textContent = countNoDeploy > 0 ? (countNoDeploy + ' Tanpa Deploy') : '0 Tanpa Deploy';
    }

    if (document.getElementById('opp-kpi-tahap')) document.getElementById('opp-kpi-tahap').innerHTML = fmtM(sumTahap);
    if (document.getElementById('opp-kpi-cash')) document.getElementById('opp-kpi-cash').innerHTML = fmtJtOrM(sumCashIn);
    if (document.getElementById('opp-kpi-wait')) document.getElementById('opp-kpi-wait').innerHTML = fmtJtOrM(sumBelum);

    // Active visual highlight on cards if filtered
    const selWbs = document.getElementById('opp-filter-wbs');
    const cardNoWbs = document.getElementById('card-filter-nowbs');
    if (cardNoWbs) {
      if (selWbs && selWbs.value === 'TIDAK') {
        cardNoWbs.style.borderColor = '#E11D48';
        cardNoWbs.style.background = '#FFF1F2';
        cardNoWbs.style.boxShadow = '0 0 0 2px #E11D48';
      } else {
        cardNoWbs.style.borderColor = '';
        cardNoWbs.style.background = '';
        cardNoWbs.style.boxShadow = '';
      }
    }

    const selUat = document.getElementById('opp-filter-uat');
    const cardNoUat = document.getElementById('card-filter-nouat');
    if (cardNoUat) {
      if (selUat && selUat.value === 'TIDAK') {
        cardNoUat.style.borderColor = '#EA580C';
        cardNoUat.style.background = '#FFF7ED';
        cardNoUat.style.boxShadow = '0 0 0 2px #EA580C';
      } else {
        cardNoUat.style.borderColor = '';
        cardNoUat.style.background = '';
        cardNoUat.style.boxShadow = '';
      }
    }

    const selDeploy = document.getElementById('opp-filter-deploy');
    const cardNoDeploy = document.getElementById('card-filter-nodeploy');
    if (cardNoDeploy) {
      if (selDeploy && selDeploy.value === 'TIDAK') {
        cardNoDeploy.style.borderColor = '#0284C7';
        cardNoDeploy.style.background = '#F0F9FF';
        cardNoDeploy.style.boxShadow = '0 0 0 2px #0284C7';
      } else {
        cardNoDeploy.style.borderColor = '';
        cardNoDeploy.style.background = '';
        cardNoDeploy.style.boxShadow = '';
      }
    }

    // Update Badges Count
    document.getElementById('badge-rekap-count').textContent = rekap.length;
    document.getElementById('badge-detil-count').textContent = detil.length;

    // Filter controls
    const keyword = (document.getElementById('opp-search-input')?.value || '').toLowerCase().trim();
    const fLayanan = document.getElementById('opp-filter-layanan')?.value || '';
    const fTipe = document.getElementById('opp-filter-tipe')?.value || '';
    const fAngsuran = document.getElementById('opp-filter-angsuran')?.value || '';
    const fWbs = document.getElementById('opp-filter-wbs')?.value || '';
    const fUat = document.getElementById('opp-filter-uat')?.value || '';
    const fDeploy = document.getElementById('opp-filter-deploy')?.value || '';

    // Filter Rekap
    const filteredRekap = rekap.filter(r => {
      if (fLayanan && r.LAYANAN !== fLayanan) return false;
      if (fTipe && r.TIPE !== fTipe) return false;
      if (fAngsuran && (r.ANGSURAN || '').toUpperCase() !== fAngsuran) return false;
      if (fWbs) {
        const stWbs = (r.STATUS_WBS || '').toString().trim().toUpperCase();
        const nilaiWbs = Number(r.NILAI_WBS) || 0;
        const hasWbs = (stWbs === 'YA' || stWbs === 'ADA') && nilaiWbs > 0;
        if (fWbs === 'YA' && !hasWbs) return false;
        if (fWbs === 'TIDAK' && hasWbs) return false;
      }
      if (fUat) {
        const hasUat = checkHasUat(r.STATUS_UAT);
        if (fUat === 'YA' && !hasUat) return false;
        if (fUat === 'TIDAK' && hasUat) return false;
      }
      if (fDeploy) {
        const hasDeploy = checkHasDeploy(r.STATUS_DEPLOY);
        if (fDeploy === 'YA' && !hasDeploy) return false;
        if (fDeploy === 'TIDAK' && hasDeploy) return false;
      }
      if (keyword) {
        const text = ((r.URAIAN || '') + ' ' + (r.NOSURAT_STI || '') + ' ' + (r.ID_TRANSAKSI || '')).toLowerCase();
        if (!text.includes(keyword)) return false;
      }
      return true;
    });

    renderRekapTable(filteredRekap);

    // Filter Detil
    const filteredDetil = detil.filter(d => {
      if (fLayanan && d.LAYANAN !== fLayanan) return false;
      if (fTipe && d.TIPE !== fTipe) return false;
      if (fWbs) {
        const stWbs = (d.STATUS_WBS || '').toString().trim().toUpperCase();
        const hasWbs = (stWbs === 'YA' || stWbs === 'ADA');
        if (fWbs === 'YA' && !hasWbs) return false;
        if (fWbs === 'TIDAK' && hasWbs) return false;
      }
      if (fUat) {
        const hasUat = checkHasUat(d.STATUS_UAT);
        if (fUat === 'YA' && !hasUat) return false;
        if (fUat === 'TIDAK' && hasUat) return false;
      }
      if (fDeploy) {
        const hasDeploy = checkHasDeploy(d.STATUS_DEPLOY);
        if (fDeploy === 'YA' && !hasDeploy) return false;
        if (fDeploy === 'TIDAK' && hasDeploy) return false;
      }
      if (keyword) {
        const text = ((d.URAIAN || '') + ' ' + (d.ID_TRANSAKSI || '') + ' ' + (d.STATUS_LUNAS || '')).toLowerCase();
        if (!text.includes(keyword)) return false;
      }
      return true;
    });

    renderDetilAngsuranTable(filteredDetil);

    // Update Footers
    document.getElementById('opp-rekap-footer-sum').textContent = 'Total Opportunity (' + filteredRekap.length + ' item): Rp ' + (filteredRekap.reduce((a, b) => a + (Number(b.NILAI_FULL) || 0), 0)).toLocaleString('id-ID');
    document.getElementById('opp-detil-footer-sum').textContent = 'Cash In Lunas: Rp ' + sumCashIn.toLocaleString('id-ID') + ' | Belum Cair: Rp ' + sumBelum.toLocaleString('id-ID');

    syncOpportunityToDashboard(sumFull, sumCashIn, sumBelum);
  }

  function renderRekapTable(items) {
    const tbody = document.getElementById('opp-rekap-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="12" style="text-align: center; padding: 24px; color: var(--muted);">Tidak ada data Rekap CR/JR yang sesuai dengan filter.</td></tr>';
      return;
    }

    items.forEach(r => {
      const isAngsuran = (r.ANGSURAN || '').toUpperCase() === 'YA';
      const badgeTipe = r.TIPE === 'JR' ? '<span class="badge-jr">JR</span>' : '<span class="badge-cr">CR</span>';
      const badgeAngsuran = isAngsuran
        ? '<span class="badge-belum" style="font-size: 10px;">YA (Termin)</span>'
        : '<span class="kpi-tag" style="background: #F1F5F9; color: var(--sub);">TIDAK</span>';

      const wbsVal = Number(r.NILAI_WBS) || 0;
      const wbsText = r.STATUS_WBS === 'YA'
        ? '<span class="badge-wbs">WBS: Rp ' + (wbsVal / 1e6).toLocaleString('id-ID') + ' Jt</span>'
        : '<span style="color: var(--muted);">-</span>';

      const hasUat = checkHasUat(r.STATUS_UAT);
      const uatBadge = hasUat
        ? `<span class="kpi-tag" style="background: #ECFDF5; color: #047857; font-weight: 700;">${r.STATUS_UAT || 'YA'}</span>`
        : `<span class="kpi-tag" style="background: #FFF7ED; color: #EA580C; font-weight: 700; border: 1px solid rgba(234, 88, 12, 0.25);">Belum Ada</span>`;

      const hasDeploy = checkHasDeploy(r.STATUS_DEPLOY);
      const deployBadge = hasDeploy
        ? `<span class="kpi-tag" style="background: #ECFDF5; color: #047857; font-weight: 700;">${r.STATUS_DEPLOY || 'YA'}</span>`
        : `<span class="kpi-tag" style="background: #F0F9FF; color: #0284C7; font-weight: 700; border: 1px solid rgba(2, 132, 199, 0.25);">Belum Ada</span>`;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align: center; font-weight: 800; color: var(--navy);">${r.ID_TRANSAKSI}</td>
        <td><span style="font-weight: 700; color: var(--navy);">${r.LAYANAN || '-'}</span></td>
        <td style="text-align: center;">${badgeTipe}</td>
        <td>
          <div style="font-weight: 700; color: var(--navy); line-height: 1.35; max-width: 380px;">${r.URAIAN || '-'}</div>
          ${r.PATHFOLDER_STI ? '<div class="table-subtext">Arsip: ' + r.PATHFOLDER_STI + '</div>' : ''}
        </td>
        <td style="font-size: 11px; color: var(--sub); white-space: pre-line;">${r.NOSURAT_STI || '-'}</td>
        <td style="text-align: right;">${wbsText}</td>
        <td style="text-align: center;">${uatBadge}</td>
        <td style="text-align: center;">${deployBadge}</td>
        <td style="text-align: center;">${badgeAngsuran}</td>
        <td class="num" style="color: var(--cyan-deep); font-weight: 700;">Rp ${(Number(r.NILAI_TAHAP) || 0).toLocaleString('id-ID')}</td>
        <td class="num" style="color: var(--purple); font-weight: 800;">Rp ${(Number(r.NILAI_FULL) || 0).toLocaleString('id-ID')}</td>
        <td style="text-align: right; white-space: nowrap;">
          <button class="btn-sm-edit" onclick="openEditRekapModal(${r.ID_TRANSAKSI})" title="Edit seluruh 36 kolom">Edit 36 Kolom</button>
          ${isAngsuran ? `<button class="btn-sm-edit" style="background: #EDE9FE; color: #6D28D9; border-color: #C4B5FD; margin-left: 4px;" onclick="openAddDetilModal(${r.ID_TRANSAKSI})" title="Tambah Termin Angsuran">+ Termin</button>` : ''}
          <button class="btn-sm-edit" style="color: var(--rose); margin-left: 4px;" onclick="deleteRekap(${r.ID_TRANSAKSI})" title="Hapus transaksi">&times;</button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }

  function renderDetilAngsuranTable(items) {
    const tbody = document.getElementById('opp-detil-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 24px; color: var(--muted);">Tidak ada data Detil Angsuran yang sesuai.</td></tr>';
      return;
    }

    items.forEach((d, idx) => {
      const isLunas = (d.STATUS_LUNAS || '').toUpperCase() === 'LUNAS';
      const badgeLunas = isLunas
        ? '<span class="badge-lunas">✓ LUNAS (CASH IN)</span>'
        : '<span class="badge-belum">⏳ BELUM CAIR</span>';

      const badgeTipe = d.TIPE === 'JR' ? '<span class="badge-jr">JR</span>' : '<span class="badge-cr">CR</span>';

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align: center; font-weight: 800; color: var(--navy);">${d.ID_TRANSAKSI}</td>
        <td><span style="font-weight: 700;">${d.LAYANAN || '-'}</span> ${badgeTipe}</td>
        <td>
          <div style="font-weight: 700; color: var(--navy); line-height: 1.3; max-width: 360px;">${d.URAIAN || '-'}</div>
        </td>
        <td style="text-align: center; font-weight: 800; color: #6D28D9; font-size: 13px;">Tahap ${d.TAHAP || 1}</td>
        <td class="num" style="color: var(--cyan-deep); font-weight: 800;">Rp ${(Number(d.NILAI_TAHAP) || 0).toLocaleString('id-ID')}</td>
        <td class="num" style="color: var(--purple); font-weight: 700;">Rp ${(Number(d.NILAI_FULL) || 0).toLocaleString('id-ID')}</td>
        <td style="text-align: center;">
          <div style="display: flex; flex-direction: column; align-items: center; gap: 4px;">
            ${badgeLunas}
            <button class="btn-sm-edit" style="font-size: 9px; padding: 1px 6px;" onclick="toggleDetilLunas(${idx})">Ubah Status</button>
          </div>
        </td>
        <td style="font-size: 11px; color: var(--sub);">${d.STATUS_BAKB || '-'}</td>
        <td style="text-align: right; white-space: nowrap;">
          <button class="btn-sm-edit" onclick="openEditDetilModal(${idx})" title="Edit seluruh 37 kolom">Edit 37 Kolom</button>
          <button class="btn-sm-edit" style="color: var(--rose); margin-left: 4px;" onclick="deleteDetil(${idx})" title="Hapus termin">&times;</button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }

  // =========================================================================
  // REKAP MODAL (36 KOLOM) HANDLERS
  // =========================================================================
  function openAddRekapModal() {
    editingRekapId = null;
    document.getElementById('rekapModalTitle').textContent = 'Tambah Data Rekap CR / JR Baru (36 Kolom)';

    // Auto next ID
    const maxId = (appData.rekap || []).reduce((max, r) => Math.max(max, Number(r.ID_TRANSAKSI) || 0), 0);
    document.getElementById('in_rek_id').value = maxId + 1;
    document.getElementById('in_rek_layanan').value = 'AP2T';
    document.getElementById('in_rek_tipe').value = 'CR';
    document.getElementById('in_rek_uraian').value = '';

    // Clear others
    ['tgl_bpo', 'no_bpo', 'path_bpo', 'tgl_mdg', 'no_mdg', 'path_mdg', 'tgl_sti', 'no_sti', 'path_sti',
     'tgl_bak', 'no_bak', 'path_bak', 'status_uat', 'path_baut', 'status_deploy', 'path_deploy',
     'status_wbs', 'path_wbs', 'nilai_wbs', 'status_bakb', 'path_bakb', 'nilai_tahap', 'nilai_full',
     'status_mom', 'path_mom', 'tgl_created', 'created_by', 'tgl_koreksi', 'koreksi_by', 'tgl_batal', 'batal_by'
    ].forEach(f => {
      const el = document.getElementById('in_rek_' + f);
      if (el) el.value = '';
    });

    document.getElementById('in_rek_status_wbs').value = 'YA';
    document.getElementById('in_rek_path_wbs').value = 'ADA';
    document.getElementById('in_rek_angsuran').value = 'TIDAK';

    document.getElementById('modalRekap').classList.add('open');
  }

  function openEditRekapModal(id) {
    editingRekapId = id;
    const r = (appData.rekap || []).find(it => it.ID_TRANSAKSI === id);
    if (!r) return;

    document.getElementById('rekapModalTitle').textContent = 'Edit Data Rekap CR / JR #' + id + ' (36 Kolom)';

    document.getElementById('in_rek_id').value = r.ID_TRANSAKSI || '';
    document.getElementById('in_rek_layanan').value = r.LAYANAN || 'AP2T';
    document.getElementById('in_rek_tipe').value = r.TIPE || 'CR';
    document.getElementById('in_rek_uraian').value = r.URAIAN || '';

    // Rantai surat
    document.getElementById('in_rek_tgl_bpo').value = r.TGLSURAT_BPO || '';
    document.getElementById('in_rek_no_bpo').value = r.NOSURAT_BPO || '';
    document.getElementById('in_rek_path_bpo').value = r.PATHFOLDER_BPO || '';
    document.getElementById('in_rek_tgl_mdg').value = r.TGLSURAT_MDG || '';
    document.getElementById('in_rek_no_mdg').value = r.NOSURAT_MDG || '';
    document.getElementById('in_rek_path_mdg').value = r.PATHFOLDER_MDG || '';
    document.getElementById('in_rek_tgl_sti').value = r.TGLSURAT_STI || '';
    document.getElementById('in_rek_no_sti').value = r.NOSURAT_STI || '';
    document.getElementById('in_rek_path_sti').value = r.PATHFOLDER_STI || '';

    // BA
    document.getElementById('in_rek_tgl_bak').value = r.TGL_BAK || '';
    document.getElementById('in_rek_no_bak').value = r.NOSURAT_BAK || '';
    document.getElementById('in_rek_path_bak').value = r.PATHFOLDER_BAK || '';
    document.getElementById('in_rek_status_uat').value = r.STATUS_UAT || '';
    document.getElementById('in_rek_path_baut').value = r.PATHFOLDER_BAUT || '';
    document.getElementById('in_rek_status_deploy').value = r.STATUS_DEPLOY || '';
    document.getElementById('in_rek_path_deploy').value = r.PATHFOLDER_DEPLOY || '';
    document.getElementById('in_rek_status_mom').value = r.STATUS_MOM || '';
    document.getElementById('in_rek_path_mom').value = r.PATHFOLDER_MOM || '';

    // WBS & Finansial
    document.getElementById('in_rek_status_wbs').value = r.STATUS_WBS || 'YA';
    document.getElementById('in_rek_path_wbs').value = r.PATHFOLDER_WBS || 'ADA';
    document.getElementById('in_rek_nilai_wbs').value = r.NILAI_WBS || '';
    document.getElementById('in_rek_angsuran').value = r.ANGSURAN || 'TIDAK';
    document.getElementById('in_rek_status_bakb').value = r.STATUS_BAKB || '';
    document.getElementById('in_rek_path_bakb').value = r.PATHFOLDER_BAKB || '';
    document.getElementById('in_rek_nilai_tahap').value = r.NILAI_TAHAP || '';
    document.getElementById('in_rek_nilai_full').value = r.NILAI_FULL || '';

    // Audit
    document.getElementById('in_rek_tgl_created').value = r.TGL_CREATED || '';
    document.getElementById('in_rek_created_by').value = r.CREATED_BY || '';
    document.getElementById('in_rek_tgl_koreksi').value = r.TGL_KOREKSI || '';
    document.getElementById('in_rek_koreksi_by').value = r.KOREKSI_BY || '';
    document.getElementById('in_rek_tgl_batal').value = r.TGL_BATAL || '';
    document.getElementById('in_rek_batal_by').value = r.BATAL_BY || '';

    document.getElementById('modalRekap').classList.add('open');
  }

  function closeRekapModal() {
    document.getElementById('modalRekap').classList.remove('open');
  }

  function saveRekapForm() {
    const id = parseInt(document.getElementById('in_rek_id').value);
    if (!id) {
      alert('ID Transaksi harus diisi!');
      return;
    }

    const rekapObj = {
      "ID_TRANSAKSI": id,
      "LAYANAN": document.getElementById('in_rek_layanan').value || 'AP2T',
      "TIPE": document.getElementById('in_rek_tipe').value || 'CR',
      "URAIAN": document.getElementById('in_rek_uraian').value || '',
      "TGLSURAT_BPO": document.getElementById('in_rek_tgl_bpo').value || '',
      "NOSURAT_BPO": document.getElementById('in_rek_no_bpo').value || '',
      "PATHFOLDER_BPO": document.getElementById('in_rek_path_bpo').value || '',
      "TGLSURAT_MDG": document.getElementById('in_rek_tgl_mdg').value || '',
      "NOSURAT_MDG": document.getElementById('in_rek_no_mdg').value || '',
      "PATHFOLDER_MDG": document.getElementById('in_rek_path_mdg').value || '',
      "TGLSURAT_STI": document.getElementById('in_rek_tgl_sti').value || '',
      "NOSURAT_STI": document.getElementById('in_rek_no_sti').value || '',
      "PATHFOLDER_STI": document.getElementById('in_rek_path_sti').value || '',
      "TGL_BAK": document.getElementById('in_rek_tgl_bak').value || '',
      "NOSURAT_BAK": document.getElementById('in_rek_no_bak').value || '',
      "PATHFOLDER_BAK": document.getElementById('in_rek_path_bak').value || '',
      "STATUS_UAT": document.getElementById('in_rek_status_uat').value || '',
      "PATHFOLDER_BAUT": document.getElementById('in_rek_path_baut').value || '',
      "STATUS_DEPLOY": document.getElementById('in_rek_status_deploy').value || '',
      "PATHFOLDER_DEPLOY": document.getElementById('in_rek_path_deploy').value || '',
      "STATUS_WBS": document.getElementById('in_rek_status_wbs').value || '',
      "PATHFOLDER_WBS": document.getElementById('in_rek_path_wbs').value || '',
      "NILAI_WBS": parseFloat(document.getElementById('in_rek_nilai_wbs').value) || 0,
      "ANGSURAN": document.getElementById('in_rek_angsuran').value || 'TIDAK',
      "STATUS_BAKB": document.getElementById('in_rek_status_bakb').value || '',
      "PATHFOLDER_BAKB": document.getElementById('in_rek_path_bakb').value || '',
      "NILAI_TAHAP": parseFloat(document.getElementById('in_rek_nilai_tahap').value) || 0,
      "NILAI_FULL": parseFloat(document.getElementById('in_rek_nilai_full').value) || 0,
      "STATUS_MOM": document.getElementById('in_rek_status_mom').value || '',
      "PATHFOLDER_MOM": document.getElementById('in_rek_path_mom').value || '',
      "TGL_CREATED": document.getElementById('in_rek_tgl_created').value || '',
      "CREATED_BY": document.getElementById('in_rek_created_by').value || '',
      "TGL_KOREKSI": document.getElementById('in_rek_tgl_koreksi').value || '',
      "KOREKSI_BY": document.getElementById('in_rek_koreksi_by').value || '',
      "TGL_BATAL": document.getElementById('in_rek_tgl_batal').value || '',
      "BATAL_BY": document.getElementById('in_rek_batal_by').value || ''
    };

    if (!appData.rekap) appData.rekap = [];

    const existingIdx = appData.rekap.findIndex(it => it.ID_TRANSAKSI === id);
    if (existingIdx >= 0) {
      appData.rekap[existingIdx] = rekapObj;
      showToast('Data Rekap #' + id + ' berhasil diperbarui!');
    } else {
      appData.rekap.push(rekapObj);
      showToast('Data Rekap #' + id + ' baru berhasil ditambahkan!');
    }

    saveAppData(appData);
    closeRekapModal();
    renderOpportunityView();
  }

  function deleteRekap(id) {
    if (!confirm('Hapus transaksi Rekap #' + id + '?')) return;
    appData.rekap = (appData.rekap || []).filter(it => it.ID_TRANSAKSI !== id);
    // Also remove related detil angsuran
    appData.detilAngsuran = (appData.detilAngsuran || []).filter(it => it.ID_TRANSAKSI !== id);
    saveAppData(appData);
    showToast('Data Rekap #' + id + ' berhasil dihapus!');
    renderOpportunityView();
  }

  // =========================================================================
  // DETIL ANGSURAN MODAL (37 KOLOM) HANDLERS
  // =========================================================================
  function populateDetilParentDropdown(selectedId) {
    const sel = document.getElementById('in_det_id');
    if (!sel) return;
    sel.innerHTML = '';

    const rekapList = appData.rekap || [];
    rekapList.forEach(r => {
      const opt = document.createElement('option');
      opt.value = r.ID_TRANSAKSI;
      opt.textContent = '#' + r.ID_TRANSAKSI + ' - ' + (r.URAIAN || '').slice(0, 50) + (r.ANGSURAN === 'YA' ? ' [Termin: YA]' : '');
      if (r.ID_TRANSAKSI === selectedId) opt.selected = true;
      sel.appendChild(opt);
    });

    onDetilParentChange();
  }

  function onDetilParentChange() {
    const parentId = parseInt(document.getElementById('in_det_id').value);
    const r = (appData.rekap || []).find(it => it.ID_TRANSAKSI === parentId);
    if (r) {
      document.getElementById('in_det_layanan').value = r.LAYANAN || 'AP2T';
      document.getElementById('in_det_tipe').value = r.TIPE || 'CR';
      document.getElementById('in_det_uraian').value = r.URAIAN || '';
      if (!editingDetilIdx && editingDetilIdx !== 0) {
        document.getElementById('in_det_nilai_full').value = r.NILAI_FULL || '';
      }
    }
  }

  function openAddDetilModal(defaultParentId) {
    editingDetilIdx = null;
    document.getElementById('detilModalTitle').textContent = 'Tambah Detil Angsuran Baru (37 Kolom)';

    const rekapList = appData.rekap || [];
    const firstId = defaultParentId || (rekapList.find(r => r.ANGSURAN === 'YA')?.ID_TRANSAKSI || (rekapList[0]?.ID_TRANSAKSI || 1));

    populateDetilParentDropdown(firstId);

    // Count existing stages for this parent
    const existStages = (appData.detilAngsuran || []).filter(d => d.ID_TRANSAKSI === firstId);
    document.getElementById('in_det_tahap').value = existStages.length + 1;
    document.getElementById('in_det_nilai_tahap').value = '';
    document.getElementById('in_det_status_lunas').value = 'BELUM';

    // Clear other fields
    ['tgl_bpo', 'no_bpo', 'path_bpo', 'tgl_mdg', 'no_mdg', 'path_mdg', 'tgl_sti', 'no_sti', 'path_sti',
     'tgl_bak', 'no_bak', 'path_bak', 'status_uat', 'path_baut', 'status_deploy', 'path_deploy',
     'status_wbs', 'path_wbs', 'status_bakb', 'path_bakb', 'status_mom', 'path_mom',
     'tgl_created', 'created_by', 'tgl_koreksi', 'koreksi_by', 'tgl_batal', 'batal_by'
    ].forEach(f => {
      const el = document.getElementById('in_det_' + f);
      if (el) el.value = '';
    });

    document.getElementById('modalDetil').classList.add('open');
  }

  function openEditDetilModal(idx) {
    editingDetilIdx = idx;
    const d = (appData.detilAngsuran || [])[idx];
    if (!d) return;

    document.getElementById('detilModalTitle').textContent = 'Edit Detil Angsuran Tahap ' + (d.TAHAP || '') + ' #' + (d.ID_TRANSAKSI || '') + ' (37 Kolom)';

    populateDetilParentDropdown(d.ID_TRANSAKSI);

    document.getElementById('in_det_tahap').value = d.TAHAP || 1;
    document.getElementById('in_det_nilai_tahap').value = d.NILAI_TAHAP || '';
    document.getElementById('in_det_nilai_full').value = d.NILAI_FULL || '';
    document.getElementById('in_det_status_lunas').value = d.STATUS_LUNAS || 'BELUM';

    // Rantai surat
    document.getElementById('in_det_tgl_bpo').value = d.TGLSURAT_BPO || '';
    document.getElementById('in_det_no_bpo').value = d.NOSURAT_BPO || '';
    document.getElementById('in_det_path_bpo').value = d.PATHFOLDER_BPO || '';
    document.getElementById('in_det_tgl_mdg').value = d.TGLSURAT_MDG || '';
    document.getElementById('in_det_no_mdg').value = d.NOSURAT_MDG || '';
    document.getElementById('in_det_path_mdg').value = d.PATHFOLDER_MDG || '';
    document.getElementById('in_det_tgl_sti').value = d.TGLSURAT_STI || '';
    document.getElementById('in_det_no_sti').value = d.NOSURAT_STI || '';
    document.getElementById('in_det_path_sti').value = d.PATHFOLDER_STI || '';

    // BA
    document.getElementById('in_det_tgl_bak').value = d.TGL_BAK || '';
    document.getElementById('in_det_no_bak').value = d.NOSURAT_BAK || '';
    document.getElementById('in_det_path_bak').value = d.PATHFOLDER_BAK || '';
    document.getElementById('in_det_status_uat').value = d.STATUS_UAT || '';
    document.getElementById('in_det_path_baut').value = d.PATHFOLDER_BAUT || '';
    document.getElementById('in_det_status_deploy').value = d.STATUS_DEPLOY || '';
    document.getElementById('in_det_path_deploy').value = d.PATHFOLDER_DEPLOY || '';
    document.getElementById('in_det_status_wbs').value = d.STATUS_WBS || '';
    document.getElementById('in_det_path_wbs').value = d.PATHFOLDER_WBS || '';
    document.getElementById('in_det_status_bakb').value = d.STATUS_BAKB || '';
    document.getElementById('in_det_path_bakb').value = d.PATHFOLDER_BAKB || '';
    document.getElementById('in_det_status_mom').value = d.STATUS_MOM || '';
    document.getElementById('in_det_path_mom').value = d.PATHFOLDER_MOM || '';

    // Audit
    document.getElementById('in_det_tgl_created').value = d.TGL_CREATED || '';
    document.getElementById('in_det_created_by').value = d.CREATED_BY || '';
    document.getElementById('in_det_tgl_koreksi').value = d.TGL_KOREKSI || '';
    document.getElementById('in_det_koreksi_by').value = d.KOREKSI_BY || '';
    document.getElementById('in_det_tgl_batal').value = d.TGL_BATAL || '';
    document.getElementById('in_det_batal_by').value = d.BATAL_BY || '';

    document.getElementById('modalDetil').classList.add('open');
  }

  function closeDetilModal() {
    document.getElementById('modalDetil').classList.remove('open');
  }

  function saveDetilForm() {
    const parentId = parseInt(document.getElementById('in_det_id').value);
    const tahap = parseInt(document.getElementById('in_det_tahap').value) || 1;

    const detilObj = {
      "ID_TRANSAKSI": parentId,
      "LAYANAN": document.getElementById('in_det_layanan').value || 'AP2T',
      "TIPE": document.getElementById('in_det_tipe').value || 'CR',
      "URAIAN": document.getElementById('in_det_uraian').value || '',
      "TGLSURAT_BPO": document.getElementById('in_det_tgl_bpo').value || '',
      "NOSURAT_BPO": document.getElementById('in_det_no_bpo').value || '',
      "PATHFOLDER_BPO": document.getElementById('in_det_path_bpo').value || '',
      "TGLSURAT_MDG": document.getElementById('in_det_tgl_mdg').value || '',
      "NOSURAT_MDG": document.getElementById('in_det_no_mdg').value || '',
      "PATHFOLDER_MDG": document.getElementById('in_det_path_mdg').value || '',
      "TGLSURAT_STI": document.getElementById('in_det_tgl_sti').value || '',
      "NOSURAT_STI": document.getElementById('in_det_no_sti').value || '',
      "PATHFOLDER_STI": document.getElementById('in_det_path_sti').value || '',
      "TGL_BAK": document.getElementById('in_det_tgl_bak').value || '',
      "NOSURAT_BAK": document.getElementById('in_det_no_bak').value || '',
      "PATHFOLDER_BAK": document.getElementById('in_det_path_bak').value || '',
      "STATUS_UAT": document.getElementById('in_det_status_uat').value || '',
      "PATHFOLDER_BAUT": document.getElementById('in_det_path_baut').value || '',
      "STATUS_DEPLOY": document.getElementById('in_det_status_deploy').value || '',
      "PATHFOLDER_DEPLOY": document.getElementById('in_det_path_deploy').value || '',
      "STATUS_WBS": document.getElementById('in_det_status_wbs').value || '',
      "PATHFOLDER_WBS": document.getElementById('in_det_path_wbs').value || '',
      "ANGSURAN": "YA",
      "STATUS_BAKB": document.getElementById('in_det_status_bakb').value || '',
      "PATHFOLDER_BAKB": document.getElementById('in_det_path_bakb').value || '',
      "TAHAP": tahap,
      "NILAI_TAHAP": parseFloat(document.getElementById('in_det_nilai_tahap').value) || 0,
      "NILAI_FULL": parseFloat(document.getElementById('in_det_nilai_full').value) || 0,
      "STATUS_LUNAS": document.getElementById('in_det_status_lunas').value || 'BELUM',
      "STATUS_MOM": document.getElementById('in_det_status_mom').value || '',
      "PATHFOLDER_MOM": document.getElementById('in_det_path_mom').value || '',
      "TGL_CREATED": document.getElementById('in_det_tgl_created').value || '',
      "CREATED_BY": document.getElementById('in_det_created_by').value || '',
      "TGL_KOREKSI": document.getElementById('in_det_tgl_koreksi').value || '',
      "KOREKSI_BY": document.getElementById('in_det_koreksi_by').value || '',
      "TGL_BATAL": document.getElementById('in_det_tgl_batal').value || '',
      "BATAL_BY": document.getElementById('in_det_batal_by').value || ''
    };

    if (!appData.detilAngsuran) appData.detilAngsuran = [];

    if (editingDetilIdx !== null && editingDetilIdx >= 0) {
      appData.detilAngsuran[editingDetilIdx] = detilObj;
      showToast('Detil Angsuran Tahap ' + tahap + ' berhasil diperbarui!');
    } else {
      appData.detilAngsuran.push(detilObj);
      showToast('Detil Angsuran Tahap ' + tahap + ' berhasil ditambahkan!');
    }

    saveAppData(appData);
    closeDetilModal();
    renderOpportunityView();
  }

  function toggleDetilLunas(idx) {
    const d = (appData.detilAngsuran || [])[idx];
    if (!d) return;
    const cur = (d.STATUS_LUNAS || '').toUpperCase();
    d.STATUS_LUNAS = (cur === 'LUNAS') ? 'BELUM' : 'LUNAS';
    saveAppData(appData);
    showToast('Status Tahap ' + (d.TAHAP || '') + ' diubah menjadi: ' + d.STATUS_LUNAS);
    renderOpportunityView();
  }

  function deleteDetil(idx) {
    if (!confirm('Hapus baris angsuran ini?')) return;
    appData.detilAngsuran.splice(idx, 1);
    saveAppData(appData);
    showToast('Baris angsuran berhasil dihapus!');
    renderOpportunityView();
  }

  function resetSamplingRev3Data() {
    if (confirm('Kembalikan data Opportunity & PO ke standar Sampling_data_rev3.xlsx?')) {
      appData.rekap = JSON.parse(JSON.stringify(DEFAULT_REKAP));
      appData.detilAngsuran = JSON.parse(JSON.stringify(DEFAULT_DETIL));
      saveAppData(appData);
      showToast('Data berhasil dikembalikan ke standar Sampling Rev3!');
      renderOpportunityView();
    }
  }

  function syncOpportunityToDashboard(sumFull, sumCashIn, sumBelum) {
    if (!appData.annual) appData.annual = {};
    appData.annual.opp = parseFloat(((sumFull || 0) / 1e9).toFixed(2));
    appData.annual.cash = parseFloat(((sumCashIn || 0) / 1e9).toFixed(2));
    appData.annual.wait = parseFloat(((sumBelum || 0) / 1e9).toFixed(2));
    appData.annual.real = parseFloat((((sumCashIn || 0) + (sumBelum || 0)) / 1e9).toFixed(2));
    appData.annual.risk = 0.00;
    saveAppData(appData);
  }

  // =========================================================================
  // EXCEL UPLOAD PARSER (CLIENT-SIDE WITH SHEETJS)
  // =========================================================================
  function handleExcelOppUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (typeof XLSX === 'undefined') {
      alert('SheetJS belum siap. Pastikan koneksi internet aktif untuk memuat parser Excel.');
      return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        showToast('Membaca sheet dari ' + file.name + '...');

        // 1. Read Sheet REKAP
        let sheetRekapName = workbook.SheetNames.find(n => n.toUpperCase() === 'REKAP') || workbook.SheetNames[0];
        let wsRekap = workbook.Sheets[sheetRekapName];
        let rawRekap = XLSX.utils.sheet_to_json(wsRekap, { defval: '' });

        // Filter out summary/formula rows
        const cleanedRekap = [];
        rawRekap.forEach(row => {
          const id = parseInt(row['ID_TRANSAKSI']);
          if (id && !isNaN(id)) {
            cleanedRekap.push(row);
          }
        });

        if (cleanedRekap.length) {
          appData.rekap = cleanedRekap;
        }

        // 2. Read Sheet DETIl_ANGSURAN
        let sheetDetilName = workbook.SheetNames.find(n => n.toUpperCase().includes('ANGSURAN') || n.toUpperCase().includes('DETIL'));
        if (sheetDetilName) {
          let wsDetil = workbook.Sheets[sheetDetilName];
          let rawDetil = XLSX.utils.sheet_to_json(wsDetil, { defval: '' });

          const cleanedDetil = [];
          let lastId = 6;
          let lastLayanan = 'AP2T';
          let lastTipe = 'CR';
          let lastUraian = '';

          rawDetil.forEach(row => {
            if (row['ID_TRANSAKSI']) {
              lastId = parseInt(row['ID_TRANSAKSI']);
              lastLayanan = row['LAYANAN'] || lastLayanan;
              lastTipe = row['TIPE'] || lastTipe;
              lastUraian = row['URAIAN'] || lastUraian;
            } else {
              row['ID_TRANSAKSI'] = lastId;
              row['LAYANAN'] = lastLayanan;
              row['TIPE'] = lastTipe;
              row['URAIAN'] = lastUraian;
            }

            if (row['TAHAP'] || row['NILAI_TAHAP']) {
              cleanedDetil.push(row);
            }
          });

          if (cleanedDetil.length) {
            appData.detilAngsuran = cleanedDetil;
          }
        }

        saveAppData(appData);
        renderOpportunityView();
        showToast('Berhasil mengimpor ' + (cleanedRekap.length || 0) + ' Rekap dan ' + ((appData.detilAngsuran || []).length) + ' Angsuran dari ' + file.name + '!');
      } catch (err) {
        alert('Gagal membaca file Excel: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = '';
  }



