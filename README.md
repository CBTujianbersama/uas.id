# uas.id
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>CBT-PRO | Sistem Ujian Berbasis Komputer</title>
    
    <!-- Tailwind CSS CDN (Render Dinamis & Konsisten di GitHub Pages) -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'sans-serif'],
                    }
                }
            }
        };
    </script>

    <!-- Google Fonts: Inter -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    
    <!-- Lucide Icons CDN -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- SheetJS (XLSX) CDN untuk Import/Export Excel -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <!-- Mammoth.js CDN untuk membaca file Word (.docx) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
    <!-- PDF.js CDN untuk membaca file PDF (.pdf) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>

    <script>
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
    </script>

    <style>
        /* Reset & Stabilizer Layout agar tidak saling timpa saat di-hosting di GitHub Pages */
        *, *::before, *::after {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
            width: 100%;
            min-height: 100vh;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            overflow-x: hidden;
            -webkit-font-smoothing: antialiased;
        }

        img, svg, video, canvas {
            max-width: 100%;
            display: block;
        }

        table {
            border-collapse: collapse;
            width: 100%;
        }

        /* Pastikan elemen tersembunyi benar-benar tersembunyi tanpa tumpang tindih */
        .hidden {
            display: none !important;
        }

        /* Mencegah seleksi teks saat siswa mengerjakan ujian */
        .no-select {
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
        }

        /* Animasi latar belakang layar terkunci anti-curang */
        @keyframes lockPulse {
            0%, 100% { background-color: rgba(127, 29, 29, 0.98); }
            50% { background-color: rgba(185, 28, 28, 0.98); }
        }
        .lock-overlay-bg {
            animation: lockPulse 2s infinite;
        }

        /* Stabilizer Sticky Header & Sidebar */
        #mainHeader {
            position: sticky;
            top: 0;
            z-index: 40;
            background-color: #ffffff;
        }

        /* Aturan Cetak (Print A4 2 Kolom) Khusus Kartu Login Peserta */
        @media print {
            body * {
                visibility: hidden;
            }
            #printCardsContainer, #printCardsContainer * {
                visibility: visible;
            }
            #printCardsContainer {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                display: grid !important;
                grid-template-columns: repeat(2, 1fr) !important;
                gap: 12px !important;
                padding: 8px !important;
                background: #ffffff !important;
            }
            .no-print {
                display: none !important;
            }
            .login-card-item {
                break-inside: avoid;
                page-break-inside: avoid;
                border: 1.5px solid #0f172a !important;
                box-shadow: none !important;
            }
        }

        /* Scrollbar halus */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col">

    <!-- ===================================================================== -->
    <!-- 1. SISTEM TOAST & DIALOG KONFIRMASI (PENGGANTI ALERT & CONFIRM)       -->
    <!-- ===================================================================== -->
    <div id="toastContainer" class="fixed top-4 right-4 z-[9999] flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

    <div id="confirmModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9990] items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center gap-3 mb-3">
                <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                    <i data-lucide="alert-triangle" class="w-5 h-5"></i>
                </div>
                <h3 id="confirmTitle" class="text-lg font-bold text-slate-900">Konfirmasi</h3>
            </div>
            <p id="confirmMessage" class="text-sm text-slate-600 mb-6 leading-relaxed">Apakah Anda yakin?</p>
            <div class="flex justify-end gap-2.5">
                <button id="confirmCancelBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 transition">Batal</button>
                <button id="confirmOkBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-white bg-blue-600 hover:bg-blue-700 transition shadow-sm">Ya, Lanjutkan</button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 2. OVERLAY LAYAR TERKUNCI ANTI-CURANG + SIRINE ALARM OTOMATIS         -->
    <!-- ===================================================================== -->
    <div id="antiCheatLockOverlay" class="fixed inset-0 lock-overlay-bg z-[99999] flex-col items-center justify-center p-4 sm:p-6 text-white text-center select-none hidden">
        <div class="max-w-lg w-full bg-red-950/95 border-2 border-red-400 rounded-3xl p-6 sm:p-8 shadow-2xl">
            <div class="w-20 h-20 rounded-full bg-red-600/30 border-2 border-red-400 flex items-center justify-center mx-auto mb-4 animate-bounce">
                <i data-lucide="siren" class="w-10 h-10 text-amber-300"></i>
            </div>
            <div class="flex items-center justify-center gap-2 mb-3">
                <span class="inline-block px-3 py-1 rounded-full bg-red-600 text-white text-xs font-extrabold uppercase tracking-wider">
                    SIRINE ALARM AKTIF • PELANGGARAN TERDETEKSI
                </span>
            </div>
            <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight mb-2">LAYAR UJIAN TERKUNCI!</h1>
            <p id="lockReasonText" class="text-red-200 text-xs sm:text-sm mb-5 leading-relaxed">Layar ujian kehilangan fokus! Dilarang membuka jendela lain, split-screen, atau berpindah aplikasi.</p>

            <div class="bg-red-900/70 border border-red-700 rounded-2xl p-4 mb-5 text-left text-xs space-y-2">
                <div class="flex justify-between">
                    <span class="text-red-300">Nama Peserta:</span>
                    <span id="lockStudentName" class="font-bold text-white">-</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">NISN / Kelas:</span>
                    <span id="lockStudentMeta" class="font-semibold text-white">-</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">Mata Pelajaran Ujian:</span>
                    <span id="lockExamSubject" class="font-semibold text-amber-200">-</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">Total Pelanggaran Tercatat:</span>
                    <span id="lockViolationBadge" class="font-extrabold text-amber-300">1 Kali</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-red-300">Status Alarm &amp; Log:</span>
                    <span class="font-semibold text-emerald-300">Sirine Menyala • Terkirim ke Pengawas &amp; Admin</span>
                </div>
            </div>

            <div class="space-y-3">
                <label class="block text-xs font-bold uppercase tracking-wider text-red-200">
                    Masukkan Kode PIN Pengawas untuk Mematikan Sirine &amp; Membuka Kunci
                </label>
                <input type="password" id="unlockPinField" placeholder="Masukkan Kode PIN Pengawas" class="w-full px-4 py-3.5 rounded-xl bg-white text-slate-900 font-extrabold text-center text-lg tracking-widest focus:outline-none focus:ring-4 focus:ring-amber-400" onkeypress="if(event.key === 'Enter') verifyPinUnlock()">
                <button onclick="verifyPinUnlock()" class="w-full py-3.5 px-6 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold text-sm transition shadow-lg flex items-center justify-center gap-2">
                    <i data-lucide="unlock" class="w-4 h-4"></i> MATIKAN SIRINE &amp; BUKA KUNCI
                </button>
            </div>
            <p class="mt-4 text-[11px] text-red-300/90">
                Sirine peringatan akan terus berbunyi sampai Pengawas Ruangan memverifikasi PIN atau membuka kunci jarak jauh.
            </p>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 3. HEADER / NAVBAR UTAMA                                              -->
    <!-- ===================================================================== -->
    <header id="mainHeader" class="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-sm no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between min-h-[4rem] py-2 gap-3">
                <div class="flex items-center gap-3 cursor-pointer min-w-0" onclick="handleHeaderBrandClick()">
                    <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-700 text-white flex items-center justify-center shadow-md shrink-0">
                        <i data-lucide="shield-check" class="w-6 h-6"></i>
                    </div>
                    <div class="min-w-0">
                        <div class="flex items-center gap-2">
                            <h1 id="headerSchoolTitle" class="font-extrabold text-sm sm:text-base text-slate-900 truncate">CBT-PRO | Sistem Ujian Berbasis Komputer</h1>
                            <span class="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold shrink-0">
                                <span class="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping"></span> LIVE SYNC
                            </span>
                        </div>
                        <p id="headerSchoolSubtitle" class="text-[11px] text-slate-500 truncate">Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027</p>
                    </div>
                </div>

                <!-- Tombol Aksi Header -->
                <div class="flex items-center gap-2 shrink-0">
                    <div id="headerGuestButtons" class="flex items-center gap-1.5 sm:gap-2">
                        <button onclick="openLoginModal('siswa')" class="px-3 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                            <i data-lucide="graduation-cap" class="w-4 h-4"></i>
                            <span>Login Siswa</span>
                        </button>
                        <button onclick="openLoginModal('pengawas')" class="px-3 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition hidden sm:flex items-center gap-1.5">
                            <i data-lucide="user-check" class="w-4 h-4"></i>
                            <span>Pengawas</span>
                        </button>
                        <button onclick="openLoginModal('admin')" class="px-3 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition hidden sm:flex items-center gap-1.5">
                            <i data-lucide="settings" class="w-4 h-4"></i>
                            <span>Admin</span>
                        </button>
                    </div>

                    <div id="headerAuthControls" class="hidden items-center gap-2">
                        <span id="headerUserBadge" class="hidden sm:inline-block px-3 py-1.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">
                            -
                        </span>
                        <button id="headerBackDashBtn" onclick="returnToRoleDashboard()" class="hidden px-3 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition">
                            Dashboard Saya
                        </button>
                        <button onclick="navigateTo('portal')" class="px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition">
                            Portal
                        </button>
                        <button onclick="logoutSession()" class="px-3 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1">
                            <i data-lucide="log-out" class="w-3.5 h-3.5"></i> Keluar
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <!-- ===================================================================== -->
    <!-- 4. KONTEN UTAMA (PORTAL, DASHBOARD SISWA, PENGAWAS, ADMIN, UJIAN)     -->
    <!-- ===================================================================== -->
    <main class="flex-grow">

        <!-- ================================================================= -->
        <!-- VIEW A: HALAMAN PORTAL UTAMA (LANDING PORTAL)                     -->
        <!-- ================================================================= -->
        <section id="view-portal" class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-10">
            <div class="bg-gradient-to-br from-slate-900 via-blue-950 to-indigo-950 rounded-3xl p-6 sm:p-10 text-white shadow-xl relative overflow-hidden">
                <div class="absolute -top-24 -right-24 w-80 h-80 bg-blue-500/15 rounded-full blur-3xl pointer-events-none"></div>
                <div class="relative z-10 grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
                    <div class="lg:col-span-7 space-y-5">
                        <div class="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-blue-500/20 border border-blue-400/30 text-blue-200 text-xs font-semibold">
                            <i data-lucide="shield-alert" class="w-4 h-4 text-amber-300"></i>
                            <span>Sistem Ujian Multi-Mapel • Token Sesi • Sirine Anti-Curang</span>
                        </div>
                        <h2 class="text-2xl sm:text-4xl font-extrabold tracking-tight leading-tight">
                            Selamat Datang di Portal <span class="text-blue-400" id="heroExamTitle">Ujian Akhir Semester (UAS) Berbasis Komputer</span>
                        </h2>
                        <p class="text-slate-300 text-sm sm:text-base leading-relaxed">
                            Silakan login ke <strong>Dashboard Siswa</strong> menggunakan <strong>NISN / Username</strong> dan <strong>Password</strong> Kartu Login untuk melihat profil, jadwal penugasan mata pelajaran, serta memasukkan <strong>Token Ujian</strong> saat soal sudah siap dikerjakan.
                        </p>

                        <!-- Tombol Login 3 Peran -->
                        <div class="flex flex-wrap gap-3 pt-1">
                            <button onclick="openLoginModal('siswa')" class="px-5 py-3 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs sm:text-sm transition flex items-center gap-2 shadow-lg shadow-blue-950/60">
                                <i data-lucide="graduation-cap" class="w-4 h-4"></i> Login Dashboard Siswa
                            </button>
                            <button onclick="openLoginModal('pengawas')" class="px-4 py-3 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="eye" class="w-4 h-4"></i> Portal Pengawas
                            </button>
                            <button onclick="openLoginModal('admin')" class="px-4 py-3 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/20 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="sliders" class="w-4 h-4"></i> Administrator
                            </button>
                        </div>

                        <!-- Tombol Salin Link Akses -->
                        <div class="pt-3 border-t border-white/10 flex flex-wrap items-center gap-2 text-xs">
                            <span class="text-slate-400 font-medium mr-1">Bagikan Link Akses:</span>
                            <button onclick="copyPortalLink('siswa')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-blue-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Portal Siswa
                            </button>
                            <button onclick="copyPortalLink('pengawas')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-amber-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Pengawas
                            </button>
                            <button onclick="copyPortalLink('admin')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-indigo-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Administrator
                            </button>
                        </div>
                    </div>

                    <!-- Kartu Motivasi & Statistik Live -->
                    <div class="lg:col-span-5">
                        <div class="bg-white/10 backdrop-blur-md border border-white/20 rounded-3xl p-6 space-y-4">
                            <div class="flex items-center justify-between">
                                <span class="inline-flex items-center gap-1.5 text-xs font-bold uppercase tracking-wider text-amber-300">
                                    <i data-lucide="sparkles" class="w-4 h-4"></i> Motivasi Belajar &amp; Kejujuran
                                </span>
                                <button onclick="nextMotivationQuote()" class="px-2.5 py-1 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-white transition flex items-center gap-1">
                                    <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i> Acak
                                </button>
                            </div>
                            <blockquote class="space-y-3">
                                <p id="motivationQuoteText" class="text-sm sm:text-base font-medium italic text-white leading-relaxed">
                                    "Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."
                                </p>
                                <footer id="motivationQuoteAuthor" class="text-xs font-bold text-blue-300">
                                    — Pesan Akademik Sekolah
                                </footer>
                            </blockquote>
                            <div class="pt-3 border-t border-white/10 grid grid-cols-3 gap-2 text-center">
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSoal" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Paket Ujian</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSiswa" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Peserta Aktif</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatDurasi" class="text-lg font-extrabold text-amber-300">0</div>
                                    <div class="text-[10px] text-slate-400">Terjadwal</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- BAGIAN PANDUAN SISTEM UJIAN CBT (SISWA, PENGAWAS, ADMIN) -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Tata Tertib &amp; Alur Pelaksanaan</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Alur Sistem Ujian CBT Terpadu</h2>
                    </div>
                    <p class="text-xs text-slate-500">Sinkronisasi real-time antara Admin, Pengawas, dan Siswa</p>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <!-- Panduan Siswa -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                                <i data-lucide="user-check" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">1. Alur Peserta / Siswa</h3>
                                <span class="text-xs text-blue-600 font-semibold">Dashboard &amp; Pemilihan Soal</span>
                            </div>
                        </div>
                        <ol class="space-y-2.5 text-xs text-slate-600 list-decimal pl-4 leading-relaxed">
                            <li><strong>Login ke Dashboard Siswa:</strong> Masuk dengan NISN/Username &amp; Password Kartu Login (tanpa langsung masuk soal).</li>
                            <li><strong>Cek Profil &amp; Jadwal Penugasan:</strong> Lihat mata pelajaran yang sudah di-assign oleh Admin atau masih menunggu jadwal.</li>
                            <li><strong>Masukkan Token saat Siap:</strong> Klik ujian yang berstatus <em>Siap Dikerjakan</em> dan input <strong>Token Aktif</strong> dari Pengawas.</li>
                            <li><strong>Kunci &amp; Sirine Anti-Curang:</strong> Dilarang membuka aplikasi/tab lain saat ujian berlangsung atau layar akan terkunci disertai bunyi sirine.</li>
                        </ol>
                    </div>

                    <!-- Panduan Pengawas -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                                <i data-lucide="shield" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">2. Alur Pengawas Ruang</h3>
                                <span class="text-xs text-amber-600 font-semibold">3 Menu Kendali Pengawasan</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Pantau Keluar/Masuk &amp; Token:</strong> Lihat log kehadiran siswa secara live beserta Token Ujian yang sinkron otomatis dengan Admin.</li>
                            <li><strong>Monitor Pelanggaran &amp; Sirine:</strong> Deteksi siswa yang keluar aplikasi, matikan alarm &amp; buka kunci layar jarak jauh.</li>
                            <li><strong>Input &amp; Upload Soal:</strong> Input soal manual atau upload file Word (.docx) / PDF (.pdf) per mata pelajaran yang langsung tersinkron ke Bank Soal.</li>
                        </ul>
                    </div>

                    <!-- Panduan Administrator -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                                <i data-lucide="cpu" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">3. Alur Administrator</h3>
                                <span class="text-xs text-indigo-600 font-semibold">Penjadwalan &amp; Audit Log</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Bank Soal Per Mata Pelajaran:</strong> Siapkan paket soal per mata pelajaran lewat upload Word/PDF atau input manual.</li>
                            <li><strong>Assign Jadwal Mata Pelajaran:</strong> Pilih soal yang sudah disiapkan di menu Assignment, tentukan tanggal/jam jadwal &amp; kelas target.</li>
                            <li><strong>Log Aktivitas Masuk/Keluar:</strong> Pantau riwayat login, logout, mulai ujian, dan pelanggaran seluruh Siswa &amp; Pengawas.</li>
                        </ul>
                    </div>
                </div>
            </div>

            <!-- BAGIAN GALERI FOTO KEGIATAN UJIAN -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-emerald-600">Dokumentasi &amp; Suasana Ujian</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Galeri Kegiatan Ujian Berbasis Komputer</h2>
                    </div>
                    <p class="text-xs text-slate-500">Dokumentasi kegiatan dapat dikelola melalui menu Setelan Administrator</p>
                </div>
                <div id="portalGalleryGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6"></div>
            </div>

            <!-- BAGIAN DOWNLOAD & INSTALASI APLIKASI -->
            <div class="bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 shadow-sm space-y-6">
                <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-5 border-b border-slate-100">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Aplikasi Multi-Platform</span>
                        <h2 class="text-xl font-extrabold text-slate-900 mt-1">Download &amp; Pasang Aplikasi di HP Android, iPhone, maupun Komputer</h2>
                        <p class="text-xs sm:text-sm text-slate-500 mt-1">Pasang sebagai aplikasi mandiri layar penuh untuk kenyamanan dan keamanan ujian maksimal.</p>
                    </div>
                    <button onclick="downloadStandaloneApp()" class="px-4 py-3 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-2 self-start md:self-center shrink-0 shadow-sm">
                        <i data-lucide="download" class="w-4 h-4"></i> Unduh File Aplikasi (.HTML)
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div class="rounded-2xl border border-emerald-200 bg-emerald-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">HP / Tablet Android</h3>
                                    <span class="text-[11px] font-semibold text-emerald-700">Android 8.0+ (Chrome / WebApp)</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang aplikasi CBT langsung ke layar utama Android tanpa bilah alamat browser sehingga berjalan seperti aplikasi mandiri.
                            </p>
                        </div>
                        <button onclick="openInstallModal('android')" class="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="download-cloud" class="w-4 h-4"></i> Install / Panduan Android
                        </button>
                    </div>

                    <div class="rounded-2xl border border-slate-200 bg-slate-50/70 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-slate-900 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="tablet-smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">iPhone &amp; iPad (iOS)</h3>
                                    <span class="text-[11px] font-semibold text-slate-600">iOS Safari WebClip / PWA</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Gunakan fitur <em>"Tambahkan ke Layar Utama"</em> di Safari iPhone/iPad untuk menjalankan CBT dalam mode layar penuh yang stabil.
                            </p>
                        </div>
                        <button onclick="openInstallModal('ios')" class="w-full py-2.5 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="smartphone" class="w-4 h-4"></i> Install / Panduan iPhone (iOS)
                        </button>
                    </div>

                    <div class="rounded-2xl border border-blue-200 bg-blue-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-blue-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="monitor" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">Komputer / Laptop (PC &amp; Mac)</h3>
                                    <span class="text-[11px] font-semibold text-blue-700">Windows / macOS / Chromebook</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang sebagai aplikasi desktop mandiri lewat Chrome/Edge atau unduh file paket aplikasi offline untuk laboratorium komputer.
                            </p>
                        </div>
                        <button onclick="openInstallModal('desktop')" class="w-full py-2.5 px-4 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="laptop" class="w-4 h-4"></i> Install / Download Komputer
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW B: DASHBOARD SISWA (PROFIL, PENUGASAN MAPEL, PANDUAN & TATIB) -->
        <!-- ================================================================= -->
        <section id="view-siswa-dashboard" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
            <!-- Header Sapaan Siswa -->
            <div class="bg-gradient-to-r from-slate-900 via-blue-950 to-indigo-950 rounded-3xl p-6 text-white shadow-lg flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div class="flex items-center gap-4">
                    <div class="w-14 h-14 rounded-2xl bg-blue-600/30 border border-blue-400/40 flex items-center justify-center font-extrabold text-xl text-blue-200 shrink-0">
                        <i data-lucide="user-check" class="w-7 h-7"></i>
                    </div>
                    <div>
                        <span class="px-2.5 py-0.5 rounded-full bg-emerald-500/20 border border-emerald-400/30 text-emerald-300 text-[10px] font-bold uppercase tracking-wider">
                            Akun Peserta Terverifikasi
                        </span>
                        <h2 id="stuDashWelcomeName" class="text-xl sm:text-2xl font-extrabold mt-1">Selamat Datang, Peserta</h2>
                        <p id="stuDashWelcomeSub" class="text-xs text-blue-200 mt-0.5">Silakan pilih mata pelajaran yang sudah di-assign untuk memulai ujian.</p>
                    </div>
                </div>
                <div class="flex items-center gap-2 self-start md:self-center">
                    <span class="px-3 py-1.5 rounded-xl bg-white/10 border border-white/15 text-xs font-semibold flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span> Sinkron Jadwal Live
                    </span>
                </div>
            </div>

            <!-- Navigasi Tab Dashboard Siswa -->
            <div class="flex flex-wrap gap-2 bg-white p-2 rounded-2xl border border-slate-200 shadow-sm">
                <button onclick="switchStudentDashTab('ujian')" id="stuTabBtn-ujian" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-blue-600 text-white transition flex items-center gap-2">
                    <i data-lucide="book-open-check" class="w-4 h-4"></i> Penugasan &amp; Jadwal Ujian
                </button>
                <button onclick="switchStudentDashTab('profil')" id="stuTabBtn-profil" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="id-card" class="w-4 h-4"></i> Profil &amp; Kartu Login Saya
                </button>
                <button onclick="switchStudentDashTab('panduan')" id="stuTabBtn-panduan" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="shield-alert" class="w-4 h-4"></i> Panduan Menjawab &amp; Tata Tertib
                </button>
            </div>

            <!-- TAB 1 SISWA: PENUGASAN / UJIAN MATA PELAJARAN TERJADWAL -->
            <div id="stuDashTab-ujian" class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <!-- Daftar Mata Pelajaran yang Di-Assign & Menunggu Assign -->
                <div class="lg:col-span-8 space-y-4 min-w-0">
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-extrabold text-slate-900 text-base">Daftar Ujian &amp; Penugasan Mata Pelajaran</h3>
                            <p class="text-xs text-slate-500">Pilih ujian yang berstatus <strong>Aktif / Siap Dikerjakan</strong>, lalu masukkan Token Ujian saat Anda siap.</p>
                        </div>
                        <span id="stuAssignedCountBadge" class="px-3 py-1 rounded-full bg-blue-50 text-blue-700 text-xs font-bold self-start">0 Paket</span>
                    </div>

                    <div id="studentAssignmentsList" class="space-y-4">
                        <!-- Diisi via JS: Kartu Mata Pelajaran Aktif (Siap Mulai + Token) & Menunggu Assign -->
                    </div>
                </div>

                <!-- Ringkasan Profil Kartu Login & Pengingat Tata Tertib -->
                <div class="lg:col-span-4 space-y-4 min-w-0">
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-2.5">
                            <span class="text-xs font-extrabold uppercase tracking-wider text-blue-600">Ringkasan Kartu Peserta</span>
                            <button onclick="switchStudentDashTab('profil')" class="text-[11px] font-bold text-slate-500 hover:text-blue-600">Lihat Kartu &rarr;</button>
                        </div>
                        <div class="space-y-2 text-xs">
                            <div class="flex justify-between">
                                <span class="text-slate-500">Nama Lengkap:</span>
                                <span id="stuSideName" class="font-bold text-slate-900">-</span>
                            </div>
                            <div class="flex justify-between">
                                <span class="text-slate-500">NISN / Username:</span>
                                <span id="stuSideNisn" class="font-mono font-bold text-slate-800">-</span>
                            </div>
                            <div class="flex justify-between">
                                <span class="text-slate-500">Kelas:</span>
                                <span id="stuSideClass" class="font-bold text-indigo-700">-</span>
                            </div>
                            <div class="flex justify-between">
                                <span class="text-slate-500">Ruangan Ujian:</span>
                                <span id="stuSideRoom" class="font-bold text-slate-900">-</span>
                            </div>
                        </div>
                    </div>

                    <div class="bg-amber-50 border border-amber-200 rounded-2xl p-5 space-y-2.5 text-xs text-amber-950">
                        <div class="font-extrabold text-amber-900 flex items-center gap-1.5 text-sm">
                            <i data-lucide="siren" class="w-4 h-4 text-red-600"></i> Peringatan Keamanan Ruang Ujian
                        </div>
                        <p class="leading-relaxed text-amber-900">
                            Setelah Anda menekan tombol <strong>Mulai Ujian</strong> dan memasukkan Token:
                        </p>
                        <ul class="list-disc pl-4 space-y-1 text-amber-900">
                            <li>Aplikasi akan mengunci mode layar penuh.</li>
                            <li>Dilarang berpindah tab, membuka WhatsApp/Google, atau membagi layar.</li>
                            <li>Jika keluar sistem, <strong>layar akan terkunci otomatis &amp; sirine alarm berbunyi keras</strong> serta tercatat di Pengawas &amp; Admin.</li>
                        </ul>
                        <button onclick="switchStudentDashTab('panduan')" class="mt-2 inline-flex items-center gap-1 text-xs font-bold text-red-700 hover:underline">
                            Baca Selengkapnya Tata Tertib &amp; Panduan &rarr;
                        </button>
                    </div>
                </div>
            </div>

            <!-- TAB 2 SISWA: PROFIL LENGKAP SESUAI KARTU LOGIN -->
            <div id="stuDashTab-profil" class="hidden grid grid-cols-1 lg:grid-cols-12 gap-6">
                <div class="lg:col-span-6 bg-white rounded-3xl min-w-0 border-2 border-slate-300 p-6 shadow-sm space-y-4">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-200">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-xl bg-slate-900 text-white flex items-center justify-center font-extrabold text-sm">CBT</div>
                            <div>
                                <div class="text-xs font-extrabold uppercase tracking-wider text-slate-900">KARTU IDENTITAS PESERTA UJIAN</div>
                                <div id="stuCardSchoolName" class="text-[11px] text-slate-500">CBT-PRO | Sistem Ujian Berbasis Komputer</div>
                            </div>
                        </div>
                        <span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold">AKTIF</span>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs pt-1">
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Nama Lengkap Peserta</span>
                            <span id="stuProfName" class="text-sm font-extrabold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">NISN</span>
                            <span id="stuProfNisn" class="text-sm font-mono font-extrabold text-blue-700 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Kelas / Jurusan</span>
                            <span id="stuProfClass" class="text-sm font-bold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Ruangan Ujian</span>
                            <span id="stuProfRoom" class="text-sm font-bold text-indigo-700 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                            <span class="text-[10px] uppercase font-bold text-blue-500 block">Username Kartu Login</span>
                            <span id="stuProfUser" class="text-sm font-mono font-extrabold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                            <span class="text-[10px] uppercase font-bold text-blue-500 block">Status Sesi Login</span>
                            <span id="stuProfLoginTime" class="text-xs font-bold text-emerald-700 mt-0.5 block">Terhubung Real-Time</span>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-6 bg-white rounded-3xl min-w-0 border border-slate-200 p-6 shadow-sm space-y-4">
                    <h3 class="font-extrabold text-slate-900 text-base flex items-center gap-2">
                        <i data-lucide="history" class="w-4 h-4 text-blue-600"></i> Riwayat Pengerjaan Ujian Saya
                    </h3>
                    <div id="stuMyHistoryList" class="space-y-2.5 text-xs">
                        <!-- Riwayat ujian siswa -->
                    </div>
                </div>
            </div>

            <!-- TAB 3 SISWA: PANDUAN MENJAWAB SOAL & TATA TERTIB RUANG UJIAN -->
            <div id="stuDashTab-panduan" class="hidden grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="bg-white rounded-3xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div class="flex items-center gap-3 pb-3 border-b border-slate-100">
                        <div class="w-10 h-10 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center">
                            <i data-lucide="help-circle" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h3 class="font-extrabold text-slate-900 text-base">Panduan Cara Menjawab Soal CBT</h3>
                            <span class="text-xs text-slate-500">Langkah teknis pengerjaan soal pada sistem</span>
                        </div>
                    </div>
                    <ol class="space-y-3 text-xs sm:text-sm text-slate-700 list-decimal pl-5 leading-relaxed">
                        <li><strong>Pilih Mata Pelajaran Terjadwal:</strong> Pada tab <em>Penugasan &amp; Jadwal Ujian</em>, pastikan status ujian sudah <strong>Siap Dikerjakan (Assigned)</strong>, lalu klik tombol <strong>Mulai Ujian</strong>.</li>
                        <li><strong>Masukkan Token Ujian:</strong> Ketikkan kode <strong>Token Ujian</strong> yang diberikan oleh Pengawas Ruangan pada layar konfirmasi token.</li>
                        <li><strong>Memilih Jawaban:</strong> Klik langsung pada pilihan jawaban <strong>A, B, C, D, atau E</strong>. Jawaban tersimpan otomatis pada sistem.</li>
                        <li><strong>Navigasi Nomor &amp; Ragu-Ragu:</strong> Gunakan panel nomor soal di sebelah kanan untuk melompat ke nomor tertentu. Gunakan tombol <em>Tandai Ragu-Ragu</em> (warna kuning) jika ingin meninjau ulang soal tersebut.</li>
                        <li><strong>Pengaturan Ukuran Teks:</strong> Gunakan tombol <strong>A- / A / A+</strong> di bagian atas layar ujian untuk menyesuaikan ukuran huruf soal.</li>
                        <li><strong>Mengakhiri Ujian:</strong> Setelah seluruh soal terjawab dan diperiksa, klik tombol <strong>Selesai Ujian</strong> untuk mengirim hasil secara permanen.</li>
                    </ol>
                </div>

                <div class="bg-white rounded-3xl border border-red-200 p-6 shadow-sm space-y-4">
                    <div class="flex items-center gap-3 pb-3 border-b border-red-100">
                        <div class="w-10 h-10 rounded-xl bg-red-100 text-red-600 flex items-center justify-center">
                            <i data-lucide="shield-alert" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h3 class="font-extrabold text-red-900 text-base">Tata Tertib Peserta di Ruangan Ujian</h3>
                            <span class="text-xs text-red-600 font-semibold">Wajib Dipatuhi Seluruh Peserta Ujian</span>
                        </div>
                    </div>
                    <ul class="space-y-3 text-xs sm:text-sm text-slate-700 list-disc pl-5 leading-relaxed">
                        <li>Peserta wajib hadir di ruangan ujian minimal 15 menit sebelum jadwal dimulai dengan membawa <strong>Kartu Login Peserta</strong>.</li>
                        <li>Peserta dilarang keras membuka tab browser lain, aplikasi pesan, kalkulator eksternal, catatan digital, atau melakukan <em>split-screen</em> selama ujian berlangsung.</li>
                        <li>Peserta dilarang keluar dari mode layar penuh (<em>Fullscreen</em>), menekan tombol kembali (<em>Back</em>), atau mengambil tangkapan layar (<em>Screenshot</em>).</li>
                        <li><strong>Sanksi Sistem Otomatis:</strong> Setiap upaya keluar dari halaman ujian akan langsung <strong>mengunci layar ujian, membunyikan sirine alarm peringatan</strong>, serta mencatat pelanggaran secara real-time di layar Pengawas dan Administrator.</li>
                        <li>Pembukaan kunci layar hanya dapat dilakukan oleh <strong>Pengawas Ruangan</strong> setelah memeriksa perangkat peserta.</li>
                        <li>Kerjakan soal secara mandiri, jujur, dan tenang hingga waktu berakhir.</li>
                    </ul>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW C: DASHBOARD ADMINISTRATOR (7 MENU SIDEBAR TERPISAH)         -->
        <!-- ================================================================= -->
        <section id="view-admin" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                
                <!-- SIDEBAR NAVIGASI -->
                <aside class="lg:col-span-3 bg-white min-w-0 rounded-2xl border border-slate-200 p-4 shadow-sm space-y-2 lg:sticky lg:top-20 no-print">
                    <div class="px-3 py-2 border-b border-slate-100 mb-2">
                        <span class="text-[10px] font-extrabold uppercase tracking-wider text-indigo-600">Panel Kendali Utama</span>
                        <h2 class="font-extrabold text-slate-900 text-base">Menu Administrator</h2>
                    </div>

                    <nav class="flex lg:flex-col overflow-x-auto lg:overflow-visible gap-1.5 pb-1 lg:pb-0">
                        <button onclick="switchAdminMenu('monitoring')" id="adminNav-monitoring" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap">
                            <i data-lucide="activity" class="w-4 h-4 shrink-0"></i>
                            <span>1. Monitoring Siswa</span>
                        </button>
                        <button onclick="switchAdminMenu('soal')" id="adminNav-soal" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-text" class="w-4 h-4 shrink-0"></i>
                            <span>2. Bank Soal (Per Mapel)</span>
                        </button>
                        <button onclick="switchAdminMenu('assign')" id="adminNav-assign" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="calendar-check" class="w-4 h-4 shrink-0"></i>
                            <span>3. Assignment &amp; Jadwal</span>
                        </button>
                        <button onclick="switchAdminMenu('logs')" id="adminNav-logs" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="scroll-text" class="w-4 h-4 shrink-0"></i>
                            <span>4. Log Masuk/Keluar &amp; Pelanggaran</span>
                        </button>
                        <button onclick="switchAdminMenu('nilai')" id="adminNav-nilai" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-spreadsheet" class="w-4 h-4 shrink-0"></i>
                            <span>5. Data &amp; Download Nilai</span>
                        </button>
                        <button onclick="switchAdminMenu('kartu')" id="adminNav-kartu" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="id-card" class="w-4 h-4 shrink-0"></i>
                            <span>6. Akun Siswa &amp; Kartu</span>
                        </button>
                        <button onclick="switchAdminMenu('setelan')" id="adminNav-setelan" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="settings-2" class="w-4 h-4 shrink-0"></i>
                            <span>7. Setelan Keseluruhan</span>
                        </button>
                    </nav>

                    <!-- Info Token Aktif di Sidebar -->
                    <div class="hidden lg:block mt-4 pt-4 border-t border-slate-100">
                        <div class="bg-slate-900 text-white rounded-xl p-3.5">
                            <div class="flex items-center justify-between text-[10px] text-slate-400 uppercase font-bold">
                                <span>Token Ujian Aktif</span>
                                <span class="text-emerald-400">SINKRON</span>
                            </div>
                            <div class="flex items-center justify-between mt-1">
                                <span id="sidebarTokenDisplay" class="font-mono font-extrabold text-xl text-amber-400 tracking-widest">------</span>
                                <button onclick="copyTextToClipboard(document.getElementById('sidebarTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white" title="Salin Token">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                </aside>

                <!-- KONTEN KANAN ADMINISTRATOR -->
                <div class="lg:col-span-9 space-y-6 min-w-0">

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 1: MONITORING SISWA (REAL-TIME)                  -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-monitoring" class="space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Monitoring Peserta Ujian Real-Time</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Pantau status siswa di Dashboard maupun ruang ujian, matikan sirine/buka kunci layar, atau reset sesi.</p>
                            </div>
                            <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start">
                                <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Sinkronisasi Real-Time
                            </span>
                        </div>

                        <div class="grid grid-cols-2 sm:grid-cols-4 gap-4">
                            <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm">
                                <span class="text-xs text-slate-500 font-medium">Total Peserta</span>
                                <div id="monStatTotal" class="text-2xl font-extrabold text-slate-900 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-blue-200 p-4 shadow-sm">
                                <span class="text-xs text-blue-600 font-semibold">Sedang Ujian / Online</span>
                                <div id="monStatActive" class="text-2xl font-extrabold text-blue-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-emerald-200 p-4 shadow-sm">
                                <span class="text-xs text-emerald-600 font-semibold">Selesai Ujian</span>
                                <div id="monStatDone" class="text-2xl font-extrabold text-emerald-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-red-200 p-4 shadow-sm">
                                <span class="text-xs text-red-600 font-semibold">Terkunci / Sirine</span>
                                <div id="monStatLocked" class="text-2xl font-extrabold text-red-600 mt-1">0</div>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="monSearchInput" oninput="renderAdminMonitoring()" placeholder="Cari nama atau NISN siswa..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-64 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <select id="monRoomFilter" onchange="renderAdminMonitoring()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Kelas / Ruangan</option>
                                </select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Peserta &amp; NISN</th>
                                            <th class="py-3 px-4 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-4 font-bold">Mapel Aktif</th>
                                            <th class="py-3 px-4 font-bold">Status Live</th>
                                            <th class="py-3 px-4 font-bold">Progres</th>
                                            <th class="py-3 px-4 font-bold">Pelanggaran</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Cepat</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminMonitoringTbody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 2: BANK SOAL PER MATA PELAJARAN / PAKET          -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-soal" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Bank Soal &amp; Paket Mata Pelajaran</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Kelola paket soal per mata pelajaran (Upload Word/PDF atau Input Manual), lalu tugaskan di menu <strong>Assignment &amp; Jadwal</strong>.</p>
                            </div>
                        </div>

                        <!-- Kelola / Pilih Paket Mata Pelajaran -->
                        <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-4">
                            <div class="flex flex-col md:flex-row md:items-end justify-between gap-4">
                                <div class="flex-1">
                                    <label class="block text-xs font-extrabold uppercase tracking-wider text-indigo-700 mb-1.5">Pilih Paket Soal Mata Pelajaran yang Dikelola:</label>
                                    <select id="adminPackageSelector" onchange="selectAdminQuestionPackage(this.value)" class="w-full px-4 py-2.5 rounded-xl border-2 border-indigo-300 bg-indigo-50/40 text-sm font-bold text-slate-900">
                                        <!-- Diisi via JS -->
                                    </select>
                                </div>
                                <form onsubmit="createNewSubjectPackage(event)" class="flex flex-wrap sm:flex-nowrap gap-2">
                                    <input type="text" id="newPkgSubjectName" required placeholder="Nama Mapel Baru (Misal: Matematika)" class="px-3.5 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <button type="submit" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold whitespace-nowrap transition">
                                        + Buat Paket Mapel
                                    </button>
                                </form>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6">
                            <!-- Kolom Input Soal -->
                            <div class="xl:col-span-5 space-y-6 min-w-0">
                                <!-- Upload Word (.docx) & PDF (.pdf) -->
                                <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-3">
                                    <div class="flex items-center justify-between">
                                        <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                            <i data-lucide="file-up" class="w-4 h-4 text-indigo-600"></i> Upload Soal Word (.docx) / PDF (.pdf)
                                        </h3>
                                        <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-indigo-600 hover:underline">
                                            Unduh Contoh Format
                                        </button>
                                    </div>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Soal akan masuk ke paket mata pelajaran yang dipilih di atas.<br>
                                        <code class="block bg-slate-50 p-2 rounded-lg border border-slate-200 text-slate-700 mt-1">
                                            1. Pertanyaan soal di sini?<br>
                                            A. Pilihan A &nbsp; B. Pilihan B<br>
                                            C. Pilihan C &nbsp; D. Pilihan D &nbsp; E. Pilihan E<br>
                                            Kunci: B
                                        </code>
                                    </p>
                                    <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-indigo-300 rounded-2xl bg-indigo-50/40 hover:bg-indigo-50 cursor-pointer transition p-4 text-center">
                                        <i data-lucide="upload-cloud" class="w-7 h-7 text-indigo-600 mb-1"></i>
                                        <span class="text-xs font-bold text-indigo-900">Klik untuk Pilih File .DOCX, .PDF, atau .TXT</span>
                                        <span class="text-[10px] text-slate-500 mt-0.5">Otomatis mengekstrak soal, opsi A-E, dan kunci jawaban</span>
                                        <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event, 'admin')" class="hidden">
                                    </label>
                                    <div id="uploadParseStatus" class="hidden text-xs p-2.5 rounded-xl bg-blue-50 text-blue-700 font-medium"></div>
                                </div>

                                <!-- Form Input Soal Manual -->
                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Soal Manual ke Paket Ini</h3>
                                    <form onsubmit="handleAddManualQuestion(event, 'admin')" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan *</label>
                                            <textarea id="mQText" required rows="2" placeholder="Tuliskan pertanyaan..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"></textarea>
                                        </div>
                                        <div class="space-y-2">
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">A.</span>
                                                <input type="text" id="mQOptA" required placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">B.</span>
                                                <input type="text" id="mQOptB" required placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">C.</span>
                                                <input type="text" id="mQOptC" required placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">D.</span>
                                                <input type="text" id="mQOptD" required placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">E.</span>
                                                <input type="text" id="mQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-3 pt-1">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban</label>
                                                <select id="mQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                                                    <option value="A">Opsi A</option>
                                                    <option value="B">Opsi B</option>
                                                    <option value="C">Opsi C</option>
                                                    <option value="D">Opsi D</option>
                                                    <option value="E">Opsi E</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin</label>
                                                <input type="number" id="mQPoints" value="10" min="1" max="100" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Soal ke Paket
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <!-- Daftar & Edit Soal -->
                            <div class="xl:col-span-7 bg-white min-w-0 rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col">
                                <div class="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Soal Paket Terpilih (<span id="adminBankSoalCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Soal yang sudah siap di sini dapat langsung di-assign di Menu 3</p>
                                    </div>
                                    <button onclick="confirmClearPackageQuestions()" class="px-3 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                        Kosongkan Soal Paket
                                    </button>
                                </div>
                                <div id="adminSoalListContainer" class="space-y-3 overflow-y-auto max-h-[680px] pr-1"></div>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 3: ASSIGNMENT MATA PELAJARAN, JADWAL & TOKEN     -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-assign" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Assignment Mata Pelajaran, Jadwal Ujian &amp; Token</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Pilih soal mata pelajaran yang telah disiapkan di Bank Soal, atur jadwal pelaksanaan &amp; status Assignment (Aktif / Menunggu Assign), serta kelola Token sinkron.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                            <!-- Manajemen Token Ujian (Sinkron Pengawas & Siswa) -->
                            <div class="lg:col-span-5 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5 self-start">
                                <div>
                                    <span class="px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">Sinkron Otomatis ke Pengawas &amp; Siswa</span>
                                    <h3 class="font-bold text-slate-900 text-base mt-2">Token Akses Ujian Aktif</h3>
                                    <p class="text-xs text-slate-500 mt-0.5">Siswa memasukkan token ini di Dashboard Siswa saat menekan tombol Mulai Ujian pada mapel yang sudah di-assign.</p>
                                </div>

                                <div class="bg-slate-900 text-white p-5 rounded-2xl flex items-center justify-between">
                                    <div>
                                        <span class="text-[10px] uppercase text-slate-400 font-bold block">Token Saat Ini</span>
                                        <span id="assignTokenDisplay" class="text-3xl font-mono font-extrabold text-amber-400 tracking-widest">CBT26X</span>
                                    </div>
                                    <button onclick="copyTextToClipboard(document.getElementById('assignTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-3 rounded-xl bg-white/10 hover:bg-white/20 text-white transition" title="Salin Token">
                                        <i data-lucide="copy" class="w-5 h-5"></i>
                                    </button>
                                </div>

                                <div class="space-y-2 pt-2 border-t border-slate-100">
                                    <label class="block text-xs font-bold text-slate-700">Ubah Token Secara Manual</label>
                                    <div class="flex gap-2">
                                        <input type="text" id="customTokenInput" maxlength="10" placeholder="Misal: UAS2026" class="flex-1 px-3.5 py-2.5 rounded-xl border border-slate-300 font-mono font-bold uppercase text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                        <button onclick="saveManualToken()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                            Simpan
                                        </button>
                                    </div>
                                </div>

                                <button onclick="generateRandomToken()" class="w-full py-3 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition flex items-center justify-center gap-2 shadow-sm">
                                    <i data-lucide="refresh-cw" class="w-4 h-4"></i> Acak / Generate Token Baru Otomatis
                                </button>
                            </div>

                            <!-- Form Penugasan (Assign) Paket Soal Mata Pelajaran -->
                            <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="calendar-plus" class="w-4 h-4 text-indigo-600"></i> Jadwalkan &amp; Assign Paket Soal ke Siswa
                                </h3>
                                <form onsubmit="savePackageAssignment(event)" class="space-y-4">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Pilih Paket Soal dari Bank Soal *</label>
                                        <select id="assignPackageSelect" onchange="loadSelectedPackageToAssignForm()" required class="w-full px-3.5 py-2.5 rounded-xl border-2 border-indigo-200 bg-indigo-50/30 text-sm font-bold text-slate-900">
                                            <!-- Diisi via JS -->
                                        </select>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Judul Sesi Ujian *</label>
                                            <input type="text" id="assignExamTitle" required placeholder="Misal: UAS Ganjil 2026" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Mata Pelajaran *</label>
                                            <input type="text" id="assignSubject" required placeholder="Misal: Akuntansi Dasar" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Tanggal Jadwal *</label>
                                            <input type="date" id="assignScheduleDate" required class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Jam Mulai - Selesai *</label>
                                            <input type="text" id="assignScheduleTime" required placeholder="08:00 - 09:30 WITA" class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Durasi (Menit) *</label>
                                            <input type="number" id="assignDuration" min="1" max="360" required value="45" class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-sm font-bold">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Target Kelas / Ruangan</label>
                                            <select id="assignTargetGroup" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-semibold">
                                                <option value="ALL">Semua Kelas &amp; Ruangan</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Status Penugasan (Assign)</label>
                                            <select id="assignStatusSelect" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-bold text-emerald-700">
                                                <option value="ASSIGNED">DI-ASSIGN (Siap Dikerjakan Siswa)</option>
                                                <option value="WAITING">TERJADWAL (Menunggu Assign)</option>
                                            </select>
                                        </div>
                                    </div>

                                    <div class="p-4 rounded-2xl bg-slate-50 border border-slate-200 grid grid-cols-1 sm:grid-cols-2 gap-3">
                                        <label class="flex items-center gap-2.5 cursor-pointer">
                                            <input type="checkbox" id="assignShuffleQuestions" class="w-4 h-4 rounded text-indigo-600">
                                            <span class="text-xs font-bold text-slate-800">Acak Urutan Soal</span>
                                        </label>
                                        <label class="flex items-center gap-2.5 cursor-pointer">
                                            <input type="checkbox" id="assignShuffleOptions" class="w-4 h-4 rounded text-indigo-600">
                                            <span class="text-xs font-bold text-slate-800">Acak Pilihan Jawaban</span>
                                        </label>
                                    </div>

                                    <div class="flex justify-end pt-1">
                                        <button type="submit" class="px-6 py-3 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-bold transition shadow-sm flex items-center gap-2">
                                            <i data-lucide="check-circle" class="w-4 h-4"></i> Simpan &amp; Sinkronkan Penugasan
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>

                        <!-- Tabel Daftar Penugasan Mata Pelajaran -->
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex items-center justify-between">
                                <div>
                                    <h3 class="font-bold text-slate-900 text-base">Daftar Jadwal &amp; Status Assignment Mata Pelajaran</h3>
                                    <p class="text-xs text-slate-500">Ubah status antara <strong>Siap Dikerjakan (Assigned)</strong> atau <strong>Menunggu Assign</strong> kapan saja sesuai jadwal.</p>
                                </div>
                            </div>
                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Mata Pelajaran &amp; Paket</th>
                                            <th class="py-3 px-4 font-bold">Jadwal Pelaksanaan</th>
                                            <th class="py-3 px-4 font-bold">Target</th>
                                            <th class="py-3 px-4 font-bold">Jumlah Soal</th>
                                            <th class="py-3 px-4 font-bold">Status</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Assignment</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminAssignmentsTableBody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 4: LOG MASUK/KELUAR SISWA & PENGAWAS             -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-logs" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Audit Log Aktivitas Masuk/Keluar &amp; Pelanggaran</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Rekam jejak real-time seluruh aktivitas login, logout, mulai ujian, selesai ujian, dan pelanggaran keluar sistem oleh Siswa &amp; Pengawas.</p>
                            </div>
                            <div class="flex items-center gap-2">
                                <button onclick="exportActivityLogsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="download" class="w-3.5 h-3.5"></i> Export Log (.XLSX)
                                </button>
                                <button onclick="confirmClearActivityLogs()" class="px-3.5 py-2.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                    Bersihkan Log
                                </button>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="logSearchInput" oninput="renderAdminLogsTable()" placeholder="Cari nama pengguna, NISN, atau aktivitas..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-72 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <div class="flex gap-2">
                                    <select id="logRoleFilter" onchange="renderAdminLogsTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                        <option value="">Semua Peran (Siswa &amp; Pengawas)</option>
                                        <option value="siswa">Khusus Siswa</option>
                                        <option value="pengawas">Khusus Pengawas</option>
                                        <option value="admin">Administrator</option>
                                    </select>
                                    <select id="logTypeFilter" onchange="renderAdminLogsTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                        <option value="">Semua Jenis Log</option>
                                        <option value="LOGIN">Login / Masuk Sistem</option>
                                        <option value="LOGOUT">Logout / Keluar Sistem</option>
                                        <option value="EXAM_START">Mulai Ujian</option>
                                        <option value="EXAM_FINISH">Selesai Ujian</option>
                                        <option value="VIOLATION">Pelanggaran / Sirine</option>
                                        <option value="UNLOCK">Buka Kunci Layar</option>
                                    </select>
                                </div>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[540px] overflow-y-auto">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                            <th class="py-3 px-4 font-bold">Waktu</th>
                                            <th class="py-3 px-4 font-bold">Peran</th>
                                            <th class="py-3 px-4 font-bold">Nama Pengguna &amp; ID</th>
                                            <th class="py-3 px-4 font-bold">Kategori</th>
                                            <th class="py-3 px-4 font-bold">Keterangan Aktivitas</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminLogsTableBody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 5: DATA SISWA & DOWNLOAD NILAI (.XLSX)           -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-nilai" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Data Peserta &amp; Rekap Nilai Ujian</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Nilai masuk secara real-time dan dapat diunduh langsung dalam format Microsoft Excel (.xlsx).</p>
                            </div>
                            <button onclick="downloadExcelResults()" class="px-5 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2 shadow-sm self-start">
                                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Download Nilai Excel (.XLSX)
                            </button>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                <input type="text" id="nilaiSearchInput" oninput="renderAdminNilaiTable()" placeholder="Cari Nama atau NISN..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                <select id="nilaiClassFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Kelas</option>
                                </select>
                                <select id="nilaiRoomFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <option value="">Semua Ruangan</option>
                                </select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-3 font-bold">No</th>
                                            <th class="py-3 px-3 font-bold">Nama Peserta</th>
                                            <th class="py-3 px-3 font-bold">NISN</th>
                                            <th class="py-3 px-3 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-3 font-bold">Mata Pelajaran</th>
                                            <th class="py-3 px-3 font-bold">Status</th>
                                            <th class="py-3 px-3 font-bold">Benar</th>
                                            <th class="py-3 px-3 font-bold">Salah</th>
                                            <th class="py-3 px-3 font-bold">Nilai Akhir</th>
                                            <th class="py-3 px-3 font-bold">Pelanggaran</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminNilaiTbody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 6: AKUN SISWA SIAP DIGUNAKAN & KARTU LOGIN        -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-kartu" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Manajemen Akun Siswa &amp; Kartu Login Resmi</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Kelola akun siswa siap digunakan berbasis NISN, edit/reset password, import Excel, dan cetak Kartu Login A4.</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-2">
                                <button onclick="exportStudentAccountsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Export Akun (.XLSX)
                                </button>
                                <button onclick="window.print()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="printer" class="w-4 h-4"></i> Cetak Kartu Login (A4)
                                </button>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6 no-print">
                            <div class="xl:col-span-4 space-y-5 min-w-0">
                                <div class="bg-white rounded-2xl border border-emerald-200 p-5 shadow-sm space-y-3">
                                    <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                        <i data-lucide="file-spreadsheet" class="w-4 h-4 text-emerald-600"></i> Import Data Siswa via Excel
                                    </h3>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Format kolom: <strong>Nama, NISN, Kelas, Ruangan</strong>. Username otomatis menggunakan NISN dan Password dibuat otomatis.
                                    </p>
                                    <button onclick="downloadStudentExcelTemplate()" class="w-full py-2.5 px-4 rounded-xl bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 text-xs font-bold transition flex items-center justify-center gap-2">
                                        <i data-lucide="download" class="w-4 h-4"></i> Download Template Excel (.XLSX)
                                    </button>
                                    <label class="flex flex-col items-center justify-center w-full py-3.5 px-3 border-2 border-dashed border-emerald-300 rounded-xl bg-emerald-50/30 hover:bg-emerald-50 cursor-pointer transition text-center">
                                        <span class="text-xs font-bold text-emerald-900"><i data-lucide="upload" class="w-3.5 h-3.5 inline mr-1"></i> Upload File Excel (.XLSX / .CSV)</span>
                                        <input type="file" accept=".xlsx,.xls,.csv" onchange="handleStudentExcelUpload(event)" class="hidden">
                                    </label>
                                </div>

                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Akun Siswa Manual</h3>
                                    <form onsubmit="handleAddManualStudent(event)" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Nama Lengkap *</label>
                                            <input type="text" id="kStuName" required placeholder="Contoh: Nadia Putri" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">NISN *</label>
                                            <input type="text" id="kStuNisn" required placeholder="Contoh: 0081234504" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-mono">
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kelas *</label>
                                                <input type="text" id="kStuClass" required placeholder="XII IPA 1" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Ruangan *</label>
                                                <input type="text" id="kStuRoom" required placeholder="Ruang 01" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Username (Opsional)</label>
                                                <input type="text" id="kStuUser" placeholder="Otomatis NISN" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Password (Opsional)</label>
                                                <input type="text" id="kStuPass" placeholder="Otomatis Acak" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Akun &amp; Buat Kartu
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <div class="xl:col-span-8 bg-white min-w-0 rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col space-y-4">
                                <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Akun Siswa Aktif (<span id="studentAccountCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Gunakan NISN/Username dan Password di bawah untuk login ke Dashboard Siswa</p>
                                    </div>
                                    <div class="flex items-center gap-2">
                                        <button onclick="toggleShowStudentPasswords()" id="btnTogglePassVis" class="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="eye" class="w-3.5 h-3.5"></i> <span>Tampilkan Password</span>
                                        </button>
                                        <button onclick="confirmClearAllStudents()" class="px-3 py-1.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                            Kosongkan
                                        </button>
                                    </div>
                                </div>

                                <input type="text" id="accountSearchInput" oninput="renderAdminStudentAccountTable()" placeholder="Cari berdasarkan Nama, NISN, Kelas, atau Ruangan..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full focus:ring-2 focus:ring-indigo-500 focus:outline-none">

                                <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[380px] overflow-y-auto">
                                    <table class="w-full text-left border-collapse text-xs">
                                        <thead>
                                            <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                                <th class="py-2.5 px-3 font-bold">Nama &amp; NISN</th>
                                                <th class="py-2.5 px-3 font-bold">Kelas / Ruang</th>
                                                <th class="py-2.5 px-3 font-bold">Username</th>
                                                <th class="py-2.5 px-3 font-bold">Password</th>
                                                <th class="py-2.5 px-3 font-bold text-right">Kelola Akun</th>
                                            </tr>
                                        </thead>
                                        <tbody id="adminStudentAccountTbody" class="divide-y divide-slate-100"></tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex items-center justify-between no-print">
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm sm:text-base">Pratinjau Kartu Login Peserta Siap Cetak (<span id="cardTotalCount">0</span> Kartu)</h3>
                                    <p class="text-[11px] text-slate-500">Format cetak otomatis 2 kolom per halaman A4</p>
                                </div>
                                <button onclick="window.print()" class="px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="printer" class="w-3.5 h-3.5"></i> Cetak Sekarang
                                </button>
                            </div>
                            <div id="printCardsContainer" class="grid grid-cols-1 sm:grid-cols-2 gap-4"></div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 7: SETELAN KESELURUHAN                           -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-setelan" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Setelan Keseluruhan Sistem, Pengawas &amp; Keamanan</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Kelola identitas sekolah, kredensial administrator, akun pengawas ruangan, sirine &amp; keamanan anti-curang, galeri portal, serta backup database.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="school" class="w-4 h-4 text-indigo-600"></i> Identitas Sekolah &amp; Akun Administrator
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Nama Sekolah / Lembaga</label>
                                        <input type="text" id="setSchoolName" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Subjudul / Alamat &amp; Tahun Pelajaran</label>
                                        <input type="text" id="setSchoolSubtitle" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div class="pt-2 border-t border-slate-100 grid grid-cols-2 gap-2.5">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Username Admin</label>
                                            <input type="text" id="setAdminUsername" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Password Admin</label>
                                            <input type="text" id="setAdminPassword" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                    </div>
                                    <button onclick="saveSchoolIdentity()" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition">
                                        Simpan Identitas &amp; Akun Admin
                                    </button>
                                </div>
                            </div>

                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="shield-alert" class="w-4 h-4 text-red-600"></i> Setelan Keamanan, Sirine &amp; Perilaku Siswa
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-red-700 mb-1">Kode PIN Pembuka Kunci Layar (Default: 123)</label>
                                        <input type="text" id="setUnlockPin" class="w-full px-3.5 py-2 rounded-xl border border-red-300 bg-red-50/40 font-mono font-extrabold text-red-700 text-sm">
                                    </div>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700 pt-1">
                                        <span>Wajib Mode Layar Penuh (Fullscreen)</span>
                                        <input type="checkbox" id="setRequireFullscreen" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Kunci Otomatis Jika Pindah Tab / Aplikasi</span>
                                        <input type="checkbox" id="setDetectTabSwitch" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Bunyikan Sirine Alarm Keras Saat Pelanggaran</span>
                                        <input type="checkbox" id="setEnableSiren" class="w-4 h-4 rounded text-red-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Tampilkan Skor Akhir ke Siswa Setelah Selesai</span>
                                        <input type="checkbox" id="setShowScoreToStudent" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <div class="flex flex-wrap gap-2 pt-1">
                                        <button onclick="saveSecuritySettings()" class="px-5 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                            Simpan Setelan Keamanan
                                        </button>
                                        <button type="button" onclick="testSirenSound()" class="px-4 py-2.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Tes Suara Sirine (2 Detik)
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="user-cog" class="w-4 h-4 text-amber-600"></i> Manajemen Akun Pengawas Ruangan
                                </h3>
                                <form onsubmit="handleAddProctor(event)" class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="procName" required placeholder="Nama Pengawas" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procRoom" required placeholder="Ruang Tugas (Misal: Ruang 01 / Semua)" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procUser" required placeholder="Username Login" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                    <div class="flex gap-2">
                                        <input type="text" id="procPass" required placeholder="Password" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                        <button type="submit" class="px-3 py-2 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold shrink-0">
                                            + Tambah
                                        </button>
                                    </div>
                                </form>
                                <div id="proctorListContainer" class="space-y-2 max-h-52 overflow-y-auto"></div>
                            </div>

                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="image" class="w-4 h-4 text-emerald-600"></i> Galeri Foto Portal &amp; Manajemen Database
                                </h3>
                                <form onsubmit="handleAddGalleryItem(event)" class="space-y-2 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="galTitle" required placeholder="Judul Foto Kegiatan (Misal: Pelaksanaan Ujian Sesi 1)" class="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <div class="flex gap-2">
                                        <input type="url" id="galUrl" placeholder="URL Gambar Eksternal (Kosongkan untuk Ilustrasi SVG)" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                        <button type="submit" class="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shrink-0">
                                            + Tambah Foto
                                        </button>
                                    </div>
                                </form>

                                <div class="pt-3 border-t border-slate-100 space-y-2">
                                    <span class="text-xs font-bold text-slate-700 block">Cadangan, Pemulihan &amp; Bersihkan Data Awal</span>
                                    <div class="flex flex-wrap gap-2">
                                        <button onclick="exportDatabaseJSON()" class="px-3.5 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="download" class="w-3.5 h-3.5"></i> Backup JSON
                                        </button>
                                        <label class="px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer">
                                            <i data-lucide="upload" class="w-3.5 h-3.5"></i> Restore JSON
                                            <input type="file" accept=".json" onchange="importDatabaseJSON(event)" class="hidden">
                                        </label>
                                        <button onclick="confirmCleanStartEmpty()" class="px-3.5 py-2 rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="eraser" class="w-3.5 h-3.5"></i> Mulai Kosong
                                        </button>
                                        <button onclick="confirmFactoryReset()" class="px-3.5 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i> Reset Bawaan
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW D: DASHBOARD PENGAWAS RUANG (3 MENU TERSTRUKTUR & SINKRON)   -->
        <!-- ================================================================= -->
        <section id="view-pengawas" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
            <!-- Header Pengawas -->
            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                    <div class="flex items-center gap-2">
                        <span class="px-2.5 py-0.5 rounded-md bg-amber-100 text-amber-800 text-xs font-bold">DASHBOARD PENGAWAS RUANG</span>
                        <span id="proctorAssignedRoomBadge" class="px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-700 text-xs font-semibold">Semua Ruangan</span>
                    </div>
                    <h2 id="proctorDashboardName" class="text-2xl font-extrabold text-slate-900 mt-1">Pengawas Ujian</h2>
                    <p class="text-xs text-slate-500 mt-0.5">Terhubung real-time dengan Administrator &amp; Peserta Ujian di ruangan.</p>
                </div>

                <!-- Kartu Token Ruang Aktif (Sinkron dengan Admin) -->
                <div class="bg-slate-900 text-white px-5 py-3.5 rounded-2xl flex items-center gap-4 self-start md:self-center">
                    <div>
                        <span class="text-[10px] uppercase text-emerald-400 font-bold block">Token Sinkron Admin</span>
                        <span id="proctorTokenDisplay" class="text-2xl font-mono font-extrabold text-amber-400 tracking-widest">------</span>
                    </div>
                    <button onclick="copyTextToClipboard(document.getElementById('proctorTokenDisplay').innerText, 'Token berhasil disalin!')" class="px-3 py-2 rounded-xl bg-white/10 hover:bg-white/20 text-xs font-bold transition flex items-center gap-1.5">
                        <i data-lucide="copy" class="w-4 h-4"></i> Salin
                    </button>
                </div>
            </div>

            <!-- 3 MENU UTAMA PENGAWAS -->
            <div class="flex flex-wrap gap-2 bg-white p-2 rounded-2xl border border-slate-200 shadow-sm">
                <button onclick="switchProctorMenu('kehadiran')" id="procMenuBtn-kehadiran" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-amber-500 text-white transition flex items-center gap-2">
                    <i data-lucide="arrow-left-right" class="w-4 h-4"></i> 1. Pantau Siswa Keluar/Masuk &amp; Token
                </button>
                <button onclick="switchProctorMenu('pelanggaran')" id="procMenuBtn-pelanggaran" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="siren" class="w-4 h-4"></i> 2. Pelanggaran &amp; Buka Kunci Sirine
                    <span id="proctorViolationCountPill" class="px-2 py-0.5 rounded-full bg-red-100 text-red-700 text-[10px] font-extrabold">0</span>
                </button>
                <button onclick="switchProctorMenu('soal')" id="procMenuBtn-soal" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="file-plus-2" class="w-4 h-4"></i> 3. Input &amp; Upload Soal (Word / PDF)
                </button>
            </div>

            <!-- MENU 1 PENGAWAS: PANTAU SISWA KELUAR MASUK & TOKEN SINKRON -->
            <div id="proctorSection-kehadiran" class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-8 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-bold text-slate-900 text-base">Status Kehadiran &amp; Pengerjaan Siswa Real-Time</h3>
                            <p class="text-xs text-slate-500">Memantau siswa yang masuk dashboard, sedang mengerjakan soal, atau sudah selesai.</p>
                        </div>
                        <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start">
                            <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Live Sync
                        </span>
                    </div>

                    <div id="proctorStudentCards" class="space-y-3"></div>
                </div>

                <!-- Panel Token & Log Keluar Masuk Siswa Ruang -->
                <div class="lg:col-span-4 space-y-4 min-w-0">
                    <div class="bg-slate-900 text-white rounded-2xl p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <span class="text-xs font-bold uppercase tracking-wider text-amber-300">Token Ujian Aktif</span>
                            <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 text-[10px] font-bold">SINKRON ADMIN</span>
                        </div>
                        <div class="text-3xl font-mono font-extrabold tracking-widest text-white" id="proctorCardTokenBig">------</div>
                        <p class="text-[11px] text-slate-300 leading-relaxed">
                            Berikan token ini kepada peserta di ruangan saat mereka siap menekan tombol <strong>Mulai Ujian</strong> pada mata pelajaran yang telah di-assign.
                        </p>
                        <div class="pt-2 border-t border-white/10 flex justify-between text-xs">
                            <span class="text-slate-400">PIN Buka Kunci Layar:</span>
                            <span id="proctorPinDisplay" class="font-mono font-extrabold text-amber-400">123</span>
                        </div>
                    </div>

                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                        <h4 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-blue-600"></i> Log Keluar / Masuk Siswa Terbaru
                        </h4>
                        <div id="proctorInOutLogList" class="space-y-2 max-h-72 overflow-y-auto text-xs"></div>
                    </div>
                </div>
            </div>

            <!-- MENU 2 PENGAWAS: PELANGGARAN & KENDALI SIRINE -->
            <div id="proctorSection-pelanggaran" class="hidden space-y-6">
                <div class="bg-white rounded-2xl border border-red-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                        <h3 class="text-lg font-extrabold text-slate-900 flex items-center gap-2">
                            <i data-lucide="siren" class="w-5 h-5 text-red-600"></i> Monitoring Pelanggaran &amp; Layar Terkunci (Sirine)
                        </h3>
                        <p class="text-xs text-slate-500 mt-0.5">Daftar peserta yang terdeteksi keluar dari halaman ujian, berpindah tab, atau membuka aplikasi lain.</p>
                    </div>
                    <div class="bg-red-50 border border-red-200 px-4 py-2 rounded-xl text-xs text-red-900">
                        PIN Pembuka Kunci Manual: <strong id="proctorPinDisplay2" class="font-mono text-sm font-extrabold">123</strong>
                    </div>
                </div>

                <!-- Siswa Sedang Terkunci Saat Ini -->
                <div id="proctorActiveLockedBox" class="space-y-3"></div>

                <!-- Riwayat Log Pelanggaran -->
                <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-3">
                    <h4 class="font-bold text-slate-900 text-sm">Riwayat Catatan Pelanggaran Peserta</h4>
                    <div class="overflow-x-auto border border-slate-200 rounded-xl">
                        <table class="w-full text-left border-collapse text-xs sm:text-sm">
                            <thead>
                                <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                    <th class="py-3 px-4 font-bold">Waktu</th>
                                    <th class="py-3 px-4 font-bold">Peserta &amp; NISN</th>
                                    <th class="py-3 px-4 font-bold">Keterangan Pelanggaran</th>
                                </tr>
                            </thead>
                            <tbody id="proctorViolationLogsTbody" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- MENU 3 PENGAWAS: INPUT & UPLOAD SOAL WORD / PDF -->
            <div id="proctorSection-soal" class="hidden grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-5 space-y-5">
                    <!-- Upload Word / PDF Pengawas -->
                    <div class="bg-white rounded-2xl border border-amber-200 p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                <i data-lucide="file-up" class="w-4 h-4 text-amber-600"></i> Upload File Word (.docx) / PDF (.pdf)
                            </h3>
                            <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-amber-700 hover:underline">
                                Unduh Contoh Format
                            </button>
                        </div>

                        <div>
                            <label class="block text-xs font-bold text-slate-700 mb-1">Pilih Paket Mata Pelajaran Tujuan:</label>
                            <select id="proctorPackageSelector" onchange="selectProctorQuestionPackage(this.value)" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-bold bg-white">
                                <!-- Diisi via JS -->
                            </select>
                        </div>

                        <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-amber-300 rounded-2xl bg-amber-50/40 hover:bg-amber-50 cursor-pointer transition p-4 text-center">
                            <i data-lucide="upload-cloud" class="w-7 h-7 text-amber-600 mb-1"></i>
                            <span class="text-xs font-bold text-amber-900">Klik Pilih File .DOCX, .PDF, atau .TXT</span>
                            <span class="text-[10px] text-slate-500 mt-0.5">Langsung tersinkron ke Bank Soal Administrator</span>
                            <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event, 'pengawas')" class="hidden">
                        </label>
                    </div>

                    <!-- Form Input Soal Manual oleh Pengawas -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                        <h3 class="font-bold text-slate-900 text-sm mb-3">Input Soal Manual (Sinkron ke Bank Soal)</h3>
                        <form onsubmit="handleAddManualQuestion(event, 'pengawas')" class="space-y-3">
                            <div>
                                <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan *</label>
                                <textarea id="pQText" required rows="2" placeholder="Tuliskan pertanyaan soal..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-amber-500 focus:outline-none"></textarea>
                            </div>
                            <div class="space-y-2">
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">A.</span>
                                    <input type="text" id="pQOptA" required placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">B.</span>
                                    <input type="text" id="pQOptB" required placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">C.</span>
                                    <input type="text" id="pQOptC" required placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">D.</span>
                                    <input type="text" id="pQOptD" required placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">E.</span>
                                    <input type="text" id="pQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                            </div>
                            <div class="grid grid-cols-2 gap-3 pt-1">
                                <div>
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban</label>
                                    <select id="pQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-amber-800">
                                        <option value="A">Opsi A</option>
                                        <option value="B">Opsi B</option>
                                        <option value="C">Opsi C</option>
                                        <option value="D">Opsi D</option>
                                        <option value="E">Opsi E</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin</label>
                                    <input type="number" id="pQPoints" value="10" min="1" max="100" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                </div>
                            </div>
                            <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition shadow-sm">
                                + Simpan Soal ke Paket Terpilih
                            </button>
                        </form>
                    </div>
                </div>

                <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-100">
                        <div>
                            <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Soal Paket Terpilih (<span id="proctorPkgSoalCount">0</span> Soal)</h3>
                            <p class="text-[11px] text-slate-500">Tersinkronisasi langsung dengan Bank Soal Administrator</p>
                        </div>
                    </div>
                    <div id="proctorSoalListContainer" class="space-y-3 max-h-[640px] overflow-y-auto pr-1"></div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW E: HALAMAN UJIAN SISWA (ANTI-CHEAT LOCKDOWN + SIRINE)        -->
        <!-- ================================================================= -->
        <section id="view-siswa" class="min-h-screen bg-slate-100 flex flex-col no-select hidden">
            <div class="bg-slate-900 text-white px-4 sm:px-6 py-3.5 shadow-md sticky top-0 z-30">
                <div class="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center font-extrabold shrink-0">
                            <i data-lucide="shield-check" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h2 id="studentExamTitle" class="font-bold text-sm sm:text-base leading-tight">Ujian Berbasis Komputer</h2>
                            <p class="text-xs text-slate-300">
                                <strong id="studentExamName" class="text-white">-</strong> • NISN: <span id="studentExamNisn">-</span> • <span id="studentExamRoom">-</span>
                            </p>
                        </div>
                    </div>

                    <div class="flex items-center gap-2 sm:gap-3">
                        <div class="hidden sm:flex items-center bg-slate-800 border border-slate-700 rounded-xl p-1">
                            <button onclick="changeExamFontSize('sm')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perkecil Teks">A-</button>
                            <button onclick="changeExamFontSize('base')" class="px-2 py-1 rounded-lg text-xs font-bold bg-slate-700" title="Ukuran Normal">A</button>
                            <button onclick="changeExamFontSize('lg')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perbesar Teks">A+</button>
                        </div>

                        <div class="bg-slate-800 border border-slate-700 px-3.5 py-1.5 rounded-xl flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-amber-400"></i>
                            <span id="studentTimerDisplay" class="font-mono font-extrabold text-base sm:text-lg text-amber-400">00:00</span>
                        </div>

                        <button onclick="confirmSubmitExam()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-bold transition shadow">
                            Selesai Ujian
                        </button>
                    </div>
                </div>
            </div>

            <div class="max-w-7xl w-full mx-auto px-4 sm:px-6 py-6 grid grid-cols-1 lg:grid-cols-12 gap-6 flex-grow">
                <div class="lg:col-span-8 bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 shadow-sm flex flex-col justify-between">
                    <div>
                        <div class="flex flex-wrap items-center justify-between gap-2 pb-4 mb-6 border-b border-slate-100">
                            <span class="px-3.5 py-1.5 rounded-xl bg-blue-50 text-blue-700 font-extrabold text-xs sm:text-sm">
                                SOAL NO. <span id="examCurrentNum">1</span> DARI <span id="examTotalNum">1</span>
                            </span>
                            <button id="doubtToggleBtn" onclick="toggleDoubtQuestion()" class="px-3.5 py-1.5 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-bold transition flex items-center gap-1.5">
                                <i data-lucide="flag" class="w-3.5 h-3.5"></i> <span>Tandai Ragu-Ragu</span>
                            </button>
                        </div>

                        <div id="examQuestionBody" class="font-semibold text-slate-900 leading-relaxed mb-6 text-base"></div>
                        <div id="examOptionsList" class="space-y-3"></div>
                    </div>

                    <div class="flex items-center justify-between pt-8 mt-8 border-t border-slate-100">
                        <button id="btnPrevQuestion" onclick="goPrevQuestion()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs sm:text-sm font-bold transition flex items-center gap-2 disabled:opacity-40">
                            <i data-lucide="chevron-left" class="w-4 h-4"></i> Sebelumnya
                        </button>
                        <button id="btnNextQuestion" onclick="goNextQuestion()" class="px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2">
                            Selanjutnya <i data-lucide="chevron-right" class="w-4 h-4"></i>
                        </button>
                    </div>
                </div>

                <div class="lg:col-span-4 space-y-4 min-w-0">
                    <div class="bg-white rounded-3xl border border-slate-200 p-6 shadow-sm">
                        <h3 class="font-bold text-slate-900 text-sm mb-3">Navigasi Nomor Soal</h3>
                        <div id="examNumberGrid" class="grid grid-cols-5 gap-2"></div>
                        <div class="mt-5 pt-4 border-t border-slate-100 grid grid-cols-3 gap-2 text-[11px] text-slate-600">
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-blue-600 inline-block"></span> Dijawab</span>
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-amber-400 inline-block"></span> Ragu</span>
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-slate-100 border border-slate-300 inline-block"></span> Kosong</span>
                        </div>
                    </div>

                    <div class="bg-red-50 border border-red-200 rounded-2xl p-4 text-xs text-red-900 space-y-1">
                        <div class="font-bold flex items-center gap-1.5 text-red-700">
                            <i data-lucide="siren" class="w-4 h-4"></i> Proteksi Layar &amp; Sirine Aktif
                        </div>
                        <p class="leading-relaxed">
                            Dilarang berpindah tab, membuka aplikasi lain, mengecilkan layar, atau keluar halaman ujian. Pelanggaran akan <strong>mengunci layar &amp; membunyikan sirine alarm otomatis</strong>.
                        </p>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- ===================================================================== -->
    <!-- 5. MODAL LOGIN MULTI-PERAN (SISWA KE DASHBOARD, PENGAWAS, ADMIN)      -->
    <!-- ===================================================================== -->
    <div id="loginModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl border border-slate-100 relative">
            <button onclick="closeLoginModal()" class="absolute top-5 right-5 w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition">
                <i data-lucide="x" class="w-4 h-4"></i>
            </button>

            <div class="mb-6">
                <span id="loginModalBadge" class="px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold uppercase">Peserta Ujian</span>
                <h2 id="loginModalTitle" class="text-2xl font-extrabold text-slate-900 mt-2">Login Dashboard Siswa</h2>
                <p id="loginModalDesc" class="text-xs text-slate-500 mt-1">Masukkan NISN/Username dan Password sesuai Kartu Login Peserta.</p>
            </div>

            <form onsubmit="processLogin(event)" class="space-y-4">
                <input type="hidden" id="loginRoleField" value="siswa">

                <div>
                    <label id="loginUsernameLabel" class="block text-xs font-bold text-slate-700 mb-1">NISN / Username Peserta</label>
                    <input type="text" id="loginUsernameField" required placeholder="Masukkan NISN atau Username..." class="w-full px-4 py-3 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Password</label>
                    <input type="password" id="loginPasswordField" required placeholder="Masukkan password..." class="w-full px-4 py-3 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>

                <button type="submit" id="loginSubmitButton" class="w-full py-3.5 px-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-lg mt-2">
                    Masuk ke Dashboard
                </button>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 6. MODAL INPUT TOKEN SAAT SISWA SUDAH SIAP MEMULAI UJIAN              -->
    <!-- ===================================================================== -->
    <div id="studentTokenEntryModal" class="fixed inset-0 bg-slate-900/70 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl border border-slate-200 relative">
            <button onclick="closeStudentTokenModal()" class="absolute top-5 right-5 w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition">
                <i data-lucide="x" class="w-4 h-4"></i>
            </button>

            <div class="space-y-4">
                <div class="w-12 h-12 rounded-2xl bg-blue-100 text-blue-600 flex items-center justify-center">
                    <i data-lucide="key-round" class="w-6 h-6"></i>
                </div>
                <div>
                    <span class="text-[11px] font-extrabold uppercase tracking-wider text-blue-600">Konfirmasi Mulai Ujian</span>
                    <h3 id="tokenModalExamTitle" class="text-xl font-extrabold text-slate-900 mt-0.5">Mata Pelajaran</h3>
                    <p id="tokenModalExamMeta" class="text-xs text-slate-500 mt-1">Durasi: 45 Menit • Jumlah Soal: 5 Butir</p>
                </div>

                <div class="p-3.5 rounded-2xl bg-amber-50 border border-amber-200 text-xs text-amber-900 leading-relaxed">
                    <strong>Perhatian:</strong> Setelah token dikonfirmasi, waktu hitung mundur berjalan dan sistem akan mengaktifkan <strong>Mode Layar Penuh + Sirine Anti-Curang</strong>. Pastikan Anda sudah siap.
                </div>

                <form onsubmit="confirmStartExamWithToken(event)" class="space-y-4">
                    <input type="hidden" id="tokenModalPackageId">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1.5">Masukkan Token Ujian dari Pengawas *</label>
                        <input type="text" id="studentExamTokenInput" required maxlength="10" placeholder="CONTOH: CBT26X" class="w-full px-4 py-3.5 rounded-xl border-2 border-blue-400 bg-blue-50/40 font-mono font-extrabold text-center text-lg uppercase tracking-widest text-blue-950 focus:outline-none focus:ring-4 focus:ring-blue-500/30">
                    </div>
                    <div class="flex gap-2">
                        <button type="button" onclick="closeStudentTokenModal()" class="flex-1 py-3 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition">
                            Batal
                        </button>
                        <button type="submit" class="flex-1 py-3 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition shadow-md">
                            Mulai Kerjakan Soal &rarr;
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 7. MODAL EDIT AKUN & RESET PASSWORD SISWA                             -->
    <!-- ===================================================================== -->
    <div id="editStudentModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                <h3 class="text-lg font-extrabold text-slate-900">Edit Akun &amp; Password Siswa</h3>
                <button onclick="closeEditStudentModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <form onsubmit="saveEditedStudent(event)" class="space-y-3">
                <input type="hidden" id="editStuId">
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Nama Lengkap *</label>
                    <input type="text" id="editStuName" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">NISN *</label>
                    <input type="text" id="editStuNisn" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm font-mono">
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Kelas *</label>
                        <input type="text" id="editStuClass" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Ruangan *</label>
                        <input type="text" id="editStuRoom" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3 pt-1">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Username Login *</label>
                        <input type="text" id="editStuUser" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Password *</label>
                        <div class="flex gap-1.5">
                            <input type="text" id="editStuPass" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold text-blue-700">
                            <button type="button" onclick="generateRandomStudentPassInModal()" class="px-2.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold" title="Acak Password Baru">
                                <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
                            </button>
                        </div>
                    </div>
                </div>
                <div class="flex justify-end gap-2 pt-3">
                    <button type="button" onclick="closeEditStudentModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Batal</button>
                    <button type="submit" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold">Simpan Perubahan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 8. MODAL EDIT SOAL                                                    -->
    <!-- ===================================================================== -->
    <div id="editQuestionModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                <h3 class="text-lg font-extrabold text-slate-900">Edit Soal Ujian</h3>
                <button onclick="closeEditQuestionModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <form onsubmit="saveEditedQuestion(event)" class="space-y-3">
                <input type="hidden" id="editQId">
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Pertanyaan</label>
                    <textarea id="editQText" required rows="3" class="w-full p-2.5 rounded-xl border border-slate-300 text-sm"></textarea>
                </div>
                <div class="space-y-2">
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">A.</span>
                        <input type="text" id="editQOptA" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">B.</span>
                        <input type="text" id="editQOptB" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">C.</span>
                        <input type="text" id="editQOptC" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">D.</span>
                        <input type="text" id="editQOptD" required class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">E.</span>
                        <input type="text" id="editQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3 pt-1">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Kunci Jawaban</label>
                        <select id="editQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                            <option value="A">Opsi A</option>
                            <option value="B">Opsi B</option>
                            <option value="C">Opsi C</option>
                            <option value="D">Opsi D</option>
                            <option value="E">Opsi E</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Bobot Poin</label>
                        <input type="number" id="editQPoints" min="1" max="100" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                    </div>
                </div>
                <div class="flex justify-end gap-2 pt-3">
                    <button type="button" onclick="closeEditQuestionModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Batal</button>
                    <button type="submit" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold">Simpan Perubahan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 9. MODAL PANDUAN INSTALASI PERANGKAT                                  -->
    <!-- ===================================================================== -->
    <div id="installPlatformModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-100">
            <div class="flex items-center justify-between mb-3">
                <h3 id="installPlatformTitle" class="text-lg font-extrabold text-slate-900">Instalasi Aplikasi</h3>
                <button onclick="closeInstallModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <div id="installPlatformContent" class="text-xs sm:text-sm text-slate-600 space-y-3 leading-relaxed"></div>
            <div class="mt-6 flex justify-between items-center gap-2">
                <button onclick="downloadStandaloneApp()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-bold">
                    Unduh File .HTML
                </button>
                <button onclick="closeInstallModal()" class="px-5 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-bold">
                    Tutup
                </button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 10. MODAL HASIL SELESAI UJIAN SISWA                                   -->
    <!-- ===================================================================== -->
    <div id="examFinishedModal" class="fixed inset-0 bg-slate-900/75 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl text-center space-y-4">
            <div class="w-16 h-16 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center mx-auto">
                <i data-lucide="check-circle-2" class="w-9 h-9"></i>
            </div>
            <h2 class="text-2xl font-extrabold text-slate-900">Ujian Telah Selesai!</h2>
            <p class="text-xs sm:text-sm text-slate-600 leading-relaxed">
                Terima kasih telah mengerjakan ujian dengan jujur. Seluruh jawaban Anda telah dikirim secara real-time ke server Administrator.
            </p>
            <div id="studentScoreSummaryBox" class="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-1">
                <span class="text-xs text-slate-500 block">Nilai Akhir Anda</span>
                <span id="studentFinalScoreText" class="text-3xl font-extrabold text-blue-600">0</span>
                <span id="studentFinalDetailText" class="text-xs text-slate-600 block">Benar 0 • Salah 0</span>
            </div>
            <button onclick="closeExamFinishedModal()" class="w-full py-3 px-6 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm transition">
                Kembali ke Dashboard Siswa
            </button>
        </div>
    </div>

    <!-- Footer -->
    <footer id="mainFooter" class="bg-white border-t border-slate-200 py-6 mt-auto no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
            <div>
                <strong id="footerSchoolName" class="text-slate-800">CBT-PRO | Sistem Ujian Berbasis Komputer</strong> © 2026 • Siap Hosting GitHub Pages.
            </div>
            <div class="flex items-center gap-4">
                <span>Anti-Cheat Lockdown &amp; Sirine</span>
                <span>•</span>
                <span>Real-Time Multi-Mapel Sync</span>
            </div>
        </div>
    </footer>

    <!-- ===================================================================== -->
    <!-- 11. JAVASCRIPT ENGINE LENGKAP                                         -->
    <!-- ===================================================================== -->
    <script>
        // =================================================================
        // A. KONSTANTA DATABASE LOKAL (LOCALSTORAGE) & DATA AWAL PRODUKSI
        // =================================================================
        const STORAGE_KEYS = {
            SETTINGS: 'cbt_prod_settings_v2',
            PACKAGES: 'cbt_prod_packages_v2',
            STUDENTS: 'cbt_prod_students_v2',
            PROCTORS: 'cbt_prod_proctors_v2',
            LIVE_STATUS: 'cbt_prod_live_status_v2',
            LOGS: 'cbt_prod_activity_logs_v2',
            GALLERY: 'cbt_prod_gallery_v2',
            ACTIVE_SESSION: 'cbt_prod_active_session_v2',
            UI_STATE: 'cbt_prod_ui_state_v2',
            ACTIVE_EXAM: 'cbt_prod_active_exam_v2'
        };

        const DEFAULT_SETTINGS = {
            schoolName: 'CBT-PRO | Sistem Ujian Berbasis Komputer',
            schoolSubtitle: 'Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027',
            adminUsername: 'admin',
            adminPassword: 'admin123',
            examTitle: 'Ujian Akhir Semester (UAS) Berbasis Komputer',
            activeToken: 'CBT26X',
            unlockPin: '123',
            requireFullscreen: true,
            detectTabSwitch: true,
            enableSiren: true,
            showScoreToStudent: true
        };

        // Paket Soal Multi-Mata Pelajaran (Siap di-Assign atau Menunggu Assign)
        const DEFAULT_PACKAGES = [
            {
                id: 'pkg_001',
                examTitle: 'Ujian Akhir Semester (UAS) 2026',
                subjectName: 'Kompetensi Akademik & Informatika',
                scheduleDate: '2026-11-20',
                scheduleTime: '08:00 - 09:30 WITA',
                durationMinutes: 45,
                targetGroup: 'ALL',
                status: 'ASSIGNED', // ASSIGNED = Siap dikerjakan siswa dengan token; WAITING = Menunggu assign
                shuffleQuestions: true,
                shuffleOptions: true,
                questions: [
                    {
                        id: 'q_init_1',
                        question: 'Komponen perangkat keras komputer yang berfungsi sebagai unit pemroses utama seluruh instruksi logika dan aritmatika adalah...',
                        options: { A: 'RAM (Random Access Memory)', B: 'CPU (Central Processing Unit)', C: 'Solid State Drive (SSD)', D: 'Power Supply Unit', E: 'Network Interface Card' },
                        answer: 'B',
                        points: 20
                    },
                    {
                        id: 'q_init_2',
                        question: 'Protokol keamanan standar pada web yang mengenkripsi pertukaran data antara browser dan server ditandai dengan awalan...',
                        options: { A: 'HTTP://', B: 'FTP://', C: 'HTTPS://', D: 'SMTP://', E: 'TELNET://' },
                        answer: 'C',
                        points: 20
                    },
                    {
                        id: 'q_init_3',
                        question: 'Jika sebuah segitiga siku-siku memiliki panjang alas 6 cm dan tinggi 8 cm, maka panjang sisi miringnya adalah...',
                        options: { A: '10 cm', B: '12 cm', C: '14 cm', D: '15 cm', E: '16 cm' },
                        answer: 'A',
                        points: 20
                    },
                    {
                        id: 'q_init_4',
                        question: 'Teks Proklamasi Kemerdekaan Republik Indonesia dibacakan oleh Ir. Soekarno pada tanggal...',
                        options: { A: '17 Agustus 1944', B: '17 Agustus 1945', C: '18 Agustus 1945', D: '28 Oktober 1928', E: '1 Juni 1945' },
                        answer: 'B',
                        points: 20
                    },
                    {
                        id: 'q_init_5',
                        question: 'Sikap utama yang harus dijunjung tinggi oleh setiap peserta didik dalam mengikuti ujian berbasis komputer adalah...',
                        options: { A: 'Mencari celah membuka tab lain', B: 'Kejujuran dan kemandirian', C: 'Menyalin jawaban teman', D: 'Mengabaikan tata tertib ruang', E: 'Membagikan token ke luar ruang' },
                        answer: 'B',
                        points: 20
                    }
                ]
            },
            {
                id: 'pkg_002',
                examTitle: 'Ujian Akhir Semester (UAS) 2026',
                subjectName: 'Akuntansi Keuangan & Perpajakan',
                scheduleDate: '2026-11-21',
                scheduleTime: '10:00 - 11:30 WITA',
                durationMinutes: 60,
                targetGroup: 'ALL',
                status: 'WAITING',
                shuffleQuestions: true,
                shuffleOptions: true,
                questions: [
                    {
                        id: 'q_akt_1',
                        question: 'Persamaan dasar akuntansi yang benar dalam pencatatan transaksi keuangan perusahaan adalah...',
                        options: {
                            A: 'Aset = Liabilitas + Ekuitas',
                            B: 'Aset + Liabilitas = Ekuitas',
                            C: 'Liabilitas = Aset + Ekuitas',
                            D: 'Ekuitas = Aset + Pendapatan',
                            E: 'Aset = Pendapatan - Beban'
                        },
                        answer: 'A',
                        points: 50
                    },
                    {
                        id: 'q_akt_2',
                        question: 'Bukti transaksi yang digunakan untuk mencatat pembelian barang dagang secara kredit dari pemasok adalah...',
                        options: {
                            A: 'Kuitansi',
                            B: 'Nota Kontan',
                            C: 'Faktur Pembelian',
                            D: 'Memo Internal',
                            E: 'Bilyet Giro'
                        },
                        answer: 'C',
                        points: 50
                    }
                ]
            }
        ];

        const DEFAULT_STUDENTS = [
            { id: 'stu_001', name: 'Ahmad Fauzi', nisn: '0081234501', className: 'XII IPA 1', room: 'Ruang 01', username: '0081234501', password: '84921' },
            { id: 'stu_002', name: 'Siti Nurhaliza', nisn: '0081234502', className: 'XII IPA 1', room: 'Ruang 01', username: '0081234502', password: '37194' },
            { id: 'stu_003', name: 'Budi Santoso', nisn: '0081234503', className: 'XII IPS 2', room: 'Ruang 02', username: '0081234503', password: '62508' }
        ];

        const DEFAULT_PROCTORS = [
            { id: 'proc_1', name: 'Drs. Hendra Wijaya', room: 'Semua Ruangan', username: 'pengawas', password: 'pengawas123' },
            { id: 'proc_2', name: 'Rina Marlina, S.Pd', room: 'Ruang 01', username: 'pengawas_r1', password: '74125' }
        ];

        const DEFAULT_GALLERY = [
            { id: 'gal_1', title: 'Pelaksanaan Ujian Berbasis Komputer di Laboratorium 1', svgType: 'lab', url: '' },
            { id: 'gal_2', title: 'Pengawasan Ruang Ujian & Verifikasi Token Peserta', svgType: 'proctor', url: '' },
            { id: 'gal_3', title: 'Fokus & Kejujuran Akademik Siswa Saat Mengerjakan Soal', svgType: 'student', url: '' }
        ];

        const MOTIVATION_QUOTES = [
            {
                text: '"Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."',
                author: '— Pesan Akademik Sekolah'
            },
            {
                text: '"Jangan takut pada soal yang sulit; percayalah pada proses belajar yang sudah kamu lalui setiap hari. Kerjakan dengan tenang dan teliti!"',
                author: '— Motivasi Belajar CBT-PRO'
            },
            {
                text: '"Pendidikan adalah paspor menuju masa depan, karena hari esok dimiliki oleh mereka yang mempersiapkannya hari ini."',
                author: '— Malcolm X'
            },
            {
                text: '"Ujian bukan sekadar mengukur seberapa banyak yang kamu hafal, tetapi melatih integritas, fokus, dan ketangguhan dirimu."',
                author: '— Dewan Guru & Pengawas'
            }
        ];

        let currentQuoteIdx = 0;
        let showStudentPasswords = false;
        let selectedAdminPkgId = 'pkg_001';
        let selectedProctorPkgId = 'pkg_001';

        const getDB = (key, fallback) => {
            try {
                const raw = localStorage.getItem(key);
                return raw ? JSON.parse(raw) : fallback;
            } catch (e) {
                return fallback;
            }
        };
        const setDB = (key, value) => {
            localStorage.setItem(key, JSON.stringify(value));
        };

        function persistSessionState(extraUi = {}) {
            if (activeSession) {
                setDB(STORAGE_KEYS.ACTIVE_SESSION, activeSession);
            } else {
                localStorage.removeItem(STORAGE_KEYS.ACTIVE_SESSION);
            }
            const prevUi = getDB(STORAGE_KEYS.UI_STATE, {});
            setDB(STORAGE_KEYS.UI_STATE, {
                ...prevUi,
                adminMenu: currentAdminMenu,
                proctorMenu: currentProctorMenu,
                studentDashTab: currentStudentDashTab,
                ...extraUi
            });
        }

        function persistActiveExamState() {
            if (examSession && examSession.isRunning && activeSession?.role === 'siswa') {
                setDB(STORAGE_KEYS.ACTIVE_EXAM, {
                    studentId: activeSession.user.id,
                    packageId: examSession.packageId,
                    packageSubject: examSession.packageSubject,
                    packageTitle: examSession.packageTitle,
                    questions: examSession.questions,
                    answers: examSession.answers,
                    doubts: examSession.doubts,
                    currentIdx: examSession.currentIdx,
                    remainingSeconds: examSession.remainingSeconds,
                    violations: examSession.violations,
                    isLocked: examSession.isLocked,
                    savedAt: Date.now()
                });
            } else {
                localStorage.removeItem(STORAGE_KEYS.ACTIVE_EXAM);
            }
        }

        function initializeAppDatabase() {
            if (!localStorage.getItem(STORAGE_KEYS.SETTINGS)) setDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (!localStorage.getItem(STORAGE_KEYS.PACKAGES)) setDB(STORAGE_KEYS.PACKAGES, DEFAULT_PACKAGES);
            if (!localStorage.getItem(STORAGE_KEYS.STUDENTS)) setDB(STORAGE_KEYS.STUDENTS, DEFAULT_STUDENTS);
            if (!localStorage.getItem(STORAGE_KEYS.PROCTORS)) setDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);
            if (!localStorage.getItem(STORAGE_KEYS.LIVE_STATUS)) setDB(STORAGE_KEYS.LIVE_STATUS, {});
            if (!localStorage.getItem(STORAGE_KEYS.LOGS)) setDB(STORAGE_KEYS.LOGS, []);
            if (!localStorage.getItem(STORAGE_KEYS.GALLERY)) setDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);

            const pkgs = getDB(STORAGE_KEYS.PACKAGES, []);
            if (pkgs.length > 0) {
                selectedAdminPkgId = pkgs[0].id;
                selectedProctorPkgId = pkgs[0].id;
            }
        }

        // =================================================================
        // B. SISTEM LOG AKTIVITAS MASUK / KELUAR & PELANGGARAN REAL-TIME
        // =================================================================
        function recordSystemLog(role, userName, userIdentifier, roomOrClass, category, detail) {
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            logs.unshift({
                id: 'log_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
                timestamp: new Date().toLocaleString('id-ID'),
                role, // 'siswa' | 'pengawas' | 'admin'
                userName,
                userIdentifier,
                roomOrClass,
                category, // 'LOGIN' | 'LOGOUT' | 'EXAM_START' | 'EXAM_FINISH' | 'VIOLATION' | 'UNLOCK'
                detail
            });
            // Simpan maksimal 300 log terbaru agar ringan
            if (logs.length > 300) logs.length = 300;
            setDB(STORAGE_KEYS.LOGS, logs);
        }

        // =================================================================
        // C. ENGINE SUARA SIRINE ALARM OTOMATIS (WEB AUDIO API)
        // =================================================================
        let sirenAudioCtx = null;
        let sirenOscillator = null;
        let sirenGain = null;
        let sirenModInterval = null;
        let isSirenPlaying = false;

        function startSirenAlarm() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (settings.enableSiren === false) return;
            if (isSirenPlaying) return;

            try {
                const AudioCtx = window.AudioContext || window.webkitAudioContext;
                if (!AudioCtx) return;
                sirenAudioCtx = new AudioCtx();
                if (sirenAudioCtx.state === 'suspended') {
                    sirenAudioCtx.resume();
                }

                sirenOscillator = sirenAudioCtx.createOscillator();
                sirenGain = sirenAudioCtx.createGain();

                sirenOscillator.type = 'sawtooth';
                sirenOscillator.frequency.setValueAtTime(680, sirenAudioCtx.currentTime);
                sirenGain.gain.setValueAtTime(0.32, sirenAudioCtx.currentTime);

                sirenOscillator.connect(sirenGain);
                sirenGain.connect(sirenAudioCtx.destination);
                sirenOscillator.start();
                isSirenPlaying = true;

                // Efek sirine naik-turun (wee-woo)
                let high = false;
                sirenModInterval = setInterval(() => {
                    if (!sirenOscillator || !sirenAudioCtx) return;
                    const now = sirenAudioCtx.currentTime;
                    high = !high;
                    sirenOscillator.frequency.cancelScheduledValues(now);
                    sirenOscillator.frequency.linearRampToValueAtTime(high ? 1150 : 620, now + 0.35);
                }, 380);
            } catch (e) {
                console.warn('Web Audio API siren blocked:', e);
            }
        }

        function stopSirenAlarm() {
            isSirenPlaying = false;
            if (sirenModInterval) {
                clearInterval(sirenModInterval);
                sirenModInterval = null;
            }
            if (sirenOscillator) {
                try { sirenOscillator.stop(); } catch (e) {}
                try { sirenOscillator.disconnect(); } catch (e) {}
                sirenOscillator = null;
            }
            if (sirenAudioCtx) {
                try { sirenAudioCtx.close(); } catch (e) {}
                sirenAudioCtx = null;
            }
        }

        function testSirenSound() {
            startSirenAlarm();
            showToast('Menguji suara sirine peringatan selama 2 detik...', 'warning');
            setTimeout(() => {
                if (!examSession.isLocked) {
                    stopSirenAlarm();
                }
            }, 2000);
        }

        // =================================================================
        // D. STATE SESI PENGGUNA & MESIN UJIAN
        // =================================================================
        let activeSession = null; // { role: 'admin'|'pengawas'|'siswa', user: {...} }
        let currentAdminMenu = 'monitoring';
        let currentProctorMenu = 'kehadiran';
        let currentStudentDashTab = 'ujian';

        let examSession = {
            isRunning: false,
            isLocked: false,
            packageId: null,
            packageSubject: '',
            packageTitle: '',
            questions: [],
            answers: {},
            doubts: {},
            currentIdx: 0,
            remainingSeconds: 0,
            timerId: null,
            violations: 0
        };

        // =================================================================
        // E. NAVIGASI SPA & RENDER UI GLOBAL
        // =================================================================
        function navigateTo(view) {
            persistSessionState({ currentView: view });
            ['portal', 'siswa-dashboard', 'admin', 'pengawas', 'siswa'].forEach(v => {
                const el = document.getElementById(`view-${v}`);
                if (el) el.classList.add('hidden');
            });
            const target = document.getElementById(`view-${view}`);
            if (target) target.classList.remove('hidden');

            const header = document.getElementById('mainHeader');
            const footer = document.getElementById('mainFooter');
            const guestBtns = document.getElementById('headerGuestButtons');
            const authControls = document.getElementById('headerAuthControls');
            const backDashBtn = document.getElementById('headerBackDashBtn');

            if (view === 'siswa') {
                header.classList.add('hidden');
                footer.classList.add('hidden');
            } else {
                header.classList.remove('hidden');
                footer.classList.remove('hidden');
                if (activeSession) {
                    guestBtns.classList.add('hidden');
                    authControls.classList.remove('hidden');
                    authControls.classList.add('flex');

                    if (activeSession.role === 'admin') {
                        document.getElementById('headerUserBadge').innerText = 'Administrator';
                    } else if (activeSession.role === 'pengawas') {
                        document.getElementById('headerUserBadge').innerText = `Pengawas: ${activeSession.user.name}`;
                    } else {
                        document.getElementById('headerUserBadge').innerText = `Siswa: ${activeSession.user.name}`;
                    }

                    if (view === 'portal') {
                        backDashBtn.classList.remove('hidden');
                    } else {
                        backDashBtn.classList.add('hidden');
                    }
                } else {
                    guestBtns.classList.remove('hidden');
                    authControls.classList.add('hidden');
                    authControls.classList.remove('flex');
                }
            }

            syncAllViews();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function handleHeaderBrandClick() {
            if (examSession.isRunning) return;
            navigateTo('portal');
        }

        function returnToRoleDashboard() {
            if (!activeSession) return navigateTo('portal');
            if (activeSession.role === 'admin') navigateTo('admin');
            else if (activeSession.role === 'pengawas') navigateTo('pengawas');
            else navigateTo('siswa-dashboard');
        }

        function syncAllViews() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const students = getDB(STORAGE_KEYS.STUDENTS, []);

            // Header & Footer Identitas Sekolah
            document.getElementById('headerSchoolTitle').innerText = settings.schoolName;
            document.getElementById('headerSchoolSubtitle').innerText = settings.schoolSubtitle;
            document.getElementById('footerSchoolName').innerText = settings.schoolName;

            // Portal Stats
            const assignedCount = packages.filter(p => p.status === 'ASSIGNED').length;
            document.getElementById('heroExamTitle').innerText = settings.examTitle;
            document.getElementById('portalStatSoal').innerText = packages.length;
            document.getElementById('portalStatSiswa').innerText = students.length;
            document.getElementById('portalStatDurasi').innerText = assignedCount;
            renderPortalGallery();

            // Student Dashboard View
            if (activeSession?.role === 'siswa' && !examSession.isRunning) {
                renderStudentDashboard();
            }

            // Admin Views
            if (activeSession?.role === 'admin') {
                document.getElementById('sidebarTokenDisplay').innerText = settings.activeToken;
                document.getElementById('assignTokenDisplay').innerText = settings.activeToken;
                populateFilterDropdowns();
                renderAdminMonitoring();
                renderAdminPackageControls();
                renderAdminQuestionList();
                renderAdminAssignmentsTable();
                renderAdminLogsTable();
                renderAdminNilaiTable();
                renderAdminStudentAccountTable();
                renderAdminLoginCards();
                renderAdminSettingsForms();
            }

            // Pengawas Views
            if (activeSession?.role === 'pengawas') {
                document.getElementById('proctorDashboardName').innerText = activeSession.user.name;
                document.getElementById('proctorAssignedRoomBadge').innerText = `Tugas: ${activeSession.user.room}`;
                document.getElementById('proctorTokenDisplay').innerText = settings.activeToken;
                document.getElementById('proctorCardTokenBig').innerText = settings.activeToken;
                document.getElementById('proctorPinDisplay').innerText = settings.unlockPin;
                document.getElementById('proctorPinDisplay2').innerText = settings.unlockPin;
                renderProctorMonitoring();
                renderProctorViolationsView();
                renderProctorSoalView();
            }

            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        // =================================================================
        // F. TOAST, KONFIRMASI & CLIPBOARD
        // =================================================================
        function showToast(message, type = 'info') {
            const container = document.getElementById('toastContainer');
            const el = document.createElement('div');
            const palette = {
                success: 'bg-emerald-600 text-white border-emerald-500',
                error: 'bg-red-600 text-white border-red-500',
                warning: 'bg-amber-400 text-slate-950 border-amber-300',
                info: 'bg-slate-900 text-white border-slate-700'
            };
            el.className = `pointer-events-auto flex items-center gap-2.5 px-4 py-3 rounded-2xl shadow-xl border text-xs sm:text-sm font-semibold transition-all duration-300 ${palette[type] || palette.info}`;
            el.innerHTML = `<span class="flex-1">${message}</span>`;
            container.appendChild(el);
            setTimeout(() => {
                el.classList.add('opacity-0', '-translate-y-2');
                setTimeout(() => el.remove(), 300);
            }, 3500);
        }

        function showConfirmDialog(title, message, onYes) {
            const modal = document.getElementById('confirmModal');
            document.getElementById('confirmTitle').innerText = title;
            document.getElementById('confirmMessage').innerText = message;
            modal.classList.remove('hidden');
            modal.classList.add('flex');

            const okBtn = document.getElementById('confirmOkBtn');
            const cancelBtn = document.getElementById('confirmCancelBtn');

            const close = () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                okBtn.onclick = null;
                cancelBtn.onclick = null;
            };
            cancelBtn.onclick = close;
            okBtn.onclick = () => {
                close();
                onYes();
            };
        }

        function copyTextToClipboard(text, successMessage = 'Berhasil disalin!') {
            const ta = document.createElement('textarea');
            ta.value = text;
            document.body.appendChild(ta);
            ta.select();
            try {
                document.execCommand('copy');
                showToast(successMessage, 'success');
            } catch (e) {
                showToast('Gagal menyalin teks.', 'error');
            }
            document.body.removeChild(ta);
        }

        function copyPortalLink(role) {
            const base = window.location.href.split('?')[0].split('#')[0];
            const fullLink = `${base}?role=${role}`;
            const label = role === 'admin' ? 'Link Administrator' : role === 'pengawas' ? 'Link Portal Pengawas' : 'Link Portal Siswa';
            copyTextToClipboard(fullLink, `${label} disalin! Siap dibagikan.`);
        }

        // =================================================================
        // G. PORTAL UTAMA: MOTIVASI, GALERI & INSTALASI
        // =================================================================
        function nextMotivationQuote() {
            currentQuoteIdx = (currentQuoteIdx + 1) % MOTIVATION_QUOTES.length;
            const q = MOTIVATION_QUOTES[currentQuoteIdx];
            document.getElementById('motivationQuoteText').innerText = q.text;
            document.getElementById('motivationQuoteAuthor').innerText = q.author;
        }

        function getGalleryIllustrationSVG(type) {
            if (type === 'proctor') {
                return `
                <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-amber-50 to-orange-100">
                    <rect x="30" y="30" width="340" height="160" rx="16" fill="#ffffff" stroke="#fde68a" stroke-width="2"/>
                    <circle cx="110" cy="95" r="28" fill="#f59e0b"/>
                    <path d="M65 165 C65 135, 155 135, 155 165" fill="#d97706"/>
                    <rect x="180" y="65" width="160" height="90" rx="10" fill="#1e293b"/>
                    <rect x="192" y="78" width="80" height="10" rx="4" fill="#fbbf24"/>
                    <rect x="192" y="96" width="130" height="8" rx="4" fill="#64748b"/>
                    <rect x="192" y="112" width="110" height="8" rx="4" fill="#38bdf8"/>
                    <circle cx="315" cy="132" r="12" fill="#10b981"/>
                </svg>`;
            }
            if (type === 'student') {
                return `
                <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-emerald-50 to-teal-100">
                    <rect x="30" y="30" width="340" height="160" rx="16" fill="#ffffff" stroke="#a7f3d0" stroke-width="2"/>
                    <rect x="85" y="65" width="230" height="105" rx="12" fill="#0f172a"/>
                    <rect x="98" y="78" width="204" height="75" rx="6" fill="#eff6ff"/>
                    <rect x="112" y="90" width="140" height="10" rx="4" fill="#2563eb"/>
                    <rect x="112" y="108" width="170" height="8" rx="4" fill="#94a3b8"/>
                    <rect x="112" y="124" width="120" height="8" rx="4" fill="#10b981"/>
                    <rect x="65" y="170" width="270" height="10" rx="5" fill="#64748b"/>
                </svg>`;
            }
            return `
            <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-blue-50 to-indigo-100">
                <rect x="25" y="25" width="350" height="170" rx="16" fill="#ffffff" stroke="#bfdbfe" stroke-width="2"/>
                <rect x="55" y="60" width="120" height="80" rx="8" fill="#1e3a8a"/>
                <rect x="63" y="68" width="104" height="60" rx="4" fill="#dbeafe"/>
                <rect x="95" y="140" width="40" height="15" fill="#64748b"/>
                <rect x="75" y="155" width="80" height="6" rx="3" fill="#475569"/>
                <rect x="225" y="60" width="120" height="80" rx="8" fill="#1e3a8a"/>
                <rect x="233" y="68" width="104" height="60" rx="4" fill="#dbeafe"/>
                <rect x="265" y="140" width="40" height="15" fill="#64748b"/>
                <rect x="245" y="155" width="80" height="6" rx="3" fill="#475569"/>
                <circle cx="200" cy="45" r="10" fill="#10b981"/>
            </svg>`;
        }

        function renderPortalGallery() {
            const gallery = getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
            const grid = document.getElementById('portalGalleryGrid');
            if (!grid) return;

            grid.innerHTML = gallery.map(item => {
                const mediaHtml = item.url
                    ? `<img src="${item.url}" alt="${item.title}" onerror="this.outerHTML=\`${getGalleryIllustrationSVG(item.svgType || 'lab').replace(/"/g, '&quot;')}\`" class="w-full h-44 object-cover">`
                    : getGalleryIllustrationSVG(item.svgType || 'lab');

                return `
                    <div class="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm hover:shadow-md transition flex flex-col">
                        ${mediaHtml}
                        <div class="p-4 flex-1 flex flex-col justify-between">
                            <p class="text-xs sm:text-sm font-bold text-slate-800 leading-snug">${item.title}</p>
                            <span class="text-[11px] text-slate-400 mt-2 inline-flex items-center gap-1">
                                <i data-lucide="camera" class="w-3 h-3"></i> Dokumentasi CBT Sekolah
                            </span>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function openInstallModal(platform) {
            const modal = document.getElementById('installPlatformModal');
            const title = document.getElementById('installPlatformTitle');
            const content = document.getElementById('installPlatformContent');

            if (platform === 'android') {
                title.innerText = 'Instalasi Aplikasi di HP / Tablet Android';
                content.innerHTML = `
                    <p>Pasang <strong>CBT-PRO</strong> di perangkat Android agar berjalan dalam mode layar penuh tanpa bilah navigasi browser:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-xs text-slate-700">
                        <li>Buka link portal ujian ini melalui <strong>Google Chrome</strong> di HP Android Anda.</li>
                        <li>Ketuk ikon menu <strong>Titik Tiga (⋮)</strong> di pojok kanan atas layar.</li>
                        <li>Pilih <strong>"Tambahkan ke Layar Utama"</strong> atau <strong>"Install Aplikasi"</strong>.</li>
                        <li>Konfirmasi dengan menekan <strong>Install / Tambah</strong>. Ikon aplikasi CBT akan muncul di layar utama HP Anda.</li>
                    </ol>
                `;
            } else if (platform === 'ios') {
                title.innerText = 'Instalasi Aplikasi di iPhone / iPad (iOS)';
                content.innerHTML = `
                    <p>Untuk memasang aplikasi ujian pada perangkat iPhone atau iPad:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-xs text-slate-700">
                        <li>Buka halaman ujian ini menggunakan browser bawaan <strong>Safari</strong>.</li>
                        <li>Ketuk tombol <strong>Bagikan (Share)</strong> di bagian bawah Safari.</li>
                        <li>Gulir menu ke bawah, lalu ketuk <strong>"Tambah ke Layar Utama" (Add to Home Screen)</strong>.</li>
                        <li>Ketuk <strong>Tambah</strong> di pojok kanan atas. Aplikasi siap dibuka langsung dari Home Screen.</li>
                    </ol>
                `;
            } else {
                title.innerText = 'Instalasi Aplikasi di Komputer (Windows / Mac)';
                content.innerHTML = `
                    <p>Untuk menjalankan aplikasi ujian secara mandiri di komputer laboratorium atau laptop:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-xs text-slate-700">
                        <li>Buka link aplikasi ini menggunakan <strong>Google Chrome</strong> atau <strong>Microsoft Edge</strong>.</li>
                        <li>Klik ikon <strong>Install App</strong> di ujung kanan bilah alamat URL.</li>
                        <li>Atau klik tombol <strong>Unduh File .HTML</strong> di bawah untuk menjalankan aplikasi secara lokal tanpa internet.</li>
                    </ol>
                `;
            }

            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeInstallModal() {
            const modal = document.getElementById('installPlatformModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function downloadStandaloneApp() {
            const htmlDoc = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
            const blob = new Blob([htmlDoc], { type: 'text/html;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'CBT-PRO-Aplikasi-Ujian.html';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('File aplikasi mandiri (.HTML) berhasil diunduh!', 'success');
        }

        // =================================================================
        // H. AUTENTIKASI 3 PERAN (SISWA KE DASHBOARD, PENGAWAS, ADMIN)
        // =================================================================
        function openLoginModal(role) {
            const modal = document.getElementById('loginModal');
            const roleField = document.getElementById('loginRoleField');
            const badge = document.getElementById('loginModalBadge');
            const title = document.getElementById('loginModalTitle');
            const desc = document.getElementById('loginModalDesc');
            const userLabel = document.getElementById('loginUsernameLabel');
            const submitBtn = document.getElementById('loginSubmitButton');

            document.getElementById('loginUsernameField').value = '';
            document.getElementById('loginPasswordField').value = '';
            roleField.value = role;

            if (role === 'admin') {
                badge.className = 'px-3 py-1 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold uppercase';
                badge.innerText = 'Administrator';
                title.innerText = 'Login Administrator';
                desc.innerText = 'Masukkan Username dan Password Administrator sekolah.';
                userLabel.innerText = 'Username Administrator';
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-sm transition shadow-lg mt-2';
            } else if (role === 'pengawas') {
                badge.className = 'px-3 py-1 rounded-full bg-amber-100 text-amber-800 text-xs font-bold uppercase';
                badge.innerText = 'Pengawas Ruang';
                title.innerText = 'Login Dashboard Pengawas';
                desc.innerText = 'Masukkan Username dan Password Pengawas Ruang yang terdaftar.';
                userLabel.innerText = 'Username Pengawas';
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-amber-500 hover:bg-amber-600 text-white font-bold text-sm transition shadow-lg mt-2';
            } else {
                badge.className = 'px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold uppercase';
                badge.innerText = 'Peserta Ujian';
                title.innerText = 'Login Dashboard Siswa';
                desc.innerText = 'Masukkan NISN/Username dan Password Kartu Login. Token dimasukkan saat Anda siap mulai ujian.';
                userLabel.innerText = 'NISN / Username Peserta';
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-lg mt-2';
            }

            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeLoginModal() {
            const modal = document.getElementById('loginModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function processLogin(e) {
            e.preventDefault();
            const role = document.getElementById('loginRoleField').value;
            const userInput = document.getElementById('loginUsernameField').value.trim();
            const passInput = document.getElementById('loginPasswordField').value.trim();
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);

            if (role === 'admin') {
                const validUser = settings.adminUsername || 'admin';
                const validPass = settings.adminPassword || 'admin123';
                if (userInput === validUser && passInput === validPass) {
                    activeSession = { role: 'admin', user: { name: 'Administrator', username: validUser } };
                    persistSessionState({ currentView: 'admin' });
                    recordSystemLog('admin', 'Administrator', validUser, 'Pusat', 'LOGIN', 'Administrator masuk ke panel kendali utama.');
                    closeLoginModal();
                    showToast('Berhasil masuk sebagai Administrator.', 'success');
                    navigateTo('admin');
                } else {
                    showToast('Username atau Password Administrator tidak valid!', 'error');
                }
                return;
            }

            if (role === 'pengawas') {
                const proctors = getDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);
                const foundProctor = proctors.find(p => p.username === userInput && p.password === passInput);
                if (foundProctor) {
                    activeSession = { role: 'pengawas', user: foundProctor };
                    persistSessionState({ currentView: 'pengawas' });
                    recordSystemLog('pengawas', foundProctor.name, foundProctor.username, foundProctor.room, 'LOGIN', `Pengawas masuk ke Dashboard Pengawas (${foundProctor.room}).`);
                    closeLoginModal();
                    showToast(`Selamat bertugas, ${foundProctor.name}!`, 'success');
                    navigateTo('pengawas');
                } else {
                    showToast('Username atau Password Pengawas tidak terdaftar!', 'error');
                }
                return;
            }

            if (role === 'siswa') {
                const students = getDB(STORAGE_KEYS.STUDENTS, []);
                const student = students.find(s =>
                    (s.username.toLowerCase() === userInput.toLowerCase() || s.nisn === userInput) &&
                    s.password === passInput
                );

                if (!student) {
                    showToast('NISN/Username atau Password tidak cocok dengan Kartu Login!', 'error');
                    return;
                }

                // Update live status: Siswa masuk ke Dashboard Siswa
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                const prev = liveMap[student.id] || {};
                liveMap[student.id] = {
                    ...prev,
                    status: prev.status === 'Selesai' ? 'Selesai' : 'Di Dashboard',
                    lastLoginAt: new Date().toLocaleTimeString('id-ID'),
                    updatedAt: new Date().toLocaleTimeString('id-ID')
                };
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

                activeSession = { role: 'siswa', user: student };
                persistSessionState({ currentView: 'siswa-dashboard' });
                recordSystemLog('siswa', student.name, student.nisn, `${student.className} • ${student.room}`, 'LOGIN', 'Siswa masuk ke Dashboard Peserta Ujian.');
                closeLoginModal();
                showToast(`Selamat datang, ${student.name}! Silakan pilih mata pelajaran ujian.`, 'success');
                navigateTo('siswa-dashboard');
            }
        }

        function logoutSession() {
            if (activeSession) {
                if (activeSession.role === 'siswa') {
                    const s = activeSession.user;
                    const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                    if (liveMap[s.id] && liveMap[s.id].status === 'Di Dashboard') {
                        liveMap[s.id].status = 'Keluar / Offline';
                        liveMap[s.id].updatedAt = new Date().toLocaleTimeString('id-ID');
                        setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                    }
                    recordSystemLog('siswa', s.name, s.nisn, `${s.className} • ${s.room}`, 'LOGOUT', 'Siswa keluar (logout) dari Dashboard Siswa.');
                } else if (activeSession.role === 'pengawas') {
                    const p = activeSession.user;
                    recordSystemLog('pengawas', p.name, p.username, p.room, 'LOGOUT', `Pengawas keluar (logout) dari sistem.`);
                } else if (activeSession.role === 'admin') {
                    recordSystemLog('admin', 'Administrator', activeSession.user.username || 'admin', 'Pusat', 'LOGOUT', 'Administrator keluar (logout) dari sistem.');
                }
            }
            activeSession = null;
            localStorage.removeItem(STORAGE_KEYS.ACTIVE_SESSION);
            localStorage.removeItem(STORAGE_KEYS.UI_STATE);
            localStorage.removeItem(STORAGE_KEYS.ACTIVE_EXAM);
            showToast('Anda telah keluar dari sistem.', 'info');
            navigateTo('portal');
        }

        // =================================================================
        // I. DASHBOARD SISWA: PROFIL KARTU, PILIH MAPEL & INPUT TOKEN
        // =================================================================
        function switchStudentDashTab(tabName) {
            currentStudentDashTab = tabName;
            persistSessionState();
            ['ujian', 'profil', 'panduan'].forEach(t => {
                document.getElementById(`stuDashTab-${t}`).classList.add('hidden');
                const btn = document.getElementById(`stuTabBtn-${t}`);
                btn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2';
            });
            document.getElementById(`stuDashTab-${tabName}`).classList.remove('hidden');
            const activeBtn = document.getElementById(`stuTabBtn-${tabName}`);
            activeBtn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-blue-600 text-white transition flex items-center gap-2';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function renderStudentDashboard() {
            if (!activeSession || activeSession.role !== 'siswa') return;
            const s = activeSession.user;
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const myLive = liveMap[s.id] || {};
            const completedPkgs = myLive.completedPackages || {};

            // Header & Sidebar Profil
            document.getElementById('stuDashWelcomeName').innerText = `Selamat Datang, ${s.name}`;
            document.getElementById('stuDashWelcomeSub').innerText = `NISN: ${s.nisn} • Kelas ${s.className} • ${s.room}`;
            document.getElementById('stuSideName').innerText = s.name;
            document.getElementById('stuSideNisn').innerText = s.nisn;
            document.getElementById('stuSideClass').innerText = s.className;
            document.getElementById('stuSideRoom').innerText = s.room;

            // Tab Profil Lengkap Sesuai Kartu Login
            document.getElementById('stuCardSchoolName').innerText = settings.schoolName;
            document.getElementById('stuProfName').innerText = s.name;
            document.getElementById('stuProfNisn').innerText = s.nisn;
            document.getElementById('stuProfClass').innerText = s.className;
            document.getElementById('stuProfRoom').innerText = s.room;
            document.getElementById('stuProfUser').innerText = s.username;
            document.getElementById('stuProfLoginTime').innerText = `Aktif (Login pukul ${myLive.lastLoginAt || new Date().toLocaleTimeString('id-ID')})`;

            // Filter paket yang relevan untuk kelas/ruang siswa
            const relevantPackages = packages.filter(pkg => {
                if (!pkg.targetGroup || pkg.targetGroup === 'ALL') return true;
                return pkg.targetGroup === s.className || pkg.targetGroup === s.room;
            });

            document.getElementById('stuAssignedCountBadge').innerText = `${relevantPackages.length} Mata Pelajaran`;

            const listContainer = document.getElementById('studentAssignmentsList');
            if (relevantPackages.length === 0) {
                listContainer.innerHTML = `
                    <div class="bg-white rounded-2xl border border-slate-200 p-8 text-center space-y-2">
                        <div class="w-12 h-12 rounded-2xl bg-slate-100 text-slate-400 flex items-center justify-center mx-auto">
                            <i data-lucide="calendar-clock" class="w-6 h-6"></i>
                        </div>
                        <h4 class="font-bold text-slate-800 text-sm">Belum Ada Penugasan Ujian</h4>
                        <p class="text-xs text-slate-500">Saat ini belum ada mata pelajaran yang dijadwalkan untuk kelas Anda. Silakan menunggu informasi dari Pengawas / Administrator.</p>
                    </div>
                `;
            } else {
                listContainer.innerHTML = relevantPackages.map(pkg => {
                    const isDone = !!completedPkgs[pkg.id];
                    const isAssigned = pkg.status === 'ASSIGNED';
                    const qCount = (pkg.questions || []).length;

                    let statusBadge = '';
                    let actionButton = '';

                    if (isDone) {
                        const res = completedPkgs[pkg.id];
                        statusBadge = `<span class="px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold flex items-center gap-1"><i data-lucide="check-circle2" class="w-3.5 h-3.5"></i> Selesai Dikerjakan</span>`;
                        actionButton = `
                            <div class="text-right">
                                <span class="text-xs font-bold text-emerald-700 block">Terkirim (${res.finishedAt || '-'})</span>
                                ${settings.showScoreToStudent ? `<span class="text-xs font-extrabold text-blue-700">Nilai: ${res.score}</span>` : ''}
                            </div>
                        `;
                    } else if (isAssigned && qCount > 0) {
                        statusBadge = `<span class="px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-blue-600 animate-ping"></span> Siap Dikerjakan (Assigned)</span>`;
                        actionButton = `
                            <button onclick="openStudentTokenModal('${pkg.id}')" class="px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition shadow-sm flex items-center gap-2">
                                <i data-lucide="key-round" class="w-4 h-4"></i> Mulai Ujian &amp; Input Token
                            </button>
                        `;
                    } else {
                        statusBadge = `<span class="px-3 py-1 rounded-full bg-amber-100 text-amber-800 text-xs font-bold flex items-center gap-1"><i data-lucide="clock" class="w-3.5 h-3.5"></i> Terjadwal • Menunggu Assign</span>`;
                        actionButton = `
                            <button disabled class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-400 text-xs font-bold cursor-not-allowed flex items-center gap-1.5">
                                <i data-lucide="lock" class="w-3.5 h-3.5"></i> Menunggu Dibuka Admin
                            </button>
                        `;
                    }

                    return `
                        <div class="bg-white rounded-2xl border ${isAssigned && !isDone ? 'border-blue-300 shadow-sm' : 'border-slate-200'} p-5 transition space-y-4">
                            <div class="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                                <div class="space-y-1">
                                    <span class="text-[11px] font-bold uppercase tracking-wider text-slate-400">${pkg.examTitle}</span>
                                    <h4 class="text-base sm:text-lg font-extrabold text-slate-900">${pkg.subjectName}</h4>
                                    <div class="flex flex-wrap items-center gap-3 text-xs text-slate-600 pt-1">
                                        <span class="inline-flex items-center gap-1"><i data-lucide="calendar" class="w-3.5 h-3.5 text-blue-600"></i> ${pkg.scheduleDate || 'Sesuai Jadwal'}</span>
                                        <span class="inline-flex items-center gap-1"><i data-lucide="clock" class="w-3.5 h-3.5 text-amber-600"></i> ${pkg.scheduleTime || '-'}</span>
                                        <span class="inline-flex items-center gap-1"><i data-lucide="timer" class="w-3.5 h-3.5 text-indigo-600"></i> Durasi: <strong>${pkg.durationMinutes} Menit</strong></span>
                                        <span class="inline-flex items-center gap-1"><i data-lucide="file-question" class="w-3.5 h-3.5 text-emerald-600"></i> <strong>${qCount} Soal</strong></span>
                                    </div>
                                </div>
                                <div class="shrink-0">
                                    ${statusBadge}
                                </div>
                            </div>
                            <div class="pt-3 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                <span class="text-[11px] text-slate-500">
                                    Target Peserta: <strong>${pkg.targetGroup === 'ALL' ? 'Semua Kelas & Ruangan' : pkg.targetGroup}</strong>
                                </span>
                                <div>
                                    ${actionButton}
                                </div>
                            </div>
                        </div>
                    `;
                }).join('');
            }

            // Render Riwayat di Tab Profil
            const historyContainer = document.getElementById('stuMyHistoryList');
            const doneKeys = Object.keys(completedPkgs);
            if (doneKeys.length === 0) {
                historyContainer.innerHTML = `<div class="p-6 text-center text-slate-400 bg-slate-50 rounded-2xl">Belum ada ujian yang diselesaikan.</div>`;
            } else {
                historyContainer.innerHTML = doneKeys.map(k => {
                    const item = completedPkgs[k];
                    return `
                        <div class="p-3.5 rounded-xl bg-emerald-50/60 border border-emerald-200 flex items-center justify-between">
                            <div>
                                <div class="font-bold text-slate-900">${item.subjectName}</div>
                                <div class="text-[11px] text-slate-500">Selesai: ${item.finishedAt} • Benar ${item.correct} / Salah ${item.wrong}</div>
                            </div>
                            ${settings.showScoreToStudent ? `<span class="px-3 py-1 rounded-lg bg-emerald-600 text-white font-extrabold text-sm">${item.score}</span>` : `<span class="text-xs font-bold text-emerald-700">Terkirim</span>`}
                        </div>
                    `;
                }).join('');
            }
        }

        function openStudentTokenModal(pkgId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg) return;

            if (pkg.status !== 'ASSIGNED') {
                showToast('Mata pelajaran ini masih menunggu di-assign oleh Administrator.', 'warning');
                return;
            }
            if (!pkg.questions || pkg.questions.length === 0) {
                showToast('Soal pada paket mata pelajaran ini belum tersedia.', 'warning');
                return;
            }

            document.getElementById('tokenModalPackageId').value = pkg.id;
            document.getElementById('tokenModalExamTitle').innerText = pkg.subjectName;
            document.getElementById('tokenModalExamMeta').innerText = `${pkg.examTitle} • Durasi: ${pkg.durationMinutes} Menit • ${pkg.questions.length} Soal`;
            document.getElementById('studentExamTokenInput').value = '';

            const modal = document.getElementById('studentTokenEntryModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
            setTimeout(() => document.getElementById('studentExamTokenInput').focus(), 100);
        }

        function closeStudentTokenModal() {
            const modal = document.getElementById('studentTokenEntryModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function confirmStartExamWithToken(e) {
            e.preventDefault();
            const pkgId = document.getElementById('tokenModalPackageId').value;
            const enteredToken = document.getElementById('studentExamTokenInput').value.trim().toUpperCase();
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);

            if (!enteredToken || enteredToken !== settings.activeToken.toUpperCase()) {
                showToast('Token Ujian tidak cocok! Periksa kembali token aktif dari Pengawas Ruangan.', 'error');
                return;
            }

            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg || pkg.status !== 'ASSIGNED') {
                showToast('Paket ujian sudah tidak aktif atau belum di-assign.', 'error');
                closeStudentTokenModal();
                return;
            }

            closeStudentTokenModal();
            beginStudentExam(activeSession.user, pkg);
        }

        // =================================================================
        // J. DASHBOARD ADMINISTRATOR: 7 MENU TERPISAH
        // =================================================================
        function switchAdminMenu(menuKey) {
            currentAdminMenu = menuKey;
            persistSessionState();
            const menus = ['monitoring', 'soal', 'assign', 'logs', 'nilai', 'kartu', 'setelan'];

            menus.forEach(m => {
                document.getElementById(`adminSection-${m}`).classList.add('hidden');
                const btn = document.getElementById(`adminNav-${m}`);
                btn.className = 'admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap';
            });

            document.getElementById(`adminSection-${menuKey}`).classList.remove('hidden');
            const activeBtn = document.getElementById(`adminNav-${menuKey}`);
            activeBtn.className = 'admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap';

            syncAllViews();
        }

        function populateFilterDropdowns() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const classes = [...new Set(students.map(s => s.className).filter(Boolean))];
            const rooms = [...new Set(students.map(s => s.room).filter(Boolean))];

            const monFilter = document.getElementById('monRoomFilter');
            const currentMonVal = monFilter.value;
            monFilter.innerHTML = `<option value="">Semua Kelas & Ruangan</option>` +
                classes.map(c => `<option value="class:${c}">Kelas: ${c}</option>`).join('') +
                rooms.map(r => `<option value="room:${r}">Ruangan: ${r}</option>`).join('');
            monFilter.value = currentMonVal;

            const nClass = document.getElementById('nilaiClassFilter');
            const curNClass = nClass.value;
            nClass.innerHTML = `<option value="">Semua Kelas</option>` + classes.map(c => `<option value="${c}">${c}</option>`).join('');
            nClass.value = curNClass;

            const nRoom = document.getElementById('nilaiRoomFilter');
            const curNRoom = nRoom.value;
            nRoom.innerHTML = `<option value="">Semua Ruangan</option>` + rooms.map(r => `<option value="${r}">${r}</option>`).join('');
            nRoom.value = curNRoom;

            const assignTarget = document.getElementById('assignTargetGroup');
            const curAssignTarget = assignTarget.value;
            assignTarget.innerHTML = `<option value="ALL">Semua Kelas & Ruangan</option>` +
                classes.map(c => `<option value="${c}">Khusus Kelas: ${c}</option>`).join('') +
                rooms.map(r => `<option value="${r}">Khusus Ruangan: ${r}</option>`).join('');
            if (curAssignTarget) assignTarget.value = curAssignTarget;
        }

        // --- MENU 1 ADMIN: MONITORING SISWA ---
        function renderAdminMonitoring() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const search = (document.getElementById('monSearchInput')?.value || '').toLowerCase();
            const filterVal = document.getElementById('monRoomFilter')?.value || '';

            let activeCount = 0, doneCount = 0, lockedCount = 0;
            students.forEach(s => {
                const st = liveMap[s.id]?.status;
                if (st === 'Mengerjakan' || st === 'Di Dashboard') activeCount++;
                else if (st === 'Selesai') doneCount++;
                else if (st === 'TERKUNCI') lockedCount++;
            });

            document.getElementById('monStatTotal').innerText = students.length;
            document.getElementById('monStatActive').innerText = activeCount;
            document.getElementById('monStatDone').innerText = doneCount;
            document.getElementById('monStatLocked').innerText = lockedCount;

            const filtered = students.filter(s => {
                const matchText = s.name.toLowerCase().includes(search) || s.nisn.includes(search);
                let matchGroup = true;
                if (filterVal.startsWith('class:')) matchGroup = s.className === filterVal.replace('class:', '');
                if (filterVal.startsWith('room:')) matchGroup = s.room === filterVal.replace('room:', '');
                return matchText && matchGroup;
            });

            const tbody = document.getElementById('adminMonitoringTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="py-8 text-center text-slate-400 text-xs">Tidak ada data peserta yang sesuai.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(s => {
                const st = liveMap[s.id] || { status: 'Belum Masuk', progress: '-', currentSubject: '-', violations: 0 };
                let badge = `<span class="px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 text-[11px] font-bold">Belum Masuk</span>`;
                if (st.status === 'Di Dashboard') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-indigo-100 text-indigo-700 text-[11px] font-bold">Di Dashboard</span>`;
                } else if (st.status === 'Mengerjakan') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-blue-100 text-blue-700 text-[11px] font-bold animate-pulse">Mengerjakan</span>`;
                } else if (st.status === 'TERKUNCI') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-red-600 text-white text-[11px] font-bold animate-bounce">TERKUNCI + SIRINE</span>`;
                } else if (st.status === 'Selesai') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">Selesai</span>`;
                } else if (st.status === 'Keluar / Offline') {
                    badge = `<span class="px-2.5 py-1 rounded-full bg-slate-200 text-slate-700 text-[11px] font-bold">Keluar / Offline</span>`;
                }

                return `
                    <tr class="hover:bg-slate-50 transition ${st.status === 'TERKUNCI' ? 'bg-red-50/70' : ''}">
                        <td class="py-3 px-4">
                            <div class="font-bold text-slate-900">${s.name}</div>
                            <div class="text-[11px] font-mono text-slate-500">NISN: ${s.nisn}</div>
                        </td>
                        <td class="py-3 px-4 text-xs text-slate-700">${s.className} • <strong>${s.room}</strong></td>
                        <td class="py-3 px-4 text-xs font-semibold text-slate-700">${st.currentSubject || '-'}</td>
                        <td class="py-3 px-4">${badge}</td>
                        <td class="py-3 px-4 font-mono text-xs font-semibold">${st.progress || '-'}</td>
                        <td class="py-3 px-4">
                            <span class="px-2 py-0.5 rounded-full text-xs font-bold ${st.violations > 0 ? 'bg-red-100 text-red-700' : 'bg-slate-100 text-slate-600'}">
                                ${st.violations || 0}x
                            </span>
                        </td>
                        <td class="py-3 px-4 text-right space-x-1">
                            ${st.status === 'TERKUNCI' ? `
                                <button onclick="remoteUnlockStudentById('${s.id}')" class="px-2.5 py-1.5 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-bold transition">
                                    Matikan Sirine &amp; Buka Kunci
                                </button>
                            ` : ''}
                            <button onclick="resetStudentExamSession('${s.id}')" class="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-bold transition">
                                Reset Sesi
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function remoteUnlockStudentById(studentId) {
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const stu = students.find(s => s.id === studentId);
            if (liveMap[studentId]) {
                liveMap[studentId].status = 'Mengerjakan';
                liveMap[studentId].remoteUnlocked = true;
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

                const actorName = activeSession?.role === 'pengawas' ? `Pengawas (${activeSession.user.name})` : 'Administrator';
                recordSystemLog(
                    activeSession?.role || 'admin',
                    stu ? stu.name : studentId,
                    stu ? stu.nisn : '-',
                    stu ? `${stu.className} • ${stu.room}` : '-',
                    'UNLOCK',
                    `Kunci layar & sirine dibuka jarak jauh oleh ${actorName}.`
                );

                syncAllViews();
                showToast('Sirine dimatikan & kunci layar peserta dibuka jarak jauh!', 'success');
            }
        }

        function resetStudentExamSession(studentId) {
            showConfirmDialog('Reset Sesi Ujian Peserta?', 'Status pengerjaan dan riwayat selesai peserta ini akan direset agar dapat mengikuti ujian kembali.', () => {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                delete liveMap[studentId];
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                syncAllViews();
                showToast('Sesi peserta berhasil direset.', 'info');
            });
        }

        // --- MENU 2 ADMIN: BANK SOAL PER MATA PELAJARAN ---
        function renderAdminPackageControls() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const selector = document.getElementById('adminPackageSelector');
            if (!selector) return;

            if (packages.length === 0) {
                selector.innerHTML = `<option value="">-- Belum Ada Paket Mata Pelajaran --</option>`;
                return;
            }

            if (!packages.some(p => p.id === selectedAdminPkgId)) {
                selectedAdminPkgId = packages[0].id;
            }

            selector.innerHTML = packages.map(p => `
                <option value="${p.id}">${p.subjectName} (${(p.questions || []).length} Soal) — [${p.status === 'ASSIGNED' ? 'DI-ASSIGN' : 'MENUNGGU ASSIGN'}]</option>
            `).join('');
            selector.value = selectedAdminPkgId;
        }

        function selectAdminQuestionPackage(pkgId) {
            selectedAdminPkgId = pkgId;
            renderAdminQuestionList();
        }

        function createNewSubjectPackage(e) {
            e.preventDefault();
            const subjectName = document.getElementById('newPkgSubjectName').value.trim();
            if (!subjectName) return;

            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const newPkg = {
                id: 'pkg_' + Date.now(),
                examTitle: settings.examTitle || 'Ujian Berbasis Komputer',
                subjectName,
                scheduleDate: new Date().toISOString().slice(0, 10),
                scheduleTime: '08:00 - 09:30 WITA',
                durationMinutes: 45,
                targetGroup: 'ALL',
                status: 'WAITING',
                shuffleQuestions: true,
                shuffleOptions: true,
                questions: []
            };
            packages.push(newPkg);
            setDB(STORAGE_KEYS.PACKAGES, packages);
            selectedAdminPkgId = newPkg.id;
            selectedProctorPkgId = newPkg.id;
            document.getElementById('newPkgSubjectName').value = '';
            syncAllViews();
            showToast(`Paket Mata Pelajaran "${subjectName}" berhasil dibuat! Silakan isi atau upload soal.`, 'success');
        }

        function handleAddManualQuestion(e, sourceRole = 'admin') {
            e.preventDefault();
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const targetPkgId = sourceRole === 'pengawas' ? selectedProctorPkgId : selectedAdminPkgId;
            const pkgIdx = packages.findIndex(p => p.id === targetPkgId);

            if (pkgIdx === -1) {
                showToast('Pilih atau buat paket mata pelajaran terlebih dahulu!', 'warning');
                return;
            }

            const prefix = sourceRole === 'pengawas' ? 'pQ' : 'mQ';
            const optE = document.getElementById(`${prefix}OptE`).value.trim();
            const options = {
                A: document.getElementById(`${prefix}OptA`).value.trim(),
                B: document.getElementById(`${prefix}OptB`).value.trim(),
                C: document.getElementById(`${prefix}OptC`).value.trim(),
                D: document.getElementById(`${prefix}OptD`).value.trim()
            };
            if (optE) options.E = optE;

            const ans = document.getElementById(`${prefix}Answer`).value;
            if (ans === 'E' && !optE) {
                showToast('Opsi E kosong, silakan isi Opsi E atau pilih kunci A-D.', 'warning');
                return;
            }

            packages[pkgIdx].questions.push({
                id: 'q_' + Date.now(),
                question: document.getElementById(`${prefix}Text`).value.trim(),
                options,
                answer: ans,
                points: parseInt(document.getElementById(`${prefix}Points`).value, 10) || 10
            });

            setDB(STORAGE_KEYS.PACKAGES, packages);
            e.target.reset();
            document.getElementById(`${prefix}Points`).value = 10;
            syncAllViews();
            showToast(`Soal berhasil ditambahkan ke ${packages[pkgIdx].subjectName}!`, 'success');
        }

        function parseExtractedTextToQuestions(rawText) {
            const normalized = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
            const parsedList = [];

            const lines = normalized.split('\n').map(l => l.trim()).filter(Boolean);
            lines.forEach((line, idx) => {
                if (line.includes('|')) {
                    const parts = line.split('|').map(p => p.trim());
                    if (parts.length >= 6) {
                        const hasE = parts.length >= 7 && ['A', 'B', 'C', 'D', 'E'].includes(parts[6].toUpperCase());
                        const ans = hasE ? parts[6].toUpperCase() : parts[5].toUpperCase();
                        if (['A', 'B', 'C', 'D', 'E'].includes(ans)) {
                            const opts = { A: parts[1], B: parts[2], C: parts[3], D: parts[4] };
                            if (hasE && parts[5]) opts.E = parts[5];
                            parsedList.push({
                                id: `q_imp_${Date.now()}_${idx}`,
                                question: parts[0],
                                options: opts,
                                answer: ans,
                                points: 10
                            });
                        }
                    }
                }
            });

            if (parsedList.length > 0) return parsedList;

            const blocks = normalized.split(/(?=\n?\s*\d+[\.\)]\s+)/g).map(b => b.trim()).filter(Boolean);
            blocks.forEach((block, idx) => {
                const keyMatch = block.match(/(?:Kunci|Jawaban)\s*[:=\-]\s*([A-Ea-e])/i);
                const optAMatch = block.match(/(?:^|\n|\s)A[\.\)]\s*([\s\S]*?)(?=(?:\n|\s)B[\.\)]\s*)/i);
                const optBMatch = block.match(/(?:^|\n|\s)B[\.\)]\s*([\s\S]*?)(?=(?:\n|\s)C[\.\)]\s*)/i);
                const optCMatch = block.match(/(?:^|\n|\s)C[\.\)]\s*([\s\S]*?)(?=(?:\n|\s)D[\.\)]\s*)/i);
                const optDMatch = block.match(/(?:^|\n|\s)D[\.\)]\s*([\s\S]*?)(?=(?:(?:\n|\s)E[\.\)]\s*|(?:Kunci|Jawaban)\s*[:=\-]|$))/i);
                const optEMatch = block.match(/(?:^|\n|\s)E[\.\)]\s*([\s\S]*?)(?=(?:(?:Kunci|Jawaban)\s*[:=\-]|$))/i);

                if (optAMatch && optBMatch && optCMatch && optDMatch) {
                    const firstOptIdx = block.search(/(?:^|\n|\s)A[\.\)]\s*/i);
                    let qText = block.substring(0, firstOptIdx).replace(/^\d+[\.\)]\s*/, '').trim();
                    const opts = {
                        A: optAMatch[1].trim(),
                        B: optBMatch[1].trim(),
                        C: optCMatch[1].trim(),
                        D: optDMatch[1].trim()
                    };
                    if (optEMatch && optEMatch[1].trim()) {
                        opts.E = optEMatch[1].trim();
                    }
                    const ans = keyMatch ? keyMatch[1].toUpperCase() : 'A';
                    if (qText) {
                        parsedList.push({
                            id: `q_doc_${Date.now()}_${idx}`,
                            question: qText,
                            options: opts,
                            answer: ans,
                            points: 10
                        });
                    }
                }
            });

            return parsedList;
        }

        async function handleDocumentQuestionUpload(event, sourceRole = 'admin') {
            const file = event.target.files[0];
            if (!file) return;

            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const targetPkgId = sourceRole === 'pengawas' ? selectedProctorPkgId : selectedAdminPkgId;
            const pkgIdx = packages.findIndex(p => p.id === targetPkgId);

            if (pkgIdx === -1) {
                showToast('Pilih paket mata pelajaran tujuan terlebih dahulu!', 'warning');
                event.target.value = '';
                return;
            }

            const statusBox = document.getElementById('uploadParseStatus');
            if (statusBox && sourceRole === 'admin') {
                statusBox.classList.remove('hidden');
                statusBox.innerText = `Sedang membaca file "${file.name}"...`;
            }

            const ext = file.name.split('.').pop().toLowerCase();
            let extractedText = '';

            try {
                if (ext === 'docx') {
                    if (typeof mammoth === 'undefined') {
                        throw new Error('Library Mammoth.js tidak tersedia (mode offline). Gunakan file .TXT atau input manual.');
                    }
                    const arrayBuffer = await file.arrayBuffer();
                    const result = await mammoth.extractRawText({ arrayBuffer });
                    extractedText = result.value;
                } else if (ext === 'pdf') {
                    if (typeof pdfjsLib === 'undefined') {
                        throw new Error('Library PDF.js tidak tersedia (mode offline). Gunakan file .TXT atau input manual.');
                    }
                    const arrayBuffer = await file.arrayBuffer();
                    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                    for (let p = 1; p <= pdf.numPages; p++) {
                        const page = await pdf.getPage(p);
                        const content = await page.getTextContent();
                        const pageText = content.items.map(item => item.str).join(' ');
                        extractedText += '\n' + pageText;
                    }
                } else {
                    extractedText = await file.text();
                }

                const newQuestions = parseExtractedTextToQuestions(extractedText);
                if (newQuestions.length > 0) {
                    packages[pkgIdx].questions = [...(packages[pkgIdx].questions || []), ...newQuestions];
                    setDB(STORAGE_KEYS.PACKAGES, packages);
                    syncAllViews();
                    showToast(`Berhasil mengimpor ${newQuestions.length} soal ke "${packages[pkgIdx].subjectName}"!`, 'success');
                    if (statusBox && sourceRole === 'admin') {
                        statusBox.innerText = `Sukses: ${newQuestions.length} soal ditambahkan ke ${packages[pkgIdx].subjectName}.`;
                    }
                } else {
                    showToast('Tidak ditemukan pola soal yang valid dalam file. Unduh contoh format untuk panduan.', 'warning');
                    if (statusBox && sourceRole === 'admin') {
                        statusBox.innerText = 'Format soal tidak dikenali. Pastikan menggunakan penomoran 1., opsi A-D, dan Kunci: A.';
                    }
                }
            } catch (err) {
                showToast(err.message || 'Gagal memproses file dokumen.', 'error');
                if (statusBox && sourceRole === 'admin') statusBox.innerText = `Error: ${err.message}`;
            } finally {
                event.target.value = '';
            }
        }

        function downloadSampleQuestionFormat() {
            const sample = [
                "1. Ibu kota negara Indonesia adalah...",
                "A. Bandung",
                "B. Jakarta",
                "C. Surabaya",
                "D. Medan",
                "E. Makassar",
                "Kunci: B",
                "",
                "2. Hasil dari 25 x 4 adalah...",
                "A. 80",
                "B. 90",
                "C. 100",
                "D. 110",
                "Kunci: C"
            ].join('\n');

            const blob = new Blob([sample], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Contoh_Format_Soal_CBT.txt';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('Contoh format soal berhasil diunduh!', 'success');
        }

        function renderAdminQuestionList() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === selectedAdminPkgId);
            const questions = pkg ? (pkg.questions || []) : [];
            const container = document.getElementById('adminSoalListContainer');
            document.getElementById('adminBankSoalCount').innerText = questions.length;

            if (questions.length === 0) {
                container.innerHTML = `<div class="text-center py-12 text-xs text-slate-400">Belum ada soal pada paket mata pelajaran ini.</div>`;
                return;
            }

            container.innerHTML = questions.map((q, idx) => {
                const optionKeys = Object.keys(q.options);
                return `
                    <div class="p-4 rounded-2xl border border-slate-200 bg-slate-50/60 hover:bg-white transition space-y-2.5">
                        <div class="flex items-start justify-between gap-3">
                            <div class="text-xs sm:text-sm font-bold text-slate-900">
                                <span class="text-indigo-600 mr-1">${idx + 1}.</span> ${q.question}
                                <span class="ml-2 px-2 py-0.5 rounded bg-slate-200 text-slate-700 text-[10px] font-semibold">${q.points || 10} Poin</span>
                            </div>
                            <div class="flex items-center gap-1.5 shrink-0">
                                <button onclick="openEditQuestionModal('${pkg.id}', '${q.id}')" class="px-2.5 py-1.5 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-800 text-xs font-bold transition flex items-center gap-1">
                                    <i data-lucide="edit-3" class="w-3.5 h-3.5"></i> Edit
                                </button>
                                <button onclick="deleteSingleQuestion('${pkg.id}', '${q.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 transition" title="Hapus Soal">
                                    <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-1.5 text-xs">
                            ${optionKeys.map(k => `
                                <div class="px-2.5 py-1.5 rounded-lg border ${q.answer === k ? 'bg-emerald-100 border-emerald-300 text-emerald-900 font-bold' : 'bg-white border-slate-200 text-slate-600'}">
                                    <strong>${k}.</strong> ${q.options[k]}
                                </div>
                            `).join('')}
                        </div>
                    </div>
                `;
            }).join('');
        }

        let editingPkgId = null;

        function openEditQuestionModal(pkgId, qId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg) return;
            const q = (pkg.questions || []).find(item => item.id === qId);
            if (!q) return;

            editingPkgId = pkgId;
            document.getElementById('editQId').value = q.id;
            document.getElementById('editQText').value = q.question;
            document.getElementById('editQOptA').value = q.options.A || '';
            document.getElementById('editQOptB').value = q.options.B || '';
            document.getElementById('editQOptC').value = q.options.C || '';
            document.getElementById('editQOptD').value = q.options.D || '';
            document.getElementById('editQOptE').value = q.options.E || '';
            document.getElementById('editQAnswer').value = q.answer;
            document.getElementById('editQPoints').value = q.points || 10;

            const modal = document.getElementById('editQuestionModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeEditQuestionModal() {
            const modal = document.getElementById('editQuestionModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function saveEditedQuestion(e) {
            e.preventDefault();
            const qId = document.getElementById('editQId').value;
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkgIdx = packages.findIndex(p => p.id === editingPkgId);
            if (pkgIdx === -1) return;

            const idx = packages[pkgIdx].questions.findIndex(item => item.id === qId);
            if (idx === -1) return;

            const optE = document.getElementById('editQOptE').value.trim();
            const options = {
                A: document.getElementById('editQOptA').value.trim(),
                B: document.getElementById('editQOptB').value.trim(),
                C: document.getElementById('editQOptC').value.trim(),
                D: document.getElementById('editQOptD').value.trim()
            };
            if (optE) options.E = optE;

            const ans = document.getElementById('editQAnswer').value;
            if (ans === 'E' && !optE) {
                showToast('Opsi E kosong, pilih kunci A-D atau isi teks Opsi E.', 'warning');
                return;
            }

            packages[pkgIdx].questions[idx] = {
                ...packages[pkgIdx].questions[idx],
                question: document.getElementById('editQText').value.trim(),
                options,
                answer: ans,
                points: parseInt(document.getElementById('editQPoints').value, 10) || 10
            };

            setDB(STORAGE_KEYS.PACKAGES, packages);
            closeEditQuestionModal();
            syncAllViews();
            showToast('Perubahan soal berhasil disimpan!', 'success');
        }

        function deleteSingleQuestion(pkgId, qId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkgIdx = packages.findIndex(p => p.id === pkgId);
            if (pkgIdx === -1) return;
            packages[pkgIdx].questions = packages[pkgIdx].questions.filter(q => q.id !== qId);
            setDB(STORAGE_KEYS.PACKAGES, packages);
            syncAllViews();
            showToast('Soal dihapus dari paket.', 'info');
        }

        function confirmClearPackageQuestions() {
            showConfirmDialog('Kosongkan Soal Paket Ini?', 'Seluruh soal dalam paket mata pelajaran terpilih akan dihapus permanen.', () => {
                const packages = getDB(STORAGE_KEYS.PACKAGES, []);
                const pkgIdx = packages.findIndex(p => p.id === selectedAdminPkgId);
                if (pkgIdx !== -1) {
                    packages[pkgIdx].questions = [];
                    setDB(STORAGE_KEYS.PACKAGES, packages);
                    syncAllViews();
                    showToast('Soal dalam paket berhasil dikosongkan.', 'info');
                }
            });
        }

        // --- MENU 3 ADMIN: ASSIGNMENT MATA PELAJARAN, JADWAL & TOKEN ---
        function generateRandomToken() {
            const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
            let token = '';
            for (let i = 0; i < 6; i++) token += chars.charAt(Math.floor(Math.random() * chars.length));
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.activeToken = token;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast(`Token baru berhasil dibuat & disinkronkan: ${token}`, 'success');
        }

        function saveManualToken() {
            const val = document.getElementById('customTokenInput').value.trim().toUpperCase();
            if (val.length < 4) {
                showToast('Token minimal terdiri dari 4 karakter!', 'warning');
                return;
            }
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.activeToken = val;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            document.getElementById('customTokenInput').value = '';
            syncAllViews();
            showToast(`Token berhasil diubah & disinkronkan: ${val}`, 'success');
        }

        function loadSelectedPackageToAssignForm() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const sel = document.getElementById('assignPackageSelect');
            if (!sel) return;
            const pkg = packages.find(p => p.id === sel.value) || packages[0];
            if (!pkg) return;

            document.getElementById('assignExamTitle').value = pkg.examTitle || '';
            document.getElementById('assignSubject').value = pkg.subjectName || '';
            document.getElementById('assignScheduleDate').value = pkg.scheduleDate || new Date().toISOString().slice(0, 10);
            document.getElementById('assignScheduleTime').value = pkg.scheduleTime || '08:00 - 09:30 WITA';
            document.getElementById('assignDuration').value = pkg.durationMinutes || 45;
            document.getElementById('assignTargetGroup').value = pkg.targetGroup || 'ALL';
            document.getElementById('assignStatusSelect').value = pkg.status || 'ASSIGNED';
            document.getElementById('assignShuffleQuestions').checked = !!pkg.shuffleQuestions;
            document.getElementById('assignShuffleOptions').checked = !!pkg.shuffleOptions;
        }

        function savePackageAssignment(e) {
            e.preventDefault();
            const pkgId = document.getElementById('assignPackageSelect').value;
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const idx = packages.findIndex(p => p.id === pkgId);
            if (idx === -1) {
                showToast('Pilih paket soal terlebih dahulu!', 'warning');
                return;
            }

            const newStatus = document.getElementById('assignStatusSelect').value;
            if (newStatus === 'ASSIGNED' && (!packages[idx].questions || packages[idx].questions.length === 0)) {
                showToast('Paket soal masih kosong! Tambahkan soal terlebih dahulu sebelum di-assign ke siswa.', 'warning');
                return;
            }

            packages[idx] = {
                ...packages[idx],
                examTitle: document.getElementById('assignExamTitle').value.trim(),
                subjectName: document.getElementById('assignSubject').value.trim(),
                scheduleDate: document.getElementById('assignScheduleDate').value,
                scheduleTime: document.getElementById('assignScheduleTime').value.trim(),
                durationMinutes: parseInt(document.getElementById('assignDuration').value, 10) || 45,
                targetGroup: document.getElementById('assignTargetGroup').value,
                status: newStatus,
                shuffleQuestions: document.getElementById('assignShuffleQuestions').checked,
                shuffleOptions: document.getElementById('assignShuffleOptions').checked
            };

            setDB(STORAGE_KEYS.PACKAGES, packages);
            syncAllViews();
            showToast(`Penugasan "${packages[idx].subjectName}" berhasil disimpan & disinkronkan ke Dashboard Siswa!`, 'success');
        }

        function togglePackageAssignStatus(pkgId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const idx = packages.findIndex(p => p.id === pkgId);
            if (idx === -1) return;

            const nextStatus = packages[idx].status === 'ASSIGNED' ? 'WAITING' : 'ASSIGNED';
            if (nextStatus === 'ASSIGNED' && (!packages[idx].questions || packages[idx].questions.length === 0)) {
                showToast('Tidak dapat meng-assign karena paket soal ini belum memiliki butir soal.', 'warning');
                return;
            }

            packages[idx].status = nextStatus;
            setDB(STORAGE_KEYS.PACKAGES, packages);
            syncAllViews();
            showToast(
                nextStatus === 'ASSIGNED'
                    ? `Mata pelajaran "${packages[idx].subjectName}" kini DI-ASSIGN & siap dikerjakan siswa!`
                    : `Mata pelajaran "${packages[idx].subjectName}" diubah ke status Menunggu Assign.`,
                'info'
            );
        }

        function renderAdminAssignmentsTable() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const sel = document.getElementById('assignPackageSelect');
            if (sel) {
                const prevVal = sel.value;
                sel.innerHTML = packages.map(p => `
                    <option value="${p.id}">${p.subjectName} (${(p.questions || []).length} Soal)</option>
                `).join('');
                if (prevVal && packages.some(p => p.id === prevVal)) {
                    sel.value = prevVal;
                } else if (packages.length > 0) {
                    sel.value = packages[0].id;
                    loadSelectedPackageToAssignForm();
                }
            }

            const tbody = document.getElementById('adminAssignmentsTableBody');
            if (!tbody) return;

            if (packages.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="py-8 text-center text-slate-400 text-xs">Belum ada paket soal mata pelajaran. Buat paket terlebih dahulu di Menu 2.</td></tr>`;
                return;
            }

            tbody.innerHTML = packages.map(p => {
                const isAssigned = p.status === 'ASSIGNED';
                return `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-4">
                            <div class="font-bold text-slate-900">${p.subjectName}</div>
                            <div class="text-[11px] text-slate-500">${p.examTitle}</div>
                        </td>
                        <td class="py-3 px-4 text-xs">
                            <div class="font-semibold text-slate-800">${p.scheduleDate || '-'}</div>
                            <div class="text-slate-500">${p.scheduleTime || '-'} (${p.durationMinutes} mnt)</div>
                        </td>
                        <td class="py-3 px-4 text-xs font-bold text-indigo-700">${p.targetGroup === 'ALL' ? 'Semua Kelas' : p.targetGroup}</td>
                        <td class="py-3 px-4 font-bold text-xs">${(p.questions || []).length} Soal</td>
                        <td class="py-3 px-4">
                            ${isAssigned
                                ? `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">DI-ASSIGN (Aktif)</span>`
                                : `<span class="px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">Menunggu Assign</span>`
                            }
                        </td>
                        <td class="py-3 px-4 text-right space-x-1">
                            <button onclick="togglePackageAssignStatus('${p.id}')" class="px-3 py-1.5 rounded-lg ${isAssigned ? 'bg-amber-100 hover:bg-amber-200 text-amber-900' : 'bg-blue-600 hover:bg-blue-700 text-white'} text-[11px] font-bold transition">
                                ${isAssigned ? 'Tunda / Set Menunggu' : 'Assign Sekarang'}
                            </button>
                            <button onclick="document.getElementById('assignPackageSelect').value='${p.id}'; loadSelectedPackageToAssignForm();" class="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-bold transition">
                                Edit Jadwal
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        // --- MENU 4 ADMIN: LOG MASUK/KELUAR SISWA & PENGAWAS ---
        function renderAdminLogsTable() {
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            const search = (document.getElementById('logSearchInput')?.value || '').toLowerCase();
            const roleFilter = document.getElementById('logRoleFilter')?.value || '';
            const typeFilter = document.getElementById('logTypeFilter')?.value || '';

            const filtered = logs.filter(item => {
                const matchText =
                    (item.userName || '').toLowerCase().includes(search) ||
                    (item.userIdentifier || '').toLowerCase().includes(search) ||
                    (item.detail || '').toLowerCase().includes(search);
                const matchRole = !roleFilter || item.role === roleFilter;
                const matchType = !typeFilter || item.category === typeFilter;
                return matchText && matchRole && matchType;
            });

            const tbody = document.getElementById('adminLogsTableBody');
            if (!tbody) return;

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="py-8 text-center text-slate-400 text-xs">Belum ada catatan aktivitas masuk/keluar yang sesuai filter.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(item => {
                let roleBadge = `<span class="px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-[10px] font-bold uppercase">${item.role}</span>`;
                if (item.role === 'siswa') roleBadge = `<span class="px-2 py-0.5 rounded bg-blue-100 text-blue-700 text-[10px] font-bold uppercase">Siswa</span>`;
                else if (item.role === 'pengawas') roleBadge = `<span class="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold uppercase">Pengawas</span>`;
                else if (item.role === 'admin') roleBadge = `<span class="px-2 py-0.5 rounded bg-indigo-100 text-indigo-700 text-[10px] font-bold uppercase">Admin</span>`;

                let catBadge = `<span class="px-2 py-0.5 rounded-full bg-slate-100 text-slate-700 text-[10px] font-bold">${item.category}</span>`;
                if (item.category === 'LOGIN') catBadge = `<span class="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold">MASUK (LOGIN)</span>`;
                else if (item.category === 'LOGOUT') catBadge = `<span class="px-2 py-0.5 rounded-full bg-slate-200 text-slate-700 text-[10px] font-bold">KELUAR (LOGOUT)</span>`;
                else if (item.category === 'EXAM_START') catBadge = `<span class="px-2 py-0.5 rounded-full bg-blue-100 text-blue-700 text-[10px] font-bold">MULAI UJIAN</span>`;
                else if (item.category === 'EXAM_FINISH') catBadge = `<span class="px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-700 text-[10px] font-bold">SELESAI UJIAN</span>`;
                else if (item.category === 'VIOLATION') catBadge = `<span class="px-2 py-0.5 rounded-full bg-red-600 text-white text-[10px] font-bold">PELANGGARAN + SIRINE</span>`;
                else if (item.category === 'UNLOCK') catBadge = `<span class="px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 text-[10px] font-bold">BUKA KUNCI</span>`;

                return `
                    <tr class="hover:bg-slate-50 transition ${item.category === 'VIOLATION' ? 'bg-red-50/50' : ''}">
                        <td class="py-2.5 px-4 font-mono text-xs text-slate-600 whitespace-nowrap">${item.timestamp}</td>
                        <td class="py-2.5 px-4">${roleBadge}</td>
                        <td class="py-2.5 px-4">
                            <div class="font-bold text-slate-900">${item.userName}</div>
                            <div class="text-[11px] font-mono text-slate-500">${item.userIdentifier} • ${item.roomOrClass}</div>
                        </td>
                        <td class="py-2.5 px-4">${catBadge}</td>
                        <td class="py-2.5 px-4 text-xs text-slate-700">${item.detail}</td>
                    </tr>
                `;
            }).join('');
        }

        function exportActivityLogsExcel() {
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            if (logs.length === 0) {
                showToast('Belum ada data log aktivitas untuk diunduh.', 'warning');
                return;
            }
            const rows = logs.map((l, idx) => ({
                'No': idx + 1,
                'Waktu': l.timestamp,
                'Peran': l.role.toUpperCase(),
                'Nama Pengguna': l.userName,
                'ID / NISN': l.userIdentifier,
                'Kelas / Ruangan': l.roomOrClass,
                'Kategori': l.category,
                'Keterangan': l.detail
            }));

            const fileName = `Log_Aktivitas_CBT_${new Date().toISOString().slice(0, 10)}`;
            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Log Masuk Keluar');
                XLSX.writeFile(wb, `${fileName}.xlsx`);
                showToast('Log aktivitas masuk/keluar (.XLSX) berhasil diunduh!', 'success');
            }
        }

        function confirmClearActivityLogs() {
            showConfirmDialog('Bersihkan Seluruh Log Aktivitas?', 'Semua riwayat log masuk/keluar dan pelanggaran akan dihapus permanen.', () => {
                setDB(STORAGE_KEYS.LOGS, []);
                renderAdminLogsTable();
                showToast('Log aktivitas berhasil dibersihkan.', 'info');
            });
        }

        // --- MENU 5 ADMIN: DATA SISWA & DOWNLOAD NILAI EXCEL (.XLSX) ---
        function renderAdminNilaiTable() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const search = (document.getElementById('nilaiSearchInput')?.value || '').toLowerCase();
            const classFilter = document.getElementById('nilaiClassFilter')?.value || '';
            const roomFilter = document.getElementById('nilaiRoomFilter')?.value || '';

            const filtered = students.filter(s => {
                const matchText = s.name.toLowerCase().includes(search) || s.nisn.includes(search);
                const matchClass = !classFilter || s.className === classFilter;
                const matchRoom = !roomFilter || s.room === roomFilter;
                return matchText && matchClass && matchRoom;
            });

            const tbody = document.getElementById('adminNilaiTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="10" class="py-8 text-center text-slate-400 text-xs">Tidak ada data peserta yang ditemukan.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map((s, idx) => {
                const st = liveMap[s.id] || { status: 'Belum Masuk', currentSubject: '-', correct: '-', wrong: '-', score: '-', violations: 0 };
                return `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-3 text-slate-500 font-semibold">${idx + 1}</td>
                        <td class="py-3 px-3 font-bold text-slate-900">${s.name}</td>
                        <td class="py-3 px-3 font-mono text-xs text-slate-600">${s.nisn}</td>
                        <td class="py-3 px-3 text-xs">${s.className} • <strong>${s.room}</strong></td>
                        <td class="py-3 px-3 text-xs font-semibold text-slate-700">${st.currentSubject || '-'}</td>
                        <td class="py-3 px-3">
                            <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${st.status === 'Selesai' ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'}">
                                ${st.status}
                            </span>
                        </td>
                        <td class="py-3 px-3 font-bold text-emerald-700">${st.correct !== undefined ? st.correct : '-'}</td>
                        <td class="py-3 px-3 font-bold text-red-600">${st.wrong !== undefined ? st.wrong : '-'}</td>
                        <td class="py-3 px-3">
                            <span class="px-2.5 py-1 rounded-lg font-extrabold text-xs ${typeof st.score === 'number' ? 'bg-blue-100 text-blue-800' : 'bg-slate-100 text-slate-400'}">
                                ${st.score !== undefined ? st.score : '-'}
                            </span>
                        </td>
                        <td class="py-3 px-3 text-xs font-bold ${st.violations > 0 ? 'text-red-600' : 'text-slate-500'}">${st.violations || 0}x</td>
                    </tr>
                `;
            }).join('');
        }

        function downloadExcelResults() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});

            const rows = students.map((s, idx) => {
                const st = liveMap[s.id] || {};
                return {
                    'No': idx + 1,
                    'Nama Peserta': s.name,
                    'NISN': s.nisn,
                    'Kelas': s.className,
                    'Ruangan': s.room,
                    'Mata Pelajaran': st.currentSubject || '-',
                    'Status Ujian': st.status || 'Belum Masuk',
                    'Jumlah Benar': st.correct !== undefined ? st.correct : 0,
                    'Jumlah Salah': st.wrong !== undefined ? st.wrong : 0,
                    'Nilai Akhir': st.score !== undefined ? st.score : 0,
                    'Pelanggaran': st.violations || 0,
                    'Waktu Selesai': st.finishedAt || '-'
                };
            });

            const fileName = `Rekap_Nilai_CBT_${new Date().toISOString().slice(0, 10)}`;

            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Hasil Nilai CBT');
                XLSX.writeFile(wb, `${fileName}.xlsx`);
                showToast('File rekap nilai Excel (.XLSX) berhasil diunduh!', 'success');
            }
        }

        // --- MENU 6 ADMIN: AKUN SISWA & KARTU LOGIN ---
        function toggleShowStudentPasswords() {
            showStudentPasswords = !showStudentPasswords;
            const btn = document.getElementById('btnTogglePassVis');
            btn.innerHTML = showStudentPasswords
                ? `<i data-lucide="eye-off" class="w-3.5 h-3.5"></i> <span>Sembunyikan Password</span>`
                : `<i data-lucide="eye" class="w-3.5 h-3.5"></i> <span>Tampilkan Password</span>`;
            renderAdminStudentAccountTable();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function renderAdminStudentAccountTable() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const search = (document.getElementById('accountSearchInput')?.value || '').toLowerCase();
            document.getElementById('studentAccountCount').innerText = students.length;

            const filtered = students.filter(s =>
                s.name.toLowerCase().includes(search) ||
                s.nisn.toLowerCase().includes(search) ||
                s.className.toLowerCase().includes(search) ||
                s.room.toLowerCase().includes(search)
            );

            const tbody = document.getElementById('adminStudentAccountTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="py-8 text-center text-slate-400 text-xs">Tidak ada akun siswa yang ditemukan.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(s => `
                <tr class="hover:bg-slate-50 transition">
                    <td class="py-2.5 px-3">
                        <div class="font-bold text-slate-900">${s.name}</div>
                        <div class="text-[11px] font-mono text-slate-500">NISN: ${s.nisn}</div>
                    </td>
                    <td class="py-2.5 px-3 text-slate-700">${s.className} • <strong>${s.room}</strong></td>
                    <td class="py-2.5 px-3 font-mono font-bold text-slate-900">${s.username}</td>
                    <td class="py-2.5 px-3 font-mono font-extrabold text-blue-700">
                        ${showStudentPasswords ? s.password : '•••••'}
                    </td>
                    <td class="py-2.5 px-3 text-right space-x-1">
                        <button onclick="openEditStudentModal('${s.id}')" class="px-2.5 py-1 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-800 text-[11px] font-bold transition">
                            Edit / Reset Pass
                        </button>
                        <button onclick="deleteSingleStudent('${s.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 transition" title="Hapus Akun Siswa">
                            <i data-lucide="trash-2" class="w-3.5 h-3.5 inline"></i>
                        </button>
                    </td>
                </tr>
            `).join('');
        }

        function openEditStudentModal(stuId) {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const s = students.find(item => item.id === stuId);
            if (!s) return;

            document.getElementById('editStuId').value = s.id;
            document.getElementById('editStuName').value = s.name;
            document.getElementById('editStuNisn').value = s.nisn;
            document.getElementById('editStuClass').value = s.className;
            document.getElementById('editStuRoom').value = s.room;
            document.getElementById('editStuUser').value = s.username;
            document.getElementById('editStuPass').value = s.password;

            const modal = document.getElementById('editStudentModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function closeEditStudentModal() {
            const modal = document.getElementById('editStudentModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function generateRandomStudentPassInModal() {
            document.getElementById('editStuPass').value = String(Math.floor(10000 + Math.random() * 90000));
        }

        function saveEditedStudent(e) {
            e.preventDefault();
            const stuId = document.getElementById('editStuId').value;
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const idx = students.findIndex(item => item.id === stuId);
            if (idx === -1) return;

            students[idx] = {
                ...students[idx],
                name: document.getElementById('editStuName').value.trim(),
                nisn: document.getElementById('editStuNisn').value.trim(),
                className: document.getElementById('editStuClass').value.trim(),
                room: document.getElementById('editStuRoom').value.trim(),
                username: document.getElementById('editStuUser').value.trim(),
                password: document.getElementById('editStuPass').value.trim()
            };

            setDB(STORAGE_KEYS.STUDENTS, students);
            closeEditStudentModal();
            syncAllViews();
            showToast('Data akun & password siswa berhasil diperbarui!', 'success');
        }

        function exportStudentAccountsExcel() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            if (students.length === 0) {
                showToast('Belum ada data akun siswa untuk diunduh.', 'warning');
                return;
            }
            const rows = students.map((s, idx) => ({
                'No': idx + 1,
                'Nama Lengkap': s.name,
                'NISN': s.nisn,
                'Kelas': s.className,
                'Ruangan': s.room,
                'Username Login': s.username,
                'Password Login': s.password
            }));

            const fileName = `Daftar_Akun_Login_Siswa_${new Date().toISOString().slice(0, 10)}`;
            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(rows);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Akun Login Siswa');
                XLSX.writeFile(wb, `${fileName}.xlsx`);
                showToast('Daftar akun & password siswa (.XLSX) berhasil diunduh!', 'success');
            }
        }

        function downloadStudentExcelTemplate() {
            const templateData = [
                { Nama: 'Nadia Putri', NISN: '0081234504', Kelas: 'XII IPA 1', Ruangan: 'Ruang 01' },
                { Nama: 'Rizky Pratama', NISN: '0081234505', Kelas: 'XII IPA 2', Ruangan: 'Ruang 02' }
            ];
            if (typeof XLSX !== 'undefined') {
                const ws = XLSX.utils.json_to_sheet(templateData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Template Siswa');
                XLSX.writeFile(wb, 'Template_Import_Siswa_CBT.xlsx');
                showToast('Template Excel (.XLSX) berhasil diunduh!', 'success');
            }
        }

        function handleStudentExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(firstSheet);

                    const students = getDB(STORAGE_KEYS.STUDENTS, []);
                    let added = 0;

                    rows.forEach((row, idx) => {
                        const name = String(row.Nama || row.NAMA || row.Name || '').trim();
                        const nisn = String(row.NISN || row.nisn || row.Username || '').trim();
                        const className = String(row.Kelas || row.KELAS || row.Class || 'Umum').trim();
                        const room = String(row.Ruangan || row.RUANGAN || row.Ruang || 'Ruang 01').trim();

                        if (name && nisn && !students.some(st => st.nisn === nisn)) {
                            const randomPass = row.Password ? String(row.Password).trim() : String(Math.floor(10000 + Math.random() * 90000));
                            students.push({
                                id: `stu_imp_${Date.now()}_${idx}`,
                                name,
                                nisn,
                                className,
                                room,
                                username: nisn,
                                password: randomPass
                            });
                            added++;
                        }
                    });

                    setDB(STORAGE_KEYS.STUDENTS, students);
                    syncAllViews();
                    showToast(`Berhasil mengimpor ${added} akun siswa & membuat kartu login!`, 'success');
                } catch (err) {
                    showToast('Gagal membaca file Excel. Pastikan format kolom sesuai template.', 'error');
                } finally {
                    event.target.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function handleAddManualStudent(e) {
            e.preventDefault();
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const nisn = document.getElementById('kStuNisn').value.trim();

            if (students.some(st => st.nisn === nisn)) {
                showToast('NISN sudah terdaftar pada kartu login!', 'warning');
                return;
            }

            const username = document.getElementById('kStuUser').value.trim() || nisn;
            const password = document.getElementById('kStuPass').value.trim() || String(Math.floor(10000 + Math.random() * 90000));

            students.push({
                id: 'stu_' + Date.now(),
                name: document.getElementById('kStuName').value.trim(),
                nisn,
                className: document.getElementById('kStuClass').value.trim(),
                room: document.getElementById('kStuRoom').value.trim(),
                username,
                password
            });

            setDB(STORAGE_KEYS.STUDENTS, students);
            e.target.reset();
            syncAllViews();
            showToast(`Kartu login untuk ${nisn} berhasil dibuat!`, 'success');
        }

        function deleteSingleStudent(stuId) {
            const students = getDB(STORAGE_KEYS.STUDENTS, []).filter(s => s.id !== stuId);
            setDB(STORAGE_KEYS.STUDENTS, students);
            syncAllViews();
            showToast('Akun peserta dihapus.', 'info');
        }

        function confirmClearAllStudents() {
            showConfirmDialog('Kosongkan Seluruh Daftar Siswa?', 'Semua akun peserta dan kartu login saat ini akan dihapus agar siap diisi dengan data peserta sekolah Anda.', () => {
                setDB(STORAGE_KEYS.STUDENTS, []);
                setDB(STORAGE_KEYS.LIVE_STATUS, {});
                syncAllViews();
                showToast('Seluruh data siswa berhasil dikosongkan.', 'info');
            });
        }

        function generateMiniQRSVG(codeText) {
            let hash = 0;
            for (let i = 0; i < codeText.length; i++) hash = ((hash << 5) - hash) + codeText.charCodeAt(i);
            let rects = '';
            for (let r = 0; r < 6; r++) {
                for (let c = 0; c < 6; c++) {
                    const bit = ((hash >> ((r * 6 + c) % 24)) & 1) || ((r === 0 || r === 5) && (c === 0 || c === 5));
                    if (bit) rects += `<rect x="${4 + c * 6}" y="${4 + r * 6}" width="5" height="5" fill="#0f172a"/>`;
                }
            }
            return `<svg viewBox="0 0 44 44" class="w-16 h-16 border border-slate-300 rounded-lg bg-white p-0.5 shrink-0">${rects}</svg>`;
        }

        function renderAdminLoginCards() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const container = document.getElementById('printCardsContainer');
            document.getElementById('cardTotalCount').innerText = students.length;

            if (students.length === 0) {
                container.innerHTML = `<div class="col-span-2 py-10 text-center text-xs text-slate-400">Belum ada data siswa. Tambahkan atau import Excel di atas.</div>`;
                return;
            }

            container.innerHTML = students.map(s => `
                <div class="login-card-item bg-white border-2 border-slate-300 rounded-2xl p-4 shadow-sm flex flex-col justify-between space-y-3">
                    <div class="flex items-center justify-between border-b border-slate-200 pb-2.5">
                        <div class="flex items-center gap-2.5 min-w-0">
                            <div class="w-8 h-8 rounded-lg bg-slate-900 text-white flex items-center justify-center font-extrabold text-xs shrink-0">CBT</div>
                            <div class="min-w-0">
                                <div class="text-[11px] font-extrabold text-slate-900 uppercase truncate">${settings.schoolName}</div>
                                <div class="text-[10px] text-slate-500 truncate">KARTU LOGIN PESERTA UJIAN</div>
                            </div>
                        </div>
                    </div>

                    <div class="flex items-center justify-between gap-3">
                        <div class="space-y-1 text-xs flex-1 min-w-0">
                            <div>
                                <span class="text-[10px] text-slate-400 block uppercase">Nama Peserta</span>
                                <span class="font-extrabold text-slate-900 text-sm truncate block">${s.name}</span>
                            </div>
                            <div class="grid grid-cols-2 gap-1 pt-0.5">
                                <div>
                                    <span class="text-[10px] text-slate-400 block">NISN</span>
                                    <span class="font-mono font-bold text-slate-800">${s.nisn}</span>
                                </div>
                                <div>
                                    <span class="text-[10px] text-slate-400 block">Kelas / Ruang</span>
                                    <span class="font-bold text-slate-800">${s.className} • ${s.room}</span>
                                </div>
                            </div>
                        </div>
                        ${generateMiniQRSVG(s.nisn + s.username)}
                    </div>

                    <div class="bg-slate-100 rounded-xl p-2.5 grid grid-cols-2 gap-2 text-xs border border-slate-200">
                        <div>
                            <span class="text-[10px] text-slate-500 block">Username</span>
                            <span class="font-mono font-extrabold text-slate-900">${s.username}</span>
                        </div>
                        <div>
                            <span class="text-[10px] text-slate-500 block">Password</span>
                            <span class="font-mono font-extrabold text-blue-700">${s.password}</span>
                        </div>
                    </div>
                </div>
            `).join('');
        }

        // --- MENU 7 ADMIN: SETELAN KESELURUHAN ---
        function renderAdminSettingsForms() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const proctors = getDB(STORAGE_KEYS.PROCTORS, DEFAULT_PROCTORS);

            document.getElementById('setSchoolName').value = settings.schoolName;
            document.getElementById('setSchoolSubtitle').value = settings.schoolSubtitle;
            document.getElementById('setAdminUsername').value = settings.adminUsername || 'admin';
            document.getElementById('setAdminPassword').value = settings.adminPassword || 'admin123';
            document.getElementById('setUnlockPin').value = settings.unlockPin;
            document.getElementById('setRequireFullscreen').checked = settings.requireFullscreen;
            document.getElementById('setDetectTabSwitch').checked = settings.detectTabSwitch;
            document.getElementById('setEnableSiren').checked = settings.enableSiren !== false;
            document.getElementById('setShowScoreToStudent').checked = settings.showScoreToStudent;

            const pContainer = document.getElementById('proctorListContainer');
            pContainer.innerHTML = proctors.map(p => `
                <div class="flex items-center justify-between p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs">
                    <div>
                        <div class="font-bold text-slate-900">${p.name} <span class="text-amber-700">(${p.room})</span></div>
                        <div class="text-[11px] font-mono text-slate-500">User: ${p.username} • Pass: ${p.password}</div>
                    </div>
                    <button onclick="deleteProctorAccount('${p.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600">
                        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                    </button>
                </div>
            `).join('');
        }

        function saveSchoolIdentity() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const newAdminUser = document.getElementById('setAdminUsername').value.trim();
            const newAdminPass = document.getElementById('setAdminPassword').value.trim();
            if (!newAdminUser || !newAdminPass) {
                showToast('Username dan Password Administrator tidak boleh kosong!', 'warning');
                return;
            }
            settings.schoolName = document.getElementById('setSchoolName').value.trim() || DEFAULT_SETTINGS.schoolName;
            settings.schoolSubtitle = document.getElementById('setSchoolSubtitle').value.trim() || DEFAULT_SETTINGS.schoolSubtitle;
            settings.adminUsername = newAdminUser;
            settings.adminPassword = newAdminPass;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast('Identitas sekolah & kredensial Administrator diperbarui!', 'success');
        }

        function saveSecuritySettings() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.unlockPin = document.getElementById('setUnlockPin').value.trim() || '123';
            settings.requireFullscreen = document.getElementById('setRequireFullscreen').checked;
            settings.detectTabSwitch = document.getElementById('setDetectTabSwitch').checked;
            settings.enableSiren = document.getElementById('setEnableSiren').checked;
            settings.showScoreToStudent = document.getElementById('setShowScoreToStudent').checked;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            syncAllViews();
            showToast('Setelan keamanan, sirine & PIN pembuka kunci berhasil disimpan!', 'success');
        }

        function handleAddProctor(e) {
            e.preventDefault();
            const proctors = getDB(STORAGE_KEYS.PROCTORS, []);
            proctors.push({
                id: 'proc_' + Date.now(),
                name: document.getElementById('procName').value.trim(),
                room: document.getElementById('procRoom').value.trim(),
                username: document.getElementById('procUser').value.trim(),
                password: document.getElementById('procPass').value.trim()
            });
            setDB(STORAGE_KEYS.PROCTORS, proctors);
            e.target.reset();
            syncAllViews();
            showToast('Akun pengawas baru berhasil ditambahkan!', 'success');
        }

        function deleteProctorAccount(id) {
            const proctors = getDB(STORAGE_KEYS.PROCTORS, []).filter(p => p.id !== id);
            setDB(STORAGE_KEYS.PROCTORS, proctors);
            syncAllViews();
            showToast('Akun pengawas dihapus.', 'info');
        }

        function handleAddGalleryItem(e) {
            e.preventDefault();
            const gallery = getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
            gallery.unshift({
                id: 'gal_' + Date.now(),
                title: document.getElementById('galTitle').value.trim(),
                url: document.getElementById('galUrl').value.trim(),
                svgType: 'lab'
            });
            setDB(STORAGE_KEYS.GALLERY, gallery);
            e.target.reset();
            syncAllViews();
            showToast('Foto kegiatan ditambahkan ke Galeri Portal!', 'success');
        }

        function exportDatabaseJSON() {
            const payload = {
                settings: getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS),
                packages: getDB(STORAGE_KEYS.PACKAGES, []),
                students: getDB(STORAGE_KEYS.STUDENTS, []),
                proctors: getDB(STORAGE_KEYS.PROCTORS, []),
                liveStatus: getDB(STORAGE_KEYS.LIVE_STATUS, {}),
                logs: getDB(STORAGE_KEYS.LOGS, []),
                gallery: getDB(STORAGE_KEYS.GALLERY, [])
            };
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Backup_CBT_PRO_${new Date().toISOString().slice(0, 10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('File backup database JSON berhasil diunduh!', 'success');
        }

        function importDatabaseJSON(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const parsed = JSON.parse(e.target.result);
                    if (parsed.settings) setDB(STORAGE_KEYS.SETTINGS, parsed.settings);
                    if (parsed.packages) setDB(STORAGE_KEYS.PACKAGES, parsed.packages);
                    if (parsed.students) setDB(STORAGE_KEYS.STUDENTS, parsed.students);
                    if (parsed.proctors) setDB(STORAGE_KEYS.PROCTORS, parsed.proctors);
                    if (parsed.liveStatus) setDB(STORAGE_KEYS.LIVE_STATUS, parsed.liveStatus);
                    if (parsed.logs) setDB(STORAGE_KEYS.LOGS, parsed.logs);
                    if (parsed.gallery) setDB(STORAGE_KEYS.GALLERY, parsed.gallery);
                    syncAllViews();
                    showToast('Database berhasil dipulihkan dari file backup!', 'success');
                } catch (err) {
                    showToast('Format file backup JSON tidak valid!', 'error');
                } finally {
                    event.target.value = '';
                }
            };
            reader.readAsText(file);
        }

        function confirmCleanStartEmpty() {
            showConfirmDialog('Kosongkan Data Awal untuk Produksi Sekolah?', 'Paket soal awal, daftar siswa awal, dan log akan dikosongkan sehingga Anda dapat mengunggah data sekolah yang asli.', () => {
                setDB(STORAGE_KEYS.PACKAGES, []);
                setDB(STORAGE_KEYS.STUDENTS, []);
                setDB(STORAGE_KEYS.LIVE_STATUS, {});
                setDB(STORAGE_KEYS.LOGS, []);
                syncAllViews();
                showToast('Data awal dikosongkan. Sistem siap diisi data sekolah Anda!', 'success');
            });
        }

        function confirmFactoryReset() {
            showConfirmDialog('Reset Seluruh Data ke Pengaturan Awal?', 'Tindakan ini akan mengatur ulang seluruh database lokal ke kondisi bawaan sistem.', () => {
                Object.values(STORAGE_KEYS).forEach(k => localStorage.removeItem(k));
                initializeAppDatabase();
                syncAllViews();
                showToast('Sistem berhasil direset ke pengaturan bawaan.', 'info');
            });
        }

        // =================================================================
        // K. DASHBOARD PENGAWAS RUANG (3 MENU TERPISAH & SINKRON)
        // =================================================================
        function switchProctorMenu(menuKey) {
            currentProctorMenu = menuKey;
            persistSessionState();
            ['kehadiran', 'pelanggaran', 'soal'].forEach(m => {
                document.getElementById(`proctorSection-${m}`).classList.add('hidden');
                const btn = document.getElementById(`procMenuBtn-${m}`);
                btn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2';
            });
            document.getElementById(`proctorSection-${menuKey}`).classList.remove('hidden');
            const activeBtn = document.getElementById(`procMenuBtn-${menuKey}`);
            activeBtn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-amber-500 text-white transition flex items-center gap-2';
            syncAllViews();
        }

        function getProctorRoomStudents() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const assignedRoom = activeSession?.user?.room || 'Semua Ruangan';
            return assignedRoom.toLowerCase().includes('semua')
                ? students
                : students.filter(s => s.room.toLowerCase() === assignedRoom.toLowerCase());
        }

        function renderProctorMonitoring() {
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const roomStudents = getProctorRoomStudents();

            const container = document.getElementById('proctorStudentCards');
            if (roomStudents.length === 0) {
                container.innerHTML = `<div class="py-10 text-center text-xs text-slate-400">Belum ada peserta terdaftar untuk ruangan ini.</div>`;
            } else {
                container.innerHTML = roomStudents.map(s => {
                    const st = liveMap[s.id] || { status: 'Belum Masuk', progress: '-', currentSubject: '-', violations: 0 };
                    const isLocked = st.status === 'TERKUNCI';
                    let statusBadge = `<span class="px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-600 text-[11px] font-bold">Belum Masuk</span>`;
                    if (st.status === 'Di Dashboard') statusBadge = `<span class="px-2.5 py-0.5 rounded-full bg-indigo-100 text-indigo-700 text-[11px] font-bold">Di Dashboard Siswa</span>`;
                    else if (st.status === 'Mengerjakan') statusBadge = `<span class="px-2.5 py-0.5 rounded-full bg-blue-100 text-blue-700 text-[11px] font-bold">Sedang Ujian</span>`;
                    else if (isLocked) statusBadge = `<span class="px-2.5 py-0.5 rounded-full bg-red-600 text-white text-[11px] font-bold animate-bounce">TERKUNCI + SIRINE</span>`;
                    else if (st.status === 'Selesai') statusBadge = `<span class="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">Selesai</span>`;
                    else if (st.status === 'Keluar / Offline') statusBadge = `<span class="px-2.5 py-0.5 rounded-full bg-slate-200 text-slate-700 text-[11px] font-bold">Keluar / Offline</span>`;

                    return `
                        <div class="p-4 rounded-2xl border ${isLocked ? 'border-red-400 bg-red-50' : 'border-slate-200 bg-slate-50/60'} flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div>
                                <div class="flex flex-wrap items-center gap-2">
                                    <span class="font-bold text-slate-900 text-sm">${s.name}</span>
                                    ${statusBadge}
                                </div>
                                <div class="text-xs text-slate-500 mt-1">
                                    NISN: <span class="font-mono font-semibold">${s.nisn}</span> • ${s.className} • <strong>${s.room}</strong> • Mapel: <strong>${st.currentSubject || '-'}</strong>
                                </div>
                                <div class="text-[11px] text-slate-500 mt-0.5">
                                    Progres: <strong>${st.progress || '-'}</strong> • Pelanggaran: <strong class="${st.violations > 0 ? 'text-red-600' : ''}">${st.violations || 0}x</strong>
                                </div>
                            </div>
                            ${isLocked ? `
                                <button onclick="remoteUnlockStudentById('${s.id}')" class="px-4 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-extrabold transition shadow shrink-0">
                                    Matikan Sirine &amp; Buka Kunci
                                </button>
                            ` : ''}
                        </div>
                    `;
                }).join('');
            }

            // Render log keluar masuk siswa terbaru di sidebar pengawas
            const logs = getDB(STORAGE_KEYS.LOGS, []).filter(l => l.role === 'siswa').slice(0, 15);
            const inOutBox = document.getElementById('proctorInOutLogList');
            if (logs.length === 0) {
                inOutBox.innerHTML = `<div class="text-slate-400 text-center py-6">Belum ada aktivitas masuk/keluar siswa.</div>`;
            } else {
                inOutBox.innerHTML = logs.map(l => `
                    <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-0.5">
                        <div class="flex justify-between text-[10px] text-slate-400 font-mono">
                            <span>${l.timestamp}</span>
                            <span class="font-bold ${l.category === 'VIOLATION' ? 'text-red-600' : 'text-blue-600'}">${l.category}</span>
                        </div>
                        <div class="font-bold text-slate-800">${l.userName}</div>
                        <div class="text-[11px] text-slate-600">${l.detail}</div>
                    </div>
                `).join('');
            }
        }

        function renderProctorViolationsView() {
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const roomStudents = getProctorRoomStudents();
            const lockedStudents = roomStudents.filter(s => liveMap[s.id]?.status === 'TERKUNCI');

            const violationLogs = getDB(STORAGE_KEYS.LOGS, []).filter(l => l.category === 'VIOLATION' || l.category === 'UNLOCK');
            document.getElementById('proctorViolationCountPill').innerText = violationLogs.filter(v => v.category === 'VIOLATION').length;

            const activeBox = document.getElementById('proctorActiveLockedBox');
            if (lockedStudents.length === 0) {
                activeBox.innerHTML = `
                    <div class="bg-emerald-50 border border-emerald-200 rounded-2xl p-4 text-xs text-emerald-800 font-semibold flex items-center gap-2">
                        <i data-lucide="check-circle2" class="w-4 h-4 text-emerald-600"></i> Tidak ada layar peserta yang terkunci (sirine tidak aktif) saat ini.
                    </div>
                `;
            } else {
                activeBox.innerHTML = lockedStudents.map(s => {
                    const st = liveMap[s.id];
                    return `
                        <div class="bg-red-50 border-2 border-red-400 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div>
                                <span class="px-2 py-0.5 rounded bg-red-600 text-white text-[10px] font-extrabold uppercase">SIRINE ALARM BERBUNYI</span>
                                <h4 class="font-extrabold text-red-950 text-sm mt-1">${s.name} (NISN: ${s.nisn} • ${s.className} • ${s.room})</h4>
                                <p class="text-xs text-red-800 mt-0.5">Alasan: ${st.lastViolationReason || 'Keluar dari sistem ujian'} • Total: <strong>${st.violations || 1}x</strong></p>
                            </div>
                            <button onclick="remoteUnlockStudentById('${s.id}')" class="px-4 py-2.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-extrabold transition shadow shrink-0">
                                Matikan Sirine &amp; Buka Kunci
                            </button>
                        </div>
                    `;
                }).join('');
            }

            const tbody = document.getElementById('proctorViolationLogsTbody');
            if (violationLogs.length === 0) {
                tbody.innerHTML = `<tr><td colspan="3" class="py-8 text-center text-slate-400 text-xs">Belum ada catatan pelanggaran peserta.</td></tr>`;
            } else {
                tbody.innerHTML = violationLogs.map(v => `
                    <tr class="hover:bg-slate-50">
                        <td class="py-2.5 px-4 font-mono text-xs text-slate-600">${v.timestamp}</td>
                        <td class="py-2.5 px-4">
                            <div class="font-bold text-slate-900">${v.userName}</div>
                            <div class="text-[11px] text-slate-500">NISN: ${v.userIdentifier} • ${v.roomOrClass}</div>
                        </td>
                        <td class="py-2.5 px-4 text-xs font-semibold ${v.category === 'VIOLATION' ? 'text-red-700' : 'text-emerald-700'}">${v.detail}</td>
                    </tr>
                `).join('');
            }
        }

        function selectProctorQuestionPackage(pkgId) {
            selectedProctorPkgId = pkgId;
            renderProctorSoalView();
        }

        function renderProctorSoalView() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const sel = document.getElementById('proctorPackageSelector');
            if (!sel) return;

            if (packages.length === 0) {
                sel.innerHTML = `<option value="">-- Belum Ada Paket Soal (Buat di Admin) --</option>`;
                document.getElementById('proctorPkgSoalCount').innerText = '0';
                document.getElementById('proctorSoalListContainer').innerHTML = `<div class="text-center py-8 text-xs text-slate-400">Belum ada paket mata pelajaran.</div>`;
                return;
            }

            if (!packages.some(p => p.id === selectedProctorPkgId)) {
                selectedProctorPkgId = packages[0].id;
            }

            sel.innerHTML = packages.map(p => `
                <option value="${p.id}">${p.subjectName} (${(p.questions || []).length} Soal)</option>
            `).join('');
            sel.value = selectedProctorPkgId;

            const pkg = packages.find(p => p.id === selectedProctorPkgId);
            const questions = pkg ? (pkg.questions || []) : [];
            document.getElementById('proctorPkgSoalCount').innerText = questions.length;

            const listBox = document.getElementById('proctorSoalListContainer');
            if (questions.length === 0) {
                listBox.innerHTML = `<div class="text-center py-8 text-xs text-slate-400">Belum ada soal pada paket ini. Silakan input manual atau upload Word/PDF di samping.</div>`;
            } else {
                listBox.innerHTML = questions.map((q, idx) => `
                    <div class="p-3.5 rounded-2xl border border-slate-200 bg-slate-50/70 space-y-2">
                        <div class="font-bold text-xs sm:text-sm text-slate-900">
                            <span class="text-amber-700 mr-1">${idx + 1}.</span> ${q.question}
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-1 text-xs">
                            ${Object.keys(q.options).map(k => `
                                <div class="px-2.5 py-1 rounded-lg border ${q.answer === k ? 'bg-emerald-100 border-emerald-300 text-emerald-900 font-bold' : 'bg-white border-slate-200 text-slate-600'}">
                                    <strong>${k}.</strong> ${q.options[k]}
                                </div>
                            `).join('')}
                        </div>
                    </div>
                `).join('');
            }
        }

        // =================================================================
        // L. MESIN UJIAN SISWA & LOCKDOWN ANTI-CURANG + SIRINE ALARM
        // =================================================================
        function shuffleArray(arr) {
            const copy = [...arr];
            for (let i = copy.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [copy[i], copy[j]] = [copy[j], copy[i]];
            }
            return copy;
        }

        function beginStudentExam(student, pkg) {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            let preparedQuestions = JSON.parse(JSON.stringify(pkg.questions || []));

            if (pkg.shuffleQuestions) {
                preparedQuestions = shuffleArray(preparedQuestions);
            }

            if (pkg.shuffleOptions) {
                preparedQuestions = preparedQuestions.map(q => {
                    const keys = Object.keys(q.options);
                    const entries = keys.map(k => ({ key: k, text: q.options[k], isCorrect: k === q.answer }));
                    const shuffled = shuffleArray(entries);
                    const newOptions = {};
                    let newAnswer = 'A';
                    shuffled.forEach((item, i) => {
                        const letter = keys[i];
                        newOptions[letter] = item.text;
                        if (item.isCorrect) newAnswer = letter;
                    });
                    return { ...q, options: newOptions, answer: newAnswer };
                });
            }

            examSession = {
                isRunning: true,
                isLocked: false,
                packageId: pkg.id,
                packageSubject: pkg.subjectName,
                packageTitle: pkg.examTitle,
                questions: preparedQuestions,
                answers: {},
                doubts: {},
                currentIdx: 0,
                remainingSeconds: (pkg.durationMinutes || 45) * 60,
                timerId: null,
                violations: 0
            };

            document.getElementById('studentExamTitle').innerText = `${pkg.subjectName} — ${pkg.examTitle}`;
            document.getElementById('studentExamName').innerText = student.name;
            document.getElementById('studentExamNisn').innerText = student.nisn;
            document.getElementById('studentExamRoom').innerText = `${student.className} • ${student.room}`;
            document.getElementById('examTotalNum').innerText = preparedQuestions.length;

            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const prev = liveMap[student.id] || {};
            liveMap[student.id] = {
                ...prev,
                status: 'Mengerjakan',
                currentSubject: pkg.subjectName,
                progress: `0 / ${preparedQuestions.length}`,
                violations: 0,
                remoteUnlocked: false,
                updatedAt: new Date().toLocaleTimeString('id-ID')
            };
            setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

            recordSystemLog('siswa', student.name, student.nisn, `${student.className} • ${student.room}`, 'EXAM_START', `Siswa memulai ujian mata pelajaran: ${pkg.subjectName}.`);

            persistActiveExamState();
            navigateTo('siswa');
            renderCurrentExamQuestion();
            startExamCountdownTimer();

            if (settings.requireFullscreen && document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
        }

        function renderCurrentExamQuestion() {
            const q = examSession.questions[examSession.currentIdx];
            if (!q) return;

            document.getElementById('examCurrentNum').innerText = examSession.currentIdx + 1;
            document.getElementById('examQuestionBody').innerText = q.question;

            const isDoubt = !!examSession.doubts[q.id];
            const doubtBtn = document.getElementById('doubtToggleBtn');
            doubtBtn.className = isDoubt
                ? 'px-3.5 py-1.5 rounded-xl bg-amber-400 text-slate-950 text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm'
                : 'px-3.5 py-1.5 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-bold transition flex items-center gap-1.5';

            const selectedOpt = examSession.answers[q.id];
            const optionKeys = Object.keys(q.options);
            const optionsContainer = document.getElementById('examOptionsList');

            optionsContainer.innerHTML = optionKeys.map(letter => {
                const isSelected = selectedOpt === letter;
                return `
                    <div onclick="selectExamOption('${q.id}', '${letter}')" class="p-4 rounded-2xl border-2 cursor-pointer transition flex items-center gap-3.5 ${isSelected ? 'border-blue-600 bg-blue-50/80 shadow-sm' : 'border-slate-200 hover:border-blue-300 bg-white'}">
                        <div class="w-9 h-9 rounded-xl flex items-center justify-center font-extrabold text-sm shrink-0 ${isSelected ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-700'}">
                            ${letter}
                        </div>
                        <div class="text-sm font-medium text-slate-800 flex-1">${q.options[letter]}</div>
                    </div>
                `;
            }).join('');

            document.getElementById('btnPrevQuestion').disabled = examSession.currentIdx === 0;
            const nextBtn = document.getElementById('btnNextQuestion');
            if (examSession.currentIdx === examSession.questions.length - 1) {
                nextBtn.innerHTML = `Selesai &amp; Kumpulkan <i data-lucide="check-circle" class="w-4 h-4"></i>`;
                nextBtn.className = 'px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2';
                nextBtn.onclick = confirmSubmitExam;
            } else {
                nextBtn.innerHTML = `Selanjutnya <i data-lucide="chevron-right" class="w-4 h-4"></i>`;
                nextBtn.className = 'px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2';
                nextBtn.onclick = goNextQuestion;
            }

            renderExamNumberGrid();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function selectExamOption(qId, letter) {
            if (examSession.isLocked) return;
            examSession.answers[qId] = letter;

            if (activeSession?.user) {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                const answeredCount = Object.keys(examSession.answers).length;
                if (liveMap[activeSession.user.id]) {
                    liveMap[activeSession.user.id].progress = `${answeredCount} / ${examSession.questions.length}`;
                    liveMap[activeSession.user.id].updatedAt = new Date().toLocaleTimeString('id-ID');
                    setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                }
            }

            persistActiveExamState();
            renderCurrentExamQuestion();
        }

        function toggleDoubtQuestion() {
            if (examSession.isLocked) return;
            const q = examSession.questions[examSession.currentIdx];
            examSession.doubts[q.id] = !examSession.doubts[q.id];
            renderCurrentExamQuestion();
        }

        function goPrevQuestion() {
            if (examSession.currentIdx > 0) {
                examSession.currentIdx--;
                renderCurrentExamQuestion();
            }
        }

        function goNextQuestion() {
            if (examSession.currentIdx < examSession.questions.length - 1) {
                examSession.currentIdx++;
                renderCurrentExamQuestion();
            }
        }

        function jumpToQuestion(idx) {
            examSession.currentIdx = idx;
            renderCurrentExamQuestion();
        }

        function renderExamNumberGrid() {
            const grid = document.getElementById('examNumberGrid');
            grid.innerHTML = examSession.questions.map((q, idx) => {
                const isCurrent = idx === examSession.currentIdx;
                const isDoubt = !!examSession.doubts[q.id];
                const isAnswered = !!examSession.answers[q.id];

                let cls = 'bg-slate-100 text-slate-700 border-slate-300';
                if (isDoubt) cls = 'bg-amber-400 text-slate-950 border-amber-500';
                else if (isAnswered) cls = 'bg-blue-600 text-white border-blue-600';

                const ring = isCurrent ? 'ring-2 ring-slate-900 ring-offset-2' : '';
                return `
                    <button onclick="jumpToQuestion(${idx})" class="h-10 rounded-xl border font-bold text-xs transition ${cls} ${ring}">
                        ${idx + 1}
                    </button>
                `;
            }).join('');
        }

        function changeExamFontSize(size) {
            const body = document.getElementById('examQuestionBody');
            body.classList.remove('text-sm', 'text-base', 'text-lg', 'text-xl');
            if (size === 'sm') body.classList.add('text-sm');
            else if (size === 'lg') body.classList.add('text-xl');
            else body.classList.add('text-base');
        }

        function startExamCountdownTimer() {
            if (examSession.timerId) clearInterval(examSession.timerId);
            updateTimerDisplay();
            examSession.timerId = setInterval(() => {
                examSession.remainingSeconds--;
                if (examSession.remainingSeconds % 5 === 0) persistActiveExamState();
                updateTimerDisplay();
                if (examSession.remainingSeconds <= 0) {
                    clearInterval(examSession.timerId);
                    finalizeStudentExam(true);
                }
            }, 1000);
        }

        function updateTimerDisplay() {
            const mins = Math.floor(Math.max(0, examSession.remainingSeconds) / 60);
            const secs = Math.max(0, examSession.remainingSeconds) % 60;
            const display = document.getElementById('studentTimerDisplay');
            display.innerText = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
            if (examSession.remainingSeconds <= 300) {
                display.classList.remove('text-amber-400');
                display.classList.add('text-red-400', 'animate-pulse');
            }
        }

        // --- PROTEKSI KETAT ANTI-CURANG + SIRINE ALARM OTOMATIS ---
        function triggerAntiCheatLock(reasonText) {
            if (!examSession.isRunning || examSession.isLocked) return;
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (!settings.detectTabSwitch) return;

            examSession.isLocked = true;
            examSession.violations++;

            // Nyalakan Sirine Alarm Keras!
            startSirenAlarm();

            const student = activeSession?.user;
            if (student) {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                const prev = liveMap[student.id] || {};
                liveMap[student.id] = {
                    ...prev,
                    status: 'TERKUNCI',
                    violations: examSession.violations,
                    lastViolationReason: reasonText,
                    remoteUnlocked: false,
                    updatedAt: new Date().toLocaleTimeString('id-ID')
                };
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

                recordSystemLog(
                    'siswa',
                    student.name,
                    student.nisn,
                    `${student.className} • ${student.room}`,
                    'VIOLATION',
                    `PELANGGARAN (${examSession.violations}x) pada ${examSession.packageSubject}: ${reasonText} (Sirine Aktif)`
                );
            }

            document.getElementById('lockReasonText').innerText = reasonText;
            document.getElementById('lockStudentName').innerText = student ? student.name : '-';
            document.getElementById('lockStudentMeta').innerText = student ? `${student.nisn} • ${student.className} (${student.room})` : '-';
            document.getElementById('lockExamSubject').innerText = examSession.packageSubject || '-';
            document.getElementById('lockViolationBadge').innerText = `${examSession.violations} Kali`;
            document.getElementById('unlockPinField').value = '';

            const overlay = document.getElementById('antiCheatLockOverlay');
            overlay.classList.remove('hidden');
            overlay.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function unlockExamScreen() {
            examSession.isLocked = false;
            stopSirenAlarm();

            const overlay = document.getElementById('antiCheatLockOverlay');
            overlay.classList.add('hidden');
            overlay.classList.remove('flex');

            const student = activeSession?.user;
            if (student) {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                if (liveMap[student.id]) {
                    liveMap[student.id].status = 'Mengerjakan';
                    liveMap[student.id].remoteUnlocked = false;
                    setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                }
            }

            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (settings.requireFullscreen && document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
            showToast('Sirine dimatikan & kunci layar dibuka. Silakan lanjutkan ujian dengan jujur.', 'info');
        }

        function verifyPinUnlock() {
            const enteredPin = document.getElementById('unlockPinField').value.trim();
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (enteredPin === settings.unlockPin) {
                const student = activeSession?.user;
                if (student) {
                    recordSystemLog('siswa', student.name, student.nisn, `${student.className} • ${student.room}`, 'UNLOCK', 'Kunci layar & sirine dibuka menggunakan PIN Pengawas.');
                }
                unlockExamScreen();
            } else {
                showToast('Kode PIN Pengawas salah! Minta Pengawas Ruangan memasukkan PIN yang benar.', 'error');
                document.getElementById('unlockPinField').value = '';
            }
        }

        // Listener deteksi keluar tab, berpindah aplikasi (blur), keluar fullscreen, klik kanan, dan shortcut terlarang
        document.addEventListener('visibilitychange', () => {
            if (document.hidden && examSession.isRunning) {
                triggerAntiCheatLock('Terdeteksi berpindah tab browser atau meminimalkan aplikasi ujian!');
            }
        });

        window.addEventListener('blur', () => {
            if (examSession.isRunning) {
                triggerAntiCheatLock('Layar ujian kehilangan fokus! Dilarang membuka aplikasi lain, split-screen, atau jendela mengambang.');
            }
        });

        document.addEventListener('fullscreenchange', () => {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (examSession.isRunning && settings.requireFullscreen && !document.fullscreenElement) {
                triggerAntiCheatLock('Peserta keluar dari mode layar penuh (Fullscreen)!');
            }
        });

        document.addEventListener('contextmenu', (e) => {
            if (examSession.isRunning) {
                e.preventDefault();
                showToast('Klik kanan dinonaktifkan selama ujian berlangsung.', 'warning');
            }
        });

        document.addEventListener('keydown', (e) => {
            if (!examSession.isRunning) return;
            const forbiddenCtrl = (e.ctrlKey || e.metaKey) && ['c', 'v', 'x', 'u', 's', 'p', 't', 'n', 'w', 'tab'].includes(e.key.toLowerCase());
            if (forbiddenCtrl || e.key === 'F12' || e.key === 'PrintScreen' || (e.altKey && e.key === 'Tab')) {
                e.preventDefault();
                triggerAntiCheatLock(`Percobaan menggunakan pintasan keyboard terlarang (${e.key}) saat ujian!`);
            }
        });

        function confirmSubmitExam() {
            const answered = Object.keys(examSession.answers).length;
            const total = examSession.questions.length;
            showConfirmDialog(
                'Kumpulkan Jawaban Ujian?',
                `Anda telah menjawab ${answered} dari ${total} soal pada mata pelajaran ${examSession.packageSubject}. Yakin ingin mengakhiri ujian dan mengirim jawaban?`,
                () => finalizeStudentExam(false)
            );
        }

        function finalizeStudentExam(isAutoTimeout) {
            if (examSession.timerId) clearInterval(examSession.timerId);
            examSession.isRunning = false;
            examSession.isLocked = false;
            localStorage.removeItem(STORAGE_KEYS.ACTIVE_EXAM);
            stopSirenAlarm();

            if (document.fullscreenElement && document.exitFullscreen) {
                document.exitFullscreen().catch(() => {});
            }

            let correct = 0;
            let wrong = 0;
            let earnedPoints = 0;
            let maxPoints = 0;

            examSession.questions.forEach(q => {
                const pts = q.points || 10;
                maxPoints += pts;
                const ans = examSession.answers[q.id];
                if (ans && ans === q.answer) {
                    correct++;
                    earnedPoints += pts;
                } else {
                    wrong++;
                }
            });

            const finalScore = maxPoints > 0 ? Math.round((earnedPoints / maxPoints) * 100) : 0;
            const student = activeSession?.user;
            const finishedTimeStr = new Date().toLocaleTimeString('id-ID');

            if (student) {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                const prev = liveMap[student.id] || {};
                const completedPackages = prev.completedPackages || {};
                completedPackages[examSession.packageId] = {
                    subjectName: examSession.packageSubject,
                    examTitle: examSession.packageTitle,
                    correct,
                    wrong,
                    score: finalScore,
                    finishedAt: finishedTimeStr
                };

                liveMap[student.id] = {
                    ...prev,
                    status: 'Selesai',
                    currentSubject: examSession.packageSubject,
                    progress: `${examSession.questions.length} / ${examSession.questions.length}`,
                    correct,
                    wrong,
                    score: finalScore,
                    violations: examSession.violations,
                    finishedAt: finishedTimeStr,
                    completedPackages
                };
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

                recordSystemLog(
                    'siswa',
                    student.name,
                    student.nisn,
                    `${student.className} • ${student.room}`,
                    'EXAM_FINISH',
                    `Siswa menyelesaikan ujian ${examSession.packageSubject} (Nilai: ${finalScore}, Pelanggaran: ${examSession.violations}x).`
                );
            }

            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const scoreBox = document.getElementById('studentScoreSummaryBox');
            if (settings.showScoreToStudent) {
                scoreBox.classList.remove('hidden');
                document.getElementById('studentFinalScoreText').innerText = finalScore;
                document.getElementById('studentFinalDetailText').innerText = `Benar ${correct} • Salah ${wrong} • Pelanggaran ${examSession.violations}x`;
            } else {
                scoreBox.classList.add('hidden');
            }

            const modal = document.getElementById('examFinishedModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            if (typeof lucide !== 'undefined') lucide.createIcons();

            if (isAutoTimeout) {
                showToast('Waktu ujian habis! Jawaban otomatis dikumpulkan.', 'warning');
            }
        }

        function closeExamFinishedModal() {
            const modal = document.getElementById('examFinishedModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
            navigateTo('siswa-dashboard');
        }

        // =================================================================
        // M. SINKRONISASI REAL-TIME ANTAR TAB / PERANGKAT
        // =================================================================
        window.addEventListener('storage', (e) => {
            if (Object.values(STORAGE_KEYS).includes(e.key)) {
                if (examSession.isRunning && examSession.isLocked && activeSession?.user) {
                    const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                    if (liveMap[activeSession.user.id]?.remoteUnlocked) {
                        unlockExamScreen();
                    }
                }
                syncAllViews();
            }
        });

        setInterval(() => {
            if (activeSession?.role === 'admin' || activeSession?.role === 'pengawas' || (activeSession?.role === 'siswa' && !examSession.isRunning)) {
                syncAllViews();
            }
            if (examSession.isRunning && examSession.isLocked && activeSession?.user) {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                if (liveMap[activeSession.user.id]?.remoteUnlocked) {
                    unlockExamScreen();
                }
            }
        }, 2500);

        // =================================================================
        // N. INISIALISASI SAAT HALAMAN DIMUAT
        // =================================================================
        function restorePersistedSessionOnLoad() {
            const savedSession = getDB(STORAGE_KEYS.ACTIVE_SESSION, null);
            const savedUi = getDB(STORAGE_KEYS.UI_STATE, {});
            if (!savedSession || !savedSession.role) {
                navigateTo('portal');
                const params = new URLSearchParams(window.location.search);
                const roleParam = params.get('role');
                if (roleParam && ['siswa', 'pengawas', 'admin'].includes(roleParam)) {
                    openLoginModal(roleParam);
                }
                return;
            }

            // Validasi ulang user dari database agar tetap sinkron
            if (savedSession.role === 'admin') {
                activeSession = savedSession;
                const menu = savedUi.adminMenu || 'monitoring';
                if (savedUi.currentView === 'portal') {
                    navigateTo('portal');
                } else {
                    navigateTo('admin');
                    switchAdminMenu(menu);
                }
                return;
            }

            if (savedSession.role === 'pengawas') {
                const proctors = getDB(STORAGE_KEYS.PROCTORS, []);
                const freshProctor = proctors.find(p => p.id === savedSession.user?.id || p.username === savedSession.user?.username);
                if (freshProctor) {
                    activeSession = { role: 'pengawas', user: freshProctor };
                    const menu = savedUi.proctorMenu || 'kehadiran';
                    if (savedUi.currentView === 'portal') {
                        navigateTo('portal');
                    } else {
                        navigateTo('pengawas');
                        switchProctorMenu(menu);
                    }
                    return;
                }
            }

            if (savedSession.role === 'siswa') {
                const students = getDB(STORAGE_KEYS.STUDENTS, []);
                const freshStudent = students.find(s => s.id === savedSession.user?.id || s.nisn === savedSession.user?.nisn);
                if (freshStudent) {
                    activeSession = { role: 'siswa', user: freshStudent };

                    // Cek apakah siswa sedang dalam sesi ujian aktif sebelum refresh
                    const savedExam = getDB(STORAGE_KEYS.ACTIVE_EXAM, null);
                    if (savedExam && savedExam.studentId === freshStudent.id && Array.isArray(savedExam.questions) && savedExam.questions.length > 0) {
                        const elapsedSec = Math.floor((Date.now() - (savedExam.savedAt || Date.now())) / 1000);
                        const remSec = Math.max(0, (savedExam.remainingSeconds || 0) - elapsedSec);
                        examSession = {
                            isRunning: true,
                            isLocked: !!savedExam.isLocked,
                            packageId: savedExam.packageId,
                            packageSubject: savedExam.packageSubject,
                            packageTitle: savedExam.packageTitle,
                            questions: savedExam.questions,
                            answers: savedExam.answers || {},
                            doubts: savedExam.doubts || {},
                            currentIdx: savedExam.currentIdx || 0,
                            remainingSeconds: remSec,
                            timerId: null,
                            violations: savedExam.violations || 0
                        };

                        document.getElementById('studentExamTitle').innerText = `${examSession.packageSubject} — ${examSession.packageTitle}`;
                        document.getElementById('studentExamName').innerText = freshStudent.name;
                        document.getElementById('studentExamNisn').innerText = freshStudent.nisn;
                        document.getElementById('studentExamRoom').innerText = `${freshStudent.className} • ${freshStudent.room}`;
                        document.getElementById('examTotalNum').innerText = examSession.questions.length;

                        navigateTo('siswa');
                        renderCurrentExamQuestion();
                        startExamCountdownTimer();
                        if (examSession.isLocked) {
                            document.getElementById('lockReasonText').innerText = 'Halaman ujian dimuat ulang saat layar masih dalam status terkunci.';
                            document.getElementById('lockStudentName').innerText = freshStudent.name;
                            document.getElementById('lockStudentMeta').innerText = `${freshStudent.nisn} • ${freshStudent.className} (${freshStudent.room})`;
                            document.getElementById('lockExamSubject').innerText = examSession.packageSubject || '-';
                            document.getElementById('lockViolationBadge').innerText = `${examSession.violations} Kali`;
                            const overlay = document.getElementById('antiCheatLockOverlay');
                            overlay.classList.remove('hidden');
                            overlay.classList.add('flex');
                            startSirenAlarm();
                        }
                        return;
                    }

                    const tab = savedUi.studentDashTab || 'ujian';
                    if (savedUi.currentView === 'portal') {
                        navigateTo('portal');
                    } else {
                        navigateTo('siswa-dashboard');
                        switchStudentDashTab(tab);
                    }
                    return;
                }
            }

            // Jika akun sudah dihapus dari database, bersihkan sesi
            localStorage.removeItem(STORAGE_KEYS.ACTIVE_SESSION);
            navigateTo('portal');
        }

        document.addEventListener('DOMContentLoaded', () => {
            initializeAppDatabase();
            restorePersistedSessionOnLoad();
        });
    </script>

</body></html>
