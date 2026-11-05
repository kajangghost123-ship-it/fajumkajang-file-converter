const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const libreoffice = require('libreoffice-convert');
const ffmpeg = require('fluent-ffmpeg');
const PDFDocument = require('pdfkit-table');


const app = express();
const PORT = process.env.PORT || 3000;
// Wadah database komentar pengunjung sementara untuk menaikkan SEO AdSense
let databaseKomentar = [
    { nama: "Randi Audio", isi: "Hasil mastering audionya jernih banget bang! Sukses terus Fajum Kajang!", waktu: "04/11/2026, 09.00" }
];
// Jalur API agar website bisa mengambil list komentar
app.get('/api/komentar', (req, res) => {
    res.json(databaseKomentar);
});

// Jalur API agar pengunjung bisa mengirim komentar baru
app.post('/api/komentar', express.json(), (req, res) => {
    const { nama, isi } = req.body;
    if (nama && isi) {
        databaseKomentar.push({
            nama: nama,
            isi: isi,
            waktu: new Date().toLocaleString('id-ID')
        });
        res.json({ success: true });
    } else {
        res.status(400).json({ error: 'Nama dan isi wajib diisi ya bang!' });
    }
});
// Jalur agar sistem website bisa membaca file komentar.html
app.get('/halaman-komentar', (req, res) => {
    res.sendFile(path.join(__dirname, 'komentar.html'));
});

// PENTING: Baris ini untuk mengizinkan website membaca file logo.png di folder proyek kamu
app.use(express.static(__dirname));

// Setup folder penyimpanan sementara
const upload = multer({ dest: 'uploads/' });

// 1. TAMPILAN FRONTEND WEBSITE (Tema Studio Musik / Musisi)
app.get('/', (req, res) => {
    res.send(`
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>FAJUM KAJANG - Studio Converter</title>
        <style>
            body { 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                background: linear-gradient(135deg, #0f0f14 0%, #1a1a24 100%); 
                color: #e0e0e0; 
                text-align: center; 
                padding: 40px 20px; 
            }
            .brand-header {
                margin-bottom: 30px;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
            }
                            .spark-container {
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                z-index: 1; 
                overflow: hidden;
                pointer-events: none;
            }
            .spark {
                position: absolute;
                bottom: -20px;
                width: 4px;
                height: 4px;
                background: radial-gradient(circle, #ffaa00 0%, #ff0055 100%);
                border-radius: 50%;
                box-shadow: 0 0 10px #ff3300, 0 0 20px #ff0055;
                animation: naikDanRedup linear infinite;
            }
            @keyframes naikDanRedup {
                0% { transform: translateY(0) scale(1); opacity: 1; }
                80% { opacity: 0.8; }
                100% { transform: translateY(-105vh) scale(0.2); opacity: 0; }
            }
            /* --- TOMBOL MENU GARIS TIGA ☰ DI POJOK KIRI ATAS --- */
            .menu-profile-btn {
                position: absolute;
                top: 20px;
                left: 20px;
                background: none;
                border: none;
                color: #ff0055;
                font-size: 32px;
                cursor: pointer;
                z-index: 100;
                transition: 0.3s;
                width: auto;
                margin: 0;
                padding: 5px;
            }
            .menu-profile-btn:hover {
                transform: scale(1.1);
                color: #00ffcc;
                text-shadow: 0 0 10px #00ffcc;
            }

            /* --- SIDEBAR MENU TRANSPARAN SLIDE DARI KIRI --- */
            .sidebar-menu {
                position: fixed;
                top: 0;
                left: -350px; /* Tersembunyi di luar layar saat awal */
                width: 300px;
                height: 100%;
                background: rgba(15, 15, 20, 0.65); /* Efek Transparan */
                backdrop-filter: blur(15px); /* Efek Blur Kaca */
                -webkit-backdrop-filter: blur(15px);
                border-right: 1px solid rgba(255, 0, 85, 0.3);
                box-shadow: 5px 0 25px rgba(0, 0, 0, 0.7);
                z-index: 1000;
                transition: 0.4s cubic-bezier(0.1, 0.7, 0.1, 1);
                padding: 30px;
                box-sizing: border-box;
                text-align: center;
                display: flex;
                flex-direction: column;
                align-items: center;
            }

            /* Kelas pembantu penting agar menu bergeser masuk layar saat diklik */
            .sidebar-menu.active {
                left: 0 !important;
            }

            /* Tombol silang untuk menutup menu samping */
            .close-sidebar-btn {
                align-self: flex-end;
                color: #aaa;
                font-size: 30px;
                cursor: pointer;
                transition: 0.3s;
                margin-bottom: 20px;
            }
            .close-sidebar-btn:hover { 
                color: #ff0055; 
                transform: rotate(90deg);
            }
            
            /* --- FOTO PROFIL DI DALAM SIDEBAR --- */
            .profile-img {
                width: 130px;
                height: 130px;
                border-radius: 50%;
                border: 3px solid #ff0055;
                object-fit: cover;
                margin-bottom: 15px;
                box-shadow: 0 0 15px rgba(255, 0, 85, 0.4);
                transition: 0.3s;
            }
            .profile-img:hover {
                border-color: #00ffcc;
                box-shadow: 0 0 20px rgba(0, 255, 204, 0.6);
                transform: scale(1.03);
            }

                        /* --- TOMBOL MENU GARIS TIGA ☰ DI POJOK KIRI ATAS --- */
            .menu-profile-btn {
                position: absolute;
                top: 20px;
                left: 20px;
                background: none;
                border: none;
                color: #ff0055;
                font-size: 32px;
                cursor: pointer;
                z-index: 100;
                transition: 0.3s;
                width: auto;
                margin: 0;
                padding: 5px;
            }
            .menu-profile-btn:hover {
                transform: scale(1.1);
                color: #00ffcc;
                text-shadow: 0 0 10px #00ffcc;
            }

            /* --- SIDEBAR MENU TRANSPARAN SLIDE DARI KIRI --- */
            .sidebar-menu {
                position: fixed;
                top: 0;
                left: -350px; /* Menyembunyikan menu di awal */
                width: 300px;
                height: 100%;
                background: rgba(15, 15, 20, 0.65); /* Efek Transparan */
                backdrop-filter: blur(15px); /* Efek Kaca Kabur */
                -webkit-backdrop-filter: blur(15px);
                border-right: 1px solid rgba(255, 0, 85, 0.3);
                box-shadow: 5px 0 25px rgba(0, 0, 0, 0.7);
                z-index: 1000;
                transition: 0.4s cubic-bezier(0.1, 0.7, 0.1, 1);
                padding: 30px;
                box-sizing: border-box;
                text-align: center;
                display: flex;
                flex-direction: column;
                align-items: center;
            }

            /* Properti krusial pemicu geser masuk layar */
            .sidebar-menu.active {
                left: 0 !important;
            }

            /* Tombol silang tutup menu samping */
            .close-sidebar-btn {
                align-self: flex-end;
                color: #aaa;
                font-size: 30px;
                cursor: pointer;
                transition: 0.3s;
                margin-bottom: 20px;
            }
            .close-sidebar-btn:hover { 
                color: #ff0055; 
                transform: rotate(90deg);
            }
            
            /* --- FOTO PROFIL DI DALAM SIDEBAR --- */
            .profile-img {
                width: 130px;
                height: 130px;
                border-radius: 50%;
                border: 3px solid #ff0055;
                object-fit: cover;
                margin-bottom: 15px;
                box-shadow: 0 0 15px rgba(255, 0, 85, 0.4);
                transition: 0.3s;
            }
            .profile-img:hover {
                border-color: #00ffcc;
                box-shadow: 0 0 20px rgba(0, 255, 204, 0.6);
                transform: scale(1.03);
            }

            .sidebar-menu h3 { margin: 10px 0 5px 0; color: #fff; font-size: 1.5rem; }
            .sidebar-menu p { margin: 5px 0; color: #00ffcc; font-size: 0.9rem; font-style: italic; }
            .sidebar-menu .bio { color: #bbb; font-size: 0.95rem; margin-top: 20px; line-height: 1.6; }

            /* --- DESAIN TOMBOL MEDIA SOSIAL --- */
            .social-container {
                display: flex;
                justify-content: center;
                gap: 15px;
                margin-top: auto; 
                margin-bottom: 20px;
                border-top: 1px solid rgba(255, 255, 255, 0.1);
                padding-top: 20px;
                width: 100%;
            }
            .social-link {
                display: flex;
                align-items: center;
                justify-content: center;
                width: 45px;
                height: 45px;
                border-radius: 50%;
                color: white;
                text-decoration: none;
                font-size: 14px;
                font-weight: bold;
                transition: 0.3s;
                background-color: rgba(255, 255, 255, 0.05);
                border: 1px solid rgba(255, 255, 255, 0.1);
            }
            .social-link:hover { transform: translateY(-4px); }
            .social-link.youtube:hover { background-color: #ff0000; box-shadow: 0 0 15px #ff0000; }
            .social-link.facebook:hover { background-color: #1877f2; box-shadow: 0 0 15px #1877f2; }
            .social-link.instagram:hover { 
                background: linear-gradient(45deg, #f09433 0%, #e6683c 25%, #dc2743 50%, #cc2366 75%, #bc1888 100%); 
                box-shadow: 0 0 15px #dc2743;
            }
            .brand-title-container {
                display: flex;
                align-items: center;
                justify-content: center;
            }
            .brand-header h1 { 
                font-size: 3rem; 
                color: #ff0055; 
                text-shadow: 0 0 15px rgba(255, 0, 85, 0.6);
                margin: 0;
                letter-spacing: 2px;
            }
            .brand-logo {
                height: 70px; /* Atur tinggi logo di sini */
                width: auto;
                margin-right: 15px;
                filter: drop-shadow(0 0 10px rgba(255, 0, 85, 0.6));
            }
            .brand-header p { color: #00ffcc; font-style: italic; margin-top: 5px; }
            
            .studio-box { 
                background-color: rgba(30, 30, 40, 0.85); 
                padding: 40px; 
                border-radius: 15px; 
                display: inline-block; 
                width: 100%;
                max-width: 500px;
                border: 1px solid #333344;
                box-shadow: 0 10px 30px rgba(0,0,0,0.7);
                text-align: left; 
            }
            
            .section-title {
                font-size: 1.2rem;
                color: #ffaa00;
                margin-bottom: 20px;
                border-bottom: 1px solid #444;
                padding-bottom: 5px;
                text-align: left;
            }

            input[type="file"], select, button { 
                width: 100%;
                box-sizing: border-box;
                margin: 10px 0 20px 0; 
                padding: 12px; 
                font-size: 16px; 
                border-radius: 6px;
                background-color: #252535;
                border: 1px solid #444455;
                color: #fff;
            }
            
            select:focus, input:focus {
                outline: none;
                border-color: #00ffcc;
            }

            button { 
                background: linear-gradient(90deg, #ff0055 0%, #990033 100%); 
                color: white; 
                border: none; 
                cursor: pointer; 
                font-weight: bold; 
                transition: 0.3s;
                letter-spacing: 1px;
                margin-top: 10px;
            }
            button:hover { 
                transform: translateY(-2px);
                box-shadow: 0 5px 15px rgba(255, 0, 85, 0.4);
            }
            label {
                display: block;
                text-align: left;
                font-size: 14px;
                color: #aaa;
                margin-bottom: 2px;
            }
                    /* --- TOMBOL TITIK TIGA DI POJOK KIRI ATAS --- */
            .menu-profile-btn {
                position: absolute;
                top: 20px;
                left: 20px;
                background: none;
                border: none;
                color: #ff0055;
                font-size: 30px;
                cursor: pointer;
                z-index: 100;
                transition: 0.3s;
                width: auto;
                margin: 0;
                padding: 5px;
            }
            .menu-profile-btn:hover {
                transform: scale(1.1);
                color: #00ffcc;
                text-shadow: 0 0 10px #00ffcc;
            }

            /* --- MODAL POP-UP PROFIL --- */
            .profile-modal {
                display: none; 
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(0, 0, 0, 0.8);
                z-index: 1000;
                align-items: center;
                justify-content: center;
            }
            .profile-card {
                background: rgba(25, 25, 35, 0.95);
                border: 2px solid #ff0055;
                box-shadow: 0 0 25px rgba(255, 0, 85, 0.5);
                padding: 30px;
                border-radius: 15px;
                width: 320px;
                text-align: center;
                position: relative;
                animation: popIn 0.3s ease;
            }
            @keyframes popIn {
                from { transform: scale(0.8); opacity: 0; }
                to { transform: scale(1); opacity: 1; }
            }
            .close-btn {
                position: absolute;
                top: 10px;
                right: 15px;
                color: #aaa;
                font-size: 24px;
                cursor: pointer;
            }
            .close-btn:hover { color: #ff0055; }
            
            .profile-img {
                width: 120px;
                height: 120px;
                border-radius: 50%;
                border: 3px solid #00ffcc;
                object-fit: cover;
                margin-bottom: 15px;
                box-shadow: 0 0 15px rgba(0, 255, 204, 0.4);
            }
            .profile-card h3 { margin: 10px 0 5px 0; color: #fff; font-size: 1.4rem; }
            .profile-card p { margin: 5px 0; color: #00ffcc; font-size: 0.9rem; font-style: italic; }
            .profile-card .bio { color: #bbb; font-size: 0.9rem; margin-top: 15px; margin-bottom: 20px; font-style: normal; }

            /* --- DESAIN TOMBOL MEDIA SOSIAL --- */
            .social-container {
                display: flex;
                justify-content: center;
                gap: 15px;
                margin-top: 20px;
                border-top: 1px solid #333;
                padding-top: 15px;
            }
            .social-link {
                display: flex;
                align-items: center;
                justify-content: center;
                width: 40px;
                height: 40px;
                border-radius: 50%;
                color: white;
                text-decoration: none;
                font-size: 18px;
                font-weight: bold;
                transition: 0.3s;
                background-color: #222;
            }
            .social-link:hover {
                transform: translateY(-3px);
            }
            .social-link.youtube:hover { background-color: #ff0000; box-shadow: 0 0 10px #ff0000; }
            .social-link.facebook:hover { background-color: #1877f2; box-shadow: 0 0 10px #1877f2; }
            .social-link.instagram:hover { 
                background: linear-gradient(45deg, #f09433 0%, #e6683c 25%, #dc2743 50%, #cc2366 75%, #bc1888 100%); 
                box-shadow: 0 0 10px #dc2743;
            }
        </style>
    </head>
    <body>
        <!-- Tombol Titik Tiga -->
        <button class="menu-profile-btn" onclick="bukaSidebar()">☰</button>
        <div class="sidebar-overlay" id="sidebarOverlay" onclick="tutupSidebar()"></div>
        <!-- Struktur Pop-Up Profil Pemilik -->
        <div class="profile-modal" id="modalProfil">
            <div class="profile-card">
                <span class="close-btn" onclick="tutupProfil()">&times;</span>
                <!-- FOTO PROFIL ABANG -->
                <img src="/fajumkajangg.jpeg" alt="Foto Pemilik" class="profile-img" onerror="this.src='https://placeholder.com'">
                <h3>Fajum Kajang</h3>
                <p>🎛️ Studio Founder & Producer</p>
                <div class="bio">
                    Selamat datang di stasiun konversi digital resmi Fajum Kajang. Tempat mixing format file audio, musik, dan dokumen dengan kualitas mastering terbaik.
                </div>
                                <!-- BAGIAN LINK MEDIA SOSIAL ASLI -->
                <!-- BAGIAN LINK MEDIA SOSIAL ASLI -->
                <div class="social-container">
                    <a href="https://youtube.com/@fajumkajang1?si=11PN8dLNmdfP2kd-" target="_blank" class="social-link youtube" title="YouTube">YT</a>
                    <a href="https://www.facebook.com/faj.um.7?" target="_blank" class="social-link facebook" title="Facebook">FB</a>
                    <a href="https://www.instagram.com/fajumkajang?vrfl=MTM1Z3pjZ2ExdXZp" target="_blank" class="social-link instagram" title="Instagram">IG</a>
                </div>
            <!-- --- KOTAK TESTIMONI MURNI ANTI-MERAH --- -->
            <div id="areaKomentarFajum" style="margin-top: 15px; border-top: 1px solid #333; padding-top: 10px; text-align: left; width: 100%;">
                <h4 style="color: #00ffcc; margin: 0 0 10px 0; font-size: 13px;">💬 Testimoni Pengunjung</h4>
                
                <div id="kotakListKmt" style="max-height: 100px; overflow-y: auto; background: rgba(0,0,0,0.4); padding: 6px; border-radius: 4px; margin-bottom: 8px; font-size: 11px;">
                    <div style="color: #666; text-align: center; padding: 5px;">Memuat testimoni...</div>
                </div>

                <div style="display: flex; flex-direction: column; gap: 4px; width: 100%;">
                    <input type="text" id="inputNamaKmt" placeholder="Nama Anda..." style="width: 100%; background: #222; border: 1px solid #444; color: #fff; padding: 4px 6px; border-radius: 4px; font-size: 11px; box-sizing: border-box;">
                    <textarea id="inputIsiKmt" placeholder="Tulis testimoni di sini..." rows="2" style="width: 100%; background: #222; border: 1px solid #444; color: #fff; padding: 4px 6px; border-radius: 4px; font-size: 11px; box-sizing: border-box; resize: none;"></textarea>
                    <button type="button" onclick="eksekusiKirimKmt()" style="background: #00ffcc; color: #000; border: none; padding: 5px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 11px; width: 100%;">Kirim Testimoni 🚀</button>
                </div>
            </div>

            <script>
                async function eksekusiMuatKmt() {
                    try {
                        const res = await fetch('/api/komentar');
                        const data = await res.json();
                        const list = document.getElementById('kotakListKmt');
                        if (list) {
                            list.innerHTML = '';
                            if(data.length === 0) {
                                list.innerHTML = '<div style="color:#666;text-align:center;">Belum ada testimoni.</div>';
                                return;
                            }
                            data.forEach(k => {
                                list.innerHTML += '<div style="border-bottom: 1px solid #222; padding: 4px 0; line-height: 1.3;"><strong style="color:#00ffcc;">' + k.nama + '</strong> <span style="font-size:9px;color:#555;">(' + k.waktu + ')</span><p style="margin:2px 0 0 0;color:#ddd;">' + k.isi + '</p></div>';
                            });
                            list.scrollTop = list.scrollHeight;
                        }
                    } catch (e) { console.error("Error muat:", e); }
                }

                async function eksekusiKirimKmt() {
                    const nama = document.getElementById('inputNamaKmt').value;
                    const isi = document.getElementById('inputIsiKmt').value;
                    if(!nama || !isi) return alert('Nama dan isi testimoni jangan kosong ya bang!');
                    try {
                        await fetch('/api/komentar', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ nama, isi })
                        });
                        document.getElementById('inputNamaKmt').value = '';
                        document.getElementById('inputIsiKmt').value = '';
                        eksekusiMuatKmt();
                    } catch (e) { console.error("Error kirim:", e); }
                }
                document.addEventListener('DOMContentLoaded', eksekusiMuatKmt);
            </script>
        </div>
    </div>
        <div class="brand-header">
            <div class="brand-title-container">
                <!-- LOGO ANDA DIPANGGIL DI SINI -->
                <img src="/fajumkajang.jpg" alt="Logo Fajum Kajang" class="brand-logo" onerror="this.style.display='none'">
                <h1>FAJUM KAJANG 🎧</h1>
            </div>
            <p>Mastering & Document Multi-Converter Station</p>
        </div>

        <div class="studio-box">
            <form action="/convert" method="POST" enctype="multipart/form-data">
                <div class="section-title">🎛️ INPUT TRACK / FILE</div>
                <label>Pilih File (Dokumen / Audio):</label>
                <input type="file" name="file" required>

                <div class="section-title">⚙️ MIXING & AUDIO TUNING</div>
                <label>Format Target:</label>
                <select name="target_format">
                <optgroup label="📄 FORMAT DOKUMEN">
                    <option value="docx_to_pdf">Word (.docx) ke PDF</option>
                    <option value="pdf_to_docx">PDF ke Word (.docx)</option>
                    <option value="xlsx_to_pdf">Excel (.xlsx) ke PDF</option>
                    <option value="pdf_to_xlsx">PDF ke Excel (.xlsx)</option>
                    <option value="img_to_pdf">Foto (JPG/PNG) ke PDF</option>
                </optgroup>

                    <optgroup label="🎵 FORMAT AUDIO & MUSIK">
                        <option value="wav">Convert Audio Lain ke WAV (.wav)</option>
                        <option value="mp3">Convert WAV / Audio Lain ke MP3 (.mp3)</option>
                        <option value="flac">Convert Audio ke FLAC (Studio Master)</option>
                    </optgroup>
                    <optgroup label="🎬 FORMAT VIDEO">
                        <option value="mp4">Convert ke MP4 (Video)</option>
                        <option value="extract_audio">Ekstrak Video menjadi Audio (MP3)</option>
                    </optgroup>
                </select>

                <label>Audio Bitrate / Kualitas (Khusus Audio):</label>
                <select name="audio_quality">
                    <option value="320k">High Quality Studio (320 kbps)</option>
                    <option value="192k">Standard Quality (192 kbps)</option>
                    <option value="128k">Radio/Eco Quality (128 kbps)</option>
                </select>

                <button type="submit">RENDER FILE 🚀</button>
            </form>
        </div>
    </body>
            <div class="spark-container" id="sparks"></div>
        <script>
            function buatPercikan() {
                const kontainer = document.getElementById('sparks');
                for (let i = 0; i < 20; i++) {
                    setTimeout(() => {
                        const spark = document.createElement('div');
                        spark.classList.add('spark');
                        spark.style.left = Math.random() * 100 + 'vw';
                        const ukuran = Math.random() * 5 + 2;
                        spark.style.width = ukuran + 'px';
                        spark.style.height = ukuran + 'px';
                        spark.style.animationDuration = (Math.random() * 3 + 4) + 's';
                        kontainer.appendChild(spark);
                        setTimeout(() => { spark.remove(); }, 7000);
                    }, Math.random() * 3000);
                }
            }
            buatPercikan();
            setInterval(buatPercikan, 4000);
        function bukaSidebar() {
            const sidebar = document.getElementById('modalProfil');
            const overlay = document.getElementById('sidebarOverlay');
            
            if (sidebar) {
                sidebar.classList.add('active');
                sidebar.style.display = 'block';
                sidebar.style.visibility = 'visible';
                sidebar.style.opacity = '1';
            }
            if (overlay) {
                overlay.style.display = 'block';
                overlay.style.visibility = 'visible';
                overlay.style.opacity = '1';
            }
        }

        function tutupSidebar() {
            const sidebar = document.getElementById('modalProfil');
            const overlay = document.getElementById('sidebarOverlay');
            
            if (sidebar) {
                sidebar.classList.remove('active');
                sidebar.style.display = 'none';
            }
            if (overlay) {
                overlay.style.display = 'none';
            }
        }

        function tutupProfil() {
            tutupSidebar();
        }

        </script>
    </body>
</html>

    `);
});

// 2. LOGIKA BACKEND PROSES KONVERSI MULTI-FILE
app.post('/convert', upload.single('file'), async (req, res) => {
    // // --- MEMASTIKAN FOLDER UPLOADS SUDAH TERBUAT (ANTI-ERROR) ---
    const dirUploads = path.join(__dirname, 'uploads');
    if (!fs.existsSync(dirUploads)){
        fs.mkdirSync(dirUploads, { recursive: true });
    }

    if (!req.file) {
        return res.status(400).send('Tidak ada file yang di-upload.');
    }

    const inputPath = req.file.path;
    const originalName = req.file.originalname;
    const fileExt = path.extname(originalName).toLowerCase();
    
    // Menangkap data pilihan menu dari depan website abang
    let targetFormat = req.body.target_format || req.body.targetFormat || req.body.format;
    const audioQuality = req.body.audio_quality || '192k';

    // Wadah nama file hasil render akhir
    let outputPath = '';

    try {
        const outputDir = path.join(__dirname, 'uploads');
        const outputFilename = `fajum_${Date.now()}`;

        // ==========================================
        // 📄 KELOMPOK 1: FORMAT DOKUMEN (LIBREOFFICE CORE MURNI)
        // ==========================================
        const { exec } = require('child_process');

        // A. Jalur Excel (.xlsx) ke PDF Murni
        if (targetFormat === 'xlsx_to_pdf' || (targetFormat === 'pdf' && fileExt === '.xlsx')) {
            const perintah = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe" --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`;
            exec(perintah, (err) => {
                if (err) return res.status(500).send('Gagal konversi Excel ke PDF. Pastikan LibreOffice terinstall.');
                const namaTanpaEkstensi = originalName.substring(0, originalName.lastIndexOf('.'));
                const fileHasilPdf = path.join(outputDir, `${path.basename(inputPath, fileExt)}.pdf`);
                const fileUnduhanFinal = path.join(outputDir, `fajum_${namaTanpaEkstensi}.pdf`);

                if (fs.existsSync(fileHasilPdf)) {
                    fs.renameSync(fileHasilPdf, fileUnduhanFinal);
                    res.setHeader('Content-Type', 'application/pdf');
                    return res.download(fileUnduhanFinal, `fajum_${namaTanpaEkstensi}.pdf`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(fileUnduhanFinal)) fs.unlinkSync(fileUnduhanFinal);
                    });
                } else {
                    return res.status(500).send('Berkas PDF tidak ditemukan.');
                }
            });
        }

        // B. Jalur Word (.docx) ke PDF Murni
        else if (targetFormat === 'docx_to_pdf' || (fileExt === '.docx' && targetFormat === 'pdf')) {
            const perintah = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe" --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`;
            exec(perintah, (err) => {
                if (err) return res.status(500).send('Gagal konversi Word ke PDF.');
                const namaTanpaEkstensi = originalName.substring(0, originalName.lastIndexOf('.'));
                const fileHasilPdf = path.join(outputDir, `${path.basename(inputPath, fileExt)}.pdf`);
                const fileUnduhanFinal = path.join(outputDir, `fajum_${namaTanpaEkstensi}.pdf`);

                if (fs.existsSync(fileHasilPdf)) {
                    fs.renameSync(fileHasilPdf, fileUnduhanFinal);
                    res.setHeader('Content-Type', 'application/pdf');
                    return res.download(fileUnduhanFinal, `fajum_${namaTanpaEkstensi}.pdf`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(fileUnduhanFinal)) fs.unlinkSync(fileUnduhanFinal);
                    });
                }
            });
        }

        // C. Jalur PDF ke Word (.docx) atau PDF ke Excel (.xlsx)
        else if (targetFormat === 'pdf_to_docx' || targetFormat === 'pdf_to_xlsx') {
            const formatTujuan = targetFormat === 'pdf_to_docx' ? 'docx' : 'xlsx';
            const perintah = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe" --headless --convert-to ${formatTujuan} "${inputPath}" --outdir "${outputDir}"`;
            exec(perintah, (err) => {
                if (err) return res.status(500).send(`Gagal konversi PDF ke ${formatTujuan.toUpperCase()}.`);
                const namaTanpaEkstensi = originalName.substring(0, originalName.lastIndexOf('.'));
                const fileHasil = path.join(outputDir, `${path.basename(inputPath, fileExt)}.${formatTujuan}`);
                const fileUnduhanFinal = path.join(outputDir, `fajum_${namaTanpaEkstensi}.${formatTujuan}`);

                if (fs.existsSync(fileHasil)) {
                    fs.renameSync(fileHasil, fileUnduhanFinal);
                    return res.download(fileUnduhanFinal, `fajum_${namaTanpaEkstensi}.${formatTujuan}`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(fileUnduhanFinal)) fs.unlinkSync(fileUnduhanFinal);
                    });
                }
            });
        }

        // D. Jalur Foto (JPG/PNG) ke PDF Murni
        else if (targetFormat === 'img_to_pdf' || (['.jpg', '.jpeg', '.png'].includes(fileExt) && targetFormat === 'pdf')) {
            const perintah = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe" --headless --convert-to pdf "${inputPath}" --outdir "${outputDir}"`;
            exec(perintah, (err) => {
                if (err) return res.status(500).send('Gagal konversi Foto ke PDF.');
                const namaTanpaEkstensi = originalName.substring(0, originalName.lastIndexOf('.'));
                const fileHasilPdf = path.join(outputDir, `${path.basename(inputPath, fileExt)}.pdf`);
                const fileUnduhanFinal = path.join(outputDir, `fajum_${namaTanpaEkstensi}.pdf`);

                if (fs.existsSync(fileHasilPdf)) {
                    fs.renameSync(fileHasilPdf, fileUnduhanFinal);
                    res.setHeader('Content-Type', 'application/pdf');
                    return res.download(fileUnduhanFinal, `fajum_${namaTanpaEkstensi}.pdf`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(fileUnduhanFinal)) fs.unlinkSync(fileUnduhanFinal);
                    });
                }
            });
        }

        // ==========================================
        // 🎵 BAGIAN 2: FORMAT AUDIO & MUSIK (FFmpeg)
        // ==========================================
        
        // 1. Convert Audio Lain ke WAV (.wav)
        else if (targetFormat === 'wav' || targetFormat === 'to_wav') {
            outputPath = path.join(__dirname, 'uploads', `${outputFilename}.wav`);
            ffmpeg(inputPath).toFormat('wav').audioBitrate(audioQuality)
                .on('end', () => {
                    res.download(outputPath, `fajum_${originalName.substring(0, originalName.lastIndexOf('.'))}.wav`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
                    });
                })
                .on('error', (err) => { console.error(err); res.status(500).send('Gagal render ke WAV.'); })
                .save(outputPath);
        }

        // 2. Convert WAV / Audio Lain ke MP3 (.mp3)
        else if (targetFormat === 'mp3' || targetFormat === 'to_mp3') {
            outputPath = path.join(__dirname, 'uploads', `${outputFilename}.mp3`);
            ffmpeg(inputPath).toFormat('mp3').audioBitrate(audioQuality)
                .on('end', () => {
                    res.download(outputPath, `fajum_${originalName.substring(0, originalName.lastIndexOf('.'))}.mp3`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
                    });
                })
                .on('error', (err) => { console.error(err); res.status(500).send('Gagal render ke MP3.'); })
                .save(outputPath);
        }

        // 3. Convert Audio ke FLAC (Studio Master)
        else if (targetFormat === 'flac' || targetFormat === 'to_flac') {
            outputPath = path.join(__dirname, 'uploads', `${outputFilename}.flac`);
            ffmpeg(inputPath).toFormat('flac')
                .on('end', () => {
                    res.download(outputPath, `fajum_${originalName.substring(0, originalName.lastIndexOf('.'))}.flac`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
                    });
                })
                .on('error', (err) => { console.error(err); res.status(500).send('Gagal render ke FLAC.'); })
                .save(outputPath);
        }

        // ==========================================
        // 🎬 BAGIAN 3: FORMAT VIDEO (FFmpeg)
        // ==========================================
        
        // 1. Convert ke MP4 (Video)
        else if (targetFormat === 'mp4' || targetFormat === 'to_mp4') {
            outputPath = path.join(__dirname, 'uploads', `${outputFilename}.mp4`);
            ffmpeg(inputPath).toFormat('mp4').videoCodec('libx264').audioCodec('aac')
                .on('end', () => {
                    res.download(outputPath, `fajum_${originalName.substring(0, originalName.lastIndexOf('.'))}.mp4`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
                    });
                })
                .on('error', (err) => { console.error(err); res.status(500).send('Gagal render ke MP4 Video.'); })
                .save(outputPath);
        }

        // 2. Ekstrak Video menjadi Audio (MP3)
        else if (targetFormat === 'video_to_mp3' || targetFormat === 'extract_mp3') {
            outputPath = path.join(__dirname, 'uploads', `${outputFilename}.mp3`);
            ffmpeg(inputPath).toFormat('mp3').noVideo().audioBitrate(audioQuality)
                .on('end', () => {
                    res.download(outputPath, `fajum_${originalName.substring(0, originalName.lastIndexOf('.'))}_audio.mp3`, () => {
                        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
                        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
                    });
                })
                .on('error', (err) => { console.error(err); res.status(500).send('Gagal mengekstrak audio dari video.'); })
                .save(outputPath);
        }

        // Jika tidak ada format yang cocok
        else {
            if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
            return res.status(400).send('Pilihan konversi tidak dikenali sistem.');
        }

    } catch (error) {
        console.error(error);
        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
        res.status(500).send('Terjadi kesalahan sistem di Studio Fajum Kajang.');
    }
});

// Jalankan Server
app.listen(PORT, () => {
    console.log(`Server Fajum Kajang berjalan di http://localhost:3000`);
});
