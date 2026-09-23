document.addEventListener('DOMContentLoaded', () => {
    // 1. Logika Login dengan Enkripsi Base64
    const formLogin = document.getElementById('formLogin');
    if (formLogin) {
        formLogin.addEventListener('submit', (e) => {
            e.preventDefault();
            const role = document.getElementById('role').value;
            const kode = document.getElementById('kodeAkses').value;

            // Mengubah input password ke format Base64 ('walas123' -> 'd2FsYXMxMjM=', 'bk123' -> 'YmsxMjM=')
            const encodedKode = btoa(kode);

            if (role === 'walas' && encodedKode === 'd2FsYXMxMjM=') {
                localStorage.setItem('sessionRole', 'walas');
                window.location.href = 'walas.html';
            } else if (role === 'bk' && encodedKode === 'YmsxMjM=') {
                localStorage.setItem('sessionRole', 'bk');
                window.location.href = 'bk.html';
            } else {
                alert('Kode akses salah atau peran tidak sesuai!');
            }
        });
    }

    // 2. Proteksi Halaman Wali Kelas
    if (window.location.pathname.includes('walas.html')) {
        if (localStorage.getItem('sessionRole') !== 'walas') {
            alert('Akses ditolak! Silakan login sebagai Wali Kelas.');
            window.location.href = 'index.html';
            return;
        }

        // Set tanggal default hari ini di filter
        const today = new Date().toISOString().split('T')[0];
        const filterEl = document.getElementById('filterTanggal');
        if (filterEl) {
            filterEl.value = today;
        }

        loadDashboardWaliKelas();

        // Form Tambah Siswa Satuan
        const formTambah = document.getElementById('formTambahSiswa');
        if (formTambah) {
            formTambah.addEventListener('submit', (e) => {
                e.preventDefault();
                const nis = document.getElementById('inputNis').value;
                const nama = document.getElementById('inputNama').value;
                const gender = document.getElementById('inputGender').value;

                let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
                daftarSiswa.push({ nis, nama, gender, statusBk: 'Aman' });
                localStorage.setItem('daftarSiswa', JSON.stringify(daftarSiswa));

                // Tutup modal & reset form
                const modalEl = document.getElementById('modalTambahSiswa');
                const modal = bootstrap.Modal.getInstance(modalEl);
                if (modal) modal.hide();
                formTambah.reset();
                loadDashboardWaliKelas();
            });
        }
    }

    // 3. Proteksi Halaman Guru BK (Mencakup bk.html dan bk-datasiswa.html)
    if (window.location.pathname.includes('bk.html') || window.location.pathname.includes('bk-datasiswa.html')) {
        if (localStorage.getItem('sessionRole') !== 'bk') {
            alert('Akses ditolak! Khusus Guru BK.');
            window.location.href = 'index.html';
            return;
        }

        // Jika berada di halaman data siswa, muat tabel dan tombol filter/export
        if (window.location.pathname.includes('bk-datasiswa.html')) {
            const filterEl = document.getElementById('filterTanggalBK');
            const today = new Date().toISOString().split('T')[0];
            if (filterEl && !filterEl.value) {
                filterEl.value = today;
            }

            loadDataBK();

            // Event listener untuk tombol filter BK agar berfungsi dengan baik
            const btnFilterBK = document.getElementById('btnFilterBK');
            if (btnFilterBK) {
                btnFilterBK.addEventListener('click', () => {
                    loadDataBK();
                });
            }

            const btnExport = document.getElementById('btnExport');
            if (btnExport) {
                btnExport.addEventListener('click', () => {
                    let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
                    if (daftarSiswa.length === 0) {
                        alert('Belum ada data siswa untuk diexport!');
                        return;
                    }

                    const worksheet = XLSX.utils.json_to_sheet(daftarSiswa);
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, "Rekap BK");

                    XLSX.writeFile(workbook, `Rekap_Data_BK_${new Date().toISOString().slice(0,10)}.xlsx`);
                });
            }
        }
    }
});

// ==========================================
// FUNGSI-FUNGSI WALI KELAS (DASHBOARD WALAS)
// ==========================================

// Memuat data ke Dashboard Wali Kelas berdasarkan tanggal yang dipilih
function loadDashboardWaliKelas() {
    const daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
    const filterEl = document.getElementById('filterTanggal');
    const today = new Date().toISOString().split('T')[0];
    
    if (filterEl && !filterEl.value) {
        filterEl.value = today;
    }
    const tanggalAktif = filterEl ? filterEl.value : today;

    // Perbarui teks label statistik secara real-time mengikuti tanggal aktif
    const labelEl = document.getElementById('labelTanggalHadir');
    const infoEl = document.getElementById('infoTanggalAbsen');
    if (labelEl) labelEl.innerText = `HADIR PADA TANGGAL ${tanggalAktif}`;
    if (infoEl) infoEl.innerText = `Mengelola absensi untuk tanggal: ${tanggalAktif}`;

    // Cek apakah data pada tanggal ini sudah disimpan
    const sudahDisimpan = localStorage.getItem('saved_absensi_' + tanggalAktif) === 'true';
    let absensiHarian = JSON.parse(localStorage.getItem('absensi_' + tanggalAktif)) || {};

    const statTotal = document.getElementById('statTotal');
    if (statTotal) statTotal.innerText = daftarSiswa.length;
    
    let totalHadir = 0;
    let totalBk = 0;

    // Hitung jumlah kehadiran secara akurat
    daftarSiswa.forEach((siswa) => {
        const statusAbsen = absensiHarian[siswa.nis] || 'Hadir'; 
        if (statusAbsen === 'Hadir' || statusAbsen === 'Terlambat') {
            totalHadir++;
        }
        if (siswa.statusBk === 'Dalam Penanganan') {
            totalBk++;
        }
    });

    const statHadir = document.getElementById('statHadir');
    const statBk = document.getElementById('statBk');
    if (statHadir) statHadir.innerText = totalHadir;
    if (statBk) statBk.innerText = totalBk;

    const tbody = document.querySelector('#tabelSiswa tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (daftarSiswa.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">Data Tidak Tersedia</td></tr>`;
        if (statHadir) statHadir.innerText = 0;
        if (statBk) statBk.innerText = 0;
        return;
    }

    // Jika sudah disimpan, sembunyikan tabel input interaktif dan tampilkan pemberitahuan
    if (sudahDisimpan) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center text-success py-4 fw-bold">
            ✓ Data absensi untuk tanggal ${tanggalAktif} telah berhasil disimpan dan dikunci.<br>
            <span class="text-muted small fw-normal">Silakan klik tombol "👁️ Lihat Data Tersimpan" di atas untuk melihat rekapitulasi data.</span>
        </td></tr>`;
        return;
    }

    daftarSiswa.forEach((siswa, index) => {
        const statusAbsen = absensiHarian[siswa.nis] || 'Hadir'; 

        const row = `<tr>
            <td>${index + 1}</td>
            <td>${siswa.nis}</td>
            <td>${siswa.nama}</td>
            <td><span class="badge bg-secondary">${siswa.gender}</span></td>
            <td>
                <select class="form-select form-select-sm status-absen-select" data-nis="${siswa.nis}">
                    <option value="Hadir" ${statusAbsen === 'Hadir' ? 'selected' : ''}>Hadir</option>
                    <option value="Terlambat" ${statusAbsen === 'Terlambat' ? 'selected' : ''}>Terlambat</option>
                    <option value="Izin" ${statusAbsen === 'Izin' ? 'selected' : ''}>Izin</option>
                    <option value="Sakit" ${statusAbsen === 'Sakit' ? 'selected' : ''}>Sakit</option>
                    <option value="Alpha" ${statusAbsen === 'Alpha' ? 'selected' : ''}>Alpha</option>
                </select>
            </td>
            <td><span class="badge ${siswa.statusBk === 'Aman' ? 'bg-success' : 'bg-warning text-dark'}">${siswa.statusBk}</span></td>
            <td>
                <button class="btn btn-sm btn-outline-primary" onclick="ubahStatusBk(${index})">Rujuk ke BK</button>
            </td>
        </tr>`;
        tbody.innerHTML += row;
    });
}

// Simpan Semua Absen di Layar & Kunci Form Input
function simpanSemuaAbsen() {
    const filterEl = document.getElementById('filterTanggal');
    const tanggalAktif = filterEl ? filterEl.value : new Date().toISOString().split('T')[0];
    const selects = document.querySelectorAll('.status-absen-select');
    let absensiHarian = {};

    selects.forEach(sel => {
        const nis = sel.getAttribute('data-nis');
        absensiHarian[nis] = sel.value;
    });

    localStorage.setItem('absensi_' + tanggalAktif, JSON.stringify(absensiHarian));
    localStorage.setItem('saved_absensi_' + tanggalAktif, 'true'); // Menandai data sudah disimpan

    alert('Data absensi tanggal ' + tanggalAktif + ' berhasil disimpan!');
    loadDashboardWaliKelas(); // Refresh agar tabel input otomatis hilang
}

// Menampilkan atau Menyembunyikan Kartu Data Tersimpan Wali Kelas
function toggleLihatDataTersimpan() {
    const card = document.getElementById('cardDataTersimpan');
    if (!card) return;

    if (card.style.display === 'none' || card.style.display === '') {
        const filterEl = document.getElementById('filterTanggal');
        const tanggalAktif = filterEl ? filterEl.value : new Date().toISOString().split('T')[0];
        
        document.getElementById('labelTanggalTersimpan').innerText = tanggalAktif;
        
        let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
        let absensiHarian = JSON.parse(localStorage.getItem('absensi_' + tanggalAktif)) || {};
        
        const tbody = document.querySelector('#tabelDataTersimpan tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (daftarSiswa.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted">Belum ada data siswa.</td></tr>`;
        } else {
            daftarSiswa.forEach((siswa, index) => {
                const statusAbsen = absensiHarian[siswa.nis] || 'Hadir';
                
                let badgeColor = 'bg-success';
                if (statusAbsen === 'Terlambat') badgeColor = 'bg-warning text-dark';
                else if (statusAbsen === 'Izin') badgeColor = 'bg-info text-dark';
                else if (statusAbsen === 'Sakit') badgeColor = 'bg-primary';
                else if (statusAbsen === 'Alpha') badgeColor = 'bg-danger';

                tbody.innerHTML += `<tr>
                    <td>${index + 1}</td>
                    <td>${siswa.nis}</td>
                    <td>${siswa.nama}</td>
                    <td><span class="badge bg-secondary">${siswa.gender}</span></td>
                    <td><span class="badge ${badgeColor}">${statusAbsen}</span></td>
                </tr>`;
            });
        }

        card.style.display = 'block';
        card.scrollIntoView({ behavior: 'smooth' });
    } else {
        card.style.display = 'none';
    }
}

// Tombol Filter Tanggal Absensi dipencet (Wali Kelas)
function filterDataAbsen() {
    const filterEl = document.getElementById('filterTanggal');
    if (!filterEl) return;
    
    // Sembunyikan kartu data tersimpan jika sedang membuka tanggal lain
    const card = document.getElementById('cardDataTersimpan');
    if (card) card.style.display = 'none';

    loadDashboardWaliKelas();
}

// Hapus Semua Siswa
function hapusSemuaSiswa() {
    if (confirm('Yakin ingin menghapus seluruh data siswa binaan?')) {
        localStorage.removeItem('daftarSiswa');
        loadDashboardWaliKelas();
    }
}

// Ubah Status BK dari Wali Kelas Berdasarkan Index Baris
function ubahStatusBk(index) {
    let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
    if (daftarSiswa[index]) {
        daftarSiswa[index].statusBk = daftarSiswa[index].statusBk === 'Aman' ? 'Dalam Penanganan' : 'Aman';
    }
    localStorage.setItem('daftarSiswa', JSON.stringify(daftarSiswa));
    loadDashboardWaliKelas();
}

// ==========================================
// FUNGSI-FUNGSI GURU BK
// ==========================================

// Memuat data ke tabel di halaman BK (bk-datasiswa.html) berdasarkan filter tanggal
function loadDataBK() {
    let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
    const filterEl = document.getElementById('filterTanggalBK');
    const today = new Date().toISOString().split('T')[0];
    
    if (filterEl && !filterEl.value) {
        filterEl.value = today;
    }
    const tanggalAktif = filterEl ? filterEl.value : today;

    // Cek apakah data penanganan BK pada tanggal ini sudah dikunci/disimpan
    const sudahDisimpan = localStorage.getItem('saved_penanganan_' + tanggalAktif) === 'true';
    let dataPenanganan = JSON.parse(localStorage.getItem('penanganan_' + tanggalAktif)) || {};

    const tableElement = document.querySelector('#tabelRekapBK') || document.querySelector('#tabelRekap');
    if (!tableElement) return;
    
    const tbody = tableElement.querySelector('tbody') || tableElement;
    tbody.innerHTML = '';

    if (daftarSiswa.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">Data Tidak Tersedia</td></tr>`;
        return;
    }

    // Jika sudah disimpan/dikunci, sembunyikan tabel interaktif dan tampilkan pesan sukses
    if (sudahDisimpan) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center text-success py-4 fw-bold">
            ✓ Data penanganan BK untuk tanggal ${tanggalAktif} telah berhasil disimpan dan dikunci.<br>
            <span class="text-muted small fw-normal">Silakan klik tombol "👁️ Lihat Data Tersimpan" di atas untuk melihat rekapitulasinya.</span>
        </td></tr>`;
        return;
    }

    daftarSiswa.forEach((siswa, index) => {
        const record = dataPenanganan[index] || { statusBk: 'Aman', tanggal: '-' };
        let statusBk = record.statusBk;
        let tglPenanganan = record.tanggal;

        let tombolAksi = '';
        if (statusBk === 'Dalam Penanganan') {
            tombolAksi = `<button class="btn btn-sm btn-success" onclick="selesaiPenangananBK(${index})">Selesai Penanganan</button>`;
        } else {
            tombolAksi = `<button class="btn btn-sm btn-warning text-dark" onclick="selesaiPenangananBK(${index})">Proses Penanganan</button>`;
        }

        const row = `<tr>
            <td>${index + 1}</td>
            <td>${siswa.nis}</td>
            <td>${siswa.nama}</td>
            <td><span class="badge bg-secondary">${siswa.gender}</span></td>
            <td><span class="badge ${statusBk === 'Aman' ? 'bg-success' : 'bg-warning text-dark'}">${statusBk}</span></td>
            <td>${tglPenanganan}</td>
            <td>${tombolAksi}</td>
        </tr>`;
        tbody.innerHTML += row;
    });
}

// Fungsi Guru BK Mengubah Status Penanganan Berdasarkan Tanggal Aktif
function selesaiPenangananBK(index) {
    const filterEl = document.getElementById('filterTanggalBK');
    const today = new Date().toISOString().split('T')[0];
    const tanggalAktif = filterEl ? (filterEl.value || today) : today;

    let dataPenanganan = JSON.parse(localStorage.getItem('penanganan_' + tanggalAktif)) || {};
    let currentRecord = dataPenanganan[index] || { statusBk: 'Aman', tanggal: '-' };

    if (currentRecord.statusBk === 'Aman') {
        currentRecord.statusBk = 'Dalam Penanganan';
        currentRecord.tanggal = tanggalAktif; 
    } else {
        currentRecord.statusBk = 'Aman';
        currentRecord.tanggal = '-'; 
    }

    dataPenanganan[index] = currentRecord;
    localStorage.setItem('penanganan_' + tanggalAktif, JSON.stringify(dataPenanganan));
    loadDataBK();
}

// Fungsi Guru BK Menyimpan & Mengunci Data Penanganan
function simpanDataBK() {
    const filterEl = document.getElementById('filterTanggalBK');
    const today = new Date().toISOString().split('T')[0];
    const tanggalAktif = filterEl ? (filterEl.value || today) : today;

    localStorage.setItem('saved_penanganan_' + tanggalAktif, 'true');
    alert('Data penanganan BK tanggal ' + tanggalAktif + ' berhasil disimpan dan dikunci!');
    loadDataBK();
}

// Menampilkan atau Menyembunyikan Kartu Data BK Tersimpan
function toggleLihatDataTersimpanBK() {
    const card = document.getElementById('cardDataTersimpanBK');
    if (!card) return;

    if (card.style.display === 'none' || card.style.display === '') {
        const filterEl = document.getElementById('filterTanggalBK');
        const today = new Date().toISOString().split('T')[0];
        const tanggalAktif = filterEl ? (filterEl.value || today) : today;
        
        document.getElementById('labelTanggalTersimpanBK').innerText = tanggalAktif;
        
        let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
        let dataPenanganan = JSON.parse(localStorage.getItem('penanganan_' + tanggalAktif)) || {};
        
        const tbody = document.querySelector('#tabelDataTersimpanBK tbody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (daftarSiswa.length === 0) {
            tbody.innerHTML = `<tr><td colspan="6" class="text-center text-muted">Belum ada data siswa.</td></tr>`;
        } else {
            daftarSiswa.forEach((siswa, index) => {
                const record = dataPenanganan[index] || { statusBk: 'Aman', tanggal: '-' };
                tbody.innerHTML += `<tr>
                    <td>${index + 1}</td>
                    <td>${siswa.nis}</td>
                    <td>${siswa.nama}</td>
                    <td><span class="badge bg-secondary">${siswa.gender}</span></td>
                    <td><span class="badge ${record.statusBk === 'Aman' ? 'bg-success' : 'bg-warning text-dark'}">${record.statusBk}</span></td>
                    <td>${record.tanggal}</td>
                </tr>`;
            });
        }

        card.style.display = 'block';
        card.scrollIntoView({ behavior: 'smooth' });
    } else {
        card.style.display = 'none';
    }
}

// ==========================================
// FUNGSI UMUM & UPLOAD EXCEL
// ==========================================

// Fungsi Logout Universal
function logout() {
    localStorage.removeItem('sessionRole');
    window.location.href = 'index.html';
}

// Fungsi untuk Memproses Upload File Excel secara Otomatis
function prosesUploadExcel() {
    const fileInput = document.getElementById('fileExcelInput');
    if (!fileInput.files || fileInput.files.length === 0) {
        alert('Silakan pilih file Excel terlebih dahulu!');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            let daftarSiswa = JSON.parse(localStorage.getItem('daftarSiswa')) || [];
            let count = 0;
            let nisCounter = daftarSiswa.length + 101;

            rows.forEach((row) => {
                let nama = '';
                let gender = 'L';

                for (let i = 0; i < row.length; i++) {
                    let val = String(row[i]).trim().toUpperCase();
                    if (val === 'L' || val === 'P') {
                        gender = val;
                        if (i > 0) {
                            nama = String(row[i - 1]).trim();
                        }
                        break;
                    }
                }

                if (nama && !['NAMA', 'L/P', 'NO', 'DATA PRIBADI SISWA'].includes(nama.toUpperCase())) {
                    let nis = 'NIS-' + nisCounter++;
                    const sudahAda = daftarSiswa.some(s => s.nama.toLowerCase() === nama.toLowerCase());
                    if (!sudahAda) {
                        daftarSiswa.push({ nis, nama, gender, statusBk: 'Aman' });
                        count++;
                    }
                }
            });

            if (count === 0) {
                alert('Gagal mendeteksi kolom Nama dan L/P. Pastikan format tabel Excel sesuai.');
                return;
            }

            localStorage.setItem('daftarSiswa', JSON.stringify(daftarSiswa));

            const modalEl = document.getElementById('modalPasteExcel');
            const modal = bootstrap.Modal.getInstance(modalEl);
            if (modal) modal.hide();
            fileInput.value = '';

            alert(`Berhasil mengimpor ${count} data siswa dari Excel secara otomatis!`);
            loadDashboardWaliKelas();

        } catch (error) {
            console.error(error);
            alert('Terjadi kesalahan saat membaca file Excel!');
        }
    };

    reader.readAsArrayBuffer(file);
}

// ==========================================
// FUNGSI JADWAL & DOKUMENTASI BK
// ==========================================

// Proteksi & Inisialisasi Halaman Jadwal BK
document.addEventListener('DOMContentLoaded', () => {
    if (window.location.pathname.includes('bk-jadwal.html')) {
        if (localStorage.getItem('sessionRole') !== 'bk') {
            alert('Akses ditolak! Khusus Guru BK.');
            window.location.href = 'index.html';
            return;
        }

        // Set tanggal default hari ini pada form
        const tglInput = document.getElementById('agendaTanggal');
        if (tglInput) {
            tglInput.value = new Date().toISOString().split('T')[0];
        }

        loadDaftarAgenda();
    }
});

// Berpindah Tab (Tambah vs Daftar)
function tampilTab(tab) {
    const secTambah = document.getElementById('sectionTambah');
    const secDaftar = document.getElementById('sectionDaftar');
    const btnTambah = document.getElementById('btnTabTambah');
    const btnDaftar = document.getElementById('btnTabDaftar');

    if (tab === 'tambah') {
        secTambah.style.display = 'block';
        secDaftar.style.display = 'none';
        btnTambah.className = 'btn btn-primary';
        btnDaftar.className = 'btn btn-outline-secondary';
    } else {
        secTambah.style.display = 'none';
        secDaftar.style.display = 'block';
        btnTambah.className = 'btn btn-outline-secondary';
        btnDaftar.className = 'btn.btn-primary' || 'btn btn-primary';
        loadDaftarAgenda();
    }
}

// Menyimpan Agenda Baru
function simpanAgenda(e) {
    e.preventDefault();
    const tanggal = document.getElementById('agendaTanggal').value;
    const kategori = document.getElementById('agendaKategori').value;
    const judul = document.getElementById('agendaJudul').value;
    const keterangan = document.getElementById('agendaKeterangan').value;
    const fotoInput = document.getElementById('agendaFoto');

    let daftarAgenda = JSON.parse(localStorage.getItem('daftarAgendaBK')) || [];

    if (fotoInput.files && fotoInput.files[0]) {
        const reader = new FileReader();
        reader.onload = function(uploadEvent) {
            const fotoBase64 = uploadEvent.target.result;
            simpanKeStorage(daftarAgenda, { tanggal, kategori, judul, keterangan, foto: fotoBase64 });
        };
        reader.readAsDataURL(fotoInput.files[0]);
    } else {
        simpanKeStorage(daftarAgenda, { tanggal, kategori, judul, keterangan, foto: null });
    }
}

function simpanKeStorage(daftarAgenda, agendaBaru) {
    daftarAgenda.unshift.apply(daftarAgenda, [agendaBaru]);
    localStorage.setItem('daftarAgendaBK', JSON.stringify(daftarAgenda));

    alert('Agenda & Dokumentasi berhasil disimpan!');
    document.getElementById('formAgenda').reset();
    document.getElementById('agendaTanggal').value = new Date().toISOString().split('T')[0];
    tampilTab('daftar');
}

// Memuat dan Menampilkan Daftar Agenda
function loadDaftarAgenda() {
    let daftarAgenda = JSON.parse(localStorage.getItem('daftarAgendaBK')) || [];
    const wadah = document.getElementById('wadahDaftarAgenda');
    if (!wadah) return;

    wadah.innerHTML = '';

    if (daftarAgenda.length === 0) {
        wadah.innerHTML = `<div class="col-12 text-center text-muted py-4">Belum ada agenda atau dokumentasi kegiatan yang tersimpan.</div>`;
        return;
    }

    daftarAgenda.forEach((item, index) => {
        let badgeColor = 'bg-primary';
        if (item.kategori === 'Bimbingan Kelompok') badgeColor = 'bg-warning text-dark';
        if (item.kategori === 'Konseling Individual') badgeColor = 'bg-danger';
        if (item.kategori === 'Sosialisasi / Klasikal') badgeColor = 'bg-success';

        let fotoTag = item.foto ? `<img src="${item.foto}" class="img-fluid rounded mb-3" style="max-height: 200px; object-fit: cover; width: 100%;">` : '';

        const card = `
            <div class="col-md-6">
                <div class="card h-100 shadow-sm border">
                    <div class="card-body">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <span class="badge ${badgeColor}">${item.kategori}</span>
                            <small class="text-muted">📅 ${item.tanggal}</small>
                        </div>
                        <h5 class="card-title fw-bold">${item.judul}</h5>
                        <p class="card-text text-muted small" style="white-space: pre-line;">${item.keterangan}</p>
                        ${fotoTag}
                    </div>
                    <div class="card-footer bg-transparent border-top-0 text-end">
                        <button class="btn btn-sm btn-outline-danger" onclick="hapusAgenda(${index})">Hapus Agenda</button>
                    </div>
                </div>
            </div>
        `;
        wadah.innerHTML += card;
    });
}

// Hapus Agenda Satuan
function hapusAgenda(index) {
    if (confirm('Yakin ingin menghapus agenda ini?')) {
        let daftarAgenda = JSON.parse(localStorage.getItem('daftarAgendaBK')) || [];
        daftarAgenda.splice(index, 1);
        localStorage.setItem('daftarAgendaBK', JSON.stringify(daftarAgenda));
        loadDaftarAgenda();
    }
}

// Hapus Semua Agenda
function hapusSemuaAgenda() {
    if (confirm('Yakin ingin menghapus seluruh daftar agenda dan dokumentasi?')) {
        localStorage.removeItem('daftarAgendaBK');
        loadDaftarAgenda();
    }
}

// ==========================================
// FUNGSI ARSIP MATERI BK
// ==========================================

// Proteksi & Inisialisasi Halaman Materi BK
document.addEventListener('DOMContentLoaded', () => {
    if (window.location.pathname.includes('bk-materi.html')) {
        if (localStorage.getItem('sessionRole') !== 'bk') {
            alert('Akses ditolak! Khusus Guru BK.');
            window.location.href = 'index.html';
            return;
        }

        loadDaftarMateri();
    }
});

// Berpindah Tab (Tambah vs Daftar Materi)
function tampilTabMateri(tab) {
    const secTambah = document.getElementById('sectionTambahMateri');
    const secDaftar = document.getElementById('sectionDaftarMateri');
    const btnTambah = document.getElementById('btnTabTambahMateri');
    const btnDaftar = document.getElementById('btnTabDaftarMateri');

    if (tab === 'tambah') {
        secTambah.style.display = 'block';
        secDaftar.style.display = 'none';
        btnTambah.className = 'btn btn-primary';
        btnDaftar.className = 'btn btn-outline-secondary';
    } else {
        secTambah.style.display = 'none';
        secDaftar.style.display = 'block';
        btnTambah.className = 'btn btn-outline-secondary';
        btnDaftar.className = 'btn btn-primary';
        loadDaftarMateri();
    }
}

// Menyimpan Materi Baru
function simpanMateri(e) {
    e.preventDefault();
    const judul = document.getElementById('materiJudul').value;
    const deskripsi = document.getElementById('materiDeskripsi').value;
    const status = document.getElementById('materiStatus').value;

    let daftarMateri = JSON.parse(localStorage.getItem('daftarMateriBK')) || [];
    
    daftarMateri.unshift({ judul, deskripsi, status });
    localStorage.setItem('daftarMateriBK', JSON.stringify(daftarMateri));

    alert('Arsip materi berhasil disimpan!');
    document.getElementById('formMateri').reset();
    tampilTabMateri('daftar');
}

// Memuat dan Menampilkan Daftar Materi
function loadDaftarMateri() {
    let daftarMateri = JSON.parse(localStorage.getItem('daftarMateriBK')) || [];
    const wadah = document.getElementById('wadahDaftarMateri');
    if (!wadah) return;

    wadah.innerHTML = '';

    if (daftarMateri.length === 0) {
        wadah.innerHTML = `<div class="col-12 text-center text-muted py-4">Belum ada arsip materi bimbingan yang tersimpan.</div>`;
        return;
    }

    daftarMateri.forEach((item, index) => {
        let badgeColor = 'bg-success';
        if (item.status === 'Belum Disampaikan') badgeColor = 'bg-warning text-dark';
        if (item.status === 'Draft') badgeColor = 'bg-secondary';

        const card = `
            <div class="col-md-6">
                <div class="card h-100 shadow-sm border">
                    <div class="card-body">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <span class="badge ${badgeColor}">${item.status}</span>
                        </div>
                        <h5 class="card-title fw-bold">${item.judul}</h5>
                        <p class="card-text text-muted small" style="white-space: pre-line;">${item.deskripsi}</p>
                    </div>
                    <div class="card-footer bg-transparent border-top-0 text-end">
                        <button class="btn btn-sm btn-outline-danger" onclick="hapusMateri(${index})">Hapus Materi</button>
                    </div>
                </div>
            </div>
        `;
        wadah.innerHTML += card;
    });
}

// Hapus Materi Satuan
function hapusMateri(index) {
    if (confirm('Yakin ingin menghapus arsip materi ini?')) {
        let daftarMateri = JSON.parse(localStorage.getItem('daftarMateriBK')) || [];
        daftarMateri.splice(index, 1);
        localStorage.setItem('daftarMateriBK', JSON.stringify(daftarMateri));
        loadDaftarMateri();
    }
}

// Hapus Semua Materi
function hapusSemuaMateri() {
    if (confirm('Yakin ingin menghapus seluruh arsip materi bimbingan?')) {
        localStorage.removeItem('daftarMateriBK');
        loadDaftarMateri();
    }
}

// ==========================================
// FUNGSI CATATAN & PENGUMUMAN INTERNAL (INFORMASI)
// ==========================================

// Proteksi & Inisialisasi Halaman Informasi BK
document.addEventListener('DOMContentLoaded', () => {
    if (window.location.pathname.includes('bk-informasi.html')) {
        if (localStorage.getItem('sessionRole') !== 'bk') {
            alert('Akses ditolak! Khusus Guru BK.');
            window.location.href = 'index.html';
            return;
        }

        loadDaftarInformasi();
    }
});

// Berpindah Tab (Buat vs Daftar Informasi)
function tampilTabInfo(tab) {
    const secBuat = document.getElementById('sectionBuatInfo');
    const secDaftar = document.getElementById('sectionDaftarInfo');
    const btnBuat = document.getElementById('btnTabBuatInfo');
    const btnDaftar = document.getElementById('btnTabDaftarInfo');
    const alertSukses = document.getElementById('alertSuksesInfo');

    if (alertSukses) alertSukses.style.display = 'none';

    if (tab === 'buat') {
        secBuat.style.display = 'block';
        secDaftar.style.display = 'none';
        btnBuat.className = 'btn btn-primary';
        btnDaftar.className = 'btn btn-outline-secondary';
    } else {
        secBuat.style.display = 'none';
        secDaftar.style.display = 'block';
        btnBuat.className = 'btn btn-outline-secondary';
        btnDaftar.className = 'btn btn-primary';
        loadDaftarInformasi();
    }
}

// Menyimpan Catatan / Pengumuman Baru
function simpanInformasi(e) {
    e.preventDefault();
    const judul = document.getElementById('infoJudul').value;
    const pesan = document.getElementById('infoPesan').value;
    const tanggal = new Date().toISOString().split('T')[0];

    let daftarInfo = JSON.parse(localStorage.getItem('daftarInformasiBK')) || [];
    
    daftarInfo.unshift({ judul, pesan, tanggal });
    localStorage.setItem('daftarInformasiBK', JSON.stringify(daftarInfo));

    document.getElementById('formInformasi').reset();
    
    // Pindah ke tab daftar dan tampilkan alert sukses
    tampilTabInfo('daftar');
    const alertSukses = document.getElementById('alertSuksesInfo');
    if (alertSukses) {
        alertSukses.style.display = 'block';
        setTimeout(() => { alertSukses.style.display = 'none'; }, 4000);
    }
}

// Memuat dan Menampilkan Daftar Informasi
function loadDaftarInformasi() {
    let daftarInfo = JSON.parse(localStorage.getItem('daftarInformasiBK')) || [];
    const wadah = document.getElementById('wadahDaftarInfo');
    if (!wadah) return;

    wadah.innerHTML = '';

    if (daftarInfo.length === 0) {
        wadah.innerHTML = `<div class="text-center text-muted py-4">Belum ada catatan atau pengumuman internal yang dipublikasikan.</div>`;
        return;
    }

    daftarInfo.forEach((item, index) => {
        const rowCard = `
            <div class="card p-3 shadow-sm border">
                <div class="d-flex justify-content-between align-items-start">
                    <div>
                        <div class="d-flex align-items-center gap-2 mb-1">
                            <h5 class="fw-bold mb-0">${item.judul}</h5>
                            <small class="text-muted">(${item.tanggal})</small>
                        </div>
                        <p class="text-muted small mb-0" style="white-space: pre-line;">${item.pesan}</p>
                    </div>
                    <div>
                        <button class="btn btn-sm btn-outline-danger px-3" onclick="hapusInformasi(${index})">Hapus</button>
                    </div>
                </div>
            </div>
        `;
        wadah.innerHTML += rowCard;
    });
}

// Hapus Informasi Satuan
function hapusInformasi(index) {
    if (confirm('Yakin ingin menghapus catatan/pengumuman ini?')) {
        let daftarInfo = JSON.parse(localStorage.getItem('daftarInformasiBK')) || [];
        daftarInfo.splice(index, 1);
        localStorage.setItem('daftarInformasiBK', JSON.stringify(daftarInfo));
        loadDaftarInformasi();
    }
}

// Hapus Semua Informasi
function hapusSemuaInformasi() {
    if (confirm('Yakin ingin menghapus seluruh catatan dan pengumuman internal?')) {
        localStorage.removeItem('daftarInformasiBK');
        loadDaftarInformasi();
    }
}
