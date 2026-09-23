document.addEventListener('DOMContentLoaded', () => {
    // 1. Logika Login
    const formLogin = document.getElementById('formLogin');
    if (formLogin) {
        formLogin.addEventListener('submit', (e) => {
            e.preventDefault();
            const role = document.getElementById('role').value;
            const kode = document.getElementById('kodeAkses').value;

            if (role === 'walas' && kode === 'walas123') {
                localStorage.setItem('sessionRole', 'walas');
                window.location.href = 'walas.html';
            } else if (role === 'bk' && kode === 'bk123') {
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
            const infoEl = document.getElementById('infoTanggalAbsen');
            const labelEl = document.getElementById('labelTanggalHadir');
            if (infoEl) infoEl.innerText = `Mengelola absensi untuk tanggal: ${today}`;
            if (labelEl) labelEl.innerText = `Hadir Pada Tanggal ${today}`;
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

    // Cek apakah data pada tanggal ini sudah disimpan
    const sudahDisimpan = localStorage.getItem('saved_absensi_' + tanggalAktif) === 'true';
    let absensiHarian = JSON.parse(localStorage.getItem('absensi_' + tanggalAktif)) || {};

    const statTotal = document.getElementById('statTotal');
    if (statTotal) statTotal.innerText = daftarSiswa.length;
    
    let totalHadir = 0;
    let totalBk = 0;

    const tbody = document.querySelector('#tabelSiswa tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (daftarSiswa.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">Data Tidak Tersedia</td></tr>`;
        const statHadir = document.getElementById('statHadir');
        const statBk = document.getElementById('statBk');
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
        if (statusAbsen === 'Hadir' || statusAbsen === 'Terlambat') {
            totalHadir++;
        }
        if (siswa.statusBk === 'Dalam Penanganan') {
            totalBk++;
        }

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

    const statHadir = document.getElementById('statHadir');
    const statBk = document.getElementById('statBk');
    if (statHadir) statHadir.innerText = totalHadir;
    if (statBk) statBk.innerText = totalBk;
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
                const statusAbsen = absensiHarian[siswa.nis] || '-';
                tbody.innerHTML += `<tr>
                    <td>${index + 1}</td>
                    <td>${siswa.nis}</td>
                    <td>${siswa.nama}</td>
                    <td><span class="badge bg-secondary">${siswa.gender}</span></td>
                    <td><span class="badge bg-primary">${statusAbsen}</span></td>
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
    const tgl = filterEl.value;
    
    const infoEl = document.getElementById('infoTanggalAbsen');
    const labelEl = document.getElementById('labelTanggalHadir');
    if (infoEl) infoEl.innerText = `Mengelola absensi untuk tanggal: ${tgl}`;
    if (labelEl) labelEl.innerText = `Hadir Pada Tanggal ${tgl}`;
    
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