# KeuanganKu — Google Apps Script

Aplikasi pencatatan keuangan pribadi berbasis Google Apps Script dan Google Sheets.

## Isi paket
- `Code.gs`: backend, pembuatan sheet, validasi, transaksi, anggaran, laporan, CSV.
- `Index.html`: antarmuka responsif HP/laptop.

## Instalasi
1. Buat Google Spreadsheet baru bernama `KeuanganKu Database`.
2. Di spreadsheet, pilih **Extensions / Ekstensi → Apps Script**.
3. Hapus kode contoh di `Code.gs`, lalu salin isi file `Code.gs` dari paket ini.
4. Tekan tombol **+** di sebelah Files → **HTML**, beri nama persis `Index`, lalu salin isi `Index.html`.
5. Di Apps Script, pilih fungsi `setupKeuanganKu` dari dropdown fungsi, lalu tekan **Run / Jalankan**.
6. Ikuti dialog izin Google. Fungsi ini membuat sheet `Transaksi`, `Anggaran`, `Pengaturan`, dan `Kategori`.
7. Kembali ke spreadsheet dan periksa apakah keempat sheet sudah dibuat.
8. (Opsional) Buka **Project Settings → Script properties** dan tambahkan `ALLOWED_EMAIL` dengan alamat akun Google milikmu. Ini adalah lapisan tambahan; deployment tetap harus dibatasi ke diri sendiri.
9. Klik **Deploy → New deployment**.
10. Klik ikon roda gigi / Select type → **Web app**.
11. Set **Execute as: Me** dan **Who has access: Only myself**.
12. Klik **Deploy**, selesaikan otorisasi, lalu buka URL Web App yang diberikan.

## Pembaruan kode
Setelah mengedit kode, buka **Deploy → Manage deployments → Edit** (ikon pensil), pilih **New version**, lalu deploy lagi. URL Web App biasanya tetap sama.

## Keamanan
- Jangan pilih “Anyone” atau “Anyone with Google account”.
- Jangan bagikan spreadsheet atau URL web app ke orang lain jika ingin tetap pribadi.
- Jangan simpan password, PIN, nomor kartu lengkap, atau data autentikasi di catatan transaksi.
- Jika memakai `ALLOWED_EMAIL`, masukkan email Google yang benar. Jika identitas aktif tidak tersedia, kontrol akses utama tetap harus melalui pilihan “Only myself” pada deployment.
- Batasi siapa yang memiliki akses editor ke spreadsheet dan proyek Apps Script.
- Data tersimpan di Google Sheets. Buat salinan cadangan berkala.

## Logika saldo
`Saldo saat ini = Saldo awal + seluruh pemasukan tercatat - seluruh pengeluaran tercatat`.
Saldo awal adalah saldo pada tanggal sebelum mulai memakai pencatatan, sehingga transaksi lama jangan dimasukkan lagi jika sudah tercakup di saldo awal.
Anggaran membandingkan batas kategori dengan total pengeluaran kategori pada bulan yang dipilih.
Transaksi “Tabungan / Dana Darurat” saat ini dicatat sebagai pengeluaran/alokasi, sehingga saldo kas berkurang.

## Catatan
- Tanggal dan waktu mengikuti zona waktu proyek Apps Script. Buka **Project Settings** dan pilih zona waktu `Asia/Jakarta`.
- Aplikasi ini tidak terhubung otomatis ke rekening bank. Semua transaksi dimasukkan manual.
- Sheet `Kategori` dapat diedit langsung untuk menambah kategori. Kolom `Aktif` harus bernilai TRUE/FALSE.
- CSV hanya mengekspor transaksi bulan terpilih.

# catatpengeluaran