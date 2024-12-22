# UAS_PEMDAS_SEMESTER1

📚 Repository ini lahir dari panggilan tugas, sebuah perjalanan memahami inti kode. Sebagai ketua kelompok, saya awalnya mengimpikan sistem yang kompleks. Namun, saya disadarkan bahwa esensi sejati tugas ini terletak pada pemahaman dan kesederhanaan. Di sinilah kami menemukan keindahan dalam setiap baris yang kami ciptakan.

## 🎯 Deskripsi Proyek
Proyek ini adalah sebuah aplikasi **Sistem Pembayaran SPP** berbasis terminal yang dikembangkan menggunakan Python dan pustaka OpenPyXL untuk berinteraksi dengan file Excel. Aplikasi ini bertujuan untuk membantu pihak sekolah dalam mencatat data siswa, mengelola tagihan, memproses pembayaran, serta menghasilkan laporan terkait pembayaran SPP secara efisien.

## ✨ Fitur Utama

### 1. 👨‍🎓 Manajemen Siswa
- **Tambah Siswa:** Menambahkan siswa baru dengan informasi seperti nomor siswa, nama siswa, tahun angkatan, kelas, nomor telepon, dan alamat.
- **Tampil Siswa:** Menampilkan daftar siswa yang terdaftar di sekolah.
- **Edit Siswa:** Mengedit informasi siswa, termasuk nama, tahun angkatan, dan detail lainnya.
- **Hapus Siswa:** Menghapus data siswa berdasarkan nomor siswa.

### 2. 💸 Manajemen Tagihan
- **Tambah Tagihan:** Menambahkan tagihan baru dengan informasi kode tagihan, nomor siswa, jumlah tagihan, bulan SPP, status tagihan, dan tanggal.
- **Tampil Tagihan:** Menampilkan seluruh daftar tagihan.
- **Tampil Tagihan Belum Dibayar:** Menampilkan daftar tagihan yang belum lunas.
- **Tampil Tagihan Sudah Dibayar:** Menampilkan daftar tagihan yang sudah lunas.
- **Edit Tagihan:** Mengedit informasi tagihan, seperti nomor siswa, jumlah tagihan, bulan SPP, dan tanggal.
- **Hapus Tagihan:** Menghapus tagihan dari daftar berdasarkan kode tagihan.

### 3. 💳 Pembayaran
- **Tambah Pembayaran:** Mencatat pembayaran SPP siswa, sekaligus mengubah status tagihan menjadi "Sudah Dibayar."
- **Tampil Pembayaran:** Menampilkan riwayat pembayaran SPP siswa.

### 4. 📊 Laporan Servis
- **Tampil Laporan Servis:** Menampilkan laporan transaksi pembayaran SPP untuk seluruh siswa.
- **Tampil Laporan Servis Siswa:** Menampilkan laporan transaksi pembayaran SPP berdasarkan nomor siswa tertentu.

## 🛠️ Teknologi yang Digunakan
- **Python:** Untuk implementasi logika utama aplikasi.
- **OpenPyXL:** Untuk bekerja dengan file Excel, menyimpan, dan mengelola data siswa, tagihan, pembayaran, dan laporan.
- **Excel:** Sebagai basis data untuk menyimpan informasi.

## 🗂️ Struktur File
- `data_spp.xlsx`: File utama yang digunakan untuk menyimpan data aplikasi.
  - **Sheet Siswa:** Menyimpan data siswa yang terdaftar di sekolah.
  - **Sheet Tagihan:** Menyimpan data tagihan SPP.
  - **Sheet Pembayaran:** Menyimpan data pembayaran tagihan SPP.

## 📋 Hirarki Menu Aplikasi

### Menu Utama
1. **Siswa**
   - Tambah Siswa
   - Tampil Siswa
   - Edit Siswa
   - Hapus Siswa
   - Kembali ke Menu Utama
2. **Tagihan**
   - Tambah Tagihan
   - Tampil Tagihan
   - Tampil Tagihan Sudah Dibayar
   - Tampil Tagihan Belum Dibayar
   - Edit Tagihan
   - Hapus Tagihan
   - Kembali ke Menu Utama
3. **Pembayaran**
   - Tambah Pembayaran
   - Tampil Pembayaran
   - Kembali ke Menu Utama
4. **Laporan Servis**
   - Tampil Seluruh Laporan Servis Pembayaran
   - Tampil Laporan Servis Pembayaran per Siswa

## 🚀 Cara Menjalankan Proyek
1. Pastikan Python sudah terinstal di komputer Anda.
2. Instal pustaka **OpenPyXL** dengan perintah berikut:
   ```bash
   pip install openpyxl
   ```
3. Unduh atau clone repository ini:
   ```bash
   git clone https://github.com/Zackisaeful/UAS_PEMDAS_SEMESTER1.git
   ```
4. Jalankan file utama aplikasi:
   ```bash
   python main.py
   ```
5. Ikuti menu interaktif di terminal untuk menggunakan fitur aplikasi.

## Kontributor
- Ketua Kelompok: [Nama Ketua Kelompok]
- Anggota Kelompok:
  1. Dhela Nurastia (10602010)
  2. Salsa Amelia (10602050)
  3. Sulton Sugiarta (10602052)
  4. Zacki Saeful Bahri (10602062)

Kami percaya bahwa proyek ini adalah wujud kolaborasi dan pembelajaran yang tak ternilai. Setiap kontribusi kecil memiliki peran besar dalam menyukseskan tugas ini.

## Lisensi 📜
Proyek ini dilisensikan di bawah [MIT License](LICENSE).

---
✨ "Kesederhanaan adalah kecanggihan tertinggi." - Leonardo da Vinci.


