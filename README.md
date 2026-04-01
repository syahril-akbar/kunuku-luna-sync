# Dokumentasi Sistem Sinkronisasi Data Warehouse & LUNA POS (KUNUKU BABY FOOD)

Sistem `generate_sync.py` ini dirancang khusus untuk mengotomatiskan proses entri data dari format Excel Surat Jalan (dari warehouse/gudang KUNUKU BABY FOOD) menjadi format baku impor untuk diaplikasikan ke dalam sistem LUNA POS.

---

## ⚠️ PROTOKOL KEAMANAN DATA (WAJIB DIBACA)

Ketidakpatuhan terhadap prosedur di bawah ini dapat mengakibatkan **kerusakan integritas data inventaris pada sistem LUNA POS (selisih stok atau duplikasi produk)**.

Sistem otomasi ini sangat bergantung pada berkas rujukan (sumber data). Apabila terdapat pembaruan produk di aplikasi LUNA POS yang tidak diperbarui pada berkas sumber di komputer ini, sistem akan mengidentifikasinya sebagai produk baru.

**STANDAR OPERASIONAL PROSEDUR (SOP) PERSIAPAN:**
1. **Pembaruan Basis Data:** Anda **DIWAJIBKAN** mengakses *dashboard* LUNA POS setiap kali sebelum menggunakan program ini. Lakukan *Export Data Produk* terbaru dan simpan/timpa berkas tersebut di folder ini dengan nama absolut: `Produk.xlsx`.
2. **Penamaan Surat Jalan:** Pastikan berkas Surat Jalan dari gudang berformat Excel (`.xlsx`) dan disimpan di direktori yang sama dengan program.

---

## 🛠 Panduan Eksekusi Program

Langkah-langkah untuk menjalankan proses sinkronisasi:

1. Pastikan Anda telah berada di direktori aplikasi.
2. Buka antarmuka baris perintah (Terminal / Command Prompt / PowerShell).
3. Jalankan perintah eksekusi berikut:
   ```cmd
   python generate_sync.py
   ```
4. **Input Nama File:** Program akan meminta nama berkas Surat Jalan. Ketikkan nama berkas secara akurat (contoh: `SURAT JALAN BARU.xlsx`) lalu tekan Enter.
5. **Input Kode Cabang:** Program akan meminta awalan/kode *Warehouse*. Ketikkan kode bersangkutan (contoh: `PRT` untuk Perintis, `HRT` untuk Hertasning, dsb) lalu tekan Enter.

---

## 📂 Penjelasan Berkas Output (Hasil Pemrosesan)

Setelah eksekusi berhasil, sistem akan menghasilkan **3 Berkas Baru**. Jangan langsung melakukan impor ke dalam LUNA POS tanpa memahami fungsi dan melakukan validasi:

### 1. `Hasil_Mapping_Review.xlsx` (Tahap Validasi Utama)
* **Fungsi:** Papan kontrol validasi visual. Lembar kerja ini menunjukkan hasil analisis kecocokan (pencocokan nama Surat Jalan dengan nama baku SKU LUNA).
* **Tindakan Wajib:** Buka berkas ini dan periksa kolom "Prediksi Nama LUNA". Pastikan kecocokan tidak keliru. Apabila statusnya tercatat sebagai **BARU (Tidak ada di LUNA)**, berarti SKU tersebut tidak ditemukan persamaannya dan akan dicetak sebagai produk baru.

### 2. `Siap_Product_Baru.xlsx` (Pendaftaran SKU Baru)
* **Fungsi:** Template data untuk mendaftarkan Master Product baru (barang dari Surat Jalan yang dipastikan belum pernah di-input ke sistem LUNA rujukan).
* **Input Manual Wajib:**
  1. **Kategori:** Secara *default*, sistem mengisinya dengan label `"General"`. Ubah manual nama kategori ini sesuai standar manajemen kategori di LUNA POS Anda.
  2. **Harga Jual:** Sistem akan secara otomatis menetapkan mark-up bawaan sebesar **1.5x (naik 50%)** dari Harga Modal. **ANDA WAJIB** mengubah kolom ini secara manual sesuai dengan standar harga ritel yang berlaku.
  3. **Batas Minimum Stok:** Secara bawaan diisi "1".
* **Impor LUNA:** Unggah berkas yang sudah direvisi ini ke menu **Product Import / Master Data** di LUNA.

### 3. `Siap_Warehouse_Transfer.xlsx` (Registrasi Pemindahan/Penambahan Stok)
* **Fungsi:** Template untuk melakukan *Mutasi Stok* atau mutasi barang yang sudah terdaftar lama di dalam sistem. Sistem secara cerdas akan menjumlahkan (akumulasi) stok apabila gudang mengetikkan suatu produk lebih dari satu baris secara terpisah.
* **Tindakan Wajib:** Verifikasi total angka mutasi sebelum mengimpor.
* **Impor LUNA:** Unggah berkas ini secara utuh (tidak perlu diedit) ke menu **Warehouse Transfer / Mutasi Stok** di LUNA.

---

## 🛑 Limitasi & Mitigasi Kesalahan
Sistem dibekali algoritma validasi ketat. Sistem mewajibkan indikator **Umur (11+, 9+, 6+)** dan **Takaran/Volume (100ML, 200ML, dll)** sama persis (`Exact Match Attribute`) meskipun rentetan teks bervariasi. 

**Catatan Khusus:**
Apabila admin gudang mengetikkan nama produk baru dengan teledor (lupa melampirkan keterangan umur atau satuan ukuran ML), program otomasi ini akan **memisahkan produk tersebut ke dalam berkas `Siap_Product_Baru.xlsx`** (sebagai produk baru yang tidak dikenali) untuk mencegah risiko pencampuran stok inventori yang fatal pada ukuran varian yang dilarang. 

*Disiplin pencatatan manual senantiasa menjadi kunci kelancaran otomasi sistem ujung-ke-ujung.*
