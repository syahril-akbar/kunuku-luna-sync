# Dokumentasi Sistem Sinkronisasi Data Warehouse & LUNA POS (KUNUKU BABY FOOD)

Sistem `generate_sync.py` ini dirancang khusus untuk mengotomatiskan proses entri data dari format Excel Surat Jalan (dari warehouse/gudang KUNUKU BABY FOOD) menjadi format baku impor untuk diaplikasikan ke dalam sistem LUNA POS.

---

## ⚠️ PROTOKOL KEAMANAN DATA (WAJIB DIBACA)

Ketidakpatuhan terhadap prosedur di bawah ini dapat mengakibatkan **kerusakan integritas data inventaris pada sistem LUNA POS (selisih stok atau duplikasi produk)**.

Sistem otomasi ini sangat bergantung pada berkas rujukan (sumber data). Apabila terdapat pembaruan produk di aplikasi LUNA POS yang tidak diperbarui pada berkas sumber di komputer ini, sistem akan mengidentifikasinya sebagai produk baru.

**STANDAR OPERASIONAL PROSEDUR (SOP) PERSIAPAN:**
1. **Pembaruan Basis Data:** Anda **DIWAJIBKAN** mengakses *dashboard* LUNA POS setiap kali sebelum menggunakan program ini. Lakukan *Export Data Produk* terbaru dan simpan/timpa berkas tersebut di folder ini dengan nama absolut: `Produk.xlsx`.
2. **Penamaan Surat Jalan:** Pastikan berkas Surat Jalan dari gudang berformat Excel (`.xlsx`) dan disimpan di direktori yang sama dengan program.
3. **Kewajiban ID Kolom:** **SANGAT PENTING**. Baris produk di Surat Jalan yang tidak memiliki **ID Barang** di kolom pertamanya akan **SECARA OTOMATIS DIABAIKAN (Dilewati/Dihapus dari Pemrosesan)**. Selalu cek hasil akhir *Total Quantitiy* untuk mencegah gagal mutasi akibat admin lupa mencantumkan ID Barang.

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

Sistem menggunakan penamaan berkas otomatis (Berdasarkan cabang & nama Surat Jalan). Misalnya jika input Anda adalah `SURAT JALAN FIX PERINTIS_30_MARET.xlsx`, maka sistem akan menghasilkan deretan *file* berikut:

### 1. `Laporan_Mutasi_PERINTIS_30_MARET.txt` (Auto-Report)
* **Fungsi:** Laporan rekapitulasi data. Berisi ringkasan matang (Total Pcs Transfer, Total Pcs Baru, Nilai Rupiah Mutasi, dan Baris Error).
* **Tindakan Wajib:** Teks di *file* ini siap disalin (*copy-paste*) ke grup WhatsApp *Daily Stock In* Anda untuk diperiksa oleh manajemen / divisi *Finance*.

### 2. `Hasil_Mapping_Review_PERINTIS_30_MARET.xlsx` (Tahap Validasi Utama)
* **Fungsi:** Papan kontrol validasi visual untuk pencocokan nama Surat Jalan dengan nama baku SKU LUNA.
* **Tindakan Wajib:** Buka berkas ini dan pastikan tidak terdeteksi status **BARU (Tidak ada di LUNA)** pada produk turunan lama. Bila terdeteksi, perhatikan apakah ini emang produk baru rilis dari pabrik atau murni karena ketikan admin warehouse terlalu parah rusaknya.

### 3. `Siap_Product_Baru_PERINTIS_30_MARET.xlsx` (Pendaftaran SKU Baru)
* **Fungsi:** Template pendaftaran Master Product Baru.
* **Kebijakan Harga (Req Finance):** Harga Modal dipatok mati menjadi `0`. Harga Jual mencaplok nilai mentah dari *Surat Jalan* (Harga yang diberikan Gudang = Harga Retail).
* **Input Manual:** Silakan periksa manual label kategori `"General"` kalau Anda memakai nama Kategori spesifik di LUNA POS Anda.
* **Impor LUNA:** Unggah ke menu **Product Import / Master Data** di LUNA.

### 4. `Siap_Warehouse_Transfer_PERINTIS_30_MARET.xlsx` (Mutasi Stok Lama)
* **Fungsi:** Template untuk menambahkan persediaan (Stok In) ke SKU yang sudah ada. Skrip secara cerdas akan mengakumulasi jumlah (*Quantity*) jika gudang mengetikkan jenis produk yang sama dua kali secara terpisah.
* **Tindakan Wajib:** Harap cek total nilai transfer dengan *Laporan_Mutasi* *TXT*.
* **Impor LUNA:** Unggah ke menu **Warehouse Transfer / Mutasi Stok** di LUNA.

---

## 🛑 Limitasi & Proteksi Kesalahan
Sistem dibekali pelindung otomatis tingkat tinggi:
* **Anti-Error Typo Huruf:** Ketikan rusak *"15 Pcs"* di kolom jumlah, atau *"Rp 15.000"* di kolom harga akan disapubersih secara aman menjadi angka utuh (`15` dan `15000`) oleh mesin validasi `safe_int()` untuk mencegah sistem *crash*.
* **Validasi Ekstra Ketat:** Sistem mewajibkan indikator **Umur (11+, 9+, 6+)** dan **Takaran/Volume (100ML, 200ML, dll)** harus cocok mutlak (Exact Match) untuk menyambung stok. Jika luput, barang tersebut diblokade dan dipindahkan ke antrean produk baru demi menghindari tragedi bercampurnya stok *LUNA POS*.
