# Dokumentasi Sistem Sinkronisasi Data Warehouse & LUNA POS (KUNUKU BABY FOOD)

Sistem `generate_sync.py` ini dirancang khusus untuk mengotomatiskan proses entri data dari format Excel Surat Jalan (dari warehouse/gudang KUNUKU BABY FOOD) menjadi format baku impor untuk diaplikasikan ke dalam sistem LUNA POS.

---

## ⚠️ PROTOKOL KEAMANAN DATA (WAJIB DIBACA)

Ketidakpatuhan terhadap prosedur di bawah ini dapat mengakibatkan **kerusakan integritas data inventaris pada sistem LUNA POS (selisih stok atau duplikasi produk)**.

Sistem ini **TIDAK MENGGUNAKAN TEBAK-TEBAKAN (NO GUESSING)**. Seluruh pencocokan dilakukan secara **SAMA PERSIS (Exact Pattern Match)** berdasarkan database LUNA. Apabila terdapat selisih satu spasi, typo, atau produk baru di Surat Jalan, mesin akan mendaftarkannya sebagai **PRODUK BARU**.

**STANDAR OPERASIONAL PROSEDUR (SOP) PERSIAPAN:**
1. **Pembaruan Basis Data:** Anda **DIWAJIBKAN** mengakses *dashboard* LUNA POS setiap kali sebelum menggunakan program ini. Lakukan *Export Data Produk* terbaru dan simpan/timpa berkas tersebut di folder ini dengan nama absolut: `Produk.xlsx`. Database ini adalah Kamus Pola Utama.
2. **Penamaan Surat Jalan:** Pastikan berkas Surat Jalan dari gudang berformat Excel (`.xlsx`) dan disimpan di direktori yang sama dengan program. Note: Mesin otomatis akan mendeteksi *Sheet/Table* mana yang berisi data terpanjang.
3. **Kewajiban ID Kolom:** **SANGAT PENTING**. Baris produk di Surat Jalan yang tidak memiliki **ID Barang** di kolom pertamanya akan **SECARA OTOMATIS DIABAIKAN (Dilewati)**. Selalu cek hasil akhir *Total Quantitiy* untuk mencegah gagal mutasi.

---

## 🛠 Panduan Eksekusi Program

1. Pastikan Anda telah berada di direktori aplikasi.
2. Buka Terminal / Command Prompt / PowerShell.
3. Jalankan perintah eksekusi berikut:
   ```cmd
   python generate_sync.py
   ```
4. **Input Nama File:** Ketikkan nama berkas secara presisi (contoh: `SURAT JALAN BARU.xlsx`).
5. **Input Kode Cabang:** Ketikkan kode bersangkutan (contoh: `PRT`, `HRT`, `MLG`).

---

## 📂 Penjelasan Berkas Output (Hasil Pemrosesan)

Sistem akan otomatis memberi nama berkas hasil (Misal *SURAT JALAN MLG_APRIL.xlsx* menjadi urutan file berikut):

### 1. `Laporan_Mutasi_MLG_APRIL.txt` (Auto-Report)
* **Fungsi:** Laporan rekapitulasi data (Total Pcs Mutasi, Pcs Barang Baru, Nilai Aset Rupiah, dan Log Error ID Kosong).
* **Tindakan Wajib:** Salin/copy isi teks di file ini ke tim *Finance* atau grup WhatsApp *Daily Stock In*.

### 2. `Hasil_Mapping_Review_MLG_APRIL.xlsx` (Tahap Validasi Utama)
* **Fungsi:** Papan kontrol validasi visual.
* **Tindakan Wajib:** Cek status barang. Jika status tercatat **BARU (Tidak ada di LUNA)**, berarti barang tersebut tulisan Surat Jalan-nya tidak sama persis dengan ejaan di LUNA POS, ATAU memang produk yang betul-betul baru diturunkan dari pabrik.

### 3. `Siap_Product_Baru_MLG_APRIL.xlsx` (Pendaftaran SKU Baru)
* **Fungsi:** Template pendaftaran Master Product Baru LUNA POS.
* **Suntikan Kuantitas Otomatis:** Meskipun ini adalah produk "Barang Baru", **Kuantitiy / Stok-nya tidak akan hangus**. Skrip otomatis memasukkan jumlah stok dari Surat Jalan ke kolom **In Stock (Stok Awal)**. Admin LUNA tidak perlu membuat transaksi mutasi tambahan.
* **Kebijakan Harga (Req Finance):** Harga Modal dikosongkan/dipatok nol (`0`). Harga Jual mutlak mengambil nilai mentah dari *Surat Jalan* (Harga Gudang).
* **Impor LUNA:** Unggah file ini ke menu **Product Import / Master Data** di LUNA.

### 4. `Siap_Warehouse_Transfer_MLG_APRIL.xlsx` (Mutasi Stok Lama)
* **Fungsi:** Template Transfer Stok untuk SKU Lama yang ada di `Produk.xlsx`. Skrip akan mengakumulasi angka *Quantity* jika gudang memecah pengiriman jadi beberapa baris untuk satu produk yang sama.
* **Impor LUNA:** Unggah file ini ke menu **Warehouse Transfer / Mutasi Stok** di LUNA.

---

## 🛑 Limitasi & Proteksi Kesalahan
Sistem menggunakan pertahanan presisi level militer:
1. **Akurasi Spasi Mutlak (No AI Prediksi):** Fitur tebak-tebakan kata dimatikan. Nama Surat Jalan dikonversi pake rumus baku baku (Contoh: `MLG BUBUR 6+ 200 ML NAMA VARIAN`). Jika ejaan di `Produk.xlsx` tidak **sama persis** (Meleset satu kata/spasi), produk itu dibuang ke Antrean Barang Baru demi mencegah kontaminasi stok lama.
2. **Anti-Error Typo Rupiah/Pcs:** Jika gudang ngetik cacat seperti `"15 Pcs"` atau `"Rp.15000"`, fungsi `safe_int()` akan menyapubersih teks jadi angka matematika polos (`15` dan `15000`) agar impor *LUNA POS* tidak menolak/crash datanya.
3. **Penyekat Antar Cabang:** Cabang PRT tidak akan sudi dan tidak bisa mengambil produk dengan label MLG. Validasi gerbang depan telah menutup kemungkinan pertukaran mutasi antar ID cabang.
