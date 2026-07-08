# 🛒 KUNUKU - Luna POS Inventory Sync

Sistem otomatisasi sinkronisasi stok antara Surat Jalan Gudang dan LUNA POS. Script ini membantu mapping barang, registrasi produk baru, dan persiapan file transfer gudang secara bulk/bulk processing.

## 🚀 Fitur Utama
- **Auto Mapping**: Mencocokkan SKU/Nama dari Surat Jalan ke database Produk LUNA secara otomatis.
- **Auto Category**: Menambahkan kategori cabang otomatis (PERINTIS, HERTASNING, MALANG) untuk produk baru.
- **Smart Formatting (Single Source of Truth)**: Merapikan nama produk dari Surat Jalan agar seragam dengan standar LUNA POS berdasarkan sub kategori (contoh: menyematkan singkatan `SUP`, `BR`, `FF`, `RB`, `SB`, `BB`, `KALDU`, `LAUK`, `PASTA`, `ABON`, `GHEE`, penyelarasan penulisan Bubur/Nasi Tim, dan kategori Mainan).
- **Smart Sheet Detection**: Otomatis mendeteksi dan memilih sheet terbaik yang berisi data produk utama di dalam file Surat Jalan.
- **Smart Header & Column Mapping**: Otomatis mendeteksi lokasi baris header serta memetakan kolom (ID/SKU, Nama, Qty, Satuan, Harga, Sub Kategori) secara adaptif menggunakan sistem alias cerdas yang meminimalkan kesalahan deteksi kolom (seperti membedakan "Total Qty" dan total nilai rupiah "Total").
- **Ground Truth Reconciliation (Wajib Nol)**: Melakukan audit rekonsiliasi total kuantitas (Qty) dan nilai rupiah (Rp) barang masuk secara real-time untuk memastikan akurasi data 100% (selisih input vs output harus nol).
- **Audit Duplikasi & Konflik**: 
  - *Auto-Merge ID*: Menggabungkan otomatis produk dengan ID/SKU yang sama dari Surat Jalan.
  - *Tabrakan SKU Luna*: Mendeteksi jika terdapat ID SJ berbeda namun terpetakan ke satu SKU LUNA yang sama.
  - *Konflik Nama*: Mendeteksi jika ada nama produk sama tetapi memiliki ID berbeda.
- **🛡️ Deteksi Konflik ID vs Nama**: Memvalidasi kesesuaian ID di Surat Jalan dengan master data Luna. Jika terjadi perbedaan nama yang signifikan, otomatis ditandai sebagai `KONFLIK ID` (berwarna merah di review Excel) untuk mencegah salah stok.
- **Auto-Generate Import Templates**: Menghasilkan file Excel siap pakai sesuai template LUNA POS:
  - `Hasil_Mapping_Review_*.xlsx` (dilengkapi warna status visual: Hijau = OK, Kuning = Baru, Merah = Konflik ID).
  - `Siap_Warehouse_Transfer_*.xlsx` (untuk impor mutasi/transfer stok cabang).
  - `Siap_Product_Baru_*.xlsx` (untuk registrasi produk baru ke LUNA dengan Harga Modal diset `0`).
- **Consistent Output Naming & Archive**: Otomatis membuat folder arsip per cabang berdasarkan tanggal dokumen, memindahkan file Surat Jalan asli, menyalin `Produk.xlsx` sebagai referensi, serta menyelaraskan nama semua file output agar sesuai dengan nama foldernya.
- **🔒 Anti-Crash File Terkunci**: Jika file Excel sedang dibuka di program lain (misalnya Excel) saat program dijalankan, script tidak akan crash melainkan menampilkan peringatan agar file ditutup lalu menekan ENTER untuk mencoba kembali.

## 📦 Prasyarat (Dependencies)
```bash
pip install openpyxl
```

## 🛠️ Cara Penggunaan
1. Pastikan file **`Produk.xlsx`** (Download dari LUNA) ada di folder root.
2. Pastikan file template berikut tersedia:
   - `warehouse-transfer-import-template.xlsx`
   - `product-import-template.xlsx`
3. Letakkan file **Surat Jalan (.xlsx)** di folder root atau masukkan path lengkapnya saat program meminta input.
4. Jalankan script:
   ```bash
   python generate_sync.py
   ```
5. Pilih cabang target (1-3) dan biarkan script bekerja.

## 📝 Aturan Perubahan Nama (Formatting Rules)
Proses sinkronisasi akan merapikan nama dari Surat Jalan agar seragam dengan database LUNA POS. Berikut adalah aturan perubahannya (dimana `{PREFIX}` adalah kode cabang, misal: `PRT`, `HRT`, `MLG`):

1. **Normalisasi Dasar:**
   Spasi berlebih otomatis dihapus dan format huruf diubah menjadi **KAPITAL** (Uppercase). Terdapat juga auto-koreksi typo standar (contoh: "GENDRANG" ➡️ "GENDERANG").

2. **Kategori Mainan:**
   - Dideteksi jika produk mengandung kata kunci: `LONCENG`, `GENDERANG`, `KENYOT`, `GELANG`, `TEPUK TANGAN`, `TEETER`, `DOT BABY` (dengan pengecualian jika ada kata `SENDOK`).
   - **Pola Konversi:** `{PREFIX} MAINAN {NAMA ASLI}`
   - *Contoh:* `Lonceng Bayi` ➡️ `HRT MAINAN LONCENG BAYI`

3. **Kategori Bubur & Nasi Tim:**
   - Dideteksi jika nama diawali variasi awalan bubur (`N TIM`, `N, TIM`, `N. TIM`, `B.`, atau `B`) dan memiliki informasi umur serta volume di akhir (contoh: `11+ 200 ML`).
   - **Pola Konversi:** `{PREFIX} BUBUR {UMUR} {VOLUME} {VARIAN RASA}`
   - *Contoh:* `N, Tim Cakalang Woku 11+ 200 ml` ➡️ `HRT BUBUR 11+ 200 ML CAKALANG WOKU`

4. **Konvensi Sub Kategori Khusus (Single Source of Truth):**
   Script mencocokkan sub kategori dari surat jalan dan menerapkan aturan singkatan/prefix berikut:
   
   | Sub Kategori | Mode Pemrosesan | Singkatan | Contoh Hasil Konversi |
   | :--- | :--- | :--- | :--- |
   | **SOUP** | `keyword` | `SUP` | `Empal Hati Sapi` ➡️ `{PREFIX} SUP EMPAL HATI SAPI`<br>`Sop Ayam Bakso` ➡️ `{PREFIX} SUP AYAM BAKSO` (Mengganti SOP/SOUP di depan menjadi SUP)<br>`Beef Mandu Soup` ➡️ `{PREFIX} SUP BEEF MANDU SOUP` (Selalu diawali SUP) |
   | **BUTTER RICE** | `strip_prefix` | `BR` | `Butter Rice Dory Cauli Flower` ➡️ `{PREFIX} BR DORY CAULI FLOWER`<br>`Butter Rice Beef` ➡️ `{PREFIX} BR BUTTER RICE BEEF` (Jika tersisa hanya 1 kata setelah strip, kata 'BUTTER RICE' tetap dipertahankan) |
   | **FINGER FOOD** | `strip_prefix` | `FF` | `Finger Food Nugget Dory` ➡️ `{PREFIX} FF NUGGET DORY` |
   | **RICE BOX** | `strip_prefix` | `RB` | `Rice Box Chicken Katsu` ➡️ `{PREFIX} RB CHICKEN KATSU` |
   | **SNACK BUAH** | `strip_prefix` | `SB` | `Snack Buah Oats Apple` ➡️ `{PREFIX} SB OATS APPLE` (Mencegah redundansi prefix) |
   | **PELENGKAP BB BOOSTER** | `prefix` | `BB` | `Kremes Hati Ayam` ➡️ `{PREFIX} BB KREMES HATI AYAM` |
   | **KALDU** | `keyword` | `KALDU` | `Salmon Immune Booster` ➡️ `{PREFIX} KALDU SALMON IMMUNE BOOSTER` |
   | **LAUK** | `keyword` | `LAUK` | `Gadon Sapi` ➡️ `{PREFIX} LAUK GADON SAPI` |
   | **PASTA** | `keyword` | `PASTA` | `Cheese Roll` ➡️ `{PREFIX} PASTA CHEESE ROLL` |
   | **ABON** | `keyword` | `ABON` | `Abon Sapi` ➡️ `{PREFIX} ABON SAPI` |
   | **GHEE & SAUS** | `keyword` | `GHEE` | `Ghee Butter Original` ➡️ `{PREFIX} GHEE BUTTER ORIGINAL` |
   | **PUREE, SEMI SOLID, SOLID** | `pass` | - | Mengikuti aturan konversi Bubur & Nasi Tim di atas. |
   | **CUP, STIKER, & PLASTIK** | `pass` | - | Non-makanan, tidak diberi singkatan sub kategori. |
   | **MAINAN/AKSESORIS** | `pass` | - | Dihandle terpisah oleh deteksi keywords mainan (contoh: Lonceng ➡️ `{PREFIX} MAINAN LONCENG`). |

    > **Note on Keyword Mode**: Untuk sub kategori dengan mode `keyword` (seperti KALDU, LAUK, PASTA, ABON, GHEE), singkatan **hanya akan ditambahkan jika nama asli produk belum mengandung kata kunci tersebut**. Khusus untuk **SOUP**, kata kunci `SOP`/`SOUP` di awal otomatis diubah menjadi `SUP` dan awalan `SUP` akan selalu disematkan di depan nama produk demi konsistensi pengelompokan.

## 📂 Struktur Output
Setelah dijalankan, script akan membuat folder baru (contoh: `PERINTIS_SENIN_2_APRIL_2026`) yang berisi:
1. **SURAT JALAN ... .xlsx** (Arsip file sumber)
2. **Hasil_Mapping_Review_... .xlsx** (WAJIB CEK: Menampilkan baris status warna-warni untuk produk cocok (Hijau), produk baru (Kuning), dan konflik ID (Merah))
3. **Siap_Warehouse_Transfer_... .xlsx** (Pake ini buat import mutasi cabang di LUNA)
4. **Siap_Product_Baru_... .xlsx** (Pake ini buat registrasi barang yang belum terdaftar di LUNA)
5. **Laporan_Mutasi_... .txt** (Ringkasan rupiah & jumlah stok buat laporan ke pimpinan)

---
*Note: Selalu cek file Review Mapping sebelum melakukan import produk baru untuk memastikan tidak ada SKU ganda.*
