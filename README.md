# 🛒 KUNUKU - Luna POS Inventory Sync

Sistem otomatisasi sinkronisasi stok antara Surat Jalan Gudang dan LUNA POS. Script ini membantu mapping barang, registrasi produk baru, dan persiapan file transfer gudang secara bulk/bulk processing.

## 🚀 Fitur Utama
- **Auto Mapping**: Mencocokkan SKU/Nama dari Surat Jalan ke database Produk LUNA.
- **Auto Category**: Menambahkan kategori otomatis (PERINTIS, HERTASNING, MALANG) untuk produk baru.
- **Smart Formatting**: Merapikan nama produk sesuai standar LUNA (misal: "PRT BUBUR 6+ 300 ML ...").
- **Folder Archive**: Otomatis menyimpan file hasil generate & file sumber ke folder cabang berbasis tanggal.
- **Summary Report**: Laporan ringkasan total Qty & Nilai Rupiah barang masuk.
- **🛡️ Deteksi Konflik ID vs Nama**: Memvalidasi kesesuaian ID di Surat Jalan dengan master data Luna. Jika terjadi perbedaan nama yang signifikan, otomatis ditandai sebagai `KONFLIK ID` (berwarna merah di review Excel) untuk mencegah salah stok.
- **🔒 Anti-Crash File Terkunci**: Jika file Excel sedang dibuka di program lain saat dijalankan, script tidak akan crash, melainkan menampilkan pesan instruksi dan opsi untuk menekan ENTER untuk retry setelah file ditutup.

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
   | **BUTTER RICE** | `prefix` | `BR` | `Chicken Rice Mentai` ➡️ `{PREFIX} BR CHICKEN RICE MENTAI` |
   | **FINGER FOOD** | `strip_prefix` | `FF` | `Finger Food Nugget Dory` ➡️ `{PREFIX} FF NUGGET DORY` |
   | **RICE BOX** | `strip_prefix` | `RB` | `Rice Box Chicken Katsu` ➡️ `{PREFIX} RB CHICKEN KATSU` |
   | **SNACK BUAH** | `prefix` | `SB` | `Pisang Regal Vla` ➡️ `{PREFIX} SB PISANG REGAL VLA` |
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
