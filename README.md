# 🛒 KUNUKU - Luna POS Inventory Sync

Sistem otomatisasi sinkronisasi stok antara Surat Jalan Gudang dan LUNA POS. Script ini membantu mapping barang, registrasi produk baru, dan persiapan file transfer gudang secara bulk/bulk processing.

## 🚀 Fitur Utama
- **Auto Mapping**: Mencocokkan SKU/Nama dari Surat Jalan ke database Produk LUNA.
- **Auto Category**: Menambahkan kategori otomatis (PERINTIS, HERTASNING, MALANG) untuk produk baru.
- **Smart Formatting**: Merapikan nama produk sesuai standar LUNA (misal: "PRT BUBUR 6+ 300 ML ...").
- **Folder Archive**: Otomatis menyimpan file hasil generate & file sumber ke folder cabang berbasis tanggal.
- **Summary Report**: Laporan ringkasan total Qty & Nilai Rupiah barang masuk.

## 📦 Prasyarat (Dependencies)
```bash
pip install openpyxl
```

## 🛠️ Cara Penggunaan
1. Pastikan file **`Produk.xlsx`** (Download dari LUNA) ada di folder root.
2. Pastikan file template berikut tersedia:
   - `warehouse-transfer-import-template.xlsx`
   - `product-import-template.xlsx`
3. Letakkan file **Surat Jalan (.xlsx)** di folder root.
4. Jalankan script:
   ```bash
   python generate_sync.py
   ```
5. Pilih cabang target (1-3) dan biarkan script bekerja.

## 📂 Struktur Output
Setelah dijalankan, script akan membuat folder baru (contoh: `PERINTIS_SENIN_2_APRIL_2026`) yang berisi:
1. **SURAT JALAN ... .xlsx** (Arsip file sumber)
2. **Hasil_Mapping_Review_... .xlsx** (WAJIB CEK: Untuk melihat ringkasan item mana yang cocok dan mana yang jadi barang baru)
3. **Siap_Warehouse_Transfer_... .xlsx** (Pake ini buat import mutasi cabang di LUNA)
4. **Siap_Product_Baru_... .xlsx** (Pake ini buat registrasi barang yang belum terdaftar di LUNA)
5. **Laporan_Mutasi_... .txt** (Ringkasan rupiah & jumlah stok buat laporan ke pimpinan)

---
*Note: Selalu cek file Review Mapping sebelum melakukan import produk baru untuk memastikan tidak ada SKU ganda.*
