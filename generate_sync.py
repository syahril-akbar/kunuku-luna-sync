import openpyxl
from openpyxl.utils import get_column_letter
import difflib
import re
import os
import sys
import shutil
import warnings

# Bungkam peringatan bawel dari openpyxl (DrawingML warning)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTLET_CATEGORIES = {
    "PRT": "PERINTIS",
    "HRT": "HERTASNING",
    "MLG": "MALANG"
}

def auto_resize_columns(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter # Dapatkan huruf kolom
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = min(adjusted_width, 60)

def safe_int(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    digits = re.sub(r'[^\d]', '', str(value))
    return int(digits) if digits else 0

print("Memulai proses sinkronisasi dan pencocokan data...")

# 1. Validasi File Surat Jalan
while True:
    sj_filename = input("\nMasukkan nama/path file Surat Jalan (.xlsx): ").strip(" '\"")
    if not sj_filename.endswith('.xlsx'):
        sj_filename += '.xlsx'
    
    if os.path.exists(sj_filename):
        break
    else:
        print(f"⚠️ ERROR: File '{sj_filename}' tidak ditemukan!")
        print("File yang tersedia di folder ini:")
        xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        for f in xlsx_files:
            print(f"  - {f}")

# 2. Pemilihan Cabang (Strict Selection)
print("\nPilih Cabang Target:")
options = list(OUTLET_CATEGORIES.keys())
for idx, code in enumerate(options, 1):
    print(f"{idx}. {code} ({OUTLET_CATEGORIES[code]})")

while True:
    choice = input(f"Pilih nomor (1-{len(options)}): ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(options):
        warehouse_prefix = options[int(choice) - 1]
        break
    print("⚠️ Input tidak valid! Pilih nomor yang tersedia.")

print(f"✅ Menggunakan awalan: {warehouse_prefix} ({OUTLET_CATEGORIES[warehouse_prefix]})")

base_name = os.path.splitext(os.path.basename(sj_filename))[0]
suffix_text = re.sub(r'(?i)SURAT JALAN', '', base_name)
suffix_text = re.sub(r'(?i)FIX', '', suffix_text)
file_suffix = re.sub(r'[^A-Za-z0-9]+', '_', suffix_text).strip('_')
if not file_suffix:
    file_suffix = "OUTPUT"

# 3. Setup Folder Output (Berdasarkan Nama Cabang + Tanggal dr Nama File)
branch_name = OUTLET_CATEGORIES[warehouse_prefix]

# BERSIHKAN file_suffix dari redundancy nama cabang
# Hapus prefix (misal 'PRT') dan nama cabang (misal 'PERINTIS') dari suffix_text
clean_suffix = file_suffix
to_remove = [warehouse_prefix, branch_name, "OUTLET", "OUTET"] # Tambahkan kata umum yang sering muncul
for word in to_remove:
    # Gunakan RegEx untuk menghapus kata secara case-insensitive
    clean_suffix = re.sub(rf'(?i){word}', '', clean_suffix).strip('_')

# Pastikan tidak ada double underscore atau underscore menggantung
clean_suffix = re.sub(r'_{2,}', '_', clean_suffix).strip('_')

if not clean_suffix:
    clean_suffix = "OUTPUT"

folder_name = f"{branch_name}_{clean_suffix}"
output_dir = os.path.join(os.getcwd(), folder_name)

if not os.path.exists(output_dir):
    os.makedirs(output_dir)
    print(f"📂 Membuat folder baru: {folder_name}")
else:
    print(f"📂 Menggunakan folder yang ada: {folder_name}")

wb_produk = openpyxl.load_workbook('Produk.xlsx', data_only=True)
ws_produk = wb_produk['Sheet1']

luna_data = {}
all_luna_skus = set()
for r in range(2, ws_produk.max_row + 1):
    sku = ws_produk.cell(r, 1).value
    nama = str(ws_produk.cell(r, 2).value or "").strip()
    qty = ws_produk.cell(r, 4).value
    harga_jual = ws_produk.cell(r, 6).value
    harga_modal = ws_produk.cell(r, 8).value
    if sku and nama and str(sku).strip().lower() != 'perintis' and nama.lower() != 'none':
        sku_str = str(sku).strip() # STRIP SKU
        all_luna_skus.add(sku_str)
        
        nama_str = nama.strip() # STRIP NAMA
        # FILTER: Blokir SKU cabang lain (misal MLG) menyusup ke data PRT
        if not nama_str.upper().startswith(warehouse_prefix):
            continue
            
        luna_data[sku_str] = {
            'nama': nama_str,
            'qty': qty,
            'harga_jual': harga_jual,
            'harga_modal': harga_modal
        }

try:
    wb_sj = openpyxl.load_workbook(sj_filename, data_only=True)
    
    # Auto-detect target sheet (Pilih sheet dengan baris data terbanyak)
    ws_sj = wb_sj.active
    biggest_row = 0
    for sheet in wb_sj.sheetnames:
        if wb_sj[sheet].max_row > biggest_row:
            biggest_row = wb_sj[sheet].max_row
            ws_sj = wb_sj[sheet]
            
    print(f"--> Target Data Terdeteksi di Sheet: '{ws_sj.title}' ({biggest_row} Baris)")
except Exception as e:
    print(f"ERROR: Gagal membaca file excel Surat Jalan. Exception: {e}")
    sys.exit(1)

sj_data = {}
baris_gagal = 0
for r in range(2, ws_sj.max_row + 1):
    id_p = ws_sj.cell(r, 1).value
    nama = str(ws_sj.cell(r, 2).value or "").strip()
    qty_raw = ws_sj.cell(r, 5).value
    qty = safe_int(qty_raw)
    sat = str(ws_sj.cell(r, 6).value or "").strip()
    harga_raw = ws_sj.cell(r, 9).value
    harga = safe_int(harga_raw)
    
    if id_p and nama and nama.lower() != 'none':
        id_str = str(id_p).strip()
        if id_str in sj_data:
            sj_data[id_str]['qty'] += qty
        else:
            sj_data[id_str] = {
                'nama': nama,
                'qty': qty,
                'satuan': sat,
                'harga': harga
            }
    else:
        if nama and nama.lower() != 'none':
            baris_gagal += 1

def normalize_spaces(text):
    return ' '.join(str(text).split()).upper()

def format_nama_luna(name, prefix):
    name = normalize_spaces(name)
    match_bubur = re.search(r'^(N\.?\s*TIM\.?|B\.+|B)\s+(.*?)\s+(\d+\+)\s+(\d+)\s*ML$', name)
    if match_bubur:
        varian = match_bubur.group(2).strip()
        umur = match_bubur.group(3).strip()
        volume = match_bubur.group(4).strip() + " ML"
        return f'{prefix} BUBUR {umur} {volume} {varian}'
    if name.startswith('KALDU'):
        return f'{prefix} {name}'
    if not name.startswith(prefix):
        return f'{prefix} {name}'
    return name

matched_items = []
unmatched_items = []
mapping_review = []

for id_sj, item_sj in sj_data.items():
    expected_luna_name = format_nama_luna(item_sj['nama'], warehouse_prefix)
    
    best_match_sku = None
    best_match_nama_luna = None
    
    # 0. Prioritaskan Pencocokan ID (SKU) Mutlak terlebih dahulu
    # Ini menjamin jika Gudang memasukkan SKU yg sudah terdaftar, ia tidak akan dijadikan Barang Baru
    if id_sj in luna_data:
        best_match_sku = id_sj
        best_match_nama_luna = luna_data[id_sj]['nama']
    
    # 1. Exact Match Strict Pattern System (Jika ID tidak cocok/kosong, coba cocokkan lewat pola Nama)
    if not best_match_sku:
        expected_norm = normalize_spaces(expected_luna_name)
        for sku, data in luna_data.items():
            if normalize_spaces(data['nama']) == expected_norm:
                best_match_sku = sku
                best_match_nama_luna = data['nama']
                break
    
    mapping_review.append({
        'id_sj': id_sj,
        'nama_sj': item_sj['nama'],
        'qty_sj': item_sj['qty'],
        'sku_luna_prediksi': best_match_sku,
        'nama_luna_prediksi': best_match_nama_luna if best_match_nama_luna else f"[NEW] {format_nama_luna(item_sj['nama'], warehouse_prefix)}",
        'harga_modal_sj': item_sj['harga']
    })
    
    if best_match_sku:
        matched_items.append({
            'sku': best_match_sku,
            'nama': best_match_nama_luna,
            'satuan': item_sj['satuan'],
            'qty': item_sj['qty']
        })
    else:
        unmatched_items.append({
            'id': id_sj,
            'nama': item_sj['nama'],
            'harga_modal': item_sj['harga'],
            'satuan': item_sj['satuan'],
            'qty': item_sj['qty']
        })

print(f"Selesai mapping. Ditemukan {len(matched_items)} prediksi sukses, {len(unmatched_items)} produk baru.")

wb_review = openpyxl.Workbook()
ws_review = wb_review.active
ws_review.title = "Review Mapping"
headers_review = ["ID SJ", "Nama SJ", "Qty Transfer", "Prediksi SKU LUNA", "Prediksi Nama LUNA", "Harga Modal dr SJ", "STATUS MATCH"]
ws_review.append(headers_review)

for m in mapping_review:
    status = "OK (Perlu Review)" if m['sku_luna_prediksi'] else "BARU (Tidak ada di LUNA)"
    ws_review.append([
        m['id_sj'], m['nama_sj'], m['qty_sj'], 
        m['sku_luna_prediksi'] or "N/A", 
        m['nama_luna_prediksi'] or "N/A", 
        m['harga_modal_sj'],
        status
    ])
auto_resize_columns(ws_review)
# Simpan ke folder output
review_file = os.path.join(output_dir, f"Hasil_Mapping_Review_{file_suffix}.xlsx")
wb_review.save(review_file)

wb_tf = openpyxl.load_workbook('warehouse-transfer-import-template.xlsx')
ws_tf = wb_tf.active
if ws_tf.max_row >= 4:
    ws_tf.delete_rows(4, ws_tf.max_row - 3)

for i, m in enumerate(matched_items, 1):
    row = [i, m['sku'], m['nama'], m['satuan'], m['qty']] + ([None] * 20)
    ws_tf.append(row)
auto_resize_columns(ws_tf)
# Simpan ke folder output
tf_file = os.path.join(output_dir, f"Siap_Warehouse_Transfer_{file_suffix}.xlsx")
wb_tf.save(tf_file)

wb_new = openpyxl.load_workbook('product-import-template.xlsx')
ws_new = wb_new.active
if ws_new.max_row >= 4:
    ws_new.delete_rows(4, ws_new.max_row - 3)

for i, u in enumerate(unmatched_items, 1):
    harga_jual = u['harga_modal']
    harga_modal_luna = 0 # Request Finance
    nama_baru = format_nama_luna(u['nama'], warehouse_prefix)
    kategori = OUTLET_CATEGORIES.get(warehouse_prefix, warehouse_prefix)
    row = [
        i, "", nama_baru, "Y", "N", 
        harga_jual, harga_modal_luna, "Y", u['qty'], 1, 
        kategori, str(u['satuan']).upper() if u['satuan'] else "PCS"
    ] + ([None] * 8)
    ws_new.append(row)
auto_resize_columns(ws_new)
# Simpan ke folder output
new_prod_file = os.path.join(output_dir, f"Siap_Product_Baru_{file_suffix}.xlsx")
wb_new.save(new_prod_file)

# ---- FITUR LAPORAN OTOMATIS AKHIR ----
total_qty_transfer = sum(m['qty'] for m in matched_items)
total_qty_baru = sum(u['qty'] for u in unmatched_items)
total_item_transfer = len(matched_items)
total_item_baru = len(unmatched_items)

total_rp_transfer = sum(safe_int(m['harga_modal_sj']) * safe_int(m['qty_sj']) for m in mapping_review if m['sku_luna_prediksi'])
total_rp_baru = sum(safe_int(m['harga_modal_sj']) * safe_int(m['qty_sj']) for m in mapping_review if not m['sku_luna_prediksi'])

rp_transfer_str = f"Rp {total_rp_transfer:,}".replace(',', '.')
rp_baru_str = f"Rp {total_rp_baru:,}".replace(',', '.')

laporan_text = f"""=========================================
LAPORAN SINKRONISASI INVENTORI LUNA POS
=========================================
File Sumber     : {os.path.basename(sj_filename)}
Cabang Target   : {warehouse_prefix}

-- STATUS PEMBACAAN DATA --
Peringatan      : Terdapat {baris_gagal} baris barang dilewati (ID Kosong)

-- RINGKASAN MUTASI (STOCK IN LAMA) --
Total SKU Terdikses                : {total_item_transfer} Varian
Total Quantitiy Masuk              : {total_qty_transfer} Pcs
Total Nilai Barang (Harga Satuan)  : {rp_transfer_str}

-- RINGKASAN REGISTRASI BARANG BARU --
Total SKU Baru Terdikses           : {total_item_baru} Varian
Total Quantitiy Masuk              : {total_qty_baru} Pcs
Total Nilai Barang (Harga Satuan)  : {rp_baru_str}

-- DAFTAR FILE HASIL --
1. Hasil_Mapping_Review_{file_suffix}.xlsx (WAJIB CEK)
2. Siap_Warehouse_Transfer_{file_suffix}.xlsx
3. Siap_Product_Baru_{file_suffix}.xlsx

========================================="""

laporan_filename = f"Laporan_Mutasi_{file_suffix}.txt"
laporan_path = os.path.join(output_dir, laporan_filename)
with open(laporan_path, "w", encoding="utf-8") as f:
    f.write(laporan_text)

# Pindahkan file Surat Jalan (SJ) asli ke folder output untuk arsip
try:
    target_sj_path = os.path.join(output_dir, os.path.basename(sj_filename))
    if os.path.exists(target_sj_path):
        os.remove(target_sj_path)
    
    shutil.move(sj_filename, target_sj_path)
    print(f"📦 File sumber '{os.path.basename(sj_filename)}' dipindahkan ke folder {branch_name} untuk arsip.")
except PermissionError:
    print(f"⚠️ PERINGATAN: File '{os.path.basename(sj_filename)}' sedang dibuka oleh program lain (Excel?).")
    print(f"   --> File GAGAL dipindahkan otomatis, tapi hasil generate DI DALAM FOLDER {branch_name} tetap aman.")
except Exception as e:
    print(f"⚠️ Peringatan: Gagal memindahkan file sumber: {e}")

print(f"\n{laporan_text}\n")
print(f"✅ Semua file berhasil disimpan di folder: {output_dir}")
