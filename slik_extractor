#!/usr/bin/env python3
"""
SLIK OJK Credit/Financing Data Extractor
Mengekstrak data kredit/pembiayaan dari PDF SLIK OJK ke format Excel/CSV
"""

import re
import sys
from pathlib import Path
from typing import List, Dict, Any
import PyPDF2
import pandas as pd


def extract_debtor_name(text: str) -> str:
    """Ekstrak nama debitur dari teks PDF"""
    patterns = [
        r'Nama\s*\n\s*([A-Z\s]+)',
        r'Nama\n([A-Z]+)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            name = match.group(1).strip()
            # Filter nama yang valid (minimal 3 karakter)
            if name and len(name) > 2 and not name.startswith('LAKI'):
                return name
    
    return "Tidak terdeteksi"


def extract_report_number(text: str) -> str:
    """Ekstrak nomor laporan"""
    pattern = r'Nomor Laporan[\s:]+(\d+/[A-Z]+/\d+/\d+)'
    match = re.search(pattern, text)
    return match.group(1) if match else "Tidak terdeteksi"


def parse_credit_block(block: str) -> Dict[str, Any]:
    """Parse satu blok kredit/pembiayaan menjadi dictionary"""
    data = {}
    
    # Bank/Pelapor - Format: "014 - PT Bank Central Asia Tbk"
    pelapor_patterns = [
        r'(\d{3,6})\s*-\s*(.+?)\s+(?:Rp|Baki Debet)',
        r'Pelapor\s+Cabang\s+Baki Debet.*?\n(.+?)\s+(.+?)\s+Rp',
    ]
    
    for pattern in pelapor_patterns:
        pelapor_match = re.search(pattern, block, re.MULTILINE | re.DOTALL)
        if pelapor_match:
            if len(pelapor_match.groups()) == 2:
                data['Kode_Pelapor'] = pelapor_match.group(1).strip()
                bank_name = pelapor_match.group(2).strip()
                # Bersihkan nama bank
                bank_name = re.sub(r'\s+', ' ', bank_name)
                data['Bank'] = bank_name
                break
    
    if 'Bank' not in data:
        data['Kode_Pelapor'] = ''
        data['Bank'] = ''
    
    # No Rekening
    no_rek_match = re.search(r'No Rekening\s+(.+?)\s+Kualitas', block, re.DOTALL)
    data['No_Rekening'] = no_rek_match.group(1).strip().replace('\n', ' ') if no_rek_match else ''
    
    # Kualitas - Format: "1 - Lancar" atau "2 - Dalam Perhatian Khusus"
    kualitas_match = re.search(r'Kualitas\s+(\d+)\s*-\s*(.+?)(?:\n|$)', block)
    if kualitas_match:
        data['Kualitas_Kode'] = kualitas_match.group(1).strip()
        data['Kualitas'] = kualitas_match.group(2).strip()
    else:
        data['Kualitas_Kode'] = ''
        data['Kualitas'] = ''
    
    # Baki Debet - Format: "Rp 9.455.927,00"
    baki_match = re.search(r'Baki Debet\s+Rp\s*([\d\.,]+)', block)
    if baki_match:
        baki_str = baki_match.group(1).replace('.', '').replace(',', '.')
        try:
            data['Baki_Debet'] = float(baki_str)
        except:
            data['Baki_Debet'] = 0
    else:
        data['Baki_Debet'] = 0
    
    # Jenis Penggunaan - Format: "Investasi", "Konsumsi", "Modal Kerja"
    penggunaan_match = re.search(r'Jenis Penggunaan\s+(.+?)(?:\s+Frekuensi|\n)', block)
    data['Jenis_Penggunaan'] = penggunaan_match.group(1).strip() if penggunaan_match else ''
    
    # Jenis Kredit/Pembiayaan
    jenis_kredit_match = re.search(r'Jenis Kredit/Pembiayaan\s+(.+?)(?:\s+Nilai Proyek|\n)', block)
    data['Jenis_Kredit'] = jenis_kredit_match.group(1).strip() if jenis_kredit_match else ''
    
    # No Akad Awal
    akad_awal_match = re.search(r'No Akad Awal\s+(.+?)(?:\s+Realisasi|\n)', block)
    data['No_Akad_Awal'] = akad_awal_match.group(1).strip() if akad_awal_match else ''
    
    # Suku Bunga/Imbalan - Format: "21 %"
    bunga_match = re.search(r'Suku Bunga/Imbalan\s+([\d\.]+)\s*%', block)
    data['Suku_Bunga'] = float(bunga_match.group(1)) if bunga_match else 0.0
    
    # Jenis Suku Bunga
    jenis_bunga_match = re.search(r'Jenis Suku Bunga/Imbalan\s+(.+?)(?:\n|Sifat)', block)
    data['Jenis_Suku_Bunga'] = jenis_bunga_match.group(1).strip() if jenis_bunga_match else ''
    
    # Jumlah Hari Tunggakan
    tunggakan_match = re.search(r'Jumlah Hari Tunggakan\s+(\d+)', block)
    data['Hari_Tunggakan'] = int(tunggakan_match.group(1)) if tunggakan_match else 0
    
    # Tanggal Akad Awal - Format: "08 Februari 2010"
    tgl_akad_match = re.search(r'Tanggal Akad Awal\s+(\d{2}\s+\w+\s+\d{4})', block)
    data['Tanggal_Akad_Awal'] = tgl_akad_match.group(1).strip() if tgl_akad_match else ''
    
    # Tanggal Jatuh Tempo - Format: "31 Januari 2030"
    tgl_tempo_match = re.search(r'Tanggal Jatuh Tempo\s+(\d{2}\s+\w+\s+\d{4})', block)
    data['Tanggal_Jatuh_Tempo'] = tgl_tempo_match.group(1).strip() if tgl_tempo_match else ''
    
    # Sektor Ekonomi
    sektor_match = re.search(r'Sektor Ekonomi\s+(.+?)(?:\s+Tanggal Restrukturisasi|\n)', block)
    data['Sektor_Ekonomi'] = sektor_match.group(1).strip() if sektor_match else ''
    
    # Kondisi Fasilitas - Format: "Aktif", "Lunas", "Dialihkan ke Fasilitas lain"
    kondisi_match = re.search(r'Kondisi\s+(.+?)(?:\n|Keterangan)', block)
    data['Kondisi'] = kondisi_match.group(1).strip() if kondisi_match else ''
    
    # Plafon Awal - Format: "Rp 10.000.000,00"
    plafon_awal_match = re.search(r'Plafon Awal\s+Rp\s*([\d\.,]+)', block)
    if plafon_awal_match:
        plafon_str = plafon_awal_match.group(1).replace('.', '').replace(',', '.')
        try:
            data['Plafon_Awal'] = float(plafon_str)
        except:
            data['Plafon_Awal'] = 0
    else:
        data['Plafon_Awal'] = 0
    
    # Plafon (current) - harus beda dari Plafon Awal
    plafon_patterns = [
        r'Perpanjangan[^\n]*\n\s*\d+\s+Plafon\s+Rp\s*([\d\.,]+)',
        r'(?<!Awal)\s+Plafon\s+Rp\s*([\d\.,]+)',
    ]
    
    for pattern in plafon_patterns:
        plafon_match = re.search(pattern, block)
        if plafon_match:
            plafon_str = plafon_match.group(1).replace('.', '').replace(',', '.')
            try:
                data['Plafon'] = float(plafon_str)
                break
            except:
                pass
    
    if 'Plafon' not in data:
        data['Plafon'] = 0
    
    # Tunggakan Pokok
    tunggakan_pokok_match = re.search(r'Tunggakan Pokok\s+Rp\s*([\d\.,]+)', block)
    if tunggakan_pokok_match:
        tunggakan_str = tunggakan_pokok_match.group(1).replace('.', '').replace(',', '.')
        try:
            data['Tunggakan_Pokok'] = float(tunggakan_str)
        except:
            data['Tunggakan_Pokok'] = 0
    else:
        data['Tunggakan_Pokok'] = 0
    
    # Frekuensi Restrukturisasi
    restruk_match = re.search(r'Frekuensi Restrukturisasi\s+(\d+)', block)
    data['Frekuensi_Restrukturisasi'] = int(restruk_match.group(1)) if restruk_match else 0
    
    # Tanggal Update
    update_match = re.search(r'Tanggal Update\s*\n\s*(\d{2}\s+\w+\s+\d{4})', block)
    data['Tanggal_Update'] = update_match.group(1).strip() if update_match else ''
    
    # Denda
    denda_match = re.search(r'Denda\s+Rp\s*([\d\.,]+)', block)
    if denda_match:
        denda_str = denda_match.group(1).replace('.', '').replace(',', '.')
        try:
            data['Denda'] = float(denda_str)
        except:
            data['Denda'] = 0
    else:
        data['Denda'] = 0
    
    return data


def extract_credits_from_pdf(pdf_path: str):
    """Ekstrak semua data kredit/pembiayaan dari PDF"""
    
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Gabungkan semua halaman
        full_text = ""
        for page in pdf_reader.pages:
            full_text += page.extract_text() + "\n"
    
    # Ekstrak nama debitur dan nomor laporan
    debtor_name = extract_debtor_name(full_text)
    report_number = extract_report_number(full_text)
    
    # Split berdasarkan pattern "Kredit/Pembiayaan"
    credit_blocks = re.split(r'Kredit/Pembiayaan\s*\n', full_text)
    
    credits = []
    for i, block in enumerate(credit_blocks[1:], 1):  # Skip blok pertama (header)
        # Ambil hanya sampai bagian "Agunan" atau "Penjamin"
        block_end = re.search(r'(?=Agunan|Penjamin|Irrevocable)', block)
        if block_end:
            clean_block = block[:block_end.start()]
        else:
            clean_block = block
        
        # Filter blok yang terlalu pendek
        if len(clean_block) > 100:
            credit_data = parse_credit_block(clean_block)
            # Hanya tambahkan jika ada data bank
            if credit_data.get('Bank'):
                credits.append(credit_data)
    
    return debtor_name, report_number, credits


def save_to_excel(debtor_name: str, report_number: str, credits: List[Dict[str, Any]], output_path: str):
    """Simpan data ke file Excel dengan format yang rapi"""
    
    # Buat DataFrame
    df = pd.DataFrame(credits)
    
    # Urutan kolom sesuai permintaan
    columns_order = [
        'Bank',
        'Jenis_Penggunaan',
        'No_Rekening',
        'Plafon_Awal',
        'Baki_Debet',
        'Suku_Bunga',
        'Tanggal_Akad_Awal',
        'Tanggal_Jatuh_Tempo',
        'Kualitas',
        'Frekuensi_Restrukturisasi',
        # Kolom tambahan
        'Jenis_Kredit',
        'No_Akad_Awal',
        'Jenis_Suku_Bunga',
        'Hari_Tunggakan',
        'Sektor_Ekonomi',
        'Kondisi',
        'Plafon',
        'Tunggakan_Pokok',
        'Denda',
        'Tanggal_Update',
        'Kualitas_Kode',
        'Kode_Pelapor'
    ]
    
    # Gunakan kolom yang ada
    available_columns = [col for col in columns_order if col in df.columns]
    df = df[available_columns]
    
    # Rename kolom agar lebih user-friendly
    rename_dict = {
        'Jenis_Penggunaan': 'Jenis Penggunaan',
        'No_Rekening': 'No Rekening',
        'Plafon_Awal': 'Plafon Awal (IDR)',
        'Baki_Debet': 'Baki Debet (IDR)',
        'Suku_Bunga': 'Suku Bunga/Imbalan (%)',
        'Tanggal_Akad_Awal': 'Tanggal Akad Awal',
        'Tanggal_Jatuh_Tempo': 'Tanggal Jatuh Tempo',
        'Frekuensi_Restrukturisasi': 'Frekuensi Restrukturisasi',
        'Jenis_Kredit': 'Jenis Kredit/Pembiayaan',
        'No_Akad_Awal': 'No Akad Awal',
        'Jenis_Suku_Bunga': 'Jenis Suku Bunga',
        'Hari_Tunggakan': 'Hari Tunggakan',
        'Sektor_Ekonomi': 'Sektor Ekonomi',
        'Plafon': 'Plafon (IDR)',
        'Tunggakan_Pokok': 'Tunggakan Pokok (IDR)',
        'Denda': 'Denda (IDR)',
        'Tanggal_Update': 'Tanggal Update',
        'Kualitas_Kode': 'Kode Kualitas',
        'Kode_Pelapor': 'Kode Pelapor'
    }
    df = df.rename(columns=rename_dict)
    
    # Simpan ke Excel dengan multiple sheets
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: Informasi Umum
        info_df = pd.DataFrame({
            'Informasi': ['Nama Debitur', 'Nomor Laporan', 'Jumlah Kredit/Pembiayaan', 'Total Baki Debet (IDR)'],
            'Nilai': [
                debtor_name, 
                report_number, 
                len(credits),
                f"Rp {df['Baki Debet (IDR)'].sum():,.2f}" if 'Baki Debet (IDR)' in df.columns else 'N/A'
            ]
        })
        info_df.to_excel(writer, sheet_name='Informasi', index=False)
        
        # Sheet 2: Data Kredit (kolom utama saja)
        main_columns = [col for col in df.columns if any(x in col for x in [
            'Bank', 'Jenis Penggunaan', 'No Rekening', 'Plafon Awal', 'Baki Debet', 
            'Suku Bunga', 'Tanggal Akad', 'Tanggal Jatuh Tempo', 'Kualitas', 'Frekuensi Restrukturisasi'
        ])]
        df[main_columns].to_excel(writer, sheet_name='Data Kredit', index=False)
        
        # Sheet 3: Data Lengkap
        df.to_excel(writer, sheet_name='Data Lengkap', index=False)
    
    print(f"\n✅ File Excel berhasil disimpan: {output_path}")


def save_to_csv(debtor_name: str, report_number: str, credits: List[Dict[str, Any]], output_path: str):
    """Simpan data ke file CSV"""
    
    df = pd.DataFrame(credits)
    
    # Kolom yang diminta
    columns_order = [
        'Bank',
        'Jenis_Penggunaan',
        'No_Rekening',
        'Plafon_Awal',
        'Baki_Debet',
        'Suku_Bunga',
        'Tanggal_Akad_Awal',
        'Tanggal_Jatuh_Tempo',
        'Kualitas',
        'Frekuensi_Restrukturisasi'
    ]
    
    available_columns = [col for col in columns_order if col in df.columns]
    df = df[available_columns]
    
    # Rename kolom
    rename_dict = {
        'Jenis_Penggunaan': 'Jenis Penggunaan',
        'No_Rekening': 'No Rekening',
        'Plafon_Awal': 'Plafon Awal',
        'Baki_Debet': 'Baki Debet',
        'Suku_Bunga': 'Suku Bunga (%)',
        'Tanggal_Akad_Awal': 'Tanggal Akad Awal',
        'Tanggal_Jatuh_Tempo': 'Tanggal Jatuh Tempo',
        'Frekuensi_Restrukturisasi': 'Frekuensi Restrukturisasi'
    }
    df = df.rename(columns=rename_dict)
    
    # Tambahkan info di header
    with open(output_path, 'w', encoding='utf-8-sig') as f:
        f.write(f"# SLIK Debitur: {debtor_name}\n")
        f.write(f"# Nomor Laporan: {report_number}\n")
        f.write(f"# Jumlah Kredit: {len(credits)}\n")
        f.write("\n")
    
    df.to_csv(output_path, mode='a', index=False, encoding='utf-8-sig')
    print(f"\n✅ File CSV berhasil disimpan: {output_path}")


def main():
    """Main function"""
    
    print(f"\n{'='*80}")
    print(" " * 20 + "SLIK OJK - CREDIT EXTRACTOR")
    print(f"{'='*80}")
    
    if len(sys.argv) < 2:
        print("\n❌ Error: File PDF tidak diberikan!")
        print("\nCara Penggunaan:")
        print("  python slik_extractor.py <nama_file.pdf> [opsi]")
        print("\nContoh:")
        print("  python slik_extractor.py BAJURI.pdf              # Output: Excel (default)")
        print("  python slik_extractor.py BAJURI.pdf --csv        # Output: CSV")
        print("  python slik_extractor.py BAJURI.pdf --xlsx       # Output: Excel")
        print("  python slik_extractor.py BAJURI.pdf --csv --xlsx # Output: CSV + Excel")
        print(f"\n{'='*80}\n")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    
    if not Path(pdf_path).exists():
        print(f"\n❌ Error: File tidak ditemukan: {pdf_path}\n")
        sys.exit(1)
    
    print(f"\n📄 Memproses file: {pdf_path}")
    print("⏳ Mengekstrak data...")
    
    try:
        # Ekstrak data
        debtor_name, report_number, credits = extract_credits_from_pdf(pdf_path)
        
        print(f"\n{'='*80}")
        print(f"✅ HASIL EKSTRAKSI")
        print(f"{'='*80}")
        print(f"  Nama Debitur         : {debtor_name}")
        print(f"  Nomor Laporan        : {report_number}")
        print(f"  Jumlah Kredit/Pembiayaan : {len(credits)}")
        print(f"{'='*80}")
        
        if not credits:
            print("\n⚠️  Peringatan: Tidak ada data kredit/pembiayaan yang berhasil diekstrak!")
            print("    Pastikan PDF berisi data kredit/pembiayaan dengan format SLIK OJK.\n")
            sys.exit(1)
        
        # Tentukan format output
        output_csv = '--csv' in sys.argv
        output_xlsx = '--xlsx' in sys.argv
        
        # Default ke XLSX jika tidak ada flag
        if not output_csv and not output_xlsx:
            output_xlsx = True
        
        # Generate output filename
        base_name = Path(pdf_path).stem
        
        # Simpan file
        if output_xlsx:
            output_file = f"{base_name}_SLIK.xlsx"
            save_to_excel(debtor_name, report_number, credits, output_file)
        
        if output_csv:
            output_file = f"{base_name}_SLIK.csv"
            save_to_csv(debtor_name, report_number, credits, output_file)
        
        print(f"\n{'='*80}")
        print("✅ PROSES SELESAI!")
        print(f"{'='*80}\n")
        
    except Exception as e:
        print(f"\n❌ Error saat memproses file: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()