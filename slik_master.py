import os
import pdfplumber
import re
import pandas as pd

from slik_extractor import extract_credit_blocks, extract_debitur_name, extract_nomor_laporan


folder_path = "slik_data"
all_records = []

for filename in os.listdir(folder_path):
    if filename.lower().endswith(".pdf"):
        full_path = os.path.join(folder_path, filename)
        print(f"Memproses: {filename}")

        # Baca seluruh teks PDF
        pages_text = []
        with pdfplumber.open(full_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                pages_text.append(text)

        full_text = "\n".join(pages_text)

        debitur = extract_debitur_name(full_text)
        nomor_laporan = extract_nomor_laporan(full_text)
        records = extract_credit_blocks(full_text)

        for r in records:
            r["Nama Debitur"] = debitur
            r["Nomor Laporan"] = nomor_laporan
            all_records.append(r)


if not all_records:
    print("Tidak ada data ditemukan.")
else:
    df = pd.DataFrame(all_records)

    df.to_excel("MASTER_SLIK.xlsx", index=False)
    print("\n✅ MASTER_SLIK.xlsx berhasil dibuat!")
    print(f"Total baris: {len(df)}")
