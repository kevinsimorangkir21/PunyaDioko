import os
import re


def clean_filename(filename):
    # Pisahkan nama dan ekstensi
    name, ext = os.path.splitext(filename)

    # Ganti spasi jadi underscore
    name = name.replace(" ", "_")

    # Hapus karakter aneh selain huruf, angka, underscore
    name = re.sub(r'[^A-Za-z0-9_]', '', name)

    return name + ext


def rename_pdfs(folder_path):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            old_path = os.path.join(folder_path, filename)
            new_filename = clean_filename(filename)
            new_path = os.path.join(folder_path, new_filename)

            if old_path != new_path:
                os.rename(old_path, new_path)
                print(f"Renamed: {filename} ➜ {new_filename}")


if __name__ == "__main__":
    folder = "."  # folder saat ini
    rename_pdfs(folder)
    print("\n✅ Semua PDF sudah direname!")
