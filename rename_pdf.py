import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def jalankan_proses():
    # 1. Pilih Folder PDF
    folder_asal = filedialog.askdirectory(title="Pilih Folder PDF Asli")
    if not folder_asal: return

    # 2. Pilih File Excel
    file_excel = filedialog.askopenfilename(
        title="Pilih File Excel", 
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_excel: return

    # 3. Tentukan Nama Folder Baru
    folder_tujuan = os.path.join(os.path.dirname(folder_asal), "HASIL_RENAME_PDF")
    
    try:
        if not os.path.exists(folder_tujuan):
            os.makedirs(folder_tujuan)

        # Ambil daftar PDF & Urutkan
        files = [f for f in os.listdir(folder_asal) if f.lower().endswith('.pdf')]
        files.sort()

        # Baca Excel
        df = pd.read_excel(file_excel, header=None)
        nama_baru_list = df[0].astype(str).tolist()

        limit = min(len(files), len(nama_baru_list))
        count = 0

        for i in range(limit):
            old_name = files[i]
            new_name = nama_baru_list[i].strip()
            
            # Bersihkan karakter terlarang
            for char in '<>:"/\\|?*':
                new_name = new_name.replace(char, '')
            
            if not new_name.lower().endswith('.pdf'):
                new_name += ".pdf"

            shutil.copy2(os.path.join(folder_asal, old_name), os.path.join(folder_tujuan, new_name))
            count += 1

        messagebox.showinfo("Selesai", f"Berhasil memproses {count} file.\nHasil ada di folder: {folder_tujuan}")
    
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")

if __name__ == "__main__":
    # Sembunyikan jendela utama tkinter
    root = tk.Tk()
    root.withdraw()
    jalankan_proses()
