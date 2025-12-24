import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Fungsi Utama ---
def mulai_proses():
    # 1. Dialog pilih folder PDF Asli
    folder_asal = filedialog.askdirectory(title="Langkah 1: Pilih Folder Tempat File PDF Asli Berada")
    if not folder_asal:
        return # Batalkan jika user menekan Cancel

    # 2. Dialog pilih file Excel
    file_excel = filedialog.askopenfilename(
        title="Langkah 2: Pilih File Excel Daftar Nama",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_excel:
        return # Batalkan jika user menekan Cancel

    # Update status di GUI
    label_status.config(text="Sedang memproses... Mohon tunggu.", fg="blue")
    root.update()

    try:
        # 3. Tentukan Nama Folder Baru secara otomatis
        parent_dir = os.path.dirname(folder_asal)
        folder_tujuan_nama = "HASIL_RENAME_" + os.path.basename(folder_asal)
        folder_tujuan = os.path.join(parent_dir, folder_tujuan_nama)

        if not os.path.exists(folder_tujuan):
            os.makedirs(folder_tujuan)

        # 4. Ambil daftar PDF & Urutkan A-Z
        files = [f for f in os.listdir(folder_asal) if f.lower().endswith('.pdf')]
        files.sort()

        if not files:
             messagebox.showwarning("Peringatan", "Tidak ada file PDF ditemukan di folder asal.")
             label_status.config(text="Siap.", fg="black")
             return

        # 5. Baca Excel (Kolom A / indeks 0)
        df = pd.read_excel(file_excel, header=None)
        nama_baru_list = df[0].astype(str).tolist()

        # 6. Proses Copy & Rename
        limit = min(len(files), len(nama_baru_list))
        count = 0
        errors = 0

        for i in range(limit):
            old_name = files[i]
            new_name_raw = nama_baru_list[i].strip()
            
            # Bersihkan karakter terlarang Windows
            new_name_clean = new_name_raw
            for char in '<>:"/\\|?*':
                new_name_clean = new_name_clean.replace(char, '')
            
            # Pastikan ekstensi .pdf
            if not new_name_clean.lower().endswith('.pdf'):
                new_name_clean += ".pdf"

            src_path = os.path.join(folder_asal, old_name)
            dst_path = os.path.join(folder_tujuan, new_name_clean)

            try:
                shutil.copy2(src_path, dst_path)
                count += 1
            except Exception:
                errors += 1

        # Pesan Selesai
        msg = f"Proses Selesai!\n\nBerhasil rename & copy: {count} file.\nGagal: {errors} file.\n\nFolder hasil:\n{folder_tujuan}"
        messagebox.showinfo("Sukses", msg)
        label_status.config(text="Selesai. Siap untuk proses berikutnya.", fg="green")

    except Exception as e:
        messagebox.showerror("Error Fatal", f"Terjadi kesalahan:\n{e}")
        label_status.config(text="Terjadi Error.", fg="red")

# --- Konfigurasi GUI ---
root = tk.Tk()
root.title("Aplikasi Rename PDF Massal")
root.geometry("500x250")
root.resizable(False, False)

# Frame Utama
main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(expand=True, fill=tk.BOTH)

# Label Instruksi
label_intro = tk.Label(main_frame, text="Aplikasi untuk mengubah nama file PDF secara massal\nberdasarkan urutan data di Excel.", font=("Arial", 10), pady=10)
label_intro.pack()

# Tombol Utama
btn_mulai = tk.Button(main_frame, text="MULAI PROSES (Pilih Folder & Excel)", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", height=2, command=mulai_proses)
btn_mulai.pack(fill=tk.X, pady=20)

# Label Status
label_status = tk.Label(main_frame, text="Siap.", font=("Arial", 9), fg="grey", bd=1, relief=tk.SUNKEN, anchor=tk.W)
label_status.pack(side=tk.BOTTOM, fill=tk.X)

# Jalankan Aplikasi
root.mainloop()
