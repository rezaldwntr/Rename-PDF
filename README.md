# 🛠️ PDF Tools Master

**Aplikasi Desktop All-in-One untuk Otomatisasi Dokumen PDF**

Project ini menggabungkan berbagai alat manipulasi PDF menjadi satu aplikasi ringan yang mudah digunakan. Dibuat khusus untuk mempercepat pekerjaan administrasi, HRD, dan pengelolaan dokumen kantor tanpa perlu mengerti koding.


## ✨ Fitur Utama (3-in-One)

### 1. Simple Renamer 📝
Mengubah nama file PDF secara massal dengan mencocokkan **urutan file** di folder dengan **baris data** di Excel.
* *Cocok untuk:* Data yang sudah urut abjad/nomor.

### 2. Advanced Renamer 📊
Fitur rename yang lebih canggih dengan tampilan tabel interaktif.
* **Live Preview:** Lihat nama lama vs nama baru sebelum dieksekusi.
* **Search & Filter:** Cari file tertentu dengan cepat.
* **Bulk Replace:** Ganti kata tertentu secara massal (Misal: ganti "2024" jadi "2025" untuk semua file).
* **Manual Edit:** Bisa edit nama file satu per satu langsung di tabel.

### 3. Smart PDF Merger 📎
Menggabungkan dua file PDF terpisah (misal: "Surat Isi" dan "Lembar Tanda Tangan") menjadi satu file utuh secara otomatis.
* Sistem pencocokan otomatis berdasarkan Nama.
* Contoh: `Budi_Isi.pdf` + `Budi_TTD.pdf` otomatis menjadi `Budi_Lengkap.pdf`.

---

## 🚀 Cara Download & Pakai (Untuk Pengguna Biasa)
Tidak perlu install Python! Cukup download aplikasinya:

1.  Buka tab **[Releases](../../releases)** di pojok kanan halaman ini.
2.  Download file terbaru: **`PDF_Tools_Master.exe`**.
3.  Simpan di folder dokumen Anda, lalu klik dua kali untuk menjalankan.

> **Catatan:** Jika muncul peringatan *"Windows protected your PC"*, itu normal untuk aplikasi buatan sendiri. Klik **More Info** -> **Run Anyway**.

---

## ⚙️ Cara Install & Modifikasi (Untuk Developer)
Jika Anda ingin melihat source code atau memodifikasi aplikasinya:

### Prasyarat
* Python 3.x
* Git

### Langkah Instalasi
1.  Clone repository ini:
    ```bash
    git clone [https://github.com/rezaldwntr/PDF-Tools-Master.git](https://github.com/rezaldwntr/PDF-Tools-Master.git)
    ```
2.  Install library yang dibutuhkan:
    ```bash
    pip install -r requirements.txt
    ```
3.  Jalankan aplikasi:
    ```bash
    python Main_App.py
    ```

### Cara Membuat File .EXE Sendiri
Cukup klik dua kali file **`Install_Apps.bat`** yang sudah disediakan. Script akan otomatis:
1.  Menginstall library kurang.
2.  Membuild aplikasi menggunakan PyInstaller.
3.  Menghasilkan file `.exe` siap pakai.

---

## 👤 Author
**Rezal Dewantara**
* GitHub: [@rezaldwntr](https://github.com/rezaldwntr)

---
*Dibuat dengan Python (Tkinter, Pandas, PyPDF).*
