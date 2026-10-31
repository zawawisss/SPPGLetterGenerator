# Generator Dokumen SPPG Joresan

Aplikasi CLI berbasis Python untuk membuat dokumen **BAST**, **Surat Jalan**, dan **Organoleptik** secara otomatis dari data `Database.xlsx`, lalu menggabungkannya menjadi satu file per jenis dokumen.

---

## Fitur

- Membuat 3 jenis dokumen sekaligus: **BAST**, **Surat Jalan**, dan **Organoleptik**.
- Data lokasi diambil dari `Database/Database.xlsx` (sekolah & posyandu).
- Nomor BAST dibuat otomatis dan berurutan, dengan **saran nomor terakhir + 1**.
- Nomor BAST bisa dimulai dari angka berapa pun (misal `020`).
- Nomor yang sudah dibuat langsung disimpan ke `Database/nomor_bast.json`.
- Lembaga + tanggal yang sama akan memakai nomor yang sama (tidak dobel).
- Bulan Romawi dan tahun mengikuti tanggal pengiriman.
- Menu navigasi panah ↑/↓ tanpa refresh seluruh layar.
- Cek dependensi otomatis; kalau belum ada, di-install otomatis lewat `pip`.

---

## Struktur Folder

```
Buat Surat/
├── main.py                  # Program utama
├── Buat_Surat.bat           # Launcher Windows
├── README.md
├── Database/
│   ├── Database.xlsx        # Data sekolah & posyandu
│   ├── Database_backup.xlsx # Backup database
│   └── nomor_bast.json      # Riwayat nomor BAST (dibuat otomatis)
├── Template/
│   ├── BAST.docx
│   ├── BAST_backup.docx
│   ├── ORGANOLEPTIK.docx
│   ├── SURAT JALAN SEKOLAH.docx
│   └── SURAT JALAN POSYANDU.docx
└── HASIL/                   # Hasil dokumen per tanggal
```

> Jangan memindahkan/mengganti nama folder `Database` dan `Template` tanpa menyesuaikan path di `main.py`.

---

## Persyaratan

- **Windows** (utama). Linux/macOS didukung sebagian lewat fallback terminal.
- **Python 3.10+** (disarankan 3.12/3.14).
- Paket Python:
  - `openpyxl`
  - `python-docx`
  - `docxcompose`

Dependensi **tidak perlu di-install manual** — `main.py` akan mengecek dan meng-install otomatis saat pertama dijalankan (butuh koneksi internet).

Install manual jika diperlukan:

```bat
python -m pip install openpyxl python-docx docxcompose
```

---

## Cara Menjalankan

1. Buka folder `Buat Surat`.
2. Klik dua kali **`Buat_Surat.bat`**.
3. Ikuti menu di jendela terminal.

`Buat_Surat.bat` otomatis:
- pindah ke folder aplikasi (`cd /d "%~dp0"`),
- mengecek Python tersedia,
- menjalankan `main.py` (yang mengecek/meng-install dependensi).

---

## Alur Penggunaan

### Langkah 1 — Pilih Tanggal

Pilih dengan panah ↑/↓ lalu **Enter**:

- `HARI INI`
- `BESOK`
- `LUSA`
- `INPUT TANGGAL MANUAL` → ketik `dd/mm/yyyy`

Tanggal ini menentukan hari, **bulan Romawi**, dan **tahun** pada nomor BAST.

### Langkah 2 — Nomor Surat BAST

Contoh tampilan:

```
Nomor terakhir tercatat : 019
Saran nomor berikutnya  : 020

Mulai nomor berapa? [Enter = 020]:
```

- **Enter** → memakai saran (`nomor terakhir + 1`).
- Ketik angka lain, misal `025` → mulai dari `025`.
- Kalau angka yang dimasukkan ≤ nomor terakhir, muncul peringatan konfirmasi `y/n`.

### Langkah 3 — Pilih Kategori Layanan

- `SEKOLAH`
- `POSYANDU BALITA`
- `POSYANDU B2 (BUMIL & BUSUI)`
- `POSYANDU SEMUA (BALITA & B2)`
- `SEMUA LAYANAN (SEKOLAH + POSYANDU)`
- `>>> SELESAI & CETAK DOKUMEN <<<`

Kategori bisa ditambahkan beberapa kali (keranjang).

### Langkah 4 — Input Daftar Menu

Ketik menu satu per satu, lalu ketik `selesai`:

```
Menu #1: Nasi Putih
Menu #2: Ayam Goreng
Menu #3: Sayur Sop
Menu #4: selesai
```

Minimal 1 menu.

### Konfirmasi

Ringkasan lokasi tampil. Tekan:

- **Enter** → cetak dokumen
- **K** → kembali ke menu

---

## Hasil Output

Dokumen disimpan di:

```
HASIL/<tanggal>-<bulan>-<tahun>/
```

Contoh untuk 3 November 2026:

```
HASIL/3-11-26/
├── BAST 3-11-26.docx
├── ORGANOLEPTIK 3-11-26.docx
└── SURAT_JALAN 3-11-26.docx
```

Keterangan:

- **BAST** → tiap lembaga 2 halaman (Penerimaan + Pengembalian) dengan nomor yang sama.
- **Surat Jalan** → digabung per lokasi.
- **Organoleptik** → mengikuti isi keranjang (bisa terpisah jika menu berbeda).

---

## Penomoran BAST

Nomor otomatis tersimpan di `Database/nomor_bast.json`.

Contoh isi:

```json
{
  "last_number": 38,
  "entries": {
    "2026-11-03|PG ARINA MANASIKANA GANDU": 20,
    "2026-11-03|PG PAS AL QOSIM": 21
  }
}
```

Cara kerja:

- `last_number` = nomor tertinggi yang pernah dipakai.
- Saran nomor = `max(last_number, semua entries) + 1`.
- Nomor baru **langsung disimpan** setiap kali dibuat.
- Kunci `entries`: `YYYY-MM-DD|NAMA LEMBAGA`.
- Lembaga + tanggal yang sudah ada akan memakai nomor lamanya.
- Urutan pemberian nomor mengikuti urutan baris di `Database.xlsx`.

Contoh: mulai dari `020` untuk 19 lokasi → `020` s.d. `038`; run berikutnya disarankan `039`.

### Reset nomor

- Hapus `Database/nomor_bast.json` → mulai dari `001`.
- Atau edit `"last_number"` sesuai kebutuhan.

---

## Database

File: `Database/Database.xlsx`

### Sheet `Sheet1`

Kolom yang dipakai aplikasi:

| Kolom | Keterangan |
|---|---|
| `Layanan` | `Sekolah` atau `Posyandu` |
| `Nama Sekolah` | Nama sekolah / posyandu |
| `TK/PAUD/SD 1-3/Balita` | Jumlah balita / TK-PAUD-SD 1-3 |
| `SMP/MTs/SD 4-5` | Jumlah siswa besar (SD 4-6, SMP, SMA) |
| `BUMIL` | Jumlah ibu hamil (posyandu) |
| `BUSUI` | Jumlah ibu menyusui (posyandu) |
| `MPASI` | Jumlah MPASI (posyandu) |
| `Staff Sekolah` | Jumlah guru/staff (sekolah) |
| `Sampel MPASI` | Sampel MPASI |
| `Sampel BUMIL BUSUI` | Sampel bumil + busui |
| `Sampel Balita` | Sampel balita |
| `Sampel Besar` | Sampel porsi besar (sekolah) |
| `Sampel Kecil` | Sampel porsi kecil (sekolah) |
| `Porsi Kecil` | Total porsi kecil |
| `Porsi Besar` | Total porsi besar |
| `Jumlah` | Total porsi (dipakai di BAST/Surat Jalan) |

### Sheet `Menu`

Berisi daftar nama menu (referensi). Aplikasi **tidak** mengambil menu otomatis dari sheet ini; menu tetap diketik manual saat Langkah 4.

### Tips edit database

- Tutup Excel sebelum menjalankan aplikasi.
- Jangan mengubah nama header kolom.
- Baris dengan `Nama Sekolah` kosong akan diabaikan.

---

## Template

File di `Template/` berisi placeholder yang diganti otomatis, misalnya:

- `«Hari»`, `«Tanggall»`, `«Nama_Sekolah»`, `«Layanan»`
- `«TKPAUDSD_13Balita»`, `«SMPMTsSD_45»`, `«Staff_Sekolah»`
- `«Sampel_Kecil»`, `«Sampel_Besar»`, `«Jumlah»`
- `<<Nomor>>`, `<<Bulan_Romawi>>`, `<<Tahun>>`

`BAST.docx` berisi 2 halaman: **Penerimaan** dan **Pengembalian**.

Untuk mengubah tampilan surat, edit file `.docx` di folder `Template/` (jangan hapus placeholder-nya).

---

## Backup

- `Database/Database_backup.xlsx` — backup database.
- `Template/BAST_backup.docx` — backup template BAST.

---

## Troubleshooting

| Masalah | Solusi |
|---|---|
| `python` tidak dikenali | Install Python 3, centang **Add Python to PATH** |
| Dependensi gagal di-install | Cek internet, lalu jalankan `python -m pip install openpyxl python-docx docxcompose` |
| Database/template tidak ditemukan | Pastikan folder `Database/` dan `Template/` ada di samping `main.py` |
| Gagal menyimpan hasil | Tutup file hasil yang sedang dibuka di Word |
| Tanggal tidak valid | Gunakan format `dd/mm/yyyy`, contoh `03/11/2026` |
| Menu terasa lambat | Pastikan dijalankan di terminal Windows modern (bukan IDE) |
| Nomor tidak sesuai | Cek/edit `Database/nomor_bast.json` |

---

## Catatan

- Surat Jalan dan Organoleptik **tidak** memakai nomor BAST.
- Nomor BAST hanya dibuat saat dokumen BAST diproses.
- Aplikasi berjalan relatif terhadap folder tempat `main.py` berada (dijamin oleh `Buat_Surat.bat`).
