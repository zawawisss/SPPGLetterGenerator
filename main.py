import os
import sys
import json
import subprocess


def _pastikan_dependensi(paket=None):
    """Cek paket pihak ketiga; install otomatis kalau belum ada."""
    if paket is None:
        paket = {
            'openpyxl': 'openpyxl',
            'docx': 'python-docx',
            'docxcompose': 'docxcompose',
        }
    kurang = []
    for modul, nama_pip in paket.items():
        try:
            __import__(modul)
        except ImportError:
            kurang.append(nama_pip)

    if not kurang:
        return

    print("Dependensi belum lengkap: " + ", ".join(kurang))
    print("Menginstall otomatis, mohon tunggu...")
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *kurang])
    except Exception as e:
        print(f"Gagal install otomatis: {e}")
        print("Jalankan manual: pip install openpyxl python-docx docxcompose")
        input("Tekan Enter untuk keluar...")
        sys.exit(1)

    for modul in paket:
        try:
            __import__(modul)
        except ImportError:
            print(f"Paket {modul} masih belum bisa dimuat setelah install.")
            print("Cek koneksi internet / izin pip, lalu jalankan lagi.")
            input("Tekan Enter untuk keluar...")
            sys.exit(1)


_pastikan_dependensi()

import openpyxl
from docx import Document
from docx.shared import Pt
from datetime import datetime, timedelta
import docx
import docx.oxml
from docxcompose.composer import Composer

# --- Cross-Platform Keystroke Reader ---
try:
    import msvcrt
    is_windows = True
except ImportError:
    import tty
    import termios
    is_windows = False

# --- CONFIGURATION ---
DB_PATH = './Database/Database.xlsx'
TEMPLATE_SJ_SEKOLAH = './Template/SURAT JALAN SEKOLAH.docx'
TEMPLATE_SJ_POSYANDU = './Template/SURAT JALAN POSYANDU.docx'
TEMPLATE_BAST = './Template/BAST.docx'
TEMPLATE_ORGANOLEPTIK = './Template/ORGANOLEPTIK.docx'
OUTPUT_DIR = './HASIL'
NOMOR_BAST_FILE = './Database/nomor_bast.json'

# Mapping Hari Bahasa Indonesia
HARI_INDO = {
    'Monday': 'Senin',
    'Tuesday': 'Selasa',
    'Wednesday': 'Rabu',
    'Thursday': 'Kamis',
    'Friday': 'Jumat',
    'Saturday': 'Sabtu',
    'Sunday': 'Minggu'
}

# Mapping Bulan Romawi
BULAN_ROMAWI = {
    1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI',
    7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X', 11: 'XI', 12: 'XII'
}

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

def get_db_data():
    wb = openpyxl.load_workbook(DB_PATH, data_only=True)
    if 'Sheet1' not in wb.sheetnames:
        raise ValueError("Sheet 'Sheet1' tidak ditemukan di Database.xlsx")
    data_ws = wb['Sheet1']
    deliveries = []
    headers = [str(cell.value).strip() if cell.value else f"Col{i}" for i, cell in enumerate(data_ws[1])]
    
    for row in data_ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 4 or not row[3]: continue 
        item = {headers[i]: row[i] for i in range(min(len(row), len(headers)))}
        deliveries.append(item)
    return deliveries

def _load_nomor_bast():
    if os.path.exists(NOMOR_BAST_FILE):
        try:
            with open(NOMOR_BAST_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            data = {}
    else:
        data = {}
    if not isinstance(data, dict):
        data = {}
    data.setdefault('last_number', 0)
    data.setdefault('entries', {})
    return data

def _save_nomor_bast(data):
    with open(NOMOR_BAST_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_last_nomor():
    """Nomor BAST tertinggi yang pernah tercatat."""
    data = _load_nomor_bast()
    last = int(data.get('last_number', 0) or 0)
    for v in data.get('entries', {}).values():
        try:
            last = max(last, int(v))
        except (TypeError, ValueError):
            pass
    return last


def input_nomor_awal():
    """Tanya nomor awal; Enter memakai saran nomor terakhir + 1."""
    last = get_last_nomor()
    saran = last + 1
    while True:
        last_txt = f"{last:03d}" if last > 0 else "-"
        print_banner()
        print(f"  {Color.BOLD}{Color.BLUE}LANGKAH 2: NOMOR SURAT BAST{Color.END}")
        print(f"  {'-' * 25}\n")
        print(f"  Nomor terakhir tercatat : {Color.BOLD}{last_txt}{Color.END}")
        print(f"  Saran nomor berikutnya  : {Color.GREEN}{Color.BOLD}{saran:03d}{Color.END}\n")
        raw = input(f"  {Color.YELLOW}Mulai nomor berapa? [Enter = {saran:03d}]: {Color.END}").strip()
        if raw == '':
            return saran
        try:
            n = int(raw)
            if n < 1:
                raise ValueError
        except ValueError:
            print(f"  {Color.RED}Nomor tidak valid. Masukkan angka bulat >= 1.{Color.END}")
            print(f"  {Color.DARKCYAN}Tekan Enter untuk mencoba lagi...{Color.END}")
            input()
            continue
        if n <= last:
            print(f"\n  {Color.YELLOW}Peringatan: nomor {n:03d} sudah pernah dipakai (terakhir {last:03d}).{Color.END}")
            yn = input(f"  Tetap lanjut dengan nomor {n:03d}? (y/n): ").strip().lower()
            if yn != 'y':
                continue
        return n


def get_nomor_bast(dt_obj, nama_lembaga, counter=None):
    """Nomor BAST per lembaga per tanggal; urutan lanjut terus (tidak reset).

    counter (list satu elemen) dipakai untuk memaksa nomor awal pada batch ini;
    entri lama tetap memakai nomor yang sudah tercatat.
    """
    data = _load_nomor_bast()
    key = f"{dt_obj.strftime('%Y-%m-%d')}|{str(nama_lembaga).strip().upper()}"
    seq = data['entries'].get(key)
    if seq is None:
        if counter is not None and counter[0] is not None:
            seq = int(counter[0])
            counter[0] = seq + 1
        else:
            seq = int(data.get('last_number', 0)) + 1
        data['last_number'] = max(int(data.get('last_number', 0) or 0), int(seq))
        data['entries'][key] = seq
        _save_nomor_bast(data)
    return {
        '<<Nomor>>': f"{int(seq):03d}",
        '<<Bulan_Romawi>>': BULAN_ROMAWI[dt_obj.month],
        '<<Tahun>>': str(dt_obj.year),
    }

def set_font_times_new_roman(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:ascii'), 'Times New Roman')
            run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:hAnsi'), 'Times New Roman')
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:ascii'), 'Times New Roman')
                        run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:hAnsi'), 'Times New Roman')

def replace_tags(doc, mapping):
    def robust_replace(parent):
        for para in parent.paragraphs:
            for tag, val in mapping.items():
                if tag in para.text:
                    full_text = para.text.replace(tag, str(val))
                    para.text = full_text

    robust_replace(doc)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                robust_replace(cell)

def process_organoleptik(entry, date_str, day_name, menus):
    doc = Document(TEMPLATE_ORGANOLEPTIK)
    svc_type = entry.get('Layanan', 'Sekolah')
    mapping = {
        '«Hari»': day_name,
        '«Tanggall»': date_str,
        '«Nama_Sekolah»': entry['Nama Sekolah'],
        '«Layanan»': svc_type
    }
    replace_tags(doc, mapping)
    
    if doc.tables:
        table = doc.tables[0]
        menu_divider_idx = -1
        for i, row in enumerate(table.rows):
            if len(row.cells) > 1 and row.cells[1].text.strip() == 'Menu':
                menu_divider_idx = i
                break
        
        if menu_divider_idx != -1:
            current_row_idx = menu_divider_idx + 1
            
            for i, menu_name in enumerate(menus):
                if current_row_idx < len(table.rows):
                    row = table.rows[current_row_idx]
                else:
                    row = table.add_row()
                
                row.cells[0].text = str(i + 1)
                row.cells[1].text = menu_name
                
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if not para.runs:
                            para.add_run()
                        for run in para.runs:
                            run.font.size = Pt(10)
                            run.font.name = 'Times New Roman'
                            run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:ascii'), 'Times New Roman')
                            run._element.rPr.rFonts.set(docx.oxml.ns.qn('w:hAnsi'), 'Times New Roman')
                current_row_idx += 1
            
            while len(table.rows) > current_row_idx:
                row_to_remove = table.rows[current_row_idx]
                table._element.remove(row_to_remove._element)

    set_font_times_new_roman(doc)
    return doc

def process_bast(entry, date_str, day_name, total_qty, nomor_parts):
    doc = Document(TEMPLATE_BAST)
    mapping = {
        '«Hari»': day_name,
        '«Tanggall»': date_str,
        '«Nama_Sekolah»': entry['Nama Sekolah'],
        '«Layanan»': entry.get('Layanan', 'Sekolah'),
        '«Jumlah»': total_qty
    }
    mapping.update(nomor_parts)
    replace_tags(doc, mapping)
    set_font_times_new_roman(doc)
    return doc

def process_sj(entry, date_str, day_name, porsi_mapping):
    is_school = entry.get('Layanan') == 'Sekolah'
    template = TEMPLATE_SJ_SEKOLAH if is_school else TEMPLATE_SJ_POSYANDU
    doc = Document(template)
    
    mapping = {
        '«Hari»': day_name,
        '«Tanggall»': date_str,
        '«Nama_Sekolah»': entry['Nama Sekolah'],
        '«Layanan»': entry.get('Layanan', 'Sekolah'),
        '«Jumlah»': porsi_mapping.get('«Jumlah»', 0)
    }
    mapping.update(porsi_mapping)
    
    replace_tags(doc, mapping)
    set_font_times_new_roman(doc)
    return doc

# --- COLORS ---
class Color:
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'


def _enable_ansi():
    """Aktifkan dukungan escape ANSI di konsol Windows."""
    if os.name != 'nt':
        return True
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            return bool(kernel32.SetConsoleMode(handle, mode.value | 0x0004))
    except Exception:
        pass
    return False


ANSI_OK = _enable_ansi()


def clear_screen():
    if ANSI_OK:
        sys.stdout.write('\033[H\033[2J')
        sys.stdout.flush()
    else:
        os.system('cls' if os.name == 'nt' else 'clear')


def get_key():
    """Membaca input keyboard (cross-platform untuk Windows & Unix)"""
    if is_windows:
        key = msvcrt.getch()
        if key in (b'\x00', b'\xe0'):  
            key = msvcrt.getch()
            if key == b'H': return 'up'
            if key == b'P': return 'down'
            return 'unknown'
        elif key in (b'\r', b'\n'): 
            return 'enter'
        elif key == b'\x03': 
            return 'ctrl_c'
        return key.decode('utf-8', 'ignore')
    else:
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(sys.stdin.fileno())
            ch = sys.stdin.read(1)
            if ch == '\x1b':
                ch2 = sys.stdin.read(1)
                if ch2 == '[':
                    ch3 = sys.stdin.read(1)
                    if ch3 == 'A': return 'up'
                    if ch3 == 'B': return 'down'
            elif ch in ('\r', '\n'): 
                return 'enter'
            elif ch == '\x03': 
                return 'ctrl_c'
            return ch
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)

def _option_lines(options, selected_idx):
    lines = []
    for i, option in enumerate(options):
        if i == selected_idx:
            lines.append(f"    {Color.YELLOW}{Color.BOLD}> {option.upper()}{Color.END}")
        else:
            lines.append(f"      {option}")
    return lines


def pick_option(options, title):
    """Navigasi menu; hanya blok pilihan yang digambar ulang (tanpa cls tiap tombol)."""
    selected_idx = 0

    if not ANSI_OK:
        while True:
            print_banner()
            print(f"  {Color.BOLD}{Color.BLUE}{title.upper()}{Color.END}")
            print(f"  {'-' * len(title)}")
            print()
            for line in _option_lines(options, selected_idx):
                print(line)
            print()
            print(f"  {Color.DARKCYAN}[ Navigasi: Panah Atas/Bawah | Pilih: Enter ]{Color.END}")
            key = get_key()
            if key == 'up':
                selected_idx = (selected_idx - 1) % len(options)
            elif key == 'down':
                selected_idx = (selected_idx + 1) % len(options)
            elif key == 'enter':
                return selected_idx
            elif key == 'ctrl_c':
                print(f"\n  {Color.RED}Proses dibatalkan oleh pengguna.{Color.END}")
                sys.exit()
        return

    print_banner()
    print(f"  {Color.BOLD}{Color.BLUE}{title.upper()}{Color.END}")
    print(f"  {'-' * len(title)}")
    print()
    for line in _option_lines(options, selected_idx):
        print(line)
    print()
    print(f"  {Color.DARKCYAN}[ Navigasi: Panah Atas/Bawah | Pilih: Enter ]{Color.END}", flush=True)

    block = len(options) + 2
    while True:
        key = get_key()
        changed = False
        if key == 'up':
            selected_idx = (selected_idx - 1) % len(options)
            changed = True
        elif key == 'down':
            selected_idx = (selected_idx + 1) % len(options)
            changed = True
        elif key == 'enter':
            return selected_idx
        elif key == 'ctrl_c':
            print(f"\n  {Color.RED}Proses dibatalkan oleh pengguna.{Color.END}")
            sys.exit()

        if not changed:
            continue

        sys.stdout.write(f"\033[{block}A")
        for line in _option_lines(options, selected_idx):
            sys.stdout.write("\033[2K" + line + "\n")
        sys.stdout.write("\033[2K\n")
        sys.stdout.write(f"\033[2K  {Color.DARKCYAN}[ Navigasi: Panah Atas/Bawah | Pilih: Enter ]{Color.END}\n")
        sys.stdout.flush()

def print_banner():
    clear_screen()
    width = 60
    border = "=" * width
    print(f"\n{Color.CYAN}{border}")
    print(f"  {Color.BOLD}{Color.YELLOW}GENERATOR DOKUMEN SPPG JORESAN{Color.END}")
    print(f"  {Color.DARKCYAN}MLARAK, PONOROGO - PROFESSIONAL EDITION{Color.END}")
    print(f"{Color.CYAN}{border}{Color.END}\n")

def add_to_basket(basket, entry, porsi, menus):
    """Menambahkan data ke keranjang dengan penggabungan jika lokasi DAN menu sama."""
    name = entry['Nama Sekolah']
    
    # Cari apakah lokasi dengan menu yang sama sudah ada di keranjang
    for existing_entry, existing_porsi, existing_menus in basket:
        # Syarat merge: Nama Sekolah sama DAN List Menu identik
        if existing_entry['Nama Sekolah'] == name and existing_menus == menus:
            # Merge Porsi (Jumlahkan nilai numerik)
            for k, v in porsi.items():
                if k in existing_porsi:
                    try:
                        curr_val = int(existing_porsi[k] or 0)
                        new_val = int(v or 0)
                        existing_porsi[k] = curr_val + new_val
                    except (ValueError, TypeError):
                        pass
                else:
                    existing_porsi[k] = v
            return

    # Jika menu berbeda atau lokasi baru, tambahkan sebagai entri terpisah
    basket.append((entry.copy(), porsi.copy(), list(menus))) # Gunakan list(menus) untuk copy

def show_summary(basket):
    """Menampilkan ringkasan isi keranjang (Satpam)."""
    print_banner()
    print(f"  {Color.BOLD}{Color.PURPLE}SATPAM: KONFIRMASI DAFTAR CETAK{Color.END}")
    print(f"  {'-' * 35}")
    
    if not basket:
        print(f"  {Color.RED}Keranjang masih kosong!{Color.END}")
        return False

    # Kelompokkan berdasarkan layanan untuk tampilan lebih rapi
    summary = {}
    for entry, porsi, menus in basket:
        layanan = entry.get('Layanan', 'Lainnya')
        if layanan not in summary:
            summary[layanan] = []
        summary[layanan].append(entry['Nama Sekolah'])

    total = 0
    for layanan, sekolah_list in summary.items():
        print(f"\n  {Color.BOLD}{Color.BLUE}[ {layanan.upper()} ]{Color.END}")
        for i, nama in enumerate(sekolah_list, 1):
            print(f"    {i}. {nama}")
            total += 1
    
    print(f"\n  {'-' * 35}")
    print(f"  {Color.BOLD}TOTAL: {total} Lokasi{Color.END}")
    print(f"\n  {Color.YELLOW}Tekan [ENTER] untuk lanjut CETAK, atau [K] untuk kembali ke menu.{Color.END}")
    
    while True:
        key = get_key()
        if key == 'enter':
            return True
        elif key.lower() == 'k':
            return False
        elif key == 'ctrl_c':
            print(f"\n  {Color.RED}Proses dibatalkan oleh pengguna.{Color.END}")
            sys.exit()

def main():
    while True:
        # 1. Input Tanggal (Smart Picker)
        today = datetime.now()
        tomorrow = today + timedelta(days=1)
        day_after = today + timedelta(days=2)

        # Mendapatkan nama hari bahasa Indonesia untuk menu
        today_day = HARI_INDO.get(today.strftime("%A"), today.strftime("%A"))
        tomorrow_day = HARI_INDO.get(tomorrow.strftime("%A"), tomorrow.strftime("%A"))
        day_after_day = HARI_INDO.get(day_after.strftime("%A"), day_after.strftime("%A"))

        date_options = [
            f"HARI INI ({today_day}, {today.strftime('%d/%m/%Y')})",
            f"BESOK    ({tomorrow_day}, {tomorrow.strftime('%d/%m/%Y')})",
            f"LUSA     ({day_after_day}, {day_after.strftime('%d/%m/%Y')})",
            f"INPUT TANGGAL MANUAL"
        ]
        
        date_idx = pick_option(date_options, "Langkah 1: Pilih Tanggal Pengiriman")

        try:
            if date_idx == 0:
                dt_obj = today
            elif date_idx == 1:
                dt_obj = tomorrow
            elif date_idx == 2:
                dt_obj = day_after
            elif date_idx == 3:
                print_banner()
                print(f"  {Color.BOLD}{Color.BLUE}LANGKAH 1: INPUT MANUAL{Color.END}")
                print(f"  {'-' * 25}\n")
                input_date = input(f"  {Color.YELLOW}Masukkan Tanggal (dd/mm/yyyy): {Color.END}")
                dt_obj = datetime.strptime(input_date, "%d/%m/%Y")

            date_str = dt_obj.strftime("%d/%m/%Y")
            date_file = f"{dt_obj.day}-{dt_obj.month}-{dt_obj.strftime('%y')}"
            day_en = dt_obj.strftime("%A")
            day_name = HARI_INDO.get(day_en, day_en)
        except (ValueError, UnboundLocalError):
            print(f"\n  {Color.RED}ERROR: Format tanggal tidak valid!{Color.END}\n")
            print(f"  {Color.DARKCYAN}Tekan tombol apapun untuk kembali ke menu...{Color.END}")
            get_key()
            continue 

        nomor_awal = input_nomor_awal()
        nomor_counter = [nomor_awal]

        # --- SISTEM KERANJANG ---
        basket = [] # List of (entry, porsi, menus)
        
        while True:
            # 2. Input Layanan
            svc_options = [
                "SEKOLAH",
                "POSYANDU BALITA",
                "POSYANDU B2 (BUMIL & BUSUI)",
                "POSYANDU SEMUA (BALITA & B2)",
                "SEMUA LAYANAN (SEKOLAH + POSYANDU)",
                f"{Color.GREEN}{Color.BOLD}>>> SELESAI & CETAK DOKUMEN <<<{Color.END}"
            ]
            
            title_prefix = f"[ {day_name}, {date_str} ] - Keranjang: {len(basket)} Lokasi"
            svc_idx = pick_option(svc_options, f"{title_prefix}\n  Langkah 3: Pilih Kategori Layanan")
            
            # Jika user pilih "Selesai & Cetak"
            if svc_idx == 5:
                if not basket:
                    print(f"\n  {Color.RED}Peringatan: Keranjang masih kosong! Tambahkan minimal 1 kategori.{Color.END}")
                    print(f"  {Color.DARKCYAN}Tekan tombol apapun untuk melanjutkan...{Color.END}")
                    get_key()
                    continue
                break

            svc_choice = str(svc_idx + 1)
            service_names = {
                '1': 'SEKOLAH', 
                '2': 'POSYANDU BALITA', 
                '3': 'POSYANDU B2', 
                '4': 'POSYANDU SEMUA',
                '5': 'SEMUA LAYANAN'
            }
            
            # 3. Input Menu
            print_banner()
            print(f"  {Color.BOLD}{Color.GREEN}KATEGORI TERPILIH:{Color.END} {service_names[svc_choice]}")
            print(f"  {'-' * 30}\n")
            
            print(f"  {Color.BOLD}{Color.BLUE}LANGKAH 4: INPUT DAFTAR MENU{Color.END}")
            print(f"  {Color.DARKCYAN}(Ketik 'selesai' jika sudah cukup){Color.END}\n")

            current_menus = []
            while True:
                m = input(f"  {Color.BOLD}Menu #{len(current_menus)+1}:{Color.END} ")
                if m.lower() == 'selesai': 
                    if not current_menus:
                        print(f"  {Color.YELLOW}Peringatan: Minimal masukkan 1 menu!{Color.END}")
                        continue
                    break
                if m.strip(): 
                    current_menus.append(m)
                else:
                    print(f"  {Color.RED}Peringatan: Nama menu tidak boleh kosong!{Color.END}")

            print(f"\n  {Color.CYAN}Sedang mencari data di Database.xlsx...{Color.END}")
            
            try:
                all_data = get_db_data()
            except Exception as e:
                print(f"  {Color.RED}ERROR: Gagal memuat file database: {e}{Color.END}")
                get_key()
                continue

            added_count = 0
            for d in all_data:
                layanan = d.get('Layanan')
                
                def get_porsi_balita():
                    qty = int(d.get('TK/PAUD/SD 1-3/Balita') or 0) + int(d.get('Sampel Balita') or 0)
                    if qty <= 0: return None
                    return {
                        '«TKPAUDSD_13Balita»': d.get('TK/PAUD/SD 1-3/Balita', 0),
                        '«BUMIL»': 0, '«BUSUI»': 0, '«MPASI»': 0,
                        '«Sampel_Balita»': d.get('Sampel Balita', 0),
                        '«Sampel_BUMIL_BUSUI»': 0, '«Sampel_MPASI»': 0,
                        '«Jumlah»': qty
                    }

                def get_porsi_b2():
                    qty = int(d.get('BUMIL') or 0) + int(d.get('BUSUI') or 0) + int(d.get('Sampel BUMIL BUSUI') or 0)
                    if qty <= 0: return None
                    return {
                        '«TKPAUDSD_13Balita»': 0,
                        '«BUMIL»': d.get('BUMIL', 0),
                        '«BUSUI»': d.get('BUSUI', 0),
                        '«MPASI»': 0,
                        '«Sampel_Balita»': 0,
                        '«Sampel_BUMIL_BUSUI»': d.get('Sampel BUMIL BUSUI', 0),
                        '«Sampel_MPASI»': 0,
                        '«Jumlah»': qty
                    }

                def get_porsi_sekolah():
                    return {
                        '«TKPAUDSD_13Balita»': d.get('TK/PAUD/SD 1-3/Balita', 0),
                        '«SMPMTsSD_45»': d.get('SMP/MTs/SD 4-5', 0),
                        '«Staff_Sekolah»': d.get('Staff Sekolah', 0),
                        '«Sampel_Kecil»': d.get('Sampel Kecil', 0),
                        '«Sampel_Besar»': d.get('Sampel Besar', 0),
                        '«Jumlah»': d.get('Jumlah', 0)
                    }

                # Logika filter dan masukkan ke keranjang menggunakan helper add_to_basket
                match_found = False
                if svc_choice == '1' and layanan == 'Sekolah':
                    add_to_basket(basket, d, get_porsi_sekolah(), current_menus)
                    match_found = True
                elif svc_choice == '2' and layanan == 'Posyandu':
                    p = get_porsi_balita()
                    if p: 
                        add_to_basket(basket, d, p, current_menus)
                        match_found = True
                elif svc_choice == '3' and layanan == 'Posyandu':
                    p = get_porsi_b2()
                    if p: 
                        add_to_basket(basket, d, p, current_menus)
                        match_found = True
                elif svc_choice == '4' and layanan == 'Posyandu':
                    p_balita = get_porsi_balita()
                    if p_balita: add_to_basket(basket, d, p_balita, current_menus)
                    p_b2 = get_porsi_b2()
                    if p_b2: add_to_basket(basket, d, p_b2, current_menus)
                    if p_balita or p_b2: match_found = True
                elif svc_choice == '5':
                    if layanan == 'Sekolah':
                        add_to_basket(basket, d, get_porsi_sekolah(), current_menus)
                        match_found = True
                    elif layanan == 'Posyandu':
                        p_balita = get_porsi_balita()
                        if p_balita: add_to_basket(basket, d, p_balita, current_menus)
                        p_b2 = get_porsi_b2()
                        if p_b2: add_to_basket(basket, d, p_b2, current_menus)
                        if p_balita or p_b2: match_found = True
                
                if match_found: added_count += 1

            print(f"  {Color.GREEN}Berhasil menambahkan {added_count} lokasi ke keranjang!{Color.END}")
            print(f"  {Color.DARKCYAN}Tekan tombol apapun untuk kembali...{Color.END}")
            get_key()

        # --- TAHAP KONFIRMASI (SATPAM) ---
        if not show_summary(basket):
            continue

        # --- PROSES PEMBUATAN DOKUMEN (TUKANG JILID) ---
        out_folder = os.path.join(OUTPUT_DIR, date_file)
        if not os.path.exists(out_folder): os.makedirs(out_folder)

        print(f"\n  {Color.YELLOW}Sedang menyusun dokumen gabungan...{Color.END}\n")
        
        for label, func in [('ORGANOLEPTIK', 'org'), ('BAST', 'bast'), ('SURAT_JALAN', 'sj')]:
            docs = []
            
            # --- KHUSUS BAST & SJ: Gabung Paksa per Lokasi ---
            if func in ['bast', 'sj']:
                temp_merged = {} # name -> (entry, combined_porsi)
                for entry, porsi, _ in basket:
                    name = entry['Nama Sekolah']
                    if name not in temp_merged:
                        temp_merged[name] = (entry, porsi.copy())
                    else:
                        target_porsi = temp_merged[name][1]
                        for k, v in porsi.items():
                            try:
                                curr = int(target_porsi.get(k, 0))
                                add = int(v or 0)
                                target_porsi[k] = curr + add
                            except (ValueError, TypeError):
                                pass
                
                # Proses dokumen dari data yang sudah digabung per lokasi
                for entry, porsi in temp_merged.values():
                    if func == 'bast': docs.append(process_bast(entry, date_str, day_name, porsi['«Jumlah»'], get_nomor_bast(dt_obj, entry['Nama Sekolah'], nomor_counter)))
                    elif func == 'sj': docs.append(process_sj(entry, date_str, day_name, porsi))
            
            else: # ORGANOLEPTIK: Tetap ikuti basket (pisah jika menu beda)
                for entry, porsi, menus_in_entry in basket:
                    docs.append(process_organoleptik(entry, date_str, day_name, menus_in_entry))
            
            if not docs: continue

            master = docs[0]
            composer = Composer(master)
            for i in range(1, len(docs)):
                master.add_page_break()
                composer.append(docs[i])
            
            # Simpan dengan nama file yang lebih bersih
            save_path = os.path.join(out_folder, f"{label} {date_file}.docx")
            composer.save(save_path)
            print(f"    - Berhasil menyusun {label}")

        print(f"\n{Color.BOLD}{Color.GREEN}PENGGABUNGAN DOKUMEN SELESAI!{Color.END}")
        print(f"{Color.CYAN}Lokasi Folder:{Color.END}")
        print(f"{Color.UNDERLINE}{os.path.abspath(out_folder)}{Color.END}\n")
        
        # --- OPSI KEMBALI / KELUAR ---
        end_options = [
            "KEMBALI KE MENU UTAMA (Buat Dokumen Lain)",
            "TUTUP APLIKASI"
        ]
        end_choice = pick_option(end_options, "Proses Selesai! Apa selanjutnya?")
        
        if end_choice == 1:
            print(f"\n  {Color.YELLOW}Terima kasih telah menggunakan aplikasi ini. Sampai jumpa!{Color.END}\n")
            break 


if __name__ == '__main__':
    main()