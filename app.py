"""
HRIS Monthly Attendance ETL (Streamlit)
---------------------------------------
Upload PDF "Monthly Attendance Report" HRIS (PT.GLOSTAR INDONESIA) -> parse -> QA -> Excel.

Run:
    pip install streamlit pdfplumber pandas openpyxl
    streamlit run hris_attendance_etl_app.py

Output workbook: README, Attendance_Daily, Employee_Monthly, Dim_Employee, Dim_Status, Dim_Period, QA_Checks

Kirim ke Google Sheets (untuk Looker Studio), setup sekali:
    pip install gspread google-auth
    1. Google Cloud: buat Service Account, aktifkan Google Sheets API, unduh key JSON.
    2. Buat Google Sheet kosong, Share ke email service account (Editor).
    3. .streamlit/secrets.toml:
         [gsheet]
         spreadsheet_id = "ID_SPREADSHEET"
         [gcp_service_account]
         type = "service_account"
         project_id = "..."
         private_key_id = "..."
         private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
         client_email = "...@....iam.gserviceaccount.com"
         client_id = "..."
         token_uri = "https://oauth2.googleapis.com/token"
Tiap bulan: upload PDF -> Proses -> Kirim. Periode yang sama otomatis DIGANTI (bukan dobel).
"""
import re
import io
import calendar
import datetime as dt
from collections import Counter

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
# Kolom ringkasan per karyawan di PDF (urut dari kiri ke kanan). Januari punya 18 kolom, Februari 19 (+SOR).
SUM_COLS = ['H', 'A', 'I', 'CT', 'S', 'SD', 'TL', 'IR', 'DL', 'CH', 'CHK', 'HD', 'PC', 'LC',
            'N', 'CL4', 'CL5', 'CL6', 'SOR']

STATUS = {  # kode: (deskripsi, grup)
    'N': ('Normal', 'Hadir'), 'TL': ('Terlambat', 'Hadir'), 'PC': ('Pulang Cepat', 'Hadir'),
    'DL': ('Dinas Luar', 'Hadir'),
    'CT': ('Cuti', 'Cuti'), 'CH': ('Cuti Hamil', 'Cuti'),
    'SD': ('Surat Dokter', 'Sakit'), 'S': ('Sakit', 'Sakit'),
    'I': ('Izin', 'Izin'), 'IR': ('Izin Resmi', 'Izin'),
    'A': ('Alpa', 'Alpa'), 'H': ('Hari Libur', 'Libur'),
    'OFF': ('Weekend (tidak ada kode di PDF)', 'Weekend'),
    'NODATA': ('Hari kerja tanpa kode di PDF', 'No Data'),
}
# Koreksi manual untuk hari kerja yang kosong di PDF HRIS (sudah dikonfirmasi user).
# Format: (NIK, 'YYYY-MM-DD'): status. Hanya berlaku jika sel di PDF memang kosong.
MANUAL_OVERRIDES = {
    ('AD25722', '2025-06-05'): 'DL',   # Ahmad Ilham Fauzi, hari ke-4 dinas luar (Juni 2025)
    ('AD32820', '2025-05-22'): 'DL',   # Farhan Nandyawan Syarief, hari ke-4 dinas luar (Mei 2025)
}
MONTH_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des']
RAW_ALIAS = {'I*': 'IR'}  # di ringkasan HRIS, I* dihitung sebagai IR



# ----------------------------------------------------------------------------
# Extract / Transform
# ----------------------------------------------------------------------------
def _near(a, b, tol=0.35):
    return abs(a - b) < tol


def _find_layout(pg):
    """Baca layout halaman dari baris header: ukuran font, posisi kolom tanggal, batas kolom ringkasan."""
    ws = pg.extract_words(x_tolerance=1.5, extra_attrs=['size'])
    # baris header tanggal: baris dengan angka 1..31 terbanyak
    groups = {}
    for w in ws:
        if w['text'].isdigit() and 1 <= int(w['text']) <= 31:
            groups.setdefault(round(w['top']), []).append(w)
    best = None
    for top, g in groups.items():
        vals = {int(w['text']) for w in g}
        if vals >= set(range(1, 32)) and (best is None or len(g) > len(best[1])):
            best = (top, g)
    if best is None:
        return None
    htop, g = best
    day_w = {}
    for w in sorted(g, key=lambda w: w['x0']):
        day_w.setdefault(int(w['text']), w)
    day_x = {d: (w['x0'] + w['x1']) / 2 for d, w in day_w.items()}
    data_size = day_w[1]['size']
    # header 'No' -> ukuran font kolom identitas & batas kolom nomor
    no_w = [w for w in ws if w['text'] == 'No' and abs(w['top'] - htop) < 8 and w['x0'] < 40]
    if not no_w:
        return None
    id_size, no_x1 = no_w[0]['size'], no_w[0]['x1']
    # header kolom pertama setelah tanggal 31 (H) -> batas antara kolom hari-31 dan kolom ringkasan/total
    after = [w for w in ws if abs(w['top'] - htop) < 3 and w['x0'] > day_w[31]['x1'] + 0.5]
    if not after:
        return None
    h_x0 = min(w['x0'] for w in after)
    split = (day_w[31]['x1'] + h_x0) / 2
    col_w = (day_x[31] - day_x[1]) / 30
    return dict(htop=htop, day_x=day_x, data_size=data_size, id_size=id_size,
                limit_no=no_x1 + 1.5, split=split, col_w=col_w)


def parse_pdf(file_obj):
    """Parse satu PDF HRIS (layout dibaca otomatis dari header tiap halaman)."""
    out, period, prev = [], None, None
    with pdfplumber.open(file_obj) as pdf:
        for pg in pdf.pages:
            m = re.search(r'Bulan\s*:\s*(\d{6})', pg.extract_text() or '')
            if m:
                period = m.group(1)
            L = _find_layout(pg)
            if L is None:           # halaman overflow / bukan halaman data
                continue
            day_x, ds, ids, split = L['day_x'], L['data_size'], L['id_size'], L['split']
            s = ds / 8.1
            chars = pg.chars
            id_chars = sorted([c for c in chars if _near(c['size'], ids)], key=lambda c: c['x0'])

            anchors = []
            for t in sorted({round(c['top'], 1) for c in id_chars
                             if c['x0'] < L['limit_no'] and c['text'].isdigit() and c['top'] > L['htop'] + 5}):
                if not anchors or t - anchors[-1] > 3:
                    anchors.append(t)

            data_l = pg.filter(lambda o: o.get('object_type') != 'char' or
                               (_near(o['size'], ds) and o['x1'] <= split)).extract_words(x_tolerance=1.5)
            data_r = pg.filter(lambda o: o.get('object_type') != 'char' or
                               (_near(o['size'], ds) and o['x1'] > split)).extract_words(x_tolerance=1.5)
            data = data_l + data_r
            x_min = day_x[1] - 0.8 * L['col_w']

            def read_ot(words):
                ot, tot = {}, None
                for w in words:
                    v = float(w['text'])
                    if w['x1'] > split:
                        tot = v
                    else:
                        d = min(day_x, key=lambda k: abs(day_x[k] - (w['x0'] + w['x1']) / 2))
                        ot[d] = v
                return ot, tot

            # baris lembur yang terpotong pergantian halaman (kode di halaman sebelumnya, lembur di awal halaman ini)
            if prev is not None and prev['ot_total'] is None and anchors:
                orphan = [w for w in data if L['htop'] + 5 < w['top'] < anchors[0] - 3 and w['x0'] >= x_min
                          and re.fullmatch(r'[\d.]+', w['text'])]
                if orphan:
                    prev['ot'], prev['ot_total'] = read_ot(orphan)

            for T in anchors:
                row = [c for c in id_chars if abs(c['top'] - T) < 2.5]
                ns = [c for c in row if c['text'].strip()]
                i = 0
                while i < len(ns) and ns[i]['text'].isdigit() and ns[i]['x0'] < L['limit_no']:
                    i += 1
                no = ''.join(c['text'] for c in ns[:i])
                nik_chars = ns[i:i + 7]
                nik = ''.join(c['text'] for c in nik_chars)
                rest = ns[i + 7:]
                name = ''
                if rest:
                    name_start = rest[0]['x0']
                    junk_until = name_start + 12 * (ids / 9.0)
                    last_x1 = nik_chars[-1]['x1']
                    for c in row:
                        if c['x0'] <= last_x1 - 0.1 and c['text'].strip():
                            continue
                        if c['x0'] < name_start:
                            continue
                        if not c['text'].strip():
                            if name and c['x0'] >= junk_until:
                                name += ' '
                            continue
                        name += c['text']
                name = re.sub(r'\s+', ' ', name).strip()

                line_a = [w for w in data if -3 <= w['top'] - T <= 4 and w['x0'] >= x_min]
                line_b = [w for w in data if 6 * s <= w['top'] - T <= 17 * s and w['x0'] >= x_min]
                codes = {}
                for w in line_a:
                    if w['x0'] < split and not re.fullmatch(r'[\d.]+', w['text']):
                        d = min(day_x, key=lambda k: abs(day_x[k] - (w['x0'] + w['x1']) / 2))
                        codes[d] = w['text']
                svals = [float(w['text']) for w in sorted((w for w in line_a if w['x0'] >= split),
                                                          key=lambda w: w['x0'])]
                ot, ot_total = read_ot(line_b)
                out.append(dict(Period=period, No=int(no) if no.isdigit() else 0, NIK=nik, Name=name,
                                svals=svals, ot_total=ot_total, codes=codes, ot=ot))
            prev = out[-1] if out else None
    return period, out


def transform(records):
    """Raw records -> (daily df, monthly df, qa-issue list)."""
    daily, summ, issues = [], [], []
    for r in records:
        y, m = int(r['Period'][:4]), int(r['Period'][4:])
        nd = calendar.monthrange(y, m)[1]
        sv = dict(zip(SUM_COLS, r['svals']))
        pkey, plabel, pdate = y * 100 + m, f'{MONTH_ID[m - 1]} {y}', dt.date(y, m, 1)

        # validasi 1: total lembur
        if r['ot_total'] is None:
            issues.append((r['Period'], r['NIK'], 'Total lembur tidak terbaca'))
        elif abs(sum(r['ot'].get(d, 0) for d in range(1, 32)) - r['ot_total']) > 0.5:   # <= 0.5 jam dianggap pembulatan HRIS
            issues.append((r['Period'], r['NIK'], f"Jumlah OT harian {sum(r['ot'].values())} != total PDF {r['ot_total']}"))
        # validasi 2: hitung kode vs ringkasan PDF
        cc = Counter(RAW_ALIAS.get(v, v) for v in r['codes'].values())
        for code in ['H', 'A', 'I', 'CT', 'S', 'SD', 'TL', 'IR', 'DL', 'CH', 'PC', 'N']:
            if cc.get(code, 0) != sv.get(code, 0):
                issues.append((r['Period'], r['NIK'], f'Hitungan {code}: harian={cc.get(code, 0)} vs PDF={sv.get(code)}'))

        for d in range(1, nd + 1):
            date = dt.date(y, m, d)
            raw = r['codes'].get(d)
            code = RAW_ALIAS.get(raw, raw)
            source = 'PDF'
            if code is None:
                manual = MANUAL_OVERRIDES.get((r['NIK'], date.isoformat()))
                if manual:
                    code, source = manual, 'Manual'
                else:
                    code = 'OFF' if date.weekday() >= 5 else 'NODATA'
            daily.append([int(r['Period']), date, date.strftime('%a'), r['NIK'], r['Name'], raw, code,
                          STATUS.get(code, ('', 'Unknown'))[1],
                          0 if code in ('OFF', 'H', 'NODATA') else 1,
                          r['ot'].get(d, 0.0), source, pkey, plabel, pdate])
        summ.append([int(r['Period']), r['No'], r['NIK'], r['Name']] +
                    [sv.get(c) for c in SUM_COLS] + [r['ot_total'], pkey, plabel, pdate])

    D = pd.DataFrame(daily, columns=['Period', 'Date', 'Day_Name', 'NIK', 'Name', 'Raw_Code', 'Status',
                                     'Status_Group', 'Is_Workday', 'OT_Hours', 'Status_Source',
                                     'Period_Key', 'Period_Label', 'Period_Date'])
    S = pd.DataFrame(summ, columns=['Period', 'No', 'NIK', 'Name'] + SUM_COLS +
                     ['OT_Total_Reported', 'Period_Key', 'Period_Label', 'Period_Date'])
    return D, S, issues


def read_team_file(upload):
    """Baca file tim (CSV/XLSX) dengan kolom NAMA, NIK, TIM -> df[NIK, Team, Nama_File]."""
    name = upload.name.lower()
    raw = pd.read_csv(upload, dtype=str, sep=None, engine='python') if name.endswith(('.csv', '.txt')) \
        else pd.read_excel(upload, dtype=str)
    raw.columns = [str(c).strip().upper() for c in raw.columns]
    alias = {'NIK': ['NIK', 'ID', 'NIP'], 'TEAM': ['TIM', 'TEAM', 'TIM KERJA'], 'NAMA': ['NAMA', 'NAME']}
    pick = {}
    for key, opts in alias.items():
        pick[key] = next((c for c in raw.columns if c in opts), None)
    if not pick['NIK'] or not pick['TEAM']:
        raise ValueError(f'Kolom wajib NIK dan TIM tidak ditemukan. Kolom terbaca: {list(raw.columns)}')
    t = pd.DataFrame({
        'NIK': raw[pick['NIK']].astype(str).str.strip().str.upper(),
        'Team': raw[pick['TEAM']].astype(str).str.strip(),
        'Nama_File': raw[pick['NAMA']].astype(str).str.strip() if pick['NAMA'] else '',
    })
    t = t[(t.NIK != '') & (t.NIK != 'NAN')]
    t['Team'] = t.Team.replace({'': 'Unmapped', 'nan': 'Unmapped', 'NaN': 'Unmapped'})
    return t


def apply_team(D, S, team):
    """Tambah kolom Team ke D (paling kanan) dan S (setelah Name). NIK tanpa tim -> 'Unmapped'.
    Return D, S, dict info (duplikat, tidak ada di PDF, tidak ada di file tim, nama beda)."""
    dup = team[team.NIK.duplicated(keep=False)]
    team1 = team.drop_duplicates('NIK', keep='first').set_index('NIK')
    D = D.copy(); S = S.copy()
    D['Team'] = D.NIK.map(team1.Team).fillna('Unmapped')
    S.insert(S.columns.get_loc('Name') + 1, 'Team', S.NIK.map(team1.Team).fillna('Unmapped'))

    pdf_nik = D.drop_duplicates('NIK').set_index('NIK').Name
    norm = lambda x: re.sub(r'[^A-Z]', '', str(x).upper())
    mism = []
    for nik, nm in pdf_nik.items():
        if nik in team1.index and team1.loc[nik, 'Nama_File'] and norm(team1.loc[nik, 'Nama_File']) != norm(nm):
            mism.append((nik, nm, team1.loc[nik, 'Nama_File']))
    info = dict(
        duplicates=dup.sort_values('NIK'),
        unmapped=pdf_nik[~pdf_nik.index.isin(team1.index)].reset_index().rename(columns={'Name': 'Name_PDF'}),
        not_in_pdf=team1[~team1.index.isin(pdf_nik.index)].reset_index()[['NIK', 'Nama_File', 'Team']],
        name_mismatch=pd.DataFrame(mism, columns=['NIK', 'Name_PDF', 'Nama_File_Tim']),
    )
    return D, S, info


def team_summary(D):
    """Ringkasan per Team per Period."""
    w = D[D.Is_Workday == 1]
    g = w.groupby(['Team', 'Period'])
    out = pd.DataFrame({
        'Karyawan': D.groupby(['Team', 'Period']).NIK.nunique(),
        'Hari_Kerja': g.size(),
        'Hadir': g.apply(lambda x: (x.Status_Group == 'Hadir').sum()),
        'Tidak_Hadir': g.apply(lambda x: x.Status_Group.isin(['Cuti', 'Sakit', 'Izin', 'Alpa']).sum()),
        'Terlambat': g.apply(lambda x: (x.Status == 'TL').sum()),
        'Jam_Lembur': D.groupby(['Team', 'Period']).OT_Hours.sum(),
    }).reset_index()
    out['Tingkat_Hadir'] = (out.Hadir / out.Hari_Kerja).round(4)
    return out


def filter_inactive(D, S, max_missing=1):
    """Hapus karyawan yang dianggap RESIGN: tidak muncul di N periode terakhir.

    Dihitung dari periode terbaru yang diupload, mundur ke belakang:
      bulan_hilang_berturut = jumlah periode terbaru yang tidak ada datanya (dihitung dari akhir).
    - Jika karyawan masih ada di periode terbaru -> 0 -> dipertahankan (bukan resign).
    - Jika bulan_hilang_berturut > N -> seluruh barisnya dihapus dari D dan S.
    Bulan kosong di tengah (misal cuti panjang lalu kembali) tidak dihitung.
    Return (D_bersih, S_bersih, df_dihapus).
    """
    periods = sorted(D.Period.unique())
    pos = {p: i for i, p in enumerate(periods)}
    last = D.groupby('NIK').Period.max().map(pos)
    trailing = (len(periods) - 1 - last).rename('Bulan_Tidak_Ada_Terakhir')
    drop_nik = trailing[trailing > max_missing].index
    info = (D[D.NIK.isin(drop_nik)].groupby(['NIK', 'Name'])
            .agg(Periode_Terakhir=('Period', 'max'),
                 Periode_Ada=('Period', lambda x: ', '.join(map(str, sorted(x.unique())))))
            .reset_index())
    info['Bulan_Tidak_Ada_Terakhir'] = info.NIK.map(trailing)
    info['Alasan'] = f'Tidak ada di {int(max_missing) + 1}+ periode terakhir (periode terbaru {periods[-1]})'
    info = info[['NIK', 'Name', 'Periode_Terakhir', 'Bulan_Tidak_Ada_Terakhir', 'Periode_Ada', 'Alasan']]
    return D[~D.NIK.isin(drop_nik)].reset_index(drop=True), S[~S.NIK.isin(drop_nik)].reset_index(drop=True), info


# ----------------------------------------------------------------------------
# Load -> Excel
# ----------------------------------------------------------------------------
def build_workbook(D, S, removed=None):
    F = 'Arial'
    hfont = Font(name=F, bold=True, color='FFFFFF', size=10)
    hfill = PatternFill('solid', fgColor='1F3864')
    bfont = Font(name=F, size=10)

    def header(ws, n):
        for c in range(1, n + 1):
            cell = ws.cell(row=1, column=c)
            cell.font, cell.fill = hfont, hfill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.freeze_panes = 'A2'

    def body(ws, n):
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=n):
            for c in row:
                c.font = bfont

    def widths(ws, w):
        for i, x in enumerate(w, 1):
            ws.column_dimensions[get_column_letter(i)].width = x

    wb = Workbook()
    ws = wb.active
    ws.title = 'README'
    periods = sorted(D.Period.unique())
    notes = [
        ('HRIS Monthly Attendance - hasil konsolidasi ETL', True),
        (f'Periode: {periods[0]} s/d {periods[-1]} | dibuat {dt.datetime.now():%Y-%m-%d %H:%M}', False),
        ('', False),
        ('Sheet', True),
        ('Attendance_Daily : 1 baris per karyawan per tanggal (tabel fakta untuk dashboard)', False),
        ('Employee_Monthly : ringkasan per karyawan per bulan (angka PDF) + kolom cek berformula', False),
        ('Dim_Employee / Dim_Status / Dim_Period : tabel dimensi. Period_Key = YYYYMM (angka, untuk sort di dashboard), Period_Label = teks tampilan, Period_Seq = urutan 1..n, Is_Latest = periode terbaru', False),
        ('Kolom Team (jika file tim diupload) : dipetakan dari NIK; NIK tanpa tim diberi Unmapped', False),
        ('Removed_Employees : karyawan yang dianggap resign (tidak ada di N periode terakhir) dan dihapus (jika filter aktif)', False),
        ('QA_Checks : rekap validasi', False),
        ('', False),
        ('Asumsi', True),
        ('1. Hanya hari kerja yang berkode di PDF; Sabtu/Minggu tanpa kode diisi Status = OFF.', False),
        ('2. Kode I* dinormalisasi ke IR (sesuai hitungan ringkasan HRIS); kode asli di Raw_Code.', False),
        ('3. Baris angka kedua per karyawan diasumsikan jam LEMBUR (OT_Hours). Check_OT toleransi 0,5 jam (pembulatan di HRIS).', False),
        ('3b. Koreksi manual (MANUAL_OVERRIDES di script) ditandai Status_Source = Manual; Raw_Code tetap kosong.', False),
        ('4. Hari kerja tanpa kode diberi Status NODATA (misal karyawan keluar/masuk di tengah bulan).', False),
        ('5. Kolom CHK, HD, LC, CL4, CL5, CL6, SOR tidak punya legenda di PDF; disimpan apa adanya.', False),
    ]
    for i, (t, b) in enumerate(notes, 1):
        ws.cell(row=i, column=1, value=t).font = Font(name=F, bold=b, size=12 if i == 1 else 10)
    ws.column_dimensions['A'].width = 120

    # Attendance_Daily
    ws = wb.create_sheet('Attendance_Daily')
    ws.append(list(D.columns))
    for r in D.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in r])
    nd = ws.max_row
    header(ws, D.shape[1]); body(ws, D.shape[1])
    for r in range(2, nd + 1):
        ws.cell(row=r, column=2).number_format = 'yyyy-mm-dd'
        ws.cell(row=r, column=10).number_format = '0.0'
    widths(ws, [10, 12, 10, 10, 30, 10, 9, 13, 11, 10, 14, 11, 12, 12] + ([16] if 'Team' in D.columns else []))
    pdc = D.columns.get_loc('Period_Date') + 1
    for r in range(2, nd + 1):
        ws.cell(row=r, column=pdc).number_format = 'yyyy-mm-dd'
    ws.auto_filter.ref = f'A1:{get_column_letter(D.shape[1])}{nd}'

    # Employee_Monthly (+ formula check)
    ws = wb.create_sheet('Employee_Monthly')
    hdr = list(S.columns) + ['N_Calc', 'OT_Calc', 'Check_N', 'Check_OT']
    ws.append(hdr)
    ncs = S.shape[1]
    cN = get_column_letter(list(S.columns).index('N') + 1)
    cOT = get_column_letter(list(S.columns).index('OT_Total_Reported') + 1)
    cNc, cOc = get_column_letter(ncs + 1), get_column_letter(ncs + 2)
    for i, r in enumerate(S.itertuples(index=False), 2):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in r] + [
            f'=COUNTIFS(Attendance_Daily!$A:$A,$A{i},Attendance_Daily!$D:$D,$C{i},Attendance_Daily!$G:$G,"N")',
            f'=SUMIFS(Attendance_Daily!$J:$J,Attendance_Daily!$A:$A,$A{i},Attendance_Daily!$D:$D,$C{i})',
            f'=IF({cNc}{i}={cN}{i},"OK","DIFF")',
            f'=IF(ABS({cOc}{i}-{cOT}{i})<=0.5,"OK","DIFF")'])
    ns = ws.max_row
    header(ws, len(hdr)); body(ws, len(hdr))
    spc = list(S.columns).index('Period_Date') + 1
    for r in range(2, ns + 1):
        ws.cell(row=r, column=spc).number_format = 'yyyy-mm-dd'
    widths(ws, [10, 6, 10, 30] + ([16] if 'Team' in S.columns else []) + [6] * len(SUM_COLS) + [12, 11, 12, 12, 9, 9, 9, 9])
    ws.auto_filter.ref = f'A1:{get_column_letter(len(hdr))}{ns}'
    ws.row_dimensions[1].height = 30

    # Dim_Employee
    ws = wb.create_sheet('Dim_Employee')
    ecols = ['NIK', 'Name'] + (['Team'] if 'Team' in D.columns else [])
    ws.append(ecols + ['Periods_Count'])
    for e in D.groupby(ecols).Period.nunique().reset_index().itertuples(index=False):
        ws.append(list(e))
    header(ws, len(ecols) + 1); body(ws, len(ecols) + 1); widths(ws, [10, 30, 16, 14][:len(ecols) + 1] if len(ecols) == 3 else [10, 30, 14])
    ne = ws.max_row

    # Dim_Period (urutan periode untuk dashboard)
    ws = wb.create_sheet('Dim_Period')
    ws.append(['Period_Key', 'Period_Label', 'Period_Date', 'Period_Seq', 'Is_Latest'])
    dp = D[['Period_Key', 'Period_Label', 'Period_Date']].drop_duplicates().sort_values('Period_Key').reset_index(drop=True)
    for i, r in enumerate(dp.itertuples(index=False), 1):
        ws.append([r.Period_Key, r.Period_Label, r.Period_Date, i, 1 if i == len(dp) else 0])
    header(ws, 5); body(ws, 5); widths(ws, [12, 14, 13, 12, 10])
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=3).number_format = 'yyyy-mm-dd'

    # Dim_Status
    ws = wb.create_sheet('Dim_Status')
    ws.append(['Status', 'Description', 'Status_Group'])
    for k, (d, g) in STATUS.items():
        ws.append([k, d, g])
    header(ws, 3); body(ws, 3); widths(ws, [10, 34, 14])

    # QA_Checks
    ws = wb.create_sheet('QA_Checks')
    ws.append(['Check', 'Result', 'Expected'])
    cCN, cCO = get_column_letter(ncs + 3), get_column_letter(ncs + 4)
    for q in [
        ['Baris Attendance_Daily', f'=COUNTA(Attendance_Daily!$A$2:$A${nd})', len(D)],
        ['Baris Employee_Monthly', f'=COUNTA(Employee_Monthly!$A$2:$A${ns})', len(S)],
        ['Karyawan unik', f'=COUNTA(Dim_Employee!$A$2:$A${ne})', D.NIK.nunique()],
        ['Check_N = DIFF', f'=COUNTIF(Employee_Monthly!${cCN}$2:${cCN}${ns},"DIFF")', 0],
        ['Check_OT = DIFF', f'=COUNTIF(Employee_Monthly!${cCO}$2:${cCO}${ns},"DIFF")', 0],
        ['Hari kerja NODATA', f'=COUNTIF(Attendance_Daily!$G$2:$G${nd},"NODATA")', int((D.Status == 'NODATA').sum())],
        ['Hari dikoreksi manual (Status_Source = Manual)', f'=COUNTIF(Attendance_Daily!$K$2:$K${nd},"Manual")', int((D.Status_Source == 'Manual').sum())],
    ] + ([['Baris tanpa tim (Unmapped)', f'=COUNTIF(Attendance_Daily!${get_column_letter(D.columns.get_loc("Team") + 1)}$2:${get_column_letter(D.columns.get_loc("Team") + 1)}${nd},"Unmapped")', int((D.Team == 'Unmapped').sum())]] if 'Team' in D.columns else []):
        ws.append(q)
    header(ws, 3); body(ws, 3); widths(ws, [30, 14, 14])

    if removed is not None and len(removed):
        ws = wb.create_sheet('Removed_Employees')
        ws.append(list(removed.columns))
        for r in removed.itertuples(index=False):
            ws.append(list(r))
        header(ws, removed.shape[1]); body(ws, removed.shape[1]); widths(ws, [10, 30, 16, 14, 40, 50])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Load -> Google Sheets (upsert per periode)
# ----------------------------------------------------------------------------
GS_TITLES = dict(daily='Attendance_Daily', monthly='Employee_Monthly', period='Dim_Period', team='Master_Team')
_GS_CHUNK = 5000


def _ws_missing(e):
    return e.__class__.__name__ == 'WorksheetNotFound'


def _cell(v, col=''):
    """Nilai python -> nilai sel Google Sheets (USER_ENTERED)."""
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)):
        return ''
    if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)):
        return v.strftime('%Y-%m-%d')
    if hasattr(v, 'item'):                 # numpy scalar
        v = v.item()
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float)):
        return v
    # "Jan 2025" akan dikira tanggal oleh Sheets; apostrof memaksa teks
    return "'" + str(v) if col == 'Period_Label' else str(v)


def _rows(df):
    return [[_cell(v, c) for v, c in zip(row, df.columns)] for row in df.itertuples(index=False, name=None)]


def period_info(key):
    y, m = divmod(int(key), 100)
    return dict(Period_Key=int(key), Period_Label=f'{MONTH_ID[m - 1]} {y}', Period_Date=dt.date(y, m, 1))


def gs_ready():
    try:
        import gspread  # noqa: F401
        from google.oauth2.service_account import Credentials  # noqa: F401
    except ImportError:
        return False, 'Library belum terpasang: pip install gspread google-auth'
    try:
        has = 'gcp_service_account' in st.secrets
    except Exception:
        has = False
    if not has:
        return False, 'Service account belum diatur di .streamlit/secrets.toml (kunci [gcp_service_account]).'
    return True, ''


def gs_default_id():
    try:
        return str(st.secrets['gsheet']['spreadsheet_id'])
    except Exception:
        return ''


def gs_open(ref):
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        dict(st.secrets['gcp_service_account']), scopes=['https://www.googleapis.com/auth/spreadsheets'])
    gc = gspread.authorize(creds)
    ref = ref.strip()
    return gc.open_by_url(ref) if ref.startswith('http') else gc.open_by_key(ref)


def gs_get_ws(sh, title, header):
    try:
        ws = sh.worksheet(title)
    except Exception as e:
        if not _ws_missing(e):
            raise
        ws = sh.add_worksheet(title=title, rows=2, cols=max(len(header), 2))
    if not ws.row_values(1):
        ws.update(range_name='A1', values=[header], raw=False)
    return ws


_DERIVED_PERIOD_COLS = ('Period_Key', 'Period_Label', 'Period_Date')


def _a1(row, col):
    letters = ''
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f'{letters}{row}'


def gs_migrate_period_cols(ws, df_cols):
    """Sheet lama tanpa Period_Key/Label/Date: tambahkan 3 kolom itu di kanan dan isi dari kolom Period.
    Hanya menambah kolom (tidak mengubah/menghapus data). Return True jika ada migrasi."""
    header = ws.row_values(1)
    missing = [c for c in _DERIVED_PERIOD_COLS if c in df_cols and c not in header]
    if not missing or 'Period' not in header or (set(df_cols) - set(header)) - set(_DERIVED_PERIOD_COLS):
        return False
    per = ws.col_values(header.index('Period') + 1)[1:]
    out = [missing]
    for v in per:
        try:
            info_ = period_info(int(float(v)))
        except (ValueError, TypeError, IndexError):
            out.append([''] * len(missing))
            continue
        cell = {'Period_Key': info_['Period_Key'], 'Period_Label': "'" + info_['Period_Label'],
                'Period_Date': info_['Period_Date'].strftime('%Y-%m-%d')}
        out.append([cell[c] for c in missing])
    ws.update(range_name=_a1(1, len(header) + 1), values=out, raw=False)
    return True


def gs_upsert(ws, df, key_col='Period_Key'):
    """Ganti semua baris dengan Period_Key yang ada di df, lalu tambahkan baris baru. Return (dihapus, ditambah)."""
    gs_migrate_period_cols(ws, set(df.columns))
    header = ws.row_values(1)
    if set(header) != set(df.columns):
        raise ValueError(f"Kolom sheet '{ws.title}' tidak sama dengan data. "
                         f"Hanya di sheet: {sorted(set(header) - set(df.columns))}; "
                         f"hanya di data: {sorted(set(df.columns) - set(header))}. "
                         "Samakan dulu (mis. upload file tim kalau sheet sudah punya kolom Team).")
    df = df[header]
    keys = {str(int(x)) for x in df[key_col].unique()}
    existing = ws.col_values(header.index(key_col) + 1)[1:]
    hit = [i + 2 for i, v in enumerate(existing) if str(v).strip() in keys]
    blocks = []                                   # blok baris berurutan, dihapus dari bawah ke atas
    for r in hit:
        if blocks and r == blocks[-1][1] + 1:
            blocks[-1][1] = r
        else:
            blocks.append([r, r])
    reqs = [{'deleteDimension': {'range': {'sheetId': ws.id, 'dimension': 'ROWS',
                                           'startIndex': a - 1, 'endIndex': b}}} for a, b in reversed(blocks)]
    if reqs:
        ws.spreadsheet.batch_update({'requests': reqs})
    rows = _rows(df)
    for i in range(0, len(rows), _GS_CHUNK):
        ws.append_rows(rows[i:i + _GS_CHUNK], value_input_option='USER_ENTERED')
    return len(hit), len(rows)


def gs_rebuild_period(sh, ws_daily):
    """Bangun ulang Dim_Period dari semua Period_Key yang ada di sheet Attendance_Daily."""
    header = ws_daily.row_values(1)
    keys = sorted({int(float(v)) for v in ws_daily.col_values(header.index('Period_Key') + 1)[1:] if str(v).strip()})
    hdr = ['Period_Key', 'Period_Label', 'Period_Date', 'Period_Seq', 'Is_Latest']
    ws = gs_get_ws(sh, GS_TITLES['period'], hdr)
    rows = []
    for i, k in enumerate(keys, 1):
        pi_ = period_info(k)
        rows.append([k, "'" + pi_['Period_Label'], pi_['Period_Date'].strftime('%Y-%m-%d'), i, 1 if i == len(keys) else 0])
    ws.clear()
    ws.update(range_name='A1', values=[hdr] + rows, raw=False)
    return len(rows)


def gs_merge_team(sh, team):
    """Gabungkan file tim ke sheet Master_Team (NIK sama -> ditimpa, NIK lain dipertahankan)."""
    hdr = ['NIK', 'NAMA', 'TIM']
    ws = gs_get_ws(sh, GS_TITLES['team'], hdr)
    cur = {}
    for r in ws.get_all_values()[1:]:
        if r and str(r[0]).strip():
            cur[str(r[0]).strip().upper()] = [str(r[0]).strip().upper(), r[1] if len(r) > 1 else '', r[2] if len(r) > 2 else '']
    for t in team.itertuples(index=False):
        old_name = cur.get(t.NIK, ['', '', ''])[1]
        cur[t.NIK] = [t.NIK, t.Nama_File if str(t.Nama_File).strip() else old_name, t.Team]
    ws.clear()
    ws.update(range_name='A1', values=[hdr] + [cur[k] for k in sorted(cur)], raw=False)
    return len(cur)


def gs_push(sh, D, S, team_df=None):
    ws_d = gs_get_ws(sh, GS_TITLES['daily'], list(D.columns))
    d_del, d_add = gs_upsert(ws_d, D)
    ws_m = gs_get_ws(sh, GS_TITLES['monthly'], list(S.columns))
    m_del, m_add = gs_upsert(ws_m, S)
    n_period = gs_rebuild_period(sh, ws_d)
    n_team = gs_merge_team(sh, team_df) if team_df is not None else None
    return dict(daily=(d_del, d_add), monthly=(m_del, m_add), periods=n_period, team=n_team)


@st.cache_data(ttl=120, show_spinner=False)
def gs_load_team(ref):
    sh = gs_open(ref)
    try:
        ws = sh.worksheet(GS_TITLES['team'])
    except Exception as e:
        if _ws_missing(e):
            return None
        raise
    vals = ws.get_all_values()
    if len(vals) < 2:
        return None
    df = pd.DataFrame(vals[1:], columns=[c.strip().upper() for c in vals[0]])
    return pd.DataFrame({'NIK': df['NIK'].str.strip().str.upper(),
                         'Team': df['TIM'].str.strip().replace('', 'Unmapped'),
                         'Nama_File': df['NAMA'].str.strip() if 'NAMA' in df else ''})


@st.cache_data(ttl=30, show_spinner=False)
def gs_existing_periods(ref):
    sh = gs_open(ref)
    try:
        ws = sh.worksheet(GS_TITLES['daily'])
    except Exception as e:
        if _ws_missing(e):
            return []
        raise
    header = ws.row_values(1)
    if 'Period_Key' not in header:
        return []
    return sorted({int(float(v)) for v in ws.col_values(header.index('Period_Key') + 1)[1:] if str(v).strip()})


def run_push(ref, D, S, team_df=None):
    """Kirim ke Google Sheets, bersihkan cache, kembalikan pesan ringkas."""
    res = gs_push(gs_open(ref), D, S, team_df)
    gs_existing_periods.clear()
    gs_load_team.clear()
    return (f"Attendance_Daily: {res['daily'][0]} baris lama diganti, {res['daily'][1]} baris ditulis. "
            f"Employee_Monthly: {res['monthly'][0]} diganti, {res['monthly'][1]} ditulis. "
            f"Dim_Period: {res['periods']} periode."
            + (f" Master_Team: {res['team']} NIK." if res['team'] is not None else ''))


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def run_etl(files):
    """files: tuple[(filename, bytes)] -> (D, S, issues, per_file_info)"""
    recs, info = [], []
    for fname, content in files:
        period, rows = parse_pdf(io.BytesIO(content))
        info.append(dict(File=fname, Period=period, Employees=len(rows)))
        recs += rows
    D, S, issues = transform(recs)
    return D, S, issues, info


def main():
    st.set_page_config(page_title='HRIS Attendance ETL', layout='wide')
    st.title('HRIS Attendance ETL')
    st.caption('PDF Monthly Attendance Report (PT.GLOSTAR INDONESIA) -> Excel siap dashboard')

    with st.sidebar:
        st.header('Master data tim')
        team_up = st.file_uploader('Upload master tim (kolom: NAMA, NIK, TIM) - CSV / XLSX',
                                   type=['csv', 'xlsx', 'xls'], key='team_file')
        st.caption('Opsional. Template berisi NAMA dan NIK dari PDF tersedia di sini setelah diproses.')
        tmpl_slot = st.container()
        st.header('Opsi')
        use_filter = st.checkbox('Hapus karyawan resign (tidak ada di N periode terakhir)', value=True)
        max_missing = st.number_input('N (bulan)', min_value=0, max_value=24, value=1, step=1)
        st.header('Google Sheets')
        sid = st.text_input('Spreadsheet URL / ID', value=gs_default_id(),
                            help='Sheet tujuan untuk Looker Studio. Bisa diisi default lewat secrets.toml.')
        auto_push = st.checkbox('Langsung kirim ke Google Sheets setelah klik Proses', value=True,
                                help='Dilewati otomatis jika ada selisih validasi; kirim manual setelah dicek.')
        use_gs_team = st.checkbox('Pakai Master_Team dari Google Sheets jika tidak upload file tim', value=True)

    up = st.file_uploader('Upload PDF HRIS (boleh banyak bulan sekaligus)', type='pdf', accept_multiple_files=True)
    files = tuple((f.name, f.getvalue()) for f in up) if up else ()

    # Tombol proses: selalu tampil (nonaktif sebelum ada PDF). ETL hanya jalan setelah diklik;
    # jika file/opsi berubah, harus klik lagi.
    sig = (tuple((n, len(b)) for n, b in files),
           (team_up.name, team_up.size) if team_up is not None else None,
           bool(use_filter), int(max_missing))
    if st.button('Proses', type='primary', disabled=not files):
        st.session_state['done_sig'] = sig
        st.session_state['push_pending'] = True
    if not files:
        st.info('Upload satu atau beberapa PDF, lalu klik **Proses**.')
        return
    if st.session_state.get('done_sig') != sig:
        if 'done_sig' in st.session_state:
            st.info('File atau opsi berubah. Klik **Proses** untuk menjalankan ulang.')
        else:
            st.info(f'{len(files)} file siap. Klik **Proses** untuk menjalankan ETL.')
        return

    try:
        with st.spinner('Memproses PDF...'):
            D, S, issues, info = run_etl(files)
    except Exception as e:
        st.error(f'Gagal memproses PDF: {e}')
        return
    if D.empty:
        st.warning('Tidak ada data karyawan yang terbaca. Pastikan format PDF sama dengan laporan HRIS.')
        return

    # Mapping tim (opsional) - uploader ada di sidebar
    st.subheader('Mapping tim')
    tmpl = D.drop_duplicates('NIK')[['Name', 'NIK']].rename(columns={'Name': 'NAMA'}).assign(TIM='')
    tmpl_slot.download_button('Download template tim (CSV)', tmpl.to_csv(index=False).encode('utf-8'),
                              'template_tim.csv', 'text/csv')
    team_df, team_from = None, None
    if team_up is not None:
        try:
            team_df, team_from = read_team_file(team_up), 'file'
        except Exception as e:
            st.error(f'Gagal membaca file tim: {e}')
    elif sid and use_gs_team and gs_ready()[0]:
        try:
            team_df = gs_load_team(sid)
            if team_df is not None:
                team_from = 'sheet'
                st.info(f'Memakai Master_Team dari Google Sheets ({len(team_df)} NIK).')
        except Exception as e:
            st.warning(f'Master_Team di Google Sheets tidak terbaca: {e}')
    if team_df is None:
        st.info('Belum ada master tim. Upload lewat sidebar kiri (Master data tim) untuk menambahkan kolom Team.')
    team_info = None
    if team_df is not None:
        try:
            D, S, team_info = apply_team(D, S, team_df)
        except Exception as e:
            st.error(f'Gagal menerapkan mapping tim: {e}')
            team_info = None
        if team_info is not None:
            n_un = len(team_info['unmapped'])
            if n_un:
                st.warning(f'{n_un} karyawan di PDF tidak ada di file tim (Team = Unmapped):')
                st.dataframe(team_info['unmapped'], use_container_width=True)
            else:
                st.success('Semua karyawan di PDF sudah punya tim.')
            if len(team_info['duplicates']):
                st.warning('NIK ganda di file tim (dipakai baris pertama):')
                st.dataframe(team_info['duplicates'], use_container_width=True)
            if len(team_info['not_in_pdf']):
                st.info(f"{len(team_info['not_in_pdf'])} NIK ada di file tim tapi tidak ada di PDF.")
                with st.expander('Lihat daftar'):
                    st.dataframe(team_info['not_in_pdf'], use_container_width=True)
            if len(team_info['name_mismatch']):
                st.info(f"{len(team_info['name_mismatch'])} NIK namanya berbeda antara PDF dan file tim (cek salah ketik / NIK tertukar).")
                with st.expander('Lihat daftar'):
                    st.dataframe(team_info['name_mismatch'], use_container_width=True)

    # Filter karyawan yang tidak ada lebih dari N bulan
    removed = pd.DataFrame()
    if use_filter and D.Period.nunique() > 1:
        D, S, removed = filter_inactive(D, S, int(max_missing))
        if len(removed):
            st.warning(f'{len(removed)} karyawan dihapus (tidak ada di {int(max_missing) + 1}+ periode terakhir, dianggap resign):')
            st.dataframe(removed, use_container_width=True)
        else:
            st.caption(f'Semua karyawan masih ada dalam {int(max_missing) + 1} periode terakhir.')

    dup = pd.DataFrame(info).Period.duplicated().any()
    if dup:
        st.warning('Ada lebih dari satu file dengan periode yang sama. Data akan terduplikasi.')

    # KPI
    w = D[D.Is_Workday == 1]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('Periode', D.Period.nunique())
    c2.metric('Karyawan unik', D.NIK.nunique())
    c3.metric('Baris harian', f'{len(D):,}')
    c4.metric('Tingkat hadir', f'{(w.Status_Group == "Hadir").mean():.1%}' if len(w) else '-')
    c5.metric('Total jam lembur', f'{D.OT_Hours.sum():,.1f}')

    # QA
    if issues:
        st.error(f'{len(issues)} selisih validasi terhadap angka ringkasan PDF')
        st.dataframe(pd.DataFrame(issues, columns=['Period', 'NIK', 'Issue']), use_container_width=True)
    else:
        st.success('Validasi lulus: hitungan kode status dan total lembur cocok dengan ringkasan PDF.')
    nodata = D[D.Status == 'NODATA']
    if len(nodata):
        rpt = (nodata.groupby(['Period', 'NIK', 'Name'])
               .agg(Hari=('Date', 'count'),
                    Tanggal=('Date', lambda x: ', '.join(d.strftime('%d') for d in sorted(x))))
               .reset_index())
        st.warning(f'LAPORAN: {len(nodata)} hari kerja kosong di PDF (NODATA) pada {rpt.NIK.nunique()} karyawan. '
                   'Cek ke HRD; kalau sudah dikonfirmasi tambahkan ke MANUAL_OVERRIDES.')
        st.dataframe(rpt, use_container_width=True)
    n_manual = int((D.Status_Source == 'Manual').sum())
    if n_manual:
        st.caption(f'{n_manual} hari dikoreksi manual (Status_Source = Manual).')

    t1, t2, t3, t4, t5 = st.tabs(['Attendance_Daily', 'Employee_Monthly', 'Ringkasan status', 'Ringkasan tim', 'File terbaca'])
    with t1:
        per = st.multiselect('Filter periode', sorted(D.Period.unique()), default=sorted(D.Period.unique()))
        view = D[D.Period.isin(per)]
        if 'Team' in D.columns:
            tms = st.multiselect('Filter tim', sorted(D.Team.unique()), default=sorted(D.Team.unique()))
            view = view[view.Team.isin(tms)]
        st.dataframe(view, use_container_width=True, height=400)
    with t2:
        st.dataframe(S, use_container_width=True, height=400)
    with t3:
        piv = w.pivot_table(index='Period', columns='Status', values='NIK', aggfunc='count', fill_value=0)
        st.dataframe(piv, use_container_width=True)
        st.bar_chart(w.pivot_table(index='Period', columns='Status_Group', values='NIK',
                                   aggfunc='count', fill_value=0))
    with t4:
        if 'Team' in D.columns:
            ts = team_summary(D)
            st.dataframe(ts, use_container_width=True, height=400,
                         column_config={'Tingkat_Hadir': st.column_config.NumberColumn(format='percent')})
            st.bar_chart(ts.pivot_table(index='Team', columns='Period', values='Tingkat_Hadir'))
        else:
            st.info('Upload file tim untuk melihat ringkasan per tim.')
    with t5:
        st.dataframe(pd.DataFrame(info), use_container_width=True)

    st.divider()
    xlsx = build_workbook(D, S, removed)
    stamp = f'{D.Period.min()}-{D.Period.max()}'
    d1, d2 = st.columns(2)
    d1.download_button('Download Excel (semua sheet)', xlsx, f'HRIS_Attendance_{stamp}.xlsx',
                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    d2.download_button('Download CSV (Attendance_Daily, untuk Looker Studio)',
                       D.to_csv(index=False).encode('utf-8'), f'Attendance_Daily_{stamp}.csv', 'text/csv')

    # ---- Kirim ke Google Sheets
    st.divider()
    st.subheader('Kirim ke Google Sheets (untuk Looker Studio)')
    ok, why = gs_ready()
    if not sid:
        st.info('Isi Spreadsheet URL / ID di sidebar (bagian Google Sheets) untuk mengaktifkan.')
    elif not ok:
        st.warning(why)
    else:
        try:
            existing = gs_existing_periods(sid)
        except Exception as e:
            existing = None
            st.error(f'Tidak bisa membuka spreadsheet: {e}')
        if existing is not None:
            new_keys = sorted(int(k) for k in D.Period_Key.unique())
            overlap = [k for k in new_keys if k in existing]
            lbl = lambda ks: ', '.join(period_info(k)['Period_Label'] for k in ks)
            st.write(f'Periode di hasil proses: **{lbl(new_keys)}**. Periode di sheet saat ini: '
                     f'{lbl(existing) if existing else "(kosong)"}.')
            if overlap:
                st.warning(f'Periode {lbl(overlap)} sudah ada di sheet. Baris lama periode itu akan DIGANTI '
                           'dengan hasil proses ini (tidak dobel).')
            else:
                st.success('Periode baru. Data akan ditambahkan di bawah data yang ada.')
            if team_from == 'file':
                st.caption('File tim yang diupload akan digabung ke sheet Master_Team.')

            team_to_push = team_df if team_from == 'file' else None
            if st.session_state.pop('push_pending', False) and auto_push:      # sekali per klik Proses
                if issues:
                    st.session_state['gs_last'] = ('warning', 'Tidak dikirim otomatis karena ada selisih validasi. '
                                                   'Periksa dulu, lalu kirim manual di bawah.')
                else:
                    try:
                        with st.spinner('Mengirim ke Google Sheets...'):
                            msg = run_push(sid, D, S, team_to_push)
                        st.session_state['gs_last'] = ('success', 'Otomatis terkirim. ' + msg)
                    except Exception as e:
                        st.session_state['gs_last'] = ('error', f'Gagal mengirim otomatis: {e}')
            if 'gs_last' in st.session_state:
                getattr(st, st.session_state['gs_last'][0])(st.session_state['gs_last'][1])

            confirm = st.checkbox('Kirim ulang manual: saya sudah cek hasilnya', key='gs_confirm')
            if st.button('Kirim ke Google Sheets', disabled=not confirm):
                try:
                    with st.spinner('Mengirim ke Google Sheets...'):
                        msg = run_push(sid, D, S, team_to_push)
                    st.session_state['gs_last'] = ('success', 'Terkirim. ' + msg)
                    st.rerun()
                except Exception as e:
                    st.error(f'Gagal mengirim: {e}')


if __name__ == '__main__':
    main()
