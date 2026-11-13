import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_24():
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    def set_cell_background(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=80, bottom=80, left=110, right=110):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_table_borders(table, color='B0BEC5', sz='4', val='single'):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    # --- KOP RESMI ---
    p_kop = doc.add_paragraph()
    p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop.paragraph_format.space_after = Pt(2)
    r = p_kop.add_run('PEMERINTAH KOTA CIREBON\nDINAS PENDIDIKAN\nSD NEGERI [NAMA SEKOLAH]')
    r.bold = True
    r.font.size = Pt(12)
    r.font.name = 'Arial'

    p_kop2 = doc.add_paragraph()
    p_kop2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop2.paragraph_format.space_after = Pt(8)
    r2 = p_kop2.add_run('Alamat: Jalan [Alamat Sekolah], Kec. [Harjamukti/Kesambi], Kota Cirebon 45132\nPosel: sdnegeri@cirebonkota.go.id | Laman: disdik.cirebonkota.go.id')
    r2.font.size = Pt(8.5)
    r2.italic = True
    r2.font.name = 'Arial'
    r2.font.color.rgb = RGBColor(100, 100, 100)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # JUDUL DOKUMEN
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 24')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: LAYANAN PSIKOSOSIAL GURU BIMBINGAN KONSELING (BK)\n(Dokumen Komprehensif: Program Kerja Layanan BK, Jurnal Log Layanan, Laporan Konseling Individual & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: PROGRAM KERJA LAYANAN BK ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: PROGRAM KERJA BIMBINGAN KONSELING & DUKUNGAN PSIKOSOSIAL')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    prog_table = doc.add_table(rows=6, cols=2)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prog_table)

    prog_data = [
        ('1. Dasar & Kerangka Kerja', 'Panduan Operasional Penyelenggaraan Bimbingan dan Konseling di Sekolah Dasar (Kemendikbudristek) dan Program Dukungan Psikososial Ramah Anak Jabar Beraksi Tahun Ajaran 2026/2027.'),
        ('2. Tujuan Layanan BK', 
         '• Membantu peserta didik mencapai kematangan emosional, penyesuaian sosial, serta kesehatan mental yang stabil.\n'
         '• Memberikan pendampingan psikososial bagi anak yang mengalami trauma, kecemasan belajar, atau hambatan pertemanan.\n'
         '• Memfasilitasi mediasi antarteman sebaya dengan prinsip tanpa menghakimi (non-judgmental).'),
        ('3. Bidang Layanan Pokok', 
         'a. Bimbingan Pribadi: Regulasi emosi, pembentukan konsep diri positif, pengelolaan kecemasan.\n'
         'b. Bimbingan Sosial: Keterampilan komunikasi empati, pencegahan perundungan, etika pergaulan.\n'
         'c. Bimbingan Belajar: Manajemen waktu tidur dan belajar, mengatasi kejenuhan, motivasi prestasi.\n'
         'd. Bimbingan Karir Awal: Pengenalan bakat, minat, dan cita-cita sejak dini.'),
        ('4. Jenis Strategi Layanan', 
         '1. Layanan Dasar (Bimbingan Klasikal): Sosialisasi pembiasaan positif di kelas 1x sebulan.\n'
         '2. Layanan Responsif: Konseling individual privat, konseling kelompok sebaya, dan advokasi mediasi.\n'
         '3. Dukungan Sistem: Kolaborasi bersama wali kelas, orang tua, dan rujukan profesional (Puskesmas/Psikolog bila diperlukan).'),
        ('5. Prinsip Etika Konseling', 'Kerahasiaan data anak terjamin penuh, berpusat pada kepentingan terbaik bagi anak (best interests of the child), sukarela, dan keterbukaan.'),
        ('6. Evaluasi & Tindak Lanjut', 'Evaluasi proses dan hasil layanan konseling dilaporkan secara berkala per caturwulan kepada Kepala Sekolah.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(prog_data):
        row = prog_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 75, 75, 100, 100)
        set_cell_margins(c1, 75, 75, 100, 100)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(col0)
        r0.bold = True
        r0.font.size = Pt(8.5)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(col1)
        r1.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 2: JURNAL LOG HARIAN LAYANAN BK ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: JURNAL LOG HARIAN PELAKSANAAN LAYANAN KONSELING & PSIKOSOSIAL')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    log_table = doc.add_table(rows=6, cols=6)
    log_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(log_table)

    log_headers = ['No', 'Hari / Tanggal', 'Inisial Siswa / Kls', 'Jenis & Topik Layanan', 'Gejala / Masalah yang Dialami', 'Hasil Layanan & Tindak Lanjut']
    header_row = log_table.rows[0]
    for idx, text in enumerate(log_headers):
        cell = header_row.cells[idx]
        set_cell_background(cell, '1565C0')
        set_cell_margins(cell, 80, 80, 50, 50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)

    log_rows = [
        ('1', 'Rabu, 09/09/2026', 'R.A. (Kls IV-A)', 'Konseling Individual\n(Sosial-Emosional)', 'Menarik diri dari lingkaran teman sebangku, merasa malu karena sering gagap saat bicara.', 'Latihan afirmasi positif dan teknik napas relaksasi. Dikoordinasikan dengan wali kelas untuk pendampingan suportif.'),
        ('2', 'Senin, 21/09/2026', 'Kelompok 5\n(Kls V-B, 4 Anak)', 'Bimbingan Kelompok\n(Keterampilan Sosial)', 'Perselisihan pembagian peran dalam tugas kelompok yang memicu perdebatan sengit.', 'Mediasi gaya komunikasi terbuka: setiap anak mengungkapkan ide secara bergantian. Masalah tuntas dengan mufakat.'),
        ('3', 'Kamis, 08/10/2026', 'D.K. (Kls VI-A)', 'Konseling Belajar &\nKesehatan Mental', 'Mengalami kecemasan berlebih menghadapi asesmen formatif, sulit tidur sebelum jam 23.00.', 'Edukasi jadwal istirahat, manajemen stres, dan pemahaman bahwa nilai ujian bukanlah penentu tunggal harga diri.'),
        ('4', 'Selasa, 20/10/2026', 'S.M. (Kls III-B)', 'Layanan Mediasi Sebaya\n(Disiplin Positif)', 'Saling berebut buku bacaan di pojok literasi kelas hingga robek.', 'Kedua siswa diajak memperbaiki buku bersama-sama, menyepakati giliran membaca tanpa emosi negatif.'),
        ('5', 'Jumat, 30/10/2026', 'Kelas IV-B\n(30 Siswa)', 'Bimbingan Klasikal\n(Pencegahan Bullying)', 'Materi pengenalan 3 jenis perundungan (verbal, fisik, siber) dan cara menjadi "Upstander".', 'Siswa berkomitmen menandatangani Pakta Kelas Ramah Sahabat Bebas Olok-olok.')
    ]

    col_l = [Inches(0.4), Inches(1.1), Inches(1.1), Inches(1.4), Inches(1.5), Inches(1.7)]
    for r_idx, data_row in enumerate(log_rows):
        row = log_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_l[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 60, 60, 50, 50)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(7.8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 3: LAPORAN KONSELING INDIVIDUAL ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: LAPORAN HASIL LAYANAN KONSELING INDIVIDUAL (STUDI KASUS)')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    case_table = doc.add_table(rows=7, cols=2)
    case_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(case_table)

    case_data = [
        ('1. Identitas Konseli (Disamarkan)', 'Inisial: R.A. | Usia: 10 Tahun | Kelas: IV-A | Jenis Kelamin: Laki-laki'),
        ('2. Sumber Rujukan', 'Rujukan Wali Kelas IV-A (Catatan: Siswa sering melamun dan nilai tugas menurun selama 3 pekan).'),
        ('3. Deskripsi Masalah & Diagnosis Awal', 'Konseli mengalami tekanan psikologis ringan pasca kehilangan hewan peliharaan kesayangan serta merasa terisolasi karena teman sebangku berpindah tempat duduk. Hal ini memicu penurunan konsentrasi belajar dan rasa enggan bersosialisasi.'),
        ('4. Pendekatan Konseling yang Digunakan', 
         '• Pendekatan Person-Centered (Client-Centered Therapy): Membangun empati, mendengarkan aktif tanpa menyela, dan menciptakan suasana aman agar anak bebas menumpahkan perasaannya.\n'
         '• Cognitive Restructuring: Mengubah pikiran negatif ("Teman-teman tidak menyukai saya") menjadi pemahaman realistis bahwa teman-teman tetap peduli.'),
        ('5. Proses & Sesi Layanan', 
         'Sesi 1 (14 Okt 2026): Membangun rapport, mengidentifikasi akar kesedihan, ventilasi emosi.\n'
         'Sesi 2 (21 Okt 2026): Eksplorasi kekuatan diri dan menetapkan target harian menyapa 2 kawan kembara bermain.\n'
         'Sesi 3 (28 Okt 2026): Evaluasi progres dan penguatan kepercayaan diri.'),
        ('6. Evaluasi Hasil & Dampak', 'Siswa menunjukkan perubahan positif nyata: kembali ceria, mulai aktif dalam diskusi kelompok, dan nafsu makan bekal kembali normal.'),
        ('7. Rekomendasi Tindak Lanjut', 'Wali kelas terus memberikan apresiasi berkala atas keterlibatan anak di kelas. Status konseling: Selesai dengan hasil memuaskan.')
    ]

    for row_idx, (col0, col1) in enumerate(case_data):
        row = case_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 75, 75, 100, 100)
        set_cell_margins(c1, 75, 75, 100, 100)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(col0)
        r0.bold = True
        r0.font.size = Pt(8.5)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(col1)
        r1.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 4: DOKUMENTASI FOTO LAYANAN BK ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: DOKUMENTASI LAYANAN BIMBINGAN KONSELING RAMAH ANAK')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: RUANG KONSELING NYAMAN & RAMAH ANAK ]\n\nKeterangan: Ruang bimbingan konseling yang bersih, nyaman, privat, dan dilengkapi media gambar ekspresi emosi anak.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: SESI KONSELING INDIVIDUAL SUPORTIF ]\n\nKeterangan: Guru BK mendengarkan keluhan peserta didik dengan penuh kehangatan, empati, dan menjaga kerahasiaan.', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: BIMBINGAN KELOMPOK DINAMIKA TEMAN SEBAYA ]\n\nKeterangan: Diskusi kelompok kecil terarah untuk melatih keterampilan komunikasi asertif dan kerja sama.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: PAPAN INFORMASI & KOTAK CURHAT SAHABAT ]\n\nKeterangan: Media edukasi poster kesehatan mental sekolah dan kotak penyampaian curahan hati murid.', photo_table.rows[1].cells[1])
    ]

    for text, cell in frames:
        cell.width = Inches(3.4)
        set_cell_background(cell, 'ECEFF1')
        set_cell_margins(cell, 180, 180, 90, 90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.font.size = Pt(8)

    # PENGESAHAN
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    p_ttd_date = doc.add_paragraph()
    p_ttd_date.paragraph_format.space_after = Pt(8)
    p_ttd_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_dt = p_ttd_date.add_run('Cirebon, [Tanggal Bulan 2026]')
    r_dt.font.size = Pt(9.5)

    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in sig_table.rows:
        for c in r.cells:
            tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="0"/><w:bottom w:w="0"/><w:left w:w="0"/><w:right w:w="0"/></w:tcMar>')
            c._tc.get_or_add_tcPr().append(tcMar)

    c0, c1, c2 = sig_table.rows[0].cells[0], sig_table.rows[0].cells[1], sig_table.rows[0].cells[2]
    c0.width, c1.width, c2.width = Inches(2.2), Inches(2.2), Inches(2.4)

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.add_run('Guru Bimbingan Konseling,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Koordinator Kesiswaan,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_24_Layanan_Psikososial_Guru_BK.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_24:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_24()
