import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_16():
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

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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

    # --- KOP SURAT ---
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

    # Double border line
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # TITLE
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI (POIN 16)')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: DETEKSI DINI DAN PENANGANAN MASALAH\n(Lampiran: Standar Operasional Prosedur, Buku Kasus, & Laporan Penanganan)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: SOP ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('A. STANDAR OPERASIONAL PROSEDUR (SOP) DETEKSI DINI & PENANGANAN MASALAH')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    sop_table = doc.add_table(rows=6, cols=2)
    sop_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sop_table)

    items = [
        ('1. Tujuan', 'Memberikan panduan baku dan terstruktur bagi pendidik serta tenaga kependidikan dalam mengidentifikasi gejala dini, mencegah, serta menyelesaikan permasalahan perilaku, kedisiplinan, dan potensi kekerasan/perundungan (bullying) peserta didik secara cepat, edukatif, adil, dan ramah anak.'),
        ('2. Ruang Lingkup', 'Berlaku bagi seluruh peserta didik, pendidik, dan tenaga kependidikan di lingkungan SD Negeri [Nama Sekolah], baik pada saat kegiatan intrakurikuler, kokurikuler, ekstrakurikuler, maupun waktu istirahat.'),
        ('3. Penanggung Jawab', '1. Kepala Sekolah (Penanggung Jawab Utama)\n2. Tim Pencegahan dan Penanganan Kekerasan (TPPK) Sekolah\n3. Guru Kelas / Wali Kelas\n4. Guru Piket & Guru Bimbingan Konseling / Pendamping'),
        ('4. Mekanisme Deteksi Dini', 'a. Observasi Harian: Pemantauan kehadiran, antusiasme, dan perubahan perilaku saat pembiasaan pagi (senam/berdoa).\nb. Kotak Saran & Curhat Sahabat: Fasilitas penyampaian informasi rahasia dari peserta didik.\nc. Buku Catatan Perkembangan Emosi: Pencatatan harian oleh wali kelas terhadap siswa yang menunjukkan sikap menarik diri atau murung.\nd. Koordinasi Paguyuban Orang Tua: Komunikasi proaktif jika terdapat anomali perilaku anak di rumah.'),
        ('5. Alur Penanganan Kasus', 'Tahap 1 (Penerimaan Informasi): Laporan dicatat secara rahasia oleh wali kelas/TPPK maksimal 1x24 jam.\nTahap 2 (Identifikasi & Klarifikasi): Pemeriksaan fakta secara terpisah dan privat tanpa ada unsur intimidasi.\nTahap 3 (Tindakan Terpadu):\n  - Kasus Ringan: Pembinaan internal, konseling persuasif bersama wali kelas.\n  - Kasus Sedang: Mediasi restoratif melibatkan orang tua, wali kelas, guru BK, dan Kepala Sekolah.\n  - Kasus Berat: Rujukan ke instansi berwenang (Puskesmas/Dinas Sosial/Dinas Pendidikan) sesuai regulasi perundang-undangan.\nTahap 4 (Evaluasi Pasca Penanganan): Pemantauan dan pendampingan selama 14 hari kerja untuk memastikan pemulihan hubungan harmonis antarsiswa.'),
        ('6. Dokumen Terkait', 'Buku Kasus Siswa, Lembar Catatan Konseling, Surat Panggilan Orang Tua, Berita Acara Mediasi Kasus.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(items):
        row = sop_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 100, 100, 120, 120)
        set_cell_margins(c1, 100, 100, 120, 120)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(col0)
        r0.bold = True
        r0.font.size = Pt(9)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(col1)
        r1.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==================== BAGIAN 2: BUKU KASUS ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('B. FORMAT BUKU KASUS SISWA (CATATAN PELANGGARAN & PENANGANAN)')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    case_table = doc.add_table(rows=4, cols=7)
    case_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(case_table)

    headers = ['No', 'Tanggal', 'Nama Siswa / Kls', 'Uraian Masalah / Perilaku', 'Pihak Terlibat', 'Tindakan / Bimbingan', 'Status']
    header_row = case_table.rows[0]
    for idx, text in enumerate(headers):
        cell = header_row.cells[idx]
        set_cell_background(cell, '1565C0')
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    sample_data = [
        ('1', '14/08/2026', 'Rian P. (Kls IV-A)\nDoni K. (Kls IV-A)', 'Perselisihan saat bermain di jam istirahat (rebutan bola)', 'Wali Kelas IV-A', 'Mediasi internal di ruang guru, saling meminta maaf dan bersalaman', 'Selesai\n(Damai)'),
        ('2', '08/09/2026', 'Siti N. (Kls V-B)', 'Menarik diri, sering murung dan penurunan nilai belajar', 'Wali Kelas & Guru BK', 'Konseling afektif individual serta konfirmasi kondisi anak kepada orang tua', 'Dalam\nPemantauan'),
        ('3', '15/10/2026', 'Fajar M. (Kls VI-A)', 'Terlambat masuk sekolah >3 kali dalam 1 minggu', 'Guru Piket & Wali Kelas', 'Komunikasi dengan orang tua terkait penyesuaian pola bangun dan tidur anak', 'Selesai')
    ]

    col_w = [Inches(0.4), Inches(0.9), Inches(1.2), Inches(1.8), Inches(1.0), Inches(1.1), Inches(0.7)]
    for r_idx, data_row in enumerate(sample_data):
        row = case_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_w[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 90, 90, 80, 80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 1, 6]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==================== BAGIAN 3: LAPORAN PENANGANAN ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('C. FORMULIR LAPORAN HASIL PENANGANAN KASUS (BERITA ACARA)')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    rep_table = doc.add_table(rows=7, cols=2)
    rep_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rep_table)

    rep_items = [
        ('1. Nomor Laporan', 'LAP-TPPK/2026/XI/007'),
        ('2. Hari / Tanggal Penanganan', '[Contoh: Rabu, 21 Oktober 2026]'),
        ('3. Identitas Siswa', 'Nama Lengkap : [Nama Siswa]\nKelas / NISN : [Kelas / Nomor Induk Siswa Nasional]'),
        ('4. Kronologi Kejadian', 'Laporan diterima dari wali kelas mengenai perselisihan antarsiswa di luar jam pembelajaran. Tim TPPK bersama wali kelas segera mengamankan situasi dan mengundang pihak terkait untuk klarifikasi di ruang pertemuan khusus secara tertutup.'),
        ('5. Upaya Mediasi & Penanganan', '1. Mendengarkan keterangan dari masing-masing siswa secara terpisah tanpa intimidasi.\n2. Mengarahkan siswa untuk memahami dampak dari perbuatan masing-masing (pendekatan disiplin positif).\n3. Pelaksanaan mediasi terbuka antarsiswa dengan disaksikan wali kelas dan orang tua/wali murid.\n4. Penandatanganan nota kesepakatan bersama untuk saling menghormati dan tidak mengulangi kejadian.'),
        ('6. Rencana Tindak Lanjut', '- Pemantauan perilaku siswa secara berkala oleh wali kelas selama 14 hari ke depan.\n- Orang tua berkomitmen memantau aktivitas dan waktu istirahat siswa di rumah.'),
        ('7. Kesimpulan & Status', 'KASUS SELESAI DENGAN MUFAKAT DAN RESTORASI HUBUNGAN SOSIAL')
    ]

    for row_idx, (col0, col1) in enumerate(rep_items):
        row = rep_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 90, 90, 120, 120)
        set_cell_margins(c1, 90, 90, 120, 120)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(col0)
        r0.bold = True
        r0.font.size = Pt(8.5)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(col1)
        r1.font.size = Pt(8.5)

    # Tanda Tangan
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
    p0.add_run('Ketua TPPK Sekolah,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Guru BK / Wali Kelas,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_16_Deteksi_Dini_dan_Penanganan_Masalah.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_16:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_16()
