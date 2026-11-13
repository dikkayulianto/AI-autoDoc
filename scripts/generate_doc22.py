import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_22():
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

    def set_cell_margins(cell, top=90, bottom=90, left=120, right=120):
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 22')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: KOORDINASI WAKA KESISWAAN / GURU KELAS\n(Berkas Lengkap: Program Kerja Pembinaan Kesiswaan, Notulen Rapat Koordinasi, & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: PROGRAM KERJA KESISWAAN ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: PROGRAM KERJA KOORDINASI PEMBINAAN KESISWAAN DAN GURU KELAS')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    prog_table = doc.add_table(rows=6, cols=2)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prog_table)

    prog_data = [
        ('1. Dasar Program', 'Program Kerja Tahunan Kesiswaan SD Negeri [Nama Sekolah] Tahun Ajaran 2026/2027 tentang Sinergi Guru Kelas dalam Penguatan Disiplin Positif & Pembiasaan Karakter Siswa.'),
        ('2. Tujuan Koordinasi', '1. Menyelaraskan standar pembiasaan karakter di seluruh rombel kelas (Fase A, B, dan C).\n2. Membangun komunikasi berkesinambungan antara Waka/Koordinator Kesiswaan dengan seluruh wali kelas mengenai perkembangan sikap dan disiplin murid.\n3. Melakukan deteksi dini dan respons cepat terhadap dinamika peserta didik.'),
        ('3. Jadwal & Frekuensi Pertemuan', 
         'a. Koordinasi Rutin Bulanan: Dilaksanakan setiap hari Jumat minggu terakhir setiap bulan pukul 13.00 - 15.00 WIB.\n'
         'b. Rapat Kasus Khusus (Insidental): Digelar maksimal 1x24 jam jika terdapat laporan permasalahan disiplin atau perilaku siswa yang membutuhkan penanganan bersama.\n'
         'c. Briefing Pagi Kesiswaan: Evaluasi singkat 10 menit setiap hari Senin setelah apel/upacara bendera.'),
        ('4. Ruang Lingkup Koordinasi', 
         '• Evaluasi Keterlaksanaan Pembiasaan 7 Hari Berkarakter (salat/doa, bangun pagi, senam, makan sehat, gemar membaca).\n'
         '• Pemantauan presensi dan penanganan siswa yang sering terlambat/absen tanpa keterangan.\n'
         '• Koordinasi pembinaan kegiatan ekstrakurikuler kepanduan (Praja Muda Karana / Pra Siaga).\n'
         '• Penjadwalan program apresiasi "Bintang Karakter Bulan Ini" di setiap kelas.'),
        ('5. Mekanisme Pelaporan', 'Setiap wali kelas mengumpulkan Rekapitulasi Catatan Perkembangan Siswa bulanan kepada Koordinator Kesiswaan untuk dianalisis bersama.'),
        ('6. Output yang Diharapkan', 'Terciptanya iklim sekolah yang tertib, harmonis, konsisten dalam penerapan disiplin positif tanpa hukuman fisik.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(prog_data):
        row = prog_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 80, 80, 110, 110)
        set_cell_margins(c1, 80, 80, 110, 110)
        
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

    # ==================== BAGIAN 2: NOTULEN RAPAT KOORDINASI ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: NOTULEN RAPAT EVALUASI KESISWAAN DAN WALI KELAS')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    notulen_table = doc.add_table(rows=6, cols=2)
    notulen_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(notulen_table)

    notulen_data = [
        ('1. Waktu & Tempat', 'Hari / Tanggal : Jumat, 25 September 2026 (Pukul 13.00 - 15.00 WIB)\nTempat         : Ruang Guru SD Negeri [Nama Sekolah]'),
        ('2. Peserta Hadir', 'Kepala Sekolah, Koordinator Kesiswaan, Guru Bimbingan Konseling, dan Seluruh Guru Kelas I - VI (Lengkap: 14 Orang Pendidik).'),
        ('3. Agenda Rapat', '1. Evaluasi Keterlaksanaan Pembiasaan Senam Pagi dan Gerakan Literasi Sekolah.\n2. Penyelarasan Penanganan Siswa yang Sering Terlambat.\n3. Rencana Penetapan Duta Karakter Ramah Anak Tingkat Sekolah.'),
        ('4. Hasil Pembahasan & Kesepakatan', 
         'a. Evaluasi Senam Anak Indonesia Hebat:\n'
         '   - Tingkat kehadiran dan keikutsertaan senam pagi mencapai 96%.\n'
         '   - Untuk kelas rendah (I & II), instruktur senam akan divariasikan melibatkan kakak kelas V sebagai tutor sebaya agar lebih ceria.\n'
         'b. Penanganan Kedisiplinan Waktu (Keterlambatan):\n'
         '   - Disepakati pendekatan disiplin positif: Siswa yang datang lewat dari pukul 07.00 WIB tidak dihukum fisik/dijemur di lapangan, melainkan diarahkan membaca buku di perpustakaan/sudut baca selama 15 menit dan mengisi lembar komitmen waktu.\n'
         '   - Wali kelas memverifikasi alasan keterlambatan kepada orang tua jika berulang lebih dari 2 kali.\n'
         'c. Apresiasi Positif:\n'
         '   - Menetapkan "Pohon Kebaikan" di setiap dinding kelas, di mana daunnya ditempeli nama siswa yang menunjukkan sikap saling tolong-menolong.'),
        ('5. Keputusan & Tindak Lanjut', 
         '• Setiap wali kelas memperbarui Buku Catatan Wali Kelas setiap hari Jumat.\n'
         '• Kesiswaan menyiapkan sertifikat apresiasi bagi kelas terbersih dan paling disiplin pada upacara Senin mendatang.'),
        ('6. Penutup', 'Rapat ditutup dengan doa bersama pada pukul 15.00 WIB.')
    ]

    for row_idx, (col0, col1) in enumerate(notulen_data):
        row = notulen_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 80, 80, 110, 110)
        set_cell_margins(c1, 80, 80, 110, 110)
        
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

    # ==================== BAGIAN 3: DAFTAR HADIR KOORDINASI ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: DAFTAR HADIR RAPAT KOORDINASI KESISWAAN DAN GURU KELAS')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    att_table = doc.add_table(rows=8, cols=5)
    att_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(att_table)

    att_headers = ['No', 'Nama Lengkap Pendidik', 'Jabatan / Tugas', 'Materi Koordinasi', 'Tanda Tangan']
    header_row = att_table.rows[0]
    for idx, text in enumerate(att_headers):
        cell = header_row.cells[idx]
        set_cell_background(cell, '1565C0')
        set_cell_margins(cell, 80, 80, 60, 60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)

    att_rows = [
        ('1', 'Dra. Hj. Siti Maryam, M.Pd', 'Kepala Sekolah', 'Arahan Kebijakan Disiplin', '1. ....................'),
        ('2', 'Agus Hendrawan, S.Pd', 'Waka / Koord. Kesiswaan', 'Pemimpin Rapat & Program', '2. ....................'),
        ('3', 'Nurul Inayah, S.Pd', 'Guru Kelas I-A', 'Evaluasi Pembiasaan Mandiri', '3. ....................'),
        ('4', 'Bambang Supriyadi, S.Pd', 'Guru Kelas II-B', 'Evaluasi Sarapan Sehat', '4. ....................'),
        ('5', 'Dewi Lestari, S.Pd', 'Guru Kelas III-A', 'Pelaporan Kehadiran Senam', '5. ....................'),
        ('6', 'Rian Hidayat, S.Pd.SD', 'Guru Kelas IV-A', 'Disiplin Waktu & Jam Belajar', '6. ....................'),
        ('7', 'Fitriani, S.Pd', 'Guru BK / Pendamping', 'Progres Bimbingan Konseling', '7. ....................')
    ]

    col_att = [Inches(0.4), Inches(2.2), Inches(1.6), Inches(1.3), Inches(1.3)]
    for r_idx, data_row in enumerate(att_rows):
        row = att_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_att[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 70, 70, 60, 60)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 4: DOKUMENTASI FOTO ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: DOKUMENTASI KEGIATAN KOORDINASI KESISWAAN')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=1, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    cell_p1 = photo_table.rows[0].cells[0]
    cell_p2 = photo_table.rows[0].cells[1]
    cell_p1.width = Inches(3.4)
    cell_p2.width = Inches(3.4)
    set_cell_background(cell_p1, 'ECEFF1')
    set_cell_background(cell_p2, 'ECEFF1')
    set_cell_margins(cell_p1, 240, 240, 90, 90)
    set_cell_margins(cell_p2, 240, 240, 90, 90)

    p_p1 = cell_p1.paragraphs[0]
    p_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p1.add_run('[ TEMPELKAN FOTO RAPAT KOORDINASI KESISWAAN DI SINI ]\n\nKeterangan: Pertemuan rutin bulanan Waka Kesiswaan bersama guru kelas membahas disiplin positif siswa.').font.size = Pt(8.5)

    p_p2 = cell_p2.paragraphs[0]
    p_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p2.add_run('[ TEMPELKAN FOTO BRIEFING HARIAN GURU KELAS DI SINI ]\n\nKeterangan: Evaluasi singkat keterlaksanaan pembiasaan pagi dan penanganan siswa di ruang guru.').font.size = Pt(8.5)

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
    p0.add_run('Waka / Koord. Kesiswaan,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Perwakilan Wali Kelas,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_22_Koordinasi_Waka_Kesiswaan_dan_Guru_Kelas.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_22:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_22()
