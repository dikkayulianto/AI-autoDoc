import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_26():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 26')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: PEMBELAJARAN BERMAKNA DAN INTEGRASI KARAKTER HOLISTIK INTEGRATIF\n(Dokumen Komprehensif: Modul Ajar Holistik, Instrumen Observasi Keterlibatan Murid, & Dokumentasi Proyek Nyata)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: DESAIN MODUL AJAR HOLISTIK INTEGRATIF ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: DESAIN MODUL AJAR PEMBELAJARAN BERMAKNA & HOLISTIK INTEGRATIF')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    mod_table = doc.add_table(rows=7, cols=2)
    mod_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(mod_table)

    mod_data = [
        ('1. Identitas Pembelajaran', 'Mata Pelajaran : IPAS & Pendidikan Karakter Terintegrasi\nFase / Kelas   : Fase B / Kelas IV (Empat)\nTopik Tematik  : "Pahlawan Lingkungan: Bijak Kelola Sampah & Menjaga Keseimbangan Alam"\nAlokasi Waktu  : 3 JP (3 x 35 Menit) - Pembelajaran Berbasis Masalah (PBL)'),
        ('2. Pendekatan Holistik Integratif', 
         '• Dimensi Kognitif (Pengetahuan): Memahami siklus sampah, dampak plastik terhadap ekosistem pesisir Cirebon.\n'
         '• Dimensi Afektif (Karakter & Nilai): Tanggung jawab ekologis, empati pada petugas kebersihan, kepedulian bersama.\n'
         '• Dimensi Psikomotorik (Tindakan Nyata): Keterampilan memilah sampah organik/anorganik dan kreasi daur ulang sederhana.\n'
         '• Dimensi Psikososial (Kesehatan Batin): Kepuasan emosi saat mampu berkontribusi positif bagi lingkungan sekitar.'),
        ('3. Tujuan Pembelajaran (TP)', 'Peserta didik mampu mengidentifikasi jenis sampah di sekitarnya, menghubungkan dampaknya dengan kehidupan nyata sehari-hari, serta merancang aksi nyata pemilahan sampah di sekolah dan rumah secara gotong royong.'),
        ('4. Langkah Kegiatan Bermakna (Meaningful Learning)', 
         'a. Eksplorasi Nyata (20 Menit): Siswa diajak melakukan "Observasi Lingkungan Sekolah" (School Walk) mengamati kondisi tempat sampah dan saluran air.\n'
         'b. Penyelidikan Kolaboratif (45 Menit): Dalam kelompok 4-5 anak, siswa menganalisis: "Berapa banyak sampah plastik kemasan jajanan yang kita hasilkan setiap hari?" dan mendiskusikan solusi pengurangan sampah.\n'
         'c. Aksi Konkret (30 Menit): Praktik pemilahan sampah langsung ke tempat sampah terpilah (Organik, Anorganik, B3) serta pembuatan poster kampanye kreatif: "Bawa Tumbler & Kotak Bekal Sendiri".'),
        ('5. Refleksi Batiniah & Integrasi Nilai', 'Refleksi terpimpin: "Menjaga kebersihan bukan sekadar menaati aturan guru, melainkan wujud rasa syukur kita kepada Tuhan atas bumi yang asri dan wujud kasih sayang kepada sesama makhluk hidup."'),
        ('6. Asesmen Otentik', 'Penilaian Kinerja Pemilahan Sampah, Lembar Observasi Kolaborasi Kelompok, dan Refleksi Portofolio Siswa.'),
        ('7. Keterkaitan Lintas Sektor', 'Terintegrasi dengan program Puskesmas (Kesehatan Lingkungan) dan Dinas Lingkungan Hidup Kota Cirebon.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(mod_data):
        row = mod_table.rows[row_idx]
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

    # ==================== BAGIAN 2: INSTRUMEN OBSERVASI KETERLIBATAN ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: INSTRUMEN OBSERVASI KETERLIBATAN SISWA DALAM BELAJAR BERMAKNA')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    p_desc2 = doc.add_paragraph()
    p_desc2.paragraph_format.space_after = Pt(3)
    r_d2 = p_desc2.add_run('Lembar Pengamatan Partisipasi Aktif, Pemahaman Konseptual, dan Aplikasi Karakter Holistik:')
    r_d2.font.size = Pt(8)
    r_d2.italic = True

    obs_table = doc.add_table(rows=7, cols=7)
    obs_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(obs_table)

    obs_headers = ['No', 'Nama Siswa', 'Koneksi Konsep Nyata (Skor 1-4)', 'Kerjasama & Gotong Royong (1-4)', 'Aksi Karakter Mandiri (1-4)', 'Refleksi Diri Batiniah (1-4)', 'Skor Akhir & Kategori']
    header_row = obs_table.rows[0]
    for idx, text in enumerate(obs_headers):
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

    obs_rows = [
        ('1', 'Ahmad Fadillah', '4', '4', '4', '4', '16 (Sangat Holistik / A)'),
        ('2', 'Citra Dewi Lestari', '4', '4', '4', '3', '15 (Sangat Holistik / A)'),
        ('3', 'Bagus Ramadhani', '3', '4', '3', '3', '13 (Baik & Berkembang / B)'),
        ('4', 'Dina Safitri', '4', '3', '4', '4', '15 (Sangat Holistik / A)'),
        ('5', 'Eko Prasetyo', '3', '3', '4', '3', '13 (Baik & Berkembang / B)'),
        ('6', 'Farhan Maulana', '4', '4', '4', '4', '16 (Sangat Holistik / A)')
    ]

    col_ob = [Inches(0.4), Inches(1.5), Inches(0.9), Inches(1.0), Inches(0.9), Inches(0.9), Inches(1.2)]
    for r_idx, data_row in enumerate(obs_rows):
        row = obs_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_ob[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 60, 60, 50, 50)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx not in [1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 3: LAPORAN REFLEKSI HASIL BELAJAR ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: LAPORAN REFLEKSI PEMBELAJARAN BERMAKNA SISWA')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    ref_table = doc.add_table(rows=5, cols=2)
    ref_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ref_table)

    ref_data = [
        ('1. Tanggal Pelaksanaan Proyek', 'Selasa - Rabu, 13 - 14 Oktober 2026'),
        ('2. Temuan Keterlibatan Siswa', '92% siswa menunjukkan antusiasme yang jauh lebih tinggi dibandingkan metode ceramah konvensional. Siswa mampu menjelaskan secara mandiri mengapa botol plastik bekas tidak boleh dibakar atau dibuang sembarangan ke selokan sekolah.'),
        ('3. Resonansi Emosional & Karakter', 'Siswa secara spontan menegur kawannya yang lupa memilah sampah saat jam istirahat dengan kalimat santun: "Yuk dibuang ke tong kuning, kan itu plastik." Hal ini membuktikan nilai karakter telah terinternalisasi menjadi budaya laku, bukan sekadar hafalan ujian.'),
        ('4. Karya Nyata yang Dihasilkan', 'Tersusun 6 set tempat sampah mini terpilah dari kardus bekas di sudut kelas dan 12 poster edukasi ajakan membawa bekal ramah lingkungan yang dipajang di majalah dinding sekolah.'),
        ('5. Kesimpulan Evaluasi Guru', 'Pembelajaran bermakna secara holistik integratif terbukti efektif meningkatkan kompetensi akademik sekaligus membentuk akhlak mulia cinta lingkungan.')
    ]

    for row_idx, (col0, col1) in enumerate(ref_data):
        row = ref_table.rows[row_idx]
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

    # ==================== BAGIAN 4: DOKUMENTASI FOTO ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: DOKUMENTASI PEMBELAJARAN BERMAKNA HOLISTIK INTEGRATIF')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: OBSERVASI NYATA LINGKUNGAN SEKOLAH ]\n\nKeterangan: Peserta didik mengamati kondisi lingkungan dan mengidentifikasi jenis sampah di area luar kelas.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: DISKUSI KOLABORATIF PEMECAHAN MASALAH ]\n\nKeterangan: Kerja sama kelompok merumuskan solusi pengurangan sampah plastik dan merancang poster kampanye.', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: PRAKTIK PEMILAHAN SAMPAH MANDIRI ]\n\nKeterangan: Penerapan nyata nilai peduli lingkungan melalui pemilahan sampah organik dan anorganik.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: PRESENTASI KARYA & REFLEKSI BATINIAH ]\n\nKeterangan: Siswa mempresentasikan hasil aksi nyata dan melakukan refleksi rasa syukur atas lingkungan yang bersih.', photo_table.rows[1].cells[1])
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
    p0.add_run('Guru Pengampu Pembelajaran,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Koordinator Kurikulum,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_26_Pembelajaran_Bermakna_Holistik_Integratif.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_26:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_26()
