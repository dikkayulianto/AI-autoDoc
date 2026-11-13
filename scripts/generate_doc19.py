import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_19():
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

    def set_cell_margins(cell, top=90, bottom=90, left=130, right=130):
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

    # Garis Pembatas
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # JUDUL DOKUMEN
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI (POIN 19)')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: PENGUATAN SPIRITUAL, KESEHATAN MENTAL, MORAL, DAN ETIKA\n(Lampiran: Ringkasan Modul Ajar Integratif, Jurnal Pembelajaran Karakter, & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: MODUL AJAR INTEGRASI KARAKTER ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('A. DESAIN MODUL AJAR BERBASIS PENGUATAN SPIRITUAL, MENTAL & MORAL')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    modul_table = doc.add_table(rows=7, cols=2)
    modul_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(modul_table)

    modul_items = [
        ('1. Identitas Pembelajaran', 'Mata Pelajaran : Pendidikan Pancasila & Pembiasaan Karakter Holistik\nFase / Kelas   : Fase B / Kelas IV (Empat)\nSemester / TA  : Ganjil / 2026/2027\nAlokasi Waktu  : 2 x 35 Menit (1 Pertemuan)'),
        ('2. Dimensi Profil Karakter', 
         '• Beriman & Bertakwa kepada Tuhan YME (Penguatan Spiritual)\n'
         '• Berakhlak Mulia kepada Sesama & Diri Sendiri (Penguatan Etika & Mental)\n'
         '• Bernalar Kritis & Bergotong Royong (Penguatan Moral Kebersamaan)'),
        ('3. Tujuan Pembelajaran', 'Peserta didik mampu membiasakan diri bersyukur, mengenali emosi diri secara sehat (kesehatan mental), serta menerapkan sopan santun (etika) dan kejujuran (moral) dalam interaksi kelas sehari-hari.'),
        ('4. Kegiatan Awal (Spiritual & Mental Check-in)', 
         'a. Berdoa bersama sesuai keyakinan dengan khusyuk dan tertib (10 menit).\n'
         'b. Kegiatan "Roda Emosi Pagi": Guru menanyakan kabar dan peserta didik mengekspresikan perasaannya hari ini secara aman tanpa dihakimi.\n'
         'c. Pembiasaan kata ajaib: "Tolong, Maaf, Terima Kasih, dan Permisi".'),
        ('5. Kegiatan Inti (Edukasi Moral & Diskusi Etika)', 
         'a. Menyimak kisah teladan berbasis video pendek tentang "Kejujuran dan Menghargai Batasan Teman".\n'
         'b. Diskusi kelompok terarah (FGD): Membahas studi kasus tentang apa yang harus dilakukan jika melihat teman yang murung atau diejek.\n'
         'c. Bermain peran (role-playing) tentang cara menyampaikan kritik atau ketidaksetujuan secara santun dan damai.'),
        ('6. Kegiatan Penutup (Refleksi Batiniah)', 
         'a. Latihan hening sejenak (mindfulness 3 menit) untuk relaksasi mental dan menumbuhkan rasa syukur.\n'
         'b. Pengisian "Catatan Kebaikan Harian": Siswa menuliskan 1 perbuatan baik yang telah dilakukan dan 1 hal yang disyukuri hari ini.'),
        ('7. Asesmen Ketercapaian', 'Asesmen Formatif Afektif: Lembar Observasi Sikap Berdoa, Rubrik Etika Diskusi, dan Jurnal Refleksi Mandiri.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(modul_items):
        row = modul_table.rows[row_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 90, 90, 110, 110)
        set_cell_margins(c1, 90, 90, 110, 110)
        
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

    # ==================== BAGIAN 2: JURNAL PEMBELAJARAN ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('B. JURNAL PEMBELAJARAN PENGUATAN KARAKTER & ETIKA KELAS')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    p_jurn = doc.add_paragraph()
    p_jurn.paragraph_format.space_after = Pt(4)
    r_jn = p_jurn.add_run('Jurnal Harian Refleksi Karakter Spiritual, Mental, dan Etika Peserta Didik:')
    r_jn.font.size = Pt(9)
    r_jn.italic = True

    jurn_table = doc.add_table(rows=5, cols=6)
    jurn_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(jurn_table)

    jurn_headers = ['No', 'Tanggal', 'Fokus Karakter', 'Aktivitas KBM & Pembiasaan', 'Capaian Peserta Didik', 'Tindak Lanjut / Catatan']
    header_row = jurn_table.rows[0]
    for idx, text in enumerate(jurn_headers):
        cell = header_row.cells[idx]
        set_cell_background(cell, '1565C0')
        set_cell_margins(cell, 90, 90, 70, 70)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    jurn_data = [
        ('1', 'Senin, 14/09/26', 'Spiritual\n(Syukur & Ibadah)', 'Doa pagi bersama, tadarus/renungan nilai kasih, dan refleksi 5 menit nikmat sehat.', 'Siswa berdoa tertib tanpa bercanda, toleransi antarpemeluk agama tampak sangat harmonis.', 'Pertahankan rutinitas doa bergantian.'),
        ('2', 'Selasa, 22/09/26', 'Kesehatan Mental\n(Kelola Emosi)', 'Praktik pernapasan teratur (STOP: Stop, Take a breath, Observe, Proceed) saat jeda belajar.', 'Siswa yang semula tegang saat latihan matematika menjadi lebih relaks dan tenang.', 'Diterapkan setiap transisi mata pelajaran berat.'),
        ('3', 'Kamis, 01/10/26', 'Moral\n(Kejujuran)', 'Kantin Kejujuran & laporan penemuan barang tercecer di lingkungan kelas/sekolah.', 'Ditemukan uang Rp 10.000,- dan alat tulis, diserahkan langsung ke guru piket oleh siswa.', 'Apresiasi nama siswa penemu saat upacara/apel.'),
        ('4', 'Jumat, 16/10/26', 'Etika & Sopan Santun', 'Penerapan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) serta cara menyela pembicaraan.', 'Siswa terbiasa mengacungkan tangan dan berterima kasih setelah menerima bantuan teman.', 'Menjadi budaya permanen seluruh civitas sekolah.')
    ]

    col_jurn = [Inches(0.4), Inches(1.1), Inches(1.2), Inches(1.8), Inches(1.3), Inches(1.0)]
    for r_idx, data_row in enumerate(jurn_data):
        row = jurn_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_jurn[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 80, 80, 70, 70)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==================== BAGIAN 3: DOKUMENTASI ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('C. LEMBAR DOKUMENTASI PENGUATAN KARAKTER SPIRITUAL & MORAL')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_after = Pt(4)
    r_bx = p_box.add_run('Lampiran Foto Kegiatan Pembiasaan Karakter (Tempelkan Foto Sekolah pada Kotak di Bawah):')
    r_bx.font.size = Pt(9)
    r_bx.italic = True

    photo_table = doc.add_table(rows=1, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')
    
    cell_p1 = photo_table.rows[0].cells[0]
    cell_p2 = photo_table.rows[0].cells[1]
    cell_p1.width = Inches(3.4)
    cell_p2.width = Inches(3.4)
    set_cell_background(cell_p1, 'ECEFF1')
    set_cell_background(cell_p2, 'ECEFF1')
    set_cell_margins(cell_p1, 240, 240, 100, 100)
    set_cell_margins(cell_p2, 240, 240, 100, 100)

    p_p1 = cell_p1.paragraphs[0]
    p_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p1.add_run('[ TEMPELKAN FOTO PEMBIASAAN DOA / TIKRAR SPIRITUAL DI SINI ]\n\nKeterangan: Kegiatan pembiasaan berdoa bersama dan penguatan akhlak mulia sebelum KBM.').font.size = Pt(8.5)

    p_p2 = cell_p2.paragraphs[0]
    p_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p2.add_run('[ TEMPELKAN FOTO DISKUSI ETIKA / MINDFULNESS DI SINI ]\n\nKeterangan: Siswa berlatih regulasi emosi positif dan diskusi budi pekerti dalam kelompok belajar.').font.size = Pt(8.5)

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
    p0.add_run('Guru PAI / Budi Pekerti,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Guru Kelas / Wali Kelas,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_19_Penguatan_Spiritual_Mental_Moral_Etika.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_19:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_19()
