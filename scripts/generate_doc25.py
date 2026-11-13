import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_25():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 25')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: PENDAMPINGAN WALI KELAS\n(Dokumen Komprehensif: Buku Administrasi Wali Kelas, Catatan Perkembangan Karakter Siswa, & Dokumentasi Pendampingan)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: BUKU ADMINISTRASI WALI KELAS ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: PROFIL DAN PROGRAM PENDAMPINGAN WALI KELAS')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    adm_table = doc.add_table(rows=6, cols=2)
    adm_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(adm_table)

    adm_data = [
        ('1. Identitas Rombongan Belajar', 'Kelas / Fase       : Kelas IV-A / Fase B\nTahun Ajaran       : 2026/2027 (Semester Ganjil)\nNama Wali Kelas    : [Nama Lengkap Wali Kelas, S.Pd]\nJumlah Peserta Didik: 28 Orang (14 Laki-laki, 14 Perempuan)'),
        ('2. Visi & Budaya Kelas', '"Menjadi Kelas Sahabat Juara: Santun dalam Tutur Kata, Tangguh dalam Belajar, Cerdas Mengelola Diri, dan Peduli terhadap Sesama."'),
        ('3. Program Pembiasaan Harian Wali Kelas', 
         '• Pukul 06.45 - 07.00: Salam Pagi & "Lingkaran Cerita" (mengecek kehadiran, kerapian seragam, dan kesiapan emosi siswa).\n'
         '• Pukul 07.00 - 07.15: Pembiasaan Doa Bersama, menyanyikan lagu Indonesia Raya, dan literasi membaca buku non-pelajaran 15 menit.\n'
         '• Pukul 09.30 - 10.00: Pendampingan makan bekal bergizi di ruang kelas dengan adab cuci tangan dan doa sebelum makan.\n'
         '• Pukul 12.30 - 12.45: Refleksi penutupan kelas "Satu Kata Kebaikan Hari Ini" sebelum kepulangan.'),
        ('4. Struktur Pengurus Kelas & Regu Piket', 
         'Ketua Kelas : Ahmad Fadil | Wakil : Citra Lestari | Sekretaris : Dimas Pratama | Bendahara : Anisa Zahra\n'
         'Regu Piket 7 Hari Berkarakter: Dibagi menjadi 5 kelompok kerja (Senin s.d Jumat) yang bertanggung jawab atas kebersihan sudut baca, papan tulis, dan tempat sampah kelas.'),
        ('5. Mekanisme Komunikasi Paguyuban Orang Tua', 'Grup komunikasi resmi WhatsApp Paguyuban Kelas IV-A untuk update berkala jurnal kegiatan, serta pertemuan tatap muka konsultasi individu setiap akhir bulan.'),
        ('6. Prinsip Pendampingan Ramah Anak', 'Mengutamakan penguatan positif (positive reinforcement), mendengarkan tanpa menghakimi, dan meniadakan hukuman fisik maupun verbal yang mempermalukan murid di depan kawan-kawannya.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(adm_data):
        row = adm_table.rows[row_idx]
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

    # ==================== BAGIAN 2: CATATAN PERKEMBANGAN KARAKTER SISWA ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: BUKU CATATAN PERKEMBANGAN KARAKTER & KEDISIPLINAN SISWA')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    p_desc2 = doc.add_paragraph()
    p_desc2.paragraph_format.space_after = Pt(3)
    r_d2 = p_desc2.add_run('Format Pemantauan Berkelanjutan Sikap Spiritual, Sosial, dan Kemandirian Belajar Siswa:')
    r_d2.font.size = Pt(8)
    r_d2.italic = True

    track_table = doc.add_table(rows=7, cols=6)
    track_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(track_table)

    track_headers = ['No', 'Nama Siswa / NIS', 'Perkembangan Sikap Spiritual (Ibadah/Syukur)', 'Perkembangan Sikap Sosial (Gotong Royong/Empati)', 'Kemandirian & Disiplin Belajar', 'Catatan Pembinaan & Apresiasi Wali Kelas']
    header_row = track_table.rows[0]
    for idx, text in enumerate(track_headers):
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

    track_rows = [
        ('1', 'Ahmad Fadillah\n(NIS: 23101)', 'Sangat tertib memimpin doa pagi, toleran terhadap teman non-muslim.', 'Memiliki jiwa kepemimpinan tinggi, selalu membantu kawan piket.', 'Mampu menyelesaikan tugas tepat waktu tanpa perlu diingatkan.', 'Diberikan apresiasi sebagai Ketua Kelas Teladan bulan September.'),
        ('2', 'Bella Cantika\n(NIS: 23102)', 'Rutin mengisi jurnal syukur harian sebelum pembelajaran usai.', 'Sangat santun dalam tutur kata, pembawa damai di kelompok belajar.', 'Rapi dalam merawat buku tulis dan perlengkapan sekolah.', 'Pertahankan sikap teladan dan kesantunannya.'),
        ('3', 'Dimas Prasetyo\n(NIS: 23103)', 'Perlu sedikit dorongan agar tidak tergesa-gesa saat berdoa.', 'Mulai aktif berbagi alat tulis dengan teman sebangku.', 'Semula sering lupa membawa buku paket, kini sudah tertib menyusun jadwal.', 'Ada peningkatan disiplin yang sangat pesat. Terus dimotivasi.'),
        ('4', 'Fathir Arkan\n(NIS: 23104)', 'Khusyuk dalam kegiatan keagamaan sekolah.', 'Sangat tanggap membantu merapikan kursi kelas yang berantakan.', 'Konsentrasi belajar meningkat tajam saat KBM proyek kelompok.', 'Sangat potensial menjadi penggerak kebersihan kelas.'),
        ('5', 'Ghea Salsabila\n(NIS: 23105)', 'Selalu tersenyum dan mengucap salam saat berpapasan dengan guru.', 'Peduli pada teman yang terlihat murung, aktif mengajak berbaur.', 'Kemandirian literasi tinggi, gemar membaca di pojok baca.', 'Menjadi Duta Membaca Kelas IV-A.'),
        ('6', 'Iqbal Ramadhan\n(NIS: 23106)', 'Disiplin waktu salat dhuha dan doa pembuka belajar.', 'Belajar mengendalikan nada bicara saat berdiskusi perbedaan pendapat.', 'Penyelesaian tugas harian konsisten meningkat.', 'Pendampingan dialog personal berjalan sangat efektif.')
    ]

    col_tr = [Inches(0.4), Inches(1.3), Inches(1.3), Inches(1.4), Inches(1.3), Inches(1.1)]
    for r_idx, data_row in enumerate(track_rows):
        row = track_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_tr[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 60, 60, 50, 50)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(7.8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 3: LAPORAN HOME VISIT / KONSULTASI WALI MURID ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: LAPORAN KUNJUNGAN RUMAH (HOME VISIT) & KONSULTASI ORANG TUA')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    hv_table = doc.add_table(rows=6, cols=2)
    hv_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(hv_table)

    hv_data = [
        ('1. Identitas Peserta Didik', 'Nama: Dimas Prasetyo | Kelas: IV-A | Alamat: Jl. Pangeran Drajat No. 14, Kesambi, Cirebon'),
        ('2. Waktu & Petugas Pelaksana', 'Hari / Tanggal : Kamis, 08 Oktober 2026 (Pukul 15.30 - 16.45 WIB)\nPetugas        : Wali Kelas IV-A bersama Guru Pendamping'),
        ('3. Alasan Kunjungan', 'Membangun sinergi pola pembiasaan bangun pagi dan pendampingan belajar di rumah setelah siswa sempat terlambat 3 hari berturut-turut pada pekan sebelumnya.'),
        ('4. Hasil Wawancara & Observasi Rumah', 
         '• Orang tua menyambut dengan sangat hangat dan terbuka.\n'
         '• Diketahui bahwa keterlambatan disebabkan anak sering diajak tidur larut malam (>22.30 WIB) karena menunggu orang tua pulang dari tempat usaha dagang.\n'
         '• Suasana rumah kondusif, namun belum ada jadwal tertulis batas jam tidur anak.'),
        ('5. Solusi & Kesepakatan Bersama', 
         'a. Orang tua menyepakati aturan batas tidur maksimal pukul 21.00 WIB didampingi oleh nenek bila orang tua belum pulang.\n'
         'b. Wali kelas memberikan Buku Jurnal Pembiasaan Tidur Tepat Waktu untuk ditandatangani orang tua setiap pagi.\n'
         'c. Memberikan motivasi kepada ananda Dimas agar lebih mandiri menyiapkan tas sekolah pada malam hari.'),
        ('6. Progres Pasca Kunjungan', 'Dalam 2 pekan pasca home visit, ananda Dimas selalu hadir pukul 06.40 WIB (tidak pernah terlambat lagi) dan tampak jauh lebih segar berenergi di kelas.')
    ]

    for row_idx, (col0, col1) in enumerate(hv_data):
        row = hv_table.rows[row_idx]
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

    # ==================== BAGIAN 4: DOKUMENTASI FOTO PENDAMPINGAN WALI KELAS ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: DOKUMENTASI KEGIATAN PENDAMPINGAN WALI KELAS')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: PENDAMPINGAN LITERASI & SALAM PAGI DI KELAS ]\n\nKeterangan: Wali kelas menyambut kedatangan peserta didik, memeriksa kesiapan belajar, dan membimbing kegiatan literasi pagi.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: REFLEKSI AKHIR PEKAN & APRESIASI SISWA ]\n\nKeterangan: Kegiatan evaluasi mingguan kelas dan penyematan bintang karakter pada "Pohon Kebaikan Kelas".', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: KUNJUNGAN RUMAH (HOME VISIT) WALI KELAS ]\n\nKeterangan: Silaturahmi edukatif wali kelas ke kediaman orang tua siswa untuk menyelaraskan pola asuh dan pembiasaan anak.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: KONSULTASI TATAP MUKA DENGAN ORANG TUA ]\n\nKeterangan: Dialog berkala wali kelas bersama orang tua murid mengenai perkembangan psikososial dan akademis anak.', photo_table.rows[1].cells[1])
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
    p0.add_run('Wali Kelas IV-A,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Koordinator Kesiswaan,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_25_Pendampingan_Wali_Kelas.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_25:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_25()
