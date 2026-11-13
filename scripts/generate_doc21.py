import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_21():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 21')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: DUKUNGAN TENDIK MENCIPTAKAN LINGKUNGAN AMAN DAN NYAMAN SECARA FISIK DAN PSIKIS\n(Berkas Lengkap: Program Kerja Tendik, Buku Log/Laporan Pelaksanaan Tugas, & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: PROGRAM KERJA TENDIK ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: PROGRAM KERJA TENAGA KEPENDIDIKAN (TENDIK) LINGKUNGAN RAMAH ANAK')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    prog_table = doc.add_table(rows=6, cols=2)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prog_table)

    prog_data = [
        ('1. Dasar Kebijakan', 'Surat Keputusan Kepala Sekolah Nomor: 421.2/028/SK-Tendik/2026 tentang Pembagian Tugas Tenaga Kependidikan dalam Menciptakan Lingkungan Satuan Pendidikan yang Aman, Sehat, Ramah Anak, dan Nyaman.'),
        ('2. Tujuan Program', '1. Mewujudkan perlindungan fisik bagi peserta didik dari ancaman bahaya luar, kecelakaan sarana, dan gangguan lingkungan.\n2. Menciptakan kenyamanan psikologis melalui budaya pelayanan ramah (5S: Senyum, Salam, Sapa, Sopan, Santun) oleh seluruh staf pendukung sekolah.'),
        ('3. Tugas Petugas Keamanan (Satpam)', 
         'a. Pukul 06.15 - 07.15 WIB: Penjagaan gerbang masuk, membantu penyeberangan jalan peserta didik, dan menerapkan salam 5S.\n'
         'b. Pukul 07.15 - 12.00 WIB: Mengunci gerbang utama, sterilisasi area dalam sekolah dari orang asing/pedagang tidak berizin, pemeriksaan identitas dan buku tamu untuk setiap tamu luar.\n'
         'c. Pukul 12.00 - 13.30 WIB: Pengamanan area penjemputan siswa, memastikan anak pulang bersama orang tua/wali resmi.'),
        ('4. Tugas Petugas Kebersihan Lingkungan', 
         'a. Sanitasi Toilet & Wastafel: Pembersihan berkala 3 kali sehari (sebelum KBM, setelah istirahat, dan setelah jam pulang) dengan disinfektan ramah anak.\n'
         'b. Pengelolaan Sampah: Pengosongan tempat sampah kelas pukul 06.30 WIB dan pemilahan sampah organik/anorganik di TPS sekolah.\n'
         'c. Keamanan Lantai: Memastikan tidak ada genangan air licin di koridor dan tangga, pemasangan tanda peringatan "Awas Licin" saat mengepel.'),
        ('5. Tugas Tenaga Administrasi (Tata Usaha)', 
         'a. Pelayanan Ramah Siswa: Membantu kebutuhan administrasi murid dengan tutur bahasa santun tanpa intimidasi.\n'
         'b. Pengamanan Arsip & Data: Mengamankan dokumen buku induk, biodata siswa, dan nilai di ruang arsip terkunci.'),
        ('6. Evaluasi Kinerja Tendik', 'Rapat koordinasi dan evaluasi mingguan bersama Kepala Sekolah dan Urusan Sarana Prasarana setiap hari Jumat siang.')
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

    # ==================== BAGIAN 2: BUKU LOG / LAPORAN TUGAS HARIAN ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: BUKU LOG PELAKSANAAN TUGAS HARIAN TENDIK (KEAMANAN & KEBERSIHAN)')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    log_table = doc.add_table(rows=6, cols=6)
    log_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(log_table)

    headers_log = ['Hari / Tanggal', 'Waktu Pelaksanaan', 'Sektor / Area Kerja', 'Uraian Aktivitas Pelayanan Tendik', 'Temuan / Status Kondisi', 'Petugas (Paraf)']
    header_row = log_table.rows[0]
    for idx, text in enumerate(headers_log):
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

    log_rows = [
        ('Senin, 14/09/2026', '06.15 - 07.15 WIB', 'Gerbang Utama & Penyeberangan', 'Menyambut siswa dengan 5S, menyeberangkan siswa, dan mengatur kelancaran antrean kendaraan jemputan.', 'Aman, tertib, tidak ada insiden.', 'Satpam\n(Bpk. Rohadi)'),
        ('Senin, 14/09/2026', '09.30 - 10.00 WIB', 'Toilet Siswa & Area Cuci Tangan', 'Pembersihan lantai toilet pasca jam istirahat pertama, pengisian ulang sabun cuci tangan cair di wastafel.', 'Lantai kering, wastafel bersih mengalir.', 'Kebersihan\n(Ibu Warsiti)'),
        ('Rabu, 23/09/2026', '10.15 - 11.30 WIB', 'Area Koridor & Halaman Belakang', 'Patroli lingkungan sekitar kelas saat KBM berlangsung, memastikan tidak ada orang asing masuk.', 'Kondusif, gerbang terkunci rapat.', 'Satpam\n(Bpk. Rohadi)'),
        ('Jumat, 02/10/2026', '06.30 - 08.00 WIB', 'Saluran Air & Taman Sekolah', 'Pembersihan daun kering dan pengecekan saluran air pembuangan dari sampah plastik.', 'Saluran lancar, tidak ada jentik nyamuk.', 'Kebersihan\n(Bpk. Slamet)'),
        ('Rabu, 21/10/2026', '12.00 - 13.00 WIB', 'Zona Penjemputan Gerbang Depan', 'Pengawasan siswa kelas rendah (I-III) yang menunggu jemputan di dalam batas aman gerbang.', 'Seluruh siswa pulang bersama keluarga terverifikasi.', 'Satpam & Guru Piket')
    ]

    col_l = [Inches(1.1), Inches(1.1), Inches(1.3), Inches(1.8), Inches(1.2), Inches(0.9)]
    for r_idx, data_row in enumerate(log_rows):
        row = log_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_l[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 70, 70, 60, 60)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 1, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 3: LAPORAN EVALUASI & FEEDBACK LINGKUNGAN ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: LAPORAN EVALUASI LINGKUNGAN AMAN DAN NYAMAN TENDIK')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    eval_table = doc.add_table(rows=5, cols=2)
    eval_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(eval_table)

    eval_data = [
        ('1. Periode Evaluasi', 'Bulan September - Oktober 2026 (Semester Ganjil TA 2026/2027)'),
        ('2. Capaian Keamanan Fisik', '100% tamu terdata rapi di buku pos satpam. Tidak ditemukan kasus orang asing mencurigakan masuk pekarangan sekolah. Seluruh pintu gerbang tertutup aman saat jam belajar.'),
        ('3. Capaian Kenyamanan Psikis', 'Penerapan senyum ramah petugas satpam saat menyambut anak pagi hari menumbuhkan rasa percaya diri anak masuk sekolah tanpa rasa takut atau cemas.'),
        ('4. Capaian Kebersihan & Sanitasi', 'Tersedia sabun cuci tangan aktif di seluruh wastafel kelas. Toilet siswa bebas dari bau menyengat dan lantai selalu dalam kondisi kering terjaga.'),
        ('5. Saran & Rencana Tindak Lanjut', 'Penambahan kanopi peneduh di zona tunggu penjemputan siswa agar tidak kehujanan/kepanasan saat musim penghujan.')
    ]

    for row_idx, (col0, col1) in enumerate(eval_data):
        row = eval_table.rows[row_idx]
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

    # ==================== BAGIAN 4: DOKUMENTASI FOTO ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: DOKUMENTASI KEGIATAN TENDIK RAMAH ANAK')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: PENYAMBUTAN SISWA 5S OLEH SATPAM ]\n\nKeterangan: Petugas keamanan menyambut siswa dengan senyum ramah di gerbang depan dan menyeberangkan jalan.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: PENCATATAN BUKU TAMU SECARA KETAT ]\n\nKeterangan: Verifikasi identitas dan penyerahan tanda pengenal tamu untuk setiap pengunjung luar.', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: SANITASI DAN KEBERSIHAN TOILET/WASTAFEL ]\n\nKeterangan: Petugas kebersihan melakukan pembersihan rutin wastafel dan toilet sekolah.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: PENGAMANAN AREA PENJEMPUTAN PULANG ]\n\nKeterangan: Penjagaan gerbang teratur saat kepulangan siswa untuk memastikan keselamatan seluruh murid.', photo_table.rows[1].cells[1])
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
    p0.add_run('Koordinator Tenaga Kependidikan,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Urusan Sarana & Prasarana,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_21_Dukungan_Tendik_Lingkungan_Aman_Nyaman.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_21:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_21()
