import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_27():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 27')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: KEAMANAN LINGKUNGAN DAN PERLINDUNGAN DATA PRIBADI MURID\n(Dokumen Komprehensif: Standar Operasional Prosedur (SOP), Program Pengamanan Data, & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: STANDAR OPERASIONAL PROSEDUR (SOP) DATA PRIBADI ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: STANDAR OPERASIONAL PROSEDUR (SOP) PERLINDUNGAN DATA PRIBADI MURID')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    sop_table = doc.add_table(rows=6, cols=2)
    sop_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sop_table)

    sop_data = [
        ('1. Dasar Hukum & Dokumen', 'Nomor Dokumen: SOP/TATA-USAHA/PDP/01/2026\nDasar Regulasi: Undang-Undang No. 27 Tahun 2022 tentang Pelindungan Data Pribadi (UU PDP) dan Permendikbudristek No. 46 Tahun 2023.'),
        ('2. Tujuan & Sasaran', 'Menjamin kerahasiaan, integritas, dan keamanan seluruh data pribadi peserta didik (baik data fisik dokumen maupun digital di aplikasi Dapodik/Emis) dari kebocoran, penyalahgunaan pihak ketiga, eksploitasi, serta ancaman kejahatan siber.'),
        ('3. Ruang Lingkup Data Pribadi', 
         '• Data Umum: Nama lengkap, jenis kelamin, kewarganegaraan, agama.\n'
         '• Data Spesifik/Sensitif: Nomor Induk Kependudukan (NIK), Nomor Kartu Keluarga, data riwayat medis/kesehatan, rekam psikologis/konseling, nomor rekening beasiswa/PIP, serta alamat domisili lengkap.'),
        ('4. Prosedur Pengamanan Fisik Dokumen', 
         'a. Seluruh berkas fisik (fotokopi akta lahir, KK, ijazah, rapor, buku induk) wajib disimpan dalam lemari arsip terkunci di Ruang Tata Usaha.\n'
         'b. Kunci lemari arsip hanya dipegang oleh Petugas Arsiparis dan Kepala Tata Usaha.\n'
         'c. Peminjaman dokumen fisik untuk verifikasi wajib dicatat dalam Buku Ekspedisi Peminjaman Berkas.'),
        ('5. Prosedur Pengamanan Digital (Siber)', 
         'a. Komputer server dan operator Dapodik wajib menggunakan password berkala (kombinasi huruf, angka, simbol) dan tidak boleh diakses oleh pihak selain operator resmi.\n'
         'b. Dilarang keras mempublikasikan data identitas pribadi lengkap (NIK, nama ibu kandung, nomor kontak orang tua, nilai rapor) ke media sosial atau grup WhatsApp terbuka.\n'
         'c. Publikasi foto/video kegiatan sekolah yang melibatkan siswa wajib memperhatikan etika dan tidak mencantumkan identitas sensitif anak tanpa persetujuan tertulis.'),
        ('6. Penanganan Insiden Kebocoran Data', 'Apabila terjadi dugaan kebocoran data, tim sekolah wajib melakukan investigasi dalam 1x24 jam, melaporkan kepada Dinas Pendidikan Kota Cirebon, serta memberitahukan orang tua/wali murid terkait.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(sop_data):
        row = sop_table.rows[row_idx]
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

    # ==================== BAGIAN 2: PROGRAM KERJA KEAMANAN DATA & LINGKUNGAN ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: PROGRAM KERJA PENGELOLAAN ARSIP DAN LITERASI KEAMANAN DIGITAL')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    prog_table = doc.add_table(rows=6, cols=2)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prog_table)

    prog_data = [
        ('1. Nama Program', 'Program Terpadu Pengamanan Arsip Satuan Pendidikan & Edukasi Keamanan Siber Ramah Anak Tahun 2026/2027.'),
        ('2. Penanggung Jawab', 'Kepala Sekolah (Penanggung Jawab Utama), Kepala Tata Usaha, Operator Dapodik, & Seluruh Guru Kelas.'),
        ('3. Agenda Digitalisasi Aman', 
         '• Backup berkala basis data Dapodik secara offline ke external hard drive terenkripsi setiap akhir pekan.\n'
         '• Pemasangan firewall dan antivirus berlisensi pada seluruh perangkat komputer sekolah yang terhubung jaringan internet.\n'
         '• Pembaruan sistem operasi dan penutupan port akses tidak sah pada router WiFi sekolah.'),
        ('4. Pakta Integritas Kerahasiaan', 'Seluruh pendidik dan tenaga kependidikan menandatangani Pakta Integritas Kerahasiaan Data Pribadi Siswa yang melarang pembagian data murid kepada pihak promosi komersial di luar sekolah.'),
        ('5. Edukasi Literasi Digital Murid', 'Pemberian materi pembelajaran tematik mengenai "Jaga Diri di Dunia Maya": Melarang siswa membagikan alamat rumah, nomor telepon, atau kata sandi akun belajar kepada orang yang baru dikenal di internet.'),
        ('6. Audit Keamanan Arsip', 'Pemeriksaan fisik lemari arsip dan audit log akses Dapodik dilakukan setiap bulan oleh Kepala Sekolah.')
    ]

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

    # ==================== BAGIAN 3: PAKTA INTEGRITAS & LOG PENGAMANAN DATA ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: FORMAT SURAT PAKTA INTEGRITAS PERLINDUNGAN DATA')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    p_pakta = doc.add_paragraph()
    p_pakta.paragraph_format.space_after = Pt(4)
    r_pk = p_pakta.add_run(
        'SURAT PERNYATAAN / PAKTA INTEGRITAS KERAHASIAN DATA PRIBADI SISWA\n\n'
        'Yang bertanda tangan di bawah ini:\n'
        'Nama Lengkap   : [Nama Operator / Tenaga Kependidikan]\n'
        'NIP / NUPTK    : [.................................................]\n'
        'Jabatan        : Operator Sekolah / Staf Tata Usaha\n\n'
        'Menyatakan dengan sungguh-sungguh bahwa saya berkomitmen:\n'
        '1. Menjaga kerahasiaan seluruh dokumen fisik dan data digital peserta didik SD Negeri [Nama Sekolah].\n'
        '2. Tidak akan menyalin, membagikan, menyebarluaskan, atau memperjualbelikan data pribadi murid (NIK, alamat rumah, nomor HP, foto identitas) kepada pihak manapun tanpa izin sah tertulis dari Kepala Sekolah dan orang tua/wali murid.\n'
        '3. Menggunakan data peserta didik semata-mata untuk kepentingan kedinasan pendidikan resmi.\n'
        '4. Bersedia menerima sanksi disiplin dan tindakan hukum yang berlaku apabila terbukti melakukan kelalaian atau pelanggaran atas perlindungan data pribadi ini.'
    )
    r_pk.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 4: DOKUMENTASI FOTO ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: DOKUMENTASI SISTEM KEAMANAN LINGKUNGAN & DATA PRIBADI')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: LEMARI ARSIP BERKAS BUKU INDUK TERKUNCI ]\n\nKeterangan: Penyimpanan dokumen fisik raport, buku induk, dan akta kelahiran siswa di lemari besi terkunci di ruang TU.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: PERANGKAT KOMPUTER DAPODIK BERPENGAMAN PASSWORD ]\n\nKeterangan: Komputer pengelola data siswa dengan proteksi login berpassword ganda dan antivirus aktif.', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: PENANDATANGANAN PAKTA INTEGRITAS TENDIK ]\n\nKeterangan: Penandatanganan komitmen perlindungan data pribadi murid oleh seluruh tenaga kependidikan dan guru.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: SOSIALISASI LITERASI KEAMANAN DIGITAL SISWA ]\n\nKeterangan: Edukasi di kelas tentang pentingnya menjaga kerahasiaan data pribadi saat berselancar di dunia maya.', photo_table.rows[1].cells[1])
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
    p0.add_run('Kepala Tata Usaha / Operator,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Koordinator Kurikulum,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_27_Keamanan_Lingkungan_dan_Data_Pribadi.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_27:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_27()
