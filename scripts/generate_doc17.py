import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_17():
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

    # Garis Pembatas Kop
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # JUDUL DOKUMEN
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI (POIN 17)')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: KEMITRAAN DENGAN ORANG TUA / PEMANGKU KEPENTINGAN\n(Lampiran: Notulen Rapat, Naskah MoU Kemitraan, & Lembar Dokumentasi/Daftar Hadir)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: NOTULEN RAPAT KEMITRAAN ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('A. NOTULEN RAPAT KEMITRAAN ORANG TUA & KOMITE SEKOLAH')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    notulen_table = doc.add_table(rows=6, cols=2)
    notulen_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(notulen_table)

    notulen_items = [
        ('1. Hari / Tanggal', 'Sabtu, 22 Agustus 2026 (Pukul 08.30 – 11.30 WIB)'),
        ('2. Tempat Kegiatan', 'Ruang Aula Serbaguna SD Negeri [Nama Sekolah]'),
        ('3. Agenda Rapat', 'Sosialisasi Program Jabar Beraksi (Penguatan Pendidikan Karakter), Sinergi Pembiasaan Positif Siswa di Rumah, dan Pemantauan Gizi/Waktu Tidur Anak.'),
        ('4. Peserta Rapat', 'Kepala Sekolah, Ketua Komite Sekolah, Seluruh Wali Kelas I - VI, Pengurus Paguyuban Kelas, dan 85 Orang Perwakilan Orang Tua/Wali Murid.'),
        ('5. Poin Kesepakatan & Pembahasan', 
         'a. Pembiasaan Karakter di Rumah:\n'
         '   - Orang tua sepakat mendampingi anak untuk bangun pagi (sebelum pukul 05.00 WIB) dan melaksanakan ibadah sesuai keyakinan masing-masing.\n'
         '   - Pembatasan penggunaan gadget/layar HP maksimal 1 jam per hari saat hari sekolah serta memastikan anak tidur malam sebelum pukul 21.00 WIB.\n'
         'b. Pengisian Jurnal Karakter Siswa:\n'
         '   - Orang tua menandatangani buku penghubung/jurnal pembiasaan karakter setiap akhir pekan.\n'
         'c. Program Bekal Bergizi Sehat:\n'
         '   - Mendukung Gerakan Sarapan Sehat bersama di sekolah setiap hari Rabu (membawa makanan bergizi seimbang dari rumah, membatasi makanan instan).\n'
         'd. Komitmen Komunikasi Terbuka:\n'
         '   - Menyepakati grup komunikasi kelas ramah anak yang bebas hoaks dan fokus pada pendampingan belajar.'),
        ('6. Tindak Lanjut', 'Pembentukan jadwal parenting kelas triwulanan dan evaluasi keterlaksanaan buku jurnal pembiasaan pada rapat tengah semester.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(notulen_items):
        row = notulen_table.rows[row_idx]
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

    # ==================== BAGIAN 2: NASKAH MoU KEMITRAAN ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('B. NASKAH NOTA KESEPAHAMAN (MoU) KEMITRAAN PEMANGKU KEPENTINGAN')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    mou_table = doc.add_table(rows=6, cols=2)
    mou_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(mou_table)

    mou_items = [
        ('1. Identitas MoU', 'Nomor Pihak I : 421.2/045/SD-MoU/2026\nNomor Pihak II: 440/128/PKM-KES/2026\nTentang: Program Sekolah Sehat, Pembinaan Karakter Fisik-Psikis, & Layanan Kesehatan Siswa'),
        ('2. Para Pihak', 'PIHAK PERTAMA: Kepala SD Negeri [Nama Sekolah] (Mewakili Satuan Pendidikan)\nPIHAK KEDUA  : Kepala UPTD Puskesmas [Kesambi/Harjamukti] (Mewakili Fasilitas Layanan Kesehatan)'),
        ('3. Maksud & Tujuan', 'Mewujudkan lingkungan satuan pendidikan yang sehat, aman, dan berkarakter melalui pemantauan status kesehatan fisik, kebersihan lingkungan, serta edukasi perilaku hidup bersih dan sehat (PHBS).'),
        ('4. Ruang Lingkup Kerja Sama', 
         'a. Pemeriksaan berkala kesehatan umum, gigi, mata, dan telinga peserta didik (skrining berkala per semester).\n'
         'b. Edukasi gizi seimbang, pencegahan anemia, dan pemantauan jajanan kantin sehat.\n'
         'c. Penyuluhan kesehatan reproduksi usia dini dan kesehatan mental/emosional ramah anak.\n'
         'd. Pelatihan pertolongan pertama pada kecelakaan (P3K) bagi Guru Pembina UKS dan Kader Tiwisada.'),
        ('5. Masa Berlaku', 'Perjanjian ini berlaku selama 3 (tiga) tahun terhitung sejak tanggal ditetapkan dan dapat diperpanjang atas kesepakatan bersama.'),
        ('6. Hak & Kewajiban', 'Pihak Pertama menyediakan fasilitas dan jadwal kegiatan; Pihak Kedua menyediakan tenaga medis, bahan penyuluhan, dan laporan hasil pemeriksaan kesehatan berkala.')
    ]

    for row_idx, (col0, col1) in enumerate(mou_items):
        row = mou_table.rows[row_idx]
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

    # ==================== BAGIAN 3: DOKUMENTASI & DAFTAR HADIR ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('C. LEMBAR DAFTAR HADIR DAN DOKUMENTASI KEGIATAN KEMITRAAN')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    p_dh = doc.add_paragraph()
    p_dh.paragraph_format.space_after = Pt(4)
    r_dh = p_dh.add_run('Sampel Daftar Hadir Pertemuan Paguyuban Kelas / Sosialisasi Pemangku Kepentingan:')
    r_dh.font.size = Pt(9)
    r_dh.italic = True

    att_table = doc.add_table(rows=6, cols=5)
    att_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(att_table)

    att_headers = ['No', 'Nama Lengkap Orang Tua / Tamu', 'Orang Tua dari Siswa', 'Kelas', 'Tanda Tangan']
    header_row = att_table.rows[0]
    for idx, text in enumerate(att_headers):
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

    att_samples = [
        ('1', 'Bapak Suryadi', 'Ahmad Fadil', 'Kelas IV-A', '1. ....................'),
        ('2', 'Ibu Nurhayati', 'Siti Zahra', 'Kelas IV-A', '2. ....................'),
        ('3', 'Bapak Hendra Gunawan', 'Dimas Pratama', 'Kelas V-B', '3. ....................'),
        ('4', 'Ibu Ratna Dewi', 'Anisa Putri', 'Kelas V-B', '4. ....................'),
        ('5', 'dr. Hj. Nuraeni, M.Kes (Puskesmas)', 'Narasumber Kesehatan', 'Mitra Kerja', '5. ....................')
    ]

    col_att = [Inches(0.4), Inches(2.2), Inches(1.6), Inches(1.1), Inches(1.5)]
    for r_idx, data_row in enumerate(att_samples):
        row = att_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_att[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 80, 80, 70, 70)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)

    # Frame Foto Dokumentasi
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_after = Pt(4)
    r_bx = p_box.add_run('Lampiran Foto Kegiatan Kemitraan (Tempelkan Foto Riil Sekolah pada Kotak di Bawah):')
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
    set_cell_margins(cell_p1, 250, 250, 100, 100)
    set_cell_margins(cell_p2, 250, 250, 100, 100)

    p_p1 = cell_p1.paragraphs[0]
    p_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p1.add_run('[ TEMPELKAN FOTO RAPAT PAGUYUBAN KELAS / KOMITE DI SINI ]\n\nKeterangan: Pertemuan sinergi pembiasaan karakter peserta didik bersama orang tua murid.').font.size = Pt(8.5)

    p_p2 = cell_p2.paragraphs[0]
    p_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p2.add_run('[ TEMPELKAN FOTO PENANDATANGANAN MoU / KEGIATAN MITRA DI SINI ]\n\nKeterangan: Pemeriksaan kesehatan & pembinaan perilaku hidup sehat oleh tenaga Puskesmas.').font.size = Pt(8.5)

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
    p0.add_run('Ketua Komite Sekolah,\n\n\n\n\n( ......................................... )').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Notulis / Perwakilan Guru,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_17_Kemitraan_Orang_Tua_dan_Pemangku_Kepentingan.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_17:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_17()
