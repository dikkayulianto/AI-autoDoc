import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_23():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 23')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: KEAMANAN FISIK SEKOLAH BERDASARKAN SARANA DAN PRASARANA\n(Dokumen Komprehensif: Pedoman K3 Sarpras, Checklist Audit Kelayakan, Laporan Pemeliharaan & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: PEDOMAN & STANDAR KEAMANAN SARPRAS ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: PEDOMAN STANDAR KEAMANAN SARANA & PRASARANA SEKOLAH RAMAH ANAK')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    ped_table = doc.add_table(rows=6, cols=2)
    ped_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ped_table)

    ped_data = [
        ('1. Dasar Regulasi', 'Permendikbudristek No. 46 Tahun 2023 tentang Pencegahan dan Penanganan Kekerasan di Satuan Pendidikan serta Permendiknas No. 24 Tahun 2007 tentang Standar Sarana dan Prasarana Sekolah Dasar.'),
        ('2. Prinsip Sarpras Aman', 
         '• Aksesibilitas & Ramah Anak: Bangunan tidak memiliki sudut tajam berbahaya, lantai tidak licin, serta memiliki akses yang aman bagi seluruh anak termasuk disabilitas.\n'
         '• Proteksi Keselamatan (Zero Accident): Setiap potensi bahaya fisik (arus listrik, api, benda jatuh, genangan air) diminimalisasi secara terstruktur.\n'
         '• Mitigasi Bencana: Jalur evakuasi bebas rintangan dan titik kumpul (assembly point) tertata jelas.'),
        ('3. Standar Kelistrikan', 'Semua saklar dan stop kontak berada pada ketinggian minimal 1,5 meter dari lantai atau dilengkapi penutup pengaman khusus anak (safety plug cover). Jalur kabel terlindung pipa conduit.'),
        ('4. Standar Struktur & Kaca', 'Pagar keliling sekolah berketinggian minimal 1,8 meter dengan pintu gerbang terkunci saat jam KBM. Jendela kelas menggunakan engsel kuat dan kaca dengan film pengaman.'),
        ('5. Sanitasi & Higienitas', 'Perbandingan toilet siswa terpisah putra dan putri, sirkulasi udara optimal, lantai berkeramik anti-selip, serta tersedia kran air bersih mengalir setiap hari.'),
        ('6. Jadwal Audit Berkala', 'Pemeriksaan rutin mingguan oleh Petugas Sarpras dan Audit Formal Triwulanan bersama Tim Manajemen BOSP/RKAS Sekolah.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(ped_data):
        row = ped_table.rows[row_idx]
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

    # ==================== BAGIAN 2: CHECKLIST AUDIT SARPRAS LENGKAP ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: INSTRUMEN CHECKLIST INSPEKSI FISIK SARANA & PRASARANA')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    chk_table = doc.add_table(rows=11, cols=6)
    chk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(chk_table)

    chk_headers = ['No', 'Komponen / Area Sarpras', 'Indikator Standar Keamanan', 'Kondisi Riil (Baik/Cukup/Rusak)', 'Tingkat Risiko', 'Tindakan Mitigasi / Perbaikan']
    header_row = chk_table.rows[0]
    for idx, text in enumerate(chk_headers):
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

    chk_rows = [
        ('1', 'Pagar & Gerbang Sekolah', 'Tinggi >1,8 m, engsel kokoh, gerbang terkunci saat KBM.', 'Baik (Kuat & Berfungsi)', 'Rendah', 'Pelumasan berkala pada engsel dan slot kunci gerbang.'),
        ('2', 'Instalasi Listrik & Saklar', 'Kabel rapi dalam pipa proteksi, saklar tertutup aman.', 'Baik (Standar PLN)', 'Rendah', 'Inspeksi rutin panel MCB setiap akhir bulan.'),
        ('3', 'Kusen, Pintu & Jendela', 'Buka tutup lancar, tidak ada kaca pecah/retak berisiko.', 'Baik (Terawat)', 'Rendah', 'Pemasangan penahan pintu (door stopper) agar tidak terbanting angin.'),
        ('4', 'Lantai Koridor & Selasar', 'Permukaan rata, tidak licin, tidak ada ubin pecah.', 'Baik (Kering & Rapi)', 'Rendah', 'Pemasangan rambu "Awas Licin" saat proses pembersihan lantai.'),
        ('5', 'Plafon & Atap Gedung', 'Rangka kokoh, tidak ada kebocoran air atau plafon rapuh.', 'Baik (Plafon Gypsum Aman)', 'Rendah', 'Pengecekan genteng menjelang musim penghujan.'),
        ('6', 'Toilet Siswa (Putra/Putri)', 'Kran air lancar, pintu terkunci dari dalam, lantai anti-selip.', 'Baik (Sanitasi Bersih)', 'Rendah', 'Pembersihan 3x sehari dan penggantian lampu secara cepat bila redup.'),
        ('7', 'Alat Pemadam Api (APAR)', 'Tekanan tabung di zona hijau, belum kadaluarsa, mudah dijangkau.', 'Baik (Siap Pakai, 2 Unit)', 'Rendah', 'Simulasi cara penggunaan APAR untuk tenaga pendidik dan tendik.'),
        ('8', 'Jalur Evakuasi & Titik Kumpul', 'Tanda penunjuk arah jelas, jalur steril tanpa barang bertumpuk.', 'Baik (Titik Kumpul di Lapangan)', 'Rendah', 'Pengecatan ulang marka titik kumpul lapangan yang mulai pudar.'),
        ('9', 'Kotak P3K & Perlengkapan', 'Terpasang di UKS & tiap lorong, obat standar lengkap & berizin.', 'Baik (Isi Lengkap)', 'Rendah', 'Pengecekan tanggal kadaluarsa obat antiseptik dan kasa secara berkala.'),
        ('10', 'Lapangan & Sarana Olahraga', 'Permukaan rata bebas kerikil lepas, tiang net stabil terpasang.', 'Baik (Aman untuk Siswa)', 'Rendah', 'Pembersihan rutin lapangan sebelum jam pembelajaran PJOK dimulai.')
    ]

    col_chk = [Inches(0.4), Inches(1.3), Inches(1.6), Inches(1.1), Inches(0.8), Inches(1.6)]
    for r_idx, data_row in enumerate(chk_rows):
        row = chk_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_chk[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 60, 60, 50, 50)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(7.8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 3: LAPORAN HASIL AUDIT & REKOMENDASI ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: LAPORAN AUDIT KELAYAKAN DAN REKOMENDASI TINDAKAN')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    rep_table = doc.add_table(rows=5, cols=2)
    rep_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rep_table)

    rep_data = [
        ('1. Tanggal Pelaksanaan Audit', 'Senin - Rabu, 19 - 21 Oktober 2026 (Audit Semester Ganjil TA 2026/2027)'),
        ('2. Tim Auditor Sarpras', 'Koordinator Sarpras, Petugas K3 Sekolah, Perwakilan Komite Sekolah, & Koordinator Tendik.'),
        ('3. Ringkasan Hasil Pemeriksaan', 'Dari 10 komponen utama yang diinspeksi, 100% berada dalam kondisi aman dan layak pakai bagi peserta didik. Tidak ditemukan kerusakan berat yang berpotensi membahayakan jiwa. Beberapa area membutuhkan pemeliharaan preventif ringan seperti pengecatan marka titik kumpul dan pengecekan talang air.'),
        ('4. Rekomendasi Pemeliharaan', 
         'a. Mengalokasikan dana pemeliharaan rutin pada pos RKAS untuk pengadaan penutup stop kontak cadangan dan isi ulang obat P3K.\n'
         'b. Melakukan pengecekan berkala terhadap kekuatan tiang kanopi parkir dan selasar.\n'
         'c. Memperbarui denah jalur evakuasi di setiap pintu masuk ruang kelas.'),
        ('5. Kesimpulan Kelayakan', 'STATUS: MEMENUHI STANDAR KESELAMATAN SATUAN PENDIDIKAN RAMAH ANAK (KATEGORI SANGAT BAIK)')
    ]

    for row_idx, (col0, col1) in enumerate(rep_data):
        row = rep_table.rows[row_idx]
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

    # ==================== BAGIAN 4: RENCANA TINDAK LANJUT RKAS ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: RENCANA TINDAK LANJUT PEMELIHARAAN SARPRAS (ALOKASI RKAS/BOS)')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    rkas_table = doc.add_table(rows=5, cols=5)
    rkas_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rkas_table)

    rkas_headers = ['No', 'Item Pemeliharaan Sarpras', 'Estimasi Anggaran', 'Sumber Dana', 'Target Realisasi & PJ']
    header_row = rkas_table.rows[0]
    for idx, text in enumerate(rkas_headers):
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

    rkas_rows = [
        ('1', 'Pengecatan Marka Jalur Evakuasi & Titik Kumpul', 'Rp 450.000,-', 'BOS Reguler (Pemeliharaan)', 'November 2026 / PJ Sarpras'),
        ('2', 'Pengadaan Kotak P3K Lengkap & Safety Plug Cover Listrik', 'Rp 650.000,-', 'BOS Reguler (Operasional UKS)', 'November 2026 / Koord UKS'),
        ('3', 'Inspeksi & Pemeliharaan Tahunan Tabung APAR', 'Rp 500.000,-', 'BOS Reguler (K3 Sekolah)', 'Desember 2026 / Petugas K3'),
        ('4', 'Pemasangan Kanopi Peneduh Zona Penjemputan Gerbang Depan', 'Rp 2.500.000,-', 'BOS Kinerja / Komite Sekolah', 'Januari 2027 / Tim Pembangunan')
    ]

    col_rk = [Inches(0.4), Inches(2.6), Inches(1.3), Inches(1.4), Inches(1.1)]
    for r_idx, data_row in enumerate(rkas_rows):
        row = rkas_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_rk[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 70, 70, 60, 60)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 2]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 5: DOKUMENTASI FOTO SARPRAS AMAN ====================
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(8)
    h5.paragraph_format.space_after = Pt(4)
    r_h5 = h5.add_run('BAGIAN V: DOKUMENTASI FASILITAS KEAMANAN SARANA & PRASARANA')
    r_h5.bold = True
    r_h5.font.size = Pt(10.5)
    r_h5.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: PAGAR UTAMA & PINTU GERBANG TERKUNCI AMAN ]\n\nKeterangan: Pagar keliling sekolah kokoh dan gerbang selalu terkunci selama jam KBM berlangsung.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: SAKLAR LISTRIK TERTUTUP & PANEL AMAN ]\n\nKeterangan: Instalasi kelistrikan terlindung rapi dan saklar terpasang pada ketinggian aman dari jangkauan anak.', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: KESIAPAN APAR & KOTAK P3K TERSTANDAR ]\n\nKeterangan: Tabung pemadam kebakaran siap pakai di dekat laboratorium dan kotak P3K lengkap di lorong kelas.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: JALUR EVAKUASI & TITIK KUMPUL DI LAPANGAN ]\n\nKeterangan: Petunjuk jalur evakuasi darurat yang bebas hambatan dan area titik kumpul terbuka di lapangan.', photo_table.rows[1].cells[1])
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
    p0.add_run('Koordinator Sarana & Prasarana,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Ketua Komite Sekolah,\n\n\n\n\n( ......................................... )').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_23_Keamanan_Fisik_Sekolah_Sarpras.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_23:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_23()
