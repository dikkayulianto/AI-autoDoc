import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_20_comprehensive():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI - POIN 20')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: KEAMANAN FISIK DAN SPORTIVITAS PJOK\n(Berkas Lengkap: SOP Keselamatan Fisik, Modul Ajar PJOK Ramah Anak, Rubrik Penilaian Sportivitas, Jurnal Pembelajaran & Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: SOP KESELAMATAN & KEAMANAN FISIK PJOK ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('BAGIAN I: STANDAR OPERASIONAL PROSEDUR (SOP) KESELAMATAN AKTIVITAS PJOK')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    sop_table = doc.add_table(rows=6, cols=2)
    sop_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sop_table)

    sop_data = [
        ('1. Identitas SOP', 'Nomor SOP: SOP/PJOK-K3/01/2026\nJudul: Keselamatan Fisik, Mitigasi Cedera, & Protokol Keamanan Lapangan Olahraga'),
        ('2. Tujuan & Sasaran', 'Menjamin seluruh proses pembelajaran PJOK berlangsung tanpa risiko cedera fisik berat, menciptakan rasa aman bagi siswa saat beraktivitas gerak, serta menanamkan budaya kepedulian terhadap keselamatan diri dan orang lain.'),
        ('3. Skrining Kesehatan Pra-KBM', 
         'a. Pendidik wajib menanyakan kondisi kesehatan fisik siswa sebelum menuju lapangan (pemeriksaan keluhan pusing, sesak napas, luka terbuka, atau demam).\n'
         'b. Siswa yang sedang dalam masa pemulihan sakit diberikan peran pendukung non-fisik (misal: pencatat skor, pengamat sportivitas, atau istirahat di ruang UKS).\n'
         'c. Memastikan siswa mengenakan seragam olahraga standar dan sepatu bertali aman.'),
        ('4. Prosedur Keamanan Lapangan & Alat', 
         'a. Sterilisasi Area: Pemeriksaan area lapangan 10 menit sebelum jam olahraga dari pecahan kaca, batu/kerikil tajam, paku, atau permukaan licin/becek.\n'
         'b. Pengecekan Alat: Bola menggunakan tekanan udara ramah anak (tidak terlalu keras), tiang gawang/ring terpasang kokoh dengan bantalan pengaman bila diperlukan.\n'
         'c. Penempatan P3K: Kotak P3K lapangan wajib berada di sisi lapangan dengan isi lengkap (antiseptik, kasa steril, perban elastis, es kompres instan, dan oksigen semprot).'),
        ('5. Alur Penanganan Insiden Cedera', 
         'Tahap 1 (Tanggap Darurat): Menghentikan sementara permainan, amankan posisi korban dari kerumunan.\n'
         'Tahap 2 (Prinsip RICE): Rest (istirahatkan), Ice (kompres es jika memar/keseleo), Compression (balut tekan ringan), Elevation (tinggikan bagian cedera).\n'
         'Tahap 3 (Rujukan UKS/Medis): Evakuasi ke ruang UKS. Jika dicurigai dislokasi/patah tulang, segera hubungi Puskesmas pembina dan orang tua siswa.\n'
         'Tahap 4 (Pelaporan): Pencatatan kronologi kejadian dalam Buku Insiden Olahraga maksimal 1x24 jam.'),
        ('6. Disiplin Pasca Olahraga', 'Wajib melakukan pendinginan teratur (5-10 menit) dan minum air putih yang cukup untuk mencegah dehidrasi.')
    ]

    col_widths = [Inches(1.8), Inches(5.0)]
    for row_idx, (col0, col1) in enumerate(sop_data):
        row = sop_table.rows[row_idx]
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

    # ==================== BAGIAN 2: MODUL AJAR PJOK ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('BAGIAN II: DESAIN MODUL AJAR PJOK INTEGRASI SPORTIVITAS')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    modul_table = doc.add_table(rows=7, cols=2)
    modul_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(modul_table)

    modul_data = [
        ('1. Informasi Umum', 'Penyusun: Tim Guru PJOK SD Negeri [Nama Sekolah]\nFase / Kelas: Fase B / Kelas IV (Empat)\nElemen: Keterampilan Gerak & Pengembangan Karakter (Sportivitas, Mandiri, Gotong Royong)\nAlokasi: 3 JP (3 x 35 Menit)'),
        ('2. Capaian Pembelajaran (CP)', 'Peserta didik dapat menunjukkan kemampuan mempraktikkan variasi pola gerak dasar manipulatif melalui permainan invasi/modifikasi beregu, menerapkan regulasi keselamatan diri, serta menjunjung tinggi nilai sportivitas, kejujuran, dan respek kepada teman maupun lawan.'),
        ('3. Tujuan Pembelajaran (TP)', '1. Peserta didik mampu melakukan gerak melempar dan menangkap bola dengan koordinasi baik tanpa membahayakan kawan.\n2. Peserta didik membiasakan diri bersikap jujur mengakui kesalahan permainan tanpa berdebat kusir.\n3. Peserta didik menunjukkan sikap respek dengan berjabat tangan/tos sebelum dan sesudah bertanding.'),
        ('4. Rincian Kegiatan Awal (20 Menit)', 
         '• Berbaris di tepi lapangan, berdoa dipimpin salah seorang siswa secara bergilir.\n'
         '• Presensi dan skrining fisik (cek suhu/keluhan sakit).\n'
         '• Pemanasan terstruktur: Peregangan statis (kepala hingga kaki) dan dinamis melalui permainan pemanasan "Lampu Merah-Hijau Berlari" untuk meningkatkan kelenturan sendi.'),
        ('5. Rincian Kegiatan Inti (65 Menit)', 
         '• Aktivitas 1: Latihan teknik dasar operan berpasangan dengan jarak bertahap (3 - 5 meter).\n'
         '• Aktivitas 2: Simulasi Game "Bola Tangan Modifikasi Empat Gawang" (peraturan: dilarang kontak fisik keras, bola hanya boleh dioper bukan direbut secara paksa dari tangan).\n'
         '• Penanaman Sikap: Bila bola menyentuh garis keluar, pemain yang menyentuh terakhir wajib sukarela menyerahkan bola ke lawan (latihan kejujuran spontan).\n'
         '• Bila terjadi tabrakan tidak sengaja, permainan dijeda otomatis untuk memastikan kedua pemain saling membantu bangkit dan memaafkan.'),
        ('6. Rincian Kegiatan Penutup (20 Menit)', 
         '• Pendinginan: Senam peregangan relaksasi napas dan ayunan tangan lembut berpasangan.\n'
         '• Lingkaran Refleksi (Debriefing): Guru memfasilitasi diskusi: "Bagaimana rasanya bermain sportif hari ini?", "Apa yang kita lakukan saat tim kita kebobolan?".\n'
         '• Seluruh siswa dan guru membentuk barisan saling berjabat tangan dan memberikan tepuk tangan apresiasi kepada seluruh kelompok.'),
        ('7. Penilaian Pembelajaran', 'Teknik Penilaian: Observasi Kinerja Afektif (Sportivitas & Tanggung Jawab) menggunakan Lembar Rubrik Observasi PJOK.')
    ]

    for row_idx, (col0, col1) in enumerate(modul_data):
        row = modul_table.rows[row_idx]
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

    # ==================== BAGIAN 3: RUBRIK & REKAP NILAI SIKAP SPORTIVITAS ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('BAGIAN III: RUBRIK ASESMEN DAN LEMBAR OBSERVASI SPORTIVITAS')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    p_rub = doc.add_paragraph()
    p_rub.paragraph_format.space_after = Pt(3)
    r_rb = p_rub.add_run('Kriteria Penilaian Sikap (Skala 1 - 4): 4 = Sangat Baik (Selalu konsisten), 3 = Baik (Sering konsisten), 2 = Cukup (Mulai tampak), 1 = Perlu Bimbingan.')
    r_rb.font.size = Pt(8)
    r_rb.italic = True

    score_table = doc.add_table(rows=7, cols=8)
    score_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(score_table)

    headers_score = ['No', 'Nama Siswa', 'Kelas', 'Kejujuran\n(Aturan)', 'Respek Lawan\n(Tos/Salam)', 'Menerima Skor\n(Lapang Dada)', 'Empati Kawan\n(Bantu Teman)', 'Predikat / Skor Akhir']
    header_row = score_table.rows[0]
    for idx, text in enumerate(headers_score):
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

    score_data = [
        ('1', 'Ahmad Fadillah', 'IV-A', '4', '4', '4', '4', '16 (Sangat Baik / SB)'),
        ('2', 'Budi Santoso', 'IV-A', '3', '4', '3', '4', '14 (Baik / B)'),
        ('3', 'Citra Dewi Lestari', 'IV-A', '4', '4', '4', '4', '16 (Sangat Baik / SB)'),
        ('4', 'Deni Setiawan', 'IV-A', '3', '3', '4', '3', '13 (Baik / B)'),
        ('5', 'Eka Ramadhani', 'IV-A', '4', '4', '3', '4', '15 (Sangat Baik / SB)'),
        ('6', 'Farhan Maulana', 'IV-A', '4', '4', '4', '4', '16 (Sangat Baik / SB)')
    ]

    col_sc = [Inches(0.4), Inches(1.5), Inches(0.7), Inches(0.9), Inches(1.0), Inches(1.0), Inches(0.9), Inches(1.4)]
    for r_idx, data_row in enumerate(score_data):
        row = score_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_sc[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 70, 70, 50, 50)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx not in [1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 4: JURNAL MENGAJAR & LOG KESELAMATAN ====================
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    r_h4 = h4.add_run('BAGIAN IV: JURNAL PEMBELAJARAN GURU & LOG KEAMANAN FISIK')
    r_h4.bold = True
    r_h4.font.size = Pt(10.5)
    r_h4.font.color.rgb = RGBColor(13, 71, 161)

    jurn_table = doc.add_table(rows=5, cols=5)
    jurn_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(jurn_table)

    jurn_headers = ['Hari / Tanggal', 'Kelas & Materi', 'Status Keamanan & Fasilitas', 'Catatan Sportivitas & Karakter', 'Tindakan Guru PJOK']
    header_row = jurn_table.rows[0]
    for idx, text in enumerate(jurn_headers):
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

    jurn_rows = [
        ('Selasa, 15/09/2026', 'Kelas IV-A\nBola Tangan', 'Lapangan bersih, cuaca cerah (30°C), kotak P3K siap di tepi lapangan.', 'Siswa sangat patuh pada peluit wasit. Tidak terjadi pertikaian antartim saat skor imbang.', 'Memberikan waktu rehidrasi setiap 15 menit agar siswa tidak kelelahan.'),
        ('Kamis, 24/09/2026', 'Kelas V-B\nKebugaran & Lari', 'Area lintasan bebas kerikil. Garis batas jelas menggunakan marka kerucut.', 'Siswa saling menyemangati pelari terakhir yang mulai melambat.', 'Apresiasi kepada regu yang sportif mendampingi temannya hingga garis finish.'),
        ('Selasa, 06/10/2026', 'Kelas IV-B\nPermainan Kasti', 'Alat pemukul terlapis busa lembut. Bola kasti karet bertekanan rendah.', 'Ada 1 siswa tersenggol bola tanpa sengaja, pelempar langsung meminta maaf spontan.', 'Memuji sikap tanggung jawab pelempar di depan seluruh kelas.'),
        ('Kamis, 22/10/2026', 'Kelas VI-A\nSenam SKJ 2022', 'Lantai semen kering, jarak rentang tangan antarbaris siswa tertib 1,5 meter.', 'Kekompakan kelompok tinggi, saling mengoreksi posisi tubuh dengan ramah.', 'Refleksi manfaat kelenturan tubuh untuk kesehatan batin dan fisik.')
    ]

    col_j = [Inches(1.2), Inches(1.2), Inches(1.6), Inches(1.8), Inches(1.4)]
    for r_idx, data_row in enumerate(jurn_rows):
        row = jurn_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_j[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 70, 70, 60, 60)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ==================== BAGIAN 5: DOKUMENTASI FOTO ====================
    h5 = doc.add_paragraph()
    h5.paragraph_format.space_before = Pt(8)
    h5.paragraph_format.space_after = Pt(4)
    r_h5 = h5.add_run('BAGIAN V: DOKUMENTASI KEGIATAN KEAMANAN & SPORTIVITAS PJOK')
    r_h5.bold = True
    r_h5.font.size = Pt(10.5)
    r_h5.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=2, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table, color='90A4AE', sz='6')

    frames = [
        ('[ FOTO 1: PEMANASAN TERSTRUKTUR & SKRINING AWAL ]\n\nKeterangan: Guru memeriksa kesiapan fisik peserta didik dan memimpin peregangan statis-dinamis sebelum KBM.', photo_table.rows[0].cells[0]),
        ('[ FOTO 2: PEMERIKSAAN KELAYAKAN ALAT & AREA LAPANGAN ]\n\nKeterangan: Pengecekan kebersihan permukaan lapangan serta kesiapan kotak P3K di pinggir lapangan.', photo_table.rows[0].cells[1]),
        ('[ FOTO 3: PENERAPAN SIKAP SPORTIF DALAM PERMAINAN ]\n\nKeterangan: Peserta didik menunjukkan kepatuhan aturan, bermain adil tanpa kekerasan fisik kasar.', photo_table.rows[1].cells[0]),
        ('[ FOTO 4: PENDINGINAN & SALAM PERSAHABATAN PASCA TANDING ]\n\nKeterangan: Pembiasaan saling berjabat tangan dan refleksi bersama setelah jam pelajaran olahraga usai.', photo_table.rows[1].cells[1])
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
    p0.add_run('Guru Pengampu PJOK,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Koordinator Kurikulum,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_20_Keamanan_Fisik_dan_Sportivitas_PJOK.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_20_COMPREHENSIVE:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_20_comprehensive()
