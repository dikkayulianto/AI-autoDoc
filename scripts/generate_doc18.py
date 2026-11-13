import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

folder = r'Bukti_Fisik_Jabar_Beraksi'
os.makedirs(folder, exist_ok=True)

def create_doc_18():
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
    r_t = p_title.add_run('BUKTI FISIK INSTRUMEN JABAR BERAKSI (POIN 18)')
    r_t.bold = True
    r_t.font.size = Pt(12)
    r_t.font.name = 'Arial'

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    r_s = p_sub.add_run('INDIKATOR: PEMANTAUAN FISIK DAN EMOSI PESERTA DIDIK\n(Lampiran: Catatan Guru / Jurnal Anekdot, Buku Pemantauan Tumbuh Kembang & Emosi, serta Dokumentasi)')
    r_s.font.size = Pt(9.5)
    r_s.bold = True
    r_s.font.name = 'Arial'
    r_s.font.color.rgb = RGBColor(13, 71, 161)

    # ==================== BAGIAN 1: CATATAN ANEKDOT GURU ====================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run('A. LEMBAR CATATAN GURU (JURNAL HARIAN SIKAP & DINAMIKA EMOSI SISWA)')
    r_h1.bold = True
    r_h1.font.size = Pt(10.5)
    r_h1.font.color.rgb = RGBColor(13, 71, 161)

    p_desc1 = doc.add_paragraph()
    p_desc1.paragraph_format.space_after = Pt(4)
    r_d1 = p_desc1.add_run('Pencatatan perilaku emosional spesifik siswa saat kegiatan pembelajaran dan istirahat:')
    r_d1.font.size = Pt(9)
    r_d1.italic = True

    anec_table = doc.add_table(rows=5, cols=6)
    anec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(anec_table)

    headers_anec = ['No', 'Hari/Tgl', 'Nama Siswa / Kelas', 'Peristiwa / Perilaku yang Teramati', 'Analisis Emosi & Respon', 'Tindak Lanjut Guru']
    header_row = anec_table.rows[0]
    for idx, text in enumerate(headers_anec):
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

    anec_data = [
        ('1', 'Senin, 07/09/26', 'Alif Pratama\n(Kls IV-A)', 'Tampak antusias menyambut teman yang baru pulih dari sakit, meminjamkan buku catatan tanpa diminta.', 'Emosi positif, empati tinggi dan kepedulian sosial kuat.', 'Memberikan apresiasi verbal di depan kelas sebagai contoh pembiasaan baik.'),
        ('2', 'Rabu, 16/09/26', 'Dina Safitri\n(Kls IV-A)', 'Menunduk saat jam istirahat, menolak makan bekal dan enggan berbaur dengan kelompok bermain.', 'Mengalami kecemasan / stres ringan karena tugas rumah belum tuntas.', 'Pendekatan personal, mendengarkan cerita anak, dan memberi waktu jeda relaksasi.'),
        ('3', 'Jumat, 25/09/26', 'Bagus Ramadhan\n(Kls V-B)', 'Cepat marah dan melempar pensil saat gambar buatannya tidak sengaja tersenggol kawan sebangku.', 'Kesulitan regulasi emosi spontan (frustrasi sesaat).', 'Mengajak anak latihan teknik pernapasan menenangkan, difasilitasi saling bermaafan.'),
        ('4', 'Selasa, 13/10/26', 'Clara Amanda\n(Kls VI-A)', 'Menangis tersedu setelah presentasi kelompok karena merasa penampilannya kurang maksimal.', 'Perfeksionisme berlebih dan kecemasan evaluasi.', 'Konseling penguatan konsep diri: belajar adalah proses bertumbuh tanpa harus sempurna.')
    ]

    col_anec = [Inches(0.4), Inches(1.1), Inches(1.2), Inches(1.8), Inches(1.2), Inches(1.1)]
    for r_idx, data_row in enumerate(anec_data):
        row = anec_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_anec[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 80, 80, 70, 70)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==================== BAGIAN 2: BUKU PEMANTAUAN TUMBUH KEMBANG & FISIK ====================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run('B. BUKU PEMANTAUAN KESEHATAN FISIK DAN KEBERSIHAN SISWA')
    r_h2.bold = True
    r_h2.font.size = Pt(10.5)
    r_h2.font.color.rgb = RGBColor(13, 71, 161)

    p_desc2 = doc.add_paragraph()
    p_desc2.paragraph_format.space_after = Pt(4)
    r_d2 = p_desc2.add_run('Instrumen Pemantauan Fisik Berkala (Kerjasama Guru Kelas & Tim UKS Sekolah):')
    r_d2.font.size = Pt(9)
    r_d2.italic = True

    phys_table = doc.add_table(rows=5, cols=7)
    phys_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(phys_table)

    phys_headers = ['No', 'Nama Siswa', 'TB / BB (cm/kg)', 'Kebersihan Diri (Kuku/Gigi/Rambut)', 'Penglihatan / Pendengaran', 'Status Gizi & Kebugaran', 'Catatan Guru UKS']
    header_row = phys_table.rows[0]
    for idx, text in enumerate(phys_headers):
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

    phys_data = [
        ('1', 'Ahmad Farhan\n(NIS: 23101)', '138 cm / 32 kg', 'Sangat Bersih (Kuku rapi, gigi terawat)', 'Normal (Duduk baris tengah jelas)', 'Normal / Bugar aktif', 'Pertahankan pola makan bergizi.'),
        ('2', 'Citra Kirana\n(NIS: 23102)', '132 cm / 26 kg', 'Bersih dan rapi', 'Sering memicingkan mata saat baca papan tulis', 'Gizi Cukup, postur ideal', 'Disarankan periksa mata ke Puskesmas, pindah duduk baris depan.'),
        ('3', 'Eko Saputra\n(NIS: 23103)', '141 cm / 45 kg', 'Cukup Bersih, pakaian sedikit kusut', 'Normal', 'Kecenderungan Gizi Lebih (Overweight)', 'Edukasi kurangi minuman manis kemasan, dorong aktif di senam pagi.'),
        ('4', 'Fitri Handayani\n(NIS: 23104)', '135 cm / 29 kg', 'Sangat Rapi & Bersih', 'Normal', 'Sangat Baik / Bugar', 'Kondisi fisik optimal, selalu sarapan pagi.')
    ]

    col_phys = [Inches(0.4), Inches(1.2), Inches(1.1), Inches(1.3), Inches(1.1), Inches(1.0), Inches(0.7)]
    for r_idx, data_row in enumerate(phys_data):
        row = phys_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8F9FA'
        for c_idx, val in enumerate(data_row):
            cell = row.cells[c_idx]
            cell.width = col_phys[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 80, 80, 70, 70)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 2]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # ==================== BAGIAN 3: DOKUMENTASI PEMANTAUAN ====================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    r_h3 = h3.add_run('C. LEMBAR DOKUMENTASI PEMANTAUAN FISIK DAN EMOSI SISWA')
    r_h3.bold = True
    r_h3.font.size = Pt(10.5)
    r_h3.font.color.rgb = RGBColor(13, 71, 161)

    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_after = Pt(4)
    r_bx = p_box.add_run('Dokumentasi Kegiatan Pemantauan Rutin (Tempelkan Foto Sekolah pada Kotak di Bawah):')
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
    p_p1.add_run('[ TEMPELKAN FOTO PEMERIKSAAN KESEHATAN FISIK SISWA DI SINI ]\n\nKeterangan: Pengukuran tinggi badan, berat badan, serta cek kuku & gigi oleh Guru UKS.').font.size = Pt(8.5)

    p_p2 = cell_p2.paragraphs[0]
    p_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_p2.add_run('[ TEMPELKAN FOTO ICE BREAKING / REFLEKSI EMOSI DI SINI ]\n\nKeterangan: Guru memandu "Emoticon Check-in" sebelum pembelajaran untuk memantau mood siswa.').font.size = Pt(8.5)

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
    p0.add_run('Guru Pembina UKS,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run('Wali Kelas / Guru Kelas,\n\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run('Mengetahui,\nKepala SD Negeri [........],\n\n\n\n( ......................................... )\nNIP. .....................................').font.size = Pt(9)

    out_file = os.path.join(folder, 'Bukti_Fisik_Poin_18_Pemantauan_Fisik_Emosi_Peserta_Didik.docx')
    doc.save(out_file)
    print("SUCCESS_FILE_18:", os.path.abspath(out_file))

if __name__ == '__main__':
    create_doc_18()
