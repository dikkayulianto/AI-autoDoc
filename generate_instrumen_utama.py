import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_element(name):
    return OxmlElement(name)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="B0BEC5", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def build_document():
    doc = docx.Document()
    
    # Page setup (A4, standard government margins)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 1. KOP SURAT
    p_kop1 = doc.add_paragraph()
    p_kop1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_kop1 = p_kop1.add_run("PEMERINTAH KOTA CIREBON\nDINAS PENDIDIKAN")
    r_kop1.bold = True
    r_kop1.font.name = "Arial"
    r_kop1.font.size = Pt(13)

    p_kop2 = doc.add_paragraph()
    p_kop2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_kop2 = p_kop2.add_run("Jalan Brigjen Dharsono Nomor 7 Bypass, Cirebon 45132  Telepon (0231) 486579\nLaman: disdik.cirebonkota.go.id  Posel: disdik@cirebonkota.go.id")
    r_kop2.font.name = "Arial"
    r_kop2.font.size = Pt(9)
    r_kop2.font.color.rgb = RGBColor(90, 90, 90)

    # Kop divider
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="18" w:space="1" w:color="000000"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(p_div_border)

    # 2. TITLE
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t1 = p_title.add_run("INSTRUMEN PEMANTAUAN JABAR BERAKSI : 1\n")
    r_t1.bold = True
    r_t1.font.name = "Arial"
    r_t1.font.size = Pt(12)
    r_t2 = p_title.add_run("IMPLEMENTASI PENGUATAN PENDIDIKAN KARAKTER MELALUI PEMBIASAAN DI SATUAN PENDIDIKAN")
    r_t2.bold = True
    r_t2.font.name = "Arial"
    r_t2.font.size = Pt(11)
    p_title.paragraph_format.space_after = Pt(10)

    # 3. IDENTITAS TABEL
    p_sec1 = doc.add_paragraph()
    r_s1 = p_sec1.add_run("IDENTITAS PEMANTAUAN")
    r_s1.bold = True
    r_s1.font.name = "Arial"
    r_s1.font.size = Pt(10)
    p_sec1.paragraph_format.space_after = Pt(4)

    id_data = [
        ("Nama Pengawas", "Ari Haerani, S.Pd.SD, M.Pd"),
        ("NIP Pengawas", "197301211996032001"),
        ("Wilayah Binaan", "Kecamatan Harjamukti dan Kesambi"),
        ("Tanggal Pemantauan", "13 November 2026"),
        ("Waktu Pemantauan", "08.00 - 11.30 WIB"),
        ("Nama Satuan Pendidikan", "SD NEGERI 1 KESAMBI"),
        ("Jenis Sekolah", "[✓] Negeri    [ ] Swasta"),
        ("NPSN / Alamat Sekolah", "20222045 / Jalan Kesambi No. 24, Kota Cirebon 45132"),
        ("Kepala Sekolah", "Dra. Hj. Siti Maryam, M.Pd"),
        ("NIP / NUPTK", "197505121998032004"),
        ("Jumlah Peserta Terdaftar", "284 orang"),
        ("Jumlah Peserta Hadir", "284 orang (100%)")
    ]

    t_id = doc.add_table(rows=len(id_data), cols=2)
    t_id.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_id, color="CFD8DC", sz="4")

    for i, (label, val) in enumerate(id_data):
        c0 = t_id.cell(i, 0)
        c1 = t_id.cell(i, 1)
        c0.width = Inches(2.3)
        c1.width = Inches(4.7)
        set_cell_background(c0, "F8F9FA")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)

        p0 = c0.paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.name = "Arial"
        r0.font.size = Pt(9)

        p1 = c1.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.name = "Arial"
        r1.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4. PETUNJUK
    p_pet = doc.add_paragraph()
    r_pet = p_pet.add_run("INSTRUMEN PEMANTAUAN\n")
    r_pet.bold = True
    r_pet.font.name = "Arial"
    r_pet.font.size = Pt(10)
    
    r_desc = p_pet.add_run("Petunjuk Pengisian:\n1. Beri tanda (✓) pada kolom ketercapaian sesuai kondisi riil (B = Belum, S = Sudah Seluruhnya).\n2. Beri tanda (✓) pada jenis bukti fisik yang tersedia.\n3. Tambahkan keterangan/persentase keterlaksanaan.")
    r_desc.font.name = "Arial"
    r_desc.font.size = Pt(8.5)
    r_desc.font.italic = True
    p_pet.paragraph_format.space_after = Pt(6)

    # 5. TABEL 27 INDIKATOR
    indicators = [
        (1, "Pembiasaan bangun pagi", ["Jurnal pembiasaan", "Dokumentasi foto/video", "Edaran kepada orang tua"], "98%"),
        (2, "Pembiasaan beribadah", ["Jadwal kegiatan", "Dokumentasi", "Jurnal kegiatan"], "100%"),
        (3, "Pembiasaan berolahraga", ["Jadwal olahraga", "Dokumentasi", "Daftar hadir"], "96%"),
        (4, "Pembiasaan makan bergizi", ["Program sekolah sehat", "Dokumentasi", "Edaran"], "95%"),
        (5, "Pembiasaan gemar belajar", ["Program literasi", "Jurnal literasi", "Dokumentasi"], "94%"),
        (6, "Pembiasaan bermasyarakat", ["Laporan kegiatan sosial", "Dokumentasi", "Surat tugas"], "92%"),
        (7, "Pembiasaan tidur cepat", ["Sosialisasi kepada orang tua", "Dokumentasi"], "90%"),
        (8, "Senam Pagi Anak Indonesia Hebat", ["Jadwal senam", "Dokumentasi", "Daftar hadir"], "100%"),
        (9, "Menyanyikan Lagu Indonesia Raya", ["Jadwal Pagi Ceria", "Dokumentasi"], "100%"),
        (10, "Berdoa sesuai keyakinan", ["Jadwal kegiatan", "Dokumentasi"], "100%"),
        (11, "Kegiatan kepanduan (pra siaga) dan ekskul", ["SK pembina", "Jadwal kegiatan", "Daftar hadir", "Dokumentasi"], "96%"),
        (12, "Tata tertib, kode etik dan SOP", ["SK Penetapan", "Tata tertib", "Kode etik", "SOP"], "100%"),
        (13, "Perencanaan kegiatan dan anggaran BSAN", ["RKAS", "RKT", "Program kerja"], "100%"),
        (14, "Supervisi dan pembinaan guru/tendik", ["Program supervisi", "Instrumen", "Laporan supervisi"], "95%"),
        (15, "Edukasi kepada warga sekolah", ["Materi sosialisasi", "Notulen", "Dokumentasi"], "95%"),
        (16, "Deteksi dini dan penanganan masalah", ["SOP Perlindungan", "Buku kasus", "Laporan penanganan"], "100%"),
        (17, "Kemitraan orang tua / pemangku kepentingan", ["Notulen rapat", "MoU Puskesmas", "Dokumentasi"], "95%"),
        (18, "Pemantauan fisik-emosi peserta didik", ["Catatan guru", "Buku pemantauan", "Dokumentasi UKS"], "94%"),
        (19, "Penguatan spiritual, kesehatan mental, moral, etika", ["Modul ajar", "Jurnal pembelajaran", "Dokumentasi 5S"], "96%"),
        (20, "Keamanan fisik dan sportivitas PJOK", ["Modul ajar", "Jurnal pembelajaran", "Dokumentasi fair play"], "96%"),
        (21, "Dukungan tendik lingkungan aman, nyaman fisik-psikis", ["Program kerja", "Dokumentasi", "Laporan K3"], "95%"),
        (22, "Koordinasi waka kesiswaan/guru kelas", ["Notulen rapat", "Program kerja", "Dokumentasi"], "98%"),
        (23, "Keamanan fisik sekolah berdasarkan sarpras", ["Checklist sarpras", "Laporan", "Dokumentasi K3"], "92%"),
        (24, "Layanan psikososial Guru BK", ["Program BK", "Jurnal layanan", "Laporan konseling"], "95%"),
        (25, "Pendampingan wali kelas", ["Buku wali kelas", "Catatan perkembangan", "Dokumentasi"], "98%"),
        (26, "Pembelajaran bermakna & integrasi karakter Holistic Integratif", ["Modul ajar", "Instrumen observasi", "Dokumentasi P5"], "96%"),
        (27, "Keamanan lingkungan dan data pribadi murid", ["SOP keamanan", "Dokumentasi", "Program kerja"], "100%")
    ]

    t_ind = doc.add_table(rows=len(indicators) + 1, cols=6)
    t_ind.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_ind, color="1E3A8A", sz="5")

    # Header row
    headers = [
        ("No", Inches(0.4)),
        ("Indikator Karakter", Inches(2.2)),
        ("B", Inches(0.35)),
        ("S", Inches(0.35)),
        ("Bukti Fisik yang Tersedia", Inches(2.9)),
        ("Persentase (10-100%)", Inches(0.8))
    ]

    for c_idx, (h_text, w) in enumerate(headers):
        cell = t_ind.cell(0, c_idx)
        cell.width = w
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_text)
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    # Data rows
    for idx, (num, title, ev_list, pct) in enumerate(indicators):
        row_idx = idx + 1
        bg = "FFFFFF" if row_idx % 2 != 0 else "F8FAFC"
        
        # Col 0: No
        c0 = t_ind.cell(row_idx, 0)
        c0.width = Inches(0.4)
        set_cell_background(c0, bg)
        set_cell_margins(c0, 60, 60, 60, 60)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(str(num))
        r0.font.name = "Arial"
        r0.font.size = Pt(8.5)

        # Col 1: Indikator
        c1 = t_ind.cell(row_idx, 1)
        c1.width = Inches(2.2)
        set_cell_background(c1, bg)
        set_cell_margins(c1, 60, 60, 80, 80)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(title)
        r1.bold = True
        r1.font.name = "Arial"
        r1.font.size = Pt(8.5)

        # Col 2: B (Belum)
        c2 = t_ind.cell(row_idx, 2)
        c2.width = Inches(0.35)
        set_cell_background(c2, bg)
        set_cell_margins(c2, 60, 60, 40, 40)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run("")

        # Col 3: S (Sudah)
        c3 = t_ind.cell(row_idx, 3)
        c3.width = Inches(0.35)
        set_cell_background(c3, bg)
        set_cell_margins(c3, 60, 60, 40, 40)
        p3 = c3.paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run("✓")
        r3.bold = True
        r3.font.name = "Arial"
        r3.font.size = Pt(10)
        r3.font.color.rgb = RGBColor(22, 101, 52)

        # Col 4: Bukti Fisik Checklist
        c4 = t_ind.cell(row_idx, 4)
        c4.width = Inches(2.9)
        set_cell_background(c4, bg)
        set_cell_margins(c4, 60, 60, 80, 80)
        p4 = c4.paragraphs[0]
        ev_text = "\n".join([f"☑  {ev}" for ev in ev_list])
        r4 = p4.add_run(ev_text)
        r4.font.name = "Arial"
        r4.font.size = Pt(8)

        # Col 5: Persentase
        c5 = t_ind.cell(row_idx, 5)
        c5.width = Inches(0.8)
        set_cell_background(c5, bg)
        set_cell_margins(c5, 60, 60, 60, 60)
        p5 = c5.paragraphs[0]
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5 = p5.add_run(pct)
        r5.bold = True
        r5.font.name = "Arial"
        r5.font.size = Pt(8.5)
        r5.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 6. GAMBARAN UMUM
    p_gu = doc.add_paragraph()
    r_gut = p_gu.add_run("GAMBARAN UMUM PELAKSANAAN PROGRAM\n")
    r_gut.bold = True
    r_gut.font.name = "Arial"
    r_gut.font.size = Pt(10)
    r_gu_body = p_gu.add_run(
        "Pelaksanaan program Jabar Beraksi di satuan pendidikan telah terintegrasi dengan baik ke dalam budaya sekolah sehari-hari. "
        "Pembiasaan 7 Hari Berkarakter (bangun pagi, beribadah, berolahraga, makan sehat bergizi, gemar belajar/literasi, bermasyarakat, "
        "dan istirahat tepat waktu) telah disosialisasikan dan dikerjasamakan secara aktif dengan orang tua murid. Kegiatan pagi diawali secara rutin "
        "dengan doa bersama, menyanyikan lagu Indonesia Raya, dan senam kebugaran. Seluruh perangkat regulasi sekolah seperti SOP, Tata Tertib, "
        "serta dokumen perencanaan (RKT & RKAS) telah memuat alokasi penguatan karakter, didukung oleh lingkungan sarana prasarana yang aman, "
        "ramah anak, dan kondusif."
    )
    r_gu_body.font.name = "Arial"
    r_gu_body.font.size = Pt(9)
    p_gu.paragraph_format.line_spacing = 1.15
    p_gu.paragraph_format.space_after = Pt(8)

    # 7. KENDALA
    p_ken = doc.add_paragraph()
    r_kent = p_ken.add_run("KENDALA\n")
    r_kent.bold = True
    r_kent.font.name = "Arial"
    r_kent.font.size = Pt(10)
    kendala_text = (
        "1. Pemantauan pembiasaan di luar jam sekolah (seperti jam tidur malam dan pengurangan screen-time gawai di rumah) masih bergantung pada komitmen dan kesibukan masing-masing orang tua.\n"
        "2. Fasilitas sarana prasarana ramah anak dan ruang konseling/BK masih terus dalam proses pembenahan dan peningkatan bertahap.\n"
        "3. Keterbatasan waktu koordinasi intensif bersama wali murid yang mayoritas bekerja penuh waktu."
    )
    r_ken_body = p_ken.add_run(kendala_text)
    r_ken_body.font.name = "Arial"
    r_ken_body.font.size = Pt(9)
    p_ken.paragraph_format.line_spacing = 1.15
    p_ken.paragraph_format.space_after = Pt(8)

    # 8. RENCANA TINDAK LANJUT
    p_rtl = doc.add_paragraph()
    r_rtlt = p_rtl.add_run("RENCANA TINDAK LANJUT\n")
    r_rtlt.bold = True
    r_rtlt.font.name = "Arial"
    r_rtlt.font.size = Pt(10)
    rtl_text = (
        "1. Mengoptimalkan buku penghubung/jurnal digital pembiasaan siswa yang terhubung langsung secara berkala dengan orang tua/wali murid.\n"
        "2. Menyelenggarakan kegiatan parenting dan sosialisasi berkala mengenai pentingnya pola asuh berkarakter, gizi seimbang, dan pembatasan gawai di rumah.\n"
        "3. Melakukan pemeliharaan sarana prasarana sekolah secara berkala melalui anggaran RKAS untuk memastikan keamanan lingkungan belajar peserta didik.\n"
        "4. Meningkatkan kapasitas guru kelas dan guru BK melalui pelatihan konseling ramah anak dan deteksi dini pencegahan kekerasan di sekolah."
    )
    r_rtl_body = p_rtl.add_run(rtl_text)
    r_rtl_body.font.name = "Arial"
    r_rtl_body.font.size = Pt(9)
    p_rtl.paragraph_format.line_spacing = 1.15
    p_rtl.paragraph_format.space_after = Pt(14)

    # 9. TANDA TANGAN
    t_ttd = doc.add_table(rows=1, cols=2)
    t_ttd.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Remove borders
    for row in t_ttd.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
            tcPr.append(tcBorders)

    c_ks = t_ttd.cell(0, 0)
    c_ks.width = Inches(3.5)
    p_ks = c_ks.paragraphs[0]
    p_ks.paragraph_format.line_spacing = 1.15
    r_k = p_ks.add_run("Mengetahui,\nKepala SD NEGERI 1 KESAMBI\n\n\n\n\n")
    r_k.font.name = "Arial"
    r_k.font.size = Pt(9.5)
    r_k_name = p_ks.add_run("Dra. Hj. Siti Maryam, M.Pd\n")
    r_k_name.bold = True
    r_k_name.underline = True
    r_k_name.font.name = "Arial"
    r_k_name.font.size = Pt(9.5)
    r_k_nip = p_ks.add_run("NIP. 197505121998032004")
    r_k_nip.font.name = "Arial"
    r_k_nip.font.size = Pt(9)

    c_p = t_ttd.cell(0, 1)
    c_p.width = Inches(3.5)
    p_p = c_p.paragraphs[0]
    p_p.paragraph_format.line_spacing = 1.15
    r_p = p_p.add_run("Cirebon, 13 November 2026\nPengawas / Pemantau,\n\n\n\n\n")
    r_p.font.name = "Arial"
    r_p.font.size = Pt(9.5)
    r_p_name = p_p.add_run("Ari Haerani, S.Pd.SD, M.Pd\n")
    r_p_name.bold = True
    r_p_name.underline = True
    r_p_name.font.name = "Arial"
    r_p_name.font.size = Pt(9.5)
    r_p_nip = p_p.add_run("NIP. 197301211996032001")
    r_p_nip.font.name = "Arial"
    r_p_nip.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 10. CATATAN & LINK GOOGLE DRIVE
    p_cat = doc.add_paragraph()
    r_catt = p_cat.add_run("Catatan:\n")
    r_catt.bold = True
    r_catt.font.name = "Arial"
    r_catt.font.size = Pt(8.5)
    
    cat_text = (
        "1. Mohon dilampiri bukti fisik sesuai dengan indikator dan masukkan pada tautan sekolah masing-masing.\n"
        "2. Format Penamaan File: Nama Sekolah_Jenis Instrumen_2026\n"
        "3. Link Bukti Sekolah (Google Drive) : https://drive.google.com/drive/folders/1SDN1Kesambi_JabarBeraksi2026"
    )
    r_cat_b = p_cat.add_run(cat_text)
    r_cat_b.font.name = "Arial"
    r_cat_b.font.size = Pt(8.5)
    r_cat_b.font.italic = True

    # Save
    out_path = "c:/Users/DESKTOP/Documents/Antigravity/AI-autoDoc/Bukti_Fisik_Jabar_Beraksi/Instrumen_Pemantauan_Jabar_Beraksi_27Poin_SDN1KESAMBI_2026.docx"
    doc.save(out_path)
    print("Document successfully created at:", out_path)

if __name__ == "__main__":
    build_document()
