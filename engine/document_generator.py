import io
import re
import requests
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- XML HELPER FUNCTIONS ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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

def add_double_divider(paragraph):
    p_div_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="14" w:space="1" w:color="000000"/></w:pBdr>')
    paragraph._p.get_or_add_pPr().append(p_div_border)

# --- 27 INDICATORS METADATA ---
INDICATORS_27 = [
    (1, "Pembiasaan bangun pagi", ["Jurnal pembiasaan pagi", "Dokumentasi foto/video kegiatan subuh", "Edaran kepada orang tua"], "98%"),
    (2, "Pembiasaan beribadah", ["Jadwal kegiatan ibadah rutin", "Buku jurnal shalat berjamaah", "Dokumentasi pembiasaan ibadah"], "100%"),
    (3, "Pembiasaan berolahraga", ["Jadwal olahraga & kebugaran", "Daftar hadir senam rutin", "Dokumentasi pembiasaan jasmani"], "96%"),
    (4, "Pembiasaan makan bergizi", ["Program sekolah sehat (Isi Piringku)", "Surat edaran bekal sehat", "Dokumentasi sarapan bersama"], "95%"),
    (5, "Pembiasaan gemar belajar", ["Program literasi 15 menit pagi", "Jurnal baca & resensi siswa", "Dokumentasi pojok baca kelas"], "94%"),
    (6, "Pembiasaan bermasyarakat", ["Laporan kegiatan gotong royong", "Surat tugas bakti sosial", "Dokumentasi kepedulian lingkungan"], "92%"),
    (7, "Pembiasaan tidur cepat", ["Materi sosialisasi screen-time", "Surat edaran kepada orang tua", "Dokumentasi monitoring pola tidur"], "90%"),
    (8, "Senam Pagi Anak Indonesia Hebat", ["Jadwal senam pagi terprogram", "Daftar hadir guru & siswa", "Dokumentasi senam bersama"], "100%"),
    (9, "Menyanyikan Lagu Indonesia Raya", ["Jadwal rutin pagi ceria & upacara", "Teks panduan dirigen lagu wajib", "Dokumentasi pelaksanaan di kelas"], "100%"),
    (10, "Berdoa sesuai keyakinan", ["Jadwal doa bersama sebelum/sesudah belajar", "Panduan doa religius & toleransi", "Dokumentasi pembiasaan"], "100%"),
    (11, "Kegiatan kepanduan (pra siaga) dan ekskul", ["SK pembina Pramuka / Pra-Siaga", "Jadwal latihan mingguan", "Daftar hadir & dokumentasi"], "96%"),
    (12, "Tata tertib, kode etik dan SOP", ["SK penetapan kepala sekolah", "Buku tata tertib & kode etik guru/murid", "SOP budaya positif sekolah"], "100%"),
    (13, "Perencanaan kegiatan dan anggaran BSAN", ["Rencana Kegiatan & Anggaran Sekolah (RKAS)", "Rencana Kerja Tahunan (RKT)", "Program kerja alokasi karakter"], "100%"),
    (14, "Supervisi dan pembinaan guru/tendik", ["Program supervisi penguatan karakter", "Instrumen penilaian kinerja guru", "Laporan supervisi & tindak lanjut"], "95%"),
    (15, "Edukasi kepada warga sekolah", ["Materi sosialisasi anti-bullying / PPKSP", "Notulen rapat sosialisasi warga sekolah", "Dokumentasi pertemuan parenting"], "95%"),
    (16, "Deteksi dini dan penanganan masalah", ["SOP perlindungan siswa & TPPK", "Buku log penanganan kasus", "Laporan mediasi restoratif"], "100%"),
    (17, "Kemitraan orang tua / pemangku kepentingan", ["Notulen pertemuan paguyuban kelas", "Naskah MoU kemitraan Puskesmas", "Dokumentasi sinergi tri pusat"], "95%"),
    (18, "Pemantauan fisik-emosi peserta didik", ["Jurnal anekdot emosi siswa", "Buku catatan pemeriksaan fisik UKS", "Instrumen mood meter harian"], "94%"),
    (19, "Penguatan spiritual, kesehatan mental, moral, etika", ["Modul ajar integrasi karakter", "Jurnal pembelajaran budi pekerti", "Dokumentasi pembiasaan 5S"], "96%"),
    (20, "Keamanan fisik dan sportivitas PJOK", ["SOP keselamatan aktivitas jasmani", "Modul ajar PJOK edukasi sportivitas", "Rubrik observasi fair play & P3K"], "96%"),
    (21, "Dukungan tendik lingkungan aman, nyaman fisik-psikis", ["Program kerja tenaga kependidikan", "Buku log keamanan & kebersihan", "Laporan sanitasi & kenyamanan"], "95%"),
    (22, "Koordinasi waka kesiswaan/guru kelas", ["Program kerja bidang kesiswaan", "Notulen rapat koordinasi bulanan", "Dokumentasi evaluasi siswa"], "98%"),
    (23, "Keamanan fisik sekolah berdasarkan sarpras", ["Pedoman standar K3 lingkungan", "Instrumen checklist sarpras ramah anak", "Laporan pemeliharaan gedung"], "92%"),
    (24, "Layanan psikososial Guru BK", ["Program bimbingan & konseling", "Jurnal layanan konseling individu", "Laporan penanganan psikososial"], "95%"),
    (25, "Pendampingan wali kelas", ["Buku catatan administrasi wali kelas", "Buku riwayat perkembangan murid", "Laporan kunjungan rumah (home visit)"], "98%"),
    (26, "Pembelajaran bermakna & integrasi karakter Holistic Integratif", ["Modul ajar deep learning berdiferensiasi", "Instrumen observasi keterlibatan siswa", "Dokumentasi gelar karya P5"], "96%"),
    (27, "Keamanan lingkungan dan data pribadi murid", ["SOP keamanan lingkungan & tamu", "SOP perlindungan data pribadi murid", "Dokumentasi penyimpanan arsip rapor"], "100%")
]

DEFAULT_SCHOOL = {
    "name": "SD NEGERI 1 KESAMBI",
    "npsn": "20222165",
    "kec": "Kesambi",
    "city": "Kota Cirebon",
    "address": "Jalan Kesambi No. 24, Kota Cirebon",
    "head": "Drs. H. Mulyono, M.Pd.",
    "nip": "196805141992031004",
    "supervisor": "Ari Haerani, S.Pd.SD, M.Pd",
    "supervisor_nip": "197301211996032001",
    "supervisor_region": "Wilayah Kerja Harjamukti & Kesambi",
    "year": "2026"
}

def generate_main_instrument_docx(sch=None):
    sch = {**DEFAULT_SCHOOL, **(sch or {})}
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 1. KOP SURAT
    p_kop1 = doc.add_paragraph()
    p_kop1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_kop1.add_run("PEMERINTAH KOTA CIREBON\nDINAS PENDIDIKAN")
    r1.bold = True
    r1.font.name = "Arial"
    r1.font.size = Pt(13)

    p_kop2 = doc.add_paragraph()
    p_kop2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop2.paragraph_format.space_after = Pt(8)
    r2 = p_kop2.add_run("Jalan Brigjen Dharsono Nomor 7 Bypass, Cirebon 45132  Telepon (0231) 486579\nLaman: disdik.cirebonkota.go.id  Posel: disdik@cirebonkota.go.id")
    r2.font.name = "Arial"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(90, 90, 90)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    add_double_divider(p_div)

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
    r_t2.font.size = Pt(10.5)
    r_t2.font.color.rgb = RGBColor(30, 58, 138)
    p_title.paragraph_format.space_after = Pt(10)

    # 3. IDENTITAS
    p_id_h = doc.add_paragraph()
    r_idh = p_id_h.add_run("IDENTITAS PEMANTAUAN")
    r_idh.bold = True
    r_idh.font.name = "Arial"
    r_idh.font.size = Pt(10)
    p_id_h.paragraph_format.space_after = Pt(4)

    id_data = [
        ("Nama Pengawas", sch.get('pengawas', 'Ari Haerani, S.Pd.SD, M.Pd')),
        ("NIP Pengawas", sch.get('pengawas_nip', '197301211996032001')),
        ("Wilayah Binaan", sch.get('wilayah', 'Kecamatan Harjamukti dan Kesambi')),
        ("Tanggal Pemantauan", f"13 November {sch.get('year', '2026')}"),
        ("Waktu Pemantauan", "08.00 - 11.30 WIB"),
        ("Nama Satuan Pendidikan", sch.get('name', 'SD NEGERI 1 KESAMBI')),
        ("Jenis Sekolah", "[✓] Negeri    [ ] Swasta"),
        ("NPSN / Alamat Sekolah", f"{sch.get('npsn', '20222045')} / {sch.get('address', 'Jalan Kesambi No. 24, Kota Cirebon')}"),
        ("Kepala Sekolah", sch.get('ks', 'Dra. Hj. Siti Maryam, M.Pd')),
        ("NIP / NUPTK", sch.get('ks_nip', '197505121998032004')),
        ("Jumlah Peserta Terdaftar", f"{sch.get('peserta', '284')} orang"),
        ("Jumlah Peserta Hadir", f"{sch.get('peserta', '284')} orang (100%)")
    ]

    t_id = doc.add_table(rows=len(id_data), cols=2)
    t_id.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_id, color="CFD8DC", sz="4")

    for i, (label, val) in enumerate(id_data):
        c0, c1 = t_id.cell(i, 0), t_id.cell(i, 1)
        c0.width, c1.width = Inches(2.3), Inches(4.7)
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
    r_desc = p_pet.add_run("Petunjuk Pengisian:\n1. Beri tanda (✓) pada kolom ketercapaian sesuai kondisi riil (B = Belum, S = Sudah Seluruhnya).\n2. Beri tanda (✓) pada jenis bukti fisik yang tersedia.\n3. Seluruh 27 indikator telah terimplementasi dengan predikat 'Sudah Seluruhnya (S)'.")
    r_desc.font.name = "Arial"
    r_desc.font.size = Pt(8.5)
    r_desc.font.italic = True
    p_pet.paragraph_format.space_after = Pt(6)

    # 5. TABEL 27 INDIKATOR
    t_ind = doc.add_table(rows=len(INDICATORS_27) + 1, cols=6)
    t_ind.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_ind, color="1E3A8A", sz="5")

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

    for idx, (num, title, ev_list, pct) in enumerate(INDICATORS_27):
        row_idx = idx + 1
        bg = "FFFFFF" if row_idx % 2 != 0 else "F8FAFC"
        
        # Col 0: No
        c0 = t_ind.cell(row_idx, 0)
        c0.width = Inches(0.4)
        set_cell_background(c0, bg)
        set_cell_margins(c0, 60, 60, 60, 60)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run(str(num)).font.size = Pt(8.5)

        # Col 1: Indikator
        c1 = t_ind.cell(row_idx, 1)
        c1.width = Inches(2.2)
        set_cell_background(c1, bg)
        set_cell_margins(c1, 60, 60, 80, 80)
        r1 = c1.paragraphs[0].add_run(title)
        r1.bold = True
        r1.font.size = Pt(8.5)

        # Col 2: B
        c2 = t_ind.cell(row_idx, 2)
        c2.width = Inches(0.35)
        set_cell_background(c2, bg)
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Col 3: S [✓]
        c3 = t_ind.cell(row_idx, 3)
        c3.width = Inches(0.35)
        set_cell_background(c3, bg)
        c3.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = c3.paragraphs[0].add_run("✓")
        r3.bold = True
        r3.font.size = Pt(10)
        r3.font.color.rgb = RGBColor(22, 101, 52)

        # Col 4: Bukti Fisik Checklist
        c4 = t_ind.cell(row_idx, 4)
        c4.width = Inches(2.9)
        set_cell_background(c4, bg)
        set_cell_margins(c4, 60, 60, 80, 80)
        c4.paragraphs[0].add_run("\n".join([f"☑  {ev}" for ev in ev_list])).font.size = Pt(8)

        # Col 5: Persentase
        c5 = t_ind.cell(row_idx, 5)
        c5.width = Inches(0.8)
        set_cell_background(c5, bg)
        c5.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5 = c5.paragraphs[0].add_run(pct)
        r5.bold = True
        r5.font.size = Pt(8.5)
        r5.font.color.rgb = RGBColor(30, 58, 138)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 6. GAMBARAN UMUM
    p_gu = doc.add_paragraph()
    r_gut = p_gu.add_run("GAMBARAN UMUM PELAKSANAAN PROGRAM\n")
    r_gut.bold = True
    r_gut.font.size = Pt(10)
    p_gu.add_run(
        "Pelaksanaan program Jabar Beraksi di satuan pendidikan telah terintegrasi dengan baik ke dalam budaya sekolah sehari-hari. "
        "Pembiasaan 7 Hari Berkarakter (bangun pagi, beribadah, berolahraga, makan sehat bergizi, gemar belajar/literasi, bermasyarakat, "
        "dan istirahat tepat waktu) telah disosialisasikan dan dikerjasamakan secara aktif dengan orang tua murid. Kegiatan pagi diawali secara rutin "
        "dengan doa bersama, menyanyikan lagu Indonesia Raya, dan senam kebugaran. Seluruh perangkat regulasi sekolah seperti SOP, Tata Tertib, "
        "serta dokumen perencanaan (RKT & RKAS) telah memuat alokasi penguatan karakter, didukung oleh lingkungan sarana prasarana yang aman, "
        "ramah anak, dan kondusif."
    ).font.size = Pt(9)
    p_gu.paragraph_format.space_after = Pt(8)

    # 7. KENDALA
    p_ken = doc.add_paragraph()
    r_kent = p_ken.add_run("KENDALA\n")
    r_kent.bold = True
    r_kent.font.size = Pt(10)
    p_ken.add_run(
        "1. Pemantauan pembiasaan di luar jam sekolah (seperti jam tidur malam dan pengurangan screen-time gawai di rumah) masih bergantung pada komitmen dan kesibukan masing-masing orang tua.\n"
        "2. Fasilitas sarana prasarana ramah anak dan ruang konseling/BK masih terus dalam proses pembenahan dan peningkatan bertahap.\n"
        "3. Keterbatasan waktu koordinasi intensif bersama wali murid yang mayoritas bekerja penuh waktu."
    ).font.size = Pt(9)
    p_ken.paragraph_format.space_after = Pt(8)

    # 8. RENCANA TINDAK LANJUT
    p_rtl = doc.add_paragraph()
    r_rtlt = p_rtl.add_run("RENCANA TINDAK LANJUT\n")
    r_rtlt.bold = True
    r_rtlt.font.size = Pt(10)
    p_rtl.add_run(
        "1. Mengoptimalkan buku penghubung/jurnal digital pembiasaan siswa yang terhubung langsung secara berkala dengan orang tua/wali murid.\n"
        "2. Menyelenggarakan kegiatan parenting dan sosialisasi berkala mengenai pentingnya pola asuh berkarakter, gizi seimbang, dan pembatasan gawai di rumah.\n"
        "3. Melakukan pemeliharaan sarana prasarana sekolah secara berkala melalui anggaran RKAS untuk memastikan keamanan lingkungan belajar peserta didik.\n"
        "4. Meningkatkan kapasitas guru kelas dan guru BK melalui pelatihan konseling ramah anak dan deteksi dini pencegahan kekerasan di sekolah."
    ).font.size = Pt(9)
    p_rtl.paragraph_format.space_after = Pt(14)

    # 9. TANDA TANGAN
    t_ttd = doc.add_table(rows=1, cols=2)
    t_ttd.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in t_ttd.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))

    c_ks = t_ttd.cell(0, 0)
    c_ks.width = Inches(3.5)
    p_ks = c_ks.paragraphs[0]
    p_ks.add_run(f"Mengetahui,\nKepala {sch.get('name', 'SD NEGERI 1 KESAMBI')}\n\n\n\n\n").font.size = Pt(9.5)
    r_kn = p_ks.add_run(f"{sch.get('ks', 'Dra. Hj. Siti Maryam, M.Pd')}\n")
    r_kn.bold = True
    r_kn.underline = True
    r_kn.font.size = Pt(9.5)
    p_ks.add_run(f"NIP. {sch.get('ks_nip', '197505121998032004')}").font.size = Pt(9)

    c_p = t_ttd.cell(0, 1)
    c_p.width = Inches(3.5)
    p_p = c_p.paragraphs[0]
    p_p.add_run(f"Cirebon, 13 November {sch.get('year', '2026')}\nPengawas / Pemantau,\n\n\n\n\n").font.size = Pt(9.5)
    r_pn = p_p.add_run(f"{sch.get('pengawas', 'Ari Haerani, S.Pd.SD, M.Pd')}\n")
    r_pn.bold = True
    r_pn.underline = True
    r_pn.font.size = Pt(9.5)
    p_p.add_run(f"NIP. {sch.get('pengawas_nip', '197301211996032001')}").font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 10. CATATAN
    p_cat = doc.add_paragraph()
    p_cat.add_run("Catatan Tambahan:\n").bold = True
    p_cat.add_run(
        f"1. Mohon dilampiri berkas bukti fisik sesuai indikator dan diunggah pada tautan sekolah.\n"
        f"2. Format Penamaan File: {sch.get('name', 'SDN1KESAMBI').replace(' ', '_')}_Instrumen_JabarBeraksi_{sch.get('year', '2026')}\n"
        f"3. Tautan Bukti Sekolah (Google Drive): {sch.get('link_drive', 'https://drive.google.com/drive/folders/1SDN1Kesambi_JabarBeraksi2026')}"
    ).font.size = Pt(8.5)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_indicator_docx(poin_num, sch=None):
    sch = {**DEFAULT_SCHOOL, **(sch or {})}
    from engine.indicator_data import INDICATOR_DETAILS_MAP

    details = INDICATOR_DETAILS_MAP.get(poin_num, {})
    matched = next((item for item in INDICATORS_27 if item[0] == poin_num), (poin_num, f"Indikator Poin {poin_num}", ["SOP", "Log", "Laporan"], "95%"))
    num, default_title, default_ev, pct = matched
    
    title = details.get('sub', default_title)
    ev_desc = details.get('ev', ', '.join(default_ev))

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # KOP SURAT
    p_kop = doc.add_paragraph()
    p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop.paragraph_format.space_after = Pt(2)
    r = p_kop.add_run(f"PEMERINTAH KOTA CIREBON\nDINAS PENDIDIKAN\n{sch.get('name', 'SD NEGERI 1 KESAMBI')}")
    r.bold = True
    r.font.size = Pt(12)
    r.font.name = 'Arial'

    p_kop2 = doc.add_paragraph()
    p_kop2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop2.paragraph_format.space_after = Pt(8)
    r2 = p_kop2.add_run(f"Alamat: {sch.get('address', 'Jalan Kesambi No. 24, Kota Cirebon')}, Kec. {sch.get('kec', 'Kesambi')}\nPosel: sdnegeri@cirebonkota.go.id | Laman: disdik.cirebonkota.go.id")
    r2.font.size = Pt(8.5)
    r2.italic = True
    r2.font.color.rgb = RGBColor(100, 100, 100)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    add_double_divider(p_div)

    # TITLE
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_after = Pt(2)
    rt = p_t.add_run(f"BUKTI FISIK INSTRUMEN JABAR BERAKSI (POIN {num})")
    rt.bold = True
    rt.font.size = Pt(12)

    p_s = doc.add_paragraph()
    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_s.paragraph_format.space_after = Pt(14)
    rs = p_s.add_run(f"INDIKATOR: {title.upper()}\n(Bukti Fisik Terverifikasi: {ev_desc})")
    rs.bold = True
    rs.font.size = Pt(9.5)
    rs.font.color.rgb = RGBColor(13, 71, 161)

    # BAGIAN 1: SURAT TUGAS / SOP
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(8)
    h1.paragraph_format.space_after = Pt(4)
    st_title_text = details.get('st_title', "A. SURAT TUGAS PELAKSANAAN PEMBIASAAN / STANDAR OPERASIONAL PROSEDUR")
    rh1 = h1.add_run(st_title_text)
    rh1.bold = True
    rh1.font.size = Pt(10.5)
    rh1.font.color.rgb = RGBColor(13, 71, 161)

    st_table = doc.add_table(rows=6, cols=2)
    st_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(st_table)

    st_task_text = details.get('st_task', f"1. Mengorganisasikan pembiasaan {title} secara konsisten dan terukur.\n2. Melakukan pendampingan afektif dan pencatatan harian di buku log.\n3. Melakukan evaluasi mingguan dan memberikan apresiasi bagi siswa berkarakter teladan.")
    role_lead = details.get('role_lead', 'Pembina Kesiswaan & Koordinator Program')

    head_name = sch.get('head') or sch.get('ks') or 'Drs. H. Mulyono, M.Pd.'
    head_nip = sch.get('nip') or sch.get('ks_nip') or '196805141992031004'

    st_data = [
        ("1. Nomor Regulasi / Surat Tugas", f"Nomor: 421.2 / 0{num + 70} / SD-KS / X / {sch.get('year', '2026')}"),
        ("2. Dasar Pelaksanaan", f"a. Panduan Penguatan Pendidikan Karakter (PPK) Jabar Beraksi Kota Cirebon {sch.get('year', '2026')}.\nb. Rencana Kerja Tahunan (RKT) {sch.get('name', 'SD NEGERI 1 KESAMBI')} Bidang Kesiswaan."),
        ("3. Pejabat Penanggung Jawab", f"Nama : {head_name}\nJabatan : Kepala Sekolah\nNIP : {head_nip}"),
        ("4. Personel yang Ditugaskan", f"1. {role_lead}\n2. Seluruh Guru Kelas (I s.d. VI) dan Guru Mata Pelajaran\n3. Pengurus Komite / Paguyuban Orang Tua Murid\n4. Seluruh Peserta Didik {sch.get('name', 'SD NEGERI 1 KESAMBI')}"),
        ("5. Uraian Prosedur & Tugas", st_task_text),
        ("6. Waktu & Lokasi", f"Waktu : Terjadwal Rutin Semester Ganjil Tahun Ajaran {sch.get('year', '2026')}/2027\nLokasi : Lingkungan {sch.get('name', 'SD NEGERI 1 KESAMBI')} dan Lingkungan Keluarga")
    ]

    col_widths = [Inches(2.0), Inches(4.8)]
    for r_idx, (c0_text, c1_text) in enumerate(st_data):
        row = st_table.rows[r_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 80, 80, 110, 110)
        set_cell_margins(c1, 80, 80, 110, 110)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(c0_text)
        r0.bold = True
        r0.font.size = Pt(8.5)
        p1 = c1.paragraphs[0]
        p1.add_run(c1_text).font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # BAGIAN 2: LAPORAN KEGIATAN
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(8)
    h2.paragraph_format.space_after = Pt(4)
    rh2 = h2.add_run("B. LAPORAN PELAKSANAAN PEMBIASAAN KARAKTER")
    rh2.bold = True
    rh2.font.size = Pt(10.5)
    rh2.font.color.rgb = RGBColor(13, 71, 161)

    rep_table = doc.add_table(rows=6, cols=2)
    rep_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rep_table)

    rep_title_text = details.get('rep_title', f"Program Pembiasaan Karakter Unggul: {title}")
    rep_bg_text = details.get('rep_bg', f"Menjamin keterlaksanaan pembiasaan {title} secara tertib, aman, terukur, dan berdampak positif bagi pembentukan budi pekerti luhur peserta didik.")
    rep_scope_text = details.get('rep_scope', "a. Sosialisasi awal tahun kepada seluruh warga sekolah.\nb. Pelaksanaan terintegrasi dalam budaya sekolah harian.\nc. Pendampingan berkala wali kelas dan pencatatan buku jurnal.")
    rep_impact_text = details.get('rep_impact', f"Tingkat partisipasi aktif peserta didik mencapai {pct}, menunjukkan peningkatan kedisiplinan dan respon positif.")
    rep_obstacle_text = details.get('rep_obstacle', "Perbedaan latar belakang kebiasaan di rumah yang membutuhkan komunikasi intensif secara berkala.")
    rep_followup_text = details.get('rep_followup', "Mengembangkan lembar refleksi mandiri siswa serta penguatan kolaborasi melalui buku penghubung dan pertemuan paguyuban.")

    rep_data = [
        ("1. Nama Program", rep_title_text),
        ("2. Latar Belakang & Tujuan", rep_bg_text),
        ("3. Ruang Lingkup Aktivitas", rep_scope_text),
        ("4. Capaian & Keterlaksanaan", rep_impact_text),
        ("5. Kendala Lapangan", rep_obstacle_text),
        ("6. Rencana Tindak Lanjut", rep_followup_text)
    ]

    for r_idx, (c0_text, c1_text) in enumerate(rep_data):
        row = rep_table.rows[r_idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, 'F5F7FA')
        set_cell_margins(c0, 80, 80, 110, 110)
        set_cell_margins(c1, 80, 80, 110, 110)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(c0_text)
        r0.bold = True
        r0.font.size = Pt(8.5)
        p1 = c1.paragraphs[0]
        p1.add_run(c1_text).font.size = Pt(8.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # BAGIAN 3: JURNAL LOG HARIAN
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)
    rh3 = h3.add_run("C. JURNAL LOG HARIAN & BUKU CATATAN AKTIVITAS PEMBIASAAN")
    rh3.bold = True
    rh3.font.size = Pt(10.5)
    rh3.font.color.rgb = RGBColor(13, 71, 161)

    log_headers = details.get('log_headers', ['No', 'Hari/Tanggal', 'Aktivitas Pembiasaan', 'Lokasi / Sasaran', 'Catatan Hasil Pengamatan', 'Status'])
    sample_logs = details.get('logs', [
        ('1', f"Senin, 14/09/{sch.get('year', '2026')}", f"Pelaksanaan pembiasaan {title} terprogram", "Lingkungan Sekolah / Kelas", "Peserta didik antusias, tertib, dan mematuhi panduan guru.", "Terlaksana\n(100%)"),
        ('2', f"Rabu, 07/10/{sch.get('year', '2026')}", f"Monitoring berkala dan pengisian jurnal refleksi", "Wali Kelas & Peserta Didik", f"{pct} siswa konsisten mempraktikkan kebiasaan positif secara mandiri.", "Terlaksana\n(98%)"),
        ('3', f"Jumat, 23/10/{sch.get('year', '2026')}", f"Evaluasi bulanan dan pemberian apresiasi karakter", "Dewan Guru & Murid", "Terjadi peningkatan kedisiplinan dan kebiasaan positif yang konsisten.", "Terlaksana\n(100%)")
    ])

    log_col_w_nums = details.get('log_col_w')
    if not log_col_w_nums or len(log_col_w_nums) != len(log_headers):
        log_col_w_nums = [6.8 / len(log_headers)] * len(log_headers)

    log_table = doc.add_table(rows=len(sample_logs) + 1, cols=len(log_headers))
    log_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(log_table)

    for idx, text in enumerate(log_headers):
        cell = log_table.rows[0].cells[idx]
        cell.width = Inches(log_col_w_nums[idx])
        set_cell_background(cell, '1565C0')
        set_cell_margins(cell, 90, 90, 70, 70)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, data_row in enumerate(sample_logs):
        row = log_table.rows[r_idx + 1]
        bg_col = 'FFFFFF' if r_idx % 2 == 0 else 'F8FAFC'
        for c_idx, val in enumerate(data_row):
            if c_idx >= len(log_headers):
                break
            cell = row.cells[c_idx]
            cell.width = Inches(log_col_w_nums[c_idx])
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, 80, 80, 70, 70)
            p = cell.paragraphs[0]
            if c_idx in [0, 1, len(log_headers) - 1]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(val).font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # BAGIAN 4: DOKUMENTASI FOTO
    h4 = doc.add_paragraph()
    h4.paragraph_format.space_before = Pt(8)
    h4.paragraph_format.space_after = Pt(4)
    rh4 = h4.add_run("D. LEMBAR DOKUMENTASI FOTO BUKTI FISIK")
    rh4.bold = True
    rh4.font.size = Pt(10.5)
    rh4.font.color.rgb = RGBColor(13, 71, 161)

    photo_table = doc.add_table(rows=1, cols=2)
    photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(photo_table)

    photo_labels = details.get('photo_labels', (
        f"[ TEMPELKAN FOTO DOKUMENTASI 1 ]\n\nPelaksanaan kegiatan {title} bersama peserta didik di sekolah.",
        f"[ TEMPELKAN FOTO DOKUMENTASI 2 ]\n\nMonitoring dan evaluasi bersama dewan guru dan paguyuban orang tua murid."
    ))

    c0, c1 = photo_table.rows[0].cells[0], photo_table.rows[0].cells[1]
    c0.width, c1.width = Inches(3.4), Inches(3.4)
    set_cell_background(c0, 'FAFAFA')
    set_cell_background(c1, 'FAFAFA')
    set_cell_margins(c0, 120, 120, 120, 120)
    set_cell_margins(c1, 120, 120, 120, 120)
    c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c0.paragraphs[0].add_run(photo_labels[0]).font.size = Pt(8)
    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c1.paragraphs[0].add_run(photo_labels[1]).font.size = Pt(8)

    # TANDA TANGAN RESMI
    doc.add_paragraph().paragraph_format.space_after = Pt(14)
    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_after = Pt(6)
    p_dt.add_run(f"Cirebon, 13 November {sch.get('year', '2026')}").font.size = Pt(9.5)

    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in sig_table.rows:
        for c in r.cells:
            tcPr = c._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))

    c0, c1, c2 = sig_table.rows[0].cells[0], sig_table.rows[0].cells[1], sig_table.rows[0].cells[2]
    c0.width, c1.width, c2.width = Inches(2.2), Inches(2.4), Inches(2.2)
    c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c0.paragraphs[0].add_run(f"{role_lead},\n\n\n\n\n( ......................................... )\nNIP. .....................................").font.size = Pt(8.5)

    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c1.paragraphs[0].add_run(f"Pengawas Pembina,\n{sch.get('supervisor_region', 'Wilayah Kerja Harjamukti & Kesambi')}\n\n\n\n{sch.get('supervisor', 'Ari Haerani, S.Pd.SD, M.Pd')}\nNIP. {sch.get('supervisor_nip', '197301211996032001')}").font.size = Pt(8.5)

    c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c2.paragraphs[0].add_run(f"Mengetahui,\nKepala {sch.get('name', 'SD NEGERI 1 KESAMBI')},\n\n\n\n{head_name}\nNIP. {head_nip}").font.size = Pt(8.5)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_custom_ai_docx(title, evidence_types, raw_ai_text, sch=None):
    sch = {**DEFAULT_SCHOOL, **(sch or {})}
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    is_modul_ajar = any(k in title.lower() for k in [
        'modul', 'rpp', 'perangkat', 'buku', 'kurikulum', 'pembelajaran', 'bab', 
        'pai', 'pendidikan agama', 'ipas', 'matematika', 'bahasa', 'pjok', 'seni'
    ])

    # KOP SURAT
    p_kop = doc.add_paragraph()
    p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop.paragraph_format.space_after = Pt(2)
    r = p_kop.add_run(f"PEMERINTAH KOTA CIREBON\nDINAS PENDIDIKAN\n{sch.get('name', 'SD NEGERI 1 KESAMBI')}")
    r.bold = True
    r.font.size = Pt(12)
    r.font.name = 'Arial'

    p_kop2 = doc.add_paragraph()
    p_kop2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop2.paragraph_format.space_after = Pt(8)
    r2 = p_kop2.add_run(f"Alamat: {sch.get('address', 'Jalan Kesambi No. 24, Kota Cirebon')}, Kec. {sch.get('kec', 'Kesambi')}\nPosel: sdnegeri@cirebonkota.go.id | Laman: disdik.cirebonkota.go.id")
    r2.font.size = Pt(8.5)
    r2.italic = True
    r2.font.color.rgb = RGBColor(100, 100, 100)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(12)
    add_double_divider(p_div)

    # TITLE
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_after = Pt(2)
    if is_modul_ajar:
        rt = p_t.add_run("PERANGKAT PEMBELAJARAN / MODUL AJAR KURIKULUM MERDEKA")
    else:
        rt = p_t.add_run("BERKAS BUKTI FISIK PEMANTAUAN PENDIDIKAN")
    rt.bold = True
    rt.font.size = Pt(12)

    p_s = doc.add_paragraph()
    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_s.paragraph_format.space_after = Pt(14)
    if is_modul_ajar:
        rs = p_s.add_run(f"MATA PELAJARAN / TOPIK: {title.upper()}\n(Pendekatan Deep Learning & Penguatan Profil Pelajar Pancasila)")
        rs.font.color.rgb = RGBColor(46, 125, 50) # Emerald
    else:
        rs = p_s.add_run(f"PROGRAM / INDIKATOR: {title.upper()}\n(Bukti Fisik: {evidence_types or 'Surat Tugas, Laporan Kegiatan, Jurnal Log, Dokumentasi'})")
        rs.font.color.rgb = RGBColor(13, 71, 161) # Deep Blue
    rs.bold = True
    rs.font.size = Pt(9.5)

    # Parse AI sections into tables and structured blocks
    lines = raw_ai_text.split('\n')
    current_section = ''
    section_items = []
    table_rows = []

    def flush_section():
        nonlocal current_section, section_items, table_rows
        if not current_section:
            return

        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(4)
        rh = h.add_run(current_section)
        rh.bold = True
        rh.font.size = Pt(10.5)
        rh.font.color.rgb = RGBColor(46, 125, 50) if is_modul_ajar else RGBColor(13, 71, 161)

        if table_rows:
            num_cols = len(table_rows[0])
            tbl = doc.add_table(rows=len(table_rows), cols=num_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(tbl)
            for r_idx, r_cells in enumerate(table_rows):
                row = tbl.rows[r_idx]
                is_header = (r_idx == 0)
                bg = ('2E7D32' if is_modul_ajar else '1565C0') if is_header else ('F8FAFC' if r_idx % 2 == 0 else 'FFFFFF')
                for c_idx, val in enumerate(r_cells):
                    if c_idx >= num_cols:
                        break
                    cell = row.cells[c_idx]
                    set_cell_background(cell, bg)
                    set_cell_margins(cell, 80, 80, 70, 70)
                    p = cell.paragraphs[0]
                    if is_header or c_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(val)
                    r.font.size = Pt(8)
                    if is_header:
                        r.bold = True
                        r.font.color.rgb = RGBColor(255, 255, 255)
            table_rows = []
            doc.add_paragraph().paragraph_format.space_after = Pt(8)
        elif section_items:
            tbl = doc.add_table(rows=len(section_items), cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            set_table_borders(tbl)
            col_widths = [Inches(2.0), Inches(4.8)]
            for r_idx, (c0_text, c1_text) in enumerate(section_items):
                row = tbl.rows[r_idx]
                c0, c1 = row.cells[0], row.cells[1]
                c0.width, c1.width = col_widths[0], col_widths[1]
                set_cell_background(c0, 'F5F7FA')
                set_cell_margins(c0, 80, 80, 110, 110)
                set_cell_margins(c1, 80, 80, 110, 110)
                p0 = c0.paragraphs[0]
                r0 = p0.add_run(c0_text)
                r0.bold = True
                r0.font.size = Pt(8.5)
                p1 = c1.paragraphs[0]
                p1.add_run(c1_text).font.size = Pt(8.5)
            section_items = []
            doc.add_paragraph().paragraph_format.space_after = Pt(8)

    def is_header_line(t):
        if not t:
            return False
        if t.startswith('#'):
            return True
        if t.startswith('BAGIAN') or t.startswith('BAB') or t.startswith('KOMPONEN'):
            return True
        if (t.startswith('**') and t.endswith('**') and len(t) < 80) or t.startswith('**BAGIAN') or t.startswith('**BAB'):
            return True
        if re.match(r'^[A-Z]\.\s+[A-Z\s]{3,}', t):
            return True
        if re.match(r'^[0-9]+\.\s+[A-Z\s]{3,}', t):
            return True
        return False

    for line in lines:
        t = line.strip()
        if not t:
            continue

        if is_header_line(t):
            flush_section()
            current_section = re.sub(r'^[#*]+', '', t).replace('*', '').strip()
            continue

        if t.startswith('|') and t.endswith('|'):
            cells = [c.strip() for c in t.split('|')[1:-1]]
            if any('---' in c for c in cells):
                continue
            table_rows.append(cells)
            continue

        if (t.startswith('- ') or t.startswith('* ') or re.match(r'^[0-9]+\.', t)) and ':' in t:
            clean = re.sub(r'^[-*0-9.]+\s*', '', t)
            parts = clean.split(':', 1)
            section_items.append((parts[0].strip(), parts[1].strip()))
            continue

        if current_section:
            section_items.append(("Uraian", t))

    flush_section()

    # DOKUMENTASI FOTO (Hanya jika bukti fisik atau jika ada kata foto)
    has_photo_req = not is_modul_ajar or ('foto' in raw_ai_text.lower() or 'dokumentasi' in raw_ai_text.lower())
    if has_photo_req:
        h4 = doc.add_paragraph()
        h4.paragraph_format.space_before = Pt(8)
        h4.paragraph_format.space_after = Pt(4)
        h4.add_run("LEMBAR DOKUMENTASI FOTO").bold = True
        h4.runs[0].font.size = Pt(10.5)
        h4.runs[0].font.color.rgb = RGBColor(13, 71, 161)

        photo_table = doc.add_table(rows=1, cols=2)
        photo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(photo_table)
        c0, c1 = photo_table.rows[0].cells[0], photo_table.rows[0].cells[1]
        c0.width, c1.width = Inches(3.4), Inches(3.4)
        set_cell_background(c0, 'FAFAFA')
        set_cell_background(c1, 'FAFAFA')
        set_cell_margins(c0, 120, 120, 120, 120)
        set_cell_margins(c1, 120, 120, 120, 120)
        c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c0.paragraphs[0].add_run(f"[ TEMPELKAN FOTO DOKUMENTASI 1 ]\n\nPelaksanaan aktivitas {title} di kelas/sekolah.").font.size = Pt(8)
        c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c1.paragraphs[0].add_run(f"[ TEMPELKAN FOTO DOKUMENTASI 2 ]\n\nEvaluasi dan unjuk kerja hasil belajar peserta didik.").font.size = Pt(8)

    # TANDA TANGAN RESMI
    doc.add_paragraph().paragraph_format.space_after = Pt(14)
    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.paragraph_format.space_after = Pt(6)
    p_dt.add_run(f"Cirebon, 13 November {sch.get('year', '2026')}").font.size = Pt(9.5)

    sig_table = doc.add_table(rows=1, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r in sig_table.rows:
        for c in r.cells:
            tcPr = c._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))

    head_name = sch.get('head') or sch.get('ks') or 'Drs. H. Mulyono, M.Pd.'
    head_nip = sch.get('nip') or sch.get('ks_nip') or '196805141992031004'

    c0, c1, c2 = sig_table.rows[0].cells[0], sig_table.rows[0].cells[1], sig_table.rows[0].cells[2]
    c0.width, c1.width, c2.width = Inches(2.2), Inches(2.4), Inches(2.2)
    
    role_c0 = "Guru Pengampu / Wali Kelas" if is_modul_ajar else "Pembina Program / Kesiswaan"
    c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c0.paragraphs[0].add_run(f"{role_c0},\n\n\n\n\n( ......................................... )\nNIP. .....................................").font.size = Pt(8.5)

    c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c1.paragraphs[0].add_run(f"Pengawas Pembina,\n{sch.get('supervisor_region', 'Wilayah Kerja Harjamukti & Kesambi')}\n\n\n\n{sch.get('supervisor', 'Ari Haerani, S.Pd.SD, M.Pd')}\nNIP. {sch.get('supervisor_nip', '197301211996032001')}").font.size = Pt(8.5)

    c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c2.paragraphs[0].add_run(f"Mengetahui,\nKepala {sch.get('name', 'SD NEGERI 1 KESAMBI')},\n\n\n\n{head_name}\nNIP. {head_nip}").font.size = Pt(8.5)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
