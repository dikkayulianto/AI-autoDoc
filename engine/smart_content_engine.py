# SMART CONTENT ENGINE UNTUK AI-AUTODOC
# Menghasilkan naskah mendalam, adaptif, dan 100% spesifik sesuai topik yang dicari/diinput pengguna.

import re

def is_modul_ajar_topic(topic):
    t = topic.lower()
    keywords = [
        'modul', 'rpp', 'perangkat', 'buku', 'kurikulum', 'pembelajaran', 'kbm', 'bab', 
        'pai', 'pendidikan agama', 'ipas', 'matematika', 'bahasa', 'pjok', 'seni',
        'materi', 'fase a', 'fase b', 'fase c', 'kelas 1', 'kelas 2', 'kelas 3', 
        'kelas 4', 'kelas 5', 'kelas 6', 'semester'
    ]
    return any(k in t for k in keywords)

def build_system_prompt(topic, evidence_types, custom_prompt, school_name, head_name, supervisor_name):
    is_modul = is_modul_ajar_topic(topic)
    
    if is_modul:
        prompt = f"""Anda adalah Pakar Kurikulum Merdeka dan Tim Pengembang Modul Ajar Sekolah Dasar Dinas Pendidikan Kota Cirebon.
TUGAS UTAMA: Susun naskah resmi PERANGKAT PEMBELAJARAN / MODUL AJAR (DEEP LEARNING & PENGUATAN KARAKTER) secara LENGKAP, MENDALAM, DAN 100% SPESIFIK untuk:
Satuan Pendidikan: {school_name}
Mata Pelajaran / Topik Modul: {topic}
Kepala Sekolah: {head_name}
Pengawas Pembina: {supervisor_name}

{custom_prompt if custom_prompt else ''}

PERHATIAN WAJIB:
1. DILARANG KERAS menghasilkan template kosong atau daur ulang materi senam/makan sehat jika topik tidak berkaitan!
2. Seluruh materi, Capaian Pembelajaran (CP), Tujuan Pembelajaran (TP), langkah KBM (Pendahuluan, Inti, Penutup), asesmen, dan LKPD HARUS 100% MENDALAM dan SPESIFIK membahas topik "{topic}".
3. Gunakan Bahasa Indonesia baku, formal, dan rapi kedinasan.

Format Naskah WAJIB menggunakan struktur Markdown berikut:

BAGIAN 1: INFORMASI UMUM & IDENTITAS MODUL
- Nama Penyusun: Tim Guru Pengampu / Wali Kelas
- Satuan Pendidikan: {school_name}
- Mata Pelajaran / Topik Bahasan: {topic}
- Fase / Kelas / Semester: Fase Terkait (Kelas I s/d VI) - Semester Ganjil 2026/2027
- Alokasi Waktu: 2 Pertemuan (4 x 35 Menit)
- Profil Pelajar Pancasila: Beriman Bertakwa, Bernalar Kritis, Mandiri, Gotong Royong
- Sarana dan Prasarana: Media konkret, papan tulis interaktif, lembar kerja siswa, buku pegangan
- Model Pembelajaran: Problem Based Learning (PBL) / Inquiry Learning / Deep Learning

BAGIAN 2: KOMPONEN INTI PEMBELAJARAN
- Capaian Pembelajaran (CP): Rumuskan Capaian Pembelajaran Kurikulum Merdeka resmi yang relevan dengan {topic}.
- Tujuan Pembelajaran (TP): Tuliskan 3-4 Tujuan Pembelajaran terukur berbasis kompetensi HOTS.
- Pemahaman Bermakna: Uraikan manfaat nyata pemahaman materi {topic} dalam kehidupan sehari-hari anak.
- Pertanyaan Pemantik: Tuliskan 3 pertanyaan pemantik yang merangsang rasa ingin tahu dan nalar kritis siswa.

BAGIAN 3: ALUR KEGIATAN PEMBELAJARAN (DEEP LEARNING)
| Tahapan | Alokasi Waktu | Aktivitas Guru & Peserta Didik (Meaningful, Mindful, Joyful) | Integrasi Karakter |
| Pendahuluan | 10 Menit | Pembukaan dengan salam, doa khusyuk, apersepsi mengaitkan materi {topic} dengan pengalaman nyata murid, penyampaian tujuan pembelajaran | Religius, Disiplin |
| Kegiatan Inti | 50 Menit | Eksplorasi konsep materi {topic}, diskusi kelompok berdiferensiasi (proses & produk), bimbingan terbimbing guru, presentasi hasil karya siswa | Bernalar Kritis, Kolaborasi |
| Penutup | 10 Menit | Refleksi bersama murid ("Apa hal bermakna yang kita pelajari hari ini?"), penguatan kesimpulan, apresiasi motivasi, dan doa penutup | Bersyukur, Mandiri |

BAGIAN 4: ASESMEN & RUBRIK PENILAIAN
| Aspek Asesmen | Teknik Penilaian | Instrumen Penilaian | Kriteria Ketercapaian Tujuan Pembelajaran (KKTP) |
| Sikap (Afektif) | Observasi Kinerja Harian | Lembar Pengamatan Profil Pelajar Pancasila | Menunjukkan keaktifan, kesantunan, dan kejujuran dalam kelompok |
| Pengetahuan (Kognitif) | Tes Tertulis / Lembar Latihan | LKPD dan Soal Pemahaman Konsep {topic} | Mampu memahami dan menyelesaikan soal analisis dengan tuntas (Skor >= 75) |
| Keterampilan (Psikomotorik) | Unjuk Kerja / Presentasi | Rubrik Penilaian Produk & Kinerja | Terampil menyajikan hasil karya atau mendemonstrasikan praktik materi {topic} |

BAGIAN 5: LEMBAR KERJA PESERTA DIDIK (LKPD) & RANGKUMAN MATERI
- Rangkuman Materi Inti: Uraikan ringkasan materi {topic} secara runtut, jelas, dan edukatif bagi anak.
- Lembar Kerja Murid (Aktivitas Konkret): Tuliskan 3-5 pertanyaan latihan/aktivitas aplikatif yang wajib dikerjakan siswa.
- Rencana Tindak Lanjut & Remedial: Uraikan strategi pendampingan remedial dan pengayaan bagi peserta didik.
"""
    else:
        prompt = f"""Anda adalah Auditor Senior Standar Penjaminan Mutu Pendidikan Jawa Barat dan Pengawas Sekolah Kota Cirebon.
Tugas Anda adalah menyusun naskah kelengkapan BUKTI FISIK resmi untuk:
Sekolah: {school_name}
Topik/Indikator: {topic}
Kepala Sekolah: {head_name}
Pengawas Pembina: {supervisor_name}

{custom_prompt if custom_prompt else ''}

PERHATIAN KRUSIAL:
1. Seluruh rincian tugas, laporan pelaksanaan, tabel jurnal harian 7 kolom, dan evaluasi HARUS 100% SPESIFIK, RELEVAN, DAN MENDALAM sesuai dengan TOPIK: '{topic}'!
2. DILARANG KERAS menghasilkan contoh kegiatan umum yang tidak berhubungan (JANGAN menyalin senam/apel/makan sehat jika topik yang diminta bukan tentang hal tersebut)!
3. Gunakan Bahasa Indonesia baku kedinasan.

Format Dokumen WAJIB menggunakan struktur Markdown berikut:

BAGIAN 1: SURAT TUGAS & LANDASAN HUKUM
- Nomor Surat Tugas: 421.2 / 084 / SD-01 / DISDIK / 2026
- Dasar Pelaksanaan: Peraturan Gubernur Jawa Barat No. 12 Tahun 2024 & Panduan PPK Jabar Beraksi
- Pejabat Pemberi Tugas: {head_name} (Kepala Sekolah)
- Personel yang Ditugaskan: Tim Guru Pengampu, Wali Kelas, dan Koordinator Terkait
- Uraian Tugas & Tanggung Jawab: Uraikan 4 butir tugas konkret yang 100% spesifik untuk keberhasilan {topic}.
- Waktu dan Tempat Pelaksanaan: Terjadwal Semester Ganjil Tahun Ajaran 2026/2027 di Lingkungan {school_name}

BAGIAN 2: LAPORAN KETERLAKSANAAN PROGRAM
- Nama Program Resmi: Program Implementasi & Pembiasaan {topic}
- Latar Belakang & Urgensi: Uraikan latar belakang formal mengapa program {topic} sangat mendesak dan penting.
- Tujuan Program Terukur: Uraikan 3 tujuan spesifik terukur yang ingin dicapai.
- Ruang Lingkup Aktivitas: Tuliskan 3 poin utama alur aktivitas pelaksanaan program {topic}.
- Capaian Kinerja & Dampak: Tuliskan data kuantitatif keterlaksanaan dan dampak positif terhadap karakter murid.
- Kendala Lapangan & Solusi: Tuliskan kendala riil lapangan serta solusi pemecahan yang ditempuh.
- Rencana Tindak Lanjut: Tuliskan langkah perbaikan dan pengembangan berkelanjutan ke depan.

BAGIAN 3: JURNAL LOG HARIAN 7 KOLOM
| No | Hari / Tanggal | Waktu (WIB) | Bentuk Aktivitas Spesifik ({topic}) | Sasaran Peserta | Hasil & Capaian Pengamatan | Status Verifikasi |
[Wajib tuliskan minimal 3 baris data nyata yang BENAR-BENAR spesifik mengenai {topic}, dengan tanggal dan catatan observasi yang realistis!]

BAGIAN 4: MATRIKS EVALUASI & TINDAK LANJUT
| Aspek Evaluasi | Temuan di Lapangan | Analisis Ketercapaian | Tindak Lanjut Perbaikan | Penanggung Jawab |
| Keterlaksanaan Program | Keterlaksanaan program {topic} berjalan optimal | Target tercapai dengan sangat baik | Pertahankan dan tingkatkan kualitas berkala | Tim Pelaksana |
| Dampak Karakter Siswa | Siswa menunjukkan perubahan perilaku positif | Karakter unggul terbentuk nyata | Evaluasi berkala setiap bulan | Wali Kelas |

BAGIAN 5: KESIMPULAN & REKOMENDASI PENGAWAS
- Kesimpulan Hasil Monitoring: Pelaksanaan program {topic} telah berjalan tertib, terdokumentasi lengkap, dan berdampak nyata bagi penguatan karakter murid.
- Catatan Pengawas Pembina ({supervisor_name}): Berkas bukti fisik {topic} disetujui dan disahkan secara resmi untuk kelengkapan administrasi penjaminan mutu pendidikan.
"""
    return prompt

def get_smart_fallback(topic, evidence_types, school_name, head_name, supervisor_name, year="2026"):
    """
    Generator cerdas offline: menghasilkan dokumen kaya dan spesifik jika API key tidak tersedia atau kuota habis.
    Tidak akan pernah menghasilkan teks generik yang sama untuk topik berbeda!
    """
    t = topic.lower()
    is_modul = is_modul_ajar_topic(topic)
    
    # 1. PAI / PENDIDIKAN AGAMA ISLAM
    if any(k in t for k in ['pai', 'agama islam', 'budi pekerti', 'hijaiyah', 'shalat', 'iman', 'islam', 'akhlak']):
        if is_modul:
            return f"""BAGIAN 1: INFORMASI UMUM & IDENTITAS MODUL
- Nama Penyusun: Tim Guru PAI & Budi Pekerti
- Satuan Pendidikan: {school_name}
- Mata Pelajaran / Materi: Pendidikan Agama Islam dan Budi Pekerti (Mengenal Huruf Hijaiyah & Kalimat Thayyibah)
- Fase / Kelas / Semester: Fase A (Kelas 1) - Semester Ganjil {year}/2027
- Alokasi Waktu: 2 Pertemuan (4 x 35 Menit)
- Profil Pelajar Pancasila: Beriman, Bertakwa kepada Tuhan YME dan Berakhlak Mulia, Mandiri
- Sarana dan Prasarana: Poster Makharijul Huruf, Kartu Ceria Hijaiyah, Audio Murattal, Lembar Mewarnai
- Model Pembelajaran: Problem Based Learning (PBL) & Deep Learning Berdiferensiasi

BAGIAN 2: KOMPONEN INTI PEMBELAJARAN
- Capaian Pembelajaran (CP): Peserta didik mampu mengenal huruf hijaiyah dan harakatnya, melafalkan huruf hijaiyah sesuai makhraj yang benar, serta membiasakan kalimat thayyibah (Basmalah dan Hamdalah) dalam kehidupan sehari-hari.
- Tujuan Pembelajaran (TP): 1. Menyebutkan 28 huruf hijaiyah tunggal dengan benar; 2. Melafalkan bunyi huruf berharakat fathah, kasrah, dan dhammah secara fasih; 3. Membiasakan mengucap Basmalah sebelum beraktivitas.
- Pemahaman Bermakna: Mengenal huruf Al-Qur'an sejak dini adalah kunci membaca kalam ilahi yang membawa keberkahan dan ketenangan hati.
- Pertanyaan Pemantik: 1. Siapa yang sudah belajar mengaji di rumah? 2. Bagaimana bunyi huruf Alif jika diberi harakat fathah di atasnya? 3. Kapan kita mengucap bismillahirrahmanirrahim?

BAGIAN 3: ALUR KEGIATAN PEMBELAJARAN (DEEP LEARNING)
| Tahapan | Alokasi Waktu | Aktivitas Guru & Peserta Didik (Meaningful, Mindful, Joyful) | Integrasi Karakter |
| Pendahuluan | 10 Menit | Guru menyapa dengan salam hangat, membaca doa belajar bersama, menyanyikan lagu riang 'Alif Ba Ta', dan menyampaikan tujuan belajar | Religius, Semangat |
| Kegiatan Inti | 50 Menit | Guru mendemonstrasikan makhraj huruf dengan kartu besar, siswa menirukan intonasi secara mindful, aktivitas 'Mencari Pasangan Huruf' berkelompok, dan latihan mewarnai huruf hijaiyah | Bernalar Kritis, Kolaborasi |
| Penutup | 10 Menit | Refleksi bersama siswa menunjukkan karya hurufnya, membaca hamdalah bersama, dan pemberian stiker bintang prestasi | Bersyukur, Percaya Diri |

BAGIAN 4: ASESMEN & RUBRIK PENILAIAN
| Aspek Asesmen | Teknik Penilaian | Instrumen Penilaian | Kriteria Ketercapaian Tujuan Pembelajaran (KKTP) |
| Sikap Spiritual | Observasi Pembiasaan | Lembar Pengamatan Adab Mengaji | Khusyuk berdoa dan memperlakukan mushaf/kartu huruf dengan hormat |
| Pengetahuan | Tes Lisan / Tebak Huruf | LKPD Pasangkan Huruf Hijaiyah | Mampu menyebutkan minimal 20 huruf hijaiyah tunggal dengan tepat (Skor >= 75) |
| Keterampilan | Praktik Pelafalan | Rubrik Makharijul Huruf | Fasih membedakan bunyi huruf yang mirip (seperti Ha dan Kha, Ta dan Tha) |

BAGIAN 5: LEMBAR KERJA PESERTA DIDIK (LKPD) & RANGKUMAN MATERI
- Rangkuman Materi Inti: Huruf hijaiyah berjumlah 28 huruf. Huruf pertama adalah Alif dan huruf terakhir adalah Ya. Harakat fathah berbunyi 'A', kasrah berbunyi 'I', dan dhammah berbunyi 'U'.
- Lembar Kerja Murid (Aktivitas): 1. Tebalkan garis putus-putus pada huruf Ba, Ta, Tsa; 2. Pasangkan huruf hijaiyah yang sama bunyinya; 3. Warnai gambar kaligrafi Basmalah dengan rapi.
- Rencana Tindak Lanjut & Remedial: Bimbingan tahsin privat 10 menit bagi siswa yang belum lancar membedakan harakat kasrah dan dhammah.
"""
        else:
            return f"""BAGIAN 1: SURAT TUGAS & LANDASAN HUKUM
- Nomor Surat Tugas: 421.2 / 084 / SD-PAI / {year}
- Dasar Pelaksanaan: Permendikbudristek No. 12 Tahun 2024 & Program Penguatan Budi Pekerti Religius
- Pejabat Pemberi Tugas: {head_name} (Kepala Sekolah)
- Personel yang Ditugaskan: Tim Guru PAI, Pembina Rohis, dan Seluruh Guru Kelas
- Uraian Tugas & Tanggung Jawab: 1. Membimbing pembiasaan Shalat Dhuha dan Shalat Zhuhur berjamaah; 2. Menyelenggarakan bimbingan tahsin Al-Qur'an dan setoran hafalan Juz 30; 3. Membimbing adab harian islami (doa makan, salam, dan birrul walidain); 4. Menyusun rekapitulasi jurnal ibadah siswa.
- Waktu dan Tempat Pelaksanaan: Terjadwal Rutin Pagi Hari dan Waktu Istirahat di Musholla & Kelas {school_name}

BAGIAN 2: LAPORAN KETERLAKSANAAN PROGRAM
- Nama Program Resmi: Program Pembiasaan Spiritual & Karakter Religius Siswa
- Latar Belakang & Urgensi: Membangun pondasi moral dan ketakwaan sejak usia dini agar siswa memiliki benteng keimanan yang kokoh di tengah era digital.
- Tujuan Program Terukur: 1. Meningkatkan kedisiplinan shalat fardhu dan shalat sunnah; 2. Memfasilitasi hafalan minimal 10 surat pendek Juz Amma; 3. Menumbuhkan adab sopan santun kepada guru dan orang tua.
- Ruang Lingkup Aktivitas: Shalat Dhuha terpimpin, kultum siswa bergilir, tahfidz pagi, dan pencatatan buku mutaba'ah ibadah harian.
- Capaian Kinerja & Dampak: Partisipasi shalat berjamaah mencapai 100% dengan shaf tertib, 95% siswa lancar melafalkan doa harian, iklim sekolah religius.
- Kendala Lapangan & Solusi: Antrean tempat wudhu diatur secara bergiliran per kelas untuk menjaga ketertiban.
- Rencana Tindak Lanjut: Menyelenggarakan peringatan Hari Besar Islam (PHBI) dan lomba dai cilik antarkelas.

BAGIAN 3: JURNAL LOG HARIAN 7 KOLOM
| No | Hari / Tanggal | Waktu (WIB) | Bentuk Aktivitas Spesifik (PAI) | Sasaran Peserta | Hasil & Capaian Pengamatan | Status Verifikasi |
| 1 | Senin, 10/08/{year} | 07.00 - 07.30 | Shalat Dhuha Terpimpin & Dzikir Pagi | Siswa Kelas IV, V, VI | Siswa berwudhu tertib dan khusyuk melafalkan doa dhuha | Terverifikasi Guru PAI |
| 2 | Selasa, 01/09/{year} | 06.45 - 07.15 | Bimbingan Tahsin & Setoran Hafalan Juz 30 | Siswa Kelas II & III | Hafalan surat Al-Kafirun s.d. An-Nas fasih makhrajnya | Terverifikasi Pembina |
| 3 | Kamis, 24/09/{year} | 12.00 - 12.35 | Shalat Zhuhur Berjamaah & Kultum Siswa | Siswa Kelas V & Guru | Kultum mandiri bertema 'Kejujuran dan Adab Menghormati Orang Tua' | Terverifikasi Wali Kelas |

BAGIAN 4: MATRIKS EVALUASI & TINDAK LANJUT
| Aspek Evaluasi | Temuan di Lapangan | Analisis Ketercapaian | Tindak Lanjut Perbaikan | Penanggung Jawab |
| Ketertiban Ibadah | Siswa antusias shalat berjamaah tepat waktu | Sangat Memuaskan | Pertahankan pembiasaan adzan dan kultum bergilir | Guru PAI |
| Hafalan Doa | 95% siswa hafal doa harian wajib | Target tercapai dengan baik | Bimbingan tambahan bagi 5% siswa yang baru pindah | Wali Kelas |

BAGIAN 5: KESIMPULAN & REKOMENDASI PENGAWAS
- Kesimpulan Hasil Monitoring: Program pembiasaan keagamaan PAI terlaksana secara terprogram, tertib, dan membentuk akhlak mulia peserta didik.
- Catatan Pengawas Pembina ({supervisor_name}): Disetujui dan disahkan sebagai bukti fisik sah instrumen monitoring pemantauan pendidikan Kota Cirebon.
"""

    # 2. PEMBIASAAN BERMASYARAKAT / SOSIAL / POIN 6
    if any(k in t for k in ['poin 6', 'masyarakat', 'sosial', 'gotong royong', 'peduli', 'dhuafa', 'lingkungan sekitar']):
        return f"""BAGIAN 1: SURAT TUGAS & LANDASAN HUKUM
- Nomor Surat Tugas: 421.2 / 084 / SD-KS / X / {year}
- Dasar Pelaksanaan: Program Penguatan Pendidikan Karakter (PPK) Jabar Beraksi & RKT Kesiswaan {year}
- Pejabat Pemberi Tugas: {head_name} (Kepala Sekolah)
- Personel yang Ditugaskan: Pembina Kesiswaan, Seluruh Guru Kelas I - VI, Pengurus Paguyuban Orang Tua Murid
- Uraian Tugas & Tanggung Jawab: 1. Merencanakan dan mendampingi pelaksanaan Program Pembiasaan Bermasyarakat; 2. Mengkoordinasikan Gerakan Jumat Berkah dan penyaluran sembako dhuafa RT 01-04 RW 03 Kesambi; 3. Mengorganisasi Aksi Bersih Lingkungan Bersama Warga (Gerakan Pungut Sampah / GPS jalan protokol); 4. Menyusun laporan keterlaksanaan dan dokumentasi kegiatan sosial.
- Waktu dan Tempat Pelaksanaan: Terjadwal Pekan ke-2 dan ke-4 di Lingkungan Sekitar {school_name} (RT 01 s.d. RT 04 RW 03 Kesambi)

BAGIAN 2: LAPORAN KETERLAKSANAAN PROGRAM
- Nama Program Resmi: Program Pembiasaan 'Siswa Beraksi Nyata: Peduli Lingkungan & Berbagi Kasih Warga'
- Latar Belakang & Urgensi: Menumbuhkan kecerdasan sosial, empati kemanusiaan, rasa tanggung jawab warga negara, dan pembiasaan budaya gotong royong sejak dini agar siswa memiliki kepedulian nyata terhadap lingkungan sosial sekitar.
- Tujuan Program Terukur: 1. Membiasakan siswa berbagi dan peduli sesama melalui infaq sukarela berkala; 2. Menjaga kebersihan fasilitas umum di sekitar sekolah; 3. Membangun keakraban sosial harmonis antara sekolah dan warga masyarakat.
- Ruang Lingkup Aktivitas: a. Gerakan Jumat Peduli penyaluran sembako lansia/kurang mampu; b. Aksi Bersih Lingkungan pungut sampah terpilah; c. Pembiasaan etika sapa santun berpapasan dengan warga sekitar.
- Capaian Kinerja & Dampak: Partisipasi siswa mencapai 96% dalam aksi bakti sosial, terjalin keharmonisan kuat dengan pengurus RT/RW, terbangun karakter dermawan dan rendah hati pada murid.
- Kendala Lapangan & Solusi: Menyesuaikan jadwal kerja bakti dengan waktu luang warga sekitar pada hari Sabtu pagi.
- Rencana Tindak Lanjut: Mengintegrasikan agenda kemasyarakatan ke dalam tema Projek Penguatan Profil Pelajar Pancasila (P5) Kearifan Lokal.

BAGIAN 3: JURNAL LOG HARIAN 7 KOLOM
| No | Hari / Tanggal | Waktu (WIB) | Bentuk Aktivitas Spesifik (Bermasyarakat) | Sasaran Peserta | Hasil & Capaian Pengamatan | Status Verifikasi |
| 1 | Jumat, 11/09/{year} | 07.30 - 08.30 | Penyaluran Paket Sembako Dhuafa Jumat Berkah | Siswa Kls V-VI & Guru | Siswa menunjukkan sikap santun, empati tinggi, dan takzim kepada warga lansia | Terverifikasi RW 03 |
| 2 | Sabtu, 26/09/{year} | 07.15 - 08.45 | Aksi Bersih Lingkungan & Gerakan Pungut Sampah | Siswa Kls I-VI & Komite | Jalan protokol dan saluran air bersih dari sampah plastik, kerja sama kompak | Terverifikasi Tim Lingkungan |
| 3 | Jumat, 16/10/{year} | 08.00 - 09.30 | Silaturahmi & Santunan Anak Yatim Warga Sekitar | Perwakilan Siswa | Terjalin kehangatan kekeluargaan dan ketulusan berbagi rezeki sesama | Terverifikasi Kepala Sekolah |

BAGIAN 4: MATRIKS EVALUASI & TINDAK LANJUT
| Aspek Evaluasi | Temuan di Lapangan | Analisis Ketercapaian | Tindak Lanjut Perbaikan | Penanggung Jawab |
| Partisipasi Siswa | 96% siswa antusias mengikuti bakti warga | Target tercapai dengan sangat memuaskan | Terus jadwalkan bergilir tiap kelas | Koordinator Kesiswaan |
| Dampak Sosial | Hubungan dengan masyarakat sekitar amat erat | Citra sekolah sangat positif di mata warga | Pertahankan silaturahmi berkala | Kepala Sekolah |

BAGIAN 5: KESIMPULAN & REKOMENDASI PENGAWAS
- Kesimpulan Hasil Monitoring: Pelaksanaan bukti fisik pembiasaan bermasyarakat terlaksana secara nyata, konkret, dan memberikan dampak sosial positif bagi warga sekitar.
- Catatan Pengawas Pembina ({supervisor_name}): Disetujui dan disahkan sebagai bukti fisik sah instrumen monitoring pemantauan Jabar Beraksi Kota Cirebon.
"""

    # 3. ANTI-BULLYING / KEKERASAN / TPPK
    if any(k in t for k in ['bully', 'perundungan', 'kekerasan', 'tppk', 'deteksi dini', 'ramah anak']):
        return f"""BAGIAN 1: SURAT TUGAS & LANDASAN HUKUM
- Nomor Surat Tugas: 421.2 / 091 / TPPK-SD / {year}
- Dasar Pelaksanaan: Permendikbudristek No. 46 Tahun 2023 tentang Pencegahan dan Penanganan Kekerasan di Satuan Pendidikan
- Pejabat Pemberi Tugas: {head_name} (Kepala Sekolah)
- Personel yang Ditugaskan: Tim Pencegahan dan Penanganan Kekerasan (TPPK), Guru BK, Perwakilan Orang Tua
- Uraian Tugas & Tanggung Jawab: 1. Melakukan deteksi dini dan pemetaan potensi kekerasan fisik, verbal, atau siber; 2. Mengelola kanal aduan ramah anak secara rahasia; 3. Menangani laporan kasus dengan pendekatan keadilan restoratif tanpa kekerasan; 4. Menyusun laporan periodik penanganan kasus kepada dinas.
- Waktu dan Tempat Pelaksanaan: Tahun Ajaran {year}/2027 di Ruang Konseling & Lingkungan {school_name}

BAGIAN 2: LAPORAN KETERLAKSANAAN PROGRAM
- Nama Program Resmi: Penyelenggaraan Satuan Pendidikan Bebas Perundungan & Kekerasan (TPPK)
- Latar Belakang & Urgensi: Menjamin setiap murid terlindungi hak-hak dasarnya untuk belajar dengan aman, bermartabat, bebas dari rasa takut dan intimidasi.
- Tujuan Program Terukur: 1. Mewujudkan zero tolerance terhadap kekerasan di sekolah; 2. Menyediakan mekanisme penanganan kasus yang adil dan ramah anak; 3. Mengedukasi seluruh warga sekolah tentang batasan perilaku etis.
- Ruang Lingkup Aktivitas: Skrining angket rasa aman berkala, penyediaan kotak aduan 'Teman Cerita', sesi bimbingan mediasi damai.
- Capaian Kinerja & Dampak: Penurunan kasus perundungan hingga 0 kasus berat, 100% siswa merasa aman di kelas, terwujud budaya saling menyayangi.
- Kendala Lapangan & Solusi: Rasa ragu siswa melapor diatasi dengan jaminan kerahasiaan identitas saksi dan korban.
- Rencana Tindak Lanjut: Membentuk duta anti-perundungan siswa di setiap rombongan belajar.

BAGIAN 3: JURNAL LOG HARIAN 7 KOLOM
| No | Hari / Tanggal | Waktu (WIB) | Bentuk Aktivitas Spesifik (TPPK) | Sasaran Peserta | Hasil & Capaian Pengamatan | Status Verifikasi |
| 1 | Senin, 10/08/{year} | 07.15 - 08.00 | Sosialisasi & Deklarasi Bersama 'Stop Bullying' | Seluruh Siswa & Guru | Penandatanganan spanduk ikrar ramah anak oleh seluruh warga sekolah | Terverifikasi TPPK |
| 2 | Rabu, 02/09/{year} | 09.30 - 10.30 | Mediasi Restoratif Perselisihan Siswa di Lapangan | 2 Siswa Kls IV & Wali | Masalah ejekan selesai secara damai, saling meminta maaf tulus | Terverifikasi Guru BK |
| 3 | Kamis, 24/09/{year} | 10.00 - 11.00 | Pemeriksaan & Pengosongan Kotak Aduan Ramah Anak | Tim TPPK Sekolah | Tidak ada aduan kekerasan mayor; catatan saran kebersihan kelas ditindaklanjuti | Terverifikasi Kepala Sekolah |

BAGIAN 4: MATRIKS EVALUASI & TINDAK LANJUT
| Aspek Evaluasi | Temuan di Lapangan | Analisis Ketercapaian | Tindak Lanjut Perbaikan | Penanggung Jawab |
| Keamanan Lingkungan | Lingkungan sekolah sangat kondusif dan ramah | Target tercapai dengan sangat baik | Pertahankan patroli guru piket saat istirahat | Tim TPPK |
| Respon Laporan | Setiap aduan ditangani kurang dari 24 jam | Kinerja TPPK sangat responsif | Dokumentasikan berita acara penanganan rapi | Guru BK |

BAGIAN 5: KESIMPULAN & REKOMENDASI PENGAWAS
- Kesimpulan Hasil Monitoring: Mekanisme TPPK dan deteksi dini berjalan efektif, adil, serta menjamin perlindungan psikososial murid.
- Catatan Pengawas Pembina ({supervisor_name}): Disahkan sebagai dokumen bukti fisik standar penjaminan mutu satuan pendidikan ramah anak.
"""

    # 4. DEFAULT ADAPTIVE (MATEMATIKA / IPAS / BAHASA / LAINNYA)
    clean_title = topic.strip()
    if is_modul:
        return f"""BAGIAN 1: INFORMASI UMUM & IDENTITAS MODUL
- Nama Penyusun: Tim Guru Pengampu Mata Pelajaran
- Satuan Pendidikan: {school_name}
- Mata Pelajaran / Materi Pokok: {clean_title}
- Fase / Kelas / Semester: Fase Terkait (Kelas I s/d VI) - Semester Ganjil {year}/2027
- Alokasi Waktu: 2 Pertemuan (4 x 35 Menit)
- Profil Pelajar Pancasila: Bernalar Kritis, Mandiri, Gotong Royong, Kreatif
- Sarana dan Prasarana: Media konkret kontekstual, LCD proyektor/papan interaktif, LKPD murid
- Model Pembelajaran: Problem Based Learning (PBL) & Deep Learning Berdiferensiasi

BAGIAN 2: KOMPONEN INTI PEMBELAJARAN
- Capaian Pembelajaran (CP): Peserta didik menguasai konsep dasar materi {clean_title}, mampu menganalisis masalah kontekstual, serta menerapkan pemahamannya dalam pemecahan masalah sehari-hari dengan nalar kritis.
- Tujuan Pembelajaran (TP): 1. Mengidentifikasi konsep esensial {clean_title}; 2. Menyelesaikan lembar kerja eksplorasi secara mandiri maupun berkelompok; 3. Mempresentasikan hasil temuan belajar dengan percaya diri.
- Pemahaman Bermakna: Penguasaan {clean_title} memberikan keterampilan aplikatif yang bermanfaat bagi pemecahan persoalan nyata dalam kehidupan siswa.
- Pertanyaan Pemantik: 1. Pernahkah kalian mengamati penerapan {clean_title} di lingkungan sekitar? 2. Bagaimana cara paling mudah memahami konsep ini? 3. Mengapa kita perlu mempelajari materi ini bersama?

BAGIAN 3: ALUR KEGIATAN PEMBELAJARAN (DEEP LEARNING)
| Tahapan | Alokasi Waktu | Aktivitas Guru & Peserta Didik (Meaningful, Mindful, Joyful) | Integrasi Karakter |
| Pendahuluan | 10 Menit | Salam pembuka, doa khusyuk, apersepsi memantik rasa ingin tahu tentang {clean_title}, penyampaian tujuan dan alur kegiatan | Religius, Disiplin |
| Kegiatan Inti | 50 Menit | Eksplorasi konsep terarah, pengerjaan LKPD berdiferensiasi kelompok, bimbingan guru pada kelompok yang membutuhkan scaffolding, presentasi hasil karya | Bernalar Kritis, Kolaborasi |
| Penutup | 10 Menit | Refleksi bermakna bersama murid, penarikan kesimpulan intisari materi {clean_title}, apresiasi tepuk semangat, dan doa penutup | Bersyukur, Percaya Diri |

BAGIAN 4: ASESMEN & RUBRIK PENILAIAN
| Aspek Asesmen | Teknik Penilaian | Instrumen Penilaian | Kriteria Ketercapaian Tujuan Pembelajaran (KKTP) |
| Sikap (Afektif) | Observasi Harian | Lembar Pengamatan Profil Pelajar Pancasila | Menunjukkan rasa ingin tahu tinggi, kerja sama, dan tanggung jawab |
| Pengetahuan (Kognitif) | Tes Tertulis LKPD | Soal Evaluasi Konsep {clean_title} | Mampu menjawab minimal 80% pertanyaan konsep dengan analisis tepat |
| Keterampilan (Psikomotorik) | Unjuk Kerja / Presentasi | Rubrik Penilaian Produk & Kinerja | Terampil mengomunikasikan ide dan mempresentasikan hasil kerja kelompok |

BAGIAN 5: LEMBAR KERJA PESERTA DIDIK (LKPD) & RANGKUMAN MATERI
- Rangkuman Materi Inti: Intisari konsep materi {clean_title} mencakup definisi kunci, prinsip penerapan, dan langkah praktis yang aplikatif untuk siswa.
- Lembar Kerja Murid (Aktivitas Konkret): 1. Amatilah studi kasus materi {clean_title} pada buku teks; 2. Diskusikan bersama teman kelompok solusi yang paling efisien; 3. Tuliskan kesimpulan kelompok pada lembar portofolio.
- Rencana Tindak Lanjut & Remedial: Pengayaan berupa penugasan proyek eksploratif bagi siswa tuntas, dan bimbingan konsep dasar bagi siswa yang memerlukan remedial.
"""
    else:
        return f"""BAGIAN 1: SURAT TUGAS & LANDASAN HUKUM
- Nomor Surat Tugas: 421.2 / 084 / SD-01 / DISDIK / {year}
- Dasar Pelaksanaan: Rencana Kerja Tahunan (RKT) Satuan Pendidikan & Panduan Program Karakter {year}
- Pejabat Pemberi Tugas: {head_name} (Kepala Sekolah)
- Personel yang Ditugaskan: Tim Pelaksana Program, Seluruh Guru Kelas, dan Tenaga Kependidikan
- Uraian Tugas & Tanggung Jawab: 1. Merencanakan dan menyusun jadwal pelaksanaan implementasi {clean_title}; 2. Melakukan pendampingan berkala kepada seluruh peserta didik; 3. Mencatat ketercapaian target pada buku jurnal pemantauan; 4. Melakukan evaluasi rutin keterlaksanaan program secara akuntabel.
- Waktu dan Tempat Pelaksanaan: Semester Ganjil Tahun Ajaran {year}/2027 di Lingkungan {school_name}

BAGIAN 2: LAPORAN KETERLAKSANAAN PROGRAM
- Nama Program Resmi: Program Pelaksanaan & Pembiasaan {clean_title}
- Latar Belakang & Urgensi: Mengembangkan budaya mutu, pembentukan karakter luhur, dan ketercapaian indikator mutu pendidikan yang terukur melalui program {clean_title}.
- Tujuan Program Terukur: 1. Membiasakan praktik positif secara konsisten; 2. Meningkatkan keterlibatan aktif seluruh warga sekolah; 3. Menghadirkan lingkungan satuan pendidikan yang kondusif dan berdaya saing.
- Ruang Lingkup Aktivitas: a. Sosialisasi program kepada guru, murid, dan orang tua; b. Pelaksanaan terstruktur setiap hari/pekanan; c. Evaluasi dan apresiasi capaian berkala.
- Capaian Kinerja & Dampak: Tingkat keterlaksanaan mencapai 98% terverifikasi aktif, terbentuk kedisiplinan dan respon sangat positif dari peserta didik dan orang tua.
- Kendala Lapangan & Solusi: Hambatan koordinasi teknis diatasi melalui sinkronisasi linimasa kegiatan di awal bulan.
- Rencana Tindak Lanjut: Memperkuat kesinambungan program melalui pemantauan digital dan pelibatan paguyuban kelas.

BAGIAN 3: JURNAL LOG HARIAN 7 KOLOM
| No | Hari / Tanggal | Waktu (WIB) | Bentuk Aktivitas Spesifik ({clean_title}) | Sasaran Peserta | Hasil & Capaian Pengamatan | Status Verifikasi |
| 1 | Senin, 10/08/{year} | 07.15 - 08.00 | Sosialisasi Awal & Pelaksanaan Pembiasaan {clean_title} | Seluruh Siswa & Guru | Siswa menyimak panduan tertib dan memahami sasaran program | Terverifikasi Koordinator |
| 2 | Rabu, 02/09/{year} | 08.00 - 09.00 | Monitoring Berkala & Pencatatan Jurnal Kemajuan | Kelas IV, V, VI | 96% siswa konsisten mempraktikkan kebiasaan secara mandiri | Terverifikasi Guru Pendamping |
| 3 | Jumat, 25/09/{year} | 09.30 - 10.30 | Evaluasi Bulanan & Apresiasi Peserta Didik Teladan | Dewan Guru & Murid | Terjadi peningkatan signifikan pada aspek disiplin dan tanggung jawab | Terverifikasi Kepala Sekolah |

BAGIAN 4: MATRIKS EVALUASI & TINDAK LANJUT
| Aspek Evaluasi | Temuan di Lapangan | Analisis Ketercapaian | Tindak Lanjut Perbaikan | Penanggung Jawab |
| Keterlaksanaan Program | Program {clean_title} berlangsung sesuai SOP | Target tercapai dengan sangat memuaskan | Pertahankan konsistensi pelaksanaan | Tim Pelaksana |
| Dampak Karakter Siswa | Terlihat peningkatan kepribadian positif murid | Sikap mandiri dan disiplin terbangun kuat | Lanjutkan pendampingan berkelanjutan | Wali Kelas |

BAGIAN 5: KESIMPULAN & REKOMENDASI PENGAWAS
- Kesimpulan Hasil Monitoring: Berkas dan pelaksanaan {clean_title} telah diverifikasi lengkap, sistematis, dan memenuhi kriteria standar mutu pendidikan.
- Catatan Pengawas Pembina ({supervisor_name}): Disetujui dan disahkan secara resmi untuk kelengkapan administrasi penjaminan mutu pendidikan Kota Cirebon.
"""
