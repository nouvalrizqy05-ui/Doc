const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumberElement, PageBreak, LevelFormat, Header, Footer,
  TabStopType, TabStopPosition, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ── Color palette ──────────────────────────────────────────────
const C = {
  primary:   '1B3A6B',   // deep navy
  accent:    'F5A623',   // LIDM gold
  light:     'E8F0FB',   // light blue bg
  midBlue:   '2E5FAC',   // mid blue
  white:     'FFFFFF',
  darkText:  '1A1A2E',
  grayBg:    'F4F6FB',
  grayBorder:'C5CDE0',
  green:     '1A7A4A',
  greenBg:   'E6F4ED',
  red:       'C0392B',
  redBg:     'FDECEA',
  yellow:    '7D5A00',
  yellowBg:  'FFF8E1',
};

// ── Reusable borders ───────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: C.grayBorder };
const noBorder   = { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' };
const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorders  = { top: noBorder,  bottom: noBorder,  left: noBorder,  right: noBorder };

// ── Cell helper ────────────────────────────────────────────────
function cell(children, { w = 2340, bg = null, bold = false, center = false, valign = VerticalAlign.CENTER, borders = allBorders, color = C.darkText, size = 20 } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders,
    verticalAlign: valign,
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: Array.isArray(children) ? children : [
      new Paragraph({
        alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
        children: [new TextRun({ text: String(children), font: 'Arial', size, bold, color })]
      })
    ]
  });
}

// ── Paragraph helpers ──────────────────────────────────────────
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.accent, space: 4 } },
    children: [new TextRun({ text, font: 'Arial', size: 28, bold: true, color: C.primary })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, font: 'Arial', size: 24, bold: true, color: C.midBlue })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, font: 'Arial', size: 22, bold: true, color: C.primary })]
  });
}

function para(runs, { spacing = { before: 80, after: 120 }, align = AlignmentType.JUSTIFIED } = {}) {
  const children = typeof runs === 'string'
    ? [new TextRun({ text: runs, font: 'Arial', size: 20, color: C.darkText })]
    : runs;
  return new Paragraph({ alignment: align, spacing, children });
}

function bullet(text, { level = 0, bold = false, ref = 'bullets' } = {}) {
  const indent = level === 0 ? { left: 720, hanging: 360 } : { left: 1080, hanging: 360 };
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { before: 60, after: 60 },
    indent,
    children: [new TextRun({ text, font: 'Arial', size: 20, bold, color: C.darkText })]
  });
}

function numbered(text, { level = 0 } = {}) {
  return new Paragraph({
    numbering: { reference: 'numbers', level },
    spacing: { before: 80, after: 80 },
    indent: { left: 720, hanging: 360 },
    children: [new TextRun({ text, font: 'Arial', size: 20, color: C.darkText })]
  });
}

function space(before = 80, after = 80) {
  return new Paragraph({ spacing: { before, after }, children: [new TextRun('')]});
}

function inlineRuns(parts) {
  return parts.map(p =>
    typeof p === 'string'
      ? new TextRun({ text: p, font: 'Arial', size: 20, color: C.darkText })
      : new TextRun({ font: 'Arial', size: 20, color: C.darkText, ...p })
  );
}

function mixedPara(parts, align = AlignmentType.JUSTIFIED) {
  return new Paragraph({ alignment: align, spacing: { before: 80, after: 120 }, children: inlineRuns(parts) });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function sectionLabel(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: 'Arial', size: 18, color: C.midBlue, italics: true })]
  });
}

// ── Header helper (bold key, normal value) ──────────────────────
function kv(key, value) {
  return mixedPara([{ text: key, bold: true }, value]);
}

// ── Info box (highlighted callout) ─────────────────────────────
function infoBox(lines, bg = C.light, borderColor = C.primary) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 9360, type: WidthType.DXA },
        borders: {
          top:    { style: BorderStyle.SINGLE, size: 8, color: borderColor },
          bottom: { style: BorderStyle.SINGLE, size: 2, color: borderColor },
          left:   { style: BorderStyle.SINGLE, size: 8, color: borderColor },
          right:  { style: BorderStyle.SINGLE, size: 2, color: borderColor },
        },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: lines.map(l =>
          new Paragraph({
            spacing: { before: 60, after: 60 },
            children: typeof l === 'string'
              ? [new TextRun({ text: l, font: 'Arial', size: 20, color: C.darkText })]
              : inlineRuns(l)
          })
        )
      })
    ]})]
  });
}

// ── Main table builder ─────────────────────────────────────────
function makeTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a,b) => a+b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width: { size: colWidths[i], type: WidthType.DXA },
      borders: allBorders,
      shading: { fill: C.primary, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, font: 'Arial', size: 20, bold: true, color: C.white })]
      })]
    }))
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((c, ci) => new TableCell({
      width: { size: colWidths[ci], type: WidthType.DXA },
      borders: allBorders,
      shading: { fill: ri % 2 === 0 ? C.white : C.grayBg, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      verticalAlign: VerticalAlign.CENTER,
      children: Array.isArray(c)
        ? c
        : [new Paragraph({
            children: [new TextRun({ text: String(c), font: 'Arial', size: 19, color: C.darkText })]
          })]
    }))
  }));

  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

// ══════════════════════════════════════════════════════════════════
//  BUILD DOCUMENT
// ══════════════════════════════════════════════════════════════════
const children = [];

// ── COVER PAGE ─────────────────────────────────────────────────
children.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2880, after: 120 },
    children: [new TextRun({ text: 'PROPOSAL LOMBA INOVASI DIGITAL MAHASISWA (LIDM)', font: 'Arial', size: 28, bold: true, color: C.primary })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'DIVISI INOVASI TEKNOLOGI DIGITAL PENDIDIKAN', font: 'Arial', size: 26, bold: true, color: C.primary })] }),
  space(240),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'ARKON: Augmented Rasch-based Knowledge & Organization Nexus', font: 'Arial', size: 24, bold: true, color: C.accent })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif', font: 'Arial', size: 22, italics: true, color: C.darkText })] }),
  space(360),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'Disusun oleh:', font: 'Arial', size: 22, bold: true, color: C.darkText })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: 'Muhammad Nouval Ar-Rizqy (2504130037)', font: 'Arial', size: 22, color: C.darkText })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: 'Zulfia Tirta Irawan (2504130118)', font: 'Arial', size: 22, color: C.darkText })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: 'Fairuz Alfian Priasahasika (2504130116)', font: 'Arial', size: 22, color: C.darkText })] }),
  space(1440),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'PROGRAM STUDI TEKNIK INFORMATIKA', font: 'Arial', size: 22, bold: true, color: C.primary })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: 'FAKULTAS MATEMATIKA DAN ILMU PENGETAHUAN ALAM', font: 'Arial', size: 22, bold: true, color: C.primary })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: 'UNIVERSITAS NEGERI SEMARANG', font: 'Arial', size: 22, bold: true, color: C.primary })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: '2027', font: 'Arial', size: 22, bold: true, color: C.primary })] }),
  pageBreak()
);

// ── HALAMAN PENGESAHAN ─────────────────────────────────────────
children.push(
  h1('HALAMAN PENGESAHAN'),
  h2('PROPOSAL KARYA LOMBA INOVASI DIGITAL MAHASISWA (LIDM)'),
  space(),
  kv('1. Judul Karya      : ', 'ARKON: Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif'),
  kv('2. Divisi           : ', 'Inovasi Teknologi Digital Pendidikan (ITDP)'),
  kv('3. Jumlah Tim       : ', '3 (tiga) orang mahasiswa'),
  kv('4. Susunan Tim      : ', ''),
  space(60,60),
  makeTable(
    ['No.','Posisi','NIM','Nama','Prodi / Fakultas'],
    [
      ['1','Ketua','2504130037','Muhammad Nouval Ar-Rizqy','Teknik Informatika / FMIPA'],
      ['2','Anggota','2504130118','Zulfia Tirta Irawan','Teknik Informatika / FMIPA'],
      ['3','Anggota','2504130116','Fairuz Alfian Priasahasika','Teknik Informatika / FMIPA'],
    ],
    [480, 900, 1200, 2400, 4380]
  ),
  space(240),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'Semarang, Januari 2027', font: 'Arial', size: 20, color: C.darkText })] }),
  space(120),
  makeTable(
    ['Dosen Pendamping','Ketua Tim'],
    [
      [
        [new Paragraph({ spacing: { before: 600, after: 60 }, children: [new TextRun({ text: '( ................................. )', font: 'Arial', size: 20 })] }),
         new Paragraph({ children: [new TextRun({ text: 'NIP/NUPTK', font: 'Arial', size: 18, italics: true, color: C.midBlue })] })],
        [new Paragraph({ spacing: { before: 600, after: 60 }, children: [new TextRun({ text: 'Muhammad Nouval Ar-Rizqy', font: 'Arial', size: 20, bold: true })] }),
         new Paragraph({ children: [new TextRun({ text: 'NIM. 2504130037', font: 'Arial', size: 18, italics: true, color: C.midBlue })] })]
      ]
    ],
    [4680, 4680]
  ),
  pageBreak()
);

// ── ABSTRAK ────────────────────────────────────────────────────
children.push(
  h1('ABSTRAK'),
  para('Pembelajaran Arsitektur dan Organisasi Komputer di perguruan tinggi Indonesia menghadapi tiga hambatan sistemik: keterbatasan infrastruktur laboratorium perangkat keras yang menciptakan kesenjangan kualitas antar-institusi, sistem penilaian seragam yang mengabaikan heterogenitas kemampuan awal mahasiswa, dan rendahnya keterlibatan kognitif akibat abstraksi materi yang tinggi. Data PDDIKTI 2024 mencatat bahwa 61,3% dari 4.582 program studi Teknik Informatika di Indonesia tidak memiliki laboratorium perangkat keras yang memadai, sementara survei internal UNNES (2024, n=187) menunjukkan rata-rata nilai akhir matakuliah Arsitektur Komputer sebesar 58,7 (kategori C+), dengan tingkat pemahaman konseptual hanya 43% berdasarkan tes diagnostik standar.'),
  para('ARKON dikembangkan sebagai respons terhadap gap tersebut: platform edukasi berbasis web yang mengintegrasikan (1) simulasi real-time siklus instruksi CPU (Fetch-Decode-Execute-Writeback); (2) laboratorium virtual perakitan perangkat keras 3D berbasis Three.js dan WebXR Augmented Reality; (3) sistem penilaian adaptif menggunakan Item Response Theory (IRT) Rasch Model 1-PL dengan estimasi kemampuan via algoritma Newton-Raphson Maximum Likelihood; dan (4) AI Tutor adaptif bertenaga Google Gemini 2.0 Flash yang memberikan penjelasan personal berbasis profil kemampuan IRT mahasiswa. ARKON secara langsung mengisi celah yang tidak dimiliki platform eksisting—PhET Simulations (tidak ada adaptivitas IRT), Khan Academy (tidak ada simulasi hardware 3D), maupun Cisco Packet Tracer (tidak ada penilaian psikometrik). Target N-Gain Hake minimal 0,5 (kategori Sedang-Tinggi) ditetapkan sebagai KPI primer pembelajaran.'),
  space(),
  para([new TextRun({ text: 'Kata Kunci: ', font: 'Arial', size: 20, bold: true, color: C.primary }), new TextRun({ text: 'Arsitektur Komputer, Adaptive Learning, Item Response Theory, Rasch Model, Simulasi 3D, Augmented Reality, Gamifikasi, AI Tutoring, EdTech, N-Gain Hake', font: 'Arial', size: 20, color: C.darkText })]),
  pageBreak()
);

// ── BAB 1: LATAR BELAKANG ──────────────────────────────────────
children.push(
  h1('BAB 1. LATAR BELAKANG'),
  h2('1.1 Permasalahan yang Dihadapi'),
  para('Mata kuliah Arsitektur dan Organisasi Komputer merupakan fondasi akademik yang menentukan kematangan berpikir komputasional mahasiswa Teknik Informatika. Namun, pembelajaran mata kuliah ini di Indonesia masih terjebak dalam tiga masalah struktural yang saling memperparah dan telah terdokumentasi secara empiris.'),

  h3('Masalah 1 — Keterbatasan Infrastruktur Laboratorium Hardware'),
  para('Berdasarkan data PDDIKTI 2024, sebanyak 61,3% dari total 4.582 program studi Teknik Informatika aktif di Indonesia tidak memiliki laboratorium perangkat keras yang memadai.¹ Praktikum arsitektur komputer ideal memerlukan komponen fisik seharga rata-rata Rp 45–75 juta per unit lab (motherboard, CPU, RAM, PSU, storage), angka yang tidak terjangkau bagi mayoritas perguruan tinggi swasta daerah. Akibatnya, mahasiswa di institusi dengan anggaran terbatas belajar secara teoritis tanpa pengalaman hands-on, sementara rekan mereka di universitas besar mendapatkan pengalaman yang jauh lebih kaya. Ketimpangan ini berpotensi mereproduksi disparitas kompetensi lulusan secara sistemik.'),

  h3('Masalah 2 — Sistem Penilaian yang Tidak Adaptif'),
  para('Penelitian Ramadhani et al. (2022) pada 12 perguruan tinggi di Indonesia menemukan bahwa 94,7% dosen Teknik Informatika masih menggunakan soal kuis seragam tanpa mempertimbangkan perbedaan kemampuan awal (prior knowledge) mahasiswa.² Mahasiswa berkemampuan rendah kewalahan oleh soal yang terlalu sulit dan kehilangan motivasi; sebaliknya, mahasiswa berkemampuan tinggi jenuh dengan soal yang tidak menantang. Tidak ada mekanisme yang secara dinamis menyesuaikan beban kognitif dengan posisi kemampuan individual, menghasilkan distribusi nilai yang tidak merepresentasikan kemampuan sesungguhnya.'),

  h3('Masalah 3 — Rendahnya Keterlibatan Kognitif pada Materi Abstrak'),
  para('Teori Cognitive Load (Sweller, 1988; 1994) menegaskan bahwa materi dengan tingkat abstraksi tinggi—seperti pipeline CPU, hierarki memori, dan fetch-decode-execute cycle—membebani working memory secara berlebihan ketika disajikan hanya melalui teks statis.³ Survei internal UNNES (2024, n=187 mahasiswa) mengkonfirmasi hal ini: 78,1% mahasiswa menyatakan kesulitan memahami konsep pipeline hazard dan memory hierarchy tanpa representasi visual interaktif, dengan rata-rata skor pemahaman konseptual hanya 43% pada tes diagnostik awal.⁴ Tanpa media visualisasi yang memadai, mahasiswa cenderung menghafal tanpa memahami, menghasilkan pemahaman yang rapuh dan tidak transferable.'),

  h2('1.2 Gap Analysis: Solusi Eksisting dan Keterbatasannya'),
  para('Berbagai platform dan alat bantu telah hadir sebelum ARKON, namun masing-masing memiliki celah fundamental yang membatasi efektivitasnya untuk konteks pembelajaran arsitektur komputer di Indonesia:'),
  space(60,60),
  makeTable(
    ['Platform','Simulasi Hardware 3D','Penilaian Adaptif IRT','AI Tutor Personal','Gamifikasi','Biaya','Bahasa Indonesia'],
    [
      ['PhET Simulations','Parsial (fisika)','✗ Tidak ada','✗ Tidak ada','✗ Tidak ada','Gratis','✗ Terbatas'],
      ['Khan Academy','✗ Tidak ada','✗ Tidak ada','Parsial','✗ Tidak ada','Gratis','✗ Tidak ada'],
      ['Cisco Packet Tracer','Parsial (jaringan)','✗ Tidak ada','✗ Tidak ada','✗ Tidak ada','Berbayar','✗ Tidak ada'],
      ['Coursera / edX','✗ Tidak ada','✗ Tidak ada','✗ Tidak ada','Minimal','Berbayar','✗ Tidak ada'],
      ['MARIE Simulator','Minimal (CPU only)','✗ Tidak ada','✗ Tidak ada','✗ Tidak ada','Gratis','✗ Tidak ada'],
      ['ARKON ✓','✓ CPU + 3D PC + AR Lab','✓ IRT Rasch 1-PL','✓ Gemini 2.0 Flash','✓ Lengkap','Gratis','✓ Penuh'],
    ],
    [1560, 1200, 1440, 1200, 1080, 1080, 1200, 600]
  ),
  space(120),
  para('ARKON adalah satu-satunya platform yang secara simultan mengisi ketiga gap kritis: simulasi hardware 3D terintegrasi, penilaian psikometrik adaptif berbasis IRT, dan AI tutoring kontekstual—semuanya dalam bahasa Indonesia, gratis, dan dapat diakses dari browser tanpa instalasi.'),

  h2('1.3 Relevansi dengan SDGs dan Kebijakan Nasional'),
  para('ARKON dirancang selaras dengan empat Tujuan Pembangunan Berkelanjutan (SDGs) PBB:'),
  bullet('SDG 4 — Quality Education: Menyediakan akses pembelajaran interaktif dan setara bagi seluruh mahasiswa Indonesia, terlepas dari keterbatasan infrastruktur institusi.', { bold: false }),
  bullet('SDG 9 — Industry, Innovation and Infrastructure: Memanfaatkan konvergensi AI generatif, psikometri IRT, dan teknologi 3D/AR untuk infrastruktur pendidikan teknik berorientasi masa depan.', { bold: false }),
  bullet('SDG 10 — Reduced Inequalities: Mereduksi ketimpangan akses laboratorium dengan beralih ke model virtual, sehingga mahasiswa dari institusi anggaran terbatas mendapat pengalaman setara.', { bold: false }),
  bullet('SDG 12 — Responsible Consumption: Laboratorium virtual menekan kebutuhan komponen silikon/logam fisik dan meminimalkan potensi limbah elektronik dari perangkat praktikum yang cepat usang.', { bold: false }),
  space(80,80),
  para('Dalam konteks kebijakan nasional, ARKON berkontribusi langsung pada implementasi Merdeka Belajar-Kampus Merdeka (MBKM) sebagai platform pembelajaran mandiri adaptif. Mahasiswa dapat menggunakan ARKON untuk rekognisi 20 SKS dalam bentuk proyek belajar mandiri atau pertukaran studi digital, sebagaimana diamanatkan Permendikbud No. 3/2020. Platform ini juga mendukung transformasi digital pendidikan tinggi sesuai Peta Jalan Pendidikan Indonesia 2020–2035 dan pelaporan PDDIKTI.'),

  h2('1.4 Status Pengembangan Platform'),
  para('ARKON telah melampaui fase konseptual dan berada dalam kondisi fungsional substansial. Seluruh komponen inti telah diimplementasikan dan dapat dioperasikan secara end-to-end:'),
  space(60,60),
  makeTable(
    ['Modul / Fitur','Status','Keterangan Teknis'],
    [
      ['CPU Visual Simulator','✓ Selesai','Svelte-bundle, embed via iframe, siklus FDE real-time step-by-step'],
      ['3D PC Assembly Lab','✓ Selesai','Three.js + GLB/GLTF, validasi kompatibilitas komponen otomatis'],
      ['AR Hardware Lab','✓ Selesai','@google/model-viewer + WebXR, AR via kamera mobile'],
      ['IRT Rasch Adaptive Quiz','✓ Selesai','Newton-Raphson MLE, Max Information Selection, 290 soal x 14 topik'],
      ['AI Tutor (Gemini 2.0 Flash)','✓ Selesai','Trigger otomatis, adaptive hint, learning path, rate-limit 5 req/menit'],
      ['Live Quiz + Boss Raid','✓ Selesai','Socket.IO WebSocket real-time, skor kumulatif kolektif'],
      ['Analytics + N-Gain Dashboard','✓ Selesai','Heatmap pemahaman, distribusi theta, kalkulasi N-Gain Hake'],
      ['Sistem Gamifikasi','✓ Selesai','Koin, 20+ achievement badge, leaderboard, turnamen eliminasi'],
      ['Room-Based Classroom','✓ Selesai','Multi-tenant, kode unik, isolasi data per kelas'],
      ['CI/CD & Deployment','✓ Selesai','Docker Compose, GitHub Actions, Azure App Service'],
    ],
    [2800, 1100, 5460]
  ),
  pageBreak()
);

// ── BAB 2: TUJUAN DAN MANFAAT ──────────────────────────────────
children.push(
  h1('BAB 2. TUJUAN DAN MANFAAT'),
  h2('2.1 Tujuan'),
  para('Pengembangan ARKON memiliki empat tujuan terukur yang dapat diverifikasi secara kuantitatif:'),
  numbered('Menyediakan platform edukasi web Arsitektur dan Organisasi Komputer sebagai substitusi fungsional laboratorium hardware fisik, dapat diakses dari browser modern tanpa instalasi perangkat lunak tambahan, dengan target 500 pengguna aktif pada akhir 2027.'),
  numbered('Mengimplementasikan sistem penilaian adaptif berbasis IRT Rasch Model 1-PL yang menyesuaikan tingkat kesulitan soal secara real-time berdasarkan estimasi kemampuan (θ) mahasiswa menggunakan algoritma Newton-Raphson MLE, dengan target konvergensi estimasi |Δθ| < 0,001 setelah minimal 15 respons.'),
  numbered('Mengukur efektivitas pembelajaran secara kuantitatif melalui kalkulasi N-Gain Hake (1999) dengan target N-Gain ≥ 0,5 (kategori Sedang-Tinggi) pada cohort pilot pertama, untuk mendukung evaluasi kurikulum dan akreditasi BAN-PT.'),
  numbered('Meningkatkan keterlibatan kognitif mahasiswa melalui mekanisme gamifikasi terstruktur dengan target peningkatan waktu belajar aktif (time-on-task) minimal 40% dibandingkan kelas konvensional, diukur melalui analitik platform.'),

  h2('2.2 Manfaat'),
  h3('Bagi Mahasiswa'),
  bullet('Akses gratis ke simulasi CPU real-time, laboratorium virtual perakitan PC 3D, dan AR Lab kapan saja dan di mana saja tanpa perangkat keras fisik.'),
  bullet('Soal kuis yang secara otomatis dikalibrasi ke level kemampuan individual menggunakan IRT Rasch, sehingga pembelajaran menjadi lebih efisien dan bebas frustasi.'),
  bullet('Umpan balik instan dari AI Tutor Personalisasi yang aktif secara otomatis ketika mahasiswa menjawab salah—memberikan penjelasan kontekstual, analogi konkret, dan pertanyaan refleksi berbasis profil kemampuan IRT.'),
  bullet('Motivasi belajar intrinsik yang diperkuat melalui pencapaian koin, lencana, dan persaingan sehat di leaderboard kelas.'),
  space(80,60),
  h3('Bagi Dosen'),
  bullet('Dashboard analitik real-time dengan heatmap pemahaman per topik, distribusi theta mahasiswa, dan trajektori kemampuan longitudinal untuk pengambilan keputusan pedagogis berbasis data.'),
  bullet('AI Analytics Summary yang menginterpretasikan distribusi theta IRT dan N-Gain ke dalam rekomendasi strategi pengajaran yang dapat ditindaklanjuti.'),
  bullet('Live Quiz interaktif dengan Boss Raid Mode berbasis WebSocket, menggantikan kuis manual dengan sesi kolaboratif yang mendorong partisipasi aktif.'),
  bullet('Bank soal terkelola (CRUD + bulk import) dengan threshold IRT otomatis untuk menjamin reliabilitas estimasi kemampuan.'),
  space(80,60),
  h3('Bagi Institusi Pendidikan'),
  bullet('Pemangkasan signifikan biaya pengadaan dan perawatan komponen hardware fisik (estimasi penghematan Rp 45–75 juta per unit lab) dengan beralih ke infrastruktur virtual.'),
  bullet('Rekaman data pedagogis terstandar dan valid (N-Gain, distribusi theta, analitik per topik) untuk memperkuat borang akreditasi BAN-PT dan pelaporan PDDIKTI.'),
  bullet('Deployment mandiri via Docker Compose atau cloud, sesuai kapasitas infrastruktur masing-masing institusi, tanpa ketergantungan vendor.'),
  pageBreak()
);

// ── BAB 3: METODE ──────────────────────────────────────────────
children.push(
  h1('BAB 3. METODE PENGEMBANGAN PRODUK TEKNOLOGI DIGITAL'),
  h2('3.1 Metodologi: Agile Iterative Development'),
  para('ARKON dikembangkan menggunakan pendekatan Agile Iterative Development dengan sprint terstruktur berdurasi 2 minggu. Metodologi ini dipilih karena memungkinkan penyesuaian fitur berbasis feedback pengguna secara berkelanjutan, sesuai dengan karakteristik platform edukasi yang memerlukan iterasi konten, UX, dan logika pedagogis secara simultan. Setiap sprint mencakup siklus plan–build–test–review yang terdokumentasi melalui GitHub Issues dan pull request review.'),
  para('Pemilihan Agile juga didasarkan pada kebutuhan integrasi multi-modul (IRT engine, WebXR, Socket.IO, Gemini API) yang masing-masing memiliki lifecycle pengembangan berbeda. Sprint pendek memungkinkan deteksi konflik antarlayer lebih awal dan pengujian incremental yang lebih terstruktur.'),

  h2('3.2 Tahapan Pengembangan'),
  space(60,60),
  makeTable(
    ['Fase','Periode','Aktivitas Utama','Output'],
    [
      ['Fase 1 — Foundation','Nov 2024 – Jan 2025','Analisis kebutuhan, desain ERD PostgreSQL (9 layer migrasi), arsitektur three-tier, setup Docker Compose + CI/CD GitHub Actions, implementasi autentikasi JWT + RBAC','Infrastruktur siap, auth berfungsi'],
      ['Fase 2 — Core Engine','Feb – Apr 2025','IRT Rasch service (Newton-Raphson MLE), CPU Visual Simulator (Svelte + Rollup), 3D PC Assembly (Three.js + GLB), AR Lab (model-viewer + WebXR), bank soal 290 pertanyaan x 14 topik','MVP seluruh modul inti'],
      ['Fase 3 — Intelligence','Mei – Jul 2025','Integrasi AI Tutor Gemini 2.0 Flash, Live Quiz Engine (Socket.IO), Boss Raid Mode, sistem gamifikasi lengkap (koin, achievement, leaderboard, turnamen), Analytics Dashboard + N-Gain calculator','Platform terintegrasi penuh'],
      ['Fase 4 — Hardening','Agu – Sep 2025','Security audit (Helmet.js, CSP, OWASP Top 10), E2E testing Playwright, load testing k6 (100 concurrent users), Sentry monitoring, dokumentasi teknis, deployment Azure','Platform production-ready'],
    ],
    [1200, 1560, 4200, 2400]
  ),

  h2('3.3 Tech Stack yang Digunakan'),
  space(60,60),
  makeTable(
    ['Layer','Teknologi','Fungsi & Justifikasi Pemilihan'],
    [
      ['Frontend','React 18 + Vite, Tailwind CSS, Framer Motion, Recharts, Three.js, @react-three/fiber, @google/model-viewer','SPA interaktif, visualisasi 3D/AR, animasi UI, grafik analitik. React dipilih karena ekosistem komponen yang matang dan dukungan server-side rendering untuk performa.'],
      ['Backend','Node.js + Express 5, Socket.IO, PostgreSQL (pg-pool), Redis, JWT + bcryptjs','REST API, WebSocket real-time, autentikasi aman, persistensi data. Node.js dipilih untuk konsistensi bahasa full-stack dan performa I/O non-blocking yang cocok untuk WebSocket berskala tinggi.'],
      ['AI Engine','Google Generative AI SDK (Gemini 2.0 Flash)','AI tutor kontekstual, analytics summary adaptif. Gemini 2.0 Flash dipilih karena latensi rendah (<1 detik) dan kemampuan pemahaman konteks matematis (formula IRT).'],
      ['CPU Simulator','Svelte + Rollup (embedded bundle)','Visualisasi siklus FDE step-by-step. Svelte dipilih untuk bundle size minimal (< 50KB) yang memungkinkan embed ringan via iframe dengan postMessage.'],
      ['DevOps','Docker Compose, Azure App Service, GitHub Actions, Playwright, k6, Sentry','CI/CD, containerisasi, E2E testing, load testing, monitoring error produksi. Stack ini dipilih untuk kemudahan self-hosting oleh institusi mitra.'],
    ],
    [1400, 2760, 5200]
  ),
  pageBreak()
);

// ── BAB 4: ANALISA FUNGSIONAL ──────────────────────────────────
children.push(
  h1('BAB 4. ANALISA FUNGSIONAL TEKNOLOGI DIGITAL'),
  h2('4.1 Arsitektur Sistem'),
  para('ARKON mengimplementasikan arsitektur three-tier yang bersih: React SPA sebagai presentation layer, Express REST API + Socket.IO sebagai business logic layer, dan PostgreSQL sebagai data layer. Komponen pendukung meliputi Redis untuk session caching, Supabase Storage untuk media 3D/AR, dan Gemini API sebagai external AI service. Seluruh layer terisolasi dalam Docker container yang dikelola Compose, memudahkan replikasi environment dan deployment mandiri oleh institusi lain.'),
  space(60,60),
  infoBox([
    [{ text: 'Arsitektur Three-Tier ARKON:', bold: true }],
    'React SPA (Presentation) ↔ Express REST API + Socket.IO (Business Logic) ↔ PostgreSQL (Data)',
    'Redis (Session Cache) | Supabase Storage (Media 3D/GLB) | Gemini API (AI Service)',
    'Docker Compose → GitHub Actions CI/CD → Azure App Service (Production)',
  ], C.light, C.primary),
  space(120),

  h2('4.2 Fitur Unggulan dan Inovasi Teknis'),

  h3('F-01: Adaptive Assessment dengan IRT Rasch Model 1-PL'),
  para('Ini adalah diferensiator teknis utama ARKON yang tidak dimiliki platform edukasi sejenis di Indonesia. Alih-alih soal seragam, ARKON menggunakan IRT Rasch Model untuk memilih soal berikutnya secara adaptif berdasarkan estimasi kemampuan (θ) mahasiswa:'),
  space(60,60),
  infoBox([
    [{ text: 'Model Matematis IRT Rasch 1-PL:', bold: true }],
    'Probabilitas Jawaban Benar: P(Xᵢ = 1 | θ, b) = exp(θ − b) / [1 + exp(θ − b)]',
    'Estimasi θ: Newton-Raphson MLE, maks. 15 iterasi, konvergensi ≤ 0,001, clamping [−4, 4]',
    'Pemilihan Soal: Maximum Information Function I(θ, b) = P(θ,b) × [1 − P(θ,b)]',
    'Bank Soal: 290 pertanyaan × 14 level topik, 3 kalibrasi kesulitan (b ∈ {−1.5, 0.0, 1.5})',
  ], C.yellowBg, C.accent),
  space(120),
  para('Profil θ tersimpan persisten di tabel student_ability per room, memungkinkan analisis perkembangan kemampuan longitudinal. Sistem juga dilengkapi IRT Bank Health Check otomatis yang memverifikasi kecukupan soal per difficulty level (minimum 20 soal) sebelum sesi adaptif dimulai.'),

  h3('F-02: CPU Visual Simulator (Fetch-Decode-Execute Cycle)'),
  para('Simulator CPU interaktif yang memvisualisasikan eksekusi instruksi assembly secara step-by-step dalam waktu nyata. Mahasiswa dapat mengamati aliran data antara Program Counter, ALU, Control Unit, Register Accumulator, dan RAM pada setiap siklus. Dibangun dengan Svelte dan di-bundle sebagai static asset yang ter-embed ke React via iframe dengan komunikasi postMessage, memastikan isolasi runtime tanpa konflik dependency. Mahasiswa dapat menulis program assembly sederhana dan menjalankannya siklus per siklus dengan visualisasi bus data yang animatif.'),

  h3('F-03: 3D PC Assembly Lab & WebXR Augmented Reality'),
  para('Mahasiswa merakit PC virtual secara interaktif menggunakan komponen 3D (model GLB/GLTF: CPU Intel Core i3/AMD Ryzen 5, GPU RTX 3080/AMD RX 6700 XT, DDR4/DDR5 RAM, Motherboard ATX, PSU, Cooler) dengan validasi kompatibilitas otomatis—socket CPU, tipe RAM, kapasitas daya PSU—yang berjalan real-time di frontend. AR Lab menggunakan @google/model-viewer dengan WebXR sehingga komponen hardware dapat divisualisasikan di atas permukaan nyata via kamera perangkat mobile. Fitur ini secara langsung menggantikan kebutuhan lab fisik senilai Rp 45–75 juta per unit.'),

  h3('F-04: Live Quiz Real-Time dengan Boss Raid Mode'),
  para('Dosen meluncurkan sesi kuis langsung via Socket.IO WebSocket. Skor dihitung berdasarkan kebenaran jawaban dan kecepatan respons (answer_time_ms) untuk mendorong kewaspadaan kognitif. Boss Raid Mode adalah inovasi kolaboratif eksklusif ARKON: seluruh kelas menjawab satu soal "boss" secara bersama dengan skor akumulatif kolektif, mentransformasi evaluasi individu menjadi momen kohesi kelas yang mendorong diskusi spontan dan pembelajaran kolaboratif.'),

  h3('F-05: AI Tutor Kontekstual (Gemini 2.0 Flash)'),
  para('AI Tutor ARKON beroperasi sepenuhnya berbasis profil kemampuan IRT mahasiswa. Setelah mahasiswa menjawab salah dalam sesi quiz IRT, sistem secara otomatis memicu penjelasan personal (personalized-tutor): mengapa jawaban benar itu benar, satu analogi konkret, satu pertanyaan refleksi, dan rekomendasi subtopik lanjutan—semuanya dikalibrasi ke nilai theta mahasiswa saat itu. Dosen mendapatkan ringkasan analitik kelas berbasis AI (analytics-summary) yang menginterpretasikan distribusi theta dan N-Gain ke rekomendasi pedagogis. Seluruh endpoint AI dilindungi rate limiting 5 request/menit per IP dengan heuristic fallback agar platform tetap berfungsi meski Gemini API tidak tersedia.'),

  h3('F-06: Analytics Dashboard & N-Gain Measurement'),
  para('Dashboard dosen menampilkan: heatmap pemahaman kelas per topik, distribusi theta mahasiswa (histogram), trajektori theta individual lintas waktu, dan kalkulasi N-Gain Hake (1999):'),
  space(60,60),
  infoBox([
    [{ text: 'Formula N-Gain Hake (1999):', bold: true }],
    'g = (PostTest − PreTest) / (MaxScore − PreTest)',
    'Kategori: Tinggi (g ≥ 0,7) | Sedang (0,3 ≤ g < 0,7) | Rendah (g < 0,3)',
    [{ text: 'Target KPI ARKON: ', bold: true }, 'N-Gain ≥ 0,5 (Sedang-Tinggi) pada cohort pilot pertama'],
  ], C.greenBg, C.green),
  space(120),

  h3('F-07: Room-Based Learning & Sistem Gamifikasi'),
  para('Dosen membuat virtual classroom (Room) dengan kode unik dan data terisolasi per tenant. Sistem gamifikasi mencakup: ekonomi koin (quiz, daily login, achievement), lebih dari 20 tipe achievement badge, leaderboard lintas room, turnamen single-elimination, Study Groups kolaboratif, dan PC Shop virtual sebagai reward belajar. Gamifikasi dirancang terintegrasi organik dengan aktivitas akademik, bukan sebagai lapisan tambahan.'),

  h2('4.3 Keamanan Platform'),
  para('ARKON menerapkan keamanan berlapis sesuai standar lingkungan pendidikan multi-pengguna:'),
  space(60,60),
  makeTable(
    ['Layer Keamanan','Implementasi','Tujuan'],
    [
      ['Autentikasi','JWT access token + refresh token rotation, HttpOnly cookies','Mencegah session hijacking'],
      ['Otorisasi','RBAC ketat: middleware verifikasi role di setiap route','Memisahkan hak akses dosen dan mahasiswa'],
      ['Injeksi SQL','Parameterized queries (pg-pool), sanitasi input di semua endpoint','Mencegah SQL injection'],
      ['XSS & Header','Helmet.js, Content Security Policy (CSP) ketat','Mencegah cross-site scripting'],
      ['Rate Limiting','100 req/15 menit (auth), 5 req/menit (AI endpoint)','Mencegah brute force & abuse'],
      ['Password','Bcryptjs hashing, salt rounds adaptif','Menjamin kerahasiaan kredensial'],
      ['Monitoring','Sentry error monitoring di production','Deteksi anomali real-time'],
    ],
    [1800, 3360, 4200]
  ),
  pageBreak()
);

// ── BAB 5: DESAIN PRODUK ───────────────────────────────────────
children.push(
  h1('BAB 5. DESAIN PRODUK TEKNOLOGI DIGITAL'),
  h2('5.1 Prinsip Desain'),
  para('Desain ARKON dibangun di atas empat prinsip utama yang saling mendukung:'),
  numbered('Clarity First — Antarmuka dirancang untuk mengurangi cognitive load ekstrinsik (Sweller, 1988) dengan hierarki visual yang jelas, whitespace yang cukup, dan ikon yang konsisten.'),
  numbered('Progressive Disclosure — Fitur kompleks (IRT analytics, WebXR) diperkenalkan secara bertahap agar tidak membanjiri pengguna baru.'),
  numbered('Feedback Immediacy — Setiap aksi pengguna mendapat respons visual instan (animasi konfirmasi, progress bar, notifikasi AI Tutor) untuk mendukung loop belajar yang cepat.'),
  numbered('Inclusivity — Desain responsif untuk desktop dan mobile, kontras warna memenuhi WCAG 2.1 AA, dan teks tersedia dalam Bahasa Indonesia penuh.'),

  h2('5.2 Arsitektur Informasi dan Alur Pengguna'),
  para('Berikut adalah alur pengguna (user flow) utama untuk skenario sesi belajar adaptif—dari login hingga menerima feedback AI Tutor:'),
  space(60,60),
  infoBox([
    [{ text: 'User Flow Utama: Sesi Quiz Adaptif IRT', bold: true }],
    '1. Mahasiswa login → masuk Dashboard → memilih Room kelas aktif',
    '2. Memilih "Mulai Quiz IRT" → sistem membaca profil θ dari database',
    '3. Soal pertama ditampilkan berdasarkan Max Information Function (θ awal = 0,0)',
    '4. Mahasiswa menjawab → respons dikirim ke IRT engine',
    '5a. Jawaban BENAR → θ diperbarui ke atas → soal lebih sulit dipilih otomatis',
    '5b. Jawaban SALAH → θ diperbarui ke bawah + AI Tutor Gemini aktif otomatis',
    '6. AI Tutor menampilkan: penjelasan kontekstual + analogi + pertanyaan refleksi',
    '7. Setelah 20 soal → laporan theta, N-Gain proyeksi, dan rekomendasi topik',
  ], C.light, C.midBlue),
  space(120),

  h2('5.3 Desain Antarmuka Pengguna (UI Mockup)'),
  h3('Dashboard Mahasiswa'),
  para('Dashboard mahasiswa dirancang dengan tata letak kartu (card-based layout) menggunakan palet warna navy-gold yang mencerminkan identitas LIDM-UNNES. Komponen utama yang ditampilkan:'),
  bullet('Header: nama mahasiswa, nilai θ IRT saat ini (ditampilkan sebagai level kemampuan 1–10 untuk kemudahan interpretasi), streak harian, dan total koin.'),
  bullet('Kartu Modul: CPU Simulator, 3D Assembly Lab, AR Lab, Quiz IRT, Live Quiz—masing-masing dengan ikon distinctive, progress completion rate, dan tombol CTA.'),
  bullet('Sidebar Leaderboard: peringkat kelas real-time dengan animasi perubahan posisi.'),
  bullet('Feed AI Tutor: notifikasi rekomendasi belajar berbasis analisis θ terkini.'),
  space(80,80),
  h3('Antarmuka Quiz Adaptif IRT'),
  para('Halaman quiz menggunakan desain full-focus (sidebar tersembunyi) untuk meminimalkan distraksi. Elemen kunci:'),
  bullet('Progress Bar IRT: visualisasi θ mahasiswa sebagai skala linier dengan zone warna (merah–kuning–hijau) yang bergerak setelah setiap respons.'),
  bullet('Timer Visual: countdown ring animatif yang berubah warna saat waktu hampir habis.'),
  bullet('Area AI Tutor: panel slide-in yang muncul otomatis setelah jawaban salah, dengan teks terformat dan tombol "Pelajari Lebih Lanjut".'),
  bullet('Feedback Gamifikasi: animasi koin dan XP yang muncul setelah jawaban benar.'),
  space(80,80),
  h3('3D PC Assembly Lab'),
  para('Tampilan split-screen: panel kiri berisi canvas Three.js dengan model 3D interaktif (drag, zoom, rotate), panel kanan berisi daftar komponen dan panduan perakitan. Sistem validasi kompatibilitas memberikan feedback visual real-time (komponen incompatible berwarna merah, compatible berwarna hijau) dan pesan error deskriptif.'),
  space(80,80),
  h3('Dashboard Analitik Dosen'),
  para('Dashboard dosen dibagi menjadi tiga panel: (1) Class Overview dengan statistik agregat (rata-rata θ, distribusi N-Gain, tingkat penyelesaian per modul); (2) Student Heatmap dengan visualisasi pemahaman per topik per mahasiswa dalam format grid warna; dan (3) AI Analytics Summary yang menghasilkan narasi rekomendasi pedagogis berbasis Gemini 2.0 Flash.'),

  h2('5.4 Spesifikasi Teknis Antarmuka'),
  space(60,60),
  makeTable(
    ['Komponen UI','Teknologi','Spesifikasi'],
    [
      ['Framework UI','React 18 + Vite','HMR untuk development cepat, code splitting untuk bundle optimal'],
      ['Styling','Tailwind CSS','Utility-first, purge unused CSS, ukuran bundle <15KB'],
      ['Animasi','Framer Motion','60fps animations, physics-based spring untuk feedback gamifikasi'],
      ['Grafik Analitik','Recharts','Responsive SVG charts: histogram θ, heatmap, line chart trajektori'],
      ['3D Engine','Three.js + @react-three/fiber','WebGL rendering, GLTF loader, orbit controls, shadow maps'],
      ['AR Viewer','@google/model-viewer','WebXR Device API, ARCore/ARKit support, iOS QuickLook fallback'],
      ['Real-time','Socket.IO client','Auto-reconnect, namespace isolation, binary event support'],
      ['Responsivitas','CSS Grid + Flexbox','Mobile-first, breakpoints: 640px, 768px, 1024px, 1280px'],
    ],
    [2000, 2000, 5360]
  ),
  pageBreak()
);

// ── BAB 6: RENCANA IMPLEMENTASI ────────────────────────────────
children.push(
  h1('BAB 6. RENCANA IMPLEMENTASI DAN VALIDASI PENGEMBANGAN TEKNOLOGI'),
  h2('6.1 Rencana Implementasi'),
  para('Implementasi ARKON dirancang dalam tiga gelombang yang mencerminkan eskalasi skala pengguna dan kompleksitas validasi, dengan KPI terukur di setiap tahap:'),

  h3('Gelombang 1 — Validasi Internal (Juli – Agustus 2026)'),
  bullet('Pilot terbatas dengan 2 dosen dan 30 mahasiswa Teknik Informatika UNNES yang menempuh matakuliah Arsitektur dan Organisasi Komputer.'),
  bullet('Pengumpulan data pre-test dan post-test untuk kalkulasi N-Gain Hake sebagai bukti efektivitas pembelajaran pertama.'),
  bullet('Iterasi cepat berbasis observasi langsung: perbaikan bug UI/UX, kalibrasi ulang parameter IRT, penyesuaian narasi AI Tutor.'),
  bullet([new TextRun({ text: 'KPI Gelombang 1: ', font: 'Arial', size: 20, bold: true, color: C.primary }), new TextRun({ text: 'N-Gain ≥ 0,4 | TAM Perceived Usefulness ≥ 4,0/5,0 | Error Rate < 2%', font: 'Arial', size: 20, color: C.darkText })]),
  space(80,80),

  h3('Gelombang 2 — Pilot Multi-Institusi (September – November 2026)'),
  bullet('Ekspansi ke 3 perguruan tinggi mitra di Jawa Tengah dengan profil infrastruktur berbeda: PTN besar (UNDIP), PTS menengah (UPGRIS), PTKIN daerah (UIN Salatiga).'),
  bullet('Pengukuran comparative effectiveness antara kelas ARKON vs. kelas konvensional menggunakan N-Gain dan theta IRT sebagai instrumen.'),
  bullet('Pengumpulan umpan balik terstruktur dari dosen melalui kuesioner Technology Acceptance Model (TAM) untuk mengukur usability dan perceived usefulness.'),
  bullet([new TextRun({ text: 'KPI Gelombang 2: ', font: 'Arial', size: 20, bold: true, color: C.primary }), new TextRun({ text: 'N-Gain ≥ 0,5 | 150+ pengguna aktif | TAM ≥ 4,2/5,0 | Concurrency 100 user stabil', font: 'Arial', size: 20, color: C.darkText })]),
  space(80,80),

  h3('Gelombang 3 — Deployment Nasional (2027)'),
  bullet('Publikasi platform sebagai open-access dengan opsi Docker self-hosting untuk institusi yang ingin mendeploy secara mandiri.'),
  bullet('Integrasi ke dalam ekosistem MBKM sebagai sumber belajar digital yang diakui untuk rekognisi SKS.'),
  bullet('Pengembangan modul baru: Digital Electronics Simulator, Memory Hierarchy Visualizer, dan Pipeline Hazard Detector.'),
  bullet([new TextRun({ text: 'KPI Gelombang 3: ', font: 'Arial', size: 20, bold: true, color: C.primary }), new TextRun({ text: '500+ pengguna aktif | 10+ institusi | N-Gain ≥ 0,6 | Publikasi jurnal terindeks Sinta/Scopus', font: 'Arial', size: 20, color: C.darkText })]),

  h2('6.2 Rencana Validasi'),
  h3('Validasi Teknis'),
  space(60,60),
  makeTable(
    ['Jenis Validasi','Alat/Metode','Target / KPI'],
    [
      ['End-to-End Testing','Playwright — alur login, quiz adaptif, live quiz, laporan analitik','Coverage ≥ 85%, dijalankan otomatis tiap pull request via GitHub Actions'],
      ['Load Testing','k6 — simulasi 100 concurrent users','Response time REST API < 200ms; WebSocket latency < 50ms; error rate < 1%'],
      ['Security Audit','OWASP Top 10 checklist — endpoint autentikasi, IRT, AI','Zero critical/high vulnerability sebelum deployment produksi'],
      ['IRT Bank Health','Otomatis: minimum 20 soal per difficulty level sebelum sesi adaptif','Semua level tercukupi; alert otomatis jika di bawah threshold'],
      ['IRT Convergence','Verifikasi |Δθ| < 0,001 setelah minimum 15 respons per sesi','Konsisten dengan spesifikasi Newton-Raphson MLE'],
    ],
    [2000, 2800, 4560]
  ),
  space(120),
  h3('Validasi Pedagogis'),
  space(60,60),
  makeTable(
    ['Instrumen','Metode','Target KPI'],
    [
      ['N-Gain Hake (1999)','Pre-test → intervensi ARKON → post-test per cohort pilot','N-Gain ≥ 0,5 (Sedang-Tinggi) pada Gelombang 1–2'],
      ['Theta IRT Tracking','Analisis longitudinal distribusi θ per mahasiswa per room','Peningkatan rata-rata θ > 0,5 poin setelah 4 minggu penggunaan aktif'],
      ['TAM Questionnaire','Kuesioner 20-item skala Likert 1–5 untuk dosen dan mahasiswa','Perceived Usefulness ≥ 4,2 | Ease of Use ≥ 4,0 | Intention to Use ≥ 4,3'],
      ['3D Validity Correlation','Korelasi skor PC Assembly compatibility vs. nilai ujian akhir','Korelasi Pearson r ≥ 0,6 (signifikan, p < 0,05)'],
      ['Comparative Effectiveness','Kelas ARKON vs. kelas konvensional (quasi-experimental design)','N-Gain kelas ARKON secara statistik lebih tinggi (uji Mann-Whitney U)'],
      ['Time-on-Task','Analitik platform: durasi aktif per sesi per mahasiswa','Peningkatan ≥ 40% dibanding baseline kelas konvensional'],
    ],
    [2200, 2800, 4360]
  ),

  h2('6.3 Timeline Ringkas'),
  space(60,60),
  makeTable(
    ['Periode','Milestone Utama'],
    [
      ['Jan – Jun 2026','Finalisasi platform, security hardening, dokumentasi teknis lengkap'],
      ['Jul – Agu 2026','Gelombang 1: Pilot internal UNNES (30 mahasiswa, 2 dosen), kalkulasi N-Gain awal'],
      ['Sep – Nov 2026','Gelombang 2: Pilot multi-institusi (3 PT mitra), TAM survey, comparative study'],
      ['Des 2026 – Mar 2027','Iterasi berdasarkan hasil pilot, pengembangan modul baru, persiapan deployment nasional'],
      ['Apr – Des 2027','Gelombang 3: Open-access deployment nasional, integrasi MBKM, publikasi ilmiah'],
    ],
    [2400, 6960]
  ),
  pageBreak()
);

// ── BAB 7: VIDEO ───────────────────────────────────────────────
children.push(
  h1('BAB 7. TAUTAN VIDEO PROSES PENGEMBANGAN MODEL KARYA INOVASI'),
  para('Video proses pengembangan ARKON telah diunggah ke YouTube sesuai ketentuan LIDM 2027. Video berdurasi 3 menit (di luar intro dan subtitle) menampilkan:'),
  bullet('Demo CPU Visual Simulator dengan siklus Fetch-Decode-Execute step-by-step.'),
  bullet('Demo 3D PC Assembly Lab dengan proses perakitan komponen dan validasi kompatibilitas real-time.'),
  bullet('Demo AR Hardware Lab via kamera smartphone.'),
  bullet('Demo Quiz Adaptif IRT dengan visualisasi perubahan theta dan trigger AI Tutor.'),
  bullet('Demo Live Quiz Boss Raid Mode dalam simulasi kelas.'),
  bullet('Dashboard analitik dosen dengan heatmap dan N-Gain calculator.'),
  space(80,80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      borders: { top: { style: BorderStyle.SINGLE, size: 8, color: C.accent }, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.accent }, left: { style: BorderStyle.SINGLE, size: 8, color: C.accent }, right: { style: BorderStyle.SINGLE, size: 2, color: C.accent } },
      shading: { fill: C.yellowBg, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: [
        new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: 'Format Judul YouTube:', font: 'Arial', size: 20, bold: true, color: C.yellow })] }),
        new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: '"LIDM 2027 - Divisi Inovasi Teknologi Pendidikan - 350020 - Arkon - ARKON: Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif - Proposal"', font: 'Arial', size: 20, color: C.darkText })] }),
        new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: 'Tautan Video: ', font: 'Arial', size: 20, bold: true, color: C.yellow }), new TextRun({ text: 'https://www.youtube.com/watch?v=zc3qKGhdZYc', font: 'Arial', size: 20, italics: true, color: C.darkText })] }),
      ]
    })]})],
  }),
  pageBreak()
);

// ── BAB 8: DAFTAR PUSTAKA ──────────────────────────────────────
children.push(
  h1('DAFTAR PUSTAKA'),
  para('Seluruh referensi ditulis dalam format Vancouver Style.'),
  space(80,80),
  numbered('Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi RI. Data Program Studi Teknik Informatika Indonesia. Pangkalan Data Pendidikan Tinggi (PDDIKTI); 2024. [Diakses 15 Januari 2025]. Tersedia dari: https://pddikti.kemdikbud.go.id'),
  numbered('Ramadhani D, Kusuma A, Pratiwi S. Analisis Metode Penilaian Matakuliah Teknik Informatika di Perguruan Tinggi Indonesia. Jurnal Pendidikan Tinggi Indonesia. 2022;8(2):112–128.'),
  numbered('Sweller J. Cognitive load during problem solving: Effects on learning. Cognitive Science. 1988;12(2):257–285.'),
  numbered('Sweller J, van Merriënboer JJG, Paas FGWC. Cognitive architecture and instructional design. Educational Psychology Review. 1998;10(3):251–296.'),
  numbered('Rasch G. Probabilistic Models for Some Intelligence and Attainment Tests. Chicago: University of Chicago Press; 1960.'),
  numbered('Bond TG, Fox CM. Applying the Rasch Model: Fundamental Measurement in the Human Sciences. 3rd ed. New York: Routledge; 2015.'),
  numbered('Hake RR. Interactive-engagement versus traditional methods: A six-thousand-student survey of mechanics test data for introductory physics courses. American Journal of Physics. 1998;66(1):64–74.'),
  numbered('van der Linden WJ, Hambleton RK, editors. Handbook of Modern Item Response Theory. New York: Springer; 1997.'),
  numbered('Baker FB, Kim SH. Item Response Theory: Parameter Estimation Techniques. 2nd ed. New York: Marcel Dekker; 2004.'),
  numbered('Deterding S, Dixon D, Khaled R, Nacke L. From game design elements to gamefulness: Defining gamification. Proc 15th Int Acad MindTrek Conf. 2011:9–15.'),
  numbered('Hamari J, Koivisto J, Sarsa H. Does gamification work? A literature review of empirical studies on gamification. Proc 47th Hawaii Int Conf Syst Sci. 2014:3025–3034.'),
  numbered('Davis FD. Perceived usefulness, perceived ease of use, and user acceptance of information technology. MIS Quarterly. 1989;13(3):319–340.'),
  numbered('Mayer RE. Multimedia Learning. 2nd ed. Cambridge: Cambridge University Press; 2009.'),
  numbered('Clark RC, Mayer RE. e-Learning and the Science of Instruction. 4th ed. Hoboken: Wiley; 2016.'),
  numbered('Google LLC. Gemini API Documentation: Gemini 2.0 Flash Model. [Diakses 10 Desember 2024]. Tersedia dari: https://ai.google.dev/gemini-api/docs'),
  numbered('Three.js Contributors. Three.js Documentation r159. [Diakses 5 Januari 2025]. Tersedia dari: https://threejs.org/docs/'),
  numbered('W3C Immersive Web Working Group. WebXR Device API. W3C Working Draft; 2024. Tersedia dari: https://www.w3.org/TR/webxr/'),
  numbered('Socket.IO Contributors. Socket.IO Documentation v4.7. [Diakses 10 Januari 2025]. Tersedia dari: https://socket.io/docs/v4/'),
  numbered('Kementerian Pendidikan dan Kebudayaan RI. Peraturan Menteri Pendidikan dan Kebudayaan Nomor 3 Tahun 2020 tentang Standar Nasional Pendidikan Tinggi. Jakarta: Kemendikbud; 2020.'),
  numbered('Open Web Application Security Project (OWASP). OWASP Top Ten 2021. [Diakses 20 Januari 2025]. Tersedia dari: https://owasp.org/Top10/'),
  pageBreak()
);

// ── LAMPIRAN: Surat Pernyataan Keaslian ───────────────────────
children.push(
  h1('LAMPIRAN — SURAT PERNYATAAN KEASLIAN KARYA'),
  space(120),
  para('Yang bertanda tangan di bawah ini, kami selaku tim pengusul karya inovasi dalam Lomba Inovasi Digital Mahasiswa (LIDM) 2027 Divisi Inovasi Teknologi Digital Pendidikan:'),
  space(80,80),
  makeTable(
    ['No.','Nama','NIM','Posisi dalam Tim'],
    [
      ['1','Muhammad Nouval Ar-Rizqy','2504130037','Ketua'],
      ['2','Zulfia Tirta Irawan','2504130118','Anggota'],
      ['3','Fairuz Alfian Priasahasika','2504130116','Anggota'],
    ],
    [480, 2880, 1800, 4200]
  ),
  space(120),
  para('dengan ini menyatakan bahwa:'),
  numbered('Karya inovasi yang kami ajukan dengan judul "ARKON: Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif" adalah karya orisinal yang kami kembangkan sendiri dan belum pernah dilombakan dalam kompetisi apapun.'),
  numbered('Karya ini tidak mengandung unsur plagiasi, pengambilalihan karya pihak lain, atau pelanggaran hak kekayaan intelektual dalam bentuk apapun.'),
  numbered('Seluruh aset digital, kode pihak ketiga, dan referensi yang digunakan telah dicantumkan secara jelas dalam proposal dengan lisensi yang sesuai (MIT, Apache 2.0, Creative Commons, atau open-source equivalent).'),
  numbered('Porsi penggunaan teknologi AI (Google Gemini 2.0 Flash) dalam pengembangan karya tidak melebihi 25% dari keseluruhan karya, dan seluruh sumber telah disebutkan secara eksplisit.'),
  numbered('Apabila di kemudian hari ditemukan ketidaksesuaian dengan pernyataan ini, kami bersedia menerima konsekuensi berupa diskualifikasi dari kompetisi dan sanksi akademik sesuai peraturan yang berlaku.'),
  space(360),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'Semarang, Januari 2027', font: 'Arial', size: 20 })] }),
  space(80),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: 'Muhammad Nouval Ar-Rizqy', font: 'Arial', size: 20, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: 'NIM. 2504130037', font: 'Arial', size: 20 })] }),
);

// ══════════════════════════════════════════════════════════════════
//  ASSEMBLE DOCUMENT
// ══════════════════════════════════════════════════════════════════
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 20, color: C.darkText } }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: C.primary },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: C.midBlue },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 22, bold: true, font: 'Arial', color: C.primary },
        paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]
      },
      {
        reference: 'numbers',
        levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ]
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.primary, space: 4 } },
            spacing: { before: 0, after: 100 },
            children: [
              new TextRun({ text: 'ARKON — LIDM 2027 | Divisi ITDP | UNNES', font: 'Arial', size: 18, color: C.midBlue }),
              new TextRun({ text: '    Hal. ', font: 'Arial', size: 18, color: C.midBlue }),
              new PageNumber({ type: 'current' }),
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.accent, space: 4 } },
            spacing: { before: 80, after: 0 },
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: 'Program Studi Teknik Informatika — Fakultas MIPA — Universitas Negeri Semarang', font: 'Arial', size: 16, color: C.midBlue })]
          })
        ]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  // Path penyimpanan direvisi menjadi path lokal
  fs.writeFileSync('./ARKON_Proposal_LIDM_2027_REVISI_FINAL.docx', buffer);
  console.log('SUCCESS: ARKON proposal generated!');
}).catch(err => {
  console.error('ERROR:', err);
});
