const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat, Header, Footer,
  PageNumberElement, ImageRun
} = require('docx');
const fs = require('fs');

// ─── Palette ────────────────────────────────────────────────────
const C = {
  primary:   '1B3A6B',
  accent:    'D4870A',
  midBlue:   '2E5FAC',
  light:     'EBF1FA',
  grayBg:    'F5F7FA',
  grayBdr:   'BDC8DC',
  white:     'FFFFFF',
  dark:      '1A1A2E',
  green:     '0F5C32',
  greenBg:   'E6F4ED',
  gold:      '7D5A00',
  goldBg:    'FFF8E1',
};

// ─── Border presets ─────────────────────────────────────────────
const B1 = (color = C.grayBdr) => ({ style: BorderStyle.SINGLE, size: 4, color });
const allB = (color = C.grayBdr) => ({ top: B1(color), bottom: B1(color), left: B1(color), right: B1(color) });
const noneB = () => ({ top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'} });

// ─── Helpers ────────────────────────────────────────────────────
function pb() { return new Paragraph({ children: [new PageBreak()] }); }
function sp(before=80, after=80) { return new Paragraph({ spacing:{before,after}, children:[new TextRun('')] }); }

function run(text, opts={}) {
  return new TextRun({ text: String(text), font:'Arial', size: opts.size||20, bold:!!opts.bold,
    italics:!!opts.italics, color: opts.color||C.dark, underline: opts.underline ? {} : undefined,
    superScript: !!opts.sup });
}

function p(content, opts={}) {
  const align = opts.center ? AlignmentType.CENTER : opts.right ? AlignmentType.RIGHT : AlignmentType.JUSTIFIED;
  let children;
  if (typeof content === 'string') children = [run(content, opts)];
  else if (Array.isArray(content)) children = content;
  else children = [content];
  return new Paragraph({ alignment: align, spacing:{ before: opts.before||80, after: opts.after||120 },
    indent: opts.indent ? { left: opts.indent } : undefined, children });
}

// Caption paragraph (centered, italic, small)
function caption(text) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing:{ before:60, after:160 },
    children:[run(text, { size:18, italics:true, color:C.midBlue })] });
}

// Section heading styles
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1,
    spacing:{ before:400, after:180 },
    border:{ bottom:{ style:BorderStyle.SINGLE, size:8, color:C.accent, space:4 } },
    children:[run(text, { size:28, bold:true, color:C.primary })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing:{ before:300, after:140 },
    children:[run(text, { size:24, bold:true, color:C.midBlue })] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing:{ before:240, after:100 },
    children:[run(text, { size:22, bold:true, color:C.primary })] });
}

// Inline-mixed paragraph
function mp(parts, opts={}) {
  const children = parts.map(pt => typeof pt === 'string' ? run(pt) : run(pt.text, pt));
  const align = opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED;
  return new Paragraph({ alignment:align, spacing:{before:opts.before||80, after:opts.after||120}, children });
}

// Bullet (ref='bul') or numbered (ref='num')
function bul(content, level=0) {
  let children;
  if (typeof content === 'string') children = [run(content)];
  else if (Array.isArray(content)) children = content;
  else children = [content];
  return new Paragraph({ numbering:{reference:'bul', level},
    spacing:{before:60,after:60},
    children });
}
function num(content) {
  let children;
  if (typeof content === 'string') children = [run(content)];
  else if (Array.isArray(content)) children = content;
  else children = [content];
  return new Paragraph({ numbering:{reference:'num', level:0}, spacing:{before:80,after:80}, children });
}

// ─── Table builder ───────────────────────────────────────────────
// headers: string[], rows: (string|Paragraph[])[][], colWidths: number[]
// tableNum: e.g. "Tabel 1. ..."
function mkTable(tableNum, headers, rows, colWidths, note=null) {
  const total = colWidths.reduce((a,b)=>a+b,0);
  const hdrRow = new TableRow({ tableHeader:true, children: headers.map((h,i) =>
    new TableCell({
      width:{size:colWidths[i], type:WidthType.DXA},
      borders: allB(C.primary),
      shading:{fill:C.primary, type:ShadingType.CLEAR},
      margins:{top:100,bottom:100,left:140,right:140},
      verticalAlign:VerticalAlign.CENTER,
      children:[new Paragraph({ alignment:AlignmentType.CENTER,
        children:[run(h, {bold:true, color:C.white, size:19})] })]
    })
  )});

  const dataRows = rows.map((row, ri) => new TableRow({ children: row.map((c, ci) => {
    const bg = ri % 2 === 0 ? C.white : C.grayBg;
    return new TableCell({
      width:{size:colWidths[ci], type:WidthType.DXA},
      borders: allB(),
      shading:{fill:bg, type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:140,right:140},
      verticalAlign:VerticalAlign.CENTER,
      children: Array.isArray(c) ? c :
        [new Paragraph({ children:[run(String(c), {size:19})] })]
    });
  })}));

  const parts = [];
  if (tableNum) {
    parts.push(new Paragraph({ spacing:{before:180,after:60},
      children:[run(tableNum, {bold:true, size:20, color:C.primary})] }));
  }
  parts.push(new Table({ width:{size:total, type:WidthType.DXA}, columnWidths:colWidths, rows:[hdrRow,...dataRows] }));
  if (note) parts.push(new Paragraph({ spacing:{before:60,after:160},
    children:[run(note, {size:17, italics:true, color:C.midBlue})] }));
  else parts.push(sp(40,120));
  return parts;
}

// ─── Image placeholder box ───────────────────────────────────────
// Returns array of paragraphs: [label, box-table, caption]
function imgPlaceholder(figNum, captionText, heightRows=6) {
  // Build a single-cell table as a visible placeholder box
  const innerLines = [
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:80,after:40},
      children:[run('[ GAMBAR ]', {bold:true, size:24, color:C.midBlue})] }),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:40},
      children:[run(figNum, {bold:true, size:20, color:C.primary})] }),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:80},
      children:[run(captionText, {size:18, italics:true, color:C.dark})] }),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:40,after:0},
      children:[run('(Sumber: Dokumentasi Tim Pengembang ARKON, 2026)', {size:16, italics:true, color:'888888'})] }),
  ];

  const boxRow = new TableRow({ children:[
    new TableCell({
      width:{size:9026, type:WidthType.DXA},
      borders:{
        top:{style:BorderStyle.DASHED, size:6, color:C.midBlue},
        bottom:{style:BorderStyle.DASHED, size:6, color:C.midBlue},
        left:{style:BorderStyle.DASHED, size:6, color:C.midBlue},
        right:{style:BorderStyle.DASHED, size:6, color:C.midBlue},
      },
      shading:{fill:C.light, type:ShadingType.CLEAR},
      margins:{top:200,bottom:200,left:300,right:300},
      verticalAlign:VerticalAlign.CENTER,
      children: innerLines
    })
  ]});

  return [
    sp(120,40),
    new Table({ width:{size:9026, type:WidthType.DXA}, columnWidths:[9026], rows:[boxRow] }),
    sp(40,120),
  ];
}

// ─── Callout / info box ──────────────────────────────────────────
function callout(lines, bg=C.light, borderColor=C.primary) {
  const innerParas = lines.map(l =>
    new Paragraph({ spacing:{before:60,after:60},
      children: typeof l === 'string' ? [run(l)] :
        Array.isArray(l) ? l.map(x => typeof x==='string' ? run(x) : run(x.text,x)) : [l]
    })
  );
  return [
    new Table({ width:{size:9026, type:WidthType.DXA}, columnWidths:[9026], rows:[
      new TableRow({ children:[new TableCell({
        width:{size:9026, type:WidthType.DXA},
        borders:{
          top:{style:BorderStyle.SINGLE,size:12,color:borderColor},
          bottom:{style:BorderStyle.SINGLE,size:4,color:borderColor},
          left:{style:BorderStyle.SINGLE,size:12,color:borderColor},
          right:{style:BorderStyle.SINGLE,size:4,color:borderColor},
        },
        shading:{fill:bg, type:ShadingType.CLEAR},
        margins:{top:160,bottom:160,left:240,right:240},
        children:innerParas
      })]})
    ]}),
    sp(60,120),
  ];
}

// ═══════════════════════════════════════════════════════════════════
//  DOCUMENT CONTENT
// ═══════════════════════════════════════════════════════════════════
const children = [];

// ── HALAMAN SAMPUL ──────────────────────────────────────────────
children.push(
  sp(2000,120),
  p('PROPOSAL LOMBA INOVASI DIGITAL MAHASISWA (LIDM)', {bold:true,size:28,color:C.primary,center:true,before:0,after:80}),
  p('DIVISI INOVASI TEKNOLOGI DIGITAL PENDIDIKAN', {bold:true,size:26,color:C.primary,center:true,before:0,after:200}),
  ...imgPlaceholder('Logo UNNES & LIDM', 'Logo Universitas Negeri Semarang dan Logo LIDM', 3),
  sp(200,80),
  p('ARKON: Augmented Rasch-based Knowledge & Organization Nexus', {bold:true,size:26,color:C.accent,center:true,before:0,after:80}),
  p('Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI,\nIRT Rasch Model, dan Simulasi 3D Interaktif', {size:22,italics:true,color:C.dark,center:true,before:0,after:400}),
  p('Disusun oleh:', {bold:true,size:22,color:C.dark,center:true,before:0,after:80}),
  p('Muhammad Nouval Ar-Rizqy (2504130037)', {size:22,color:C.dark,center:true,before:0,after:60}),
  p('Zulfia Tirta Irawan (2504130118)', {size:22,color:C.dark,center:true,before:0,after:60}),
  p('Fairuz Alfian Priasahasika (2504130116)', {size:22,color:C.dark,center:true,before:0,after:600}),
  p('PROGRAM STUDI TEKNIK INFORMATIKA', {bold:true,size:22,color:C.primary,center:true,before:0,after:60}),
  p('FAKULTAS MATEMATIKA DAN ILMU PENGETAHUAN ALAM', {bold:true,size:22,color:C.primary,center:true,before:0,after:60}),
  p('UNIVERSITAS NEGERI SEMARANG', {bold:true,size:22,color:C.primary,center:true,before:0,after:60}),
  p('2027', {bold:true,size:22,color:C.primary,center:true,before:0,after:0}),
  pb()
);

// ── HALAMAN PENGESAHAN ──────────────────────────────────────────
children.push(
  h1('HALAMAN PENGESAHAN'),
  p('PROPOSAL KARYA LOMBA INOVASI DIGITAL MAHASISWA (LIDM)', {bold:true,size:22,color:C.primary,center:true}),
  sp(120,60),
  mp([{text:'1. Judul Karya    : ', bold:true}, 'ARKON — Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif']),
  mp([{text:'2. Divisi          : ', bold:true}, 'Inovasi Teknologi Digital Pendidikan (ITDP)']),
  mp([{text:'3. Jumlah Tim      : ', bold:true}, '3 (tiga) orang mahasiswa aktif']),
  mp([{text:'4. Susunan Tim     : ', bold:true}, '']),
  sp(40,60),
  ...mkTable('Tabel P.1 Susunan Tim Peserta ARKON — LIDM 2027',
    ['No.','Posisi','NIM','Nama Lengkap','Program Studi / Fakultas'],
    [
      ['1','Ketua','2504130037','Muhammad Nouval Ar-Rizqy','Teknik Informatika / FMIPA'],
      ['2','Anggota','2504130118','Zulfia Tirta Irawan','Teknik Informatika / FMIPA'],
      ['3','Anggota','2504130116','Fairuz Alfian Priasahasika','Teknik Informatika / FMIPA'],
    ],
    [480,1080,1440,2640,3386]
  ),
  sp(240,80),
  p('Semarang, Januari 2027', {right:true}),
  sp(80,80),
  ...mkTable(null,
    ['Dosen Pendamping','Ketua Tim Pengusul'],
    [[
      [new Paragraph({spacing:{before:720,after:60},children:[run('( ...................................... )', {size:20})]}),
       new Paragraph({children:[run('NIP/NUPTK :', {size:18,italics:true,color:C.midBlue})]})],
      [new Paragraph({spacing:{before:720,after:60},children:[run('Muhammad Nouval Ar-Rizqy', {size:20,bold:true})]}),
       new Paragraph({children:[run('NIM. 2504130037', {size:18,italics:true,color:C.midBlue})]})]
    ]],
    [4513,4513]
  ),
  pb()
);

// ── ABSTRAK ─────────────────────────────────────────────────────
children.push(
  h1('ABSTRAK'),
  p([run('Kata Kunci — ', {bold:true, color:C.primary}),
     run('Arsitektur Komputer · Adaptive Learning · IRT Rasch Model · Simulasi 3D · WebXR · AI Tutoring · N-Gain Hake · Gamifikasi · EdTech Indonesia', {italics:true})],
    {before:0,after:160}),
  p('Pembelajaran Arsitektur dan Organisasi Komputer di perguruan tinggi Indonesia menghadapi tiga hambatan sistemik: (1) keterbatasan infrastruktur laboratorium perangkat keras yang menciptakan kesenjangan kualitas antar-institusi; (2) sistem penilaian seragam yang mengabaikan heterogenitas kemampuan awal mahasiswa; dan (3) rendahnya keterlibatan kognitif akibat abstraksi materi yang tinggi. Data PDDIKTI 2024 mencatat bahwa 61,3% dari 4.582 program studi Teknik Informatika aktif di Indonesia tidak memiliki laboratorium perangkat keras yang memadai.¹ Survei diagnostik internal UNNES (2024, n = 187 mahasiswa) menunjukkan rata-rata nilai akhir matakuliah Arsitektur Komputer sebesar 58,7/100 dengan tingkat pemahaman konseptual hanya 43%.⁴'),
  p('Sebagai respons terhadap gap tersebut, ARKON dikembangkan sebagai platform edukasi berbasis web yang mengintegrasikan empat inovasi digital secara sinergis: (1) simulasi real-time siklus instruksi CPU (Fetch-Decode-Execute-Writeback); (2) laboratorium virtual perakitan perangkat keras 3D berbasis Three.js dan WebXR Augmented Reality; (3) sistem penilaian adaptif menggunakan Item Response Theory (IRT) Rasch Model 1-PL dengan estimasi kemampuan via algoritma Newton-Raphson Maximum Likelihood; serta (4) AI Tutor adaptif bertenaga Google Gemini 2.0 Flash yang memberikan penjelasan personal berbasis profil kemampuan IRT mahasiswa. ARKON secara langsung mengisi celah yang tidak dimiliki platform eksisting — PhET Simulations, Khan Academy, maupun Cisco Packet Tracer — dengan menyatukan simulasi hardware 3D, psikometri adaptif, dan AI tutoring dalam satu ekosistem berbahasa Indonesia yang sepenuhnya gratis. Target primer yang ditetapkan adalah N-Gain Hake ≥ 0,5 (kategori Sedang-Tinggi) pada cohort pilot pertama sebagai bukti efektivitas pembelajaran yang terukur.'),
  pb()
);

// ── DAFTAR ISI (manual, karena TOC butuh field update) ──────────
children.push(
  h1('DAFTAR ISI'),
  mp([{text:'Halaman Pengesahan', color:C.dark}, {text:' .........................................................', color:'999999'}, {text:'  2', color:C.dark}]),
  mp([{text:'Abstrak', color:C.dark}, {text:' ..............................................................................................', color:'999999'}, {text:'  3', color:C.dark}]),
  mp([{text:'Daftar Isi', color:C.dark}, {text:' ..........................................................................................', color:'999999'}, {text:'  4', color:C.dark}]),
  mp([{text:'Daftar Tabel', color:C.dark}, {text:' ......................................................................................', color:'999999'}, {text:'  5', color:C.dark}]),
  mp([{text:'Daftar Gambar', color:C.dark}, {text:' ...................................................................................', color:'999999'}, {text:'  5', color:C.dark}]),
  mp([{text:'BAB 1.  Latar Belakang', color:C.dark, bold:true}, {text:' ....................................................................', color:'999999'}, {text:'  6', color:C.dark}]),
  mp([{text:'   1.1  Permasalahan yang Dihadapi', color:C.dark}, {text:' .................................................', color:'999999'}, {text:'  6', color:C.dark}]),
  mp([{text:'   1.2  Gap Analysis: Solusi Eksisting dan Keterbatasannya', color:C.dark}, {text:' ..................', color:'999999'}, {text:'  8', color:C.dark}]),
  mp([{text:'   1.3  Relevansi SDGs dan Kebijakan Nasional', color:C.dark}, {text:' ....................................', color:'999999'}, {text:'  9', color:C.dark}]),
  mp([{text:'   1.4  Status Pengembangan Platform', color:C.dark}, {text:' ..............................................', color:'999999'}, {text:' 10', color:C.dark}]),
  mp([{text:'BAB 2.  Tujuan dan Manfaat', color:C.dark, bold:true}, {text:' ...................................................................', color:'999999'}, {text:' 11', color:C.dark}]),
  mp([{text:'BAB 3.  Metode Pengembangan Produk Teknologi Digital', color:C.dark, bold:true}, {text:' ..............................', color:'999999'}, {text:' 13', color:C.dark}]),
  mp([{text:'BAB 4.  Analisa Fungsional Teknologi Digital', color:C.dark, bold:true}, {text:' ............................................', color:'999999'}, {text:' 15', color:C.dark}]),
  mp([{text:'BAB 5.  Desain Produk Teknologi Digital', color:C.dark, bold:true}, {text:' ..................................................', color:'999999'}, {text:' 20', color:C.dark}]),
  mp([{text:'BAB 6.  Rencana Implementasi dan Validasi', color:C.dark, bold:true}, {text:' ................................................', color:'999999'}, {text:' 23', color:C.dark}]),
  mp([{text:'BAB 7.  Tautan Video Proses Pengembangan', color:C.dark, bold:true}, {text:' .................................................', color:'999999'}, {text:' 26', color:C.dark}]),
  mp([{text:'Daftar Pustaka', color:C.dark, bold:true}, {text:' ...............................................................................', color:'999999'}, {text:' 27', color:C.dark}]),
  mp([{text:'Lampiran — Surat Pernyataan Keaslian', color:C.dark, bold:true}, {text:' ...............................................', color:'999999'}, {text:' 29', color:C.dark}]),
  pb()
);

// ── DAFTAR TABEL & GAMBAR ────────────────────────────────────────
children.push(
  h1('DAFTAR TABEL'),
  p('Tabel P.1   Susunan Tim Peserta ARKON — LIDM 2027 .......................................................   2'),
  p('Tabel 1.1   Perbandingan Platform Edukasi Arsitektur Komputer (Gap Analysis) .......................   8'),
  p('Tabel 1.2   Status Pengembangan Modul dan Fitur ARKON ...................................................  10'),
  p('Tabel 3.1   Fase dan Aktivitas Pengembangan ARKON .......................................................  13'),
  p('Tabel 3.2   Tech Stack dan Justifikasi Pemilihan Teknologi ...............................................  14'),
  p('Tabel 4.1   Parameter dan Notasi IRT Rasch Model 1-PL ....................................................  16'),
  p('Tabel 4.2   Spesifikasi Keamanan Berlapis Platform ARKON .................................................  19'),
  p('Tabel 5.1   Spesifikasi Teknis Antarmuka Pengguna ARKON ..................................................  22'),
  p('Tabel 6.1   Rencana Implementasi Tiga Gelombang dan KPI Terukur .......................................  23'),
  p('Tabel 6.2   Instrumen Validasi Teknis ARKON ...............................................................  24'),
  p('Tabel 6.3   Instrumen Validasi Pedagogis ARKON ............................................................  25'),
  p('Tabel 6.4   Timeline Pengembangan dan Deployment ........................................................  26'),
  sp(120,120),
  h1('DAFTAR GAMBAR'),
  p('Gambar P.1  Logo Universitas Negeri Semarang dan Logo LIDM ..............................................   1'),
  p('Gambar 1.1  Persentase Program Studi Tanpa Lab Hardware Memadai (PDDIKTI, 2024) .......................   7'),
  p('Gambar 1.2  Distribusi Nilai Akhir Arsitektur Komputer UNNES (2024, n = 187) ...........................   7'),
  p('Gambar 3.1  Alur Metodologi Agile Iterative Development ARKON ...........................................  13'),
  p('Gambar 4.1  Arsitektur Sistem Three-Tier ARKON ............................................................  15'),
  p('Gambar 4.2  Antarmuka CPU Visual Simulator — Siklus Fetch-Decode-Execute-Writeback ....................  17'),
  p('Gambar 4.3  Antarmuka 3D PC Assembly Lab — Proses Perakitan dan Validasi Kompatibilitas ................  17'),
  p('Gambar 4.4  Tampilan AR Hardware Lab via Kamera Perangkat Mobile ........................................  18'),
  p('Gambar 4.5  Alur Kerja AI Tutor Adaptif Berbasis Profil IRT Mahasiswa ..................................  18'),
  p('Gambar 4.6  Dashboard Analitik Dosen — Heatmap Pemahaman dan Distribusi Theta ..........................  19'),
  p('Gambar 5.1  Wireframe Dashboard Mahasiswa .................................................................  20'),
  p('Gambar 5.2  Wireframe Antarmuka Quiz Adaptif IRT .........................................................  21'),
  p('Gambar 5.3  Wireframe 3D PC Assembly Lab ..................................................................  21'),
  p('Gambar 5.4  Wireframe Dashboard Analitik Dosen ............................................................  22'),
  p('Gambar 6.1  Rancangan Quasi-Experimental Design Validasi Pedagogis ARKON ...............................  25'),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 1: LATAR BELAKANG
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 1. LATAR BELAKANG'),
  h2('1.1 Permasalahan yang Dihadapi'),
  p('Mata kuliah Arsitektur dan Organisasi Komputer merupakan fondasi akademik yang menentukan kematangan berpikir komputasional mahasiswa Teknik Informatika. Tanpa penguasaan konsep ini, mahasiswa tidak dapat secara mendalam memahami bagaimana perangkat lunak berinteraksi dengan perangkat keras, sebuah kompetensi yang bersifat lintas-program dan sangat dibutuhkan industri. Namun, di balik urgensinya, pembelajaran matakuliah ini di Indonesia masih terjebak dalam tiga masalah struktural yang saling memperparah dan telah terdokumentasi secara empiris.'),

  h3('1.1.1  Masalah Pertama: Keterbatasan Infrastruktur Laboratorium Hardware'),
  p('Data PDDIKTI 2024 mencatat bahwa 61,3% dari total 4.582 program studi Teknik Informatika aktif di Indonesia tidak memiliki laboratorium perangkat keras yang memadai.¹ Praktikum arsitektur komputer yang ideal memerlukan komponen fisik seperti motherboard, CPU, RAM, GPU, PSU, dan media penyimpanan, dengan estimasi biaya pengadaan Rp 45—75 juta per unit laboratorium dan biaya perawatan tahunan berkisar Rp 10—15 juta. Angka tersebut tidak terjangkau bagi mayoritas perguruan tinggi swasta daerah. Akibatnya, mahasiswa di institusi dengan anggaran terbatas belajar secara teoritis tanpa pengalaman hands-on, sementara rekan mereka di universitas besar mendapatkan pengalaman yang jauh lebih kaya. Ketimpangan ini berpotensi mereproduksi disparitas kompetensi lulusan secara sistemik dan berkelanjutan.'),

  ...imgPlaceholder(
    'Gambar 1.1 Persentase Program Studi Teknik Informatika Tanpa Laboratorium Hardware Memadai di Indonesia (PDDIKTI, 2024)',
    'Bar chart distribusi per wilayah: Jawa 47,2% · Sumatra 68,9% · Kalimantan 74,1% · Sulawesi 79,3% · Papua & Maluku 88,6%',
    5
  ),

  h3('1.1.2  Masalah Kedua: Sistem Penilaian yang Tidak Adaptif'),
  p('Penelitian Ramadhani et al. (2022) pada 12 perguruan tinggi di Indonesia menemukan bahwa 94,7% dosen Teknik Informatika masih menggunakan soal kuis seragam tanpa mempertimbangkan perbedaan kemampuan awal (prior knowledge) mahasiswa.² Mahasiswa berkemampuan rendah kewalahan oleh soal yang terlalu sulit dan kehilangan motivasi belajar; sebaliknya, mahasiswa berkemampuan tinggi mengalami kebosanan karena tidak mendapat tantangan yang memadai. Tidak ada mekanisme yang secara dinamis menyesuaikan beban kognitif dengan posisi kemampuan individual, sehingga distribusi nilai yang dihasilkan tidak merepresentasikan kemampuan sesungguhnya.'),

  h3('1.1.3  Masalah Ketiga: Rendahnya Keterlibatan Kognitif pada Materi Abstrak'),
  p('Teori Cognitive Load yang dikembangkan Sweller (1988, 1994) menegaskan bahwa materi dengan tingkat abstraksi tinggi — seperti pipeline CPU, hierarki memori, dan fetch-decode-execute cycle — membebani working memory secara berlebihan ketika disajikan hanya melalui teks atau diagram statis.³ Tanpa representasi visual yang interaktif dan eksperimen langsung, mahasiswa cenderung menghafal tanpa memahami, menghasilkan pengetahuan yang rapuh dan tidak transferable ke konteks nyata. Survei internal UNNES (2024, n = 187 mahasiswa) mengkonfirmasi hal ini: 78,1% responden menyatakan kesulitan memahami konsep pipeline hazard dan memory hierarchy, dengan rata-rata skor pemahaman konseptual hanya 43% pada tes diagnostik awal.⁴'),

  ...imgPlaceholder(
    'Gambar 1.2 Distribusi Nilai Akhir Matakuliah Arsitektur Komputer UNNES (2024, n = 187 Mahasiswa)',
    'Histogram distribusi nilai: A (10,2%) · B (22,4%) · C (38,7%) · D (19,8%) · E (8,9%) · Rata-rata = 58,7/100',
    5
  ),

  h2('1.2 Gap Analysis: Solusi Eksisting dan Keterbatasannya'),
  p('Berbagai platform dan alat bantu telah hadir sebelum ARKON, namun masing-masing memiliki celah fundamental yang membatasi efektivitasnya untuk konteks pembelajaran arsitektur komputer di Indonesia. Analisis komparatif berikut menunjukkan posisi ARKON sebagai solusi yang mengisi gap yang tidak dimiliki satu pun platform eksisting secara simultan.'),
  sp(40,60),
  ...mkTable(
    'Tabel 1.1 Perbandingan Platform Edukasi Arsitektur Komputer dan Gap Analysis',
    ['Platform','Simulasi HW 3D','IRT Adaptif','AI Tutor','Gamifikasi','Bahasa Indonesia','Biaya'],
    [
      ['PhET Simulations','Parsial (fisika)','✗','✗','✗','Terbatas','Gratis'],
      ['Khan Academy','✗','✗','Parsial','Minimal','✗','Gratis'],
      ['Cisco Packet Tracer','Parsial (jaringan)','✗','✗','✗','✗','Berbayar'],
      ['Coursera / edX','✗','✗','✗','Minimal','✗','Berbayar'],
      ['MARIE Simulator','CPU saja','✗','✗','✗','✗','Gratis'],
      ['ARKON ✓','✓ CPU + 3D + AR','✓ Rasch 1-PL','✓ Gemini Flash','✓ Lengkap','✓ Penuh','Gratis'],
    ],
    [1800,1280,1120,1040,1040,1360,1386],
    'Keterangan: ✓ = tersedia penuh; ✗ = tidak tersedia; Parsial = tersedia sebagian dengan keterbatasan signifikan.'
  ),
  p('Dari analisis di atas tampak bahwa ARKON adalah satu-satunya platform yang secara simultan mengisi tiga gap kritis: (1) simulasi hardware 3D terintegrasi yang mencakup seluruh komponen PC, (2) penilaian psikometrik adaptif berbasis IRT yang dikalibrasi secara individual, dan (3) AI tutoring kontekstual yang terhubung dengan profil kemampuan mahasiswa — semuanya dalam Bahasa Indonesia yang lengkap, dapat diakses gratis, dan tanpa instalasi perangkat lunak tambahan.'),

  h2('1.3 Relevansi dengan SDGs dan Kebijakan Nasional'),
  p('ARKON dirancang selaras dengan empat Tujuan Pembangunan Berkelanjutan (SDGs) PBB yang relevan:'),
  bul([run('SDG 4 — Quality Education: ', {bold:true, color:C.primary}), run('Menyediakan akses pembelajaran interaktif dan setara bagi seluruh mahasiswa Indonesia, terlepas dari keterbatasan infrastruktur fisik institusi masing-masing.')]),
  bul([run('SDG 9 — Industry, Innovation and Infrastructure: ', {bold:true, color:C.primary}), run('Memanfaatkan konvergensi AI generatif, psikometri IRT, dan teknologi 3D/AR untuk membangun infrastruktur pendidikan teknik yang tangguh dan berorientasi masa depan.')]),
  bul([run('SDG 10 — Reduced Inequalities: ', {bold:true, color:C.primary}), run('Mereduksi ketimpangan akses laboratorium komputer dengan mengalihkan infrastruktur fisik ke model virtual, sehingga mahasiswa dari institusi dengan anggaran terbatas mendapatkan pengalaman belajar yang setara.')]),
  bul([run('SDG 12 — Responsible Consumption and Production: ', {bold:true, color:C.primary}), run('Laboratorium virtual 3D dan AR menekan kebutuhan komponen silikon, plastik, dan logam fisik, sekaligus meminimalkan potensi akumulasi limbah elektronik dari perangkat praktikum yang cepat usang.')]),
  sp(80,80),
  p('Dalam konteks kebijakan nasional, ARKON berkontribusi langsung pada implementasi Merdeka Belajar-Kampus Merdeka (MBKM) sebagaimana diamanatkan Permendikbud No. 3 Tahun 2020. Mahasiswa dapat menggunakan ARKON untuk rekognisi 20 SKS dalam bentuk proyek belajar mandiri atau pertukaran studi digital, karena platform ini menyediakan kurikulum terstruktur, rekaman aktivitas belajar yang dapat diverifikasi, dan laporan kompetensi berbasis IRT yang memenuhi standar PDDIKTI. Platform ini juga mendukung transformasi digital pendidikan tinggi sebagaimana tercantum dalam Peta Jalan Pendidikan Indonesia 2020—2035.'),

  h2('1.4 Status Pengembangan Platform'),
  p('ARKON telah melampaui fase konseptual dan berada dalam kondisi fungsional yang substansial. Seluruh komponen inti telah diimplementasikan dan dapat dioperasikan secara end-to-end. Backend telah diuji dengan 9 layer migrasi database (PostgreSQL) dan seluruh route API — meliputi autentikasi, IRT, analitik, gamifikasi, live quiz, dan simulator — beroperasi dengan validasi input ketat dan penanganan error berlapis.'),
  sp(40,60),
  ...mkTable(
    'Tabel 1.2 Status Pengembangan Modul dan Fitur ARKON (per Januari 2026)',
    ['Modul / Fitur','Status','Keterangan Teknis'],
    [
      ['CPU Visual Simulator','✓ Selesai','Svelte bundle, embed via iframe postMessage, siklus FDE step-by-step'],
      ['3D PC Assembly Lab','✓ Selesai','Three.js + GLB/GLTF, validasi kompatibilitas komponen real-time'],
      ['AR Hardware Lab','✓ Selesai','@google/model-viewer + WebXR, AR overlay via kamera mobile'],
      ['IRT Rasch Adaptive Quiz','✓ Selesai','Newton-Raphson MLE, Max Information Selection, 290 soal × 14 topik'],
      ['AI Tutor — Gemini 2.0 Flash','✓ Selesai','Trigger otomatis pasca-jawaban-salah, adaptive hint, rate-limit 5 req/mnt'],
      ['Live Quiz + Boss Raid Mode','✓ Selesai','Socket.IO WebSocket, skor akumulatif kolektif, mode Boss Raid'],
      ['Analytics + N-Gain Dashboard','✓ Selesai','Heatmap pemahaman, distribusi theta, kalkulasi N-Gain Hake otomatis'],
      ['Sistem Gamifikasi Lengkap','✓ Selesai','Koin, 20+ achievement badge, leaderboard, turnamen eliminasi'],
      ['Room-Based Classroom','✓ Selesai','Multi-tenant, kode unik per kelas, isolasi data per room'],
      ['CI/CD & Deployment','✓ Selesai','Docker Compose, GitHub Actions, Azure App Service'],
    ],
    [2520,1080,5426],
    'Keterangan: Platform dapat diakses via browser modern tanpa instalasi tambahan.'
  ),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 2: TUJUAN DAN MANFAAT
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 2. TUJUAN DAN MANFAAT'),
  h2('2.1 Tujuan'),
  p('Pengembangan ARKON memiliki empat tujuan yang dirumuskan secara SMART (Specific, Measurable, Achievable, Relevant, Time-bound):'),
  num('Menyediakan platform edukasi web Arsitektur dan Organisasi Komputer sebagai substitusi fungsional laboratorium hardware fisik yang dapat diakses dari browser modern tanpa instalasi perangkat lunak tambahan, dengan target 500 pengguna aktif pada akhir tahun 2027.'),
  num('Mengimplementasikan sistem penilaian adaptif berbasis IRT Rasch Model 1-PL yang menyesuaikan tingkat kesulitan soal secara real-time berdasarkan estimasi kemampuan (θ) mahasiswa menggunakan algoritma Newton-Raphson Maximum Likelihood, dengan target konvergensi estimasi |Δθ| < 0,001 setelah minimal 15 respons per sesi.'),
  num('Mengukur efektivitas pembelajaran secara kuantitatif melalui kalkulasi N-Gain Hake (1998) dengan target N-Gain ≥ 0,5 (kategori Sedang-Tinggi) pada cohort pilot pertama, untuk mendukung evaluasi kurikulum dan borang akreditasi BAN-PT.'),
  num('Meningkatkan keterlibatan kognitif mahasiswa melalui mekanisme gamifikasi terstruktur dengan target peningkatan waktu belajar aktif (time-on-task) minimal 40% dibandingkan kelas konvensional, diukur melalui analitik platform.'),

  h2('2.2 Manfaat'),
  h3('2.2.1  Bagi Mahasiswa'),
  bul('Akses gratis ke simulasi CPU real-time, laboratorium virtual perakitan PC 3D, dan AR Lab kapan saja dan di mana saja tanpa perangkat keras fisik, menghilangkan hambatan infrastruktur yang selama ini menciptakan ketimpangan kompetensi.'),
  bul('Soal kuis yang secara otomatis dikalibrasi ke level kemampuan individual menggunakan IRT Rasch Model, sehingga setiap mahasiswa mengerjakan soal yang tepat pada level yang tepat — tidak terlalu mudah hingga membosankan, tidak terlalu sulit hingga memfrustasi.'),
  bul('Umpan balik instan dari AI Tutor yang aktif secara otomatis ketika mahasiswa menjawab salah — memberikan penjelasan kontekstual, analogi konkret, pertanyaan refleksi, dan rekomendasi subtopik lanjutan yang semuanya dikalibrasi ke profil kemampuan IRT mahasiswa saat itu.'),
  bul('Motivasi belajar intrinsik yang diperkuat melalui pencapaian koin, lencana, dan persaingan sehat di leaderboard, mengubah belajar dari kewajiban menjadi pengalaman yang dinantikan.'),
  sp(80,60),
  h3('2.2.2  Bagi Dosen'),
  bul('Dashboard analitik real-time dengan heatmap pemahaman per topik, distribusi theta mahasiswa, dan trajektori kemampuan longitudinal — memungkinkan pengambilan keputusan pedagogis berbasis data, bukan intuisi semata.'),
  bul('AI Analytics Summary yang secara otomatis menginterpretasikan distribusi theta IRT dan nilai N-Gain ke dalam rekomendasi strategi pengajaran yang dapat langsung ditindaklanjuti dalam pertemuan berikutnya.'),
  bul('Live Quiz interaktif dengan Boss Raid Mode berbasis WebSocket, menggantikan kuis manual yang memakan waktu dengan sesi kolaboratif yang mendorong partisipasi aktif seluruh kelas.'),
  bul('Bank soal terkelola (CRUD + bulk import CSV) dengan threshold IRT otomatis untuk menjamin reliabilitas estimasi kemampuan sebelum sesi adaptif dimulai.'),
  sp(80,60),
  h3('2.2.3  Bagi Institusi Pendidikan'),
  bul('Penghematan signifikan biaya pengadaan dan perawatan komponen hardware fisik (estimasi Rp 45—75 juta per unit lab) dengan beralih ke infrastruktur virtual yang terjangkau dan mudah diperbarui.'),
  bul('Rekaman data pedagogis terstandar dan valid — N-Gain, distribusi theta, analitik per topik — untuk memperkuat borang akreditasi BAN-PT dan pelaporan PDDIKTI secara otomatis.'),
  bul('Deployment mandiri via Docker Compose atau cloud provider mana pun, sesuai kapasitas infrastruktur masing-masing institusi, tanpa ketergantungan vendor dan tanpa biaya lisensi.'),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 3: METODE
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 3. METODE PENGEMBANGAN PRODUK TEKNOLOGI DIGITAL'),
  h2('3.1 Metodologi: Agile Iterative Development'),
  p('ARKON dikembangkan menggunakan pendekatan Agile Iterative Development yang membagi proses menjadi sprint-sprint terstruktur dengan durasi dua minggu. Metodologi ini dipilih karena memungkinkan penyesuaian fitur berbasis feedback pengguna secara berkelanjutan, sesuai dengan karakteristik platform edukasi yang memerlukan iterasi konten, UX, dan logika pedagogis secara simultan. Setiap sprint mencakup siklus plan → build → test → review yang terdokumentasi melalui GitHub Issues dan pull request review. Pemilihan Agile juga dilatarbelakangi oleh kebutuhan integrasi multi-modul (IRT engine, WebXR, Socket.IO, Gemini API) yang masing-masing memiliki lifecycle pengembangan berbeda, sehingga sprint pendek memungkinkan deteksi konflik antarlayer lebih awal dan pengujian incremental yang lebih terstruktur.'),

  ...imgPlaceholder(
    'Gambar 3.1 Alur Metodologi Agile Iterative Development ARKON',
    'Diagram siklus: Plan → Build → Test → Review → Deploy → Feedback → Plan (iterasi berikutnya), dengan 4 fase bertumpuk',
    5
  ),

  h2('3.2 Tahapan Pengembangan'),
  sp(40,60),
  ...mkTable(
    'Tabel 3.1 Fase, Periode, dan Aktivitas Utama Pengembangan ARKON',
    ['Fase','Periode','Aktivitas Utama','Output Terukur'],
    [
      ['Fase 1 — Foundation','Nov 2024 – Jan 2025','Analisis kebutuhan; desain ERD PostgreSQL (9 layer migrasi); arsitektur three-tier; setup Docker Compose + CI/CD GitHub Actions; implementasi autentikasi JWT + RBAC multi-role','Infrastruktur production-ready; auth berfungsi dengan refresh token rotation'],
      ['Fase 2 — Core Engine','Feb – Apr 2025','IRT Rasch service (Newton-Raphson MLE); CPU Visual Simulator (Svelte + Rollup bundle); 3D PC Assembly (Three.js + GLB/GLTF); AR Lab (@google/model-viewer + WebXR); bank soal 290 pertanyaan × 14 level topik','MVP seluruh modul inti; konvergensi IRT terverifikasi'],
      ['Fase 3 — Intelligence','Mei – Jul 2025','Integrasi AI Tutor Gemini 2.0 Flash; Live Quiz Engine (Socket.IO); Boss Raid Mode; sistem gamifikasi lengkap (koin, achievement, leaderboard, turnamen); Analytics Dashboard + N-Gain calculator; Quiz Bank Manager CRUD','Platform terintegrasi penuh; semua modul terkoneksi'],
      ['Fase 4 — Hardening','Agu – Sep 2025','Security audit OWASP Top 10; E2E testing Playwright; load testing k6 (100 concurrent users); Sentry error monitoring; dokumentasi teknis lengkap; deployment Azure App Service','Platform production-ready; zero critical vulnerability'],
    ],
    [960,1440,4000,2626]
  ),

  h2('3.3 Tech Stack yang Digunakan'),
  sp(40,60),
  ...mkTable(
    'Tabel 3.2 Tech Stack ARKON dan Justifikasi Pemilihan Teknologi',
    ['Layer','Teknologi','Fungsi','Justifikasi Pemilihan'],
    [
      ['Frontend','React 18 + Vite, Tailwind CSS, Framer Motion, Recharts, Three.js, @react-three/fiber, @google/model-viewer','SPA interaktif, visualisasi 3D/AR, animasi UI, grafik analitik','React dipilih karena ekosistem komponen matang; Vite untuk HMR dan bundle optimal; Three.js untuk WebGL 3D yang performant'],
      ['Backend','Node.js + Express 5, Socket.IO, PostgreSQL (pg-pool), Redis, JWT + bcryptjs','REST API, WebSocket real-time, autentikasi aman, persistensi data','Node.js untuk konsistensi bahasa full-stack dan performa I/O non-blocking yang optimal untuk koneksi WebSocket berskala tinggi'],
      ['AI Engine','Google Generative AI SDK — Gemini 2.0 Flash','AI tutor kontekstual, analytics summary adaptif, generator soal','Gemini 2.0 Flash dipilih karena latensi <1 detik dan kemampuan pemahaman konteks matematis (formula IRT dan terminologi teknik komputer)'],
      ['CPU Simulator','Svelte + Rollup (embedded bundle)','Visualisasi siklus Fetch-Decode-Execute-Writeback step-by-step','Svelte dipilih untuk bundle size minimal (<50 KB) yang memungkinkan embed ringan via iframe dengan komunikasi postMessage'],
      ['DevOps','Docker Compose, Azure App Service, GitHub Actions, Playwright, k6, Sentry','CI/CD, containerisasi, E2E testing, load testing, monitoring','Stack ini dipilih untuk kemudahan self-hosting oleh institusi mitra tanpa ketergantungan vendor'],
    ],
    [1000,1800,1800,3426]
  ),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 4: ANALISA FUNGSIONAL
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 4. ANALISA FUNGSIONAL TEKNOLOGI DIGITAL'),
  h2('4.1 Arsitektur Sistem'),
  p('ARKON mengimplementasikan arsitektur three-tier yang bersih: React SPA sebagai presentation layer, Express REST API + Socket.IO sebagai business logic layer, dan PostgreSQL sebagai data layer. Komponen pendukung meliputi Redis untuk session caching, Supabase Storage untuk aset media 3D (file GLB/GLTF berukuran 5—15 MB per model), dan Google Gemini API sebagai layanan AI eksternal. Seluruh layer terisolasi dalam Docker container yang dikelola oleh Compose, memudahkan replikasi environment dan deployment mandiri oleh institusi lain.'),

  ...imgPlaceholder(
    'Gambar 4.1 Arsitektur Sistem Three-Tier ARKON — Diagram Komponen dan Alur Data',
    'Diagram: React SPA ↔ Express REST API + Socket.IO ↔ PostgreSQL | Redis (cache) | Supabase (media) | Gemini API | Docker Compose → Azure',
    6
  ),

  h2('4.2 Fitur Unggulan dan Inovasi Teknis'),
  h3('F-01: Adaptive Assessment dengan IRT Rasch Model 1-PL'),
  p('Ini adalah diferensiator teknis utama ARKON yang tidak dimiliki platform edukasi arsitektur komputer mana pun di Indonesia. Alih-alih soal seragam, ARKON menggunakan IRT Rasch Model untuk memilih soal berikutnya secara adaptif berdasarkan estimasi kemampuan (θ) mahasiswa yang diperbarui setelah setiap respons. Pendekatan ini berakar pada tradisi psikometri yang dikembangkan Rasch (1960) dan diformalisasi lebih lanjut oleh van der Linden dan Hambleton (1997).⁸'),
  sp(40,60),
  ...mkTable(
    'Tabel 4.1 Parameter dan Notasi IRT Rasch Model 1-PL dalam ARKON',
    ['Komponen','Notasi','Nilai / Spesifikasi','Keterangan'],
    [
      ['Fungsi Probabilitas','P(Xᵢ=1|θ,b)','exp(θ−b) / [1+exp(θ−b)]','Probabilitas menjawab benar item i'],
      ['Kemampuan Mahasiswa','θ (theta)','Rentang [−4, 4], real number','Estimasi kemampuan laten mahasiswa'],
      ['Tingkat Kesulitan Soal','b (beta)','−1,5 | 0,0 | 1,5 (3 level)','Parameter kesulitan butir soal'],
      ['Algoritma Estimasi','Newton-Raphson MLE','Maks. 15 iterasi, konvergensi ≤ 0,001','Maximum Likelihood Estimation iteratif'],
      ['Pemilihan Soal','Max Information','I(θ,b) = P(θ,b) × [1−P(θ,b)]','Soal dengan informasi tertinggi diprioritaskan'],
      ['Bank Soal','290 butir × 14 topik','Min. 20 soal per level kesulitan','IRT Bank Health Check otomatis sebelum sesi'],
    ],
    [2000,1200,2000,3826]
  ),

  ...imgPlaceholder(
    'Gambar 4.2 Antarmuka CPU Visual Simulator — Siklus Fetch-Decode-Execute-Writeback Real-Time',
    'Screenshot animasi aliran data: PC → MAR → RAM → MDR → IR → CU → ALU → ACC, dengan highlight step aktif berwarna biru',
    6
  ),

  h3('F-02: CPU Visual Simulator (Fetch-Decode-Execute Cycle)'),
  p('Simulator CPU interaktif yang memvisualisasikan eksekusi instruksi assembly secara step-by-step dalam waktu nyata. Mahasiswa dapat mengamati aliran data antara Program Counter (PC), Memory Address Register (MAR), Random Access Memory (RAM), Memory Data Register (MDR), Instruction Register (IR), Control Unit (CU), Arithmetic Logic Unit (ALU), dan Register Accumulator (ACC) pada setiap siklus instruksi. Simulator dibangun dengan Svelte dan di-bundle sebagai static asset yang ter-embed ke React via iframe dengan komunikasi postMessage, memastikan isolasi runtime tanpa konflik dependency.'),

  ...imgPlaceholder(
    'Gambar 4.3 Antarmuka 3D PC Assembly Lab — Proses Perakitan Komponen dan Validasi Kompatibilitas',
    'Screenshot: panel kiri = canvas Three.js dengan model 3D motherboard + CPU + RAM; panel kanan = checklist kompatibilitas; indikator merah/hijau',
    6
  ),

  h3('F-03: 3D PC Assembly Lab & WebXR Augmented Reality'),
  p('Mahasiswa merakit PC virtual secara interaktif menggunakan komponen 3D beresolusi tinggi (model GLB/GLTF: CPU Intel Core i3/AMD Ryzen 5, GPU RTX 3080/AMD RX 6700 XT, DDR4/DDR5 RAM, Motherboard ATX, PSU, CPU Cooler) dengan validasi kompatibilitas otomatis — socket CPU, tipe RAM, kapasitas daya PSU — yang berjalan secara real-time di frontend tanpa round-trip ke server. AR Lab menggunakan @google/model-viewer dengan WebXR Device API sehingga komponen hardware dapat divisualisasikan di atas permukaan nyata via kamera perangkat mobile, dengan dukungan ARCore (Android) dan ARKit (iOS). Fitur ini secara langsung mengeliminasi kebutuhan lab fisik senilai Rp 45—75 juta per unit.'),

  ...imgPlaceholder(
    'Gambar 4.4 Tampilan AR Hardware Lab — Komponen Hardware Divisualisasikan di Atas Meja Nyata',
    'Screenshot: model 3D CPU Intel Core i3 tertampil di atas meja fisik via kamera smartphone, dengan anotasi nama komponen',
    5
  ),

  h3('F-04: Live Quiz Real-Time dengan Boss Raid Mode'),
  p('Dosen meluncurkan sesi kuis langsung via Socket.IO WebSocket. Skor dihitung berdasarkan kebenaran jawaban dan kecepatan respons (answer_time_ms) untuk mendorong kewaspadaan kognitif. Boss Raid Mode adalah inovasi kolaboratif eksklusif ARKON: seluruh kelas menjawab satu soal "bos" secara bersama dengan skor akumulatif kolektif, mentransformasi evaluasi individu menjadi momen kohesi kelas yang mendorong diskusi spontan dan pembelajaran kolaboratif. Desain ini terinspirasi penelitian Hamari et al. (2014) yang menunjukkan bahwa elemen kooperatif dalam gamifikasi secara signifikan meningkatkan engagement dan retensi pengetahuan.¹¹'),

  h3('F-05: AI Tutor Kontekstual (Gemini 2.0 Flash)'),
  p('AI Tutor ARKON beroperasi sepenuhnya berbasis profil kemampuan IRT mahasiswa — sebuah keunggulan yang tidak dimiliki platform AI tutoring generik mana pun. Setelah mahasiswa menjawab salah dalam sesi quiz IRT, sistem secara otomatis memicu endpoint personalized-tutor dengan payload yang menyertakan: (a) soal yang dijawab salah, (b) jawaban mahasiswa, (c) jawaban benar, (d) nilai theta saat ini, dan (e) riwayat topik lemah mahasiswa. Gemini 2.0 Flash kemudian menghasilkan respons yang mencakup: penjelasan mengapa jawaban benar itu benar, satu analogi konkret yang dikalibrasi ke level abstraksi mahasiswa, satu pertanyaan refleksi yang mendorong berpikir lebih dalam, dan rekomendasi subtopik lanjutan. Seluruh endpoint AI dilindungi rate limiting 5 request/menit per IP dengan heuristic fallback sehingga platform tetap berfungsi penuh meski Gemini API tidak tersedia.'),

  ...imgPlaceholder(
    'Gambar 4.5 Alur Kerja AI Tutor Adaptif ARKON — Dari Jawaban Salah hingga Penjelasan Personal',
    'Flowchart: Jawaban salah → Extract context (θ, topik lemah, soal) → Gemini API → Tampilkan: penjelasan + analogi + refleksi + rekomendasi',
    5
  ),

  h3('F-06: Analytics Dashboard & N-Gain Measurement'),
  p([run('Dashboard dosen menampilkan: (1) heatmap pemahaman kelas per topik, (2) distribusi theta mahasiswa (histogram), (3) trajektori theta individual lintas waktu, dan (4) kalkulasi N-Gain Hake (1998)⁷ yang mengukur efektivitas pembelajaran pre-post. Formula N-Gain: '), run('g = (PostTest − PreTest) / (MaxScore − PreTest)', {bold:true}), run('. Kategori interpretasi: Tinggi (g ≥ 0,7), Sedang (0,3 ≤ g < 0,7), Rendah (g < 0,3). Target KPI ARKON adalah N-Gain ≥ 0,5 pada cohort pilot pertama.')]),

  ...imgPlaceholder(
    'Gambar 4.6 Dashboard Analitik Dosen — Heatmap Pemahaman Kelas dan Distribusi Theta IRT',
    'Screenshot: grid heatmap warna merah-kuning-hijau per topik per mahasiswa; histogram distribusi theta; grafik trajektori θ individual',
    6
  ),

  h3('F-07: Room-Based Learning & Sistem Gamifikasi'),
  p('Dosen membuat virtual classroom (Room) dengan kode unik dan data terisolasi per tenant, memungkinkan penggunaan multi-kelas secara bersamaan tanpa interferensi data. Sistem gamifikasi yang terintegrasi organik dengan aktivitas akademik — bukan sekadar lapisan tambahan — mencakup: ekonomi koin (diperoleh dari menjawab kuis, login harian, menyelesaikan achievement), lebih dari 20 tipe achievement badge dengan kriteria yang terhubung ke pencapaian akademik nyata, leaderboard lintas room, turnamen single-elimination berbatas waktu, Study Groups kolaboratif, dan PC Shop virtual di mana mahasiswa dapat "membeli" komponen PC dengan koin sebagai reward belajar. Desain ini mengacu pada prinsip Self-Determination Theory (Ryan & Deci, 2000) yang menekankan motivasi intrinsik melalui otonomi, kompetensi, dan keterkaitan sosial.'),

  h2('4.3 Keamanan Platform'),
  sp(40,60),
  ...mkTable(
    'Tabel 4.2 Spesifikasi Keamanan Berlapis Platform ARKON',
    ['Layer Keamanan','Implementasi Teknis','Ancaman yang Dimitigasi'],
    [
      ['Autentikasi','JWT access token + refresh token rotation; HttpOnly cookies untuk refresh token','Session hijacking, token replay attack'],
      ['Otorisasi','RBAC ketat: middleware verifikasi role di setiap route; endpoint dosen dan mahasiswa dipisah tegas','Privilege escalation, unauthorized data access'],
      ['Injeksi SQL','Parameterized queries (pg-pool); sanitasi input di semua endpoint via middleware','SQL injection, data tampering'],
      ['XSS & Header Security','Helmet.js; Content Security Policy (CSP) ketat; X-Frame-Options: DENY','Cross-site scripting, clickjacking'],
      ['Rate Limiting','100 req/15 menit (auth endpoint); 5 req/menit (AI endpoint)','Brute force attack, API abuse'],
      ['Password Storage','Bcryptjs hashing dengan salt rounds adaptif (cost factor 12)','Credential theft, rainbow table attack'],
      ['Error Monitoring','Sentry error monitoring + alerting di production environment','Anomali real-time, silent failures'],
    ],
    [1800,3400,4026]
  ),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 5: DESAIN PRODUK
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 5. DESAIN PRODUK TEKNOLOGI DIGITAL'),
  h2('5.1 Prinsip Desain'),
  p('Desain antarmuka ARKON dibangun di atas empat prinsip utama yang saling mendukung dan berakar pada riset Human-Computer Interaction (HCI) serta teori pembelajaran multimedia:'),
  num('Clarity First — Antarmuka dirancang untuk meminimalkan cognitive load ekstrinsik (Sweller, 1988) melalui hierarki visual yang jelas, whitespace yang memadai, dan sistem ikon yang konsisten. Setiap elemen UI memiliki satu fungsi yang jelas dan tidak ambigu.'),
  num('Progressive Disclosure — Fitur kompleks seperti IRT analytics dan WebXR diperkenalkan secara bertahap melalui onboarding interaktif, sehingga pengguna baru tidak dibanjiri opsi pada kunjungan pertama. Pengguna lanjutan dapat mengaktifkan fitur tambahan secara sadar.'),
  num('Feedback Immediacy — Setiap aksi pengguna mendapat respons visual instan: animasi konfirmasi, progress bar theta yang bergerak setelah setiap respons, dan notifikasi AI Tutor yang muncul dalam hitungan detik. Prinsip ini mendukung loop belajar yang cepat dan efektif.'),
  num('Inclusivity & Accessibility — Desain responsif untuk desktop dan mobile; kontras warna memenuhi standar WCAG 2.1 AA; seluruh konten tersedia dalam Bahasa Indonesia; dan platform berfungsi penuh pada koneksi internet minimal 1 Mbps.'),

  h2('5.2 Alur Pengguna (User Flow) Utama'),
  p('Berikut adalah user flow untuk skenario belajar adaptif yang merupakan alur penggunaan terpenting dalam ARKON — dari login hingga menerima umpan balik AI Tutor:'),
  sp(40,60),
  ...callout([
    [{text:'User Flow: Sesi Quiz Adaptif IRT (Alur Utama)', bold:true, color:C.primary}],
    '1.  Mahasiswa login → masuk Dashboard → memilih Room kelas aktif yang diberikan dosen',
    '2.  Memilih "Mulai Quiz IRT" → sistem membaca profil θ persisten dari database (θ awal = 0,0 jika sesi pertama)',
    '3.  IRT engine memilih soal pertama menggunakan Max Information Function terhadap θ saat ini',
    '4.  Mahasiswa membaca soal, memilih jawaban → respons dikirim ke IRT engine backend',
    '5a. Jawaban BENAR → θ diperbarui ke atas via Newton-Raphson MLE → soal lebih sulit dipilih otomatis',
    '5b. Jawaban SALAH → θ diperbarui ke bawah → AI Tutor Gemini 2.0 Flash aktif otomatis dalam <1 detik',
    '6.  AI Tutor menampilkan: penjelasan kontekstual + analogi + pertanyaan refleksi + rekomendasi subtopik',
    '7.  Setelah 20 soal → laporan akhir: nilai θ terkini, N-Gain proyeksi, heatmap topik, rekomendasi belajar selanjutnya',
  ], C.light, C.primary),

  h2('5.3 Desain Antarmuka Pengguna (Wireframe)'),
  h3('5.3.1  Dashboard Mahasiswa'),
  p('Dashboard mahasiswa dirancang dengan tata letak kartu (card-based layout) menggunakan palet warna navy-gold yang mencerminkan identitas akademik UNNES. Prinsip gestalt proximity dan similarity diterapkan untuk mengelompokkan fitur terkait secara visual.'),
  bul('Header navigasi: nama mahasiswa, nilai θ ditampilkan sebagai "Level Kemampuan 1—10" untuk kemudahan interpretasi non-teknis, streak harian, dan total koin dengan animasi incremental.'),
  bul('Grid modul (2×3): CPU Simulator, 3D Assembly Lab, AR Lab, Quiz IRT, Live Quiz, Leaderboard — masing-masing dengan ikon distinctive, persentase completion, dan tombol Call-to-Action primer.'),
  bul('Panel samping: leaderboard real-time kelas dengan animasi perubahan peringkat, notifikasi achievement terbaru, dan feed rekomendasi AI Tutor.'),

  ...imgPlaceholder(
    'Gambar 5.1 Wireframe Dashboard Mahasiswa ARKON',
    'Mockup: header (nama + θ-level + streak + koin) | grid 2×3 modul | sidebar leaderboard + feed AI Tutor; palet navy-gold-white',
    7
  ),

  h3('5.3.2  Antarmuka Quiz Adaptif IRT'),
  p('Halaman quiz menggunakan desain full-focus (sidebar tersembunyi secara otomatis) untuk meminimalkan distraksi eksternal dan memaksimalkan perhatian kognitif pada konten soal.'),
  bul('Progress bar IRT: visualisasi θ mahasiswa sebagai skala linier dengan zona warna merah—kuning—hijau yang bergerak secara animatif setelah setiap respons, memberikan feedback visual kemajuan yang intuitif.'),
  bul('Timer visual: countdown ring animatif yang berubah warna dari hijau (>30 detik) menjadi kuning (10—30 detik) dan merah (<10 detik) untuk mendorong respons yang tepat waktu.'),
  bul('Panel AI Tutor: slide-in drawer yang muncul otomatis setelah jawaban salah, dengan teks terformat (markdown rendering) dan tombol "Pelajari Lebih Lanjut" yang menautkan ke subtopik relevan.'),
  bul('Feedback gamifikasi: animasi burst koin dan partikel XP yang muncul setelah jawaban benar, disertai sound effect opsional.'),

  ...imgPlaceholder(
    'Gambar 5.2 Wireframe Antarmuka Quiz Adaptif IRT ARKON',
    'Mockup: area soal (tengah) | progress bar theta (atas) | timer ring (kanan atas) | panel AI Tutor slide-in (kanan) | animasi koin (bawah)',
    7
  ),

  h3('5.3.3  3D PC Assembly Lab'),
  p('Tampilan split-screen horizontal dengan rasio 60:40: panel kiri berisi canvas Three.js dengan model 3D interaktif beresolusi tinggi (drag untuk memutar, scroll untuk zoom, klik komponen untuk informasi detail); panel kanan berisi daftar komponen tersedia, panduan perakitan step-by-step, dan checker kompatibilitas real-time.'),

  ...imgPlaceholder(
    'Gambar 5.3 Wireframe 3D PC Assembly Lab ARKON',
    'Mockup: panel kiri = canvas 3D motherboard + komponen terapung; panel kanan = daftar komponen + step guide + checklist kompatibilitas hijau/merah',
    7
  ),

  h3('5.3.4  Dashboard Analitik Dosen'),
  p('Dashboard dosen dibagi menjadi tiga panel vertikal: (1) Class Overview dengan statistik agregat (rata-rata θ kelas, rentang distribusi, N-Gain terkini, tingkat penyelesaian per modul) dalam kartu metrik berwarna; (2) Student Heatmap dengan grid warna merah—kuning—hijau yang menampilkan pemahaman per topik per mahasiswa secara simultan; dan (3) AI Analytics Summary yang menghasilkan narasi rekomendasi pedagogis berbasis interpretasi Gemini 2.0 Flash terhadap pola data kelas.'),

  ...imgPlaceholder(
    'Gambar 5.4 Wireframe Dashboard Analitik Dosen ARKON',
    'Mockup: panel kiri = kartu metrik (θ rata-rata, N-Gain, completion); panel tengah = heatmap; panel kanan = AI Analytics Summary narasi',
    7
  ),

  h2('5.4 Spesifikasi Teknis Antarmuka Pengguna'),
  sp(40,60),
  ...mkTable(
    'Tabel 5.1 Spesifikasi Teknis Antarmuka Pengguna ARKON',
    ['Komponen','Teknologi','Spesifikasi Teknis'],
    [
      ['Framework UI','React 18 + Vite','Code splitting otomatis; bundle utama < 200 KB (gzipped); HMR untuk development'],
      ['Sistem Styling','Tailwind CSS v3','Utility-first; purge unused CSS; ukuran CSS produksi < 15 KB'],
      ['Animasi & Transisi','Framer Motion v10','60 fps physics-based spring; GPU-accelerated transforms; reduced-motion support'],
      ['Grafik Analitik','Recharts v2','Responsive SVG: histogram θ, heatmap, line chart trajektori, bar chart N-Gain'],
      ['3D Rendering Engine','Three.js r159 + @react-three/fiber','WebGL 2.0; GLTF/GLB loader; orbit controls; PCF shadow maps; Draco compression'],
      ['AR Viewer','@google/model-viewer v3','WebXR Device API; ARCore/ARKit support; iOS USDZ QuickLook fallback'],
      ['Real-time Communication','Socket.IO client v4','Auto-reconnect; namespace isolation; binary event; heartbeat 25 detik'],
      ['Responsivitas','CSS Grid + Flexbox','Mobile-first; breakpoints: 640 / 768 / 1024 / 1280 px; min-width 320 px'],
    ],
    [1800,1800,5426]
  ),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 6: RENCANA IMPLEMENTASI & VALIDASI
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 6. RENCANA IMPLEMENTASI DAN VALIDASI PENGEMBANGAN TEKNOLOGI'),
  h2('6.1 Rencana Implementasi'),
  p('Implementasi ARKON dirancang dalam tiga gelombang yang mencerminkan eskalasi skala pengguna dan kompleksitas validasi. Setiap gelombang memiliki KPI terukur yang berfungsi sebagai gate untuk melanjutkan ke gelombang berikutnya.'),
  sp(40,60),
  ...mkTable(
    'Tabel 6.1 Rencana Implementasi Tiga Gelombang ARKON dan KPI Terukur',
    ['Gelombang','Periode','Cakupan','Aktivitas Utama','KPI Target'],
    [
      ['Gelombang 1\nValidasi Internal','Jul – Agu 2026','UNNES (2 dosen, 30 mahasiswa)','Pilot terbatas; pre-test & post-test N-Gain; observasi UX; kalibrasi ulang IRT; penyesuaian narasi AI Tutor','N-Gain ≥ 0,4; TAM PU ≥ 4,0/5,0; error rate < 2%'],
      ['Gelombang 2\nPilot Multi-Institusi','Sep – Nov 2026','3 PT mitra Jawa Tengah (PTN, PTS, PTKIN)','Ekspansi 3 institusi berbeda profil; comparative effectiveness vs. kelas konvensional; TAM survey dosen','N-Gain ≥ 0,5; 150+ pengguna; TAM ≥ 4,2/5,0; concurrency 100 user stabil'],
      ['Gelombang 3\nDeployment Nasional','Jan – Des 2027','Open-access nasional; integrasi MBKM','Docker self-hosting; integrasi ekosistem MBKM; modul baru (Digital Electronics, Memory Hierarchy); publikasi ilmiah','500+ pengguna; 10+ institusi; N-Gain ≥ 0,6; 1 publikasi Sinta/Scopus'],
    ],
    [1200,1200,1200,2200,3226]
  ),

  h2('6.2 Rencana Validasi Teknis'),
  sp(40,60),
  ...mkTable(
    'Tabel 6.2 Instrumen dan Target Validasi Teknis ARKON',
    ['Jenis Validasi','Alat / Metode','Skenario Pengujian','Target / Kriteria Lulus'],
    [
      ['End-to-End Testing','Playwright v1.40','Alur: login → quiz adaptif → live quiz → laporan analitik → logout','Coverage ≥ 85%; dijalankan otomatis di setiap pull request via GitHub Actions'],
      ['Load Testing','k6 v0.49','Simulasi 100 concurrent users selama 10 menit; campuran REST API dan WebSocket','REST API < 200 ms; WebSocket latency < 50 ms; error rate < 1%'],
      ['Security Audit','OWASP Top 10 Checklist 2021','Endpoint autentikasi, IRT, AI; injection, broken auth, XSS, IDOR','Zero critical/high vulnerability; medium dimitigasi sebelum deployment'],
      ['IRT Bank Health','Otomatis sebelum sesi adaptif','Verifikasi kecukupan soal per difficulty level; flag jika di bawah threshold','Minimum 20 soal per level; alert email otomatis jika threshold terancam'],
      ['IRT Convergence','Analisis log theta per sesi','Verifikasi |Δθ| < 0,001 setelah minimum 15 respons pada 100 sesi acak','≥ 95% sesi konvergen; konsisten dengan spesifikasi Newton-Raphson MLE'],
    ],
    [1400,1200,2200,4226]
  ),

  h2('6.3 Rencana Validasi Pedagogis'),
  p('Validasi pedagogis dilakukan menggunakan quasi-experimental design dengan pre-test — intervensi ARKON — post-test, disertai kelompok kontrol (kelas konvensional) untuk menghasilkan bukti kausal efektivitas platform. Desain ini mengikuti standar penelitian pendidikan yang dikemukakan Creswell (2014).'),

  ...imgPlaceholder(
    'Gambar 6.1 Rancangan Quasi-Experimental Design Validasi Pedagogis ARKON',
    'Diagram: Kelompok Eksperimen (Pre-test → Intervensi ARKON → Post-test) vs Kelompok Kontrol (Pre-test → Pembelajaran Konvensional → Post-test) → Komparasi N-Gain + Uji Mann-Whitney U',
    5
  ),

  sp(40,60),
  ...mkTable(
    'Tabel 6.3 Instrumen Validasi Pedagogis ARKON',
    ['Instrumen','Metode','Sampel','Target KPI'],
    [
      ['N-Gain Hake (1998)','Pre-test (sebelum ARKON) → intervensi 4 minggu → post-test; formula g = (Post−Pre)/(Max−Pre)','30 mahasiswa Gelombang 1; 150+ Gelombang 2','N-Gain ≥ 0,5 (Sedang-Tinggi) pada Gelombang 1; ≥ 0,6 pada Gelombang 2'],
      ['Theta IRT Tracking','Analisis longitudinal distribusi θ per mahasiswa per room selama 4 minggu','Seluruh peserta aktif per gelombang','Peningkatan rata-rata θ > 0,5 poin setelah 4 minggu penggunaan aktif'],
      ['TAM Questionnaire','20-item skala Likert 1—5; konstruk: Perceived Usefulness, Ease of Use, Intention to Use','Dosen dan mahasiswa tiap gelombang','PU ≥ 4,2; EoU ≥ 4,0; ITU ≥ 4,3'],
      ['3D Validity Correlation','Korelasi skor PC Assembly compatibility vs. nilai ujian akhir matakuliah','Peserta Gelombang 1 & 2','Korelasi Pearson r ≥ 0,6 (p < 0,05)'],
      ['Comparative Effectiveness','Kelas ARKON vs. kelas konvensional; uji Mann-Whitney U untuk N-Gain','Min. 30 per kelompok per gelombang','N-Gain kelas ARKON secara statistik lebih tinggi (p < 0,05, efek sedang)'],
      ['Time-on-Task','Log analitik platform: durasi sesi aktif per mahasiswa per hari','Seluruh peserta terdaftar','Peningkatan ≥ 40% dibanding baseline kelas konvensional'],
    ],
    [1600,2200,1600,3626]
  ),

  h2('6.4 Timeline Pengembangan dan Deployment'),
  sp(40,60),
  ...mkTable(
    'Tabel 6.4 Timeline Pengembangan, Validasi, dan Deployment ARKON',
    ['Periode','Fase / Milestone','Output Terukur'],
    [
      ['Nov 2024 – Sep 2025','Fase 1—4: Foundation, Core Engine, Intelligence, Hardening','Platform production-ready; seluruh 10 modul selesai; CI/CD aktif'],
      ['Okt – Des 2025','Finalisasi dokumentasi teknis; security audit final; penyempurnaan UX berdasarkan usability testing internal','Dokumentasi lengkap; zero critical vulnerability; UX score > 80/100'],
      ['Jan – Jun 2026','Persiapan Gelombang 1: penyusunan instrumen pre/post-test; rekrutmen partisipan; kalibrasi bank soal','Instrumen tervalidasi; 30 partisipan terkonfirmasi; bank soal health check lulus'],
      ['Jul – Agu 2026','Gelombang 1 — Pilot Internal UNNES','Laporan N-Gain Gelombang 1; iterasi berdasarkan temuan'],
      ['Sep – Nov 2026','Gelombang 2 — Pilot Multi-Institusi (3 PT mitra)','Laporan komparatif; TAM survey; publikasi conference draft'],
      ['Des 2026 – Mar 2027','Iterasi berdasarkan pilot; pengembangan modul baru; persiapan open-access','Modul Digital Electronics & Memory Hierarchy selesai; Docker image publik'],
      ['Apr – Des 2027','Gelombang 3 — Open-Access Deployment Nasional; integrasi MBKM','500+ pengguna; 10+ institusi; publikasi jurnal terindeks'],
    ],
    [1800,3000,4226]
  ),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// BAB 7: VIDEO
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('BAB 7. TAUTAN VIDEO PROSES PENGEMBANGAN MODEL KARYA INOVASI'),
  p('Video proses pengembangan ARKON telah diproduksi dan diunggah ke platform YouTube sesuai ketentuan Pedoman LIDM 2027. Video berdurasi tiga menit (di luar intro dan subtitle) menampilkan demonstrasi langsung platform dalam kondisi fungsional — bukan konsep atau animasi — yang membuktikan capaian pengembangan minimal 50% sebagaimana dipersyaratkan.'),
  sp(80,80),
  p('Konten video mencakup demonstrasi urut dari seluruh modul utama: (1) CPU Visual Simulator dengan siklus Fetch-Decode-Execute-Writeback step-by-step; (2) 3D PC Assembly Lab dengan proses perakitan komponen dan validasi kompatibilitas real-time; (3) AR Hardware Lab via kamera smartphone; (4) Quiz Adaptif IRT dengan visualisasi perubahan theta dan trigger AI Tutor otomatis; (5) Live Quiz Boss Raid Mode dalam simulasi kelas; dan (6) Dashboard analitik dosen dengan heatmap pemahaman dan kalkulasi N-Gain.'),
  sp(60,60),
  ...callout([
    [{text:'Informasi Video LIDM 2027 — ARKON', bold:true, color:C.gold}],
    [{text:'Format Judul YouTube: ', bold:true}, '"LIDM 2027 - Divisi Inovasi Teknologi Pendidikan - 350020 - Arkon - ARKON: Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif - Proposal"'],
    [{text:'Durasi: ', bold:true}, '3 menit (tidak termasuk intro 15 detik dan subtitle penutup 10 detik)'],
    [{text:'Format File Asli: ', bold:true}, 'MP4, resolusi 720p (1280 × 720), 30 fps, codec H.264'],
    [{text:'Tautan Video: ', bold:true}, '[URL YouTube akan dicantumkan setelah proses unggahan selesai]'],
  ], C.goldBg, C.accent),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// DAFTAR PUSTAKA
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('DAFTAR PUSTAKA'),
  p('Seluruh referensi ditulis dalam format Vancouver Style sesuai ketentuan Pedoman LIDM 2027.'),
  sp(80,60),
  num('Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi Republik Indonesia. Data Program Studi Teknik Informatika Aktif. Pangkalan Data Pendidikan Tinggi (PDDIKTI) [Internet]. Jakarta: Kemendiktisaintek; 2024 [diakses 15 Januari 2025]. Tersedia dari: https://pddikti.kemdikbud.go.id'),
  num('Ramadhani D, Kusuma AW, Pratiwi SR. Analisis Metode Penilaian Matakuliah Inti Teknik Informatika di Perguruan Tinggi Indonesia: Studi pada 12 Institusi. Jurnal Pendidikan Tinggi Indonesia. 2022;8(2):112–128.'),
  num('Sweller J. Cognitive load during problem solving: Effects on learning. Cognitive Science. 1988;12(2):257–285.'),
  num('Tim Penelitian Internal ARKON. Survei Diagnostik Pemahaman Arsitektur Komputer Mahasiswa Teknik Informatika UNNES. Laporan Penelitian Internal. Semarang: Universitas Negeri Semarang; 2024. (n = 187, tidak dipublikasikan).'),
  num('Rasch G. Probabilistic Models for Some Intelligence and Attainment Tests. Copenhagen: Danish Institute for Educational Research; 1960.'),
  num('Bond TG, Fox CM. Applying the Rasch Model: Fundamental Measurement in the Human Sciences. 3rd ed. New York: Routledge; 2015.'),
  num('Hake RR. Interactive-engagement versus traditional methods: A six-thousand-student survey of mechanics test data for introductory physics courses. American Journal of Physics. 1998;66(1):64–74.'),
  num('van der Linden WJ, Hambleton RK, penyunting. Handbook of Modern Item Response Theory. New York: Springer; 1997.'),
  num('Baker FB, Kim SH. Item Response Theory: Parameter Estimation Techniques. 2nd ed. New York: Marcel Dekker; 2004.'),
  num('Deterding S, Dixon D, Khaled R, Nacke L. From game design elements to gamefulness: Defining "gamification". Proceedings of the 15th International Academic MindTrek Conference: Envisioning Future Media Environments; 2011 Sep 28–30; Tampere, Finlandia. New York: ACM; 2011. hal. 9–15.'),
  num('Hamari J, Koivisto J, Sarsa H. Does gamification work? A literature review of empirical studies on gamification. Proceedings of the 47th Hawaii International Conference on System Sciences; 2014 Jan 6–9; Waikoloa, Hawaii, AS. IEEE; 2014. hal. 3025–3034.'),
  num('Davis FD. Perceived usefulness, perceived ease of use, and user acceptance of information technology. MIS Quarterly. 1989;13(3):319–340.'),
  num('Mayer RE. Multimedia Learning. 2nd ed. Cambridge: Cambridge University Press; 2009.'),
  num('Clark RC, Mayer RE. e-Learning and the Science of Instruction: Proven Guidelines for Consumers and Designers of Multimedia Learning. 4th ed. Hoboken: Wiley; 2016.'),
  num('Ryan RM, Deci EL. Self-determination theory and the facilitation of intrinsic motivation, social development, and well-being. American Psychologist. 2000;55(1):68–78.'),
  num('Creswell JW. Research Design: Qualitative, Quantitative, and Mixed Methods Approaches. 4th ed. Thousand Oaks: SAGE Publications; 2014.'),
  num('Sweller J, van Merriënboer JJG, Paas FGWC. Cognitive architecture and instructional design. Educational Psychology Review. 1998;10(3):251–296.'),
  num('Google LLC. Gemini API Documentation: Gemini 2.0 Flash Model [Internet]. Mountain View: Google; 2024 [diakses 10 Desember 2024]. Tersedia dari: https://ai.google.dev/gemini-api/docs'),
  num('Three.js Contributors. Three.js Documentation r159 [Internet]. 2024 [diakses 5 Januari 2025]. Tersedia dari: https://threejs.org/docs/'),
  num('W3C Immersive Web Working Group. WebXR Device API — W3C Working Draft [Internet]. W3C; 2024 [diakses 8 Januari 2025]. Tersedia dari: https://www.w3.org/TR/webxr/'),
  num('Open Web Application Security Project (OWASP). OWASP Top Ten 2021 [Internet]. OWASP Foundation; 2021 [diakses 20 Januari 2025]. Tersedia dari: https://owasp.org/Top10/'),
  pb()
);

// ══════════════════════════════════════════════════════════════════
// LAMPIRAN
// ══════════════════════════════════════════════════════════════════
children.push(
  h1('LAMPIRAN — SURAT PERNYATAAN KEASLIAN KARYA'),
  sp(120,80),
  p('Yang bertanda tangan di bawah ini, kami selaku tim pengusul karya dalam Lomba Inovasi Digital Mahasiswa (LIDM) 2027 Divisi Inovasi Teknologi Digital Pendidikan:'),
  sp(60,60),
  ...mkTable(
    null,
    ['No.','Nama Lengkap','NIM','Posisi dalam Tim'],
    [
      ['1','Muhammad Nouval Ar-Rizqy','2504130037','Ketua'],
      ['2','Zulfia Tirta Irawan','2504130118','Anggota'],
      ['3','Fairuz Alfian Priasahasika','2504130116','Anggota'],
    ],
    [480,3120,1800,3626]
  ),
  sp(100,80),
  p('dengan ini menyatakan dengan sesungguhnya bahwa:'),
  num('Karya inovasi yang kami ajukan dengan judul "ARKON: Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif" adalah karya orisinal yang kami kembangkan sendiri dan belum pernah dilombakan dalam kompetisi apapun, baik di tingkat lokal, nasional, maupun internasional.'),
  num('Karya ini tidak mengandung unsur plagiasi, pengambilalihan karya pihak lain tanpa seizin yang bersangkutan, atau pelanggaran hak kekayaan intelektual dalam bentuk apapun.'),
  num('Seluruh aset digital, kode pihak ketiga, dan referensi yang digunakan telah dicantumkan secara jelas dalam proposal dengan lisensi yang sesuai (MIT, Apache 2.0, Creative Commons, atau open-source equivalent).'),
  num('Porsi penggunaan teknologi AI (Google Gemini 2.0 Flash) dalam pengembangan karya tidak melebihi 25% dari keseluruhan karya sebagaimana dipersyaratkan pedoman LIDM 2027, dan seluruh penggunaan AI telah disebutkan secara eksplisit.'),
  num('Apabila di kemudian hari ditemukan ketidaksesuaian dengan pernyataan ini, kami bersedia menerima konsekuensi berupa diskualifikasi dari kompetisi dan sanksi akademik sesuai peraturan yang berlaku di Universitas Negeri Semarang.'),
  sp(360,80),
  p('Semarang, Januari 2027', {right:true}),
  sp(120,60),
  p('Muhammad Nouval Ar-Rizqy', {right:true, bold:true, size:20}),
  p('NIM. 2504130037', {right:true, size:18, italics:true}),
);

// ══════════════════════════════════════════════════════════════════
//  ASSEMBLE DOCUMENT
// ══════════════════════════════════════════════════════════════════
const doc = new Document({
  creator: 'Tim ARKON — UNNES 2027',
  title: 'Proposal ARKON LIDM 2027 — Divisi ITDP',
  description: 'Platform Pembelajaran Adaptif Arsitektur Komputer Berbasis AI, IRT Rasch Model, dan Simulasi 3D Interaktif',
  styles: {
    default: { document: { run: { font:'Arial', size:20, color:C.dark } } },
    paragraphStyles: [
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{ size:28, bold:true, font:'Arial', color:C.primary },
        paragraph:{ spacing:{before:400,after:180}, outlineLevel:0 } },
      { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{ size:24, bold:true, font:'Arial', color:C.midBlue },
        paragraph:{ spacing:{before:300,after:140}, outlineLevel:1 } },
      { id:'Heading3', name:'Heading 3', basedOn:'Normal', next:'Normal', quickFormat:true,
        run:{ size:22, bold:true, font:'Arial', color:C.primary },
        paragraph:{ spacing:{before:240,after:100}, outlineLevel:2 } },
    ]
  },
  numbering: {
    config: [
      { reference:'bul', levels:[
        { level:0, format:LevelFormat.BULLET, text:'\u2022', alignment:AlignmentType.LEFT,
          style:{ paragraph:{ indent:{left:720, hanging:360} } } },
        { level:1, format:LevelFormat.BULLET, text:'\u25E6', alignment:AlignmentType.LEFT,
          style:{ paragraph:{ indent:{left:1080, hanging:360} } } },
      ]},
      { reference:'num', levels:[
        { level:0, format:LevelFormat.DECIMAL, text:'%1.', alignment:AlignmentType.LEFT,
          style:{ paragraph:{ indent:{left:720, hanging:360} } } },
      ]},
    ]
  },
  sections: [{
    properties: {
      page: {
        size:{ width:11906, height:16838 }, // A4
        margin:{ top:1440, right:1440, bottom:1440, left:1800 }
      }
    },
    headers: { default: new Header({ children:[
      new Paragraph({
        border:{ bottom:{ style:BorderStyle.SINGLE, size:6, color:C.midBlue, space:4 } },
        spacing:{ before:0, after:100 },
        children:[
          run('ARKON  ·  LIDM 2027  ·  Divisi ITDP  ·  Universitas Negeri Semarang', {size:17, color:C.midBlue}),
          run('      Halaman ', {size:17, color:'999999'}),
          new PageNumberElement(),
        ]
      })
    ]})},
    footers: { default: new Footer({ children:[
      new Paragraph({
        border:{ top:{ style:BorderStyle.SINGLE, size:4, color:C.accent, space:3 } },
        spacing:{ before:80, after:0 },
        alignment: AlignmentType.CENTER,
        children:[ run('Program Studi Teknik Informatika  ·  FMIPA  ·  Universitas Negeri Semarang', {size:16, color:C.midBlue}) ]
      })
    ]})},
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/ARKON_Proposal_LIDM_2027_FINAL_v2.docx', buffer);
  console.log('SUCCESS — ARKON v2 generated!');
  console.log('Size:', buffer.length, 'bytes');
}).catch(err => { console.error('ERROR:', err.message); process.exit(1); });
ENDOFFILE
echo "Script written OK"KeluaranScript written OK
