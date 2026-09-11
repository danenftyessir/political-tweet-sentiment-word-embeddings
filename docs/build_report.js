// Membuat 02-13523136-laporan.docx (gaya Praktikum 1). Figur = screenshot sel notebook final.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, TabStopType, LeaderType, LevelFormat, LineRuleType,
} = require("docx");

const DIR = __dirname;
const TNR = "Times New Roman";
const CW = 9026;

function body(text) {
  return new Paragraph({ alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 120, line: 276, lineRule: LineRuleType.AUTO },
    children: [new TextRun({ text, font: TNR, size: 24 })] });
}
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 160 },
    children: [new TextRun({ text, font: TNR, size: 30, bold: true, color: "000000" })] });
}
function caption(text) {
  return new Paragraph({ alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 160, line: 240, lineRule: LineRuleType.AUTO },
    children: [new TextRun({ text, font: TNR, size: 18, italics: true })] });
}
function pngSize(file) { const b = fs.readFileSync(file); return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) }; }
function figFit(file, targetIn) {
  const abs = path.join(DIR, file); const { w, h } = pngSize(abs);
  const wpx = Math.round(targetIn * 96); const hpx = Math.round(wpx * h / w);
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 }, keepNext: true,
    children: [new ImageRun({ type: "png", data: fs.readFileSync(abs), transformation: { width: wpx, height: hpx },
      altText: { title: file, description: file, name: path.basename(file) } })] });
}
function bulletLead(lead, rest) {
  return new Paragraph({ numbering: { reference: "bul", level: 0 }, alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 100, line: 276, lineRule: LineRuleType.AUTO },
    children: [ new TextRun({ text: lead, font: TNR, size: 24, bold: true }),
                new TextRun({ text: " " + rest, font: TNR, size: 24 }) ] });
}
function tline(text, opts = {}) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: opts.spacing || { after: 60 },
    children: (opts.runs || [{ text }]).map(r => new TextRun({ font: TNR, size: opts.size || 28, bold: opts.bold !== false, ...r })) });
}
function toc(text, page, bold) {
  return new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: CW, leader: LeaderType.DOT }],
    spacing: { after: 100, line: 276, lineRule: LineRuleType.AUTO }, indent: bold ? undefined : { left: 360 },
    children: [ new TextRun({ text, font: TNR, size: 24, bold: !!bold }),
                new TextRun({ text: "\t" + page, font: TNR, size: 24, bold: !!bold }) ] });
}
const bd = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const cellBorders = { top: bd, bottom: bd, left: bd, right: bd };
function cell(text, w, { header = false, align = AlignmentType.LEFT, bold = false } = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: cellBorders,
    margins: { top: 60, bottom: 60, left: 110, right: 110 }, verticalAlign: VerticalAlign.CENTER,
    shading: header ? { fill: "F2F2F2", type: ShadingType.CLEAR, color: "auto" } : undefined,
    children: [new Paragraph({ alignment: align, spacing: { after: 0, line: 252, lineRule: LineRuleType.AUTO },
      children: [new TextRun({ text, font: TNR, size: 22, bold: header || bold })] })] });
}
function makeTable(widths, rows) {
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((cells, ri) => new TableRow({ tableHeader: ri === 0,
      children: cells.map((c, ci) => cell(c.t, widths[ci], { header: ri === 0, align: c.a || AlignmentType.LEFT, bold: c.b })) })) });
}

const children = [];

// ---- Halaman judul ----
children.push(new Paragraph({ spacing: { before: 1400, after: 60 }, alignment: AlignmentType.CENTER, children: [
  new TextRun({ text: "Laporan Klasifikasi Teks dengan ", font: TNR, size: 28, bold: true }),
  new TextRun({ text: "Deep Learning", font: TNR, size: 28, bold: true, italics: true }),
  new TextRun({ text: " dan ", font: TNR, size: 28, bold: true }),
  new TextRun({ text: "Word Embedding", font: TNR, size: 28, bold: true, italics: true }) ] }));
children.push(tline("IF5153 Pemrosesan Bahasa Alami"));
children.push(tline("Praktikum 2", { spacing: { after: 120 } }));
children.push(new Paragraph({ text: "" }));
children.push(new Paragraph({ text: "" }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(DIR, "itb_logo.png")),
    transformation: { width: 200, height: 200 }, altText: { title: "Logo ITB", description: "Logo ITB", name: "LogoITB" } })] }));
children.push(tline("Oleh", { bold: false, size: 24, spacing: { after: 40 } }));
children.push(tline("Danendra Shafi Athallah", { bold: false, size: 24, spacing: { after: 40 } }));
children.push(tline("13523136", { bold: false, size: 24, spacing: { after: 120 } }));
children.push(new Paragraph({ text: "" }));
children.push(new Paragraph({ text: "" }));
children.push(new Paragraph({ text: "" }));
children.push(tline("Program Studi Teknik Informatika", { size: 24, spacing: { after: 40 } }));
children.push(tline("Sekolah Teknik Elektro dan Informatika, Institut Teknologi Bandung", { size: 24, spacing: { after: 40 } }));
children.push(tline("Jl. Ganesha 10, Bandung 40132", { size: 24, spacing: { after: 40 } }));
children.push(tline("2026/2027", { size: 24, spacing: { after: 40 } }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ---- Daftar Isi ----
children.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "Daftar Isi", font: TNR, size: 28, bold: true })] }));
children.push(toc("1. Pendahuluan", "3", true));
children.push(toc("2. Preprocessing dan Argumentasi", "3", true));
children.push(toc("3. Word Embedding", "5", true));
children.push(toc("4. Arsitektur Deep Learning", "6", true));
children.push(toc("5. Perbandingan Hasil, Dampak, dan Strategi Inferensi", "6", true));
children.push(toc("6. Insight", "9", true));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ---- 1. Pendahuluan ----
children.push(h1("1. Pendahuluan"));
children.push(body(
  "Praktikum 2 melanjutkan tugas klasifikasi sentimen tiga kelas pada dataset tweet politik India yang sama seperti Praktikum 1, dengan label 0 untuk Negatif, 1 untuk Netral, dan 2 untuk Positif. Pendekatannya kini deep learning yang dipadukan dengan word embedding. Data latih berjumlah 168.801 baris (dipecah stratified 90:10) dan data uji 72.344 baris."
));
children.push(body(
  "Eksperimen dilakukan melalui tahap preprocessing dan tokenisasi, perbandingan beberapa word embedding, pelatihan CNN/RNN/LSTM, lalu inferensi untuk submission Kaggle. Prediksi akhir memakai ensemble CNN dengan bagging K-fold dan multi-seed, penyetelan pengali kelas pada out-of-fold, serta strategi inferensi sadar-duplikat yang dibahas pada Bagian 5. Seluruh keacakan dikunci pada seed 42 dan pelatihan embedding memakai satu worker agar hasil reproducible."
));

// ---- 2. Preprocessing ----
children.push(h1("2. Preprocessing dan Argumentasi"));
children.push(body(
  "Seluruh pembersihan dirangkum dalam fungsi bersihkan_teks pada Gambar 1. Dataset ini hasil scraping yang sudah huruf kecil dan minim tanda baca, tetapi masih menyimpan artefak encoding (mojibake)."
));
children.push(figFit("img/cell_preprocess.png", 5.3));
children.push(caption("Gambar 1. Fungsi preprocessing: perbaikan encoding, ekspansi kontraksi, pemetaan emoji, dan normalisasi elongasi."));
children.push(body(
  "Langkah pertama adalah memperbaiki encoding, karena sekitar 7,8% baris mengandung mojibake. Kalau pembuangan karakter non-alfanumerik dijalankan lebih dulu, bentuk rusak seperti “isn’t” berubah menjadi token yatim “isn” dan negasinya hilang pada sekitar 2% baris. Saya memakai ftfy bila tersedia, dengan pemetaan regex sebagai cadangan. Setelah itu teks disamakan menjadi huruf kecil, kontraksi seperti “dont” dan “isnt” dikembalikan menjadi “do not” dan “is not” supaya kata “not” tetap terbaca, emoji sentimen dipetakan ke token kata sebelum simbol lain dibuang, dan huruf berulang seperti “gooood” diciutkan menjadi “good”."
));
children.push(body(
  "Baris berteks kosong tidak dibuang, hanya diisi string kosong, supaya jumlah prediksi tetap 72.344 sesuai format submission. Tokenizer di-fit hanya pada data latih, dengan num_words = 40.000 dan token “<OOV>”, agar tidak ada informasi data uji yang bocor ke pelatihan. Panjang urutan dipatok 50 token; distribusi panjang data uji tergolong pendek (median 11, persentil 99 = 39), dan pada MAX_LEN = 50 hanya 0,46% baris yang terpotong, sehingga menaikkannya lagi hampir tidak berdampak. Word2Vec dan FastText dilatih pada gabungan teks train dan test tanpa memakai labelnya selama 12 epoch, sehingga vektornya turut mencakup kosakata yang hanya muncul di data uji."
));

// ---- 3. Word Embedding ----
children.push(h1("3. Word Embedding"));
children.push(body(
  "Saya membandingkan empat teknik embedding pada arsitektur CNN yang sama, lalu menambahkan gabungan GloVe dan FastText. Gambar 2 menampilkan sel perbandingannya dan Tabel 1 merangkum akurasinya."
));
children.push(body(
  "Word2Vec dan FastText dilatih langsung pada korpus dan keduanya mencakup seluruh kosakata, sedangkan GloVe memuat vektor Twitter-200 pra-latih dan SVD dibentuk dari matriks ko-okurensi lalu direduksi dengan Truncated SVD. Word2Vec menangkap relasi semantik dari konteks lokal, dan FastText menambahkan n-gram karakter sehingga lebih tahan pada kata OOV, salah ketik, serta slang. GloVe membawa statistik ko-okurensi global yang matang, tetapi kosakatanya tetap sehingga hanya sekitar 79,5% token dataset ini yang tercakup dan sisanya menjadi vektor nol. SVD berbeda paradigma karena representasinya count-based dan statis, dengan variansi terjelaskan 95,26%. Terakhir, GloVe dan FastText saya gabungkan dengan meng-concat kedua vektor menjadi 300 dimensi untuk memadukan statistik global dan ketahanan sub-kata, dan kombinasi inilah yang memberi hasil validasi terbaik."
));
children.push(figFit("img/cell_embed_compare.png", 5.6));
children.push(caption("Gambar 2. Sel perbandingan lima representasi embedding pada CNN; keluaran menunjukkan GloVe+FastText terbaik."));
children.push(makeTable([3626, 2700, 2700], [
  [{ t: "Embedding (arsitektur CNN)" }, { t: "Val. Accuracy", a: AlignmentType.CENTER }, { t: "Val. Macro-F1", a: AlignmentType.CENTER }],
  [{ t: "GloVe + FastText", b: true }, { t: "0,8661", a: AlignmentType.CENTER, b: true }, { t: "0,8592", a: AlignmentType.CENTER, b: true }],
  [{ t: "FastText" }, { t: "0,8635", a: AlignmentType.CENTER }, { t: "0,8562", a: AlignmentType.CENTER }],
  [{ t: "GloVe (Twitter-200)" }, { t: "0,8622", a: AlignmentType.CENTER }, { t: "0,8551", a: AlignmentType.CENTER }],
  [{ t: "Word2Vec" }, { t: "0,8591", a: AlignmentType.CENTER }, { t: "0,8515", a: AlignmentType.CENTER }],
  [{ t: "SVD" }, { t: "0,4274", a: AlignmentType.CENTER }, { t: "0,1996", a: AlignmentType.CENTER }],
]));
children.push(caption("Tabel 1. Performa validasi lima representasi embedding memakai arsitektur CNN yang sama."));

// ---- 4. Arsitektur ----
children.push(h1("4. Arsitektur Deep Learning"));
children.push(body(
  "Ketiga arsitektur memakai lapisan Embedding yang diinisialisasi dari matriks pra-latih dan tetap dilatih (trainable), optimizer Adam, loss sparse categorical crossentropy, EarlyStopping, dan ReduceLROnPlateau. CNN memakai konvolusi banyak-kernel (3, 4, 5) dengan global max pooling, sedangkan RNN dan LSTM memakai lapisan dua arah (Gambar 3 menampilkan sel LSTM). Keempat embedding diuji lebih dahulu pada CNN, lalu embedding terbaik dipakai ulang untuk RNN dan LSTM agar pengaruh arsitektur terisolasi. Peringkat lengkapnya ada pada Tabel 2."
));
children.push(figFit("img/cell_lstm.png", 6.1));
children.push(caption("Gambar 3. Sel arsitektur LSTM dua arah (Bidirectional LSTM)."));
children.push(body(
  "Median panjang teks hanya 11 token, sehingga sebagian besar tweet relatif pendek. Pada eksperimen ini CNN memberikan hasil terbaik, sedangkan LSTM dan RNN berada di bawahnya. Hasil tersebut mengindikasikan model dengan convolution sudah cukup efektif untuk dataset ini, meskipun bukan berarti dependensi urutan tidak berpengaruh sama sekali."
));

// ---- 5. Perbandingan Hasil, Dampak, dan Strategi Inferensi ----
children.push(h1("5. Perbandingan Hasil, Dampak, dan Strategi Inferensi"));
children.push(body(
  "Tabel 2 merangkum keenam eksperimen (di luar SVD yang gagal), diurutkan dari akurasi validasi tertinggi."
));
children.push(makeTable([2160, 2860, 2003, 2003], [
  [{ t: "Arsitektur" }, { t: "Embedding" }, { t: "Val. Accuracy", a: AlignmentType.CENTER }, { t: "Val. Macro-F1", a: AlignmentType.CENTER }],
  [{ t: "CNN", b: true }, { t: "GloVe + FastText", b: true }, { t: "0,8661", a: AlignmentType.CENTER, b: true }, { t: "0,8592", a: AlignmentType.CENTER, b: true }],
  [{ t: "CNN" }, { t: "FastText" }, { t: "0,8635", a: AlignmentType.CENTER }, { t: "0,8562", a: AlignmentType.CENTER }],
  [{ t: "CNN" }, { t: "GloVe (Twitter-200)" }, { t: "0,8622", a: AlignmentType.CENTER }, { t: "0,8551", a: AlignmentType.CENTER }],
  [{ t: "LSTM" }, { t: "GloVe + FastText" }, { t: "0,8593", a: AlignmentType.CENTER }, { t: "0,8516", a: AlignmentType.CENTER }],
  [{ t: "CNN" }, { t: "Word2Vec" }, { t: "0,8591", a: AlignmentType.CENTER }, { t: "0,8515", a: AlignmentType.CENTER }],
  [{ t: "RNN" }, { t: "GloVe + FastText" }, { t: "0,8346", a: AlignmentType.CENTER }, { t: "0,8251", a: AlignmentType.CENTER }],
]));
children.push(caption("Tabel 2. Hasil validasi seluruh kombinasi embedding dan arsitektur (baris terbaik ditebalkan)."));
children.push(body(
  "Perbedaan ketiga embedding padat kecil, sekitar 0,004 pada accuracy, dan gabungan GloVe+FastText berada tipis di atasnya (0,8661). SVD jauh tertinggal (0,4274; macro-F1 0,1996), kemungkinan karena representasinya yang berbeda dari embedding prediktif lain. Antar-arsitektur selisihnya sedikit lebih besar: CNN (0,8661) di atas LSTM (0,8593) dan RNN (0,8346), sehingga pada eksperimen ini pilihan arsitektur terasa lebih berpengaruh daripada pilihan embedding padat."
));
children.push(body(
  "Prediksi akhir memakai ensemble CNN dari bagging StratifiedKFold dan beberapa seed, dengan pengali per-kelas yang disetel pada prediksi out-of-fold untuk mengimbangi kelas yang timpang; estimasi akurasi out-of-fold model akhir 0,8678. Selain penyetelan model, saya memanfaatkan kemiripan antara data latih dan uji yang dirangkum pada Tabel 3."
));
children.push(body(
  "Sekitar 14,92% baris uji identik persis dengan baris latih dan 15,56% identik setelah normalisasi, dengan konflik label yang kecil (78 key). Untuk baris uji yang teksnya sudah ada di data latih, saya memakai label dari data latih sebagai prediksi, dengan pencocokan exact lebih dahulu lalu normalized, mengambil suara terbanyak dan mengabaikan kasus seri. Sel inferensinya ditampilkan pada Gambar 4."
));
children.push(figFit("img/cell_infer.png", 4.4));
children.push(caption("Gambar 4. Sel inferensi final: ensemble CNN K-fold multi-seed, penyetelan pengali kelas, dan override sadar-duplikat."));
children.push(makeTable([6026, 3000], [
  [{ t: "Metrik overlap train–test (terukur)" }, { t: "Nilai", a: AlignmentType.CENTER }],
  [{ t: "Baris uji identik persis dengan train" }, { t: "10.795 (14,92%)", a: AlignmentType.CENTER }],
  [{ t: "Baris uji identik setelah normalisasi" }, { t: "11.255 (15,56%)", a: AlignmentType.CENTER }],
  [{ t: "Kesesuaian prediksi model final vs label train pada subset itu" }, { t: "±95,0%", a: AlignmentType.CENTER }],
  [{ t: "Prediksi berubah bila override diterapkan (model final)" }, { t: "565 (0,78% test)", a: AlignmentType.CENTER }],
  [{ t: "Duplikat internal train (indikasi OOF optimistik)" }, { t: "27.080 (16,04%)", a: AlignmentType.CENTER }],
]));
children.push(caption("Tabel 3. Analisis overlap train–test yang diukur langsung dari train.csv dan test.csv."));
children.push(body(
  "Model final sudah benar pada sekitar 95% baris duplikat, jadi override hanya mengubah 565 prediksi (0,78% data uji). Angka ini lebih kecil daripada versi model sebelumnya karena pipeline yang diperbaiki sudah menebak lebih banyak duplikat dengan benar. Cakupannya terukur, sedangkan besar kenaikan pada data uji tersembunyi masih berupa perkiraan."
));

// ---- 6. Insight ----
children.push(h1("6. Insight"));
children.push(bulletLead("Preprocessing.",
  "Urutan perbaikan encoding sebelum tokenisasi penting karena beberapa negasi hilang ketika teks langsung dibersihkan."));
children.push(bulletLead("Embedding.",
  "Word2Vec, GloVe, dan FastText menghasilkan performa yang relatif dekat; kombinasi GloVe dan FastText memberikan hasil validasi terbaik, sedangkan SVD jauh di bawahnya."));
children.push(bulletLead("Arsitektur.",
  "Pada dataset ini CNN memperoleh validation accuracy tertinggi dibandingkan LSTM dan RNN."));
children.push(bulletLead("Duplikasi data.",
  "Ditemukan cukup banyak teks yang sama antara train dan test (sekitar 15,56%), sehingga informasi dari data latih dapat dimanfaatkan untuk sebagian sampel uji."));
children.push(bulletLead("Validasi.",
  "Duplikat di dalam train membuat skor cross-validation perlu diinterpretasikan dengan hati-hati karena teks identik dapat muncul pada fold pelatihan dan validasi."));

const doc = new Document({
  styles: { default: { document: { run: { font: TNR, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: TNR, color: "000000" }, paragraph: { spacing: { before: 260, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: TNR, color: "000000" }, paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } },
    ] },
  numbering: { config: [{ reference: "bul",
    levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 620, hanging: 320 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }],
});
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(path.join(DIR, "02-13523136-laporan.docx"), buf);
  console.log("OK wrote 02-13523136-laporan.docx (" + buf.length + " bytes)");
});
