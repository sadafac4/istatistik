const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Karma Yontemler",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2A30", lt1: "FFFFFF", dk2: "0B4F5C", lt2: "EAF4F4",
    accent1: "0B4F5C", accent2: "2A9D8F", accent3: "E76F51",
    accent4: "F4A261", accent5: "5B6C73", accent6: "8AB17D",
    hlink: "2A9D8F", folHlink: "5B6C73",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "10. Karma Yöntemler";
pres.author = "istatistik";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 1.9, w: 12, h: 0.5, fontSize: 20, bold: true, color: C.accent4, charSpacing: 3, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.5, w: 12, h: 1.6, fontSize: 54, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.3, w: 11, h: 1.2, fontSize: 24, color: C.background2, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.35, w: 12, h: 0.4, fontSize: 16, bold: true, color: C.accent3, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.75, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 12, color: C.accent5, align: "right" },
});
pres.defineSlideMaster({
  title: "SUMMARY_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.35, w: 12, h: 0.4, fontSize: 16, bold: true, color: C.accent4, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.75, w: 12.1, h: 0.8, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 12, color: C.background2, align: "right" },
});

// ---------- helpers ----------
async function iconPng(Icon, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Icon, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function iconCircle(slide, Icon, x, y, d, fill, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + " circle" });
  const p = d * 0.25;
  slide.addImage({ data: await iconPng(Icon, "FFFFFF"), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name, objectName: name + " icon" });
}
function card(slide, x, y, w, h, fill, name) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" },
    shadow: { type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 }, objectName: name,
  });
}
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 18, color: C.text1, valign: "top" }, opts));
}
function arrow(slide, x, y, w, color, name) {
  slide.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h: 0.45, fill: { color: color || C.accent5 }, line: { type: "none" }, objectName: name || "arrow" });
}
function content(section, kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}

function chip(slide, text, x, y, w, h, fill, fs, name) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, objectName: name || ("Etiket " + text) });
  txt(slide, text, { x: x + 0.08, y, w: w - 0.16, h, fontSize: fs || 17, bold: true, color: C.background1, align: "center", valign: "middle" });
}

const hdrT = (t, fill) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: fill }, fontSize: 20, align: "left" } });

(async () => {
  // ===== 1. Kapak =====
  pres.addSection({ title: "Temel kavramlar" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Temel kavramlar" });
  s.addText("BÖLÜM 10", { placeholder: "kicker" });
  s.addText("Karma Yöntemler", { placeholder: "title" });
  s.addText("Sayıların gücü ile hasta sesinin derinliğini tek çalışmada birleştirmek", { placeholder: "body" });
  await iconCircle(s, fa.FaChartBar, 8.9, 5.6, 1.1, C.accent2, "Nicel");
  s.addText("+", { isTextBox: true, x: 10.05, y: 5.65, w: 0.5, h: 1.0, fontSize: 40, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  await iconCircle(s, fa.FaComments, 10.6, 5.6, 1.1, C.accent3, "Nitel");
  s.addNotes("Bu bölümde nicel ve nitel yöntemleri tek bir araştırmada birleştiren karma yöntemleri anlatacağım. Önce tanım ve gerekçeler, sonra dört temel desen ve her biri için tıbbi bir örnek, en sonda da özet.");

  // ===== Gündem =====
  s = content("Temel kavramlar", "GÜNDEM", "Bu bölümde neler var?");
  const agenda = [
    ["Temel kavramlar", "Tanım, felsefi temel, gerekçeler"],
    ["Desenler", "Dört temel desen ve gösterim sistemi"],
    ["Tıbbi örnekler", "Beş klinik araştırma senaryosu"],
    ["Analiz", "Örnekleme, bütünleştirme, veri dönüştürme"],
    ["Kalite ve planlama", "Raporlama, adımlar, sık hatalar"],
    ["Pekiştirme", "Mini test ve özet"],
  ];
  for (let i = 0; i < agenda.length; i++) {
    const x = M + (i % 3) * 4.14, y = 1.85 + Math.floor(i / 3) * 2.6;
    card(s, x, y, 3.84, 2.3, C.background2, "Gündem " + (i + 1));
    txt(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: y + 0.2, w: 1.5, h: 0.8, fontSize: 40, bold: true, color: i % 2 ? C.accent3 : C.accent2, fontFace: THEME.headFontFace });
    txt(s, agenda[i][0], { x: x + 0.3, y: y + 1.0, w: 3.3, h: 0.5, fontSize: 22, bold: true, color: C.text2 });
    txt(s, agenda[i][1], { x: x + 0.3, y: y + 1.5, w: 3.3, h: 0.7, fontSize: 17, color: C.accent5 });
  }
  s.addNotes("Sunum altı bölümden oluşuyor: temel kavramlar, desenler, tıbbi örnekler, analiz ve bütünleştirme, kalite ve planlama, son olarak mini test ve özet.");

  // ===== 2. Tanım =====
  s = content("Temel kavramlar", "TANIM", "Karma yöntem nedir?");
  txt(s, [
    { text: "Aynı araştırmada ", options: {} },
    { text: "nicel", options: { bold: true, color: C.accent2 } },
    { text: " ve ", options: {} },
    { text: "nitel", options: { bold: true, color: C.accent3 } },
    { text: " verilerin toplanması, analiz edilmesi ve ", options: {} },
    { text: "bütünleştirilmesidir.", options: { bold: true } },
  ], { x: M, y: 1.9, w: 6.0, h: 1.6, fontSize: 24 });
  txt(s, [
    { text: "Tek başına ne sayılar ne de sözcükler sağlık sorunlarının tamamını açıklayabilir.", options: { bullet: true, breakLine: true } },
    { text: "Amaç: bir yöntemin zayıf yönünü diğerinin güçlü yönüyle dengelemek.", options: { bullet: true, breakLine: true } },
    { text: "Sağlık hizmetleri, hemşirelik ve halk sağlığı araştırmalarında giderek yaygınlaşıyor.", options: { bullet: true } },
  ], { x: M, y: 3.7, w: 6.0, h: 3.0, fontSize: 20, paraSpaceAfter: 12 });
  // Venn
  s.addShape(pres.shapes.OVAL, { x: 7.1, y: 2.0, w: 3.8, h: 3.8, fill: { color: C.accent2, transparency: 25 }, line: { type: "none" }, objectName: "Nicel daire" });
  s.addShape(pres.shapes.OVAL, { x: 9.1, y: 2.0, w: 3.8, h: 3.8, fill: { color: C.accent3, transparency: 25 }, line: { type: "none" }, objectName: "Nitel daire" });
  txt(s, [{ text: "NİCEL", options: { bold: true, fontSize: 22, breakLine: true } }, { text: "Ne kadar?", options: { breakLine: true } }, { text: "Ne sıklıkta?", options: {} }],
    { x: 7.35, y: 3.0, w: 1.75, h: 1.8, color: C.background1, fontSize: 17, valign: "middle" });
  txt(s, [{ text: "NİTEL", options: { bold: true, fontSize: 22, breakLine: true } }, { text: "Neden?", options: { breakLine: true } }, { text: "Nasıl?", options: {} }],
    { x: 11.0, y: 3.0, w: 1.7, h: 1.8, color: C.background1, fontSize: 17, valign: "middle", align: "right" });
  txt(s, "KARMA", { x: 9.1, y: 3.65, w: 1.8, h: 0.5, color: C.background1, fontSize: 22, bold: true, align: "center", valign: "middle" });
  txt(s, "Kaynak: Creswell & Plano Clark, Designing and Conducting Mixed Methods Research", { x: 7.1, y: 6.2, w: 5.8, h: 0.4, fontSize: 12, color: C.accent5, italic: true, align: "center" });
  s.addNotes("Karma yöntem; nicel ve nitel verilerin tek çalışmada toplanması, analizi ve en önemlisi bütünleştirilmesidir. Nicel yöntem 'ne kadar' sorusunu, nitel yöntem 'neden ve nasıl' sorusunu yanıtlar. Kesişim noktası karma yöntemdir.");

  // ===== Felsefi temel =====
  s = content("Temel kavramlar", "FELSEFİ TEMEL", "Pragmatizm: soruyu yanıtlayan yöntem doğrudur");
  const para = [
    ["Post-pozitivizm", "Nicel araştırmanın temeli", [["Gerçeklik", "Tek ve ölçülebilir"], ["Yöntem", "Hipotez testi, ölçme"], ["Odak", "Genelleme, neden-sonuç"]], false],
    ["Pragmatizm", "Karma yöntemin temeli", [["Gerçeklik", "Soruya göre tek ya da çoklu"], ["Yöntem", "Soruya en uygun yöntemlerin birleşimi"], ["Odak", "Uygulamaya dönük, işe yarayan bilgi"]], true],
    ["Yapılandırmacılık", "Nitel araştırmanın temeli", [["Gerçeklik", "Çoklu, kişiye göre kurulur"], ["Yöntem", "Görüşme, gözlem, yorumlama"], ["Odak", "Anlam ve deneyim"]], false],
  ];
  for (let i = 0; i < 3; i++) {
    const [h, sub, rowsP, hi] = para[i];
    const x = M + i * 4.13, y = 1.85;
    card(s, x, y, 3.84, 4.4, hi ? C.text2 : C.background2, "Paradigma " + h);
    txt(s, h, { x: x + 0.3, y: y + 0.25, w: 3.3, h: 0.5, fontSize: 24, bold: true, color: hi ? C.background1 : C.text2 });
    txt(s, sub, { x: x + 0.3, y: y + 0.75, w: 3.3, h: 0.4, fontSize: 16, italic: true, color: hi ? C.accent4 : C.accent5 });
    const runs = [];
    rowsP.forEach(([l, v], k) => {
      runs.push({ text: l, options: { bold: true, fontSize: 16, color: hi ? C.accent4 : (i === 0 ? C.accent2 : C.accent3), breakLine: true } });
      runs.push({ text: v, options: { fontSize: 19, color: hi ? C.background1 : C.text1, breakLine: k < 2, paraSpaceAfter: 12 } });
    });
    txt(s, runs, { x: x + 0.3, y: y + 1.35, w: 3.3, h: 2.9 });
  }
  txt(s, [{ text: "Pratikte: ", options: { bold: true, color: C.accent3 } }, { text: "“Hangi felsefe doğru?” yerine “Bu klinik soruyu en iyi hangi veri yanıtlar?” diye sorulur.", options: {} }], { x: M, y: 6.45, w: 12.1, h: 0.5, fontSize: 19, valign: "middle" });
  s.addNotes("Nicel araştırma genellikle post-pozitivizme, nitel araştırma yapılandırmacılığa dayanır. Karma yöntemin felsefi temeli ise çoğunlukla pragmatizmdir: araştırma sorusu merkezdedir ve soruyu en iyi yanıtlayan yöntemler birlikte kullanılır.");

  // ===== 3. Neden =====
  s = content("Temel kavramlar", "GEREKÇELER", "Neden karma yöntem kullanırız?");
  const reasons = [
    [fa.FaCheckDouble, "Üçgenleme", "Farklı yöntemlerle aynı sonuca ulaşıp bulguyu doğrulamak"],
    [fa.FaPuzzlePiece, "Tamamlayıcılık", "Bir olgunun farklı yönlerini birlikte görmek"],
    [fa.FaSearch, "Açıklama", "Beklenmedik bir istatistiksel sonucun nedenini anlamak"],
    [fa.FaTools, "Geliştirme", "Nitel bulgulardan ölçek, anket veya müdahale tasarlamak"],
    [fa.FaExpandArrowsAlt, "Genişletme", "Araştırmanın kapsamını ve genellenebilirliğini artırmak"],
    [fa.FaUserInjured, "Hasta sesi", "Klinik sonuçlara hastanın deneyimini ve önceliklerini eklemek"],
  ];
  const cw = 3.84, ch = 2.45, gx = 0.3;
  for (let i = 0; i < reasons.length; i++) {
    const [Ic, h, d] = reasons[i];
    const x = M + (i % 3) * (cw + gx), y = 1.85 + Math.floor(i / 3) * (ch + 0.3);
    card(s, x, y, cw, ch, C.background2, "Kart " + h);
    await iconCircle(s, Ic, x + 0.3, y + 0.3, 0.75, i % 2 ? C.accent3 : C.accent2, h);
    txt(s, h, { x: x + 1.25, y: y + 0.42, w: cw - 1.45, h: 0.5, fontSize: 22, bold: true, color: C.text2, valign: "middle" });
    txt(s, d, { x: x + 0.3, y: y + 1.25, w: cw - 0.6, h: 1.05, fontSize: 18 });
  }
  s.addNotes("Karma yöntemi seçmek için bir gerekçe olmalı. Üçgenleme ile bulguyu doğrularız; tamamlayıcılık ile farklı yönleri görürüz; açıklama ile beklenmedik sonucu anlarız; geliştirme ile nitel bulgudan araç üretiriz; genişletme ile kapsamı büyütürüz; ve klinik sonuçlara hastanın sesini katarız.");

  // ===== Ne zaman =====
  s = content("Temel kavramlar", "KARAR", "Ne zaman karma yöntem seçmeli?");
  const when = [
    [fa.FaQuestionCircle, "Nicel sonuç beklenmedik, açıklanması gerekiyor", "Açıklayıcı sıralı", C.accent2],
    [fa.FaClipboardList, "Konuya ya da gruba uygun ölçme aracı yok", "Keşfedici sıralı", C.accent3],
    [fa.FaBalanceScale, "Bulguyu iki farklı kaynakla doğrulamak istiyoruz", "Yakınsayan paralel", C.text2],
    [fa.FaFlask, "Müdahalenin etkisini ve işleyişini birlikte görmek", "Gömülü", C.accent2],
    [fa.FaRoute, "Uzun soluklu bir programı aşama aşama geliştireceğiz", "Çok aşamalı", C.accent3],
  ];
  for (let i = 0; i < when.length; i++) {
    const [Ic, d, dz, col] = when[i];
    const y = 1.85 + i * 0.95;
    await iconCircle(s, Ic, M, y, 0.72, C.accent5, "Durum " + (i + 1));
    txt(s, d, { x: M + 1.0, y, w: 6.9, h: 0.72, fontSize: 20, valign: "middle" });
    arrow(s, 8.2, y + 0.14, 0.5, C.accent5, "Ok durum " + (i + 1));
    chip(s, dz, 8.9, y + 0.01, 3.8, 0.7, col, 19);
  }
  txt(s, "Tek yöntem soruyu tam yanıtlıyorsa karma yönteme gerek yoktur.", { x: M, y: 6.65, w: 11, h: 0.4, fontSize: 18, italic: true, color: C.accent5 });
  s.addNotes("Karma yöntem bir moda değil, ihtiyaçtır. Beklenmedik bir nicel sonucu açıklamak için açıklayıcı sıralı, ölçme aracı yoksa keşfedici sıralı, doğrulama için yakınsayan paralel, müdahalenin işleyişini görmek için gömülü, uzun soluklu programlar için çok aşamalı desen uygundur. Tek yöntem yetiyorsa karma yönteme gerek yoktur.");

  // ===== 4. Nicel vs Nitel =====
  s = content("Temel kavramlar", "KARŞILAŞTIRMA", "Nicel ve nitel yaklaşım yan yana");
  const hdr = (t, fill) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: fill }, fontSize: 20, align: "left" } });
  const lab = (t) => ({ text: t, options: { bold: true, color: C.text2, fill: { color: C.background2 } } });
  const rows = [
    [hdr("Özellik", C.text2), hdr("Nicel", C.accent2), hdr("Nitel", C.accent3)],
    [lab("Temel soru"), "Ne kadar? Ne sıklıkta? İlişki var mı?", "Neden? Nasıl? Nasıl deneyimleniyor?"],
    [lab("Veri"), "Ölçümler, anketler, laboratuvar değerleri", "Görüşme, odak grup, gözlem notları"],
    [lab("Örneklem"), "Büyük, olasılıklı, temsil edici", "Küçük, amaçlı, bilgi açısından zengin"],
    [lab("Analiz"), "İstatistiksel testler, modeller", "Tematik / içerik analizi"],
    [lab("Tıbbi örnek"), "Tedavi sonrası sistolik KB 12 mmHg düştü", "Hasta: “İlaç beni yoruyor, bazen atlıyorum”"],
  ];
  s.addTable(rows, { x: M, y: 1.85, w: 12.1, colW: [2.5, 4.8, 4.8], rowH: 0.78, fontSize: 18, fontFace: THEME.bodyFontFace, color: C.text1, valign: "middle", margin: [0.06, 0.15, 0.06, 0.15], border: { type: "solid", pt: 1, color: "D3E3E3" } });
  txt(s, "Karma yöntem bu iki sütunu rakip değil, ortak olarak görür.", { x: M, y: 6.75, w: 11, h: 0.4, fontSize: 18, italic: true, color: C.accent5 });
  s.addNotes("İki yaklaşımı karşılaştıralım. Nicel yaklaşım büyük örneklemde sayısal ölçüm ve istatistik kullanır; nitel yaklaşım küçük ama bilgi açısından zengin örneklemde görüşme ve tematik analiz kullanır. Örneğin kan basıncındaki 12 mmHg düşüş nicel bir bulgudur; hastanın ilacı neden atladığı ise nitel bir bulgudur.");

  // ===== 5. Desenler =====
  pres.addSection({ title: "Desenler" });
  s = content("Desenler", "TEMEL DESENLER", "Dört temel karma yöntem deseni");
  const designs = [
    ["Yakınsayan paralel", "İki veri eş zamanlı, sonra karşılaştırılır", [["NİCEL + NİTEL\naynı anda", C.accent2], ["Birleştir ve karşılaştır", C.text2], ["Ortak yorum", C.accent5]]],
    ["Açıklayıcı sıralı", "Sayılar önce, nedenler sonra", [["NİCEL", C.accent2], ["nitel", C.accent3], ["Açıklayıcı yorum", C.accent5]]],
    ["Keşfedici sıralı", "Önce keşfet, sonra ölç", [["NİTEL", C.accent3], ["Araç geliştir", C.text2], ["NİCEL test", C.accent2]]],
  ];
  const ry0 = 1.85, rh = 1.3;
  for (let i = 0; i < designs.length; i++) {
    const [n, d, boxes] = designs[i];
    const y = ry0 + i * rh;
    txt(s, [{ text: n, options: { bold: true, fontSize: 22, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 16, color: C.accent5 } }], { x: M, y: y + 0.1, w: 4.3, h: 0.95 });
    for (let j = 0; j < 3; j++) {
      const bx = 5.1 + j * 2.6;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 0.1, w: 2.1, h: 0.9, rectRadius: 0.1, fill: { color: boxes[j][1] }, line: { type: "none" }, objectName: n + " kutu " + (j + 1) });
      txt(s, boxes[j][0], { x: bx + 0.08, y: y + 0.1, w: 1.94, h: 0.9, fontSize: 17, bold: true, color: C.background1, align: "center", valign: "middle" });
      if (j < 2) arrow(s, bx + 2.15, y + 0.33, 0.4, C.accent5, n + " ok " + (j + 1));
    }
  }
  // Gömülü
  const gy = ry0 + 3 * rh;
  txt(s, [{ text: "Gömülü (iç içe)", options: { bold: true, fontSize: 22, color: C.text2, breakLine: true } }, { text: "Ana çalışmanın içine ikincil bir bileşen", options: { fontSize: 16, color: C.accent5 } }], { x: M, y: gy + 0.1, w: 4.3, h: 0.95 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.1, y: gy + 0.05, w: 7.3, h: 1.0, rectRadius: 0.1, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Gömülü dış kutu" });
  txt(s, "NİCEL ana çalışma (ör. RKÇ)", { x: 5.3, y: gy + 0.05, w: 3.8, h: 1.0, fontSize: 17, bold: true, color: C.background1, valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.5, y: gy + 0.2, w: 2.7, h: 0.7, rectRadius: 0.1, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "Gömülü iç kutu" });
  txt(s, "nitel bileşen", { x: 9.5, y: gy + 0.2, w: 2.7, h: 0.7, fontSize: 17, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, "Gösterim: BÜYÜK HARF = öncelikli yöntem, küçük harf = destekleyici yöntem; ok (→) = sıra", { x: M, y: 6.95, w: 11.4, h: 0.35, fontSize: 14, italic: true, color: C.accent5 });
  s.addNotes("Dört temel desen var. Yakınsayan paralel desende iki veri aynı anda toplanır ve sonra karşılaştırılır. Açıklayıcı sıralı desende önce nicel, sonra nicel sonucu açıklamak için nitel veri toplanır. Keşfedici sıralı desende önce nitel keşif yapılır, buna dayanarak bir araç geliştirilir ve nicel olarak test edilir. Gömülü desende ise ana çalışmanın, örneğin bir randomize kontrollü çalışmanın içine nitel bir bileşen yerleştirilir. Büyük harf öncelikli yöntemi gösterir.");

  // ===== Üç karar =====
  s = content("Desenler", "DESEN SEÇİMİ", "Deseni belirleyen üç temel karar");
  const dec = [
    [fa.FaClock, "Zamanlama", "Veriler aynı anda mı, art arda mı toplanacak?", ["Eş zamanlı", "Sıralı"]],
    [fa.FaBalanceScale, "Öncelik", "Soruyu yanıtlamada hangi yöntem daha baskın?", ["Eşit", "Nicel ağırlıklı", "Nitel ağırlıklı"]],
    [fa.FaCodeBranch, "Bütünleştirme noktası", "İki yöntem nerede buluşacak?", ["Tasarım", "Veri toplama", "Analiz", "Yorum"]],
  ];
  for (let i = 0; i < 3; i++) {
    const [Ic, h, d, opts] = dec[i];
    const x = M + i * 4.13, y = 1.85;
    card(s, x, y, 3.84, 4.75, C.background2, "Karar " + h);
    await iconCircle(s, Ic, x + 0.3, y + 0.3, 0.8, [C.accent2, C.accent3, C.text2][i], h);
    txt(s, h, { x: x + 1.25, y: y + 0.3, w: 2.45, h: 0.8, fontSize: 21, bold: true, color: C.text2, valign: "middle" });
    txt(s, d, { x: x + 0.3, y: y + 1.3, w: 3.3, h: 0.9, fontSize: 18 });
    opts.forEach((o, k) => chip(s, o, x + 0.3, y + 2.3 + k * 0.6, 3.24, 0.48, C.accent5, 17, h + " seçenek " + (k + 1)));
  }
  txt(s, "Bu üç kararın birleşimi, kullanılan deseni belirler.", { x: M, y: 6.75, w: 11, h: 0.4, fontSize: 18, italic: true, color: C.accent5 });
  s.addNotes("Bir deseni seçerken üç karar verilir. Zamanlama: veriler eş zamanlı mı, sıralı mı toplanacak? Öncelik: iki yöntem eşit mi, yoksa biri baskın mı? Bütünleştirme noktası: yöntemler tasarımda mı, veri toplamada mı, analizde mi, yorumda mı buluşacak?");

  // ===== Gösterim =====
  s = content("Desenler", "GÖSTERİM SİSTEMİ", "Desenleri kısaca yazmak: Morse gösterimi");
  const mono = (t) => ({ text: t, options: { fontFace: "Courier New", bold: true, fontSize: 21, color: C.text2, fill: { color: C.background2 } } });
  s.addTable([
    [hdrT("Gösterim", C.text2), hdrT("Anlamı", C.accent2), hdrT("Desen", C.accent3)],
    [mono("QUAN + QUAL"), "Eş zamanlı, eşit öncelik", "Yakınsayan paralel"],
    [mono("QUAN → qual"), "Önce nicel (baskın), sonra nitel", "Açıklayıcı sıralı"],
    [mono("QUAL → quan"), "Önce nitel (baskın), sonra nicel", "Keşfedici sıralı"],
    [mono("QUAN(qual)"), "Nitel bileşen nicel çalışmanın içinde", "Gömülü"],
    [mono("QUAL → QUAN → …"), "Birbirine bağlı ardışık aşamalar", "Çok aşamalı"],
  ], { x: M, y: 1.85, w: 12.1, colW: [3.6, 5.0, 3.5], rowH: 0.72, fontSize: 19, fontFace: THEME.bodyFontFace, color: C.text1, valign: "middle", margin: [0.06, 0.15, 0.06, 0.15], border: { type: "solid", pt: 1, color: "D3E3E3" } });
  const legend = [["+  eş zamanlı", C.accent2], ["→  sıralı", C.accent3], ["BÜYÜK  baskın", C.text2], ["( )  gömülü", C.accent5]];
  legend.forEach(([t, c], k) => chip(s, t, M + k * 3.08, 6.35, 2.85, 0.55, c, 18, "Açıklama " + (k + 1)));
  s.addNotes("Karma yöntem desenleri Morse'un gösterimiyle kısaca yazılır. Artı eş zamanlı, ok sıralı demektir. Büyük harf baskın yöntemi, küçük harf destekleyici yöntemi gösterir. Parantez bir yöntemin diğerinin içine gömüldüğünü belirtir.");

  // ===== 6. Örnek 1 açıklayıcı =====
  pres.addSection({ title: "Tıbbi örnekler" });
  s = content("Tıbbi örnekler", "ÖRNEK 1 · AÇIKLAYICI SIRALI DESEN", "Tip 2 diyabette ilaç uyumu neden düşük?");
  async function steps(slide, items) {
    const sw = 3.75, sg = 0.42;
    for (let i = 0; i < items.length; i++) {
      const [Ic, fill, head, body] = items[i];
      const x = M + i * (sw + sg), y = 1.85;
      card(slide, x, y, sw, 3.65, C.background2, "Aşama " + (i + 1));
      await iconCircle(slide, Ic, x + 0.3, y + 0.3, 0.8, fill, head);
      txt(slide, head, { x: x + 1.25, y: y + 0.3, w: sw - 1.4, h: 0.8, fontSize: 20, bold: true, color: C.text2, valign: "middle" });
      txt(slide, body.map((b, k) => ({ text: b, options: { bullet: true, breakLine: k < body.length - 1 } })), { x: x + 0.3, y: y + 1.3, w: sw - 0.5, h: 2.25, fontSize: 18, paraSpaceAfter: 6 });
      if (i < items.length - 1) arrow(slide, x + sw + 0.02, y + 1.6, 0.38, C.accent5, "Ok " + (i + 1));
    }
  }
  await steps(s, [
    [fa.FaChartBar, C.accent2, "1. NİCEL aşama", ["400 hastaya ilaç uyum ölçeği + HbA1c ölçümü", "Her 3 hastadan 1’inde düşük uyum", "Düşük uyum → daha yüksek HbA1c"]],
    [fa.FaComments, C.accent3, "2. nitel aşama", ["Uyumu düşük 20 hastayla derinlemesine görüşme", "Temalar: yan etki korkusu, ilaç maliyeti, “kendimi iyi hissediyorum” inancı"]],
    [fa.FaLink, C.text2, "3. Bütünleştirme", ["Sayılar sorunun büyüklüğünü, görüşmeler nedenini gösterdi", "Sonuç: hedefe yönelik hasta eğitim programı"]],
  ]);
  card(s, M, 5.85, 12.1, 0.95, C.text2, "Mesaj kutusu");
  txt(s, [{ text: "Neden karma?  ", options: { bold: true, color: C.accent4 } }, { text: "Anket “ne kadar” sorusunu yanıtladı; “neden” sorusunu yalnızca hastalarla konuşarak öğrenebildik.", options: { color: C.background1 } }], { x: M + 0.3, y: 5.85, w: 11.5, h: 0.95, fontSize: 19, valign: "middle" });
  txt(s, "Eğitim amaçlı örnek senaryo; sayılar temsilidir.", { x: M, y: 6.95, w: 8, h: 0.35, fontSize: 12, italic: true, color: C.accent5 });
  s.addNotes("Açıklayıcı sıralı desene örnek: Önce 400 diyabet hastasında uyum ölçeği ve HbA1c ile nicel veri topluyoruz ve uyumu düşük grubu belirliyoruz. Ardından bu gruptan seçtiğimiz hastalarla görüşerek nedenleri araştırıyoruz. Sayılar sorunun büyüklüğünü, görüşmeler nedenini gösteriyor ve bu bilgiyle hedefli bir eğitim programı tasarlıyoruz.");

  // ===== 7. Örnek 2 keşfedici =====
  s = content("Tıbbi örnekler", "ÖRNEK 2 · KEŞFEDİCİ SIRALI DESEN", "Yoğun bakım hemşirelerinde tükenmişlik ölçeği");
  await steps(s, [
    [fa.FaUsers, C.accent3, "1. NİTEL keşif", ["Hemşirelerle 4 odak grup görüşmesi", "Temalar: gece nöbeti, hasta kaybı, ekipman eksikliği"]],
    [fa.FaClipboardList, C.text2, "2. Araç geliştirme", ["Temalardan maddeler yazıldı", "Uzman görüşü ve pilot uygulama ile ölçek taslağı"]],
    [fa.FaChartLine, C.accent2, "3. NİCEL test", ["300 hemşirede uygulama", "Faktör analizi ile geçerlik", "Cronbach alfa ile güvenirlik"]],
  ]);
  card(s, M, 5.85, 12.1, 0.95, C.text2, "Mesaj kutusu");
  txt(s, [{ text: "Neden karma?  ", options: { bold: true, color: C.accent4 } }, { text: "Hazır ölçek yoksa maddeleri hemşirelerin kendi sözlerinden türetmek, ölçeği bu gruba gerçekten uygun kılar.", options: { color: C.background1 } }], { x: M + 0.3, y: 5.85, w: 11.5, h: 0.95, fontSize: 19, valign: "middle" });
  txt(s, "Eğitim amaçlı örnek senaryo.", { x: M, y: 6.95, w: 8, h: 0.35, fontSize: 12, italic: true, color: C.accent5 });
  s.addNotes("Keşfedici sıralı desende sıra tersine döner. Yoğun bakım hemşirelerinin tükenmişliğine özgü bir ölçek yoksa, önce odak gruplarla stres kaynaklarını keşfederiz. Bu temalardan ölçek maddeleri yazarız, uzman görüşü ve pilot uygulama yaparız. Son olarak 300 hemşirede nicel olarak geçerlik ve güvenirlik analizi yaparız.");

  // ===== 8. Örnek 3 yakınsayan =====
  s = content("Tıbbi örnekler", "ÖRNEK 3 · YAKINSAYAN PARALEL DESEN", "Aşı tereddüdü: birleşik gösterim tablosu");
  const tag = (t, fill) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: fill }, align: "center" } });
  const jrows = [
    [hdr("Tema", C.text2), hdr("Nicel bulgu (anket, n = 1000)", C.accent2), hdr("Nitel bulgu (görüşme, n = 25)", C.accent3), hdr("Sonuç", C.text2)],
    [lab("Güvenlik endişesi"), "Katılımcıların yarısına yakını yan etkiden korkuyor", "“Çok hızlı geliştirildi, uzun vadede ne olacak?”", tag("Doğruluyor", C.accent6)],
    [lab("Bilgi kaynağı"), "Sosyal medya kullanımı arttıkça tereddüt artıyor", "“Doktordan çok mesaj gruplarına güveniyorum”", tag("Genişletiyor", C.accent2)],
    [lab("Hekime güven"), "Hekime güven puanı yüksek (5 üzerinden 4)", "“Doktoruma güvenirim ama sağlık sistemine güvenmem”", tag("Çelişiyor", C.accent3)],
  ];
  s.addTable(jrows, { x: M, y: 1.85, w: 12.1, colW: [2.3, 3.9, 4.05, 1.85], rowH: [0.65, 1.1, 1.1, 1.1], fontSize: 17, fontFace: THEME.bodyFontFace, color: C.text1, valign: "middle", margin: [0.06, 0.12, 0.06, 0.12], border: { type: "solid", pt: 1, color: "D3E3E3" } });
  await iconCircle(s, fa.FaSyringe, M, 6.1, 0.7, C.accent3, "Aşı");
  txt(s, [{ text: "Çelişki bir hata değil, bir bulgudur: ", options: { bold: true, color: C.text2 } }, { text: "kişiye güven ile kuruma güven ayrı ele alınmalı.", options: {} }], { x: M + 0.95, y: 6.1, w: 11, h: 0.7, fontSize: 19, valign: "middle" });
  txt(s, "Eğitim amaçlı örnek senaryo; sayılar temsilidir.", { x: M, y: 6.95, w: 8, h: 0.35, fontSize: 12, italic: true, color: C.accent5 });
  s.addNotes("Yakınsayan paralel desende anket ve görüşmeler aynı dönemde yapılır, sonra birleşik gösterim tablosunda (joint display) yan yana konur. Güvenlik endişesinde iki veri birbirini doğruluyor; bilgi kaynağında nitel veri nicel bulguyu genişletiyor; hekime güvende ise çelişki var. Bu çelişki yeni bir bulgu: kişiye güven ile sisteme güven farklı.");

  // ===== 9. Örnek 4 gömülü =====
  s = content("Tıbbi örnekler", "ÖRNEK 4 · GÖMÜLÜ DESEN", "Diz osteoartritinde ev egzersiz programı RKÇ’si");
  card(s, M, 1.85, 7.6, 4.85, C.accent2, "RKÇ dış kutu");
  await iconCircle(s, fa.FaWalking, M + 0.3, 2.1, 0.8, C.text2, "Egzersiz");
  txt(s, "NİCEL ana çalışma: RKÇ", { x: M + 1.3, y: 2.1, w: 6, h: 0.8, fontSize: 22, bold: true, color: C.background1, valign: "middle" });
  txt(s, [
    { text: "120 hasta: egzersiz ve kontrol grubu", options: { bullet: true, breakLine: true } },
    { text: "12. haftada ağrı (VAS) ve fonksiyon (WOMAC)", options: { bullet: true } },
  ], { x: M + 0.3, y: 3.1, w: 7.0, h: 1.0, fontSize: 19, color: C.background1, paraSpaceAfter: 6 });
  card(s, M + 0.3, 4.25, 7.0, 2.2, C.background1, "Nitel iç kutu");
  txt(s, [
    { text: "nitel bileşen (gömülü)", options: { bold: true, fontSize: 20, color: C.accent3, breakLine: true } },
    { text: "Programı sürdüren ve bırakan 15 hastayla görüşme", options: { bullet: true, breakLine: true } },
    { text: "Engeller: ağrı artışı korkusu, zaman yokluğu", options: { bullet: true, breakLine: true } },
    { text: "Kolaylaştırıcı: telefonla hatırlatma", options: { bullet: true } },
  ], { x: M + 0.55, y: 4.4, w: 6.6, h: 1.95, fontSize: 18, paraSpaceAfter: 4 });
  // right side
  txt(s, "Ne kazandık?", { x: 8.65, y: 1.85, w: 4.1, h: 0.6, fontSize: 24, bold: true, color: C.text2 });
  txt(s, [
    { text: "RKÇ “Etkili mi?” sorusunu yanıtlar.", options: { bullet: true, breakLine: true } },
    { text: "Nitel bileşen “Neden işe yaradı ya da yaramadı?” sorusunu yanıtlar.", options: { bullet: true, breakLine: true } },
    { text: "Nitel veri çalışmadan önce, sırasında veya sonrasında toplanabilir.", options: { bullet: true, breakLine: true } },
    { text: "Sonuçların klinik uygulamaya aktarımı kolaylaşır.", options: { bullet: true } },
  ], { x: 8.65, y: 2.55, w: 4.1, h: 4.2, fontSize: 19, paraSpaceAfter: 10 });
  s.addNotes("Gömülü desende bir randomize kontrollü çalışmanın içine nitel bir bileşen yerleştiririz. RKÇ egzersiz programının ağrı ve fonksiyon üzerindeki etkisini ölçer. Gömülü görüşmeler ise hastaların programı neden sürdürdüğünü veya bıraktığını açıklar. Böylece yalnızca 'etkili mi' değil, 'neden ve kimde işe yarıyor' sorusunu da yanıtlarız.");

  // ===== Örnek 5 çok aşamalı =====
  s = content("Tıbbi örnekler", "ÖRNEK 5 · ÇOK AŞAMALI DESEN", "Mamografi taramasına katılımı artırma programı");
  s.addShape(pres.shapes.LINE, { x: M + 1.4, y: 2.45, w: 9.3, h: 0, line: { color: C.accent5, width: 2.5 }, objectName: "Zaman çizgisi" });
  const phases = [
    [fa.FaComments, C.accent3, "Aşama 1 · NİTEL", "Kadınlarla görüşme: utanma, ağrı korkusu, ulaşım güçlüğü"],
    [fa.FaTools, C.text2, "Aşama 2 · Geliştirme", "Aile hekimi daveti, SMS hatırlatma, kadın sağlık personeli"],
    [fa.FaChartBar, C.accent2, "Aşama 3 · NİCEL", "Küme randomize çalışma: tarama katılım oranları karşılaştırılır"],
    [fa.FaSyncAlt, C.accent5, "Aşama 4 · KARMA", "Süreç değerlendirmesi: anket + görüşme, yaygınlaştırma"],
  ];
  for (let i = 0; i < 4; i++) {
    const [Ic, col, h, d] = phases[i];
    const x = M + i * 3.1;
    await iconCircle(s, Ic, x + 1.0, 2.05, 0.8, col, h);
    card(s, x, 3.15, 2.8, 2.75, C.background2, h);
    txt(s, h, { x: x + 0.2, y: 3.3, w: 2.4, h: 0.5, fontSize: 18, bold: true, color: col });
    txt(s, d, { x: x + 0.2, y: 3.85, w: 2.45, h: 1.95, fontSize: 18 });
  }
  card(s, M, 6.05, 12.1, 0.8, C.text2, "Mesaj kutusu");
  txt(s, [{ text: "Neden karma?  ", options: { bold: true, color: C.accent4 } }, { text: "Her aşamanın sorusu farklıdır ve her aşama bir öncekinin bulgusuna dayanır.", options: { color: C.background1 } }], { x: M + 0.3, y: 6.05, w: 11.5, h: 0.8, fontSize: 19, valign: "middle" });
  txt(s, "Eğitim amaçlı örnek senaryo.", { x: M, y: 6.95, w: 8, h: 0.35, fontSize: 12, italic: true, color: C.accent5 });
  s.addNotes("Çok aşamalı desen, birbirine bağlı birkaç çalışmadan oluşan ileri düzey bir desendir. Önce kadınlarla görüşerek taramaya katılım engellerini öğreniriz. Bu engellere yönelik bir müdahale geliştiririz. Müdahaleyi küme randomize bir çalışmada test ederiz. Son aşamada hem anket hem görüşme ile programın nasıl uygulandığını değerlendirip yaygınlaştırırız.");

  // ===== Örnekleme =====
  pres.addSection({ title: "Analiz ve bütünleştirme" });
  s = content("Analiz ve bütünleştirme", "ÖRNEKLEME", "Karma yöntemde örnekleme stratejileri");
  const samp = [
    [fa.FaClone, "Özdeş", "Aynı katılımcılar her iki veriyi de sağlar", "Anketi dolduran 30 hastanın hepsiyle görüşme"],
    [fa.FaBullseye, "İç içe", "Nitel örneklem, nicel örneklemin alt kümesidir", "400 hastadan uyumu en düşük 20 hasta"],
    [fa.FaColumns, "Paralel", "Farklı kişiler, ama aynı popülasyondan", "A hastanesinde anket, B hastanesinde görüşme"],
    [fa.FaSitemap, "Çok düzeyli", "Farklı düzeylerden katılımcılar", "Hastalara anket; hekim ve yöneticilerle görüşme"],
  ];
  for (let i = 0; i < 4; i++) {
    const [Ic, h, d, e] = samp[i];
    const x = M + (i % 2) * 6.2, y = 1.85 + Math.floor(i / 2) * 2.5;
    card(s, x, y, 5.9, 2.25, C.background2, "Örnekleme " + h);
    await iconCircle(s, Ic, x + 0.3, y + 0.3, 0.8, i % 2 ? C.accent3 : C.accent2, h);
    txt(s, [{ text: h, options: { bold: true, fontSize: 22, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 18 } }], { x: x + 1.3, y: y + 0.25, w: 4.4, h: 1.1 });
    txt(s, [{ text: "Örnek: ", options: { bold: true, color: C.accent3 } }, { text: e, options: {} }], { x: x + 1.3, y: y + 1.4, w: 4.4, h: 0.7, fontSize: 17 });
  }
  txt(s, "Kaynak: Teddlie & Yu, Journal of Mixed Methods Research, 2007", { x: M, y: 6.85, w: 8, h: 0.35, fontSize: 12, italic: true, color: C.accent5 });
  s.addNotes("Karma yöntemde nicel ve nitel örneklemler arasındaki ilişki önemlidir. Özdeş örneklemde aynı kişiler iki veriyi de sağlar. İç içe örneklemde nitel katılımcılar nicel örneklemden seçilir; açıklayıcı sıralı desende en sık bu kullanılır. Paralel örneklemde farklı kişiler ama aynı popülasyon vardır. Çok düzeyli örneklemde hasta, hekim, yönetici gibi farklı düzeylerden veri toplanır.");

  // ===== 10. Bütünleştirme =====
  s = content("Analiz ve bütünleştirme", "BÜTÜNLEŞTİRME", "Bütünleştirme: karma yöntemin kalbi");
  const integ = [
    [fa.FaLink, "Bağlama", "Bir aşamanın örneklemi diğer aşamadan seçilir (ör. uyumu düşük hastalar)"],
    [fa.FaHammer, "İnşa etme", "Bir aşamanın bulguları diğerinin aracını oluşturur (ör. ölçek maddeleri)"],
    [fa.FaCodeBranch, "Birleştirme", "İki veri seti analizde yan yana konur (ör. birleşik gösterim tablosu)"],
    [fa.FaLayerGroup, "Gömme", "Bir veri, diğerinin çerçevesi içinde toplanır (ör. RKÇ içinde görüşme)"],
  ];
  for (let i = 0; i < integ.length; i++) {
    const [Ic, h, d] = integ[i];
    const y = 1.85 + i * 1.2;
    await iconCircle(s, Ic, M, y, 0.85, i % 2 ? C.accent3 : C.accent2, h);
    txt(s, [{ text: h, options: { bold: true, fontSize: 21, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 18 } }], { x: M + 1.1, y: y - 0.05, w: 6.4, h: 1.1 });
  }
  card(s, 8.6, 1.85, 4.15, 4.6, C.text2, "Uyarı kutusu");
  txt(s, "“", { x: 8.9, y: 1.95, w: 1, h: 0.9, fontSize: 72, bold: true, color: C.accent4, fontFace: THEME.headFontFace });
  txt(s, "Bütünleştirme yoksa elinizde karma yöntem değil, aynı konuda yapılmış iki ayrı çalışma vardır.", { x: 8.9, y: 2.85, w: 3.6, h: 2.6, fontSize: 22, bold: true, color: C.background1 });
  txt(s, "Fetters, Curry & Creswell, Health Services Research, 2013", { x: 8.9, y: 5.6, w: 3.6, h: 0.6, fontSize: 13, italic: true, color: C.background2 });
  s.addNotes("Karma yöntemi iki ayrı çalışmadan ayıran şey bütünleştirmedir. Fetters ve arkadaşlarının tanımladığı dört yol var: bağlama, inşa etme, birleştirme ve gömme. Her biri önceki örneklerimizden birine karşılık geliyor.");

  // ===== Veri dönüştürme =====
  s = content("Analiz ve bütünleştirme", "ANALİZ", "Veri dönüştürme: sözcükten sayıya, sayıdan sözcüğe");
  const conv = [
    [fa.FaSortNumericUp, "Nicelleştirme", C.accent2, ["Görüşme metni", "Kod ve tema", "Sayı / oran"], "25 görüşmenin 15’inde (%60) “ilaç maliyeti” teması geçti; bu değişken nicel analize eklendi."],
    [fa.FaUserFriends, "Niteleştirme", C.accent3, ["Anket puanları", "Kümeleme", "Hasta profilleri"], "Puanlardan “bilgili ama kaygılı” ve “ilgisiz” gibi anlatısal hasta tipleri tanımlandı."],
  ];
  for (let i = 0; i < 2; i++) {
    const [Ic, h, col, boxes, ex] = conv[i];
    const x = M + i * 6.2;
    await iconCircle(s, Ic, x, 1.85, 0.75, col, h);
    txt(s, h, { x: x + 0.95, y: 1.85, w: 4.8, h: 0.75, fontSize: 24, bold: true, color: C.text2, valign: "middle" });
    for (let j = 0; j < 3; j++) {
      const bx = x + j * 2.05;
      chip(s, boxes[j], bx, 2.9, 1.6, 0.95, j === 2 ? col : C.accent5, 16, h + " kutu " + (j + 1));
      if (j < 2) arrow(s, bx + 1.63, 3.15, 0.38, C.accent5, h + " ok " + (j + 1));
    }
    card(s, x, 4.15, 5.9, 1.85, C.background2, h + " örnek");
    txt(s, [{ text: "Örnek: ", options: { bold: true, color: col } }, { text: ex, options: {} }], { x: x + 0.3, y: 4.3, w: 5.3, h: 1.6, fontSize: 18, valign: "middle" });
  }
  txt(s, [{ text: "Dikkat: ", options: { bold: true, color: C.accent3 } }, { text: "nitel veriyi sayıya çevirirken bağlam kaybolabilir; küçük örneklemden elde edilen oranlar genellenemez.", options: {} }], { x: M, y: 6.25, w: 12.1, h: 0.7, fontSize: 18, valign: "middle" });
  s.addNotes("Bütünleştirmenin analiz aşamasındaki yollarından biri veri dönüştürmedir. Nicelleştirmede nitel temalar kodlanıp sayılır ve istatistiksel analize eklenir. Niteleştirmede nicel veriler, örneğin kümeleme analiziyle, anlatısal hasta profillerine dönüştürülür. Ancak dönüştürme sırasında bağlam kaybolabileceğini unutmamak gerekir.");

  // ===== Kalite ve raporlama =====
  pres.addSection({ title: "Kalite ve planlama" });
  s = content("Kalite ve planlama", "KALİTE", "Raporlama ve kalite değerlendirme araçları");
  const qual = [
    [fa.FaFileAlt, "GRAMMS", "Raporlama kontrol listesi · O’Cathain ve ark., 2008", C.accent2, [
      "Karma yöntemin gerekçesi", "Desen: amaç, öncelik, sıra", "Her yöntemin örneklem, veri ve analizi",
      "Bütünleştirme nerede, nasıl, kim tarafından", "Bir yöntemin diğerine getirdiği sınırlılık", "Bütünleştirmeden doğan yeni içgörü"]],
    [fa.FaClipboardCheck, "MMAT", "Kritik değerlendirme aracı · Hong ve ark., 2018", C.accent3, [
      "Karma yöntem gerekçesi yeterli mi?", "Bileşenler etkili biçimde bütünleşmiş mi?", "Bütünleşik sonuçlar yeterince yorumlanmış mı?",
      "Nicel-nitel tutarsızlıklar ele alınmış mı?", "Her bileşen kendi kalite ölçütlerine uyuyor mu?"]],
  ];
  for (let i = 0; i < 2; i++) {
    const [Ic, h, sub, col, items] = qual[i];
    const x = M + i * 6.2;
    card(s, x, 1.85, 5.9, 5.0, C.background2, h + " kartı");
    await iconCircle(s, Ic, x + 0.3, 2.05, 0.8, col, h);
    txt(s, [{ text: h, options: { bold: true, fontSize: 24, color: C.text2, breakLine: true } }, { text: sub, options: { fontSize: 14, italic: true, color: C.accent5 } }], { x: x + 1.3, y: 2.0, w: 4.4, h: 0.95 });
    txt(s, items.map((t, k) => ({ text: t, options: { bullet: { type: "number" }, breakLine: k < items.length - 1 } })), { x: x + 0.35, y: 3.15, w: 5.3, h: 3.55, fontSize: 18, paraSpaceAfter: 8 });
  }
  s.addNotes("Karma yöntem çalışmalarının raporlanmasında GRAMMS kontrol listesi kullanılır; gerekçe, desen, her yöntemin ayrıntıları ve özellikle bütünleştirmenin nasıl yapıldığı raporlanmalıdır. Yayımlanmış bir karma yöntem çalışmasını değerlendirirken ise MMAT aracının karma yöntem soruları kullanılabilir.");

  // ===== Planlama adımları =====
  s = content("Kalite ve planlama", "PLANLAMA", "Adım adım karma yöntem çalışması");
  const plan = [
    ["Soru", "Araştırma sorusu ve karma yöntem gerekçesi"],
    ["Desen", "Zamanlama, öncelik, bütünleştirme noktası"],
    ["Örneklem", "Özdeş, iç içe, paralel ya da çok düzeyli"],
    ["Veri", "Nicel ve nitel veri toplama planı, etik izin"],
    ["Analiz", "Ayrı analizler + bütünleştirme"],
    ["Rapor", "Meta-çıkarım, GRAMMS ile raporlama"],
  ];
  const pw = 2.02;
  s.addShape(pres.shapes.LINE, { x: M + pw / 2, y: 2.4, w: pw * 5, h: 0, line: { color: C.accent5, width: 2.5 }, objectName: "Adım çizgisi" });
  for (let i = 0; i < plan.length; i++) {
    const cx = M + i * pw + pw / 2;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.45, y: 1.95, w: 0.9, h: 0.9, fill: { color: i % 2 ? C.accent3 : C.accent2 }, line: { color: C.background1, width: 3 }, objectName: "Adım " + (i + 1) });
    txt(s, String(i + 1), { x: cx - 0.45, y: 1.95, w: 0.9, h: 0.9, fontSize: 26, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, plan[i][0], { x: cx - pw / 2 + 0.1, y: 3.05, w: pw - 0.2, h: 0.5, fontSize: 22, bold: true, color: C.text2, align: "center" });
    txt(s, plan[i][1], { x: cx - pw / 2 + 0.12, y: 3.6, w: pw - 0.24, h: 1.9, fontSize: 17, align: "center" });
  }
  card(s, M, 5.75, 12.1, 0.95, C.text2, "İpucu kutusu");
  txt(s, [{ text: "İpucu:  ", options: { bold: true, color: C.accent4 } }, { text: "Desen şemasını (kutular ve oklar) çalışma protokolüne ve makaleye mutlaka ekleyin.", options: { color: C.background1 } }], { x: M + 0.3, y: 5.75, w: 11.5, h: 0.95, fontSize: 19, valign: "middle" });
  s.addNotes("Bir karma yöntem çalışması altı adımda planlanabilir: araştırma sorusu ve gerekçe, desen seçimi, örnekleme stratejisi, veri toplama ve etik izin, ayrı analizler ve bütünleştirme, son olarak meta-çıkarım ve GRAMMS ile raporlama. Meta-çıkarım, iki veri türünün birlikte yorumlanmasıyla ulaşılan bütünleşik sonuçtur.");

  // ===== Sık hatalar =====
  s = content("Kalite ve planlama", "DİKKAT", "Sık yapılan hatalar ve doğruları");
  txt(s, "Sık hata", { x: M + 0.8, y: 1.8, w: 4, h: 0.4, fontSize: 18, bold: true, color: C.accent3 });
  txt(s, "Doğrusu", { x: 7.65, y: 1.8, w: 4, h: 0.4, fontSize: 18, bold: true, color: C.accent6 });
  const errs = [
    ["“Daha zengin olsun” diye gerekçesiz karma", "Yöntem seçimini araştırma sorusuna bağlayın"],
    ["İki ayrı rapor, bütünleştirme yok", "Birleşik gösterim ve meta-çıkarım kullanın"],
    ["Göstermelik nitel bileşen (2–3 görüşme)", "Doygunluğa ulaşan amaçlı örneklem planlayın"],
    ["Desen adı ve şeması belirtilmemiş", "Deseni adlandırın, şemasını çizin"],
    ["Çelişen bulgular görmezden gelinmiş", "Çelişkiyi raporlayın ve açıklamaya çalışın"],
  ];
  const xIcon = await iconPng(fa.FaTimes, "FFFFFF"), cIcon = await iconPng(fa.FaCheck, "FFFFFF");
  for (let i = 0; i < errs.length; i++) {
    const y = 2.3 + i * 0.9;
    card(s, M, y, 5.75, 0.78, C.background2, "Hata " + (i + 1));
    s.addShape(pres.shapes.OVAL, { x: M + 0.15, y: y + 0.14, w: 0.5, h: 0.5, fill: { color: C.accent3 }, line: { type: "none" }, objectName: "Hata işareti " + (i + 1) });
    s.addImage({ data: xIcon, x: M + 0.27, y: y + 0.26, w: 0.26, h: 0.26, altText: "Hata", objectName: "Hata ikon " + (i + 1) });
    txt(s, errs[i][0], { x: M + 0.8, y, w: 4.85, h: 0.78, fontSize: 18, valign: "middle" });
    arrow(s, 6.45, y + 0.17, 0.4, C.accent5, "Ok hata " + (i + 1));
    card(s, 6.95, y, 5.75, 0.78, C.background2, "Doğru " + (i + 1));
    s.addShape(pres.shapes.OVAL, { x: 7.1, y: y + 0.14, w: 0.5, h: 0.5, fill: { color: C.accent6 }, line: { type: "none" }, objectName: "Doğru işareti " + (i + 1) });
    s.addImage({ data: cIcon, x: 7.22, y: y + 0.26, w: 0.26, h: 0.26, altText: "Doğru", objectName: "Doğru ikon " + (i + 1) });
    txt(s, errs[i][1], { x: 7.75, y, w: 4.85, h: 0.78, fontSize: 18, valign: "middle" });
  }
  s.addNotes("Karma yöntem çalışmalarında sık görülen hatalar: gerekçesiz karma yöntem kullanmak, iki veriyi hiç bütünleştirmeden ayrı ayrı raporlamak, nitel bileşeni birkaç görüşmeyle göstermelik bırakmak, deseni adlandırmamak ve çelişen bulguları görmezden gelmek. Her birinin doğrusu sağ sütunda.");

  // ===== 11. Avantaj / sınırlılık =====
  s = content("Kalite ve planlama", "DEĞERLENDİRME", "Güçlü yönler ve sınırlılıklar");
  async function col(x, Ic, fill, head, items, name) {
    card(s, x, 1.85, 5.9, 4.1, C.background2, name);
    await iconCircle(s, Ic, x + 0.35, 2.1, 0.8, fill, head);
    txt(s, head, { x: x + 1.35, y: 2.1, w: 4.3, h: 0.8, fontSize: 24, bold: true, color: C.text2, valign: "middle" });
    txt(s, items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })), { x: x + 0.35, y: 3.1, w: 5.3, h: 2.75, fontSize: 20, paraSpaceAfter: 10 });
  }
  await col(M, fa.FaThumbsUp, C.accent6, "Güçlü yönler", [
    "Daha kapsamlı ve inandırıcı kanıt",
    "İstatistiksel sonuca anlam ve bağlam katar",
    "Hasta merkezli müdahale tasarımına uygun",
    "Bir yöntemin zayıflığını diğeri dengeler",
  ], "Güçlü yönler kartı");
  await col(M + 6.2, fa.FaExclamationTriangle, C.accent3, "Sınırlılıklar", [
    "Daha fazla zaman, bütçe ve emek gerektirir",
    "Ekipte her iki yöntemde de uzmanlık gerekir",
    "Bütünleştirme ve raporlama zordur",
    "Çelişen bulguları yorumlamak beceri ister",
  ], "Sınırlılıklar kartı");
  txt(s, [{ text: "Ne zaman seçmeli?  ", options: { bold: true, color: C.accent3 } }, { text: "Araştırma sorusu tek bir yöntemle tam yanıtlanamıyorsa.", options: {} }], { x: M, y: 6.25, w: 12.1, h: 0.55, fontSize: 20, valign: "middle" });
  s.addNotes("Güçlü yönler: daha kapsamlı kanıt, sayılara bağlam, hasta merkezli tasarım. Sınırlılıklar: zaman ve maliyet, iki yöntemde de uzmanlık ihtiyacı, bütünleştirme ve çelişkili bulguları yorumlama güçlüğü.");

  // ===== Mini test =====
  pres.addSection({ title: "Pekiştirme ve özet" });
  const quiz = [
    ["Hastanede düşme sayısında beklenmedik artış saptandı; nedenini anlamak için hemşirelerle görüşme yapılıyor.", "Açıklayıcı sıralı · QUAN → qual", C.accent2],
    ["Palyatif bakım hastalarında yaşam kalitesi anketi ve aile görüşmeleri aynı dönemde yapılıp karşılaştırılıyor.", "Yakınsayan paralel · QUAN + QUAL", C.text2],
    ["Göçmen kadınlarda gebelik takibine uygun anket yok; önce görüşmeler yapılıp anket geliştiriliyor.", "Keşfedici sıralı · QUAL → quan", C.accent3],
    ["Dijital rehabilitasyon uygulamasının RKÇ’si sırasında hastaların deneyimleri görüşmelerle inceleniyor.", "Gömülü · QUAN(qual)", C.accent6],
  ];
  for (const reveal of [false, true]) {
    s = content("Pekiştirme ve özet", reveal ? "MİNİ TEST · CEVAPLAR" : "MİNİ TEST", reveal ? "Cevaplar" : "Hangi desen?");
    for (let i = 0; i < 4; i++) {
      const [sc, ans, col] = quiz[i];
      const x = M + (i % 2) * 6.2, y = 1.85 + Math.floor(i / 2) * 2.5;
      card(s, x, y, 5.9, 2.25, C.background2, "Senaryo " + "ABCD"[i]);
      s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Harf " + "ABCD"[i] });
      txt(s, "ABCD"[i], { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fontSize: 24, bold: true, color: C.background1, align: "center", valign: "middle" });
      txt(s, sc, { x: x + 1.2, y: y + 0.22, w: 4.5, h: 1.3, fontSize: 18 });
      chip(s, reveal ? ans : "?", x + 1.2, y + 1.6, 4.5, 0.48, reveal ? col : C.accent5, 17, "Cevap " + "ABCD"[i]);
    }
    s.addNotes(reveal
      ? "Cevaplar: A açıklayıcı sıralı, B yakınsayan paralel, C keşfedici sıralı, D gömülü desen."
      : "Dinleyicilere her senaryonun hangi desene uyduğunu soralım. Bir dakika düşünme süresi verip cevapları bir sonraki slaytta gösterelim.");
  }

  // ===== 12. Özet =====
  s = pres.addSlide({ masterName: "SUMMARY_DARK", sectionTitle: "Pekiştirme ve özet" });
  s.addText("ÖZET", { placeholder: "kicker" });
  s.addText("Akılda kalması gerekenler", { placeholder: "title" });
  const sum = [
    ["Tanım", "Karma yöntem = nicel + nitel + bütünleştirme"],
    ["Temel", "Pragmatizm: araştırma sorusu yöntemi belirler"],
    ["Desenler", "Yakınsayan, açıklayıcı, keşfedici, gömülü, çok aşamalı"],
    ["Tıpta", "İlaç uyumu, ölçek geliştirme, aşı tereddüdü, RKÇ, tarama"],
    ["Kalite", "GRAMMS ile raporla, MMAT ile değerlendir"],
    ["Kilit nokta", "Bütünleştirme yoksa karma yöntem de yoktur"],
  ];
  for (let i = 0; i < sum.length; i++) {
    const y = 1.75 + i * 0.84;
    s.addShape(pres.shapes.OVAL, { x: M, y, w: 0.7, h: 0.7, fill: { color: i % 2 ? C.accent3 : C.accent2 }, line: { type: "none" }, objectName: "Numara " + (i + 1) });
    txt(s, String(i + 1), { x: M, y, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, sum[i][0], { x: M + 1.0, y, w: 2.4, h: 0.7, fontSize: 22, bold: true, color: C.accent4, valign: "middle" });
    txt(s, sum[i][1], { x: M + 3.4, y, w: 8.7, h: 0.7, fontSize: 22, color: C.background1, valign: "middle" });
  }
  txt(s, "Kaynaklar: Creswell & Plano Clark (2018), Designing and Conducting Mixed Methods Research · Fetters, Curry & Creswell (2013), Health Serv Res · O’Cathain ve ark. (2008), J Health Serv Res Policy · Hong ve ark. (2018), MMAT · Teddlie & Yu (2007), J Mix Methods Res", { x: M, y: 6.8, w: 11.5, h: 0.55, fontSize: 11, italic: true, color: C.background2 });
  s.addNotes("Özetle: karma yöntem nicel ve nitel veriyi bütünleştirir ve felsefi temeli pragmatizmdir; kalitesi GRAMMS ve MMAT ile değerlendirilir; ne kadar ve neden sorularını birlikte yanıtlar; dört temel deseni vardır; tıpta ilaç uyumundan RKÇ süreç değerlendirmesine kadar geniş kullanım alanı bulur. En önemli nokta: bütünleştirme yoksa karma yöntem yoktur.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
