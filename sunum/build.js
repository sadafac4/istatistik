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

(async () => {
  // ===== 1. Kapak =====
  pres.addSection({ title: "Giriş" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Giriş" });
  s.addText("BÖLÜM 10", { placeholder: "kicker" });
  s.addText("Karma Yöntemler", { placeholder: "title" });
  s.addText("Sayıların gücü ile hasta sesinin derinliğini tek çalışmada birleştirmek", { placeholder: "body" });
  await iconCircle(s, fa.FaChartBar, 8.9, 5.6, 1.1, C.accent2, "Nicel");
  s.addText("+", { isTextBox: true, x: 10.05, y: 5.65, w: 0.5, h: 1.0, fontSize: 40, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  await iconCircle(s, fa.FaComments, 10.6, 5.6, 1.1, C.accent3, "Nitel");
  s.addNotes("Bu bölümde nicel ve nitel yöntemleri tek bir araştırmada birleştiren karma yöntemleri anlatacağım. Önce tanım ve gerekçeler, sonra dört temel desen ve her biri için tıbbi bir örnek, en sonda da özet.");

  // ===== 2. Tanım =====
  s = content("Giriş", "TANIM", "Karma yöntem nedir?");
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

  // ===== 3. Neden =====
  s = content("Giriş", "GEREKÇELER", "Neden karma yöntem kullanırız?");
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

  // ===== 4. Nicel vs Nitel =====
  s = content("Giriş", "KARŞILAŞTIRMA", "Nicel ve nitel yaklaşım yan yana");
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

  // ===== 10. Bütünleştirme =====
  pres.addSection({ title: "Değerlendirme" });
  s = content("Değerlendirme", "BÜTÜNLEŞTİRME", "Bütünleştirme: karma yöntemin kalbi");
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

  // ===== 11. Avantaj / sınırlılık =====
  s = content("Değerlendirme", "DEĞERLENDİRME", "Güçlü yönler ve sınırlılıklar");
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

  // ===== 12. Özet =====
  pres.addSection({ title: "Özet" });
  s = pres.addSlide({ masterName: "SUMMARY_DARK", sectionTitle: "Özet" });
  s.addText("ÖZET", { placeholder: "kicker" });
  s.addText("Akılda kalması gerekenler", { placeholder: "title" });
  const sum = [
    ["Tanım", "Karma yöntem = nicel + nitel + bütünleştirme"],
    ["Amaç", "“Ne kadar?” ve “Neden?” sorularını birlikte yanıtlamak"],
    ["Desenler", "Yakınsayan paralel, açıklayıcı sıralı, keşfedici sıralı, gömülü"],
    ["Tıpta", "İlaç uyumu, ölçek geliştirme, aşı tereddüdü, RKÇ süreç değerlendirmesi"],
    ["Kilit nokta", "Bütünleştirme yoksa karma yöntem de yoktur"],
  ];
  for (let i = 0; i < sum.length; i++) {
    const y = 1.85 + i * 1.0;
    s.addShape(pres.shapes.OVAL, { x: M, y, w: 0.7, h: 0.7, fill: { color: i % 2 ? C.accent3 : C.accent2 }, line: { type: "none" }, objectName: "Numara " + (i + 1) });
    txt(s, String(i + 1), { x: M, y, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, sum[i][0], { x: M + 1.0, y, w: 2.4, h: 0.7, fontSize: 22, bold: true, color: C.accent4, valign: "middle" });
    txt(s, sum[i][1], { x: M + 3.4, y, w: 8.7, h: 0.7, fontSize: 22, color: C.background1, valign: "middle" });
  }
  txt(s, "Kaynaklar: Creswell & Plano Clark (2018), Designing and Conducting Mixed Methods Research; Fetters, Curry & Creswell (2013), Health Services Research 48(6).", { x: M, y: 6.85, w: 11.5, h: 0.45, fontSize: 12, italic: true, color: C.background2 });
  s.addNotes("Özetle: karma yöntem nicel ve nitel veriyi bütünleştirir; ne kadar ve neden sorularını birlikte yanıtlar; dört temel deseni vardır; tıpta ilaç uyumundan RKÇ süreç değerlendirmesine kadar geniş kullanım alanı bulur. En önemli nokta: bütünleştirme yoksa karma yöntem yoktur.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
