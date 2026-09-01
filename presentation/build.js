const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";               // 13.33 x 7.5
pres.author = "";
pres.title = "DYNATREKを使った分析の紹介";

/* ---------- design tokens ---------- */
const INK   = "13292E";
const INK2  = "45585B";
const MUTED = "7B8C8E";
const TEAL  = "15616D";
const TEALD = "0C2E33";
const TEALL = "E3EFF0";
const ORG   = "AD5717";
const ORGL  = "F8EDE2";
const LINE  = "D3DCDA";
const PANEL = "F2F6F5";
const WHITE = "FFFFFF";
const DMUT  = "9FC1C5";
const F     = "Meiryo";

const M = 0.62, W = 13.33 - M * 2, SW = 13.33, SH = 7.5;

/* ---------- helpers ---------- */
function bg(s, dark) { s.background = { color: dark ? TEALD : WHITE }; }

function foot(s, n, dark) {
  s.addText(String(n).padStart(2, "0"), {
    isTextBox: true, x: SW - 1.12, y: SH - 0.70, w: 0.6, h: 0.3, margin: 0,
    align: "right", fontFace: F, fontSize: 10, color: dark ? DMUT : MUTED, charSpacing: 1
  });
}

// role: {label, color}
function header(s, o) {
  const dark = !!o.dark;
  const rc = o.color || (dark ? DMUT : TEAL);
  s.addShape(pres.ShapeType.ellipse, {
    x: M, y: 0.64, w: 0.46, h: 0.46, fill: { color: rc },
    line: { color: rc, width: 0 }
  });
  s.addText(String(o.num).padStart(2, "0"), {
    isTextBox: true, x: M, y: 0.64, w: 0.46, h: 0.46, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 12, bold: true, color: WHITE
  });
  s.addText(o.role, {
    isTextBox: true, x: M + 0.66, y: 0.46, w: 6, h: 0.26, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: rc, charSpacing: 1.4
  });
  s.addText(o.title, {
    isTextBox: true, x: M + 0.66, y: 0.70, w: W - 0.66, h: 0.62, margin: 0,
    fontFace: F, fontSize: o.titleSize || 27, bold: true,
    color: dark ? WHITE : INK, valign: "top"
  });
  if (o.lead) s.addText(o.lead, {
    isTextBox: true, x: M + 0.66, y: 1.38, w: W - 0.66, h: 0.4, margin: 0,
    fontFace: F, fontSize: 13, color: dark ? DMUT : MUTED
  });
}

function bullets(s, items, o) {
  s.addText(items.map((t, i) => ({
    text: t, options: { bullet: true, breakLine: i !== items.length - 1 }
  })), Object.assign({
    isTextBox: true, margin: 0, fontFace: F, fontSize: 14, color: INK2,
    lineSpacing: 24, paraSpaceAfter: 10, valign: "top"
  }, o));
}

// dashed placeholder frame
function frame(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.05,
    fill: { color: o.tint || PANEL },
    line: { color: o.border || LINE, width: 1.25, dashType: "dash" }
  });
  const hasSub = !!o.sub;
  s.addText(o.label, {
    isTextBox: true, x: o.x, y: o.y + o.h / 2 - (hasSub ? 0.42 : 0.2), w: o.w, h: 0.4, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: o.size || 13, bold: true, color: o.fg || TEAL
  });
  if (hasSub) s.addText(o.sub, {
    isTextBox: true, x: o.x + 0.25, y: o.y + o.h / 2 - 0.02, w: o.w - 0.5, h: 0.75, margin: 0,
    align: "center", valign: "top", fontFace: F, fontSize: 11, color: MUTED, lineSpacing: 17
  });
}

function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.05,
    fill: { color: o.fill || PANEL }, line: { color: o.line || "F2F6F5", width: 0.75 }
  });
}

function arrow(s, x, y, w, color) {
  s.addShape(pres.ShapeType.rightArrow, {
    x: x, y: y, w: w, h: 0.26, fill: { color: color || LINE }, line: { color: color || LINE, width: 0 }
  });
}

function label(s, t, o) {
  s.addText(t, Object.assign({
    isTextBox: true, margin: 0, fontFace: F, fontSize: 10.5, bold: true, color: MUTED, charSpacing: 1.1
  }, o));
}

/* ================= SLIDE 1 ================= */
{
  const s = pres.addSlide(); bg(s, true);
  s.addText("成果発表 ／ 実際のテーマ内容", {
    isTextBox: true, x: M, y: 0.78, w: 8, h: 0.3, margin: 0,
    fontFace: F, fontSize: 12, bold: true, color: DMUT, charSpacing: 2
  });
  s.addText("専門知識がなくても、\nデータで問いに答える", {
    isTextBox: true, x: M, y: 1.28, w: 11.6, h: 1.7, margin: 0,
    fontFace: F, fontSize: 38, bold: true, color: WHITE, lineSpacing: 52
  });
  s.addText("DYNATREKを使って行った分析の紹介", {
    isTextBox: true, x: M, y: 3.08, w: 10, h: 0.36, margin: 0,
    fontFace: F, fontSize: 16, color: DMUT
  });

  const bw = 3.69, bg2 = 0.50, by = 4.15, bh = 1.0;
  const bx = [M, M + bw + bg2, M + (bw + bg2) * 2];
  const boxes = ["業務上の疑問", "DYNATREK", "答え"];
  boxes.forEach((t, i) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: bx[i], y: by, w: bw, h: bh, rectRadius: 0.05,
      fill: { color: i === 1 ? TEAL : "17383E" }, line: { color: i === 1 ? TEAL : "27484E", width: 1 }
    });
    s.addText(t, {
      isTextBox: true, x: bx[i], y: by, w: bw, h: bh, margin: 0, align: "center", valign: "middle",
      fontFace: F, fontSize: 15, bold: true, color: WHITE
    });
  });
  arrow(s, bx[0] + bw + 0.09, by + 0.37, 0.32, "43666C");
  arrow(s, bx[1] + bw + 0.09, by + 0.37, 0.32, "43666C");

  s.addText("所属 ／ 氏名　（発表日）", {
    isTextBox: true, x: M, y: 6.55, w: 6, h: 0.3, margin: 0,
    fontFace: F, fontSize: 12, color: DMUT
  });
  s.addNotes("挨拶と自己紹介のあと、この発表の主役を先に宣言する。\n「今日の主役はDYNATREKという製品でもなければ、あいち銀行の分析結果でもありません。“疑問を持ってから、データで確かめるまでの距離が短くなる”という一点です」と言い切ってから本編に入る。");
  foot(s, 1, true);
}

/* ================= SLIDE 2 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 2, role: "課 題", title: "「知りたいこと」はある。でも、データは遠い",
    lead: "業務のデータベースは、分析のためだけに作られているわけではありません。" });

  bullets(s, [
    "入出金のデータベースには、分析に使う項目以外にも多くの項目が含まれる",
    "目的のデータを取り出すには、テーブル構造の理解と、抽出・集計の手続きが必要になる",
    "結果として「気になるけれど、確かめるまでに手間がかかる」状態が生まれる"
  ], { x: M, y: 2.15, w: 6.15, h: 3.1 });

  const rx = 7.35, rw = 5.36;
  card(s, { x: rx, y: 2.05, w: rw, h: 4.35, fill: PANEL });
  s.addShape(pres.ShapeType.roundRect, {
    x: rx + 0.42, y: 2.32, w: rw - 0.84, h: 0.82, rectRadius: 0.12,
    fill: { color: WHITE }, line: { color: ORG, width: 1.25 }
  });
  s.addText("IBって、実際に使われるように\nなったのかな？", {
    isTextBox: true, x: rx + 0.42, y: 2.32, w: rw - 0.84, h: 0.82, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: ORG, lineSpacing: 18
  });
  s.addShape(pres.ShapeType.downArrow, {
    x: rx + rw / 2 - 0.13, y: 3.28, w: 0.26, h: 0.42, fill: { color: "C3CECC" }, line: { color: "C3CECC", width: 0 }
  });
  s.addShape(pres.ShapeType.rect, {
    x: rx + 0.42, y: 3.82, w: rw - 0.84, h: 0.62,
    fill: { color: "DDE4E2" }, line: { color: "C9D3D1", width: 1 }
  });
  s.addText("テーブル構造の理解 ／ 抽出・集計の手続き", {
    isTextBox: true, x: rx + 0.42, y: 3.82, w: rw - 0.84, h: 0.62, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 12, bold: true, color: INK2
  });
  s.addShape(pres.ShapeType.downArrow, {
    x: rx + rw / 2 - 0.13, y: 4.56, w: 0.26, h: 0.36, fill: { color: "C3CECC" }, line: { color: "C3CECC", width: 0 }
  });
  s.addShape(pres.ShapeType.rect, {
    x: rx + 0.42, y: 5.04, w: rw - 0.84, h: 1.06,
    fill: { color: WHITE }, line: { color: LINE, width: 1 }
  });
  s.addText("入出金データベース\n（分析に使わない項目も多数）", {
    isTextBox: true, x: rx + 0.42, y: 5.04, w: rw - 0.84, h: 1.06, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 12.5, color: INK2, lineSpacing: 19
  });
  s.addNotes("自分の立ち位置をここで正直に置く。\n「大学の研究でデータを扱った経験はありますが、銀行の業務も、このデータベースの構造も知らない状態から始めました」\n※“分析ができない人”ではなく“業務とDBの前提知識がない人”という設定にしておくと、後半の主張と矛盾しない。");
  foot(s, 2, false);
}

/* ================= SLIDE 3 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 3, role: "使 っ た ツ ー ル", title: "DYNATREK ── 条件を指定してデータを見る",
    lead: "機能の説明はしません。今日の話に必要な一点だけをお伝えします。" });

  const cw = 3.66, gap = 0.555, cy = 2.35, ch = 2.0;
  const cx = [M, M + cw + gap, M + (cw + gap) * 2];
  const steps = [
    ["条件を指定する", "期間・取引の種類などを\n画面上で指定する"],
    ["集計される", "指定した条件のデータが\nそのまま集計される"],
    ["表・グラフが出る", "結果が表やグラフとして\n出力される"]
  ];
  steps.forEach((st, i) => {
    card(s, { x: cx[i], y: cy, w: cw, h: ch, fill: i === 2 ? TEALL : PANEL });
    s.addText(String(i + 1), {
      isTextBox: true, x: cx[i] + 0.3, y: cy + 0.26, w: 0.5, h: 0.34, margin: 0,
      fontFace: F, fontSize: 17, bold: true, color: i === 2 ? TEAL : "AFBEBC"
    });
    s.addText(st[0], {
      isTextBox: true, x: cx[i] + 0.3, y: cy + 0.72, w: cw - 0.6, h: 0.36, margin: 0,
      fontFace: F, fontSize: 15, bold: true, color: INK
    });
    s.addText(st[1], {
      isTextBox: true, x: cx[i] + 0.3, y: cy + 1.14, w: cw - 0.6, h: 0.7, margin: 0,
      fontFace: F, fontSize: 12, color: INK2, lineSpacing: 18
    });
    if (i < 2) arrow(s, cx[i] + cw + 0.12, cy + ch / 2 - 0.13, 0.31, "C3CECC");
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.75, w: W, h: 0.78, rectRadius: 0.06,
    fill: { color: ORGL }, line: { color: ORGL, width: 0 }
  });
  s.addText("条件を変えれば、すぐに別の切り口で見直せる　── この後の話は、すべてこの繰り返しです", {
    isTextBox: true, x: M, y: 4.75, w: W, h: 0.78, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 14, bold: true, color: ORG
  });
  s.addNotes("30秒で通過する。ただし3つ目だけは強調する。\n「大事なのは3つ目です。条件を変えるのが軽いので、“じゃあ、これはどうなんだろう”をその場で試せます。この後の話は、全部これの繰り返しです」\n※ここで先に種を明かしておくと、後半の展開が“予告どおり”に見えて効く。");
  foot(s, 3, false);
}

/* ================= SLIDE 4 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 4, role: "題 材", title: "題材 ── あいち銀行の入出金データ",
    lead: "手元にあるデータの期間と、銀行が置かれている状況が、きれいに重なっていました。" });

  // timeline
  const tx = M + 0.4, tw = W - 0.8, ty = 2.78;
  s.addShape(pres.ShapeType.rect, { x: tx, y: ty, w: tw, h: 0.035, fill: { color: "C9D3D1" }, line: { color: "C9D3D1", width: 0 } });
  s.addShape(pres.ShapeType.ellipse, { x: tx - 0.09, y: ty - 0.075, w: 0.19, h: 0.19, fill: { color: TEAL }, line: { color: TEAL, width: 0 } });
  s.addText("2025年4月　第2次中期経営計画スタート", {
    isTextBox: true, x: tx - 0.05, y: ty - 0.62, w: 6, h: 0.3, margin: 0,
    fontFace: F, fontSize: 12, bold: true, color: TEAL
  });
  s.addText("2028年3月", {
    isTextBox: true, x: tx + tw - 2.2, y: ty - 0.52, w: 2.2, h: 0.28, margin: 0,
    align: "right", fontFace: F, fontSize: 11.5, color: MUTED
  });

  const winW = 2.5;
  const wins = [[tx + 0.12, "2025年4〜6月", "計画スタート直後"], [tx + 4.9, "2026年4〜6月", "1年経過時点"]];
  wins.forEach(w => {
    s.addShape(pres.ShapeType.roundRect, {
      x: w[0], y: ty + 0.28, w: winW, h: 0.92, rectRadius: 0.06,
      fill: { color: TEALL }, line: { color: "BBD5D7", width: 1 }
    });
    s.addText(w[1], {
      isTextBox: true, x: w[0], y: ty + 0.4, w: winW, h: 0.32, margin: 0,
      align: "center", fontFace: F, fontSize: 14, bold: true, color: TEAL
    });
    s.addText(w[2], {
      isTextBox: true, x: w[0], y: ty + 0.76, w: winW, h: 0.28, margin: 0,
      align: "center", fontFace: F, fontSize: 11, color: MUTED
    });
  });
  s.addText("← 使用したデータの期間 →", {
    isTextBox: true, x: tx + 0.12, y: ty + 1.28, w: winW * 2 + 2.28, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 11, color: MUTED
  });

  bullets(s, [
    "使用データ：あいち銀行の入出金データ（2025年4〜6月／2026年4〜6月）",
    "背景：あいちフィナンシャルグループは第2次中期経営計画（2025年4月〜2028年3月）で「DX戦略の加速化」を基本戦略に掲げている",
    "重点施策には、バンキングアプリやインターネット支店など非対面チャネルの活用が挙げられている",
    "この2つの期間は、計画のスタート時点と1年経過時点にあたる"
  ], { x: M, y: 4.62, w: W, h: 2.1, fontSize: 13.5 });
  s.addNotes("「データの期間を見たとき、ちょうど中期経営計画のスタート直後と、その1年後になっていることに気づきました。だとすれば、施策の手応えがこのデータに出ているはずです」\n\n※中期経営計画の「重点施策⑦」という番号は未確認。公式PDFで採番を確認できない場合は「重点施策のひとつ」と言うこと。");
  foot(s, 4, false);
}

/* ================= SLIDE 5 ================= */
{
  const s = pres.addSlide(); bg(s, true);
  header(s, { num: 5, role: "問 い ①", title: "IBの利用は、本当に広がっているのか",
    lead: "施策が動いているなら、その手応えはデータに出ているはずです。", dark: true, color: "D98A45", titleSize: 30 });

  bullets(s, [
    "非対面チャネルを進めているなら、インターネットバンキング（IB）の使われ方が変わっているはず",
    "2025年4〜6月と2026年4〜6月を比べれば、1年間の変化が見える",
    "同じ4〜6月どうしを比べるので、季節による差の影響を避けられる"
  ], { x: M + 0.66, y: 2.35, w: 6.45, h: 2.6, color: "D6E4E6" });

  const px = 8.35, pw = 4.35;
  card(s, { x: px, y: 2.35, w: pw, h: 2.55, fill: "17383E", line: "27484E" });
  const bys = [2.72, 3.92];
  [["2025年4〜6月", "計画スタート直後"], ["2026年4〜6月", "1年経過時点"]].forEach((t, i) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: px + 0.42, y: bys[i], w: pw - 0.84, h: 0.72, rectRadius: 0.05,
      fill: { color: "0C2E33" }, line: { color: "3A5C61", width: 1 }
    });
    s.addText(t[0], {
      isTextBox: true, x: px + 0.42, y: bys[i] + 0.08, w: pw - 0.84, h: 0.32, margin: 0,
      align: "center", fontFace: F, fontSize: 14, bold: true, color: WHITE
    });
    s.addText(t[1], {
      isTextBox: true, x: px + 0.42, y: bys[i] + 0.4, w: pw - 0.84, h: 0.26, margin: 0,
      align: "center", fontFace: F, fontSize: 10.5, color: DMUT
    });
  });
  s.addShape(pres.ShapeType.downArrow, {
    x: px + pw / 2 - 0.13, y: 3.5, w: 0.26, h: 0.34, fill: { color: "D98A45" }, line: { color: "D98A45", width: 0 }
  });
  s.addText("この1年で何が変わったか", {
    isTextBox: true, x: px, y: 4.98, w: pw, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 11.5, color: DMUT
  });
  s.addNotes("問いを声に出して読み上げる。そのうえで、なぜ同じ四半期どうしの比較にしたのかを一言添える（給与・賞与・年度替わりなど、月によって取引の量は変わるため）。\nここで比較設計に触れておくと、後半の解釈が軽くならない。");
  foot(s, 5, true);
}

/* ================= SLIDE 6 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 6, role: "準 備", title: "準備 ── 入出金テーブルをDYNATREK上に作る",
    lead: "調べ始める前に、「調べやすい形」を作るところから始めました。" });

  bullets(s, [
    "元のデータベースは項目が多く、そのままでは入出金の状況を追いにくい",
    "DYNATREK上で、入出金に関する項目を集めたテーブルを作成した",
    "以降の分析は、このテーブルに条件を指定するだけで行える"
  ], { x: M, y: 2.15, w: 5.6, h: 2.6 });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.25, w: 5.6, h: 0.85, rectRadius: 0.06,
    fill: { color: ORGL }, line: { color: ORGL, width: 0 }
  });
  s.addText("この土台を一度作れば、\n後から出た疑問は条件指定だけで答えられる", {
    isTextBox: true, x: M + 0.25, y: 5.25, w: 5.1, h: 0.85, margin: 0,
    valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: ORG, lineSpacing: 18
  });

  const lx = 6.75, lw = 2.55, rx2 = 10.15, rw2 = 2.55, dy = 2.35, dh = 3.4;
  card(s, { x: lx, y: dy, w: lw, h: dh, fill: PANEL });
  s.addText("元のデータ", {
    isTextBox: true, x: lx, y: dy + 0.24, w: lw, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 12.5, bold: true, color: INK2
  });
  s.addText("項目が多く、入出金の\n状況を追いにくい", {
    isTextBox: true, x: lx + 0.15, y: dy + 0.58, w: lw - 0.3, h: 0.55, margin: 0,
    align: "center", fontFace: F, fontSize: 10.5, color: MUTED, lineSpacing: 15
  });
  for (let r = 0; r < 6; r++) for (let c = 0; c < 3; c++) {
    s.addShape(pres.ShapeType.rect, {
      x: lx + 0.28 + c * 0.68, y: dy + 1.25 + r * 0.3, w: 0.56, h: 0.19,
      fill: { color: "D9E1DF" }, line: { color: "D9E1DF", width: 0 }
    });
  }
  arrow(s, lx + lw + 0.16, dy + dh / 2 - 0.13, 0.52, TEAL);
  s.addText("必要な項目\nを集める", {
    isTextBox: true, x: 9.28, y: 3.42, w: 0.88, h: 0.44, margin: 0,
    align: "center", fontFace: F, fontSize: 8.5, color: TEAL, lineSpacing: 12
  });

  card(s, { x: rx2, y: dy, w: rw2, h: dh, fill: TEALL, line: "BBD5D7" });
  s.addText("入出金テーブル", {
    isTextBox: true, x: rx2, y: dy + 0.24, w: rw2, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 12.5, bold: true, color: TEAL
  });
  s.addText("DYNATREK上に作成", {
    isTextBox: true, x: rx2, y: dy + 0.56, w: rw2, h: 0.28, margin: 0,
    align: "center", fontFace: F, fontSize: 10.5, color: TEAL
  });
  ["日　付", "チャネル", "取引種別", "年　代"].forEach((t, i) => {
    s.addShape(pres.ShapeType.rect, {
      x: rx2 + 0.28, y: dy + 1.05 + i * 0.5, w: rw2 - 0.56, h: 0.36,
      fill: { color: WHITE }, line: { color: "C6DCDE", width: 1 }
    });
    s.addText(t, {
      isTextBox: true, x: rx2 + 0.28, y: dy + 1.05 + i * 0.5, w: rw2 - 0.56, h: 0.36, margin: 0,
      align: "center", valign: "middle", fontFace: F, fontSize: 11.5, color: INK2
    });
  });
  s.addText("※ 項目名は実際に使ったものに\n　 差し替えてください", {
    isTextBox: true, x: rx2, y: dy + 3.05, w: rw2, h: 0.4, margin: 0,
    align: "center", fontFace: F, fontSize: 9, color: MUTED, lineSpacing: 13
  });
  s.addNotes("見た目は地味だが、この発表で一番効いている場所だと明言する。\n「一度この土台を作ってしまえば、後から出てきた疑問は、すべて条件を指定するだけで答えられるようになりました」\n※S13の締めで回収する伏線なので、ここは急がずに置く。");
  foot(s, 6, false);
}

/* ================= SLIDE 7 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 7, role: "実 際 の 画 面", title: "条件を指定する。グラフが出る。",
    lead: "実際の画面です。やっていることは、これだけです。" });

  const fw = 5.35, fy = 2.2, fh = 3.5;
  const fx1 = M, fx2 = 13.33 - M - fw;
  frame(s, { x: fx1, y: fy, w: fw, h: fh, label: "スクリーンショット①", sub: "条件設定画面\n（期間・チャネル・取引種別を指定）" });
  frame(s, { x: fx2, y: fy, w: fw, h: fh, label: "スクリーンショット②", sub: "出力されたグラフ\n（指定した条件で集計された結果）", tint: TEALL, border: "BBD5D7" });
  s.addShape(pres.ShapeType.rightArrow, {
    x: fx1 + fw + 0.28, y: fy + fh / 2 - 0.2, w: 0.7, h: 0.4,
    fill: { color: TEAL }, line: { color: TEAL, width: 0 }
  });
  s.addText("条件を指定すると", {
    isTextBox: true, x: fx1, y: fy + fh + 0.16, w: fw, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 12.5, bold: true, color: INK2
  });
  s.addText("そのまま結果が出てくる", {
    isTextBox: true, x: fx2, y: fy + fh + 0.16, w: fw, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 12.5, bold: true, color: TEAL
  });
  s.addText("注釈の吹き出しは多くても2つまで。操作手順は追わない。", {
    isTextBox: true, x: M, y: 6.5, w: W, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 10, color: "B7C2C1"
  });
  s.addNotes("1分以内で切り上げる。「左で条件を指定すると、右がそのまま出てきます」だけで足りる。操作手順を追わない。\n次の1枚から結果が始まる、と予告して抜ける。\n※「SQLは書いていません」は、実際に書いていない場合のみ使う。\n※スライド上の注意書き（グレーの1行）は、貼り込みが終わったら削除すること。");
  foot(s, 7, false);
}

/* ================= SLIDE 8 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 8, role: "確 認 ①", title: "IBの利用件数は増えていた",
    lead: "1年間で、IBで行われた取引の件数はこう変わりました。" });

  const lx = M, lw = 4.6;
  label(s, "対象：IBで行われた取引", { x: lx, y: 2.2, w: lw, h: 0.28 });
  s.addText("2025年4〜6月", { isTextBox: true, x: lx, y: 2.68, w: lw, h: 0.3, margin: 0, fontFace: F, fontSize: 12, color: MUTED });
  s.addText("◯◯件", { isTextBox: true, x: lx, y: 2.98, w: lw, h: 0.6, margin: 0, fontFace: F, fontSize: 30, bold: true, color: INK2 });
  s.addShape(pres.ShapeType.downArrow, { x: lx + 0.35, y: 3.68, w: 0.28, h: 0.38, fill: { color: "C3CECC" }, line: { color: "C3CECC", width: 0 } });
  s.addText("2026年4〜6月", { isTextBox: true, x: lx, y: 4.16, w: lw, h: 0.3, margin: 0, fontFace: F, fontSize: 12, color: MUTED });
  s.addText("◯◯件", { isTextBox: true, x: lx, y: 4.46, w: lw, h: 0.6, margin: 0, fontFace: F, fontSize: 30, bold: true, color: TEAL });

  s.addShape(pres.ShapeType.roundRect, {
    x: lx, y: 5.28, w: lw, h: 0.82, rectRadius: 0.06, fill: { color: TEALL }, line: { color: TEALL, width: 0 }
  });
  s.addText("＋◯◯％", {
    isTextBox: true, x: lx, y: 5.28, w: lw, h: 0.82, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 24, bold: true, color: TEAL
  });

  frame(s, { x: 6.15, y: 2.2, w: 6.56, h: 3.9, label: "グラフ貼付エリア",
    sub: "IB利用件数の2期間比較（棒グラフ）\n※ここで決めた軸・単位・色ルールを\nスライド10・11でもそのまま使うこと" });
  s.addText("※ ◯◯ は実測値に差し替えてください", {
    isTextBox: true, x: 6.15, y: 6.25, w: 6.56, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 10, color: "B7C2C1"
  });
  s.addNotes("数字を読み上げ、増えていることを確認する。そして必ず引きを作る。\n「増えています。……ここで発表を終えてもいいのですが、私はこの時点で、ひとつ引っかかりました」\n※ここが発表全体で一番の転換点。間を置いて次へ。");
  foot(s, 8, false);
}

/* ================= SLIDE 9 ================= */
{
  const s = pres.addSlide(); bg(s, true);
  header(s, { num: 9, role: "問 い ②", title: "増えた。では「普及した」と言えるか？",
    lead: "件数が増えたことと、IBに置き換わったことは、同じではありません。", dark: true, color: "D98A45", titleSize: 30 });

  bullets(s, [
    "IBの件数が増えても、ATMでの取引がそのまま残っていれば「増えただけ」",
    "「置き換わっている」と言うには、ATM側が減っているかも見る必要がある",
    "そこで、ATMで行われている取引のうち、IBでも同じことができる取引（振込・振替など）を数えた"
  ], { x: M + 0.66, y: 2.35, w: 6.15, h: 2.7, color: "D6E4E6" });

  const px = 7.9, pw = 4.8;
  card(s, { x: px, y: 2.35, w: pw, h: 3.1, fill: "17383E", line: "27484E" });
  const cw2 = 1.75;
  const cols = [[px + 0.45, "ATM", "IBでも行える取引\n（振込・振替 など）", "3A5C61"], [px + pw - 0.45 - cw2, "IB", "インターネット\nバンキング", "D98A45"]];
  cols.forEach(c => {
    s.addShape(pres.ShapeType.roundRect, {
      x: c[0], y: 2.78, w: cw2, h: 1.55, rectRadius: 0.06,
      fill: { color: "0C2E33" }, line: { color: c[3], width: 1.25 }
    });
    s.addText(c[1], {
      isTextBox: true, x: c[0], y: 2.94, w: cw2, h: 0.36, margin: 0,
      align: "center", fontFace: F, fontSize: 15, bold: true, color: WHITE
    });
    s.addText(c[2], {
      isTextBox: true, x: c[0] + 0.08, y: 3.34, w: cw2 - 0.16, h: 0.8, margin: 0,
      align: "center", fontFace: F, fontSize: 9.5, color: DMUT, lineSpacing: 14
    });
  });
  arrow(s, px + 0.45 + cw2 + 0.12, 3.42, 0.66, "D98A45");
  s.addText("置き換わっているなら、\nATM側は減っているはず", {
    isTextBox: true, x: px + 0.3, y: 4.5, w: pw - 0.6, h: 0.7, margin: 0,
    align: "center", fontFace: F, fontSize: 12, bold: true, color: "D98A45", lineSpacing: 18
  });
  s.addText("※ これは考え方の図です。結果ではありません。", {
    isTextBox: true, x: px, y: 5.52, w: pw, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 10, color: DMUT
  });
  s.addNotes("思いつき自体は特別ではない、と正直に言うほうが強い。\n「比較対象を置こう、という発想自体は普通のことだと思います。大きかったのは、思いついたときにすぐ試せたことでした。新しくデータをもらう必要も、誰かに依頼する必要もなく、条件を変えるだけでした」\n※この図は概念図であって結果ではない、と口頭でも明示する。");
  foot(s, 9, true);
}

/* ================= SLIDE 10 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 10, role: "確 認 ②", title: "ATM側は減っていた",
    lead: "IBでもできる取引に絞ってATMを見ると、逆向きに動いていました。" });

  frame(s, { x: M, y: 2.2, w: 6.56, h: 3.9, label: "グラフ貼付エリア",
    sub: "ATMのIB代替可能取引 件数の2期間比較\n※スライド8と同じ軸・単位・棒の太さで、\n色だけ減少方向に変える" });

  const rx = 7.55, rw = 5.16;
  label(s, "対象：ATMでの取引のうち、IBでも行えるもの", { x: rx, y: 2.2, w: rw, h: 0.28 });
  s.addText("振込 ／ 振替 　など", {
    isTextBox: true, x: rx, y: 2.5, w: rw, h: 0.3, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: INK2
  });
  s.addText("2025年4〜6月", { isTextBox: true, x: rx, y: 3.02, w: rw, h: 0.3, margin: 0, fontFace: F, fontSize: 12, color: MUTED });
  s.addText("◯◯件", { isTextBox: true, x: rx, y: 3.32, w: rw, h: 0.6, margin: 0, fontFace: F, fontSize: 30, bold: true, color: INK2 });
  s.addShape(pres.ShapeType.downArrow, { x: rx + 0.35, y: 4.02, w: 0.28, h: 0.38, fill: { color: "C3CECC" }, line: { color: "C3CECC", width: 0 } });
  s.addText("2026年4〜6月", { isTextBox: true, x: rx, y: 4.5, w: rw, h: 0.3, margin: 0, fontFace: F, fontSize: 12, color: MUTED });
  s.addText("◯◯件", { isTextBox: true, x: rx, y: 4.8, w: rw, h: 0.6, margin: 0, fontFace: F, fontSize: 30, bold: true, color: ORG });
  s.addShape(pres.ShapeType.roundRect, {
    x: rx, y: 5.55, w: rw, h: 0.72, rectRadius: 0.06, fill: { color: ORGL }, line: { color: ORGL, width: 0 }
  });
  s.addText("−◯◯％", {
    isTextBox: true, x: rx, y: 5.55, w: rw, h: 0.72, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 22, bold: true, color: ORG
  });
  s.addNotes("結果を読み上げる。ここでは解釈をまだ言わない。\n「では、この2つを並べてみます」で次へ渡すと、次の1枚が効く。\n※「IBでも行える取引」に何を含めたかは口頭でも一度言っておく。質疑でほぼ確実に聞かれる。");
  foot(s, 10, false);
}

/* ================= SLIDE 11 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 11, role: "解 釈", title: "2つを並べると、見えること",
    lead: "片方だけでは言えなかったことが、2つ並べると言えるようになります。" });

  const gw = 5.72, gy = 2.15, gh = 2.55;
  const gx1 = M, gx2 = 13.33 - M - gw;
  frame(s, { x: gx1, y: gy, w: gw, h: gh, label: "スライド8のグラフを再掲", tint: TEALL, border: "BBD5D7" });
  frame(s, { x: gx2, y: gy, w: gw, h: gh, label: "スライド10のグラフを再掲", tint: ORGL, border: "E4C8AC", fg: ORG });

  s.addText("IB　▲ 増加", {
    isTextBox: true, x: gx1, y: gy + gh + 0.14, w: gw, h: 0.36, margin: 0,
    align: "center", fontFace: F, fontSize: 15, bold: true, color: TEAL
  });
  s.addText("ATM（IBでも可能な取引）　▼ 減少", {
    isTextBox: true, x: gx2, y: gy + gh + 0.14, w: gw, h: 0.36, margin: 0,
    align: "center", fontFace: F, fontSize: 15, bold: true, color: ORG
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.28, w: W, h: 0.86, rectRadius: 0.06, fill: { color: PANEL }, line: { color: PANEL, width: 0 }
  });
  s.addText("同じ期間・同じデータで逆方向に動いている　→　「増えた」ではなく「置き換わりつつある」と読める", {
    isTextBox: true, x: M + 0.3, y: 5.28, w: W - 0.6, h: 0.86, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 15, bold: true, color: INK
  });
  s.addText("ただし件数は取引全体の増減にも影響されるため、断定ではなく傾向として捉えている", {
    isTextBox: true, x: M, y: 6.3, w: W, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 11.5, color: MUTED
  });
  s.addNotes("「ひとつの数字だけを見ていると、都合よく読めてしまいます。比較対象を置くと、言えることの範囲が変わります」\n※ここは学びに寄りすぎず1文で止める。学びの掘り下げは後続パートの担当。\n※最後の但し書きも、逃げではなく“どこまで言えるかを自分で線引きした”という姿勢として、はっきり口に出す。");
  foot(s, 11, false);
}

/* ================= SLIDE 12 ================= */
{
  const s = pres.addSlide(); bg(s, false);
  header(s, { num: 12, role: "問 い ③　→　確 認 ③", title: "では、どの年代で進んでいるのか",
    lead: "全体が動いているなら、次に気になるのは「誰が」です。", color: ORG });

  bullets(s, [
    "同じテーブルに「年代」の条件を加えるだけで確認できた",
    "新しいデータの準備も、誰かへの依頼も発生していない"
  ], { x: M, y: 2.2, w: 5.1, h: 1.3 });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 3.62, w: 5.1, h: 0.78, rectRadius: 0.06, fill: { color: ORGL }, line: { color: ORGL, width: 0 }
  });
  s.addText("結果：◯代で特に◯◯", {
    isTextBox: true, x: M, y: 3.62, w: 5.1, h: 0.78, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 17, bold: true, color: ORG
  });

  frame(s, { x: M, y: 4.62, w: 5.1, h: 1.5, label: "スクリーンショット③（任意）",
    sub: "条件設定画面の「年代」部分だけを切り出す\n──「条件を1つ足しただけ」が絵で伝わる", size: 11.5 });

  frame(s, { x: 6.35, y: 2.2, w: 6.36, h: 3.92, label: "グラフ貼付エリア",
    sub: "年代別の増減（横軸に年代をとった棒グラフ）\n※IB側・ATM側のどちらを見た結果かを必ず明記" });
  s.addText("※ 結果の記述は実際に出た内容に差し替えてください", {
    isTextBox: true, x: 6.35, y: 6.25, w: 6.36, h: 0.3, margin: 0,
    align: "center", fontFace: F, fontSize: 10, color: "B7C2C1"
  });
  s.addNotes("「ここまで、新しいデータの準備も、誰かへの依頼も発生していません。条件を1つ足しただけです」\n※3周目にしてこの台詞が出ると、スライド3で予告した「全部これの繰り返しです」がきれいに回収される。\n※結果の中身そのものより、たどり着く速さを強調する。");
  foot(s, 12, false);
}

/* ================= SLIDE 13 ================= */
{
  const s = pres.addSlide(); bg(s, true);
  header(s, { num: 13, role: "ま と め", title: "今回できたこと",
    lead: "特別なことはしていません。浮かんだ疑問を、その場で確かめ続けただけです。", dark: true, color: "5FB6BC" });

  const cw = 3.66, gap = 0.555, cy = 2.3, ch = 2.45;
  const cx = [M, M + cw + gap, M + (cw + gap) * 2];
  const loops = [
    ["IBの利用は、本当に\n広がっているのか？", "IB利用件数を2期間で比較\n→ 増えていた"],
    ["増えた。でも\n「置き換わった」のか？", "ATMのIB代替可能取引を比較\n→ 減っていた"],
    ["では、どの年代で\n進んでいるのか？", "年代の条件を足すだけで\n確認できた"]
  ];
  loops.forEach((lp, i) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: cx[i], y: cy, w: cw, h: ch, rectRadius: 0.05,
      fill: { color: "17383E" }, line: { color: "27484E", width: 1 }
    });
    s.addText("問 い " + "①②③"[i], {
      isTextBox: true, x: cx[i] + 0.28, y: cy + 0.2, w: cw - 0.56, h: 0.26, margin: 0,
      fontFace: F, fontSize: 10.5, bold: true, color: "D98A45", charSpacing: 1.2
    });
    s.addText(lp[0], {
      isTextBox: true, x: cx[i] + 0.28, y: cy + 0.5, w: cw - 0.56, h: 0.72, margin: 0,
      fontFace: F, fontSize: 13.5, bold: true, color: WHITE, lineSpacing: 20
    });
    s.addShape(pres.ShapeType.rect, {
      x: cx[i] + 0.28, y: cy + 1.32, w: cw - 0.56, h: 0.012,
      fill: { color: "3A5C61" }, line: { color: "3A5C61", width: 0 }
    });
    s.addText("確 認 " + "①②③"[i], {
      isTextBox: true, x: cx[i] + 0.28, y: cy + 1.44, w: cw - 0.56, h: 0.26, margin: 0,
      fontFace: F, fontSize: 10.5, bold: true, color: "5FB6BC", charSpacing: 1.2
    });
    s.addText(lp[1], {
      isTextBox: true, x: cx[i] + 0.28, y: cy + 1.72, w: cw - 0.56, h: 0.6, margin: 0,
      fontFace: F, fontSize: 12, color: "D6E4E6", lineSpacing: 17
    });
    if (i < 2) arrow(s, cx[i] + cw + 0.12, cy + ch / 2 - 0.13, 0.31, "43666C");
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.05, w: W, h: 0.72, rectRadius: 0.06,
    fill: { color: "17383E" }, line: { color: "17383E", width: 0 }
  });
  s.addText("この3往復のために新しく準備したのは、入出金テーブル1つだけ　／　新しいデータの準備・依頼：なし", {
    isTextBox: true, x: M, y: 5.05, w: W, h: 0.72, margin: 0,
    align: "center", valign: "middle", fontFace: F, fontSize: 13, bold: true, color: "5FB6BC"
  });
  s.addText("DYNATREKの価値は「分析ができること」よりも、疑問を持ってから確かめるまでが短いことにある", {
    isTextBox: true, x: M, y: 6.05, w: W, h: 0.4, margin: 0,
    align: "center", fontFace: F, fontSize: 15, bold: true, color: WHITE
  });
  s.addNotes("「業務のなかで“ちょっと気になる”が生まれたときに、その場で確かめられる。それが今回いちばん実感した価値です」で締める。\nそのうえで「ここからは、この取り組みを通じて私自身が学んだこと、苦労したことをお話しします」と後続パートへ橋渡しする。\n\n※この1枚に「今後の展望」は書かない。分析の続きは行えないので、将来の話をすると空手形になる。目標・達成度・課題はすべて後続パートの担当。");
  foot(s, 13, true);
}

pres.writeFile({ fileName: "/tmp/claude-0/-home-user-my-website/9d2b3804-e19b-5dbd-a9cd-4f44a2f53f65/scratchpad/DYNATREK_成果発表.pptx" })
  .then(f => console.log("wrote", f));
