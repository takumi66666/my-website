// 日立製作所 インターンシップ成果発表スライド生成スクリプト
// 実行: NODE_PATH=<pptxgenjs のあるディレクトリ>/node_modules node build_deck.js
const PptxGenJS = require("pptxgenjs");

const NAVY       = "13294B";
const NAVY_MID   = "1C3A66";
const NAVY_SOFT  = "2A4A7A";
const INK        = "1E2A38";
const MUTED      = "5D6E82";
const ICE        = "CADCFC";
const TINT       = "EEF2F8";
const ACCENT     = "D9482F";
const ACCENT_TNT = "FBEAE4";
const WHITE      = "FFFFFF";
const LINE       = "D3DCE8";
const F          = "Meiryo";

const W = 13.33, H = 7.5, M = 0.55, CW = 12.23;

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Hitachi Internship";
pres.title  = "DYNATREKで業務の疑問をデータで確かめる";

// ---------- helpers ----------
function pageNum(slide, n, dark) {
  slide.addText(String(n), {
    isTextBox: true, x: W - 0.85, y: H - 0.52, w: 0.4, h: 0.3,
    fontFace: F, fontSize: 10, color: dark ? NAVY_SOFT : "9AA9BA", align: "right", margin: 0,
  });
}

function head(slide, title, lead) {
  slide.addText(title, {
    isTextBox: true, x: M, y: 0.38, w: CW, h: 0.62,
    fontFace: F, fontSize: 26, bold: true, color: NAVY, align: "left", margin: 0, valign: "middle",
  });
  if (lead) {
    slide.addText(lead, {
      isTextBox: true, x: M, y: 1.04, w: CW, h: 0.42,
      fontFace: F, fontSize: 13.5, bold: true, color: ACCENT, align: "left", margin: 0, valign: "middle",
    });
  }
}

function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.06,
    fill: { color: o.fill || TINT },
    line: o.line ? { color: o.line, width: 1 } : { color: o.fill || TINT, width: 0.5 },
  });
}

function badge(slide, o) {
  slide.addShape(pres.ShapeType.ellipse, {
    x: o.x, y: o.y, w: 0.36, h: 0.36, fill: { color: o.color },
    line: { color: o.color, width: 0.5 },
  });
  slide.addText(o.label, {
    isTextBox: true, x: o.x, y: o.y, w: 0.36, h: 0.36,
    fontFace: F, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0,
  });
}

function body(items, opts) {
  const o = opts || {};
  return items.map((t, i) => ({
    text: t,
    options: {
      breakLine: i !== items.length - 1,
      bullet: o.bullet ? { characterCode: "25AA" } : false,
      paraSpaceAfter: o.gap === undefined ? 7 : o.gap,
    },
  }));
}

// =========================================================
// Slide 1 — 表紙
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  s.addShape(pres.ShapeType.ellipse, { x: 9.45, y: 1.15, w: 3.0, h: 3.0, fill: { color: NAVY_MID }, line: { color: NAVY_MID, width: 0.5 } });
  s.addShape(pres.ShapeType.ellipse, { x: 10.15, y: 2.62, w: 2.9, h: 2.9, fill: { color: NAVY_SOFT }, line: { color: NAVY_SOFT, width: 0.5 } });

  s.addText("日立製作所　インターンシップ　成果発表", {
    isTextBox: true, x: M, y: 0.85, w: 8.4, h: 0.35,
    fontFace: F, fontSize: 12, color: ICE, charSpacing: 1, margin: 0, valign: "middle",
  });

  s.addText("「IBはどれくらい広がったのか」に\nデータで答える", {
    isTextBox: true, x: M, y: 1.6, w: 8.9, h: 1.9,
    fontFace: F, fontSize: 33, bold: true, color: WHITE, lineSpacingMultiple: 1.25, margin: 0, valign: "middle",
  });

  s.addText("専門的なデータ分析の知識がなくても、DYNATREKなら業務の疑問を確かめられる", {
    isTextBox: true, x: M, y: 3.72, w: 9.0, h: 0.42,
    fontFace: F, fontSize: 15.5, color: ICE, margin: 0, valign: "middle",
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.35, w: 7.3, h: 0.62, rectRadius: 0.08,
    fill: { color: NAVY_MID }, line: { color: NAVY_SOFT, width: 1 },
  });
  s.addText("事例：愛知銀行　インターネットバンキング（IB）普及状況の分析", {
    isTextBox: true, x: M + 0.22, y: 4.35, w: 6.9, h: 0.62,
    fontFace: F, fontSize: 13, color: WHITE, margin: 0, valign: "middle",
  });

  s.addText("発表者：（氏名）　／　2026年●月●日", {
    isTextBox: true, x: M, y: 6.35, w: 7.0, h: 0.35,
    fontFace: F, fontSize: 12, color: ICE, margin: 0, valign: "middle",
  });

  s.addNotes(
    "本日は、2週間のインターンシップで取り組んだ内容についてご報告します。\n" +
    "タイトルにあるとおり、今回私が扱ったのは「愛知銀行のインターネットバンキング、IBがどれくらい広がっているのか」という疑問です。\n" +
    "ただ、本日一番お伝えしたいのは分析結果そのものではありません。\n" +
    "専門的なデータ分析の知識がなくても、DYNATREKを使えば業務上の疑問をデータで確かめていける、ということを、実際に自分がやってみた分析を例にお話しします。"
  );
  pageNum(s, 1, true);
}

// =========================================================
// Slide 2 — 2つの壁
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "業務の疑問をデータで確かめるには、2つの壁がある",
        "専門知識がないと越えにくいのは「データを取り出す」壁。DYNATREKはそこを下げてくれる。");

  // 左：背景
  card(s, { x: M, y: 1.72, w: 5.7, h: 2.90, fill: TINT });
  s.addText("今回のテーマと背景", {
    isTextBox: true, x: M + 0.28, y: 1.9, w: 5.2, h: 0.34,
    fontFace: F, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s.addText(body([
    "愛知銀行では、インターネットバンキング（IB）の普及に向けた施策やキャンペーンが行われている。",
    "現場の疑問は「IBは、実際にどれくらい使われるようになったのか」。",
    "この疑問に答えるため、銀行の業務データをDYNATREKで分析した。",
  ], { bullet: true, gap: 9 }), {
    isTextBox: true, x: M + 0.28, y: 2.30, w: 5.16, h: 2.20,
    fontFace: F, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.3, margin: 0, valign: "top",
  });

  card(s, { x: M, y: 4.88, w: 5.7, h: 1.62, fill: WHITE, line: LINE });
  s.addText("DYNATREKとは", {
    isTextBox: true, x: M + 0.28, y: 5.04, w: 5.2, h: 0.32,
    fontFace: F, fontSize: 14, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("業務データベースに対して、期間・取引の種類・チャネルなどの条件を指定し、必要なデータだけを表示・集計・比較できるツール。", {
    isTextBox: true, x: M + 0.28, y: 5.44, w: 5.16, h: 0.95,
    fontFace: F, fontSize: 12, color: INK, lineSpacingMultiple: 1.3, margin: 0, valign: "top",
  });

  // 右：2つの壁
  const rx = 6.65, rw = 6.13;
  card(s, { x: rx, y: 1.72, w: rw, h: 1.9, fill: WHITE, line: LINE });
  badge(s, { x: rx + 0.28, y: 1.98, color: NAVY, label: "1" });
  s.addText("壁①　何を見れば答えになるのかを決める", {
    isTextBox: true, x: rx + 0.78, y: 1.96, w: rw - 1.05, h: 0.4,
    fontFace: F, fontSize: 15, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  s.addText("「IBの普及」を、どの数字で見れば確かめたことになるのか。ここは人が考えるところで、ツールでは代われない。", {
    isTextBox: true, x: rx + 0.78, y: 2.42, w: rw - 1.05, h: 1.0,
    fontFace: F, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.3, margin: 0, valign: "top",
  });

  card(s, { x: rx, y: 3.78, w: rw, h: 1.9, fill: ACCENT_TNT });
  badge(s, { x: rx + 0.28, y: 4.04, color: ACCENT, label: "2" });
  s.addText("壁②　必要なデータを業務DBから取り出す", {
    isTextBox: true, x: rx + 0.78, y: 4.02, w: rw - 1.05, h: 0.4,
    fontFace: F, fontSize: 15, bold: true, color: ACCENT, margin: 0, valign: "middle",
  });
  s.addText("業務DBには分析に使わない項目も数多く含まれ、目的のデータを抜き出すこと自体が難しい。ここをDYNATREKが引き受ける。", {
    isTextBox: true, x: rx + 0.78, y: 4.48, w: rw - 1.05, h: 1.0,
    fontFace: F, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.3, margin: 0, valign: "top",
  });

  card(s, { x: rx, y: 5.84, w: rw, h: 0.66, fill: NAVY });
  s.addText("壁②が下がるからこそ、壁①を考えることに時間を使える。", {
    isTextBox: true, x: rx + 0.28, y: 5.84, w: rw - 0.56, h: 0.66,
    fontFace: F, fontSize: 13.5, bold: true, color: WHITE, margin: 0, valign: "middle", align: "center",
  });

  s.addNotes(
    "まず、今回のテーマと背景です。愛知銀行ではIBの普及に向けた施策やキャンペーンが行われていて、では実際にどれくらい使われるようになったのか、を確かめたい、というのが出発点でした。\n" +
    "実際に取り組んでみて感じたのは、業務上の疑問をデータで確かめようとすると、壁が2つあるということです。\n" +
    "壁①は「そもそも何を見れば答えになるのか」を決めること。これは人が考えるしかありません。\n" +
    "壁②は「そのデータを業務DBから取り出す」こと。業務DBには分析に使わない項目も大量にあるので、ここでつまずくことが多いと思います。\n" +
    "DYNATREKは、条件を指定して必要なデータだけを表示・集計できるツールで、この壁②を下げてくれます。\n" +
    "つまり、専門的な分析の知識がなくても、壁①を考えることに時間を使える。ここが今日お伝えしたい一番のポイントです。\n" +
    "では実際にどう考えたのか、次のスライドでご説明します。"
  );
  pageNum(s, 2);
}

// =========================================================
// Slide 3 — 何と比べるか
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "IBの件数を数えるだけでは、普及は測れない",
        "「IBでも行える取引が、どれだけIBで行われているか」という見方に置き換えた。");

  // 左：思考の流れ
  const steps = [
    ["まず思いついたのは「IBの利用件数を数える」こと。",
     "しかし取引全体が増えれば件数も増えるため、それだけでは「普及した」とは言い切れない。"],
    ["IBの普及とは「IBでも行える取引が、IBで行われるようになること」ではないか。",
     ""],
    ["そこで、IBで行われた取引と、ATMで行われた取引のうちIBでも代替できるもの（振込・振替など）を比較することにした。",
     ""],
  ];
  let y = 1.78;
  const lw = 6.0;
  steps.forEach((st, i) => {
    const h = i === 0 ? 1.62 : 1.32;
    card(s, { x: M, y: y, w: lw, h: h, fill: i === 2 ? ACCENT_TNT : TINT });
    badge(s, { x: M + 0.26, y: y + 0.26, color: i === 2 ? ACCENT : NAVY, label: String(i + 1) });
    const txt = st[1] ? st[0] + "\n" + st[1] : st[0];
    s.addText(txt, {
      isTextBox: true, x: M + 0.76, y: y + 0.2, w: lw - 1.05, h: h - 0.4,
      fontFace: F, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.32, margin: 0, valign: "middle",
    });
    y += h + 0.2;
  });

  // 右：概念図
  const rx = 6.95, rw = 5.83;
  s.addText("比較の考え方", {
    isTextBox: true, x: rx, y: 1.74, w: rw, h: 0.32,
    fontFace: F, fontSize: 14, bold: true, color: NAVY, margin: 0, valign: "middle",
  });

  card(s, { x: rx, y: 2.14, w: rw, h: 2.72, fill: WHITE, line: NAVY });
  s.addText("IBでも行える取引（振込・振替 など）", {
    isTextBox: true, x: rx + 0.2, y: 2.26, w: rw - 0.4, h: 0.36,
    fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0, valign: "middle", align: "center",
  });

  card(s, { x: rx + 0.28, y: 2.72, w: rw - 0.56, h: 0.92, fill: ACCENT });
  s.addText("A　IBで実施された取引", {
    isTextBox: true, x: rx + 0.28, y: 2.72, w: rw - 0.56, h: 0.92,
    fontFace: F, fontSize: 13.5, bold: true, color: WHITE, margin: 0, valign: "middle", align: "center",
  });

  card(s, { x: rx + 0.28, y: 3.78, w: rw - 0.56, h: 0.92, fill: NAVY_SOFT });
  s.addText("B　ATMで実施された取引\n（IBに移りうる部分）", {
    isTextBox: true, x: rx + 0.28, y: 3.78, w: rw - 0.56, h: 0.92,
    fontFace: F, fontSize: 13.5, bold: true, color: WHITE, lineSpacingMultiple: 1.15, margin: 0, valign: "middle", align: "center",
  });

  card(s, { x: rx, y: 5.0, w: rw, h: 0.72, fill: "F0F2F5", line: LINE });
  s.addText("ATMでしか行えない取引（現金の入出金 など）　→　今回は対象外", {
    isTextBox: true, x: rx + 0.2, y: 5.0, w: rw - 0.4, h: 0.72,
    fontFace: F, fontSize: 11.5, color: MUTED, margin: 0, valign: "middle", align: "center",
  });

  s.addText("対象期間：2025年4〜6月／2026年4〜6月（利用できたデータの範囲、計6か月）", {
    isTextBox: true, x: rx, y: 5.88, w: rw, h: 0.4,
    fontFace: F, fontSize: 11.5, color: MUTED, margin: 0, valign: "middle",
  });

  s.addNotes(
    "壁①、つまり「何を見れば答えになるのか」を考えた部分です。今回、時間を一番使ったのはここでした。\n" +
    "最初に思いついたのは、単純にIBの利用件数を数えることでした。ただ、取引全体が増えれば件数も増えるので、件数が増えたことだけでは「普及した」とは言い切れません。\n" +
    "そこで、IBの普及とは「IBでも行える取引が、IBで行われるようになること」ではないか、と考え直しました。\n" +
    "この考え方に立つと、比べるべきものが決まります。右の図のとおり、振込・振替のようにIBでも行える取引を、IBで実施されたものと、ATMで実施されたものに分けて見る、ということです。\n" +
    "ATMでしか行えない現金の入出金などは、IBに移りようがないので今回は対象から外しました。\n" +
    "対象期間は、使えるデータの範囲である2025年と2026年の4〜6月、計6か月です。\n" +
    "ここまでが「何を見るか」を決めるところで、ここから先が実際の集計になります。"
  );
  pageNum(s, 3);
}

// =========================================================
// Slide 4 — DYNATREKで集計・結果
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "条件を指定して集計しただけ ── IB利用件数は2026年に増加",
        "統計モデルや機械学習は使っていない。見たい条件を指定して数え、並べただけ。");

  // 左：手順
  const lx = M, lw = 4.55;
  s.addText("DYNATREKで行ったこと", {
    isTextBox: true, x: lx, y: 1.74, w: lw, h: 0.32,
    fontFace: F, fontSize: 14, bold: true, color: NAVY, margin: 0, valign: "middle",
  });
  const ops = [
    "期間を指定する\n2025年4〜6月／2026年4〜6月",
    "取引の種類を指定する\n振込・振替など、IBでも行える取引",
    "チャネル別（IB／ATM）に\n件数を集計して並べる",
  ];
  let oy = 2.16;
  ops.forEach((t, i) => {
    card(s, { x: lx, y: oy, w: lw, h: 0.88, fill: TINT });
    badge(s, { x: lx + 0.22, y: oy + 0.26, color: NAVY, label: String(i + 1) });
    s.addText(t, {
      isTextBox: true, x: lx + 0.7, y: oy + 0.08, w: lw - 0.92, h: 0.72,
      fontFace: F, fontSize: 11.5, color: INK, lineSpacingMultiple: 1.25, margin: 0, valign: "middle",
    });
    oy += 1.02;
  });

  // 数値コールアウト
  card(s, { x: lx, y: 5.40, w: lw, h: 1.55, fill: NAVY });
  s.addText("IB利用件数（1か月あたり）", {
    isTextBox: true, x: lx + 0.24, y: 5.52, w: lw - 0.48, h: 0.28,
    fontFace: F, fontSize: 11, color: ICE, margin: 0, valign: "middle",
  });
  s.addText([
    { text: "2025年 4〜6月　", options: { fontSize: 11, color: ICE } },
    { text: "約15〜16", options: { fontSize: 19, bold: true, color: WHITE } },
    { text: "万件", options: { fontSize: 12, color: ICE } },
  ], {
    isTextBox: true, x: lx + 0.24, y: 5.86, w: lw - 0.48, h: 0.44,
    fontFace: F, margin: 0, valign: "middle",
  });
  s.addText([
    { text: "2026年 4〜6月　", options: { fontSize: 11, color: ICE } },
    { text: "約18〜20", options: { fontSize: 21, bold: true, color: "FFC9BC" } },
    { text: "万件", options: { fontSize: 12, color: ICE } },
  ], {
    isTextBox: true, x: lx + 0.24, y: 6.34, w: lw - 0.48, h: 0.48,
    fontFace: F, margin: 0, valign: "middle",
  });

  // 右：グラフ枠
  const gx = 5.5, gw = 7.28;
  const graphs = [
    ["グラフ①", "IB利用件数の月次推移",
     "2025年4・5・6月 と 2026年4・5・6月 を並べた棒グラフ（縦軸：件数）"],
    ["グラフ②", "IB と ATM（IB代替可能取引）の比較",
     "同じ月について、IBで行われた件数とATMで行われたIB代替可能取引の件数を横並びにした棒グラフ"],
  ];
  let gy = 1.74;
  graphs.forEach((g) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: gx, y: gy, w: gw, h: 2.35, rectRadius: 0.05,
      fill: { color: "FAFBFD" }, line: { color: NAVY_SOFT, width: 1.25, dashType: "dash" },
    });
    s.addText([
      { text: g[0] + "　", options: { fontSize: 12, bold: true, color: ACCENT } },
      { text: g[1], options: { fontSize: 13, bold: true, color: NAVY } },
    ], {
      isTextBox: true, x: gx + 0.3, y: gy + 0.66, w: gw - 0.6, h: 0.4,
      fontFace: F, margin: 0, valign: "middle", align: "center",
    });
    s.addText(g[2], {
      isTextBox: true, x: gx + 0.5, y: gy + 1.10, w: gw - 1.0, h: 0.6,
      fontFace: F, fontSize: 11.5, color: MUTED, lineSpacingMultiple: 1.25, margin: 0, valign: "top", align: "center",
    });
    gy += 2.61;
  });

  s.addText("※ グラフはDYNATREKの集計結果を貼り付ける（上の枠内）", {
    isTextBox: true, x: gx, y: 6.78, w: 6.3, h: 0.28,
    fontFace: F, fontSize: 10, color: "9AA9BA", margin: 0, valign: "middle",
  });

  s.addNotes(
    "実際にDYNATREKで行ったことは、左に書いた3つだけです。期間を指定し、取引の種類を指定し、チャネル別に件数を集計して並べる。統計モデルや機械学習は使っていません。\n" +
    "結果ですが、IBの利用件数は1か月あたり、2025年が約15万から16万件、2026年が約18万から20万件でした。前年の同じ4〜6月と比べると増えています。\n" +
    "グラフ①が月ごとの推移、グラフ②が同じ月のIBとATMのIB代替可能取引を並べたものです。（※グラフ②の見え方については、実際の数字を見ながら説明する）\n" +
    "ここで強調したいのは、この集計自体は、条件を指定して数えただけだということです。特別な分析の知識は使っていません。\n" +
    "ただ、この数字をどう読むかは、また別の話になります。"
  );
  pageNum(s, 4);
}

// =========================================================
// Slide 5 — 言えること／言えないこと／今後
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "「増えた」とは言えるが、「施策の効果」とは言い切れない",
        "数字から言えることと言えないことを分けることも、分析の一部だと考えた。");

  const cols = [
    {
      mark: "○", color: NAVY, fill: TINT, title: "このデータから言えること",
      items: [
        "前年の同時期（4〜6月）と比べて、2026年のIB利用件数は増えている。",
        "1か月あたり 約15〜16万件 → 約18〜20万件。",
      ],
    },
    {
      mark: "△", color: ACCENT, fill: ACCENT_TNT, title: "このデータからは言えないこと",
      items: [
        "比較できるのは4〜6月の3か月分×2年のみで、連続した推移が追えない。",
        "口座数の変化、取引全体の増減、季節性、社会全体でのIB利用の広がりなど、施策以外の要因を切り分けられない。",
        "→「IB普及施策によって増えた」という因果関係は、このデータだけでは判断できない。",
      ],
    },
    {
      mark: "▶", color: NAVY_SOFT, fill: TINT, title: "データが蓄積されれば確かめられること",
      items: [
        "連続した月次の推移を見て、施策の前後で変化しているかを確かめる。",
        "ATMのIB代替可能取引が減っているかを、IBの増加と対応させて見る。",
        "支店や年代などの属性別に、普及の進み方の違いを見る。",
      ],
    },
  ];

  const cw = 3.85, gap = 0.29;
  cols.forEach((c, i) => {
    const x = M + i * (cw + gap);
    card(s, { x: x, y: 1.68, w: cw, h: 4.50, fill: c.fill });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + 0.24, y: 1.92, w: 0.4, h: 0.4, fill: { color: c.color }, line: { color: c.color, width: 0.5 },
    });
    s.addText(c.mark, {
      isTextBox: true, x: x + 0.24, y: 1.92, w: 0.4, h: 0.4,
      fontFace: F, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.title, {
      isTextBox: true, x: x + 0.76, y: 1.88, w: cw - 1.0, h: 0.5,
      fontFace: F, fontSize: 13.5, bold: true, color: c.color, margin: 0, valign: "middle",
    });
    s.addText(body(c.items, { bullet: true, gap: 10 }), {
      isTextBox: true, x: x + 0.28, y: 2.54, w: cw - 0.56, h: 3.52,
      fontFace: F, fontSize: 12, color: INK, lineSpacingMultiple: 1.32, margin: 0, valign: "top",
    });
  });

  card(s, { x: M, y: 6.34, w: CW, h: 0.60, fill: NAVY });
  s.addText("「増えている」という事実は確認できた。同時に、次に何を確かめればよいかがはっきりした。", {
    isTextBox: true, x: M + 0.28, y: 6.34, w: CW - 0.56, h: 0.60,
    fontFace: F, fontSize: 13.5, bold: true, color: WHITE, margin: 0, valign: "middle", align: "center",
  });

  s.addNotes(
    "結果をどう読むかです。ここは意識して3つに分けました。\n" +
    "まず言えることは、前年の同じ4〜6月と比べてIBの利用件数は増えている、という事実です。\n" +
    "一方で、言えないこともあります。今回使えたデータは4〜6月の3か月分が2年分だけで、連続した推移が追えません。また、口座数の変化や取引全体の増減、季節性、社会全体でIB利用が広がっていること、といった施策以外の要因を切り分けることもできません。\n" +
    "ですので、「IB普及施策によって増えた」という因果関係は、このデータだけでは判断できない、というのが今回の結論です。ここは断定しないようにしました。\n" +
    "ただ、言えないことがはっきりしたことで、次に何を見ればよいかも具体的になりました。連続した月次推移で施策の前後を比べる、ATM側のIB代替可能取引が減っているかを対応させて見る、支店や年代などの属性別に見る、といったことです。\n" +
    "これらは、データが蓄積されれば、今回と同じやり方で確かめられると考えています。"
  );
  pageNum(s, 5);
}

// =========================================================
// Slide 6 — まとめ・学び
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  s.addText("難しいのは分析手法ではなく、「何を見るか」を決めること", {
    isTextBox: true, x: M, y: 0.42, w: CW, h: 0.62,
    fontFace: F, fontSize: 26, bold: true, color: WHITE, margin: 0, valign: "middle",
  });
  s.addText("DYNATREKが「取り出す」壁を下げてくれるからこそ、「何を見るか」を考えることに時間を使えた。", {
    isTextBox: true, x: M, y: 1.08, w: CW, h: 0.4,
    fontFace: F, fontSize: 13.5, bold: true, color: "FFC9BC", margin: 0, valign: "middle",
  });

  const blocks = [
    { t: "DYNATREKに感じた価値",
      b: "業務DBには分析に使わない項目も多く含まれ、目的のデータを取り出すこと自体が難しい。DYNATREKは条件を指定して必要なデータだけを表示・集計できるため、専門的な分析知識がなくても、業務上の疑問をその場でデータに当てて確かめられる。" },
    { t: "今回いちばん難しかったこと",
      b: "大学の研究ではPOSデータを扱ってきたが、今回難しかったのは手法ではなかった。目的に対して何を見るべきか、何と比べるべきか、どこまで条件を絞るか。DYNATREKは条件を細かく設定できるぶん、仮説はいくらでも立てられる。限られた時間の中では「分析できること」ではなく「目的の達成に必要な分析」を選ぶことが重要だと感じた。" },
    { t: "業務の現場で見えたこと",
      b: "期間中は複数の打ち合わせにも参加させていただいた。専門的な内容も多く、すべてを理解することは難しかったが、社員の方々が意見を出し合いながら進めていく様子から、実際の仕事の進め方を知ることができた。" },
  ];
  const bw = 3.85, bgap = 0.29;
  blocks.forEach((bl, i) => {
    const x = M + i * (bw + bgap);
    card(s, { x: x, y: 1.72, w: bw, h: 3.72, fill: NAVY_MID, line: NAVY_SOFT });
    s.addText(bl.t, {
      isTextBox: true, x: x + 0.28, y: 1.92, w: bw - 0.56, h: 0.42,
      fontFace: F, fontSize: 14, bold: true, color: "FFC9BC", margin: 0, valign: "middle",
    });
    s.addText(bl.b, {
      isTextBox: true, x: x + 0.28, y: 2.42, w: bw - 0.56, h: 2.8,
      fontFace: F, fontSize: 11.5, color: WHITE, lineSpacingMultiple: 1.35, margin: 0, valign: "top",
    });
  });

  card(s, { x: M, y: 5.66, w: CW, h: 1.24, fill: ACCENT });
  s.addText("専門的なデータ分析の知識がなくても、DYNATREKを使えば、業務上の疑問に対して\nデータで確認・比較していける ── 今回の分析を通して、それを実感しました。", {
    isTextBox: true, x: M + 0.3, y: 5.66, w: CW - 0.6, h: 1.24,
    fontFace: F, fontSize: 15, bold: true, color: WHITE, lineSpacingMultiple: 1.35, margin: 0, valign: "middle", align: "center",
  });

  s.addNotes(
    "最後にまとめです。\n" +
    "今回、DYNATREKに感じた価値は、機能の多さではなく、目的のデータを取り出すところの負担を下げてくれることでした。業務DBには分析に使わない項目も多く、そこが難しいのですが、DYNATREKは条件を指定して必要なデータだけを見られるので、専門的な分析知識がなくても、疑問をその場でデータに当てられます。\n" +
    "一方で、今回いちばん難しかったのは分析手法ではありませんでした。私は大学の研究でPOSデータを扱ってきたので、分析そのものの経験はあります。それでも、目的に対して何を見るべきか、何と比べるべきか、どこまで条件を絞るかを決めることが難しかったです。\n" +
    "DYNATREKは条件を細かく設定できるので、仮説はいくらでも立てられます。ただ実際の分析時間には限りがあるので、「分析できること」ではなく「目的の達成に必要な分析」を選ぶことが重要だと感じました。これが今回いちばんの学びです。\n" +
    "また、複数の打ち合わせにも参加させていただき、社員の方々が意見を出し合いながら進めていく様子を見られたことも、実際の業務を知るうえで貴重な経験でした。\n" +
    "専門的なデータ分析の知識がなくても、DYNATREKを使えば業務上の疑問にデータで向き合っていける。今回の分析を通して、それを実感しました。ありがとうございました。"
  );
  pageNum(s, 6, true);
}

pres.writeFile({ fileName: process.argv[2] || "dynatrek_ib_presentation.pptx" })
  .then((f) => console.log("created:", f));
