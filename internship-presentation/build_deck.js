// 日立製作所 インターンシップ成果発表スライド生成スクリプト
// 実行: NODE_PATH=<pptxgenjs のあるディレクトリ>/node_modules node build_deck.js [出力ファイル名]
const PptxGenJS = require("pptxgenjs");

const NAVY="13294B", NAVY_MID="1C3A66", NAVY_SOFT="2A4A7A";
const INK="1E2A38", MUTED="5D6E82", HINT="8A9AAD";
const ICE="CADCFC", TINT="EEF2F8", LINE="D3DCE8";
const ACCENT="D9482F", ACCENT_TNT="FBEAE4", ACCENT_LT="FFC9BC";
const WHITE="FFFFFF", F="Meiryo";
const W=13.33, H=7.5, M=0.55, CW=12.23;

const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";
pres.author = "Hitachi Internship";
pres.title  = "DYNATREKで業務の疑問をデータで確かめる";

function pageNum(s, n, dark) {
  s.addText(String(n), { isTextBox:true, x:W-0.85, y:7.02, w:0.4, h:0.3,
    fontFace:F, fontSize:10, color: dark?NAVY_SOFT:"9AA9BA", align:"right", margin:0 });
}
function head(s, title, lead) {
  s.addText(title, { isTextBox:true, x:M, y:0.38, w:CW, h:0.62,
    fontFace:F, fontSize:26, bold:true, color:NAVY, margin:0, valign:"middle" });
  if (lead) s.addText(lead, { isTextBox:true, x:M, y:1.04, w:CW, h:0.42,
    fontFace:F, fontSize:13.5, bold:true, color:ACCENT, margin:0, valign:"middle" });
}
function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, { x:o.x, y:o.y, w:o.w, h:o.h, rectRadius:0.06,
    fill:{ color:o.fill||TINT },
    line:o.line?{ color:o.line, width:1 }:{ color:o.fill||TINT, width:0.5 } });
}
function badge(s, o) {
  const d = o.d || 0.36;
  s.addShape(pres.ShapeType.ellipse, { x:o.x, y:o.y, w:d, h:d,
    fill:{ color:o.color }, line:{ color:o.color, width:0.5 } });
  s.addText(o.label, { isTextBox:true, x:o.x, y:o.y, w:d, h:d,
    fontFace:F, fontSize:o.fs||13, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
}
function frame(s, o) {
  s.addShape(pres.ShapeType.roundRect, { x:o.x, y:o.y, w:o.w, h:o.h, rectRadius:0.05,
    fill:{ color:"FAFBFD" }, line:{ color:NAVY_SOFT, width:1.25, dashType:"dash" } });
  s.addText([{ text:o.tag+"　", options:{ fontSize:12, bold:true, color:ACCENT } },
             { text:o.title, options:{ fontSize:13, bold:true, color:NAVY } }],
    { isTextBox:true, x:o.x+0.3, y:o.y+o.h/2-0.42, w:o.w-0.6, h:0.4,
      fontFace:F, margin:0, valign:"middle", align:"center" });
  s.addText(o.desc, { isTextBox:true, x:o.x+0.5, y:o.y+o.h/2, w:o.w-1.0, h:0.62,
    fontFace:F, fontSize:11.5, color:MUTED, lineSpacingMultiple:1.25, margin:0, valign:"top", align:"center" });
}
function bullets(items, gap) {
  return items.map((t,i)=>({ text:t, options:{ breakLine:i!==items.length-1,
    bullet:{ characterCode:"25AA" }, paraSpaceAfter: gap===undefined?9:gap } }));
}

// ============ 1. 表紙 ============
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.ShapeType.ellipse, { x:9.45, y:1.15, w:3.0, h:3.0, fill:{color:NAVY_MID}, line:{color:NAVY_MID,width:0.5} });
  s.addShape(pres.ShapeType.ellipse, { x:10.15, y:2.62, w:2.9, h:2.9, fill:{color:NAVY_SOFT}, line:{color:NAVY_SOFT,width:0.5} });

  s.addText("日立製作所　インターンシップ　成果発表", { isTextBox:true, x:M, y:0.85, w:8.4, h:0.35,
    fontFace:F, fontSize:12, color:ICE, charSpacing:1, margin:0, valign:"middle" });
  s.addText("分析の専門知識や経験がなくても、\nDYNATREKなら分析できる", { isTextBox:true, x:M, y:1.6, w:8.9, h:1.9,
    fontFace:F, fontSize:33, bold:true, color:WHITE, lineSpacingMultiple:1.25, margin:0, valign:"middle" });
  s.addText("業務データを集計・比較して、そこから傾向をつかむまでを実際にやってみました", {
    isTextBox:true, x:M, y:3.72, w:9.0, h:0.42, fontFace:F, fontSize:15.5, color:ICE, margin:0, valign:"middle" });

  s.addShape(pres.ShapeType.roundRect, { x:M, y:4.35, w:7.9, h:0.62, rectRadius:0.08,
    fill:{color:NAVY_MID}, line:{color:NAVY_SOFT,width:1} });
  s.addText("その一例：愛知銀行　インターネットバンキング（IB）の普及状況の分析", {
    isTextBox:true, x:M+0.22, y:4.35, w:7.5, h:0.62, fontFace:F, fontSize:13, color:WHITE, margin:0, valign:"middle" });

  s.addText("発表者：（氏名）　／　2026年●月●日", { isTextBox:true, x:M, y:6.35, w:7.0, h:0.35,
    fontFace:F, fontSize:12, color:ICE, margin:0, valign:"middle" });

  s.addNotes(
    "本日は、2週間のインターンシップで取り組んだ内容についてご報告します。\n"+
    "今日お伝えしたいのは、DYNATREKを使えば、データ分析の専門的な知識や経験がない人でも、業務データを集計・比較して傾向をつかめる、ということです。\n"+
    "そして今回は、その一例として、私が実際に行った分析をご紹介します。題材は、愛知銀行のインターネットバンキング、IBの普及状況です。\n"+
    "分析結果そのものよりも、どう考えて、どう使ったのか、という流れを見ていただければと思います。");
  pageNum(s, 1, true);
}

// ============ 2. 本発表の目的（2ビート） ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "本発表でお伝えしたいことは、次の2点です");

  card(s, { x:M, y:1.42, w:5.4, h:2.5, fill:NAVY });
  badge(s, { x:M+0.28, y:1.66, color:ACCENT, label:"1" });
  s.addText("お伝えしたいこと", { isTextBox:true, x:M+0.76, y:1.64, w:4.3, h:0.4,
    fontFace:F, fontSize:14, bold:true, color:ACCENT_LT, margin:0, valign:"middle" });
  s.addText("DYNATREKを使えば、データ分析の専門的な知識や経験がない人でも、業務データを集計・比較して、そこから傾向をつかむことができる。", {
    isTextBox:true, x:M+0.3, y:2.2, w:4.8, h:1.55,
    fontFace:F, fontSize:14, bold:true, color:WHITE, lineSpacingMultiple:1.4, margin:0, valign:"middle" });

  s.addShape(pres.ShapeType.rightArrow, { x:6.08, y:2.42, w:0.62, h:0.5,
    fill:{ color:ACCENT }, line:{ color:ACCENT, width:0.5 } });

  card(s, { x:6.9, y:1.42, w:5.88, h:2.5, fill:ACCENT_TNT, line:ACCENT });
  badge(s, { x:6.9+0.28, y:1.66, color:ACCENT, label:"2" });
  s.addText("本発表ですること", { isTextBox:true, x:6.9+0.76, y:1.64, w:4.8, h:0.4,
    fontFace:F, fontSize:14, bold:true, color:ACCENT, margin:0, valign:"middle" });
  s.addText("今回は、その一例として、私が実際に行った分析を紹介します。\n題材は、愛知銀行のインターネットバンキング（IB）の普及状況です。", {
    isTextBox:true, x:6.9+0.3, y:2.2, w:5.28, h:1.55,
    fontFace:F, fontSize:14, bold:true, color:INK, lineSpacingMultiple:1.4, margin:0, valign:"middle" });

  s.addText("なぜ、専門知識や経験がなくても分析できるのか", { isTextBox:true, x:M, y:4.22, w:CW, h:0.36,
    fontFace:F, fontSize:14, bold:true, color:NAVY, margin:0, valign:"middle" });

  const walls = [
    { x:M, fill:TINT, col:NAVY, lb:"壁①", t:"何を見れば答えになるのかを決める",
      b:"「普及したかどうか」を、どの数字で見れば確かめたことになるのか。ここは人が考えるところで、ツールでは代われない。" },
    { x:6.78, fill:ACCENT_TNT, col:ACCENT, lb:"壁②", t:"必要なデータを業務DBから取り出す",
      b:"業務DBには分析に使わない項目も数多く含まれ、目的のデータを抜き出すこと自体が難しい。ここをDYNATREKが引き受けるので、壁①を考えることに時間を使える。" },
  ];
  walls.forEach(w=>{
    card(s, { x:w.x, y:4.66, w:6.0, h:1.78, fill:w.fill });
    s.addShape(pres.ShapeType.roundRect, { x:w.x+0.26, y:4.9, w:0.72, h:0.34, rectRadius:0.05,
      fill:{color:w.col}, line:{color:w.col,width:0.5} });
    s.addText(w.lb, { isTextBox:true, x:w.x+0.26, y:4.9, w:0.72, h:0.34,
      fontFace:F, fontSize:11.5, bold:true, color:WHITE, align:"center", valign:"middle", margin:0 });
    s.addText(w.t, { isTextBox:true, x:w.x+1.1, y:4.88, w:4.6, h:0.38,
      fontFace:F, fontSize:14, bold:true, color:w.col, margin:0, valign:"middle" });
    s.addText(w.b, { isTextBox:true, x:w.x+0.28, y:5.34, w:5.44, h:0.98,
      fontFace:F, fontSize:12, color:INK, lineSpacingMultiple:1.3, margin:0, valign:"top" });
  });

  s.addNotes(
    "はじめに、今日の発表の目的をはっきりさせておきます。2点です。\n"+
    "1点目。DYNATREKを使えば、データ分析の専門的な知識や経験がない人でも、業務データを集計・比較して、傾向をつかむことができる。これが今日お伝えしたいことです。\n"+
    "2点目。今回は、その一例として、私が実際に行った分析をご紹介します。題材は愛知銀行のIBの普及状況です。\n"+
    "なぜ専門知識がなくてもできるのか、簡単に触れておきます。業務上の疑問をデータで確かめようとすると、壁が2つあります。\n"+
    "壁①は「何を見れば答えになるのか」を決めること。これは人が考えるしかありません。\n"+
    "壁②は「そのデータを業務DBから取り出す」こと。業務DBには分析に使わない項目も大量にあって、専門知識がないとここで止まります。\n"+
    "DYNATREKはこの壁②を引き受けてくれるので、壁①を考えることに時間を使えます。\n"+
    "では、その一例を見ていただきます。まずは題材の背景からです。");
  pageNum(s, 2);
}

// ============ 3. 愛知銀行の施策（背景） ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "愛知銀行では、IBの利用を広げるための取り組みが行われている",
        "まず、どのような取り組みが行われているのかを押さえておく。");

  const cw = 3.85, gap = 0.29;
  ["施策①","施策②","施策③"].forEach((lb,i)=>{
    const x = M + i*(cw+gap);
    card(s, { x:x, y:1.68, w:cw, h:3.86, fill:i===0?TINT:WHITE, line:i===0?null:LINE });
    s.addText(lb, { isTextBox:true, x:x+0.28, y:1.86, w:cw-0.56, h:0.3,
      fontFace:F, fontSize:11.5, bold:true, color:ACCENT, margin:0, valign:"middle" });
    s.addText("（施策の名称を記入）", { isTextBox:true, x:x+0.28, y:2.18, w:cw-0.56, h:0.6,
      fontFace:F, fontSize:15, bold:true, color:HINT, lineSpacingMultiple:1.2, margin:0, valign:"top" });
    let ry = 2.92;
    [["内容","どのような取り組みか"],["ねらい","何を増やしたいのか"],["実施時期","いつ行われているか"]].forEach(r=>{
      s.addText(r[0], { isTextBox:true, x:x+0.28, y:ry, w:cw-0.56, h:0.26,
        fontFace:F, fontSize:11, bold:true, color:NAVY, margin:0, valign:"middle" });
      s.addText("（"+r[1]+"を記入）", { isTextBox:true, x:x+0.28, y:ry+0.26, w:cw-0.56, h:0.52,
        fontFace:F, fontSize:11.5, color:HINT, lineSpacingMultiple:1.25, margin:0, valign:"top" });
      ry += 0.84;
    });
  });

  card(s, { x:M, y:5.78, w:CW, h:0.66, fill:NAVY });
  s.addText("では、こうした取り組みのもとで、IBの利用は実際にどう変化しているのか。", {
    isTextBox:true, x:M+0.28, y:5.78, w:CW-0.56, h:0.66,
    fontFace:F, fontSize:13.5, bold:true, color:WHITE, margin:0, valign:"middle", align:"center" });

  s.addNotes(
    "題材の背景です。愛知銀行では、IBの利用を広げるための取り組みが行われています。\n"+
    "（※ここで、実際の施策の内容・ねらい・実施時期を具体的に説明する。スライドの記入欄に入れた内容をそのまま話す）\n"+
    "こうした取り組みが行われているなかで、では実際にIBの利用はどう変化しているのか。それを確かめたい、というのが今回の出発点でした。");
  pageNum(s, 3);
}

// ============ 4. 何と比べるか（山場） ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "IBの件数を数えるだけでは、普及は測れない",
        "「IBでも行える取引が、どれだけIBで行われているか」という見方に置き換えた。");

  const steps = [
    ["まず思いついたのは「IBの利用件数を数える」こと。",
     "しかし取引全体が増えれば件数も増えるため、それだけでは「普及した」とは言い切れない。"],
    ["IBの普及とは「IBでも行える取引が、IBで行われるようになること」ではないか。",""],
    ["そこで、IBで行われた取引と、ATMで行われた取引のうちIBでも代替できるもの（振込・振替など）を比較することにした。",""],
  ];
  let y = 1.78; const lw = 6.0;
  steps.forEach((st,i)=>{
    const h = i===0 ? 1.62 : 1.32;
    card(s, { x:M, y:y, w:lw, h:h, fill:i===2?ACCENT_TNT:TINT });
    badge(s, { x:M+0.26, y:y+0.26, color:i===2?ACCENT:NAVY, label:String(i+1) });
    s.addText(st[1]?st[0]+"\n"+st[1]:st[0], { isTextBox:true, x:M+0.76, y:y+0.2, w:lw-1.05, h:h-0.4,
      fontFace:F, fontSize:12.5, color:INK, lineSpacingMultiple:1.32, margin:0, valign:"middle" });
    y += h + 0.2;
  });

  const rx = 6.95, rw = 5.83;
  s.addText("比較の考え方", { isTextBox:true, x:rx, y:1.74, w:rw, h:0.32,
    fontFace:F, fontSize:14, bold:true, color:NAVY, margin:0, valign:"middle" });
  card(s, { x:rx, y:2.14, w:rw, h:2.72, fill:WHITE, line:NAVY });
  s.addText("IBでも行える取引（振込・振替 など）", { isTextBox:true, x:rx+0.2, y:2.26, w:rw-0.4, h:0.36,
    fontFace:F, fontSize:12.5, bold:true, color:NAVY, margin:0, valign:"middle", align:"center" });
  card(s, { x:rx+0.28, y:2.72, w:rw-0.56, h:0.92, fill:ACCENT });
  s.addText("A　IBで実施された取引", { isTextBox:true, x:rx+0.28, y:2.72, w:rw-0.56, h:0.92,
    fontFace:F, fontSize:13.5, bold:true, color:WHITE, margin:0, valign:"middle", align:"center" });
  card(s, { x:rx+0.28, y:3.78, w:rw-0.56, h:0.92, fill:NAVY_SOFT });
  s.addText("B　ATMで実施された取引\n（IBに移りうる部分）", { isTextBox:true, x:rx+0.28, y:3.78, w:rw-0.56, h:0.92,
    fontFace:F, fontSize:13.5, bold:true, color:WHITE, lineSpacingMultiple:1.15, margin:0, valign:"middle", align:"center" });
  card(s, { x:rx, y:5.0, w:rw, h:0.72, fill:"F0F2F5", line:LINE });
  s.addText("ATMでしか行えない取引（現金の入出金 など）　→　今回は対象外", {
    isTextBox:true, x:rx+0.2, y:5.0, w:rw-0.4, h:0.72,
    fontFace:F, fontSize:11.5, color:MUTED, margin:0, valign:"middle", align:"center" });
  s.addText("対象期間：2025年4〜6月／2026年4〜6月（利用できたデータの範囲、計6か月）", {
    isTextBox:true, x:rx, y:5.88, w:rw, h:0.4, fontFace:F, fontSize:11.5, color:MUTED, margin:0, valign:"middle" });

  s.addNotes(
    "ここが壁①、つまり「何を見れば答えになるのか」を考えた部分です。今回いちばん時間を使ったのはここでした。\n"+
    "最初に思いついたのは、単純にIBの利用件数を数えることでした。ただ、取引全体が増えれば件数も増えるので、件数が増えたことだけでは「普及した」とは言い切れません。\n"+
    "そこで、IBの普及とは「IBでも行える取引が、IBで行われるようになること」ではないか、と考え直しました。\n"+
    "この考え方に立つと、比べるべきものが決まります。右の図のとおり、振込・振替のようにIBでも行える取引を、IBで実施されたものと、ATMで実施されたものに分けて見る、ということです。\n"+
    "ATMでしか行えない現金の入出金などは、IBに移りようがないので今回は対象から外しました。\n"+
    "対象期間は、使えるデータの範囲である2025年と2026年の4〜6月、計6か月です。\n"+
    "ここまでが人が考えるところです。ここから先、実際の画面をお見せします。");
  pageNum(s, 4);
}

// ============ 5. DYNATREKの画面（条件指定 → 出力） ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "やったことは、画面で条件を指定して集計しただけ",
        "統計モデルや機械学習は使っていない。特別な操作も必要なかった。");

  const lx = M, lw = 3.6;
  s.addText("指定した条件", { isTextBox:true, x:lx, y:1.68, w:lw, h:0.32,
    fontFace:F, fontSize:14, bold:true, color:NAVY, margin:0, valign:"middle" });
  const conds = [
    "期間\n2025年4〜6月／2026年4〜6月",
    "取引の種類\n振込・振替など、IBでも行える取引",
    "チャネル\nIB／ATM 別に集計",
  ];
  let cy = 2.08;
  conds.forEach((t,i)=>{
    card(s, { x:lx, y:cy, w:lw, h:0.88, fill:TINT });
    badge(s, { x:lx+0.2, y:cy+0.26, color:NAVY, label:String(i+1) });
    s.addText(t, { isTextBox:true, x:lx+0.66, y:cy+0.08, w:lw-0.88, h:0.72,
      fontFace:F, fontSize:11, color:INK, lineSpacingMultiple:1.25, margin:0, valign:"middle" });
    cy += 1.0;
  });
  card(s, { x:lx, y:5.2, w:lw, h:1.2, fill:NAVY });
  s.addText("この2画面だけで、\n比べたい数字が出せた。", { isTextBox:true, x:lx+0.24, y:5.2, w:lw-0.48, h:1.2,
    fontFace:F, fontSize:13, bold:true, color:WHITE, lineSpacingMultiple:1.3, margin:0, valign:"middle", align:"center" });

  const gx = 4.35, gw = 8.43;
  frame(s, { x:gx, y:1.68, w:gw, h:2.3, tag:"画面①", title:"条件を指定する画面",
    desc:"期間・取引の種類・チャネルを指定しているところのスクリーンショット" });
  s.addShape(pres.ShapeType.downArrow, { x:gx+gw/2-0.16, y:4.04, w:0.32, h:0.24,
    fill:{color:ACCENT}, line:{color:ACCENT,width:0.5} });
  frame(s, { x:gx, y:4.34, w:gw, h:2.3, tag:"画面②", title:"出力された集計結果",
    desc:"指定した条件で集計された件数が表示されているところのスクリーンショット" });

  s.addNotes(
    "実際の画面をお見せします。やったことは、左に書いた3つの条件を指定しただけです。\n"+
    "期間、取引の種類、そしてチャネルをIBとATMに分けること。この3つです。\n"+
    "画面①が、その条件を指定しているところです。（※どこに何を入れたのかを指しながら説明する）\n"+
    "画面②が、その結果として出てきた集計結果です。（※どの数字を見ているのかを指しながら説明する）\n"+
    "統計モデルや機械学習は使っていませんし、特別な操作も必要ありませんでした。この2画面だけで、比べたい数字が出せています。\n"+
    "では、出てきた数字を見ていきます。");
  pageNum(s, 5);
}

// ============ 6. 結果 ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "IB利用件数は、2025年より2026年のほうが多い",
        "出てきた数字を、月ごとに並べて比べた。");

  const lx = M, lw = 4.3;
  card(s, { x:lx, y:1.68, w:lw, h:1.95, fill:NAVY });
  s.addText("IB利用件数（1か月あたり）", { isTextBox:true, x:lx+0.24, y:1.82, w:lw-0.48, h:0.3,
    fontFace:F, fontSize:11, color:ICE, margin:0, valign:"middle" });
  s.addText([{ text:"2025年 4〜6月　", options:{ fontSize:11, color:ICE } },
             { text:"約15〜16", options:{ fontSize:19, bold:true, color:WHITE } },
             { text:"万件", options:{ fontSize:12, color:ICE } }],
    { isTextBox:true, x:lx+0.24, y:2.2, w:lw-0.48, h:0.5, fontFace:F, margin:0, valign:"middle" });
  s.addText([{ text:"2026年 4〜6月　", options:{ fontSize:11, color:ICE } },
             { text:"約18〜20", options:{ fontSize:21, bold:true, color:ACCENT_LT } },
             { text:"万件", options:{ fontSize:12, color:ICE } }],
    { isTextBox:true, x:lx+0.24, y:2.78, w:lw-0.48, h:0.55, fontFace:F, margin:0, valign:"middle" });

  card(s, { x:lx, y:3.86, w:lw, h:2.1, fill:TINT });
  s.addText("あわせて確認したこと", { isTextBox:true, x:lx+0.26, y:4.02, w:lw-0.52, h:0.3,
    fontFace:F, fontSize:12.5, bold:true, color:NAVY, margin:0, valign:"middle" });
  s.addText(bullets([
    "ATMで行われた取引のうちIBでも代替できるものも、同じ条件で集計した。",
    "同じ月どうしで並べ、IBとATMの両方から利用状況を見た。",
  ]), { isTextBox:true, x:lx+0.26, y:4.38, w:lw-0.52, h:1.45,
    fontFace:F, fontSize:11.5, color:INK, lineSpacingMultiple:1.3, margin:0, valign:"top" });

  const gx = 5.2, gw = 7.58;
  frame(s, { x:gx, y:1.68, w:gw, h:2.35, tag:"グラフ①", title:"IB利用件数の月次推移",
    desc:"2025年4・5・6月 と 2026年4・5・6月 を並べた棒グラフ（縦軸：件数）" });
  frame(s, { x:gx, y:4.28, w:gw, h:2.35, tag:"グラフ②", title:"IB と ATM（IB代替可能取引）の比較",
    desc:"同じ月について、IBの件数とATMのIB代替可能取引の件数を横並びにした棒グラフ" });
  s.addText("※ グラフは上の枠内に貼り付ける", { isTextBox:true, x:gx, y:6.7, w:6.0, h:0.28,
    fontFace:F, fontSize:10, color:"9AA9BA", margin:0, valign:"middle" });

  s.addNotes(
    "結果です。IBの利用件数は1か月あたり、2025年が約15万から16万件、2026年が約18万から20万件でした。前年の同じ4〜6月と比べると増えています。\n"+
    "グラフ①が月ごとの推移です。\n"+
    "あわせて、ATMで行われた取引のうちIBでも代替できるものも、まったく同じ条件指定で集計し、同じ月どうしで並べました。それがグラフ②です。\n"+
    "（※グラフ②については、実際に出た数字を見ながら説明する）\n"+
    "では、この数字から何が言えるのか。ここは分けて考える必要があります。");
  pageNum(s, 6);
}

// ============ 7. 言えること／言えないこと／今後 ============
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "「増えた」とは言えるが、「施策の効果」とは言い切れない",
        "数字から言えることと言えないことを分けることも、分析の一部だと考えた。");

  const cols = [
    { mark:"○", color:NAVY, fill:TINT, title:"このデータから言えること", items:[
      "前年の同時期（4〜6月）と比べて、2026年のIB利用件数は増えている。",
      "1か月あたり 約15〜16万件 → 約18〜20万件。" ] },
    { mark:"△", color:ACCENT, fill:ACCENT_TNT, title:"このデータからは言えないこと", items:[
      "比較できるのは4〜6月の3か月分×2年のみで、連続した推移が追えない。",
      "口座数の変化、取引全体の増減、季節性、社会全体でのIB利用の広がりなど、施策以外の要因を切り分けられない。",
      "→「IB普及施策によって増えた」という因果関係は、このデータだけでは判断できない。" ] },
    { mark:"▶", color:NAVY_SOFT, fill:TINT, title:"データが蓄積されれば確かめられること", items:[
      "連続した月次の推移を見て、施策の前後で変化しているかを確かめる。",
      "ATMのIB代替可能取引が減っているかを、IBの増加と対応させて見る。",
      "支店や年代などの属性別に、普及の進み方の違いを見る。" ] },
  ];
  const cw = 3.85, gap = 0.29;
  cols.forEach((c,i)=>{
    const x = M + i*(cw+gap);
    card(s, { x:x, y:1.68, w:cw, h:4.50, fill:c.fill });
    badge(s, { x:x+0.24, y:1.92, d:0.4, color:c.color, label:c.mark });
    s.addText(c.title, { isTextBox:true, x:x+0.76, y:1.88, w:cw-1.0, h:0.5,
      fontFace:F, fontSize:13.5, bold:true, color:c.color, margin:0, valign:"middle" });
    s.addText(bullets(c.items, 10), { isTextBox:true, x:x+0.28, y:2.54, w:cw-0.56, h:3.52,
      fontFace:F, fontSize:12, color:INK, lineSpacingMultiple:1.32, margin:0, valign:"top" });
  });
  card(s, { x:M, y:6.34, w:CW, h:0.60, fill:NAVY });
  s.addText("「増えている」という事実は確認できた。同時に、次に何を確かめればよいかがはっきりした。", {
    isTextBox:true, x:M+0.28, y:6.34, w:CW-0.56, h:0.60,
    fontFace:F, fontSize:13.5, bold:true, color:WHITE, margin:0, valign:"middle", align:"center" });

  s.addNotes(
    "結果をどう読むかです。ここは意識して3つに分けました。\n"+
    "まず言えることは、前年の同じ4〜6月と比べてIBの利用件数は増えている、という事実です。\n"+
    "一方で、言えないこともあります。今回使えたデータは4〜6月の3か月分が2年分だけで、連続した推移が追えません。また、口座数の変化や取引全体の増減、季節性、社会全体でIB利用が広がっていること、といった施策以外の要因を切り分けることもできません。\n"+
    "ですので、「IB普及施策によって増えた」という因果関係は、このデータだけでは判断できません。ここは断定しないようにしました。\n"+
    "ただ、言えないことがはっきりしたことで、次に何を見ればよいかも具体的になりました。連続した月次推移で施策の前後を比べる、ATM側のIB代替可能取引が減っているかを対応させて見る、支店や年代などの属性別に見る、といったことです。\n"+
    "これらも、データが蓄積されれば、今回と同じやり方で確かめられると考えています。");
  pageNum(s, 7);
}

// ============ 8. 回収（一連の流れ） ============
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("この一連の流れを、専門的な分析知識なしで実施できた", {
    isTextBox:true, x:M, y:0.42, w:CW, h:0.62,
    fontFace:F, fontSize:26, bold:true, color:WHITE, margin:0, valign:"middle" });
  s.addText("人が考えるところと、DYNATREKが引き受けるところ。", {
    isTextBox:true, x:M, y:1.08, w:CW, h:0.4,
    fontFace:F, fontSize:13.5, bold:true, color:ACCENT_LT, margin:0, valign:"middle" });

  const flow = [
    { t:"IBの普及状況を\n知りたい", who:"人" },
    { t:"何と比べれば\nよいか考える", who:"人" },
    { t:"条件を指定して\n集計する", who:"DYNATREK" },
    { t:"出てきた結果を\n見る", who:"DYNATREK" },
    { t:"何が言えるかを\n考える", who:"人" },
  ];
  const bw = 2.1, bg = 0.28, startX = M + (CW - (5*bw + 4*bg))/2;
  flow.forEach((f,i)=>{
    const x = startX + i*(bw+bg);
    const isTool = f.who === "DYNATREK";
    card(s, { x:x, y:2.06, w:bw, h:1.95, fill: isTool?ACCENT:NAVY_MID, line: isTool?ACCENT:NAVY_SOFT });
    s.addText(f.who, { isTextBox:true, x:x, y:2.24, w:bw, h:0.3,
      fontFace:F, fontSize:10.5, bold:true, color: isTool?WHITE:ACCENT_LT, align:"center", valign:"middle", margin:0 });
    s.addText(f.t, { isTextBox:true, x:x+0.14, y:2.62, w:bw-0.28, h:1.1,
      fontFace:F, fontSize:12.5, bold:true, color:WHITE, lineSpacingMultiple:1.3, align:"center", valign:"middle", margin:0 });
    if (i < 4) s.addShape(pres.ShapeType.rightArrow, { x:x+bw+0.03, y:2.88, w:0.22, h:0.3,
      fill:{color:NAVY_SOFT}, line:{color:NAVY_SOFT,width:0.5} });
  });

  card(s, { x:M, y:4.34, w:CW, h:0.86, fill:NAVY_MID, line:NAVY_SOFT });
  s.addText("専門的な知識が要るのは、業務DBから必要なデータを取り出すところ。そこはDYNATREKが引き受けてくれた。", {
    isTextBox:true, x:M+0.3, y:4.34, w:CW-0.6, h:0.86,
    fontFace:F, fontSize:13, color:WHITE, margin:0, valign:"middle", align:"center" });

  card(s, { x:M, y:5.52, w:CW, h:1.24, fill:WHITE });
  s.addText("分析の専門知識や経験がなくても、DYNATREKを使えば、\n業務上の疑問に対してデータで確認・比較していける。", {
    isTextBox:true, x:M+0.3, y:5.52, w:CW-0.6, h:1.24,
    fontFace:F, fontSize:16, bold:true, color:NAVY, lineSpacingMultiple:1.35, margin:0, valign:"middle", align:"center" });

  s.addNotes(
    "今回やったことを、はじめから並べるとこの5ステップです。\n"+
    "IBの普及状況を知りたい。では何と比べればよいか考える。条件を指定して集計する。出てきた結果を見る。そこから何が言えるかを考える。\n"+
    "このうち、赤い2つ、条件を指定して集計するところと結果を出すところは、DYNATREKが引き受けてくれた部分です。専門的な知識が要るのはまさにここで、業務DBから必要なデータを取り出すところでした。\n"+
    "残りの3つ、何を知りたいのか、何と比べるのか、何が言えるのかは、人が考えるところです。私が時間を使ったのもここでした。\n"+
    "分析の専門知識や経験がなくても、DYNATREKを使えば、業務上の疑問に対してデータで確認・比較していける。今回、それを実際にやってみることができました。");
  pageNum(s, 8, true);
}

pres.writeFile({ fileName: process.argv[2] || "dynatrek_ib_presentation.pptx" })
  .then(f=>console.log("created:", f));
