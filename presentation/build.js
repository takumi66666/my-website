const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";                 // 13.333 x 7.5
pres.title = "銀行データを次の一手につなげるまで";

/* ---------- tokens ---------- */
const INK="1A2B33", INK2="465A63", MUTED="8497A0", LINE="D8E0E3", PANEL="F3F6F7";
const PRI="0F5A6B", PRIL="E2EEF0", ACC="B4551F", ACCL="F8EBE1", WHITE="FFFFFF";
const F="Meiryo";
const M=0.70, W=13.333-M*2, SW=13.333, SH=7.5, CX=M+0.64, CW=W-0.64;
const LX=M, LW=6.55, RX=7.62, RW=5.01;       // 共通の2カラム

/* ---------- helpers ---------- */
function slide(n, section, title, lead){
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape(pres.ShapeType.ellipse,{x:M,y:0.62,w:0.44,h:0.44,fill:{color:PRI},line:{color:PRI,width:0}});
  s.addText(String(n).padStart(2,"0"),{isTextBox:true,x:M,y:0.62,w:0.44,h:0.44,margin:0,
    align:"center",valign:"middle",fontFace:F,fontSize:11.5,bold:true,color:WHITE});
  s.addText(section,{isTextBox:true,x:CX,y:0.44,w:7,h:0.25,margin:0,
    fontFace:F,fontSize:10.5,bold:true,color:PRI,charSpacing:1.5});
  s.addText(title,{isTextBox:true,x:CX,y:0.66,w:CW,h:0.6,margin:0,
    fontFace:F,fontSize:26,bold:true,color:INK});
  if(lead) s.addText(lead,{isTextBox:true,x:CX,y:1.28,w:CW,h:0.38,margin:0,
    fontFace:F,fontSize:12.5,color:MUTED});
  s.addText(String(n).padStart(2,"0"),{isTextBox:true,x:SW-1.16,y:SH-0.70,w:0.6,h:0.28,margin:0,
    align:"right",fontFace:F,fontSize:10,color:MUTED,charSpacing:1});
  return s;
}
function card(s,o){
  s.addShape(pres.ShapeType.roundRect,{x:o.x,y:o.y,w:o.w,h:o.h,rectRadius:0.05,
    fill:{color:o.fill||PANEL},line:{color:o.line||(o.fill||PANEL),width:o.lw||0.75}});
}
function txt(s,t,o){ s.addText(t,Object.assign({isTextBox:true,margin:0,fontFace:F,fontSize:13,color:INK2},o)); }
function eyebrow(s,t,o){ txt(s,t,Object.assign({fontSize:10.5,bold:true,color:MUTED,charSpacing:1,h:0.26},o)); }
function dArrow(s,x,y,h){ s.addShape(pres.ShapeType.downArrow,{x:x,y:y,w:0.24,h:h,fill:{color:"C2CFD3"},line:{color:"C2CFD3",width:0}}); }
function rArrow(s,x,y,w){ s.addShape(pres.ShapeType.rightArrow,{x:x,y:y,w:w,h:0.24,fill:{color:"C2CFD3"},line:{color:"C2CFD3",width:0}}); }

function chartBox(s,o){                       // グレー点線＝図版を貼る
  s.addShape(pres.ShapeType.roundRect,{x:o.x,y:o.y,w:o.w,h:o.h,rectRadius:0.05,
    fill:{color:o.tint||PANEL},line:{color:o.border||"BFCCD1",width:1.25,dashType:"dash"}});
  txt(s,o.label||"グラフ貼付エリア",{x:o.x,y:o.y+o.h/2-0.42,w:o.w,h:0.32,align:"center",fontSize:13,bold:true,color:PRI});
  txt(s,o.sub,{x:o.x+0.3,y:o.y+o.h/2-0.04,w:o.w-0.6,h:0.75,align:"center",fontSize:11,color:MUTED,lineSpacing:17});
}
function blank(s,o){                          // オレンジ点線＝本人が書く
  s.addShape(pres.ShapeType.roundRect,{x:o.x,y:o.y,w:o.w,h:o.h,rectRadius:0.05,
    fill:{color:ACCL},line:{color:"D9A97F",width:1.25,dashType:"dash"}});
  txt(s,o.label,{x:o.x+0.24,y:o.y+0.15,w:o.w-0.48,h:0.26,fontSize:10.5,bold:true,color:ACC,charSpacing:0.8});
  txt(s,o.hint,{x:o.x+0.24,y:o.y+0.44,w:o.w-0.48,h:o.h-0.58,fontSize:11.5,color:"A2795C",lineSpacing:17});
}
function band(s,o){
  const x=o.x||M, w=o.w||W, h=o.h||0.72;
  s.addShape(pres.ShapeType.roundRect,{x:x,y:o.y,w:w,h:h,rectRadius:0.05,
    fill:{color:o.fill||PRIL},line:{color:o.fill||PRIL,width:0}});
  txt(s,o.text,{x:x+0.3,y:o.y,w:w-0.6,h:h,align:"center",valign:"middle",
    fontSize:o.size||14,bold:true,color:o.color||PRI,lineSpacing:o.ls});
}

/* ============ 01 はじめに ============ */
{
  const s = slide(1,"は じ め に","データはあっても答えはすぐに出ない",
    "本インターンのテーマ ── 銀行内で保有する各種データを分析し、どのような価値に変換していくかのプロセスを学ぶ");

  const ch=1.05, ys=[2.05,3.38,4.71];
  const rows=[
    ["現 状","銀行には、日々の取引データが蓄積されている"],
    ["課 題","項目が多く、知りたい切り口でまとめて取り出すのは容易ではない"],
    ["道 具","DYNATREK ── 条件を指定するだけで、必要なデータを集計・表示できる"]
  ];
  rows.forEach((r,i)=>{
    const hl = i===2;
    card(s,{x:LX,y:ys[i],w:LW,h:ch,fill:hl?PRIL:PANEL});
    eyebrow(s,r[0],{x:LX+0.3,y:ys[i]+0.15,w:LW-0.6,color:hl?PRI:MUTED});
    txt(s,r[1],{x:LX+0.3,y:ys[i]+0.44,w:LW-0.6,h:0.55,fontSize:13.5,bold:hl,
      color:hl?PRI:INK,lineSpacing:19});
    if(i<2) dArrow(s,LX+0.55,ys[i]+ch+0.03,0.26);
  });

  card(s,{x:RX,y:2.05,w:RW,h:3.71,fill:WHITE,line:"D9A97F",lw:1.25});
  eyebrow(s,"例 え ば",{x:RX+0.42,y:2.42,w:RW-0.84,color:ACC});
  txt(s,"預金の増減はどうか\nIBはどれだけ使われているか",
    {x:RX+0.42,y:2.78,w:RW-0.84,h:1.1,fontSize:18,bold:true,color:ACC,lineSpacing:32});
  s.addShape(pres.ShapeType.rect,{x:RX+0.42,y:4.14,w:RW-0.84,h:0.012,
    fill:{color:"E2CDB9"},line:{color:"E2CDB9",width:0}});
  txt(s,"本ワークでは、こうした問いを一つ立て、\nデータから答えを出すまでを実際に通した",
    {x:RX+0.42,y:4.48,w:RW-0.84,h:0.95,fontSize:13.5,color:INK,lineSpacing:22});

  s.addNotes("本インターンのテーマは、銀行が持っているデータを分析して、どのような価値に変えていくかのプロセスを学ぶことでした。\n銀行には日々の取引データが蓄積されていますが、項目が多く、知りたい切り口でまとめて取り出すのは簡単ではありません。そこで使ったのがDYNATREKです。条件を指定するだけで、必要なデータを集計して表示できます。\n「預金の増減はどうか」「IBはどれだけ使われているか」──こうした問いを一つ立てて、データから答えを出すところまで実際に通してみました。その流れを紹介します。");
}

/* ============ 02 背景 ============ */
{
  const s = slide(2,"背 景","非対面チャネルの利用拡大という狙い",
    "数ある問いのうち、銀行が実際に力を入れている領域からテーマを選んだ");

  const lw=6.9, ys=[2.05,3.10,4.15];
  const rows=[
    ["計 画","あいちフィナンシャルグループ　第2次中期経営計画（2025年4月〜2028年3月）",0],
    ["基本戦略","Ⅲ「DX戦略の加速化」",0.35],
    ["重点施策","⑦　〔　計画本文の表現を記入　〕",0.70]
  ];
  rows.forEach((r,i)=>{
    const x=M+r[2], w=lw-r[2], hl=i===2;
    card(s,{x:x,y:ys[i],w:w,h:0.85,fill:hl?ACCL:PANEL,line:hl?"D9A97F":PANEL,lw:hl?1.25:0.75});
    eyebrow(s,r[0],{x:x+0.28,y:ys[i]+0.13,w:1.8,fontSize:10,color:hl?ACC:MUTED});
    txt(s,r[1],{x:x+0.28,y:ys[i]+0.40,w:w-0.56,h:0.34,fontSize:13,bold:hl,color:hl?ACC:INK});
    if(i<2) dArrow(s,x+0.5,ys[i]+0.88,0.20);
  });
  txt(s,"非対面チャネル（バンキングアプリ、インターネット支店 など）の有効活用",
    {x:M+0.70,y:5.06,w:lw-0.70,h:0.32,fontSize:12,color:INK2});

  const rx=7.95, rw=4.68;
  card(s,{x:rx,y:2.05,w:rw,h:2.95,fill:PANEL});
  eyebrow(s,"非対面チャネル",{x:rx+0.3,y:2.25,w:rw-0.6,fontSize:12.5});
  [["バンキングアプリ",false],["インターネット支店",false],["インターネットバンキング（IB）",true]].forEach((it,i)=>{
    const y=2.68+i*0.72;
    card(s,{x:rx+0.3,y:y,w:rw-0.6,h:0.6,fill:it[1]?PRI:WHITE,line:it[1]?PRI:LINE,lw:1});
    txt(s,it[0],{x:rx+0.3,y:y,w:rw-0.6,h:0.6,align:"center",valign:"middle",
      fontSize:it[1]?12.5:12,bold:it[1],color:it[1]?WHITE:INK2});
  });
  txt(s,"↑　今回の分析対象",{x:rx,y:5.06,w:rw,h:0.3,align:"center",fontSize:11.5,bold:true,color:PRI});

  band(s,{y:5.72,text:"非対面チャネルがどれだけ使われるようになったかは、計画の狙いがどこまで進んでいるかを示す手がかりになる"});

  s.addNotes("では、どの問いを選んだか。銀行が実際に力を入れている領域から選びました。\nあいちフィナンシャルグループの第2次中期経営計画では、基本戦略のひとつに「DX戦略の加速化」が掲げられ、その重点施策として、バンキングアプリやインターネット支店といった非対面チャネルの有効活用が挙げられています。\n非対面チャネルのひとつがインターネットバンキング、IBです。これがどれだけ使われるようになったかは、計画の狙いがどこまで進んでいるかを示す手がかりになります。\n\n【作成時】重点施策の番号と本文は計画の原本で確認して記入する。確認できなければ「重点施策のひとつ」と書けば事実関係は崩れない。KPIの数値には触れない。");
}

/* ============ 03 目的 ============ */
{
  const s = slide(3,"目 的","IBの利用がどれだけ広がったかを確かめる",
    "背景をふまえ、今回のワークで確かめることを次のように定めた");

  card(s,{x:M,y:2.1,w:W,h:1.55,fill:PRIL});
  eyebrow(s,"目 的",{x:M+0.45,y:2.32,w:2,charSpacing:1.5,color:PRI});
  txt(s,"2025年と2026年の同じ時期を比べ、IBの利用件数がどれだけ増えたかを確認する",
    {x:M+0.45,y:2.68,w:W-0.9,h:0.6,fontSize:19,bold:true,color:PRI});

  const cw=(W-0.5)/2, cy=4.05, ch=1.75;
  [["比 べ 方","同じ4〜6月どうしを比べ、\n時期による取引量の差を避ける"],
   ["見 る 対 象","IBの件数に加え、\nATM側の「IBでも代替できる取引」も見る"]
  ].forEach((c,i)=>{
    const x=M+i*(cw+0.5);
    card(s,{x:x,y:cy,w:cw,h:ch});
    eyebrow(s,c[0],{x:x+0.35,y:cy+0.25,w:cw-0.7});
    txt(s,c[1],{x:x+0.35,y:cy+0.62,w:cw-0.7,h:0.9,fontSize:13.5,color:INK,lineSpacing:22});
  });

  s.addNotes("背景をふまえて、確かめることをこう定めました。2025年と2026年の同じ時期を比べて、IBの利用件数がどれだけ増えたかを確認することです。\n比べ方は同じ4〜6月どうし。月によって取引の量は変わるので、同じ季節どうしで比べます。\nもうひとつ、IBの件数だけでなく、ATMで行われている「IBでも代替できる取引」も合わせて見ることにしました。理由は結果のところで説明します。");
}

/* ============ 04 使用データ ============ */
{
  const s = slide(4,"使 用 デ ー タ","取引履歴（流動）から入出金テーブルを作成",
    "分析に使ったデータと、DYNATREKで扱える形にするまでの準備");

  const cw=3.55, gp=0.64, fy=2.05, fh=1.62;
  const xs=[M, M+cw+gp, M+(cw+gp)*2];
  [["元データ","取引履歴（流動）",PANEL,INK],
   ["作成したもの","入出金テーブル",PANEL,INK],
   ["分 析","DYNATREKで\n条件を指定して集計",PRIL,PRI]
  ].forEach((f,i)=>{
    card(s,{x:xs[i],y:fy,w:cw,h:fh,fill:f[2]});
    eyebrow(s,f[0],{x:xs[i],y:fy+0.28,w:cw,align:"center",color:i===2?PRI:MUTED});
    txt(s,f[1],{x:xs[i]+0.2,y:fy+0.68,w:cw-0.4,h:0.7,align:"center",fontSize:15,bold:true,
      color:f[3],lineSpacing:22});
    if(i<2) rArrow(s,xs[i]+cw+0.18,fy+fh/2-0.12,0.28);
  });

  card(s,{x:M,y:4.05,w:W,h:1.2});
  eyebrow(s,"対 象 期 間",{x:M+0.4,y:4.28,w:1.9});
  [["2025年　4月・5月・6月",3.35],["2026年　4月・5月・6月",7.65]].forEach(p=>{
    card(s,{x:p[1],y:4.28,w:3.9,h:0.72,fill:WHITE,line:LINE,lw:1});
    txt(s,p[0],{x:p[1],y:4.28,w:3.9,h:0.72,align:"center",valign:"middle",fontSize:14,bold:true,color:INK});
  });

  blank(s,{x:M,y:5.6,w:W,h:1.0,label:"記 入 欄",
    hint:"入出金テーブルに持たせた主な項目（例：日付／チャネル／取引種別／年代 など、実際に使ったものを記載）"});

  s.addNotes("使ったデータは、2025年と2026年の4月・5月・6月です。\n元になったのは取引履歴（流動）で、そこから入出金テーブルを作成し、DYNATREKで条件を指定して集計できる形にしました。\n取引履歴そのままでは項目が多く、目的のデータを追いにくい。逆に、一度この形にしてしまえば、以降は条件を指定するだけでいろいろな角度から見られるようになります。");
}

/* ============ 05 分析方法 ============ */
{
  const s = slide(5,"分 析 方 法","条件を指定して必要な集計を取り出す",
    "作成した入出金テーブルに対し、画面上で条件を指定するだけで集計結果が得られる");

  const fw=5.50, fy=2.05, fh=3.20, fx1=M, fx2=SW-M-fw;
  chartBox(s,{x:fx1,y:fy,w:fw,h:fh,label:"スクリーンショット ①",
    sub:"DYNATREKの条件設定画面\n（期間・チャネル・取引種別などを指定）"});
  chartBox(s,{x:fx2,y:fy,w:fw,h:fh,label:"スクリーンショット ②",
    sub:"指定した条件で出力された集計結果",tint:PRIL,border:"A8C7CC"});
  s.addShape(pres.ShapeType.rightArrow,{x:6.34,y:fy+fh/2-0.20,w:0.65,h:0.40,
    fill:{color:PRI},line:{color:PRI,width:0}});
  txt(s,"条件を指定する",{x:fx1,y:fy+fh+0.16,w:fw,h:0.30,align:"center",fontSize:13,bold:true,color:INK});
  txt(s,"そのまま集計結果が出る",{x:fx2,y:fy+fh+0.16,w:fw,h:0.30,align:"center",fontSize:13,bold:true,color:PRI});

  band(s,{y:5.95,h:0.62,fill:PANEL,color:INK,
    text:"一度テーブルを整えれば、条件を変えるだけで別の切り口でも確認できる"});

  s.addNotes("実際の画面です。左で、期間やチャネル、取引の種類といった条件を指定します。すると右のように、その条件で集計された結果がそのまま出てきます。\nここは長く説明しません。お伝えしたいのは、一度テーブルを整えてしまえば、あとは条件を変えるだけで別の切り口でも確認できる、という点です。この後の分析は、すべてこの繰り返しになります。\n\n【作成時】スクリーンショット②が結果スライドのグラフと同じものになる場合は、②には集計表など別の出力を貼るか、②を小さめに配置して重複感を避ける。");
}

/* ============ 06 結果① ============ */
{
  const s = slide(6,"結 果 ①","IBの利用件数は増加していた",
    "まず、IBで行われた取引の件数を2つの期間で比較");

  chartBox(s,{x:LX,y:2.05,w:LW,h:3.5,sub:"IBの利用件数（2025年4〜6月 / 2026年4〜6月）"});
  blank(s,{x:RX,y:2.05,w:RW,h:1.05,label:"件 数",
    hint:"2025年4〜6月 合計 ◯◯件 → 2026年4〜6月 合計 ◯◯件"});
  blank(s,{x:RX,y:3.28,w:RW,h:1.05,label:"変 化",hint:"1年間で ＋◯◯件（＋◯◯％）"});
  blank(s,{x:RX,y:4.51,w:RW,h:1.04,label:"月 別 の 傾 向",
    hint:"4月・5月・6月それぞれ増えているか／特に伸びた月はどこか"});

  band(s,{y:5.85,h:0.68,fill:PANEL,color:INK,
    text:"ただし件数の増加だけでは、取引そのものが増えた可能性を否定できない"});

  s.addNotes("まずIBの利用件数です。2025年の4〜6月と、2026年の4〜6月を比べました。\n（記入した数値を読み上げる）\n1年間で件数は増えています。月別に見ても同じ傾向です。\nただ、これだけでIBが広がったと言い切ることはできません。取引そのものが増えただけ、という可能性が残るからです。そこで、もうひとつの見方を足しました。");
}

/* ============ 07 結果② ============ */
{
  const s = slide(7,"結 果 ②","ATMでIBに代替できる取引は減少していた",
    "IBへ移っているのかどうかを判断するため、ATM側の同種の取引も確認");

  chartBox(s,{x:LX,y:2.05,w:LW,h:3.5,sub:"ATMでのIB代替可能な取引の件数（2025年4〜6月 / 2026年4〜6月）"});

  card(s,{x:RX,y:2.05,w:RW,h:1.05});
  eyebrow(s,"対 象",{x:RX+0.28,y:2.20,w:2,fontSize:10});
  txt(s,"ATMでの取引のうち、IBでも行えるもの\n（振込・振替 など）",
    {x:RX+0.28,y:2.46,w:RW-0.56,h:0.55,fontSize:12,color:INK,lineSpacing:17});
  blank(s,{x:RX,y:3.28,w:RW,h:1.05,label:"件 数",
    hint:"2025年4〜6月 合計 ◯◯件 → 2026年4〜6月 合計 ◯◯件"});
  blank(s,{x:RX,y:4.51,w:RW,h:1.04,label:"変 化",hint:"1年間で −◯◯件（−◯◯％）"});

  band(s,{y:5.85,h:0.68,text:"IBは増加、ATM側は減少　→　同じ取引がIBへ移りつつあると読める"});

  s.addNotes("IBでも同じことができる取引──振込や振替など──に絞って、ATM側の件数を見ました。\n（記入した数値を読み上げる）\nこちらは減っています。IBが増え、ATM側が減っている。つまり、同じ取引がIBへ移りつつあると読めます。\n片方だけを見ていたら「増えた」で終わっていましたが、比べる相手を置いたことで、一歩踏み込んだ言い方ができるようになりました。\n\n※「IBでも行える取引」に何を含めたかは質疑で聞かれやすいので、口頭でも一度言っておく。");
}

/* ============ 08 考察 ============ */
{
  const s = slide(8,"考 察","どの年代で利用が広がったのか",
    "全体の動きが分かった次に知りたくなるのは「誰が使うようになったのか」。今回は年代を5つに分けて確認");

  const gw=2.18, gp=0.25;
  [["未成年","0〜17歳"],["若手成人","18〜29歳"],["働き手①","30〜44歳"],
   ["働き手②","45〜64歳"],["高齢者","65歳以上"]].forEach((a,i)=>{
    const x=M+i*(gw+gp);
    card(s,{x:x,y:2.05,w:gw,h:0.74});
    txt(s,a[0],{x:x,y:2.15,w:gw,h:0.28,align:"center",fontSize:12.5,bold:true,color:INK});
    txt(s,a[1],{x:x,y:2.45,w:gw,h:0.26,align:"center",fontSize:11,color:MUTED});
  });

  chartBox(s,{x:LX,y:2.98,w:LW,h:2.55,sub:"年代別 IB利用件数の増加率（棒グラフ）"});

  blank(s,{x:RX,y:2.98,w:RW,h:0.95,label:"読 み 取 り",
    hint:"どの年代の増加率が最も高かったか／他の年代はどうだったか"});

  card(s,{x:RX,y:4.13,w:RW,h:1.40,fill:PRIL});
  eyebrow(s,"こ こ か ら 考 え ら れ る 打 ち 手",{x:RX+0.28,y:4.28,w:RW-0.56,color:PRI});
  s.addText([
    {text:"伸びが小さい年代には、ATMでの振込・振替の際にIBを案内する",options:{bullet:true,breakLine:true}},
    {text:"すでに伸びている年代には、扱える手続きの幅を広げる",options:{bullet:true}}
  ],{isTextBox:true,x:RX+0.28,y:4.56,w:RW-0.56,h:0.9,margin:0,fontFace:F,fontSize:11.5,
     color:PRI,lineSpacing:16,paraSpaceAfter:6});

  band(s,{y:5.78,h:0.68,fill:PANEL,color:INK,
    text:"年代まで分かると、誰にどう案内するかという具体的な打ち手まで検討できる"});

  s.addNotes("IBへ移りつつあることが分かると、次に知りたくなるのは「では、誰が使うようになったのか」です。これも同じテーブルに年代の条件を足すだけで確認できました。\n年代は5つに分けています。未成年、若手成人、働き手①、働き手②、高齢者です。\n（読み取りを説明する）\nここまで分かると、伸びていない年代にどう案内するか、といった具体的な打ち手の話ができるようになります。\n\n【作成時】打ち手の2行は、実際の読み取りに合わせて書き換えること。");
}

/* ============ 09 まとめ ============ */
{
  const s = slide(9,"ま と め","データを次の一手につなげるまで",
    "今回のワークで行ったことを、ひとつの流れにまとめる");

  const cw=2.11, gp=0.34, cy=2.15, ch=2.45;
  [["問 い を 立 て る","「IBの利用は\n広がっているか」"],
   ["整 え る","取引履歴（流動）から\n入出金テーブルを作成"],
   ["確 か め る","IBは増加\nATM側は減少"],
   ["深 掘 る","年代別に\n増加率を確認"],
   ["次 の 一 手","案内すべき年代と\n方法を検討"]
  ].forEach((st,i)=>{
    const x=M+i*(cw+gp), last=i===4;
    card(s,{x:x,y:cy,w:cw,h:ch,fill:last?PRI:PANEL});
    txt(s,String(i+1),{x:x,y:cy+0.28,w:cw,h:0.34,align:"center",fontSize:16,bold:true,
      color:last?"9CC8CF":"B9C7CC"});
    txt(s,st[0],{x:x+0.06,y:cy+0.78,w:cw-0.12,h:0.3,align:"center",fontSize:11.5,bold:true,
      color:last?WHITE:INK,charSpacing:0.4});
    txt(s,st[1],{x:x+0.12,y:cy+1.22,w:cw-0.24,h:0.9,align:"center",fontSize:11.5,
      color:last?"D6E8EB":INK2,lineSpacing:17});
    if(i<4) rArrow(s,x+cw+0.05,cy+ch/2-0.12,0.24);
  });

  band(s,{y:5.05,h:0.95,size:15,
    text:"行内にあるデータも、条件を指定して見られる形に整えれば、業務上の問いに答え、次の施策を考える材料になる"});
  txt(s,"今回のワークで学んだのは、この一連のプロセス",
    {x:M,y:6.22,w:W,h:0.32,align:"center",fontSize:12.5,color:MUTED});

  s.addNotes("最後に、今回の流れをまとめます。\n「IBの利用は広がっているか」という問いを立て、取引履歴（流動）から入出金テーブルを整え、IBが増えATM側が減っていることを確かめ、さらに年代別まで深掘りして、最後は誰にどう案内するかという打ち手の検討まで進みました。\n行内にあるデータも、条件を指定して見られる形に整えれば、業務上の問いに答えて、次の施策を考える材料になります。今回のワークで学んだのは、この一連のプロセスです。");
}

pres.writeFile({fileName:"/tmp/claude-0/-home-user-my-website/9d2b3804-e19b-5dbd-a9cd-4f44a2f53f65/scratchpad/DYNATREK_成果発表.pptx"})
  .then(f=>console.log("wrote",f));
