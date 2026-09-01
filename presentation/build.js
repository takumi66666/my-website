const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";                 // 13.33 x 7.5
pres.title = "銀行データを価値に変えるプロセス";

/* ---------- tokens ---------- */
const INK="1A2B33", INK2="465A63", MUTED="8497A0", LINE="D8E0E3", PANEL="F3F6F7";
const PRI="0F5A6B", PRIL="E2EEF0", ACC="B4551F", ACCL="F8EBE1", WHITE="FFFFFF";
const F="Meiryo";
const M=0.70, W=13.333-M*2, SW=13.333, SH=7.5, CX=M+0.64, CW=W-0.64;

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
function dArrow(s,x,y,h,c){ s.addShape(pres.ShapeType.downArrow,{x:x,y:y,w:0.24,h:h,fill:{color:c||"C2CFD3"},line:{color:c||"C2CFD3",width:0}}); }
function rArrow(s,x,y,w,c){ s.addShape(pres.ShapeType.rightArrow,{x:x,y:y,w:w,h:0.24,fill:{color:c||"C2CFD3"},line:{color:c||"C2CFD3",width:0}}); }

/* grey dashed frame = paste a chart/image here */
function chartBox(s,o){
  s.addShape(pres.ShapeType.roundRect,{x:o.x,y:o.y,w:o.w,h:o.h,rectRadius:0.05,
    fill:{color:PANEL},line:{color:"BFCCD1",width:1.25,dashType:"dash"}});
  s.addText(o.label,{isTextBox:true,x:o.x,y:o.y+o.h/2-0.44,w:o.w,h:0.34,margin:0,
    align:"center",fontFace:F,fontSize:13,bold:true,color:PRI});
  s.addText(o.sub,{isTextBox:true,x:o.x+0.3,y:o.y+o.h/2-0.04,w:o.w-0.6,h:0.8,margin:0,
    align:"center",fontFace:F,fontSize:11,color:MUTED,lineSpacing:17});
}
/* orange dashed field = write your reading here */
function blank(s,o){
  s.addShape(pres.ShapeType.roundRect,{x:o.x,y:o.y,w:o.w,h:o.h,rectRadius:0.05,
    fill:{color:ACCL},line:{color:"D9A97F",width:1.25,dashType:"dash"}});
  s.addText(o.label,{isTextBox:true,x:o.x+0.24,y:o.y+0.15,w:o.w-0.48,h:0.26,margin:0,
    fontFace:F,fontSize:10.5,bold:true,color:ACC,charSpacing:0.8});
  s.addText(o.hint,{isTextBox:true,x:o.x+0.24,y:o.y+0.44,w:o.w-0.48,h:o.h-0.58,margin:0,
    fontFace:F,fontSize:11.5,color:"A2795C",lineSpacing:17});
}
function band(s,o){
  s.addShape(pres.ShapeType.roundRect,{x:o.x||M,y:o.y,w:o.w||W,h:o.h||0.72,rectRadius:0.05,
    fill:{color:o.fill||PRIL},line:{color:o.fill||PRIL,width:0}});
  s.addText(o.text,{isTextBox:true,x:(o.x||M)+0.3,y:o.y,w:(o.w||W)-0.6,h:o.h||0.72,margin:0,
    align:"center",valign:"middle",fontFace:F,fontSize:o.size||14,bold:true,color:o.color||PRI,lineSpacing:o.ls});
}

/* ===================== 01 きっかけ ===================== */
{
  const s = slide(1,"き っ か け","「預金の増減を調べたい」という相談から",
    "本インターンのテーマ：銀行内で保有する各種データを分析し、どのような価値に変換していくかのプロセスを学ぶ");

  const lw=6.55, ch=1.05, ys=[2.05,3.38,4.71];
  const steps=[
    ["あいち銀行様からのご相談","「預金の増減を調べたい」"],
    ["ところが","行内のデータは項目が多く、見たい条件でまとめて取り出す手段がない"],
    ["そこで","DYNATREK ── 条件を指定して、必要なデータを集計・表示できるツール"]
  ];
  steps.forEach((st,i)=>{
    card(s,{x:M,y:ys[i],w:lw,h:ch,fill:i===2?PRIL:PANEL});
    txt(s,st[0],{x:M+0.3,y:ys[i]+0.15,w:lw-0.6,h:0.24,fontSize:10.5,bold:true,
      color:i===2?PRI:MUTED,charSpacing:1});
    txt(s,st[1],{x:M+0.3,y:ys[i]+0.44,w:lw-0.6,h:0.5,fontSize:13.5,bold:i===2,
      color:i===2?PRI:INK,lineSpacing:19});
    if(i<2) dArrow(s,M+0.55,ys[i]+ch+0.03,0.26);
  });

  const rx=7.62, rw=5.01;
  card(s,{x:rx,y:2.05,w:rw,h:3.71,fill:WHITE,line:"D9A97F",lw:1.25});
  txt(s,"例えば、\nどんなことが\nできるのか？",{x:rx+0.4,y:2.45,w:rw-0.8,h:1.75,
    fontSize:23,bold:true,color:ACC,lineSpacing:38});
  s.addShape(pres.ShapeType.rect,{x:rx+0.4,y:4.32,w:rw-0.8,h:0.012,fill:{color:"E2CDB9"},line:{color:"E2CDB9",width:0}});
  txt(s,"本発表では、その一例として\n私が行った分析を紹介します",{x:rx+0.4,y:4.62,w:rw-0.8,h:0.8,
    fontSize:14,color:INK,lineSpacing:24});

  s.addNotes("インターンのテーマは、銀行が持っているデータを分析して価値に変えるプロセスを学ぶことでした。\n出発点は、あいち銀行様の「預金の増減を調べたい」というご相談です。ただ、行内のデータは項目が多く、見たい条件でまとめて取り出すのが簡単ではありません。そこで使ったのがDYNATREKです。\n「では、実際にどんなことができるのか」──本発表では、その一例として私が行った分析を紹介します。");
}

/* ===================== 02 背景 ===================== */
{
  const s = slide(2,"背 景","非対面チャネルの利用拡大は、経営計画上の狙いでもある",
    "今回のテーマは、あいち銀行様が経営計画で掲げている方向と重なっています。");

  const lw=6.9, ys=[2.05,3.10,4.15];
  const rows=[
    ["計 画","あいちフィナンシャルグループ　第2次中期経営計画（2025年4月〜2028年3月）",0],
    ["基本戦略","Ⅲ「DX戦略の加速化」",0.35],
    ["重点施策","⑦　〔　計画本文の表現を記入　〕",0.70]
  ];
  rows.forEach((r,i)=>{
    const x=M+r[2], w=lw-r[2];
    card(s,{x:x,y:ys[i],w:w,h:0.85,fill:i===2?ACCL:PANEL,line:i===2?"D9A97F":PANEL,lw:i===2?1.25:0.75});
    txt(s,r[0],{x:x+0.28,y:ys[i]+0.13,w:1.6,h:0.24,fontSize:10,bold:true,
      color:i===2?ACC:MUTED,charSpacing:1});
    txt(s,r[1],{x:x+0.28,y:ys[i]+0.40,w:w-0.56,h:0.34,fontSize:13,bold:i===2,color:i===2?ACC:INK});
    if(i<2) dArrow(s,M+r[2]+0.5,ys[i]+0.88,0.20);
  });
  txt(s,"非対面チャネル（バンキングアプリ、インターネット支店 など）の有効活用",
    {x:M+0.70,y:5.06,w:lw-0.70,h:0.32,fontSize:12,color:INK2});

  const rx=7.95, rw=4.68;
  card(s,{x:rx,y:2.05,w:rw,h:2.95,fill:PANEL});
  txt(s,"非対面チャネル",{x:rx+0.3,y:2.25,w:rw-0.6,h:0.3,fontSize:12.5,bold:true,color:MUTED,charSpacing:1});
  const items=[["バンキングアプリ",false],["インターネット支店",false],["インターネットバンキング（IB）",true]];
  items.forEach((it,i)=>{
    const y=2.68+i*0.72;
    card(s,{x:rx+0.3,y:y,w:rw-0.6,h:0.6,fill:it[1]?PRI:WHITE,line:it[1]?PRI:LINE,lw:1});
    txt(s,it[0],{x:rx+0.3,y:y,w:rw-0.6,h:0.6,align:"center",valign:"middle",
      fontSize:it[1]?12.5:12,bold:it[1],color:it[1]?WHITE:INK2});
  });
  txt(s,"↑　今回の分析対象",{x:rx,y:5.06,w:rw,h:0.3,align:"center",fontSize:11.5,bold:true,color:PRI});

  band(s,{y:5.72,text:"非対面チャネルがどれだけ使われるようになったかは、計画の狙いがどこまで進んでいるかを示す手がかりになる"});

  s.addNotes("あいちフィナンシャルグループの第2次中期経営計画では、基本戦略のひとつに「DX戦略の加速化」が掲げられ、その重点施策として、バンキングアプリやインターネット支店といった非対面チャネルの有効活用が挙げられています。\nその非対面チャネルのひとつがインターネットバンキング、IBです。\n\n【作成時の注意】重点施策の番号（⑦）と本文の表現は、計画の原本で確認してから記入すること。番号が確認できない場合は「重点施策のひとつ」と書けば事実関係は崩れない。\n※KPIの数値には触れない（今回はKPIに沿った分析をしていないため）。");
}

/* ===================== 03 目的 ===================== */
{
  const s = slide(3,"目 的","IBの利用がどれだけ増えたかを、データで確かめる",
    "背景をふまえて、今回の分析の目的を次のように置きました。");

  card(s,{x:M,y:2.1,w:W,h:1.55,fill:PRIL});
  txt(s,"目 的",{x:M+0.45,y:2.32,w:2,h:0.26,fontSize:10.5,bold:true,color:PRI,charSpacing:1.5});
  txt(s,"2025年と2026年の同じ時期を比べ、IBの利用件数がどれだけ増えたかを確認する",
    {x:M+0.45,y:2.68,w:W-0.9,h:0.6,fontSize:19,bold:true,color:PRI});

  const cw=(W-0.5)/2, cy=4.05, ch=1.75;
  const subs=[
    ["比べ方","同じ4〜6月どうしを比べることで、時期による取引量の差の影響を避ける"],
    ["見る対象","IBの件数だけでなく、ATMで行われている「IBでも代替できる取引」も合わせて見る"]
  ];
  subs.forEach((c,i)=>{
    const x=M+i*(cw+0.5);
    card(s,{x:x,y:cy,w:cw,h:ch,fill:PANEL});
    txt(s,c[0],{x:x+0.35,y:cy+0.25,w:cw-0.7,h:0.26,fontSize:10.5,bold:true,color:MUTED,charSpacing:1});
    txt(s,c[1],{x:x+0.35,y:cy+0.62,w:cw-0.7,h:0.9,fontSize:13.5,color:INK,lineSpacing:22});
  });

  s.addNotes("背景をふまえて、目的はシンプルに置きました。2025年と2026年の同じ時期を比べて、IBの利用件数がどれだけ増えたかを確認することです。\n比べ方は、同じ4〜6月どうし。時期によって取引の量は変わるので、同じ季節どうしで比べます。\nもうひとつ、IBの件数だけを見るのではなく、ATMで行われている「IBでもできる取引」も合わせて見ることにしました。理由は結果のところで説明します。");
}

/* ===================== 04 使用したデータ ===================== */
{
  const s = slide(4,"使 用 し た デ ー タ","取引履歴（流動）から入出金テーブルを作成した",
    "分析に使ったデータと、DYNATREKで扱えるようにするまでの準備です。");

  const cw=3.55, gp=0.64, fy=2.05, fh=1.62;
  const xs=[M, M+cw+gp, M+(cw+gp)*2];
  const flow=[
    ["元データ","取引履歴（流動）",PANEL,INK],
    ["作成したもの","入出金テーブル",PANEL,INK],
    ["分析","DYNATREKで\n条件を指定して集計",PRIL,PRI]
  ];
  flow.forEach((f,i)=>{
    card(s,{x:xs[i],y:fy,w:cw,h:fh,fill:f[2]});
    txt(s,f[0],{x:xs[i],y:fy+0.28,w:cw,h:0.26,align:"center",fontSize:10.5,bold:true,
      color:i===2?PRI:MUTED,charSpacing:1});
    txt(s,f[1],{x:xs[i]+0.2,y:fy+0.68,w:cw-0.4,h:0.7,align:"center",fontSize:15,bold:true,
      color:f[3],lineSpacing:22});
    if(i<2) rArrow(s,xs[i]+cw+0.18,fy+fh/2-0.12,0.28);
  });

  card(s,{x:M,y:4.05,w:W,h:1.2,fill:PANEL});
  txt(s,"対 象 期 間",{x:M+0.4,y:4.28,w:1.8,h:0.26,fontSize:10.5,bold:true,color:MUTED,charSpacing:1});
  [["2025年　4月・5月・6月",3.35],["2026年　4月・5月・6月",7.65]].forEach(p=>{
    card(s,{x:p[1],y:4.28,w:3.9,h:0.72,fill:WHITE,line:LINE,lw:1});
    txt(s,p[0],{x:p[1],y:4.28,w:3.9,h:0.72,align:"center",valign:"middle",fontSize:14,bold:true,color:INK});
  });

  blank(s,{x:M,y:5.6,w:W,h:1.0,label:"記 入 欄",
    hint:"入出金テーブルに持たせた主な項目（例：日付／チャネル／取引種別／年代 など、実際に使ったものを記載）"});

  s.addNotes("使ったデータは、2025年と2026年の4月・5月・6月です。\n元になったのは取引履歴（流動）で、そこから入出金テーブルを作成し、DYNATREKで条件を指定して集計できる形にしました。\n取引履歴そのままだと項目が多く目的のデータを追いにくいので、この準備が必要でした。逆に、一度この形にしてしまえば、以降は条件を指定するだけで色々な角度から見られるようになります。");
}

/* ===================== 05 結果① IB ===================== */
{
  const s = slide(5,"結 果 ①","IBの利用件数は増えていた",
    "まず、IBで行われた取引の件数を、2つの期間で比べました。");

  chartBox(s,{x:M,y:2.05,w:6.55,h:3.5,label:"グラフ貼付エリア",
    sub:"IBの利用件数（2025年4〜6月 / 2026年4〜6月）"});

  const rx=7.62, rw=5.01;
  blank(s,{x:rx,y:2.05,w:rw,h:1.05,label:"件 数",
    hint:"2025年4〜6月 合計 ◯◯件 → 2026年4〜6月 合計 ◯◯件"});
  blank(s,{x:rx,y:3.28,w:rw,h:1.05,label:"変 化",
    hint:"1年間で ＋◯◯件（＋◯◯％）"});
  blank(s,{x:rx,y:4.51,w:rw,h:1.04,label:"月 別 の 傾 向",
    hint:"4月・5月・6月それぞれ増えているか／特に伸びた月はどこか"});

  band(s,{y:5.85,h:0.68,text:"IBの利用件数は、1年間で増加していた"});

  s.addNotes("まずIBの利用件数です。2025年の4〜6月と、2026年の4〜6月を比べました。\n（記入した数値を読み上げる）\n1年間で件数は増えています。月別に見ても、同じ傾向が出ています。\n──ただ、これだけで「IBが広がった」と言い切れるかというと、そうではありません。次のスライドで、もうひとつの見方を足します。");
}

/* ===================== 06 結果② ATM ===================== */
{
  const s = slide(6,"結 果 ②","ATMのうち、IBでもできる取引は減っていた",
    "件数が増えた理由が「取引全体の増加」ではないことを確かめるため、ATM側も見ました。");

  chartBox(s,{x:M,y:2.05,w:6.55,h:3.5,label:"グラフ貼付エリア",
    sub:"ATMでのIB代替可能な取引の件数（2025年4〜6月 / 2026年4〜6月）"});

  const rx=7.62, rw=5.01;
  card(s,{x:rx,y:2.05,w:rw,h:1.05,fill:PANEL});
  txt(s,"対 象",{x:rx+0.28,y:2.20,w:2,h:0.24,fontSize:10,bold:true,color:MUTED,charSpacing:1});
  txt(s,"ATMでの取引のうち、IBでも行えるもの\n（振込・振替 など）",
    {x:rx+0.28,y:2.46,w:rw-0.56,h:0.55,fontSize:12,color:INK,lineSpacing:17});
  blank(s,{x:rx,y:3.28,w:rw,h:1.05,label:"件 数",
    hint:"2025年4〜6月 合計 ◯◯件 → 2026年4〜6月 合計 ◯◯件"});
  blank(s,{x:rx,y:4.51,w:rw,h:1.04,label:"変 化",
    hint:"1年間で −◯◯件（−◯◯％）"});

  band(s,{y:5.85,h:0.68,text:"IBは増え、ATM側は減っている　→　同じ取引がIBへ移りつつあると読める"});

  s.addNotes("IBの件数が増えても、それが「取引そのものが増えただけ」なら、IBが広がったとは言えません。\nそこで、ATMで行われている取引のうち、IBでも同じことができるもの──振込や振替など──に絞って件数を見ました。\n（記入した数値を読み上げる）\nこちらは減っています。IBが増え、ATM側が減っている。同じ取引がIBへ移りつつある、と読めます。\n\n※「IBでも行える取引」に何を含めたかは質疑で聞かれやすいので、口頭でも一度言っておく。");
}

/* ===================== 07 深掘り ===================== */
{
  const s = slide(7,"深 掘 り","では、どの年代でIBの利用が増えたのか",
    "「増えた」の次に知りたくなるのは「誰が使うようになったのか」です。今回は年代を5つに分けて調べました。");

  const gw=2.18, gp=0.25;
  const ages=[["未成年","0〜17歳"],["若手成人","18〜29歳"],["働き手①","30〜44歳"],["働き手②","45〜64歳"],["高齢者","65歳以上"]];
  ages.forEach((a,i)=>{
    const x=M+i*(gw+gp);
    card(s,{x:x,y:2.0,w:gw,h:0.78,fill:PANEL});
    txt(s,a[0],{x:x,y:2.12,w:gw,h:0.28,align:"center",fontSize:12.5,bold:true,color:INK});
    txt(s,a[1],{x:x,y:2.42,w:gw,h:0.26,align:"center",fontSize:11,color:MUTED});
  });

  chartBox(s,{x:M,y:3.05,w:6.55,h:2.6,label:"グラフ貼付エリア",
    sub:"年代別 IB利用件数の増加率（棒グラフ）"});

  const rx=7.62, rw=5.01;
  blank(s,{x:rx,y:3.05,w:rw,h:0.88,label:"読 み 取 り",
    hint:"どの年代の増加率が最も高かったか／他の年代はどうだったか"});

  card(s,{x:rx,y:4.10,w:rw,h:1.55,fill:PRIL});
  txt(s,"次 に 考 え ら れ る 施 策",{x:rx+0.28,y:4.25,w:rw-0.56,h:0.26,fontSize:10.5,bold:true,color:PRI,charSpacing:1});
  s.addText([
    {text:"伸びが小さい年代には、ATMでの振込・振替の際にIBを案内する",options:{bullet:true,breakLine:true}},
    {text:"すでに伸びている年代には、扱える手続きの幅を広げる",options:{bullet:true}}
  ],{isTextBox:true,x:rx+0.28,y:4.53,w:rw-0.56,h:1.05,margin:0,fontFace:F,fontSize:11.5,
     color:PRI,lineSpacing:16,paraSpaceAfter:6});

  band(s,{y:5.9,h:0.66,fill:PANEL,color:INK,
    text:"年代まで分かると、「誰に、どう案内するか」という具体的な打ち手まで考えられる"});

  s.addNotes("IBが増えていることが分かると、銀行側が次に知りたくなるのは「では、どの年代で増えているのか」です。これも同じテーブルに年代の条件を足すだけで確認できました。\n年代は5つに分けています。未成年、若手成人、働き手①、働き手②、高齢者です。\n（読み取りを説明する）\nここまで分かると、伸びていない年代にどう案内するか、といった具体的な打ち手の話ができるようになります。\n\n【作成時の注意】施策の2行は、実際の読み取りに合わせて書き換えること。");
}

/* ===================== 08 まとめ ===================== */
{
  const s = slide(8,"ま と め","データを「次の一手」に変えるまで",
    "今回のワークで行ったことを、ひとつの流れにまとめます。");

  const cw=2.11, gp=0.34, cy=2.15, ch=2.45;
  const steps=[
    ["問 う","預金やIBの動きを\n知りたい"],
    ["整 え る","取引履歴（流動）から\n入出金テーブルを作成"],
    ["確 か め る","IBは増加\nATM側は減少"],
    ["深 掘 る","年代別に\n増加率を確認"],
    ["次 の 一 手","案内すべき年代と\n方法を検討"]
  ];
  steps.forEach((st,i)=>{
    const x=M+i*(cw+gp), last=i===4;
    card(s,{x:x,y:cy,w:cw,h:ch,fill:last?PRI:PANEL});
    txt(s,String(i+1),{x:x,y:cy+0.28,w:cw,h:0.34,align:"center",fontSize:16,bold:true,
      color:last?"9CC8CF":"B9C7CC"});
    txt(s,st[0],{x:x,y:cy+0.78,w:cw,h:0.3,align:"center",fontSize:13,bold:true,
      color:last?WHITE:INK,charSpacing:0.8});
    txt(s,st[1],{x:x+0.12,y:cy+1.22,w:cw-0.24,h:0.9,align:"center",fontSize:11.5,
      color:last?"D6E8EB":INK2,lineSpacing:17});
    if(i<4) rArrow(s,x+cw+0.05,cy+ch/2-0.12,0.24);
  });

  band(s,{y:5.05,h:0.95,
    text:"行内にあるデータも、条件を指定して見られる形に整えれば、業務上の問いに答え、次の施策を考える材料になる",size:15});
  txt(s,"今回のワークで学んだのは、この一連のプロセスです",
    {x:M,y:6.22,w:W,h:0.32,align:"center",fontSize:12.5,color:MUTED});

  s.addNotes("最後に、今回の流れをまとめます。\n「預金やIBの動きを知りたい」という問いから始まり、取引履歴（流動）から入出金テーブルを整え、IBが増えATM側が減っていることを確かめ、さらに年代別まで深掘りして、最後は誰にどう案内するかという打ち手の検討まで進みました。\n行内にあるデータも、条件を指定して見られる形に整えれば、業務上の問いに答えて、次の施策を考える材料になります。今回のワークで学んだのは、この一連のプロセスです。");
}

pres.writeFile({fileName:"/tmp/claude-0/-home-user-my-website/9d2b3804-e19b-5dbd-a9cd-4f44a2f53f65/scratchpad/DYNATREK_成果発表.pptx"})
  .then(f=>console.log("wrote",f));
