const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";                 // 13.333 x 7.5
pres.title = "銀行データを次の一手につなげるまで";

/* ---------- tokens ---------- */
const INK="1A2B33", INK2="465A63", MUTED="8497A0", LINE="D8E0E3", PANEL="F3F6F7";
const PRI="0F5A6B", PRIL="E2EEF0", ACC="B4551F", ACCL="F8EBE1", WHITE="FFFFFF";
const F="Meiryo";
const M=0.70, W=13.333-M*2, SW=13.333, SH=7.5, CX=M+0.66, CW=W-0.66;
const LX=M, LW=6.55, RX=7.62, RW=5.01;

/* ---------- helpers ---------- */
function slide(n, title, lead){
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape(pres.ShapeType.ellipse,{x:M,y:0.56,w:0.46,h:0.46,fill:{color:PRI},line:{color:PRI,width:0}});
  s.addText(String(n).padStart(2,"0"),{isTextBox:true,x:M,y:0.56,w:0.46,h:0.46,margin:0,
    align:"center",valign:"middle",fontFace:F,fontSize:12,bold:true,color:WHITE});
  s.addText(title,{isTextBox:true,x:CX,y:0.52,w:CW,h:0.60,margin:0,
    fontFace:F,fontSize:26,bold:true,color:INK});
  if(lead) s.addText(lead,{isTextBox:true,x:CX,y:1.24,w:CW,h:0.42,margin:0,
    fontFace:F,fontSize:14,color:INK2});
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

function figBox(s,o){                          // グレー点線＝図版を貼る
  s.addShape(pres.ShapeType.roundRect,{x:o.x,y:o.y,w:o.w,h:o.h,rectRadius:0.05,
    fill:{color:o.tint||PANEL},line:{color:o.border||"BFCCD1",width:1.25,dashType:"dash"}});
  txt(s,o.label||"グラフ貼付エリア",{x:o.x,y:o.y+o.h/2-0.42,w:o.w,h:0.32,align:"center",
    fontSize:13,bold:true,color:PRI});
  txt(s,o.sub,{x:o.x+0.3,y:o.y+o.h/2-0.04,w:o.w-0.6,h:0.78,align:"center",
    fontSize:11,color:MUTED,lineSpacing:17});
}
function blank(s,o){                           // オレンジ点線＝本人が書く
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

/* ================= 01 はじめに ================= */
{
  const s = slide(1,"はじめに",
    "本インターンのテーマは、銀行内で保有する各種データを分析し、どのような価値に変換していくかのプロセスを学ぶこと");

  const ch=1.05, ys=[2.05,3.38,4.71];
  [["現 状","銀行には、日々の取引データが蓄積されている"],
   ["課 題","項目が多く、知りたい切り口でまとめて取り出すのは容易ではない"],
   ["道 具","DYNATREK ── 条件を指定するだけで、必要なデータを集計・表示できる"]
  ].forEach((r,i)=>{
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

/* ================= 02 背景 ================= */
{
  const s = slide(2,"背景",
    "あいち銀行では、デジタルサービスの強化を中期経営計画の重点施策に掲げている");

  const lw=6.75, ys=[2.05,3.10,4.15];
  [["計 画","あいちフィナンシャルグループ　第2次中期経営計画（2025年4月〜2028年3月）",0],
   ["基本戦略Ⅲ","DX戦略の加速化",0.35],
   ["重点施策⑦","顧客の利便性・先進性の向上",0.70]
  ].forEach((r,i)=>{
    const x=M+r[2], w=lw-r[2], hl=i===2;
    card(s,{x:x,y:ys[i],w:w,h:0.85,fill:hl?PRI:PANEL});
    eyebrow(s,r[0],{x:x+0.28,y:ys[i]+0.13,w:2.2,fontSize:10,color:hl?"9CC8CF":MUTED});
    txt(s,r[1],{x:x+0.28,y:ys[i]+0.40,w:w-0.56,h:0.34,fontSize:hl?14:13,bold:hl,color:hl?WHITE:INK});
    if(i<2) dArrow(s,x+0.5,ys[i]+0.88,0.20);
  });
  txt(s,"「すべてのお客さまに便利に使っていただける金融サービスを提供する」",
    {x:M+0.70,y:5.06,w:lw-0.70,h:0.36,fontSize:11.5,color:MUTED});

  const rx=7.72, rw=4.91;
  card(s,{x:rx,y:2.05,w:rw,h:3.30});
  eyebrow(s,"取 組 方 針 ⑯ 　顧 客 サ ー ビ ス の 強 化",{x:rx+0.3,y:2.26,w:rw-0.6,fontSize:10.5});
  [["バンキングアプリ・インターネット支店をリニューアル",false],
   ["若年層等の新たな顧客層へのアプローチを強化",true],
   ["法人向けにライト版法人IBを導入",false]
  ].forEach((it,i)=>{
    const y=2.68+i*0.80;
    card(s,{x:rx+0.3,y:y,w:rw-0.6,h:0.68,fill:it[1]?ACCL:WHITE,line:it[1]?"D9A97F":LINE,lw:1});
    txt(s,it[0],{x:rx+0.48,y:y,w:rw-0.96,h:0.68,valign:"middle",fontSize:11.5,
      bold:it[1],color:it[1]?ACC:INK2,lineSpacing:16});
  });
  txt(s,"出典：あいちFG「第2次中期経営計画」p.31",
    {x:rx,y:5.10,w:rw,h:0.26,align:"center",fontSize:9.5,color:MUTED});

  band(s,{y:5.72,size:13.5,
    text:"非対面のチャネルがどれだけ使われているかは、この狙いの進み具合を示す　→　今回はインターネットバンキング（IB）を対象にした"});

  s.addNotes("では、どの問いを選んだか。銀行が実際に力を入れている領域から選びました。\nあいちフィナンシャルグループの第2次中期経営計画では、基本戦略Ⅲ「DX戦略の加速化」のもとに、重点施策⑦「顧客の利便性・先進性の向上」が置かれています。その取組方針が「顧客サービスの強化」で、バンキングアプリやインターネット支店のリニューアル、そして若年層など新たな顧客層へのアプローチ強化が挙げられています。\nこうした非対面のチャネルがどれだけ使われているかは、この狙いがどこまで進んでいるかを示します。今回はそのうち、インターネットバンキング、IBを対象にしました。\n\n※「若年層へのアプローチ強化」は8枚目の年代別分析につながる伏線。ここで一度触れておく。");
}

/* ================= 03 目的 ================= */
{
  const s = slide(3,"目的",
    "IBの利用が1年間でどれだけ広がったかを、2つの期間の比較で確認する");

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

  s.addNotes("確かめることは、2025年と2026年の同じ時期を比べて、IBの利用件数がどれだけ増えたかを確認することです。\n比べ方は同じ4〜6月どうし。月によって取引の量は変わるので、同じ季節どうしで比べます。\nもうひとつ、IBの件数だけでなく、ATM側の「IBでも代替できる取引」も合わせて見ることにしました。理由は結果のところで説明します。");
}

/* ================= 04 使用したデータ ================= */
{
  const s = slide(4,"使用したデータ",
    "取引履歴（流動）をもとに作成された入出金テーブルを使用した");

  card(s,{x:LX,y:2.05,w:LW,h:2.30,fill:PRIL});
  eyebrow(s,"使 用 デ ー タ",{x:LX+0.4,y:2.30,w:LW-0.8,color:PRI});
  txt(s,"入出金テーブル",{x:LX+0.4,y:2.68,w:LW-0.8,h:0.5,fontSize:24,bold:true,color:PRI});
  txt(s,"取引履歴（流動）をもとに作成されたもの",
    {x:LX+0.4,y:3.32,w:LW-0.8,h:0.34,fontSize:13,color:PRI});
  txt(s,"入出金の状況を、条件を指定して確認できる形になっている",
    {x:LX+0.4,y:3.72,w:LW-0.8,h:0.34,fontSize:12,color:INK2});

  card(s,{x:RX,y:2.05,w:RW,h:2.30});
  eyebrow(s,"対 象 期 間",{x:RX+0.36,y:2.30,w:RW-0.72});
  [["2025年　4月・5月・6月",2.68],["2026年　4月・5月・6月",3.44]].forEach(p=>{
    card(s,{x:RX+0.36,y:p[1],w:RW-0.72,h:0.64,fill:WHITE,line:LINE,lw:1});
    txt(s,p[0],{x:RX+0.36,y:p[1],w:RW-0.72,h:0.64,align:"center",valign:"middle",
      fontSize:14,bold:true,color:INK});
  });

  blank(s,{x:M,y:4.65,w:W,h:1.05,label:"記 入 欄",
    hint:"入出金テーブルの主な項目（例：日付／チャネル／取引種別／年代 など、実際に使ったものを記載）"});

  band(s,{y:6.00,h:0.62,fill:PANEL,color:INK,
    text:"このテーブルに対して、DYNATREKで条件を指定して集計する"});

  s.addNotes("使ったのは、取引履歴（流動）をもとに作成された入出金テーブルです。私が作ったものではなく、既にある形のものを使わせていただきました。\n対象期間は、2025年と2026年の4月・5月・6月です。\nこのテーブルに対して、DYNATREKで条件を指定して集計していきます。");
}

/* ================= 05 分析方法 ================= */
{
  const s = slide(5,"分析方法",
    "入出金テーブルに条件を指定するだけで、必要な集計を取り出せる");

  const fw=5.50, fy=2.05, fh=3.20, fx1=M, fx2=SW-M-fw;
  figBox(s,{x:fx1,y:fy,w:fw,h:fh,label:"スクリーンショット ①",
    sub:"DYNATREKの条件設定画面\n（期間・チャネル・取引種別などを指定）"});
  figBox(s,{x:fx2,y:fy,w:fw,h:fh,label:"スクリーンショット ②",
    sub:"指定した条件で出力された集計結果",tint:PRIL,border:"A8C7CC"});
  s.addShape(pres.ShapeType.rightArrow,{x:6.34,y:fy+fh/2-0.20,w:0.65,h:0.40,
    fill:{color:PRI},line:{color:PRI,width:0}});
  txt(s,"条件を指定する",{x:fx1,y:fy+fh+0.16,w:fw,h:0.30,align:"center",fontSize:13,bold:true,color:INK});
  txt(s,"そのまま集計結果が出る",{x:fx2,y:fy+fh+0.16,w:fw,h:0.30,align:"center",fontSize:13,bold:true,color:PRI});

  band(s,{y:5.95,h:0.62,fill:PANEL,color:INK,
    text:"条件を変えるだけで、別の切り口でも確認できる"});

  s.addNotes("実際の画面です。左で、期間やチャネル、取引の種類といった条件を指定します。すると右のように、その条件で集計された結果がそのまま出てきます。\nここは長く説明しません。お伝えしたいのは、条件を変えるだけで別の切り口でも確認できる、という点です。この後の分析は、すべてこの繰り返しになります。\n\n【作成時】スクリーンショット②が6枚目のグラフと同じものになる場合は、②には集計表など別の出力を貼るか、②を小さめに配置して重複感を避ける。");
}

/* ================= 06 結果① ================= */
{
  const s = slide(6,"結果①　IBの利用件数",
    "IBで行われた取引の件数は、1年間で増加していた");

  figBox(s,{x:LX,y:2.05,w:LW,h:3.5,sub:"IBの利用件数（2025年4〜6月 / 2026年4〜6月）"});
  blank(s,{x:RX,y:2.05,w:RW,h:1.05,label:"件 数",
    hint:"2025年4〜6月 合計 ◯◯件 → 2026年4〜6月 合計 ◯◯件"});
  blank(s,{x:RX,y:3.28,w:RW,h:1.05,label:"変 化",hint:"1年間で ＋◯◯件（＋◯◯％）"});
  blank(s,{x:RX,y:4.51,w:RW,h:1.04,label:"月 別 の 傾 向",
    hint:"4月・5月・6月それぞれ増えているか／特に伸びた月はどこか"});

  band(s,{y:5.85,h:0.68,fill:PANEL,color:INK,
    text:"ただし件数の増加だけでは、取引そのものが増えた可能性を否定できない"});

  s.addNotes("まずIBの利用件数です。2025年の4〜6月と、2026年の4〜6月を比べました。\n（記入した数値を読み上げる）\n1年間で件数は増えています。月別に見ても同じ傾向です。\nただ、これだけでIBが広がったと言い切ることはできません。取引そのものが増えただけ、という可能性が残るからです。そこで、もうひとつの見方を足しました。");
}

/* ================= 07 結果② ================= */
{
  const s = slide(7,"結果②　ATMでIBに代替できる取引",
    "IBへ移っているかを判断するため、ATM側の同種の取引も確認した");

  figBox(s,{x:LX,y:2.05,w:LW,h:3.5,sub:"ATMでのIB代替可能な取引の件数（2025年4〜6月 / 2026年4〜6月）"});

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

/* ================= 08 考察 ================= */
{
  const s = slide(8,"考察",
    "計画が掲げる「若年層へのアプローチ強化」が進んでいるかを、年代別に確認した");

  const gw=2.18, gp=0.25;
  [["未成年","0〜17歳"],["若手成人","18〜29歳"],["働き手①","30〜44歳"],
   ["働き手②","45〜64歳"],["高齢者","65歳以上"]].forEach((a,i)=>{
    const x=M+i*(gw+gp);
    card(s,{x:x,y:2.05,w:gw,h:0.74});
    txt(s,a[0],{x:x,y:2.15,w:gw,h:0.28,align:"center",fontSize:12.5,bold:true,color:INK});
    txt(s,a[1],{x:x,y:2.45,w:gw,h:0.26,align:"center",fontSize:11,color:MUTED});
  });

  figBox(s,{x:LX,y:2.98,w:LW,h:2.55,sub:"年代別 IB利用件数の増加率（棒グラフ）"});

  blank(s,{x:RX,y:2.98,w:RW,h:0.95,label:"読 み 取 り",
    hint:"どの年代の増加率が最も高かったか／他の年代はどうだったか"});

  card(s,{x:RX,y:4.13,w:RW,h:1.40,fill:PRIL});
  eyebrow(s,"こ こ か ら 考 え ら れ る 打 ち 手",{x:RX+0.28,y:4.28,w:RW-0.56,color:PRI});
  s.addText([
    {text:"伸びが小さい年代には、ATMでの振込・振替の際にIBを案内する",options:{bullet:true,breakLine:true}},
    {text:"すでに伸びている年代には、扱える手続きの幅を広げる",options:{bullet:true}}
  ],{isTextBox:true,x:RX+0.28,y:4.56,w:RW-0.56,h:0.9,margin:0,fontFace:F,fontSize:11.5,
     color:PRI,lineSpacing:16,paraSpaceAfter:6});

  band(s,{y:5.78,h:0.68,
    text:"年代まで分かると、計画の狙いがどこまで進んでいるかと、次に誰へ案内すべきかが同時に見える"});

  s.addNotes("IBへ移りつつあることが分かると、次に知りたくなるのは「では、誰が使うようになったのか」です。\n2枚目で触れたとおり、計画では若年層など新たな顧客層へのアプローチ強化が掲げられています。年代別に見れば、その狙いが実際に進んでいるかを確認できます。\n年代は5つに分けました。未成年、若手成人、働き手①、働き手②、高齢者です。\n（読み取りを説明する）\nここまで分かると、計画の狙いの進み具合と、次に誰へ案内すべきかが同時に見えてきます。\n\n【作成時】打ち手の2行は、実際の読み取りに合わせて書き換えること。");
}

/* ================= 09 まとめ ================= */
{
  const s = slide(9,"まとめ","今回のワークで行ったことを、ひとつの流れにまとめる");

  const cw=2.11, gp=0.34, cy=2.15, ch=2.45;
  [["問 い を 立 て る","「IBの利用は\n広がっているか」"],
   ["条 件 を 指 定 す る","入出金テーブルから\n必要な集計を取り出す"],
   ["確 か め る","IBは増加\nATM側は減少"],
   ["深 掘 る","年代別に\n増加率を確認"],
   ["次 の 一 手","案内すべき年代と\n方法を検討"]
  ].forEach((st,i)=>{
    const x=M+i*(cw+gp), last=i===4;
    card(s,{x:x,y:cy,w:cw,h:ch,fill:last?PRI:PANEL});
    txt(s,String(i+1),{x:x,y:cy+0.28,w:cw,h:0.34,align:"center",fontSize:16,bold:true,
      color:last?"9CC8CF":"B9C7CC"});
    txt(s,st[0],{x:x+0.04,y:cy+0.78,w:cw-0.08,h:0.30,align:"center",fontSize:11,bold:true,
      color:last?WHITE:INK,charSpacing:0.2});
    txt(s,st[1],{x:x+0.12,y:cy+1.22,w:cw-0.24,h:0.9,align:"center",fontSize:11.5,
      color:last?"D6E8EB":INK2,lineSpacing:17});
    if(i<4) rArrow(s,x+cw+0.05,cy+ch/2-0.12,0.24);
  });

  band(s,{y:5.05,h:0.95,size:15,
    text:"行内にあるデータも、条件を指定して見られる形にすれば、業務上の問いに答え、次の施策を考える材料になる"});
  txt(s,"今回のワークで学んだのは、この一連のプロセス",
    {x:M,y:6.22,w:W,h:0.32,align:"center",fontSize:12.5,color:MUTED});

  s.addNotes("最後に、今回の流れをまとめます。\n「IBの利用は広がっているか」という問いを立て、入出金テーブルから必要な集計を取り出し、IBが増えATM側が減っていることを確かめ、さらに年代別まで深掘りして、最後は誰にどう案内するかという打ち手の検討まで進みました。\n行内にあるデータも、条件を指定して見られる形にすれば、業務上の問いに答えて、次の施策を考える材料になります。今回のワークで学んだのは、この一連のプロセスです。");
}

pres.writeFile({fileName:"/tmp/claude-0/-home-user-my-website/9d2b3804-e19b-5dbd-a9cd-4f44a2f53f65/scratchpad/DYNATREK_成果発表.pptx"})
  .then(f=>console.log("wrote",f));
