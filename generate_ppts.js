const pptxgen = require("pptxgenjs");
const fs = require("fs");

const LOGO = "image/png;base64," + fs.readFileSync("SM CAT 3/ppt/media/image1.png").toString("base64");

// ── Color palette ─────────────────────────────────────────────
const M   = "990033"; // maroon   (primary / border)
const GN  = "009900"; // green    (header)
const GL  = "F9DD67"; // gold     (inner border)
const WH  = "FFFFFF"; // white
const BK  = "000000"; // black
const GY  = "555555"; // gray
const LGR = "F8F8F8"; // light gray
const LGN = "E8F5E9"; // light green
const LMR = "FCE4EC"; // light maroon/pink
const BL  = "1565C0"; // blue
const LBL = "E3F2FD"; // light blue
const AM  = "F57F17"; // amber
const LYL = "FFF9C4"; // light yellow
const DGN = "006400"; // dark green

// ── Slide dimensions (LAYOUT_WIDE = 13.33 × 7.5 in) ─────────
const SW = 13.33, SH = 7.5;
const BRD = 0.26;     // maroon border width
const GLD = 0.09;     // gold accent width

// ── Header / logo / content coordinates ──────────────────────
const HX=0.35, HY=0.27, HW=11.12, HH=0.79;
const LX=11.55, LY=0.25, LW=1.03, LH=1.03;
const CX=0.50,  CW=12.35;
const CYT=1.12;   // y of slide subtitle title
const CYC=1.72;   // y where main content starts

// ─────────────────────────────────────────────────────────────
//  SHARED LAYOUT HELPERS
// ─────────────────────────────────────────────────────────────
function frame(pres, sl, hdrText) {
  sl.background = { color: WH };
  [[0,0,SW,BRD],[0,SH-BRD,SW,BRD],[0,0,BRD,SH],[SW-BRD,0,BRD,SH]]
    .forEach(([x,y,w,h]) => sl.addShape(pres.shapes.RECTANGLE,
      {x,y,w,h, fill:{color:M}, line:{color:M,width:1}}));
  const gx=BRD,gy=BRD,gw=SW-2*BRD,gh=SH-2*BRD;
  [[gx,gy,gw,GLD],[gx,gy+gh-GLD,gw,GLD],[gx,gy,GLD,gh],[gx+gw-GLD,gy,GLD,gh]]
    .forEach(([x,y,w,h]) => sl.addShape(pres.shapes.RECTANGLE,
      {x,y,w,h, fill:{color:GL}, line:{color:GL,width:1}}));
  sl.addShape(pres.shapes.RECTANGLE,
    {x:HX,y:HY,w:HW,h:HH, fill:{color:GN}, line:{color:GN,width:1}});
  sl.addText(hdrText,
    {x:HX,y:HY,w:HW,h:HH, fontSize:15,bold:true,color:WH,
     align:"center",valign:"middle",fontFace:"Arial",margin:0});
  sl.addImage({data:LOGO, x:LX,y:LY,w:LW,h:LH});
}

function titleSlide(pres, hdr, unit, topic) {
  const sl = pres.addSlide();
  frame(pres, sl, hdr);
  sl.addText(unit,
    {x:CX,y:1.18,w:CW,h:0.42, fontSize:13,italic:true,color:M,align:"center",fontFace:"Arial"});
  sl.addShape(pres.shapes.RECTANGLE,
    {x:1.80,y:1.72,w:9.73,h:1.40, fill:{color:M}, line:{color:M,width:1}});
  sl.addText(topic,
    {x:1.80,y:1.72,w:9.73,h:1.40, fontSize:28,bold:true,color:WH,
     align:"center",valign:"middle",fontFace:"Arial",margin:0});
  sl.addText("PPT PRESENTED BY",
    {x:CX,y:3.28,w:CW,h:0.40, fontSize:14,bold:true,color:M,align:"center",fontFace:"Arial"});
  ["NAME :   SRinivas Rao Tammireddy",
   "ROLL NO :   257Y1D5805",
   "YEAR :   I Year  I Semester",
   "ACADEMIC YEAR :   2025 - 2026"
  ].forEach((t,i) =>
    sl.addText(t,{x:3.20,y:3.76+i*0.44,w:7.0,h:0.42,fontSize:12,color:BK,fontFace:"Arial"}));
  sl.addText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING",
    {x:CX,y:6.65,w:CW,h:0.45, fontSize:13,bold:true,color:M,align:"center",fontFace:"Arial"});
}

function contentSlide(pres, hdr, title) {
  const sl = pres.addSlide();
  frame(pres, sl, hdr);
  sl.addText(title,
    {x:CX,y:CYT,w:CW,h:0.52, fontSize:20,bold:true,color:M,fontFace:"Arial"});
  sl.addShape(pres.shapes.RECTANGLE,
    {x:CX,y:CYT+0.52,w:CW,h:0.04, fill:{color:M}, line:{color:M,width:1}});
  return sl;
}

function thankYouSlide(pres, hdr) {
  const sl = pres.addSlide();
  frame(pres, sl, hdr);
  sl.addText("THANK YOU",
    {x:CX,y:2.6,w:CW,h:1.5, fontSize:48,bold:true,color:M,
     align:"center",valign:"middle",fontFace:"Arial"});
  sl.addText(
    "Department of Computer Science and Engineering\n"+
    "Marri Laxman Reddy Institute of Technology and Management",
    {x:CX,y:5.5,w:CW,h:0.9, fontSize:12,color:GY,align:"center",fontFace:"Arial"});
}

function box(pres, sl, x, y, w, h, bg, bd) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE,
    {x,y,w,h, fill:{color:bg}, line:{color:bd,width:1.5}, rectRadius:0.1});
}

function boxTitle(sl, x, y, w, txt, clr) {
  sl.addText(txt,
    {x:x+0.12,y:y+0.10,w:w-0.24,h:0.38, fontSize:13,bold:true,color:clr,fontFace:"Arial",margin:0});
}

function boxBody(sl, x, y, w, h, txt) {
  sl.addText(txt,
    {x:x+0.12,y:y+0.52,w:w-0.24,h:h-0.58, fontSize:11.5,color:BK,fontFace:"Arial",margin:0});
}

// ═══════════════════════════════════════════════════════════════
//  PPT 1 – PROPOSITIONAL LOGIC  (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Mathematical Foundations of Computer Science";

  // Slide 1 ── Title
  titleSlide(p, H, "Unit I: The Foundations – Logic and Proofs", "Propositional Logic");

  // Slide 2 ── What is Propositional Logic?
  {
    const s = contentSlide(p, H, "What is Propositional Logic?");
    s.addText([
      {text:"Propositional Logic ",options:{bold:true}},
      {text:"is a branch of formal logic dealing with propositions and their "+
            "combinations using logical connectives.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.52, fontSize:13.5,fontFace:"Arial",color:BK});

    // Left box
    box(p,s,CX,2.30,5.90,4.26,LGN,GN);
    s.addText("What is a Proposition?",
      {x:0.62,y:2.38,w:5.70,h:0.38, fontSize:14,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"A declarative statement that is either ",options:{}},
      {text:"TRUE",options:{bold:true,color:DGN}},{text:" or ",options:{}},
      {text:"FALSE",options:{bold:true,color:M}},{text:" — never both.",options:{breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"Propositions (have truth values):",options:{bold:true,breakLine:true}},
      {text:"   \"2 + 2 = 4\"  →  TRUE",options:{breakLine:true}},
      {text:"   \"Paris is in Germany\"  →  FALSE",options:{breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"NOT Propositions:",options:{bold:true,breakLine:true}},
      {text:"   \"Close the door!\"  (command)",options:{breakLine:true}},
      {text:"   \"What time is it?\"  (question)",options:{breakLine:true}},
      {text:"   \"x + 1 = 5\"  (depends on x)",options:{}}
    ],{x:0.62,y:2.80,w:5.70,h:3.68, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    // Right box
    box(p,s,6.60,2.30,6.25,4.26,LMR,M);
    s.addText("Notation and Key Terms",
      {x:6.72,y:2.38,w:6.03,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Variables:",options:{bold:true,breakLine:true}},
      {text:"   p, q, r, s  (lower-case letters for propositions)",options:{breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"Truth Values:",options:{bold:true,breakLine:true}},
      {text:"   T = True = 1        F = False = 0",options:{breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"Compound Proposition:",options:{bold:true,breakLine:true}},
      {text:"   Built using connectives:  ¬ ∧ ∨ → ↔",options:{breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"Example:",options:{bold:true,breakLine:true}},
      {text:"   p = \"It is raining\"",options:{breakLine:true}},
      {text:"   q = \"It is cloudy\"",options:{breakLine:true}},
      {text:"   p → q  =  \"If raining, then cloudy\"",options:{}}
    ],{x:6.72,y:2.80,w:6.03,h:3.68, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 2 – What is Propositional Logic?\n\n"+
"Propositional Logic is the foundation of all formal reasoning in computer science and mathematics. It studies how we form and reason about propositions — declarative statements that can be assigned a definite truth value of TRUE or FALSE.\n\n"+
"The keyword here is 'declarative'. The statement must make a factual claim about something. 'Two plus two equals four' is a proposition — it is always true. 'Paris is in Germany' is a proposition — it is always false. However, a command like 'Close the door' is not a proposition because we cannot assign it a truth value. A question like 'What time is it?' is not a proposition either. An open formula like 'x plus 1 equals 5' is not a proposition because its truth depends on the unknown variable x.\n\n"+
"We use lowercase letters p, q, r, s as variables to represent propositions. Each variable takes a truth value: T (True, also written as 1) or F (False, also written as 0).\n\n"+
"When we combine simple (atomic) propositions using logical connectives — NOT, AND, OR, IMPLIES, and IF-AND-ONLY-IF — we form compound propositions. For example, if p means 'it is raining' and q means 'it is cloudy', then p arrow q is the compound proposition 'if it is raining, then it is cloudy'. The entire field of propositional logic studies how to evaluate, simplify, and reason about such compound statements.");
  }

  // Slide 3 ── Logical Connectives
  {
    const s = contentSlide(p, H, "Logical Connectives");
    s.addText("Five fundamental connectives build compound propositions from atomic ones:",
      {x:CX,y:CYC,w:CW,h:0.40, fontSize:14,fontFace:"Arial",color:BK});
    const tData = [
      [
        {text:"Symbol",options:{bold:true,fill:{color:GN},color:WH,align:"center"}},
        {text:"Name",options:{bold:true,fill:{color:GN},color:WH,align:"center"}},
        {text:"Read As",options:{bold:true,fill:{color:GN},color:WH,align:"center"}},
        {text:"Example  (p = 'it is hot',  q = 'it is humid')",options:{bold:true,fill:{color:GN},color:WH}}
      ],
      [
        {text:"¬ p",options:{bold:true,align:"center",fill:{color:LGR}}},
        {text:"Negation",options:{fill:{color:LGR}}},{text:"NOT p",options:{fill:{color:LGR}}},
        {text:"\"It is NOT hot\"",options:{italic:true,fill:{color:LGR}}}
      ],
      [
        {text:"p ∧ q",options:{bold:true,align:"center",fill:{color:LBL},color:BL}},
        {text:"Conjunction",options:{fill:{color:LBL}}},{text:"p AND q",options:{fill:{color:LBL}}},
        {text:"\"It is hot AND humid\"",options:{italic:true,fill:{color:LBL}}}
      ],
      [
        {text:"p ∨ q",options:{bold:true,align:"center",fill:{color:LGN},color:DGN}},
        {text:"Disjunction",options:{fill:{color:LGN}}},{text:"p OR q  (inclusive)",options:{fill:{color:LGN}}},
        {text:"\"It is hot OR humid (or both)\"",options:{italic:true,fill:{color:LGN}}}
      ],
      [
        {text:"p → q",options:{bold:true,align:"center",fill:{color:LMR},color:M}},
        {text:"Conditional",options:{fill:{color:LMR}}},{text:"If p, then q",options:{fill:{color:LMR}}},
        {text:"\"If it is hot, then it is humid\"",options:{italic:true,fill:{color:LMR}}}
      ],
      [
        {text:"p ↔ q",options:{bold:true,align:"center",fill:{color:LYL},color:AM}},
        {text:"Biconditional",options:{fill:{color:LYL}}},{text:"p if and only if q",options:{fill:{color:LYL}}},
        {text:"\"Hot if and only if humid\"",options:{italic:true,fill:{color:LYL}}}
      ]
    ];
    s.addTable(tData,
      {x:CX,y:2.18,w:CW,h:5.0, fontSize:13,fontFace:"Arial",
       border:{pt:1,color:"CCCCCC"},colW:[1.55,2.12,2.22,6.46]});
    s.addNotes(
"Slide 3 – Logical Connectives\n\n"+
"The five logical connectives are the building blocks that allow us to form compound propositions.\n\n"+
"NEGATION (¬): The simplest connective. It inverts the truth value. If p is TRUE, ¬p is FALSE, and vice versa. It corresponds to the word NOT.\n\n"+
"CONJUNCTION (∧): Corresponds to AND. The compound statement p ∧ q is TRUE ONLY when BOTH p and q are TRUE. If either is FALSE, the conjunction is FALSE. Think of it as a chain — the whole thing fails if any link breaks.\n\n"+
"DISJUNCTION (∨): Corresponds to OR. The compound statement p ∨ q is TRUE when AT LEAST ONE of p or q is TRUE. This is the inclusive OR, meaning it is TRUE even when both are TRUE. It is FALSE only when BOTH are FALSE.\n\n"+
"CONDITIONAL (→): Represents IF...THEN, also called IMPLICATION. p → q is FALSE only when the hypothesis p is TRUE and the conclusion q is FALSE. In all other cases — especially when p is FALSE — the implication is TRUE (called 'vacuously true'). This is the most frequently used connective in mathematical proofs.\n\n"+
"BICONDITIONAL (↔): Represents IF AND ONLY IF. p ↔ q is TRUE when p and q have exactly the same truth value — both TRUE or both FALSE. It is equivalent to writing both p→q AND q→p.\n\n"+
"Using our example: p = 'it is hot', q = 'it is humid' — each connective creates a completely different compound statement as shown in the table.");
  }

  // Slide 4 ── Truth Tables
  {
    const s = contentSlide(p, H, "Truth Tables");
    s.addText(
      "A truth table systematically evaluates a compound proposition "+
      "for every possible combination of truth values:",
      {x:CX,y:CYC,w:CW,h:0.50, fontSize:13.5,fontFace:"Arial",color:BK});
    const th=(t)=>({text:t,options:{bold:true,fill:{color:M},color:WH,align:"center"}});
    const r1=(t,b,c)=>({text:t,options:{fill:{color:LGN},align:"center",bold:!!b,color:c||BK}});
    const r2=(t,b,c)=>({text:t,options:{fill:{color:LMR},align:"center",bold:!!b,color:c||BK}});
    const r3=(t,b,c)=>({text:t,options:{fill:{color:LBL},align:"center",bold:!!b,color:c||BK}});
    const r4=(t,b,c)=>({text:t,options:{fill:{color:LGR},align:"center",bold:!!b,color:c||BK}});
    const tData=[
      [th("p"),th("q"),th("¬ p"),th("p ∧ q"),th("p ∨ q"),th("p → q"),th("p ↔ q")],
      [r1("T",1),r1("T",1),r1("F",0),r1("T",1,DGN),r1("T",1,DGN),r1("T",1,DGN),r1("T",1,DGN)],
      [r2("T",1),r2("F",1),r2("F",0),r2("F",1,M), r2("T",1,DGN),r2("F",1,M), r2("F",1,M)],
      [r3("F",1),r3("T",1),r3("T",1,DGN),r3("F",0,M),r3("T",1,DGN),r3("T",1,DGN),r3("F",0,M)],
      [r4("F",1),r4("F",1),r4("T",1,DGN),r4("F",0,M),r4("F",0,M),r4("T",1,DGN),r4("T",1,DGN)]
    ];
    s.addTable(tData,
      {x:CX,y:2.28,w:CW,h:3.65, fontSize:14,fontFace:"Arial",
       border:{pt:1,color:"CCCCCC"},colW:[1.55,1.55,1.58,1.74,1.74,1.85,2.34]});
    // insight box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:6.08,w:CW,h:0.88, fill:{color:LYL}, line:{color:GL,width:2},rectRadius:0.1});
    s.addText(
      "★  Key Insight (p → q):  The conditional is FALSE only when p = TRUE and q = FALSE "+
      "(Row 2). A FALSE hypothesis makes the implication vacuously TRUE — even 'If 1=2 then pigs fly' is TRUE!",
      {x:0.62,y:6.11,w:12.12,h:0.82, fontSize:12.5,fontFace:"Arial",color:"4E342E",margin:0});
    s.addNotes(
"Slide 4 – Truth Tables\n\n"+
"The truth table is the most powerful tool in propositional logic. It exhaustively lists every possible combination of truth values for the variables and shows the resulting value of the compound proposition.\n\n"+
"With two variables p and q, we have exactly 4 rows: TT, TF, FT, and FF.\n\n"+
"¬p column: Simply flips the value of p. T becomes F and F becomes T.\n\n"+
"p∧q (AND): TRUE only in Row 1 where BOTH p and q are TRUE.\n\n"+
"p∨q (OR): FALSE only in Row 4 where BOTH are FALSE. TRUE in all other rows.\n\n"+
"p→q (IMPLIES): This is the most important. It is FALSE ONLY in Row 2 — when p is TRUE and q is FALSE. In Rows 3 and 4, p is FALSE, so the implication is TRUE regardless of q. This is called 'vacuous truth'. For example, 'If the moon is made of cheese, then 2+2=5' is technically TRUE in propositional logic, because the hypothesis is FALSE!\n\n"+
"p↔q (BICONDITIONAL): TRUE in Rows 1 and 4 where both variables have the SAME truth value. FALSE in Rows 2 and 3.\n\n"+
"Truth tables are used to verify logical equivalences, design digital circuits, simplify logical expressions, and create systematic software test cases through a method called Condition Coverage Testing.");
  }

  // Slide 5 ── Logical Equivalence
  {
    const s = contentSlide(p, H, "Logical Equivalence & Key Laws");
    s.addText(
      "Two propositions P and Q are logically equivalent (P ≡ Q) if they have "+
      "identical truth values for every possible assignment of variable values.",
      {x:CX,y:CYC,w:CW,h:0.58, fontSize:13.5,fontFace:"Arial",color:BK});
    // De Morgan's box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:2.38,w:CW,h:1.62, fill:{color:LBL}, line:{color:BL,width:1.5},rectRadius:0.1});
    s.addText("De Morgan's Laws  (most important equivalences in logic)",
      {x:0.62,y:2.44,w:CW-0.24,h:0.38, fontSize:14,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"¬(p ∧ q)  ≡  ¬p ∨ ¬q",options:{bold:true,color:M}},
      {text:"      → NOT of (AND)  =  OR of (NOTs)",options:{italic:true,color:BK,breakLine:true}},
      {text:"¬(p ∨ q)  ≡  ¬p ∧ ¬q",options:{bold:true,color:M}},
      {text:"      → NOT of (OR)  =  AND of (NOTs)",options:{italic:true,color:BK}}
    ],{x:0.62,y:2.86,w:CW-0.24,h:1.08, fontSize:14,fontFace:"Arial",margin:0});
    s.addText("Other Key Equivalences:",
      {x:CX,y:4.12,w:CW,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial"});
    const laws=[
      {n:"Contrapositive:",       f:"p → q  ≡  ¬q → ¬p",         bg:LMR},
      {n:"Double Negation:",      f:"¬(¬p)  ≡  p",                bg:LGN},
      {n:"Idempotent:",           f:"p∧p ≡ p     and     p∨p ≡ p",bg:LBL},
      {n:"Implication Rewrite:",  f:"p → q  ≡  ¬p ∨ q",          bg:LYL}
    ];
    laws.forEach((law,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=4.58+row*1.02;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:bx,y:by,w:6.05,h:0.90, fill:{color:law.bg}, line:{color:"CCCCCC",width:1},rectRadius:0.08});
      s.addText([
        {text:law.n+"  ",options:{bold:true,color:M}},
        {text:law.f,options:{color:BK}}
      ],{x:bx+0.10,y:by+0.10,w:5.85,h:0.72, fontSize:13.5,fontFace:"Arial",margin:0});
    });
    s.addNotes(
"Slide 5 – Logical Equivalence and Key Laws\n\n"+
"Two logical formulas are called logically equivalent if they produce the exact same truth value for every possible assignment of truth values to the variables. We write P ≡ Q to denote equivalence.\n\n"+
"The most famous equivalences are De Morgan's Laws, named after the 19th-century mathematician Augustus De Morgan.\n\n"+
"First De Morgan's Law: NOT(p AND q) is equivalent to (NOT p) OR (NOT q). In plain English: 'It is not the case that both p and q are true' is the same as saying 'either p is false, or q is false, or both'. For example: 'It is NOT (hot AND humid)' means 'It is not hot OR it is not humid'.\n\n"+
"Second De Morgan's Law: NOT(p OR q) is equivalent to (NOT p) AND (NOT q). 'Neither p nor q' means 'p is false AND q is false'. Example: 'It is neither hot nor humid' means 'It is not hot AND it is not humid'.\n\n"+
"De Morgan's Laws are extensively used in circuit minimization, simplifying if-else conditions in programming, and database query optimization.\n\n"+
"The Contrapositive: p→q is equivalent to NOT q→NOT p. To prove 'if p then q', we can equivalently prove 'if not q then not p'. This is used extensively in proofs by contrapositive.\n\n"+
"Double Negation: Negating twice brings back the original, just like multiplying by minus one twice.\n\n"+
"Implication Rewrite: p→q can be rewritten as NOT p OR q. This form is used in resolution proofs in automated theorem proving.");
  }

  // Slide 6 ── Applications
  {
    const s = contentSlide(p, H, "Applications in Computer Science");
    s.addText("Propositional Logic is the theoretical backbone of key areas in computer science:",
      {x:CX,y:CYC,w:CW,h:0.40, fontSize:14,fontFace:"Arial",color:BK});
    const apps=[
      {t:"⚙  Digital Circuit Design",
       b:"Logic gates (AND, OR, NOT, NAND, NOR, XOR) are direct hardware implementations of logical connectives. Boolean algebra — rooted in propositional logic — is used to design, analyze, and minimize digital circuits in processors, memory, and all digital devices.", bg:LGN,bd:GN},
      {t:"💾  Database Query Optimization",
       b:"SQL WHERE clauses use AND, OR, NOT operators: WHERE (age > 18) AND (city = 'Hyderabad'). Database engines apply De Morgan's Laws and other equivalences to transform queries into faster, equivalent forms automatically.",bg:LBL,bd:BL},
      {t:"💻  Programming & Software Testing",
       b:"Every if-else and loop condition is propositional logic in code. Software testers use truth tables to design test cases ensuring every logical combination of conditions is exercised — a technique called Condition Coverage Testing.",bg:LMR,bd:M},
      {t:"🤖  Artificial Intelligence",
       b:"Expert systems use propositional logic for knowledge representation and rule-based inference. Automated theorem provers use Resolution Refutation — a proof technique built entirely on propositional logic — to verify program correctness.",bg:LYL,bd:AM}
    ];
    apps.forEach((app,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=2.20+row*2.30;
      box(p,s,bx,by,6.05,2.14,app.bg,app.bd);
      boxTitle(s,bx,by,6.05,app.t,app.bd);
      boxBody(s,bx,by,6.05,2.14,app.b);
    });
    s.addNotes(
"Slide 6 – Applications of Propositional Logic in Computer Science\n\n"+
"Propositional logic is far from an abstract exercise — it has direct, essential applications across computer science.\n\n"+
"Digital Circuit Design: Every logic gate in a computer chip is a physical implementation of a logical connective. The AND gate outputs 1 only when both inputs are 1. The OR gate outputs 1 when at least one input is 1. The NOT gate inverts the input. Engineers use Boolean algebra — the mathematical system of propositional logic — to design circuits using a technique called Karnaugh Map minimization, which reduces the number of gates needed, saving space, power, and cost.\n\n"+
"Database Query Optimization: When you write a SQL query with a WHERE clause, you are writing propositional logic. Database management systems internally apply logical equivalences, including De Morgan's Laws, to rewrite your query into an equivalent but more efficient form before execution.\n\n"+
"Programming and Software Testing: Every conditional statement in any programming language is propositional logic. The technique of Condition Coverage Testing uses truth tables to systematically verify that all logical combinations of conditions in the code are tested, helping catch hard-to-find bugs.\n\n"+
"Artificial Intelligence: Expert systems — like early medical diagnosis systems — store knowledge as propositional rules of the form IF condition THEN conclusion. Automated theorem provers use Resolution Refutation, which is a propositional logic-based algorithm, to formally verify the correctness of software and hardware systems.");
  }

  thankYouSlide(p, H);
  await p.writeFile({fileName:"./outputs/PPT1_Unit1_Propositional_Logic.pptx"});
  console.log("✔  PPT1 written");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – RELATIONS AND THEIR PROPERTIES  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Mathematical Foundations of Computer Science";

  titleSlide(p, H, "Unit II: Basic Structures – Sets, Functions, Relations", "Relations and Their Properties");

  // Slide 2 ── What is a Relation?
  {
    const s = contentSlide(p, H, "What is a Relation?");
    s.addText([
      {text:"A binary relation R from set A to set B ",options:{bold:true}},
      {text:"is a subset of the Cartesian product A × B. "+
            "It specifies which elements of A are 'related to' which elements of B.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.55, fontSize:13.5,fontFace:"Arial",color:BK});

    box(p,s,CX,2.34,5.90,4.22,LBL,BL);
    s.addText("Definition and Notation",
      {x:0.62,y:2.42,w:5.70,h:0.38, fontSize:14,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"R ⊆ A × B",options:{bold:true,color:BL,breakLine:true}},
      {text:"(a, b) ∈ R  means  a is related to b",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Relation ON a set A:",options:{bold:true,breakLine:true}},
      {text:"R ⊆ A × A  (both elements from same set)",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Example (on A = {1, 2, 3}):",options:{bold:true,breakLine:true}},
      {text:"R = \"less than\" = {(1,2),(1,3),(2,3)}",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Representations:",options:{bold:true,breakLine:true}},
      {text:"  • Set of ordered pairs",options:{breakLine:true}},
      {text:"  • Boolean matrix (1 if related, 0 if not)",options:{breakLine:true}},
      {text:"  • Arrow (directed graph) diagram",options:{}}
    ],{x:0.62,y:2.84,w:5.70,h:3.65, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    box(p,s,6.60,2.34,6.25,4.22,LMR,M);
    s.addText("Concrete Example",
      {x:6.72,y:2.42,w:6.03,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Let A = {1, 2, 3}  and  R = \"divides\"",options:{bold:true,breakLine:true}},
      {text:"aRb  means  a divides b evenly",options:{italic:true,breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Pairs in R:",options:{bold:true,breakLine:true}},
      {text:"(1,1): 1 divides 1  ✔",options:{breakLine:true}},
      {text:"(1,2): 1 divides 2  ✔",options:{breakLine:true}},
      {text:"(1,3): 1 divides 3  ✔",options:{breakLine:true}},
      {text:"(2,2): 2 divides 2  ✔",options:{breakLine:true}},
      {text:"(3,3): 3 divides 3  ✔",options:{breakLine:true}},
      {text:"(2,3): 2 divides 3  ✘  (not in R)",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"R = {(1,1),(1,2),(1,3),(2,2),(3,3)}",options:{bold:true}}
    ],{x:6.72,y:2.84,w:6.03,h:3.65, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 2 – What is a Relation?\n\n"+
"A binary relation from set A to set B is a subset of the Cartesian product A×B. Recall that A×B is the set of all ordered pairs (a,b) where a comes from A and b comes from B. The relation R specifies which of those pairs are considered 'related'.\n\n"+
"We write (a,b) ∈ R, or equivalently aRb, to say 'a is related to b'. The number of possible relations from A to B is 2 raised to the power |A|×|B|, since each ordered pair can either be in the relation or not.\n\n"+
"When a relation is from a set A to itself — that is R ⊆ A×A — we call it a 'relation on A'. Most of the interesting properties we study (reflexive, symmetric, etc.) apply to relations on a set.\n\n"+
"In our example, A = {1,2,3} and R is the 'divides' relation. We have (1,1) in R because 1 divides 1, (1,2) because 1 divides 2, and so on. The pair (2,3) is NOT in R because 2 does not divide 3 evenly.\n\n"+
"A relation can be represented in three ways: as a set of ordered pairs (as we just did), as a Boolean matrix where entry (i,j) is 1 if (i,j)∈R and 0 otherwise, or as a directed graph (arrow diagram) where we draw an arrow from a to b whenever aRb.\n\n"+
"All three representations are equivalent — they describe the same information in different visual forms. The matrix representation is especially useful in computing the transitive closure of a relation.");
  }

  // Slide 3 ── Five Key Properties
  {
    const s = contentSlide(p, H, "Five Key Properties of Relations");
    s.addText("These properties are defined for a relation R on a set A:",
      {x:CX,y:CYC,w:CW,h:0.38, fontSize:14,fontFace:"Arial",color:BK});
    const tData=[
      [
        {text:"Property",options:{bold:true,fill:{color:M},color:WH,align:"center"}},
        {text:"Formal Definition",options:{bold:true,fill:{color:M},color:WH,align:"center"}},
        {text:"Real Example",options:{bold:true,fill:{color:M},color:WH,align:"center"}},
        {text:"Counter-example",options:{bold:true,fill:{color:M},color:WH,align:"center"}}
      ],
      [
        {text:"Reflexive",options:{bold:true,fill:{color:LGN},color:DGN}},
        {text:"∀a ∈ A: (a,a) ∈ R",options:{fill:{color:LGN}}},
        {text:"\"≤\" on integers  (1≤1, 2≤2, …)",options:{fill:{color:LGN}}},
        {text:"\"<\" on integers  (1<1 is false)",options:{fill:{color:LGN}}}
      ],
      [
        {text:"Irreflexive",options:{bold:true,fill:{color:LMR},color:M}},
        {text:"∀a ∈ A: (a,a) ∉ R",options:{fill:{color:LMR}}},
        {text:"\"<\" on integers  (no a < a)",options:{fill:{color:LMR}}},
        {text:"\"≤\" includes (a,a) pairs",options:{fill:{color:LMR}}}
      ],
      [
        {text:"Symmetric",options:{bold:true,fill:{color:LBL},color:BL}},
        {text:"(a,b)∈R ⟹ (b,a)∈R",options:{fill:{color:LBL}}},
        {text:"\"sibling of\" among people",options:{fill:{color:LBL}}},
        {text:"\"parent of\" — not symmetric",options:{fill:{color:LBL}}}
      ],
      [
        {text:"Antisymmetric",options:{bold:true,fill:{color:LYL},color:AM}},
        {text:"(a,b)∈R ∧ (b,a)∈R ⟹ a=b",options:{fill:{color:LYL}}},
        {text:"\"≤\" on integers",options:{fill:{color:LYL}}},
        {text:"\"sibling of\" (both aRb and bRa, a≠b)",options:{fill:{color:LYL}}}
      ],
      [
        {text:"Transitive",options:{bold:true,fill:{color:LGR},color:GY}},
        {text:"(a,b)∈R ∧ (b,c)∈R ⟹ (a,c)∈R",options:{fill:{color:LGR}}},
        {text:"\"<\" on integers  (if a<b and b<c then a<c)",options:{fill:{color:LGR}}},
        {text:"\"parent of\" — not transitive",options:{fill:{color:LGR}}}
      ]
    ];
    s.addTable(tData,
      {x:CX,y:2.16,w:CW,h:5.02, fontSize:12.5,fontFace:"Arial",
       border:{pt:1,color:"CCCCCC"},colW:[2.2,3.2,3.7,3.25]});
    s.addNotes(
"Slide 3 – Five Key Properties of Relations\n\n"+
"We study five fundamental properties that a relation on a set A may or may not satisfy.\n\n"+
"REFLEXIVE: Every element is related to itself. Formally, for all a in A, the pair (a,a) must be in R. The relation 'less than or equal to' on integers is reflexive because every integer is ≤ itself. In contrast, 'strictly less than' is NOT reflexive because no integer is strictly less than itself.\n\n"+
"IRREFLEXIVE: The opposite — NO element is related to itself. The pair (a,a) must NOT be in R for any a. 'Strictly less than' is irreflexive.\n\n"+
"SYMMETRIC: If a is related to b, then b must also be related to a. The 'sibling of' relation is symmetric — if Alice is Bob's sibling, then Bob is Alice's sibling. 'Parent of' is NOT symmetric — if Alice is Bob's parent, Bob is not Alice's parent.\n\n"+
"ANTISYMMETRIC: If both (a,b) and (b,a) are in R, then a and b must be the same element. This does NOT mean symmetric and antisymmetric are opposites — a relation can be both, neither, or one. 'Less than or equal to' is antisymmetric: if a≤b and b≤a, then a=b.\n\n"+
"TRANSITIVE: If a is related to b and b is related to c, then a must be related to c. 'Less than' is transitive: if a<b and b<c, then a<c. 'Parent of' is NOT transitive — Alice's parent's parent is Alice's grandparent, not another parent.");
  }

  // Slide 4 ── Equivalence Relations
  {
    const s = contentSlide(p, H, "Equivalence Relations");
    s.addText([
      {text:"A relation R on set A is an Equivalence Relation ",options:{bold:true}},
      {text:"if and only if it is ",options:{}},
      {text:"Reflexive, Symmetric, AND Transitive",options:{bold:true,color:M}},
      {text:" — all three together.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.52, fontSize:13.5,fontFace:"Arial",color:BK});

    // Three conditions
    const conds=[
      {t:"1.  Reflexive",   d:"∀a ∈ A: (a,a) ∈ R",  bg:LGN,bd:GN},
      {t:"2.  Symmetric",   d:"(a,b)∈R ⟹ (b,a)∈R",  bg:LBL,bd:BL},
      {t:"3.  Transitive",  d:"(a,b),(b,c)∈R ⟹ (a,c)∈R", bg:LMR,bd:M}
    ];
    conds.forEach((c,i)=>{
      const bx=CX+i*4.15;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:bx,y:2.34,w:4.0,h:1.05, fill:{color:c.bg}, line:{color:c.bd,width:1.5},rectRadius:0.1});
      s.addText(c.t,{x:bx+0.10,y:2.38,w:3.80,h:0.38, fontSize:14,bold:true,color:c.bd,fontFace:"Arial",margin:0});
      s.addText(c.d,{x:bx+0.10,y:2.78,w:3.80,h:0.55, fontSize:13,color:BK,fontFace:"Arial",margin:0});
    });

    // Main example
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:3.52,w:CW,h:3.04, fill:{color:LYL}, line:{color:AM,width:1.5},rectRadius:0.1});
    s.addText("Classic Example: Congruence Modulo n",
      {x:0.62,y:3.60,w:CW-0.24,h:0.40, fontSize:15,bold:true,color:AM,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Define:  a ≡ b (mod n)  iff  n divides (a − b)",options:{bold:true,breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"Example with n = 3:",options:{bold:true,breakLine:true}},
      {text:"   7 ≡ 1 (mod 3)  because  3 divides (7−1) = 6  ✔",options:{breakLine:true}},
      {text:"   10 ≡ 4 (mod 3)  because  3 divides (10−4) = 6  ✔",options:{breakLine:true}},
      {text:" ",options:{fontSize:6,breakLine:true}},
      {text:"Equivalence Classes:",options:{bold:true,breakLine:true}},
      {text:"   [0] = {…, −6, −3, 0, 3, 6, 9, …}     [1] = {…, −5, −2, 1, 4, 7, 10, …}     [2] = {…, −4, −1, 2, 5, 8, 11, …}",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Key Theorem: An equivalence relation partitions the set into disjoint equivalence classes.",options:{bold:true}}
    ],{x:0.62,y:4.04,w:CW-0.24,h:2.48, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 4 – Equivalence Relations\n\n"+
"An equivalence relation is a relation that satisfies all three properties simultaneously: reflexive, symmetric, AND transitive. It is one of the most important structures in all of mathematics.\n\n"+
"The most important example is congruence modulo n. We say a is congruent to b modulo n, written a ≡ b (mod n), if n divides the difference a−b. For n=3: 7 ≡ 1 (mod 3) because 3 divides 6. And 10 ≡ 4 (mod 3) because 3 divides 6 again.\n\n"+
"Let us verify it is an equivalence relation:\n- Reflexive: a ≡ a (mod n) because n divides a−a = 0. ✔\n- Symmetric: if n divides (a−b), then n divides (b−a) = −(a−b). ✔\n- Transitive: if n|(a−b) and n|(b−c), then n|((a−b)+(b−c)) = (a−c). ✔\n\n"+
"The key theorem about equivalence relations: Every equivalence relation partitions the set into disjoint equivalence classes — groups of elements that are all related to each other. For congruence mod 3, the three classes are all multiples of 3, all numbers ≡ 1 mod 3, and all numbers ≡ 2 mod 3. These classes are disjoint (no overlap) and exhaustive (every integer belongs to exactly one class).\n\n"+
"Other examples of equivalence relations: same nationality, same birthday (date, not year), having the same remainder when divided by k.");
  }

  // Slide 5 ── Partial Orderings
  {
    const s = contentSlide(p, H, "Partial Orderings");
    s.addText([
      {text:"A Partial Ordering ",options:{bold:true}},
      {text:"is a relation R on set A that is ",options:{}},
      {text:"Reflexive, Antisymmetric, AND Transitive",options:{bold:true,color:M}},
      {text:". The pair (A, R) is called a partially ordered set (POSET).",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.55, fontSize:13.5,fontFace:"Arial",color:BK});

    box(p,s,CX,2.35,5.90,4.22,LGN,GN);
    s.addText("Example: Divisibility on {1,2,3,4,6,12}",
      {x:0.62,y:2.43,w:5.70,h:0.38, fontSize:13.5,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"a | b  means  a divides b",options:{bold:true,breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Verify:",options:{bold:true,breakLine:true}},
      {text:"  ✔ Reflexive:    a | a  (every number divides itself)",options:{breakLine:true}},
      {text:"  ✔ Antisymmetric: if a|b and b|a, then a=b",options:{breakLine:true}},
      {text:"  ✔ Transitive:    if a|b and b|c, then a|c",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Hasse Diagram (removes reflexive & transitive edges):",options:{bold:true,breakLine:true}},
      {text:"      12",options:{breakLine:true}},
      {text:"     /  \\",options:{breakLine:true}},
      {text:"    4    6",options:{breakLine:true}},
      {text:"    |   / \\",options:{breakLine:true}},
      {text:"    2  3   2",options:{breakLine:true}},
      {text:"     \\ /",options:{breakLine:true}},
      {text:"      1  ← minimal element",options:{}}
    ],{x:0.62,y:2.85,w:5.70,h:3.65, fontSize:12,fontFace:"Arial",color:BK,margin:0,fontFace:"Courier New"});

    box(p,s,6.60,2.35,6.25,4.22,LMR,M);
    s.addText("Key Terminology",
      {x:6.72,y:2.43,w:6.03,h:0.38, fontSize:13.5,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Minimal Element:",options:{bold:true,breakLine:true}},
      {text:"   No element precedes it.",options:{breakLine:true}},
      {text:"   In divisibility: 1 (nothing divides 1 except 1 itself)",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Maximal Element:",options:{bold:true,breakLine:true}},
      {text:"   No element comes after it.",options:{breakLine:true}},
      {text:"   In divisibility on {1..12}: 12",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Total Order (Linear Order):",options:{bold:true,breakLine:true}},
      {text:"   Every pair is comparable.",options:{breakLine:true}},
      {text:"   Example: ≤ on integers",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Partial Order:",options:{bold:true,breakLine:true}},
      {text:"   Some pairs may be incomparable.",options:{breakLine:true}},
      {text:"   e.g., 4 and 6 in divisibility: neither 4|6 nor 6|4.",options:{}}
    ],{x:6.72,y:2.85,w:6.03,h:3.65, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 5 – Partial Orderings\n\n"+
"A partial ordering is a relation on a set A that satisfies reflexive, antisymmetric, AND transitive properties. Together with the set, it forms a partially ordered set or POSET.\n\n"+
"The classic example is the divisibility relation on a set of positive integers. We say a divides b (written a|b) if there exists an integer k such that b = k×a. Let us verify: Reflexive — every number divides itself (a = 1×a). Antisymmetric — if a|b and b|a, both must be positive, so a=b. Transitive — if a|b and b|c, then c = kb×(ka×a), so a|c.\n\n"+
"The Hasse diagram is a visual tool to represent a partial order. We draw it by removing the arrows that are implied by reflexivity (self-loops) and transitivity (skip-over arrows), keeping only the 'covering' relations. The result is a clean hierarchical diagram. In the divisibility example, 1 is at the bottom (minimal element — it divides everything), 12 is at the top (maximal), and intermediate elements connect by direct divisibility.\n\n"+
"A key distinction: in a PARTIAL order, some pairs may be INCOMPARABLE — neither a≤b nor b≤a. For example, 4 and 6 in the divisibility order: 4 does not divide 6, and 6 does not divide 4. In a TOTAL (or linear) order, every pair is comparable — like ≤ on integers.\n\n"+
"Partial orderings are used in task scheduling, version control systems (commit history), and type hierarchies in object-oriented programming.");
  }

  // Slide 6 ── Applications
  {
    const s = contentSlide(p, H, "Applications in Computer Science");
    s.addText("Relations are fundamental data structures underlying many CS systems:",
      {x:CX,y:CYC,w:CW,h:0.40, fontSize:14,fontFace:"Arial",color:BK});
    const apps=[
      {t:"🗄  Relational Databases",
       b:"Every table in a relational database IS a relation (set of tuples). SQL joins combine relations using Cartesian product and selection. Database normalization is based on functional dependencies — a special type of relation.",bg:LBL,bd:BL},
      {t:"🌐  Social Networks",
       b:"'Friend', 'follows', and 'blocks' are all relations on a set of users. Symmetric relations (friend) vs. asymmetric (follow). Transitivity of 'knows' gives 'Six degrees of separation'. Equivalence classes group users in communities.",bg:LGN,bd:GN},
      {t:"🔄  Compiler Design",
       b:"Variable dependencies in code form a relation. Computing the transitive closure finds all indirect dependencies. Type compatibility (subtype relation) is a partial order. Operator precedence is a partial ordering of operators.",bg:LMR,bd:M},
      {t:"📅  Task Scheduling",
       b:"Project tasks have 'must complete before' constraints — a partial order (DAG). Topological sort of the Hasse diagram gives a valid execution order. Critical path analysis uses this POSET structure to minimize project duration.",bg:LYL,bd:AM}
    ];
    apps.forEach((app,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=2.20+row*2.30;
      box(p,s,bx,by,6.05,2.14,app.bg,app.bd);
      boxTitle(s,bx,by,6.05,app.t,app.bd);
      boxBody(s,bx,by,6.05,2.14,app.b);
    });
    s.addNotes(
"Slide 6 – Applications of Relations in Computer Science\n\n"+
"Relations are not just theoretical — they appear in every major area of computer science.\n\n"+
"Relational Databases: Every table in a SQL database IS a mathematical relation — a set of tuples. The relational model, invented by E.F. Codd at IBM, directly applies set theory and relation theory. SQL operations like JOIN and SELECT correspond to relational algebra operations. Database normalization theory (1NF, 2NF, 3NF, BCNF) is entirely based on functional dependency — a special kind of relation between attributes.\n\n"+
"Social Networks: All connections between users in social platforms are modeled as relations. Facebook's 'friend' relation is symmetric. Twitter's 'follows' relation is asymmetric. The concept of equivalence classes is used to identify communities or clusters of densely connected users.\n\n"+
"Compiler Design: When a compiler analyzes your code, it builds a dependency relation between variables. Computing the transitive closure of this relation finds all indirect dependencies — crucial for optimization. The type hierarchy in object-oriented languages (int is a subtype of Number) forms a partial order used for type checking.\n\n"+
"Task Scheduling: In project management, tasks have 'must be completed before' constraints. This forms a partial order — specifically a Directed Acyclic Graph. Topological sorting finds a valid execution order. Build systems like Make and Gradle use this to compile code in the correct order.");
  }

  thankYouSlide(p, H);
  await p.writeFile({fileName:"./outputs/PPT2_Unit2_Relations_Properties.pptx"});
  console.log("✔  PPT2 written");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – MATHEMATICAL INDUCTION  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Mathematical Foundations of Computer Science";

  titleSlide(p, H, "Unit III: Algorithms, Induction and Recursion", "Mathematical Induction");

  // Slide 2 ── Introduction
  {
    const s = contentSlide(p, H, "What is Mathematical Induction?");
    s.addText([
      {text:"Mathematical Induction ",options:{bold:true}},
      {text:"is a proof technique used to establish that a statement P(n) is true for "+
            "ALL natural numbers n ≥ n₀.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.52, fontSize:13.5,fontFace:"Arial",color:BK});

    // Domino analogy box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:2.33,w:CW,h:1.72, fill:{color:LYL}, line:{color:AM,width:2},rectRadius:0.1});
    s.addText("The Domino Analogy",
      {x:0.62,y:2.41,w:CW-0.24,h:0.38, fontSize:15,bold:true,color:AM,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Imagine an infinite row of dominoes labelled 1, 2, 3, …",options:{breakLine:true}},
      {text:"If  (1) the FIRST domino falls, and (2) whenever a domino falls it knocks the NEXT one — ",options:{breakLine:true}},
      {text:"then ALL dominoes fall.",options:{bold:true}}
    ],{x:0.62,y:2.83,w:CW-0.24,h:1.14, fontSize:13.5,fontFace:"Arial",color:BK,margin:0});

    // Two-column comparison
    box(p,s,CX,4.18,5.90,2.40,LGN,GN);
    s.addText("Mathematical Induction IS:",
      {x:0.62,y:4.26,w:5.70,h:0.38, fontSize:14,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"  ✔  A rigorous mathematical proof technique",options:{breakLine:true}},
      {text:"  ✔  Proves statements for all n ≥ n₀",options:{breakLine:true}},
      {text:"  ✔  Uses logical deduction (deductive reasoning)",options:{breakLine:true}},
      {text:"  ✔  100% certain conclusion",options:{}}
    ],{x:0.62,y:4.68,w:5.70,h:1.84, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    box(p,s,6.60,4.18,6.25,2.40,LMR,M);
    s.addText("Mathematical Induction is NOT:",
      {x:6.72,y:4.26,w:6.03,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"  ✘  Inductive reasoning (science/empirical)",options:{breakLine:true}},
      {text:"  ✘  \"It works for first 100 cases, so always\"",options:{breakLine:true}},
      {text:"  ✘  A pattern guess or conjecture",options:{breakLine:true}},
      {text:"  ✘  Trial and error",options:{}}
    ],{x:6.72,y:4.68,w:6.03,h:1.84, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 2 – What is Mathematical Induction?\n\n"+
"Mathematical induction is a powerful proof technique used to prove that some statement P(n) is true for every natural number n greater than or equal to some starting value n₀, usually n₀ = 1 or n₀ = 0.\n\n"+
"The domino analogy is the best way to understand it. Imagine an infinite line of dominoes numbered 1, 2, 3, and so on. Now suppose two things are true: first, the first domino (Domino 1) falls. Second, whenever any domino k falls, it knocks over the next domino k+1. Under these two conditions, ALL dominoes will eventually fall — even domino number one million. This is exactly how mathematical induction works.\n\n"+
"A critical distinction: Mathematical induction is a DEDUCTIVE proof technique, not the 'inductive reasoning' used in science where we observe patterns and make generalizations. In mathematics, observing that a statement is true for the first million values does NOT constitute a proof — we need the formal inductive argument. There are famous examples where a pattern holds for millions of cases but fails eventually.\n\n"+
"For example, the expression n² + n + 41 is prime for n = 0, 1, 2, ..., 39 — forty consecutive values! But for n = 40, it gives 40² + 40 + 41 = 1681 = 41², which is NOT prime. Without a formal proof, we cannot trust patterns.");
  }

  // Slide 3 ── The Two Steps
  {
    const s = contentSlide(p, H, "The Two Steps of Mathematical Induction");
    s.addText("To prove P(n) is true for all n ≥ 1, we need exactly TWO steps:",
      {x:CX,y:CYC,w:CW,h:0.40, fontSize:14,fontFace:"Arial",color:BK});

    // Step 1
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:2.22,w:CW,h:1.85, fill:{color:LGN}, line:{color:GN,width:2},rectRadius:0.1});
    s.addText("Step 1:  BASE CASE  (The first domino falls)",
      {x:0.62,y:2.30,w:CW-0.24,h:0.45, fontSize:16,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Prove that P(1) is true.",options:{bold:true,breakLine:true}},
      {text:"Directly verify the statement for the smallest value n₀ (usually n₀ = 1).",options:{breakLine:true}},
      {text:"This is the 'anchor' — it grounds the proof in a concrete fact.",options:{italic:true}}
    ],{x:0.62,y:2.78,w:CW-0.24,h:1.22, fontSize:13.5,fontFace:"Arial",color:BK,margin:0});

    // Step 2
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:4.18,w:CW,h:2.06, fill:{color:LBL}, line:{color:BL,width:2},rectRadius:0.1});
    s.addText("Step 2:  INDUCTIVE STEP  (Each domino knocks the next)",
      {x:0.62,y:4.26,w:CW-0.24,h:0.45, fontSize:16,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Inductive Hypothesis (IH):",options:{bold:true}},
      {text:"  Assume P(k) is true for some arbitrary k ≥ 1.",options:{breakLine:true}},
      {text:"Goal:",options:{bold:true}},
      {text:"  Using the IH, prove that P(k+1) is also true.",options:{breakLine:true}},
      {text:"The IH is the key tool — we USE P(k) as a 'known fact' to derive P(k+1).",options:{italic:true}}
    ],{x:0.62,y:4.76,w:CW-0.24,h:1.42, fontSize:13.5,fontFace:"Arial",color:BK,margin:0});

    // Conclusion box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:6.35,w:CW,h:0.75, fill:{color:LMR}, line:{color:M,width:2},rectRadius:0.1});
    s.addText(
      "✔  CONCLUSION: If both steps succeed, then P(n) is true for ALL n ≥ 1  (by the Principle of Mathematical Induction)",
      {x:0.62,y:6.38,w:CW-0.24,h:0.68, fontSize:13,bold:true,color:M,fontFace:"Arial",margin:0});

    s.addNotes(
"Slide 3 – The Two Steps of Mathematical Induction\n\n"+
"The formal proof by mathematical induction has exactly two steps.\n\n"+
"STEP 1 — BASE CASE: We directly prove that the statement P(n) is true for the initial value n₀, which is usually 1 (or sometimes 0). This is straightforward verification — just substitute n=1 into the statement and check. This step is like pushing the first domino over. Without the base case, the proof is incomplete and invalid.\n\n"+
"STEP 2 — INDUCTIVE STEP: This is the heart of the proof. We assume the statement is true for SOME arbitrary k — this assumption is called the Inductive Hypothesis (IH). Then, using the IH as a given fact, we must PROVE that P(k+1) is also true. This step is like showing that whenever domino k falls, it will knock over domino k+1.\n\n"+
"The logic is: the base case gives us P(1). The inductive step, applied to k=1, gives us P(2). Applying it again to k=2 gives P(3). And so on — we get P(n) for every natural number n.\n\n"+
"CONCLUSION: Once both steps are proven, we can conclude P(n) is true for ALL n ≥ 1 by the Principle of Mathematical Induction.\n\n"+
"The crucial skill is in the inductive step: you must write 'Assume P(k) is true' (the hypothesis), then express P(k+1) in terms of P(k), and substitute the hypothesis to complete the argument.");
  }

  // Slide 4 ── Setting Up the Proof
  {
    const s = contentSlide(p, H, "Example: Setting Up the Proof");
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:CYC,w:CW,h:0.85, fill:{color:LYL}, line:{color:AM,width:2},rectRadius:0.1});
    s.addText(
      "Theorem:  1 + 2 + 3 + ⋯ + n  =  n(n+1)/2     for all integers n ≥ 1",
      {x:0.62,y:CYC+0.08,w:CW-0.24,h:0.68, fontSize:16,bold:true,color:AM,
       align:"center",valign:"middle",fontFace:"Arial",margin:0});

    // Base case box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:2.72,w:5.90,h:2.10, fill:{color:LGN}, line:{color:GN,width:1.5},rectRadius:0.1});
    s.addText("Step 1: Base Case  (n = 1)",
      {x:0.62,y:2.80,w:5.70,h:0.38, fontSize:14,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"LHS (Left-Hand Side):",options:{bold:true,breakLine:true}},
      {text:"   Sum = 1",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"RHS (Right-Hand Side):",options:{bold:true,breakLine:true}},
      {text:"   1 × (1+1) / 2 = 1 × 2 / 2 = 1",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"LHS = RHS = 1  ✔  Base Case holds!",options:{bold:true,color:DGN}}
    ],{x:0.62,y:3.22,w:5.70,h:1.52, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    // Inductive hypothesis box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:6.60,y:2.72,w:6.25,h:2.10, fill:{color:LBL}, line:{color:BL,width:1.5},rectRadius:0.1});
    s.addText("Step 2: Inductive Hypothesis",
      {x:6.72,y:2.80,w:6.03,h:0.38, fontSize:14,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Assume P(k) is true:",options:{bold:true,breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"   1 + 2 + 3 + ⋯ + k  =  k(k+1)/2",options:{bold:true,color:BL,breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"We call this the Inductive Hypothesis (IH).",options:{italic:true,breakLine:true}},
      {text:"We will USE this in the next step.",options:{italic:true}}
    ],{x:6.72,y:3.22,w:6.03,h:1.52, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    // Goal box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:4.95,w:CW,h:1.66, fill:{color:LMR}, line:{color:M,width:1.5},rectRadius:0.1});
    s.addText("Goal of the Inductive Step: Prove P(k+1)",
      {x:0.62,y:5.03,w:CW-0.24,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"We need to show: 1 + 2 + ⋯ + k + (k+1)  =  ",options:{}},
      {text:"(k+1)(k+2)/2",options:{bold:true,color:M}}
    ],{x:0.62,y:5.45,w:CW-0.24,h:1.10, fontSize:14,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 4 – Example: Setting Up the Proof\n\n"+
"We will prove by mathematical induction that the sum 1 + 2 + 3 + ... + n equals n(n+1)/2 for all integers n ≥ 1. This is one of the most famous formulas in mathematics — the Gaussian sum formula.\n\n"+
"STEP 1 — BASE CASE (n=1):\nWe verify the formula for n=1. The left-hand side is just 1 (the sum of the first 1 integer). The right-hand side is 1×(1+1)/2 = 1×2/2 = 1. Both sides equal 1, so the base case holds.\n\n"+
"STEP 2 — INDUCTIVE STEP:\nFirst, we state the Inductive Hypothesis (IH): Assume P(k) is true for some k ≥ 1. This means we assume that 1 + 2 + 3 + ... + k equals k(k+1)/2. We ASSUME this — we don't prove it here. It is our working assumption.\n\n"+
"Our GOAL is to prove P(k+1): that 1 + 2 + ... + k + (k+1) equals (k+1)(k+2)/2.\n\n"+
"Notice the pattern: P(k+1) is what you get by replacing n with k+1 in the original formula. The right-hand side becomes (k+1)((k+1)+1)/2 = (k+1)(k+2)/2. We need to show the left-hand side equals this.\n\n"+
"The proof continues on the next slide where we use the Inductive Hypothesis to complete the algebra.");
  }

  // Slide 5 ── Completing the Proof
  {
    const s = contentSlide(p, H, "Completing the Inductive Proof");
    s.addText("We prove P(k+1):  1 + 2 + ⋯ + k + (k+1)  =  (k+1)(k+2)/2",
      {x:CX,y:CYC,w:CW,h:0.42, fontSize:14,bold:true,color:M,fontFace:"Arial"});

    const steps=[
      {label:"Start:", expr:"1 + 2 + ⋯ + k + (k+1)", note:"(The left-hand side of P(k+1))", bg:LGR},
      {label:"Group the first k terms:", expr:"= [ 1 + 2 + ⋯ + k ]  +  (k+1)", note:"(Separate the last term)", bg:LBL},
      {label:"Apply the IH:", expr:"= k(k+1)/2  +  (k+1)", note:"(IH says the bracket = k(k+1)/2)", bg:LGN},
      {label:"Factor out (k+1):", expr:"= (k+1) × [ k/2 + 1 ]", note:"(Common factor)", bg:LYL},
      {label:"Simplify:", expr:"= (k+1) × (k+2)/2  =  (k+1)(k+2)/2  ✔", note:"(This is exactly P(k+1)!)", bg:LMR}
    ];
    steps.forEach((st,i)=>{
      const y=2.25+i*0.88;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:CX,y,w:CW,h:0.80, fill:{color:st.bg}, line:{color:"CCCCCC",width:1},rectRadius:0.08});
      s.addText([
        {text:st.label+"  ",options:{bold:true,color:M}},
        {text:st.expr,options:{color:BK}},
        {text:"    "+st.note,options:{italic:true,color:GY,fontSize:11}}
      ],{x:0.62,y:y+0.08,w:CW-0.24,h:0.64, fontSize:13.5,fontFace:"Arial",margin:0});
    });

    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:6.64,w:CW,h:0.68, fill:{color:LGN}, line:{color:GN,width:2},rectRadius:0.1});
    s.addText(
      "✔  CONCLUSION: Both steps are complete. By the Principle of Mathematical Induction, "+
      "1+2+⋯+n = n(n+1)/2  for ALL n ≥ 1.",
      {x:0.62,y:6.67,w:CW-0.24,h:0.62, fontSize:12.5,bold:true,color:GN,fontFace:"Arial",margin:0});

    s.addNotes(
"Slide 5 – Completing the Inductive Proof\n\n"+
"Now we complete the inductive step. We need to prove P(k+1): that 1 + 2 + ... + k + (k+1) = (k+1)(k+2)/2.\n\n"+
"Step 1 – Start: Write down the left-hand side of P(k+1), which is 1 + 2 + ... + k + (k+1).\n\n"+
"Step 2 – Group: Separate the last term by bracketing the first k terms: [1 + 2 + ... + k] + (k+1).\n\n"+
"Step 3 – Apply the Inductive Hypothesis: The bracket [1 + 2 + ... + k] is exactly what our Inductive Hypothesis says equals k(k+1)/2. So we substitute to get k(k+1)/2 + (k+1).\n\n"+
"Step 4 – Factor: Both terms have the common factor (k+1). Factor it out to get (k+1) × [k/2 + 1].\n\n"+
"Step 5 – Simplify: [k/2 + 1] = (k+2)/2. So the expression becomes (k+1)(k+2)/2, which is EXACTLY the right-hand side of P(k+1).\n\n"+
"The left-hand side and right-hand side of P(k+1) are equal, so the inductive step is proven!\n\n"+
"CONCLUSION: We proved the base case P(1), and we proved that P(k) implies P(k+1). By the Principle of Mathematical Induction, P(n) is true for ALL n ≥ 1. The formula 1+2+...+n = n(n+1)/2 is established for every positive integer n.\n\n"+
"Fun fact: The German mathematician Carl Friedrich Gauss discovered this formula as a child when his teacher asked him to add numbers 1 to 100. He instantly answered 5050 using the insight that pairs like (1+100), (2+99)... each sum to 101, and there are 50 such pairs.");
  }

  // Slide 6 ── Strong Induction + Applications
  {
    const s = contentSlide(p, H, "Strong Induction & Applications");
    s.addText([
      {text:"Strong Induction ",options:{bold:true}},
      {text:"(also called Complete Induction) assumes P(j) is true for ALL j ≤ k, "+
            "not just P(k). This stronger hypothesis is needed when P(k+1) depends on multiple previous cases.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.62, fontSize:13.5,fontFace:"Arial",color:BK});

    box(p,s,CX,2.45,6.05,2.12,LBL,BL);
    s.addText("Strong Induction Steps",
      {x:0.62,y:2.53,w:5.82,h:0.38, fontSize:14,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"1. Base Case(s):",options:{bold:true,breakLine:true}},
      {text:"   Prove P(1) [and possibly P(2), P(3), etc.]",options:{breakLine:true}},
      {text:"2. Strong IH:",options:{bold:true,breakLine:true}},
      {text:"   Assume P(j) is true for ALL j ≤ k",options:{breakLine:true}},
      {text:"3. Goal:",options:{bold:true,breakLine:true}},
      {text:"   Prove P(k+1) using any previous P(j), j ≤ k",options:{}}
    ],{x:0.62,y:2.95,w:5.82,h:1.55, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    box(p,s,6.60,2.45,6.25,2.12,LYL,AM);
    s.addText("Example: Fundamental Theorem of Arithmetic",
      {x:6.72,y:2.53,w:6.03,h:0.38, fontSize:13,bold:true,color:AM,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Theorem:",options:{bold:true,breakLine:true}},
      {text:"Every integer n > 1 is either prime or can be written as a product of primes.",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Why Strong Induction?",options:{bold:true,breakLine:true}},
      {text:"If n is composite, n = a×b where both a,b < n. We need induction applied to BOTH a and b.",options:{}}
    ],{x:6.72,y:2.95,w:6.03,h:1.55, fontSize:12,fontFace:"Arial",color:BK,margin:0});

    s.addText("Applications in CS:",
      {x:CX,y:4.70,w:CW,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial"});
    const apps2=[
      {t:"Algorithm Correctness", b:"Proving recursive algorithms (like Merge Sort, Binary Search) produce correct output for all inputs.",bg:LGN,bd:GN},
      {t:"Complexity Proofs",     b:"Proving that algorithms run in O(n log n) time uses recurrences solved by induction.",bg:LBL,bd:BL},
      {t:"Data Structure Props",  b:"Properties of heaps, trees, and balanced BSTs (height, node count) are proven by induction on height.",bg:LMR,bd:M},
      {t:"Formal Verification",   b:"Proving that program loops maintain invariants — induction on iteration count guarantees correctness.",bg:LYL,bd:AM}
    ];
    apps2.forEach((app,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=5.14+row*0.94;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:bx,y:by,w:6.05,h:0.82, fill:{color:app.bg}, line:{color:app.bd,width:1},rectRadius:0.08});
      s.addText([
        {text:app.t+":  ",options:{bold:true,color:app.bd}},
        {text:app.b,options:{color:BK}}
      ],{x:bx+0.10,y:by+0.08,w:5.85,h:0.66, fontSize:11.5,fontFace:"Arial",margin:0});
    });

    s.addNotes(
"Slide 6 – Strong Induction and Applications\n\n"+
"Strong induction (also called complete induction) is a more powerful variant of mathematical induction. In regular induction, the inductive hypothesis assumes only P(k). In strong induction, we assume P(j) is true for ALL j from 1 up to k.\n\n"+
"Why would we need this? Sometimes proving P(k+1) requires more than just knowing P(k) — we might need to use results for smaller values too. A classic example is the Fundamental Theorem of Arithmetic: every integer greater than 1 is either prime or a product of primes. If a number n is composite, we write n = a×b where both a and b are strictly less than n. To apply induction to both a and b, we need to assume the theorem is true for ALL values up to k, not just k.\n\n"+
"Applications of Mathematical Induction in Computer Science:\n\n"+
"Algorithm Correctness: Induction is the standard technique for proving recursive algorithms correct. For example, we prove Merge Sort correct by showing: base case (sorting 1 element is trivially correct) + inductive step (if we correctly sort two halves, merging gives the correct sort of the whole).\n\n"+
"Complexity Analysis: Proofs about running time often use recurrence relations solved by induction. The claim 'Binary search runs in O(log n) steps' is proved by induction on the size of the search space.\n\n"+
"Data Structure Properties: The claim 'A complete binary tree with n levels has 2^n - 1 nodes' is proven by induction on the number of levels.\n\n"+
"Loop Invariants: In formal program verification, a loop invariant is a property that holds before and after every iteration. We prove this by induction on the iteration count, guaranteeing the program's correctness.");
  }

  thankYouSlide(p, H);
  await p.writeFile({fileName:"./outputs/PPT3_Unit3_Mathematical_Induction.pptx"});
  console.log("✔  PPT3 written");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 4 – BAYES' THEOREM  (Unit IV)
// ═══════════════════════════════════════════════════════════════
async function ppt4() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Mathematical Foundations of Computer Science";

  titleSlide(p, H, "Unit IV: Discrete Probability and Advanced Counting Techniques", "Bayes' Theorem");

  // Slide 2 ── Conditional Probability
  {
    const s = contentSlide(p, H, "Conditional Probability — The Foundation");
    s.addText([
      {text:"Conditional Probability ",options:{bold:true}},
      {text:"P(A|B) measures the probability of event A occurring, given that event B has already occurred.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.52, fontSize:13.5,fontFace:"Arial",color:BK});

    // Formula box
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:2.35,w:CW,h:1.32, fill:{color:LBL}, line:{color:BL,width:2},rectRadius:0.1});
    s.addText("Definition:",
      {x:0.62,y:2.43,w:CW-0.24,h:0.38, fontSize:15,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"P(A | B)  =  P(A ∩ B) / P(B)",options:{bold:true,fontSize:20,color:M}},
      {text:"      provided P(B) > 0",options:{italic:true,color:GY}}
    ],{x:0.62,y:2.83,w:CW-0.24,h:0.78, fontSize:16,fontFace:"Arial",margin:0});

    // Example
    box(p,s,CX,3.80,5.90,2.80,LGN,GN);
    s.addText("Intuitive Example",
      {x:0.62,y:3.88,w:5.70,h:0.38, fontSize:14,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"A bag has: 3 Red balls, 2 Blue balls",options:{bold:true,breakLine:true}},
      {text:"Event A = 'draw a Red ball'",options:{breakLine:true}},
      {text:"Event B = 'draw a ball on Monday'",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"P(A) = 3/5 = 0.60",options:{breakLine:true}},
      {text:"P(A|B) = 3/5 = 0.60  (Monday doesn't help here)",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"But if B = 'a Red was drawn first (without replacement)':",options:{breakLine:true}},
      {text:"P(A|B) = 2/4 = 0.50  (reduced pool!)",options:{bold:true,color:DGN}}
    ],{x:0.62,y:4.30,w:5.70,h:2.22, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    box(p,s,6.60,3.80,6.25,2.80,LMR,M);
    s.addText("Key Properties",
      {x:6.72,y:3.88,w:6.03,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Multiplication Rule:",options:{bold:true,breakLine:true}},
      {text:"   P(A ∩ B) = P(A|B) × P(B) = P(B|A) × P(A)",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Independence:",options:{bold:true,breakLine:true}},
      {text:"   A and B are independent iff",options:{breakLine:true}},
      {text:"   P(A|B) = P(A)  ←  B gives no information about A",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Law of Total Probability:",options:{bold:true,breakLine:true}},
      {text:"   P(B) = P(B|A)·P(A) + P(B|¬A)·P(¬A)",options:{}}
    ],{x:6.72,y:4.30,w:6.03,h:2.22, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 2 – Conditional Probability\n\n"+
"Before we can understand Bayes' Theorem, we need to understand conditional probability.\n\n"+
"The conditional probability P(A|B) is read as 'the probability of A given B'. It answers the question: IF we know that event B has already occurred, what is the probability that A also occurs? Knowing B 'updates' our probability estimate for A.\n\n"+
"The formula is P(A|B) = P(A∩B) / P(B). We divide the probability of both events occurring by the probability of the conditioning event B. This is only defined when P(B) > 0, since we cannot condition on an impossible event.\n\n"+
"For example: imagine a bag with 3 red and 2 blue balls. If we draw one ball without replacement, and we're told the first draw was red (event B), then only 4 balls remain in the bag with only 2 being red. So the probability of drawing red in the second draw, given the first was red, is 2/4 = 0.5 — lower than the original 3/5.\n\n"+
"Two important related results: First, the Multiplication Rule rearranges the conditional probability formula to give P(A∩B) = P(A|B)·P(B). Second, the Law of Total Probability says P(B) can be broken down as P(B|A)·P(A) + P(B|¬A)·P(¬A). This will be used in the denominator of Bayes' Theorem.");
  }

  // Slide 3 ── Bayes' Theorem
  {
    const s = contentSlide(p, H, "Bayes' Theorem");
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:CYC,w:CW,h:2.35, fill:{color:LMR}, line:{color:M,width:2},rectRadius:0.1});
    s.addText("Bayes' Theorem  (Thomas Bayes, 1763)",
      {x:0.62,y:CYC+0.08,w:CW-0.24,h:0.40, fontSize:16,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"P(A | B)  =  P(B | A) × P(A)  /  P(B)",options:{bold:true,fontSize:22,color:M,breakLine:true}},
      {text:"\n",options:{fontSize:6,breakLine:true}},
      {text:"Expanded form (using Law of Total Probability):",options:{italic:true,breakLine:true}},
      {text:"P(A|B)  =  P(B|A) × P(A)  /  [ P(B|A)×P(A)  +  P(B|¬A)×P(¬A) ]",options:{bold:true,fontSize:15,color:BL}}
    ],{x:0.62,y:CYC+0.52,w:CW-0.24,h:1.75, fontSize:15,fontFace:"Arial",color:BK,margin:0});

    // Terminology boxes
    const terms=[
      {t:"P(A)  = Prior Probability",
       b:"Our INITIAL probability estimate for A BEFORE observing B. What we believe about A without any new evidence.",bg:LBL,bd:BL},
      {t:"P(B|A)  = Likelihood",
       b:"The probability of observing evidence B IF A is true. How likely is the evidence under our hypothesis?",bg:LGN,bd:GN},
      {t:"P(A|B)  = Posterior Probability",
       b:"The UPDATED probability of A AFTER observing evidence B. This is what Bayes' Theorem computes for us.",bg:LMR,bd:M},
      {t:"P(B)  = Marginal Probability",
       b:"Total probability of observing B regardless of A. Acts as a normalizing constant (denominator).",bg:LYL,bd:AM}
    ];
    terms.forEach((tm,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=4.50+row*1.06;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:bx,y:by,w:6.05,h:0.94, fill:{color:tm.bg}, line:{color:tm.bd,width:1},rectRadius:0.08});
      s.addText([
        {text:tm.t,options:{bold:true,color:tm.bd,breakLine:true}},
        {text:tm.b,options:{color:BK,fontSize:11}}
      ],{x:bx+0.10,y:by+0.06,w:5.85,h:0.80, fontSize:12.5,fontFace:"Arial",margin:0});
    });

    s.addNotes(
"Slide 3 – Bayes' Theorem\n\n"+
"Bayes' Theorem is one of the most powerful and important results in probability theory. It was developed by Reverend Thomas Bayes in 1763 and later refined by Pierre-Simon Laplace.\n\n"+
"The theorem gives us a way to UPDATE our probability estimate for a hypothesis A after observing new evidence B.\n\n"+
"The formula is: P(A|B) = P(B|A) × P(A) / P(B)\n\n"+
"This is derived directly from the definition of conditional probability. We know P(A|B) = P(A∩B)/P(B) and P(B|A) = P(A∩B)/P(A). Solving both equations for P(A∩B) and equating them gives us Bayes' Theorem.\n\n"+
"The four components:\n\n"+
"PRIOR P(A): Our initial belief about hypothesis A before seeing any evidence. It represents our background knowledge.\n\n"+
"LIKELIHOOD P(B|A): How well the hypothesis A explains the observed evidence B. High likelihood means A strongly predicts B.\n\n"+
"POSTERIOR P(A|B): The updated probability of A after observing B. This is what we want to compute.\n\n"+
"MARGINAL P(B): The overall probability of observing B under all possible scenarios. It normalizes the result so all probabilities sum to 1.\n\n"+
"The key insight of Bayes' Theorem: it tells us how to RATIONALLY UPDATE our beliefs when we receive new evidence. This is the foundation of Bayesian statistics and probabilistic machine learning.");
  }

  // Slide 4 ── Medical Test Example – Setup
  {
    const s = contentSlide(p, H, "Classic Example: Medical Diagnostic Test");
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:CYC,w:CW,h:0.88, fill:{color:LYL}, line:{color:AM,width:2},rectRadius:0.1});
    s.addText(
      "Scenario:  A rare disease D affects 1% of the population. A test T has 99% accuracy "+
      "(both sensitivity and specificity). If a person tests positive, what is the probability they actually have the disease?",
      {x:0.62,y:CYC+0.06,w:CW-0.24,h:0.76, fontSize:13.5,fontFace:"Arial",color:"4E342E",margin:0});

    // Four given values
    const given=[
      {t:"P(D) = 0.01",        d:"Prior probability of having the disease (1% of population)",bg:LMR,bd:M},
      {t:"P(¬D) = 0.99",       d:"Prior probability of NOT having the disease (99%)",bg:LGR,bd:GY},
      {t:"P(T⁺ | D) = 0.99",  d:"Sensitivity: Test correctly detects disease 99% of the time",bg:LGN,bd:GN},
      {t:"P(T⁺ | ¬D) = 0.01", d:"False Positive Rate: Test incorrectly says positive 1% of the time for healthy people",bg:LBL,bd:BL}
    ];
    given.forEach((g,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=2.98+row*1.22;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:bx,y:by,w:6.05,h:1.08, fill:{color:g.bg}, line:{color:g.bd,width:1.5},rectRadius:0.1});
      s.addText(g.t, {x:bx+0.12,y:by+0.08,w:5.82,h:0.42, fontSize:16,bold:true,color:g.bd,fontFace:"Arial",margin:0});
      s.addText(g.d, {x:bx+0.12,y:by+0.52,w:5.82,h:0.52, fontSize:11.5,color:BK,fontFace:"Arial",margin:0});
    });

    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:5.58,w:CW,h:0.88, fill:{color:LMR}, line:{color:M,width:2},rectRadius:0.1});
    s.addText(
      "Question: A patient tests positive (T⁺). What is P(D | T⁺)?  "+
      "Most people guess ≈ 99%  … the answer is SHOCKING.",
      {x:0.62,y:5.62,w:CW-0.24,h:0.78, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});

    s.addNotes(
"Slide 4 – Classic Example: Medical Diagnostic Test\n\n"+
"This is the most famous application of Bayes' Theorem and it produces a result that surprises almost everyone the first time they see it.\n\n"+
"The scenario: A rare disease D affects 1 in 100 people, so P(D) = 0.01. A diagnostic test T has been developed with 99% accuracy in BOTH directions. Specifically:\n\n"+
"Sensitivity P(T+|D) = 0.99: If you HAVE the disease, the test correctly comes back positive 99% of the time. Only 1% of sick people get a false negative.\n\n"+
"Specificity P(T-|¬D) = 0.99: If you do NOT have the disease, the test correctly comes back negative 99% of the time. Only 1% of healthy people get a false positive. So the false positive rate P(T+|¬D) = 0.01.\n\n"+
"Now the question: A randomly selected person takes the test and gets a POSITIVE result. What is the probability they actually have the disease?\n\n"+
"Most people immediately say 'about 99%' — after all, the test is 99% accurate! But this reasoning ignores the prior probability. The disease only affects 1% of people. Even with a positive test, the answer might be very different from 99%.\n\n"+
"This is the key insight of Bayes' Theorem: we must combine the evidence (the test result) with our prior knowledge (the base rate of the disease). Let us calculate on the next slide.");
  }

  // Slide 5 ── Solving the Example
  {
    const s = contentSlide(p, H, "Solving with Bayes' Theorem");
    s.addText("Apply the expanded form of Bayes' Theorem:",
      {x:CX,y:CYC,w:CW,h:0.38, fontSize:14,fontFace:"Arial",color:BK});

    const calcSteps=[
      {step:"Formula:",  val:"P(D|T⁺) = P(T⁺|D) × P(D)  /  [P(T⁺|D)×P(D)  +  P(T⁺|¬D)×P(¬D)]",bg:LBL},
      {step:"Substitute values:",val:"P(D|T⁺) = (0.99 × 0.01)  /  [(0.99 × 0.01) + (0.01 × 0.99)]",bg:LGR},
      {step:"Numerator:",  val:"0.99 × 0.01  =  0.0099",bg:LGN},
      {step:"Denominator:",val:"(0.99×0.01) + (0.01×0.99)  =  0.0099 + 0.0099  =  0.0198",bg:LYL},
      {step:"Final Answer:", val:"P(D|T⁺) = 0.0099 / 0.0198  =  0.50  =  50%  !!",bg:LMR}
    ];
    calcSteps.forEach((st,i)=>{
      const y=2.18+i*0.82;
      s.addShape(p.shapes.ROUNDED_RECTANGLE,
        {x:CX,y,w:CW,h:0.75, fill:{color:st.bg}, line:{color:"CCCCCC",width:1},rectRadius:0.08});
      s.addText([
        {text:st.step+"  ",options:{bold:true,color:M}},
        {text:st.val,options:{color:BK}}
      ],{x:0.62,y:y+0.09,w:CW-0.24,h:0.58, fontSize:i===4?15:13.5,fontFace:"Arial",margin:0});
    });

    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:6.30,w:CW,h:1.12, fill:{color:LYL}, line:{color:AM,width:2},rectRadius:0.1});
    s.addText([
      {text:"Why only 50%?  ",options:{bold:true,color:AM}},
      {text:"Of 10,000 people: 100 have disease (99 test positive) + 9,900 healthy (99 test false-positive). "+
            "So of 198 positive tests, only 99 truly have the disease = 50%!\n",options:{}},
      {text:"Key Lesson: Low prior probability dramatically reduces the posterior even with accurate tests.",options:{bold:true,color:M}}
    ],{x:0.62,y:6.34,w:CW-0.24,h:1.04, fontSize:12,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 5 – Solving the Medical Test Example with Bayes' Theorem\n\n"+
"Let us carefully apply Bayes' Theorem step by step.\n\n"+
"Step 1 — Write the formula: P(D|T+) = P(T+|D)×P(D) / [P(T+|D)×P(D) + P(T+|¬D)×P(¬D)]\n\n"+
"Step 2 — Substitute known values: Numerator = 0.99 × 0.01 = 0.0099. Denominator = (0.99×0.01) + (0.01×0.99) = 0.0099 + 0.0099 = 0.0198.\n\n"+
"Step 3 — Divide: P(D|T+) = 0.0099 / 0.0198 = 0.50 = 50%.\n\n"+
"The answer is only 50%! Only half the people who test positive actually have the disease. This is dramatically less than the 99% most people expected.\n\n"+
"Why? Let us think in terms of natural frequencies to make it intuitive. Imagine testing 10,000 people. Of these, 100 have the disease (1%) and 9,900 do not. Of the 100 sick people, 99 test positive (99% sensitivity). Of the 9,900 healthy people, 99 also test positive (1% false positive rate). So we have 99 + 99 = 198 positive tests total, but only 99 of those 198 people are actually sick. 99/198 = exactly 50%.\n\n"+
"The key lesson: the base rate (prior probability) matters enormously. The disease affects only 1% of people, which is so rare that even a 99% accurate test produces an equal number of true positives and false positives. This is called 'base rate fallacy' when people ignore the prior.\n\n"+
"This result has real consequences in medicine, law (where test evidence is presented to juries), and machine learning.");
  }

  // Slide 6 ── Applications
  {
    const s = contentSlide(p, H, "Applications in Computer Science");
    s.addText("Bayes' Theorem is the foundation of probabilistic reasoning in AI and data science:",
      {x:CX,y:CYC,w:CW,h:0.40, fontSize:14,fontFace:"Arial",color:BK});
    const apps=[
      {t:"📧  Naive Bayes Spam Filter",
       b:"Given an email, compute P(spam | words in email). Train on labeled emails to learn P(word|spam) and P(word|not-spam). The 'Naive' assumption: words are independent given the class. Very fast and still competitive.",bg:LGN,bd:GN},
      {t:"🏥  Medical Diagnosis AI",
       b:"AI systems compute P(disease | symptoms, test results). Updated with new evidence as more tests are run. IBM Watson Health uses Bayesian reasoning to recommend treatments based on patient history and literature.",bg:LMR,bd:M},
      {t:"🔍  Search & Recommendation",
       b:"Google's ranking algorithms incorporate Bayesian principles. Collaborative filtering in recommendation systems (Netflix, Amazon) estimates P(user likes item | ratings of similar users) using Bayesian updates.",bg:LBL,bd:BL},
      {t:"🛡️  Intrusion Detection",
       b:"Network security systems compute P(attack | observed traffic patterns). Start with P(attack) as the base rate, update with each suspicious packet. Bayesian Networks model complex multi-step attack scenarios.",bg:LYL,bd:AM}
    ];
    apps.forEach((app,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=2.20+row*2.30;
      box(p,s,bx,by,6.05,2.14,app.bg,app.bd);
      boxTitle(s,bx,by,6.05,app.t,app.bd);
      boxBody(s,bx,by,6.05,2.14,app.b);
    });
    s.addNotes(
"Slide 6 – Applications of Bayes' Theorem in Computer Science\n\n"+
"Bayes' Theorem is the mathematical engine behind many of the most important technologies in artificial intelligence, data science, and computer security.\n\n"+
"Naive Bayes Spam Filter: This is the classic text classification application. To classify an email as spam or not-spam, we compute P(spam|words) ∝ P(spam) × product of P(word_i|spam) for each word. We learn the likelihoods from a training dataset of labeled emails. Despite the 'naive' independence assumption between words, Naive Bayes classifiers are remarkably effective and were state-of-the-art spam filters in the early 2000s.\n\n"+
"Medical Diagnosis AI: Clinical decision support systems use Bayesian networks to reason about diagnoses. As a patient undergoes more tests, the posterior probability of each diagnosis is continuously updated using Bayes' Theorem. This models exactly how a good doctor thinks — updating beliefs with each new piece of evidence.\n\n"+
"Search and Recommendation Systems: Google and other search engines incorporate Bayesian thinking. Collaborative filtering — the technology behind 'users who liked X also liked Y' recommendations on Netflix and Amazon — is based on estimating conditional probabilities from large datasets of user behavior.\n\n"+
"Intrusion Detection: Security systems monitor network traffic and compute the probability of an ongoing attack given the observed patterns. Bayesian Networks allow modeling of complex, multi-step attack scenarios and adapting to new attack patterns by updating the prior.");
  }

  thankYouSlide(p, H);
  await p.writeFile({fileName:"./outputs/PPT4_Unit4_Bayes_Theorem.pptx"});
  console.log("✔  PPT4 written");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 5 – GRAPH COLORING  (Unit V)
// ═══════════════════════════════════════════════════════════════
async function ppt5() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Mathematical Foundations of Computer Science";

  titleSlide(p, H, "Unit V: Graphs and Trees", "Graph Coloring");

  // Slide 2 ── Introduction
  {
    const s = contentSlide(p, H, "What is Graph Coloring?");
    s.addText([
      {text:"Graph Coloring ",options:{bold:true}},
      {text:"is the assignment of labels ('colors') to the vertices of a graph such that "+
            "no two adjacent vertices (connected by an edge) receive the same color.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.55, fontSize:13.5,fontFace:"Arial",color:BK});

    box(p,s,CX,2.35,6.05,4.22,LBL,BL);
    s.addText("Formal Definition",
      {x:0.62,y:2.43,w:5.82,h:0.38, fontSize:14,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Given graph G = (V, E):",options:{bold:true,breakLine:true}},
      {text:"  V = set of vertices",options:{breakLine:true}},
      {text:"  E = set of edges",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"A proper k-coloring is a function:",options:{bold:true,breakLine:true}},
      {text:"  c : V → {1, 2, …, k}",options:{bold:true,color:BL,breakLine:true}},
      {text:"  such that c(u) ≠ c(v)  whenever  {u,v} ∈ E",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"A graph is k-colorable if a proper",options:{breakLine:true}},
      {text:"k-coloring exists for it.",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Note: Colors are just labels — the",options:{breakLine:true}},
      {text:"actual colors used don't matter,",options:{breakLine:true}},
      {text:"only that adjacent vertices differ.",options:{}}
    ],{x:0.62,y:2.85,w:5.82,h:3.65, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    box(p,s,6.60,2.35,6.25,4.22,LYL,AM);
    s.addText("Illustrative Example: K₃ (Triangle)",
      {x:6.72,y:2.43,w:6.03,h:0.38, fontSize:13.5,bold:true,color:AM,fontFace:"Arial",margin:0});

    // Draw triangle K3
    const cx3=9.72, cy3=4.25, r=1.10;
    const verts=[
      {x:cx3, y:cy3-r, label:"A", col:"FF4444"},
      {x:cx3-r*0.87, y:cy3+r*0.5, label:"B", col:"4477FF"},
      {x:cx3+r*0.87, y:cy3+r*0.5, label:"C", col:"22AA44"}
    ];
    // Edges as lines
    for(let i=0;i<3;i++){
      const a=verts[i], b=verts[(i+1)%3];
      s.addShape(p.shapes.LINE,
        {x:a.x,y:a.y,w:b.x-a.x,h:b.y-a.y, line:{color:"333333",width:2.5}});
    }
    // Vertex circles + labels
    verts.forEach(v=>{
      s.addShape(p.shapes.OVAL,
        {x:v.x-0.30,y:v.y-0.30,w:0.60,h:0.60, fill:{color:v.col}, line:{color:"333333",width:1.5}});
      s.addText(v.label,
        {x:v.x-0.28,y:v.y-0.28,w:0.56,h:0.56, fontSize:16,bold:true,color:WH,
         align:"center",valign:"middle",fontFace:"Arial"});
    });

    s.addText([
      {text:"K₃: 3 vertices, all connected",options:{breakLine:true}},
      {text:"  • A–B edge: A≠B  ✔",options:{breakLine:true}},
      {text:"  • B–C edge: B≠C  ✔",options:{breakLine:true}},
      {text:"  • A–C edge: A≠C  ✔",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Minimum colors needed = 3",options:{bold:true,color:M,breakLine:true}},
      {text:"(Cannot use only 2 colors!)",options:{italic:true}}
    ],{x:6.72,y:5.52,w:6.03,h:1.00, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 2 – What is Graph Coloring?\n\n"+
"Graph coloring is one of the most studied and most applicable problems in graph theory. Given a graph G with vertices V and edges E, we want to assign colors (or labels) to the vertices such that no two vertices connected by an edge have the same color. Such an assignment is called a 'proper coloring'.\n\n"+
"Formally, a proper k-coloring is a function c from the vertex set V to a set of k colors, such that for every edge {u,v} in E, the colors of u and v are different: c(u) ≠ c(v).\n\n"+
"The word 'colors' is just a metaphor — in practice we might use integers {1,2,3,...,k} or any k distinct labels. What matters is only that adjacent vertices get different labels.\n\n"+
"Consider the triangle K₃ — the complete graph on 3 vertices, where every pair of vertices is connected. We need to color vertices A, B, and C. Since A is connected to B, A and B must get different colors. Since B is connected to C, B and C must get different colors. Since A is connected to C, A and C must get different colors. Therefore all three must have different colors — we cannot 2-color a triangle. The minimum is 3 colors.\n\n"+
"This minimum number of colors needed to properly color a graph is called the chromatic number, which we will study on the next slide.");
  }

  // Slide 3 ── Chromatic Number
  {
    const s = contentSlide(p, H, "Chromatic Number  χ(G)");
    s.addText([
      {text:"The chromatic number χ(G) ",options:{bold:true}},
      {text:"of a graph G is the MINIMUM number of colors required for a proper vertex coloring.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.50, fontSize:13.5,fontFace:"Arial",color:BK});

    // Table of graph types
    const tData=[
      [
        {text:"Graph Type",options:{bold:true,fill:{color:M},color:WH}},
        {text:"Example",options:{bold:true,fill:{color:M},color:WH,align:"center"}},
        {text:"χ(G)",options:{bold:true,fill:{color:M},color:WH,align:"center"}},
        {text:"Reason",options:{bold:true,fill:{color:M},color:WH}}
      ],
      [
        {text:"Complete Graph Kₙ",options:{fill:{color:LMR},bold:true}},
        {text:"K₄ (4 mutual edges)",options:{fill:{color:LMR},align:"center"}},
        {text:"n",options:{fill:{color:LMR},align:"center",bold:true,color:M}},
        {text:"Every vertex connects to all others; all need distinct colors",options:{fill:{color:LMR}}}
      ],
      [
        {text:"Bipartite Graph",options:{fill:{color:LGN},bold:true}},
        {text:"Even cycle C₄",options:{fill:{color:LGN},align:"center"}},
        {text:"2",options:{fill:{color:LGN},align:"center",bold:true,color:DGN}},
        {text:"Vertices split into 2 independent sets; 2 colors suffice",options:{fill:{color:LGN}}}
      ],
      [
        {text:"Odd Cycle C₂ₖ₊₁",options:{fill:{color:LBL},bold:true}},
        {text:"Triangle C₃",options:{fill:{color:LBL},align:"center"}},
        {text:"3",options:{fill:{color:LBL},align:"center",bold:true,color:BL}},
        {text:"Odd-length cycle cannot be 2-colored (parity argument)",options:{fill:{color:LBL}}}
      ],
      [
        {text:"Tree or Forest",options:{fill:{color:LYL},bold:true}},
        {text:"Any acyclic graph",options:{fill:{color:LYL},align:"center"}},
        {text:"2",options:{fill:{color:LYL},align:"center",bold:true,color:AM}},
        {text:"Trees are bipartite; BFS/DFS gives 2-coloring",options:{fill:{color:LYL}}}
      ],
      [
        {text:"Planar Graph",options:{fill:{color:LGR},bold:true}},
        {text:"Any map graph",options:{fill:{color:LGR},align:"center"}},
        {text:"≤ 4",options:{fill:{color:LGR},align:"center",bold:true,color:GY}},
        {text:"Four Color Theorem: every planar graph is 4-colorable",options:{fill:{color:LGR}}}
      ]
    ];
    s.addTable(tData,
      {x:CX,y:2.28,w:CW,h:4.94, fontSize:12.5,fontFace:"Arial",
       border:{pt:1,color:"CCCCCC"},colW:[2.8,2.5,1.3,5.75]});
    s.addNotes(
"Slide 3 – Chromatic Number\n\n"+
"The chromatic number χ(G) — read as 'chi of G' — is the minimum number of colors needed to properly color the graph G. Finding the chromatic number is generally one of the hardest problems in computer science — it is NP-complete for arbitrary graphs.\n\n"+
"Let us survey the chromatic numbers of common graph families:\n\n"+
"COMPLETE GRAPH Kₙ: In Kₙ, every vertex is connected to every other vertex. So every pair of vertices is adjacent, and all n vertices must receive different colors. Therefore χ(Kₙ) = n.\n\n"+
"BIPARTITE GRAPH: A bipartite graph has its vertices split into two sets, say X and Y, with edges only going between X and Y (no edges within X or within Y). We can color all of X with Color 1 and all of Y with Color 2. Therefore χ(bipartite) = 2 (assuming at least one edge exists). All trees and even cycles are bipartite.\n\n"+
"ODD CYCLE C₂ₖ₊₁: An odd cycle with 2k+1 vertices cannot be 2-colored. If we try to alternate two colors around the cycle, we eventually reach a conflict when the last vertex is adjacent to the first. We need a third color. So χ(odd cycle) = 3. But EVEN cycles are bipartite, so χ(even cycle) = 2.\n\n"+
"TREES: Trees are acyclic graphs and are always bipartite. A simple BFS from any root, coloring vertices by their level parity (even levels = color 1, odd levels = color 2) gives a valid 2-coloring. So χ(tree) = 2 for any tree with at least one edge.\n\n"+
"PLANAR GRAPHS: The famous Four Color Theorem states that any planar graph can be colored with at most 4 colors.");
  }

  // Slide 4 ── Greedy Coloring
  {
    const s = contentSlide(p, H, "Greedy Coloring Algorithm");
    s.addText([
      {text:"Greedy Coloring:",options:{bold:true}},
      {text:" Process vertices one by one and assign the SMALLEST available color "+
            "not used by any already-colored neighbor.",options:{}}
    ],{x:CX,y:CYC,w:CW,h:0.50, fontSize:13.5,fontFace:"Arial",color:BK});

    // Algorithm steps
    box(p,s,CX,2.30,5.90,4.28,LGR,"AAAAAA");
    s.addText("Algorithm Steps",
      {x:0.62,y:2.38,w:5.70,h:0.38, fontSize:14,bold:true,color:M,fontFace:"Arial",margin:0});
    s.addText([
      {text:"For each vertex v in order v₁, v₂, …, vₙ:",options:{bold:true,breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"   1. Find all colors used by v's neighbors",options:{breakLine:true}},
      {text:"      that have already been colored.",options:{breakLine:true}},
      {text:"   2. Assign to v the smallest color",options:{breakLine:true}},
      {text:"      NOT in that set.",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Time Complexity:",options:{bold:true,breakLine:true}},
      {text:"   O(V + E)  — linear in graph size",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Limitation:",options:{bold:true,breakLine:true}},
      {text:"   NOT always optimal — result depends",options:{breakLine:true}},
      {text:"   on vertex ordering. May use MORE",options:{breakLine:true}},
      {text:"   colors than the chromatic number.",options:{}}
    ],{x:0.62,y:2.80,w:5.70,h:3.70, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    // Trace example
    box(p,s,6.60,2.30,6.25,4.28,LYL,AM);
    s.addText("Trace: Graph with 5 vertices",
      {x:6.72,y:2.38,w:6.03,h:0.38, fontSize:13.5,bold:true,color:AM,fontFace:"Arial",margin:0});
    const trace=[
      {v:"v₁", nbrs:"—",          col:"Red  (1)",    why:"No neighbors colored yet"},
      {v:"v₂", nbrs:"v₁ (Red)",   col:"Blue  (2)",   why:"Red is taken by v₁"},
      {v:"v₃", nbrs:"v₁,v₂",      col:"Green  (3)",  why:"Red+Blue taken"},
      {v:"v₄", nbrs:"v₂,v₃",      col:"Red  (1)",    why:"Blue+Green taken; Red free!"},
      {v:"v₅", nbrs:"v₁,v₄",      col:"Blue  (2)",   why:"Red taken by v₁,v₄"},
    ];
    const trTbl = trace.map((t,i)=>[
      {text:t.v, options:{bold:true,fill:{color:i%2===0?LGR:"EFEFEF"},align:"center"}},
      {text:t.nbrs, options:{fill:{color:i%2===0?LGR:"EFEFEF"},fontSize:11}},
      {text:t.col, options:{bold:true,fill:{color:i%2===0?LGR:"EFEFEF"},color:M}},
      {text:t.why, options:{fill:{color:i%2===0?LGR:"EFEFEF"},fontSize:11}}
    ]);
    trTbl.unshift([
      {text:"Vertex",options:{bold:true,fill:{color:AM},color:WH,align:"center"}},
      {text:"Colored Neighbors",options:{bold:true,fill:{color:AM},color:WH}},
      {text:"Assigned Color",options:{bold:true,fill:{color:AM},color:WH}},
      {text:"Reason",options:{bold:true,fill:{color:AM},color:WH}}
    ]);
    s.addTable(trTbl,
      {x:6.72,y:2.82,w:6.03,h:3.68, fontSize:12,fontFace:"Arial",
       border:{pt:1,color:"CCCCCC"},colW:[0.75,1.85,1.62,1.81]});

    s.addNotes(
"Slide 4 – Greedy Coloring Algorithm\n\n"+
"The Greedy Coloring algorithm is the simplest and most efficient approach to graph coloring, even though it is not guaranteed to find the minimum number of colors.\n\n"+
"The algorithm works as follows: We process the vertices in some fixed order v₁, v₂, ..., vₙ. For each vertex, we look at its already-colored neighbors, collect the set of colors they use, and then assign to the current vertex the smallest color number NOT in that forbidden set.\n\n"+
"Let us trace through the example in the table. We have 5 vertices with edges forming a cycle-like structure.\n\n"+
"v₁: No neighbors are colored yet, so we assign Color 1 (Red).\n"+
"v₂: Neighbor v₁ has Red. The smallest available color is Blue (2).\n"+
"v₃: Neighbors v₁ (Red) and v₂ (Blue) are both colored. We need Green (3).\n"+
"v₄: Neighbors v₂ (Blue) and v₃ (Green) are colored. Red (1) is free — assign Red!\n"+
"v₅: Neighbors v₁ (Red) and v₄ (Red) are both Red. Blue (2) is free — assign Blue.\n\n"+
"Result: We used 3 colors. For this particular graph, this is optimal.\n\n"+
"However, the Greedy algorithm's result heavily depends on the vertex ordering. A different ordering might give a different (possibly worse) coloring. For example, adversarially ordered bipartite graphs can be made to require n/2 colors by Greedy even though 2 colors suffice. More sophisticated algorithms like Welsh-Powell use degree-based ordering to improve results.\n\n"+
"The Greedy algorithm runs in O(V + E) time, making it practical for large graphs even when optimal coloring is computationally infeasible.");
  }

  // Slide 5 ── Four Color Theorem
  {
    const s = contentSlide(p, H, "The Four Color Theorem");
    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:CYC,w:CW,h:1.05, fill:{color:LMR}, line:{color:M,width:2},rectRadius:0.1});
    s.addText([
      {text:"Four Color Theorem: ",options:{bold:true,fontSize:18,color:M}},
      {text:"Every planar graph can be properly colored using at most ",options:{fontSize:15}},
      {text:"4 colors",options:{bold:true,fontSize:18,color:M}},
      {text:".",options:{fontSize:15}}
    ],{x:0.62,y:CYC+0.08,w:CW-0.24,h:0.88, fontSize:15,fontFace:"Arial",color:BK,valign:"middle",margin:0});

    box(p,s,CX,2.90,5.90,2.48,LBL,BL);
    s.addText("History of the Proof",
      {x:0.62,y:2.98,w:5.70,h:0.38, fontSize:14,bold:true,color:BL,fontFace:"Arial",margin:0});
    s.addText([
      {text:"1852:",options:{bold:true,color:BL}},
      {text:" Francis Guthrie first conjectured the theorem while coloring a map of England.",options:{breakLine:true}},
      {text:"1879:",options:{bold:true,color:BL}},
      {text:" Kempe published a 'proof' — found to be flawed by Heawood in 1890.",options:{breakLine:true}},
      {text:"1976:",options:{bold:true,color:BL}},
      {text:" Kenneth Appel and Wolfgang Haken (Univ. of Illinois) produced the first correct proof — with a computer checking 1,936 special cases.",options:{breakLine:true}},
      {text:"Significance:",options:{bold:true,color:M}},
      {text:" First major theorem PROVEN using a computer. 124 years from conjecture to proof!",options:{}}
    ],{x:0.62,y:3.40,w:5.70,h:1.90, fontSize:12,fontFace:"Arial",color:BK,margin:0});

    box(p,s,6.60,2.90,6.25,2.48,LGN,GN);
    s.addText("Map Coloring Interpretation",
      {x:6.72,y:2.98,w:6.03,h:0.38, fontSize:14,bold:true,color:GN,fontFace:"Arial",margin:0});
    s.addText([
      {text:"Any geographic map can be colored with at most 4 colors so that no two adjacent regions (sharing a border) have the same color.",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Key: A map corresponds to a PLANAR graph —",options:{bold:true,breakLine:true}},
      {text:"  • Vertices = regions/countries",options:{breakLine:true}},
      {text:"  • Edges = shared borders",options:{breakLine:true}},
      {text:"  • Planar = can be drawn without edge crossings",options:{breakLine:true}},
      {text:" ",options:{fontSize:5,breakLine:true}},
      {text:"Note: 5 Colors were known to suffice (1890). The hard part was proving 4 is enough!",options:{italic:true}}
    ],{x:6.72,y:3.40,w:6.03,h:1.90, fontSize:12,fontFace:"Arial",color:BK,margin:0});

    s.addShape(p.shapes.ROUNDED_RECTANGLE,
      {x:CX,y:5.48,w:CW,h:1.72, fill:{color:LYL}, line:{color:AM,width:1.5},rectRadius:0.1});
    s.addText("5-Color Theorem (Easy Proof via Induction):",
      {x:0.62,y:5.56,w:CW-0.24,h:0.38, fontSize:14,bold:true,color:AM,fontFace:"Arial",margin:0});
    s.addText(
      "Every planar graph has a vertex of degree ≤ 5 (by Euler's formula). Remove this vertex, "+
      "5-color the rest by induction (IH), then reinsert the vertex — with at most 5 neighbors "+
      "using ≤ 5 colors, at least one color remains available for the removed vertex.",
      {x:0.62,y:5.96,w:CW-0.24,h:1.18, fontSize:12.5,fontFace:"Arial",color:BK,margin:0});

    s.addNotes(
"Slide 5 – The Four Color Theorem\n\n"+
"The Four Color Theorem states that any planar graph can be properly vertex-colored using at most four colors. Equivalently, any geographic map can be colored with just four colors so that no two adjacent countries share the same color.\n\n"+
"The history of this theorem is one of the most fascinating in mathematics. In 1852, Francis Guthrie noticed while coloring a map of English counties that four colors seemed to always suffice. He asked his mathematician brother if this was always true.\n\n"+
"For over a century, the theorem resisted proof. In 1879, Alfred Kempe published what he claimed was a proof, but it was found to contain a subtle error 11 years later by Percy Heawood. Heawood did manage to prove the weaker 5-Color Theorem.\n\n"+
"Finally, in 1976, Kenneth Appel and Wolfgang Haken at the University of Illinois achieved the first correct proof — but it was unlike any proof before it. They reduced the problem to checking 1,936 special configurations, each of which was verified by a computer program. This was the first major theorem in mathematics proved with computer assistance, which sparked significant philosophical debate about what constitutes a valid mathematical proof.\n\n"+
"The 5-Color Theorem has an elegant inductive proof: Every planar graph has a vertex of degree at most 5 (provable using Euler's formula V - E + F = 2). Remove this vertex, 5-color the rest by induction, then re-insert the vertex. With at most 5 neighbors already using at most 5 colors, at least one color is available.\n\n"+
"The jump from 5 to 4 is what required the massive computer-assisted case analysis.");
  }

  // Slide 6 ── Applications
  {
    const s = contentSlide(p, H, "Applications in Computer Science");
    s.addText("Graph coloring models many real-world conflict-avoidance and resource-allocation problems:",
      {x:CX,y:CYC,w:CW,h:0.40, fontSize:14,fontFace:"Arial",color:BK});
    const apps=[
      {t:"🖥️  Register Allocation in Compilers",
       b:"Variables in a program form an 'interference graph' — two variables are adjacent if their lifetimes overlap (they cannot share a register). Coloring this graph assigns variables to CPU registers. Vertices needing more colors than available registers are 'spilled' to memory.",bg:LGN,bd:GN},
      {t:"📡  Frequency Assignment in Networks",
       b:"In mobile networks, nearby cell towers must operate on different frequencies to avoid interference. Model towers as vertices and proximity as edges. Graph coloring assigns frequencies (colors) while minimizing the total frequency spectrum used — saving expensive bandwidth.",bg:LBL,bd:BL},
      {t:"📅  Scheduling with Conflicts",
       b:"Exam scheduling: students (or courses) with common students are 'conflicting' — they cannot be in the same time slot. Build a conflict graph and color it. Each color represents a time slot. Minimum colors = minimum number of time slots needed.",bg:LMR,bd:M},
      {t:"🔢  Sudoku as Graph Coloring",
       b:"A 9×9 Sudoku grid is a graph with 81 vertices (cells). Two vertices are adjacent if they share a row, column, or 3×3 box. We need to color this graph with 9 colors (digits 1-9) such that each row, column, and box uses all 9 colors exactly once.",bg:LYL,bd:AM}
    ];
    apps.forEach((app,i)=>{
      const col=i%2, row=Math.floor(i/2);
      const bx=CX+col*6.22, by=2.20+row*2.30;
      box(p,s,bx,by,6.05,2.14,app.bg,app.bd);
      boxTitle(s,bx,by,6.05,app.t,app.bd);
      boxBody(s,bx,by,6.05,2.14,app.b);
    });
    s.addNotes(
"Slide 6 – Applications of Graph Coloring in Computer Science\n\n"+
"Graph coloring is not just a theoretical concept — it models real-world resource allocation and conflict resolution problems across computer science.\n\n"+
"Register Allocation in Compilers: This is one of the most important applications. When a compiler generates machine code, it must decide which CPU registers to assign to which variables. Two variables that are 'live' at the same time cannot share a register. We build an interference graph where vertices are variables and edges connect variables that are simultaneously live. Coloring this graph assigns registers, with each color representing one register. If the graph requires more colors than available registers, some variables are 'spilled' to slower memory. This is why register allocation is studied in every compiler design course.\n\n"+
"Frequency Assignment in Mobile Networks: Cell towers that are geographically close to each other use different radio frequencies to avoid interference. This is modeled as graph coloring where towers are vertices and proximity defines edges. The goal is to minimize the number of distinct frequencies used — which directly translates to minimizing the expensive electromagnetic spectrum licensed from governments.\n\n"+
"Exam Scheduling: A classic scheduling application. We build a conflict graph where courses are vertices and two courses are adjacent if they have at least one common student (who obviously cannot be in two places at once). Coloring this graph gives a schedule where each color (time slot) groups courses that can run simultaneously without conflicts. The chromatic number gives the minimum number of time slots needed.\n\n"+
"Sudoku Solving: A 9×9 Sudoku is equivalent to a graph coloring problem. Each of the 81 cells is a vertex, and two cells are connected if they are in the same row, column, or 3×3 box. Filling in the Sudoku is equivalent to finding a valid 9-coloring of this graph respecting the pre-filled cells as fixed color assignments.");
  }

  thankYouSlide(p, H);
  await p.writeFile({fileName:"./outputs/PPT5_Unit5_Graph_Coloring.pptx"});
  console.log("✔  PPT5 written");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  // Create outputs directory if it doesn't exist
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
    console.log(`Created directory: ${outputDir}`);
  }

  console.log("Generating 5 MFCS PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  await ppt4();
  await ppt5();
  console.log("\nAll 5 PPTs generated successfully!");
})();
