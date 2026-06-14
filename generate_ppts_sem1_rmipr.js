const pptxgen = require("pptxgenjs");
const fs = require("fs");

const LOGO = "image/png;base64," + fs.readFileSync("image1.png").toString("base64");

// ── Color palette ─────────────────────────────────────────────
const M = "990033"; // maroon
const GN = "009900"; // green
const GL = "F9DD67"; // gold
const WH = "FFFFFF"; // white
const BK = "000000"; // black
const GY = "555555"; // gray
const LGR = "F8F8F8"; // light gray
const LGN = "E8F5E9"; // light green
const LMR = "FCE4EC"; // light maroon/pink
const BL = "1565C0"; // blue
const LBL = "E3F2FD"; // light blue
const AM = "F57F17"; // amber
const LYL = "FFF9C4"; // light yellow
const DGN = "006400"; // dark green

// ── Slide dimensions (LAYOUT_WIDE = 13.33 × 7.5 in) ─────────
const SW = 13.33, SH = 7.5;
const BRD = 0.26;     // maroon border width
const GLD = 0.09;     // gold accent width

// ── Header / logo / content coordinates ──────────────────────
const HX = 0.35, HY = 0.27, HW = 11.12, HH = 0.79;
const LX = 11.55, LY = 0.25, LW = 1.03, LH = 1.03;
const CX = 0.50, CW = 12.35;
const CYT = 1.12;   // y of slide subtitle title
const CYC = 1.72;   // y where main content starts

// ─────────────────────────────────────────────────────────────
//  SHARED LAYOUT HELPERS
// ─────────────────────────────────────────────────────────────
function frame(pres, sl, hdrText) {
  sl.background = { color: WH };
  [[0, 0, SW, BRD], [0, SH - BRD, SW, BRD], [0, 0, BRD, SH], [SW - BRD, 0, BRD, SH]]
    .forEach(([x, y, w, h]) => sl.addShape(pres.shapes.RECTANGLE,
      { x, y, w, h, fill: { color: M }, line: { color: M, width: 1 } }));
  const gx = BRD, gy = BRD, gw = SW - 2 * BRD, gh = SH - 2 * BRD;
  [[gx, gy, gw, GLD], [gx, gy + gh - GLD, gw, GLD], [gx, gy, GLD, gh], [gx + gw - GLD, gy, GLD, gh]]
    .forEach(([x, y, w, h]) => sl.addShape(pres.shapes.RECTANGLE,
      { x, y, w, h, fill: { color: GL }, line: { color: GL, width: 1 } }));
  sl.addShape(pres.shapes.RECTANGLE,
    { x: HX, y: HY, w: HW, h: HH, fill: { color: GN }, line: { color: GN, width: 1 } });
  sl.addText(hdrText,
    {
      x: HX, y: HY, w: HW, h: HH, fontSize: 15, bold: true, color: WH,
      align: "center", valign: "middle", fontFace: "Arial", margin: 0
    });
  sl.addImage({ data: LOGO, x: LX, y: LY, w: LW, h: LH });
}

function titleSlide(pres, hdr, unit, topic) {
  const sl = pres.addSlide();
  frame(pres, sl, hdr);
  sl.addText(unit,
    { x: CX, y: 1.18, w: CW, h: 0.42, fontSize: 14, italic: true, color: M, align: "center", fontFace: "Arial" });
  sl.addShape(pres.shapes.RECTANGLE,
    { x: 1.80, y: 1.72, w: 9.73, h: 1.40, fill: { color: M }, line: { color: M, width: 1 } });
  sl.addText(topic,
    {
      x: 1.80, y: 1.72, w: 9.73, h: 1.40, fontSize: 28, bold: true, color: WH,
      align: "center", valign: "middle", fontFace: "Arial", margin: 0
    });
  sl.addText("PPT PRESENTED BY",
    { x: CX, y: 3.28, w: CW, h: 0.40, fontSize: 15, bold: true, color: M, align: "center", fontFace: "Arial" });
  ["NAME :   Srinivas Rao Tammireddy",
    "ROLL NO :   257Y1D5805",
    "YEAR :   I Year  I Semester",
    "ACADEMIC YEAR :   2025 - 2026"
  ].forEach((t, i) =>
    sl.addText(t, { x: 3.20, y: 3.76 + i * 0.44, w: 7.0, h: 0.42, fontSize: 13, color: BK, fontFace: "Arial" }));
  sl.addText("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING",
    { x: CX, y: 6.65, w: CW, h: 0.45, fontSize: 14, bold: true, color: M, align: "center", fontFace: "Arial" });
  return sl;
}

function contentSlide(pres, hdr, title) {
  const sl = pres.addSlide();
  frame(pres, sl, hdr);
  sl.addText(title,
    { x: CX, y: CYT, w: CW, h: 0.52, fontSize: 22, bold: true, color: M, fontFace: "Arial" });
  sl.addShape(pres.shapes.RECTANGLE,
    { x: CX, y: CYT + 0.52, w: CW, h: 0.04, fill: { color: M }, line: { color: M, width: 1 } });
  return sl;
}

function thankYouSlide(pres, hdr) {
  const sl = pres.addSlide();
  frame(pres, sl, hdr);
  sl.addText("THANK YOU",
    {
      x: CX, y: 2.6, w: CW, h: 1.5, fontSize: 48, bold: true, color: M,
      align: "center", valign: "middle", fontFace: "Arial"
    });
  sl.addText(
    "Department of Computer Science and Engineering\n" +
    "Marri Laxman Reddy Institute of Technology and Management",
    { x: CX, y: 5.5, w: CW, h: 0.9, fontSize: 13, color: GY, align: "center", fontFace: "Arial" });
  return sl;
}

function box(pres, sl, x, y, w, h, bg, bd) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE,
    { x, y, w, h, fill: { color: bg }, line: { color: bd, width: 1.5 }, rectRadius: 0.1 });
}

function boxTitle(sl, x, y, w, txt, clr) {
  sl.addText(txt,
    { x: x + 0.12, y: y + 0.10, w: w - 0.24, h: 0.38, fontSize: 14.5, bold: true, color: clr, fontFace: "Arial", margin: 0 });
}

function boxBody(sl, x, y, w, h, txt) {
  sl.addText(txt,
    { x: x + 0.12, y: y + 0.52, w: w - 0.24, h: h - 0.58, fontSize: 13, color: BK, fontFace: "Arial", margin: 0 });
}

// ═══════════════════════════════════════════════════════════════
//  PPT 1 – IDENTIFYING A RESEARCH PROBLEM (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Research Methodology & IPR";

  const tSl = titleSlide(p, H, "Unit I: Research Problem Formulation", "Identifying a Research Problem");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Identifying a Research Problem' as part of the Research Methodology and IPR course for my M.Tech in Computer Science and Engineering. In this presentation, we will define what a research problem is, identify its key characteristics and sources, discuss common errors researchers make in selecting a problem, outline the systematic investigation lifecycle, and highlight practical applications in academic and industrial domains.");

  // Slide 2 ── What is a Research Problem?
  {
    const s = contentSlide(p, H, "What is a Research Problem?");
    s.addText([
      { text: "A Research Problem ", options: { bold: true } },
      { text: "is a clear, precise expression of an area of concern, a gap or contradiction in literature, or a practical challenge that requires systematic investigation to solve.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Characteristics of a Good Problem",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Novelty & Originality:", options: { bold: true, color: DGN } }, { text: " Must address a genuine gap. Should not duplicate completed research.", options: { breakLine: true } },
      { text: "• Feasibility & Resources:", options: { bold: true } }, { text: " Can be solved within standard constraints (time, budget, equipment).", options: { breakLine: true } },
      { text: "• Social & Scientific Value:", options: { bold: true } }, { text: " Practical or theoretical relevance, adding value to academia or industry.", options: { breakLine: true } },
      { text: "• Scope Boundary:", options: { bold: true } }, { text: " Clearly bounded so it is manageable — avoiding topics that are too broad.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Primary Sources of Research Problems",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Literature Gaps:", options: { bold: true } }, { text: " Contradictory findings in journals, recommendations for future study.", options: { breakLine: true } },
      { text: "• Practical Challenges:", options: { bold: true, color: M } }, { text: " Real-world engineering issues, system inefficiencies, or software bugs.", options: { breakLine: true } },
      { text: "• Technological Changes:", options: { bold: true } }, { text: " New tools (AI, cloud) create fresh domains needing new research.", options: { breakLine: true } },
      { text: "• Conceptual Theories:", options: { bold: true } }, { text: " Testing established scientific theories in new, unexplored contexts.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by understanding what a research problem is. A research problem is a clear, precise expression of an area of concern, a gap or contradiction in existing literature, or a practical challenge that requires systematic investigation to solve. A well-formulated research problem must possess several key characteristics. It must be novel and original, meaning it addresses a genuine gap in knowledge and does not duplicate completed studies. It must be feasible, meaning it can be solved within standard constraints like time, budget, and available equipment. It must hold scientific or social value, offering practical or theoretical relevance to academia or industry. Additionally, it must have clear boundaries to ensure it remains manageable. We identify these problems through academic literature reviews, observing real-world engineering inefficiencies, tracking technological advancements, or testing established scientific theories in new, unexplored contexts.");
  }

  // Slide 3 ── Errors in Selecting a Problem
  {
    const s = contentSlide(p, H, "Errors in Problem Selection");
    s.addText("Researchers often commit classic errors when selecting and defining their research problem:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Common Selection Errors",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Choosing Too Broad a Topic:", options: { bold: true, color: M } }, { text: " E.g., 'Improving Cloud Database Uptime'. A single research paper cannot solve this. It must be narrowed to a specific query, layer, or mechanism.", options: { breakLine: true } },
      { text: "• Choosing Too Narrow a Topic:", options: { bold: true } }, { text: " E.g., 'Optimizing database index lookups for one local restaurant system'. Lacks general applicability or scientific value.", options: { breakLine: true } },
      { text: "• Overlooking Constraints:", options: { bold: true } }, { text: " Selecting problems requiring equipment (like supercomputers or specialized hardware) the university does not possess.", options: { breakLine: true } },
      { text: "• Pre-conceived Bias:", options: { bold: true } }, { text: " Selecting a topic to prove a personal bias rather than exploring objectively.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Approaches of Investigation",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Systematic Data Collection:", options: { bold: true } }, { text: " Formulating protocols to log experimental parameters or user metrics.", options: { breakLine: true } },
      { text: "• Analytical Interpretation:", options: { bold: true, color: DGN } }, { text: " Applying statistical validation tests (t-tests, ANOVA) to verify if results are significant.", options: { breakLine: true } },
      { text: "• Instrumentation Requirements:", options: { bold: true } }, { text: " Identifying the necessary software tools, datasets, and hardware metrics (e.g. CPU temperature, network latency) to measure results accurately.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Next, let us discuss the common errors researchers make when selecting a research problem. The first classic error is choosing a topic that is too broad, such as 'Improving Cloud Database Uptime'. A single research project cannot solve such a massive, unconstrained problem; it must be narrowed to a specific query, layer, or mechanism. Conversely, choosing too narrow a topic, like optimizing database index lookups for a single local restaurant's local system, lacks general applicability and scientific value. Another common pitfall is overlooking resource constraints, such as selecting problems that require specialized hardware or supercomputers that are unavailable. Researchers must also avoid pre-conceived biases, ensuring they explore topics objectively. To address these issues, a researcher must plan their systematic investigation early, establishing protocols for data collection, analytical interpretation using statistical validation, and identifying the necessary software tools and hardware metrics.");
  }

  // Slide 4 ── Systematic Investigation Flow
  {
    const s = contentSlide(p, H, "Systematic Investigation Flow");
    s.addText("The scientific process of investigating a research problem proceeds in six steps:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LBL, BL);
    s.addText("The Six Steps of Investigation",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Identify & Define the Problem:", options: { bold: true } }, { text: " Translate an area of concern into a specific, bounded research question.", options: { breakLine: true } },
      { text: "2. Construct Hypotheses:", options: { bold: true } }, { text: " Formulate testable assertions about relationships between variables.", options: { breakLine: true } },
      { text: "3. Design Experiments:", options: { bold: true } }, { text: " Plan the test setup, control conditions, and target metrics.", options: { breakLine: true } },
      { text: "4. Collect Data:", options: { bold: true, color: DGN } }, { text: " Run code simulations, query database testbeds, or gather hardware logs systematically.", options: { breakLine: true } },
      { text: "5. Analyze & Interpret:", options: { bold: true, color: DGN } }, { text: " Apply statistical modeling (e.g., regressions) to process values.", options: { breakLine: true } },
      { text: "6. Draw Conclusions & Publish:", options: { bold: true } }, { text: " Determine if the hypotheses were validated. Document limitations and publish findings.", options: {} }
    ], { x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("The scientific process of investigating a research problem proceeds through a structured, six-step lifecycle. The first step is to identify and define the problem, translating an area of concern into a specific, bounded research question. The second step is to construct hypotheses, formulating testable assertions about the relationships between variables. The third step is to design experiments, planning the test setup, control conditions, and target metrics. The fourth step is data collection, running code simulations, querying database testbeds, or gathering hardware logs systematically. The fifth step is to analyze and interpret the collected data, applying statistical modeling like regressions to process the results. Finally, the sixth step is to draw conclusions, determining if the hypotheses were validated, documenting limitations, and publishing the findings. This structured approach prevents research bias and ensures results are reproducible.");
  }

  // Slide 5 ── Applications
  {
    const s = contentSlide(p, H, "Applications of Problem Formulation");
    s.addText("Formulating a precise research problem is critical in academic and corporate domains:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📄  Research Grant Proposals",
        b: "Funding bodies (e.g., DST, UGC, corporate foundations) reject proposals that have vague problem definitions. A clear, well-formulated problem statement is essential for securing research funding.", bg: LGN, bd: GN
      },
      {
        t: "🎓  Master's & PhD Thesis",
        b: "Students must defend their thesis topic choice before review committees. A clearly bounded, novel problem statement guarantees a clear direction for the research.", bg: LBL, bd: BL
      },
      {
        t: "🏢  Corporate R&D Initiatives",
        b: "Tech companies frame R&D goals as precise problems (e.g., 'Reducing latency in distributed databases under 90% load'). Guides engineering teams towards clear performance targets.", bg: LMR, bd: M
      },
      {
        t: "🔬  Standardizing Patents",
        b: "Patent offices require inventions to solve a specific 'technical problem' in a non-obvious way. A precise definition of the problem is key to getting a patent granted.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Formulating a precise research problem is critical in academic and corporate domains. First, in research grant proposals, funding bodies like DST, UGC, or corporate foundations reject proposals with vague problem definitions. A clear, well-formulated problem statement is essential for securing research funding. Second, for graduate students, a clearly bounded, novel problem statement guarantees a clear direction for their master's or PhD thesis defense. Third, in corporate R&D initiatives, technology companies frame R&D goals as precise problems to guide engineering teams toward clear performance targets. Finally, patent offices require inventions to solve a specific technical problem in a non-obvious way, making a precise problem definition key to getting a patent granted.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, formulating a clear and bounded research problem is the foundation of the scientific method. By avoiding common selection errors and following a systematic investigation flow, researchers can generate high-value, reproducible solutions that advance both science and industry. Thank you for your time, and I am happy to take any questions.");
  await p.writeFile({ fileName: "./outputs/Sem1_RMIPR_Unit1_Research_Problem.pptx" });
  console.log("✔  RMIPR Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – LITERATURE REVIEW & RESEARCH ETHICS  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Research Methodology & IPR";

  const tSl = titleSlide(p, H, "Unit II: Literature Studies & Ethics", "Literature Review & Research Ethics");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Literature Review & Research Ethics' as part of the Research Methodology and IPR course for my M.Tech in Computer Science and Engineering. In this presentation, we will discuss how to conduct effective literature reviews, examine the different forms of plagiarism and how to prevent them, outline research ethics and scientific misconduct, and discuss ethical standards in academic publishing.");

  // Slide 2 ── Effective Literature Study
  {
    const s = contentSlide(p, H, "Effective Literature Studies");
    s.addText([
      { text: "A Literature Review ", options: { bold: true } },
      { text: "is a systematic search, analysis, and synthesis of published academic research. It provides background context, tracks prior studies, and identifies gaps for your project.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Objectives of Literature Review",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Map the Domain:", options: { bold: true } }, { text: " Understand the historical development and current state of research in the field.", options: { breakLine: true } },
      { text: "• Identify Gaps:", options: { bold: true, color: DGN } }, { text: " Locate contradictions, limitations, or unaddressed areas in prior studies.", options: { breakLine: true } },
      { text: "• Avoid Duplication:", options: { bold: true } }, { text: " Verify that your proposed solution has not already been published.", options: { breakLine: true } },
      { text: "• Find Methodologies:", options: { bold: true } }, { text: " Learn from experimental setups, database testbeds, and statistical techniques used by others.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LGN, GN);
    s.addText("Systematic Review Approach",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Keyword Search:", options: { bold: true } }, { text: " Formulate queries for databases like IEEE Xplore, ACM Digital Library, Google Scholar.", options: { breakLine: true } },
      { text: "• Filtering & Screening:", options: { bold: true } }, { text: " Screen papers by title, abstract, and then select candidate papers for full reading.", options: { breakLine: true } },
      { text: "• Synthesis vs. Summary:", options: { bold: true, color: M } }, { text: " Avoid simply listing papers. Group them by themes, compare their approaches, and outline trends in a structured synthesis.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by exploring effective literature studies. A literature review is a systematic search, analysis, and synthesis of published academic research. It provides background context, tracks prior studies, and identifies gaps for your project. The primary objectives are to map the domain, identify literature gaps, avoid duplication of work, and discover established research methodologies. A systematic review approach involves formulating keyword search queries for databases like IEEE Xplore, ACM Digital Library, and Google Scholar. Researchers must screen papers by title and abstract, select candidate papers for full reading, and synthesize the findings. It is crucial to focus on synthesis rather than summary: instead of simply listing papers, researchers should group them by themes, compare their approaches, and outline trends in a structured synthesis.");
  }

  // Slide 3 ── Plagiarism in Research
  {
    const s = contentSlide(p, H, "Plagiarism in Research");
    s.addText("Plagiarism is the appropriation of another person's ideas, processes, results, or words without giving appropriate credit:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Forms of Plagiarism",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Direct Plagiarism:", options: { bold: true, color: M } }, { text: " Word-for-word copy of text from a source without quotation marks or citations.", options: { breakLine: true } },
      { text: "• Paraphrased Plagiarism:", options: { bold: true } }, { text: " Copying ideas and rewriting sentences while keeping the original structure, without citation.", options: { breakLine: true } },
      { text: "• Self-Plagiarism (Duplicate Publication):", options: { bold: true } }, { text: " Re-using your own previously published code, text, or results in a new paper without citing the prior work.", options: { breakLine: true } },
      { text: "• Mosaic Plagiarism:", options: { bold: true } }, { text: " Patching together text from various sources, changing only a few words.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Prevention & Detection",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Proper Citation:", options: { bold: true, color: DGN } }, { text: " Always cite the source when referencing ideas, data, diagrams, or quotes.", options: { breakLine: true } },
      { text: "• Use Quotation Marks:", options: { bold: true } }, { text: " Enclose exact copied phrases in quotes.", options: { breakLine: true } },
      { text: "• Paraphrase properly:", options: { bold: true } }, { text: " Synthesize ideas in your own words and write original sentences.", options: { breakLine: true } },
      { text: "• Checker Tools:", options: { bold: true } }, { text: " Run manuscripts through checker tools (e.g. Turnitin, iThenticate) before submission. Academic journals reject papers exceeding strict similarity limits (typically 10-15%).", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Next, let us discuss plagiarism, which is the appropriation of another person's ideas, processes, results, or words without giving appropriate credit. Plagiarism is a serious ethical violation and takes several forms. Direct plagiarism is a word-for-word copy of text from a source without quotation marks or citations. Paraphrased plagiarism involves copying ideas and rewriting sentences while keeping the original structure, without citation. Self-plagiarism is the re-use of one's own previously published code, text, or results in a new paper without citing the prior work. Mosaic plagiarism is patching together text from various sources, changing only a few words. To prevent plagiarism, researchers must always cite sources when referencing ideas, data, diagrams, or quotes, enclose exact copied phrases in quotation marks, and write original sentences. Additionally, manuscripts must be checked using tools like Turnitin before submission.");
  }

  // Slide 4 ── Research Ethics
  {
    const s = contentSlide(p, H, "Research Ethics & Integrity");
    s.addText("Ethical guidelines protect the integrity of the scientific record and researchers:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, GN);
    s.addText("Scientific Misconduct",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Data Fabrication:", options: { bold: true, color: M } }, { text: " Making up fake experimental data or results and recording them as real.", options: { breakLine: true } },
      { text: "• Data Falsification:", options: { bold: true, color: M } }, { text: " Manipulating research equipment, omitting data points, or altering results to support a desired conclusion.", options: { breakLine: true } },
      { text: "• Selective Reporting:", options: { bold: true } }, { text: " Reporting only successful runs while omitting failed attempts or outliers.", options: { breakLine: true } },
      { text: "• Conflict of Interest:", options: { bold: true } }, { text: " Failing to disclose financial backing or relationships that could bias findings.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Publication Standards",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Authorship Ethics:", options: { bold: true } }, { text: " Only include authors who made significant contributions to the design, execution, or writing. Avoid 'gift' or 'ghost' authors.", options: { breakLine: true } },
      { text: "• Duplicate Submission:", options: { bold: true, color: M } }, { text: " Submitting the same paper to multiple journals simultaneously is strictly prohibited.", options: { breakLine: true } },
      { text: "• Peer Review Integrity:", options: { bold: true } }, { text: " Reviewers must evaluate papers objectively, keep content confidential, and declare conflicts of interest.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Beyond plagiarism, research ethics and scientific integrity require absolute honesty. Scientific misconduct includes data fabrication, which is making up fake experimental data, and data falsification, which is manipulating research equipment or omitting data points to support a desired conclusion. Selective reporting—reporting only successful runs while omitting failed attempts—is also unethical. Additionally, failing to disclose financial backing or relationships that could bias findings represents a conflict of interest. Publication standards also dictate authorship ethics, requiring that only individuals who made significant contributions to the design, execution, or writing be included. Duplicate submission—submitting the same paper to multiple journals simultaneously—is strictly prohibited, and peer reviewers must evaluate papers objectively, keep content confidential, and declare any conflicts of interest.");
  }

  // Slide 5 ── Applications in Publishing
  {
    const s = contentSlide(p, H, "Applications in Academic Publishing");
    s.addText("Ethical guidelines and literature review standards are checked throughout the publishing cycle:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📝  Journal Peer Reviewing",
        b: "Submitted papers undergo rigorous peer review. Reviewers check the literature review section to ensure the researcher has mapped the field correctly and cited the latest advances.", bg: LGN, bd: GN
      },
      {
        t: "🎓  University Thesis Defense",
        b: "Graduate schools run thesis submissions through plagiarism checkers. Students must pass validation checks (similarity thresholds) to graduate.", bg: LBL, bd: BL
      },
      {
        t: "🛡️  Retraction of Publications",
        b: "Journals retract published papers if fabrication, falsification, or plagiarism is exposed, permanently damaging the author's reputation and funding.", bg: LMR, bd: M
      },
      {
        t: "🌐  Open-Science Telemetry",
        b: "Modern research sites require uploading raw data and code alongside the manuscript, enabling public replication and verifying findings.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Ethical guidelines and literature review standards are checked throughout the publishing cycle. In journal peer review, reviewers check the literature review section to ensure the researcher has cited the latest advances and mapped the field correctly. Graduate schools run thesis submissions through plagiarism checkers, and students must pass strict similarity thresholds to graduate. Journals retract published papers if fabrication, falsification, or plagiarism is exposed, which permanently damages the author's reputation and funding. Lastly, modern open-science telemetry requires uploading raw data and code alongside the manuscript, enabling public replication and verifying findings. Adhering to these standards ensures your research is valid and trusted.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, conducting a systematic literature review and maintaining strict ethical standards are essential for the credibility of scientific research. By preventing plagiarism and avoiding scientific misconduct, we protect the integrity of the academic record. Thank you for your attention, and I am ready to answer any questions.");
  await p.writeFile({ fileName: "./outputs/Sem1_RMIPR_Unit2_Literature_Ethics.pptx" });
  console.log("✔  RMIPR Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – RESEARCH PROPOSAL STRUCTURE  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Research Methodology & IPR";

  const tSl = titleSlide(p, H, "Unit III: Technical Writing & Proposal Development", "Research Proposal Structure");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Research Proposal Structure' as part of the Research Methodology and IPR course for my M.Tech in Computer Science and Engineering. In this presentation, we will define a research proposal, examine its standard structure, outline technical writing and citation standards, and look at practical applications of research proposals in securing funding and academic approvals.");

  // Slide 2 ── What is a Research Proposal?
  {
    const s = contentSlide(p, H, "What is a Research Proposal?");
    s.addText([
      { text: "A Research Proposal ", options: { bold: true } },
      { text: "is a structured, formal document that details a plan for a research project. It outlines the research question, provides background context, details the proposed methodology, and requests budget/timeline approvals.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Purpose of the Proposal",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Secure Approvals:", options: { bold: true } }, { text: " Gets thesis topic approval from university review committees.", options: { breakLine: true } },
      { text: "• Secure Funding:", options: { bold: true, color: DGN } }, { text: " Convinces government agencies (DST, AICTE) or corporate R&D boards to fund the project.", options: { breakLine: true } },
      { text: "• Project Blueprint:", options: { bold: true } }, { text: " Acts as the structured plan for the project, keeping research focused.", options: { breakLine: true } },
      { text: "• Demonstrate Feasibility:", options: { bold: true } }, { text: " Shows that you have the skills, resources, and timeline to solve the problem.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Key Proposal Questions",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "A successful proposal must answer four basic questions:", options: { breakLine: true } },
      { text: "1. What do you want to accomplish?", options: { bold: true, color: M, breakLine: true } },
      { text: "2. Why is this problem important?", options: { bold: true, color: M, breakLine: true } },
      { text: "3. How will you solve the problem?", options: { bold: true, color: M, breakLine: true } },
      { text: "4. How much time and budget is needed?", options: { bold: true, color: M, breakLine: true } },
      { text: "• Style:", options: { bold: true } }, { text: " Written in an objective, precise, and formal technical style. Avoids marketing language or speculative claims.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us define what a research proposal is. A research proposal is a structured, formal document that details a plan for a research project. It outlines the research question, provides background context, details the proposed methodology, and requests budget and timeline approvals. The core purpose of the proposal is to secure approvals from university review committees, secure research grants from government agencies or corporate R&D boards, act as a project blueprint to keep the research focused, and demonstrate feasibility. A successful proposal must answer four basic questions: What do you want to accomplish? Why is this problem important? How will you solve the problem? And how much time and budget is needed? The proposal must be written in an objective, precise, and formal technical style, avoiding marketing language or speculative claims.");
  }

  // Slide 3 ── Typical Structure of a Proposal
  {
    const s = contentSlide(p, H, "Structure of a Research Proposal");
    s.addText("A standard research proposal contains the following structured sections:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LYL, AM);
    s.addText("Standard Proposal Layout Table",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });

    const tData = [
      [
        { text: "Proposal Section", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "Core Content & Focus", options: { bold: true, fill: { color: AM }, color: WH } },
        { text: "Key Output Details", options: { bold: true, fill: { color: AM }, color: WH } }
      ],
      [
        { text: "Title & Abstract", options: { bold: true, fill: { color: LGR }, align: "center" } }, { text: "Precise summary of the research topic and proposed approach.", options: { fill: { color: LGR } } }, { text: "Project title, summary of goals, methods, and expected results.", options: { fill: { color: LGR } } }
      ],
      [
        { text: "Introduction", options: { bold: true, fill: { color: LBL }, align: "center" } }, { text: "Formulates the problem statement, scope, and objectives.", options: { fill: { color: LBL } } }, { text: "Research questions, significance, and project limitations.", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Literature Review", options: { bold: true, fill: { color: LGN }, align: "center" } }, { text: "Critical synthesis of prior work, highlighting the research gap.", options: { fill: { color: LGN } } }, { text: "Citations of key papers, summary of gaps to be addressed.", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Methodology", options: { bold: true, fill: { color: LMR }, align: "center" } }, { text: "Step-by-step description of experiments, data, and tools.", options: { fill: { color: LMR } } }, { text: "Flowcharts, experimental setup, and statistical analysis plans.", options: { fill: { color: LMR } } }
      ],
      [
        { text: "Timeline & Budget", options: { bold: true, fill: { color: LYL }, align: "center" } }, { text: "Project schedule (Gantt chart) and itemized cost details.", options: { fill: { color: LYL } } }, { text: "Milestones, cost of equipment, travel, and personnel.", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.0, 5.5, 4.61]
      });

    s.addNotes("A standard research proposal contains several structured sections. The Title & Abstract provide a precise summary of the research topic, proposed approach, and expected results. The Introduction formulates the problem statement, scope, and objectives. The Literature Review provides a critical synthesis of prior work, highlighting the research gap and proving the novelty of the project. The Methodology section describes the step-by-step plan, experiments, datasets, and tools, utilizing flowcharts and block diagrams. The Timeline & Budget section outlines the project schedule, often using a Gantt chart, and itemizes costs for equipment, travel, and personnel. Following this structured layout is essential to satisfy professional review standards.");
  }

  // Slide 4 ── Technical Writing Standards
  {
    const s = contentSlide(p, H, "Technical Writing & Citation Standards");
    s.addText("Proposals must be written following academic style and citation guidelines:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Technical Writing Guidelines",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Objective Tone:", options: { bold: true } }, { text: " Present facts, data, and logical arguments. Avoid using emotional or subjective modifiers ('this amazing system').", options: { breakLine: true } },
      { text: "• Clarity & Conciseness:", options: { bold: true, color: M } }, { text: " Use simple, direct sentences. Avoid wordy explanations.", options: { breakLine: true } },
      { text: "• Active Voice preference:", options: { bold: true } }, { text: " Use active voice where possible for clarity:\n   - Good: 'We designed a K-means algorithm.'\n   - Avoid: 'A K-means algorithm was designed by us.'", options: { breakLine: true } },
      { text: "• Visual Elements:", options: { bold: true } }, { text: " Integrate flowcharts, block diagrams, and tables to explain complex logic.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Academic Citation Standards",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Academic Attribution:", options: { bold: true } }, { text: " Every claim of fact, dataset, or diagram from prior work must be cited.", options: { breakLine: true } },
      { text: "• Citation Formats:", options: { bold: true, color: DGN } }, { text: " Follow the style required by the funding agency or journal:\n", options: { breakLine: true } },
      { text: "   - IEEE Format:", options: { bold: true } }, { text: " Numbered bracket citations, standard in engineering: '[1] J. Smith...'.\n", options: { breakLine: true } },
      { text: "   - APA Format:", options: { bold: true } }, { text: " Author-date style, common in social sciences: '(Smith, 2023)'.\n", options: { breakLine: true } },
      { text: "• Reference Managers:", options: { bold: true } }, { text: " Use tools (e.g. Zotero, Mendeley) to generate and format references automatically.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Research proposals must be written following academic style and citation guidelines. Technical writing guidelines require an objective tone, presenting facts and logical arguments while avoiding subjective modifiers. Writers should use simple, direct sentences for clarity and conciseness, and prefer the active voice, such as 'We designed a K-means algorithm' rather than 'A K-means algorithm was designed by us.' Every claim of fact, dataset, or diagram from prior work must be cited. We follow standard citation formats required by funding agencies, such as the IEEE format, which uses numbered brackets, or the APA format, which uses the author-date style. Reference managers like Zotero or Mendeley should be utilized to generate and format references automatically, ensuring bibliography consistency.");
  }

  // Slide 5 ── Proposal Applications
  {
    const s = contentSlide(p, H, "Applications of Research Proposals");
    s.addText("Where research proposals are used in professional and academic settings:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🏛️  Government R&D Funding",
        b: "Securing research grants from agencies (like DST, DRDO, or AICTE) requires submitting a detailed proposal that details the budget, equipment costs, and milestones.", bg: LGN, bd: GN
      },
      {
        t: "🎓  PhD Committee Approval",
        b: "PhD candidates draft a proposal to define their research direction, defending it before a review committee to gain approval to begin dissertation work.", bg: LBL, bd: BL
      },
      {
        t: "🏢  Corporate R&D Funding",
        b: "Industrial researchers pitch proposals to executive boards to secure budgets for innovative product developments or patentable designs.", bg: LMR, bd: M
      },
      {
        t: "🤝  Collaborative Bids",
        b: "Consortiums of universities and industry partners submit research proposals to bid on large-scale international projects (like Horizon Europe).", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Research proposals are widely used in professional and academic settings. First, securing research grants from government agencies like DST, DRDO, or AICTE requires submitting a detailed proposal that details the budget, equipment costs, and milestones. Second, PhD candidates draft a proposal to define their research direction, defending it before a review committee to gain approval to begin dissertation work. Third, industrial researchers pitch proposals to executive boards to secure budgets for innovative product developments or patentable designs. Finally, collaborative bids by consortiums of universities and industry partners use proposals to secure funding for large-scale international projects. Writing a clear, structured proposal is the key to securing resources for any project.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, a well-structured research proposal is essential for defining a clear research plan and securing funding and approvals. By adhering to technical writing and citation standards, we ensure that our research plan is clear, professional, and feasible. Thank you for your time, and I am ready to take any questions.");
  await p.writeFile({ fileName: "./outputs/Sem1_RMIPR_Unit3_Research_Proposal.pptx" });
  console.log("✔  RMIPR Unit 3 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 4 – INTELLECTUAL PROPERTY RIGHTS (Unit IV)
// ═══════════════════════════════════════════════════════════════
async function ppt4() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Research Methodology & IPR";

  const tSl = titleSlide(p, H, "Unit IV: Intellectual Property", "Intellectual Property Rights (IPR)");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Intellectual Property Rights' as part of the Research Methodology and IPR course for my M.Tech in Computer Science and Engineering. In this presentation, we will define intellectual property, discuss the purpose of IPR, examine the four primary types of intellectual property, detail the patentability criteria for inventions, outline the patenting process, and look at real-world business applications of IPR.");

  // Slide 2 ── What is Intellectual Property?
  {
    const s = contentSlide(p, H, "What is Intellectual Property?");
    s.addText([
      { text: "Intellectual Property (IP) ", options: { bold: true } },
      { text: "refers to creations of the mind — such as inventions, literary and artistic works, designs, symbols, and brand names. IPR represents the legal rights that protect these creations.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("The Purpose of IPR",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Encourage Innovation:", options: { bold: true } }, { text: " Grants creators exclusive monopoly rights, allowing them to recover R&D costs and profit from their work.", options: { breakLine: true } },
      { text: "• Public Disclosure:", options: { bold: true, color: DGN } }, { text: " In exchange for legal protection, inventors must publish the details of their creation, expanding global scientific knowledge.", options: { breakLine: true } },
      { text: "• Economic Asset:", options: { bold: true } }, { text: " IP can be bought, sold, licensed, or used as collateral for corporate funding.", options: { breakLine: true } },
      { text: "• Consumer Protection:", options: { bold: true } }, { text: " Trademarks prevent customer confusion by verifying the origin of products.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Four Primary Types of IP",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Patents:", options: { bold: true, color: M } }, { text: " Protects novel, useful inventions and processes (e.g. computer processors). Typically granted for 20 years.", options: { breakLine: true } },
      { text: "• Industrial Designs:", options: { bold: true, color: M } }, { text: " Protects the visual, aesthetic appearance of an object (e.g. smart phone shape).", options: { breakLine: true } },
      { text: "• Trademarks:", options: { bold: true, color: M } }, { text: " Protects brand names, logos, and slogans (e.g. 'Apple' leaf logo). Prevents brand copycats.", options: { breakLine: true } },
      { text: "• Copyrights:", options: { bold: true, color: M } }, { text: " Protects original artistic, musical, or literary works. E.g. books, music, and software code.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by defining intellectual property, or IP. Intellectual property refers to creations of the mind—such as inventions, literary and artistic works, designs, symbols, and brand names. Intellectual Property Rights, or IPR, represent the legal rights that protect these creations. The purpose of IPR is to encourage innovation by granting creators exclusive monopoly rights, allowing them to recover R&D costs and profit from their work. In exchange for this legal protection, inventors must publicly disclose the details of their creation, expanding global scientific knowledge. IP acts as an economic asset that can be bought, sold, licensed, or used as collateral. It also protects consumers by using trademarks to prevent customer confusion about product origins. The four primary types of IP are Patents, which protect novel inventions; Industrial Designs, which protect aesthetics; Trademarks, which protect brand logos; and Copyrights, which protect artistic works and software code.");
  }

  // Slide 3 ── Inventions & Patents
  {
    const s = contentSlide(p, H, "Inventions and Patentability Criteria");
    s.addText("A patent is a government grant giving an inventor exclusive rights to exclude others from using the invention:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, GN);
    s.addText("Three Criteria for Patentability",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Novelty (Newness):\n", options: { bold: true, color: DGN } },
      { text: "   The invention must not be described in prior art anywhere in the world before the filing date.\n\n", options: { color: BK } },
      { text: "2. Inventive Step (Non-Obviousness):\n", options: { bold: true, color: DGN } },
      { text: "   Must not be obvious to a person with ordinary skill in that technical field.\n\n", options: { color: BK } },
      { text: "3. Industrial Applicability (Utility):\n", options: { bold: true, color: DGN } },
      { text: "   Must be useful and capable of being manufactured or used in some industry.", options: { color: BK } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("What CANNOT be Patented?",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Most patent systems (including India's) exclude:", options: { breakLine: true } },
      { text: "• Mathematical Formulas:", options: { bold: true, color: M } }, { text: " Pure algorithms or mathematical models (e.g. Einstein's E=mc²).", options: { breakLine: true } },
      { text: "• Laws of Nature:", options: { bold: true } }, { text: " Discovering a new plant or natural mineral in the wild.", options: { breakLine: true } },
      { text: "• Business Methods:", options: { bold: true } }, { text: " Abstract business concepts (e.g., standard auction steps).", options: { breakLine: true } },
      { text: "• Software as Code:", options: { bold: true, color: M } }, { text: " Raw code is protected by Copyrights, not patents. However, software that controls physical hardware can be patented in some regions.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("A patent is a government grant giving an inventor exclusive rights to exclude others from using their invention. To be patentable, an invention must meet three criteria. First, novelty: the invention must not be described in prior art anywhere in the world before the filing date. Second, inventive step, or non-obviousness: the invention must not be obvious to a person with ordinary skill in that technical field. Third, industrial applicability: the invention must be useful and capable of being manufactured or used in industry. Most patent systems, including India's, exclude mathematical formulas, laws of nature, abstract business methods, and raw software code from patentability. Raw software code is protected by copyrights rather than patents, though software that controls physical hardware functions is patentable in some countries.");
  }

  // Slide 4 ── Process of Patenting & PCT
  {
    const s = contentSlide(p, H, "Process of Patenting & PCT");
    s.addText("How patents are filed locally and globally using international treaties:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGR, GY);
    s.addText("Standard Patenting Timeline",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Innovation & Prior Art Search:", options: { bold: true } }, { text: " Verify novelty using patent databases (Espacenet, USPTO).", options: { breakLine: true } },
      { text: "2. Draft Specification:", options: { bold: true } }, { text: " Write the patent description and compile the precise 'claims' defining your invention's boundaries.", options: { breakLine: true } },
      { text: "3. File Application:", options: { bold: true, color: BL } }, { text: " File a provisional application to secure your 'priority date'.", options: { breakLine: true } },
      { text: "4. Publication & Examination:", options: { bold: true } }, { text: " Patent office publishes the application and reviews claims against prior art.", options: { breakLine: true } },
      { text: "5. Grant of Patent:", options: { bold: true, color: DGN } }, { text: " If approved, the patent is granted, requiring annual renewal fees.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Patent Cooperation Treaty (PCT)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Territorial Limitation:", options: { bold: true, color: M } }, { text: " Patents are national. A patent granted in India does not protect your invention in the USA.", options: { breakLine: true } },
      { text: "• The PCT Solution:", options: { bold: true } }, { text: " Administered by WIPO, the PCT allows inventors to file a SINGLE 'international' application to secure priority in 150+ member countries simultaneously.", options: { breakLine: true } },
      { text: "• Priority Window:", options: { bold: true, color: DGN } }, { text: " Secures a priority window of 30 months. This gives inventors time to search for foreign investors before paying expensive foreign filing fees.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Patents are territorial, meaning they are valid only within the borders of the granting nation. The patenting process involves searching patent databases to verify novelty, drafting specifications with precise claims, filing an application to secure a priority date, publication, examination, and the final grant. Because of territorial limits, a patent granted in India does not protect an invention in the United States. To address this, the Patent Cooperation Treaty, or PCT, administered by WIPO, allows inventors to file a single international application to secure priority in over 150 member countries simultaneously. This secures a priority window of 30 months, giving inventors time to search for foreign investors and evaluate market potential before paying expensive local filing fees in each individual country.");
  }

  // Slide 5 ── IPR Applications
  {
    const s = contentSlide(p, H, "Applications of IP Protection");
    s.addText("Intellectual property rights are critical business assets across the tech sector:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "💻  Software Innovations Protection",
        b: "Tech companies protect proprietary software engines (like search algorithms or video compression codecs) using patents, while protecting source code files using copyrights.", bg: LGN, bd: GN
      },
      {
        t: "🚗  Industrial Design Rights",
        b: "Automotive and consumer tech companies register visual designs (like smartphone cases, car body lines, or watch faces) to prevent competitors from mimicking their aesthetics.", bg: LBL, bd: BL
      },
      {
        t: "💊  Pharmaceutical R&D Patents",
        b: "Drug companies patent new chemical compounds (typically costing billions to develop). Exclusive rights allow them to recover development costs prior to generic competition.", bg: LMR, bd: M
      },
      {
        t: "🤝  Standard Essential Patents (SEPs)",
        b: "Telecommunications groups patent technologies essential to standards (like 5G). They license these to hardware makers on FRAND terms, earning royalties.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Intellectual property rights are critical business assets across the tech sector. First, software companies protect proprietary algorithms and codecs using patents, while protecting source code files using copyrights. Second, consumer tech companies register industrial designs to protect smartphone shapes and product aesthetics from copycats. Third, pharmaceutical firms patent new chemical compounds to recover their massive R&D costs before generic competition enters the market. Finally, telecommunications groups patent technologies essential to global standards, like 5G, and license them as Standard Essential Patents on fair, reasonable, and non-discriminatory terms, earning substantial licensing royalties. Protecting intellectual property ensures companies can secure their innovations and generate revenue.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, intellectual property rights protect creative and technical innovations, driving economic growth and scientific progress. By understanding patentability criteria and utilizing international filing systems like the PCT, inventors can secure their designs globally. Thank you for your time, and I am happy to answer any questions.");
  await p.writeFile({ fileName: "./outputs/Sem1_RMIPR_Unit4_IPR_Nature.pptx" });
  console.log("✔  RMIPR Unit 4 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 5 – TECHNOLOGY TRANSFER & PATENT RIGHTS (Unit V)
// ═══════════════════════════════════════════════════════════════
async function ppt5() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Research Methodology & IPR";

  const tSl = titleSlide(p, H, "Unit V: Patent Rights & New Developments", "Technology Transfer & Patent Rights");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Technology Transfer & Patent Rights' as part of the Research Methodology and IPR course for my M.Tech in Computer Science and Engineering. In this presentation, we will discuss the scope of patent rights, explain the mechanics of technology transfer and licensing models, define geographical indications, outline new developments in biotech and software IP, and examine traditional knowledge protections.");

  // Slide 2 ── Scope of Patent Rights
  {
    const s = contentSlide(p, H, "Scope of Patent Rights");
    s.addText([
      { text: "A Patent ", options: { bold: true } },
      { text: "grants the owner exclusionary rights to prevent others from making, using, selling, or importing the invention. It represents a legal monopoly with distinct boundaries.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Legal Boundaries of Patents",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Exclusionary Right:", options: { bold: true, color: DGN } }, { text: " Does not grant the owner the right to make the invention — it only grants the right to stop OTHERS from doing so. (Subject to regulatory and prior patent locks).", options: { breakLine: true } },
      { text: "• Territorial Scope:", options: { bold: true } }, { text: " Rights are valid only within the boundaries of the granting nation.", options: { breakLine: true } },
      { text: "• Term Limit:", options: { bold: true } }, { text: " Standard term is 20 years from the filing date, after which it enters the public domain.", options: { breakLine: true } },
      { text: "• Maintenance Fees:", options: { bold: true, color: M } }, { text: " Requires paying annual maintenance fees to keep the patent valid.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Licensing & Technology Transfer",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Technology Transfer:", options: { bold: true } }, { text: " Sharing and commercializing innovations from research labs (universities) to industrial partners.", options: { breakLine: true } },
      { text: "• Licensing Models:", options: { bold: true, color: BL } }, { text: " Legal contracts granting third parties rights to use the IP in exchange for royalties:\n", options: { breakLine: true } },
      { text: "   - Exclusive License:", options: { bold: true } }, { text: " Only one licensee has the rights.\n", options: { breakLine: true } },
      { text: "   - Non-Exclusive:", options: { bold: true } }, { text: " Multiple competitors can buy rights (e.g. video codecs).\n", options: { breakLine: true } },
      { text: "• Cross-Licensing:", options: { bold: true } }, { text: " Competitors trade patent portfolios to prevent mutual litigation.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us examine the scope of patent rights and the process of technology transfer. A patent grants the owner exclusionary rights to prevent others from making, using, selling, or importing the invention. It represents a legal monopoly with distinct boundaries. Crucially, it does not grant the owner the positive right to make the invention, but only the negative right to stop others from doing so, subject to regulatory approvals. Patent rights are territorial, valid only in the granting nation, and have a standard term limit of 20 years from the filing date, requiring annual maintenance fees. Technology transfer is the process of sharing and commercializing innovations from university research labs to industrial partners. This is done using licensing contracts, which can be exclusive, non-exclusive, or cross-licensing deals where competitors trade portfolios to avoid litigation.");
  }

  // Slide 3 ── Geographical Indications (GI)
  {
    const s = contentSlide(p, H, "Geographical Indications (GIs)");
    s.addText("Geographical Indications protect products whose quality, characteristics, or reputation are due to their origin:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("What is a GI?",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Geographical Origin:", options: { bold: true } }, { text: " A label used on agricultural, natural, or manufactured goods that have a specific geographic origin.", options: { breakLine: true } },
      { text: "• Inherent Quality Link:", options: { bold: true, color: M } }, { text: " The product's characteristics must be due to the soil, climate, or traditional methods of that origin (e.g. Darjeeling Tea).", options: { breakLine: true } },
      { text: "• Collective Right:", options: { bold: true } }, { text: " Belongs to a community of local producers, not a single private company. Prevents misuse of the name by foreign competitors.", options: { breakLine: true } },
      { text: "• Duration:", options: { bold: true } }, { text: " Typically granted for 10 years, but can be renewed indefinitely.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("GI Examples in India",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "India enacted the GIs of Goods Act in 1999 to protect local specialties:", options: { breakLine: true } },
      { text: "• Darjeeling Tea:", options: { bold: true, color: DGN } }, { text: " The first GI tag registered in India (2004). Protections prevent foreign growers from labeling inferior tea as Darjeeling.", options: { breakLine: true } },
      { text: "• Basmati Rice:", options: { bold: true } }, { text: " Aromatic long-grain rice grown in specific northern plains.", options: { breakLine: true } },
      { text: "• Pashmina Shawls:", options: { bold: true } }, { text: " Hand-crafted wool products from Kashmir.", options: { breakLine: true } },
      { text: "• Pochampally Ikat:", options: { bold: true } }, { text: " Traditional handloom sarees from Telangana.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Next, let us discuss Geographical Indications, or GIs. A geographical indication is a label used on agricultural, natural, or manufactured goods that have a specific geographic origin. The product's quality, reputation, or characteristics must be due to the soil, climate, or traditional methods of that origin. Unlike private patents, a GI tag is a collective right belonging to a community of local producers, not a single private company. It prevents misuse of the name by foreign competitors and is typically granted for 10 years, renewable indefinitely. India enacted the GIs of Goods Act in 1999 to protect local specialties. Examples include Darjeeling Tea, which was the first GI tag registered in India, Basmati Rice, Kashmir Pashmina Shawls, and Pochampally Ikat sarees from Telangana.");
  }

  // Slide 4 ── New Developments in IPR
  {
    const s = contentSlide(p, H, "New Developments in IPR");
    s.addText("IP systems are adapting to protect digital, biological, and traditional innovations:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Biotechnology & Software IP",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Biological Systems:", options: { bold: true, color: BL } }, { text: " Patenting genetically modified organisms (like GM cotton seeds), gene editing protocols (CRISPR), and vaccine formulas. Raises ethical debates.", options: { breakLine: true } },
      { text: "• Software & AI IP:", options: { bold: true } }, { text: " AI algorithms are abstract math (non-patentable). However, systems where AI controls physical robots can be patented. Generative AI raises debates on copyright ownership of machine-generated code.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Traditional Knowledge (TK)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• The Challenge:", options: { bold: true } }, { text: " Bio-pirating companies patenting ancient remedies (e.g. Turmeric, Neem) in foreign patent offices.", options: { breakLine: true } },
      { text: "• The India Case:", options: { bold: true, color: M } }, { text: " India fought and revoked US patents on Turmeric healing properties by presenting ancient Sanskrit texts as prior art.", options: { breakLine: true } },
      { text: "• TKDL Database:", options: { bold: true, color: DGN } }, { text: " Traditional Knowledge Digital Library translates ancient medical texts into 5 global languages, allowing patent offices to find prior art immediately and block invalid patents.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("IP systems are evolving to handle genetic engineering, AI generation, and traditional knowledge. In biotechnology, patenting genetically modified organisms, gene editing protocols like CRISPR, and vaccine formulas has expanded rapidly, raising ethical debates. In software, while raw AI algorithms are abstract math and non-patentable, AI systems that control physical robots can be patented. Generative AI also raises debates on copyright ownership of machine-generated code. A major challenge is bio-piracy, where foreign companies attempt to patent ancient remedies, such as turmeric or neem. To combat this, India fought and revoked US patents by presenting ancient texts as prior art. India also established the Traditional Knowledge Digital Library (TKDL) database to translate ancient texts, allowing patent offices to find prior art immediately and block invalid patents.");
  }

  // Slide 5 ── Technology Transfer Applications
  {
    const s = contentSlide(p, H, "Applications of Technology Transfer");
    s.addText("How technology transfer mechanics bridge the gap between academic labs and market sales:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🏫  University spin-offs",
        b: "Universities set up Technology Transfer Offices (TTOs) to license patents created by professors and students to startup companies, commercializing research.", bg: LGN, bd: GN
      },
      {
        t: "💉  Vaccine Licensing Deals",
        b: "Research institutes (like Oxford University) patent vaccine formulas and license manufacturing rights to global firms (like AstraZeneca) for production.", bg: LBL, bd: BL
      },
      {
        t: "🌾  Agricultural Licensing",
        b: "Agricultural universities license disease-resistant seed strains to farming cooperatives, increasing crop yields through research.", bg: LMR, bd: M
      },
      {
        t: "🌐  Open-source patent pledges",
        b: "Tech companies pledge patent portfolios to open-source consortia, allowing startups to build services without litigation risks.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Technology transfer mechanics bridge the gap between academic labs and market sales. First, universities establish Technology Transfer Offices to license student and professor patents to startup spin-offs, commercializing academic research. Second, research institutes patent vaccine formulas and license manufacturing rights to global pharmaceutical firms for mass production. Third, agricultural universities license disease-resistant crop strains to farming cooperatives, increasing food yields. Finally, tech companies make open-source patent pledges to consortia, allowing startups to build services without litigation risks. These applications demonstrate how technology transfer translates theoretical research into real-world products.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, technology transfer and patent rights are vital for moving research from labs into the real world. By understanding geographical indications and defending traditional knowledge, we protect local and historical innovations in a globalized economy. Thank you for your time, and I am ready to answer your questions.");
  await p.writeFile({ fileName: "./outputs/Sem1_RMIPR_Unit5_Tech_Transfer.pptx" });
  console.log("✔  RMIPR Unit 5 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 5 RMIPR PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  await ppt4();
  await ppt5();
  console.log("All 5 RMIPR PPTs generated successfully!");
})();
