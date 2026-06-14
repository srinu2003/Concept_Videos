const pptxgen = require("pptxgenjs");
const { convertToPdf } = require("./pdf_converter");
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
//  PPT 1 – SOFTWARE QUALITY FRAMEWORKS (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Software Quality Engineering";

  const s1 = titleSlide(p, H, "Unit I: Software Quality", "Software Quality Frameworks & ISO-9126");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. I am a first-year, first-semester M.Tech student in the Department of Computer Science and Engineering. Today, I am going to present our first topic in Software Quality Engineering: Software Quality Frameworks and the ISO-9126 standard. In this presentation, we will explore what defines software quality, examine different stakeholder perspectives, study the historical progression of defect measurement, analyze the six core characteristics of the ISO-9126 quality model, compare it to McCall's and Boehm's frameworks, and discuss how these frameworks are applied in the industry today. Let's begin by looking at the core definition of software quality.");

  // Slide 2 ── What is Software Quality?
  {
    const s = contentSlide(p, H, "What is Software Quality?");
    s.addText([
      { text: "Software Quality ", options: { bold: true } },
      { text: "is the degree to which a software system satisfies its explicit functional requirements and implicit user expectations. Correctness and defects are its primary metrics.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Perspectives of Quality",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• User Perspective:", options: { bold: true } }, { text: " Reliability, usability, and responsiveness during application execution.", options: { breakLine: true } },
      { text: "• Developer Perspective:", options: { bold: true } }, { text: " Maintainability, clear structure, and testability of the source code.", options: { breakLine: true } },
      { text: "• Business Perspective:", options: { bold: true } }, { text: " Cost-effectiveness, time-to-market, and alignment with contract specs.", options: { breakLine: true } },
      { text: "• Correctness vs. Defects:", options: { bold: true, color: DGN } }, { text: " Correctness is adherence to specs; defects are deviations (faults, errors, failures) in execution.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Historical Defect Measurements",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Defect Density:", options: { bold: true } }, { text: " Measured as the number of defects per Thousand Lines of Code (KLOC):\n   Density = Defects / KLOC.", options: { breakLine: true } },
      { text: "• Fault vs. Failure:", options: { bold: true, color: M } }, { text: " A 'fault' is a bug in the code. A 'failure' is the execution of a fault causing incorrect user results.", options: { breakLine: true } },
      { text: "• Historical Progression:", options: { bold: true } }, { text: " Quality has progressed from pure post-development testing to active quality engineering throughout the development lifecycle.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("To start, let's address the fundamental question: What is Software Quality? Software quality is the degree to which a software system satisfies its explicit functional requirements and implicit user expectations. Correctness and defects are its primary metrics. From a stakeholder perspective, quality means different things: users care about reliability and responsiveness; developers focus on maintainability, clean structure, and testability; while the business looks at cost-effectiveness and alignment with specs. Historically, we have measured quality using defect density, which is the number of defects per Thousand Lines of Code (KLOC). It is also crucial to distinguish between a fault, which is a bug in the code, and a failure, which is the actual execution of that fault causing incorrect user results. Over time, quality practices have progressed from simple post-development testing to active quality engineering throughout the software development lifecycle.");
  }

  // Slide 3 ── ISO-9126 Quality Standard
  {
    const s = contentSlide(p, H, "ISO-9126 Quality Model");
    s.addText("ISO-9126 is an international standard defining a hierarchical software quality framework:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LYL, AM);
    s.addText("The Six Core Quality Characteristics",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });

    const tData = [
      [
        { text: "Characteristic", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "Definition / Focus", options: { bold: true, fill: { color: AM }, color: WH } },
        { text: "Key Metric Focus", options: { bold: true, fill: { color: AM }, color: WH } }
      ],
      [
        { text: "Functionality", options: { bold: true, fill: { color: LGR }, align: "center" } }, { text: "Existence of functions matching stated and implicit needs.", options: { fill: { color: LGR } } }, { text: "Accuracy, interoperability, compliance, security.", options: { fill: { color: LGR } } }
      ],
      [
        { text: "Reliability", options: { bold: true, fill: { color: LBL }, align: "center" } }, { text: "Software's capability to maintain performance levels over time.", options: { fill: { color: LBL } } }, { text: "Maturity, fault tolerance, recoverability.", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Usability", options: { bold: true, fill: { color: LGN }, align: "center" } }, { text: "Ease of understanding, learning, and operating the system.", options: { fill: { color: LGN } } }, { text: "Understandability, learnability, operability.", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Efficiency", options: { bold: true, fill: { color: LMR }, align: "center" } }, { text: "Performance relative to CPU, memory, and database resources.", options: { fill: { color: LMR } } }, { text: "Time behavior, resource utilization.", options: { fill: { color: LMR } } }
      ],
      [
        { text: "Maintainability", options: { bold: true, fill: { color: LYL }, align: "center" } }, { text: "Effort needed to make specified code modifications.", options: { fill: { color: LYL } } }, { text: "Analyzability, changeability, stability, testability.", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.0, 5.5, 4.61]
      });

    s.addNotes("Now, let's explore the ISO-9126 Quality Model. It is an international standard that defines a hierarchical software quality framework using six core characteristics. First is Functionality, which checks the existence of functions matching stated and implicit needs. Second is Reliability, focusing on the software's capability to maintain performance levels over time under specified conditions. Third is Usability, which measures the ease of understanding, learning, and operating the system. Fourth is Efficiency, evaluating performance relative to CPU, memory, and database resources. Fifth is Maintainability, which assesses the effort needed to make code modifications. Finally, Portability is the software's ability to be transferred from one environment to another. This model ensures that we assess software quality comprehensively from multiple dimensions, rather than just looking at the presence of bugs.");
  }

  // Slide 4 ── Quality Models Comparison
  {
    const s = contentSlide(p, H, "Comparison of Quality Models");
    s.addText("A structural comparison of ISO-9126 with McCall's and Boehm's quality frameworks:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const tData = [
      [
        { text: "Evaluation Category", options: { bold: true, fill: { color: M }, color: WH } },
        { text: "McCall's Quality Model", options: { bold: true, fill: { color: M }, color: WH, align: "center" } },
        { text: "Boehm's Quality Model", options: { bold: true, fill: { color: M }, color: WH, align: "center" } },
        { text: "ISO-9126 Standard", options: { bold: true, fill: { color: M }, color: WH, align: "center" } }
      ],
      [
        { text: "Structure", options: { bold: true, fill: { color: LGR } } }, { text: "3 product pillars (Operations, Revision, Transition)", options: { fill: { color: LGR } } }, { text: "Hierarchical tree representing user needs", options: { fill: { color: LGR } } }, { text: "6 main characteristics, 27 sub-characteristics", options: { fill: { color: LGR } } }
      ],
      [
        { text: "Focus Area", options: { bold: true, fill: { color: LBL } } }, { text: "Code properties and maintainer ease", options: { fill: { color: LBL } } }, { text: "Hardware utilization and efficiency", options: { fill: { color: LBL } } }, { text: "External quality, internal quality, quality in use", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Portability", options: { bold: true, fill: { color: LGN } } }, { text: "Pillar of Product Transition", options: { fill: { color: LGN } } }, { text: "Classified under General Utility", options: { fill: { color: LGN } } }, { text: "Primary top-level characteristic", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Target Audience", options: { bold: true, fill: { color: LMR } } }, { text: "Software maintainers & developers", options: { fill: { color: LMR } } }, { text: "System engineers & customers", options: { fill: { color: LMR } } }, { text: "End-users, developers, and QA auditors", options: { fill: { color: LMR } } }
      ]
    ];
    s.addTable(tData,
      {
        x: CX, y: 2.20, w: CW, h: 4.80, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.5, 3.28, 3.28, 3.29]
      });

    s.addNotes("Let's compare ISO-9126 with other classical frameworks, namely McCall's and Boehm's quality models. Structurally, McCall's model organizes quality into three product pillars: Product Operations, Product Revision, and Product Transition. Boehm's model uses a hierarchical tree representing user needs, while ISO-9126 uses six main characteristics and twenty-seven sub-characteristics. In terms of focus, McCall's emphasizes code properties and maintainability; Boehm's focuses on hardware utilization and efficiency; and ISO-9126 balances external quality, internal quality, and quality in use. Portability is a sub-category under Product Transition in McCall's and general utility in Boehm's, but is a top-level characteristic in ISO-9126. Finally, the target audience differs: McCall's is mainly for maintainers, Boehm's for system engineers, and ISO-9126 is designed for end-users, developers, and QA auditors alike.");
  }

  // Slide 5 ── Applications of Frameworks
  {
    const s = contentSlide(p, H, "Applications of Quality Frameworks");
    s.addText("How software companies implement these quality models in practice:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📝  Contract Requirements Alignment",
        b: "Organizations use ISO-9126 terms to specify contractual SLA bounds. E.g., 'Reliability must guarantee 99.9% uptime (fault tolerance) and Recoverability must be within 1 hour.'", bg: LGN, bd: GN
      },
      {
        t: "⚙️  Automated Code Scanning",
        b: "Static analysis tools (e.g. SonarQube) map code issues to Maintainability, Security, and Reliability buckets, calculating 'Technical Debt' dynamically.", bg: LBL, bd: BL
      },
      {
        t: "📊  Vendor Benchmarking",
        b: "Allows IT enterprises to objectively evaluate third-party software products against standardized quality checklists before procurement.", bg: LMR, bd: M
      },
      {
        t: "🛡️  Safety-Critical Certifications",
        b: "Aerospace and medical software must demonstrate compliance with strict Reliability and Maintainability standards prior to public deployment.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("How do organizations implement these quality models in practice? First, in Contract Requirements Alignment, companies use ISO-9126 terms to specify clear service level agreement boundaries, such as defining fault tolerance or recovery times. Second, in Automated Code Scanning, static analysis tools like SonarQube map code issues to reliability or maintainability buckets to calculate technical debt. Third, in Vendor Benchmarking, IT enterprises use standardized checklists based on these models to evaluate third-party software before purchasing. Lastly, in Safety-Critical Certifications, such as aerospace and medical applications, software must strictly demonstrate compliance with reliability and maintainability standards prior to public deployment. By using these frameworks, companies translate abstract quality concepts into measurable business attributes.");
  }

  const s6 = thankYouSlide(p, H);
  s6.addNotes("That brings me to the end of this presentation on Software Quality Frameworks and ISO-9126. We have seen how defining, measuring, and applying quality models helps build more reliable and maintainable systems. I would like to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management for this opportunity. Thank you everyone for your time and attention.");

  await p.writeFile({ fileName: "./outputs/Sem1_SQE_Unit1_Quality_Frameworks.pptx" });
  convertToPdf("./outputs/Sem1_SQE_Unit1_Quality_Frameworks.pptx");
  console.log("✔  SQE Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – DEFECT PREVENTION, REDUCTION & CONTAINMENT (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Software Quality Engineering";

  const s1 = titleSlide(p, H, "Unit II: Quality Assurance", "Defect Prevention, Reduction & Containment");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. In this presentation, I am going to talk about Defect Prevention, Reduction, and Containment under Software Quality Assurance. We will look at how we classify QA activities, analyze the exponential cost multiplier of defects, explore specific prevention techniques like formal specifications and Cleanroom engineering, study static inspections versus dynamic testing for defect reduction, examine fault tolerance and safety containment strategies, and see how modern DevOps teams apply these pillars in CI/CD pipelines and embedded control systems.");

  // Slide 2 ── QA Classification: Dealing with Defects
  {
    const s = contentSlide(p, H, "QA Classification: Dealing with Defects");
    s.addText([
      { text: "Quality Assurance ", options: { bold: true } },
      { text: "consists of activities that guarantee software matches quality goals. We classify QA operations based on how they handle defects: Prevention, Reduction, or Containment.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("The Cost of Defects",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Cost Multiplier:", options: { bold: true, color: M } }, { text: " The cost of fixing bugs grows exponentially as software moves from design to code, to test, and finally to production.", options: { breakLine: true } },
      { text: "• Phase of Origin:", options: { bold: true } }, { text: " Most defects originate during requirements and design phases, but are only detected during system testing.", options: { breakLine: true } },
      { text: "• QA Goal:", options: { bold: true, color: DGN } }, { text: " Introduce defect handling activities early in the lifecycle to minimize cost and risk.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Three Pillars of Defect QA",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Defect Prevention:", options: { bold: true, color: GN } }, { text: " Stop bugs from entering the code. Focuses on requirements clarity, design rules, and developer training.", options: { breakLine: true } },
      { text: "2. Defect Reduction:", options: { bold: true, color: BL } }, { text: " Find and remove bugs present in code. Uses inspection (static reviews) and execution-based testing.", options: { breakLine: true } },
      { text: "3. Defect Containment:", options: { bold: true, color: AM } }, { text: " Limit the damage of bugs that escape into production. Focuses on fault tolerance and safety containment.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's first understand the classification of QA and the cost of defects. Quality Assurance consists of all structured activities that guarantee software matches quality goals. We classify defect-handling operations into three pillars: Prevention, Reduction, and Containment. Defect Prevention stops bugs from entering the code by focusing on requirements clarity and developer training. Defect Reduction aims to find and remove bugs already present in the code through inspections and testing. Defect Containment limits the damage when bugs escape to production. Managing these early is key because the cost of fixing a bug grows exponentially as it moves from design to code, test, and finally to production. While most bugs originate in requirements and design, they are usually found late during testing, making early intervention critical to minimize cost and risk.");
  }

  // Slide 3 ── Defect Prevention Techniques
  {
    const s = contentSlide(p, H, "Defect Prevention");
    s.addText("Defect Prevention is the most cost-effective QA strategy, acting before bugs are coded:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, GN);
    s.addText("Developer Training & Spec Clarity",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Training & Education:", options: { bold: true } }, { text: " Educating developers on secure coding standards and common architectural bugs (e.g., buffer overflows).", options: { breakLine: true } },
      { text: "• Formal Specifications:", options: { bold: true, color: DGN } }, { text: " Using formal mathematical specification languages (e.g. Z notation) to define system interfaces, preventing requirements ambiguity.", options: { breakLine: true } },
      { text: "• Root Cause Analysis (RCA):", options: { bold: true } }, { text: " Analyzing past bugs to identify process improvements, updating design rules to prevent recurrence.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Process & Technology Tools",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Design Rules & Checklists:", options: { bold: true } }, { text: " Enforcing coding conventions, strict design patterns, and type checking constraints.", options: { breakLine: true } },
      { text: "• Cleanroom Software Engineering:", options: { bold: true, color: M } }, { text: " A process model that uses formal design verification and mathematical proofs to ensure defect-free builds before testing begins.", options: { breakLine: true } },
      { text: "• Automated Linting:", options: { bold: true } }, { text: " Run code analyzers during development to catch syntax errors and poor style patterns automatically.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Defect Prevention is the most cost-effective strategy. We achieve it through developer training on secure coding and formal specification languages like Z-notation, which use math to eliminate requirements ambiguity. Root Cause Analysis is also used to study past bugs and update design rules. On the process side, we enforce coding standards, static checklists, and automated linter tools that scan code on the fly. We also have Cleanroom Software Engineering, a formal process model that uses mathematical proofs to verify code correctness before testing ever starts. By using these tools and practices, we prevent errors from being written in the first place, saving significant debugging effort.");
  }

  // Slide 4 ── Defect Reduction (Inspection & Testing)
  {
    const s = contentSlide(p, H, "Defect Reduction: Inspection & Testing");
    s.addText("Defect reduction detects and removes bugs that have already been coded:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Inspections (Static Verification)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Non-Execution Review:", options: { bold: true } }, { text: " Reviews design docs, source code, and schemas without running the application.", options: { breakLine: true } },
      { text: "• Peer Code Reviews:", options: { bold: true } }, { text: " Team walkthroughs and formal code inspections led by moderators.", options: { breakLine: true } },
      { text: "• High Fault Detection Rate:", options: { bold: true, color: DGN } }, { text: " Inspections typically detect 60% to 90% of code defects, finding root errors (faults) directly in the source.", options: { breakLine: true } },
      { text: "• Static Analyzers:", options: { bold: true } }, { text: " Automated tools scan code for logic errors, dead blocks, and compliance warnings.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LMR, M);
    s.addText("Software Testing (Dynamic Validation)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Execution-Based:", options: { bold: true } }, { text: " Evaluates system behavior by executing the code with test cases and inputs.", options: { breakLine: true } },
      { text: "• Failure Observation:", options: { bold: true, color: M } }, { text: " Observes system failures (deviations from expectations), which must be debugged to locate the root fault.", options: { breakLine: true } },
      { text: "• Testing Levels:", options: { bold: true } }, { text: " Ranges from Unit Testing (components) to Integration Testing (modules) and System Testing (entire application).", options: { breakLine: true } },
      { text: "• Limits:", options: { bold: true } }, { text: " 'Testing shows the presence of bugs, not their absence' (Dijkstra).", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("When bugs do enter the code, we rely on Defect Reduction, which combines inspections and testing. Inspections are static, non-execution reviews of design, schemas, and code. Led by moderators or done via peer code reviews, inspections are highly efficient, catching sixty to ninety percent of defects directly at their source. Testing, on the other hand, is dynamic. We execute the code with test cases to observe system failures. Failures represent deviations from expectations, which developers must then debug to locate and repair the root fault. Testing spans from unit testing components to integration and system testing. Remember, as Dijkstra famously said, testing shows the presence of bugs, not their absence, which is why static reviews and dynamic testing must work together.");
  }

  // Slide 5 ── Defect Containment
  {
    const s = contentSlide(p, H, "Defect Containment: Production Resilience");
    s.addText("Defect containment limits the impact of bugs that escape into production:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LYL, AM);
    s.addText("Fault Tolerance (Software Resilience)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• System Continuity:", options: { bold: true } }, { text: " The software detects runtime errors and handles them without crashing the entire system.", options: { breakLine: true } },
      { text: "• N-Version Programming:", options: { bold: true, color: BL } }, { text: " Executes multiple versions of the same module in parallel. A voting system compares results to select the correct output, bypassing individual code bugs.", options: { breakLine: true } },
      { text: "• Recovery Blocks:", options: { bold: true } }, { text: " Tries to run a primary module. If the check fails, it rolls back state and executes an alternate backup module.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGR, GY);
    s.addText("Safety Assurance & Containment",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: GY, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Safety Checks:", options: { bold: true } }, { text: " Code monitors parameters to detect dangerous conditions (e.g. medical devices checking dose values).", options: { breakLine: true } },
      { text: "• Fail-Safe Design:", options: { bold: true, color: M } }, { text: " If a critical failure occurs, the system defaults to a safe state (e.g., shutting down power to prevent damage).", options: { breakLine: true } },
      { text: "• Containment Boundaries:", options: { bold: true } }, { text: " Isolates subsystems so a crash in a non-critical component (e.g., user profile UI) cannot crash the primary transaction module.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("If a defect escapes all checks and reaches production, Defect Containment is our final line of defense. It prevents a local bug from causing a complete system disaster. First, we use Fault Tolerance to detect errors at runtime and handle them without crashing. Methods include N-Version Programming, which runs multiple versions of a module in parallel and uses a voting scheme to select the correct output, and Recovery Blocks, which executes an alternate backup module if the primary module's sanity check fails. Second, we use Safety Assurance and fail-safe designs. For example, medical devices or industrial controllers continuously monitor parameters and shut down safely if a critical failure is detected. We also use containment boundaries to isolate crashes, ensuring a failure in a minor module like a profile UI doesn't crash a core transaction module.");
  }

  // Slide 6 ── Applications of QA Pillars
  {
    const s = contentSlide(p, H, "Applications of QA Strategies");
    s.addText("How modern dev teams implement the three pillars of QA:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "⚙️  CI/CD Automated Pipelines",
        b: "Linting checks run on code commit (Prevention). Automated unit tests run on compilation, and integration tests run on build packaging (Reduction).", bg: LGN, bd: GN
      },
      {
        t: "👥  Rigorous Pull Request Code Reviews",
        b: "Code reviews and checklists verify that every line of code is inspected by colleagues before merge. Eliminates logic bugs and improves quality (Reduction).", bg: LBL, bd: BL
      },
      {
        t: "🌐  High-Availability Server Clusters",
        b: "Database replicas and load balancers detect hardware or software errors and automatically redirect user traffic, preserving availability (Containment).", bg: LMR, bd: M
      },
      {
        t: "🏥  Embedded Control Systems",
        b: "Medical, automotive, and avionics software run concurrent redundant processors to vote on mechanical decisions, preventing accidents (Containment).", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let's see how these QA strategies are applied in the industry. Modern teams build automated linting and code scanning directly into CI/CD pipelines to prevent syntax errors. They enforce rigorous pull request reviews where every line is inspected by colleagues before merging, implementing defect reduction. For high-availability cloud servers, database replicas and load balancers detect errors and redirect traffic, acting as containment. Finally, safety-critical systems in medical devices, automotive control, and avionics run concurrent redundant processors that vote on decisions to ensure physical safety under containment. Combining these three pillars guarantees highly resilient and high-quality software.");
  }

  const s7 = thankYouSlide(p, H);
  s7.addNotes("That concludes my presentation on Defect Prevention, Reduction, and Containment. By structuring our QA strategy across these three layers, we minimize costs, maximize quality, and ensure production resilience. I want to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management for this opportunity. Thank you all for your attention.");

  await p.writeFile({ fileName: "./outputs/Sem1_SQE_Unit2_QA_Dealing_Defects.pptx" });
  convertToPdf("./outputs/Sem1_SQE_Unit2_QA_Dealing_Defects.pptx");
  console.log("✔  SQE Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – QUALITY PLANNING & GOAL SETTING (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Software Quality Engineering";

  const s1 = titleSlide(p, H, "Unit III: Quality Engineering", "Quality Planning & Goal Setting");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. In this presentation, I am going to talk about Quality Planning and Goal Setting in Software Quality Engineering. We will explore the entire Quality Engineering Process, study the famous Goal-Question-Metric paradigm developed by Victor Basili, look at how we perform Quality Assessment and Continuous Quality Improvement, and finally, look at real-world applications of quality planning in Agile teams, medical device verification, cloud microservices, and legacy refactoring projects.");

  // Slide 2 ── Software Quality Engineering Process
  {
    const s = contentSlide(p, H, "The Quality Engineering Process");
    s.addText([
      { text: "Quality Engineering ", options: { bold: true } },
      { text: "is a disciplined approach that integrates quality goals, processes, planning, and evaluation throughout the software development life cycle.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Key Quality Engineering Activities",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Quality Planning:", options: { bold: true, color: DGN } }, { text: " Setting target quality goals and forming a validation strategy for the project.", options: { breakLine: true } },
      { text: "• Quality Assessment:", options: { bold: true } }, { text: " Measuring software characteristics against benchmarks during development.", options: { breakLine: true } },
      { text: "• Quality Improvement:", options: { bold: true } }, { text: " Feedback loops to tune the development process, preventing recurring code defects.", options: { breakLine: true } },
      { text: "• Process Integration:", options: { bold: true } }, { text: " Ensuring quality tasks are built into milestones (requirements, architecture, code, test).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Quality Planning & Strategy",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Defining Objectives:", options: { bold: true } }, { text: " Agreeing on target benchmarks (e.g. max 0.5 bugs/KLOC, 99.9% availability).", options: { breakLine: true } },
      { text: "• Selecting QA Methods:", options: { bold: true } }, { text: " Deciding where to use inspections vs dynamic unit testing or load testing.", options: { breakLine: true } },
      { text: "• Resource Budgeting:", options: { bold: true } }, { text: " Allocating developers, QA testers, and hardware environments for testing.", options: { breakLine: true } },
      { text: "• Risk Assessment:", options: { bold: true, color: M } }, { text: " Identifying critical modules needing rigorous verification (e.g. payment APIs).", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Quality Engineering is a disciplined approach that integrates quality goals, processes, planning, and evaluation throughout the software development life cycle. Rather than just finding bugs at the end, it focuses on three core activities: Quality Planning, which sets target goals and shapes validation strategies; Quality Assessment, which measures software characteristics against benchmarks during active development; and Quality Improvement, which uses feedback loops to optimize processes. During planning, we define objective benchmarks, select appropriate QA verification methods like inspections versus dynamic load testing, allocate developer and hardware testing resources, and conduct risk assessments to pinpoint critical modules that require intensive verification.");
  }

  // Slide 3 ── The GQM (Goal-Question-Metric) Paradigm
  {
    const s = contentSlide(p, H, "Goal-Question-Metric (GQM) Paradigm");
    s.addText("GQM is a structured framework developed by Basili to establish measurable software goals:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Hierarchical GQM Structure",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Goal (Conceptual Level):", options: { bold: true, color: M } }, { text: " State the objective for a project, process, or product. Must define object, purpose, focus, viewpoint, and context.\n", options: { breakLine: true } },
      { text: "2. Question (Operational Level):", options: { bold: true, color: BL } }, { text: " Formulate questions that determine if the goal is being met during development.\n", options: { breakLine: true } },
      { text: "3. Metric (Quantitative Level):", options: { bold: true, color: DGN } }, { text: " Define specific, numeric metrics to answer each question objectively.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("GQM Example in Practice",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Goal:", options: { bold: true, color: M } }, { text: " Improve the reliability of the billing API from the developer's viewpoint.\n", options: { breakLine: true } },
      { text: "• Questions to ask:\n", options: { bold: true, color: BL } },
      { text: "  - Q1: What is the current crash frequency?\n", options: { color: BK } },
      { text: "  - Q2: Are exception blocks trapping failures?\n\n", options: { color: BK } },
      { text: "• Metrics defined:\n", options: { bold: true, color: DGN } },
      { text: "  - M1: Mean Time Between Failures (MTBF) in hours.\n", options: { color: BK } },
      { text: "  - M2: Percentage of database errors logged vs. uncaught runtime exceptions.", options: { color: BK } }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("To measure quality, we need structured goals. The Goal-Question-Metric, or GQM, paradigm was developed by Victor Basili to establish measurable objectives. It works top-down across three levels. At the conceptual level is the Goal, defining the object of study, its purpose, quality focus, viewpoint, and context. At the operational level are the Questions, which check whether the goal is being met. At the quantitative level are the Metrics, providing numeric data to answer the questions. For example, if our goal is to improve billing API reliability, our questions might ask about crash frequency and exception trapping. Our selected metrics would track Mean Time Between Failures in hours and the percentage of logged exceptions, ensuring we only gather useful, actionable data.");
  }

  // Slide 4 ── Quality Assessment & Improvement
  {
    const s = contentSlide(p, H, "Quality Assessment & Improvement");
    s.addText("How companies measure quality during development and perform improvements:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Quality Assessment (Measurement)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Metrics Collection:", options: { bold: true } }, { text: " Collecting quantitative data on defect density, test coverage, and compiler warnings during the project.", options: { breakLine: true } },
      { text: "• Benchmarking:", options: { bold: true, color: BL } }, { text: " Comparing metrics against standards (e.g., requiring 80% code coverage before a pull request can be merged).", options: { breakLine: true } },
      { text: "• Quality Audits:", options: { bold: true } }, { text: " Periodic reviews by QA leads to verify code matches design guidelines.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Continuous Quality Improvement",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Root Cause Analysis (RCA):", options: { bold: true } }, { text: " When a severe defect escapes to production, the team conducts an RCA to identify the underlying process failure.", options: { breakLine: true } },
      { text: "• Retrospectives:", options: { bold: true, color: DGN } }, { text: " Post-sprint review sessions where developers identify bottlenecks and update quality checklists.", options: { breakLine: true } },
      { text: "• Process Tuning:", options: { bold: true } }, { text: " Modifying design constraints, compiler rules, or unit test targets based on RCA feedback, preventing repeat bugs.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Quality assessment and improvement close the engineering feedback loop. In Quality Assessment, we collect quantitative data during development, covering metrics like defect density, code coverage, and compiler warnings. We benchmark these values against project thresholds, such as requiring eighty percent code coverage before a pull request can be merged, and conduct periodic audits to ensure compliance. If a severe bug slips into production, we perform Continuous Quality Improvement. We conduct a Root Cause Analysis to identify the process gap, run retrospectives to gather team feedback, and tune our development process by updating design rules, compiler rules, or unit testing targets to prevent the bug from recurring.");
  }

  // Slide 5 ── Applications of Quality Planning
  {
    const s = contentSlide(p, H, "Applications of Quality Planning");
    s.addText("Software teams implement GQM and quality plans across various development domains:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🏃  Agile Sprint Planning",
        b: "Teams establish Definition of Done (DoD) goals at sprint start. E.g., 'A story is Done only if it has 80% test coverage and passes security scans (DoD Goal).'", bg: LGN, bd: GN
      },
      {
        t: "🏥  Medical Device Verification",
        b: "Quality plans detail the specific test runs, compliance audits, and safety-critical thresholds required to meet FDA validation criteria.", bg: LBL, bd: BL
      },
      {
        t: "🌐  Cloud Microservices Scaling",
        b: "Applies GQM to monitor system latency. Goal: Optimize query speeds. Metrics tracked: 99th percentile response time, CPU utilization under load.", bg: LMR, bd: M
      },
      {
        t: "🔄  Legacy Refactoring Projects",
        b: "Establishes quality plans to minimize regression bugs during migrations. Defines strict validation thresholds prior to production deployment.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let's look at how quality planning is applied. In Agile Sprint Planning, teams establish quality goals inside their Definition of Done, such as requiring automated tests to pass before merging. In Medical Device Verification, quality plans detail the exact test cases and compliance audits needed to meet FDA validation criteria. In Cloud Microservices, teams apply GQM to monitor scalability, defining goals like optimizing query speeds and tracking ninety-ninth percentile response times and CPU usage. Finally, in Legacy Refactoring, quality plans establish strict validation thresholds to minimize regression bugs during migrations. This ensures quality is treated as a core business driver across all projects.");
  }

  const s6 = thankYouSlide(p, H);
  s6.addNotes("That concludes my presentation on Quality Planning and Goal Setting. By using systematic models like GQM, software teams align their quality metrics directly with business and user needs, driving continuous improvement. I would like to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you for your time.");

  await p.writeFile({ fileName: "./outputs/Sem1_SQE_Unit3_Quality_Planning.pptx" });
  convertToPdf("./outputs/Sem1_SQE_Unit3_Quality_Planning.pptx");
  console.log("✔  SQE Unit 3 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 4 – SOFTWARE TEST AUTOMATION (Unit IV)
// ═══════════════════════════════════════════════════════════════
async function ppt4() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Software Quality Engineering";

  const s1 = titleSlide(p, H, "Unit IV: Test Activities & Automation", "Software Test Automation");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. In this presentation, I am going to talk about Software Test Automation under Software Quality Engineering. We will go through the core testing activities and how they are managed, discuss what test automation is and establish clear guidelines on when to automate versus when to keep tests manual, analyze modular test framework architecture and common challenges, and review practical applications of test automation in modern CI/CD pipelines, browser UI testing, load testing, and API verification.");

  // Slide 2 ── Test Activities & Management
  {
    const s = contentSlide(p, H, "Test Activities & Management");
    s.addText([
      { text: "Testing ", options: { bold: true } },
      { text: "is a structured process to verify code correctness. It contains multiple sequential activities that must be managed, tracked, and reported systematically.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Testing Activities",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Test Planning:", options: { bold: true, color: BL } }, { text: " Formulating the test strategy, boundaries, environments, and goals.", options: { breakLine: true } },
      { text: "2. Test Preparation:", options: { bold: true } }, { text: " Designing specific test cases, mapping parameters, and setting up test data.", options: { breakLine: true } },
      { text: "3. Test Execution:", options: { bold: true } }, { text: " Executing the test cases on build deployments.", options: { breakLine: true } },
      { text: "4. Result Checking:", options: { bold: true, color: DGN } }, { text: " Comparing actual execution outputs against expected results to detect mismatches.", options: { breakLine: true } },
      { text: "5. Analysis & Follow-up:", options: { bold: true } }, { text: " Logging bugs in trackers, debugging code, and re-testing fixes.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Managing the Testing Process",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Role Assignments:", options: { bold: true } }, { text: " Defining clear responsibilities for test leads, test engineers, and developers.", options: { breakLine: true } },
      { text: "• Traceability Matrix:", options: { bold: true, color: M } }, { text: " Mapping test cases to requirement IDs to ensure all features are fully verified.", options: { breakLine: true } },
      { text: "• Environment Management:", options: { bold: true } }, { text: " Provisioning staging environments that match production configurations.", options: { breakLine: true } },
      { text: "• Exit Criteria:", options: { bold: true } }, { text: " Defining clear thresholds (e.g. zero blocker bugs, 90% test pass rate) to release software.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's first look at core testing activities and how they are managed. Software testing is a structured process to verify code correctness, involving five sequential activities: Test Planning to define scope and strategies; Test Preparation to design test cases and set up mock data; Test Execution to run cases on build deployments; Result Checking to compare actual outputs with expected results; and Analysis and Follow-up to log bugs and verify fixes. Managing this process requires clear role assignments, using a Traceability Matrix to map test cases directly to requirements, provisioning staging environments that mirror production, and defining strict release exit criteria like zero blocker bugs.");
  }

  // Slide 3 ── What is Test Automation?
  {
    const s = contentSlide(p, H, "What is Test Automation?");
    s.addText([
      { text: "Test Automation ", options: { bold: true } },
      { text: "uses specialized software tools to execute tests, compare actual outcomes against expected results, and report findings automatically, replacing manual test tasks.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("When to Automate (Best ROI)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Regression Testing:", options: { bold: true, color: DGN } }, { text: " Running past tests repeatedly on new builds to check for regression bugs.", options: { breakLine: true } },
      { text: "• Repetitive & High-Volume Tests:", options: { bold: true } }, { text: " Tests running frequently with multiple data variations (data-driven tests).", options: { breakLine: true } },
      { text: "• Performance & Load Tests:", options: { bold: true } }, { text: " Simulating thousands of simultaneous users accessing database connections.", options: { breakLine: true } },
      { text: "• Smoke Tests:", options: { bold: true } }, { text: " Basic validation checks run immediately on every build commit.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("When NOT to Automate (Keep Manual)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Ad-Hoc / Exploratory Tests:", options: { bold: true, color: M } }, { text: " Dynamic, non-scripted tests looking for edge cases based on tester intuition.", options: { breakLine: true } },
      { text: "• Unstable / Evolving Features:", options: { bold: true } }, { text: " Features undergoing frequent design edits. Constant test script adjustments waste time.", options: { breakLine: true } },
      { text: "• One-Off Tests:", options: { bold: true } }, { text: " Tests run only once or twice. Script development cost exceeds manual execution time.", options: { breakLine: true } },
      { text: "• UX & Visual Design Checks:", options: { bold: true } }, { text: " Assessing visual aesthetics, layout alignment, and user friendliness.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("What is Test Automation? It is the use of specialized software tools to execute tests, compare actual outcomes against expected results, and report findings automatically. However, we should only automate when it yields a high return on investment. The best candidates are regression testing to check for bugs in new builds, high-volume repetitive data-driven tests, performance and load tests, and smoke tests run on every commit. We should keep tests manual for exploratory testing based on tester intuition, unstable evolving features undergoing frequent changes, one-off test runs where setup cost exceeds manual execution time, and user experience and visual design checks.");
  }

  // Slide 4 ── Test Automation Architecture & Challenges
  {
    const s = contentSlide(p, H, "Automation Architecture & Challenges");
    s.addText("How automated test frameworks are structured, and their common implementation challenges:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Modular Test Framework Architecture",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Test Runner:", options: { bold: true } }, { text: " Orchestrates test runs, maps files, and handles execution sequences.", options: { breakLine: true } },
      { text: "• Assertion Library:", options: { bold: true, color: BL } }, { text: " Compares actual variables against expected states, reporting passes or failures (e.g. expect(val).toBe(true)).", options: { breakLine: true } },
      { text: "• Page Object Model (POM):", options: { bold: true } }, { text: " Design pattern representing web pages as classes. Isolates UI change impacts.", options: { breakLine: true } },
      { text: "• Test Data Mocking:", options: { bold: true } }, { text: " Simulates databases or external APIs using mock services, keeping tests isolated.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Common Automation Challenges",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Brittle Test Scripts:", options: { bold: true, color: M } }, { text: " Minor UI changes (like changing a button ID) break automated tests, causing high maintenance overhead.", options: { breakLine: true } },
      { text: "• Test Data Staleness:", options: { bold: true } }, { text: " Managing database states so tests run with clean, predictable data consistently.", options: { breakLine: true } },
      { text: "• High Initial Setup Cost:", options: { bold: true } }, { text: " Initial framework setup and script development take substantial developer hours.", options: { breakLine: true } },
      { text: "• False Positives / Flakiness:", options: { bold: true } }, { text: " Tests failing due to network latency, timing issues, or environment glitches rather than actual bugs.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Automated test frameworks are structured modularly. They include a Test Runner to execute sequences, Assertion Libraries to evaluate pass or fail states, a Page Object Model to represent UI pages as classes and isolate UI changes, and API Mocking to keep test runs independent of live databases. However, teams face significant challenges. These include brittle test scripts that break on minor UI changes, managing stale database states across runs, high initial setup and development costs, and flaky tests that fail due to network lag rather than actual software defects. Resolving these requires choosing robust locators and designing clean, isolated data patterns.");
  }

  // Slide 5 ── Applications of Test Automation
  {
    const s = contentSlide(p, H, "Applications of Test Automation");
    s.addText("Where test automation is implemented in professional software delivery pipelines:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "⚡  CI/CD Integration",
        b: "Integrates automated test suites into code pipelines. Every pull request automatically runs unit and regression tests before it can be merged.", bg: LGN, bd: GN
      },
      {
        t: "🌐  Automated UI Testing",
        b: "Uses tools like Selenium or Playwright to simulate user clicks, form submissions, and check page flows across multiple web browsers.", bg: LBL, bd: BL
      },
      {
        t: "🏎️  Load & Stress Testing",
        b: "Uses tools like JMeter or k6 to simulate thousands of concurrent requests, measuring server latency, memory usage, and throughput limits.", bg: LMR, bd: M
      },
      {
        t: "📞  API Verification Testing",
        b: "Automates REST/SOAP endpoint verification. Sends JSON payloads, validates status codes, and checks database state changes.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("In practice, test automation is applied across four primary domains. First, in CI/CD pipelines, automated tests run automatically on every pull request before merging. Second, in UI testing, tools like Selenium or Playwright simulate browser actions to verify user flows. Third, in Load and Stress testing, tools like JMeter or k6 simulate thousands of concurrent requests to measure server latency and database throughput. Fourth, in API testing, automated scripts send JSON payloads to REST endpoints, validating responses and database states. Automating these areas allows software teams to test builds continuously and deliver reliable updates faster.");
  }

  const s6 = thankYouSlide(p, H);
  s6.addNotes("That brings me to the end of my presentation on Software Test Automation. We have seen how structuring test activities and selectively automating them helps teams deliver reliable, high-performance software at scale. I want to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you for your time.");

  await p.writeFile({ fileName: "./outputs/Sem1_SQE_Unit4_Test_Automation.pptx" });
  convertToPdf("./outputs/Sem1_SQE_Unit4_Test_Automation.pptx");
  console.log("✔  SQE Unit 4 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 5 – USAGE-BASED STATISTICAL TESTING (Unit V)
// ═══════════════════════════════════════════════════════════════
async function ppt5() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Software Quality Engineering";

  const s1 = titleSlide(p, H, "Unit V: Checklist & Usage Testing", "Usage-Based Statistical Testing");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. In this presentation, I am going to talk about Usage-Based Statistical Testing in Software Quality Engineering. We will review the core concept of testing based on user-centric operational profiles rather than code coverage, analyze limitations of traditional checklist-based testing, study John Musa's five-step model for building operational profiles, go over a cartridge support software case study, and look at practical applications of usage-based testing across cloud systems, medical devices, automotive controllers, and mobile apps.");

  // Slide 2 ── What is Usage-Based Testing?
  {
    const s = contentSlide(p, H, "Usage-Based Statistical Testing");
    s.addText([
      { text: "Usage-Based Statistical Testing ", options: { bold: true } },
      { text: "focuses testing resources on software execution paths based on how users run the system in production, rather than focusing on structural code coverage.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Concept & Theory",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• User-Centric View:", options: { bold: true } }, { text: " Traditional testing checks code branches, but users only run specific paths. Usage-based testing prioritizes paths users visit most.", options: { breakLine: true } },
      { text: "• Operational Profile (OP):", options: { bold: true, color: DGN } }, { text: " A probability distribution mapping user operations to their run frequencies.", options: { breakLine: true } },
      { text: "• Reliability Estimation:", options: { bold: true } }, { text: " Testing based on the OP provides a realistic estimate of the system's Mean Time to Failure (MTTF) in production.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Checklist Limitations",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Checklist Testing:", options: { bold: true } }, { text: " Uses static checklists of common bugs to audit features manually.", options: { breakLine: true } },
      { text: "• Limitations:", options: { bold: true, color: M } }, { text: " Checklists can become outdated, suffer from bias, and do not scale. They fail to test dynamic data interactions and edge cases.", options: { breakLine: true } },
      { text: "• Why Use OP instead:", options: { bold: true } }, { text: " Checklists test what testers remember. Operational Profiles test what users actually do, ensuring high reliability where it matters most.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's begin with the core theory of usage-based statistical testing. Unlike traditional structural testing that checks code branches, usage-based testing focuses on how users execute the system in production. It prioritizes testing paths that users visit most frequently, which are mapped using an Operational Profile. This probability distribution maps user operations to their frequency of use. Testing against this profile provides a realistic estimate of the system's Mean Time to Failure. In contrast, traditional checklist testing uses static bug lists to audit features. Checklist testing has limitations: it easily becomes outdated, suffers from tester bias, does not scale, and fails to catch dynamic edge cases. Using operational profiles ensures we test what users actually do, improving quality where it matters most.");
  }

  // Slide 3 ── Musa's Operational Profiles
  {
    const s = contentSlide(p, H, "Musa's Operational Profiles");
    s.addText("John Musa developed a structured, step-by-step model to build operational profiles:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Five Steps to Build an OP",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Find Customer Profile:", options: { bold: true } }, { text: " Identify user segments (e.g. admin, guest, subscriber).", options: { breakLine: true } },
      { text: "2. Find User Profile:", options: { bold: true } }, { text: " Identify distinct systems or components users access.", options: { breakLine: true } },
      { text: "3. Find System Modes:", options: { bold: true } }, { text: " Identify system states (e.g. startup, normal, peak load).", options: { breakLine: true } },
      { text: "4. Find Functional Profile:", options: { bold: true } }, { text: " List functions the software performs.", options: { breakLine: true } },
      { text: "5. Find Operational Profile:", options: { bold: true, color: DGN } }, { text: " Group functions into specific user operations. Assign probabilities based on log telemetry or marketing data.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Profile Representation Example",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });

    const tData = [
      [
        { text: "User Operation", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "Count / hr", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "Probability", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } }
      ],
      [
        { text: "Search Products", options: { fill: { color: LGR } } }, { text: "6,000", options: { fill: { color: LGR }, align: "center" } },
        { text: "0.60  (60%)", options: { fill: { color: LGR }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Add to Cart", options: { fill: { color: LBL } } }, { text: "2,500", options: { fill: { color: LBL }, align: "center" } },
        { text: "0.25  (25%)", options: { fill: { color: LBL }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Checkout Payment", options: { fill: { color: LGN } } }, { text: "1,000", options: { fill: { color: LGN }, align: "center" } },
        { text: "0.10  (10%)", options: { fill: { color: LGN }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Update Profile", options: { fill: { color: LMR } } }, { text: "500", options: { fill: { color: LMR }, align: "center" } },
        { text: "0.05  (5%)", options: { fill: { color: LMR }, bold: true, color: GY, align: "center" } }
      ]
    ];
    s.addTable(tData,
      {
        x: 6.72, y: 2.82, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.5, 1.5, 2.03]
      });

    s.addNotes("To build these profiles, John Musa developed a structured five-step model. First, we identify customer profiles, segmenting users into groups like guest or admin. Second, we define user profiles, mapping the distinct system modules they access. Third, we establish system modes, identifying states like normal operations or peak loads. Fourth, we outline functional profiles to list all software functions. Finally, we build the operational profile, grouping functions into specific operations and assigning probabilities using log telemetry. For example, in an e-commerce system, query logs might reveal that searching products represents sixty percent of user operations, cart additions twenty-five percent, payments ten percent, and profile updates only five percent, directing our primary testing focus.");
  }

  // Slide 4 ── Case Study: Cartridge Support Software
  {
    const s = contentSlide(p, H, "Case Study: Cartridge Support Software");
    s.addText("Applying Musa's operational profiling to printer driver utility software:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGR, GY);
    s.addText("Project Overview & Goals",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Target System:", options: { bold: true } }, { text: " PC utility software that manages ink cartridge statuses, prints nozzle test patterns, and calibrates printer alignment.", options: { breakLine: true } },
      { text: "• Context & Problem:", options: { bold: true } }, { text: " Limited verification window before software release. Standard code coverage checks were taking too long.", options: { breakLine: true } },
      { text: "• Objective:", options: { bold: true, color: BL } }, { text: " Model a 4-step functional profile and assign probabilities based on printer telemetry logs, targeting testing on the most frequent print paths.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Cartridge Software Operational Profile",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });

    const csData = [
      [
        { text: "Operation Class", options: { bold: true, fill: { color: BL }, color: WH, align: "center" } },
        { text: "Telemetry Probability", options: { bold: true, fill: { color: BL }, color: WH, align: "center" } },
        { text: "Test Effort Allocation", options: { bold: true, fill: { color: BL }, color: WH, align: "center" } }
      ],
      [
        { text: "Query Ink Levels", options: { fill: { color: LGR } } }, { text: "0.75  (75%)", options: { fill: { color: LGR }, align: "center" } },
        { text: "High  (Verify UI/cache)", options: { fill: { color: LGR }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Print Test Pattern", options: { fill: { color: LBL } } }, { text: "0.15  (15%)", options: { fill: { color: LBL }, align: "center" } },
        { text: "Medium  (Check drivers)", options: { fill: { color: LBL }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Nozzle Alignment", options: { fill: { color: LGN } } }, { text: "0.08  (8%)", options: { fill: { color: LGN }, align: "center" } },
        { text: "Low  (Basic validation)", options: { fill: { color: LGN }, bold: true, color: AM, align: "center" } }
      ],
      [
        { text: "Firmware Upgrade", options: { fill: { color: LMR } } }, { text: "0.02  (2%)", options: { fill: { color: LMR }, align: "center" } },
        { text: "Critical  (Risk focused)", options: { fill: { color: LMR }, bold: true, color: M, align: "center" } }
      ]
    ];
    s.addTable(csData,
      {
        x: 6.72, y: 2.82, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.3, 1.8, 1.93]
      });

    s.addNotes("Let's examine a case study applying Musa's model to cartridge support utility software. The software checks ink levels, prints nozzle test patterns, and calibrates alignment. With a limited testing window before release, standard code coverage checks were too slow. The team analyzed printer logs to build an operational profile. The data showed that querying ink levels made up seventy-five percent of operations, print testing fifteen percent, nozzle alignment eight percent, and firmware upgrades only two percent. Testing efforts were allocated accordingly, prioritizing the high-use ink query UI. However, for firmware upgrades, despite a two percent probability, the team allocated critical testing resources because a firmware failure would brick the printer. This shows that we must combine usage probability with business risk to distribute testing efforts.");
  }

  // Slide 5 ── Applications of Usage-Based Testing
  {
    const s = contentSlide(p, H, "Applications of Usage-Based Testing");
    s.addText("How software engineering teams utilize usage testing and profiles in practice:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📊  Telemetry-Driven Performance Tests",
        b: "Web apps analyze production traffic logs to build operational profiles. Load tests replicate this profile distribution (e.g. simulating 80% reads and 20% writes) to verify database performance under realistic loads.", bg: LGN, bd: GN
      },
      {
        t: "🩺  Medical Device Reliability Models",
        b: "Infusion pumps and pacemaker test runs are designed based on patient usage modes (operational profiles), enabling accurate Mean Time to Failure reliability estimates for FDA approval.", bg: LBL, bd: BL
      },
      {
        t: "🚗  Automotive Firmware Verification",
        b: "Engine and ADAS controllers are tested using simulated driving profiles (city, highway, extreme conditions) to ensure safety in typical user environments.", bg: LMR, bd: M
      },
      {
        t: "📲  Mobile App UI Layout Tests",
        b: "App analytics track user clicks. Teams focus manual UI testing on highly visited pages, keeping lesser-visited settings pages on automated schedules.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("In practice, usage testing and operational profiles are utilized across several domains. For web applications, teams analyze production traffic logs to run telemetry-driven performance tests, simulating eighty percent reads and twenty percent writes to test databases. In medical engineering, infusion pumps and pacemakers undergo test runs based on clinical operational profiles to estimate Mean Time to Failure for FDA approval. In automotive systems, engine and ADAS controllers are tested against simulated city and highway driving profiles to ensure safety. For mobile apps, UI testing focuses on highly visited pages tracked via analytics, ensuring the most common user paths are bug-free.");
  }

  const s6 = thankYouSlide(p, H);
  s6.addNotes("That concludes my presentation on Usage-Based Statistical Testing. By prioritizing software paths based on real user behavior and combining probability with risk, organizations achieve high reliability and maximize the return on their testing investments. I want to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you for your time.");

  await p.writeFile({ fileName: "./outputs/Sem1_SQE_Unit5_Usage_Testing.pptx" });
  convertToPdf("./outputs/Sem1_SQE_Unit5_Usage_Testing.pptx");
  console.log("✔  SQE Unit 5 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 5 SQE PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  await ppt4();
  await ppt5();
  console.log("All 5 SQE PPTs generated successfully!");
})();
