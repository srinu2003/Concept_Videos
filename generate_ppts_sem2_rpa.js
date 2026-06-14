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
    "YEAR :   I Year  II Semester",
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
//  PPT 1 – RPA USE CASES & BOT CREATION  (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Robotic Process Automation";

  const tSl = titleSlide(p, H, "Unit I: Introduction & Bot Creation", "RPA Use Cases & Bot Creation");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and I am presenting today on the topic of Robotic Process Automation Use Cases and Bot Creation as part of my M.Tech in Computer Science and Engineering. Today, we'll explore the foundational concepts of RPA, look at the primary ways software bots are constructed, and discuss how enterprises deploy them to streamline business workflows.");

  // Slide 2 ── What is RPA?
  {
    const s = contentSlide(p, H, "What is Robotic Process Automation?");
    s.addText([
      { text: "Robotic Process Automation (RPA) ", options: { bold: true } },
      { text: "is a technology that uses software robots ('bots') to automate repetitive, rules-based digital tasks by mimicking human interactions with user interfaces.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Concepts & Benefits",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Mimics Humans:", options: { bold: true } }, { text: " Interacts with applications, clicks buttons, types text, and copies files just like a human operator.", options: { breakLine: true } },
      { text: "• Rules-Based:", options: { bold: true } }, { text: " Follows predefined execution paths (if-else logic, loops) without needing human decision-making.", options: { breakLine: true } },
      { text: "• Zero Integration Cost:", options: { bold: true, color: DGN } }, { text: " Works on the user interface layer (UI), requiring no modifications to legacy backend databases or APIs.", options: { breakLine: true } },
      { text: "• Speed and Accuracy:", options: { bold: true } }, { text: " Executes tasks 24/7 with zero typographical errors, minimizing business costs.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Primary Industry Use Cases",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Finance & Accounting:", options: { bold: true, color: BL } }, { text: " Invoice processing, bank reconciliations, tax reports extraction, account audits.", options: { breakLine: true } },
      { text: "• Human Resources (HR):", options: { bold: true } }, { text: " Employee onboarding data synchronization, payroll processing, holiday tracking updates.", options: { breakLine: true } },
      { text: "• Customer Service:", options: { bold: true } }, { text: " Directing customer tickets, processing basic account queries automatically, pulling chat histories.", options: { breakLine: true } },
      { text: "• Data Operations:", options: { bold: true } }, { text: " Moving data between legacy green-screen terminals and modern cloud databases (data entry).", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by understanding the foundational concept of Robotic Process Automation, or RPA. RPA is a technology that leverages software robots, or bots, to automate highly repetitive, rules-based digital tasks. These bots mimic the exact actions of a human worker, such as clicking buttons, logging into applications, copying files, and performing data entry. A key advantage of RPA is that it operates directly at the user interface layer, meaning it has a zero integration cost. We do not need to rebuild legacy backend databases or program complex APIs; the bot simply interacts with existing software applications. By executing rules-based tasks around the clock with absolute accuracy, RPA eliminates human typographical errors, dramatically speeds up transactions, and significantly lowers operational costs. Industry use cases span finance for invoice processing, HR for payroll updates, customer service for ticket routing, and data operations for system migrations.");
  }

  // Slide 3 ── Ways to Create Bots
  {
    const s = contentSlide(p, H, "Ways to Create Bots");
    s.addText("RPA platforms (like Automation Anywhere) offer three primary methods to construct bots:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    const ways = [
      {
        t: "1. Screen Recording",
        b: "• Action: Logs keyboard inputs and mouse coordinates as you perform a task.\n• Pros: Instantly generates bot files; needs no code.\n• Cons: Brittle. Changing button positions or resolution breaks the bot.", bg: LGN, bd: GN
      },
      {
        t: "2. Object Cloning",
        b: "• Action: Captures actual UI element properties (HTML IDs, tag names, controls) rather than coordinates.\n• Pros: Highly robust. Bot finds button even if moved.\n• Target: Standard in web and desk apps.", bg: LBL, bd: BL
      },
      {
        t: "3. Workflow Designers",
        b: "• Action: Constructing bots by dragging and dropping command blocks (e.g. Open Browser, Query DB).\n• Pros: Easy to insert loops, conditions, and custom error handling.\n• Target: Complex business workflows.", bg: LYL, bd: AM
      }
    ];

    ways.forEach((wy, i) => {
      const bx = CX + i * 4.15;
      box(p, s, bx, 2.30, 4.00, 4.60, wy.bg, wy.bd);
      boxTitle(s, bx, 2.30, 4.00, wy.t, wy.bd);
      s.addText(wy.b,
        { x: bx + 0.12, y: 2.30 + 0.52, w: 4.00 - 0.24, h: 4.60 - 0.58, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });
    });

    s.addNotes("Moving on to how bots are actually constructed. Standard RPA platforms like Automation Anywhere offer three primary methods to create bots. The first method is Screen Recording, which logs keyboard inputs and mouse coordinates as a human performs the task. While this is fast and requires zero programming, it is highly brittle. If a button moves slightly on the screen or if screen resolution changes, the bot will click the wrong spot and fail. The second method, which is the industry standard, is Object Cloning. Instead of coordinates, Object Cloning captures the underlying properties of UI elements, such as HTML IDs, tag names, and class controls. This makes the bot highly robust, as it will locate and click the correct button even if it is moved or resized. The third method is using Workflow Designers. This provides a drag-and-drop interface where developers can build logical flows, insert conditional statements, build loops, and write custom error-handling routines, making it the preferred choice for complex enterprise workflows.");
  }

  // Slide 4 ── Applications of Bot Creation
  {
    const s = contentSlide(p, H, "Applications of Bot Creation");
    s.addText("How enterprises implement bot creation to solve daily data bottlenecks:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🧾  Automated Invoice Extraction",
        b: "A bot logs into email, downloads PDF invoices, uses OCR to extract total values, and inputs data into legacy accounting terminals.", bg: LGN, bd: GN
      },
      {
        t: "👥  HR Onboarding Automation",
        b: "Triggers when a new hire is flagged. Bot copies data from recruiter portals and sets up logins across AD, email, and payroll systems.", bg: LBL, bd: BL
      },
      {
        t: "🔄  Legacy System Data Sync",
        b: "Acts as a bridge. The bot reads data from old mainframe screens and copies it into modern Salesforce cloud tables, keeping systems in sync.", bg: LMR, bd: M
      },
      {
        t: "📊  Weekly Report Aggregator",
        b: "Bot logs into multiple portals overnight, runs search queries, downloads Excel files, aggregates values, and emails reports to managers.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let's look at some real-world applications of bot creation. In financial departments, bots automate invoice extraction by logging into emails, downloading PDFs, using Optical Character Recognition (OCR) to pull total billing amounts, and entering them directly into legacy ERP systems. For Human Resources, onboarding bots trigger when a new hire is marked, automatically creating accounts in active directory, provisioning email addresses, and updating payroll systems. In legacy migrations, bots act as data bridges, scraping records off old terminal screens and copying them into modern cloud systems like Salesforce. Finally, weekly report aggregators run overnight, querying multiple supplier portals, downloading spreadsheet files, performing calculations, and sending compiled PDF reports to management. These examples highlight how RPA bridges the gap between disconnected software systems.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("That brings me to the end of this presentation. In summary, RPA bot creation enables organizations to automate repetitive data entry, transfer information between disconnected software applications, and achieve massive efficiency gains with minimal integration overhead. Thank you so much for your time. I am now open to any questions you might have.");
  await p.writeFile({ fileName: "./outputs/Sem2_RPA_Unit1_Bot_Creation.pptx" });
  convertToPdf("./outputs/Sem2_RPA_Unit1_Bot_Creation.pptx");
  console.log("✔  RPA Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – WEB CONTROL ROOM FEATURES  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Robotic Process Automation";

  const tSl = titleSlide(p, H, "Unit II: Web Control Room & Client", "Web Control Room Features");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and I am presenting today on the Web Control Room and its features in Robotic Process Automation. In this session, we will explore the centralized server architecture of Automation Anywhere, understand the Web Control Room dashboard, and look at the critical role of the Credential Vault in securing automation credentials.");

  // Slide 2 ── What is the Web Control Room?
  {
    const s = contentSlide(p, H, "The Web Control Room");
    s.addText([
      { text: "The Web Control Room (Web CR) ", options: { bold: true } },
      { text: "is the centralized web-based management server in Automation Anywhere. It manages, deploys, schedules, and audits software bots across the entire enterprise network.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Management Features",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Centralized Deployment:", options: { bold: true, color: DGN } }, { text: " Upload bots from developer clients to the Web CR, then push them to local runtime runners.", options: { breakLine: true } },
      { text: "• Activity Tracker:", options: { bold: true } }, { text: " Monitors active runs, schedules tasks, and alerts admins if a bot encounters errors.", options: { breakLine: true } },
      { text: "• Device Management:", options: { bold: true } }, { text: " Monitors connections of development and runtime client machines.", options: { breakLine: true } },
      { text: "• Compliance & Auditing:", options: { bold: true } }, { text: " Logs all admin logins, bot executions, and configuration changes to secure databases.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Dashboard Panels (Insights)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Web CR features panels to monitor operations:", options: { breakLine: true } },
      { text: "• Home:", options: { bold: true } }, { text: " Overall summary of system health and active runs.", options: { breakLine: true } },
      { text: "• Bots:", options: { bold: true } }, { text: " Library of uploaded task files (.atmx) and credential configs.", options: { breakLine: true } },
      { text: "• Devices:", options: { bold: true } }, { text: " Connection states of bot runner clients.", options: { breakLine: true } },
      { text: "• Workload:", options: { bold: true, color: M } }, { text: " Status of transaction queues and SLA counters.", options: { breakLine: true } },
      { text: "• Audit Log:", options: { bold: true } }, { text: " Security compliance reports tracking all user activities.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("The Web Control Room, or Web CR, is the centralized web-based management server in Automation Anywhere. It acts as the brain of the entire RPA deployment, managing, scheduling, and auditing software bots across the entire enterprise network. Through the Web CR, administrators can upload bots from developer clients and push them to local runtime runners. It tracks home summaries of system health, active and scheduled runs, connected development and runtime machines, and transaction queues. Crucially, the Web CR records all activities in secure database logs for compliance, which provides an audit trail that cannot be modified. This ensures that organizations can safely deploy automation at scale while maintaining visibility and regulatory compliance.");
  }

  // Slide 3 ── Credential Vault
  {
    const s = contentSlide(p, H, "The Credential Vault");
    s.addText("The Credential Vault is a secure database that manages credentials for bots:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, GN);
    s.addText("Security & Encryption",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Secure Storage:", options: { bold: true } }, { text: " Stores passwords, API keys, and database logins. Prevents hardcoding passwords in bot files.", options: { breakLine: true } },
      { text: "• High Encryption:", options: { bold: true, color: DGN } }, { text: " Encrypts data using AES-256 standard, keeping keys secure.", options: { breakLine: true } },
      { text: "• Role-Based Access:", options: { bold: true } }, { text: " Admins control which bots and users can reference specific locker values.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Runtime Variable Binding",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Dynamic Reference:", options: { bold: true } }, { text: " Bots reference vault items as variables (e.g. $Locker.Password$).", options: { breakLine: true } },
      { text: "• Zero Plain Text Exposure:", options: { bold: true, color: M } }, { text: " The password is only decrypted in memory during the login step. Developers cannot see the raw password value.", options: { breakLine: true } },
      { text: "• Central Rotation:", options: { bold: true } }, { text: " Updating a password in the vault updates it for all bots instantly, avoiding script edits.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Security is one of the most critical aspects of enterprise automation, as bots often need to log into databases, legacy systems, and web applications. To avoid hardcoding passwords inside bot scripts, Automation Anywhere provides the Credential Vault. The Credential Vault is a secure database that encrypts sensitive information using the AES-256 standard. Administrators use role-based access to control which developers and bots can reference specific variables. At runtime, bots query these values dynamically using variables like Locker.Password. The credential is only decrypted in memory during the login action, meaning developers can build and maintain bots without ever seeing the raw passwords in plain text. This centralizes password rotation and keeps the entire automation platform secure.");
  }

  // Slide 4 ── Web CR Applications
  {
    const s = contentSlide(p, H, "Applications of Web Control Room");
    s.addText("How administrators manage bot deployments in enterprise environments:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📅  Overnight Job Scheduling",
        b: "Admins schedule invoice processing bots to run at 2:00 AM on runtime devices. Web CR automatically boots the bot and processes files.", bg: LGN, bd: GN
      },
      {
        t: "🔑  Secure Database Credential Provisioning",
        b: "Provides database access tokens to query bots via the Credential Vault. Bots run queries without developers seeing passwords.", bg: LBL, bd: BL
      },
      {
        t: "🛡️  Security Compliance Auditing",
        b: "Logs every bot run (username, device, duration, success status) to the secure audit catalog, meeting regulatory compliance requirements.", bg: LMR, bd: M
      },
      {
        t: "⚙️  API Integration",
        b: "Exposes Web CR features via REST APIs. Allows external platforms (like Salesforce or ServiceNow) to trigger bots dynamically.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let's look at the key applications of the Web Control Room in practice. First, it manages overnight job scheduling, allowing administrators to configure invoicing bots to run at 2:00 AM without human intervention. Second, it handles secure credential provisioning, providing database tokens to bots dynamically. Third, it logs all user modifications, bot uploads, and executions in the audit logs, which is essential for audits under standards like SOX or HIPAA. Finally, the Web CR exposes REST APIs, enabling external platforms like Salesforce or ServiceNow to trigger bots automatically when a customer ticket or order is created.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("To wrap up, the Web Control Room serves as the centralized command center that makes enterprise RPA secure, auditable, and scalable. By managing bots, devices, and credentials from a single web interface, administrators can ensure high availability and compliance across the platform. Thank you for your time, and I am ready to answer any questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_RPA_Unit2_Control_Room.pptx" });
  convertToPdf("./outputs/Sem2_RPA_Unit2_Control_Room.pptx");
  console.log("✔  RPA Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – DEVICE POOLS & ADMINISTRATION  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Robotic Process Automation";

  const tSl = titleSlide(p, H, "Unit III: Devices, Workloads & Administration", "RPA Device Pools & Administration");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on RPA Device Pools, Workloads, and Administration as part of the M.Tech curriculum. In this presentation, we will discuss how to cluster runtime machines for high availability, how to manage transaction queues using workloads, and how to configure roles and system settings securely in an RPA enterprise platform.");

  // Slide 2 ── Device Pools & Runtime Clients
  {
    const s = contentSlide(p, H, "Device Pools & Runtime Clients");
    s.addText([
      { text: "Device Pools ", options: { bold: true } },
      { text: "group multiple runtime client machines (bot runners) into clusters. They allow the Web Control Room to allocate tasks dynamically, providing scalability and high availability.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Managing Device Pools",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Runtime Clients:", options: { bold: true } }, { text: " Host machines that run compiled bots. Classified into unattended (automated) and attended (user-triggered).", options: { breakLine: true } },
      { text: "• Clustering:", options: { bold: true, color: DGN } }, { text: " Grouping multiple runners into a device pool. If runner A is busy, Web CR routes tasks to runner B.", options: { breakLine: true } },
      { text: "• High Availability:", options: { bold: true } }, { text: " If a runner machine crashes, the pool automatically redirects the queue to active runner nodes.", options: { breakLine: true } },
      { text: "• Dynamic Scaling:", options: { bold: true } }, { text: " Admins can add or remove runner nodes to the pool to scale capacity.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Workloads & Queue Management",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Transaction Queues:", options: { bold: true } }, { text: " Lists of data records (e.g. 5,000 invoices) to be processed by a bot.", options: { breakLine: true } },
      { text: "• Workload Distributor:", options: { bold: true, color: BL } }, { text: " Web CR reads items from the queue and distributes them across device pool runners concurrently.", options: { breakLine: true } },
      { text: "• SLA Calculator:", options: { bold: true } }, { text: " Calculates the estimated time to complete a queue based on processing speeds and node counts. Alerts admins if SLAs are at risk.", options: { breakLine: true } },
      { text: "• Queue Security:", options: { bold: true } }, { text: " Encrypts data elements inside database queues.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's look at how we manage devices and workloads in a production environment. Device Pools group multiple runtime client machines, or bot runners, into a single cluster. This allows the Web Control Room to allocate tasks dynamically across available resources. Runtime clients are categorized into unattended, which execute scheduled tasks on servers, and attended, which run on user machines. By clustering these devices, we achieve high availability: if one machine encounters a hardware crash or loses network connection, the Web Control Room automatically redirects the remaining items in the queue to active runner nodes. Furthermore, administrators can dynamically scale capacity by adding or removing machines. Workloads utilize transaction queues to store items to be processed. The SLA Calculator tracks completion times and alerts administrators if a backlog risks breaching processing schedules, ensuring consistent operational efficiency.");
  }

  // Slide 3 ── RPA Administration & Security
  {
    const s = contentSlide(p, H, "RPA Administration & Roles");
    s.addText("Administration tools secure, audit, and configure the enterprise RPA platform:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("User & Role Management",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Role-Based Access (RBAC):", options: { bold: true, color: M } }, { text: " Admins configure roles with specific permissions, enforcing security rules:\n", options: { breakLine: true } },
      { text: "   - Bot Creator:", options: { bold: true } }, { text: " Permissions to write scripts and edit local code. Cannot deploy to production.\n", options: { breakLine: true } },
      { text: "   - Bot Runner:", options: { bold: true } }, { text: " Permissions to run scripts. Cannot edit code.\n", options: { breakLine: true } },
      { text: "   - Control Room Admin:", options: { bold: true } }, { text: " Manages users, devices, licenses, and system configurations.\n", options: { breakLine: true } },
      { text: "• License Allocation:", options: { bold: true } }, { text: " Allocates developer and runtime runner licenses.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Audit Log Compliance",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Security Compliance Audits:", options: { bold: true } }, { text: " Logs all activities in the Web CR. Cannot be edited by users or admins.\n", options: { breakLine: true } },
      { text: "• Actions Tracked:", options: { bold: true, color: DGN } }, { text: "\n", options: { breakLine: true } },
      { text: "   - User modifications (role updates, password changes).\n", options: { color: BK } },
      { text: "   - Bot modifications (uploads, deletions, runs).\n", options: { color: BK } },
      { text: "   - Device changes (pool creation, connection edits).\n", options: { color: BK } },
      { text: "• Compliance value:", options: { bold: true } }, { text: " Essential for SOX, HIPAA, and GDPR database audits.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("RPA Administration provides the tools required to secure, configure, and maintain the Automation Anywhere platform. First, user management relies on Role-Based Access Control, or RBAC. Admins assign users to specific roles: Bot Creators can write and test scripts locally but cannot run them in production; Bot Runners can execute compiled bots but cannot alter code; and Control Room Admins manage licenses, users, and global configurations. This segregation of duties is a fundamental security practice. Second, the Web CR maintains an immutable Audit Log. It logs all configuration edits, password updates, and bot executions. Because this log cannot be altered by developers or administrators, it is a crucial resource for satisfying regulatory audits such as GDPR, HIPAA, and SOX, verifying that all system changes are authorized.");
  }

  // Slide 4 ── RPA Administration Applications
  {
    const s = contentSlide(p, H, "Applications of RPA Administration");
    s.addText("How organizations deploy device pooling and role segregation in production:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📈  Scaling for Seasonal Invoices",
        b: "During holiday sales, invoice volumes multiply. Admins add runtime servers to the billing device pool, allowing parallel processing and meeting SLAs.", bg: LGN, bd: GN
      },
      {
        t: "👥  Compliance Role Segregation",
        b: "Enforces security rules. Developers create bots but have no runner permissions in production. Operations staff run bots but cannot edit scripts.", bg: LBL, bd: BL
      },
      {
        t: "🩺  Auditing Financial Bot Runs",
        b: "Auditors review Web CR audit logs to confirm only authorized bots executed bank transfers, matching transaction timestamps.", bg: LMR, bd: M
      },
      {
        t: "🔄  System Migration & Updates",
        b: "RPA administrators manage migrations (e.g. importing bot packages from dev to staging schemas) and verify licensing compliance.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let's explore how these administration tools are applied in enterprise scenarios. First, device pools allow dynamic scaling. During high-volume periods like holiday sales, administrators add servers to the invoice processing pool to process records in parallel. Second, role segregation enforces compliance: developers are restricted to writing code in dev sandbox environments, while operations staff manage production runs. Third, financial auditors review the immutable audit logs to verify that only authorized bots executed bank transfers. Finally, migration tools help administrators move bot packages from dev to staging schemas safely, verifying license compliance at every stage.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, robust device pooling, workload management, and role-based administration are the keys to running a secure and scalable RPA platform in any major enterprise. These tools ensure that automation remains high-performing, auditable, and compliant. Thank you for your attention, and I welcome your questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_RPA_Unit3_Administration.pptx" });
  convertToPdf("./outputs/Sem2_RPA_Unit3_Administration.pptx");
  console.log("✔  RPA Unit 3 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 3 RPA PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  console.log("All 3 RPA PPTs generated successfully!");
})();
