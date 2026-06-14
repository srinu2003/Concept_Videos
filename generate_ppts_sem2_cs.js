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
//  PPT 1 – THE CIA TRIAD & ATTACK VECTORS  (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Cyber Security";

  const tSl = titleSlide(p, H, "Unit I: Cyber Security Concepts", "The CIA Triad & Threat Landscapes");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'The CIA Triad and Threat Landscapes' as part of the Cyber Security course for my M.Tech in Computer Science and Engineering. In this presentation, we will explore the foundational concepts of cybersecurity, discuss the three pillars of the CIA Triad—Confidentiality, Integrity, and Availability—compare active versus passive attacks, and examine how these security concepts are implemented in real-world environments.");

  // Slide 2 ── What is Cyber Security?
  {
    const s = contentSlide(p, H, "What is Cyber Security?");
    s.addText([
      { text: "Cyber Security ", options: { bold: true } },
      { text: "is the practice of protecting systems, networks, programs, and data from digital attacks. It requires establishing multi-layered security controls across the computing infrastructure.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Concepts & Layers",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Vulnerability:", options: { bold: true } }, { text: " A weakness or loophole in system code or configuration (e.g. unpatched software).", options: { breakLine: true } },
      { text: "• Threat:", options: { bold: true, color: M } }, { text: " A potential danger that could exploit a vulnerability (e.g. hacker group, malware).", options: { breakLine: true } },
      { text: "• Assets & Harm:", options: { bold: true } }, { text: " Assets are data, hardware, and services. Harm is the impact of a threat exploiting a vulnerability.", options: { breakLine: true } },
      { text: "• Defense in Depth:", options: { bold: true, color: DGN } }, { text: " Implementing security at multiple layers (Physical, Network, Host, Application, Data).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Motives & Attacker Profiles",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Cyber Criminals:", options: { bold: true } }, { text: " Primarily motivated by financial gain (e.g. ransomware, credit card fraud).", options: { breakLine: true } },
      { text: "• Hacktivists:", options: { bold: true } }, { text: " Motivated by political or social causes (e.g. defacing websites, leaking documents).", options: { breakLine: true } },
      { text: "• Nation-State Agents:", options: { bold: true, color: M } }, { text: " Cyber warfare, espionage, and sabotaging foreign infrastructure.", options: { breakLine: true } },
      { text: "• Insider Threats:", options: { bold: true } }, { text: " Malicious or negligent employees leaking credentials from within.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us start by defining what Cyber Security is. In modern computing, cyber security is the broad practice of protecting our systems, networks, programs, and digital data from malicious digital attacks. These attacks are often aimed at accessing, changing, or destroying sensitive information, extorting money from users, or disrupting normal business operations. To secure our assets, we look at several core concepts. First, vulnerabilities are weaknesses or loopholes in our system code or configuration, such as unpatched software. Second, threats represent the potential danger that can exploit those vulnerabilities, which include threat actors like hacker groups or malware. Third, assets refer to the valuable data, services, or hardware we want to protect, and harm is the negative business impact when a threat succeeds. To defend our systems, we use a concept called 'Defense in Depth', which means placing overlapping security controls across physical, network, host, application, and data layers, so that if one control fails, others remain to block the attacker.");
  }

  // Slide 3 ── The CIA Triad
  {
    const s = contentSlide(p, H, "The CIA Triad");
    s.addText("The CIA Triad is the core model that guides security policies and controls:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    const pillars = [
      {
        t: "1. Confidentiality",
        b: "• Goal: Prevent unauthorized access to sensitive information.\n• Violations: Data breaches, packet sniffing, unauthorized file reads.\n• Controls: Encryption (AES), multi-factor authentication (MFA), access control lists.", bg: LGN, bd: GN
      },
      {
        t: "2. Integrity",
        b: "• Goal: Prevent unauthorized or undetected modification of data.\n• Violations: Man-in-the-middle edits, database tampering, virus code injection.\n• Controls: Hashing (SHA-256), digital signatures, file checksums.", bg: LBL, bd: BL
      },
      {
        t: "3. Availability",
        b: "• Goal: Guarantee reliable, timely access to systems for authorized users.\n• Violations: DDoS attacks, server crashes, power grid failures.\n• Controls: Redundant servers, load balancers, backups, DDoS protection.", bg: LMR, bd: M
      }
    ];

    pillars.forEach((pl, i) => {
      const bx = CX + i * 4.15;
      box(p, s, bx, 2.30, 4.00, 4.60, pl.bg, pl.bd);
      boxTitle(s, bx, 2.30, 4.00, pl.t, pl.bd);
      s.addText(pl.b,
        { x: bx + 0.12, y: 2.30 + 0.52, w: 4.00 - 0.24, h: 4.60 - 0.58, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });
    });

    s.addNotes("The fundamental guiding model for security policies and controls is known as the CIA Triad. This triad is composed of three core pillars: Confidentiality, Integrity, and Availability. Confidentiality focuses on preventing unauthorized access to sensitive information. Violations of confidentiality include data breaches, packet sniffing, or unauthorized file access, and we enforce it using tools like Advanced Encryption Standard (AES), multi-factor authentication, and strict access control lists. The second pillar is Integrity, which ensures that data is not modified or tampered with in transit or storage. Violations include database tampering or code injection. We protect integrity through cryptographic hash functions like SHA-256, digital signatures, and file checksum audits. The third pillar is Availability, which guarantees that authorized users have reliable and timely access to their systems. Common threats to availability include distributed denial of service, or DDoS attacks, and hardware crashes. We secure availability using redundant servers, load balancers, and regular data backups.");
  }

  // Slide 4 ── Active vs. Passive Attacks
  {
    const s = contentSlide(p, H, "Active vs. Passive Attacks");
    s.addText("Understanding how threat actors interact with data determines our defensive controls:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const tData = [
      [
        { text: "Characteristic", options: { bold: true, fill: { color: M }, color: WH } },
        { text: "Passive Attacks", options: { bold: true, fill: { color: M }, color: WH, align: "center" } },
        { text: "Active Attacks", options: { bold: true, fill: { color: M }, color: WH, align: "center" } }
      ],
      [
        { text: "Core Action", options: { bold: true, fill: { color: LGR } } }, { text: "Monitoring, sniffing, and reading traffic", options: { fill: { color: LGR } } }, { text: "Modifying, blocking, or injecting fake data", options: { fill: { color: LGR } } }
      ],
      [
        { text: "System Impact", options: { bold: true, fill: { color: LBL } } }, { text: "No alteration of database or system state", options: { fill: { color: LBL } } }, { text: "Alters data, system states, or network operations", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Examples", options: { bold: true, fill: { color: LGN } } }, { text: "Packet sniffing, traffic analysis, eavesdropping", options: { fill: { color: LGN } } }, { text: "DDoS, man-in-the-middle, replay, spoofing", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Detection Style", options: { bold: true, fill: { color: LMR } } }, { text: "Hard to detect (requires packet audits)", options: { fill: { color: LMR } } }, { text: "Easy to detect (crashes or alerts trigger)", options: { fill: { color: LMR } } }
      ],
      [
        { text: "Primary Defense", options: { bold: true, fill: { color: LYL } } }, { text: "Encryption (hides readable content)", options: { fill: { color: LYL } } }, { text: "Firewalls, intrusion detection, authentication", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: CX, y: 2.20, w: CW, h: 4.80, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.5, 4.92, 4.93]
      });

    s.addNotes("Now, let us examine the difference between active and passive attacks, as this determines how we design our defenses. Passive attacks focus on monitoring, sniffing, and reading network traffic without altering any system state or data. Common examples are packet sniffing and traffic analysis. Because no data is modified, these attacks are extremely difficult to detect in real-time, meaning our primary defense must be strong encryption, which renders any intercepted data unreadable. On the other hand, active attacks involve modifying, blocking, or injecting fake data into the system, directly altering its state. Examples include DDoS attacks, man-in-the-middle exploits, and message replays. These attacks are usually easy to detect because they cause system failures, service disruptions, or authentication errors. We defend against active attacks using real-time controls like firewalls, intrusion detection and prevention systems, and rigorous session authentication protocols.");
  }

  // Slide 5 ── Applications
  {
    const s = contentSlide(p, H, "Applications of Security Policies");
    s.addText("How organizations deploy the CIA triad and attack countermeasures in practice:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🔐  Multi-Factor Authentication (MFA)",
        b: "Protects Confidentiality. Requires users to provide two or more verification factors (password + SMS token + biometric fingerprint) to login.", bg: LGN, bd: GN
      },
      {
        t: "🛡️  Intrusion Detection Systems (IDS)",
        b: "Protects Availability and Integrity. Actively scans incoming network packets for signatures of active attacks, blocking malicious IPs automatically.", bg: LBL, bd: BL
      },
      {
        t: "🔒  End-to-End Encryption (E2EE)",
        b: "Protects Confidentiality and Integrity. Encrypts data at the source and decrypts only at destination, preventing passive packet sniffing mid-transit.", bg: LMR, bd: M
      },
      {
        t: "🔄  System Backup & DR plans",
        b: "Protects Availability. Regularly clones database systems to offline locations, ensuring quick recovery in the event of a ransomware attack.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let us look at how these security concepts are applied in production environments. First, multi-factor authentication, or MFA, is widely used to protect confidentiality by requiring users to verify their identity using multiple factors like passwords, SMS tokens, and biometrics. Second, Intrusion Detection and Prevention Systems are deployed at the network perimeter to protect availability and integrity by scanning incoming traffic for known attack signatures and blocking malicious IP addresses. Third, End-to-End Encryption, or E2EE, is implemented in communication platforms to ensure both confidentiality and integrity, preventing passive interceptors from reading message payloads. Lastly, organizations maintain offsite data backups and disaster recovery plans to safeguard availability, allowing quick restoration of systems following a destructive ransomware attack.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, understanding the CIA Triad and the nature of active and passive threat vectors is critical to establishing robust security defenses in modern organizations. By implementing a defense-in-depth approach, we can safeguard our digital assets against evolving threat landscapes. Thank you so much for your time and attention. I am now open to any questions you might have.");
  await p.writeFile({ fileName: "./outputs/Sem2_CS_Unit1_CIA_Threats.pptx" });
  convertToPdf("./outputs/Sem2_CS_Unit1_CIA_Threats.pptx");
  console.log("✔  CS Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – DIGITAL FORENSICS LIFECYCLE  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Cyber Security";

  const tSl = titleSlide(p, H, "Unit II: Cyber Forensics & Digital Evidence", "Digital Forensics Lifecycle");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on the 'Digital Forensics Lifecycle' as part of the Cyber Security curriculum for my M.Tech in Computer Science and Engineering. In this presentation, we will explore the core concepts of cyber forensics, trace the five-phase lifecycle of digital evidence gathering, examine how email analysis is conducted to trace cybercrime, and discuss corporate and legal applications of digital forensics.");

  // Slide 2 ── Introduction to Cyber Forensics
  {
    const s = contentSlide(p, H, "What is Cyber Forensics?");
    s.addText([
      { text: "Cyber Forensics ", options: { bold: true } },
      { text: "(also called Digital Forensics) is the branch of forensic science that deals with the recovery and investigation of material found in digital devices, preparing evidence that is legally admissible in court.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Evidence Integrity & Law",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Legal Admissibility:", options: { bold: true } }, { text: " Forensic investigators must follow strict protocols so that digital evidence is accepted by judges and juries.", options: { breakLine: true } },
      { text: "• Chain of Custody:", options: { bold: true, color: DGN } }, { text: " A chronological paper log documenting who collected, handled, and analyzed the evidence, preventing tampering.", options: { breakLine: true } },
      { text: "• Copying Integrity:", options: { bold: true } }, { text: " Investigators NEVER analyze the original hardware. They extract a bit-stream copy image and analyze the copy.", options: { breakLine: true } },
      { text: "• Hash Verification:", options: { bold: true } }, { text: " Verification that copy matches original using hashing algorithms (MD5, SHA-256).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Forensic Challenges",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Anti-Forensics:", options: { bold: true, color: M } }, { text: " Techniques used by criminals to hide evidence (encryption, secure file shredding, metadata scrubbing).", options: { breakLine: true } },
      { text: "• Data Volume:", options: { bold: true } }, { text: " Analyzing terabytes of storage on phones and hard drives efficiently.", options: { breakLine: true } },
      { text: "• Volatile Memory:", options: { bold: true } }, { text: " RAM data disappears when the system is powered down, requiring live extraction techniques.", options: { breakLine: true } },
      { text: "• Cloud Storage:", options: { bold: true } }, { text: " Evidence scattered across overseas servers, creating jurisdictional issues.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by defining cyber forensics. Also known as digital forensics, it is a specialized branch of forensic science focused on identifying, recovering, analyzing, and presenting evidence from digital devices in a manner that is legally admissible in a court of law. Ensuring the legal admissibility of evidence is paramount. To achieve this, investigators must maintain a strict 'Chain of Custody', which is a detailed, unbroken paper log documenting exactly who collected, handled, stored, and analyzed the physical or digital evidence at every single step. A key rule in digital forensics is that investigators never perform analysis on the original storage media, as this could corrupt the evidence. Instead, they use hardware write-blockers to extract a bit-stream, block-by-block copy image of the storage drive, and perform all analysis on this duplicate. They run cryptographic hash algorithms like MD5 or SHA-256 before and after the process to verify that the image matches the original source perfectly. However, forensics faces challenges like anti-forensics—which includes encryption and metadata scrubbing—large volumes of data, and the volatile nature of RAM data.");
  }

  // Slide 3 ── The Digital Forensics Lifecycle
  {
    const s = contentSlide(p, H, "The Digital Forensics Lifecycle");
    s.addText("The forensic process proceeds in five distinct, documented phases:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LYL, AM);
    s.addText("The Five Forensic Phases",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Identification:", options: { bold: true } }, { text: " Identify potential evidence sources (computers, phones, USB drives, cloud logs). Document the physical scene.", options: { breakLine: true } },
      { text: "2. Preservation:", options: { bold: true, color: M } }, { text: " Isolate devices from networks to prevent remote wipes (e.g. place phones in Faraday bags). Use hardware write-blockers to copy storage bit-by-bit.", options: { breakLine: true } },
      { text: "3. Extraction:", options: { bold: true } }, { text: " Recover files, including hidden or deleted data, database records, and browser histories, using forensic software.", options: { breakLine: true } },
      { text: "4. Analysis:", options: { bold: true, color: DGN } }, { text: " Reconstruct timelines, locate keyword matches, trace emails, and link user actions to the crime. Answer: Who, What, When, and How.", options: { breakLine: true } },
      { text: "5. Reporting & Presentation:", options: { bold: true } }, { text: " Draft a detailed expert witness report explaining findings clearly for non-technical judges and juries.", options: {} }
    ], { x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("The digital forensics process is structured into five distinct, chronologically documented phases to ensure validity. First is Identification, where investigators identify potential evidence sources such as laptops, smartphones, server logs, or USB drives, and carefully document the physical and digital scene. Second is Preservation, where the devices are isolated from networks using tools like Faraday bags to prevent remote data wipes, and bit-stream copies are created using write-blockers. Third is Extraction, where specialized software is used to recover active files, log histories, and even hidden or deleted data from unallocated space. Fourth is Analysis, where the investigator reconstructs timelines, performs keyword searches, and analyzes registry entries to link a specific user identity to the digital crime, answering the critical questions of who, what, when, and how. Finally, the fifth phase is Reporting and Presentation, which involves writing a comprehensive technical report that translates complex binary structures into clear, understandable evidence for non-technical judges and juries.");
  }

  // Slide 4 ── Forensic Email Analysis
  {
    const s = contentSlide(p, H, "Forensic Email Analysis");
    s.addText("Email is a primary vector for cybercrime. Forensic analysis involves parsing headers and metadata:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Parsing Email Headers",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• SMTP Hops (Received: fields):", options: { bold: true } }, { text: " Traces the sequence of mail servers the email passed through, revealing the true sending IP address.", options: { breakLine: true } },
      { text: "• Message-ID:", options: { bold: true } }, { text: " A unique identifier assigned by the sending mail server. Useful for database log auditing.", options: { breakLine: true } },
      { text: "• Authentication Checks:", options: { bold: true, color: DGN } }, { text: " We check:\n", options: { breakLine: true } },
      { text: "   - SPF (Sender Policy Framework):", options: { bold: true } }, { text: " Checks if the sending IP is authorized to mail on behalf of the domain.\n", options: { breakLine: true } },
      { text: "   - DKIM (DomainKeys Identified Mail):", options: { bold: true } }, { text: " Verifies cryptographic signatures to ensure the email was not tampered with in transit.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Tracing Phishing & Spoofing",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Spoofed Headers:", options: { bold: true, color: M } }, { text: " Attackers modify the 'From:' field display name. However, parsing the raw headers reveals the actual sending email address.", options: { breakLine: true } },
      { text: "• Attachment Analysis:", options: { bold: true } }, { text: " Extracting hidden macro code or malware binaries from PDF/Office attachments inside sandbox test environments.", options: { breakLine: true } },
      { text: "• Mail Server Logs:", options: { bold: true, color: BL } }, { text: " Auditing SMTP server transaction logs to confirm connections, authentications, and delivery timestamps.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Email is one of the most common vectors used in cybercrimes, including phishing, financial fraud, and business email compromise. Forensic email analysis requires parsing the raw, hidden headers of an email to verify its authenticity. We look closely at SMTP hops, which are the 'Received' fields in the headers. These fields record every mail server the email passed through, allowing investigators to trace the traffic back to the sender's actual IP address, bypassing spoofed names. We also examine the unique Message-ID and verify cryptographic authentication records. This includes checking Sender Policy Framework, or SPF, which lists authorized sending IPs for a domain, and DomainKeys Identified Mail, or DKIM, which uses digital signatures to confirm the email payload was not altered during transit. We also perform attachment analysis inside secure sandbox environments to isolate any hidden macro malware, and audit mail server log catalogs to confirm transactional handshakes.");
  }

  // Slide 5 ── Applications of Cyber Forensics
  {
    const s = contentSlide(p, H, "Applications of Cyber Forensics");
    s.addText("Where digital forensics is implemented in legal and corporate environments:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "⚖️  Criminal Law Case Evidence",
        b: "Extracts call logs, GPS locations, and deleted chat histories from suspects' smartphones, presenting key timelines to prosecutors.", bg: LGN, bd: GN
      },
      {
        t: "🏢  Corporate Data Breach Audits",
        b: "Investigates corporate systems after data breaches. Identifies how hackers entered (the initial entry vector) and logs what data was stolen.", bg: LBL, bd: BL
      },
      {
        t: "💼  Insider Threat Investigations",
        b: "Audits logs of employees suspected of leaking intellectual property. Reconstructs file copy actions and USB insert histories.", bg: LMR, bd: M
      },
      {
        t: "🔄  Accidental Data Recovery",
        b: "Assists companies in recovering critical database files deleted by accident or deleted during drive corruptions.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Digital forensics is applied across legal, corporate, and operational domains. In criminal cases, investigators extract call logs, geographical coordinates, and deleted chat messages from smartphones to present hard evidence to prosecutors. In the corporate sector, forensics is crucial during data breach audits, helping security teams identify the initial entry vector of hackers and determine exactly what sensitive files were exfiltrated. It is also used to investigate insider threats, such as employees suspected of leaking intellectual property, by auditing file copy records and registry logs for external USB inserts. Finally, digital forensics assists organizations in recovering critical files lost during accidental deletions or storage drive corruptions.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("To conclude, the digital forensics lifecycle provides a standardized, scientifically sound methodology to recover and analyze digital evidence. By adhering to protocols like chain of custody and hash verification, forensics ensures that digital evidence remains untampered and legally admissible in court. Thank you for your time, and I welcome any questions or discussion on this topic.");
  await p.writeFile({ fileName: "./outputs/Sem2_CS_Unit2_Forensics.pptx" });
  convertToPdf("./outputs/Sem2_CS_Unit2_Forensics.pptx");
  console.log("✔  CS Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – MOBILE SECURITY & VULNERABILITIES  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Cyber Security";

  const tSl = titleSlide(p, H, "Unit III: Mobile & Wireless Security", "Mobile Device Security & Vulnerabilities");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on the topic of 'Mobile Device Security and Vulnerabilities' as part of the Cyber Security course for my M.Tech in Computer Science and Engineering. In this presentation, we will discuss the security risks associated with the massive proliferation of mobile devices, examine the Bring Your Own Device, or BYOD, policy challenges, analyze core mobile operating system security features, explore Mobile Device Management, or MDM, systems, and look at real-world applications of mobile security.");

  // Slide 2 ── Proliferation of Mobile Devices
  {
    const s = contentSlide(p, H, "Mobile Device Proliferation & Risks");
    s.addText([
      { text: "Mobile Proliferation ", options: { bold: true } },
      { text: "has transformed corporate operations. Bring Your Own Device (BYOD) policies allow employees to access sensitive corporate databases on personal smartphones, creating severe security challenges.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("The BYOD Security Threat",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Mixed Contexts:", options: { bold: true, color: M } }, { text: " Personal apps (games, social media) run alongside corporate emails and databases on the same device.", options: { breakLine: true } },
      { text: "• Unsecured Networks:", options: { bold: true } }, { text: " Mobile devices connect to public, unencrypted Wi-Fi networks, exposing traffic to sniffing.", options: { breakLine: true } },
      { text: "• Physical Loss / Theft:", options: { bold: true, color: M } }, { text: " Smartphones are easily lost. If unencrypted, finder has access to stored company data.", options: { breakLine: true } },
      { text: "• Rooted / Jailbroken Devices:", options: { bold: true } }, { text: " Users bypass OS security rules, enabling malware to access sandbox resources.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Common Mobile Attack Vectors",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Malicious Apps:", options: { bold: true, color: M } }, { text: " App store clones hiding spyware, tracking locations, and logging keystrokes.", options: { breakLine: true } },
      { text: "• Smishing (SMS Phishing):", options: { bold: true } }, { text: " SMS messages containing links to fake bank portals or malware downloads.", options: { breakLine: true } },
      { text: "• Credit Card / SMS Frauds:", options: { bold: true } }, { text: " Malicious apps registering users to premium SMS services or copying credit card details from tap-to-pay NFC interfaces.", options: { breakLine: true } },
      { text: "• Vulnerable OS versions:", options: { bold: true } }, { text: " Slow vendor update cycles leave Android/iOS devices unpatched against known kernel exploits.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("The rapid proliferation of mobile devices has changed how organizations operate. Employees now access sensitive corporate databases, customer records, and email accounts using personal smartphones, often guided by Bring Your Own Device, or BYOD, policies. While BYOD increases flexibility, it introduces severe security risks. First, it mixes personal and corporate contexts: employees run unvetted games, social media, and messaging apps alongside corporate databases on the same physical hardware, allowing potential malware to scrape business data. Second, mobile devices connect to unsecured public Wi-Fi networks in airports or cafes, exposing sensitive traffic to passive packet sniffing attacks. Third, smartphones are easily lost or stolen; if the device lacks strong encryption, a thief can read corporate files directly. Finally, some users jailbreak or root their devices, bypassing the built-in operating system security boundaries and exposing application memory to malicious exploits.");
  }

  // Slide 3  ── Mobile OS Security & MDM
  {
    const s = contentSlide(p, H, "Mobile OS Security & MDM Controls");
    s.addText("To secure mobile endpoints, enterprises use a combination of OS features and MDM software:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Mobile OS Security Features",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• App Sandboxing:", options: { bold: true, color: DGN } }, { text: " Each app runs in an isolated container. A game cannot read data from a banking app's memory.", options: { breakLine: true } },
      { text: "• Permission Systems:", options: { bold: true } }, { text: " Users must explicitly authorize app access to camera, contacts, location, and files.", options: { breakLine: true } },
      { text: "• Biometric Authentication:", options: { bold: true } }, { text: " Uses face/fingerprint scans to unlock cryptographic key chains dynamically.", options: { breakLine: true } },
      { text: "• Hardware Encryption:", options: { bold: true } }, { text: " Storage drives are encrypted by default, utilizing dedicated secure enclave processors.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Mobile Device Management (MDM)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Central Administration:", options: { bold: true, color: BL } }, { text: " Software (e.g. Microsoft Intune) that manages corporate configurations on devices.", options: { breakLine: true } },
      { text: "• Remote Wipe:", options: { bold: true, color: M } }, { text: " Instantly erases all corporate data if a device is reported lost or stolen.", options: { breakLine: true } },
      { text: "• Secure Containerization:", options: { bold: true } }, { text: " Separates personal apps from corporate apps, blocking copy-paste actions between them.", options: { breakLine: true } },
      { text: "• Compliance Locks:", options: { bold: true } }, { text: " Blocks devices from corporate networks if they are rooted, jailbroken, or missing password locks.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("To counter these risks, we combine native mobile operating system security features with centralized Mobile Device Management, or MDM, systems. Mobile operating systems like Android and iOS use App Sandboxing to run each application in an isolated container, preventing a game app from reading files from a secure banking app. They also use explicit runtime permission prompts, biometric authentications like face scans, and default hardware-level storage encryption powered by secure enclaves. In addition, enterprises implement Mobile Device Management software, such as Microsoft Intune. MDM platforms allow administrators to centrally manage device configurations. They support 'Remote Wipe' capabilities to erase all corporate data if a device is lost. They also create secure business containers on personal phones, blocking copy-paste actions between corporate apps and personal channels, and enforce compliance locks that block access if a device is rooted or lacks secure screen locks.");
  }

  // Slide 4 ── Applications of Mobile Security
  {
    const s = contentSlide(p, H, "Applications of Mobile Security");
    s.addText("Where mobile security protocols are implemented to secure transactions and assets:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🏦  Mobile Banking Applications",
        b: "Uses app sandboxing, mandatory MFA, biometric verification, and SSL pinning. Detects rooted devices to block execution, preventing balance tampering.", bg: LGN, bd: GN
      },
      {
        t: "🏥  Securing Electronic Health Records (EHR)",
        b: "Doctors use tablet systems to view patient records. MDM containers ensure HIPAA compliance by encrypting data and blocking screen captures.", bg: LBL, bd: BL
      },
      {
        t: "🏢  Corporate BYOD Email Access",
        b: "Companies use MDM policies to lock email access behind secure PIN containers, protecting corporate data on employee phones.", bg: LMR, bd: M
      },
      {
        t: "🛒  Mobile Point of Sale (POS) Tablets",
        b: "Retail tablets running card transactions encrypt swipe data at the card reader, preventing credit card theft over wireless networks.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let us look at how mobile security is applied in key sectors. In mobile banking applications, developers enforce app sandboxing, mandatory multi-factor authentication, and SSL pinning, and block the application from running if the OS is rooted. In healthcare, doctors use tablets to view patient records; MDM containerization encrypts these records and blocks screen capture capabilities, ensuring compliance with health regulations like HIPAA. For corporate BYOD setups, MDM separates personal apps from business emails. Finally, in mobile Point of Sale systems used by retail stores, POS applications encrypt credit card data at the hardware reader level, ensuring that transaction info is never exposed over local wireless networks.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In summary, securing mobile endpoints requires a multi-layered approach that bridges native operating system sandboxing with centralized enterprise management tools. As mobile devices continue to be the primary interface for business applications, maintaining robust MDM and OS controls is vital to defending against digital threats. Thank you very much for your time. I am now open to any questions you may have.");
  await p.writeFile({ fileName: "./outputs/Sem2_CS_Unit3_Mobile_Security.pptx" });
  convertToPdf("./outputs/Sem2_CS_Unit3_Mobile_Security.pptx");
  console.log("✔  CS Unit 3 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 3 CS PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  console.log("All 3 CS PPTs generated successfully!");
})();
