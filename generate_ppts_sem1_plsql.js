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
//  PPT 1 – PL/SQL BLOCK STRUCTURE & VARIABLES  (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Database Programming with PL/SQL";

  const s1 = titleSlide(p, H, "Unit I: PL/SQL Basics", "PL/SQL Block Structure & Variables");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. I am a first-year, first-semester M.Tech student in Computer Science and Engineering. Today, I will present PL/SQL Block Structure and Variables. In this presentation, we will explore what PL/SQL is as Oracle's procedural extension, examine the mandatory and optional parts of PL/SQL block structures, study variable declarations and type anchoring with percent-type and percent-rowtype, look at robust exception handling, and discuss practical applications of PL/SQL blocks in database application development.");

  // Slide 2 ── What is PL/SQL?
  {
    const s = contentSlide(p, H, "What is PL/SQL?");
    s.addText([
      { text: "PL/SQL ", options: { bold: true } },
      { text: "(Procedural Language/Structured Query Language) is Oracle Corporation's procedural extension to SQL. It integrates SQL queries with procedural programming constructs (loops, conditions, variables).", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Features & Integration",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Block Structured:", options: { bold: true } }, { text: " Code is structured into logically enclosed blocks. Each block is a single unit of execution.", options: { breakLine: true } },
      { text: "• Tight SQL Integration:", options: { bold: true } }, { text: " Supports all SQL statements, data types, and functions directly within the code.", options: { breakLine: true } },
      { text: "• Client-Server efficiency:", options: { bold: true } }, { text: " Groups multiple SQL queries into a single block sent to the database engine, reducing network traffic and latency.", options: { breakLine: true } },
      { text: "• Performance gains:", options: { bold: true } }, { text: " Compiled and run inside the Oracle Database kernel, maximizing data-processing speeds.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Benefits in Application Development",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Procedural Logic:", options: { bold: true } }, { text: " Allows developers to implement complex algorithms, logical checks, and loops directly on the database side.", options: { breakLine: true } },
      { text: "• Robust Exception Handling:", options: { bold: true } }, { text: " Has built-in, custom error-trapping blocks to handle query failures gracefully.", options: { breakLine: true } },
      { text: "• Portability & Security:", options: { bold: true } }, { text: " Runs on any OS hosting Oracle Database. Business logic is secured inside database schemas.", options: { breakLine: true } },
      { text: "• Modular Architecture:", options: { bold: true } }, { text: " Supports packaging code into reusable procedures, functions, and packages.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("What is PL/SQL? It stands for Procedural Language extension to SQL, developed by Oracle Corporation. While SQL is a declarative language where you specify what data you want, PL/SQL allows you to write step-by-step procedural logic using loops, conditions, and variables. Its core benefits include tight SQL integration, supporting all SQL data types and functions directly. It is block-structured, meaning code is grouped into logical, reusable units. Most importantly, it is highly efficient. Instead of sending queries one-by-one over the network, PL/SQL groups queries into a single block sent to the database engine, reducing network traffic and latency. Since it compiles and runs inside the Oracle Database kernel, it maximizes processing speed and secures business logic.");
  }

  // Slide 3 ── PL/SQL Block Structure
  {
    const s = contentSlide(p, H, "PL/SQL Block Structure");
    s.addText("PL/SQL code is written inside standardized, nested blocks containing up to three sections:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGR, GY);
    s.addText("Typical Syntax Structure",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Courier New", margin: 0 });
    s.addText([
      { text: "[DECLARE]\n", options: { bold: true, color: M } },
      { text: "   -- Declaration of variables, types, cursors\n", options: { color: GY } },
      { text: "BEGIN\n", options: { bold: true, color: M } },
      { text: "   -- Executable statements (SQL and PL/SQL)\n", options: { color: BK } },
      { text: "[EXCEPTION]\n", options: { bold: true, color: M } },
      { text: "   -- Error handling and trapping code\n", options: { color: GY } },
      { text: "END;\n", options: { bold: true, color: M } },
      { text: "/  -- Executes the block", options: { color: M } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13.5, fontFace: "Courier New", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LMR, M);
    s.addText("Section Details",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• DECLARE (Optional):", options: { bold: true, color: M } }, { text: " Allocates memory for variables, constants, cursors, and user-defined exceptions. Not accessible outside the block.", options: { breakLine: true } },
      { text: "• BEGIN (Mandatory):", options: { bold: true, color: DGN } }, { text: " Contains the core procedural instructions and SQL queries. Must have at least one executable statement (can be 'NULL;').", options: { breakLine: true } },
      { text: "• EXCEPTION (Optional):", options: { bold: true, color: AM } }, { text: " Traps errors that occur during the execution of the BEGIN section, preventing program crashes.", options: { breakLine: true } },
      { text: "• Anonymous vs. Named:", options: { bold: true } }, { text: " Anonymous blocks are compiled on the fly and not saved. Named blocks (procedures, functions) are stored in the database catalog.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's examine the PL/SQL Block Structure. Code is written in nested blocks containing three sections. The optional DECLARE section is where variables, constants, and cursors are allocated in memory. The mandatory BEGIN section contains procedural instructions and SQL queries; it must have at least one executable statement, even if it is just a NULL command. The optional EXCEPTION section traps runtime errors to prevent program crashes. Blocks can be anonymous, which are compiled and executed immediately in memory, or named, like stored procedures, functions, or triggers, which are saved permanently in the database catalog.");
  }

  // Slide 4 ── Variables and Data Types
  {
    const s = contentSlide(p, H, "Variables and Data Types");
    s.addText("PL/SQL supports robust scalar, composite, reference, and LOB data types:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Scalar Types & Anchors",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Standard SQL Scalars:", options: { bold: true } }, { text: " VARCHAR2, NUMBER, DATE, CHAR, BOOLEAN (unique to PL/SQL).", options: { breakLine: true } },
      { text: "• %TYPE Attribute (Anchored Type):", options: { bold: true, color: BL } }, { text: " Declares a variable with the exact data type of a database column. Dynamically adapts to table schema changes:\n   emp_salary  employees.salary%TYPE;", options: { breakLine: true } },
      { text: "• CONSTANT Declaration:", options: { bold: true } }, { text: " Binds a value that cannot change during block execution:\n   pi  CONSTANT  NUMBER := 3.14159;", options: { breakLine: true } },
      { text: "• NOT NULL:", options: { bold: true } }, { text: " Forces variable initialization and prevents assigning nulls.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Composite Record Types",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• PL/SQL Records:", options: { bold: true } }, { text: " Custom group of related fields. Similar to a struct in C/C++ or an object attribute list.", options: { breakLine: true } },
      { text: "• %ROWTYPE Attribute:", options: { bold: true, color: M } }, { text: " Declares a record that represents an entire row of a table or view. Fields map to table columns dynamically:\n   emp_rec  employees%ROWTYPE;\n   -- Access: emp_rec.first_name := 'Srinivas';", options: { breakLine: true } },
      { text: "• Initialization:", options: { bold: true } }, { text: " Variables are initialized using the assignment operator ':=' or the DEFAULT keyword. Uninitialized variables default to NULL.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("PL/SQL supports scalar, composite, reference, and LOB data types. Scalar types include VARCHAR2, NUMBER, DATE, and BOOLEAN. A key feature is the percent-type attribute, which anchors a variable to a database column's type. For example, declaring employee salary as employees dot salary percent-type ensures the variable automatically adapts if the DBA changes the column's definition, avoiding compilation errors. We can also declare variables as CONSTANT or NOT NULL. For composite records, we use the percent-rowtype attribute. This declares a record variable representing an entire database table row, mapping fields dynamically and allowing us to select rows directly into records using SELECT INTO.");
  }

  // Slide 5 ── Exception Handling
  {
    const s = contentSlide(p, H, "Exception Handling: Resilient Code");
    s.addText("Exceptions handle runtime errors cleanly, separating error-handling from main logic:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Types of Exceptions",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Predefined Exceptions:", options: { bold: true } }, { text: " Automatically raised by Oracle Database runtime for common errors:\n", options: { breakLine: true } },
      { text: "   - NO_DATA_FOUND:", options: { bold: true, color: M } }, { text: " SELECT INTO returned 0 rows.\n", options: { breakLine: true } },
      { text: "   - TOO_MANY_ROWS:", options: { bold: true, color: M } }, { text: " SELECT INTO returned >1 row.\n", options: { breakLine: true } },
      { text: "   - ZERO_DIVIDE:", options: { bold: true, color: M } }, { text: " Division by zero error.\n", options: { breakLine: true } },
      { text: "• User-Defined Exceptions:", options: { bold: true } }, { text: " Custom exceptions declared in DECLARE section and triggered manually using the RAISE statement when a business rule is violated.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Error Handling Block Syntax",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Courier New", margin: 0 });
    s.addText([
      { text: "EXCEPTION\n", options: { bold: true, color: M } },
      { text: "  WHEN NO_DATA_FOUND THEN\n", options: { bold: true, color: BL } },
      { text: "    dbms_output.put_line('No Employee found!');\n", options: { color: BK } },
      { text: "  WHEN OTHERS THEN\n", options: { bold: true, color: BL } },
      { text: "    -- Handles any other error\n", options: { color: GY } },
      { text: "    dbms_output.put_line('Error: ' || SQLERRM);\n", options: { color: BK } },
      { text: "    ROLLBACK;\n", options: { color: BK, bold: true } }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13.5, fontFace: "Courier New", color: BK, margin: 0 });

    s.addNotes("Exceptions handle runtime errors cleanly, separating error-handling from main logic. Oracle has predefined exceptions raised automatically, such as NO_DATA_FOUND when a SELECT INTO returns zero rows, and TOO_MANY_ROWS when it returns more than one. We can also declare user-defined exceptions and trigger them manually using the RAISE statement when business rules are violated. The EXCEPTION block handles these errors using WHEN clauses. The WHEN OTHERS block acts as a catch-all to trap any unhandled error, where developers can log error details using SQLCODE and SQLERRM and perform a ROLLBACK to maintain transaction integrity.");
  }

  // Slide 6 ── Applications of PL/SQL Blocks
  {
    const s = contentSlide(p, H, "Applications of PL/SQL Blocks");
    s.addText("PL/SQL block structures are the backbone of secure database application development:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "💳  Transactional Controls",
        b: "Groups complex SQL operations (e.g., deducting balance and adding to another account) into a single execution block. Commits transaction if all succeed, or performs a ROLLBACK in the EXCEPTION block if any step fails.", bg: LGN, bd: GN
      },
      {
        t: "📊  Batch Data Updates",
        b: "Iterates through database cursors and performs batch updates. Bulk statements load millions of rows efficiently with minimal CPU overhead.", bg: LBL, bd: BL
      },
      {
        t: "🔐  Data Validation API",
        b: "Implements business validation constraints (e.g., checking user eligibility, verifying credit bounds) prior to database insertion, protecting database integrity.", bg: LMR, bd: M
      },
      {
        t: "⚙️  ETL Processes",
        b: "Extracts, transforms, and loads data during data migrations. Exception sections capture and log malformed rows to error tables, keeping migrations running.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Finally, what are the applications of PL/SQL blocks? They are the backbone of secure database applications. In Transactional Controls, they group operations like funds transfers, committing only if all succeed, or rolling back in the exception block. In Batch Data Updates, they iterate through cursors to process millions of rows efficiently. In Data Validation APIs, they enforce business rule constraints before records are inserted, protecting database integrity. In ETL processes, they extract, transform, and load data during migrations, capturing and logging errors without halting the entire process.");
  }

  const s7 = thankYouSlide(p, H);
  s7.addNotes("That brings me to the end of this presentation on PL/SQL Block Structure and Variables. We have seen how combining SQL with procedural logic under structured blocks makes database applications highly efficient, secure, and resilient. I would like to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you all for your attention.");

  await p.writeFile({ fileName: "./outputs/Sem1_PLSQL_Unit1_Blocks_Variables.pptx" });
  convertToPdf("./outputs/Sem1_PLSQL_Unit1_Blocks_Variables.pptx");
  console.log("✔  PLSQL Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – ORACLE COLLECTIONS & ARRAYS  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Database Programming with PL/SQL";

  const s1 = titleSlide(p, H, "Unit II: Language Fundamentals & Control Structures", "Oracle Collections & Arrays");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. In this presentation, I will cover Oracle Collections and Arrays, a fundamental topic in PL/SQL. We will understand what collections are as dynamic in-memory arrays, compare the three main collection types: Varrays, Nested Tables, and Associative Arrays, look at variable-size array syntax and boundaries, examine the differences between Nested Tables and Associative Arrays, review the collection API methods, and study how collections are applied in bulk SQL operations like bulk collect and forall.");

  // Slide 2 ── What are Oracle Collections?
  {
    const s = contentSlide(p, H, "What are Oracle Collections?");
    s.addText([
      { text: "Oracle Collections ", options: { bold: true } },
      { text: "are single-dimensional, homogeneous arrays that store groups of elements. They are dynamic in-memory data structures utilized to batch and process bulk database rows.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Characteristics",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Homogeneous:", options: { bold: true } }, { text: " All elements in a collection must have the exact same data type.", options: { breakLine: true } },
      { text: "• In-Memory:", options: { bold: true } }, { text: " Standard collection variables exist during PL/SQL execution, held in the session memory.", options: { breakLine: true } },
      { text: "• Dynamic Size:", options: { bold: true } }, { text: " Collections can shrink or grow at runtime using built-in API calls.", options: { breakLine: true } },
      { text: "• Bulk Operations:", options: { bold: true, color: DGN } }, { text: " Used to fetch multiple rows in a single step (BULK COLLECT) and bind arrays to SQL (FORALL).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Three Main Collection Types",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Varrays (Variable-Size Arrays):", options: { bold: true, color: M } }, { text: " Have a fixed maximum size. Ordered lists starting at index 1. Can be stored in database table columns.", options: { breakLine: true } },
      { text: "• Nested Tables:", options: { bold: true, color: M } }, { text: " Unbounded size. Starts at index 1. Can become sparse (have gaps) when elements are deleted. Can be stored in the database.", options: { breakLine: true } },
      { text: "• Associative Arrays (Index-by Tables):", options: { bold: true, color: M } }, { text: " Unbounded size. Can be indexed by integers or strings (key-value maps). Exist in memory only. Cannot be stored in database tables.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's first define Oracle Collections. They are single-dimensional, homogeneous arrays that store groups of elements. They are dynamic in-memory structures used to batch and process bulk database records. Their core characteristics are: first, they are homogeneous, meaning all elements must share the same type; second, they exist in session memory during execution; and third, their size is dynamic, growing or shrinking at runtime. The three main collection types are Varrays, which are ordered lists with a fixed maximum size; Nested Tables, which are unbounded lists that can become sparse when elements are deleted; and Associative Arrays, also known as index-by tables, which are in-memory key-value maps indexed by integers or strings.");
  }

  // Slide 3 ── Varrays (Variable-Size Arrays)
  {
    const s = contentSlide(p, H, "Varrays (Variable-Size Arrays)");
    s.addText("Varrays represent ordered, bounded lists of items. Ideal for small, fixed collections:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, GN);
    s.addText("Syntax & Declaration",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Courier New", margin: 0 });
    s.addText([
      { text: "-- Define Varray type\n", options: { color: GY } },
      { text: "TYPE phone_list IS VARRAY(3) OF VARCHAR2(15);\n\n", options: { color: BK, bold: true } },
      { text: "-- Declare and Initialize variable\n", options: { color: GY } },
      { text: "my_phones phone_list := phone_list('123-456', '789-012');\n\n", options: { color: BK } },
      { text: "-- Accessing element\n", options: { color: GY } },
      { text: "first_phone := my_phones(1); -- 1-indexed", options: { color: BK } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Courier New", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Key Rules & Limitations",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Strict Boundary:", options: { bold: true } }, { text: " Cannot add more elements than the limit declared (e.g. VARRAY(3) cannot hold 4 elements).", options: { breakLine: true } },
      { text: "• Sequential Order:", options: { bold: true } }, { text: " Cannot delete individual elements from the middle to make it sparse. You can only delete from the end.", options: { breakLine: true } },
      { text: "• Constructor Initialization:", options: { bold: true } }, { text: " Must be initialized using a constructor before use. Using an uninitialized varray raises a COLLECTION_IS_NULL exception.", options: { breakLine: true } },
      { text: "• DB Storage:", options: { bold: true } }, { text: " Stored inline in database tables if small, preserving the exact element sequence.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's look closer at Varrays. They represent ordered, bounded lists of items, ideal when the maximum size is known. To use a varray, we first define the type, specifying the limit—such as VARRAY of three elements—then declare and initialize the variable using a constructor. If we try to access an uninitialized varray, PL/SQL throws a collection is null exception, so constructor initialization is mandatory. Varrays have strict boundaries; trying to add more than the limit raises an error. They maintain sequential order, meaning we cannot delete individual elements from the middle to make it sparse; we can only delete from the end. If stored in a database table, varrays are kept inline, preserving their exact sequence.");
  }

  // Slide 4 ── Nested Tables & Associative Arrays
  {
    const s = contentSlide(p, H, "Nested Tables vs. Associative Arrays");
    s.addText("Understanding the core differences between unbounded collection types:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const tData = [
      [
        { text: "Feature", options: { bold: true, fill: { color: M }, color: WH } },
        { text: "Nested Tables", options: { bold: true, fill: { color: M }, color: WH, align: "center" } },
        { text: "Associative Arrays", options: { bold: true, fill: { color: M }, color: WH, align: "center" } }
      ],
      [
        { text: "Database Storage", options: { bold: true, fill: { color: LGR } } }, { text: "Can be stored in DB table columns", options: { fill: { color: LGR } } }, { text: "PL/SQL memory only (cannot be in DB)", options: { fill: { color: LGR } } }
      ],
      [
        { text: "Indexing Key", options: { bold: true, fill: { color: LBL } } }, { text: "Sequential integers starting at 1", options: { fill: { color: LBL } } }, { text: "Integers or VARCHAR2 strings (key-value)", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Sparseness", options: { bold: true, fill: { color: LGN } } }, { text: "Can become sparse (using DELETE(i))", options: { fill: { color: LGN } } }, { text: "Naturally sparse (keys can be non-sequential)", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Initialization", options: { bold: true, fill: { color: LMR } } }, { text: "Requires constructor: nested_tab_var()", options: { fill: { color: LMR } } }, { text: "No constructor; initialized on first assignment", options: { fill: { color: LMR } } }
      ],
      [
        { text: "Dynamic Growth", options: { bold: true, fill: { color: LYL } } }, { text: "Requires EXTEND to allocate memory", options: { fill: { color: LYL } } }, { text: "Grows automatically as new keys are added", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: CX, y: 2.20, w: CW, h: 4.80, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [3.0, 4.67, 4.68]
      });

    s.addNotes("Next, let's contrast Nested Tables and Associative Arrays, which are both unbounded. Nested Tables can be stored in database columns, whereas Associative Arrays exist in memory only. Nested tables are indexed by sequential integers starting at one, but can become sparse if we delete elements from the middle. Associative arrays act like hash maps, allowing index keys to be strings or non-sequential integers, making them naturally sparse. While Nested Tables require constructor initialization and using the extend method to allocate memory before adding elements, Associative Arrays require no constructors or extend calls, growing automatically when we assign values to new keys.");
  }

  // Slide 5 ── Oracle Collection API
  {
    const s = contentSlide(p, H, "Oracle Collection API Methods");
    s.addText("Collections are manipulated using built-in method functions:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LGN, GN);
    s.addText("Core Collection Methods Table",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });

    const tData = [
      [
        { text: "Method", options: { bold: true, fill: { color: GN }, color: WH, align: "center" } },
        { text: "Syntax Example", options: { bold: true, fill: { color: GN }, color: WH, align: "center" } },
        { text: "Description & Functionality", options: { bold: true, fill: { color: GN }, color: WH } }
      ],
      [
        { text: "COUNT", options: { bold: true, fill: { color: LGR }, align: "center" } }, { text: "var.COUNT", options: { fill: { color: LGR }, align: "center" } },
        { text: "Returns the number of elements currently stored in the collection.", options: { fill: { color: LGR } } }
      ],
      [
        { text: "EXISTS", options: { bold: true, fill: { color: LBL }, align: "center" } }, { text: "var.EXISTS(i)", options: { fill: { color: LBL }, align: "center" } },
        { text: "Returns TRUE if element at index i is initialized; FALSE otherwise.", options: { fill: { color: LBL } } }
      ],
      [
        { text: "FIRST / LAST", options: { bold: true, fill: { color: LGN }, align: "center" } }, { text: "var.FIRST", options: { fill: { color: LGN }, align: "center" } },
        { text: "Returns the first (smallest) and last (largest) index keys in the collection.", options: { fill: { color: LGN } } }
      ],
      [
        { text: "EXTEND", options: { bold: true, fill: { color: LMR }, align: "center" } }, { text: "var.EXTEND(n)", options: { fill: { color: LMR }, align: "center" } },
        { text: "Allocates n new null elements at the end of Nested Tables / Varrays.", options: { fill: { color: LMR } } }
      ],
      [
        { text: "DELETE", options: { bold: true, fill: { color: LYL }, align: "center" } }, { text: "var.DELETE(i)", options: { fill: { color: LYL }, align: "center" } },
        { text: "Removes the element at index i, making the collection sparse.", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [1.8, 2.2, 8.11]
      });

    s.addNotes("Oracle collections provide a built-in API with helper methods for array manipulation. The COUNT method returns the number of elements currently stored. The EXISTS method checks if a specific index is initialized, which is crucial for handling sparse collections and avoiding out of bounds errors. FIRST and LAST return the smallest and largest index keys. For Nested Tables and Varrays, the EXTEND method allocates empty slots at the end. The DELETE method removes elements, which makes Nested Tables or Associative Arrays sparse. These methods allow developers to write clean loops and safely access collection elements.");
  }

  // Slide 6 ── Applications: Bulk SQL Bindings
  {
    const s = contentSlide(p, H, "Applications: Bulk SQL Processing");
    s.addText("The primary use of collections is optimizing database queries using bulk operations:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "⚡  BULK COLLECT Statement",
        b: "Fetches an entire SQL query result set directly into a collection variable in a single context switch. Dramatically faster than fetching rows one-by-one in a cursor loop.", bg: LGN, bd: GN
      },
      {
        t: "🏎️  FORALL Statement",
        b: "Sends an entire array of DML operations (INSERT, UPDATE, DELETE) to the SQL engine at once. Eliminates CPU context switches between PL/SQL and SQL engines.", bg: LBL, bd: BL
      },
      {
        t: "🌐  Data Integration Caches",
        b: "Used as temporary in-memory tables to join or aggregate data fetched from external web services before inserting into tables.", bg: LMR, bd: M
      },
      {
        t: "💾  Stored Procedure Parameters",
        b: "Enables client programs (Java, C#) to pass arrays of records to stored procedures, executing batch inserts in a single call.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("The most important use of collections is optimizing database queries using bulk operations. In database programming, the main bottleneck is the context switch between the PL/SQL engine and the SQL engine. The BULK COLLECT statement fetches an entire SQL query result set directly into a collection in a single context switch, which is dramatically faster than looping through rows one-by-one. Similarly, the FORALL statement binds an array to DML queries, sending inserts or updates to the SQL engine in a single batch, eliminating context switch overhead. Collections are also used as in-memory caches or to pass arrays of records to stored procedures, boosting speed up to ten times.");
  }

  const s7 = thankYouSlide(p, H);
  s7.addNotes("That concludes my presentation on Oracle Collections and Arrays. We have seen how using Varrays, Nested Tables, and Associative Arrays, combined with bulk operations, dramatically increases data processing efficiency. I want to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you all for your attention.");

  await p.writeFile({ fileName: "./outputs/Sem1_PLSQL_Unit2_Collections.pptx" });
  convertToPdf("./outputs/Sem1_PLSQL_Unit2_Collections.pptx");
  console.log("✔  PLSQL Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – FUNCTIONS & PROCEDURES  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Database Programming with PL/SQL";

  const s1 = titleSlide(p, H, "Unit III: Functions and Procedures", "Functions & Procedures (Subprograms)");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. In this presentation, I am going to talk about Functions and Procedures in PL/SQL, commonly referred to as subprograms. We will explore the differences between Functions and Procedures, analyze parameter modes including IN, OUT, and IN OUT, study the mechanics of parameter passing comparing pass-by-value and pass-by-reference with the NOCOPY compiler hint, look at various subprogram calling notations, and discuss how subprograms serve as secure APIs in database systems.");

  // Slide 2 ── Introduction to Subprograms
  {
    const s = contentSlide(p, H, "Functions & Procedures (Subprograms)");
    s.addText([
      { text: "Subprograms ", options: { bold: true } },
      { text: "are named PL/SQL blocks stored in the database. They promote modularity, reuse, and security in database applications. The two main types are Functions and Procedures.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Procedures Overview",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Purpose:", options: { bold: true } }, { text: " Performs actions or database modifications (DML).", options: { breakLine: true } },
      { text: "• Return value:", options: { bold: true, color: M } }, { text: " Cannot return a value directly in its header. Can return data using OUT parameters.", options: { breakLine: true } },
      { text: "• Call style:", options: { bold: true } }, { text: " Called as a standalone statement:\n   execute_payment(acc_id, amount);", options: { breakLine: true } },
      { text: "• Usage:", options: { bold: true } }, { text: " Standard block for transactional processes (inserts, updates, audits).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Functions Overview",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Purpose:", options: { bold: true } }, { text: " Calculates and returns a single data value.", options: { breakLine: true } },
      { text: "• Return value:", options: { bold: true, color: DGN } }, { text: " Must contain a RETURN statement in its header and execution path.", options: { breakLine: true } },
      { text: "• Call style:", options: { bold: true } }, { text: " Called as part of an expression:\n   net_sal := salary + calc_bonus(emp_id);", options: { breakLine: true } },
      { text: "• SQL Integration:", options: { bold: true } }, { text: " Can be called directly inside SQL queries:\n   SELECT name, calc_bonus(id) FROM employees;", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Subprograms are named PL/SQL blocks stored in the database schema. They promote modularity, code reuse, and security. The two main types are Procedures and Functions. Procedures are designed to perform actions or database DML modifications. They do not return a value directly in their header, but can return multiple datasets using OUT parameters. They are invoked as standalone statements, such as execute payment. Functions, on the other hand, calculate and return a single data value. They must have a RETURN statement in their header and execution path. Functions are called as part of expressions and can be integrated directly inside SQL queries, like select name and calc bonus from employees.");
  }

  // Slide 3 ── Parameter Modes
  {
    const s = contentSlide(p, H, "Parameter Modes: IN, OUT, and IN OUT");
    s.addText("Parameters define how data is passed into and out of stored subprograms:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    const modes = [
      {
        t: "1. IN Mode (Default)",
        b: "• Read-Only: The subprogram cannot modify the parameter's value.\n• Pass-by-value: Passes a copy of the actual parameter.\n• Supports constants, variables, or expressions as input arguments.", bg: LGN, bd: GN
      },
      {
        t: "2. OUT Mode",
        b: "• Write-Only: The variable is initialized to NULL inside the block. Used to send data back to caller.\n• Must pass a variable, not a literal or expression.\n• Subprogram must assign a value before exiting.", bg: LBL, bd: BL
      },
      {
        t: "3. IN OUT Mode",
        b: "• Read-Write: Passes an initialized value into the subprogram, modifies it, and returns the updated value to the caller.\n• Must pass a variable.\n• Ideal for updating records in place.", bg: LYL, bd: AM
      }
    ];

    modes.forEach((md, i) => {
      const bx = CX + i * 4.15;
      box(p, s, bx, 2.30, 4.00, 4.60, md.bg, md.bd);
      boxTitle(s, bx, 2.30, 4.00, md.t, md.bd);
      s.addText(md.b,
        { x: bx + 0.12, y: 2.30 + 0.52, w: 4.00 - 0.24, h: 4.60 - 0.58, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });
    });

    s.addNotes("Parameters define how data flows between the caller and the subprogram. PL/SQL supports three parameter modes. First, IN mode, which is the default. IN parameters are read-only inside the block; the subprogram cannot modify their values. Second, OUT mode, which is write-only. When passed, the variable is initialized to NULL inside the block and is used to return data back to the caller. We must pass a variable, not a literal. Third, IN OUT mode, which is read-write. It passes an initialized value into the block, allows modification, and returns the updated value to the caller. Both OUT and IN OUT require passing a variable to receive the return value.");
  }

  // Slide 4 ── Pass-by-Value vs. Pass-by-Reference & NOCOPY
  {
    const s = contentSlide(p, H, "Pass-by-Value vs. Pass-by-Reference & NOCOPY");
    s.addText("Understanding parameter passing mechanics and performance tuning:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Default Pass-by-Value Mechanics",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Copying Values:", options: { bold: true } }, { text: " By default, OUT and IN OUT parameters are passed by value. Oracle creates a copy of the argument in memory.", options: { breakLine: true } },
      { text: "• Transaction Safety:", options: { bold: true } }, { text: " If the subprogram fails with an unhandled exception, the copy is discarded and the original variable remains unchanged (clean rollback).", options: { breakLine: true } },
      { text: "• Performance Cost:", options: { bold: true, color: M } }, { text: " Copying large collections (e.g. nested tables of 100,000 rows) causes severe CPU and memory overhead.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("The NOCOPY Hint (Pass-by-Reference)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• The Syntax:", options: { bold: true, color: DGN } }, { text: " Request compiler to pass a pointer directly, skipping the copy:\n   PROCEDURE process_data(data IN OUT NOCOPY my_collection);", options: { breakLine: true } },
      { text: "• Performance Boost:", options: { bold: true } }, { text: " Speeds up execution dramatically when passing large collections, arrays, and objects.", options: { breakLine: true } },
      { text: "• The Danger (Side Effect):", options: { bold: true, color: M } }, { text: " Since the subprogram edits the original variable directly, if an error occurs mid-execution, changes remain in the variable (no automatic rollback).", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let's examine how parameters are passed. By default, IN parameters are passed by reference, but OUT and IN OUT parameters are passed by value. This means Oracle creates a copy of the argument in memory. While this ensures transaction safety—because if the subprogram fails, the copy is discarded and the original variable is unchanged—it introduces a performance cost when passing large collections like nested tables of thousands of rows. To avoid this overhead, we can use the NOCOPY hint. This requests the compiler to pass parameters by reference using a pointer. This speeds up execution, but has the side effect that if an exception occurs mid-run, changes to the variable remain, as there is no automatic rollback.");
  }

  // Slide 5 ── Calling Notations
  {
    const s = contentSlide(p, H, "Subprogram Calling Notations");
    s.addText("PL/SQL supports multiple syntaxes to bind actual arguments to formal parameters:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LYL, AM);
    s.addText("Syntax Code Comparison",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Courier New", margin: 0 });
    s.addText([
      { text: "-- Assume procedure definition:\n", options: { color: GY } },
      { text: "PROCEDURE create_user(p_id NUMBER, p_name VARCHAR2, p_role VARCHAR2 DEFAULT 'Guest');\n\n", options: { color: BK, bold: true } },
      { text: "1. Positional Notation (Order based):\n", options: { bold: true, color: M } },
      { text: "   create_user(101, 'Srinivas', 'Admin');  -- Arguments must match parameter positions exactly.\n\n", options: { color: BK } },
      { text: "2. Named Notation (Association based):\n", options: { bold: true, color: M } },
      { text: "   create_user(p_name => 'Srinivas', p_id => 101, p_role => 'Admin');  -- Order does not matter.\n\n", options: { color: BK } },
      { text: "3. Mixed Notation:\n", options: { bold: true, color: M } },
      { text: "   create_user(101, p_name => 'Srinivas', p_role => 'Admin');  -- Positional first, then named.\n\n", options: { color: BK } },
      { text: "4. Default Parameter Omission:\n", options: { bold: true, color: M } },
      { text: "   create_user(101, 'Srinivas');  -- Omitted parameters resolve to their DEFAULT values.", options: { color: BK } }
    ], { x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13, fontFace: "Courier New", color: BK, margin: 0 });

    s.addNotes("When calling subprograms, PL/SQL supports multiple notations to bind arguments to parameters. Positional notation is the standard method, binding arguments by their position order. Named notation uses the arrow operator to explicitly associate parameter names with values, allowing arguments to be passed in any order and improving readability. Mixed notation combines both, specifying positional arguments first, followed by named arguments. Finally, default parameter omission allows us to omit arguments that have a default value defined in the subprogram header, in which case Oracle automatically resolves them to their default assignments.");
  }

  // Slide 6 ── Applications of Subprograms
  {
    const s = contentSlide(p, H, "Applications of Subprograms");
    s.addText("Subprograms organize logic, secure tables, and improve database performance:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🛡️  Security & Access Control",
        b: "Rather than granting users direct read/write permissions on tables, developers grant access to execute stored procedures. This prevents direct table manipulation and SQL injection.", bg: LGN, bd: GN
      },
      {
        t: "🌐  Business Logic Centralization",
        b: "Keeps business calculations (e.g., tax math, shipping fees) in stored database functions. All client apps (web, mobile, BI) query the same function, preventing logic duplication.", bg: LBL, bd: BL
      },
      {
        t: "⚡  Performance Tuning",
        b: "Subprograms are compiled once and stored in the Shared Pool memory of Oracle Database. Subsequent calls execute immediately without parsing overhead, boosting speed.", bg: LMR, bd: M
      },
      {
        t: "📊  API Layer Integration",
        b: "Serves as an API layer for applications. Standardizing database interactions inside procedures simplifies schema modifications and system updates.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("What are the core applications of subprograms? First, they enforce security. Instead of granting users direct permissions on tables, developers grant access to execute stored procedures, preventing direct table manipulation. Second, they centralize business rules. Storing tax calculations inside database functions ensures all client apps access the same logic, avoiding duplication. Third, they optimize performance. Subprograms are compiled once and stored in Oracle's shared pool memory, executing immediately without parse overhead. Lastly, they act as an API layer, isolating database schema modifications from the frontend applications.");
  }

  const s7 = thankYouSlide(p, H);
  s7.addNotes("That concludes my presentation on Functions and Procedures in PL/SQL. We have seen how stored subprograms enhance modularity, optimize performance, and secure transactions at the database level. I would like to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you for your time.");

  await p.writeFile({ fileName: "./outputs/Sem1_PLSQL_Unit3_Subprograms.pptx" });
  convertToPdf("./outputs/Sem1_PLSQL_Unit3_Subprograms.pptx");
  console.log("✔  PLSQL Unit 3 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 4 – PL/SQL PACKAGES  (Unit IV)
// ═══════════════════════════════════════════════════════════════
async function ppt4() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Database Programming with PL/SQL";

  const s1 = titleSlide(p, H, "Unit IV: Packages", "PL/SQL Packages");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. I am a first-year, first-semester M.Tech student in Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Today, I am going to present on Unit Four of Database Programming with PL/SQL, which covers PL/SQL Packages. In this session, we will look at the definition of packages, the two-part architecture of specification and body, the differences between Definer Rights and Invoker Rights, how packages are monitored in the database catalog, the timestamp versus signature validation methods, and their primary applications in enterprise systems.");

  // Slide 2 ── What is a PL/SQL Package?
  {
    const s = contentSlide(p, H, "What is a PL/SQL Package?");
    s.addText([
      { text: "A PL/SQL Package ", options: { bold: true } },
      { text: "is a schema object that groups logically related PL/SQL types, variables, constants, subprograms, and cursors. It consists of two distinct components compiled separately.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Two-Part Architecture",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Package Specification (Spec):", options: { bold: true, color: BL } }, { text: " The public interface / header. Declares public variables, cursors, types, and subprogram signatures. Visible to other database programs.", options: { breakLine: true } },
      { text: "2. Package Body:", options: { bold: true, color: M } }, { text: " The private implementation. Contains the code for subprograms declared in the specification, as well as private variables, types, and helper methods. Hidden from outer scope.", options: { breakLine: true } },
      { text: "• Separate Compilation:", options: { bold: true } }, { text: " You can change and re-compile the body without compiling the specification, preventing invalidation of dependent objects.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Core Benefits of Packages",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Modular Development:", options: { bold: true } }, { text: " Groups related business features into clean libraries (e.g., HR_PKG, FINANCE_PKG).", options: { breakLine: true } },
      { text: "• Encapsulation (Information Hiding):", options: { bold: true } }, { text: " Keeps helper functions and variables private inside the body, protecting the API implementation.", options: { breakLine: true } },
      { text: "• Session Persistent State:", options: { bold: true, color: DGN } }, { text: " Package variables persist for the duration of a user's database session, acting as in-memory global state.", options: { breakLine: true } },
      { text: "• High Performance:", options: { bold: true } }, { text: " The entire package is loaded into memory on the first call, reducing database page reads for subsequent subprogram runs.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us start by defining what a PL/SQL package is. A package is a schema object that groups logically related PL/SQL types, variables, constants, subprograms, and cursors. As shown on the slide, it uses a two-part architecture. The Package Specification serves as the public interface, declaring public variables, cursors, types, and subprogram signatures that are visible to other database programs. The Package Body, on the other hand, contains the private implementation and actual code of these subprograms, which remains hidden from the outer scope. A major benefit of this separate compilation is that we can change and re-compile the body without compiling the specification, which prevents the invalidation of dependent objects. Overall, packages support modular development, protect implementation details through encapsulation, maintain session-persistent state in memory, and improve performance by loading the entire package on the first call.");
  }

  // Slide 3 ── Definer vs. Invoker Rights
  {
    const s = contentSlide(p, H, "Definer Rights vs. Invoker Rights");
    s.addText("Packages run under different security context privileges, controlled by AUTHID:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, GN);
    s.addText("Definer Rights (AUTHID DEFINER)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Default Privilege Behavior:", options: { bold: true } }, { text: " Run with the schema privileges of the user who CREATED (defined) the package.", options: { breakLine: true } },
      { text: "• Security Encapsulation:", options: { bold: true, color: GN } }, { text: " Allows a user to perform modifications on tables they cannot access directly, by calling a procedure owned by a privileged creator.", options: { breakLine: true } },
      { text: "• Name Resolution:", options: { bold: true } }, { text: " Unqualified table names resolve to the definer's schema.", options: { breakLine: true } },
      { text: "• Execution context:", options: { bold: true } }, { text: " Ideal for secure, transactional routines where table data modifications are highly regulated.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Invoker Rights (AUTHID CURRENT_USER)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Dynamic Privileges:", options: { bold: true } }, { text: " Run with the schema privileges of the user currently executing (invoking) the package.", options: { breakLine: true } },
      { text: "• Name Resolution:", options: { bold: true, color: M } }, { text: " Unqualified table references resolve to the invoker's current schema. The same package query fetches different data depending on who runs it.", options: { breakLine: true } },
      { text: "• Code sharing:", options: { bold: true } }, { text: " Ideal for utility packages (e.g. string formatting, mathematical tools) shared across multiple schemas.", options: { breakLine: true } },
      { text: "• Security constraint:", options: { bold: true } }, { text: " Prevents users from escalating privileges using malicious package code.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Now, let's discuss package security contexts, specifically Definer Rights versus Invoker Rights, controlled by the AUTHID clause. Definer Rights, declared using AUTHID DEFINER, is the default behavior. Under this model, the package runs with the schema privileges of the user who created it. This provides strong security encapsulation because it allows less-privileged users to perform modifications on tables they cannot access directly by executing a procedure owned by a privileged creator. Unqualified table names resolve to the definer's schema. In contrast, Invoker Rights, declared using AUTHID CURRENT_USER, executes the package with the privileges of the active caller. References to unqualified tables resolve dynamically to the caller's schema, meaning the same query fetches different data depending on who runs it. This is ideal for sharing utility packages across different schemas while preventing privilege escalation.");
  }

  // Slide 4 ── Package Catalog & Dependencies
  {
    const s = contentSlide(p, H, "Managing Packages: Invalidation & Catalog");
    s.addText("How Oracle monitors package status, invalidation, and dependencies:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Database Catalog Queries",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Finding Packages:", options: { bold: true } }, { text: " Metadata is stored in catalog views like USER_OBJECTS and USER_SOURCE:\n   SELECT object_name, status FROM user_objects WHERE object_type = 'PACKAGE';", options: { breakLine: true } },
      { text: "• Status Monitoring:", options: { bold: true } }, { text: " Status can be 'VALID' or 'INVALID'. If a dependent table changes, the package becomes invalid and must be re-compiled.", options: { breakLine: true } },
      { text: "• Describing Packages:", options: { bold: true } }, { text: " Use DESCRIBE command in SQL*Plus to check public signatures:\n   DESCRIBE hr_pkg;", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LMR, M);
    s.addText("Timestamp vs. Signature Validation",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "When a package is modified, Oracle determines if dependent programs must be recompiled using two methods:", options: { breakLine: true } },
      { text: "1. Timestamp Method (Default):", options: { bold: true, color: M } }, { text: " Compares compile timestamps. If the package spec changes, all calling programs are immediately marked invalid and must re-compile. Causes cascading invalidations.", options: { breakLine: true } },
      { text: "2. Signature Method:", options: { bold: true, color: DGN } }, { text: " Compares subprogram parameters and data types (signatures). Calling programs remain valid as long as parameter signatures are unchanged, avoiding rebuild loops.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us examine how Oracle manages packages and dependencies. Package metadata and statuses are stored in catalog views like USER_OBJECTS and USER_SOURCE. A query on USER_OBJECTS, as shown on the slide, returns whether a package is VALID or INVALID. If a dependent table changes, Oracle marks the package as invalid. When recompiling, Oracle uses two methods: Timestamp and Signature. The default Timestamp method compares compilation times; if a package specification changes, all calling programs are marked invalid immediately, which can cause massive cascading invalidations. The Signature method compares subprogram parameter signatures and data types instead. Calling programs remain valid as long as parameter signatures do not change, which significantly reduces compilation overhead and avoids rebuild loops in production environments.");
  }

  // Slide 5 ── Applications of PL/SQL Packages
  {
    const s = contentSlide(p, H, "Applications of Packages");
    s.addText("PL/SQL packages are utilized to build modular, secure database interfaces:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📦  Enterprise business API libraries",
        b: "Groups related workflows into single libraries (e.g. Ledger_Pkg, Inventory_Pkg). Standardizes how client software executes transactions, enforcing business rules.", bg: LGN, bd: GN
      },
      {
        t: "💾  Session Cache Managers",
        b: "Package variables persist throughout a database connection. They are utilized to cache static data (e.g., system settings, user permissions) in memory, avoiding redundant table queries.", bg: LBL, bd: BL
      },
      {
        t: "🌐  System Integration Modules",
        b: "Encapsulates complex database integrations, such as calling external APIs, parsing XML/JSON, or managing TCP connections inside private package body routines.", bg: LMR, bd: M
      },
      {
        t: "⚙️  Utility Toolkits",
        b: "Groups helper subprograms for data mapping, error logging, and date arithmetic. Oracle built-in packages like DBMS_SQL and UT_FILE are classic examples.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("To summarize, PL/SQL packages have vital applications in database development. They act as enterprise business API libraries, grouping related workflows like finance or HR into single repositories to standardize transactions. They also function as session cache managers. Because package variables persist throughout a database connection, they cache static data in memory, avoiding redundant table queries. Furthermore, packages encapsulate complex integrations, such as calling external APIs or parsing JSON, keeping body routines private. Finally, they serve as utility toolkits, grouping common helper subprograms for data mapping and error logging, similar to Oracle's built-in packages like DBMS_SQL.");
  }

  const s6 = thankYouSlide(p, H);
  s6.addNotes("That brings me to the end of my presentation on PL/SQL Packages. We have seen how packaging related components improves modularity, security, and performance. I would like to express my sincere gratitude to the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management for this opportunity. Thank you all for your time and attention.");

  await p.writeFile({ fileName: "./outputs/Sem1_PLSQL_Unit4_Packages.pptx" });
  convertToPdf("./outputs/Sem1_PLSQL_Unit4_Packages.pptx");
  console.log("✔  PLSQL Unit 4 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 5 – DATABASE TRIGGERS  (Unit V)
// ═══════════════════════════════════════════════════════════════
async function ppt5() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Database Programming with PL/SQL";

  const s1 = titleSlide(p, H, "Unit V: Triggers", "Database Triggers");
  s1.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and my roll number is 257Y1D5805. I am a first-year, first-semester M.Tech student in Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Today, I am going to present on Unit Five of Database Programming with PL/SQL, focusing on Database Triggers. In this presentation, we will explore what a trigger is, its core characteristics, the classifications of triggers, the differences between row-level and statement-level triggers, the timing and execution order of trigger blocks, and their common enterprise applications.");

  // Slide 2 ── What is a Trigger?
  {
    const s = contentSlide(p, H, "What is a Database Trigger?");
    s.addText([
      { text: "A Database Trigger ", options: { bold: true } },
      { text: "is a stored PL/SQL block that automatically executes (fires) in response to specific events in the database (DML modifications, DDL statements, or database system events).", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Characteristics",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Implicit Execution:", options: { bold: true } }, { text: " Fired automatically by the database engine. Cannot be called manually by user queries.", options: { breakLine: true } },
      { text: "• Event-Driven:", options: { bold: true } }, { text: " Bound directly to database events. Fires immediately when the triggering SQL statement executes.", options: { breakLine: true } },
      { text: "• Transaction Context:", options: { bold: true } }, { text: " Runs within the same transaction scope as the DML statement. A failure inside a trigger rolls back the entire transaction.", options: { breakLine: true } },
      { text: "• High Risk:", options: { bold: true, color: M } }, { text: " Poorly designed triggers can degrade database performance and cause unexpected side effects (mutating tables).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Trigger Classification",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• DML Triggers:", options: { bold: true, color: BL } }, { text: " Bound to INSERT, UPDATE, or DELETE statements on tables or views.", options: { breakLine: true } },
      { text: "• DDL Triggers:", options: { bold: true, color: M } }, { text: " Bound to schema modifications (CREATE, ALTER, DROP). Used to audit operations or restrict changes.", options: { breakLine: true } },
      { text: "• Database System Triggers:", options: { bold: true, color: AM } }, { text: " Bound to database events (STARTUP, SHUTDOWN, SERVERERROR, LOGON, LOGOFF). Ideal for tracking connections.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin with the definition of a database trigger. A trigger is a stored PL/SQL block that automatically executes, or fires, in response to specific events, such as DML modifications, DDL statements, or database system events. Triggers have several core characteristics. They execute implicitly; they are fired automatically by the database engine and cannot be called manually by user queries. They are event-driven, firing immediately when the triggering SQL statement executes. They also run within the transaction context of the DML statement. If a trigger encounters an unhandled exception, it rolls back the entire transaction. However, they carry high risk, as poorly designed triggers can degrade database performance or cause mutating table errors. In terms of classification, we have DML triggers on tables or views, DDL triggers on schema modifications, and system triggers on events like startup, shutdown, logon, or logoff.");
  }

  // Slide 3 ── Row-Level vs. Statement-Level Triggers
  {
    const s = contentSlide(p, H, "Row-Level vs. Statement-Level DML Triggers");
    s.addText("Understanding the granularity of DML triggers determines how often they fire:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Statement-Level Triggers",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Execution Frequency:", options: { bold: true } }, { text: " Fires exactly ONCE per DML statement, regardless of the number of rows affected (even if 0 rows are updated).", options: { breakLine: true } },
      { text: "• Syntax:", options: { bold: true } }, { text: " This is the default trigger behavior. Do not use 'FOR EACH ROW' clause.", options: { breakLine: true } },
      { text: "• Variables:", options: { bold: true, color: M } }, { text: " Cannot access individual row values. No access to ':OLD' or ':NEW' record values.", options: { breakLine: true } },
      { text: "• Best Use Case:", options: { bold: true } }, { text: " Enforcing global business restrictions (e.g. prevent table updates outside business hours).", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Row-Level Triggers",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Execution Frequency:", options: { bold: true } }, { text: " Fires ONCE FOR EACH ROW affected by the DML statement. If a query updates 100 rows, the trigger fires 100 times.", options: { breakLine: true } },
      { text: "• Syntax:", options: { bold: true } }, { text: " Must include the 'FOR EACH ROW' clause in trigger header.", options: { breakLine: true } },
      { text: "• Pseudo-records:", options: { bold: true, color: DGN } }, { text: " Has access to row values:\n   - :OLD points to values before query.\n   - :NEW points to values after query.", options: { breakLine: true } },
      { text: "• Best Use Case:", options: { bold: true } }, { text: " Auditing cell value changes, generating unique primary keys dynamically.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Next, let's differentiate between statement-level and row-level DML triggers. A statement-level trigger fires exactly once per SQL statement, regardless of the number of rows affected. This is the default behavior when we omit the FOR EACH ROW clause. Because statement-level triggers execute once, they cannot access individual row values, and the OLD and NEW pseudo-records are unavailable. They are ideal for enforcing global business restrictions, such as preventing table updates outside business hours. In contrast, a row-level trigger fires once for each row affected by the statement. We must include the FOR EACH ROW clause in its header. Row-level triggers have access to OLD and NEW pseudo-records, representing row values before and after the query, which makes them perfect for auditing column changes and generating primary keys.");
  }

  // Slide 4 ── Trigger Timing & Architecture
  {
    const s = contentSlide(p, H, "Trigger Timing and Execution Order");
    s.addText("Timings control when the trigger fires relative to the DML event:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Trigger Timings",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• BEFORE Trigger:", options: { bold: true, color: BL } }, { text: " Fires before DML execution. Used to validate or alter :NEW values before they are written to disk.", options: { breakLine: true } },
      { text: "• AFTER Trigger:", options: { bold: true, color: DGN } }, { text: " Fires after DML execution. Used to perform post-processing actions (like updating related tables or writing audit records). Cannot modify row values.", options: { breakLine: true } },
      { text: "• INSTEAD OF Trigger:", options: { bold: true, color: AM } }, { text: " Fires INSTEAD OF the triggering DML. Used primarily on non-updatable database views to route data to baseline tables.", options: { breakLine: true } },
      { text: "• Compound Trigger:", options: { bold: true } }, { text: " Combines BEFORE/AFTER and statement/row timings in a single block, sharing local variables to prevent mutating table errors.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Execution Order Flow",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "When a statement modifies multiple rows, Oracle executes steps in this sequence:", options: { breakLine: true } },
      { text: "1. Execute BEFORE statement-level triggers.", options: { breakLine: true } },
      { text: "2. For each row affected:\n", options: { bold: true } },
      { text: "   a. Execute BEFORE row-level triggers.\n", options: { color: BK } },
      { text: "   b. Apply DML modifications to the row.\n", options: { color: BK } },
      { text: "   c. Execute AFTER row-level triggers.\n", options: { color: BK } },
      { text: "3. Execute AFTER statement-level triggers.", options: { breakLine: true } },
      { text: "• Constraints:", options: { bold: true, color: M } }, { text: " Triggers cannot execute transaction commands (COMMIT, ROLLBACK) directly. Trigger size limit is 32KB.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us explore trigger timings and execution flow. Timings control when a trigger fires. BEFORE triggers run before the DML statement, validating or altering values before they are written to disk. AFTER triggers run after execution, performing post-processing tasks like writing audit records. INSTEAD OF triggers run on views, routing data to base tables. A Compound Trigger combines all these timings in one block, sharing local variables to prevent mutating table errors. In terms of execution order, when a query modifies multiple rows, Oracle first runs BEFORE statement-level triggers, then loop through each row to execute BEFORE row-level triggers, apply the DML, and run AFTER row-level triggers. Finally, it runs AFTER statement-level triggers. Note that triggers cannot execute transaction commands like COMMIT or ROLLBACK directly.");
  }

  // Slide 5 ── Applications of Database Triggers
  {
    const s = contentSlide(p, H, "Applications of Triggers");
    s.addText("Database triggers automate safety, auditing, and synchronization tasks:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "📝  Audit Logging & Tracking",
        b: "Row-level AFTER triggers capture changes to sensitive columns (e.g. salaries) and write old and new values to log tables with timestamps and username details.", bg: LGN, bd: GN
      },
      {
        t: "🔑  Auto-Generating Primary Keys",
        b: "BEFORE INSERT row triggers fetch sequence numbers (or UUIDs) and assign them to the :NEW.id field automatically, ensuring unique identifiers.", bg: LBL, bd: BL
      },
      {
        t: "🛡️  DDL Change Auditing",
        b: "DDL triggers track alterations to database tables (e.g. dropping columns) and log schema modifications or prevent changes during release locks.", bg: LMR, bd: M
      },
      {
        t: "🔄  Data Sync & Replication",
        b: "Synchronizes replicated tables or denormalized caches automatically when baseline records are altered, keeping data consistent.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Triggers automate critical tasks at the database level. Their primary applications include audit logging, where AFTER row-level triggers write old and new values of sensitive columns to log tables. They are also used for auto-generating primary keys, where BEFORE INSERT triggers fetch sequence numbers and assign them to the ID column. Additionally, they audit DDL modifications, tracking when schemas are modified or preventing alterations during release locks. Lastly, triggers support data synchronization, updating denormalized tables or baseline records automatically to maintain consistency across the system.");
  }

  const s6 = thankYouSlide(p, H);
  s6.addNotes("That concludes my presentation on Database Triggers. We have learned how triggers enforce security, audit changes, and ensure integrity implicitly. I would like to thank the Department of Computer Science and Engineering at the Marri Laxman Reddy Institute of Technology and Management. Thank you very much for your attention.");

  await p.writeFile({ fileName: "./outputs/Sem1_PLSQL_Unit5_Triggers.pptx" });
  convertToPdf("./outputs/Sem1_PLSQL_Unit5_Triggers.pptx");
  console.log("✔  PLSQL Unit 5 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 5 PL/SQL PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  await ppt4();
  await ppt5();
  console.log("All 5 PL/SQL PPTs generated successfully!");
})();
