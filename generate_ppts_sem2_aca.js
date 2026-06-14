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

async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Computer Architecture";

  const tSl = titleSlide(p, H, "Unit I: Theory of Parallelism", "Parallel Models & Conditions of Parallelism");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Parallel Models & Conditions of Parallelism' as part of the Advanced Computer Architecture course for my M.Tech in Computer Science and Engineering. In this presentation, we will explore the theory of parallel processing models, discuss Bernstein's conditions of parallelism, examine program dependency structures, and look at practical applications of these models in high-performance computing systems.");

  // Slide 2 ── Theory of Parallelism
  {
    const s = contentSlide(p, H, "Theory of Parallelism & Models");
    s.addText([
      { text: "Parallel Computing ", options: { bold: true } },
      { text: "aims to execute multiple calculations simultaneously. In computer architecture, we analyze parallelism using standardized parallel models to design efficient hardware configurations.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Parallel Computer Models",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Multiprocessors (Shared Memory):", options: { bold: true, color: DGN } }, { text: " Multiple CPUs sharing a single physical memory address space (e.g. SMP systems).", options: { breakLine: true } },
      { text: "• Multicomputers (Distributed Memory):", options: { bold: true, color: DGN } }, { text: " Nodes with private memory communicating via message passing (e.g. clusters).", options: { breakLine: true } },
      { text: "• SIMD Computers:", options: { bold: true } }, { text: " Single Instruction, Multiple Data. One controller executes a single command on vectors of data (e.g. GPUs).", options: { breakLine: true } },
      { text: "• PRAM Model:", options: { bold: true } }, { text: " Parallel Random Access Machine. A theoretical model to analyze parallel algorithm complexities.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Program Partitioning & Flow",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Parallelism Granularity:", options: { bold: true } }, { text: " Sizes of execution blocks. Ranges from fine-grained (instruction level) to coarse-grained (process level).", options: { breakLine: true } },
      { text: "• Static Scheduling:", options: { bold: true } }, { text: " Compiler assigns instructions to cores at compilation time. Low runtime overhead.", options: { breakLine: true } },
      { text: "• Dynamic Scheduling:", options: { bold: true } }, { text: " Hardware assigns tasks at runtime based on resource availability, handling hazards dynamically.", options: { breakLine: true } },
      { text: "• Control vs. Data Flow:", options: { bold: true, color: M } }, { text: " Control-flow: sequential, instruction-pointer driven. Data-flow: instructions execute as soon as their operands are ready.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us start by understanding the theory of parallelism and the primary models used in high-performance computing. Parallel computing is the practice of dividing a computational workload into smaller, independent tasks that can be executed simultaneously by multiple processing units. In hardware design, we classify parallel computers into three main models. First, Shared Memory Multiprocessors, where multiple processing units share a single, global physical address space, allowing communication via shared variables. Second, Distributed Memory Multicomputers, where each processor node has its own private memory and communicates via network message passing. Third, SIMD computers—Single Instruction, Multiple Data—where a single control unit broadcasts a single instruction to multiple vector processing elements, which is the model used by modern graphics processing units, or GPUs. In computer architecture, we also model execution theoretics using the Parallel Random Access Machine, or PRAM, model, and partition programs according to granularity—ranging from fine-grained instruction levels to coarse-grained process levels.");
  }

  // Slide 3 ── Bernstein's Conditions
  {
    const s = contentSlide(p, H, "Bernstein's Conditions of Parallelism");
    s.addText("A set of algebraic conditions used to determine if two software statements can execute in parallel:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Mathematical Formulation",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Let I1 and I2 be two program statements.\n", options: { breakLine: true } },
      { text: "• Read Set (R_i):", options: { bold: true } }, { text: " Set of all variables read by statement I_i.\n", options: { breakLine: true } },
      { text: "• Write Set (W_i):", options: { bold: true } }, { text: " Set of all variables modified by statement I_i.\n\n", options: { breakLine: true } },
      { text: "Statements I1 and I2 can execute in parallel (I1 || I2) if and only if they satisfy three conditions:\n", options: { bold: true } },
      { text: "1.  R1 ∩ W2 = ∅  (No Flow Dependency)\n", options: { bold: true, color: M } },
      { text: "2.  W1 ∩ R2 = ∅  (No Anti-Dependency)\n", options: { bold: true, color: M } },
      { text: "3.  W1 ∩ W2 = ∅  (No Output Dependency)", options: { bold: true, color: M } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Dependency Tracing Example",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Courier New", margin: 0 });
    s.addText([
      { text: "-- Assume two statements:\n", options: { color: GY } },
      { text: "I1: A := B + C;\n", options: { color: BK, bold: true } },
      { text: "I2: D := A * E;\n\n", options: { color: BK, bold: true } },
      { text: "• Read/Write Sets:\n", options: { color: BK, bold: true } },
      { text: "  R1 = {B, C}       W1 = {A}\n", options: { color: BK } },
      { text: "  R2 = {A, E}       W2 = {D}\n\n", options: { color: BK } },
      { text: "• Check Bernstein's Rules:\n", options: { color: BK, bold: true } },
      { text: "  R1 ∩ W2 = ∅   ✔\n", options: { color: DGN } },
      { text: "  W1 ∩ R2 = {A} ✘ (Flow Dependency)\n", options: { color: M, bold: true } },
      { text: "  W1 ∩ W2 = ∅   ✔\n\n", options: { color: DGN } },
      { text: "• Conclusion: I1 and I2 CANNOT run in parallel. I2 must wait for I1.", options: { bold: true, color: M } }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Courier New", color: BK, margin: 0 });

    s.addNotes("Bernstein's conditions are the fundamental mathematical rules used to determine whether two program statements or blocks can execute in parallel without producing data conflicts. Let us consider two statements, statement one and statement two. We define input sets I-one and I-two as the variables read by the statements, and output sets O-one and O-two as the variables written to by the statements. For these two statements to execute in parallel, three conditions must be satisfied. First, the input set of statement one must not intersect with the output set of statement two. This prevents a Read-After-Write dependency, also known as flow dependency. Second, the output set of statement one must not intersect with the input set of statement two. This prevents a Write-After-Read dependency, also known as anti-dependency. Third, the output sets of statement one and statement two must not intersect. This prevents a Write-After-Write dependency, also known as output dependency. If any of these intersections are not empty, the statements must be executed sequentially to ensure correctness.");
  }

  // Slide 4  ── Hardware Interconnection Networks
  {
    const s = contentSlide(p, H, "System Interconnect Architectures");
    s.addText("Parallel systems use different network topologies to route data between processors and memory:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    const nets = [
      {
        t: "1. Shared Bus Network",
        b: "• Description: A single shared communication line connects all processors and memory modules.\n• Cost: Very low.\n• Bottleneck: Serious bandwidth congestion as node count scales past 8-16 processors.", bg: LGN, bd: GN
      },
      {
        t: "2. Crossbar Switch",
        b: "• Description: A matrix of grid switches connecting every processor directly to every memory module.\n• Performance: Maximum bandwidth; zero blocking conflicts.\n• Cost: O(N²) switches — very expensive for large systems.", bg: LBL, bd: BL
      },
      {
        t: "3. Multistage Networks (MINs)",
        b: "• Description: Uses stages of small 2x2 switches (like Omega or Butterfly networks) to route packets.\n• Trade-off: Balanced cost O(N log N) and good scaling throughput. Standard in supercomputer grids.", bg: LYL, bd: AM
      }
    ];

    nets.forEach((nt, i) => {
      const bx = CX + i * 4.15;
      box(p, s, bx, 2.30, 4.00, 4.60, nt.bg, nt.bd);
      boxTitle(s, bx, 2.30, 4.00, nt.t, nt.bd);
      s.addText(nt.b,
        { x: bx + 0.12, y: 2.30 + 0.52, w: 4.00 - 0.24, h: 4.60 - 0.58, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });
    });

    s.addNotes("To analyze how instructions interact at runtime, compilers construct dependency graphs. There are five major types of program dependencies we must consider. Data dependency occurs when an instruction refers to data produced by a previous instruction. Control dependency is determined by conditional branch decisions, where the execution of one path depends on a prior test. Resource dependency occurs when two independent instructions attempt to use the same hardware unit simultaneously, such as a floating-point multiplier or memory port. Control flow dependency dictates the sequential order of statement execution as written in the program code. Finally, hardware must manage data hazards. A Read-After-Write, or RAW hazard, occurs when an instruction tries to read a source before a previous instruction writes to it. Write-After-Read, or WAR hazards, occur when an instruction tries to write a destination before a previous read. Write-After-Write, or WAW hazards, occur when an instruction tries to write a destination before a previous write completes.");
  }

  // Slide 5 ── Applications
  {
    const s = contentSlide(p, H, "Applications of Parallel Models");
    s.addText("Parallel architectures and dependency models drive modern high-performance systems:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "⚙️  Optimizing Compilers",
        b: "Compilers apply Bernstein's conditions to construct dependency graphs of code instructions. Allows compilers to schedule instructions to run in parallel on multicore CPUs safely.", bg: LGN, bd: GN
      },
      {
        t: "🎮  GPU Vector Processing",
        b: "Modern graphics cards partition image rendering logic into millions of vector threads. Executing them in parallel speeds up real-time 3D games.", bg: LBL, bd: BL
      },
      {
        t: "🧬  Distributed Supercomputing",
        b: "Supercomputers (like weather models) partition computations across thousands of nodes using MPI libraries to solve problems in days instead of years.", bg: LMR, bd: M
      },
      {
        t: "🗃️  Database Query Partitioning",
        b: "Cloud databases partition tables and queries. Running lookups in parallel across nodes decreases query latencies under load.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let us look at how these parallel architecture models are applied in practice. First, symmetric multiprocessing, or SMP servers, use shared-memory models to run database engines, where thread locks and semaphores manage concurrent query access. Second, high-performance computing clusters use distributed-memory message passing protocols like MPI to run scientific simulations across thousands of node cores. Third, modern compiler optimization pipelines analyze Bernstein's conditions to identify parallelizable code blocks during compilation, scheduling instructions to execute concurrently on superscalar execution units. Finally, graphics cards exploit SIMD parallelism to run identical pixel shading instructions across millions of pixels simultaneously.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, the theory of parallelism and Bernstein's conditions form the foundation of both hardware design and compiler optimization. By understanding data, control, and resource dependencies, we can design more efficient processors and algorithms that fully exploit modern multicore systems. Thank you for your time. I am now open to any questions you might have.");
  await p.writeFile({ fileName: "./outputs/Sem2_ACA_Unit1_Parallelism.pptx" });
  convertToPdf("./outputs/Sem2_ACA_Unit1_Parallelism.pptx");
  console.log("✔  ACA Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – AMDAHL'S VS. GUSTAFSON'S LAW  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Computer Architecture";

  const tSl = titleSlide(p, H, "Unit II: Principles of Scalable Performance", "Amdahl's Law vs. Gustafson's Law");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Amdahl's Law vs. Gustafson's Law' as part of the Advanced Computer Architecture course for my M.Tech in Computer Science and Engineering. In this session, we will explore the concepts of scalability and speedup metrics, discuss Amdahl's fixed-workload speedup law, contrast it with Gustafson's scaled-workload speedup law, examine their direct comparisons, and discuss their real-world applications in hardware and software scaling.");

  // Slide 2 ── Scalability and Speedup Metrics
  {
    const s = contentSlide(p, H, "Scalability & Speedup Performance");
    s.addText([
      { text: "Speedup ", options: { bold: true } },
      { text: "measures performance gains from executing code on multiple processors. Scalability evaluates how well speedup scales as both processor count and problem size increase.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Speedup Concepts",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Speedup Formula (S_p):", options: { bold: true } }, { text: " S_p = T_1 / T_p, where T_1 is execution time on 1 processor, and T_p is execution time on p processors.", options: { breakLine: true } },
      { text: "• Ideal Speedup:", options: { bold: true, color: DGN } }, { text: " Linear speedup (S_p = p). Doubling processors should halve the execution time.", options: { breakLine: true } },
      { text: "• Real-World Limits:", options: { bold: true, color: M } }, { text: " Overhead (communication, scheduling, memory latency) and serial portions of code prevent ideal scaling.", options: { breakLine: true } },
      { text: "• Scalability Analysis:", options: { bold: true } }, { text: " Studies how speedup behaves as both processor count and workload size grow together.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Serial Bottlenecks",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• The Serial Fraction (s):", options: { bold: true, color: M } }, { text: " The portion of a program that cannot be parallelized (e.g. initialization, writing output files).", options: { breakLine: true } },
      { text: "• The Parallel Fraction (p):", options: { bold: true } }, { text: " The portion of code that can be executed concurrently across cores (e.g., matrix updates). Note: s + p = 1.", options: { breakLine: true } },
      { text: "• The Challenge:", options: { bold: true } }, { text: " No matter how many processors are added, the time to run the serial fraction 's' remains constant, creating a bottleneck.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us start by understanding the concepts of scalability and speedup, which measure performance gains from parallel computing. Speedup is mathematically defined as the ratio of single-processor execution time to multi-processor execution time. Under ideal conditions, we expect linear speedup—meaning that daily scaling. Crucially, every program contains a serial fraction that cannot be parallelized, such as initialization blocks or file write operations. This serial portion remains constant regardless of the number of cores added, forming a persistent bottleneck. We analyze these performance limitations using two main scalability formulations: Amdahl's Law and Gustafson's Law.");
  }

  // Slide 3 ── Amdahl's Law
  {
    const s = contentSlide(p, H, "Amdahl's Law (Fixed-Workload)");
    s.addText("Gene Amdahl (1967) modeled speedup assuming a FIXED workload size:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Formula & Derivation",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Given serial fraction s and parallel fraction p (with s + p = 1):\n", options: { breakLine: true } },
      { text: "S_p = 1 / (s + p / N)", options: { bold: true, fontSize: 20, color: M, breakLine: true } },
      { text: "  where N is the number of processors.\n\n", options: { fontSize: 13, breakLine: true } },
      { text: "• Limit as N → ∞ (Asymptotic Speedup):\n", options: { bold: true, color: M } },
      { text: "  S_∞ = 1 / s", options: { bold: true, fontSize: 18, color: M, breakLine: true } },
      { text: "  If only 5% of your program is serial (s = 0.05), your MAXIMUM speedup is 1/0.05 = 20, even with an infinite number of processors!", options: { fontSize: 13 } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Implications of Amdahl's Law",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Diminishing Returns:", options: { bold: true } }, { text: " Adding processors yields decreasing speedup gains because the serial fraction dominates execution time.", options: { breakLine: true } },
      { text: "• Critical Bottleneck:", options: { bold: true, color: M } }, { text: " Focus on optimizing the serial fraction to achieve higher speedup rather than simply scaling core counts.", options: { breakLine: true } },
      { text: "• Design Constraint:", options: { bold: true } }, { text: " Tells chip designers that adding hundreds of cores to a CPU is useless if code remains sequential.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Gene Amdahl formulated his law in 1967, assuming a fixed problem workload size. The Amdahl's speedup formula is S-p equals one divided by the quantity of the serial fraction plus the parallel fraction divided by processor count. As the number of processors approaches infinity, the parallel fraction term vanishes, leaving the maximum speedup bounded strictly by one divided by the serial fraction. For example, if only five percent of a program is sequential, your absolute maximum speedup is twenty, even if you deploy an infinite number of processor cores. The implications of Amdahl's law are significant: it shows diminishing returns when adding processors, highlighting that sequential parts of a program represent a critical design constraint. Thus, engineers must focus on optimizing the serial code fraction itself rather than simply scaling core counts.");
  }

  // Slide 4 ── Gustafson's Law
  {
    const s = contentSlide(p, H, "Gustafson's Law (Scaled-Workload)");
    s.addText("John Gustafson (1988) argued that problem sizes scale as computing power increases:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, DGN);
    s.addText("Formula & Derivation",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Gustafson assumed parallel execution time is fixed, and problem size scales to match hardware size:\n", options: { breakLine: true } },
      { text: "S_p = s + N · (1 − s)", options: { bold: true, fontSize: 20, color: DGN, breakLine: true } },
      { text: "  where N is the number of processors, and s is the serial fraction of the scaled workload.\n\n", options: { fontSize: 13, breakLine: true } },
      { text: "• Linear Scaling Concept:", options: { bold: true, color: DGN } }, { text: " Speedup increases linearly with processor count N. E.g., if s = 0.05 (5% serial), on 100 processors:\n   S_100 = 0.05 + 100 · 0.95 = 95.05. A very high speedup is achievable!", options: { fontSize: 13 } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Implications of Gustafson's Law",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Problem Size Scaling:", options: { bold: true } }, { text: " As computing power grows, we don't run the same small problem faster. Instead, we run larger, more complex simulations (e.g. higher-res weather models).", options: { breakLine: true } },
      { text: "• Linear Speedup:", options: { bold: true, color: DGN } }, { text: " If the workload scales, the serial fraction becomes negligible relative to the massive parallel processing volume, allowing linear speedups.", options: { breakLine: true } },
      { text: "• Optimistic View:", options: { bold: true } }, { text: " Proves that massively parallel supercomputer grids are highly effective for large-scale scientific modeling.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("John Gustafson challenged Amdahl's pessimistic view in 1988, arguing that developers scale the problem size as more computing power becomes available. In other words, we do not run the same small problem faster; instead, we run larger, more complex simulations in the same amount of time. Gustafson's scaled-speedup formula is S-p equals the serial fraction plus processor count multiplied by the quantity of one minus the serial fraction. Under Gustafson's law, speedup scales linearly with processor count. E.g., if a program is five percent serial, on one hundred processors it can achieve a speedup of ninety-five. Gustafson's law provides a highly optimistic and realistic view of parallel computing, justifying the construction of massively parallel supercomputers for large-scale applications like high-resolution weather models, where the parallel fraction scales to dwarf the serial overhead.");
  }

  // Slide 5 ── Amdahl vs. Gustafson Comparison
  {
    const s = contentSlide(p, H, "Comparison: Amdahl vs. Gustafson");
    s.addText("A direct comparative summary of the two performance scaling laws:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const tData = [
      [
        { text: "Comparison Metric", options: { bold: true, fill: { color: M }, color: WH } },
        { text: "Amdahl's Law", options: { bold: true, fill: { color: M }, color: WH, align: "center" } },
        { text: "Gustafson's Law", options: { bold: true, fill: { color: M }, color: WH, align: "center" } }
      ],
      [
        { text: "Workload Model", options: { bold: true, fill: { color: LGR } } }, { text: "Fixed Size (does not change)", options: { fill: { color: LGR } } }, { text: "Scaled Size (grows with hardware)", options: { fill: { color: LGR } } }
      ],
      [
        { text: "Speedup Behavior", options: { bold: true, fill: { color: LBL } } }, { text: "Logarithmic curve; bounds speedup to 1/s", options: { fill: { color: LBL } } }, { text: "Linear growth line; no fixed limit", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Hardware Focus", options: { bold: true, fill: { color: LGN } } }, { text: "Optimize core latency (single-thread speed)", options: { fill: { color: LGN } } }, { text: "Maximize throughput (aggregate core counts)", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Typical Application", options: { bold: true, fill: { color: LMR } } }, { text: "Desktop apps, local tools, minor utilities", options: { fill: { color: LMR } } }, { text: "Supercomputers, AI training, weather models", options: { fill: { color: LMR } } }
      ],
      [
        { text: "Primary Goal", options: { bold: true, fill: { color: LYL } } }, { text: "Minimize execution latency", options: { fill: { color: LYL } } }, { text: "Maximize problem accuracy & scope", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: CX, y: 2.20, w: CW, h: 4.80, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [3.0, 4.67, 4.68]
      });

    s.addNotes("Let us compare Amdahl's and Gustafson's laws directly. Amdahl's law assumes a fixed workload model, focusing on minimizing execution latency. Its speedup curve is logarithmic, approaching a fixed asymptote of one divided by the serial fraction. It is typical for local desktop applications. In contrast, Gustafson's law assumes a scaled workload model, aiming to maximize problem accuracy and computational throughput. Its speedup grows linearly with processor count, without a fixed upper bound. Gustafson's law is typical for supercomputers, deep learning training, and climate modeling, where scaling the problem size is the primary goal.");
  }

  // Slide 6 ── Applications
  {
    const s = contentSlide(p, H, "Applications of Speedup Laws");
    s.addText("Understanding speedup laws determines how companies scale hardware and software architectures:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "💻  Multicore CPU Design",
        b: "Chip manufacturers balance single-thread clock speeds (Amdahl's bottleneck) with core counts, optimizing both sequential and parallel fractions.", bg: LGN, bd: GN
      },
      {
        t: "🌐  Cloud Computing & Auto-scaling",
        b: "Cloud platforms auto-scale nodes based on workload. If a task is parallelizable, Gustafson's law guarantees that adding instances scales throughput.", bg: LBL, bd: BL
      },
      {
        t: "🧬  HPC Scientific Simulations",
        b: "Climate models and astrophysics simulators scale grid nodes to increase grid resolution (Gustafson's law), achieving highly accurate models.", bg: LMR, bd: M
      },
      {
        t: "🤖  AI Training & LLM scaling",
        b: "Training massive AI models requires thousands of GPUs. Gustafson's law justifies this scale, as training parameters are scaled to prevent bottlenecks.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let us look at how these speedup laws guide system architecture designs. First, in multi-core CPU designs, chip manufacturers balance single-thread clock speeds, to combat Amdahl's sequential bottleneck, with high core counts to optimize parallel processing. Second, cloud computing platforms utilize Gustafson's law to auto-scale virtual machines, distributing larger incoming traffic batches concurrently. Third, high-performance computing clusters use Gustafson's law to justify scaling grid nodes, allowing simulations of physical processes at higher resolutions. Finally, training deep learning and large language models utilizes thousands of GPUs, scaling batch sizes and parameters to keep the parallel fraction close to one and maximize hardware efficiency.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, both Amdahl's and Gustafson's laws are essential for understanding the limits of parallel system performance. While Amdahl's law warns us of the danger of sequential bottlenecks, Gustafson's law shows that scaling the problem size allows us to exploit massive parallelism effectively. Thank you for your time. I am now open to any questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_ACA_Unit2_Scalability.pptx" });
  convertToPdf("./outputs/Sem2_ACA_Unit2_Scalability.pptx");
  console.log("✔  ACA Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – PIPELINING & HAZARDS  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Computer Architecture";

  const tSl = titleSlide(p, H, "Unit III: Pipelining & Superscalar Designs", "Pipelining & Hazards in Processors");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Linear Pipeline Processors' as part of the Advanced Computer Architecture course for my M.Tech in Computer Science and Engineering. In this presentation, we will explore the concepts of linear pipelining, analyze instruction pipeline stages, discuss the classification of pipeline hazards, examine hardware and compiler solutions for resolving these hazards, and look at practical applications of pipelined processors in modern computing systems.");

  // Slide 2 ── What is Pipelining?
  {
    const s = contentSlide(p, H, "What is Pipelining?");
    s.addText([
      { text: "Pipelining ", options: { bold: true } },
      { text: "is an implementation technique where multiple instructions are overlapped in execution, similar to a manufacturing assembly line. It maximizes processor instruction throughput.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Pipeline Architecture", { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Stage Division:", options: { bold: true } }, { text: " Instructions are split into sequential steps, each executed by a dedicated hardware stage.", options: { breakLine: true } },
      { text: "• Overlapped Execution:", options: { bold: true, color: DGN } }, { text: " When stage 1 finishes instruction A and moves it to stage 2, stage 1 immediately begins instruction B.", options: { breakLine: true } },
      { text: "• Ideal Speedup:", options: { bold: true } }, { text: " A pipeline of k stages can achieve a speedup factor of k in the ideal case (completing 1 instruction per clock cycle).", options: { breakLine: true } },
      { text: "• Latency vs. Throughput:", options: { bold: true } }, { text: " Pipelining does not decrease the execution latency of a single instruction. It increases the throughput of completed instructions.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Standard 5-Stage MIPS Pipeline", { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Most RISC designs use these classic 5 stages:", options: { breakLine: true } },
      { text: "1. Instruction Fetch (IF):", options: { bold: true, color: BL } }, { text: " Load instruction from memory.", options: { breakLine: true } },
      { text: "2. Instruction Decode (ID):", options: { bold: true, color: BL } }, { text: " Translate opcode and read register values.", options: { breakLine: true } },
      { text: "3. Execute (EX):", options: { bold: true, color: BL } }, { text: " Run arithmetic operations in the ALU.", options: { breakLine: true } },
      { text: "4. Memory Access (MEM):", options: { bold: true, color: BL } }, { text: " Read or write data from cache memory.", options: { breakLine: true } },
      { text: "5. Write-Back (WB):", options: { bold: true, color: BL } }, { text: " Write results back into register files.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by defining linear pipelining. Pipelining is an implementation technique where multiple instructions are overlapped in execution, similar to an assembly line. In a linear pipeline processor, the execution of an instruction is divided into sequential stages, with intermediate registers, called latches, separating each stage. A standard instruction pipeline contains five stages. First, Instruction Fetch, where the instruction is read from memory using the program counter. Second, Instruction Decode and Register Read, where the instruction is parsed and source operands are retrieved. Third, Execute, where the ALU performs the operation. Fourth, Memory Access, where data is read from or written to memory if it is a load or store instruction. Finally, Fifth is Write-Back, where the result is written to the destination register. Each stage takes one clock cycle, and once the pipeline is full, one instruction completes every cycle.");
  }

  // Slide 3 ── Pipeline Hazards
  {
    const s = contentSlide(p, H, "Pipeline Hazards");
    s.addText("Hazards are situations that prevent the next instruction in the instruction stream from executing in its designated clock cycle:", { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    const hazards = [
      {
        t: "1. Structural Hazards",
        b: "• Conflict: Hardware cannot support all stage combinations in the same clock cycle.\n• Cause: E.g., single memory port for both fetching instructions (IF) and loading data (MEM).\n• Solution: Separate instruction and data caches (Harvard Architecture).", bg: LGN, bd: GN
      },
      {
        t: "2. Data Hazards",
        b: "• Conflict: An instruction depends on the result of a prior instruction still in the pipeline.\n• Cause: Read-After-Write (RAW) dependencies.\n• Solution: Data forwarding (ALU bypass), compiler instruction reordering, or inserting bubbles.", bg: LBL, bd: BL
      },
      {
        t: "3. Control Hazards",
        b: "• Conflict: Imbalances caused by branch or jump instructions, which change the instruction flow.\n• Cause: Target address is unknown during Fetch stage.\n• Solution: Branch prediction hardware, branch delay slots.", bg: LYL, bd: AM
      }
    ];

    hazards.forEach((hz, i) => {
      const bx = CX + i * 4.15;
      box(p, s, bx, 2.30, 4.00, 4.60, hz.bg, hz.bd);
      boxTitle(s, bx, 2.30, 4.00, hz.t, hz.bd);
      s.addText(hz.b, { x: bx + 0.12, y: 2.30 + 0.52, w: 4.00 - 0.24, h: 4.60 - 0.58, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });
    });

    s.addNotes("In a pipeline processor, hazards are conflicts that prevent the next instruction in the instruction stream from executing in its designated clock cycle. These hazards cause the pipeline to stall or introduce bubbles, reducing overall performance. We classify pipeline hazards into three main categories. First, structural hazards occur when two or more instructions in the pipeline require the same hardware resource at the same time, such as a single-ported memory being accessed for both fetching instructions and reading data. Second, data hazards arise when an instruction depends on the result of a previous instruction that has not yet completed execution. The three sub-types are Read-After-Write, Write-After-Read, and Write-After-Write hazards. Third, control hazards are caused by branch and jump instructions that alter the program flow, making the next instruction address unknown during the fetch stage. Understanding these hazards is essential for designing robust processor control logic.");
  }

  // Slide 4 ── Resolving Data Hazards
  {
    const s = contentSlide(p, H, "Resolving Data Hazards");
    s.addText("How CPU designers resolve data dependencies in high-performance pipelines:", { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Data Forwarding & Bypassing", { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• The Concept:", options: { bold: true } }, { text: " Route the ALU calculation result directly to the ALU input stages of subsequent instructions, skipping the Write-Back register step.", options: { breakLine: true } },
      { text: "• Dynamic Routing:", options: { bold: true, color: DGN } }, { text: " Bypass multiplexers feed the EX output directly back to the EX input, resolving RAW hazards in one cycle.", options: { breakLine: true } },
      { text: "• Load-Use Stalls:", options: { bold: true, color: M } }, { text: " If an instruction depends on a memory LOAD (MEM stage), forwarding is not enough because data is not ready until the end of MEM. The pipeline MUST insert a 1-cycle stall (bubble).", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Compiler Solutions & Bubbles", { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Pipeline Stalling (Bubbles):", options: { bold: true } }, { text: " Hardware detects dependencies and pauses the fetch/decode stages, inserting a 'no-operation' (NOP) instruction bubble to let prior instructions finish.", options: { breakLine: true } },
      { text: "• Compiler Reordering:", options: { bold: true, color: BL } }, { text: " Optimizing compilers analyze dependencies and insert independent instructions into hazard gaps, avoiding hardware stalls.", options: { breakLine: true } },
      { text: "• Branch Delay Slot:", options: { bold: true } }, { text: " Compilers place an independent instruction after a branch, executing it regardless of the branch outcome, avoiding control stalls.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("To maintain high instruction throughput, processor designers and compilers use several techniques to resolve data and control hazards. Data forwarding, or bypassing, is a hardware solution where the output of the ALU is routed directly to the ALU inputs for subsequent instructions, skipping the write-back stage. However, for load-use data dependencies, hardware must stall the pipeline for one cycle, inserting a bubble, because the loaded data is not available until the end of the memory access stage. Compilers can also help by reordering instructions to place independent operations in hazard slots, or by utilizing branch delay slots where an instruction is executed regardless of the branch outcome. Hardware branch prediction also attempts to predict branch directions to reduce control hazard stalls.");
  }

  // Slide 5 ── Applications of Pipelining
  {
    const s = contentSlide(p, H, "Applications of Pipelining");
    s.addText("Pipelining is a fundamental design principle across modern processor cores:", { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🖥️  CISC & RISC Processors",
        b: "Every modern desktop CPU (Intel x86, AMD, ARM) utilizes deep pipelines (typically 10-20 stages) to achieve gigahertz clock speeds and high instruction throughput.", bg: LGN, bd: GN
      },
      {
        t: "🎨  GPU Vector Cores",
        b: "Graphic processors use deep, specialized pipeline channels to run vector math on millions of pixels concurrently, accelerating graphics rendering.", bg: LBL, bd: BL
      },
      {
        t: "🕸️  Network Router Switches",
        b: "Router switches apply pipelining to packet header checking and route lookup, forwarding network packets with minimal latency.", bg: LMR, bd: M
      },
      {
        t: "🔢  Supercomputer Vector Units",
        b: "HPC vector units use deep pipelining to compute floating-point arithmetic (matrix operations) rapidly, powering scientific modeling.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Let us look at how pipelining is applied in modern computer systems. First, desktop CISC and RISC processors utilize deep instruction pipelines, often ten to twenty stages, to achieve high clock speeds and high instruction throughput. Second, GPU vector cores utilize specialized arithmetic pipelines to run identical math operations on millions of pixels concurrently, which is critical for graphics rendering. Third, high-speed network switches apply pipelining to packet header checking and routing tables, allowing packets to be forwarded with minimal latency. Finally, supercomputer vector processing units use deep floating-point pipelines to accelerate massive matrix calculations, powering scientific modeling and simulation workloads.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, linear pipelining is a fundamental technique that dramatically increases processor performance without requiring excessive hardware redundancy. By analyzing pipeline hazards and implementing mitigation strategies like forwarding, branch prediction, and compiler scheduling, computer architects can design high-performance processors. Thank you for your time, and I am ready to answer any questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_ACA_Unit3_Pipelining.pptx" });
  convertToPdf("./outputs/Sem2_ACA_Unit3_Pipelining.pptx");
  console.log("✔  ACA Unit 3 PPT generated");
}
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 3 ACA PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  console.log("All 3 ACA PPTs generated successfully!");
})();
