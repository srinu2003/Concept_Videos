const pptxgen = require("pptxgenjs");
const { convertToPdf } = require("./pdf_converter");
const fs = require("fs");

const LOGO = "image/png;base64," + fs.readFileSync("image1.png").toString("base64");

// ── Color palette ─────────────────────────────────────────────
const M = "990033"; // maroon   (primary / border)
const GN = "009900"; // green    (header)
const GL = "F9DD67"; // gold     (inner border)
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
//  PPT 1 – FIBONACCI HEAPS  (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Data Structures";

  titleSlide(p, H, "Unit I: Heap Structures", "Fibonacci Heaps");

  // Slide 2 ── What is a Fibonacci Heap?
  {
    const s = contentSlide(p, H, "What is a Fibonacci Heap?");
    s.addText([
      { text: "A Fibonacci Heap ", options: { bold: true } },
      { text: "is a collection of min-heap-ordered trees. It provides faster amortized running times for several key operations compared to binary or binomial heaps.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Structure & Principles",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Root List:", options: { bold: true } }, { text: " Roots of all trees are linked using a circular doubly-linked list.", options: { breakLine: true } },
      { text: "• Min Pointer:", options: { bold: true } }, { text: " Points to the minimum root in the root list. Gives O(1) time access to the minimum element.", options: { breakLine: true } },
      { text: "• Relaxed Structure:", options: { bold: true } }, { text: " Trees do not have to be binomial. Consolidation of trees is deferred until the Extract-Min operation.", options: { breakLine: true } },
      { text: "• Node Pointers:", options: { bold: true } }, { text: " Each node contains parent, child (leftmost), left/right siblings, degree, and a mark flag.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Why 'Fibonacci'?",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Degree Bounds:", options: { bold: true } }, { text: " The maximum degree of any node in a heap of size n is bounded by O(log n).", options: { breakLine: true } },
      { text: "• Golden Ratio relation:", options: { bold: true } }, { text: " If a node in a Fibonacci heap has degree d, the subtree rooted at that node has size at least F_(d+2) (the (d+2)-th Fibonacci number).", options: { breakLine: true } },
      { text: "• Size growth:", options: { bold: true } }, { text: " Since Fibonacci numbers grow exponentially (F_k ≈ φ^k, where φ ≈ 1.618), the minimum size of a tree grows exponentially, bounding the maximum degree.", options: { breakLine: true } },
      { text: "• Node Marking:", options: { bold: true } }, { text: " A node is marked if it has lost a single child since it was made a child of another node. Helps maintain structural balance.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 2 – What is a Fibonacci Heap?\n\nFibonacci heaps were introduced by Fredman and Tarjan in 1984. They are a collection of min-heap-ordered trees, but unlike binomial heaps, they have a highly relaxed structure.\n\nThe trees in a Fibonacci heap are not constrained to be binomial trees. This relaxation allows us to perform operations like insertion and key reduction in O(1) actual or amortized time by deferring the structural consolidation of the heap.\n\nAll roots of the trees are kept in a circular doubly-linked list. We maintain a pointer directly to the minimum root. This gives us instant O(1) access to the minimum key.\n\nWhy is it named after Fibonacci? It is because Fibonacci numbers arise in the analysis of its operations. Specifically, for any node with degree d, the size of its subtree is at least the (d+2)-th Fibonacci number. This mathematical property bounds the maximum degree of any node to O(log n), which is crucial for proving the efficiency of Extract-Min.");
  }

  // Slide 3 ── Amortized Performance Comparison
  {
    const s = contentSlide(p, H, "Amortized Complexity Analysis");
    s.addText("Comparison of asymptotic running times between different heap types:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const tData = [
      [
        { text: "Operation", options: { bold: true, fill: { color: GN }, color: WH, align: "center" } },
        { text: "Binary Heap (Worst)", options: { bold: true, fill: { color: GN }, color: WH, align: "center" } },
        { text: "Binomial Heap (Amortized)", options: { bold: true, fill: { color: GN }, color: WH, align: "center" } },
        { text: "Fibonacci Heap (Amortized)", options: { bold: true, fill: { color: GN }, color: WH, align: "center", fill: { color: M } } }
      ],
      [
        { text: "Insert", options: { bold: true, fill: { color: LGR } } }, { text: "O(log n)", options: { fill: { color: LGR }, align: "center" } },
        { text: "O(1)", options: { fill: { color: LGR }, align: "center" } }, { text: "O(1)", options: { fill: { color: LGR }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Find Min", options: { bold: true, fill: { color: LBL } } }, { text: "O(1)", options: { fill: { color: LBL }, align: "center" } },
        { text: "O(log n) or O(1)", options: { fill: { color: LBL }, align: "center" } }, { text: "O(1)", options: { fill: { color: LBL }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Union", options: { bold: true, fill: { color: LGN } } }, { text: "O(n)", options: { fill: { color: LGN }, align: "center" } },
        { text: "O(log n)", options: { fill: { color: LGN }, align: "center" } }, { text: "O(1)", options: { fill: { color: LGN }, bold: true, color: DGN, align: "center" } }
      ],
      [
        { text: "Extract Min", options: { bold: true, fill: { color: LMR } } }, { text: "O(log n)", options: { fill: { color: LMR }, align: "center" } },
        { text: "O(log n)", options: { fill: { color: LMR }, align: "center" } }, { text: "O(log n)", options: { fill: { color: LMR }, bold: true, color: M, align: "center" } }
      ],
      [
        { text: "Decrease Key", options: { bold: true, fill: { color: LYL } } }, { text: "O(log n)", options: { fill: { color: LYL }, align: "center" } },
        { text: "O(log n)", options: { fill: { color: LYL }, align: "center" } }, { text: "O(1)", options: { fill: { color: LYL }, bold: true, color: DGN, align: "center" } }
      ]
    ];
    s.addTable(tData,
      {
        x: CX, y: 2.18, w: CW, h: 5.0, fontSize: 13.5, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [2.5, 3.28, 3.28, 3.29]
      });

    s.addNotes("Slide 3 – Complexity Analysis\n\nThis slide displays the power of Fibonacci Heaps. Traditional binary heaps require logarithmic time for insertions and key reductions.\n\nBinomial heaps optimize insertions to O(1) amortized, but Union and Decrease-Key still take O(log n) time.\n\nFibonacci heaps achieve O(1) amortized time for Insert, Union, and Decrease-Key. The only operation that requires O(log n) amortized time is Extract-Min, because it is forced to do all the deferred work of consolidating the trees in the root list. This makes the Fibonacci heap ideal for algorithms that perform a large number of Decrease-Key operations, such as Dijkstra's algorithm.");
  }

  // Slide 4 ── Extract-Min Operation
  {
    const s = contentSlide(p, H, "Extract-Min: Consolidating Trees");
    s.addText("Extract-Min is the most complex operation. It extracts the min node and cleans up the root list:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LGN, GN);
    s.addText("Step-by-Step Extract-Min Process",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Remove the Minimum Node:", options: { bold: true } }, { text: " Take out the node pointed to by the min pointer.", options: { breakLine: true } },
      { text: "2. Promote Children:", options: { bold: true } }, { text: " Add all children of the removed min node to the root list. They become root trees.", options: { breakLine: true } },
      { text: "3. Consolidate Root List:", options: { bold: true } }, { text: " Reduce the number of trees in the root list. We link trees of the same degree:", options: { breakLine: true } },
      { text: "   • Create an auxiliary array of size O(log n) indexed by degree.", options: { breakLine: true } },
      { text: "   • Traverse the root list. If two trees have the same degree d, make the one with the larger root key a child of the other, increasing its degree to d+1.", options: { breakLine: true } },
      { text: "   • Repeat this linking process until there is at most one tree of each degree.", options: { breakLine: true } },
      { text: "4. Re-establish Min Pointer:", options: { bold: true } }, { text: " Scan the remaining consolidated roots to find the new minimum node.", options: {} }
    ], { x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 4 – Extract-Min Operation\n\nExtract-Min is where the 'lazy' design of the Fibonacci heap catch up with us. Since we did not structure the heap during insertions, the root list might contain many single nodes.\n\nDuring Extract-Min:\n1. We remove the minimum node.\n2. We take its children and make them roots by placing them in the root list.\n3. Then, we consolidate. We use a degree array to keep track of roots by degree. If we find two roots with the same degree, we merge them by making the root with the larger key a child of the root with the smaller key.\n4. We repeat this until all roots in the root list have distinct degrees. This step bounds the size of the root list to O(log n) and guarantees the amortized log n complexity.");
  }

  // Slide 5 ── Decrease-Key Operation
  {
    const s = contentSlide(p, H, "Decrease-Key: Cascading Cuts");
    s.addText("Decrease-Key reduces a node's key value. If the heap order is violated, it performs cuts:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.70, LMR, M);
    s.addText("The Cut & Cascading Cut Algorithm",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "If node x's new key is less than its parent y's key, heap order is violated:", options: { breakLine: true } },
      { text: "1. Cut Node x:", options: { bold: true, color: M } }, { text: " Sever the link between x and y. Add x to the root list. Clear its mark flag.", options: { breakLine: true } },
      { text: "2. Cascading Cut on Parent y:", options: { bold: true, color: M } }, { text: " Check parent y:", options: { breakLine: true } },
      { text: "   • If y is a root, do nothing.", options: { breakLine: true } },
      { text: "   • If y is unmarked, mark it (indicates it has lost a child).", options: { breakLine: true } },
      { text: "   • If y is already marked, cut y from its parent, add y to the root list (unmarked), and recursively apply cascading cut to y's parent.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 4.00, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.70, LYL, AM);
    s.addText("Intuition & Purpose",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• O(1) Amortized Time:", options: { bold: true } }, { text: " By cutting immediately, we avoid traversing the tree to bubble up. The amortized cost of cascading cuts balances out over time.", options: { breakLine: true } },
      { text: "• Balanced Subtrees:", options: { bold: true } }, { text: " Node marking guarantees that no node loses more than one child while remaining a child itself. This preserves the exponential size property of trees relative to their degree.", options: { breakLine: true } },
      { text: "• Fibonacci Relation:", options: { bold: true } }, { text: " This specific rule ensures that a node of degree d has a subtree of size at least φ^d, keeping the maximum degree bounded by O(log n).", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 4.00, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 5 – Decrease-Key Operation\n\nDecrease-Key is the most elegant operation. If we decrease a node's key and it becomes smaller than its parent's key, the min-heap property is violated.\n\nInstead of bubbling up, we cut the node from its parent and add it directly to the root list as a new tree root. However, if we do this repeatedly without rules, the trees could become long thin chains, destroying the O(log n) degree bound.\n\nTo prevent this, we use the mark flag. If a parent node loses its second child, we cut the parent too, and move it to the root list. This is called a cascading cut. It recursively bubbles up until it hits an unmarked node (which it marks) or a root. This cascading process ensures trees remain bushy and the Fibonacci bounds hold.");
  }

  // Slide 6 ── Applications of Fibonacci Heaps
  {
    const s = contentSlide(p, H, "Applications in Computer Science");
    s.addText("Fibonacci heaps are primarily used in graph algorithms where Decrease-Key is frequent:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🚗  Dijkstra's Shortest Path",
        b: "Dijkstra's algorithm finds the shortest path from a source node to all others. Using a Fibonacci heap, the complexity is reduced from O(E log V) to O(E + V log V). This is highly efficient for dense graphs with many edges.", bg: LGN, bd: GN
      },
      {
        t: "🌲  Prim's Minimum Spanning Tree",
        b: "Prim's algorithm finds the MST of a weighted graph. Using a Fibonacci heap to store vertices ordered by connection weight improves the runtime to O(E + V log V), matching Dijkstra's performance.", bg: LBL, bd: BL
      },
      {
        t: "🕸️  Network Flow Algorithms",
        b: "Several advanced network flow and matching algorithms rely on min-priority queues. Using Fibonacci heaps speeds up the bottleneck operations of finding augmenting paths.", bg: LMR, bd: M
      },
      {
        t: "📊  Theoretical Algorithmics",
        b: "Fibonacci heaps are used as a building block in many advanced data structures, such as directed minimum spanning trees, and link-cut trees, where theoretical efficiency is critical.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Slide 6 – Applications of Fibonacci Heaps\n\nFibonacci heaps excel in network and graph optimization algorithms. In Dijkstra's and Prim's algorithms, the number of Decrease-Key operations is bounded by the number of edges E, whereas Extract-Min is done V times.\n\nBy optimizing Decrease-Key to O(1) amortized, Fibonacci heaps yield an overall time complexity of O(E + V log V) instead of O(E log V). For dense graphs, this is a significant speedup.\n\nHowever, in practice, Fibonacci heaps have large constant factors and complex pointer overhead. Therefore, for smaller or sparse graphs, simpler structures like binary heaps are often faster. But for very large scale networks, Fibonacci heaps are theoretically superior.");
  }

  thankYouSlide(p, H);
  await p.writeFile({ fileName: "./outputs/Sem1_ADS_Unit1_Fibonacci_Heaps.pptx" });
  convertToPdf("./outputs/Sem1_ADS_Unit1_Fibonacci_Heaps.pptx");
  console.log("✔  ADS Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – COLLISION RESOLUTION IN HASHING  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Data Structures";

  titleSlide(p, H, "Unit II: Hashing and Collisions", "Collision Resolution in Hashing");

  // Slide 2 ── Hashing and the Collision Problem
  {
    const s = contentSlide(p, H, "Hashing and the Collision Problem");
    s.addText([
      { text: "Hashing ", options: { bold: true } },
      { text: "maps keys of arbitrary size to fixed-size values (hash codes) using a hash function. A collision occurs when two distinct keys hash to the same table slot.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Core Concepts",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Hash Table:", options: { bold: true } }, { text: " An array of size m that stores key-value pairs.", options: { breakLine: true } },
      { text: "• Hash Function h(k):", options: { bold: true } }, { text: " Maps key k to an index in [0, m-1]. Ideal function is uniform (distributes keys equally).", options: { breakLine: true } },
      { text: "• Load Factor (α):", options: { bold: true } }, { text: " α = n / m, where n is the number of keys and m is the table size. Indicates how full the table is.", options: { breakLine: true } },
      { text: "• Collision Definition:", options: { bold: true } }, { text: " If k1 ≠ k2 but h(k1) = h(k2), they contend for the same array index.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("Why Collisions are Unavoidable",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Pigeonhole Principle:", options: { bold: true } }, { text: " If the universe of keys U is larger than table size m, multiple keys MUST map to the same slot.", options: { breakLine: true } },
      { text: "• Birthday Paradox:", options: { bold: true } }, { text: " In a table of size 365, there is a 50% chance of collision with only 23 keys inserted, even with a perfect random hash function.", options: { breakLine: true } },
      { text: "• Resolution Requirement:", options: { bold: true } }, { text: " Hashing requires an efficient collision resolution strategy to maintain O(1) average lookup speed.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 2 – Hashing and the Collision Problem\n\nHashing is a classic data structure technique that achieves O(1) average time complexity for insertion, deletion, and lookup. We use a hash function to map keys to slots in an array.\n\nHowever, because the set of possible keys is usually much larger than the array size, collisions are mathematically guaranteed to happen. This is a consequence of the Pigeonhole Principle.\n\nFurthermore, the Birthday Paradox shows that collisions occur much sooner than intuitively expected. For instance, with only 23 keys in a table of size 365, the probability of at least one collision is already over 50%.\n\nTherefore, the design of a hash table consists of two parts: a good hash function to minimize collisions, and a collision resolution algorithm to handle them when they happen.");
  }

  // Slide 3 ── Separate Chaining
  {
    const s = contentSlide(p, H, "Separate Chaining (Open Hashing)");
    s.addText("Separate Chaining resolves collisions by maintaining a linked list of elements at each slot:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("How it Works",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Slots contain pointers:", options: { bold: true } }, { text: " Each array entry points to the head of a linked list (or other dynamic container).", options: { breakLine: true } },
      { text: "• Insert:", options: { bold: true } }, { text: " Hash key to slot i. Insert new element at the head of list i. O(1) time.", options: { breakLine: true } },
      { text: "• Search:", options: { bold: true } }, { text: " Hash key to slot i. Traverse list i to find key. O(1 + α) average time.", options: { breakLine: true } },
      { text: "• Delete:", options: { bold: true } }, { text: " Hash key to slot i. Find key in list i and remove the node. O(1 + α) average time.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGR, GY);
    s.addText("Pros & Cons",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: GY, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "✔ Simple to implement.", options: { bold: true, color: DGN, breakLine: true } },
      { text: "✔ Table never fills up:", options: { bold: true } }, { text: " Load factor α can exceed 1 without catastrophic failure.", options: { breakLine: true } },
      { text: "✔ Graceful degradation:", options: { bold: true } }, { text: " Performance degrades linearly as α grows.", options: { breakLine: true } },
      { text: "✘ Memory overhead:", options: { bold: true, color: M } }, { text: " Extra memory is wasted storing pointers for linked list nodes.", options: { breakLine: true } },
      { text: "✘ Cache inefficiency:", options: { bold: true, color: M } }, { text: " Linked list nodes are scattered in memory, causing CPU cache misses.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 3 – Separate Chaining\n\nSeparate Chaining is a simple and robust collision resolution method. It is often referred to as open hashing because keys are stored outside the primary table array.\n\nIn separate chaining, each table index points to a linked list. When keys collide, they are simply appended to the list at that index. The load factor alpha represents the average length of these lists. Search time is proportional to 1 + alpha.\n\nThe main benefit is that the table can never get full — we can store more elements than the array size (alpha > 1). However, the pointer overhead is a disadvantage, and traversing linked lists causes poor CPU cache locality.");
  }

  // Slide 4 ── Open Addressing
  {
    const s = contentSlide(p, H, "Open Addressing (Closed Hashing)");
    s.addText("Open Addressing stores all elements directly within the array. It probes subsequent slots on collision:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    const probes = [
      {
        t: "1. Linear Probing",
        b: "Checks consecutive slots: h(k, i) = (h'(k) + i) % m.\n• Easy to implement.\n• High cache locality.\n• Primary Clustering: Long runs of occupied slots build up, slowing down searches.", bg: LGN, bd: GN
      },
      {
        t: "2. Quadratic Probing",
        b: "Checks slots quadratically: h(k, i) = (h'(k) + c1·i + c2·i²) % m.\n• Reduces primary clustering.\n• Secondary Clustering: Keys that hash to the same index traverse the same probe path.", bg: LBL, bd: BL
      },
      {
        t: "3. Double Hashing",
        b: "Uses a second hash function to determine step size: h(k, i) = (h1(k) + i·h2(k)) % m.\n• Best distribution: step size depends on key.\n• Eliminates clustering.\n• h2(k) must be coprime to m.", bg: LYL, bd: AM
      }
    ];

    probes.forEach((pb, i) => {
      const bx = CX + i * 4.15;
      box(p, s, bx, 2.30, 4.00, 4.60, pb.bg, pb.bd);
      boxTitle(s, bx, 2.30, 4.00, pb.t, pb.bd);
      s.addText(pb.b,
        { x: bx + 0.12, y: 2.30 + 0.52, w: 4.00 - 0.24, h: 4.60 - 0.58, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });
    });

    s.addNotes("Slide 4 – Open Addressing\n\nIn Open Addressing, all elements are stored directly inside the array, meaning we have no linked lists and no pointers. If a slot is occupied, we search (probe) the table in a systematic sequence until we find an empty slot.\n\nWe look at three major probing strategies:\n1. Linear Probing: Checks indices sequentially. It is very fast for CPU caches, but suffers from primary clustering — when keys group together in long chains, making lookup O(n) in the worst case.\n2. Quadratic Probing: Steps grow quadratically to avoid linear clusters, but still has secondary clustering since keys with the same initial hash index share the same path.\n3. Double Hashing: The best open addressing method. It uses a second hash function to calculate a custom step size for each key, eliminating both primary and secondary clustering.");
  }

  // Slide 5 ── Chaining vs. Open Addressing
  {
    const s = contentSlide(p, H, "Comparison: Chaining vs. Open Addressing");
    s.addText("A direct structural and performance comparison between the two strategies:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const tData = [
      [
        { text: "Feature", options: { bold: true, fill: { color: M }, color: WH } },
        { text: "Separate Chaining", options: { bold: true, fill: { color: M }, color: WH, align: "center" } },
        { text: "Open Addressing", options: { bold: true, fill: { color: M }, color: WH, align: "center" } }
      ],
      [
        { text: "Storage Location", options: { bold: true, fill: { color: LGR } } }, { text: "Outside array (pointers/lists)", options: { fill: { color: LGR } } }, { text: "Inside the primary array slots", options: { fill: { color: LGR } } }
      ],
      [
        { text: "Max Load Factor (α)", options: { bold: true, fill: { color: LBL } } }, { text: "α can be > 1 (limitless)", options: { fill: { color: LBL } } }, { text: "α strictly ≤ 1 (typically keeps α < 0.7)", options: { fill: { color: LBL } } }
      ],
      [
        { text: "Cache Performance", options: { bold: true, fill: { color: LGN } } }, { text: "Poor (linked list traversals)", options: { fill: { color: LGN } } }, { text: "Excellent (array elements adjacent)", options: { fill: { color: LGN } } }
      ],
      [
        { text: "Deletion Complexity", options: { bold: true, fill: { color: LMR } } }, { text: "Simple (standard node deletion)", options: { fill: { color: LMR } } }, { text: "Complex (requires 'deleted' placeholders)", options: { fill: { color: LMR } } }
      ],
      [
        { text: "Memory Overhead", options: { bold: true, fill: { color: LYL } } }, { text: "Extra memory spent on pointers", options: { fill: { color: LYL } } }, { text: "Wasted array space if table is sparse", options: { fill: { color: LYL } } }
      ]
    ];
    s.addTable(tData,
      {
        x: CX, y: 2.20, w: CW, h: 4.80, fontSize: 13, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [3.0, 4.67, 4.68]
      });

    s.addNotes("Slide 5 – Chaining vs. Open Addressing\n\nThis slide summarizes the tradeoffs. Chaining is robust against poor hash functions and doesn't break if load factor exceeds 1. Deletion is also simple.\n\nOpen Addressing is highly cache-friendly, but requires load factor to be kept small (usually below 0.7) to prevent slow lookups. Deletion in open addressing is notoriously tricky because removing an element can break the search path of other keys. We have to use special 'Deleted' marker values, which complicate the insert and search algorithms.");
  }

  // Slide 6 ── Real-World Hashing Applications
  {
    const s = contentSlide(p, H, "Applications of Hash Tables");
    s.addText("Hashing and collision resolution are critical parts of high-performance software engineering:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🗃  Symbol Tables in Compilers",
        b: "Compilers and interpreters use hash tables to store identifier names (variables, functions) and their metadata. High lookup speeds are vital during syntactic analysis and code generation.", bg: LGN, bd: GN
      },
      {
        t: "💾  Database Query Caching",
        b: "Database engines cache parsed queries and execution plans in hash tables for instant retrieval, avoiding costly re-compilation of query strings.", bg: LBL, bd: BL
      },
      {
        t: "🌐  IP Routing Table Lookups",
        b: "Routers use hash tables to store IP routing routes, enabling quick route matching to direct network packets with minimum latency.", bg: LMR, bd: M
      },
      {
        t: "🔒  Set Implementations & deduplication",
        b: "Language runtimes implement 'Set' classes (e.g., HashSet in Java/C#) using hash tables to support O(1) membership checks and fast duplicate elimination.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Slide 6 – Applications\n\nWe wrap up by seeing where hash tables are used. They are everywhere — in compilers (symbol tables), database systems (query caches), network routers (IP routes), and in core data structure implementations like hash sets and hash maps in modern programming languages. Implementing efficient collision resolution is what makes these systems performant in practice.");
  }

  thankYouSlide(p, H);
  await p.writeFile({ fileName: "./outputs/Sem1_ADS_Unit2_Collision_Hashing.pptx" });
  convertToPdf("./outputs/Sem1_ADS_Unit2_Collision_Hashing.pptx");
  console.log("✔  ADS Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – AVL TREES  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Data Structures";

  titleSlide(p, H, "Unit III: Search Structures", "AVL Trees");

  // Slide 2 ── What is an AVL Tree?
  {
    const s = contentSlide(p, H, "What is an AVL Tree?");
    s.addText([
      { text: "An AVL Tree ", options: { bold: true } },
      { text: "(named after Adelson-Velsky and Landis) is a self-balancing Binary Search Tree (BST) where the heights of the two child subtrees of any node differ by at most one.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Key Rules & Properties",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Balance Factor (BF):", options: { bold: true } }, { text: " For any node N, BF(N) = Height(LeftSubtree) − Height(RightSubtree).", options: { breakLine: true } },
      { text: "• AVL Property:", options: { bold: true } }, { text: " The balance factor of EVERY node must be in {-1, 0, 1}.", options: { breakLine: true } },
      { text: "• Height Bound:", options: { bold: true } }, { text: " The height of an AVL tree with n nodes is strictly bounded by 1.44 log₂(n), making it highly balanced.", options: { breakLine: true } },
      { text: "• Time Complexity:", options: { bold: true } }, { text: " Search, Insert, and Delete operations all take O(log n) worst-case time.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LMR, M);
    s.addText("The Need for Balancing",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Standard BST Problem:", options: { bold: true } }, { text: " If keys are inserted in sorted order (e.g. 1, 2, 3, 4), a normal BST degenerates into a linear linked list with O(n) search time.", options: { breakLine: true } },
      { text: "• Active Balancing:", options: { bold: true } }, { text: " AVL trees monitor the balance factor during insertions and deletions.", options: { breakLine: true } },
      { text: "• Self-Correction:", options: { bold: true } }, { text: " If any node's BF becomes +2 or -2, the tree performs localized rotations to restore the AVL balance property immediately.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 2 – What is an AVL Tree?\n\nAn AVL Tree is the very first self-balancing binary search tree ever invented. In a standard BST, if we insert sorted keys, the tree becomes highly skewed, resulting in O(n) search operations.\n\nAVL trees guarantee logarithmic time by maintaining a balance condition. The heights of the left and right subtrees of every node can differ by at most 1. This condition is monitored using a Balance Factor.\n\nIf an insertion or deletion causes the Balance Factor to deviate to +2 or -2, we perform rotations to restore balance. This guarantees that search, insertion, and deletion are always O(log n) in the worst case.");
  }

  // Slide 3 ── Single Rotations
  {
    const s = contentSlide(p, H, "Single Rotations (LL & RR Cases)");
    s.addText("Single rotations are used when the insertion occurs in the 'outer' subtrees (Left-Left or Right-Right):",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Right Rotation (LL Case)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Trigger:", options: { bold: true } }, { text: " Node y is inserted into the left subtree of the left child x of root z (Left-of-Left violation). z has BF = +2.", options: { breakLine: true } },
      { text: "• Operation:", options: { bold: true } }, { text: " Pivot around x. Pull z down to become the right child of x. Move x's right child to become z's left child.", options: { breakLine: true } },
      { text: "• Result:", options: { bold: true } }, { text: " x becomes the new root of this subtree, with y on its left and z on its right. Height is balanced.", options: { breakLine: true } },
      { text: "• Code Logic:", options: { bold: true, color: BL } }, { text: " z.left = x.right; x.right = z;", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Left Rotation (RR Case)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Trigger:", options: { bold: true } }, { text: " Node y is inserted into the right subtree of the right child x of root z (Right-of-Right violation). z has BF = -2.", options: { breakLine: true } },
      { text: "• Operation:", options: { bold: true } }, { text: " Pivot around x. Pull z down to become the left child of x. Move x's left child to become z's right child.", options: { breakLine: true } },
      { text: "• Result:", options: { bold: true } }, { text: " x becomes the new root of this subtree, with z on its left and y on its right. Balance restored.", options: { breakLine: true } },
      { text: "• Code Logic:", options: { bold: true, color: DGN } }, { text: " z.right = x.left; x.left = z;", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 3 – Single Rotations\n\nSingle rotations handle straight-line imbalances. There are two symmetric cases:\n1. Left-Left (LL) case: The tree is heavy on the left-of-the-left. We perform a Single Right Rotation to balance it. The middle node becomes the new root, and the old root becomes its right child.\n2. Right-Right (RR) case: The tree is heavy on the right-of-the-right. We perform a Single Left Rotation to restore balance.\n\nBoth single rotations take O(1) time since they only involve reassigning a few pointers. They do not alter the BST ordering property, so keys remain sorted.");
  }

  // Slide 4 ── Double Rotations
  {
    const s = contentSlide(p, H, "Double Rotations (LR & RL Cases)");
    s.addText("Double rotations are required when the insertion occurs in the 'inner' subtrees (Left-Right or Right-Left):",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Left-Right Rotation (LR Case)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Trigger:", options: { bold: true } }, { text: " Node is inserted into the right subtree of the left child of root z (Left-Right violation). z has BF = +2.", options: { breakLine: true } },
      { text: "• Two Steps to Resolve:", options: { bold: true, color: M } }, { text: "\n", options: { breakLine: true } },
      { text: "  1. Left Rotation on Left Child:", options: { bold: true } }, { text: " Align the imbalance into a straight line (turns LR case into LL case).", options: { breakLine: true } },
      { text: "  2. Right Rotation on Root z:", options: { bold: true } }, { text: " Restores the actual height balance of the tree.", options: { breakLine: true } },
      { text: "• Final state:", options: { bold: true } }, { text: " The deepest node is promoted to become the root of this subtree.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Right-Left Rotation (RL Case)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Trigger:", options: { bold: true } }, { text: " Node is inserted into the left subtree of the right child of root z (Right-Left violation). z has BF = -2.", options: { breakLine: true } },
      { text: "• Two Steps to Resolve:", options: { bold: true, color: AM } }, { text: "\n", options: { breakLine: true } },
      { text: "  1. Right Rotation on Right Child:", options: { bold: true } }, { text: " Align the imbalance into a straight line (turns RL case into RR case).", options: { breakLine: true } },
      { text: "  2. Left Rotation on Root z:", options: { bold: true } }, { text: " Restores the actual height balance of the tree.", options: { breakLine: true } },
      { text: "• Final state:", options: { bold: true } }, { text: " The deepest node is promoted to become the root of this subtree.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 4 – Double Rotations\n\nWhen imbalances are 'zig-zag' (inner subtrees), a single rotation is not enough to balance the tree. We must perform a double rotation.\n\nFor the Left-Right (LR) case, we first perform a Left Rotation on the left child, which transforms the structure into a Left-Left (LL) shape. Then, we perform a Right Rotation on the parent node to restore balance.\n\nThe Right-Left (RL) case is the exact mirror image. We first perform a Right Rotation on the right child, and then a Left Rotation on the parent node. Both operations are O(1) in terms of pointer changes, maintaining the theoretical efficiency of AVL trees.");
  }

  // Slide 5 ── Tracing AVL Insertion and Deletion
  {
    const s = contentSlide(p, H, "AVL Insertion & Deletion Complexity");
    s.addText("How AVL trees maintain balance during dynamic insertions and deletions:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGR, GY);
    s.addText("Insertion Process & Rebalancing",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GY, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Standard BST Insert:", options: { bold: true } }, { text: " Insert the node recursively at a leaf slot. O(log n) search steps.", options: { breakLine: true } },
      { text: "2. Backtrack and Update:", options: { bold: true } }, { text: " Retrace the path upwards towards the root, updating balance factors.", options: { breakLine: true } },
      { text: "3. Detect Imbalance:", options: { bold: true } }, { text: " Identify the first node where BF becomes +2 or -2.", options: { breakLine: true } },
      { text: "4. Perform Rotations:", options: { bold: true } }, { text: " Execute LL, RR, LR, or RL rotation as appropriate. At most ONE rotation is needed to restore balance to the entire tree after an insertion.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Deletion Process & Rebalancing",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Standard BST Delete:", options: { bold: true } }, { text: " Remove the node. If it has two children, swap with its inorder predecessor or successor, then remove.", options: { breakLine: true } },
      { text: "2. Retrace and Rebalance:", options: { bold: true } }, { text: " Retrace upwards, updating balance factors.", options: { breakLine: true } },
      { text: "3. Multiple Rotations:", options: { bold: true, color: M } }, { text: " Unlike insertion, a single rotation after a deletion may balance a subtree but reduce its height, causing imbalances further up. We may need to perform up to O(log n) rotations up to the root.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 5 – Dynamic Operations\n\nWhen inserting into an AVL tree, we insert the node as a leaf, and then walk back up to the root updating balance factors. If we detect an imbalance, we perform a single or double rotation. Once a rotation is performed at the lowest unbalanced node, the height of that subtree is restored to its pre-insertion state, so no further rotations are required.\n\nDeletion is more complex. Removing a node can shorten a path, which can cause imbalances above it. Retracing back up, we may need to perform rotations at multiple levels, up to O(log n) times. However, each rotation still takes O(1) time, so the overall deletion complexity remains O(log n).");
  }

  // Slide 6 ── Applications of AVL Trees
  {
    const s = contentSlide(p, H, "Applications of AVL Trees");
    s.addText("AVL trees are highly efficient for lookup-heavy, dynamic key-value storage:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🔍  High-Performance Databases",
        b: "Used in database indexing where read operations outnumber write operations. AVL trees are more strictly balanced than Red-Black trees, providing faster search lookups (at the cost of slightly slower insertions).", bg: LGN, bd: GN
      },
      {
        t: "💻  Language Symbol Lookups",
        b: "Used in language runtimes and memory tables to search for variable names, functions, and object keys where quick access is crucial.", bg: LBL, bd: BL
      },
      {
        t: "🌐  Network Routing Tables",
        b: "Router tables use AVL structures to index subnet IPs and prefixes, ensuring rapid packet forwarding routing paths.", bg: LMR, bd: M
      },
      {
        t: "🎮  Real-time Systems",
        b: "Real-time simulation and gaming engines use AVL trees for spatial indexing and quick nearest-neighbor queries where worst-case lookup bounds are strict.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Slide 6 – Applications\n\nAVL trees are used when lookup speed is paramount. Because AVL trees have a strict balance factor condition, their average and worst-case heights are closer to log₂ n. In comparison, Red-Black trees have a more relaxed balance condition, making Red-Black trees faster for insertions but slightly slower for search lookups.\n\nThis makes AVL trees the preferred choice for indexing datasets that are mostly read-only or have minimal write operations, such as search indexes, network route tables, and lookup maps in real-time systems.");
  }

  thankYouSlide(p, H);
  await p.writeFile({ fileName: "./outputs/Sem1_ADS_Unit3_AVL_Trees.pptx" });
  convertToPdf("./outputs/Sem1_ADS_Unit3_AVL_Trees.pptx");
  console.log("✔  ADS Unit 3 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 4 – TRIES  (Unit IV)
// ═══════════════════════════════════════════════════════════════
async function ppt4() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Data Structures";

  titleSlide(p, H, "Unit IV: Digital Search Structures", "Tries & Digital Search Structures");

  // Slide 2 ── What is a Trie?
  {
    const s = contentSlide(p, H, "Introduction to Tries");
    s.addText([
      { text: "A Trie ", options: { bold: true } },
      { text: "(derived from 'retrieval') is an ordered tree-like data structure used to store a dynamic set of strings, where keys are usually strings. It is also known as a prefix tree.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Structure & Design",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Key Sharing:", options: { bold: true } }, { text: " Unlike BSTs, nodes do not store their associated keys. Instead, a node's position in the tree defines its key.", options: { breakLine: true } },
      { text: "• Character Edges:", options: { bold: true } }, { text: " Edges outgoing from a node represent characters in the alphabet. All descendants of a node share a common string prefix.", options: { breakLine: true } },
      { text: "• End of Word Marker:", options: { bold: true } }, { text: " A boolean flag (isEndOfWord) indicates if the path from root to that node represents a complete word in the set.", options: { breakLine: true } },
      { text: "• Root Node:", options: { bold: true } }, { text: " Always represents the empty string (root of all search prefixes).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LGN, GN);
    s.addText("Key Complexity & Benefits",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• O(L) Operations:", options: { bold: true } }, { text: " Lookup, insert, and delete run in O(L) time, where L is the length of the string. Completely independent of table size n!", options: { breakLine: true } },
      { text: "• Fast Misses:", options: { bold: true } }, { text: " Searches terminate immediately when a character edge is missing — much faster than hashing on key mismatches.", options: { breakLine: true } },
      { text: "• Prefix Queries:", options: { bold: true } }, { text: " Can easily retrieve all keys starting with a given prefix, which is impossible with standard hash tables.", options: { breakLine: true } },
      { text: "• Space Trade-off:", options: { bold: true, color: M } }, { text: " Tries can use significant memory because of pointer overhead for empty edges at each node.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 2 – Introduction to Tries\n\nA Trie, also known as a prefix tree, is a specialized tree structure for storing strings. The name comes from the word 're-trie-val'.\n\nIn a trie, keys are represented by paths from the root rather than being stored in individual nodes. Every node represents a prefix. The root represents the empty string. Edges represent characters. As you traverse down, you form words.\n\nSearching for a word of length L takes O(L) time. This is incredibly fast, especially for large dictionaries, because the lookup time is independent of the number of words stored in the trie. However, the downside is space. Each node must have pointers to all possible child characters, which can waste memory.");
  }

  // Slide 3 ── Standard Tries vs. Binary Tries
  {
    const s = contentSlide(p, H, "Standard Tries vs. Binary Tries");
    s.addText("Tries can be implemented using different alphabet sizes and structures:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Standard Tries (Multiway Tries)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Alphabet size (Σ):", options: { bold: true } }, { text: " Typically large (e.g. Σ = 26 for lowercase English, Σ = 256 for ASCII).", options: { breakLine: true } },
      { text: "• Child Representation:", options: { bold: true } }, { text: " Each node stores an array of size Σ pointing to child nodes: Node* children[26].", options: { breakLine: true } },
      { text: "• Pointer overhead:", options: { bold: true, color: M } }, { text: " Wastes substantial memory. If a node only has one child, 25 array slots are empty pointers.", options: { breakLine: true } },
      { text: "• Search Speed:", options: { bold: true } }, { text: " Extremely fast. Direct array access child = node.children[char] in O(1) step time.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Binary Tries",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Alphabet size (Σ):", options: { bold: true } }, { text: " Binary alphabet Σ = {0, 1}.", options: { breakLine: true } },
      { text: "• Bit-level Representation:", options: { bold: true } }, { text: " Keys are split into their binary bit strings. The edges represent 0 or 1 bits.", options: { breakLine: true } },
      { text: "• Child Representation:", options: { bold: true } }, { text: " Node only has two child pointers: left (for 0) and right (for 1).", options: { breakLine: true } },
      { text: "• Space Efficiency:", options: { bold: true } }, { text: " High. Minimum pointer waste compared to standard multiway tries.", options: { breakLine: true } },
      { text: "• Bit-level search:", options: { bold: true } }, { text: " Searching involves traversing bit by bit. Useful in routing tables where IP addresses are represented as 32-bit strings.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 3 – Standard vs. Binary Tries\n\nWe compare standard multiway tries and binary tries. Standard tries use a child pointer array for each character of the alphabet. This is very fast because char indices map directly to array elements, but it is extremely wasteful in memory.\n\nBinary tries restrict the alphabet to {0,1}. Strings are parsed at the bit level. Each node has only two pointers (left and right), which dramatically reduces empty pointer waste. Binary tries are heavily used in low-level networking, particularly for IP address prefix matching in routers.");
  }

  // Slide 4 ── Patricia Tries
  {
    const s = contentSlide(p, H, "Compressed Tries & Patricia Tries");
    s.addText("To solve the space overhead of tries, we compress paths with single-child nodes:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LGN, DGN);
    s.addText("Compressed Tries",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• The Idea:", options: { bold: true } }, { text: " If a chain of single-child nodes exists (e.g. p -> a -> t -> h), collapse them into a single node with edge label 'path'.", options: { breakLine: true } },
      { text: "• Structural Benefit:", options: { bold: true } }, { text: " Reduces the number of internal nodes. Every internal node in a compressed trie now has at least two children.", options: { breakLine: true } },
      { text: "• Size Bound:", options: { bold: true } }, { text: " For a set of n words, the number of leaf nodes is n, and the number of internal nodes is at most n-1. Total space is O(n).", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Patricia Tries",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Definition:", options: { bold: true } }, { text: " PATRICIA stands for: 'Practical Algorithm to Retrieve Information Coded in Alphanumeric'.", options: { breakLine: true } },
      { text: "• Compact Binary Trie:", options: { bold: true } }, { text: " A compressed binary trie. Instead of storing edge characters, nodes store the bit-index to be tested next.", options: { breakLine: true } },
      { text: "• Fast Bit Testing:", options: { bold: true } }, { text: " When searching, we skip bit positions that do not distinguish between keys, jumping directly to the bit position where they diverge.", options: { breakLine: true } },
      { text: "• Application:", options: { bold: true } }, { text: " Highly efficient for prefix search, IP routing lookups, and text processing.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 4 – Patricia Tries\n\nTo solve the space problem, we can compress the trie. If a node has only one child, we merge it with its child, combining the edge characters. This is a Compressed Trie. It guarantees that every internal node has at least two children, which bounds the node count to O(n).\n\nA Patricia Trie is a specialized compressed binary trie. It is designed to speed up search by storing the bit-index to test at each node. Instead of testing every bit sequentially, the search jumps directly to the bit positions where keys differ, making lookups extremely efficient.");
  }

  // Slide 5 ── Suffix Trees
  {
    const s = contentSlide(p, H, "Suffix Trees");
    s.addText([
      { text: "A Suffix Tree ", options: { bold: true } },
      { text: "is a compressed trie containing all suffixes of a single string S. It is a powerful index for text searching.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LYL, AM);
    s.addText("How it is Built",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Suffix Set:", options: { bold: true } }, { text: " For a string S of length n, append a termination character '$'. Generate all n suffixes of S.", options: { breakLine: true } },
      { text: "• Example (S = 'BANANA$'):", options: { bold: true } }, { text: " Suffixes are:\n  'BANANA$', 'ANANA$', 'NANA$', 'ANA$', 'NA$', 'A$', '$'.", options: { breakLine: true } },
      { text: "• Trie Compression:", options: { bold: true } }, { text: " Insert all these suffixes into a compressed trie.", options: { breakLine: true } },
      { text: "• Linear Construction:", options: { bold: true, color: DGN } }, { text: " Ukkonen's algorithm builds a suffix tree in O(n) time and space.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LMR, M);
    s.addText("String Operations",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Pattern Matching:", options: { bold: true } }, { text: " Check if pattern P is a substring of S in O(|P|) time. We just trace P from root.", options: { breakLine: true } },
      { text: "• Longest Repeated Substring:", options: { bold: true } }, { text: " Find the deepest internal node in the tree. Can be solved in O(n) time.", options: { breakLine: true } },
      { text: "• Longest Common Substring:", options: { bold: true } }, { text: " Build a generalized suffix tree for S1 and S2. Deepest node with descendants from both strings is the solution. O(n1 + n2) time.", options: { breakLine: true } },
      { text: "• Substring Count:", options: { bold: true } }, { text: " Find the number of occurrences of P in O(|P|) time.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 5 – Suffix Trees\n\nA Suffix Tree is a spectacular data structure. If you insert all possible suffixes of a string S of length n into a compressed trie, you get a suffix tree.\n\nFor example, the suffixes of 'BANANA$' are inserted. The '$' character is used to ensure no suffix is a prefix of another, so all suffixes end at leaf nodes.\n\nWhile a naive construction would take O(n²) time, Ukkonen's algorithm builds it in O(n) time. Once constructed, we can perform advanced string queries, like finding substrings, finding the longest common substring, or checking repetition, in time proportional to the query length, completely independent of the size of the original text.");
  }

  // Slide 6 ── Applications of Tries
  {
    const s = contentSlide(p, H, "Applications of Digital Search Structures");
    s.addText("Tries and suffix trees form the core of modern search, routing, and text analysis:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "⌨️  Autocomplete & Typeahead",
        b: "Tries store dictionary terms. As the user types, the prefix acts as a path to a node. The sub-tree under that node is traversed to yield suggestions in O(L) time, where L is prefix length.", bg: LGN, bd: GN
      },
      {
        t: "🌐  IP Routing Table Lookups",
        b: "Patricia tries are used in network routers to store routing tables. The router performs a 'Longest Prefix Match' on the incoming packet's IP address to find the forwarding route in O(1) time.", bg: LBL, bd: BL
      },
      {
        t: "📖  Spell Checkers & Dictionaries",
        b: "Tries act as spelling dictionaries. It is extremely fast to detect if a word exists, and to list nearby word candidates using spelling edit distance algorithms on the trie.", bg: LMR, bd: M
      },
      {
        t: "🧬  Bioinformatics & DNA Indexing",
        b: "Suffix trees index massive genomes. Researchers search for DNA sub-sequences, repeating genes, and similarities across species in linear time using suffix tree lookups.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Slide 6 – Applications\n\nDigital search structures are incredibly practical. Autocomplete is the classic application, where tries are used to quickly find words matching a prefix.\n\nIn networking, Patricia tries are used for routing tables. Spell checkers use tries because they are compact and fast. Finally, in genomics, suffix trees are used to align massive DNA sequences in linear time, which is essential for medical research.");
  }

  thankYouSlide(p, H);
  await p.writeFile({ fileName: "./outputs/Sem1_ADS_Unit4_Tries.pptx" });
  convertToPdf("./outputs/Sem1_ADS_Unit4_Tries.pptx");
  console.log("✔  ADS Unit 4 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 5 – KMP ALGORITHM  (Unit V)
// ═══════════════════════════════════════════════════════════════
async function ppt5() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Data Structures";

  titleSlide(p, H, "Unit V: Pattern Matching", "The Knuth-Morris-Pratt (KMP) Algorithm");

  // Slide 2 ── The Pattern Matching Problem
  {
    const s = contentSlide(p, H, "The Pattern Matching Problem");
    s.addText([
      { text: "Pattern Matching ", options: { bold: true } },
      { text: "seeks all occurrences of a pattern string P of length m within a text T of length n.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("The Naive Search Algorithm",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• How it works:", options: { bold: true } }, { text: " Aligns the pattern P at the start of text T. Compares character by character.", options: { breakLine: true } },
      { text: "• Mismatch action:", options: { bold: true } }, { text: " Shift the pattern by exactly ONE position. Reset comparison index and start over.", options: { breakLine: true } },
      { text: "• Worst-case Complexity:", options: { bold: true, color: M } }, { text: " O(n · m). Occurs when text and pattern have highly repetitive prefixes (e.g. T = 'AAAAAAAB', P = 'AAAB').", options: { breakLine: true } },
      { text: "• Limitation:", options: { bold: true } }, { text: " Wastes information. It forgets character matches that have already occurred, resetting the text pointer backwards.", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("The KMP Innovation",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Linear Complexity:", options: { bold: true, color: DGN } }, { text: " O(n + m) worst-case running time. Preprocesses the pattern in O(m) time and searches in O(n) time.", options: { breakLine: true } },
      { text: "• No Text Backtracking:", options: { bold: true } }, { text: " The index scanning the text T NEVER goes backward. We only move the pattern index forward using prior match data.", options: { breakLine: true } },
      { text: "• Key Idea:", options: { bold: true } }, { text: " When a mismatch occurs, we know the prefix of the pattern matched up to that point. We look at the pattern itself to determine where to shift, skipping useless alignments.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 2 – The Pattern Matching Problem\n\nIn string processing, pattern matching is a fundamental task. The naive algorithm aligns the pattern with the text and compares characters. If a mismatch occurs, the pattern is shifted by one slot, and the text pointer backtracks.\n\nIn the worst case, this leads to O(n·m) time complexity, which is highly inefficient. The Knuth-Morris-Pratt (KMP) algorithm, published in 1977, solves this problem. It achieves O(n+m) linear time complexity.\n\nThe core breakthrough of KMP is that the text pointer never backtracks. It only moves forward. When a mismatch occurs, KMP uses a precomputed table to slide the pattern past positions that are guaranteed to fail.");
  }

  // Slide 3 ── The Prefix Function
  {
    const s = contentSlide(p, H, "The Prefix Function (π-table / Failure Function)");
    s.addText("The π-table stores the lengths of the longest proper prefix that is also a suffix of the pattern:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Mathematical Definition",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• π[q]:", options: { bold: true } }, { text: " The length of the longest proper prefix of pattern P[0..q] that is also a suffix of P[0..q].", options: { breakLine: true } },
      { text: "• Proper Prefix:", options: { bold: true } }, { text: " A prefix of string S that is not equal to S itself.", options: { breakLine: true } },
      { text: "• Purpose:", options: { bold: true } }, { text: " If a mismatch occurs at P[q+1], we do not start over. The value π[q] tells us the length of the matching prefix that is already aligned, allowing us to resume search at index P[π[q]].", options: { breakLine: true } },
      { text: "• Preprocessing time:", options: { bold: true, color: DGN } }, { text: " Computed in O(m) time using dynamic programming.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Example: P = 'ABABC'",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });

    const tData = [
      [
        { text: "i", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "P[i]", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "π[i]", options: { bold: true, fill: { color: AM }, color: WH, align: "center" } },
        { text: "Longest Prefix/Suffix Match", options: { bold: true, fill: { color: AM }, color: WH } }
      ],
      [
        { text: "0", options: { align: "center", bold: true } }, { text: "A", options: { align: "center" } },
        { text: "0", options: { align: "center", bold: true, color: M } }, { text: "No proper prefix exists", options: { italic: true } }
      ],
      [
        { text: "1", options: { align: "center", bold: true } }, { text: "B", options: { align: "center" } },
        { text: "0", options: { align: "center", bold: true, color: M } }, { text: "No prefix-suffix match for 'AB'", options: { italic: true } }
      ],
      [
        { text: "2", options: { align: "center", bold: true } }, { text: "A", options: { align: "center" } },
        { text: "1", options: { align: "center", bold: true, color: DGN } }, { text: "Prefix 'A' matches suffix 'A'", options: {} }
      ],
      [
        { text: "3", options: { align: "center", bold: true } }, { text: "B", options: { align: "center" } },
        { text: "2", options: { align: "center", bold: true, color: DGN } }, { text: "Prefix 'AB' matches suffix 'AB'", options: {} }
      ],
      [
        { text: "4", options: { align: "center", bold: true } }, { text: "C", options: { align: "center" } },
        { text: "0", options: { align: "center", bold: true, color: M } }, { text: "No match for 'ABABC'", options: { italic: true } }
      ]
    ];
    s.addTable(tData,
      {
        x: 6.72, y: 2.82, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial",
        border: { pt: 1, color: "CCCCCC" }, colW: [0.6, 0.8, 0.8, 3.83]
      });

    s.addNotes("Slide 3 – The Prefix Function\n\nThe secret to KMP is the prefix function, represented as the pi-table. It tracks repeating patterns within the pattern itself.\n\nFor each index i, we calculate the longest proper prefix of the subpattern P[0..i] that is also a suffix of P[0..i].\n\nLet us trace P = 'ABABC':\n- At i=0 ('A'), pi[0] is 0.\n- At i=1 ('AB'), no prefix matches suffix, so pi[1] is 0.\n- At i=2 ('ABA'), the prefix 'A' matches the suffix 'A', so pi[2] is 1.\n- At i=3 ('ABAB'), the prefix 'AB' matches the suffix 'AB', so pi[3] is 2.\n- At i=4 ('ABABC'), no prefix matches suffix, so pi[4] is 0.\n\nIf we mismatch after matching 'ABAB', we do not go back to the start. The pi-value of 2 tells us that 'AB' is already matched at the start of our pattern, so we align the third character ('A') with the current text position and keep going.");
  }

  // Slide 4 ── KMP Matching Algorithm
  {
    const s = contentSlide(p, H, "KMP Matcher Algorithm");
    s.addText("How the KMP algorithm matches patterns in linear time:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LGN, GN);
    s.addText("KMP Search Loop Logic",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Let text be T of length n, pattern be P of length m. Maintain indices i for T and j for P:", options: { breakLine: true } },
      { text: "1. Initialize:", options: { bold: true } }, { text: " i = 0 (text scan pointer), j = 0 (pattern match pointer). Compute π-table for P.", options: { breakLine: true } },
      { text: "2. Search Loop:", options: { bold: true } }, { text: " While i < n:", options: { breakLine: true } },
      { text: "   • If T[i] == P[j]:", options: { bold: true, color: DGN } }, { text: " Characters match! Increment both: i++ and j++.", options: { breakLine: true } },
      { text: "   • If j == m:", options: { bold: true, color: DGN } }, { text: " Pattern found! Report occurrence at index i - m. Slide pattern: set j = π[j - 1].", options: { breakLine: true } },
      { text: "   • If T[i] != P[j] (Mismatch):", options: { bold: true, color: M } }, { text: " Slide pattern without backtracking i:", options: { breakLine: true } },
      { text: "       - If j > 0: set j = π[j - 1] (shift pattern to align the longest matching prefix).", options: { breakLine: true } },
      { text: "       - If j == 0: cannot shift. Simply move text pointer forward: i++.", options: {} }
    ], { x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 4 – KMP Search Loop\n\nThis slide displays the core loop logic of the KMP search phase. We maintain index i in the text and index j in the pattern.\n\nOn character match, we increment both indices. If the pattern is fully matched (j == m), we report a match and slide the pattern using the pi-table value.\n\nOn mismatch: if j is greater than 0, we update j to pi[j-1]. This shifts the pattern to align the next best prefix. If j is 0, we cannot shift the pattern, so we simply increment the text pointer i.\n\nNotice that there is no statement decrementing i. This guarantees that the text is scanned in a single forward pass, which is ideal for streaming large files.");
  }

  // Slide 5 ── Complexity and Correctness
  {
    const s = contentSlide(p, H, "Complexity & Correctness Proof");
    s.addText("Mathematical analysis of KMP's performance bounds:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Time Complexity Proof",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Split phase costs:", options: { bold: true } }, { text: " Preprocessing takes O(m) steps. Searching takes O(n) steps. Total time is O(n + m).", options: { breakLine: true } },
      { text: "• Search Cost Analysis (Amortized):", options: { bold: true, color: BL } }, { text: " Let's count the number of comparisons. The loop runs at most 2n times because:", options: { breakLine: true } },
      { text: "   - The index i increases in every step (either by matching or by j == 0 mismatch). It never decreases. Hence, at most n increments.", options: { breakLine: true } },
      { text: "   - The index j changes. We increment j when characters match (at most n times). When we reset j = π[j-1] on mismatch, j decreases.", options: { breakLine: true } },
      { text: "   - Since j starts at 0, is incremented at most n times, and can never drop below 0, it can decrease at most n times. Hence, j = π[j-1] can be executed at most n times.", options: { breakLine: true } },
      { text: "• Total operations ≤ 2n.", options: { bold: true } }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LMR, M);
    s.addText("Correctness and Safety",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• No Missed Matches:", options: { bold: true } }, { text: " We must prove that KMP never skips a valid alignment of the pattern.", options: { breakLine: true } },
      { text: "• Shift Safety Proof:", options: { bold: true, color: M } }, { text: " Suppose we mismatch at P[j]. We know P[0..j-1] matches T[i-j..i-1]. We shift P by setting j = π[j-1].", options: { breakLine: true } },
      { text: "   - This is the largest shift size where the new alignment could possibly match.", options: { breakLine: true } },
      { text: "   - Any smaller shift (which corresponds to a longer prefix matching a suffix of P[0..j-1]) is checked because π[j-1] is defined as the MAXIMUM prefix-suffix overlap.", options: { breakLine: true } },
      { text: "   - Any intermediate shift cannot match because it would require a suffix of P[0..j-1] to match a prefix of P, which is mathematically impossible since π[j-1] is the longest such match.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Slide 5 – Complexity & Correctness\n\nHow do we prove KMP is O(n+m)? We use aggregate analysis. The preprocessing takes O(m) steps. For the search phase, we track the values of i and j.\n\nIn each iteration, either i increases (which can happen at most n times), or j decreases (via the pi table). Since j is only incremented when i is incremented, the total number of decrements to j is bounded by the total number of increments, which is at most n.\n\nTherefore, the total number of steps in the search loop is at most 2n, proving the O(n) complexity.\n\nWe also prove correctness: why is KMP safe? Because by shifting by the longest proper prefix-suffix match, we guarantee that we do not skip any possible alignment that could match, while skipping all alignments that are mathematically guaranteed to fail.");
  }

  // Slide 6 ── Applications of KMP
  {
    const s = contentSlide(p, H, "Applications of KMP in Computer Science");
    s.addText("KMP is widely used in systems requiring fast, streaming pattern searching:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🔍  Text Editors & IDE Search",
        b: "Modern editors (like VS Code, Notepad++) use KMP or Boyer-Moore algorithms to implement 'Find' operations, ensuring instantaneous search and replace in large text files.", bg: LGN, bd: GN
      },
      {
        t: "🛡️  Intrusion Detection Systems (IDS)",
        b: "Network packet scanners (like Snort) use KMP to scan live packets against database signatures of malware and cyber threats. Since KMP does not backtrack, it works on streaming network traffic.", bg: LBL, bd: BL
      },
      {
        t: "🧬  Bioinformatics & DNA Alignments",
        b: "Genomics software uses KMP to search for specific gene signatures and mutations inside massive DNA sequences containing billions of base pairs.", bg: LMR, bd: M
      },
      {
        t: "🌐  Log Analysis Tools",
        b: "Tools like grep or log parsers apply KMP to filter out error logs or extract warnings from massive server log files in a single pass.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Slide 6 – Applications\n\nKMP is the algorithm of choice for streaming data. In Network Intrusion Detection Systems (IDS), we cannot afford to buffer entire packet streams to backtrack. KMP's forward-only text scanning is perfect here.\n\nIt is also used in IDE search-and-replace, DNA sequence matching, and command-line search utilities like grep. This wraps up our presentations on Advanced Data Structures.");
  }

  thankYouSlide(p, H);
  await p.writeFile({ fileName: "./outputs/Sem1_ADS_Unit5_KMP_Algorithm.pptx" });
  convertToPdf("./outputs/Sem1_ADS_Unit5_KMP_Algorithm.pptx");
  console.log("✔  ADS Unit 5 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 5 ADS PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  await ppt4();
  await ppt5();
  console.log("All 5 ADS PPTs generated successfully!");
})();
