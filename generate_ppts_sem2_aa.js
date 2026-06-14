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
//  PPT 1 – STRONGLY CONNECTED COMPONENTS  (Unit I)
// ═══════════════════════════════════════════════════════════════
async function ppt1() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Algorithms";

  const tSl = titleSlide(p, H, "Unit I: Sorting & Graph Algorithms", "Strongly Connected Components (SCC)");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Strongly Connected Components' as part of the Advanced Algorithms course for my M.Tech in Computer Science and Engineering. In this presentation, we will define strongly connected components in directed graphs, detail Kosaraju's algorithm for finding them in linear time, analyze the correctness of the two-pass DFS approach, compare it to Tarjan's single-pass method, and discuss various real-world applications in network and database management.");

  // Slide 2 ── What is a Strongly Connected Component?
  {
    const s = contentSlide(p, H, "What is a Strongly Connected Component?");
    s.addText([
      { text: "A Strongly Connected Component (SCC) ", options: { bold: true } },
      { text: "of a directed graph G = (V, E) is a maximal set of vertices U ⊆ V such that for every pair of vertices u, v ∈ U, there is a directed path from u to v and a directed path from v to u.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Mathematical Properties",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Connectivity equivalence:", options: { bold: true } }, { text: " The 'path connection' relation (u ~ v iff u is reachable from v and v from u) is an equivalence relation on V.", options: { breakLine: true } },
      { text: "• Partitioning:", options: { bold: true, color: BL } }, { text: " The equivalence classes of this relation partition the graph G into disjoint SCCs.", options: { breakLine: true } },
      { text: "• Component Graph (DAG):", options: { bold: true } }, { text: " If we collapse each SCC into a single super-node, the resulting graph of components is a Directed Acyclic Graph (DAG).", options: { breakLine: true } },
      { text: "• Cycles:", options: { bold: true } }, { text: " An SCC contains at least one cycle (unless it is a single vertex with no self-loop).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LGN, GN);
    s.addText("Reachability & DFS Trees",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• DFS Tree structure:", options: { bold: true } }, { text: " Running Depth-First Search (DFS) on G partitions G into a forest of DFS trees.", options: { breakLine: true } },
      { text: "• Reachability Bounds:", options: { bold: true } }, { text: " If two vertices u and v belong to the same SCC, they are guaranteed to be in the same DFS tree during any search.", options: { breakLine: true } },
      { text: "• Transpose Graph Gᵀ:", options: { bold: true, color: M } }, { text: " Gᵀ is the same graph G but with all edge directions reversed. G and Gᵀ have the EXACT same SCCs because reachability is symmetric.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by defining what a strongly connected component, or SCC, is. In a directed graph, connectivity is not symmetric. An SCC is a maximal subset of vertices where every vertex in the set can reach every other vertex. Under this definition, reachability acts as an equivalence relation, which partitions the graph's vertices into disjoint components. If we contract each SCC into a single super-node, the resulting graph is guaranteed to be a Directed Acyclic Graph, or DAG, because any cycle between components would merge them. To analyze these components, we run Depth-First Searches to build DFS trees. Additionally, we consider the transpose graph G-T, which is the same graph but with reversed edges. G and G-T share the exact same SCCs because reversing all edges in a cycle does not break the reachability between any pair of vertices.");
  }

  // Slide 3 ── Kosaraju's Algorithm
  {
    const s = contentSlide(p, H, "Kosaraju's DFS-Based Algorithm");
    s.addText("Kosaraju-Sharir algorithm computes all SCCs of a directed graph in O(V + E) time:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 12.35, 4.80, LYL, AM);
    s.addText("The Two-Pass DFS Algorithm Steps",
      { x: 0.62, y: 2.38, w: 12.11, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Run First DFS on Graph G:", options: { bold: true } }, { text: " Traverse G using DFS. As each vertex finishes execution (when all its outgoing edges are explored), push it onto an empty stack S.", options: { breakLine: true } },
      { text: "2. Reverse Graph to Get Gᵀ:", options: { bold: true, color: M } }, { text: " Compute the transpose graph by reversing the directions of all edges in G.", options: { breakLine: true } },
      { text: "3. Run Second DFS on Gᵀ in Stack Order:", options: { bold: true } }, { text: " While stack S is not empty:\n", options: { breakLine: true } },
      { text: "   • Pop the top vertex v from stack S.\n", options: { color: BK } },
      { text: "   • If v has not been visited in this second pass, it is the root of a new SCC.\n", options: { color: BK } },
      { text: "   • Run DFS starting at v in Gᵀ. All vertices reached from v in Gᵀ belong to the same SCC as v.\n", options: { color: BK } },
      { text: "   • Output this group as a completed SCC and mark them as visited.\n", options: { breakLine: true } },
      { text: "4. Time Complexity:", options: { bold: true, color: DGN } }, { text: " O(V + E) since it runs two standard DFS scans and one graph transposition.", options: {} }
    ], { x: 0.62, y: 2.85, w: 12.11, h: 4.00, fontSize: 13.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Next, let us explore Kosaraju's algorithm, which computes all strongly connected components of a directed graph in linear time, specifically O of V plus E. The algorithm is remarkably elegant and operates in two DFS passes. In the first pass, we run a standard Depth-First Search on the original graph G. As each vertex finishes—meaning all its outgoing edges have been fully explored—we push the vertex onto a global stack. This stack tracks the post-order finishing times. In step two, we construct the transpose graph G-T by reversing every edge. In the third step, we run a second DFS, but we process vertices by popping them from the stack. Each unvisited node we pop becomes the root of a new strongly connected component. We run a DFS from this root in G-T, and all reachable vertices belong to its SCC. Because we run two DFS searches, the total time complexity is linear.");
  }

  // Slide 4 ── Correctness and Transpose Proof
  {
    const s = contentSlide(p, H, "Why It Works: Correctness Proof");
    s.addText("Understanding the mathematical mechanism behind Kosaraju's algorithm:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("The Leakage Problem & Transpose",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Why not run DFS twice on G?", options: { bold: true, color: M } }, { text: " If we simply run DFS on G, the search will leak from one SCC to another connected downstream, merging them incorrectly.", options: { breakLine: true } },
      { text: "• The Transpose Shield:", options: { bold: true, color: DGN } }, { text: " In Gᵀ, edges go backward. If there is a directed path SCC1 → SCC2 in G, then in Gᵀ the path goes SCC2 → SCC1.", options: { breakLine: true } },
      { text: "• The Stack Order:", options: { bold: true } }, { text: " The first DFS finish order guarantees that the top node of the stack belongs to a 'source' SCC in G (meaning it has no incoming edges from other SCCs in Gᵀ).", options: { breakLine: true } },
      { text: "• Clean Isolation:", options: { bold: true } }, { text: " When we run DFS in Gᵀ starting at this popped root, it cannot escape to other components, isolating the SCC perfectly.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LBL, BL);
    s.addText("Kosaraju's vs. Tarjan's Algorithm",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Tarjan's SCC Algorithm:", options: { bold: true } }, { text: " Another O(V + E) method that finds SCCs in a SINGLE DFS pass.", options: { breakLine: true } },
      { text: "• How Tarjan's Works:", options: { bold: true, color: BL } }, { text: " Tracks DFS discovery numbers and the lowest reachable node index ('low-link' values) on an active stack.", options: { breakLine: true } },
      { text: "• Comparison:", options: { bold: true } }, { text: "\n", options: { breakLine: true } },
      { text: "  - Kosaraju's:", options: { bold: true } }, { text: " Conceptually simpler, easier to implement, but requires transposing the graph and running two DFS passes.\n", options: { breakLine: true } },
      { text: "  - Tarjan's:", options: { bold: true } }, { text: " More complex logic, but faster in practice since it only traverses the graph once and needs no transpose step.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("To understand why Kosaraju's algorithm works, we must analyze the role of the transpose graph and stack order. If we ran a second DFS on the original graph G without reversing edges, the search would leak from one SCC to downstream SCCs, incorrectly merging them. The transpose graph acts as a shield: reversing the edges turns a path from SCC-one to SCC-two into a path from SCC-two to SCC-one. Since the stack order guarantees that the top vertex belongs to a source component in G, it becomes a sink component in G-T. When we start DFS from this node in G-T, the reversed edges prevent the search from leaking to any other components, isolating the SCC perfectly. We can compare Kosaraju's algorithm to Tarjan's algorithm, which computes SCCs in a single pass by tracking discovery times and low-link values on an active stack. While Tarjan's is slightly faster, Kosaraju's is much simpler to implement.");
  }

  // Slide 5 ── Applications of SCCs
  {
    const s = contentSlide(p, H, "Applications of SCCs");
    s.addText("Strongly Connected Components model dependencies and structures in complex networks:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🌐  Social Network Communities",
        b: "Identifies tightly knit groups of users where everyone interacts or shares info. Used to analyze information spread and recommend connections.", bg: LGN, bd: GN
      },
      {
        t: "💻  Software Dependency Loops",
        b: "Compilers build module import graphs. Finding SCCs reveals circular imports (e.g. A imports B, B imports C, C imports A) which must be resolved.", bg: LBL, bd: BL
      },
      {
        t: "🔍  Web Page Ranking (Google PageRank)",
        b: "The web forms a directed graph of links. PageRank partitions the web into SCCs to calculate page importances without getting stuck in link cycles.", bg: LMR, bd: M
      },
      {
        t: "⚙️  Deadlock Detection in DBMS",
        b: "Database engines maintain 'wait-for' graphs of transactions waiting for locks. An SCC in this graph indicates a cycle, revealing a system deadlock.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Finally, let us discuss the practical applications of strongly connected components in computer science. First, in social networks, SCCs identify tightly knit communities where information flows bidirectionally. Second, in software development, compilers build module dependency graphs. Finding SCCs in these graphs helps detect circular imports, such as class A importing B, B importing C, and C importing A, which must be refactored. Third, Google's PageRank algorithm uses SCC partitions to navigate the web graph without getting stuck in infinite link loops. Lastly, database engines construct wait-for graphs to manage transactions. A strongly connected component in a wait-for graph indicates a transaction cycle, which reveals a system deadlock that the database manager must resolve by aborting a transaction.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, strongly connected components are a fundamental concept for analyzing directed graph connectivity. Algorithms like Kosaraju's and Tarjan's allow us to partition complex networks and detect cycles in O(V + E) time. Thank you for your time, and I am happy to take any questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_AA_Unit1_SCC.pptx" });
  convertToPdf("./outputs/Sem2_AA_Unit1_SCC.pptx");
  console.log("✔  AA Unit 1 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 2 – MATROIDS & THE GREEDY PARADIGM  (Unit II)
// ═══════════════════════════════════════════════════════════════
async function ppt2() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Algorithms";

  const tSl = titleSlide(p, H, "Unit II: Greedy Paradigm & Matroids", "Matroids & The Greedy Paradigm");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on 'Matroids and the Greedy Paradigm' as part of the Advanced Algorithms course for my M.Tech in Computer Science and Engineering. In this presentation, we will define the mathematical structure of matroids, examine the general greedy algorithm on matroids, prove the correctness using the Rado-Edmonds theorem, apply these concepts to Kruskal's Minimum Spanning Tree algorithm, and discuss wider applications of matroids in scheduling and coding theory.");

  // Slide 2 ── What is a Matroid?
  {
    const s = contentSlide(p, H, "What is a Matroid?");
    s.addText([
      { text: "A Matroid ", options: { bold: true } },
      { text: "is a mathematical structure that generalizes the concept of linear independence in vector spaces. It defines a family of subsets where a simple greedy algorithm is guaranteed to find an optimal solution.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LGN, GN);
    s.addText("Formal Definition M = (S, I)",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "A matroid is an ordered pair (S, I) satisfying three axioms:", options: { breakLine: true } },
      { text: "1. Finite Ground Set S:", options: { bold: true, color: DGN } }, { text: " A non-empty finite set of elements.", options: { breakLine: true } },
      { text: "2. Hereditary Property:", options: { bold: true, color: DGN } }, { text: " I is a non-empty family of subsets of S (independent sets) such that if B ∈ I and A ⊆ B, then A ∈ I. (Subsets of independent sets are independent).", options: { breakLine: true } },
      { text: "3. Exchange Property:", options: { bold: true, color: DGN } }, { text: " If A, B ∈ I and |A| < |B|, there exists some element x ∈ B - A such that A ∪ {x} ∈ I. (We can extend independent sets).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LBL, BL);
    s.addText("Vector Space Analogy",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Matroid axioms match properties of vectors:", options: { breakLine: true } },
      { text: "• Ground Set S:", options: { bold: true } }, { text: " A set of vectors in a vector space.", options: { breakLine: true } },
      { text: "• Independent Sets I:", options: { bold: true } }, { text: " Subsets of vectors that are linearly independent (no vector is a linear combination of others).", options: { breakLine: true } },
      { text: "• Hereditary:", options: { bold: true } }, { text: " Subsets of linearly independent vectors are also linearly independent. ✔", options: { breakLine: true } },
      { text: "• Exchange:", options: { bold: true, color: M } }, { text: " If you have a basis of size 3 and another of size 4, you can find a vector in the larger set to extend the smaller one. ✔", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us begin by defining what a matroid is. A matroid is a mathematical structure that generalizes the concept of linear independence from vector spaces to general sets. Formally, a matroid M is represented as a pair containing a ground set S and a family of independent subsets I. It must satisfy three fundamental axioms. First, the ground set S must be finite. Second, the hereditary property, which states that if a set B is independent, then any subset A of B must also be independent. Third, the exchange property, which states that if we have two independent sets A and B, where A is smaller than B, there exists an element in B minus A that can be added to A to form a larger independent set. This matches vector space properties, where subsets of linearly independent vectors are independent, and a smaller set of vectors can be extended using vectors from a larger spanning set.");
  }

  // Slide 3 ── The Greedy Algorithm on Matroids
  {
    const s = contentSlide(p, H, "The Matroid Greedy Algorithm");
    s.addText("If we assign weights to ground elements, a greedy algorithm finds a maximum-weight independent set:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("The General Greedy Algorithm",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Given Matroid M = (S, I) and weight function w(x) > 0 for x ∈ S:\n", options: { breakLine: true } },
      { text: "1. Sort Ground elements:", options: { bold: true } }, { text: " Sort all x ∈ S in descending order of weights:\n   w(x₁) ≥ w(x₂) ≥ … ≥ w(xₙ).", options: { breakLine: true } },
      { text: "2. Initialize Set:", options: { bold: true, color: BL } }, { text: " Set A = ∅ (empty set).", options: { breakLine: true } },
      { text: "3. Greedy Loop:", options: { bold: true } }, { text: " For i = 1 to n:\n", options: { breakLine: true } },
      { text: "   • If A ∪ {x_i} ∈ I (independent):\n", options: { color: BK } },
      { text: "   • Update: A = A ∪ {x_i}.\n", options: { color: BK } },
      { text: "4. Return A:", options: { bold: true, color: DGN } }, { text: " A is guaranteed to be a maximum-weight maximal independent set (basis) of the Matroid.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Why Greedy Works (Matroid Theorem)",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Rado-Edmonds Theorem:", options: { bold: true, color: AM } }, { text: " A hereditary system (S, I) is a matroid if and only if the greedy algorithm finds an optimal independent set for all positive weight functions.", options: { breakLine: true } },
      { text: "• Optimization Guarantee:", options: { bold: true } }, { text: " The exchange property guarantees that a local greedy choice never blocks reaching the global optimum.", options: { breakLine: true } },
      { text: "• Contrast with Non-Matroids:", options: { bold: true, color: M } }, { text: " In non-matroid systems (e.g. Knapsack, TSP), greedy decisions can lead to sub-optimal solutions. Matroids are the boundary of greedy solvability.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Now, let us examine the greedy algorithm as applied to matroids. If we assign positive weights to the elements of our ground set S, we want to find an independent set in I that has the maximum total weight. The matroid greedy algorithm works by first sorting all elements in descending order of their weights. We initialize our target set A as an empty set. Then, we loop through the sorted elements. For each element, we check if adding it to A maintains independence. If it does, we add it to A; otherwise, we skip it. The Rado-Edmonds theorem states that a hereditary system is a matroid if and only if this simple greedy approach is guaranteed to find an optimal maximum-weight independent set. In non-matroid systems, like the knapsack problem, local greedy choices lead to suboptimal results, meaning matroids define the mathematical limits of greedy solvability.");
  }

  // Slide 4 ── Application to Minimum Spanning Tree
  {
    const s = contentSlide(p, H, "Application: Kruskal's MST Algorithm");
    s.addText("Kruskal's Minimum Spanning Tree algorithm is a direct application of the Matroid Greedy algorithm:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Graphic Matroids",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "We can model graphs as matroids. Given G = (V, E):\n", options: { breakLine: true } },
      { text: "• Ground Set S:", options: { bold: true } }, { text: " The edge set E of G.", options: { breakLine: true } },
      { text: "• Independent Sets I:", options: { bold: true, color: BL } }, { text: " Any subset of edges that is acyclic (contains no cycles, forming a forest).", options: { breakLine: true } },
      { text: "• Hereditary Property:", options: { bold: true } }, { text: " Subsets of acyclic edges are also acyclic. ✔", options: { breakLine: true } },
      { text: "• Exchange Property:", options: { bold: true } }, { text: " If forest A has fewer edges than forest B, we can always find an edge in B to add to A without forming a cycle. ✔", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LGN, DGN);
    s.addText("Kruskal's Code Map",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: DGN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "Kruskal's steps map directly to the Matroid algorithm:", options: { breakLine: true } },
      { text: "1. Sort Edges:", options: { bold: true } }, { text: " Sort edges by weight in ascending order (we find the minimum spanning tree, which is equivalent to maximizing negative weights).", options: { breakLine: true } },
      { text: "2. Select Edges:", options: { bold: true } }, { text: " Add edges one-by-one to our MST.", options: { breakLine: true } },
      { text: "3. Cycle Check (Independence):", options: { bold: true, color: DGN } }, { text: " Only add an edge if it does not form a cycle with already-selected edges (checking independence in the graphic matroid).", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("A classic application of matroid theory is Kruskal's algorithm for finding a Minimum Spanning Tree. We can model a graph G as a graphic matroid where the ground set S represents the graph's edges, and the independent sets I represent any subset of edges that is acyclic. This structure satisfies the hereditary property because any subset of a forest is also a forest. It satisfies the exchange property because if we have two forests A and B, and A has fewer edges, we can always find an edge in B that connects two separate trees in A, allowing us to add it without creating a cycle. Kruskal's algorithm sorts edges by weight and greedily adds them only if they do not create a cycle. This matches the matroid greedy template exactly, proving mathematically why Kruskal's algorithm is guaranteed to find the optimal MST.");
  }

  // Slide 5 ── Applications of Matroids
  {
    const s = contentSlide(p, H, "Applications of Matroids in CS");
    s.addText("Matroid structures guarantee optimal greedy solutions across scheduling and network designs:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "🌲  Kruskal's MST Networks",
        b: "Graphic matroids prove the correctness of Kruskal's MST algorithm. Used to construct optimal telecommunication, power grid, and computer network wiring systems.", bg: LGN, bd: GN
      },
      {
        t: "📅  Task Scheduling with Deadlines",
        b: "Scheduling unit-time jobs with deadlines and penalties is modeled as a matroid. The greedy algorithm yields the optimal schedule that minimizes penalties.", bg: LBL, bd: BL
      },
      {
        t: "🕸️  Network Flow & Matrix Matching",
        b: "Algebraic matroids solve structural matching in matrix computations, optimizing index allocations for high-performance solvers.", bg: LMR, bd: M
      },
      {
        t: "🔒  Coding Theory & Cryptography",
        b: "Matroid theory is utilized in error-correcting codes and secret sharing schemes to design linear bounds for secure communications.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Beyond minimum spanning trees, matroids have several key applications in computer science. First, they are used to solve unit-time task scheduling problems with deadlines and penalties, where greedily scheduling jobs with the highest penalties first minimizes total loss. Second, algebraic matroids help optimize sparse matrix matchings and row reductions in high-performance computing solvers. Third, matroid theory is applied in coding theory and cryptography to establish linear bounds for error-correcting codes and design secure secret-sharing schemes. By framing these complex problems within matroid theory, designers can replace complex dynamic programming solvers with simple, efficient greedy algorithms that run in O of n log n time while guaranteeing global optimality.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, matroids provide the theoretical foundation that guarantees when greedy algorithms will find globally optimal solutions. By recognizing matroid properties in graph structures, scheduling problems, and vector spaces, we can design simple and highly efficient algorithms. Thank you for your attention, and I am ready to answer any questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_AA_Unit2_Matroids.pptx" });
  convertToPdf("./outputs/Sem2_AA_Unit2_Matroids.pptx");
  console.log("✔  AA Unit 2 PPT generated");
}

// ═══════════════════════════════════════════════════════════════
//  PPT 3 – MAXFLOW-MINCUT THEOREM  (Unit III)
// ═══════════════════════════════════════════════════════════════
async function ppt3() {
  const p = new pptxgen(); p.layout = "LAYOUT_WIDE";
  const H = "Advanced Algorithms";

  const tSl = titleSlide(p, H, "Unit III: Flow Networks & Matrix Computations", "Maxflow-Mincut Theorem & Ford-Fulkerson");
  tSl.addNotes("Hello everyone, my name is Srinivas Rao Tammireddy, and today I will be presenting on the 'Maxflow-Mincut Theorem and the Ford-Fulkerson Method' as part of the Advanced Algorithms course for my M.Tech in Computer Science and Engineering. In this presentation, we will define flow networks and their constraints, explain the Ford-Fulkerson method and its Edmonds-Karp optimization, prove correctness using the Maxflow-Mincut theorem, and look at practical applications such as bipartite matching, image segmentation, and network routing.");

  // Slide 2 ── Flow Networks & Constraints
  {
    const s = contentSlide(p, H, "Flow Networks & Constraints");
    s.addText([
      { text: "A Flow Network ", options: { bold: true } },
      { text: "is a directed graph G = (V, E) where each edge has a positive capacity c(u, v) > 0. We model the flow of material from a source s to a sink t under two constraints.", options: {} }
    ], { x: CX, y: CYC, w: CW, h: 0.55, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.26, LBL, BL);
    s.addText("Core Mathematical Constraints",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "A flow f on G must satisfy two properties for all u, v ∈ V:", options: { breakLine: true } },
      { text: "1. Capacity Constraint:", options: { bold: true, color: M } }, { text: " The flow along any edge cannot exceed its capacity:\n   0 ≤ f(u, v) ≤ c(u, v).", options: { breakLine: true } },
      { text: "2. Flow Conservation:", options: { bold: true, color: DGN } }, { text: " For all nodes except source s and sink t, the total incoming flow must equal the total outgoing flow:\n   ∑ f(w, u) = ∑ f(u, w). (No material is stored or lost).", options: {} }
    ], { x: 0.62, y: 2.80, w: 5.70, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.26, LGN, GN);
    s.addText("Residual Networks & Augmentation",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: GN, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Residual Capacity c_f:", options: { bold: true } }, { text: " The remaining capacity on an edge:\n   c_f(u, v) = c(u, v) − f(u, v) (forward capacity),\n   c_f(v, u) = f(u, v) (unused flow that can be pushed back).", options: { breakLine: true } },
      { text: "• Residual Graph G_f:", options: { bold: true } }, { text: " A graph containing edges with positive residual capacity.", options: { breakLine: true } },
      { text: "• Augmenting Path:", options: { bold: true, color: DGN } }, { text: " A simple path from s to t in the residual graph G_f. Allows us to push additional flow through the network.", options: {} }
    ], { x: 6.72, y: 2.80, w: 6.03, h: 3.68, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Let us start by defining flow networks. A flow network is a directed graph where each edge has a positive capacity. We model the flow of material from a source node s to a sink node t. The flow along any edge must satisfy two core constraints. First, the capacity constraint, which states that the flow on an edge cannot exceed its capacity. Second, flow conservation, which requires that for every node other than the source and the sink, the total incoming flow must equal the total outgoing flow. To optimize this flow, we define the residual network, which represents the remaining capacity on each edge, and includes reverse edges to push back existing flow. An augmenting path is a simple path from source to sink in the residual network. Finding these paths allows us to increase the total flow through the network.");
  }

  // Slide 3 ── The Ford-Fulkerson Method
  {
    const s = contentSlide(p, H, "The Ford-Fulkerson Method");
    s.addText("Ford-Fulkerson iteratively increases flow along augmenting paths until none remain:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LMR, M);
    s.addText("Ford-Fulkerson Method Steps",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "1. Initialize Flow:", options: { bold: true } }, { text: " Set f(u, v) = 0 for all edges.", options: { breakLine: true } },
      { text: "2. Find Augmenting Path:", options: { bold: true, color: BL } }, { text: " Search G_f for a path P from s to t.", options: { breakLine: true } },
      { text: "3. Find Bottleneck Capacity (c_f(P)):", options: { bold: true } }, { text: " Locate the minimum residual capacity in path P:\n   c_f(P) = min { c_f(u, v) for (u, v) in P }.", options: { breakLine: true } },
      { text: "4. Augment Flow:", options: { bold: true, color: DGN } }, { text: " Push c_f(P) along the path. For each edge (u, v) in P, set f(u, v) += c_f(P).", options: { breakLine: true } },
      { text: "5. Repeat:", options: { bold: true } }, { text: " Reconstruct G_f and repeat until no path from s to t exists.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LYL, AM);
    s.addText("Edmonds-Karp Optimization",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: AM, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Path Search Selection:", options: { bold: true } }, { text: " Ford-Fulkerson does not define how to search for paths. Using DFS can lead to slow execution times if capacity values are large.", options: { breakLine: true } },
      { text: "• Edmonds-Karp Improvement:", options: { bold: true, color: DGN } }, { text: " Uses Breadth-First Search (BFS) to select the shortest augmenting path (minimum number of edges) in each iteration.", options: { breakLine: true } },
      { text: "• Time Complexity:", options: { bold: true } }, { text: " Edmonds-Karp is guaranteed to run in O(V · E²) time, completely independent of the edge capacities.", options: {} }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("Next, let us look at the Ford-Fulkerson method. This method starts with zero flow on all edges and iteratively increases flow. In each iteration, we search for an augmenting path from source to sink in the residual network. Once a path is found, we identify the bottleneck edge with the minimum residual capacity along that path. We then augment the flow by adding this bottleneck value to forward edges and subtracting it from reverse edges. We repeat this process until no augmenting paths exist. The Edmonds-Karp algorithm improves Ford-Fulkerson by using Breadth-First Search to find the shortest augmenting path. This BFS selection guarantees that the algorithm terminates in O of V times E squared time, making the execution time independent of edge capacities.");
  }

  // Slide 4 ── The Maxflow-Mincut Theorem
  {
    const s = contentSlide(p, H, "The Maxflow-Mincut Theorem");
    s.addText("The Maxflow-Mincut theorem proves the correctness of the Ford-Fulkerson method:",
      { x: CX, y: CYC, w: CW, h: 0.50, fontSize: 14, fontFace: "Arial", color: BK });

    box(p, s, CX, 2.30, 5.90, 4.50, LBL, BL);
    s.addText("Cuts in Flow Networks",
      { x: 0.62, y: 2.38, w: 5.70, h: 0.38, fontSize: 15, bold: true, color: BL, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "• Defining a Cut (S, T):", options: { bold: true } }, { text: " A partition of V into two disjoint sets S and T such that source s ∈ S and sink t ∈ T.", options: { breakLine: true } },
      { text: "• Capacity of a Cut c(S, T):", options: { bold: true, color: M } }, { text: " The sum of capacities of all edges crossing the cut from S to T:\n   c(S, T) = ∑ c(u, v) for u in S, v in T.", options: { breakLine: true } },
      { text: "• Flow Across a Cut f(S, T):", options: { bold: true } }, { text: " Net flow is the sum of flows from S to T minus the return flows from T to S. For any flow, f(S, T) is equal to the total value of the flow |f|.", options: {} }
    ], { x: 0.62, y: 2.85, w: 5.70, h: 3.80, fontSize: 13, fontFace: "Arial", color: BK, margin: 0 });

    box(p, s, 6.60, 2.30, 6.25, 4.50, LMR, M);
    s.addText("The Theorem Statement & Equivalence",
      { x: 6.72, y: 2.38, w: 6.03, h: 0.38, fontSize: 15, bold: true, color: M, fontFace: "Arial", margin: 0 });
    s.addText([
      { text: "The Theorem (L.R. Ford & D.R. Fulkerson, 1956):\n", options: { bold: true, color: M } },
      { text: "The maximum value of a flow in a network is equal to the minimum capacity of a cut that separates the source from the sink.\n\n", options: { bold: true, color: BK } },
      { text: "Equivalence Conditions:\n", options: { bold: true, color: DGN } },
      { text: "1. f is a maximum flow in G.\n", options: { color: BK } },
      { text: "2. The residual network G_f contains no augmenting paths.\n", options: { color: BK } },
      { text: "3. |f| = c(S, T) for some cut (S, T).", options: { color: BK } }
    ], { x: 6.72, y: 2.85, w: 6.03, h: 3.80, fontSize: 12.5, fontFace: "Arial", color: BK, margin: 0 });

    s.addNotes("The correctness of the Ford-Fulkerson method is established by the Maxflow-Mincut theorem. To understand the theorem, we define a cut as a partition of the vertices into two sets, S and T, with the source in S and the sink in T. The capacity of a cut is the sum of all edge capacities going from S to T. The net flow across any cut equals the total flow of the network. The Maxflow-Mincut theorem states that the maximum flow in a network is exactly equal to the minimum capacity of any cut separating the source and sink. At maximum flow, the residual network contains no augmenting paths, and the edges crossing the minimum cut are completely saturated. This equivalence bridges flow maximization with cut capacity minimization.");
  }

  // Slide 5 ── Applications of Maxflow
  {
    const s = contentSlide(p, H, "Applications of Maxflow");
    s.addText("Maxflow-Mincut models many scheduling and resource allocation problems in CS:",
      { x: CX, y: CYC, w: CW, h: 0.40, fontSize: 14, fontFace: "Arial", color: BK });

    const apps = [
      {
        t: "👥  Maximum Bipartite Matching",
        b: "Matches candidates to jobs, or tasks to processors. Model as a flow network by connecting source s to candidates, candidates to jobs, and jobs to sink t with capacity 1. Max flow yields the maximum matching size.", bg: LGN, bd: GN
      },
      {
        t: "🌐  Data Routing in Networks",
        b: "Determines the maximum data transmission rate between server routers over channels with capacity limits, routing traffic without congestion.", bg: LBL, bd: BL
      },
      {
        t: "🎨  Image Segmentation in CV",
        b: "Divides an image into foreground and background. Pixels are nodes; edges are color similarities. The min-cut partition separates the foreground from the background optimally.", bg: LMR, bd: M
      },
      {
        t: "📅  Airline Crew Scheduling",
        b: "Schedules flight crews to flight paths while respecting constraints (rest hours, locations). Modeled as a circulation flow network with demands.", bg: LYL, bd: AM
      }
    ];

    apps.forEach((app, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const bx = CX + col * 6.22, by = 2.20 + row * 2.30;
      box(p, s, bx, by, 6.05, 2.14, app.bg, app.bd);
      boxTitle(s, bx, by, 6.05, app.t, app.bd);
      boxBody(s, bx, by, 6.05, 2.14, app.b);
    });

    s.addNotes("Finally, let us explore the diverse applications of the Maxflow-Mincut theorem. First, it solves the maximum bipartite matching problem, which pairs elements of two sets, like matching job candidates to open positions. Second, it optimizes data routing in computer networks to maximize bandwidth without exceeding channel limits. Third, in computer vision, min-cut algorithms segment images into foreground and background by partitioning pixels based on color similarity. Fourth, it schedules airline crews to flights while meeting work hour and location constraints. By modeling these scheduling and optimization problems as flow networks, we can solve them efficiently using standard max-flow algorithms.");
  }

  const tySl = thankYouSlide(p, H);
  tySl.addNotes("In conclusion, the Maxflow-Mincut theorem is a powerful optimization duality that guarantees the correctness of flow algorithms. From network routing to image processing, modeling systems as flow networks provides elegant and efficient solutions to complex problems. Thank you for your time, and I am happy to take any questions.");
  await p.writeFile({ fileName: "./outputs/Sem2_AA_Unit3_Maxflow.pptx" });
  convertToPdf("./outputs/Sem2_AA_Unit3_Maxflow.pptx");
  console.log("✔  AA Unit 3 PPT generated");
}

// ─────────────────────────────────────────────────────────────
//  MAIN
// ─────────────────────────────────────────────────────────────
(async () => {
  const outputDir = "./outputs";
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log("Generating 3 AA PPTs …");
  await ppt1();
  await ppt2();
  await ppt3();
  console.log("All 3 AA PPTs generated successfully!");
})();
