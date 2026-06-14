const { execSync } = require("child_process");
const fs = require("fs");

console.log("=========================================");
console.log("STARTING SEMESTER 1 POWERPOINT GENERATION");
console.log("=========================================\n");

// Ensure outputs directory exists
if (!fs.existsSync("./outputs")) {
  fs.mkdirSync("./outputs");
}

const scripts = [
  "generate_ppts.js", // Existing MFCS generator
  "generate_ppts_sem1_ads.js",
  "generate_ppts_sem1_plsql.js",
  "generate_ppts_sem1_sqe.js",
  "generate_ppts_sem1_rmipr.js"
];

scripts.forEach((script) => {
  console.log(`Running: node ${script}...`);
  try {
    execSync(`node ${script}`, { stdio: "inherit" });
    console.log(`Finished ${script} successfully!\n`);
  } catch (error) {
    console.error(`Error running ${script}:`, error.message);
    process.exit(1);
  }
});

console.log("=========================================");
console.log("ALL SEMESTER 1 POWERPOINTS GENERATED SUCCESSFULLY!");
console.log("=========================================");
