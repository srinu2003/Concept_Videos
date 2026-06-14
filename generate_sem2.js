const { execSync } = require("child_process");
const fs = require("fs");

console.log("=========================================");
console.log("STARTING SEMESTER 2 POWERPOINT GENERATION");
console.log("=========================================\n");

// Ensure outputs directory exists
if (!fs.existsSync("./outputs")) {
  fs.mkdirSync("./outputs");
}

const scripts = [
  "generate_ppts_sem2_aa.js",
  "generate_ppts_sem2_aca.js",
  "generate_ppts_sem2_cs.js",
  "generate_ppts_sem2_rpa.js"
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
console.log("ALL SEMESTER 2 POWERPOINTS GENERATED SUCCESSFULLY!");
console.log("=========================================");
