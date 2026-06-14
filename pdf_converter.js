const { execSync } = require("child_process");
const path = require("path");
const fs = require("fs");

function convertToPdf(pptxPath) {
  try {
    const absolutePptxPath = path.resolve(pptxPath);
    const absolutePdfPath = absolutePptxPath.replace(/\.pptx$/i, ".pdf");
    const scriptPath = path.join(__dirname, "convert.ps1");
    
    // Run the powershell script to convert pptx to pdf
    execSync(`powershell -ExecutionPolicy Bypass -File "${scriptPath}" -pptxPath "${absolutePptxPath}" -pdfPath "${absolutePdfPath}"`, {
      stdio: "inherit"
    });
  } catch (error) {
    console.error(`Failed to convert ${pptxPath} to PDF:`, error.message);
  }
}

module.exports = { convertToPdf };
