param (
    [string]$pptxPath,
    [string]$pdfPath
)

# Resolve paths to absolute paths
$pptxPath = (Resolve-Path $pptxPath).Path
$pdfPath = [System.IO.Path]::GetFullPath($pdfPath)

Write-Output "Converting '$pptxPath' to '$pdfPath'..."

$ppt = New-Object -ComObject PowerPoint.Application
try {
    # Open the presentation (ReadOnly = $true, Untitled = $true, WithWindow = $false)
    $presentation = $ppt.Presentations.Open($pptxPath, $true, $true, $false)
    # Save as PDF (32 is ppSaveAsPDF)
    $presentation.SaveAs($pdfPath, 32)
    $presentation.Close()
    Write-Output "Successfully converted to PDF!"
} catch {
    Write-Error "Failed to convert: $_"
    exit 1
} finally {
    # Make sure we clean up the PowerPoint application instance
    $ppt.Quit()
    # Release COM reference
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
