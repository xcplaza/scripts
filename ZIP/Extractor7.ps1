# Specify the full path to the 7-Zip executable
$7ZipPath = ".\7z.exe"

# Check if 7-Zip executable exists
if (-not (Test-Path $7ZipPath)) {
    Write-Host "7-Zip executable not found. Please make sure 7-Zip is installed at '$7ZipPath' or specify the correct path."
    Exit
}

# Specify the ARJ file path
$ArjFilePath = "*.ARJ"

# Check if the ARJ file exists
if (-not (Test-Path $ArjFilePath)) {
    Write-Host "ARJ file not found at $ArjFilePath."
    Exit
}

# Specify the output directory
$OutputDirectory = ".\"

# Execute 7-Zip to extract ARJ file
try {
    & $7ZipPath x "$ArjFilePath" -o"$OutputDirectory" -aoa
    Write-Host "Extraction successful."
} catch {
    Write-Host "Error occurred while extracting ARJ archive: $_"
}
