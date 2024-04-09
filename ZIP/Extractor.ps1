# Specify the full path to the UnRAR executable
$UnrarPath = ".\Arj64.exe"

# Check if UnRAR executable exists
if (-not (Test-Path $UnrarPath)) {
    Write-Host "UnRAR executable not found. Please make sure UnRAR is installed at '$UnrarPath' or specify the correct path."
    Exit
}

# Specify the ARJ file path
$ArjFilePath = "*.ARJ"

# Check if the ARJ file exists
if (-not (Test-Path $ArjFilePath)) {
    Write-Host "ARJ file not found at $ArjFilePath."
    Exit
}

# Execute UnRAR to extract ARJ file
try {
    & $UnrarPath x "$ArjFilePath" -y
    Write-Host "Extraction successful."
} catch {
    Write-Host "Error occurred while extracting ARJ archive: $_"
}
