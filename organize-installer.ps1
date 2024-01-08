.SYNOPSIS
    This script organizes files into folders based on product name and version.

.DESCRIPTION
    This script scans files in the current folder, extracts product name and version information,
    and organizes them into folders named "ProductName_ProductVersion".

.NOTES
    File Name      : organize-installer.ps1
    Author         : Seele Speicher (sp3icher@gmail.com)
    Prerequisite   : PowerShell V5
#>
# Get the current folder
$currentFolder = Get-Location

# Get all files in the current folder
$files = Get-ChildItem -File

# Iterate through each file
foreach ($file in $files) {
    # Get information about the file
    $versionInfo = $file.VersionInfo

    # Check if the file has product name and product version
    if ($versionInfo.ProductName -and $versionInfo.ProductVersion) {
        # Display product name and version, removing extra spaces
        $productName = $versionInfo.ProductName.Trim()
        $productVersion = $versionInfo.ProductVersion.Trim()

        Write-Host "File: $($file.Name)"
        Write-Host "Product Name: $productName"
        Write-Host "Product Version: $productVersion"

        # Replace problematic characters with underscores
        $invalidCharacters = [System.IO.Path]::GetInvalidFileNameChars() -join ''
        $folderName = ($productName -replace "[$invalidCharacters]", "_") + "_$($productVersion)"

        # Create folder only if the name is not empty
        if (-not [string]::IsNullOrWhiteSpace($folderName)) {
            $folderPath = Join-Path $currentFolder $folderName

            # Check if the folder already exists
            if (-not (Test-Path $folderPath -PathType Container)) {
                # Create the folder
                New-Item -ItemType Directory -Path $folderPath

                Write-Host "Folder created: $folderPath"
            }

            # Move the file to the new folder
            Move-Item -Path $file.FullName -Destination $folderPath
            Write-Host "File moved to: $folderPath"
        } else {
            Write-Host "Invalid folder name. Skipping file."
        }
    }
}
