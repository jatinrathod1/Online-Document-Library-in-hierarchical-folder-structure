# Define SharePoint site and document library details
$siteUrl = "https://futurrizoninterns.sharepoint.com/sites/MentalHealthCareWebApplication1"
$libraryName = "CustomDocumentLibrary"
$libraryDisplayName = "Custom Document Library"
$localPath = "E:\Work FT\Hierarchical_Files_Library_5355_TEST"  

# Connect to SharePoint Online (Interactive login)
Connect-PnPOnline -URL $siteUrl -UseWebLogin

# Check if the document library exists
$library = Get-PnPList -Identity $libraryName -ErrorAction SilentlyContinue
if (-not $library) {
    Write-Host "Creating document library: $libraryDisplayName..."
    New-PnPList -Title $libraryDisplayName -Url $libraryName -Template DocumentLibrary -OnQuickLaunch
} else {
    Write-Host "Document library '$libraryDisplayName' already exists."
}

# Add custom columns if they don't already exist
Write-Host "Adding custom columns..."
$existingFields = Get-PnPField -List $libraryName

if ($existingFields.InternalName -notcontains "FolderCount") {
    Add-PnPField -List $libraryName -InternalName "FolderCount" -DisplayName "FolderCount" -Type Number -AddToDefaultView -Required:$false
}
if ($existingFields.InternalName -notcontains "PageCount") {
    Add-PnPField -List $libraryName -InternalName "PageCount" -DisplayName "PageCount" -Type Number -AddToDefaultView -Required:$false
}

# Rename 'Title' column to 'Name'
Write-Host "Renaming 'Title' column to 'Name'..."
Set-PnPField -List $libraryName -Identity "Title" -Values @{Title="Name"}

# Function to create folder structure in SharePoint Document Library
function Create-SharePointFolders {
    param (
        [string]$folderPath
    )

    # Get all subfolders recursively
    $folders = Get-ChildItem -Path $folderPath -Directory -Recurse

    foreach ($folder in $folders) {
        # Get relative path
        $relativePath = $folder.FullName.Replace($localPath, "").TrimStart("\")
        $relativePath = $relativePath -replace "\\", "/"

        Write-Host "Creating folder in SharePoint: $relativePath"

        # Split the relative path into folder levels
        $folderLevels = $relativePath -split "/"
        $currentPath = ""

        foreach ($level in $folderLevels) {
            $parentFolder = $currentPath
            $currentPath = if ($currentPath -eq "") { $level } else { "$currentPath/$level" }

            # Check if folder exists before creating
            $existingFolder = Get-PnPFolder -Url "$libraryName/$currentPath" -ErrorAction SilentlyContinue
            if (-not $existingFolder) {
                Write-Host "Creating folder: $currentPath"

                if ($parentFolder -eq "") {
                    # Create top-level folder inside the document library
                    Add-PnPFolder -Name $level -Folder $libraryName
                } else {
                    # Create subfolder inside the correct parent folder
                    Add-PnPFolder -Name $level -Folder "$libraryName/$parentFolder"
                }
            }
        }

        # Upload PDF files in the current folder to SharePoint
        $pdfFiles = Get-ChildItem -Path $folder.FullName -Filter *.pdf
        foreach ($pdfFile in $pdfFiles) {
            $sharePointPath = "$libraryName/$relativePath/$($pdfFile.Name)"
            Write-Host "Uploading file: $($pdfFile.Name) to $sharePointPath"

            # Upload the PDF file to the corresponding SharePoint folder
            Add-PnPFile -Path $pdfFile.FullName -Folder "$libraryName/$relativePath"
        }
    }
}

# Call the function to upload folder structure and PDF files
Write-Host "Creating folder structure and uploading PDF files to SharePoint..."
Create-SharePointFolders -folderPath $localPath

Write-Host "Process completed successfully!"    