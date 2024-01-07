# Create a Shell Application object
$objApp = New-Object -comobject Shell.Application

# Open a dialog for the user to select a destination folder
$dialog = $objApp.BrowseForFolder(0, "コピー先を指定", 0)

# If no folder is selected, exit the script
if ($null -eq $dialog) {
    exit
}

# Get the path of the selected folder
$folderpath = $dialog.Self.Path

# Find all shortcut files (*.lnk), get their shortcuts, and copy them to the selected folder
Get-ChildItem -Recurse -Include "*.lnk" | Get-Shortcut | Copy-Item -Destination $folderpath 

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objApp)