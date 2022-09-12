# Copy script file to the root folder containing .ppt/.pptx docs
# Batch convert all .ppt/.pptx files encountered in folder including subfolders
#
# If PowerShell exits with an error, check if unsigned scripts are allowed in your system.
# You can allow them by calling PowerShell as an Administrator and typing
# ```
# Set-ExecutionPolicy Unrestricted
# ```
# Get invocation path ($curr_path)
$curr_path = Split-Path -parent $MyInvocation.MyCommand.Path
# Create a PowerPoint object
$ppt_app = New-Object -ComObject PowerPoint.Application
# Get all objects of type .ppt? in $curr_path and its subfolders
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | ForEach-Object {
    Write-Host "Processing" $_.FullName "..."
    # Open it in PowerPoint
    $document = $ppt_app.Presentations.Open($_.FullName)
    # Create a name for the PDF document
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    # Save as PDF
    $opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
    $document.SaveAs($pdf_filename, $opt)
    # Close PowerPoint file
    $document.Close()
}
# Exit and release the PowerPoint object
$ppt_app.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt_app)
# Create a folder named "powerpoint_archive" in the $curr_path
New-Item -Path $curr_path -Name "powerpoint_archive" -ItemType "directory"
# Move all .ppt? files to the newly created "powerpoint_archive" folder
Get-ChildItem -Path $curr_path -Recurse -Filter *.ppt? | Move-Item -Destination $curr_path\powerpoint_archive
# 
# If newly created files show as a Chrome HTML Document, right click on a newly created
# document, properties > general tab > "type of file" / "open with" any Adobe software, should fix (all)
#
