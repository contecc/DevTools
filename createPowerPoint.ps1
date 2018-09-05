# Add the PowerPoint assemblies that we'll need
Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue

# Start PowerPoint
$ppt = new-object -com powerpoint.application
$ppt.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Set the locations where to find the PowerPoint files, and where to store the thumbnails
$pptPath = "C:\SPC2012\"
$jpgPath = "C:\SPC2012\Thumbnails\"
New-Item $jpgPath -ItemType directory -ea SilentlyContinue #create if it doesn't exist

# Loop through each PowerPoint File
Foreach($iFile in $(ls $pptPath -Filter "*.pptx")){
    $filename = Split-Path $iFile -leaf
    $file = $filename.Split(".")[0]
    $oFile = $pptPath + $file

    # Open the PowerPoint file
    $pres = $ppt.Presentations.Open($pptPath + $iFile)
    # Export the entire file to JPG images of size 591x333 pixels (pick your size here!)
    $pres.Export($oFile, "JPG", 591, 333);
    $pres.Close();

    #get filename for Slide2.jpg - rename it to {PowerPoint filename}.jpg
    $slideFile = $oFile + "\Slide2.jpg"
    $slideNewFile = $jpgPath + $file + ".jpg"
    Move-Item $slideFile $slideNewFile -Force # Move the slide to the output path
    Remove-Item $oFile -Recurse -Confirm:$false # Delete the Temporary export of all Slides
}

#Clean Up
$ppt.quit();
$ppt = $null
[gc]::Collect();
[gc]::WaitForPendingFinalizers();