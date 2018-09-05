Add-PSSnapin Microsoft.SharePoint.PowerShell -ea 0

#Uploads a folder of Files to SharePoint
function UploadFilesToSharePoint([string] $siteCollectionUrl, [string] $libraryName, [string] $SourceFilesLocation){
 
    $spSourceWeb = Get-SPWeb $siteCollectionUrl
    $spSourceList = $spSourceWeb.Lists[$libraryName]
    
    if($spSourceList -eq $null)
    {
        Write-Host "The Library $libraryName could not be found."
        return;
    }
 
    $files = ([System.IO.DirectoryInfo] (Get-Item $SourceFilesLocation)).GetFiles()
    foreach($file in $files)
    {
        #Open file
        $fileStream = ([System.IO.FileInfo] (Get-Item $file.FullName)).OpenRead()
 
        #Add file
        $folder =  $spSourceWeb.getfolder($libraryName)
 
        Write-Host -ForegroundColor Yellow  "Copying file $file to $libraryName..."
 
        $spFile = $folder.Files.Add($folder.Url + "/" + $file.Name, [System.IO.Stream]$fileStream, $true)
 
        #Close file stream
        $fileStream.Close();
    }
    $spSourceWeb.dispose();
 
    Write-Host -ForegroundColor Cyan "Files have been uploaded to $libraryName."
 
}

$siteCollectionUrl = "http://portal.corp.contoso.com"
$libraryName = "SampleDocs"
$SourceFilesLocation = "C:\SourceDocs"

 
UploadFilesToSharePoint $siteCollectionUrl $libraryName $SourceFilesLocation
 
