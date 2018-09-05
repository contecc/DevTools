

function createWordDoc ([string] $Title, [string] $Keywords)
{
$filePath = "C:\temp\" + $Title + ".docx"

[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
$word=new-object -ComObject "Word.Application"
$doc=$word.documents.Add()
$word.Visible=$True

$selection=$word.Selection

$selection.Font.Name="Segoe UI Light"
$selection.Font.Size=36
$selection.TypeText(($Title))

$selection.Font.Name="Times New Roman"
$selection.Font.Size=12



#Get the Extract to fill the document
$request = "https://en.wikipedia.org/w/api.php?format=json&action=query&prop=extracts&exintro&explaintext&redirects=1&titles=" + $Title
$Description = Invoke-WebRequest $request 



$paragraph = "`n`n This is a document concerning " + $Title
$paragraph += "`n " + $Description
$paragraph += "`n" + $Keywords


$selection.TypeText(($paragraph))

$doc.saveas([ref] $filePath, [ref]$saveFormat::wdFormatDocument)
}


createWordDoc "Pleading" "legal, attorney, litigation, civil liberties"
createWordDoc "procedural law" "legal, attorney, litigation, civil liberties"
createWordDoc "administrative law" "legal, attorney, litigation, civil liberties"
createWordDoc "cause of action" "legal, attorney, litigation, civil liberties"
createWordDoc "civil law" "legal, attorney, litigation, civil liberties"
createWordDoc "compensatory damages" "legal, attorney, litigation, civil liberties"
createWordDoc "constitutional law" "legal, attorney, litigation, civil liberties"
createWordDoc "demurrer" "legal, attorney, litigation, civil liberties"
createWordDoc "depose" "legal, attorney, litigation, civil liberties"
createWordDoc "misdemeanor" "legal, attorney, litigation, civil liberties"
createWordDoc "malfeasance" "legal, attorney, litigation, civil liberties"
createWordDoc "provisional remedy" "legal, attorney, litigation, civil liberties"
createWordDoc "title abstract" "legal, attorney, litigation, civil liberties"
createWordDoc "title search" "legal, attorney, litigation, civil liberties"
createWordDoc "tort" "legal, attorney, litigation, civil liberties"
createWordDoc "wobbler" "legal, attorney, litigation, civil liberties"
