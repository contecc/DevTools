
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


$paragraph = "`n`n This is a document concerning " + $Title
$paragraph += "`nLorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna. Nunc viverra imperdiet enim."
$paragraph += "`nFusce est. Vivamus a tellus. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede."
$paragraph += "`nMauris et orci. Aenean nec lorem. In porttitor. Donec laoreet nonummy augue."
$paragraph += "`n" + $Keywords


$selection.TypeText(($paragraph))

$doc.saveas([ref] $filePath, [ref]$saveFormat::wdFormatDocument)
}


createWordDoc "Civil Liberties" "legal, attorney, litigation, civil liberties"
