
function addWelcomeSlide($ppt, $title, $subtitle)
{
$slideStyle = 1 #Title Slide Style
$slide = $ppt.Slides.Add($ppt.Slides.Count+1, $slideStyle)
$slide.Shapes.Title.TextFrame.TextRange = $title
$slide.Shapes | ?{$_.Name -eq 'Subtitle 2'} | %{$_.TextFrame.TextRange = $subtitle}
}
function addBulletSlide($ppt, $title, $body)
{
$slideStyle = 2 #Bullet Slide Style
$slide = $ppt.Slides.Add($ppt.Slides.Count+1, $slideStyle)
$slide.Shapes.Title.TextFrame.TextRange = $title
$slide.Shapes | ?{$_.Name -eq 'Text Placeholder 2'} | %{$_.TextFrame.TextRange = $body}
}

#Create PowerPoint
$app = New-Object -ComObject PowerPoint.Application -strict -property @{visible=$true}
$ppt = $app.Presentations.Add($true)
$themePath = "c:\temp\circuit.thmx"
$ppt.ApplyTemplate($themePath)


#Add Slides
addWelcomeSlide $ppt "Welcome to my Presentation" "Chris Conte"
addBulletSlide $ppt "Fruits" "Apple`nBanana`nPear`nOrange"
addBulletSlide $ppt "Nuts" "Walnut`nPeanut`nCashew`nGrape Nut"


$ppt.SaveAs("C:\temp\Sneeze.pptx",11)
$ppt.Close()
$app.quit()
$app = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers() 