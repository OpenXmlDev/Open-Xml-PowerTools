New-Docx Out-1 -Bookmark
New-Docx Out-2.docx -Comment
New-Docx Out-3.docx -CoverPage
New-Docx Out-All.docx -All -LoadAndSaveUsingWord
[OpenXmlPowerTools.WmlDocument]$wml = New-Docx -Bookmark
$wml.SaveAs("Out-4.docx")

