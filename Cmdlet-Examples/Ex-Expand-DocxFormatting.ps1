[environment]::CurrentDirectory = $(Get-Location)

New-Docx WithStyles.docx -Headings -LoadAndSaveUsingWord
Expand-DocxFormatting WithStyles.docx
