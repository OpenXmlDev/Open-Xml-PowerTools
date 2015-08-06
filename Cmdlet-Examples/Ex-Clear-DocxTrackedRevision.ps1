[environment]::CurrentDirectory = $(Get-Location)

Copy-Item .\TrackedRevision.docx .\Out-NoTrackedRevisions.docx
Clear-DocxTrackedRevision Out-NoTrackedRevisions.docx


Copy-Item .\TrackedRevision.docx .\Out-NoTrackedRevisions2.docx
Clear-DocxTrackedRevision Out-NoTrackedRevisions*.docx
