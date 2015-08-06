[environment]::CurrentDirectory = $(Get-Location)
if (-not $(Test-Path .\GenerateNewDocxCmdlet.ps1))
{
    Throw "You must run this script from within the NewDocxDocuments directory"
}

$dx = "..\Cmdlets\DocxLib.ps1"
if (Test-Path $dx) { del $dx }

$lineBreak = [System.Environment]::NewLine

[System.Text.StringBuilder]$sbDxl = New-Object -TypeName System.Text.StringBuilder

$copyrightString = @"
<#***************************************************************************

Copyright (c) Microsoft Corporation 2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************#>

"@

[void]$sbDxl.Append($copyrightString + $lineBreak)

dir *.docx | % {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbDxl.Append("`$SampleDocx$($_.BaseName) =" + $lineBreak)
    $b64 = $(ConvertTo-Base64 $_ -PowerShellLiteral)
    [void]$sbDxl.Append($b64 + $lineBreak)
    [void]$sbDxl.Append("" + $lineBreak)
}
Set-Content -Value $sbDxl.ToString() -Path $dx -Encoding UTF8

$template = [System.IO.File]::ReadAllLines("..\Cmdlets\New-Docx-Template.ps1")
$paramDocs = -1;
$paramDecl = -1;
$paramUse = -1;
for ($i = 0; $i -lt $template.Length; $i++)
{
    $t = $template[$i]
    if ($t.Contains("ParameterDocumentation")) { $paramDocs = $i }
    if ($t.Contains("ParameterDeclaration")) { $paramDecl = $i }
    if ($t.Contains("ParameterUse")) { $paramUse = $i }
}

$ndx = "..\Cmdlets\New-Docx.ps1"
if (Test-Path $ndx)
{
    Remove-Item $ndx
}

$sbGenNewDocx = New-Object System.Text.StringBuilder;

$template[0..($paramDocs - 1)] | % { [void]$sbGenNewDocx.Append($_ + $lineBreak) }
dir *.docx | % {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbGenNewDocx.Append("    .PARAMETER $($fi.BaseName)" + $lineBreak)
    $fiDesc = New-Object System.IO.FileInfo $($_.BaseName + ".txt")
    if ($fiDesc.Exists)
    {
        Get-Content $($fiDesc.FullName) | % { [void]$sbGenNewDocx.Append('    ' + $_ + $lineBreak) }
    }
    else
    {
        $errMessage = "Error: $($fi.BaseName).docx does not have a corresponding $($fi.BaseName).txt"
        Write-Host -ForegroundColor Red $errMessage
        [void]$sbGenNewDocx.Append('    ' + $errMessage + $lineBreak)
    }
}
$start = $paramDocs + 1
$end = $paramDecl - 1
$template[$start..$end] | % { [void]$sbGenNewDocx.Append($_ + $lineBreak) }
$last = (($(dir *.docx) | measure).Count) - 1
$count = 0
dir *.docx | % {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbGenNewDocx.Append('        [Parameter(Mandatory=$False)]' + $lineBreak)
    [void]$sbGenNewDocx.Append('        [Switch]' + $lineBreak)
    if ($count -ne $last)
    {
        [void]$sbGenNewDocx.Append("        [bool]`$$($_.BaseName)," + $lineBreak)
    }
    else
    {
        [void]$sbGenNewDocx.Append("        [bool]`$$($_.BaseName)" + $lineBreak)
    }
    [void]$sbGenNewDocx.Append($lineBreak)
    $count++
}
$start = $paramDecl + 1
$end = $paramUse - 1
$template[$start..$end] | % { [void]$sbGenNewDocx.Append($_ + $lineBreak) }
dir *.docx | % {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbGenNewDocx.Append("    if (`$All -or `$$($fi.BaseName)) { AppendDoc `$srcList `$SampleDocx$($fi.BaseName) `"$($fi.BaseName)`" }" + $lineBreak)
}
$start = $paramUse + 1
$template[$start..99999] | % { [void]$sbGenNewDocx.Append($_ + $lineBreak) }

Set-Content -Value $sbGenNewDocx.ToString() -Path $ndx -Encoding UTF8
