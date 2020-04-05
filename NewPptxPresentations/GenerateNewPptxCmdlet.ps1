[environment]::CurrentDirectory = $(Get-Location)
if (-not $(Test-Path .\GenerateNewPptxCmdlet.ps1)) {
    Throw "You must run this script from within the NewPptxPresentations directory"
}

$dx = "..\Cmdlets\PptxLib.ps1"
if (Test-Path $dx) { Remove-Item $dx }

$lineBreak = [System.Environment]::NewLine

[System.Text.StringBuilder]$sbDxl = New-Object -TypeName System.Text.StringBuilder

$copyrightString = @"
<#Copyright (c) Microsoft. All rights reserved.
Licensed under the MIT license. See LICENSE file in the project root for full license information.#>

"@

[void]$sbDxl.Append($copyrightString + $lineBreak)

Get-ChildItem *.pptx | ForEach-Object {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbDxl.Append("`$SamplePptx$($_.BaseName) =" + $lineBreak)
    $b64 = $(ConvertTo-Base64 $_ -PowerShellLiteral)
    [void]$sbDxl.Append($b64 + $lineBreak)
    [void]$sbDxl.Append("" + $lineBreak)
}
Set-Content -Value $sbDxl.ToString() -Path $dx -Encoding UTF8

$template = [System.IO.File]::ReadAllLines("..\Cmdlets\New-Pptx-Template.ps1")
$paramDocs = -1;
$paramDecl = -1;
$paramUse = -1;
for ($i = 0; $i -lt $template.Length; $i++) {
    $t = $template[$i]
    if ($t.Contains("ParameterDocumentation")) { $paramDocs = $i }
    if ($t.Contains("ParameterDeclaration")) { $paramDecl = $i }
    if ($t.Contains("ParameterUse")) { $paramUse = $i }
}

$npx = "..\Cmdlets\New-Pptx.ps1"
if (Test-Path $npx) {
    Remove-Item $npx
}

$sbGenNewPptx = New-Object System.Text.StringBuilder;

$template[0..($paramDocs - 1)] | ForEach-Object { [void]$sbGenNewPptx.Append($_ + $lineBreak) }
Get-ChildItem *.pptx | ForEach-Object {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbGenNewPptx.Append("    .PARAMETER $($fi.BaseName)" + $lineBreak)
    $fiDesc = New-Object System.IO.FileInfo $($_.BaseName + ".txt")
    if ($fiDesc.Exists) {
        Get-Content $($fiDesc.FullName) | ForEach-Object { [void]$sbGenNewPptx.Append('    ' + $_ + $lineBreak) }
    }
    else {
        $errMessage = "Error: $($fi.BaseName).pptx does not have a corresponding $($fi.BaseName).txt"
        Write-Error $errMessage
        [void]$sbGenNewPptx.Append('    ' + $errMessage + $lineBreak)
    }
}
$start = $paramDocs + 1
$end = $paramDecl - 1
$template[$start..$end] | ForEach-Object { [void]$sbGenNewPptx.Append($_ + $lineBreak) }
$last = (($(Get-ChildItem *.pptx) | Measure-Object).Count) - 1
$count = 0
Get-ChildItem *.pptx | ForEach-Object {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbGenNewPptx.Append('        [Parameter(Mandatory=$False)]' + $lineBreak)
    [void]$sbGenNewPptx.Append('        [Switch]' + $lineBreak)
    if ($count -ne $last) {
        [void]$sbGenNewPptx.Append("        [bool]`$$($_.BaseName)," + $lineBreak)
    }
    else {
        [void]$sbGenNewPptx.Append("        [bool]`$$($_.BaseName)" + $lineBreak)
    }
    [void]$sbGenNewPptx.Append($lineBreak)
    $count++
}
$start = $paramDecl + 1
$end = $paramUse - 1
$template[$start..$end] | ForEach-Object { [void]$sbGenNewPptx.Append($_ + $lineBreak) }
Get-ChildItem *.pptx | ForEach-Object {
    $fi = New-Object System.IO.FileInfo $_
    [void]$sbGenNewPptx.Append("    if (`$All -or `$$($fi.BaseName)) { AppendPresentation `$srcList `$SamplePptx$($fi.BaseName) `"$($fi.BaseName)`" }" + $lineBreak)
}
$start = $paramUse + 1
$template[$start..99999] | ForEach-Object { [void]$sbGenNewPptx.Append($_ + $lineBreak) }

Set-Content -Value $sbGenNewPptx.ToString() -Path $npx -Encoding UTF8
