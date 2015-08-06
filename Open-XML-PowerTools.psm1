<#***************************************************************************

Copyright (c) Microsoft Corporation 2014.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************#>

<#

Various advanced functions contain prototypical ways of accomplishing certain things:

####################### All Open XML

Get-OpenXmlValidationErrors
- Supports wildcards
- Supports piping a collection of files into it
- Calls a C# function that returns a collection of objects
- Creates a hash table and calls Write-Output
- ValidateSet Parameter (Office version)

Test-OpenXmlValid
- Supports wildcards
- Supports piping a collection of files into it
- Calls a C# function (that returns bool) that was introduced by Add-Type in AddTypes.ps1
- Returns a boolean value
- ValidateSet Parameter (Office version)

AddTypes.ps1
- Add and use C# types
- Link to assemblies when compiling the C#

####################### DOCX

Clear-DocxTrackedRevision
- Supports wildcards
- Supports piping a collection of files into it
- ShouldProcess=$True, ConfirmImpact='Medium'
- Switch parameter (supports -Force)
- Calls directly into OpenXmlPowerTools using WmlDocument
- Defines aliases

Merge-Docx
- Calls directly into OpenXmlPowerTools using WmlDocument
- Defines aliases
- OutputPath parameter
- Returns WmlDocument object

#>

$ver = $PSVersionTable.PSVersion
if ($ver.Major -lt 3) { throw "You must be running PowerShell 3.0 or later" }
if (Get-Module Open-XML-PowerTools) { return }

# AddTypes.ps1 is in the same directory as Open-XML-PowerTools.psm1
# needs to access both Cmdlets and OpenXmlPowerTools
. "$PSScriptRoot\AddTypes.ps1"
. "$PSScriptRoot\Cmdlets\Utils.ps1"

## Applies to any file
. "$PSScriptRoot\Cmdlets\ConvertTo-Base64.ps1"
. "$PSScriptRoot\Cmdlets\ConvertFrom-Base64.ps1"
. "$PSScriptRoot\Cmdlets\ConvertTo-FlatOpc.ps1"
. "$PSScriptRoot\Cmdlets\ConvertFrom-FlatOpc.ps1"

## Applies to all Open XML document types
. "$PSScriptRoot\Cmdlets\Get-OpenXmlValidationErrors.ps1"
. "$PSScriptRoot\Cmdlets\Test-OpenXmlValid.ps1"
. "$PSScriptRoot\Cmdlets\Test-OpenXmlPowerToolsCmdlets.ps1"

# DOCX
. "$PSScriptRoot\Cmdlets\Convert-DocxToHtml.ps1"
. "$PSScriptRoot\Cmdlets\Clear-DocxTrackedRevision.ps1"
. "$PSScriptRoot\Cmdlets\Expand-DocxFormatting.ps1"
. "$PSScriptRoot\Cmdlets\Merge-Docx.ps1"
. "$PSScriptRoot\Cmdlets\Complete-DocxTemplateFromXml.ps1"
. "$PSScriptRoot\Cmdlets\New-Docx.ps1"
. "$PSScriptRoot\Cmdlets\Add-DocxText.ps1"
. "$PSScriptRoot\Cmdlets\New-WmlDocument.ps1"
. "$PSScriptRoot\Cmdlets\DocxLib.ps1"
. "$PSScriptRoot\Cmdlets\Get-DocxMetrics.ps1"

# XLSX

# PPTX
. "$PSScriptRoot\Cmdlets\New-Pptx.ps1"
. "$PSScriptRoot\Cmdlets\Merge-Pptx.ps1"
. "$PSScriptRoot\Cmdlets\PptxLib.ps1"
. "$PSScriptRoot\Cmdlets\New-PmlDocument.ps1"

Export-ModuleMember `
    -Alias @(
        'AcceptRevision',
        'Accept-DocxTrackedRevision',
        'BuildDocument',
        'BuildPresentation',
        'AssembleFormatting'
    ) `
    -Function @(
        # All Files
        'ConvertTo-Base64',
        'ConvertFrom-Base64',
        'ConvertTo-FlatOpc',
        'ConvertFrom-FlatOpc',
        'Convert-DocxToHtml',
        'Format-Xml',

        # All Open XML
        'Get-OpenXmlValidationErrors',
        'Test-OpenXmlValid',
        'Test-OpenXmlPowerToolsCmdlets',

        # DOCX
        'Clear-DocxTrackedRevision',
        'Expand-DocxFormatting',
        'Merge-Docx',
        'Complete-DocxTemplateFromXml',
        'New-Docx',
        'Add-DocxText',
        'Get-DocxMetrics',
        'New-WmlDocument',

        # XLSX

        # PPTX
        'New-Pptx',
        'Merge-Pptx',
        'New-PmlDocument'
    )
