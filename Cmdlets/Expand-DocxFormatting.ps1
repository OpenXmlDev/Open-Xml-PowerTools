<#***************************************************************************

Copyright (c) Microsoft Corporation 2014.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

Version: 3.0.0

***************************************************************************#>

function Expand-DocxFormatting {
    <#
    .SYNOPSIS
    Expands all formatting from styles and document defaults into direct formatting on paragraphs and runs.
    .DESCRIPTION
    Expands all formatting from styles and document defaults into direct formatting on paragraphs and runs.
    This Cmdlet will add attributes in ignored namespaces that give auxiliary information about paragraphs
    and runs, including the actual font family for each run.
    .EXAMPLE
    # Simple use
    Expand-DocxFormatting MyFile.docx
    .EXAMPLE
    # Pipes DOCX into Expand-DocxFormatting
    Get-ChildItem *.docx | Expand-DocxFormatting
    .EXAMPLE
    # Uses wildcard
    Expand-DocxFormatting *.docx
    .PARAMETER FileName
    The document to expand formatting.
    .PARAMETER Force
    If set, suppresses confirmation.
    #>
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Medium')]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True, HelpMessage='What document would you like to expand formatting for?')]
        [ValidateScript(
        {
            $prevCurrentDirectory = [Environment]::CurrentDirectory
            [environment]::CurrentDirectory = $(Get-Location)
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_))
            {
            	[environment]::CurrentDirectory = $prevCurrentDirectory
                return $True
            }
            else
            {
                if (Test-Path $_)
                {
                	[environment]::CurrentDirectory = $prevCurrentDirectory
                    return $True
                }
                else
                {
                	[environment]::CurrentDirectory = $prevCurrentDirectory
                    Throw "$_ is not a valid filename"
                }
            }
        })]
        [SupportsWildcards()]
        [string[]]$FileName,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Force
    )
  
    begin {

        $prevCurrentDirectory = [Environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)

        write-verbose "Expanding formatting in $fileName"
    }
  
    process {
        write-verbose "Beginning process loop"
        foreach ($argItem in $fileName) {
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($argItem))
            {
                $dir = New-Object -TypeName System.IO.DirectoryInfo $(Get-Location)
                foreach ($fi in $dir.GetFiles($argItem))
                {
                    if ($force -or $pscmdlet.ShouldProcess($fi)) {
                        Expand-DocxFormatting-Helper($fi)
                    }
                }
            }
            else
            {
                if ($force -or $pscmdlet.ShouldProcess($argItem)) {
                    $fi = New-Object System.IO.FileInfo $argItem
                    Expand-DocxFormatting-Helper($fi)
                }
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Expand-DocxFormatting-Helper {
    param (
        [System.IO.FileInfo]$fi
    )
    if (IsWordprocessingML($fi.Extension))
    {
        Write-Verbose "Expanding formatting for $fi.FullName"
        $settings = New-Object OpenXmlPowerTools.FormattingAssemblerSettings
        $settings.ClearStyles = $true
        $settings.RemoveStyleNamesFromParagraphAndRunProperties = $true
        $settings.CreateHtmlConverterAnnotationAttributes = $true
        $settings.OrderElementsPerStandard = $true
        $settings.RestrictToSupportedLanguages = $false
        $settings.RestrictToSupportedNumberingFormats = $false
        $wml = new-object OpenXmlPowerTools.WmlDocument $fi.FullName
        $newWml = [OpenXmlPowerTools.FormattingAssembler]::AssembleFormatting($wml, $settings)
        $newWml.SaveAs($fi.FullName)
    }
    else
    {
        Throw "Invalid Open XML file type for expanding formatting"
    }
}

New-Alias AssembleFormatting Expand-DocxFormatting
