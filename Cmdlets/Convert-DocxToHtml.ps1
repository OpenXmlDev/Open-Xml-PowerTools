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

function Convert-DocxToHtml {
    <#
    .SYNOPSIS
    Converts a DOCX to HTML/CSS.
    .DESCRIPTION
    Converts a DOCX to HTML/CSS, outputting images to a related directory.
    .EXAMPLE
    # Demonstrates the Verbose argument
    Convert-DocxToHtml Valid.docx -Verbose
    .PARAMETER FileName
    The DOCX file to convert to HTML.
    .PARAMETER OutputPath
    The directory that will contain the converted HTML file.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='What DOCX file would you like to transform to HTML?',
        Position=0)]
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
        [ValidateScript({Test-Path $_ -PathType Container})]
        [string]$OutputPath
    )
   
    begin {
        $prevCurrentDirectory = [Environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)
    }
   
    process {
        foreach ($argItem in $FileName) {
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($argItem))
            {
                $dir = New-Object -TypeName System.IO.DirectoryInfo $(Get-Location)
                foreach ($fi in $dir.GetFiles($argItem))
                {
                    Convert-DocxToHtml-Helper $fi $OutputPath
                }
            }
            else
            {
                $fi = New-Object System.IO.FileInfo($argItem)
                Convert-DocxToHtml-Helper $fi $OutputPath
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Convert-DocxToHtml-Helper {
    param (
        [System.IO.FileInfo]$fi,
        [string]$OutputPath
    )
    [OpenXmlPowerTools.HtmlConverterHelper]::ConvertToHtml($fi.FullName, $OutputPath);
}
