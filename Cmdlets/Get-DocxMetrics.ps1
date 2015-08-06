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

function Get-DocxMetrics {
    <#
    .SYNOPSIS
    Gets a variety of metrics and measurements for an Open XML document.
    .DESCRIPTION
    Gets a variety of metrics and measurements for an Open XML document, including the hierarchy of content controls and
    the style hierarchy for WordprocessingML documents, table information for SpreadsheetML documents, and images for all
    varieties of documents.
    .EXAMPLE
    # Demonstrate 
    Get-DocxMetrics Test.docx | fl
    .EXAMPLE
    # Demonstrates piping files into Test-OpenXmlValid
    Get-ChildItem *.xlsx | Get-DocxMetrics
    .EXAMPLE
    # Demonstrates wildcards
    Get-DocxMetrics *.docx
    .PARAMETER FileName
    The Open XML file to retrieve metrics for.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True,
                   HelpMessage='What Open XML file would you like to get metrics for?',
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
        [string[]]$FileName
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
                    Get-OpenXmlMetrics-Helper $fi
                }
            }
            else
            {
                $fi = New-Object System.IO.FileInfo($argItem)
                Get-OpenXmlMetrics-Helper $fi
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Get-OpenXmlMetrics-Helper {
    param (
        [System.IO.FileInfo]$fi
    )
    $x = [OpenXmlPowerTools.GetMetricsHelper]::GetDocxMetrics($fi.FullName);
    $x
}
