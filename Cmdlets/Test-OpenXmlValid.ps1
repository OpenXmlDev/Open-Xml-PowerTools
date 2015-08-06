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

function Test-OpenXmlValid {
    <#
    .SYNOPSIS
    Test an Open XML document for validity.
    .DESCRIPTION
    Test a document for validity using the Open XML SDK validation functionality.  Returns $true
    if the document is valid; $false otherwise.
    .EXAMPLE
    # Demonstrates the Verbose argument
    Test-OpenXmlValid Valid.docx -Verbose
    Test-OpenXmlValid Invalid.docx -Verbose
    .EXAMPLE
    # Demonstrates piping files into Test-OpenXmlValid
    Get-ChildItem *.xlsx | Test-OpenXmlValid
    .EXAMPLE
    # Demonstrates wildcards
    Test-OpenXmlValid *.docx -OfficeVersion 2010
    .PARAMETER FileName
    The Open XML file to validate.
    .PARAMETER OfficeVersion
    Specifies the version of Office to validate against.  The default is 2013.

    Valid values are:

    -- 2007
    -- 2010
    -- 2013
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='What Open XML file would you like to validate?',
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
        [ValidateSet('2007', '2010', '2013')]
        [string]$OfficeVersion
    )
   
    begin {
        $prevCurrentDirectory = [Environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)
        if ($OfficeVersion -eq [string]::Empty)
        {
            $OfficeVersion = '2013'
        }
        $ffv = "Office" + $OfficeVersion;
    }
   
    process {
        foreach ($argItem in $FileName) {
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($argItem))
            {
                $dir = New-Object -TypeName System.IO.DirectoryInfo $(Get-Location)
                foreach ($fi in $dir.GetFiles($argItem))
                {
                    Test-OpenXmlValid-Helper $fi $ffv
                }
            }
            else
            {
                $fi = New-Object System.IO.FileInfo($argItem)
                Test-OpenXmlValid-Helper $fi $ffv
            }
        }
    }
	
    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Test-OpenXmlValid-Helper {
    param (
        [System.IO.FileInfo]$fi,
        [string]$ffv
    )
    $b = [OpenXmlPowerTools.ValidationHelper]::IsValid($fi.FullName, $ffv);
    if ($b)
    {
        Write-Verbose "$fi is valid ($ffv)"
    }
    else
    {
        Write-Verbose "$fi is invalid ($ffv)"
    }
    $b
}
