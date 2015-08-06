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

function Get-OpenXmlValidationErrors {
    <#
    .SYNOPSIS
    Test an Open XML document for validity and returns a list of errors.
    .DESCRIPTION
    Test a document for validity using the Open XML SDK validation functionality.  If the document is
    valid, returns an empty collection.
    .EXAMPLE
    Get-OpenXmlValidationErrors Valid.docx -Office2010 -Verbose
    Get-OpenXmlValidationErrors Invalid.docx -Office2007 -Verbose
    .EXAMPLE
    Get-ChildItem *.xlsx | Get-OpenXmlValidationErrors -Verbose
    .EXAMPLE
    Get-OpenXmlValidationErrors *.docx
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
                foreach ($f in $dir.GetFiles($argItem))
                {
                    Get-OpenXmlValidationErrors-Helper $f $ffv
                }
            }
            else
            {
                $fi = New-Object -TypeName System.IO.FileInfo $argItem
                Get-OpenXmlValidationErrors-Helper $fi $ffv
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Get-OpenXmlValidationErrors-Helper {
    param (
        [System.IO.FileInfo]$fi,
        [string]$ffv
    )
    $e = [OpenXmlPowerTools.ValidationHelper]::GetOpenXmlValidationErrors($fi.FullName, $ffv);
    $errCount = $($e | measure).Count
    if ($errCount -eq 0)
    {
        Write-Verbose "$fi is valid ($ffv)"
    }
    elseif ($errCount -eq 1)
    {
        Write-Verbose "$fi has 1 error ($ffv)"
    }
    else
    {
        Write-Verbose "$fi has $errCount errors ($ffv)"
    }
    $e | % {
        # create hashtable
        $output = @{
            'FileName' = $fi.FullName;
            'Description' = $_.Description;
            'ErrorType' = $_.ErrorType;
            'Id' = $_.Id;
            'Node' = $_.Node;
            'Part' = $_.Part;
            'XPath' = $_.Path.XPath
        }
        Write-Output (New-Object –Typename PSObject –Prop $output)
    }
}
