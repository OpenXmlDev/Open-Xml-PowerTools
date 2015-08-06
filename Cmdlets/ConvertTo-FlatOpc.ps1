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

function ConvertTo-FlatOpc {
    <#
    .SYNOPSIS
    Converts an OPC file to Flat OPC.
    .DESCRIPTION
    Converts an OPC binary file to Flat OPC.  See http://bit.ly/1yyG5rq for more information about the Flat OPC format.
    .EXAMPLE
    # Demonstrates simple conversion to Flat OPC
    ConvertTo-FlatOpc Input1.docx
    .PARAMETER FileName
    The file to convert to Flat OPC.
    .PARAMETER OutputFormat
    Specifies the output format.  The default is XmlDocument.

    Valid values are:

    -- XmlDocument:  Returns an XmlDocument object
    -- XDocument:  Returns an XDocument object
    -- Text:  Returns an array of string objects
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='What OPC file would you like to convert to Flat OPC?',
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
        [ValidateSet('XmlDocument', 'XDocument', 'Text')]
        [string]$OutputFormat
    )
  
    begin {

        $prevCurrentDirectory = [Environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)

        if ($OutputFormat -eq [string]::Empty)
        {
            $OutputFormat = 'XmlDocument'
        }
    }
  
    process {
        foreach ($argItem in $FileName) {
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($argItem))
            {
                $dir = New-Object -TypeName System.IO.DirectoryInfo $(Get-Location)
                foreach ($f in $dir.GetFiles($argItem))
                {
                    if ($OutputFormat -eq 'Text')
                    {
                        [OpenXmlPowerTools.FlatOpc]::OpcToText($f.FullName)
                    }
                    elseif ($OutputFormat -eq 'XDocument')
                    {
                        [OpenXmlPowerTools.FlatOpc]::OpcToXDocument($f.FullName)
                    }
                    else
                    {
                        [OpenXmlPowerTools.FlatOpc]::OpcToXmlDocument($f.FullName)
                    }
                }
            }
            else
            {
                $fi = New-Object -TypeName System.IO.FileInfo $argItem
                if ($OutputFormat -eq 'Text')
                {
                    [OpenXmlPowerTools.FlatOpc]::OpcToText($fi.FullName)
                }
                elseif ($OutputFormat -eq 'XDocument')
                {
                    [OpenXmlPowerTools.FlatOpc]::OpcToXDocument($fi.FullName)
                }
                else
                {
                    [OpenXmlPowerTools.FlatOpc]::OpcToXmlDocument($fi.FullName)
                }
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}
