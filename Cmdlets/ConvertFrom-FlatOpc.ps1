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

function ConvertFrom-FlatOpc {
    <#
    .SYNOPSIS
    Converts a Flat OPC file to a binary OPC file.
    .DESCRIPTION
    Converts a Flat OPC file to a binary OPC file.
    .EXAMPLE
    # Demonstrates saving a Flat OPC file (in an XmlDocument object) to a file.
    $a = ConvertTo-FlatOpc Input1.docx
    ConvertFrom-FlatOpc -OutputPath Output1.docx -FlatOpc $a
    .PARAMETER FlatOpc
    The Flat OPC file to convert to binary OPC.
    .PARAMETER OutputPath
    The filename of the binary OPC file.
    #>
    param(
        [Parameter(Mandatory=$True)]
        $FlatOpc,

        [Parameter(Mandatory=$True)]
        [string]$OutputPath
	)

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    $t = $FlatOpc.GetType()
    $ns = $t.Namespace
    $name = $t.Name
    $bt = $t.BaseType
    if ($ns -eq 'System.Xml' -and $name -eq 'XmlDocument')
    {
        [OpenXmlPowerTools.FlatOpc]::FlatToOpc($FlatOpc, $OutputPath)
        Write-Verbose "Input file is XmlDocument"
    }
    elseif ($ns -eq 'System.Xml.Linq' -and $name -eq 'XDocument')
    {
        [OpenXmlPowerTools.FlatOpc]::FlatToOpc($FlatOpc, $OutputPath)
        Write-Verbose "Input file is XDocument"
    }
    elseif ($ns -eq 'System' -and $name -eq 'Object[]' -and $bt.Name -eq "Array")
    {
        $sb = New-Object System.Text.StringBuilder
        $z = $FlatOpc | % { $sb.Append($_) }
        [OpenXmlPowerTools.FlatOpc]::FlatToOpc($sb.ToString(), $OutputPath)
        Write-Verbose "Input file is array of text"
    }
    else
    {
        $z = $($FlatOpc.GetType() | fl)
        throw "Error, type is $z"
    }

	[environment]::CurrentDirectory = $prevCurrentDirectory
}
