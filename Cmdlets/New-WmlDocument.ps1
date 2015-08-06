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

function New-WmlDocument {
    <#
    .SYNOPSIS
    Creates a WmlDocument from a file.
    .DESCRIPTION
    Creates a WmlDocument from a file, ready for further processing.
    .EXAMPLE
    # Create a new document, open it, and add content to it.
    New-Docx Test.docx -Empty
    $w = New-WmlDocument Test.docx
    Add-DocxText $w "Hello World"
    $w.SaveAs("Test2.docx")
    .PARAMETER FileName
    The file to open.
    #>
    param(
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateScript({Test-Path $_})]
        [string]$FileName
	)

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    $w = New-Object -TypeName OpenXmlPowerTools.WmlDocument $FileName

	[environment]::CurrentDirectory = $prevCurrentDirectory
    return $w
}
