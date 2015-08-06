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

function ConvertFrom-Base64 {
    <#
    .SYNOPSIS
    Saves a base 64 encoded ASCII string to a binary file.
    .DESCRIPTION
    Saves a base 64 encoded ASCII string to a binary file.
    .EXAMPLE
    # Demonstrates saving a base 64 encoded ASCII string to a file.
    ConvertTo-Base64 Empty.docx
    .EXAMPLE
    # Demonstrates conversion to a Powershell literal string
    $b64 =
    @"
    //5OAG8AdwAgAGkAcwAgAHQAaABlACAAdABpAG0AZQAgAGYAbwByACAAYQBsAGwAIABnAG8AbwBk
    ACAAbQBlAG4AIAB0AG8AIABjAG8AbQBlACAAdABvACAAdABoAGUAIABhAGkAZAAgAG8AZgAgAHQA
    aABlAGkAcgAgAGMAbwB1AG4AdAByAHkALgANAAoA
    "@
    ConvertFrom-Base64 "Hello.txt" $b64
    .PARAMETER FileName
    The file to convert to base 64.
    .PARAMETER Base64EncodedString
    The base 64 encoded string to save.
    #>
    param(
        [Parameter(Mandatory=$True,
        HelpMessage='What file would you like to save the base 64 encoded ASCII to?')]
        [string]$FileName,
        [Parameter(Mandatory=$True)]
        [string]$Base64EncodedString
	)

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    $b64b = $Base64EncodedString.Replace("\r\n", "")
    $ba = [System.Convert]::FromBase64String($b64b)
    [System.IO.File]::WriteAllBytes($FileName, $ba)

	[environment]::CurrentDirectory = $prevCurrentDirectory
}
