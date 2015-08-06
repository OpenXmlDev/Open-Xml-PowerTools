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

function IsWordprocessingML {
    param(
        [string]$ext
    )
    $lcext = $ext.ToLower();
    $lcext -eq ".docx" -or $lcext -eq ".docm" -or $lcext -eq ".dotx" -or $lcext -eq ".dotm"
}

function IsSpreadsheetML {
    param(
        [string]$ext
    )
    $lcext = $ext.ToLower();
    $lcext -eq ".xlsx" -or $lcext -eq ".xlsm" -or $lcext -eq ".xltx" -or $lcext -eq ".xltm" -or $lcext -eq ".xlam"
}

function IsPresentationML {
    param(
        [string]$ext
    )
    $lcext = $ext.ToLower();
    $lcext -eq ".pptx" -or $lcext -eq ".potx" -or $lcext -eq ".ppsx" -or $lcext -eq ".pptm" `
        -or $lcext -eq ".potm" -or $lcext -eq ".ppsm" -or $lcext -eq ".ppam"
}

function ReleaseComObject {
    param (
        $comObject
    )
    try
    {
        while ($true)
        {
            $a = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($comObject)
            if ($a -eq 0)
            {
                break
            }
        }
    }
    catch
    {
    }
}

function Format-XML ([xml]$Xml) 
{
    try
    {
        $stringWriter = New-Object System.IO.StringWriter 
        try
        {
            $xmlWriter = New-Object System.XMl.XmlTextWriter $stringWriter 
            $xmlWriter.Formatting = "indented" 
            $xmlWriter.Indentation = 4 
            $Xml.WriteContentTo($xmlWriter) 
            $xmlWriter.Flush() 
            $stringWriter.Flush() 
            Write-Output $stringWriter.ToString() 
        }
        finally
        {
            $xmlWriter.Dispose();
        }
    }
    finally
    {
        $stringWriter.Dispose();
    }
}