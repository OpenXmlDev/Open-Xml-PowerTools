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

function ConvertTo-Base64 {
    <#
    .SYNOPSIS
    Converts a file to base 64 encoded ASCII.
    .DESCRIPTION
    Converts a file to base 64 encoded ASCII.  There are options for a variety of output formats.
    .EXAMPLE
    # Demonstrates simple conversion to base 64 encoded ASCII
    ConvertTo-Base64 Empty.docx
    .EXAMPLE
    # Demonstrates conversion to a Powershell literal string
    ConvertTo-Base64 Empty.docx -PowerShellLiteral
    .EXAMPLE
    # Demonstrates conversion to a C# literal string, along with the code to convert and open the document.
    ConvertTo-Base64 Empty.docx -CSharpLiteral -WithCode
    .PARAMETER FileName
    The file to convert to base 64.
    .PARAMETER JavaScriptLiteral
    Convert to a JavaScript literal.
    .PARAMETER PowerShellLiteral
    Convert to a PowerShell literal.
    .PARAMETER CSharpLiteral
    Convert to a C# literal.
    .PARAMETER VBLiteral
    Convert to a VB literal.
    .PARAMETER WithCode
    Outputs sample code that shows how to work with the base 64 encoded ASCII string.
    #>
    param(
        [Parameter(Mandatory=$True,
        HelpMessage='What file would you like to convert to base 64 encoded ASCII?')]
        [ValidateScript({Test-Path $_})]
        [string]$FileName,
        
        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$JavaScriptLiteral,
        
        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$PowerShellLiteral,
        
        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$CSharpLiteral,
        
        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$VBLiteral,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$WithCode
	)

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [environment]::CurrentDirectory = $(Get-Location)

    $ofCount = 0
    $lit = ""
    if ($JavaScriptLiteral) { $ofCount++ }
    if ($PowerShellLiteral) { $ofCount++ }
    if ($CSharpLiteral) { $ofCount++ }
    if ($VBLiteral) { $ofCount++ }
    if ($ofCount -gt 1) { Throw "Only one output format can be specified" }
    if ($JavaScriptLiteral)
    {
        $lit = ConvertToJavaScriptLiteral $FileName $WithCode
    }
    elseif ($PowerShellLiteral)
    {
        $lit = ConvertToPowerShellLiteral $FileName $WithCode
    }
    elseif ($CSharpLiteral)
    {
        $lit = ConvertToCSharpLiteral $FileName $WithCode
    }
    elseif ($VBLiteral)
    {
        $lit = ConvertToVBLiteral $FileName $WithCode
    }
    else
    {
        $lit = [OpenXmlPowerTools.Base64]::ConvertToBase64($FileName)
    }

	[environment]::CurrentDirectory = $prevCurrentDirectory
    return $lit
}


function ConvertToCSharpLiteral {
    param (
        [string]$path,
        [bool]$withCode
    )
    $f = Resolve-Path $path
    $a = [OpenXmlPowerTools.Base64]::ConvertToBase64($f)
    if ($withCode)
    {
        #'@"' + $a + '";'
@"
            string b64 =
@"$a"; 
            string b64b = b64.Replace("\r\n", "");
            byte[] ba = System.Convert.FromBase64String(b64b);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var cnt = wDoc.MainDocumentPart.Document.Descendants().Count();
                    Console.WriteLine(cnt);
                }
            }
"@
    }
    else
    {
@"
@"$a";
"@
    }
}

function ConvertToJavaScriptLiteral {
    param (
        [string]$path,
        [bool]$withCode
    )
    $f = Resolve-Path $path
    $a = ConvertTo-Base64 $f
    $b = $a -split "\r\n"
    $count = $b.Count
    if ($withCode)
    {
        "var document ="
        $c = $b | Select-Object -First $($count - 1) | % { "`"$_`" +" }
        $d = $b | Select-Object -Last 1 | % { "`"$_`"" }
        $c + $d
        "var doc = new openXml.OpenXmlPackage(document);"
        "var p = new XElement(W.p,"
        "    new XElement(W.r,"
        "        new XElement(W.t, ""Hello Open XML World"")));"
        "doc.mainDocumentPart().getXDocument().descendants(W.p).firstOrDefault().replaceWith(p);"
        "var modifiedDoc = doc.saveToBase64();"
    }
    else
    {
        $c = $b | Select-Object -First $($count - 1) | % { "`"$_`" +" }
        $d = $b | Select-Object -Last 1 | % { "`"$_`"" }
        $c + $d
    }
}

function ConvertToPowerShellLiteral {
    param (
        $path,
        $withCode
    )
    $f = Resolve-Path $path
    $a = ConvertTo-Base64 $f
    if ($withCode)
    {
        $sbConvertToLiteral = New-Object -TypeName "System.Text.StringBuilder";
        [void]$sbConvertToLiteral.Append('$b64 = @"' + [System.Environment]::NewLine)
        [void]$sbConvertToLiteral.Append('@"' + [System.Environment]::NewLine)
        [void]$sbConvertToLiteral.Append($a + [System.Environment]::NewLine);
        [void]$sbConvertToLiteral.Append('"@' + [System.Environment]::NewLine)
        [void]$sbConvertToLiteral.Append('ConvertFrom-Base64 "Test.docx" $b64' + [System.Environment]::NewLine)
        return $sbConvertToLiteral.ToString()
    }
    else
    {
        $sbConvertToLiteral = New-Object -TypeName "System.Text.StringBuilder";
        [void]$sbConvertToLiteral.Append('@"' + [System.Environment]::NewLine)
        [void]$sbConvertToLiteral.Append($a + [System.Environment]::NewLine);
        [void]$sbConvertToLiteral.Append('"@' + [System.Environment]::NewLine)
        return $sbConvertToLiteral.ToString()
    }
}

function ConvertToVBLiteral {
    param (
        $path,
        $withCode
    )
    $f = Resolve-Path $path
    $a = ConvertTo-Base64 $f
    $b = $a -split "\r\n"
    $count = $b.Count
    if ($withCode)
    {
        "Dim b64 = _"
        $c = $b | Select-Object -First $($count - 1) | % { "`"$_`" `& vbCrLf `& _" }
        $d = $b | Select-Object -Last 1 | % { "`"$_`"" }
        $c + $d

        "Dim b64b = b64.Replace(""\r\n"", """")"
        "Dim ba = System.Convert.FromBase64String(b64b)"
        "Using ms As New MemoryStream()"
        "    ms.Write(ba, 0, ba.Length)"
        "    Dim wDoc = WordprocessingDocument.Open(ms, False)"
        "    Dim cnt = wDoc.MainDocumentPart.Document.Descendants().Count()"
        "    Console.WriteLine(cnt)"
        "End Using"
    }
    else
    {
        $c = $b | Select-Object -First $($count - 1) | % { "`"$_`" `& vbCrLf `& _" }
        $d = $b | Select-Object -Last 1 | % { "`"$_`"" }
        $c + $d
    }
}
