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

function New-Docx {
    <#
    .SYNOPSIS
    Creates a new sample DOCX.
    .DESCRIPTION
    #TBD
    .EXAMPLE
    #TBD
    .PARAMETER OutputPath
    Path and name of the sample file to create.
    .PARAMETER All
    Creates a document with every feature in it.
    .PARAMETER LoadAndSaveUsingWord
    Uses Word Automation to load and save the document after DocumentBuilder finishes building it.
    # ParameterDocumentation
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$False)]
        [string]$OutputPath,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$All,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$LoadAndSaveUsingWord,

        [Parameter(Mandatory = $false)]
        [scriptblock]
        $ScriptBlock,

        # ParameterDeclaration
    )

    Write-Verbose "Creating a sample DOCX"
    if ($OutputPath -ne [string]::Empty)
    {
        Write-Verbose "  Output document: $OutputPath"
    }
    else
    {
        Write-Verbose "  No output document, returning WmlDocument object"
    }

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    $srcList = New-Object 'System.Collections.Generic.List[OpenXmlPowerTools.Source]'

    # ParameterUse

    $randomizedSources = $srcList | % {
        @{
            Source = $_
            Rand = Get-Random
        }} | Sort-Object { $_.Rand } | % { $_.Source }

    $mergedWmlDocument = [OpenXmlPowerTools.DocumentBuilder]::BuildDocument($randomizedSources)
    if ($OutputPath -ne [string]::Empty)
    {
        $outputFi = New-Object System.IO.FileInfo $OutputPath
        if ($outputFi.Extension.ToLower() -ne '.docx')
        {
            $newFileName = $(Join-Path $outputFi.Directory ($outputFi.BaseName + ".docx"))
            $outputFi = New-Object System.IO.FileInfo $newFileName
        }
        if ($LoadAndSaveUsingWord)
        {
            $tempDocName = $(Join-Path $outputFi.DirectoryName ($outputFi.BaseName + "-Temp.docx"))
            #zzz
            $mergedWmlDocument.SaveAs($tempDocName)
            $Word = New-Object -Com Word.Application
            $Word.Visible = $false
            $wdFormatDocumentDefault=16 # http://msdn.microsoft.com/en-us/library/bb238158.aspx
            $Doc = $Word.Documents.Open($tempDocName)
            $wordFilename = $outputFi.FullName
            Try
            {
                $Doc.SaveAs($wordFilename, $wdFormatDocumentDefault)
            }
            Catch
            {
                $Doc.SaveAs([ref]$wordFilename, [ref]$wdFormatDocumentDefault)
            }
            $Doc.Close()
            $Word.Quit()
            Remove-Item $tempDocName
        }
        else
        {
            #zzz
            $mergedWmlDocument.SaveAs($outputFi.FullName)
        }
    }
    else
    {
        $mergedWmlDocument
    }

    [Environment]::CurrentDirectory = $prevCurrentDirectory
}

function AppendDoc {
    param (
        $srcList,
        $ba,
        [string]$argName
    )
    $b64b = $ba.Replace("\r\n", "")
    $ba = [System.Convert]::FromBase64String($b64b)
    $wml = (New-Object OpenXmlPowerTools.WmlDocument("DummyName.docx", $ba))
    if ($argName.Contains('Section'))
    {
        $src = (New-Object OpenXmlPowerTools.Source($wml, $true))
    }
    else
    {
        $src = (New-Object OpenXmlPowerTools.Source($wml))
    }
    $srcList.Add($src)
}
