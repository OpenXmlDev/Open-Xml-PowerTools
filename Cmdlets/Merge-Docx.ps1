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

function Merge-Docx {
    <#
    .SYNOPSIS
    Merge multiple source documents into a new document.
    .DESCRIPTION
    Create a new document that contains specified contents of multiple source documents. All or part
    of the source documents may be merged into the final document. You detail which parts you want
    from multiple source documents using the Sources parameter.
    
    There are some parts of a document that cannot be merged, such as settings. These will all come
    from the first source document. You can specify a 0 Count for the first source document to get
    those settings without including any content.
    .EXAMPLE
    # Takes all of both documents, using the section properties of the second document.
    $doc1 = New-Object OpenXmlPowerTools.Source("Input1.docx")
    $doc2 = New-Object OpenXmlPowerTools.Source("Input2.docx", $True)
    $sources = ($doc1, $doc2)
    Merge-Docx -OutputPath merge01.docx -Sources $sources
    .EXAMPLE
    # Takes the first two paragraphs of both documents.
    $doc1 = New-Object OpenXmlPowerTools.Source("Input1.docx", 1, 2, $false)
    $doc2 = New-Object OpenXmlPowerTools.Source("Input2.docx", 1, 2, $false)
    $sources = ($doc1, $doc2)
    Merge-Docx -OutputPath merge02.docx -Sources $sources
    .EXAMPLE
    # Takes all of the first document and the first two paragraphs of the second document.
    $doc1 = New-Object OpenXmlPowerTools.Source("Input1.docx")
    $doc2 = New-Object OpenXmlPowerTools.Source("Input2.docx", 1, 2, $false)
    $newWml = Merge-Docx -Sources ($doc1, $doc2) -Verbose
    $newWml.SaveAs("out-merge03.docx");
    .PARAMETER Sources
    This parameter is an array of one or more OpenXmlPowerTools.Source objects that define each source
    document and range of elements (in most cases paragraphs) that will be copied from that source
    document into the merged document. The Source object may specify an entire document or part of
    the document.

    The Start argument of the Source object specifies by index the first child element of the w:body
    element, and the Count argument specifies the number of child elements of the w:body element.  Note
    that w:bookmarkStart, w:bookmarkEnd, and other elements may be children of the w:body element.
    You need to include them in the count when creating OpenXmlPowerTools.Source objects.

    Start is a zero-based index.
    .PARAMETER OutputPath
    Path and name of the file to create with the new content.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True)]
        [OpenXmlPowerTools.Source[]]$Sources,

        [Parameter(Mandatory=$False)]
        [string]$OutputPath
    )

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    Write-Verbose "Merging DOCX documents"
    if ($OutputPath -ne [string]::Empty)
    {
        Write-Verbose "  Output document: $OutputPath"
    }
    else
    {
        Write-Verbose "  No output document, returning WmlDocument object"
    }
    Write-Verbose "  Number of sources: $($Sources.Length)"
    Write-Verbose ""
    for($i = 0; $i -lt $Sources.Length; $i++)
    {
        $s = $Sources[$i]
        Write-Verbose "  Source $($i + 1)"
        Write-Verbose "  File Name: $($s.WmlDocument.FileName)"
        Write-Verbose "  Start: $($s.Start)"
        if ($s.Count -eq 2147483647)
        {
            Write-Verbose "  Count: end of document"
        }
        else
        {
            Write-Verbose "  Count: $($s.Count)"
        }
        Write-Verbose "  Keep Sections: $($s.KeepSections)"
        Write-Verbose ""
    }

    $mergedWmlDocument = [OpenXmlPowerTools.DocumentBuilder]::BuildDocument($Sources)
    if ($OutputPath -ne [string]::Empty)
    {
        $mergedWmlDocument.SaveAs($OutputPath)
    }
    else
    {
        $mergedWmlDocument
    }

	[environment]::CurrentDirectory = $prevCurrentDirectory
}

New-Alias BuildDocument Merge-Docx
