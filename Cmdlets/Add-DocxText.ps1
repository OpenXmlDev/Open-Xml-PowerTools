<#***************************************************************************

Copyright (c) Microsoft Corporation 2014.
 
This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:
 
http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx
 
Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************#>

function Add-DocxText {

    <#
    .SYNOPSIS
    Appends given text to a specified DOCX document.
    .DESCRIPTION
    Add-DocxText cmdlet appends given text to a specified DOCX document.  Supports adding text with
    bold, italic, underline, forecolor, backcolor, paragraph style.
    .EXAMPLE
    # Simple use
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor"
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with variety of styles / colors
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Title of Document" -Style Title
    Add-DocxText $w "Heading" -Style Heading1
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor"
    Add-DocxText $w "Bold Text" -Bold
    Add-DocxText $w "Italic Text" -Italic
    Add-DocxText $w "Heading2" -Style Heading2
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor"
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with forecolor
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -ForeColor Red
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with backcolor
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -BackColor Green
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with bold, italic, underline
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -Bold -Italic -Underline
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with forecolor, backcolor
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -ForeColor White -BackColor Red
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with paragraph style
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -Style Title
    $w.Save()
    .EXAMPLE
    # Demonstrates adding text with bold, italic, underline, forcolor, backcolor, style
    $fn = "Doc1.docx"
    New-Docx $fn -Empty
    $fn = Resolve-Path $fn
    $w = New-WmlDocument $fn
    Add-DocxText $w "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor" -Bold -Italic -Underline -ForeColor White -BackColor Red -Style Heading1
    $w.Save()
    #>     
  
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Medium')]
    param
    (
        [Parameter(Mandatory=$True, Position=0)]
        [OpenXmlPowerTools.WmlDocument]$WmlDocument,

        [Parameter(Mandatory=$True, Position=1)]
        [string]$Content,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Bold,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Italic,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Underline,

        [Parameter(Mandatory=$False)]
        [string]$ForeColor,

        [Parameter(Mandatory=$False)]
        [string]$BackColor,

        [Parameter(Mandatory=$False)]
        [string]$Style
    )

    write-verbose "Appending Text to a document"     
  
    if ($WmlDocument -ne $null)
    {
        $newWmlDoc = [OpenXmlPowerTools.AddDocxTextHelper]::AppendParagraphToDocument(
			$WmlDocument, 
			$Content,
			$Bold,
			$Italic,
			$Underline,
			$ForeColor,
			$BackColor,
			$Style)
        $WmlDocument.DocumentByteArray = $newWmlDoc.DocumentByteArray
    }
}
