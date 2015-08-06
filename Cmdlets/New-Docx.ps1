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
    .PARAMETER Bookmark
    Bookmark and hyperlink to the bookmark.
    .PARAMETER BulletedText
    Bulleted text.
    .PARAMETER Chart
    Chart.
    .PARAMETER Comment
    Paragraph that contains a comment.
    .PARAMETER ContentControls
    Content controls of various types.
    .PARAMETER ContentControlsNested
    Nested Content Controls
    .PARAMETER CoverPage
    Cover page.
    .PARAMETER EmbeddedWorkbook
    Embedded Excel workbook.
    .PARAMETER Empty
    Single blank paragraph.
    .PARAMETER EndNote
    Endnote.
    .PARAMETER Equation
    Math equation.
    .PARAMETER Fields
    Date-time field
    .PARAMETER Fonts
    Various fonts.
    .PARAMETER FootNote
    Footnote.
    .PARAMETER FormattedText
    Two paragraphs with text formatting.
    .PARAMETER Headings
    Headings
    .PARAMETER HierarchicalNumberedList
    Hierarchical numbered list.
    .PARAMETER HorizontalWhiteSpace
    First line indent, hanging indent, indented paragraphs.
    .PARAMETER Hyperlink
    Bookmark and hyperlink to the bookmark.
    .PARAMETER Image
    Image.
    .PARAMETER Justified
    Justified text.
    .PARAMETER NumberedList
    Simple numbered list.
    .PARAMETER NumberedListRomanNumerals
    List with Roman Numerals
    .PARAMETER ParagraphBorder
    Paragraph border.
    .PARAMETER Plain
    Plain text.
    .PARAMETER RevisionTracking
    Revision tracking.
    .PARAMETER RightJustified
    Right justified text.
    .PARAMETER Section
    Landscape section.
    .PARAMETER SectionWithWatermark
    Section with watermark.
    .PARAMETER Shading
    Shaded paragraph.
    .PARAMETER Shape
    Shapes.
    .PARAMETER SmartArt
    SmartArt.
    .PARAMETER Symbols
    Symbols.
    .PARAMETER Table
    Table that has a table style applied.
    .PARAMETER TableOfContents
    Table of contents.
    .PARAMETER TextBox
    Text box.
    .PARAMETER TextEffects
    Text effects.
    .PARAMETER Theme
    Theme.
    .PARAMETER VerticalWhiteSpace
    Vertical space before and after paragraphs.
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

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Bookmark,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$BulletedText,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Chart,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Comment,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$ContentControls,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$ContentControlsNested,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$CoverPage,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$EmbeddedWorkbook,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Empty,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$EndNote,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Equation,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Fields,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Fonts,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$FootNote,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$FormattedText,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Headings,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$HierarchicalNumberedList,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$HorizontalWhiteSpace,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Hyperlink,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Image,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Justified,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$NumberedList,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$NumberedListRomanNumerals,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$ParagraphBorder,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Plain,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$RevisionTracking,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$RightJustified,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Section,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$SectionWithWatermark,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Shading,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Shape,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$SmartArt,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Symbols,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Table,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TableOfContents,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TextBox,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TextEffects,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Theme,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$VerticalWhiteSpace

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

    if ($All -or $Bookmark) { AppendDoc $srcList $SampleDocxBookmark "Bookmark" }
    if ($All -or $BulletedText) { AppendDoc $srcList $SampleDocxBulletedText "BulletedText" }
    if ($All -or $Chart) { AppendDoc $srcList $SampleDocxChart "Chart" }
    if ($All -or $Comment) { AppendDoc $srcList $SampleDocxComment "Comment" }
    if ($All -or $ContentControls) { AppendDoc $srcList $SampleDocxContentControls "ContentControls" }
    if ($All -or $ContentControlsNested) { AppendDoc $srcList $SampleDocxContentControlsNested "ContentControlsNested" }
    if ($All -or $CoverPage) { AppendDoc $srcList $SampleDocxCoverPage "CoverPage" }
    if ($All -or $EmbeddedWorkbook) { AppendDoc $srcList $SampleDocxEmbeddedWorkbook "EmbeddedWorkbook" }
    if ($All -or $Empty) { AppendDoc $srcList $SampleDocxEmpty "Empty" }
    if ($All -or $EndNote) { AppendDoc $srcList $SampleDocxEndNote "EndNote" }
    if ($All -or $Equation) { AppendDoc $srcList $SampleDocxEquation "Equation" }
    if ($All -or $Fields) { AppendDoc $srcList $SampleDocxFields "Fields" }
    if ($All -or $Fonts) { AppendDoc $srcList $SampleDocxFonts "Fonts" }
    if ($All -or $FootNote) { AppendDoc $srcList $SampleDocxFootNote "FootNote" }
    if ($All -or $FormattedText) { AppendDoc $srcList $SampleDocxFormattedText "FormattedText" }
    if ($All -or $Headings) { AppendDoc $srcList $SampleDocxHeadings "Headings" }
    if ($All -or $HierarchicalNumberedList) { AppendDoc $srcList $SampleDocxHierarchicalNumberedList "HierarchicalNumberedList" }
    if ($All -or $HorizontalWhiteSpace) { AppendDoc $srcList $SampleDocxHorizontalWhiteSpace "HorizontalWhiteSpace" }
    if ($All -or $Hyperlink) { AppendDoc $srcList $SampleDocxHyperlink "Hyperlink" }
    if ($All -or $Image) { AppendDoc $srcList $SampleDocxImage "Image" }
    if ($All -or $Justified) { AppendDoc $srcList $SampleDocxJustified "Justified" }
    if ($All -or $NumberedList) { AppendDoc $srcList $SampleDocxNumberedList "NumberedList" }
    if ($All -or $NumberedListRomanNumerals) { AppendDoc $srcList $SampleDocxNumberedListRomanNumerals "NumberedListRomanNumerals" }
    if ($All -or $ParagraphBorder) { AppendDoc $srcList $SampleDocxParagraphBorder "ParagraphBorder" }
    if ($All -or $Plain) { AppendDoc $srcList $SampleDocxPlain "Plain" }
    if ($All -or $RevisionTracking) { AppendDoc $srcList $SampleDocxRevisionTracking "RevisionTracking" }
    if ($All -or $RightJustified) { AppendDoc $srcList $SampleDocxRightJustified "RightJustified" }
    if ($All -or $Section) { AppendDoc $srcList $SampleDocxSection "Section" }
    if ($All -or $SectionWithWatermark) { AppendDoc $srcList $SampleDocxSectionWithWatermark "SectionWithWatermark" }
    if ($All -or $Shading) { AppendDoc $srcList $SampleDocxShading "Shading" }
    if ($All -or $Shape) { AppendDoc $srcList $SampleDocxShape "Shape" }
    if ($All -or $SmartArt) { AppendDoc $srcList $SampleDocxSmartArt "SmartArt" }
    if ($All -or $Symbols) { AppendDoc $srcList $SampleDocxSymbols "Symbols" }
    if ($All -or $Table) { AppendDoc $srcList $SampleDocxTable "Table" }
    if ($All -or $TableOfContents) { AppendDoc $srcList $SampleDocxTableOfContents "TableOfContents" }
    if ($All -or $TextBox) { AppendDoc $srcList $SampleDocxTextBox "TextBox" }
    if ($All -or $TextEffects) { AppendDoc $srcList $SampleDocxTextEffects "TextEffects" }
    if ($All -or $Theme) { AppendDoc $srcList $SampleDocxTheme "Theme" }
    if ($All -or $VerticalWhiteSpace) { AppendDoc $srcList $SampleDocxVerticalWhiteSpace "VerticalWhiteSpace" }

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

