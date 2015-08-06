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

***************************************************************************#>

function New-Pptx {
    <#
    .SYNOPSIS
    Creates a new sample Pptx.
    .DESCRIPTION
    #TBD
    .EXAMPLE
    #TBD
    .PARAMETER OutputPath
    Path and name of the sample file to create.
    .PARAMETER All
    Creates a presentation with every feature in it.
    .PARAMETER LoadAndSaveUsingPowerPoint
    Uses PowerPoint Automation to load and save the presentation after PresentationBuilder finishes building it.
    .PARAMETER AdjacencyTheme
    Adjacency Theme
    .PARAMETER AnglesTheme
    Angles Theme.
    .PARAMETER ApexTheme
    Apex Theme.
    .PARAMETER BlankLayout
    Slide with Blank layout.
    .PARAMETER ComparisonLayout
    Slide with Comparison layout.
    .PARAMETER ContentWithCaptionLayout
    Slide with Content With Caption layout.
    .PARAMETER Empty
    No slides.
    .PARAMETER FiveSlides
    No slides.
    .PARAMETER PictureWithCaptionLayout
    Slide with Picture With Caption layout.
    .PARAMETER SectionHeaderLayout
    Slide with Section Header layout.
    .PARAMETER SlideTransitions
    Slide Transitions.
    .PARAMETER TenSlides
    Ten Slides.
    .PARAMETER TitleAndContentLayout
    Slide with Title and Content layout.
    .PARAMETER TitleLayout
    Slide with Title layout.
    .PARAMETER TitleOnlyLayout
    Slide with Title Only layout.
    .PARAMETER TwoContentLayout
    Slide with Two Content layout.
    .PARAMETER WaveFormTheme
    Wave Form Theme.
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
        [bool]$LoadAndSaveUsingPowerPoint,

        [Parameter(Mandatory = $false)]
        [scriptblock]
        $ScriptBlock,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$AdjacencyTheme,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$AnglesTheme,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$ApexTheme,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$BlankLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$ComparisonLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$ContentWithCaptionLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Empty,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$FiveSlides,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$PictureWithCaptionLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$SectionHeaderLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$SlideTransitions,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TenSlides,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TitleAndContentLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TitleLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TitleOnlyLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$TwoContentLayout,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$WaveFormTheme

    )

    Write-Verbose "Creating a sample PPTX"
    if ($OutputPath -ne [string]::Empty)
    {
        Write-Verbose "  Output document: $OutputPath"
    }
    else
    {
        Write-Verbose "  No output document, returning PmlDocument object"
    }

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    $srcList = New-Object 'System.Collections.Generic.List[OpenXmlPowerTools.SlideSource]'

    if ($All -or $AdjacencyTheme) { AppendPresentation $srcList $SamplePptxAdjacencyTheme "AdjacencyTheme" }
    if ($All -or $AnglesTheme) { AppendPresentation $srcList $SamplePptxAnglesTheme "AnglesTheme" }
    if ($All -or $ApexTheme) { AppendPresentation $srcList $SamplePptxApexTheme "ApexTheme" }
    if ($All -or $BlankLayout) { AppendPresentation $srcList $SamplePptxBlankLayout "BlankLayout" }
    if ($All -or $ComparisonLayout) { AppendPresentation $srcList $SamplePptxComparisonLayout "ComparisonLayout" }
    if ($All -or $ContentWithCaptionLayout) { AppendPresentation $srcList $SamplePptxContentWithCaptionLayout "ContentWithCaptionLayout" }
    if ($All -or $Empty) { AppendPresentation $srcList $SamplePptxEmpty "Empty" }
    if ($All -or $FiveSlides) { AppendPresentation $srcList $SamplePptxFiveSlides "FiveSlides" }
    if ($All -or $PictureWithCaptionLayout) { AppendPresentation $srcList $SamplePptxPictureWithCaptionLayout "PictureWithCaptionLayout" }
    if ($All -or $SectionHeaderLayout) { AppendPresentation $srcList $SamplePptxSectionHeaderLayout "SectionHeaderLayout" }
    if ($All -or $SlideTransitions) { AppendPresentation $srcList $SamplePptxSlideTransitions "SlideTransitions" }
    if ($All -or $TenSlides) { AppendPresentation $srcList $SamplePptxTenSlides "TenSlides" }
    if ($All -or $TitleAndContentLayout) { AppendPresentation $srcList $SamplePptxTitleAndContentLayout "TitleAndContentLayout" }
    if ($All -or $TitleLayout) { AppendPresentation $srcList $SamplePptxTitleLayout "TitleLayout" }
    if ($All -or $TitleOnlyLayout) { AppendPresentation $srcList $SamplePptxTitleOnlyLayout "TitleOnlyLayout" }
    if ($All -or $TwoContentLayout) { AppendPresentation $srcList $SamplePptxTwoContentLayout "TwoContentLayout" }
    if ($All -or $WaveFormTheme) { AppendPresentation $srcList $SamplePptxWaveFormTheme "WaveFormTheme" }

    $randomizedSources = $srcList | % {
        @{
            Source = $_
            Rand = Get-Random
        }} | Sort-Object { $_.Rand } | % { $_.Source }

    $mergedPmlDocument = [OpenXmlPowerTools.PresentationBuilder]::BuildPresentation($randomizedSources)
    if ($OutputPath -ne [string]::Empty)
    {
        $outputFi = New-Object System.IO.FileInfo $OutputPath
        if ($outputFi.Extension.ToLower() -ne '.pptx')
        {
            $newFileName = $(Join-Path $outputFi.Directory ($outputFi.BaseName + ".pptx"))
            $outputFi = New-Object System.IO.FileInfo $newFileName
        }
        if ($LoadAndSaveUsingPowerPoint)
        {
            $powerPointIsRunning = $false
            try
            {
                $pptProcessId = (Get-Process powerpnt -WarningAction SilentlyContinue -ErrorAction SilentlyContinue) 2> $null
                if ($pptProcessId -ne $null)
                {
                    $powerPointIsRunning = $true
                }
            }
            catch
            {
            }

            if (-not $powerPointIsRunning)
            {
                $tempDocName = $(Join-Path $outputFi.DirectoryName ($outputFi.BaseName + "-Temp.pptx"))
                #zzz
                $mergedPmlDocument.SaveAs($tempDocName)
                $PowerPoint = New-Object -ComObject Powerpoint.Application
                # $PowerPoint.Visible = $false # not allowed with PowerPoint
                $Pres = $PowerPoint.Presentations.Open($tempDocName)
                $presFilename = $outputFi.FullName
                Try
                {
                    $Pres.SaveAs($presFilename)
                }
                Catch
                {
                    $Pres.SaveAs([ref]$presFilename)
                }
                $Pres.Close()
                ReleaseComObject($Pres)
                $PowerPoint.Quit()
                ReleaseComObject($PowerPoint)
                Remove-Item $tempDocName
                while ($true)
                {
                    try
                    {
                        $pptProcessId = (Get-Process powerpnt -WarningAction SilentlyContinue -ErrorAction SilentlyContinue) 2> $null
                        if ($pptProcessId -ne $null)
                        {
                            # strange behavior - if write to host, then waiting for process to quit works.
                            # if never write to host, then powerpoint runs forever.
                            # We can put this down to weird behavior on the part of automation of PowerPoint
                            # This is a common problem out there.

                            $c = [char]0x0008
                            [System.Console]::Write($c.ToString());
                            [System.Threading.Thread]::Sleep(50)
                            Write-Verbose "PowerPoint still running, wait a bit"
                            continue
                        }
                        break
                    }
                    catch
                    {
                    }
                    break
                }
            }
            else
            {
                Write-Error "Can't load and save using PowerPoint if PowerPoint is already running"
                #zzz
                $mergedPmlDocument.SaveAs($outputFi.FullName)
            }
        }
        else
        {
            #zzz
            $mergedPmlDocument.SaveAs($outputFi.FullName)
        }
    }
    else
    {
        $mergedPmlDocument
    }

    [Environment]::CurrentDirectory = $prevCurrentDirectory
}

function AppendPresentation {
    param (
        $srcList,
        $ba,
        [string]$argName
    )
    $b64b = $ba.Replace("\r\n", "")
    $ba = [System.Convert]::FromBase64String($b64b)
    $pml = (New-Object OpenXmlPowerTools.PmlDocument("DummyName.pptx", $ba))
    $src = (New-Object OpenXmlPowerTools.SlideSource($pml, $true))
    $srcList.Add($src)
}

