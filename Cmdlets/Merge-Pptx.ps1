<#***************************************************************************

Copyright (c) Microsoft Corporation 2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************#>

function Merge-Pptx {
    <#
    .SYNOPSIS
    Merge multiple source presentations into a new presentation.
    .DESCRIPTION
    Create a new presentation that contains specified contents of multiple source presentations. All or part
    of the source presentations may be merged into the final presentation. You detail which parts you want
    from multiple source documents using the Sources parameter.
    
    There are some parts of a presentation that cannot be merged, such as settings. These will all come
    from the first source presentation. You can specify a 0 Count for the first source document to get
    those settings without including any content.
    .EXAMPLE
    # Takes all of both presentations, using the slide master of the first presentation.
    $pml1 = New-PmlDocument "Input1.pptx"
    $pres1 = New-Object OpenXmlPowerTools.SlideSource($pml1, $true)
    $pml2 = New-PmlDocument "Input2.pptx"
    $pres2 = New-Object OpenXmlPowerTools.SlideSource($pml2, $false)
    $sources = ($pres1, $pres2)
    Merge-Pptx -OutputPath merge01.pptx -Sources $sources
    .EXAMPLE
    # Takes the first two slides of both presentations, including the master slides from both presentations.
    $pres1 = New-Object OpenXmlPowerTools.SlideSource($(New-PmlDocument "Input1.pptx"), 0, 2, $true)
    $pres2 = New-Object OpenXmlPowerTools.SlideSource($(New-PmlDocument "Input2.pptx"), 0, 2, $true)
    $sources = ($pres1, $pres2)
    Merge-Pptx -OutputPath merge02.pptx -Sources $sources
    .EXAMPLE
    # Produces a new presentation with the second slide removed from the presentation.
    $pres1 = New-Object OpenXmlPowerTools.SlideSource($(New-PmlDocument "Input1.pptx"), 0, 1, $true)
    $pres2 = New-Object OpenXmlPowerTools.SlideSource($(New-PmlDocument "Input1.pptx"), 2, 9999, $false)
    $newPml = Merge-Pptx -Sources ($pres1, $pres2) -Verbose
    $newPml.SaveAs($(Join-Path $pwd.Path "merge03.pptx"))
    .PARAMETER Sources
    This parameter is an array of one or more OpenXmlPowerTools.SlideSource objects that define each source
    presentation and range of slides that will be copied from that source presentation into the merged
    presentation. The Source object may specify an entire presentation or part of the presentation.

    The Start argument of the SlideSource object specifies by zero-based index the first slide to include,
	and the Count argument specifies the number of slides to include.

    Important note: the Start parameter in the SlideSource object is zero-based.
    .PARAMETER OutputPath
    Path and name of the file to create with the new content.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True)]
        [OpenXmlPowerTools.SlideSource[]]$Sources,

        [Parameter(Mandatory=$False)]
        [string]$OutputPath
    )

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    Write-Verbose "Merging PPTX documents"
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
        Write-Verbose "  File Name: $($s.PmlDocument.FileName)"
        Write-Verbose "  Start: $($s.Start)"
        if ($s.Count -eq 2147483647)
        {
            Write-Verbose "  Count: end of presentation"
        }
        else
        {
            Write-Verbose "  Count: $($s.Count)"
        }
        Write-Verbose "  Keep Master: $($s.KeepMaster)"
        Write-Verbose ""
    }

    $mergedPmlDocument = [OpenXmlPowerTools.PresentationBuilder]::BuildPresentation($Sources)
    if ($OutputPath -ne [string]::Empty)
    {
        $mergedPmlDocument.SaveAs($OutputPath)
    }
    else
    {
        $mergedPmlDocument
    }

	[environment]::CurrentDirectory = $prevCurrentDirectory
}

New-Alias BuildPresentation Merge-Pptx
