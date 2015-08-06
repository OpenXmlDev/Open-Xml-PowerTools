<#***************************************************************************

Copyright (c) Microsoft Corporation 2014.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************#>

function Test-OpenXmlPowerToolsCmdlets {
    <#
    .SYNOPSIS
    Test all PowerTools for Open XML Cmdlets.
    .DESCRIPTION
    Test all PowerTools for Open XML Cmdlets, making sure that functionality is correct.
    .EXAMPLE
    # Demonstrates the Verbose argument
    Test-OpenXmlPowerToolsCmdlets -Verbose
    .EXAMPLE
    # Demonstrates the basic test of Cmdlets
    Test-OpenXmlPowerToolsCmdlets
    .EXAMPLE
    # Demonstrates execution of a single test
    Test-OpenXmlPowerToolsCmdlets -Test 3
    #>

    [CmdletBinding()]
    param
    (
        [Parameter()]
        [ValidateScript(
        {
            return $_ -is [int16]
        })]
        [int16]$Test
    )

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    $maxTest = 16
  
    if ($Test -ne 0)
    {
        $b = RunTest $Test
        Cleanup
    }
    else  
    {
        $passed = 0
        $failed = 0
        for ($i = 1; $i -le $maxTest; ++$i)
        {
       
            if ($(RunTest $i))
            {
                ++$passed
            }
            else
            {
                ++$failed
            }
        }
        ""
        "Summary"
        "Passed: $passed"
        "Failed: $failed"
    }
	[environment]::CurrentDirectory = $prevCurrentDirectory
}

function RunTest {
    param (
        [int]$test
    )
        
    $testId = [string]::Format("Test {0:0000}", $test)
    $officeVersion = "2010"
    switch ($test)
    {
        #New-Docx Cmdlet
        1 {
            Write-Host "Testing -- New-Docx" -ForegroundColor Magenta
            ""
            Write-Verbose "New-Docx -All"
            New-Docx OxPtTemp-New-Docx.docx -All
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -All,Bookmark,BulletedText,Comment"
            New-Docx OxPtTemp-New-Docx.docx -Comment -Bookmark -All
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Bookmark"
            New-Docx OxPtTemp-New-Docx.docx -Bookmark
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -BulletedText"
            New-Docx OxPtTemp-New-Docx.docx -BulletedText
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Comment"
            New-Docx OxPtTemp-New-Docx.docx -Comment
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -CoverPage"
            New-Docx OxPtTemp-New-Docx.docx -CoverPage
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }
            
            Write-Verbose "New-Docx -Chart"
            New-Docx OxPtTemp-New-Docx.docx -Chart
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -ContentControls"
            New-Docx OxPtTemp-New-Docx.docx -ContentControls
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -EmbeddedWorkbook"
            New-Docx OxPtTemp-New-Docx.docx -EmbeddedWorkbook
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Empty"
            New-Docx OxPtTemp-New-Docx.docx -Empty
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -EndNote"
            New-Docx OxPtTemp-New-Docx.docx -EndNote
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Equation"
            New-Docx OxPtTemp-New-Docx.docx -Equation
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Fields"
            New-Docx OxPtTemp-New-Docx.docx -Fields
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Fonts"
            New-Docx OxPtTemp-New-Docx.docx -Fonts
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -FootNote"
            New-Docx OxPtTemp-New-Docx.docx -FootNote
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -FormattedText"
            New-Docx OxPtTemp-New-Docx.docx -FormattedText
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Headings"
            New-Docx OxPtTemp-New-Docx.docx -Headings
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -HierarchicalNumberedList"
            New-Docx OxPtTemp-New-Docx.docx -HierarchicalNumberedList
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -HorizontalWhiteSpace"
            New-Docx OxPtTemp-New-Docx.docx -HorizontalWhiteSpace
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Hyperlink"
            New-Docx OxPtTemp-New-Docx.docx -Hyperlink
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Image"
            New-Docx OxPtTemp-New-Docx.docx -Image
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Justified"
            New-Docx OxPtTemp-New-Docx.docx -Justified
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -NumberedList"
            New-Docx OxPtTemp-New-Docx.docx -NumberedList
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -ParagraphBorder"
            New-Docx OxPtTemp-New-Docx.docx -ParagraphBorder
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Plain"
            New-Docx OxPtTemp-New-Docx.docx -Plain
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -RevisionTracking"
            New-Docx OxPtTemp-New-Docx.docx -RevisionTracking
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -RightJustified"
            New-Docx OxPtTemp-New-Docx.docx -RightJustified
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Section"
            New-Docx OxPtTemp-New-Docx.docx -Section
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -SectionWithWatermark"
            New-Docx OxPtTemp-New-Docx.docx -SectionWithWatermark
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Shading"
            New-Docx OxPtTemp-New-Docx.docx -Shading
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Shape"
            New-Docx OxPtTemp-New-Docx.docx -Shape
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -SmartArt"
            New-Docx OxPtTemp-New-Docx.docx -SmartArt
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Symbols"
            New-Docx OxPtTemp-New-Docx.docx -Symbols
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Table"
            New-Docx OxPtTemp-New-Docx.docx -Table
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -TableOfContents"
            New-Docx OxPtTemp-New-Docx.docx -TableOfContents
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -TextBox"
            New-Docx OxPtTemp-New-Docx.docx -TextBox
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -TextEffects"
            New-Docx OxPtTemp-New-Docx.docx -TextEffects
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -Theme"
            New-Docx OxPtTemp-New-Docx.docx -Theme
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx -VerticalWhiteSpace"
            New-Docx OxPtTemp-New-Docx.docx -VerticalWhiteSpace
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "New-Docx (no extension)"
            New-Docx OxPtTemp-New-Docx -VerticalWhiteSpace
            $pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            #Write-Verbose "New-Docx (load and save using Word)"
            #New-Docx OxPtTemp-New-Docx -VerticalWhiteSpace -LoadAndSaveUsingWord
            #$pass = Test-OpenXmlValid OxPtTemp-New-Docx.docx -OfficeVersion $officeVersion
            #if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Convert-DocxToHtml
        2 {
            Write-Host "Testing -- Convert-DocxToHtml" -ForegroundColor Magenta
            ""
            Write-Verbose "New-Docx -All"
            New-Docx OxPtTemp-00.docx -All

            Write-Verbose "New-Docx -Bookmark"
            New-Docx OxPtTemp-01.docx -Bookmark

            Write-Verbose "New-Docx -Comment"
            New-Docx OxPtTemp-02.docx -Comment

            Write-Verbose "New-Docx -ContentControls"
            New-Docx OxPtTemp-03.docx -ContentControls

            Write-Verbose "New-Docx -Image"
            New-Docx OxPtTemp-04.docx -Image

            Write-Verbose "New-Docx -BulletedText"
            New-Docx OxPtTemp-05.docx -BulletedText

            Write-Verbose "New-Docx -CoverPage"
            New-Docx OxPtTemp-06.docx -CoverPage
           
            Write-Verbose "New-Docx -Chart"
            New-Docx OxPtTemp-07.docx -Chart

            Write-Verbose "New-Docx -ContentControls"
            New-Docx OxPtTemp-08.docx -ContentControls

            Write-Verbose "New-Docx -EmbeddedWorkbook"
            New-Docx OxPtTemp-09.docx -EmbeddedWorkbook

            Write-Verbose "New-Docx -Empty"
            New-Docx OxPtTemp-10.docx -Empty 

            Write-Verbose "New-Docx -EndNote"
            New-Docx OxPtTemp-11.docx -EndNote

            Write-Verbose "New-Docx -Equation"
            New-Docx OxPtTemp-12.docx -Equation

            Write-Verbose "New-Docx -Fields"
            New-Docx OxPtTemp-13.docx -Fields

            Write-Verbose "New-Docx -Fonts"
            New-Docx OxPtTemp-14.docx -Fonts

            Write-Verbose "New-Docx -FootNote"
            New-Docx OxPtTemp-15.docx -FootNote

            Write-Verbose "New-Docx -FormattedText"
            New-Docx OxPtTemp-16.docx -FormattedText

            Write-Verbose "New-Docx -Headings"
            New-Docx OxPtTemp-17.docx -Headings

            Write-Verbose "New-Docx -HierarchicalNumberedList"
            New-Docx OxPtTemp-18.docx -HierarchicalNumberedList

            Write-Verbose "New-Docx -HorizontalWhiteSpace"
            New-Docx OxPtTemp-19.docx -HorizontalWhiteSpace

            Write-Verbose "New-Docx -Hyperlink"
            New-Docx OxPtTemp-20.docx -Hyperlink

            Write-Verbose "New-Docx -Justified"
            New-Docx OxPtTemp-21.docx -Justified

            Write-Verbose "New-Docx -NumberedList"
            New-Docx OxPtTemp-22.docx -NumberedList

            Write-Verbose "New-Docx -ParagraphBorder"
            New-Docx OxPtTemp-23.docx -ParagraphBorder

            Write-Verbose "New-Docx -Plain"
            New-Docx OxPtTemp-24.docx -Plain

            Write-Verbose "New-Docx -RevisionTracking"
            New-Docx OxPtTemp-25.docx -RevisionTracking

            Write-Verbose "New-Docx -RightJustified"
            New-Docx OxPtTemp-26.docx -RightJustified

            Write-Verbose "New-Docx -Section"
            New-Docx OxPtTemp-27.docx -Section

            Write-Verbose "New-Docx -SectionWithWatermark"
            New-Docx OxPtTemp-28.docx -SectionWithWatermark

            Write-Verbose "New-Docx -Shading"
            New-Docx OxPtTemp-29.docx -Shading

            Write-Verbose "New-Docx -Shape"
            New-Docx OxPtTemp-30.docx -Shape

            Write-Verbose "New-Docx -SmartArt"
            New-Docx OxPtTemp-31.docx -SmartArt

            Write-Verbose "New-Docx -Symbols"
            New-Docx OxPtTemp-32.docx -Symbols

            Write-Verbose "New-Docx -Table"
            New-Docx OxPtTemp-33.docx -Table

            Write-Verbose "New-Docx -TableOfContents"
            New-Docx OxPtTemp-34.docx -TableOfContents

            Write-Verbose "New-Docx -TextBox"
            New-Docx OxPtTemp-35.docx -TextBox

            Write-Verbose "New-Docx -TextEffects"
            New-Docx OxPtTemp-36.docx -TextEffects

            Write-Verbose "New-Docx -Theme"
            New-Docx OxPtTemp-37.docx -Theme

            Write-Verbose "New-Docx -VerticalWhiteSpace"
            New-Docx OxPtTemp-38.docx -VerticalWhiteSpace

            Dir OxPtTemp*.docx | % {
                Write-Verbose "Converting $_ to Html"
                Convert-DocxToHtml $_
            }

            $pass = $false
            try
            {
                Convert-DocxToHtml NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }
                        

            if ((Test-Path OxPtTemp-00.html) -and `
                (Test-Path OxPtTemp-01.html) -and `
                (Test-Path OxPtTemp-02.html) -and `
                (Test-Path OxPtTemp-03.html) -and `
                (Test-Path OxPtTemp-04.html) -and `
                (Test-Path OxPtTemp-05.html) -and `
                (Test-Path OxPtTemp-06.html) -and `
                (Test-Path OxPtTemp-07.html) -and `
                (Test-Path OxPtTemp-08.html) -and `
                (Test-Path OxPtTemp-09.html) -and `
                (Test-Path OxPtTemp-10.html) -and `
                (Test-Path OxPtTemp-11.html) -and `
                (Test-Path OxPtTemp-12.html) -and `
                (Test-Path OxPtTemp-13.html) -and `
                (Test-Path OxPtTemp-14.html) -and `
                (Test-Path OxPtTemp-15.html) -and `
                (Test-Path OxPtTemp-16.html) -and `
                (Test-Path OxPtTemp-17.html) -and `
                (Test-Path OxPtTemp-18.html) -and `
                (Test-Path OxPtTemp-19.html) -and `
                (Test-Path OxPtTemp-20.html) -and `
                (Test-Path OxPtTemp-21.html) -and `
                (Test-Path OxPtTemp-22.html) -and `
                (Test-Path OxPtTemp-23.html) -and `
                (Test-Path OxPtTemp-24.html) -and `
                (Test-Path OxPtTemp-25.html) -and `
                (Test-Path OxPtTemp-26.html) -and `
                (Test-Path OxPtTemp-27.html) -and `
                (Test-Path OxPtTemp-28.html) -and `
                (Test-Path OxPtTemp-29.html) -and `
                (Test-Path OxPtTemp-30.html) -and `
                (Test-Path OxPtTemp-31.html) -and `
                (Test-Path OxPtTemp-32.html) -and `
                (Test-Path OxPtTemp-33.html) -and `
                (Test-Path OxPtTemp-34.html) -and `
                (Test-Path OxPtTemp-35.html) -and `
                (Test-Path OxPtTemp-36.html) -and `
                (Test-Path OxPtTemp-37.html) -and `
                (Test-Path OxPtTemp-38.html))
            {
                Cleanup
            }
            else
            {
                Cleanup
                Write-Host "$testId - Fail"
                return $false
            }

            Write-Verbose "New-Docx -Symbols"
            New-Docx OxPtTemp-32.docx -Symbols

            Write-Verbose "New-Docx -Table"
            New-Docx OxPtTemp-33.docx -Table

            Convert-DocxToHtml *.docx

            Write-Host "$testId - Pass"
            return $true
        }

        #Clear-DocxTrackedRevision
        3 {
            Write-host "Testing -- Clear-DocxTrackedRevision"  -ForegroundColor Magenta
            ""
            $fn = "OxPtTemp-Tracked.docx"
            Write-Verbose "Creating and testing $fn"
            New-Docx $fn -RevisionTracking
            Clear-DocxTrackedRevision $fn
            $pass = Test-OpenXmlValid $fn -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }
            
            $fn = "OxPtTemp-ContentControls.docx"
            Write-Verbose "Creating and testing $fn"
            New-Docx $fn -ContentControls
            Clear-DocxTrackedRevision $fn
            $pass = Test-OpenXmlValid $fn -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }
            
            $fn = "OxPtTemp-NoRevision.docx"
            Write-Verbose "Creating and testing $fn"
            New-Docx $fn -Plain
            Clear-DocxTrackedRevision $fn
            $pass = Test-OpenXmlValid $fn -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Host "Testing -- Get-ChildItem *.docx | Clear-DocxTrackedRevision" -ForegroundColor Magenta
            $fn1 = "OxPtTemp-Tracked1.docx"
            Write-Verbose "Creating and testing $fn1"
            New-Docx $fn1 -RevisionTracking
            
            $fn2 = "OxPtTemp-Tracked2.docx"
            Write-Verbose "Creating and testing $fn2"
            New-Docx $fn2 -RevisionTracking

            $fn3 = "OxPtTemp-Tracked3.docx"
            Write-Verbose "Creating and testing $fn3"
            New-Docx $fn3 -RevisionTracking

            $fn4 = "OxPtTemp-Tracked4.docx"
            Write-Verbose "Creating and testing $fn4"
            New-Docx $fn4 -RevisionTracking

            $fn5 = "OxPtTemp-Tracked5.docx"
            Write-Verbose "Creating and testing $fn5"
            New-Docx $fn5 -RevisionTracking

            Get-ChildItem *.docx | Clear-DocxTrackedRevision
            Cleanup

            Write-Host "Testing -- Clear-DocxTrackedRevision *.docx" -ForegroundColor Magenta
            $fn1 = "OxPtTemp-Tracked1.docx"
            Write-Verbose "Creating and testing $fn1"
            New-Docx $fn1 -RevisionTracking
            
            $fn2 = "OxPtTemp-Tracked2.docx"
            Write-Verbose "Creating and testing $fn2"
            New-Docx $fn2 -RevisionTracking

            $fn3 = "OxPtTemp-Tracked3.docx"
            Write-Verbose "Creating and testing $fn3"
            New-Docx $fn3 -RevisionTracking

            $fn4 = "OxPtTemp-Tracked4.docx"
            Write-Verbose "Creating and testing $fn4"
            New-Docx $fn4 -RevisionTracking

            $fn5 = "OxPtTemp-Tracked5.docx"
            Write-Verbose "Creating and testing $fn5"
            New-Docx $fn5 -RevisionTracking

            Clear-DocxTrackedRevision *.docx
            
            $pass = $false
            try
            {
                Clear-DocxTrackedRevision NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }
                        
            Cleanup

            Write-Host "$testId - Pass"
            return $true
        }

        #Add-DocxText
        4 {
            Write-host "Testing -- Add-DocxText"  -ForegroundColor Magenta
            ""
            $fn = "OxPtTemp-Add-DocxText.docx"
            Write-Verbose "Creating and testing for Empty parameter $fn"
            New-Docx $fn -Empty
            $w = New-WmlDocument $fn
            Add-DocxText $w "Title" -Style "Title"
            Add-DocxText $w "Heading1" -Style "Heading1"
            Add-DocxText $w "Hello World" -Bold -Italic -ForeColor Red
            Add-DocxText $w "Hello World Hello World" -Bold -Italic -Underline -ForeColor White -BackColor Red -Style Heading1
            $w.Save()
            $pass = Test-OpenXmlValid $fn -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #ConvertFrom-Base64
        5 {
            Write-host "Testing -- ConvertFrom-Base64"  -ForegroundColor Magenta
            New-Docx OxPtTemp-New-Docx-Empty.docx -All
            $b64 = ConvertTo-Base64 OxPtTemp-New-Docx-Empty.docx
            ConvertFrom-Base64 OxPtTemp-From-B64.docx $b64

            $pass = Test-OpenXmlValid OxPtTemp-From-B64.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #ConvertFrom-FlatOpc
        6 {
            Write-host "Testing -- ConvertFrom-FlatOpc"  -ForegroundColor Magenta
            New-Docx OxPtTemp-New-Docx.docx -Bookmark
            $a = ConvertTo-FlatOpc OxPtTemp-New-Docx.docx
            ConvertFrom-FlatOpc -OutputPath OxPtTemp-Output1.docx -FlatOpc $a
            $pass = Test-OpenXmlValid OxPtTemp-Output1.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            $t = ConvertTo-FlatOpc OxPtTemp-New-Docx.docx -OutputFormat Text
            ConvertFrom-FlatOpc -OutputPath OxPtTemp-Output2.docx -FlatOpc $t
            $pass = Test-OpenXmlValid OxPtTemp-Output2.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            $x = ConvertTo-FlatOpc OxPtTemp-New-Docx.docx -OutputFormat XDocument
            ConvertFrom-FlatOpc -OutputPath OxPtTemp-Output3.docx -FlatOpc $x
            $pass = Test-OpenXmlValid OxPtTemp-Output3.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            $pass = $false
            Try
            {
                $s = "This is not flat OPC"
                ConvertFrom-FlatOpc -OutputPath OxPtTemp-Output4.docx -FlatOpc $s
            }
            Catch
            {
                $pass = $true
            }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #ConvertTo-Base64
        7 {
            Write-host "Testing -- ConvertTo-Base64"  -ForegroundColor Magenta
            New-Docx OxPtTemp-New-Docx.docx -Empty

            ConvertTo-Base64 OxPtTemp-New-Docx.docx -PowerShellLiteral
            ConvertTo-Base64 OxPtTemp-New-Docx.docx -JavaScriptLiteral
            ConvertTo-Base64 OxPtTemp-New-Docx.docx -CSharpLiteral
            ConvertTo-Base64 OxPtTemp-New-Docx.docx -VBLiteral

            ConvertTo-Base64 OxPtTemp-New-Docx.docx -PowerShellLiteral -WithCode
            ConvertTo-Base64 OxPtTemp-New-Docx.docx -JavaScriptLiteral -WithCode
            ConvertTo-Base64 OxPtTemp-New-Docx.docx -CSharpLiteral -WithCode
            ConvertTo-Base64 OxPtTemp-New-Docx.docx -VBLiteral -WithCode

            $pass = $false
            try
            {
                ConvertTo-Base64 NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #ConvertTo-FlatOpc
        8 {
            Write-host "Testing -- ConvertTo-FlatOpc"  -ForegroundColor Magenta
            New-Docx OxPtTemp-Input1.docx -Empty
            ConvertTo-FlatOpc OxPtTemp-Input1.docx

            $pass = $false
            try
            {
                ConvertTo-FlatOpc NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx OxPtTemp-Input1.docx -Headings
            ConvertTo-FlatOpc OxPtTemp-Input1.docx

            $t = ConvertTo-FlatOpc *.docx -OutputFormat Text
            $x = ConvertTo-FlatOpc *.docx -OutputFormat XDocument
            $z = ConvertTo-FlatOpc *.docx -OutputFormat XmlDocument

            $t = ConvertTo-FlatOpc OxPtTemp-Input1.docx -OutputFormat Text
            $x = ConvertTo-FlatOpc OxPtTemp-Input1.docx -OutputFormat XDocument
            $Z = ConvertTo-FlatOpc OxPtTemp-Input1.docx -OutputFormat XmlDocument

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Merge-Docx
        9 {
            Write-host "Testing -- Merge-Docx"  -ForegroundColor Magenta
            New-Docx OxPtTemp-Input1.docx -Bookmark
            New-Docx OxPtTemp-Input2.docx -BulletedText
            $doc1 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input1.docx")
            $doc2 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input2.docx", $True)
            $sources = ($doc1, $doc2)
            Merge-Docx -OutputPath OxPtTemp-merge01.docx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge01.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx OxPtTemp-Input1.docx -Bookmark
            New-Docx OxPtTemp-Input2.docx -BulletedText
            $doc1 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input1.docx", 1, 2, $false)
            $doc2 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input2.docx", 1, 2, $false)
            $sources = ($doc1, $doc2)
            Merge-Docx -OutputPath OxPtTemp-merge02.docx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge02.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx OxPtTemp-Input1.docx -Bookmark
            New-Docx OxPtTemp-Input2.docx -BulletedText
            $doc1 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input1.docx", 1, 2, $false)
            $doc2 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input2.docx", 1, 2, $false)
            $sources = ($doc1, $doc2)
            Merge-Docx -OutputPath OxPtTemp-merge02.docx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge02.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx OxPtTemp-Input1.docx -Bookmark
            New-Docx OxPtTemp-Input2.docx -BulletedText
            $doc1 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input1.docx")
            $doc2 = New-Object OpenXmlPowerTools.Source("OxPtTemp-Input2.docx", 1, 2, $false)
            $newWml = Merge-Docx -Sources ($doc1, $doc2)
            $newWml.SaveAs("OxPtTemp-out-merge03.docx");

            $pass = Test-OpenXmlValid OxPtTemp-out-merge03.docx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Get-OpenXmlValidationErrors
        10 {
            Write-host "Testing -- Get-OpenXmlValidationErrors"  -ForegroundColor Magenta
            New-Docx OxPtTemp-Input1.docx -Bookmark
            Get-OpenXmlValidationErrors OxPtTemp-Input1.docx -OfficeVersion $officeVersion

            $pass = $false
            try
            {
                Get-OpenXmlValidationErrors NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx $fn -Headings
            Get-OpenXmlValidationErrors *.docx

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Test-OpenXmlValid
        11 {
            Write-host "Testing -- Test-OpenXmlValid"  -ForegroundColor Magenta
            New-Docx OxPtTemp-Input1.docx -Bookmark
            Test-OpenXmlValid OxPtTemp-Input1.docx
               
            $pass = $false
            try
            {
                Test-OpenXmlValid NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx $fn -Headings
            Test-OpenXmlValid *.docx

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Get-DocxMetrics
        12 {
            Write-host "Testing -- Get-DocxMetrics" -ForegroundColor Magenta

            #todo needs fleshed out
            New-Docx OxPtTemp-Plain.docx -Plain
            if ($(Get-DocxMetrics OxPtTemp-Plain.docx).ActiveX) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx OxPtTemp-TrackedRevs.docx -RevisionTracking
            if ($(Get-DocxMetrics OxPtTemp-Plain.docx).RevisionTracking) { Cleanup; Write-Host "$testId - Fail"; return $false; }
            if (-not $(Get-DocxMetrics OxPtTemp-TrackedRevs.docx).RevisionTracking) { Cleanup; Write-Host "$testId - Fail"; return $false; }
               
            $pass = $false
            try
            {
                Get-DocxMetrics NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx $fn -Headings
            Get-DocxMetrics *.docx

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Expand-DocxFormatting
        13 {
            Write-host "Testing -- Expand-DocxFormatting"  -ForegroundColor Magenta
            ""
            $fn = "OxPtTemp-Headings.docx"
            Write-Verbose "Creating and testing $fn"
            New-Docx $fn -Headings
            Expand-DocxFormatting $fn
            $pass = Test-OpenXmlValid $fn -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            $pass = $false
            try
            {
                Expand-DocxFormatting NonExistentFile.docx
            }
            catch
            {
                $pass = $true
            }
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Docx $fn -Headings
            Expand-DocxFormatting *.docx

            Cleanup

            Write-Host "$testId - Pass"
            return $true
        }

        #Format-XML
        14 {
            Write-host "Testing -- Format-XML"  -ForegroundColor Magenta
            ""
            $fn = "OxPtTemp-Headings.docx"
            Write-Verbose "Creating and testing $fn"
            New-Docx $fn -Headings

            $z = $(Format-XML $(Get-DocxMetrics OxPtTemp-Headings.docx).StyleHierarchy)

            Cleanup

            Write-Host "$testId - Pass"
            return $true
        }

        #Merge-Pptx
        15 {
            Write-host "Testing -- Merge-Pptx"  -ForegroundColor Magenta
            New-Pptx OxPtTemp-Input1.pptx -AdjacencyTheme -FiveSlides
            New-Pptx OxPtTemp-Input2.pptx -BlankLayout -ComparisonLayout -TenSlides
            $pres1 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input1.pptx", $true)
            $pres2 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input2.pptx", $true)
            $sources = ($pres1, $pres2)
            Merge-Pptx -OutputPath OxPtTemp-merge01.pptx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge01.pptx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            New-Pptx OxPtTemp-Input1.pptx -AdjacencyTheme -FiveSlides
            New-Pptx OxPtTemp-Input2.pptx -BlankLayout -ComparisonLayout -TenSlides
            $pres1 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input1.pptx", 0, 2, $true)
            $pres2 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input2.pptx", 0, 2, $true)
            $sources = ($pres1, $pres2)
            Merge-Pptx -OutputPath OxPtTemp-merge02.pptx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge02.pptx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "Creating and testing OxPtTemp-Input1.pptx"
            New-Pptx OxPtTemp-Input1.pptx -AdjacencyTheme -FiveSlides
            Write-Verbose "Creating and testing OxPtTemp-Input2.pptx"
            New-Pptx OxPtTemp-Input2.pptx -BlankLayout -ComparisonLayout -TenSlides
            $pres1 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input1.pptx", $true)
            $pres2 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input2.pptx", $true)
            $sources = ($pres1, $pres2)
            Write-Verbose "Calling Merge-Pptx"
            Merge-Pptx -OutputPath OxPtTemp-merge01.pptx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge01.pptx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Write-Verbose "Creating and testing OxPtTemp-Input1.pptx"
            New-Pptx OxPtTemp-Input1.pptx -AdjacencyTheme -FiveSlides
            Write-Verbose "Creating and testing OxPtTemp-Input2.pptx"
            New-Pptx OxPtTemp-Input2.pptx -BlankLayout -ComparisonLayout -TenSlides
            $pres1 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input1.pptx", 0, 2, $true)
            $pres2 = New-Object OpenXmlPowerTools.SlideSource("OxPtTemp-Input2.pptx", 0, 2, $true)
            $sources = ($pres1, $pres2)
            Write-Verbose "Calling Merge-Pptx"
            Merge-Pptx -OutputPath OxPtTemp-merge02.pptx -Sources $sources

            $pass = Test-OpenXmlValid OxPtTemp-merge02.pptx -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup
            Write-Host "$testId - Pass"
            return $true
        }

        #Out-Xlsx
        16 {
            Write-host "Testing -- Out-Xlsx"  -ForegroundColor Magenta
            ""
            $fn = "OxPtTemp-Directory.xlsx"
            Write-Verbose "Creating and testing $fn"
            dir .. -Directory | Out-Xlsx -FileName $fn -SheetName "SheetOne" -TableName "Table"

            $pass = Test-OpenXmlValid $fn -OfficeVersion $officeVersion
            if (-not $pass) { Cleanup; Write-Host "$testId - Fail"; return $false; }

            Cleanup

            Write-Host "$testId - Pass"
            return $true
        }

    }
    Cleanup
    Write-Host "Invalid test number ($test) passed to RunTest"
    return $false
}

function Cleanup {
    Write-Verbose "Cleanup"
    del OxPtTemp-* -Recurse -Force -ErrorAction SilentlyContinue
}
