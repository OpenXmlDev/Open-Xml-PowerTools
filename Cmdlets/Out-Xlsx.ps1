<#***************************************************************************

Copyright (c) Microsoft Corporation 2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

***************************************************************************#>

function Out-Xlsx {
    <#
    .SYNOPSIS
    Sends output to an Excel workbook.
    .DESCRIPTION
    This cmdlet sends output to an Excel workbook.  You can either store data simply as cells in a sheet, or as an Excel named table.
    .EXAMPLE
    # Simple use
    Get-ChildItem | Out-Xlsx Out1.xlsx
    .EXAMPLE
    # Outputting as a table, with table name and sheet name specified.
    Get-ChildItem | Out-Xlsx -FileName Out2.xlsx -SheetName "MySheet" -TableName "Files"
    # Outputting as a data in a sheet, no table, then open the resulting XLSX with Excel.
    Get-ChildItem | Out-Xlsx -FileName Out3.xlsx -SheetName "MySheet" -OpenWithExcel
    .PARAMETER FileName
    The Open XML spreadsheet to generate
    .PARAMETER SheetName
    Specifies the sheet name for the sheet that will contain the data.
    .PARAMETER TableName
    If set, creates an Excel table with the specified name, and populates the table with the data.
    .PARAMETER OpenWithExcel
    If set, upon completion of the creation of the XLSX, opens the file using Excel.
    #>
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Medium')]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True)]
        [object[]]$InputObject,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False)]
        [ValidateScript(
        {
            $prevCurrentDirectory = [Environment]::CurrentDirectory
            [environment]::CurrentDirectory = $(Get-Location)

            if (Test-Path $_)
            {
                [environment]::CurrentDirectory = $prevCurrentDirectory
                Throw "$_ already exists"
            }
            else
            {
                [environment]::CurrentDirectory = $prevCurrentDirectory
                return $True
            }
        })]
        [string]$FileName,

        [Parameter(Mandatory=$False)]
        [string]$SheetName,

        [Parameter(Mandatory=$False)]
        [string]$TableName,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$OpenWithExcel
    )
  
    begin {
        $prevCurrentDirectory = [environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)

        $Rows = New-Object Collections.Generic.LinkedList[OpenXmlPowerTools.SpreadsheetWriter.Row]
        if ($SheetName -eq $null -or $SheetName -eq '')
        {
            $SheetName = 'Sheet1'
        }

        if ($FileName -eq $null -or $FileName -eq '')
        {
            $FileName = $(Join-Path $env:TEMP ([OpenXmlPowerTools.FileUtils]::GetDateTimeStampedFileInfo('Temp-Xlsx-', ".xlsx")))
            $OpenWithExcel = $True
        }

        $outputFi = New-Object System.IO.FileInfo $FileName
        if ($outputFi.Extension.ToLower() -ne '.xlsx')
        {
            $newFileName = $(Join-Path $outputFi.Directory ($outputFi.BaseName + ".xlsx"))
            $outputFi = New-Object System.IO.FileInfo $newFileName
        }
        write-verbose "Writing to $fileName"
    }
  
    process {
       write-verbose "Beginning process loop"
        
       $InputObject | % {
            $propertyValues = New-Object System.Collections.Generic.List[String]

            $row = New-Object -TypeName OpenXmlPowerTools.SpreadsheetWriter.Row

            $dataCells = New-Object Collections.Generic.LinkedList[OpenXmlPowerTools.SpreadsheetWriter.Cell]

            $properties = Get-Member -InputObject $_ -MemberType Property, Properties, CodeProperty, NoteProperty
            $listOfPropNames = $(ForEach-Object -InputObject $properties { $_.Name }) | Sort-Object
            $object = $_

            $listOfPropNames | % {
                $dataCell = New-Object -TypeName OpenXmlPowerTools.SpreadsheetWriter.Cell
                if ($object.$_ -eq $null)
                {
                    $dataCell.Value = ''
                }
                else
                {
                    $dataCell.Value = $object.$_.ToString()
                }
                $dataCell.CellDataType =  [OpenXmlPowerTools.SpreadsheetWriter.CellDataType]::String
                $dataCell.HorizontalCellAlignment = [OpenXmlPowerTools.SpreadsheetWriter.HorizontalCellAlignment]::Left
                $dataCells.Add($dataCell)
            }
           
            $row.Cells = $dataCells
            $Rows.Add($row)
       }
    }
     
    end {

        [OpenXmlPowerTools.SpreadsheetWriter.Workbook]$workBook = New-Object -TypeName OpenXmlPowerTools.SpreadsheetWriter.Workbook
        $workBook.Worksheets = New-Object Collections.Generic.LinkedList[OpenXmlPowerTools.SpreadsheetWriter.Worksheet]
        $workSheet = New-Object -TypeName OpenXmlPowerTools.SpreadsheetWriter.Worksheet
        
        $workSheet.Name = $SheetName
        if ($TableName -ne $null -and $TableName -ne '')
        {
            $workSheet.TableName = $TableName
        }

        $workSheet.ColumnHeadings = New-Object Collections.Generic.LinkedList[OpenXmlPowerTools.SpreadsheetWriter.Cell]
        
        if ($listOfPropNames -ne $null -and $listOfPropNames -ne '')
        {
            $columnValues = $listOfPropNames.Split(' ') | Sort-Object

            $columnValues | % {
                $cell = New-Object -TypeName OpenXmlPowerTools.SpreadsheetWriter.Cell
                $cell.Value = [string]($_.ToString())
                $cell.Bold = $true
                $cell.HorizontalCellAlignment = [OpenXmlPowerTools.SpreadsheetWriter.HorizontalCellAlignment]::Left
                $cell.CellDataType = [OpenXmlPowerTools.SpreadsheetWriter.CellDataType]::String
                $workSheet.ColumnHeadings.Add($cell)
            }

            $workSheet.Rows = $Rows
            $workBook.Worksheets.Add($workSheet)
            [OpenXmlPowerTools.SpreadsheetWriter.SpreadsheetWriter]::Write($outputFi.FullName, $workBook);

            if ($OpenWithExcel)
            {
                 Invoke-Item $outputFi.FullName
            }
        }
        
        [environment]::CurrentDirectory = $prevCurrentDirectory
     }
}
