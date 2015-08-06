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

function Clear-DocxTrackedRevision {
    <#
    .SYNOPSIS
    Accepts revisions (tracked changes) in a DOCX document.
    .DESCRIPTION
    This cmdlet accepts revisions (tracked changes) in a DOCX document.
    .EXAMPLE
    # Simple use
    Clear-DocxTrackedRevision MyFile.docx
    .EXAMPLE
    # Pipes DOCX into Clear-DocxTrackedRevision
    Get-ChildItem *.docx | Clear-DocxTrackedRevision
    .EXAMPLE
    # Uses wildcard
    Clear-DocxTrackedRevision *.docx
    .PARAMETER FileName
    The document to accept tracked changes
    .PARAMETER Force
    If set, suppresses confirmation
    #>
    [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Medium')]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True, HelpMessage='What document would you like to remove tracked revisions from?')]
        [ValidateScript(
        {
            $prevCurrentDirectory = [Environment]::CurrentDirectory
            [environment]::CurrentDirectory = $(Get-Location)
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_))
            {
            	[environment]::CurrentDirectory = $prevCurrentDirectory
                return $True
            }
            else
            {
                if (Test-Path $_)
                {
                	[environment]::CurrentDirectory = $prevCurrentDirectory
                    return $True
                }
                else
                {
                	[environment]::CurrentDirectory = $prevCurrentDirectory
                    Throw "$_ is not a valid filename"
                }
            }
        })]
        [SupportsWildcards()]
        [string[]]$FileName,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$Force
    )
  
    begin {
        $prevCurrentDirectory = [environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)
        write-verbose "Accepting revisions in $fileName"
    }
  
    process {
        write-verbose "Beginning process loop"
        foreach ($argItem in $fileName) {
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($argItem))
            {
                $dir = New-Object -TypeName System.IO.DirectoryInfo $(Get-Location)
                foreach ($fi in $dir.GetFiles($argItem))
                {
                    if ($force -or $pscmdlet.ShouldProcess($fi)) {
                        Clear-DocxTrackedRevision-Helper($fi)
                    }
                }
            }
            else
            {
                if ($force -or $pscmdlet.ShouldProcess($argItem)) {
                    $fi = New-Object System.IO.FileInfo $argItem
                    Clear-DocxTrackedRevision-Helper($fi)
                }
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Clear-DocxTrackedRevision-Helper {
    param (
        [System.IO.FileInfo]$fi
    )
    if (IsWordprocessingML($fi.Extension))
    {
        Write-Verbose "Accepting tracked revisions for $fi.FullName"
        $wml = new-object OpenXmlPowerTools.WmlDocument $fi.FullName
        $newWml = [OpenXmlPowerTools.RevisionAccepter]::AcceptRevisions($wml)
        $newWml.SaveAs($fi.FullName)
    }
    else
    {
        Throw "Invalid Open XML file type for clearing tracked revisions"
    }
}

New-Alias Accept-DocxTrackedRevision Clear-DocxTrackedRevision
New-Alias AcceptRevision Clear-DocxTrackedRevision
