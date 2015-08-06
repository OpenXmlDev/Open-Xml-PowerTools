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

function Complete-DocxTemplateFromXml {
    <#
    .SYNOPSIS
    Assemble a DOCX from a template document and an XML data file.
    .DESCRIPTION
    Processes a template DOCX and an XML datafile, producing a new document where the content of
    content controls are replaced with data from the XML file.  The contents of the content controls are
    XPath expressions that identify the elements and attributes that are the source of the replacement
    contents of the content controls.  It supports generation of populated tables from repeating data
    in the XML file, repeating sections from repeating data, and conditional sections, where the content
    can be optionally included in the generated document based on a test of data in the XML file.

    This Cmdlet is a thin wrapper of the DocumentAssembler.cs module.
    .EXAMPLE
    # Replaces the contents of content controls that have ./Name and ./Age XPath expressions in them.
    # You can find the template document for this example in the Cmdlet-Examples directory.
    $data = [XML] "
    <Data>
        <Name>Eric White</Name>
        <Age>53</Age>
    </Data>
    "
    Complete-DocxTemplateFromXml -OutputPath Generated01.docx -Template Template.docx -XmlData $data
    .EXAMPLE
    # Replaces the contents of content controls that have ./Name and ./Age XPath expressions in them.
    # You can find the template document and the XML file for this example in the Cmdlet-Examples directory.
    [xml]$data = Get-Content TemplateData.xml
    Complete-DocxTemplateFromXml -OutputPath Generated01.docx -Template Template.docx -XmlData $data
    .PARAMETER Template
    Path and file name to the template document that contains content controls
    with XPath expressions in them.
    .PARAMETER OutputPath
    Path and file name of the file to create with the new content.
    .PARAMETER XmlData
    PowerShell XML variable that contains the data that will replace the contents of content controls.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True)]
        [string]$Template,

        [Parameter(Mandatory=$True)]
        [xml]$XmlData,

        [Parameter(Mandatory=$False)]
        [string]$OutputPath
    )

    $prevCurrentDirectory = [Environment]::CurrentDirectory
    [Environment]::CurrentDirectory = $(pwd).Path

    Write-Verbose "Assembling DOCX documents"
    if ($OutputPath -ne [string]::Empty)
    {
        Write-Verbose "  Output document: $OutputPath"
    }
    else
    {
        Write-Verbose "  No output document, returning WmlDocument object"
    }

    [bool]$templateError = $false
    $wmlTemplate = New-WmlDocument $Template
    $assembledWmlDocument = [OpenXmlPowerTools.DocumentAssembler]::AssembleDocument($wmlTemplate, $XmlData, [ref] $templateError)
    if ($templateError)
    {
        Write-Error "Template document contains errors"
    }
    if ($OutputPath -ne [string]::Empty)
    {
        $assembledWmlDocument.SaveAs($OutputPath)
    }
    else
    {
        $assembledWmlDocument
    }

	[environment]::CurrentDirectory = $prevCurrentDirectory
}
