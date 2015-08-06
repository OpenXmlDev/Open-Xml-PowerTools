Open-XML-PowerTools
===================
The Open XML PowerTools provides guidance and example code for programming with Open XML
Documents (DOCX, XLSX, and PPTX).  It is based on, and extends the functionality
in the Open XML SDK (https://github.com/OfficeDev/Open-XML-SDK).

It supports scenarios such as:
- Splitting DOCX/PPTX files into multiple files.
- Combining multiple DOCX/PPTX files into a single file.
- Populating content in template DOCX files with data from XML.
- High-fidelity conversion of DOCX to HTML.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.

Copyright (c) Microsoft Corporation 2012-2015
Licensed under the Microsoft Public License.
See License.txt in the project root for license information.

News
====
We are happy to announce the release of the Open XML PowerTools Version 4.0.  There are lots of new features in 4.0, including:
- Renaming the project and the PowerShell module to Open-Xml-PowerTools, to be consistent with the Open-Xml-Sdk.
- DocumentAssembler module, which enables populating a template DOCX with data from an XML file.
- SpreadsheetWriter module, which enables writing far simpler code to generate an XLSX file, and enables a streaming approach.
- Many xUnit tests!!!, which will enable a far nimbler process for accepting contributes to PowerTools via Git pull requests.
- New PowerShell Cmdlet: Complete-DocxTemplateFromXml, which populates a template document from XML
- New PowerShell Cmdlet: Out-Xlsx, which produces an 

Build Instructions
==================

To use the PowerShell Cmdlets, you need not install Visual Studio.  The following video shows how to install and use PowerTools
so that you can use the Cmdlets:

<insert video here>

The short form of the installation instructions are:
1)  Make sure you are running PowerShell 3.0 or later
2)  If necessary, run PowerShell as administrator, Set-ExecutionPolicy Unrestricted (or RemoteSigned)
3)  Clone Open-Xml-PowerTools to %HOMEDRIVE%/%HOMEUSER%/Documents/WindowsPowerShell/Modules/Open-Xml-PowerTools
4)  Import the module in PowerShell:  Import-Module Open-Xml-PowerTools

To build the library, you must have some version of Visual Studio
installed.  Visual Studio Community Edition will work just fine:
https://www.visualstudio.com/en-us/products/visual-studio-community-vs.aspx

To build the Open XML SDK:
- clone the repo at https://github.com/OfficeDev/Open-XML-SDK
- Start a Visual Studio command prompt, and check into the directory that contains the repo
- Use MSBUILD to build the SDK  (C:> MSBUILD Open-Xml-Sdk.sln)
- In your program that uses the Open XML SDK, add references to the newly built libraries in bin/Debug

Instead of using MSBUILD, you can also open the solution using Visual Studio and build it.

Change Log
==========

Version 4.0.0 : August 6, 2015
- New DocumentAssember module
- New SpreadsheetWriter module
- New Cmdlet: Complete-DocxTemplateFromXml
- Fix DocumentBuilder: deal with headers / footers more rationally
- Enhance DocumentBuilder: add option to discard headers / footers from section (but keep layout of section)
- Fix RevisionAccepter: deal with w:moveTo immediately before a table
- New test document library in the TestFiles directory
- Cleaned up build system
- Build using the open source Open-Xml-SDK and the new System.IO.Packaging by default
- Back port to .NET 3.5
- Rename the PowerShell module to Open-Xml-PowerTools

Version 3.1.11 : June 30, 2015
- Updated projects and solutions to build with the open source Open XML SDK and new System.IO.Packaging

Version 3.1.10 : June 14, 2015
- Changed Out-Xlsx Cmdlet to C# implementation
- Fix Add-DocxText

Version 3.1.09 : April 20, 2015
- Fix OpenXmlRegex: PowerPoint 2007 and xml:space issues, causing 2007 to not open PPTX's

Version 3.1.08 : March 13, 2015
- Added Out-Xlsx Cmdlet

Version 3.1.07 : February 9, 2015
- Added Merge-Pptx Cmdlet
- Added New-Pptx Cmdlet
- Added New-PmlDocument
- Fixed help for Merge-Docx
- Don't throw duplicate attribute exception when running FormattingAssembler.AssembleFormatting
  twice on same document.

Version 3.1.06 : February 7, 2015
- Added Expand-DocxFormatting Cmdlet
- Cmdlets do not keep a handle to the current directory, preventing deletion of the directory.
- Added additional tests to Test-OxPtCmdlets

Version 3.1.05 : January 29, 2015
- Added GetListItemText_zh_CN.cs
- Fixed GetListItemText_fr_FR.cs
- Partially fixed GetListItemText_ru_RU.cs
- Fixed GetListItemText_Default.cs
- Added better support in ListItemRetriever.cs
- Added FileUtils class in PtUtil.cs

Version 3.1.04 : December 17, 2014
- Added Get-DocxMetrics Cmdlet
- Added New-WmlDocument Cmdlet
- Added MetricsGetter.cs module
- Added MettricsGetter01.cs module, along with sample documents
- Reworked Add-DocxText, new style of using it with New-WmlDocument

Version 3.1.03 : December 9, 2014
- Added ChartUpdater.cs module
- Added ChartUpdater01.cs module, along with sample documents
- Added Test-OxPtCmdlets Cmdlet

Version 3.1.02 : December 1, 2014
- Added Add-DocxText Cmdlet

Version 3.1.01 : November 23, 2014
- Added Convert-DocxToHtml Cmdlet
- Added Chinese and Hebrew sample documents
- Cmdlets in this release
	Clear-DocxTrackedRevision
	Convert-DocxToHtml
	ConvertFrom-Base64
	ConvertFrom-FlatOpc
	ConvertTo-Base64
	ConvertTo-FlatOpc
	Get-OpenXmlValidationErrors
	Merge-Docx
	New-Docx
	Test-OpenXmlValid

Version 3.1.00 : November 13, 2014
- Changed installation process - no longer requires compilation using Visual Studio
- Added ConvertTo-FlatOpc Cmdlet
- Added ConvertFrom-FlatOpc Cmdlet
- Changed parameters for Test-OpenXmlValid, Get-OpenXmlValidationErrors
- Removed the unnecessary 1/2 second sleep when doing Word automation in the New-Docx Cmdlet

Version 3.0.00 : October 29, 2014
- New release of cmdlets that are written as 'Advanced Functions' instead of in C#.

Procedures for enhancing OxPt
-----------------------------
There are a variety of things to do when adding a new CmdLet to OxPt:
- Write the new CmdLet.  Put it in OxPtCmdlets
- Modify OxPt.psm1
    Call the new Cmdlet script to make the function available
    Modify Export-ModuleMember function to export the Cmdlet and any aliases
- Update Readme.txt, describing the enhancement
- Add a new test to Test-OxPtCmdlets.ps1
- Update Downloads page on CodePlex