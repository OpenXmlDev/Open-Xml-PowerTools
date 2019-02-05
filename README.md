[ARCHIVED] Open-XML-PowerTools
===================
This repository is no longer maintained by Microsoft. It has been archived and can still be forked and cloned for use and continued developement. 

If you're looking for a fork of this project that is actively maintained, try the following: 

[https://github.com/EricWhiteDev/Open-Xml-PowerTools](https://github.com/EricWhiteDev/Open-Xml-PowerTools)

The Open XML PowerTools provides guidance and example code for programming with Open XML
Documents (DOCX, XLSX, and PPTX).  It is based on, and extends the functionality
of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

It supports scenarios such as:
- Splitting DOCX/PPTX files into multiple files.
- Combining multiple DOCX/PPTX files into a single file.
- Populating content in template DOCX files with data from XML.
- High-fidelity conversion of DOCX to HTML/CSS.
- High-fidelity conversion of HTML/CSS to DOCX.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Comparing two DOCX files, producing a DOCX with revision tracking markup, and enabling retrieving a list of revisions.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.
- Extracting data (along with formatting) from spreadsheets.

Copyright (c) Microsoft Corporation 2012-2017
Licensed under the MIT License.
See License in the project root for license information.

News
====
New Release!  Version 4.4.

This version has a completely re-written WmlComparer.cs, which now supports nested tables and text boxes.  WmlComparer.cs is a module that compares two DOCX files and
produces a DOCX with revision tracking markup.  It enables retrieving a list of revisions.

Open-Xml-PowerTools Content
===========================

There is a lot of content about Open-Xml-PowerTools at the [Open-Xml-PowerTools Resource Center at OpenXmlDeveloper.org](http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx)

See:
- [DocumentBuilder Resource Center](http://openxmldeveloper.org/wiki/w/wiki/documentbuilder.aspx)
- [PresentationBuilder Resource Center](http://openxmldeveloper.org/wiki/w/wiki/presentationbuilder.aspx)
- [HtmlConverter Resource Center](http://openxmldeveloper.org/wiki/w/wiki/htmlconverter.aspx)
- [Introduction to DocumentAssembler](https://www.youtube.com/watch?v=9QqzCgfqA2Y)
- [Contributing to Open-Xml-PowerTools via GitHub](https://www.youtube.com/watch?v=Ii7z9L6Dkko)
- [Gitting, Building, and Installing Open-Xml-PowerTools](https://www.youtube.com/watch?v=60w-yPDSQD0)

Build Instructions
==================

**Prerequisites:**

- Visual Studio 2017 Update 5 or .NET CLI toolchain

**Build**
 
 With Visual Studio:

- Open `OpenXmlPowerTools.sln` in Visual Studio
- Rebuild the project
- Build the solution.  To validate the build, open the Test Explorer.  Click Run All.
- To run an example, set the example as the startup project, and press F5.

With .NET CLI toolchain:

- Run `dotnet build OpenXmlPowerTools.sln`

Change Log
==========

Version 4.3 : June 13, 2016
- New WmlComparer module

Version 4.2 : December 11, 2015
- New SmlDataRetriever module
- New SmlCellFormatter module

Version 4.1.3 : November 2, 2015
- DocumentAssembler: Fix bug associated with duplicate bookmarks.
- DocumentAssembler: Enable processing of content controls / metadata in footer rows.
- DocumentAssembler: Avoid processing content controls used for purposes other than the DocumentAssembler template, including page numbers in footers, etc.

Version 4.1.2 : October 31, 2015
- HtmlToWmlConverter: Handle unknown elements by recursively processing descendants

Version 4.1.1 : October 21, 2015
- Fix to AddTypes.ps1 to compile WmlToHtmlConverter.cs instead of HtmlConverter.cs
- Fix to MettricsGetter.ps1 to correctly report whether a document contains tracked revisions
- Added some unit tests for PresentationBuilder

Version 4.1.0 : September 27, 2015
- New HtmlToWmlConverter module
- HtmlConverter generates non breaking spaces as #00a0 unicode charater, not &nbsp; entity.

Version 4.0.0 : August 6, 2015
- New DocumentAssember module
- New SpreadsheetWriter module
- New Cmdlet: Complete-DocxTemplateFromXml
- Fix DocumentBuilder: deal with headers / footers more rationally
- Enhance DocumentBuilder: add option to discard headers / footers from section (but keep layout of section)
- Fix RevisionAccepter: deal with w:moveTo immediately before a table
- New test document library in the TestFiles directory
- XUnit tests
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

Procedures for enhancing Open-Xml-PowerTools
--------------------------------------------
There are a variety of things to do when adding a new CmdLet to Open-Xml-PowerTools:
- Write the new CmdLet.  Put it in the Cmdlets directory
- Modify Open-Xml-PowerTools.psm1
  - Call the new Cmdlet script to make the function available
  - Modify Export-ModuleMember function to export the Cmdlet and any aliases
- Update Readme.txt, describing the enhancement
- Add a new test to Test-OpenXmlPowerToolsCmdlets.ps1

Procedures for enhancing the core C# modules
- Modify the code
- Write xUnit tests
- Write an example if necessary
- Run xUnit tests
