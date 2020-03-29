# Open-XML-PowerTools

[![Build status](https://ci.appveyor.com/api/projects/status/2wq6a4b0q6akvy4n?svg=true)](https://ci.appveyor.com/project/stesee/open-xml-powertools) [![Nuget](https://img.shields.io/nuget/v/OpenXmlPowerToolsStandard.svg)](https://www.nuget.org/packages/OpenXmlPowerToolsStandard/)

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

## Open-Xml-PowerTools Content

There is a lot of content about Open-Xml-PowerTools at the [Open-Xml-PowerTools Resource Center at OpenXmlDeveloper.org](http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx)

See:

- [DocumentBuilder Resource Center](http://www.ericwhite.com/blog/documentbuilder-developer-center/)
- [PresentationBuilder Resource Center](http://www.ericwhite.com/blog/presentationbuilder-developer-center/)
- [WmlToHtmlConverter Resource Center](http://www.ericwhite.com/blog/wmltohtmlconverter-developer-center/)
- [DocumentAssembler Resource Center](http://www.ericwhite.com/blog/documentassembler-developer-center/)

## Build Instructions

### Prerequisites

- Visual Studio 2019 or .NET CLI toolchain

### Build

 With Visual Studio:

- Open `OpenXmlPowerTools.sln` in Visual Studio
- Rebuild the project
- Build the solution.  To validate the build, open the Test Explorer.  Click Run All.
- To run an example, set the example as the startup project, and press F5.

With .NET CLI toolchain:

- Run `dotnet build OpenXmlPowerTools.sln`
