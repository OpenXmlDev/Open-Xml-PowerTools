# Open-XML-PowerTools

[![Build status](https://ci.appveyor.com/api/projects/status/2wq6a4b0q6akvy4n/branch/master?svg=true)](https://ci.appveyor.com/project/stesee/open-xml-powertools/branch/master)
 [![Nuget](https://img.shields.io/nuget/v/OpenXmlPowerToolsStandard.svg)](https://www.nuget.org/packages/OpenXmlPowerToolsStandard/) [![Codacy Badge](https://api.codacy.com/project/badge/Grade/73ab9db4912449f28d961e3ddad189b1)](https://app.codacy.com/gh/Codeuctivity/Open-Xml-PowerTools?utm_source=github.com&utm_medium=referral&utm_content=Codeuctivity/Open-Xml-PowerTools&utm_campaign=Badge_Grade_Dashboard)

The Open XML PowerTools provides guidance and example code for programming with Open XML
Documents (DOCX, XLSX, and PPTX). It is based on, and extends the functionality
of the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

It supports scenarios such as:

- Splitting DOCX/PPTX files into multiple files.
- Combining multiple DOCX/PPTX files into a single file.
- Populating content in template DOCX files with data from XML.
- Conversion of DOCX to HTML/CSS.
- Conversion of HTML/CSS to DOCX.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Comparing two DOCX files, producing a DOCX with revision tracking markup, and enabling retrieving a list of revisions.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.
- Extracting data (along with formatting) from spreadsheets.

## Build Instructions

### Prerequisites

- Visual Studio 2019 or .NET CLI toolchain

### Build

With Visual Studio:

- Open `OpenXmlPowerTools.sln` in Visual Studio
- Rebuild the project
- Build the solution. To validate the build, open the Test Explorer. Click Run All.
- To run an example, set the example as the startup project, and press F5.

With .NET CLI toolchain:

- Run `dotnet build OpenXmlPowerTools.sln`
