# OpenXmlPowerTools

[![.NET build and test](https://github.com/Codeuctivity/OpenXmlPowerTools/actions/workflows/dotnet.yml/badge.svg)](https://github.com/Codeuctivity/OpenXmlPowerTools/actions/workflows/dotnet.yml) [![Nuget](https://img.shields.io/nuget/v/Codeuctivity.OpenXmlPowerTools.svg)](https://www.nuget.org/packages/Codeuctivity.OpenXmlPowerTools/) [![Codacy Badge](https://app.codacy.com/project/badge/Grade/91883269775a4333aa78d8a911dfbaf5)](https://www.codacy.com/gh/Codeuctivity/OpenXmlPowerTools/dashboard?utm_source=github.com&utm_medium=referral&utm_content=Codeuctivity/OpenXmlPowerTools&utm_campaign=Badge_Grade)

## Features

- Working on Linux an Windows
- Conversion of DOCX to HTML/CSS.
- Splitting DOCX/PPTX files into multiple files.
- Combining multiple DOCX/PPTX files into a single file.
- Populating content in template DOCX files with data from XML.
- Conversion of HTML/CSS to DOCX.
- Searching and replacing content in DOCX/PPTX using regular expressions.
- Managing tracked-revisions, including detecting tracked revisions, and accepting tracked revisions.
- Updating Charts in DOCX/PPTX files, including updating cached data, as well as the embedded XLSX.
- Comparing two DOCX files, producing a DOCX with revision tracking markup, and enabling retrieving a list of revisions.
- Retrieving metrics from DOCX files, including the hierarchy of styles used, the languages used, and the fonts used.
- Writing XLSX files using far simpler code than directly writing the markup, including a streaming approach that
  enables writing XLSX files with millions of rows.
- Extracting data (along with formatting) from spreadsheets.

## Development

- Run `dotnet build OpenXmlPowerTools.sln`
