# OpenXmlPowerTools

[![.NET build and test](https://github.com/Codeuctivity/OpenXmlPowerTools/actions/workflows/dotnet.yml/badge.svg)](https://github.com/Codeuctivity/OpenXmlPowerTools/actions/workflows/dotnet.yml) [![Nuget](https://img.shields.io/nuget/v/Codeuctivity.OpenXmlPowerTools.svg)](https://www.nuget.org/packages/Codeuctivity.OpenXmlPowerTools/) [![Codacy Badge](https://app.codacy.com/project/badge/Grade/91883269775a4333aa78d8a911dfbaf5)](https://www.codacy.com/gh/Codeuctivity/OpenXmlPowerTools/dashboard?utm_source=github.com&utm_medium=referral&utm_content=Codeuctivity/OpenXmlPowerTools&utm_campaign=Badge_Grade)

This is a fork of https://www.nuget.org/packages/OpenXmlPowerTools/

## Focus of this fork

- Linux and Windows support
- Conversion of DOCX to HTML/CSS.

## Example - Convert DOCX to HTML

``` csharp
var sourceDocxFileContent = File.ReadAllBytes("./source.docx");
using var memoryStream = new MemoryStream();
await memoryStream.WriteAsync(sourceDocxFileContent, 0, sourceDocxFileContent.Length);
using var wordProcessingDocument = WordprocessingDocument.Open(memoryStream, true);
var settings = new WmlToHtmlConverterSettings("htmlPageTile");
var html = WmlToHtmlConverter.ConvertToHtml(wordProcessingDocument, settings);
var htmlString = html.ToString(SaveOptions.DisableFormatting);
File.WriteAllText("./target.html", htmlString, Encoding.UTF8);
```