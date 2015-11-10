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

$assemblies = (
  "$PSScriptRoot\..\Open-Xml-Sdk\DocumentFormat.OpenXml\bin\debug\DocumentFormat.OpenXml.dll",
  "$PSScriptRoot\..\Open-Xml-Sdk\DocumentFormat.OpenXml\bin\debug\System.IO.Packaging.dll",
  "System.Xml.Linq, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a",
  "System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
  "System.IO.Compression, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
)

$sources = @(
    "$PSScriptRoot\OpenXmlPowerTools\DocumentAssembler.cs",
    "$PSScriptRoot\OpenXmlPowerTools\DocumentBuilder.cs",
    "$PSScriptRoot\OpenXmlPowerTools\ExcelFormula.cs",
    "$PSScriptRoot\OpenXmlPowerTools\FieldRetriever.cs",
    "$PSScriptRoot\OpenXmlPowerTools\FormattingAssembler.cs",
    "$PSScriptRoot\OpenXmlPowerTools\GetListItemText_Default.cs",
    "$PSScriptRoot\OpenXmlPowerTools\GetListItemText_fr_FR.cs",
    "$PSScriptRoot\OpenXmlPowerTools\GetListItemText_ru_RU.cs",
    "$PSScriptRoot\OpenXmlPowerTools\GetListItemText_sv_SE.cs",
    "$PSScriptRoot\OpenXmlPowerTools\GetListItemText_tr_TR.cs",
    "$PSScriptRoot\OpenXmlPowerTools\GetListItemText_zh_CN.cs",
    "$PSScriptRoot\OpenXmlPowerTools\WmlToHtmlConverter.cs",
    "$PSScriptRoot\OpenXmlPowerTools\ListItemRetriever.cs",
    "$PSScriptRoot\OpenXmlPowerTools\MarkupSimplifier.cs",
    "$PSScriptRoot\OpenXmlPowerTools\MetricsGetter.cs",
    "$PSScriptRoot\OpenXmlPowerTools\OpenXmlRegex.cs",
    "$PSScriptRoot\OpenXmlPowerTools\OxPtHelpers.cs",
    "$PSScriptRoot\OpenXmlPowerTools\PegBase.cs",
    "$PSScriptRoot\OpenXmlPowerTools\PresentationBuilder.cs",
    "$PSScriptRoot\OpenXmlPowerTools\PtOpenXmlDocument.cs",
    "$PSScriptRoot\OpenXmlPowerTools\PtOpenXmlUtil.cs",
    "$PSScriptRoot\OpenXmlPowerTools\PtUtil.cs",
    "$PSScriptRoot\OpenXmlPowerTools\ReferenceAdder.cs",
    "$PSScriptRoot\OpenXmlPowerTools\RevisionAccepter.cs",
    "$PSScriptRoot\OpenXmlPowerTools\SSFormula.cs",
    "$PSScriptRoot\OpenXmlPowerTools\SpreadsheetWriter.cs",
    "$PSScriptRoot\OpenXmlPowerTools\TextReplacer.cs",
    "$PSScriptRoot\OpenXmlPowerTools\WmlDocument.cs",
    "$PSScriptRoot\OpenXmlPowerTools\WorksheetAccessor.cs",
    "$PSScriptRoot\OpenXmlPowerTools\XlsxTables.cs",
    "$PSScriptRoot\OpenXmlPowerTools\OpenXmlCmdlet.cs",
    "$PSScriptRoot\OpenXmlPowerTools\SpreadsheetDocumentManager.cs",
    "$PSScriptRoot\OpenXmlPowerTools\OutXlsxCmdlet.cs",
    "$PSScriptRoot\OpenXmlPowerTools\ScalarTypes.cs"
)

Add-Type -Path "$PSScriptRoot\..\Open-Xml-Sdk\DocumentFormat.OpenXml\bin\debug\DocumentFormat.OpenXml.dll"
Add-Type -Path "$PSScriptRoot\..\Open-Xml-Sdk\DocumentFormat.OpenXml\bin\debug\System.IO.Packaging.dll"
Add-Type -ReferencedAssemblies $assemblies -Path $sources -OutputAssembly "$PSScriptRoot\OpenXmlPowerTools.dll"
