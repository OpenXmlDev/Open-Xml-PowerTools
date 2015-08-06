Remove-Item .\OpenXmlPowerTools\bin -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools\obj -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools.Net35\bin -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools.Net35\obj -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools.Tests\bin -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools.Tests\obj -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools.Tests.OA\bin -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerTools.Tests.OA\obj -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerToolsExamples\*\bin -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item .\OpenXmlPowerToolsExamples\*\obj -Recurse -Force -ErrorAction SilentlyContinue
dir *.suo -Recurse -Force | Remove-Item -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\ChartUpdater01\Updated-Chart*'
Remove-Item $d -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\DocumentAssembler\AssembledDoc*'
Remove-Item $d -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\DocumentAssembler01\AssembledDoc*'
Remove-Item $d -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\FormattingAssembler01\*out.docx'
Remove-Item $d -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\HtmlConverter01\*.html'
Remove-Item $d -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\HtmlConverter01\*_files'
Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue
$d = '.\OpenXmlPowerToolsExamples\TextReplacer01\*out*'
Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item OpenXmlPowerTools.dll -ErrorAction SilentlyContinue
