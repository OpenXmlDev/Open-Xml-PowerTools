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
Remove-Item OpenXmlPowerTools.dll -ErrorAction SilentlyContinue
