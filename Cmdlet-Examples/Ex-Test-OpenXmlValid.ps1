[environment]::CurrentDirectory = $(Get-Location)

function TestOpenXmlValidExample {
    param (
        [int]$exampleNum
    )
    switch ($exampleNum) {
        1 {
            Test-OpenXmlValid Valid.docx -OfficeVersion 2010 -Verbose
            Test-OpenXmlValid Invalid.docx -OfficeVersion 2007 -Verbose
        }
        2 {
            Get-ChildItem *.xlsx | Test-OpenXmlValid -Verbose
        }
        3 {
            Test-OpenXmlValid *.docx -Verbose -OfficeVersion 2010
        }
    }
}

for ($e = 1; $e -le 3; $e += 1)
{
    TestOpenXmlValidExample($e)
}
