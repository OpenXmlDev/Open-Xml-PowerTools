[environment]::CurrentDirectory = $(Get-Location)

function GetOpenXmlValidationErrorsExample {
    param (
        [int]$exampleNum
    )
    switch ($exampleNum) {
        1 {
            Get-OpenXmlValidationErrors Valid.docx -OfficeVersion 2010 -Verbose
            Get-OpenXmlValidationErrors Invalid.docx -OfficeVersion 2007 -Verbose
        }
        2 {
            Get-ChildItem *.xlsx | Get-OpenXmlValidationErrors -Verbose
        }
        3 {
            Get-OpenXmlValidationErrors *.docx -Verbose -OfficeVersion 2010
        }
    }
}

for ($e = 1; $e -le 3; $e += 1)
{
    GetOpenXmlValidationErrorsExample($e)
}
