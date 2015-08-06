[environment]::CurrentDirectory = $(Get-Location)

function MergeDocxExample {
    param (
        [int]$exampleNum
    )
    switch ($exampleNum) {
        1 {
            $doc1 = New-Object OpenXmlPowerTools.Source("Input1.docx")
            $doc2 = New-Object OpenXmlPowerTools.Source("Input2.docx", $True)
            $sources = ($doc1, $doc2)
            Merge-Docx -OutputPath Out-Merge01.docx -Sources $sources -Verbose
        }
        2 {
            $doc1 = New-Object OpenXmlPowerTools.Source("Input1.docx", 1, 2, $false)
            $doc2 = New-Object OpenXmlPowerTools.Source("Input2.docx", 1, 2, $false)
            $sources = ($doc1, $doc2)
            Merge-Docx -OutputPath Out-Merge02.docx -Sources $sources -Verbose
        }
        3 {
            $doc1 = New-Object OpenXmlPowerTools.Source("Input1.docx")
            $doc2 = New-Object OpenXmlPowerTools.Source("Input2.docx", 1, 2, $false)
            $newWml = Merge-Docx -Sources ($doc1, $doc2) -Verbose
            $newWml.SaveAs("Out-Merge03.docx");
        }
    }
}

for ($e = 1; $e -le 3; $e += 1)
{
    MergeDocxExample($e)
}
