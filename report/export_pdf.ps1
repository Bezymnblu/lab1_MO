# Updates the table of contents and exports PDF via Microsoft Word (COM).
# Run from lab1:  powershell -ExecutionPolicy Bypass -File report\export_pdf.ps1
$docx = (Get-ChildItem (Join-Path $PSScriptRoot "..") -Filter "*7.docx" | Where-Object { $_.Name -notlike "~*" } | Select-Object -First 1).FullName
$pdf = [System.IO.Path]::ChangeExtension($docx, ".pdf")
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docx)
    $doc.TablesOfContents(1).Update()
    $doc.Fields.Update() | Out-Null
    $doc.TablesOfContents(1).Update()
    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Output ("Pages: " + $doc.ComputeStatistics(2))
    $doc.Close()
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Output "PDF: $pdf"
