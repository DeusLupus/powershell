#declare path where excel is stored.
$file =
$sheetname = "sheet1"

$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$objExcel.Visible=$false

$rowMax = ($sheet.UsedRange.Rows).count

$rowName.$colName = 1,1
$rowAge.$colAge = 1,2
$rowCity,$colCity = 1,3

