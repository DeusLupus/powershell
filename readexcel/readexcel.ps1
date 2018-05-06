#declare path where excel is stored.
$file = "C:\Users\Deus_\Desktop\PowerShell Tutorials\powershell\readexcel\readcsvtest.xlsx"
$sheetname = "readcsvtest"

#create an instance and open excel
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.workbooks,Open($file)
$sheet = $workbook.Worksheets.Item($sheetname)
$objExcel.Visible=$false

#get max rows
$rowMax = ($sheet.UsedRange.Rows).count

#declare starting position for each column
$rowSerial,$colSerial = 1,1
$rowAsset,$colAsset = 1,2

#loop through each row and store each variable
for ($i=1; $i -le $rowMax - 1; $i++) {
    $serial = $sheet.Cells.Item($rowSerial + $i,$colSerial).text
    $asset = $sheet.Cells.Item($rowAsset + $i,$colAsset).text

    #use write host to check data
    Write-Host ("Serial Number: " + $serial)
    Write-Host ("Asset Tag: " + $asset)
}

#close excel or it will be locked for editing
$objExcel.quit()