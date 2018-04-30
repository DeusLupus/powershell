#declare path where excel is stored.
$file = "C:\Users\Deus_\Desktop\PowerShell Tutorials\powershell\readexcel\readcsvtest.xlsx"
$sheetname = "readcsvtest"

#create xl object
$xl = New-Object -ComObject Excel.Application

#disable the visible property
$xl.visible = $false

#open excel file data in $wb variable
$wb = $objExcel.Workbooks.Open($file)

#select correct worksheet
$ws = $wb.sheets.item("readcsvtest")

