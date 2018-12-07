#prep file for sendkeys
$wshell = New-Object -ComObject wscript.shell
$wshell.AppActivate('EDGE')

#declare path where excel is stored.
$file = "C:\Users\dbissessar\Desktop\Basic Receiving\Basic Receive.xlsx"
$sheetname = "Sheet1"

#create an instance and open excel
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetname)
$objExcel.Visible=$false

#get max rows
$rowMax = ($sheet.UsedRange.Rows).count

#declare starting position for each column
$rowItem,$colItem = 1,1
$rowQTY,$colQTY = 1,2
$rowReason, $colReason = 1,3
$rowLocation, $colLocation = 1,4
$rowSO, $colSO = 1,5

#loop through each row and store each variable, add to clipboard, paste and wait a second
for ($i=1; $i -le $rowMax - 1; $i++) {
    $item = $sheet.Cells.Item($rowItem + $i,$colItem).text
    $qty = $sheet.Cells.Item($rowQTY + $i,$colQTY).text
    $reason = $sheet.Cells.Item($rowReason + $i,$colReason).text
    $location = $sheet.Cells.Item($rowLocation + $i,$colLocation).text
    $SO = $sheet.Cells.Item($rowSO + $i,$colSO).text

    #send to external
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait($item)
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait($qty)
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait($reason)
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait($location)
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait($SO)
    Sleep 3
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    Sleep 1
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait('~')
      
    
    #use write host to check data, eventually replace with copy and paste
    #Write-Host ("Serial Number: " + $serial)
    #Write-Host ("Asset Tag: " + $asset)
}

#close excel or it will be locked for editing
$objExcel.quit()
