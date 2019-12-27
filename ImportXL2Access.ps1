# Constants
$acImport = 0
# acSpreadsheetTypeExcel = 9 (Excel 2010), 10 : Microsoft Excel 2010/2013/2016 XML format (.xlsx, .xlsm, .xlsb)
$acSpreadsheetTypeExcel = 10
$acHasFieldNames = $True

# Variables
$accessDBName = 'Report.accdb'
$xlFileName = 'Data to import.xlsx'

# Access DB
$accessDBFile = Join-Path -Path $(get-location) -ChildPath $accessDBName
#$strDBConn = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=$accessDBFile"

if (Test-Path $accessDBFile) {
    Remove-Item $accessDBFile
    Write-Host 'Deleted DB:'$accessDBName
}

# Create a new Access DB
$access = New-Object -ComObject Access.Application
$access.NewCurrentDataBase($accessDBFile)

Write-Host 'Created DB:'$accessDBName


# Import data from Excel
$targetFile = Join-Path -Path $(get-location) -ChildPath $xlFileName
$xl = New-Object -ComObject Excel.Application
$xl.visible = $false
# $xl.DisplayAlerts = $false
$xlWorkBook = $xl.Workbooks.Open($targetFile)
$xlWorkSheet = $xlWorkBook.WorkSheets

Write-Host 'Number of WorkSheets:'$xlWorkSheet.Count

foreach ($ws in $xlWorkbook.Worksheets) { 
    Write-Host 'Importing WorkSheet:'$ws.Name 
    $TableName = $ws.Name
    $Range = $ws.Name + "!" 
    
    # $access.OpenCurrentDataBase($accessDBFile)
    $access.DoCmd.TransferSpreadSheet($acImport, $acSpreadsheetTypeExcel, $TableName, $targetFile, $acHasFieldNames, $Range)
}

# close objects
$xl.Quit()
Stop-Process -processname EXCEL

$access.CloseCurrentDataBase()
$access.Quit()