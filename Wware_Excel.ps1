$filename = "c:\users\admin\desktop\SecurityCentralSupportedProducts.xlsx"
$objExcel = New-Object -ComObject Excel.Application
#$objExcel.DisplayAlerts = $False
$objWrkBk  = $objExcel.Workbooks.Open($Filename)
$objExcel.Visible = $True
# one way of selecting a sheet e.g. 3rd worksheet
# $objWrkBk.Sheets(3).Select()
# another way of selecting a sheet e.g. 1st worksheet
# $objWrkBk.Worksheets.item(1).Select()
$objWrkSht1 = $objWrkBk.Worksheets.item(1)
$objWrkSht1.Select()
# filter on 7th column, with string matching 'ABC'
$objWrkSht1.Range('$A$nn:$P$mm').AutoFilter(3,'=Supported')
$objWrkSht1.Range('$A$nn:$P$mm').AutoFilter(1,'>=$thirtyAgo')
# $A$nn => range first cell position, $P$mm = range last cell position