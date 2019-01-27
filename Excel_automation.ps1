#Creting a Excel object
$Excel = New-Object -ComObject Excel.application
#making the excel window visible
$Excel.visible = $true
$Excel.DisplayAlerts = $false
#Adding workbooks
$Excel.Workbooks.add()
$Excel.Workbooks.add()
$Excel.Workbooks.add()
#To view all the workbook's name
$Excel.Workbooks | select -Property name
#To activate a workbook
$Excel.Workbooks.Item(1).activate()
$Excel.Workbooks.Item("Book5").activate()
#To activate random workbook in a Excel file
$Excel.Workbooks.Item((Get-Random -Minimum 1 -Maximum($Excel.Workbooks.Count+1))).activate()
#To open an existing workbook
$Excel.Workbooks.open("C:\Users\Suraj\Desktop\Excel Automation\demo.xlsx")
#To close a workbook
$Excel.Workbooks.Item(3).Close()
$Excel.Workbooks.Item("demo").Close()
#To save a workbook
$Excel.Workbooks.Item(1).saveas("C:\Users\Suraj\Desktop\Excel Automation\aaaa.xlsx")

#####################################
##########Working with workboooks####
#####################################

#To add a worksheet
$Excel.Worksheets.Add()
#To see the worksheets
$Excel.Worksheets | select -ExpandProperty name
#To change the name of the worksheet in a workbook
$Excel.Worksheets.Item(2).name = "Test_2"
$Excel.Worksheets.Item("Sheet2").name = "Test_2"
#To activate a worksheet in a workbook
$Excel.Worksheets.Item("Test_2").activate()
#To activate random worksheet
$Excel.Worksheets.Item((Get-Random -Minimum 1 -Maximum($Excel.Workbooks.Count+1))).activate()
#To delete a worksheet
$Excel.Worksheets.Item(1).delete()
$Excel.Worksheets.Item("Sheet1").delete()

##############################################
###############Cleaning up the Excel object###
##############################################
$Excel.Workbooks.Close()
$Excel.Quit()

#Checking the excel process still running in the background even after quiting
Get-Process excel | Stop-Process -Force

#For garbage collection
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)