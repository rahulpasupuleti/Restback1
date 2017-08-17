Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("E:\Bala\Rahul\xls\PlotJavaResultV6.xlsm",0,True)

objExcel.Application.Visible = False
'objExcel.Workbooks.Add
'objExcel.Cells(1, 1).Value = "Test value"

objExcel.Application.Run "uploadFile"
objExcel.Application.Run "PrintPlotResult"
objExcel.ActiveWorkbook.Close


objExcel.Application.Quit
'WScript.Echo "Finished."
WScript.Quit


