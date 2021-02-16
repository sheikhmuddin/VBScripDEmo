strPath="D:\vbscript\testRead.xlsx"
Set xlAPP=CreateObject("Excel.Application")
xlApp.DisplayAlerts=False
set xlWb=xlApp.Workbooks.Open(strPath)
 
Set sht1=xlWb.Sheets(1)
intRowCount=sht1.UsedRange.Rows.Count
intCoulmnCount=sht1.UsedRange.Columns.Count
'print sht1.UsedRange.Rows.Counts.Count
MsgBox sht1.cells(1,1)
MsgBox intRowCount
MsgBox intCoulmnCount
 
Set sht1 = nothing
Set xlWb=nothing
Set xlApp=nothing