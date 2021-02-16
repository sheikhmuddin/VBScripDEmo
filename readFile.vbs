strPath="D:\vbscript\testRead.xlsx"
SetxlAPP=CreateObject("Excel.Application")
setxWb=xlApp.Workbooks.Open(strPath)
xlApp.DisplayAlerts=false
Set sht1=zlWb.Sheets(1)
print sht1.UsedRange.Rows.Count
Printsht1.cells(1,1)

Set sht1 = nothing
Set xlWb=nothing
Set xlApp=nothing