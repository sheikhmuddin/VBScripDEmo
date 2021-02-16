'MP Expert 
'Create excl Workbooks
'Set objxl=CreateObject("Excel.Application")
'objxl.Visible=true 'TO show flash for creating file
'objxl.Workbooks.Add
'objxl.ActiveWorkbook.SaveAs"D:\vbscript\Mp.xlsx"
'objxl.Quit
'Set objxl=Nothing

'Write into the Workbooks

'Set objxl=CreateObject("Excel.Application")
'objxl.Visible=true 'TO show flash for creating file
'objxl.Workbooks.Open "D:\vbscript\Mp.xlsx"
'objxl.Worksheets(1).Cells(1,1)="SL NO"
'objxl.Worksheets(1).Cells(1,2)="Description"
'objxl.ActiveWorkbook.Save 
'objxl.Quit
'Set objxl=Nothing

'Read from the Workbooks
Set objxl=CreateObject("Excel.Application")
objxl.Visible=true 'TO show flash for creating file
objxl.Workbooks.Open "D:\vbscript\Mp.xlsx"

Colcnt=objxl.ActiveWorkbook.Sheets(1).UsedRange.Columns.count
'MsgBox RWcnt
For i=1 To Colcnt
RWcnt=objxl.ActiveWorkbook.Sheets(1).UsedRange.Rows.count
For j =1 To RWcnt
'val=objxl.ActiveWorkbook.Sheets(1).Range("A1").Value
val=objxl.ActiveWorkbook.Sheets(1).Cells(j,i)
MsgBox val
Next
Next
 
objxl.ActiveWorkbook.Save 
objxl.Quit
Set objxl=Nothing