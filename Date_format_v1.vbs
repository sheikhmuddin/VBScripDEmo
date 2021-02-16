Option Explicit
Dim dt1,dt2,dt3
dt1=CDate("6/05/2020")
dt2=CDate("06/05/2020")
dt3=CDate("6/5/2020")

MsgBox dt1&vbCrLf&dt2
If dt1=dt2 Then
MsgBox"True"
Else
MsgBox"False"
End If

MsgBox now