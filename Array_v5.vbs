
Dim Names()
ReDim Names(5)
Names(0)="Sheikh"
Names(1)="Mohd"
Names(2)="Ashraf"
Names(3)="Uddin"
Names(4)="Riyang"
Names(5)="Sheikh Uddin"

dim str
str =""
For i =0 to 5
str=str&Names(i)&vbNewLine
Next
MsgBox str
ReDim preserve Names(6)
Names(6)="SM"
str =""
For i =0 to 6
str=str&Names(i)&vbNewLine
Next
MsgBox str
