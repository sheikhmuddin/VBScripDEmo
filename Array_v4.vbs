
Dim Names(5)
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

str =""
Dim Names2
Names2 = Array("Sheik","Mohd","Ashraf","Uddin","Riyang","Sheikh Uddin")
for i=0 To 5
str=str&Names2(i)&vbNewLine
Next
MsgBox str