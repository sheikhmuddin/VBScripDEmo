
Dim str
Dim Names()
ReDim Names(2,5)
Names(0,0)="Sheikh"
Names(0,1)="Mohd"
Names(0,2)="Ashraf"
Names(0,3)="Uddin"
Names(0,4)="Riyang"
Names(0,5)="Sheikh Uddin"

Names(1,0)="Sheikh1"
Names(1,1)="Mohd1"
Names(1,2)="Ashraf1"
Names(1,3)="Uddin1"
Names(1,4)="Riyang1"
Names(1,5)="Sheikh Uddin1"

Names(2,0)="Sheikh2"
Names(2,1)="Mohd2"
Names(2,2)="Ashraf2"
Names(2,3)="Uddin2"
Names(2,4)="Riyang2"
Names(2,5)="Sheikh Uddin2"


str =""
For i =0 to 2
For j =0 to 5
str=str&Names(i,j)&vbNewLine
Next
Next
MsgBox str

