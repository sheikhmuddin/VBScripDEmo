'Expert Youtube channel
Dim objFso,Objtextreader,MyData
Set objFso=CreateObject("Scripting.FileSystemObject")

'Read the Text File
Set Objtextreader=objFso.OpenTextFile("D:\vbscript\Data_File\test.txt")
'Read a Line
Do While Objtextreader.AtEndOfLine=False
MyData=Objtextreader.ReadLine
 
MsgBox MyData

Loop
'Release the Object
Set objFso=Nothing
