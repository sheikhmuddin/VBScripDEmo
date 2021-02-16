'Expert Youtube channel
Dim objFso,Objtextreader,MyData
Set objFso=CreateObject("Scripting.FileSystemObject")

'Read the Text File
Set Objtextreader=objFso.OpenTextFile("D:\vbscript\Data_File\test.txt")
'Read a Line
'MyData=Objtextreader.ReadLine
MyData=Objtextreader.ReadALL
MsgBox MyData

'Release the Object
Set objFso=Nothing
