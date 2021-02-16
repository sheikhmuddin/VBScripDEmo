'Expert Youtube channel
Dim objFso
Set objFso=CreateObject("Scripting.FileSystemObject")

'Use that object
If objFso.FolderExists("D:\vbscript\Data_File")Then
MsgBox"Existing"
objFso.CreateTextFile("D:\vbscript\Data_File\test.txt")
objFso.CreateTextFile("D:\vbscript\Data_File\test.xlsx")
objFso.CreateTextFile("D:\vbscript\Data_File\test.docx")
Else
objFso.CreateFolder("D:\vbscript\Data_File")
End if
'Release the Object
Set objFso=Nothing
