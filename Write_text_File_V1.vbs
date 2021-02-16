'Expert Youtube channel
Dim objFso,Objtextwriter,MyData
Set objFso=CreateObject("Scripting.FileSystemObject")

'Write the Text File
Set Objtextwriter=objFso.CreateTextFile("D:\vbscript\Data_File\test3.txt",true)
Objtextwriter.WriteLine("This Line")
Objtextwriter.Close
 Set Objtextwriter= Nothing
 
 
'Release the Object
Set objFso=Nothing
