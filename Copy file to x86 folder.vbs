On Error Resume Next
dim filesys, newfolder,  fso
 

Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell") 


 

set filesys=CreateObject("Scripting.FileSystemObject")
newfolder = filesys.CreateFolder ("C:\Program Files (x86)\folder1\folder2\")
newfolder = filesys.CreateFolder ("C:\ProgramData\folder1\")

 

fso.CopyFile "C:\users\administrator\desktop\*.txt", "C:\ProgramData\folder1\"
fso.CopyFile "C:\users\administrator\desktop\*.txt", "C:\Program Files (x86)\folder1\folder2\"


fso.CopyFile "C:\users\administrator\desktop\*.txt", "C:\ProgramData\folder1\"


WScript.Quit