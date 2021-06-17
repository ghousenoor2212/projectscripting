On Error Resume Next

 Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell") 
dim filesys, newfolder,  fso

 

set filesys=CreateObject("Scripting.FileSystemObject")
newfolder = filesys.CreateFolder ("C:\Program Files (x86)\Pegasystems\Deployment\")

 

fso.CopyFile "CommonConfig.xml", "C:\ProgramData\Pegasystems\"
fso.CopyFile "RuntimeConfig.xml", "C:\Program Files (x86)\Pegasystems\Pega Robot Runtime\"
fso.CopyFile "Welbit_RDA_Solution_Release.manifest", "C:\Program Files (x86)\Pegasystems\Deployment\"
fso.CopyFile "Welbit_RDA_Solution_Release.OpenSpan", "C:\Program Files (x86)\Pegasystems\Deployment\"
