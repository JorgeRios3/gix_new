dim sh,  fs 
set sh = WScript.CreateObject("WScript.Shell")
set fs = WScript.CreateObject("Scripting.FileSystemObject")

if fs.FileExists("C:\Archivos de programa\gixr\gix.pyc") then
	fs.DeleteFile "C:\Archivos de programa\gixr\*.pyc",True
end if
sh.run """C:\Archivos de programa\Subversion\bin\svn.exe"" update --username gix --password gixpythont --no-auth-cache ""C:\Archivos de programa\gixr""",,True
fs.DeleteFile "C:\Archivos de programa\gixr\*.py",True

fs.DeleteFile "C:\Archivos de programa\gixr\force*",True
fs.CopyFile "C:\smartics\forcer\force*", "C:\Archivos de programa\gixr\", true
sh.run "c:\python25\pythonw.exe ""C:\Archivos de programa\gixr\gix.pyc""" 

