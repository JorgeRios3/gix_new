dim sh,  fs 
set sh = WScript.CreateObject("WScript.Shell")
set fs = WScript.CreateObject("Scripting.FileSystemObject")

if fs.FileExists("C:\Archivos de programa\gix\gix.pyc") then
	fs.DeleteFile "C:\Archivos de programa\gix\*.pyc",True
end if
sh.run """C:\Archivos de programa\Subversion\bin\svn.exe"" update --username gix --password gixpythont --no-auth-cache ""C:\Archivos de programa\gix""",,True
fs.DeleteFile "C:\Archivos de programa\gix\*.py",True

fs.DeleteFile "C:\Archivos de programa\gix\force*",True
fs.CopyFile "C:\smartics\force\force*", "C:\Archivos de programa\gix\", true
sh.run "c:\python25\pythonw.exe ""C:\Archivos de programa\gix\gix.pyc""" 

