Option Explicit
Dim objFSO,objFile,myFile,line,Result
myFile=".\accounts.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(myFile,1)
Do
	line=objFile.ReadLine
	WScript.Echo line
Loop Until objFile.AtEndOfStream
Set objFSO=Nothing
Set objFile=Nothing