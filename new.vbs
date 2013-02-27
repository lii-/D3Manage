Option Explicit
Dim objFSO,objAccounts,myAccounts,objYAR,myYAR,STDIn,STDOut,str
myAccounts=".\accounts.txt"
myYAR="..\_YAR\Settings\Bots.xml"
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set STDIn=WScript.STDIn
Set STDOut=WScript.STDOut
STDOut.WriteLine "Account Name: "
str = STDIn.ReadLine
WScript.Echo str

'end
Set objFSO=Nothing



Function getTextFile(path)
	Set objFile = objFSO.OpenTextFile(path,1)
	getTextFile=objFile
	Set objFile=Nothing
End Function

Function ReadByLine(obj)
	Do
		line=obj.ReadLine
		WScript.Echo line
	Loop Until obj.AtEndOfStream
End Function