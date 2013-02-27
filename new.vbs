Option Explicit
Dim objFSO,objAccounts,myAccounts,objRegions,myRegions,objYAR,myYAR,objFile,STDIn,STDOut,str,line
myAccounts=".\accounts.txt"
myYAR="..\_YAR\Settings\Bots.xml"
myRegions=".\regions.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set STDIn=WScript.STDIn
Set STDOut=WScript.STDOut




STDOut.WriteLine "Account Name: "
str = STDIn.ReadLine
Set objRegions = objFSO.OpenTextFile(myRegions,1)

Do
	line=objRegions.ReadLine
	WScript.Echo "Creating "+str+" "+line+" folder."
	objFSO.CopyFolder ".\Diablo III", str+" "+line
Loop Until objRegions.AtEndOfStream


'end
Set objFSO=Nothing
