Option Explicit
Dim WshShell, sourceWritten, strPath, strCommand, objFSO,objAccounts,myAccounts,objRegions,myRegions,objYARSource,myYARSource,myYARTarget,objYARTarget,objFile,STDIn,STDOut,account,password,line,line1,line2
dim objArgs: Set objArgs = WScript.Arguments

If LCase(Right(Wscript.FullName, 11)) <> "cscript.exe" Then
    strPath = Wscript.ScriptFullName
    strCommand = "%comspec% /k cscript  " & Chr(34) & strPath & chr(34)
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run(strCommand)
    Wscript.Quit
End If

myAccounts="accounts.txt"
myYARSource="YAR\Settings\Bots"
myYARTarget="YAR\Settings\_Bots"
myRegions="regions.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set STDIn=WScript.STDIn
Set STDOut=WScript.STDOut
Call GetInputs()
Call CheckExisting()

If Not objFSO.FileExists(myYARTarget & ".xml") Then
	Set objYARTarget = objFSO.CreateTextFile(myYARTarget & ".xml")
Else
	Set objYARTarget = objFSO.OpenTextFile(myYARTarget & ".xml",2)
End if
Set objRegions = objFSO.OpenTextFile(myRegions,1)
Set objYARSource = objFSO.OpenTextFile(myYARSource & ".xml",1)

Do Until InStr(line,"</ArrayOfBotClass>") > 0
	line=objYARSource.ReadLine
	if not InStr(line,"</ArrayOfBotClass>") > 0 then objYARTarget.WriteLine line
Loop
objYARSource.Close

Do
	Set objYARSource = objFSO.OpenTextFile(myYARSource & ".xml",1)
	Call SkipTo(objYARSource,"stephen.poole@hotmail.com")
	line1 = objYARSource.ReadLine

	line=objRegions.ReadLine
	WScript.Echo "Creating '"+account+" "+line+"' folder."
	objFSO.CopyFolder "Diablo III", account+" "+line
	
	WScript.Echo "Writing to Bots.xml"
	objYARTarget.WriteLine "<BotClass>"
		objYARTarget.WriteLine "<Name>"&account&"</Name>"
		objYARTarget.WriteLine "<Description>"&line&"</Description>"
		objYARTarget.WriteLine "<IsEnabled>"+SkipToAndGetXMLValue(objYARSource,"IsEnabled")+"</IsEnabled>"
		objYARTarget.WriteLine "<Demonbuddy>"
		  objYARTarget.WriteLine "<BuddyAuthUsername />"
		  objYARTarget.WriteLine "<BuddyAuthPassword />"
		  objYARTarget.WriteLine "<Location>e:\Applications\Diablo III\"&account&" "&line&"\Demonbuddy.exe</Location>"
		  objYARTarget.WriteLine "<Key>"+SkipToAndGetXMLValue(objYARSource,"Key")+"</Key>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<CombatRoutine>"+GetXMLValue(line1)+"</CombatRoutine>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<NoFlash>"+GetXMLValue(line1)+"</NoFlash>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<AutoUpdate>"+GetXMLValue(line1)+"</AutoUpdate>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<NoUpdate>"+GetXMLValue(line1)+"</NoUpdate>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<Priority>"+GetXMLValue(line1)+"</Priority>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<CpuCount>"+GetXMLValue(line1)+"</CpuCount>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<ProcessorAffinity>"+GetXMLValue(line1)+"</ProcessorAffinity>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<ManualPosSize>"+GetXMLValue(line1)+"</ManualPosSize>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<X>"+GetXMLValue(line1)+"</X>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<Y>"+GetXMLValue(line1)+"</Y>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<W>"+GetXMLValue(line1)+"</W>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<H>"+GetXMLValue(line1)+"</H>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<ForceEnableAllPlugins>"+GetXMLValue(line1)+"</ForceEnableAllPlugins>"
		  line1 = objYARSource.ReadLine
		objYARTarget.WriteLine "</Demonbuddy>"
		objYARTarget.WriteLine "<Diablo>"
		  objYARTarget.WriteLine "<Username>"&account&"</Username>"
		  objYARTarget.WriteLine "<Password>"&password&"</Password>"
		  objYARTarget.WriteLine "<Location>E:\Applications\Diablo III\"&account&" "&line&"\Diablo III.exe</Location>"
		  objYARTarget.WriteLine "<Language>"+SkipToAndGetXMLValue(objYARSource,"Language")+"</Language>"
		  if line = "US" Then
			objYARTarget.WriteLine "<Region>America</Region>"
		  Elseif line = "EU" Then
			objYARTarget.WriteLine "<Region>Europe</Region>"
		  end if
		  objYARTarget.WriteLine "<Priority>"+SkipToAndGetXMLValue(objYARSource,"Priority")+"</Priority>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<UseIsBoxer>"+GetXMLValue(line1)+"</UseIsBoxer>"
		  objYARTarget.WriteLine "<DisplaySlot />"
		  objYARTarget.WriteLine "<CharacterSet />"
		  objYARTarget.WriteLine "<ManualPosSize>"+SkipToAndGetXMLValue(objYARSource,"ManualPosSize")+"</ManualPosSize>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<X>"+GetXMLValue(line1)+"</X>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<Y>"+GetXMLValue(line1)+"</Y>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<W>"+GetXMLValue(line1)+"</W>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<H>"+GetXMLValue(line1)+"</H>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<NoFrame>"+GetXMLValue(line1)+"</NoFrame>"
		  objYARTarget.WriteLine "<UseAuthenticator>false</UseAuthenticator>"
		  objYARTarget.WriteLine "<Serial>---</Serial>"
		  objYARTarget.WriteLine "<RestoreCode />"
		  objYARTarget.WriteLine "<CpuCount>"+SkipToAndGetXMLValue(objYARSource,"CpuCount")+"</CpuCount>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<ProcessorAffinity>"+GetXMLValue(line1)+"</ProcessorAffinity>"
		objYARTarget.WriteLine "</Diablo>"
		objYARTarget.WriteLine "<Week>"
		  objYARTarget.WriteLine "<Monday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Monday>"
		  objYARTarget.WriteLine "<Tuesday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Tuesday>"
		  objYARTarget.WriteLine "<Wednesday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Wednesday>"
		  objYARTarget.WriteLine "<Thursday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Thursday>"
		  objYARTarget.WriteLine "<Friday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Friday>"
		  objYARTarget.WriteLine "<Saturday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Saturday>"
		  objYARTarget.WriteLine "<Sunday>"
			objYARTarget.WriteLine "<Hours>"
				Call PrintBooleans(24,90)
			objYARTarget.WriteLine "</Hours>"
		  objYARTarget.WriteLine "</Sunday>"
		  objYARTarget.WriteLine "<MinRandom>"+SkipToAndGetXMLValue(objYARSource,"MinRandom")+"</MinRandom>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<MaxRandom>"+GetXMLValue(line1)+"</MaxRandom>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<Shuffle>"+GetXMLValue(line1)+"</Shuffle>"
		objYARTarget.WriteLine "</Week>"
		objYARTarget.WriteLine "<ProfileSchedule>"
		  objYARTarget.WriteLine "<UseThirdPartyPlugin>"+SkipToAndGetXMLValue(objYARSource,"UseThirdPartyPlugin")+"</UseThirdPartyPlugin>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<MaxRandomRuns>"+GetXMLValue(line1)+"</MaxRandomRuns>"
		  line1 = objYARSource.ReadLine
		  objYARTarget.WriteLine "<MaxRandomTime>"+GetXMLValue(line1)+"</MaxRandomTime>"
		  objYARTarget.WriteLine "<Profiles>"
			objYARTarget.WriteLine "<Profile>"
			  objYARTarget.WriteLine "<Name>"+SkipToAndGetXMLValue(objYARSource,"Name")+"</Name>"
			  line1 = objYARSource.ReadLine
			  objYARTarget.WriteLine "<Location>"+GetXMLValue(line1)+"</Location>"
			  line1 = objYARSource.ReadLine
			  objYARTarget.WriteLine "<Runs>"+GetXMLValue(line1)+"</Runs>"
			  line1 = objYARSource.ReadLine
			  objYARTarget.WriteLine "<Minutes>"+GetXMLValue(line1)+"</Minutes>"
			  line1 = objYARSource.ReadLine
			  objYARTarget.WriteLine "<MonsterPowerLevel>"+GetXMLValue(line1)+"</MonsterPowerLevel>"
			objYARTarget.WriteLine "</Profile>"
		  objYARTarget.WriteLine "</Profiles>"
		  objYARTarget.WriteLine "<Random>"+SkipToAndGetXMLValue(objYARSource,"Random")+"</Random>"
		objYARTarget.WriteLine "</ProfileSchedule>"
		objYARTarget.WriteLine "<UseWindowsUser>"+SkipToAndGetXMLValue(objYARSource,"UseWindowsUser")+"</UseWindowsUser>"
		line1 = objYARSource.ReadLine
		objYARTarget.WriteLine "<CreateWindowsUser>"+GetXMLValue(line1)+"</CreateWindowsUser>"
		objYARTarget.WriteLine "<WindowsUserName />"
		objYARTarget.WriteLine "<WindowsUserPassword />"
		objYARTarget.WriteLine "<D3PrefsLocation>"+SkipToAndGetXMLValue(objYARSource,"D3PrefsLocation")+"</D3PrefsLocation>"
	  objYARTarget.WriteLine "</BotClass>"
	WScript.Echo "Success"
	objYARSource.Close
Loop Until objRegions.AtEndOfStream

objYARTarget.WriteLine "</ArrayOfBotClass>"
objYARTarget.Close
objRegions.Close
If objFSO.FileExists(myYARSource & ".old.xml") Then
	objFSO.DeleteFile myYARSource & ".old.xml"
End if
objFSO.MoveFile myYARSource+".xml",myYARSource+".old.xml"
objFSO.MoveFile myYARTarget+".xml",myYARSource+".xml" 
Call WriteNewAccount()
Set objFSO=Nothing
'end

Function PrintBooleans(num, percent)
	dim i,rnd
	For i = 1 to num
		rnd=Int(percent-1)*Rnd
		If rnd > 50 Then
			objYarTarget.WriteLine "<boolean>true</boolean>"
		Else
			objYarTarget.WriteLine "<boolean>false</boolean>"
		End If
	Next
End Function

Function CheckExisting()
	dim i
	Set objYARSource = objFSO.OpenTextFile(myAccounts,1)
	Do
		i = objYARSource.ReadLine
		If i = account Then
			WScript.Echo "Account Exists."
			WScript.Quit 1
		End If
	Loop Until objYARSource.AtEndOfStream
	objYARSource.Close
End Function

Function WriteNewAccount()
	Set objYARSource = objFSO.OpenTextFile(myAccounts,8)
	objYARSource.Write vbCrLf & account
	objYARSource.Close
End Function

Function GetXMLValue(node)
	node=Cstr(node)
	If InStr(node,">") < InStrRev(node,"<") then
		GetXMLValue = Mid(node,InStr(node,">")+1, Len(node)-InStr(node,">")-(Len(node)-InStrRev(node,"<"))-1)
	else
		GetXMLValue = "NIL"
	End IF
	WScript.Echo "Migrating "&node&" found "&GetXMLValue
End Function

Function GetXMLTag(node)
	GetXMLTag = Mid(node,InStr(node,"<")+1, InStr(node,">")-2)
End Function

Function SkipAhead(file,numLines)
	dim i,e
	For i = 1 to numLines
		e = file.ReadLine
	Next
End Function

Function SkipToAndGetXMLValue(file,str)
	dim i
	Do While InStr(i,str) = 0
		i = file.ReadLine
	Loop
	SkipToAndGetXMLValue = GetXMLValue(i)
End Function

Function SkipTo(file,str)
	dim i
	Do While InStr(i,str) = 0
		i = file.ReadLine
	Loop
End Function

Function GetInputs()
	dim i
	For Each i in objArgs
		If Left(i,7) = "-revert" or Left(i,7) = "/revert" Then
			If objFSO.FileExists(myYARSource & ".old.xml") Then
				objFSO.DeleteFile myYARSource+".xml"
				objFSO.MoveFile myYARSource+".old.xml",myYARSource+".xml"
				WScript.Echo "Reverted Bots.xml."
				WScript.Quit 0
			Else
				WScript.Echo "Cannot revert. No backup file exists."
				WScript.Quit 0
			End If
		End If
		If Left(i,8) = "-account" or Left(i,8) = "/account" Then
			account = Right(i,Len(i)-8)
			WScript.Echo "Account: "&account
		End If
		If Left(i,9) = "-password" or Left(i,9) = "/password" Then
			password = Right(i,Len(i)-9)
			WScript.Echo "Password: "&password
		End If
	Next
	If account = "" Then
		STDOut.Write "Blizzard Account Name: "
		account = STDIn.ReadLine
	End If
	If password = "" Then
		STDOut.Write "Blizzard Account Password: "
		password = STDIn.ReadLine
	End If
End Function
