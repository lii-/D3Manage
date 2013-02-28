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
	line=objRegions.ReadLine
	WScript.Echo "Creating '"+account+" "+line+"' folder."
	objFSO.CopyFolder "Diablo III", account+" "+line
	
	WScript.Echo "Writing to Bots.xml"
	objYARTarget.WriteLine "<BotClass>"
		objYARTarget.WriteLine "<Name>"&account&"</Name>"
		objYARTarget.WriteLine "<Description>"&line&"</Description>"
		objYARTarget.WriteLine "<IsEnabled>true</IsEnabled>"
		objYARTarget.WriteLine "<Demonbuddy>"
		  objYARTarget.WriteLine "<BuddyAuthUsername />"
		  objYARTarget.WriteLine "<BuddyAuthPassword />"
		  objYARTarget.WriteLine "<Location>e:\Applications\Diablo III\"&account&" "&line&"\Demonbuddy.exe</Location>"
		  objYARTarget.WriteLine "<Key>123</Key>"
		  objYARTarget.WriteLine "<CombatRoutine>Generic</CombatRoutine>"
		  objYARTarget.WriteLine "<NoFlash>true</NoFlash>"
		  objYARTarget.WriteLine "<AutoUpdate>false</AutoUpdate>"
		  objYARTarget.WriteLine "<NoUpdate>true</NoUpdate>"
		  objYARTarget.WriteLine "<Priority>2</Priority>"
		  objYARTarget.WriteLine "<CpuCount>4</CpuCount>"
		  objYARTarget.WriteLine "<ProcessorAffinity>15</ProcessorAffinity>"
		  objYARTarget.WriteLine "<ManualPosSize>false</ManualPosSize>"
		  objYARTarget.WriteLine "<X>0</X>"
		  objYARTarget.WriteLine "<Y>0</Y>"
		  objYARTarget.WriteLine "<W>320</W>"
		  objYARTarget.WriteLine "<H>200</H>"
		  objYARTarget.WriteLine "<ForceEnableAllPlugins>true</ForceEnableAllPlugins>"
		objYARTarget.WriteLine "</Demonbuddy>"
		objYARTarget.WriteLine "<Diablo>"
		  objYARTarget.WriteLine "<Username>"&account&"</Username>"
		  objYARTarget.WriteLine "<Password>"&password&"</Password>"
		  objYARTarget.WriteLine "<Location>E:\Applications\Diablo III\"&account&" "&line&"\Diablo III.exe</Location>"
		  objYARTarget.WriteLine "<Language>English (US)</Language>"
		  if line = "US" Then
			objYARTarget.WriteLine "<Region>America</Region>"
		  Elseif line = "EU" Then
			objYARTarget.WriteLine "<Region>Europe</Region>"
		  end if
		  objYARTarget.WriteLine "<Priority>2</Priority>"
		  objYARTarget.WriteLine "<UseIsBoxer>false</UseIsBoxer>"
		  objYARTarget.WriteLine "<DisplaySlot />"
		  objYARTarget.WriteLine "<CharacterSet />"
		  objYARTarget.WriteLine "<ManualPosSize>false</ManualPosSize>"
		  objYARTarget.WriteLine "<X>0</X>"
		  objYARTarget.WriteLine "<Y>0</Y>"
		  objYARTarget.WriteLine "<W>320</W>"
		  objYARTarget.WriteLine "<H>200</H>"
		  objYARTarget.WriteLine "<NoFrame>true</NoFrame>"
		  objYARTarget.WriteLine "<UseAuthenticator>false</UseAuthenticator>"
		  objYARTarget.WriteLine "<Serial>---</Serial>"
		  objYARTarget.WriteLine "<RestoreCode />"
		  objYARTarget.WriteLine "<CpuCount>4</CpuCount>"
		  objYARTarget.WriteLine "<ProcessorAffinity>15</ProcessorAffinity>"
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
		  objYARTarget.WriteLine "<MinRandom>0</MinRandom>"
		  objYARTarget.WriteLine "<MaxRandom>0</MaxRandom>"
		  objYARTarget.WriteLine "<Shuffle>false</Shuffle>"
		objYARTarget.WriteLine "</Week>"
		objYARTarget.WriteLine "<ProfileSchedule>"
		  objYARTarget.WriteLine "<UseThirdPartyPlugin>false</UseThirdPartyPlugin>"
		  objYARTarget.WriteLine "<MaxRandomRuns>0</MaxRandomRuns>"
		  objYARTarget.WriteLine "<MaxRandomTime>0</MaxRandomTime>"
		  objYARTarget.WriteLine "<Profiles>"
			objYARTarget.WriteLine "<Profile>"
			  objYARTarget.WriteLine "<Name>Questing</Name>"
			  objYARTarget.WriteLine "<Location>E:\Applications\Diablo III\"&account&" "&line&"\questing\Profiles\Questing\Profile.xml</Location>"
			  objYARTarget.WriteLine "<Runs>0</Runs>"
			  objYARTarget.WriteLine "<Minutes>0</Minutes>"
			  objYARTarget.WriteLine "<MonsterPowerLevel>Disabled</MonsterPowerLevel>"
			objYARTarget.WriteLine "</Profile>"
		  objYARTarget.WriteLine "</Profiles>"
		  objYARTarget.WriteLine "<Random>false</Random>"
		objYARTarget.WriteLine "</ProfileSchedule>"
		objYARTarget.WriteLine "<UseWindowsUser>false</UseWindowsUser>"
		objYARTarget.WriteLine "<CreateWindowsUser>false</CreateWindowsUser>"
		objYARTarget.WriteLine "<WindowsUserName />"
		objYARTarget.WriteLine "<WindowsUserPassword />"
		objYARTarget.WriteLine "<D3PrefsLocation>E:\Applications\Diablo III\D3Prefs.txt</D3PrefsLocation>"
	  objYARTarget.WriteLine "</BotClass>"
	WScript.Echo "Success"
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
