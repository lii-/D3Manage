Option Explicit
Dim objFSO,objFSO1,objFSO2,objAccounts,myAccounts,objRegions,myRegions,objYARSource,myYARSource,myYARTarget,objYARTarget,objFile,STDIn,STDOut,account,password,line,line2,line3
myAccounts=".\accounts.txt"
myYARSource=".\YAR\Settings\Bots"
myYARTarget=".\YAR\Settings\_Bots"
myRegions=".\regions.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO1 = CreateObject("Scripting.FileSystemObject")
Set objFSO2 = CreateObject("Scripting.FileSystemObject")
Set STDIn=WScript.STDIn
Set STDOut=WScript.STDOut




STDOut.WriteLine "Blizzard Account Name: "
account = STDIn.ReadLine
STDOut.WriteLine "Blizzard Account Password: "
password = STDIn.ReadLine
Set objRegions = objFSO.OpenTextFile(myRegions,1)

Do
	line=objRegions.ReadLine
	WScript.Echo "Creating "+account+" "+line+" folder."
	objFSO.CopyFolder ".\Diablo III", account+" "+line
	
	Set objYARSource = objFSO1.OpenTextFile(myYARSource & ".xml",1)
	If Not objYARTarget.FileExists(myYARTarget & ".xml") Then
		Set objYARTarget = objFSO2.CreateTextFile(myYARTarget & ".xml", true)
	Else
		Set objYARTarget = objFSO2.OpenTextFile(myYARTarget & ".xml",0)
	End if
	WScript.Echo "Writing to Bots.xml"
	Do
		line1=objYARSource.ReadLine
		if StrComp(Left(line1,9),"<BotClass>") = 1 then
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
				  objYARTarget.WriteLine "<CpuCount>8</CpuCount>"
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
				  objYARTarget.WriteLine "<CpuCount>8</CpuCount>"
				  objYARTarget.WriteLine "<ProcessorAffinity>15</ProcessorAffinity>"
				objYARTarget.WriteLine "</Diablo>"
				objYARTarget.WriteLine "<Week>"
				  objYARTarget.WriteLine "<Monday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
					objYARTarget.WriteLine "</Hours>"
				  objYARTarget.WriteLine "</Monday>"
				  objYARTarget.WriteLine "<Tuesday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
					objYARTarget.WriteLine "</Hours>"
				  objYARTarget.WriteLine "</Tuesday>"
				  objYARTarget.WriteLine "<Wednesday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
					objYARTarget.WriteLine "</Hours>"
				  objYARTarget.WriteLine "</Wednesday>"
				  objYARTarget.WriteLine "<Thursday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
					objYARTarget.WriteLine "</Hours>"
				  objYARTarget.WriteLine "</Thursday>"
				  objYARTarget.WriteLine "<Friday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
					objYARTarget.WriteLine "</Hours>"
				  objYARTarget.WriteLine "</Friday>"
				  objYARTarget.WriteLine "<Saturday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
					objYARTarget.WriteLine "</Hours>"
				  objYARTarget.WriteLine "</Saturday>"
				  objYARTarget.WriteLine "<Sunday>"
					objYARTarget.WriteLine "<Hours>"
						Call PrintBooleans()
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
		else
			objYARTarget.WriteLine line1
		end if
	Loop Until objYAR.AtEndOfStream
	WScript.Echo "Success"
	objYARSource.Close
	objYARTarget.Close
Loop Until objRegions.AtEndOfStream

'fso.MoveFile myYARSource+".xml",myYARSource+".old.xml"
'fso.MoveFile myYARTarget+".xml",myYARSource+".xml" 
objRegions.Close


'end
Set objFSO=Nothing

Function PrintBooleans()
	For i = 1 to 7
		objYarTarget.WriteLine "<boolean>true</boolean>"
	Next
End Function
