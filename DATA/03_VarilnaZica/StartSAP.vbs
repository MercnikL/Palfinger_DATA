
'Start SAP
Set winShell = CreateObject("WScript.Shell") 'Oppening cmd
WaitOnReturn = False
windowStyle = 1

    'Preberi ime in geslo iz SAP_Credentials.txt
	'Read username and password from SAP_Credentials.txt
	
    Dim f, sUser, sPass
    Set f = CreateObject("Scripting.FileSystemObject").OpenTextFile("..\SAP_Credentials.txt", 1) '1 je za read / 1 means read
    sUser = f.ReadLine
    sPass = f.ReadLine
    Set f = Nothing

	'Zaženi zaželjene ukaze v cmd
    'Run the desired pair of shell commands in command prompt.
    command1 = "start sapshcut -sysname=PD1 -client=400 -user=" & sUser & " -pw=" & sPass &" "
    command2 = "exit "
	
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
	
	dteWait = DateAdd("s", 7, Now())
    Do Until (Now() > dteWait)
    Loop

	If Not IsObject(Ap) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Ap = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Ap.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Ap, "on"
    End If