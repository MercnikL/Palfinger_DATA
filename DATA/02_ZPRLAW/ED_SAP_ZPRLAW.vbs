sPath = SetPath()
'Sets path to current directory

StartSAP
'Start SAP

ErrCatch 
'SAP procedura ZMATERIAL
'SAP procedure ZMATERIAL

Python
'Transforms Data gathered from SAP via Python
'Transformira podatke pridobljene iz SAP z Pythonom

EndSAP
'END SAP

Sub Python

Set fso = CreateObject("Scripting.FileSystemObject")
'Setting path for Python
'Nastavljanje poti Pythona

path_student = "C:/Users/student5/Anaconda/Python.exe"
path_home = "D:/OneDrive/Dokumenti/Python"
path_work = "C:/Users/slanad/OneDrive/Dokumenti/Python"
path_daniela = "C:/Users/bedernjakd/Documents"

If (fso.FileExists(path_home)) Then
	pathPython = path_home

ElseIf (fso.FileExists(path_work)) Then
	pathPython = path_work
	
ElseIf (fso.FileExists(path_daniela)) Then
	pathPython = path_daniela

ElseIf (fso.FileExists(path_student)) Then
	pathPython = path_student

Else
    WScript.Echo "Could not find path"
End If

'Runs cmd line
Set winShell = CreateObject("WScript.Shell")
WaitOnReturn = False
windowStyle = 1

'Define the command to run the python file and exit when done
command1 = pathPython & " TransformZPRLAW.py" 
command2 = "exit"

'Run the commands
Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)

End Sub

Function SetPath()

	'Pot vzame iz starša datoteke, kjer se skripta nahaja. Doda še \ za lazje zdruzevanje
	'Path is taken from parent of the file, where script is located. Adds a \ for easier combining
	SetPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"

End Function


Sub ErrCatch() 'If Error occurs, save the log about it

	dim Res, CurrentStep, sLog
	LogFile = "ED_log.txt" 'Log file
	
	on error resume next
	Res = UnSafeCode(CurrentStep)
	
	'Creates a string containing on which error step did Error occur and provides a description of the error
	sLog = cstr(now()) & "; ErrStep= " & CurrentStep & "; " & Err.Description 

	Set fs = CreateObject("Scripting.FileSystemObject")
	'Set f = fs.CreateTextFile(LogFile, True)
	Set f = fs.OpenTextFile(LogFile, 8, True, 0) '8=append
	
	f.WriteLine(sLog) 'Write in the file

	f.Close
	Set f = Nothing 'Clearing memory
	Set fs = Nothing 'Clearing memory

end sub


Function UnSafeCode(ErrStep)

ErrStep = 1
'Oppening SAPGUI

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


ErrStep = 2

'Data location
'Lokacije datotek


sPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"
'sFile1 = "data_1D.txt" <-- test file
sFile2 = "data_120D.txt"

'There were errors with pure Date function so the function is transformed
'Variables contain the current date, yesterday's date, and date for one month back

'Dim today, yesterday, monthBack 'Creating variables
yesterday = DateAdd("d","-1",Date) 'Transforming current date to yesterday
yesterday = Day(yesterday) & "." & Month(yesterday) & "." & Year(yesterday)' Formatting date
	
back127 = DateAdd("d","-127",Date) 'Transforming current date to 127 days back
back127 = Day(back127) & "." & Month(back127) & "." & Year(back127)' Formatting date


''izvozi podatke za 1 dan
'   session.FindById("wnd[0]").Maximize
'   session.FindById("wnd[0]/tbar[0]/okcd").Text = "zprlaw"
'    session.FindById("wnd[0]").SendVKey 0 'ENTER
'    session.FindById("wnd[0]/usr/ctxtP_WERKS-LOW").Text = "5101"
'    session.FindById("wnd[0]/usr/ctxtP_BUDAT-LOW").Text = CStr(Date - 1)   'vcerajnji datum "10.12.2018"
'    session.FindById("wnd[0]").SendVKey 8 'F8
'    session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT" 'Export
'    session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem "&PC"
'    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select 'Text s tabulatorji
'    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
'    session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Enter
'	session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = sPath 'lokacija
'    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = sFile1 'ime datoteke
''    session.FindById("wnd[2]/usr/ctxtDY_FILENAME").CaretPosition = 12 'novi file
'    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 16 'nadomesti
'    session.FindById("wnd[1]").SendVKey 11
'    session.FindById("wnd[0]").SendVKey 12 'F12 - * zakljuci transakcijo
'    session.FindById("wnd[0]").SendVKey 12 'F12 - * zakljuci transakcijo

ErrStep = 3
'izvozi podatke za zadnjih 120 dni
	Session.findById("wnd[0]").maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "zprlaw"
    session.FindById("wnd[0]").SendVKey 0 'ENTER
    session.FindById("wnd[0]/usr/ctxtP_WERKS-LOW").Text = "5101"
    session.FindById("wnd[0]/usr/ctxtP_BUDAT-LOW").Text = back127   'danes - 127 dni "10.12.2018"	
	session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").text = yesterday	'danes - 1 dan
    session.FindById("wnd[0]").SendVKey 8 'F8
    session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT" 'Export
    session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem "&PC"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select 'Text s tabulatorji
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Enter
	session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = sPath 'lokacija
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = sFile2 'ime datoteke
'    session.FindById("wnd[2]/usr/ctxtDY_FILENAME").CaretPosition = 12 'novi file
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 16 'nadomesti
    session.FindById("wnd[1]").SendVKey 11
    session.FindById("wnd[0]").SendVKey 12 'F12 - * zakljuci transakcijo
    session.FindById("wnd[0]").SendVKey 12 'F12 - * zakljuci transakcijo

ErrStep = -1
	'If everything is OK
	'Če vse dela


End Function
	
Private Sub StartSAP() 'Start SAP Session and login
'Start SAP
	
	'Kreiramo cmd objekt
	'Creating a cmd object
	Set winShell = CreateObject("WScript.Shell") 
	WaitOnReturn = False
	windowStyle = 1

    'Preberi ime in geslo iz SAP_Credentials.txt
	'Read username and password from SAP_Credentials.txt
    Dim f, sUser, sPass
    Set f = CreateObject("Scripting.FileSystemObject").OpenTextFile("..\SAP_Credentials.txt", 1) '1 je za read / 1 means read
    sUser = f.ReadLine
    sPass = f.ReadLine
    Set f = Nothing

    command1 = "start sapshcut -sysname=PD1 -client=400 -user=" & sUser & " -pw=" & sPass &" " 'Defining command to log-in into SAP
    command2 = "exit " 'Defining command to exit cmd
	
	'Zaženi zaželjene ukaze v cmd 
    'Run the desired pair of shell commands in command prompt.
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
	
	'Pocakaj 7 sekund da se SAP nalozi, ker v prejsnjem stavku WaitOnReturn ne dela vedno
	'Wait 7 seconds for SAP to load, because WaitOnReturn does not work everytime
    dteWait = DateAdd("s", 7, Now())
    Do Until (Now() > dteWait)
    Loop
	
End Sub

Private Sub EndSAP() 'Kills SAP Session

	'Kreiramo cmd objekt
	 'Creating a cmd object
	Set winShell = CreateObject("WScript.Shell")
	WaitOnReturn = False
	windowStyle = 1
	
    command1 = "taskkill /F /IM saplogon.exe " 'Defining command to kill the program
    command2 = "exit " 'Defining command to exit cmd
	
	'Zaženi ukaze v cmd
	'Running the commands in cmd
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
	
	'Sprosti kazalec
	'Release pointer to the command prompt.
    Set winShell = Nothing 
	
End Sub