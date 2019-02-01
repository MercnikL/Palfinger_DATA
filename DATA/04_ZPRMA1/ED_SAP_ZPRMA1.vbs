sPath = SetPath() 'Sets path to current directory

'Sets the path for python.exe
pPath = PythonPath()

'Start SAP
StartSAP	

'Check Multiple logins pop-up window
'Preveri okno, ki se prikaže ob veèkratni prijavi
MultipleLogins

'Izvede ZPRMA1
ErrCatch	

'END SAP
EndSAP		 


Function SetPath()

	'Pot vzame iz starša datoteke, kjer se skripta nahaja. Doda še \ za lazje zdruzevanje
	'Path is taken from parent of the file, where script is located. Adds a \ for easier combining
	SetPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"

End Function

Function PythonPath()

	Set fso = CreateObject("Scripting.FileSystemObject")
	'Setting path for Python
	'Nastavljanje poti Pythona

	path_student = "C:/Users/student5/Anaconda/Python.exe"
	path_slanad = "C:/Users/slanad/AppData/Local/Continuum/anaconda3/Python.exe"
	path_work = "C:/Users/slanad/OneDrive/Dokumenti/Python"

	If (fso.FileExists(path_slanad)) Then
		PythonPath = path_slanad

	ElseIf (fso.FileExists(path_work)) Then
		PythonPath = path_work

	ElseIf (fso.FileExists(path_student)) Then
		PythonPath = path_student

	Else
		WScript.Echo "Could not find path"
	End If

End Function

Sub MultipleLogins

	'Runs cmd line
	Set winShell = CreateObject("WScript.Shell")
	WaitOnReturn = True
	windowStyle = 1

	'Define the command to run the python file and exit when done
	command1 = pPath & " autoClose.py" 
	command2 = "exit"

	'Run the commands
	Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
	
	dteWait = DateAdd("s", 3, Now())
    Do Until (Now() > dteWait)
    Loop

End Sub


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

End Sub
	
Private Sub StartSAP() 'Start SAP Session and login
'Start SAP
	
	'Kreiramo cmd objekt
	'Creating a cmd object
	
	Set winShell = CreateObject("WScript.Shell") 
	WaitOnReturn = True
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


Function UnSafeCode(ErrStep)

ErrStep = 1
'Opening SAPGUI

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
'There were errors with pure Date function so the function is transformed
'Variables contain the current date, yesterday's date, and date for one month back
	
	Dim today, yesterday, monthBack 'Creating variables
	
	today = Day(Date) & "." & Month(Date) & "." & Year(Date) 'Formatting today's date
	
	yesterday = DateAdd("d","-1",Date) 'Transforming current date to yesterday
	yesterday =  Day(yesterday) & "." & Month(yesterday) & "." & Year(yesterday)' Formatting date
	
	monthBack = DateAdd("m","-1",Date)'Transforming current date to a month back
	monthBack = Day(monthBack) & "." & Month(monthBack) & "." & Year(monthBack) ' Formatting date

	
ErrStep = 3

'SAP Commands

exportFileName = "data_ZPRMA1.txt"

	session.findById("wnd[0]").maximize
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPRMA1" 'Transaction date
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/radP_FILE").select
	session.findById("wnd[0]/usr/ctxtP_WERKS-LOW").text = "5101"
	session.findById("wnd[0]/usr/ctxtP_BUDAT-LOW").text = yesterday 'From which date
	session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").text = today 'To which date
	session.findById("wnd[0]/usr/chkP_PERJN").selected = false
	session.findById("wnd[0]/usr/radP_FILE").setFocus
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[1]").sendVKey 4
	session.findById("wnd[2]/usr/ctxtDY_PATH").text = sPath 'Sets path according to the location of this file
	session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = exportFileName 'File name
	session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 9
	session.findById("wnd[2]/tbar[0]/btn[11]").press
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	
	
ErrStep = -1
	'If everything is OK
	'Èe vse dela

End Function