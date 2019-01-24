sPath = SetPath()
'Sets path to current directory

StartSAP
'Start SAP

ErrCatch 
'SAP procedura ZMATERIAL
'SAP procedure ZMATERIAL

'EndSAP
'END SAP

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
'There were errors with pure Date function so the function is transformed
'Variables contain the current date, yesterday's date, and date for one month back

	'Dim today, yesterday, monthBack 'Creating variables
	
	today = Day(Date) & "." & Month(Date) & "." & Year(Date)  'Formatting today's date
	
	yesterday = DateAdd("d","-1",Date) 'Transforming current date to yesterday
	yesterday =  Day(yesterday) & "." & Month(yesterday) & "." & Year(yesterday)' Formatting date
	
	'monthBack = DateAdd("m","-1",Date)'Transforming current date to a month back
	'monthBack = Day(monthBack) & "." & Month(monthBack) & "." & Year(monthBack) ' Formatting date
	
	'today = Format(now(),"dd.mm.yyyy")
	'today = FormatDateTime(Date, "dd.mm.yyyy")
	
	'yesterday = DateAdd("d","-1",Date) 'Transforming current date to yesterday
	
	'yesterday = Format(now(), "dd.mm.yyyy")
	'yesterday = FormatDateTime(yesterday, "dd.mm.yyyy")
	
ErrStep = 3

'SAP Commands

exportFileName = "test_data.txt"

    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZPRMA1"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/chkP_PERJN").Selected = False  'izkljucim samo za tiste z maticno
    Session.findById("wnd[0]/usr/ctxtP_WERKS-LOW").Text = "5101"
    Session.findById("wnd[0]/usr/txtP_ARBPL-LOW").Text = ""
    Session.findById("wnd[0]/usr/ctxtP_PLANR-LOW").Text = "M30"
    Session.findById("wnd[0]/usr/ctxtP_PLANR-HIGH").Text = "M98"
    Session.findById("wnd[0]/usr/ctxtP_BUDAT-LOW").Text = yesterday   'datum od
    Session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").Text = today   'datum do
    Session.findById("wnd[0]/usr/chkP_PERJN").SetFocus
    Session.findById("wnd[0]").sendVKey 8 'F8
    
'Tukaj shranis:
    Session.findById("wnd[0]/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    Session.findById("wnd[0]/shellcont/shell").selectContextMenuItem "&PC"
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    Session.findById("wnd[1]").sendVKey 0
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = sPath 'lokacija
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportFileName 'ime datoteke
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    Session.findById("wnd[1]/tbar[0]/btn[11]").press
    Session.findById("wnd[0]").sendVKey 12  'F12 - * zakljuci transakcijo
	
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