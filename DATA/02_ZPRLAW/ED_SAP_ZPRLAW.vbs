'Iz SAP s trasn. ZPRLAW potegni podatke za celo leto in jih shrani
'Slana 2018-12-11

ErrCatch() 'SAP procedura ZPRLAW

'Transformiraj v Pythonu
Set winShell = CreateObject("WScript.Shell")
WaitOnReturn = False
windowStyle = 1
'Run the desired pair of shell commands in command prompt.
command1 = "C:\Users\slanad\AppData\Local\Continuum\anaconda3\python Transform.py data_120D.txt data_120D.csv" 
command2 = "exit"
Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)



Sub ErrCatch()
	dim Res, CurrentStep, sLog
	LogFile = "ED_log.txt"
	on error resume next
	Res = UnSafeCode(CurrentStep)

	sLog = cstr(now()) & "; ErrStep= " & CurrentStep & "; " & Err.Description
	
	Set fs = CreateObject("Scripting.FileSystemObject")
	'Set f = fs.CreateTextFile(LogFile, True)
	Set f = fs.OpenTextFile(LogFile, 8, True, 0) '8=append

	f.WriteLine(sLog)

	f.Close
	Set f = Nothing
	Set fs = Nothing

end sub


Function UnSafeCode(ErrStep)
ErrStep = 1
'Start SAP
Set winShell = CreateObject("WScript.Shell")
WaitOnReturn = False
windowStyle = 1

    'Preberi geslo iz Credentials.txt
    Dim f, sUser, sPass
    Set f = CreateObject("Scripting.FileSystemObject").OpenTextFile("..\SAP_Credentials.txt", 1) '1 je za read
    sUser = f.ReadLine
    sPass = f.ReadLine
    Set f = Nothing

    'Run the desired pair of shell commands in command prompt.
    command1 = "start sapshcut -sysname=PD1 -client=400 -user=" & sUser & " -pw=" & sPass &" "
    command2 = "exit "
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)

ErrStep = 2
'Pocakaj 7 sekund da se SAP nalozi, ker v prejsnjem stavku WaitOnReturn ne dela vedno
    dteWait = DateAdd("s", 7, Now())
    Do Until (Now() > dteWait)
    Loop

'Lokacije datotek
sPath = "C:\Users\slanad\OneDrive\Dokumenti\Palfinger\DATA\02_ZPRLAW\"
'sFile1 = "data_1D.txt"
sFile2 = "data_120D.txt"

ErrStep = 3
'SAP ukazi
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

ErrStep = 4
'izvozi podatke za zadnjih 120 dni
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "zprlaw"
    session.FindById("wnd[0]").SendVKey 0 'ENTER
    session.FindById("wnd[0]/usr/ctxtP_WERKS-LOW").Text = "5101"
    session.FindById("wnd[0]/usr/ctxtP_BUDAT-LOW").Text = CStr(Date - 127)   'danes - 127 dni "10.12.2018"	
	session.findById("wnd[0]/usr/ctxtP_BUDAT-HIGH").text = CStr(Date - 1)	'danes - 1 dan
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
	
ErrStep = 5	
'END SAP
    command1 = "taskkill /F /IM saplogon.exe "
    command2 = "exit "
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
        
    'Release pointer to the command prompt.
    Set winShell = Nothing

ErrStep = -1

End Function



