'Iz SAP s trasn. ZMATERIAL potegni podatke za varilno zico - W.nr se vpise v seznam
'Iz SAP s trasn. ZMATERIAL potegni podatke za varilno zico - W.nr se vpise v seznam
'Get Data from SAP with ZMATERIAL for Welding Wire - W.nr is written in the list

'Slana 2018-12-11

sPath = SetPath()
'Sets path to current directory

pPath = PythonPath()
'Sets the path for python.exe

StartSAP
'Start SAP

MultipleLogins
'Check Multiple logins pop-up window
'Preveri okno, ki se prikaže ob veèkratni prijavi

ErrCatch 
'SAP procedura ZMATERIAL
'SAP procedure ZMATERIAL

Macro_MD04_VarilnaZica
'SAP procedura MD04
'SAP procedure MD04

'EndSAP
'END SAP

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
'Pocakaj 7 sekund da se SAP nalozi, ker v prejsnjem stavku WaitOnReturn ne dela vedno
'Wait 7 seconds for SAP to load, because WaitOnReturn does not work everytime

    dteWait = DateAdd("s", 7, Now())
    Do Until (Now() > dteWait)
    Loop

'Lokacije datotek
'Data location

sFile = "data_ZMATERIAL.txt"
sFileExcelData = "W_varilna_zica.xlsx"

ErrStep = 2
'SAP ukazi
'SAP commands

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
 
	
ErrStep = 3

    Session.findById("wnd[0]").maximize 'Maximizes window
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "zmaterial"
    Session.findById("wnd[0]").sendVKey 0 'Enter
    Session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press 'pritisnem za izbiro materiala
	
	
	'Odprem excel dokument  / Oppening Excel document defined in sPath and sFileExcelData
    Set oExcel = CreateObject("Excel.Application") 
    Set oWB = oExcel.Workbooks.Open(sPath & sFileExcelData)

	
    iRow = 2
	
    Do Until (oExcel.Cells(iRow, 1).Value = "" Or iRow > 50) 'Max 50 rows 
        Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = oExcel.Cells(iRow, 1).Value
        Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - Enters rows
        iRow = iRow + 1
    Loop
    oExcel.Quit 'Clossing Excel Workbook
    Set oExcel = Nothing
		
    Session.findById("wnd[1]").sendVKey 0   'Enter
    Session.findById("wnd[1]").sendVKey 8   'F8
    Session.findById("wnd[0]").sendVKey 8   'F8
    Session.findById("wnd[0]/tbar[1]/btn[45]").press 'Save As
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select  'kot teksti s tabolatorji
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus 'kot teksti s tabolatorji
    Session.findById("wnd[1]").sendVKey 0 'Enter
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = sPath
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = sFile
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
    Session.findById("wnd[1]/tbar[0]/btn[11]").press 'Replace existing file
    Session.findById("wnd[0]").sendVKey 12 'F12 - * End Transaction
    Session.findById("wnd[0]").sendVKey 12 'F12 - * Zakljuci transakcijo

'MD04
    Session.findById("wnd[0]").maximize 'Maximizes window
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
    Session.findById("wnd[0]").sendVKey 0  'Enter
    Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = "W10012089"
    'Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").caretPosition = 9
    Session.findById("wnd[0]").sendVKey 0 'Enter
    
    'PROSTA ZALOGA
    ProstaZaloga = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[9,0]").Text
    'Set myGrid = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ")
    'ProstaZaloga = Val(myGrid.getcell(0, 9).Text) * 1000
    'ProstaZaloga = Val(ProstaZaloga)*1000

    Session.findById("wnd[0]").sendVKey 12 'F12 - * End Transaction
    Session.findById("wnd[0]").sendVKey 12 'F12 - * Zakljuci transakcijo

	
ErrStep = 4
    'Zapisi rezultat
	'Outputing results
	
    sDatoteka = sPath & "Data_MD04.txt" 'Path and name of file
    'Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(sDatoteka) 'Creating File
        oFile.WriteLine "W, Razpolozljiva_zaloga" 'Adding header
        oFile.WriteLine "W10012089, " & CStr(ProstaZaloga)
    oFile.Close

ErrStep = -1
	'If everything is OK
	'Èe vse dela


End Function



Private Sub Macro_MD04_VarilnaZica()
'Izvede transakcijo ZMATERIAL za varilne žice
'prijavljen v SAP ze moras biti!

'Running transaction ZMATERIAL for welding wires
'Log on in SAP required


'Lokacije datotek
'Data location


	sFileExcelData = "W_varilna_zica.xlsx"

    
	'Oppening SAPGUI
	
    If Not IsObject(ap) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set ap = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = ap.Children(0)
    End If
    If Not IsObject(Session) Then
       Set Session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject Session, "on"
       WScript.ConnectObject ap, "on"
    End If
 
 'MD04
    Session.findById("wnd[0]").maximize 'Maximizes window
    
    Set oExcel = CreateObject("Excel.Application")
    Set oWB = oExcel.Workbooks.Open(sPath & sFileExcelData) 	'Odprem Excel dokument  / Oppening Excel document defined in sPath and sFileExcelData
    
    sBesedilo = "W, Razpolozljiva_zaloga, Datum_naslednjega_narocila, Kolicina_narocila " & Chr(13) & Chr(10) 'Text for header
    sLocilo = ", " 'Divider
	
    iRow = 2
    Do Until (oExcel.Cells(iRow, 1).Value = "" Or iRow > 50) ' Max 50 rows
        sMaterial = CStr(oExcel.Cells(iRow, 1).Value)
        
        'MD04
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
        Session.findById("wnd[0]").sendVKey 0 'Enter
        Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = sMaterial
        Session.findById("wnd[0]").sendVKey 0 'Enter
    
        'MD04 iz tabele / MDO4 from Table
        
        ProstaZaloga = f_ReadMD04TableCell(Session, 0, 9) '(0,9) Prosta zaloga
        ProstaZaloga = f_RemoveDot(CStr(ProstaZaloga))
        
        NaslednjeNarociloDatum = f_ReadMD04TableCell(Session, 2, 1) '(2,1) Naslednje narocilo - Datum
		
        NaslednjeNarociloKol = f_ReadMD04TableCell(Session, 2, 8) '(2,4) Naslednje narocilo - kolicina
        NaslednjeNarociloKol = f_RemoveDot(CStr(NaslednjeNarociloKol)) 'Deleting dot
        
        sBesedilo = sBesedilo & CStr(sMaterial) & sLocilo & CStr(ProstaZaloga) & sLocilo & CStr(NaslednjeNarociloDatum) & sLocilo & CStr(NaslednjeNarociloKol)
        sBesedilo = sBesedilo & Chr(13) & Chr(10) 'Compositing string
        
        iRow = iRow + 1
    Loop
    oExcel.Quit 'Clossing Excel Workbook
    Set oExcel = Nothing
     
    Session.findById("wnd[0]").sendVKey 12 'F12 - * End Transaction
    Session.findById("wnd[0]").sendVKey 12 'F12 - * Zakljuci transakcijo

    'Zapisi rezultat v datoteko
	'Write result into file
    sDatoteka = sPath & "Data_MD04_ALL.txt" 'Path and name of file
    'Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile2 = CreateObject("Scripting.FileSystemObject").CreateTextFile(sDatoteka) 'Creating File
        oFile2.Write sBesedilo
    oFile2.Close
    
End Sub

Function f_ReadMD04TableCell(Session, Vrstica, Stolpec)
'Session je SAP session
'Vrstice se zacnejo pri 0
'Stolpci se zacnejo pri 0
'Nekatere lokacije v tabeli vrnejo napako in je potem vrednost 0

'Session is SAP Session
'Rows start at 0
'Columns start at 0
'Some locations in table return an error, so their value is 0

On Error Resume Next
    Set myGrid = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ")
    Vrednost = myGrid.getcell(Vrstica, Stolpec).Text  'Read from table
	
    If Not IsNull(Vrednost) Then 'If Error, set value to 0
        f_ReadMD04TableCell = Vrednost
    Else
        f_ReadMD04TableCell = 0
    End If
End Function

Function f_RemoveNonNumeric(strSearchString)
    Set objRegEx = CreateObject("VBScript.RegExp") 'Creating RegEx Object
    objRegEx.Global = True
    'objRegEx.Pattern = "[^A-Za-z\n\r]"
    objRegEx.Pattern = "\D" 'Matches all non-digits
    strSearchString = objRegEx.Replace(strSearchString, "")
    f_RemoveNonNumeric = Val(strSearchString)
End Function

Function f_RemoveDot(strSearchString)
      f_RemoveDot = Replace(strSearchString, ".", "") 'Searches for . in strSearchString and deletes it
End Function

Private Sub EndSAP() 'Kills SAP Session

	Set winShell = CreateObject("WScript.Shell") 'Creating a cmd object
	WaitOnReturn = False
	windowStyle = 1
    command1 = "taskkill /F /IM saplogon.exe " 'Defining command to kill the program
    command2 = "exit " 'Defining command to exit cmd
	
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
	'Running the commands in cmd
        
    'Release pointer to the command prompt.
    Set winShell = Nothing     'Release pointer to the command prompt.
	
End Sub

Private Sub StartSAP() 'Start SAP Session and login

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
	
End Sub


Function PythonPath()

	Set fso = CreateObject("Scripting.FileSystemObject")
	'Setting path for Python
	'Nastavljanje poti Pythona

	path_student = "C:/Users/student5/Anaconda/Python.exe"
	path_home = "D:/OneDrive/Dokumenti/Python"
	path_work = "C:/Users/slanad/OneDrive/Dokumenti/Python"
	path_daniela = "C:/Users/bedernjakd/Documents"

	If (fso.FileExists(path_home)) Then
		PythonPath = path_home

	ElseIf (fso.FileExists(path_work)) Then
		PythonPath = path_work
		
	ElseIf (fso.FileExists(path_daniela)) Then
		PythonPath = path_daniela

	ElseIf (fso.FileExists(path_student)) Then
		PythonPath = path_student
		WScript.Echo "Path found"

	Else
		WScript.Echo "Could not find path"
	End If

End Function

Sub MultipleLogins

'Runs cmd line
Set winShell = CreateObject("WScript.Shell")
WaitOnReturn = False
windowStyle = 1

'Define the command to run the python file and exit when done
command1 = pPath & " autoClose.py" 

command2 = "exit"

'Run the commands
Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)

End Sub

Function SetPath()

	'Pot vzame iz starša datoteke, kjer se skripta nahaja. Doda še \ za lazje zdruzevanje
	'Path is taken from parent of the file, where script is located. Adds a \ for easier combining
	SetPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"

End Function
