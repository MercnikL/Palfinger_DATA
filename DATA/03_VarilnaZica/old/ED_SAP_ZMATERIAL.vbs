'Iz SAP s trasn. ZMATERIAL potegni podatke za varilno zico - W.nr se vpise v seznam
'Slana 2018-12-11

ErrCatch 
'SAP procedura ZMATERIAL

Macro_MD04_VarilnaZica
'SAP procedura MD04

'END SAP
	Set winShell = CreateObject("WScript.Shell")
	WaitOnReturn = False
	windowStyle = 1
    command1 = "taskkill /F /IM saplogon.exe "
    command2 = "exit "
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
        
    'Release pointer to the command prompt.
    Set winShell = Nothing
	

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

    'Preberi geslo iz SAP_Credentials.txt
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
sPath = "C:\Users\slanad\OneDrive\Dokumenti\Palfinger\DATA\03_VarilnaZica\"
sFile = "data_ZMATERIAL.txt"

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
 
ErrStep = 4

    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "zmaterial"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press 'pritisnem za izbiro materila
    
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00003"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006701+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006701+00003"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006701+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006121"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10000626+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10002968"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10000623+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007347"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007241+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007241+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007241+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007243+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007243+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007243+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007242+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007363+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007363+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007363+00003"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007363+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007364"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10012089"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10008495"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10052933"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007362+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	
    Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00003"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006119+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006701+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006701+00003"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006701+00005"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10006121"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10000623+00002"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico
	Session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "W10007362+00001"
    Session.findById("wnd[1]").sendVKey 13 'SHIFT+F1 - vstavi vrstico    
		
    Session.findById("wnd[1]").sendVKey 0   'Enter
    Session.findById("wnd[1]").sendVKey 8   'F8
    Session.findById("wnd[0]").sendVKey 8   'F8
    Session.findById("wnd[0]/tbar[1]/btn[45]").press 'shrani datoteko kot
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select  'kot teksti s tabolatorji
    Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus 'kot teksti s tabolatorji
    Session.findById("wnd[1]").sendVKey 0 'Enter
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = sPath
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = sFile
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
    Session.findById("wnd[1]/tbar[0]/btn[11]").press 'nadomesti
    Session.findById("wnd[0]").sendVKey 12 'F12 - * zakljuci transakcijo
    Session.findById("wnd[0]").sendVKey 12 'F12 - * zakljuci transakcijo

'MD04
    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = "W10012089"
    'Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").caretPosition = 9
    Session.findById("wnd[0]").sendVKey 0
    
    'PROSTA ZALOGA
    ProstaZaloga = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[9,0]").Text
    'Set myGrid = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ")
    'ProstaZaloga = Val(myGrid.getcell(0, 9).Text) * 1000
    'ProstaZaloga = Val(ProstaZaloga)*1000

    Session.findById("wnd[0]").sendVKey 12
    Session.findById("wnd[0]").sendVKey 12

	
ErrStep = 6

    'zapisi rezultat
    sDatoteka = "C:\Users\slanad\OneDrive\Dokumenti\Palfinger\DATA\03_VarilnaZica\Data_MD04.txt"
    'Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(sDatoteka)
        oFile.WriteLine "W, Razpolozljiva_zaloga"
        oFile.WriteLine "W10012089, " & CStr(ProstaZaloga)
    oFile.Close

ErrStep = -1


End Function



Private Sub Macro_MD04_VarilnaZica()
'Izvede transakcijo ZMATERIAL za varilne žice
'prijavljen v SAP ze moras biti!

'Lokacije datotek
    sPath = "C:\Users\slanad\OneDrive\Dokumenti\Palfinger\DATA\03_VarilnaZica\"
    sFileExcelData = "W_varilna_zica.xlsx"
    
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
    Session.findById("wnd[0]").maximize
    
    
    Set oExcel = CreateObject("Excel.Application")
    Set oWB = oExcel.Workbooks.Open(sPath & sFileExcelData)
    
    sBesedilo = "W, Razpolozljiva_zaloga, Datum_naslednjega_narocila, Kolicina_narocila " & Chr(13) & Chr(10)
    sLocilo = ", "
    iRow = 2
    Do Until (oExcel.Cells(iRow, 1).Value = "" Or iRow > 50) ' omejim na max 50 vrstic
        sMaterial = CStr(oExcel.Cells(iRow, 1).Value)
        
        'MD04
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = sMaterial
        Session.findById("wnd[0]").sendVKey 0
    
        'MD04 iz tabele
        
        ProstaZaloga = f_ReadMD04TableCell(Session, 0, 9) '(0,9) Prosta zaloga
        ProstaZaloga = f_RemoveDot(CStr(ProstaZaloga))
        
        NaslednjeNarociloDatum = f_ReadMD04TableCell(Session, 2, 1) '(2,1) Naslednje narocilo - Datum
		
        NaslednjeNarociloKol = f_ReadMD04TableCell(Session, 2, 8) '(2,4) Naslednje narocilo - kolicina
        NaslednjeNarociloKol = f_RemoveDot(CStr(NaslednjeNarociloKol))
        
        sBesedilo = sBesedilo & CStr(sMaterial) & sLocilo & CStr(ProstaZaloga) & sLocilo & CStr(NaslednjeNarociloDatum) & sLocilo & CStr(NaslednjeNarociloKol)
        sBesedilo = sBesedilo & Chr(13) & Chr(10)
        
        iRow = iRow + 1
    Loop
    oExcel.Quit
    Set oExcel = Nothing
     
    Session.findById("wnd[0]").sendVKey 12
    Session.findById("wnd[0]").sendVKey 12
    

    'zapisi rezultat
    sDatoteka = "C:\Users\slanad\OneDrive\Dokumenti\Palfinger\DATA\03_VarilnaZica\Data_MD04_ALL.txt"
    'Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile2 = CreateObject("Scripting.FileSystemObject").CreateTextFile(sDatoteka)
        oFile2.Write sBesedilo
    oFile2.Close
    
End Sub

Function f_ReadMD04TableCell(Session, Vrstica, Stolpec)
'Session je SAP session
'Vrstice se zacnejo pri 0
'Stolpci se zacnejo pri 0
'Nekatere lokacije v tabeli vrnejo napako in je potem vrednost 0
On Error Resume Next
    Set myGrid = Session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ")
    Vrednost = myGrid.getcell(Vrstica, Stolpec).Text
    If Not IsNull(Vrednost) Then
        f_ReadMD04TableCell = Vrednost
    Else
        f_ReadMD04TableCell = 0
    End If
End Function

Function f_RemoveNonNumeric(strSearchString)
    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Global = True
    'objRegEx.Pattern = "[^A-Za-z\n\r]"
    objRegEx.Pattern = "\D" 'matches all non-digits
    strSearchString = objRegEx.Replace(strSearchString, "")
    f_RemoveNonNumeric = Val(strSearchString)
End Function

Function f_RemoveDot(strSearchString)
      f_RemoveDot = Replace(strSearchString, ".", "")
End Function
