'Extract Data from MIS10_DATA_10.MDB table 80_Poskodbe_pri_delu (//SSIMARFILE01/MIS10/....)
'and save into data.csv
'SLD 2018-12-02

'Saves path of current directory
'Shrani pot trenutne mape
sPath = SetPath()

ErrCatch()

Function SetPath()

	'Pot vzame iz starša datoteke, kjer se skripta nahaja. Doda še \ za lazje zdruzevanje
	'Path is taken from parent of the file, where script is located. Adds a \ for easier combining
	SetPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\"

End Function

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
    Set fs = CreateObject("Scripting.FileSystemObject")
    todaysdate = Now()
    OutputFile = sPath & "data.csv"
    sLocilo = "; "
    sString = Chr(39)

    connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=\\ssimarfile01\mis10\Proizvodnja\BAZA\00_ACCESS\MIS_DATA_10.MDB"
    Set objConn = CreateObject("ADODB.Connection")
    objConn.Open connStr

ErrStep = 2
    'Define recordset and SQL query
    Set rs = objConn.Execute("SELECT * FROM 80_Poskodbe_pri_delu ")

ErrStep = 3
    sTxt = ""
    For i = 0 To rs.Fields.Count - 1
        sTxt = sTxt & sString & rs.Fields(i).Name & sString & sLocilo
    Next
    sTxt = Left(sTxt, Len(sTxt) - 2) 'odstranim zadnje podpicje
	sTxt = sString & "Prenos" & sString & sLocilo & sTxt 'Tukaj bo vpisan now()

ErrStep = 4    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(OutputFile, True)

    f.WriteLine (sTxt)
    
        
ErrStep = 5
	sSedaj = Cstr(now())
    While Not rs.EOF
        sTxt = ""
        For i = 0 To rs.Fields.Count - 1
            If rs.Fields(i).Type = 202 Then 'text
                If IsNull(rs.Fields(i).Value) Then
                    sTxt = sTxt & sString & sString & sLocilo
                Else
                    sTxt = sTxt & sString & CStr(rs.Fields(i).Value) & sString & sLocilo
                    sTxt = Replace(sTxt, vbCr, "")
                    sTxt = Replace(sTxt, vbLf, " ")
                End If
            ElseIf rs.Fields(i).Type = 11 Then 'boolean
                If IsNull(rs.Fields(i).Value) Then
                    sTxt = sTxt & "0" & sLocilo
                Else
                    If rs.Fields(i).Value = True Then 'iff dela probleme v vbscriptu, zato celi if stavek
                        sTxt = sTxt & CStr(1) & sLocilo
                    Else
                        sTxt = sTxt & CStr(0) & sLocilo
                    End If
                    'sTxt = sTxt & CStr(IIf(rs.Fields(i).Value = true, 1, 0)) & sLocilo
                End If
            Else
                sTxt = sTxt & rs.Fields(i).Value & sLocilo
            End If
        Next
        sTxt = Left(sTxt, Len(sTxt) - 2) 'odstranim zadnje podpicje
		sTxt = sSedaj & sLocilo & sTxt
		
        f.WriteLine (sTxt)
        rs.MoveNext
    Wend
	
ErrStep = 6        
    'Close connection and release objects
    objConn.Close
    Set rs = Nothing
    Set objConn = Nothing
    f.Close
    Set f = Nothing
    Set fs = Nothing

ErrStep = -1

End Function

