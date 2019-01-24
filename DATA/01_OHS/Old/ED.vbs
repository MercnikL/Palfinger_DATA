'Extract Data from MIS10_DATA_10.MDB table 80_Poskodbe_pri_delu (//SSIMARFILE01/MIS10/....)
'and save into data.csv
'SLD 2018-11-25

    Set fs = CreateObject("Scripting.FileSystemObject")
    todaysdate = Now()
    OutputFile = "C:\Users\slanad\OneDrive\Dokumenti\Palfinger\DATA\01_OHS\data.csv"
    sLocilo = "; "
    sString = Chr(39)

    connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=\\ssimarfile01\mis10\Proizvodnja\BAZA\00_ACCESS\MIS_DATA_10.MDB"
    Set objConn = CreateObject("ADODB.Connection")
    objConn.Open connStr
    'Define recordset and SQL query
    Set rs = objConn.Execute("SELECT * FROM 80_Poskodbe_pri_delu ")
 
    sTxt = ""
    For i = 0 To rs.Fields.Count - 1
        sTxt = sTxt & sString & rs.Fields(i).Name & sString & sLocilo
    Next
    sTxt = Left(sTxt, Len(sTxt) - 2) 'odstranim zadnje podpicje
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(OutputFile, True)

    f.WriteLine (sTxt)
    
'    For i = 0 To rs.Fields.Count - 1
'        Debug.Print rs.Fields(i).Name, rs.Fields(i).Type
'    Next
        
    
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

        f.WriteLine (sTxt)
        rs.MoveNext
    Wend
        
    'Close connection and release objects
    objConn.Close
    Set rs = Nothing
    Set objConn = Nothing
    Set f = Nothing
    Set fs = Nothing
