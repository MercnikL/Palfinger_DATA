'Iz spiced shrani s porocilo DOGODKI za industrijski cas v vhodno datoteko Spica_I.txt
'Slana 2019-01-20

'Transform the report from spica located in Spica_I to a excel file located in file results, named data_Spica_dogodki
pPath = PythonPath()

TransformPython
'Get Python path, depending on the computer we are running this script from

Function PythonPath()

Set fso = CreateObject("Scripting.FileSystemObject")

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

Sub TransformPython

Set winShell = CreateObject("WScript.Shell")
WaitOnReturn = False
windowStyle = 1

'Define the command to run the python file and exit when done
command1 = pPath & " -i powerBItoPDF_Function.py" 
command2 = "\ZPRLAW_00"


'Run the commands
Call winShell.Run("cmd /k " & command1 & " " & command2, windowStyle, WaitOnReturn)

End Sub