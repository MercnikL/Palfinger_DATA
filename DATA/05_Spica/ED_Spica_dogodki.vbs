'Iz spiced shrani s porocilo DOGODKI za industrijski cas v vhodno datoteko Spica_I.txt
'Slana 2019-01-20

'Transform the report from spica located in Spica_I to a excel file located in file results, named data_Spica_dogodki

'Get Python path, depending on the computer we are running this script from
Set fso = CreateObject("Scripting.FileSystemObject")

path_student = "C:/Users/student5/Anaconda/Python.exe"
path_home = "D:/OneDrive/Dokumenti/Python"
path_work = "C:/Users/slanad/OneDrive/Dokumenti/Python"
path_daniela = "C:/Users/bedernjakd/Documents"

If (fso.FileExists(path_home)) Then
	path = path_home
	WScript.Echo "Found for home"
ElseIf (fso.FileExists(path_work)) Then
	path = path_work
	WScript.Echo "Found for work"
ElseIf (fso.FileExists(path_daniela)) Then
	path = path_daniela
	WScript.Echo "Found for daniela"
ElseIf (fso.FileExists(path_student)) Then
	path = path_student
	WScript.Echo "Found for student"
Else
    WScript.Echo "Could not find path"
End If

Set winShell = CreateObject("WScript.Shell")
WaitOnReturn = False
windowStyle = 1

'Define the command to run the python file and exit when done
command1 = path & " TransformSpica.py" 
command2 = "exit"

'Run the commands
Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
