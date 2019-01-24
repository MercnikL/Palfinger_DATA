
	Set winShell = CreateObject("WScript.Shell") 'Creating a cmd object
	WaitOnReturn = False
	windowStyle = 1
    command1 = "taskkill /F /IM saplogon.exe " 'Defining command to kill the program
    command2 = "exit " 'Defining command to exit cmd
	
    Call winShell.Run("cmd /k " & command1 & " & " & command2, windowStyle, WaitOnReturn)
	'Running the commands in cmd
        
    'Release pointer to the command prompt.
    Set winShell = Nothing     'Release pointer to the command prompt.