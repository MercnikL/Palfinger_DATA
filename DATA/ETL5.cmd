@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM PODATKE O ZPRMA1 IZ SAP"
REM Prenos iz SAP s transakcijo ZPRMA1
CD .\05_Spica\
C:\Windows\SysWOW64\CScript.exe "ED_Spica_dogodki.vbs"
CD ..
ECHO '05_Spica'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END"

REM Da ne zapre okna
timeout 10
pause