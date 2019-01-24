@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM PODATKE O ZPRMA1 IZ SAP"
REM Prenos iz SAP s transakcijo ZPRMA1
CD .\04_ZPRMA1\
C:\Windows\SysWOW64\CScript.exe "ED_SAP_ZPRMA1.vbs"
CD ..
ECHO '04_ZPRMA1'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END"

REM Da ne zapre okna
timeout 10
pause