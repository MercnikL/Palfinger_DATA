@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM URE ZPRLW iz SAP"
REM Prenos iz SAP s transakcijo ZPRLAW in potem transformacija s Transform.py
CD .\02_ZPRLAW\
C:\Windows\SysWOW64\CScript.exe "ED_SAP_ZPRLAW.vbs"
CD ..
ECHO '02_ZPRLAW'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END"

REM Da ne zapre okna
timeout 10
pause