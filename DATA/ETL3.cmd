@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM PODATKE O VARILNI ŽICI IZ SAP"
REM Prenos iz SAP s transakcijo ZMATERIAL
CD .\03_VarilnaZica\
C:\Windows\SysWOW64\CScript.exe "ED_SAP_ZMATERIAL.vbs"
CD ..
ECHO '03_VarilnaZica'; %DATE%; %TIME% >>ETLcmd_log.txt

Echo "END"

REM Da ne zapre okna
timeout 10
pause