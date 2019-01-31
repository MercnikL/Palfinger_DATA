@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM PODATKE IZ SIS10 - Q22_MS-TDM"
REM Podatki o poškodbah iz SIS10, standarni 32bitni cscript ne deluje na službenem raèunalniku z adodb.connect, zato ta verzija
CD .\06_SIS\
C:\Windows\SysWOW64\CScript.exe "ED_SIS.vbs"
CD ..
ECHO '06_SIS'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END"

REM Da ne zapre okna
timeout 10
pause