@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM PODATKE O POSKODBAH IZ MIS10"
REM Podatki o poškodbah iz MIS10, standarni 32bitni cscript ne deluje na službenem raèunalniku z adodb.connect, zato ta verzija
CD .\01_OHS\
C:\Windows\SysWOW64\CScript.exe "ED.vbs"
CD ..
ECHO '01_OHS'; %DATE%; %TIME% >>ETLcmd_log.txt

ECHO "PRENASAM 01_Zaposleni_Kadrovska_ver02.xlsx"
xcopy "\\ssimarfile01\MIS10\Proizvodnja\BAZA\01_Zaposleni_Kadrovska_ver02.xlsx" /Y
ECHO '01_Zaposleni_Kadrovska_ver02.xlsx'; %DATE%; %TIME% >>ETLcmd_log.txt

Echo "END"

REM Da ne zapre okna
timeout 10
pause