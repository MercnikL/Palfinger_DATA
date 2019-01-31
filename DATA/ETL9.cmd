@ECHO OFF
Echo "START EXTRACT TRANSFER AND LOAD"

ECHO
ECHO "PRENASAM PODATKE O ZAPOSLENIH 01_Zaposleni_Kadrovska_ver02.xlsx "
REM 
CD .\09_Zaposleni_kadrovska\

ECHO "PRENASAM 01_Zaposleni_Kadrovska_ver02.xlsx"
xcopy "\\ssimarfile01\MIS10\Proizvodnja\BAZA\01_Zaposleni_Kadrovska_ver02.xlsx" /Y
CD ..

ECHO '09_Zaposleni_kadrovska'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END"

REM Da ne zapre okna
timeout 10
pause