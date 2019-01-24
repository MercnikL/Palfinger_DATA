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

Echo "END 1"

ECHO
ECHO "PRENASAM URE ZPRLW iz SAP"
REM Prenos iz SAP s transakcijo ZPRLAW in potem transformacija s Transform.py
CD .\02_ZPRLAW\
C:\Windows\SysWOW64\CScript.exe "ED_SAP_ZPRLAW.vbs"
CD ..
ECHO '02_ZPRLAW'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END 2"

ECHO
ECHO "PRENASAM PODATKE O VARILNI ŽICI IZ SAP"
REM Prenos iz SAP s transakcijo ZMATERIAL
CD .\03_VarilnaZica\
C:\Windows\SysWOW64\CScript.exe "ED_SAP_ZMATERIAL.vbs"
CD ..
ECHO '03_VarilnaZica'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END 3"


ECHO
ECHO "PRENASAM PODATKE O ZPRMA1 IZ SAP"
REM Prenos iz SAP s transakcijo ZPRMA1
CD .\04_ZPRMA1\
C:\Windows\SysWOW64\CScript.exe "ED_SAP_ZPRMA1.vbs"
CD ..
ECHO '04_ZPRMA1'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END 4"


ECHO
ECHO "PRENASAM PODATKE O ZPRMA1 IZ SAP"
REM Prenos iz SAP s transakcijo ZPRMA1
CD .\05_Spica\
C:\Windows\SysWOW64\CScript.exe "ED_Spica_dogodki.vbs"
CD ..
ECHO '05_Spica'; %DATE%; %TIME% >>ETLcmd_log.txt


Echo "END 5"

REM Da ne zapre okna
timeout 10
