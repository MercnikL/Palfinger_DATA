# -*- coding: utf-8 -*-
"""
Created on Fri Jan  4 11:41:44 2019

@author: slanad

Vhodni podatki so:
     Vhodna datoteka je iz SPICE, porocilo-dnevno-dogodki in nadure
"""

#%reset -f

# %% Common routines
# %% SET_PATH
import os 
import sys
import pandas as pd
from datetime import datetime
import re

def f_setPath(folder):
    "Set or create working folder"
    
    path_home = "D:/OneDrive/Dokumenti/Python"
    path_work = "C:/Users/slanad/OneDrive/Dokumenti/Python"
    path_daniela = "C:/Users/bedernjakd/Documents"
    #path_student = "//ssimarfile01/CC/CF_WERK/TMP/Studenti/2018_Leon/DATA"
    path_studentDesktop = "C:/Users/student5/Desktop/Palfinger_DATA/DATA"
    

    if os.path.exists(path_home):
        path = path_home
    elif os.path.exists(path_work):
        path = path_work
    elif os.path.exists(path_daniela):
        path = path_daniela
    elif os.path.exists(path_studentDesktop):
        path = path_studentDesktop
    #elif os.path.exists(path_student):
    #    path = path_student

    else:
        # print("NO CORRECT PATH")
        return("No path found")

    os.makedirs(os.path.join(path, folder), exist_ok=True)
    os.chdir(os.path.join(path, folder))
    # print('Working folder "pwd":', os.getcwd())

    del((folder, path_home, path_work, path))
    return(os.getcwd())
    
def isNotEmpty(s):
    # check if string is empty
    #print isNotEmpty("")    # False
    #print isNotEmpty("   ") # False
    #print isNotEmpty("ok")  # True
    #print isNotEmpty(None)  # False
    if isinstance(s, str): #if it is string
        return bool(s and s.strip())
    else:
        return False

# %% Commmon 2

def f_validate_date(date_text):
    "Preverim ali gre za datum v obliki kot jo pricakujem"
    #need import datetime
    try:
        if not datetime.strptime(date_text, "%d.%m.%Y"):
            raise ValueError
        return True
    except ValueError:
        return False

def f_int(s):
    "vrne stevilko ali 0, ce ni stevilke"
    try:
        int(s)
        return int(s)
    except ValueError:
        # Handle the exception
        return 0

def f_Zaokrozi_na_30min(s_ura):
    "Zaokrozim uro na 30 minut. Mora biti v pravem formatu"
#need import re !
#zaokrozi uro v string formatu "12:12" na 12:30, ali 12:31 na 13:00. 12:00 ostane 12:00 
    if (s_ura.find(':') != -1):
        try:
            hhmm = list(map(int, re.findall('\d+', s_ura)))  # poisci vse stevilke v tekstu
            if (len(hhmm)!=2): # ce ni ura ampak nekaj kaj ima : vmes 
                return s_ura
            if ( hhmm[1] > 0) and (hhmm[1] <= 29) : 
                rez = str(hhmm[0]) + ':' + '30'
            elif hhmm[1] >= 30 :
                rez = str((hhmm[0]+1)) + ':' + '00'
            elif  hhmm[1] >= 0 : 
                rez = str(hhmm[0]) + ':' + '00'
            #return s[poz-2:poz]
        except ValueError:
            rez = s_ura
    else:
        rez = s_ura      
    return rez


def fUraDec(s):
    #vstopni format so ure in minute: s="10:34", vrne pa decimalni zapis v urah npr. 10.57 (double)
    #ob napaki vrnem 0
    if not isinstance(s, str): #if it is not string
        return 0
    else:
        try:
            if s[:1] == '-' : #ce je prvi znak minus
                sign = -1
                s = s[1:] #odrezem prvi znak
            elif s[:1] == '+' : #ce je prvi znak plus
                sign = 1
                s = s[1:] #odrezem prvi znak
            else:
                sign = 1
            
            t1 = datetime.strptime(s, '%H:%M')
            return sign*(t1.hour + t1.minute/60)
        except:
            return 0

def fRazlikaH(t1, t2):
    #izracuna razliko v urah npr. od 6:00 do 14:00 je 8ur in 22:00 do 6:00 je 8 ur
    #vstop in izstop so decimalna stevila: 6.00 - 14.75 = 8.75
    #ob napaki vrnem 0
    try: 
        if t2 > t1 :
            return (t2 - t1)
        elif t2 == t1:
            return 0
        else:
            return (24 - t1 + t2)
    except:
        return 0

# =============================================================================
# def fNocneH(t1, t2, ts1 = 0.00, te1 = 6.00, ts2 = 22.00, te2 = (22.00+8)):
#     #'''izracuna stevilo nocnih ur npr. od 22:00 do 8:00 je 8 nocnih ur.
#     #vstop in izstop so decimalna stevila: : 6.00 - 14.75 -> 0
#     #ob napaki vrnem 0
#     #na enak nacin lahko izracunas tudi druge ure, samo premaknes interval
#     #'''
#     try: 
#         #Ce je 24 ur - ni možno
#         if t1 == t2:
#             return 0
#         tn=0
#         if t1 >= t2: t2 = t2 + 24 #prekoracitev polnoci    
#         #t1:I, t2: I
#         if (t1 >= ts1 and t1 < te1) and (t2 > ts1 and t2 <= te1): 
#             tn = t2 - t1
#         #t1:I, t2: II
#         if (t1 >= ts1 and t1 < te1) and (t2 > te1 and t2 <= ts2):
#             tn = te1 - t1
#         #t1:I, t2: III
#         if (t1 >= ts1 and t1 < te1) and (t2 > ts2 and t2 <= te2):
#             tn =(te1 - t1) + (t2 - ts2)   
#         #t1:I, t2: IV
#         if (t1 >= ts1 and t1 < te1) and (t2 > te2 and t2 <= 48):
#             tn =(te1 - t1) + (te2 - ts2)
#         #t1:II, t2: II
#         if (t1 >= te1 and t1 < ts2) and (t2 > te1 and t2 <= ts2):
#             tn = 0 
#         #t1:II, t2: III
#         if (t1 >= te1 and t1 < ts2) and (t2 > ts2 and t2 <= te2):
#             tn = t2 - ts2      
#         #t1:II, t2: IV
#         if (t1 >= te1 and t1 < ts2) and (t2 > te2 and t2 <= 48):
#             tn = te2 - ts2
#         #t1:III, t2: III
#         if (t1 >= ts2 and t1 < te2) and (t2 > ts2 and t2 <= te2):
#             tn = t2 - t1  
#         #t1:III, t2: IV
#         if (t1 >= ts2 and t1 < te2) and (t2 > te2 and t2 <= 48):
#             tn = te2 - t1
#         #t1:IV, t2: IV
#         #ni mozen
#         return tn
#     except:
#         return 0
# 
# =============================================================================
        
def fInterval(t1, t2):
    '''izracuna stevilo ur po intervalih 
    Vstop: zacetni cas, koncni cas npr. 10.0 14.0
    Rezultat: tuple (4,0,0)
    '''
    ure_po_izmenah = [0,0,0]
    interval_ime = [3, 1, 2, 3, 3, 1, 2, 3]
    interval = [
            [0, 6],
            [6, 14],
            [14, 22],
            [22, 24],
            [24, 6+24],
            [6+24, 14+24],
            [14+24, 22+24],
            [22+24, 24+24]
            ]  
    interval_znotraj = [0, 0, 0, 0, 0, 0, 0, 0]

    try: 
        #Ce je 24 ur - ni možno
        if t1 == t2:
            return [0,0,0] 
   
        if t1 >= t2: t2 = t2 + 24 #prekoracitev polnoci    

        for i in range(0, len(interval)):
            ts = interval[i][0]
            te = interval[i][1] 
            if (t1 < ts and t2 < ts):
                tn = 0
            elif (t1 >= ts and t1 < te) and (t2 <= te ):
                tn = t2 - t1
            elif (t1 > te):
                tn = 0
            elif (t1 < ts) and (t2 > ts and t2 <= te):
                tn = t2 - ts
            elif (t1 >= ts and t1 < te) and t2 > te:
                tn = te - t1
            elif (t1 <= ts) and (t2 > te):
                tn = te - ts
            else:
                tn = 0

            interval_znotraj[i] = tn    
                    
        for i in range(0, len(interval)):
            j = interval_ime[i]-1
            ure_po_izmenah[j] += interval_znotraj[i]
    
        return ure_po_izmenah
    except:
        return [0,0,0]   

# %% READ FILE LINE BY LINE
# Izdelaj izhodno urejeno datoteko.
# Vhodna datoteka je iz SPICE, porocilo-dnevno-dogodki in nadure

def f_Uredi_SpicaTXT(datoteka_vhodna, datoteka_izhodna):
    "Vstopna datoteka je izvoz iz porocilo-dnevno-dogodki in nadure v txt obliki"
    
    priimek_ime = ''
    priimek_ime_n = '*'
    datum = ''
    datum_n = ''
    ms = 0  # maticna stevilka
    ms_n = -1 
    m = 0   # stevec
    n = 0   # stevec
    cc = '' # CC
    cc_n = '*'
    lines = []
    prvi = False
    dogodki = ''
    dogodki_z = ''
    t = ['Odhod', 'Prihod',	'Zaokroževanje', 'Korekcija nadur', 'Zamenjava urnika',	'Korekcija poizmenske m', 'Sindikalno delo', 
                         'Kompenzacijski prihod', 'Plačan izhod',	'Službeni izhod brez pr', 'Službeni izhod',	
                         'Kompenzacijski izhod', 'Korekcija poizmenske ', 'Službeni izhod brez p', 'Vnos dopusta',
                         'Bolniška do 30 dni 90', 'Refundacija bolniška ', 'redni dopust po urah', 
                         'Refunac.bol.90 dni po', 'Bolniška do 30 dni 10']
    tt = ''
    
    
    with open(datoteka_izhodna,'w') as fw:
        lines.append('Priimek, '+'MS, '+'CC, '+ 'm, '+'n, '+'Datum, '+'Dan, '+'Urnik, '+'Prisotnost ,'+'Plan, '
                     +'Nadure, '+'Saldo_nadur, '+'Dopust' +'\n')
        
        with open(datoteka_vhodna, 'r') as f:
            for line_terminated in f:
                s = line_terminated.rstrip('\n')
                # Vrstica s priimkom
                if 'Priimek, Ime' in s:
                    priimek_ime = s[26:62].strip()
                elif 'Evidenčna številka' in s:
                    ms = f_int(s[25:33].strip())
                elif 'Org. enota' in s:
                    cc = s[26:62].strip()
                # Vrstica z datumom
                elif f_validate_date(s[1:15].strip()) & (len(s[16:25].strip())>2) & (len(s)>30):
                    if priimek_ime_n != priimek_ime:
                        if (prvi == True):
                            lines.append(priimek_ime_n + ',' + str(ms_n) + ',' + cc_n + ',' + str(m) + ',' + str(n) + 
                                     ',' + datum_n + ',' + 'AKC' + ',' + str(dogodki) + '\n')
                            lines.append(priimek_ime_n + ',' + str(ms_n) + ',' + cc_n + ',' + str(m) + ',' + str(n) + 
                                     ',' + datum_n + ',' + 'AKC_Z' + ',' + str(dogodki_z) + '\n')
                            datum_n = datum
                            dogodki = ''
                            dogodki_z = ''
                        m = m + 1
                        n = 0
                        prvi = True
                        lines.append( '{},{},{},{},{}\n'.format(priimek_ime, str(ms), cc, str(m),str(n)))
                        priimek_ime_n = priimek_ime
                        ms_n = ms
                        cc_n = cc
                    n = n + 1
                    datum = s.split()[0]
                    #dogodki.append()
                    if n==1: 
                        datum_n = datum
                    if (datum != datum_n) & (prvi == True):
                        lines.append(priimek_ime + ',' + str(ms) + ',' + cc + ',' + str(m) + ',' + str(n-1) + 
                                     ',' + datum_n + ',' + 'AKC' + ',' + str(dogodki) + '\n')
                        lines.append(priimek_ime + ',' + str(ms) + ',' + cc + ',' + str(m) + ',' + str(n-1) + 
                                     ',' + datum_n + ',' + 'AKC_Z' + ',' + str(dogodki_z) + '\n')
                        datum_n = datum
                        dogodki = ''
                        dogodki_z = ''
                        
                    lines.append(priimek_ime + ',' + str(ms) + ',' + cc + ',' + str(m) + ',' + str(n) + ',' + ','.join(s.split()) + '\n' )
                    #print('   ' + s[1:15].strip())
                # Vrstica z dogodki
                elif (s[0:14] == 14 * ' ') & (len(s[15:35].strip())>0) & (s[15:70].strip() != 'urnika'):
                    s1 = s[15:23].strip()
                    s2 = s[28:52].strip()
                    s3 = s[53:].strip()
                    
                    s1x = f_Zaokrozi_na_30min(s1)
                    
                    if (s1 == '') and ((s2[:21]) in t):
                        tt = s2
                    elif ((tt!='') and (s2=='')):
                        lines.append(priimek_ime + ',' + str(ms) + ',' + cc + ',' + str(m) + ',' + str(n) + ',' + datum + ',' + s1 + ',' + tt + ',' + s3 + '\n')
                        dogodki = dogodki + ('->' + s1 + ':' + tt)
                        dogodki_z = dogodki_z + ('->' + s1x + ':' + tt)
                        tt = ''
                    else:
                        lines.append(priimek_ime + ',' + str(ms) + ',' + cc + ',' + str(m) + ',' + str(n) + ',' + datum + ',' + s1 + ',' + s2 + ',' + s3 + '\n')
                        dogodki = dogodki + ('->' + s1 + ':' + s2)
                        dogodki_z = dogodki_z + ('->' + s1x + ':' + s2)
    
                        
            fw.writelines(lines)
                
            #print("---KONEC---")
    return True


#%% Predelaj datoteko z dogodtki v drugacno obliko
    

def f_Pretvori_v_dataframe(datoteka_vhodna):
    '''VER2, VSTOP: Vstopna datoteka je urejen izvoz iz porocilo-dnevno-dogodki in nadure 
    v txt obliki, urejeno s pomocjo f_Uredi_SpicaTXT
    Izstop: DataFrame
    '''
    import pandas as pd
    #preberi vhodno datoteko v pandas dataframe. 
    #Vhodna datoteka je oblike, kot jo dobis iz txt_08.py ali iz f_Uredi_SpicaTXT()
    df = pd.read_csv(datoteka_vhodna, sep=',', index_col=None, encoding='cp1250', na_values=[" ", ""])

    #novi df2 v katerega pises rezultate
    columns = ['datum', 'MS', 'priimek', 'prihod', 'odhod', 'err', 'korekcija', 'urnik', 'prisotnost', 'plan', 'bolniska', 'dopust', 'razno' ]
    df2  = pd.DataFrame(columns=columns)
    dfA = pd.DataFrame(columns=columns)

    #iniciinicializiraj nekatere spremenenljivke
    n0 = 0          #spremlja ali smo ze skozi del dan dogodkov
    sPrihod = ''    #ura prihoda
    sOdhod = ''     #ura odhoda
    flagZ = False   #ali je bilo izvedeno zaokrozevanje
    flagOdhod = True #potrebujes, da oznacis, da si nasel polje odhod
    flagPrihod = False #oznacit, da se je prej pojavil prihod kot odhod. vse ostale odhode pred prihodom pozabis
    sErr = ''       #sem prepisem AKZ iz vhodne datoteke - seznam vseh dogodkov npr. 8:00:prihod->14:00:odhod
    korekcija = ''  #korekcije - zapisejo se k zadnjemu paru dogotkov prihod-odhod v dnevu
    nPrihod = 0     #steje stevilo odhodov v enem dnevu
    dnevi = ('pon', 'tor', 'sre', 'čet', 'pet', 'sob', 'ned')
    
    for index, row in df.iterrows():
        #skozi vse vrstice v df, kjer je kopija vhodne datoteke
        priimek = row[0]
        MS = row[1]
        #CC = row[2]
        #m = row[3]
        n = row[4]
        sDatum = row[5]
        dan = row[6] #tu se nahaja ura ali dan v tednu ali kode AKC ali AKC_Z
        urnik = row[7] #tu se nahajajo kode: prihod, odhodk, celoten AKC, zaokrozevanje, korekcije ...
        prisotnost = row[8]
        plan = row[9]
        
        B_R = ("olni", "efundacija")
        D_C = ("opust", "elodnevno")
        
        if urnik == 'Korekcija poizmenske mej': 
            korekcija = dan #v dan je vrednost korekcije npr. "+2:00"
        if dan in dnevi:
            sTipUrnika = urnik
            sPrisotnost = prisotnost
            sPlan = plan
            
        if n == n0: #dokler si znotraj enega dne (n šteje dneve)
            if urnik == 'Prihod' or urnik == 'Kompenzacijski prihod':
                flagZ = False
                flagPrihod = True
                sPrihod = dan 
            elif urnik == 'Zaokroževanje':
                flagZ = True
                sOdhod = dan
            elif (urnik == 'Odhod' or urnik == 'Kompenzacijski izhod') and flagPrihod == True: #ko najdem odhod ustvarim zapis, ki ga kasneje korigiram, ker se AKC vedno nahaj za zapisom odhod
                flagPrihod = False
                if flagZ == True:
                    flagZ = False
                else:
                    sOdhod = dan
                nPrihod += 1
                dfA.loc[0] = [sDatum, MS, priimek, sPrihod, sOdhod, sErr, korekcija, sTipUrnika, sPrisotnost, sPlan, "0:00", "0:00", "0:00" ]
                df2 = df2.append(dfA, ignore_index=True)
                flagOdhod = True
                korekcija = ''
                if nPrihod > 1: #korekcijo premaknem k zadnjemu zapisu v enem dnevu
                    df2['korekcija'].iloc[-1] = df2['korekcija'].iloc[-2]
                    df2['korekcija'].iloc[-2] = ''
            elif any(s in str(urnik) for s in B_R) and ("->" not in str(urnik)):
                dfA.loc[0] = [sDatum, MS, priimek, "0:00", "0:00", sErr, "0", sTipUrnika, sPrisotnost, sPlan, sPrisotnost, "0:00", "0:00" ]
                df2 = df2.append(dfA, ignore_index=True)   
            elif any(s in str(urnik) for s in D_C) and ("->" not in str(urnik)):
                dfA.loc[0] = [sDatum, MS, priimek, "0:00", "0:00", sErr, "0", sTipUrnika, sPrisotnost, sPlan, "0:00", sPrisotnost, "0:00" ]
                df2 = df2.append(dfA, ignore_index=True)
                    
        else: #tukaj korigiram zapis z AKC kodo
            if (df2.shape[0]>0) and (isNotEmpty(sErr)) and (flagOdhod==True): 
                df2['err'].iloc[-1] = sErr
                if nPrihod > 1:
                    df2['err'].iloc[-2] = df2['err'].iloc[-1]
                flagOdhod = False
            nPrihod = 0
            n0 = n
        
        if (dan == 'AKC') and (isNotEmpty(urnik)): sErr = urnik
    
    return df2

#%% MAIN
# Delovna mapa

def main():
    try:
        print(sys.version)
        f_setPath("05_Spica")
        try:
        #arguments from cmd line
            datoteka_vhodna = sys.argv[1]
            datoteka_izhodna = sys.argv[2]
            datoteka_izhodnaTest = sys.argv[3]

        except:
            datoteka_vhodna = './Spica_I.TXT'
            datoteka_izhodna = './results/data_Spica_dogodki.csv'
            datoteka_izhodnaTest = './data_Spica_dogodkiTest.csv'

        datoteka_tmp1 = './results/output_report.txt'  
        datoteka_uporabniki = "./01_Zaposleni_Kadrovska_ver02.xlsx"

        # Uvozi in uredi podatke iz SPICE in shrani urejeno datoteko
		# Import and sort data from SPICA and save the sorted file
		
        f_Uredi_SpicaTXT(datoteka_vhodna, datoteka_tmp1) #rezultat predelave shranim v tmp1 / Save the results of transformation into tmp1
        
        df2 = f_Pretvori_v_dataframe(datoteka_tmp1)#transform file tmp1 into a dataframe
        
        df2['H_start'] = df2['prihod'].apply(fUraDec) #zacetek dela v decimalni obliki / Start of work in decimal
        df2['H_end'] = df2['odhod'].apply(fUraDec) #konec dela v decimalni obliki / End of work in decimal
		
        df2['H_1'] = df2.apply(lambda x: fInterval(x.H_start, x.H_end)[0], axis=1) #prva izmena / First shift
        df2['H_2'] = df2.apply(lambda x: fInterval(x.H_start, x.H_end)[1], axis=1) #druga izmena / Second shift
        df2['H_3'] = df2.apply(lambda x: fInterval(x.H_start, x.H_end)[2], axis=1) #tretja izmena / Third shift
		
        df2['H_bolniska'] = df2['bolniska'].apply(fUraDec) #Bolniška / Hospital leave
        df2['H_dopust'] = df2['dopust'].apply(fUraDec) #Dopust / Vaccation
        df2['H_skupaj'] = df2['H_1'] + df2['H_2'] + df2['H_3'] + df2['H_bolniska'] + df2['H_dopust'] #Skupaj / Together
    
        #Združi po dnevih / Combine by days      
        zdruzi = ['datum',
                 'MS',           
                 'priimek',
                 'urnik',
                 ]
        
        df3 = df2.groupby(zdruzi)[['H_skupaj', 'H_1', 'H_2', 'H_3', 'H_bolniska', 'H_dopust']].agg('sum')
        df3.reset_index(level=df3.index.names, inplace=True) #Convert a given whole multiindex into columns 
        
        df3['OsemUr'] = 8 # 8 - dolzina rednega delovnega dne
        
        #Nadure / Overtime
        df3['H_nadure'] = df3['H_skupaj']-df3['OsemUr']        
        df3.loc[df3.H_nadure < 0, 'H_nadure'] = 0
        df3['H_minus'] = df3['H_skupaj']-df3['OsemUr']
        df3.loc[df3.H_minus > 0, 'H_minus'] = 0
        
        df_u = pd.read_excel(datoteka_uporabniki, sheet_name='DATA') #Open excel file containing workforce information
        df4 = pd.merge(df3, df_u, how='left', left_on=['MS'], right_on=['MS'] )

        zdruzi = [
                 'datum',
                 'MS',
                 'Tim',
                 'CC',
                 'Enote',
                 'Podjetje',
                 'Poklic'
                 ] 
        df5 = df4.groupby(zdruzi)[['H_skupaj', 'H_1', 'H_2', 'H_3', 'H_bolniska', 'H_dopust', 'H_nadure', 'H_minus']].agg('sum')
		
        df5['Stej'] = df4.groupby(zdruzi)['H_skupaj'].agg('count')
		
        df5.reset_index(level=df5.index.names, inplace=True) #Convert a given whole multiindex into columns 
        df5.to_csv(datoteka_izhodna, mode='w', sep=';', header=True, index=False, encoding='cp1250', decimal=',' ) #Save as an excel file
        
        print ('Made it till end')

    except ImportError as error:
        # Output expected ImportErrors.
        print(error.__class__.__name__ + ": " + error.message)
    except Exception as exception:
        # Output unexpected Exceptions.
        print(exception, False)
        print(exception.__class__.__name__ + ": " + exception.message)
 
 

#%% START
myStart = True
if (__name__=='__main__' and myStart == True):
    main()

   