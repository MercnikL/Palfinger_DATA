# -*- coding: utf-8 -*-
"""
Created on Tue Dec 11 19:09:14 2018

@author: Slana
DODAM STOLPEC CC v datoteko (dolocim ga glede na stolpec tim)
"""

import os
import sys
import pandas as pd


# %% Common routines
# %% SET_PATH
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
# %% Transform

f_setPath("02_ZPRLAW") #Set working folder

try:
#arguments from cmd line: python data_1D.txt data_1D.csv
    sFile1 = sys.argv[1]
    sFileOut1 = sys.argv[2]
except:
    sFile1 = "data_120D.txt"
    sFileOut1 ="data_120D.csv"
    
sLogFile = "ED_log.txt"

#Read file
df = pd.read_csv(sFile1, 
                   sep='\t', 
                   skipinitialspace = True, 
                   encoding = 'latin2',
                   dialect = None,
                   index_col = False,
                   skip_blank_lines= True,
                   engine = 'c',  
                   skiprows = 1,
                   header = 1
                   )

df.drop('Unnamed: 0', axis=1, inplace=True) # ali kot df.drop(df.columns[[0]], axis=1, inplace=True)

def fSetCC(Tim):
    "Return CC according to TeamNm"
    CC = [["BF", 40, 41, 42, 43, 44],
          ["KS", 60, 62, 63, 64, 65, 66, 68],
          ["GG", 50, 51, 52, 53, 54, 55, 56],
          ["EPS", 30, 57, 58, 59, 67],
          ["GTC", 80, 76, 81, 82, 83, 84, 85],
          ["GKC", 70, 71, 72, 77, 78, 79, 61],
          ["KSR", 61],
          ["LAK", 91],
          ["MGK", 90, 92, 93, 94, 95, 96, 97, 98],
          ["WERK", 11],
          ["LOG", 12],
          ["EK", 13],
          ["CONTR", 14],
          ["PPS", 15],
          ["AT", 16],
          ["ST", 17],
          ["HR", 18],
          ["WZA", 19],
          ["INST", 20],
          ["QM", 21],
          ["IT", 22]]
    try:
        search = int(Tim[-2:]) #z desne dva znaka
    except:
        return('TeamNm Error')
        
    for sublist in CC:
        if search in sublist:
            return(sublist[0])           
            break

#doloci CC glede na tim
df['CC'] = df['TeamNm'].apply(lambda x: fSetCC(x))

#Zapisi rezultat v csv datoteko
df.to_csv(sFileOut1, mode='w',sep=';', encoding = 'latin2', header=True, index=False)    

#Zapisi porocilo v log
from datetime import datetime
sSedaj = datetime.now().strftime('%d.%m.%Y %H:%M:%S')

with open(sLogFile, "a") as myfile:
    myfile.write(sSedaj + "; transform.py ; " + sFile1 + " \n")
