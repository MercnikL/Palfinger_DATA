# %%PYWINAUTAO za powerBI export to PDF
from pywinauto import application
import subprocess
import time


def main():
    ime = input('>')
    
    fToPDF(ime)

def fToPDF(fileName):

    #p = subprocess.Popen('WHERE pbi*.exe', stdout=subprocess.PIPE, shell=False)
    p = subprocess.Popen('WHERE /R "c:\Program Files" "pbi*.exe"', stdout=subprocess.PIPE, shell=False)
    
    (output, err) = p.communicate()
    
    if not err:
        exeFile = str(output.strip(), 'utf-8')
        dataFile = r"\\ssimarfile01\CC\CF_WERK\TMP\Studenti\2018_Leon\Vizualizacija\PBI" + fileName+ ".pbix"
        
        app = application.Application(backend="uia").start(exeFile +" " + dataFile, timeout=30)
        app.wait_cpu_usage_lower(threshold=10) # wait until CPU usage is lower than 10%
        time.sleep(10) 
        #dlg = app.top_window
        dlg = app.window(title_re=".*Power BI.*")
        #app.dlg.print_control_identifiers()
        
        osvezi = dlg.window(title_re=".*Osveži.*")
        osvezi.click()
        
        time.sleep(10)
        
        datoteka_button = dlg.window(title_re=".*vihek datote.*")
        datoteka_button.click()
        wnd = app.top_window()
        export_menu = wnd.window(title_re=".*Izvozi v PDF.*")
     
        export_menu.select()
        time.sleep(5) 
        
        exit_menu = wnd.window(title_re=".*Izhod.*")
        
        datoteka_button.click()
        exit_menu.select()
        
        time.sleep(1) 
        app.Dialog.Shrani.click()
        
        time.sleep(5) 
        
        PDFapp = application.Application(backend="uia").connect(title_re=".*crobat.*")
        
        time.sleep(3) 
        
        PDFapp.Dialog.menu_select("Datoteka->Shrani kot ... Shift+Ctrl+S")
        
        time.sleep(1) 
        PDFapp.Dialog.ShraniKot.Shrani.click()
        
        time.sleep(1) 
        PDFapp.Dialog.Da.click()
        
        time.sleep(2)
        PDFapp.Dialog.menu_select("Datoteka->Izhod Ctrl+Q")
    
    else:
        print(output, err)  
    
    
    print('konec')
    
myStart = True
if (__name__=='__main__' and myStart == True):
    main()
