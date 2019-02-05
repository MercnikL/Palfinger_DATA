from pywinauto import Application

try:
    
    #Finds the pop-up window
    app = Application().connect(title_re=".*Ribbon.*") 
    
    #Sets the window as a specific element in the whole dialog
    dlg = app.window(best_match = ".*Ribbon.*")
    app.dlg.print_control_identifiers()

    #Selects the element
    #dlg.click()

    #Sends TAB key and UP key, so we select the right option
    #dlg.type_keys('{TAB}')
    #dlg.type_keys('{UP}')

    #Wait's so there is no problems
    #time.sleep(1)

    #Clicks the checkmark
    #app.window(best_match = "Dialog").child_window(best_match='2').click()

except:
    #If there is no window found
    print("No Window Found")
    
    