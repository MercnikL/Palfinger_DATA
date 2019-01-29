from pywinauto import Application

try:
    
    #Finds the pop-up window
    app = Application().connect(title_re=".*večkratni.*") 
    
    #Sets the window as a specific element in the whole dialog
    groupText = app.window(best_match = "Dialog0").child_window(best_match='Afx:67020000:b')

    #Selects the element
    groupText.click()

    #Sends TAB key and UP key, so we select the right option
    groupText.type_keys('{TAB}')
    groupText.type_keys('{UP}')

    #Wait's so there is no problems
    #time.sleep(1)

    #Clicks the checkmark
    app.window(best_match = "Dialog").child_window(best_match='2').click()

except:
    #If there is no window found
    print("No Window Found")
    
    