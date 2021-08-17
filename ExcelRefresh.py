#
# Open an existing workbook
#
import win32com.client as win32
import datetime 
#start ab excel process
excel = win32.gencache.EnsureDispatch('Excel.Application')
#open workbook
wb = excel.Workbooks.Open(r'C:\Users\Callu\OneDrive\Documents\Python Scripts\Data Analysis\Excel-Automation\Grades2.xlsx')

#
# Refresh All Data connections + Save 
#
wb.RefreshAll()
wb.SaveAs(r'C:\Users\Callu\OneDrive\Documents\Python Scripts\Data Analysis\Excel-Automation\Grades2.xlsx')
excel.Application.Quit()
