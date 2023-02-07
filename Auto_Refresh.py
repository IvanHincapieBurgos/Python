import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

wb = excel.Workbooks.Open('C:/Users/ivanj/Desktop/Data_Projets/PQS.xlsx')

# selecciona la hoja de trabajo que contiene el Power Query
ws = wb.Worksheets("Sheet1")

# actualiza la query
ws.ListObjects("super").QueryTable.Refresh()

# guarda y cierra el libro
wb.Save()
wb.Close()

# cierra Excel
excel.Quit()