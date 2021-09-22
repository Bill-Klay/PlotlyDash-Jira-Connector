import sys
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open("D:\My Projects\Jira Connector\Jira Connector Application\Output\{}".format(str(sys.argv[1] + ' ' + sys.argv[2] + ' ' + sys.argv[3])))
ws = wb.Worksheets("Delivery Report")
ws.Columns.AutoFit()
ws = wb.Worksheets("Jira Data")
ws.Columns.AutoFit()
wb.Save()
excel.Application.Quit()