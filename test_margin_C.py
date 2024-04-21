#import Openpyxl library
import openpyxl

#load workbook into python
wb = openpyxl.load_workbook('..\\door_formatted_C.xlsx')

#call the active worksheet, which is the first sheet in the workbook
ws = wb.active

#change the scale of the worksheet
ws.page_setup.scale = 65

#change the margin of the worksheet
ws.page_margins.left = 0.15
ws.page_margins.right = 0.15
ws.page_margins.top = 0.6
ws.page_margins.bottom = 0.4

#save the workbook
wb.save('..\\door_cluster_C.xlsx')

