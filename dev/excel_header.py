from openpyxl import Workbook
import datetime
# Create workbooks
wb = Workbook()
ws = wb.active

month_dic = {'1':'januar','2':'februar','3':'marts','4':'april','5':'maj','6':'juni',
            '7':'juli','8':'august','9':'september','10':'oktober','11':'november','12':'december'}
currentMonth = datetime.datetime.now().month
currentYear = datetime.datetime.now().year

Header_center_text = "Voxmeter A/S"
Header_center_size = 20
Header_right_text = "Side &[Page] af &N"
ws.oddHeader.center.text = Header_center_text
ws.oddHeader.center.size = Header_center_size
ws.oddHeader.right.text = Header_right_text
ws.evenHeader.center.text = Header_center_text
ws.evenHeader.center.size = Header_center_size
ws.evenHeader.right.text = Header_right_text
Footer_center_text = "Udarbejdet af Voxmeter A/S %s %i"%(month_dic[str(currentMonth)], currentYear)
ws.oddFooter.center.text = Footer_center_text
ws.evenFooter.center.text = Footer_center_text

# Write header
ws.append(['Hello', 'World!'])

#print(dir(ws))

wb_fname = 'excel_header.xlsx'
wb.save(wb_fname)
print("Workbook saved!\n")
print(wb_fname)
    