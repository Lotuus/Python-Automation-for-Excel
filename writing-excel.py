from openpyxl import Workbook

wb = Workbook()
st = wb.active

st['a1'] = 'hello'
st['a2'] = 'world'

wb.save(filename='hello-world.xlsx')
