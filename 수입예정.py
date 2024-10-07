import xlwings as xw

app = xw.App(add_book=False)
wb = app.books.add()
ws = wb.sheets['sheet1']

ws.name = "수입예정"
ws.range('A1').value = ['제품명', '박스당수량', '단가', '현재고', '수입예정수량', '카톤수', '소계']
ws.range('A2').value = ['WH-B06', 40, 10, 13, 80, '=E2/B2', '=C2*E2']