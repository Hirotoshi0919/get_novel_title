import openpyxl


class WriteExcel:
    @staticmethod
    def main():
        wb = openpyxl.Workbook()
        ws = wb.active

        # ws.cell(row=count, column=1).value = s.text
        # ws.cell(row=count, column=2).value = s.get_attribute("href")
        # ws.cell(row=count, column=3).value = len(s.text)

        wb.save("Book1.xlsx")
