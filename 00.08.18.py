import xlsxwriter

from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet


workbook = xlsxwriter.Workbook("Notebook-1.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1','NAME')
worksheet.write('B1', 'DEPARTMENT')
worksheet.write('C1', 'BARCOD')

row = 1
col = 0


class Main:
      def __init__(self):
          pass

      def menu(self):
        choose = input('1-add\n')
        if choose == '1':
            self.add()


      def add(self):
        Data = []
        for name, department, barcod in (Data):
                worksheet.write(row, col, name)
                worksheet.write(row, col +1, department)
                worksheet.write(row, col +2, barcod)
                Data.append(input('name:'))
                Data.append(input('department:'))
                Data.append(input('barcod:'))

row += 1
workbook.close()
if __name__ == "__main__":
  Main()




