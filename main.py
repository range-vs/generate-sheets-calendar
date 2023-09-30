import calendar
import datetime
import openpyxl
import locale

from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors


class CalendarExcelGenerator:

    def __init__(self):
        self.book = openpyxl.Workbook()
        self.column_width = 30
        self.headerFill = PatternFill(start_color='FF7d7777',
                    end_color='FF7d7777',
                    fill_type='solid')
        self.dateFill = PatternFill(start_color='FF6d87ab',
                    end_color='FF6d87ab',
                    fill_type='solid')
        self.cells = ["a", "b", "c", "d", "e", "f", "g"]
        self.weekdays = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
        self.mounts = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

    def createCellParams(self, sheet, cellIdent, fontSize, color):
        headerCell = sheet[cellIdent] 
        headerCell.alignment = Alignment(horizontal='center')
        headerCell.font = Font(size=fontSize)
        headerCell.fill = color

    def generateMount(self, indexMount):
        sheet = self.book.create_sheet(self.mounts[indexMount])
        rowColumn = 1
        # заголовок
        sheet.merge_cells(start_row=rowColumn, start_column=1, end_row=rowColumn, end_column=len(self.weekdays))
        self.createCellParams(sheet, "A"+str(rowColumn), 20, self.headerFill)
        sheet["A"+str(rowColumn)] = self.mounts[indexMount] + " " + str(datetime.datetime.now().year) + " года"
        rowColumn += 1

        # дни недели
        for i in range(0, 7):
            self.createCellParams(sheet, self.cells[i] + str(rowColumn), 16, self.headerFill)
            sheet.column_dimensions[get_column_letter(i+1)].width = self.column_width
            sheet[self.cells[i] + str(rowColumn)] = self.weekdays[i]
        rowColumn += 1

        # ячейки
        for week in calendar.monthcalendar(datetime.datetime.now().year, indexMount + 1):
            index = 0
            for day in week:
                self.createCellParams(sheet, self.cells[index] + str(rowColumn), 14, self.dateFill)
                if(day != 0):
                    date = datetime.datetime(datetime.datetime.now().year, indexMount + 1, day)
                    sheet[self.cells[date.weekday()] + str(rowColumn)]  = date.strftime("%d.%m")
                index += 1
            sheet.column_dimensions[get_column_letter(index + 1)].width = self.column_width
            self.createCellParams(sheet, "H" + str(rowColumn), 14, self.dateFill)
            sheet["H" + str(rowColumn)]  = "Объём за неделю"
            rowColumn += 1
            sheet.row_dimensions[rowColumn].height = 70
            rowColumn += 1

    def generateFile(self):
        for i in range(0, len(self.mounts)):
           self.generateMount(i) 
        self.book.remove(self.book["Sheet"])
        self.book.save('calendar_training.xlsx')

def main():
    locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
    generator = CalendarExcelGenerator()
    generator.generateFile()

if __name__ == "__main__":
    main()