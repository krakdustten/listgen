from enum import Enum

from datetime import datetime
import xlsxwriter as xlsx
import csv

class file(Enum):
    CSV = (1, ".csv")
    XLS = (2, ".xls")
    XLSX = (3, ".xlsx")

    def __init__(self, number, extension):
        self.number = number
        self.extension = extension

    @staticmethod
    def from_str(label):
        if label in ('svc', '.svc'):
            return file.SVC
        elif label in ('xls', '.xls'):
            return file.XLS
        elif label in ('xlsx', '.xlsx'):
            return file.XLSX
        return None

class workbookWriter:
    def __init__(self, fileType, filename, num_format="€#,##"):
        now = datetime.now()
        self.fileType = fileType
        self.fileName = "files/" + filename + now.strftime("-%m-%d-%Y,%H-%M-%S") + fileType.extension
        self.name = filename
        if fileType == file.CSV:
            raise NotImplementedError
        if fileType == file.XLS:
            raise NotImplementedError
        if fileType == file.XLSX:
            self.workbook = xlsx.Workbook(self.fileName)
            self.worksheet = self.workbook.add_worksheet()
            self.format_bold = self.workbook.add_format({'bold': True})
            self.format_money = self.workbook.add_format({'num_format': num_format})
            return

    def writeCell(self, x, y, data, format=""):
        if self.fileType == file.CSV:
            raise NotImplementedError
        if self.fileType == file.XLS:
            raise NotImplementedError
        if self.fileType == file.XLSX:
            if format == "":
                self.worksheet.write(y, x, data)
            else:
                self.worksheet.write(y, x, data, format)

    def writeCellFormula(self, x, y, formula="", format=""):
        if self.fileType == file.CSV:
            raise NotImplementedError
        if self.fileType == file.XLS:
            raise NotImplementedError
        if self.fileType == file.XLSX:
            if format == "":
                self.worksheet.write_formula(y, x, formula)
            else:
                self.worksheet.write_formula(y, x, formula, format)

    def setFormatColorScale(self, x, y, width=1, height=1, min_value=1, mid_value=20, max_value=100):
        if self.fileType == file.CSV:
            raise NotImplementedError
        if self.fileType == file.XLS:
            raise NotImplementedError
        if self.fileType == file.XLSX:
            self.worksheet.conditional_format(y, x, y + height - 1, x + width - 1,
                                              {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num',
                                               'max_value': min_value, 'mid_value': mid_value, 'min_value': max_value,
                                               'min_color': "green", 'mid_color': "yellow", 'max_color': "red"})

    def close(self):
        if self.fileType == file.CSV:
            raise NotImplementedError
        if self.fileType == file.XLS:
            raise NotImplementedError
        if self.fileType == file.XLSX:
            self.workbook.close()


def intToCol(i):
    string = ""
    string = str(chr(65 + int(i % 26))) + string
    i = int(i / 26)
    if i > 0:
        string = str(chr(64 + int(i % 26))) + string
    return string
