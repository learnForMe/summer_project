import openpyxl
import datetime
from datetime import timedelta
from calendar import mdays
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import coordinate_from_string
from openpyxl.styles import Font


def formular():
    wb=load_workbook('testing.xlsx',read_only = False, data_only = True)
    worksheet= wb.get_sheet_names()
    wb.active
    sheet = wb.get_sheet_by_name('Sheet')
    if "Monthly_STAT" not in worksheet:
        wb.create_sheet (index=2, title="Monthly_STAT")
        monthly_STAT=wb.get_sheet_by_name('Monthly_STAT')
        monthly_STAT.column_dimensions['A'].width = 25
        max_col=sheet.max_column-1
        max_row = sheet.max_row
        stat_row =monthly_STAT.max_row
        formular_col=get_column_letter(max_col)
        monthly_STAT['b%s' % stat_row]= '=SUM(Sheet!%s2:%s%d)' % (formular_col,formular_col,max_row)

        wb.save('testing.xlsx')

    else:
        monthly_STAT=wb.get_sheet_by_name('Monthly_STAT')
        monthly_STAT.column_dimensions['A'].width = 25
        max_col=sheet.max_column-1
        max_row = sheet.max_row
        stat_row =monthly_STAT.max_row
        formular_col=get_column_letter(max_col)
        monthly_STAT['b%s' % stat_row]= '=SUM(Sheet!%s2:%s%d)' % (formular_col,formular_col,max_row)
        wb.save('testing.xlsx')




