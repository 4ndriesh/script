# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys
# import xlwings as xw
from openpyxl import load_workbook, Workbook
# from win32com import client
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles.borders import Border, Side
from logging_err import *
from config import Config


class Sheets():

    Data_for_record = dict()
    # Шаблон записи по группам
    choose_param = list()
    # зачитка листа excel
    # --------------------------------------------------------------------------------------------------------------
    def read_sheet_by_excel(self,workSheets=None,work_path_project=None, nrows=None, usecols=None):
        try:
            setting = Config.inst()
            DataFrame_Sheet = pd.read_excel(work_path_project,
                                                 sheet_name=workSheets,
                                                 usecols=usecols, nrows=nrows, skiprows=[1])

        except Exception as e:
            exeption_print(e)
            return 0
        return DataFrame_Sheet
    # --------------------------------------------------------------------------------------------------------------
    def write_to_excel(self, pd_C=None, startrow=0, index=0, file_name=None, startcol=None, workSheets=None):

        mv = load_workbook(file_name)
        with pd.ExcelWriter(file_name, engine='openpyxl') as write_to_report:
            write_to_report.book = mv
            write_to_report.sheets = dict((ws.title, ws) for ws in mv.worksheets)
            pd_C.to_excel(write_to_report,
                                                     sheet_name=workSheets,
                                                     startrow=startrow,
                                                     startcol=startcol,
                                                     header=False,
                                                     index=index)
            write_to_report.save()
