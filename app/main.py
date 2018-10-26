# -*- coding: utf-8 -*-
__author__ = 'BiziurAA'
from class_Sheets import Sheets
from config import Config
import os

if __name__ == "__main__":
    sheets=Sheets()
    setting = Config.inst()
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    work_path_project=os.path.join(BASE_DIR, setting.version[2])

    A=sheets.read_sheet_by_excel(work_path_project=work_path_project, workSheets=setting.version[3], nrows=36,
                                 usecols=setting.version[0])
    B = sheets.read_sheet_by_excel(work_path_project=work_path_project,
                                   workSheets=setting.version[4],nrows=None,usecols='A:E')

    # print(B['Старый адрес пультовой'])
    pd_C=A.merge(B, left_on=setting.version[5], right_on='Старый адрес пультовой',how='left')
    pd_D=A.merge(B, left_on=setting.version[6], right_on='Старый адрес пультовой', how='left')
    #
    sheets.write_to_excel(pd_C=pd_C[['Верхушка','Пульт']],startrow=2, index=0,file_name=work_path_project,
                          workSheets=setting.version[3],startcol=setting.version[5]-3)
    sheets.write_to_excel(pd_C=pd_D[['Пульт.1','Верхушка.1']], startrow=2, index=0, file_name=work_path_project,
                          workSheets=setting.version[3], startcol=setting.version[5] + 1)
    print(pd_C[['Верхушка','Пульт',setting.version[5]]])
