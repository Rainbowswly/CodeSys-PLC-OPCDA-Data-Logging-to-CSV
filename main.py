import OpenOPC
import pywintypes # To avoid timeout error
pywintypes.datetime=pywintypes.TimeType
import pandas as pd
from datetime import date
import shutil
import os
import csv
from csv import writer


if os.path.isfile(f'Data\Date_{date.today()}.csv') == False:
    src= r'Template\Template.csv'
    dst = f'Data\Date_{date.today()}.csv'
    shutil.copyfile(src, dst)
else:
    pass


opc = OpenOPC.client()
opc.connect('CoDeSys.OPC.DA')


tags = ['MCR.Application.GVL.WMS.GII',\
        'MCR.Application.GVL.WMS.GHI',\
        'MCR.Application.GVL.WMS.MOD_TEMP',\
        'MCR.Application.GVL.WMS.AMB_TEMP',\
        'MCR.Application.GVL.WMS.WS',\
        'MCR.Application.GVL.WMS.WD',\
        'MCR.Application.GVL.WMS.RAIN',\
        'MCR.Application.GVL.WMS.HUMIDITY',\
        'MCR.Application.GVL.PLANT_RUN_INV',\
        'MCR.Application.GVL.PLANT_KW',\
        'MCR.Application.GVL.PLANT_KW',\
        'MCR.Application.GVL.PLANT_TODAY_KWH',\
        'MCR.Application.GVL.PLANT_EXP_MWh',\
        'MCR.Application.GVL.PLANT_TOTAL_INV',\
        'MCR.Application.GVL.ACTIVE_SETPOINT',\
        'MCR.Application.GVL.WMS.AVG_GII',\
        'MCR.Application.GVL.WMS.AVG_GHI',\
        'MCR.Application.GVL.AVG_KW',\
        'MCR.Application.GVL.AVG_TOTAL_KWH',\
        'MCR.Application.GVL.PLANT_CODE']



DF = pd.DataFrame(opc.read(tags, group='test'), columns = ['TAG_NAME','VALUE','QUALITY','TimeStamp'])


Record = DF.VALUE.to_list()
Record.insert(0, DF.TimeStamp.to_list()[0])


with open(f'Data\Date_{date.today()}.csv', 'a', newline='') as f_object:  
    writer_object = writer(f_object)
    writer_object.writerow(Record)  
    f_object.close()
