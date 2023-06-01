# Author :    Oumayma Azennoud
# Date:       01/06/2023
# Function:   Create phases
# ----------------------------------------------------------------------------------------------------------------------------------------------
# Import fonctions
# ----------------------------------------------------------------------------------------------------------------------------------------------
import Excel_fonctions as EF
from Excel_fonctions import max_row, max_col
# ----------------------------------------------------------------------------------------------------------------------------------------------
# General libreries
# ----------------------------------------------------------------------------------------------------------------------------------------------
from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import pandas as pd
import os

# ----------------------------------------------------------------------------------------------------------------------------------------------
# Global Variables
# ----------------------------------------------------------------------------------------------------------------------------------------------
#CSV2TABLE-------------------
Col_time=1
Col_voltage=2
Col_current = 3
Col_power = 4
#Charging IOP----------------
Col_Nlog=3
Col_Type=4
Col_ChargingType=5
Col_USBVersion=6
Col_EDTRef=7
Col_TypeOfProduct=8
Col_Brand=9
Col_Model=10
Col_MaxCurrent=11
Col_Led=12
Col_PrechargeCrurrent=13
Col_PrechargeEndVoltage=14
Col_PrechargeDuration=15
Col_NormalChargeCurrent=16
Col_NormalChargePower=17
Col_PrechargeExitDuration=18
Col_CurretAfterFlip=19
Col_PowerAfterFlip=20
Col_VoltageNormalCharge=21
Col_TotalChargeDuration=22
Col_PassRemarkFail=23
Col_Sample=24
Col_Cable=25
Col_Comment=26
#-------------------------------
Log_line_number=16
IOP_file="Template"
Table_file="CSV2Table"
precharge_start=0.05
precharge_end=0.2
# ----------------------------------------------------------------------------------------------------------------------------------------------
# Workspaces and worksheets
# ----------------------------------------------------------------------------------------------------------------------------------------------
wb1 = load_workbook(
    'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/Template.xlsx')
ws1 = wb1['Charging IOP']

wb2 = load_workbook(
    'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/CSV2Table.xlsx')
ws2 = wb2['Sheet1']
# ----------------------------------------------------------------------------------------------------------------------------------------------

EF.Import_CSV(14)
EF.Fill_IOP_row(22)


