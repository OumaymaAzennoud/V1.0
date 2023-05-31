# Author :    Oumayma Azennoud
# Date:       31/05/2023
# Function:   Create phases
#----------------------------------------------------------------------------------------------------------------------------------------------
#Import fonctions
#----------------------------------------------------------------------------------------------------------------------------------------------
import Excel_fonctions as EF
from Excel_fonctions import max_row,max_col
#----------------------------------------------------------------------------------------------------------------------------------------------
#General libreries
#----------------------------------------------------------------------------------------------------------------------------------------------
from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import pandas as pd
import os
#----------------------------------------------------------------------------------------------------------------------------------------------
#Global Variables
#----------------------------------------------------------------------------------------------------------------------------------------------
Col_current=3

#----------------------------------------------------------------------------------------------------------------------------------------------
wb = load_workbook(
    'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/CSV2Table.xlsx')
ws = wb['Sheet1']

Max_ph1=None
for i in range (1, max_row+1, 1):
    char = get_column_letter(Col_current)
    if ws[char+str(i)].value !=0 :
        EndPh1=i
        break
    data_ph1=0 if ws[char + str(i)].value == None else ws[char + str(i)].value
    if Max_ph1 is None or data_ph1 > Max_ph1:Max_ph1=data_ph1
EF.FillCell(8,9,Max_ph1,"Sheet1","CSV2Table")


Max_ph2=None
for i in range (EndPh1, max_row+1, 1):
    char = get_column_letter(Col_current)
    if ws[char+str(i)].value ==0 :
        EndPh2=i
        break
    data_ph2=0 if ws[char + str(i)].value == None else ws[char + str(i)].value
    if Max_ph2 is None or data_ph1 > Max_ph2:Max_ph2=data_ph2
EF.FillCell(8,9,Max_ph2,"Sheet1","CSV2Table")
print(EndPh2)