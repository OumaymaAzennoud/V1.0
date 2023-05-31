# Author :    Oumayma Azennoud
# Date:       31/05/2023
# Function:   Get&Print the data from colum A=1, Fill the column B=2 with data
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


col_letter=1
for i in range (1,max_row+1,1):
    print(EF.GetCellValue(i,col_letter,"Sheet1","CSV2Table"))

col_letter=2
for i in range (1,max_row+1,1):
    EF.FillCell(i,col_letter,i,"Sheet1","CSV2Table")