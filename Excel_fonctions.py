from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import pandas as pd
import os

#----------------------------------------------------------------------------------------------------------------------------------------------
#Import CSV file and copy it into CSV2Table Excel file.
#Arguments: Log_to_import : This is the number of the log & the name of the CSV file, Please ensure that all the Logs names are numbers
#----------------------------------------------------------------------------------------------------------------------------------------------
def Import_CSV(Log_to_import):
    path = os.path.dirname(r'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Excel (1)\Excel/Excel/Logs/')
    csv_file = os.path.join(path, str(Log_to_import) + '.csv')
    read_file = pd.read_csv(csv_file)
    read_file.to_excel(
        r'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/CSV2Table.xlsx',
        index=None, header=True)
#----------------------------------------------------------------------------------------------------------------------------------------------
#Fill a cell with data, data can be everything
#Arguments: Row : Line of the cell
#Arguments: Col : Column of the cell
#Arguments: Value : The data you want to put in the Cell
#Arguments: Sheet : the "Sheet" name
#Arguments: File : the "File" name
#----------------------------------------------------------------------------------------------------------------------------------------------

def FillCell(Row,Col,Value,Sheet,File):
    path=os.path.dirname(r'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/')
    Xl_File = os.path.join(path, str(File) + '.xlsx')
    wb = load_workbook(Xl_File)
    ws = wb[str(Sheet)]
    char = get_column_letter(Col)
    ws[char + str(Row)].value = Value
    wb.save(Xl_File)

#----------------------------------------------------------------------------------------------------------------------------------------------
#Get date from a cell
#Arguments: Row : Line of the cell
#Arguments: Col : Column of the cell
#Arguments: Sheet : the "Sheet" name
#Arguments: File : the "File" name
#----------------------------------------------------------------------------------------------------------------------------------------------
def GetCellValue(Row, Col,Sheet,File):
    path=os.path.dirname(r'C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/')
    Xl_File = os.path.join(path, str(File) + '.xlsx')
    wb = load_workbook(Xl_File)
    ws = wb[str(Sheet)]
    char = get_column_letter(Col)
    return ws[char+str(Row)].value
#----------------------------------------------------------------------------------------------------------------------------------------------
#Calculate the number of cells containing data in a Row/Column
#Arguments: df : Keep it df, nothing to enter
#Arguments: Sheet name : "Sheet name"
#----------------------------------------------------------------------------------------------------------------------------------------------
def get_max_row_column(df, sheet_name):
    global max_row
    global max_col
    max_row = 1
    max_col = 1
    for sh_name, sh_content in df.items():
        if sh_name == sheet_name:
            max_row = len(sh_content) + 1
            max_col = len(sh_content.columns)
            break
    coordinates = {'max_row': max_row, 'max_col': max_col}
    return coordinates
df = pd.read_excel('C:/Users/oumayma.azennoud/OneDrive - Eurofins Digital Testing International/Desktop/Python/CSV2Table.xlsx', sheet_name=None)
max_row = get_max_row_column(df, 'Sheet1')['max_row']
max_col = get_max_row_column(df, 'Sheet1')['max_col']