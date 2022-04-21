import win32com.client as win32
from pywintypes import com_error
import sys
win32c = win32.constants

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list,
                pt_filters: list, pt_fields: list,location:int):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """
    # pivot table location
    pt_loc = len(pt_filters) +location
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)
    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in (
    (pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1],
                                                field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True

    pt_ws.Cells(pt_loc, 2).Value=pt_cols[0]
    pt_ws.Cells(pt_loc+1, 1).Value = pt_rows[0]

def create_excel_pivot(filename, sheet_name: str,exclude_from_pivot:list):
    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # excel can be visible or not
    excel.Visible = True  # False
    # try except for file / path
    try:
        wb = excel.Workbooks.Open(filename)
    except com_error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
        else:
            raise e
        sys.exit(1)
    # set worksheet
    ws1 = wb.Sheets(sheet_name)
    # Get header columns
    headers=[]
    for i in range(1,ws1.UsedRange.Columns.Count):
        headers.append(ws1.Cells(1,i).Value)
    # Setup and call pivot_table
    ws2_name = 'pivot_table'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    pt_fields = [['EVENT_ID', 'Count of EVENT_ID', win32c.xlCount, '0']]
    count=2
    for pivot_cols in set(headers) - set(exclude_from_pivot):
        # Creating initial pivot table
        pt_name = f"{pivot_cols} Analysis "  # must be a string
        pt_rows = [pivot_cols]  # must be a list
        pt_cols = ['DECISION']  # must be a list
        pt_filters = []  # must be a list
        pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields,count)
        count+=len(get_unique_column_values(ws1,headers.index(pivot_cols)+1))+5
        # print(pivot_cols,count)
        # break
    # wb.Close(True)
    # excel.Quit()

def get_unique_column_values(ws: object,column_count:int):
    unique_cols=set()
    for i in range(2,ws.UsedRange.Rows.Count):
        unique_cols.add(ws.Cells(i,column_count).Value)
    # print(unique_cols)
    return unique_cols


if __name__=="__main__":
    groupby_col="DECISION"
    group_on_cols="EVENT_ID"
    exclude_from_pivot=['EVENT_ID','RECV_DT','EVENT_CREATED_DT','DECISION']
    # create_pivot(r"C:\Users\Ssaurabh\Documents\Release_New_Small.xlsx",0,groupby_col,group_on_cols,exclude_from_pivot)
    create_excel_pivot(r"C:\Users\Ssaurabh\Documents\Release_New_Small_1.xlsx","Release new",exclude_from_pivot)
