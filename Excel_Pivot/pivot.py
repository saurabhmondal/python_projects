import win32com.client as win32
from pywintypes import com_error
import sys

win32c = win32.constants
from config import *
import os


excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel can be visible or not
excel.Visible = True  # False

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list,
                pt_filters: list, pt_fields: list, location: int, position: int):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """
    # pivot table location
    pt_loc = len(pt_filters) + location
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C{position}', TableName=pt_name)
    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, position).Select()

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

    pt_ws.Cells(pt_loc, position + 1).Value = pt_cols[0]
    pt_ws.Cells(pt_loc + 1, position).Value = pt_rows[0]
    print(f"pt_loc:{pt_loc}, position:{position}")


def read_excel(filename, sheet_name="Sheet1"):
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
    ws = wb.Sheets(sheet_name)
    return wb, ws

def get_excel_col_index(num):
    threshold=(ord("Z")-ord("A")+1)
    if int(num)<=threshold:
        return chr(num + ord("A")-1)
    else:
        return get_excel_col_index(int(num/threshold))+get_excel_col_index(num%threshold)

def write_excel(filename, sheet_name="Sheet1",sheet_obj=None,req_cols=[]):
    if isinstance(filename,str):
        # try except for file / path
        try:
            wb = excel.Workbooks.Open(filename)
        except com_error as e:
            if e.excepinfo[5] == -2146827284:
                print(f'Existing file not found.  Hence creating the file: {filename}')
                wb = excel.Workbooks.Add()
    else:
        wb=filename
    wb.Sheets.Add().Name = sheet_name
    ws2 = wb.Sheets(sheet_name)
    if sheet_obj is not None:
        ws2 = wb.Worksheets(sheet_name)
        max_source_row_count=sheet_obj.UsedRange.Rows.Count
        if sheet_obj.FilterMode:
            sheet_obj.ShowAllData()
        if len(req_cols)==0:
            max_col_in_source=get_excel_col_index(sheet_obj.UsedRange.Columns.Count)
            sheet_obj.Range(f"A1:{max_col_in_source}{sheet_obj.UsedRange.Rows.Count}").Copy(ws2.Range(f"A1:{max_col_in_source}{sheet_obj.UsedRange.Rows.Count}"))
        else:
            destination_col_index=ord("A")
            for col_index in req_cols:
                sheet_obj.Range(f"{col_index}1:{col_index}{max_source_row_count}").Copy(ws2.Range(f"{chr(destination_col_index)}1:{chr(destination_col_index)}{max_source_row_count}"))
                destination_col_index+=1
    return wb, ws2


def copy_excel(source_file, source_sheet, destination_file="", destination_sheet="",req_cols=[]):
    if destination_file == "":
        destination_file = source_file
        if destination_sheet == "":
            destination_sheet = source_sheet + " copy"
    old_wb, old_ws = read_excel(source_file, source_sheet)
    req_cols_index = [ get_excel_col_index(i) for i in range(1, old_ws.UsedRange.Columns.Count) if old_ws.Cells(1, i).Value in req_cols]
    wb_new,ws_new=write_excel(destination_file, sheet_name=destination_sheet, sheet_obj=old_ws,req_cols=req_cols_index)
    return old_wb,wb_new,ws_new

def add_yes_percent(ws,unique_values,row,col):
    row =+ 1
    col+4
    ws.Cells(row, col).Value="Yes %"
    for i in range(len(unique_values)):
        # ws.Cells(i+row,col).Value=f"=({get_excel_col_index(col-2)}{i+row}/{get_excel_col_index(col-1)}{i+row})*100"
        print(row+i+1,col,f"=({get_excel_col_index(col-2)}{i+row}/{get_excel_col_index(col-1)}{i+row})*100")
def create_excel_pivot():
    # Delete old file if already exist if remove_output_if_exist is True
    if remove_output_if_exist:
        if os.path.isfile(output):
            try:
                os.remove(output)
                print(f"{output} removed successfully")
            except OSError as error:
                print(error)
                print(f"File :{output} can not be removed")
    # Selecting columns only required to for our calculation
    required_cols=[group_on_cols,groupby_col]
    required_cols.extend(list_of_columns)
    wb, ws_pvt = write_excel(output, pivot_table_sheet_name)
    wb_old, wb, ws_old = copy_excel(old_file_name, old_data_sheet_name,destination_file=wb,destination_sheet="old",req_cols=required_cols)
    wb_new,wb, ws_new = copy_excel(new_file_name, new_data_sheet_name,destination_file=wb,destination_sheet="new",req_cols=required_cols)
    wb_old.Close(True)
    wb_new.Close(True)
    # Format pivot sheet
    ws_pvt.Activate()
    # Adding level
    ws_pvt.Cells(1, 1).Value="Old Data Pivot"
    ws_pvt.Cells(1, 8).Value="New Data Pivot"
    # Setting color
    ws_pvt.Range("A1:D1").Interior.Color = int("FFFF00",16)
    ws_pvt.Range("H1:K1").Interior.Color = int("FFFF00",16)
    ws_pvt.Range("A1").Font.Bold = True
    ws_pvt.Range("H1").Font.Bold = True
    # Freeze 2nd row
    ws_pvt.Range("A2").Select()
    excel.ActiveWindow.FreezePanes = True

    # sys.exit(1)
    # Setup and call pivot_table
    pt_fields = [[group_on_cols, f'Count of {group_on_cols}', win32c.xlCount, '0']]
    count = 2
    for pivot_cols in list_of_columns:
        # Creating initial pivot table
        pt_name = f"{pivot_cols} Analysis "  # must be a string
        pt_rows = [pivot_cols]  # must be a list
        pt_cols = [groupby_col]  # must be a list
        pt_filters = []  # must be a list
        pivot_table(wb, ws_old, ws_pvt, pivot_table_sheet_name, f"{pt_name} Old", pt_rows, pt_cols, pt_filters, pt_fields, count, 1)
        pivot_table(wb, ws_new, ws_pvt, pivot_table_sheet_name, f"{pt_name} New", pt_rows, pt_cols, pt_filters, pt_fields, count, 8)
        unique_values=get_unique_column_values(ws_new, required_cols.index(pivot_cols) + 1)
        add_yes_percent(ws_pvt,unique_values,count, 1)
        add_yes_percent(ws_pvt, unique_values, count, 8)
        count += len(unique_values) + 5
    wb.SaveAs(str(output))
    wb.Close(True)
    # excel.Quit()


def get_unique_column_values(ws: object, column_count: int):
    unique_cols = set()
    for i in range(2, ws.UsedRange.Rows.Count):
        unique_cols.add(ws.Cells(i, column_count).Value)
    # print(unique_cols)
    return unique_cols


if __name__ == "__main__":
    create_excel_pivot()
