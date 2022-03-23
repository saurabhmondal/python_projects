import win32com.client as win32
from pywintypes import com_error
import sys
from utils.commonUtils import readFlatFile, get_excel_col_index

win32c = win32.constants

class excelPivot:
    def __init__(self, config_filename):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        # excel can be visible or not
        self.excel.Visible = True  # False
        self.config = readFlatFile(config_filename)
        self.groupby_col = self.config["pivot_table_conf"]["columns"]["groupby_col"]
        self.group_on_cols = self.config["pivot_table_conf"]["columns"]["group_on_cols"]
        self.list_of_columns = self.config["pivot_table_conf"]["columns"]["list_of_columns"]

    def pivot_table(self, wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list,
                    pt_cols: list,
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

    def read_excel(self, filename, sheet_name="Sheet1"):
        # try except for file / path
        try:
            wb = self.excel.Workbooks.Open(filename)
        except com_error as e:
            if e.excepinfo[5] == -2146827284:
                print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
            else:
                raise e
            sys.exit(1)
        # set worksheet
        ws = wb.Sheets(sheet_name)
        return wb, ws

    def write_excel(self, filename, sheet_name="Sheet1", sheet_obj=None, req_cols=[]):
        if isinstance(filename, str):
            # try except for file / path
            try:
                wb = self.excel.Workbooks.Open(filename)
            except com_error as e:
                if e.excepinfo[5] == -2146827284:
                    print(f'Existing file not found.  Hence creating the file: {filename}')
                    wb = self.excel.Workbooks.Add()
        else:
            wb = filename
        wb.Sheets.Add().Name = sheet_name
        ws2 = wb.Sheets(sheet_name)
        if sheet_obj is not None:
            ws2 = wb.Worksheets(sheet_name)
            max_source_row_count = sheet_obj.UsedRange.Rows.Count
            if sheet_obj.FilterMode:
                sheet_obj.ShowAllData()
            if len(req_cols) == 0:
                max_col_in_source = get_excel_col_index(sheet_obj.UsedRange.Columns.Count)
                sheet_obj.Range(f"A1:{max_col_in_source}{sheet_obj.UsedRange.Rows.Count}").Copy(
                    ws2.Range(f"A1:{max_col_in_source}{sheet_obj.UsedRange.Rows.Count}"))
            else:
                destination_col_index = ord("A")
                for col_index in req_cols:
                    sheet_obj.Range(f"{col_index}1:{col_index}{max_source_row_count}").Copy(
                        ws2.Range(f"{chr(destination_col_index)}1:{chr(destination_col_index)}{max_source_row_count}"))
                    destination_col_index += 1
        return wb, ws2

    def copy_excel(self, source_file, source_sheet, destination_file="", destination_sheet="", req_cols=[]):
        if destination_file == "":
            destination_file = source_file
            if destination_sheet == "":
                destination_sheet = source_sheet + " copy"
        old_wb, old_ws = self.read_excel(source_file, source_sheet)
        req_cols_index = [get_excel_col_index(i) for i in range(1, old_ws.UsedRange.Columns.Count) if
                          old_ws.Cells(1, i).Value in req_cols]
        wb_new, ws_new = self.write_excel(destination_file, sheet_name=destination_sheet, sheet_obj=old_ws,
                                          req_cols=req_cols_index)
        return old_wb, wb_new, ws_new

    def get_unique_column_values(self, ws: object, column_count: int):
        return set(ws.Range(
            f"{get_excel_col_index(column_count)}2:{get_excel_col_index(column_count)}{ws.UsedRange.Rows.Count}").Value)

