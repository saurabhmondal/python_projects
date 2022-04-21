'''
Base Excel (using win32 library) class containing common excel manupulation functions
'''
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
        self.groupby_col = self.config["pivot_table_conf"]["columns"]["columns_list"]
        self.group_on_cols = self.config["pivot_table_conf"]["columns"]["values_list"]
        self.filter_columns = self.config["pivot_table_conf"]["columns"]["filter_list"]
        self.include_columns = self.config["pivot_table_conf"]["columns"]["rows_list"]
        self.exclude_columns = self.config["pivot_table_conf"]["columns"]["exclude_columns"]

    def pivot_table(self, wb: object, ws: object, pt_ws: object, pt_name: str, pt_rows: list,  start_row: int, start_col: int):
        """
        wb = workbook reference
        ws = data worksheet reference
        pt_ws = pivot table worksheet reference
        pt_name = name given to pivot table
        pt_rows: values selected for filling the pivot tables rows
        start_row: starting row of pivot table
        start_col: starting cols of pivot table
        """
        # Creating initial pivot table
        pt_cols = self.groupby_col  # must be a list
        pt_filters=self.filter_columns
        # Setup and call pivot_table
        pt_fields = [[group_on_cols, f'Count of {group_on_cols}', win32c.xlCount, '0'] for group_on_cols in self.group_on_cols]
        # Pivot table worksheet name
        ws_name=pt_ws.Name
        # pivot table location
        pt_loc = len(pt_filters) + start_row
        # print(ws_name,pt_loc,start_col,pt_name)
        # grab the pivot table source data
        pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws.UsedRange)
        # create the pivot table object
        pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C{start_col}', TableName=pt_name)
        # selecte the pivot table work sheet and location to create the pivot table
        pt_ws.Select()
        pt_ws.Cells(pt_loc, start_col).Select()

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

        pt_ws.Cells(pt_loc, start_col + 1).Value = pt_cols[0]
        pt_ws.Cells(pt_loc + 1, start_col).Value = pt_rows[0]

    def read_excel(self, filename, sheet_name="Sheet1"):
        ''' Open, read and return workbook and work sheet objects'''
        # try except for file / path
        try:
            wb = self.excel.Workbooks.Open(filename)
        except com_error as e:
            if e.excepinfo[5] == -2146827284:
                print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
            else:
                raise e
            sys.exit(1)
        ws = wb.Sheets(sheet_name)
        return wb, ws

    def write_excel(self, filename, sheet_name="Sheet1", sheet_obj=None, req_cols=[]):
        '''write content to new excel sheet or just create blank excel if (sheet_obj is None) sheet with name sheet_name'''

        if isinstance(filename, str):
            # If filename is string (e.g. excel path)
            try:
                # try to open it
                wb = self.excel.Workbooks.Open(filename)
            except com_error as e:
                if e.excepinfo[5] == -2146827284:
                    print(f'Existing file not found.  Hence creating the file: {filename}')
                    # create new work book object
                    wb = self.excel.Workbooks.Add()
        else:
            # In case filename is not a path then it should be existing workbook object
            wb = filename
        # Add sheet with given name (in sheet_name variable)
        wb.Sheets.Add().Name = sheet_name
        ws2 = wb.Sheets(sheet_name)
        # in case sheet object provided
        if sheet_obj is not None:
            ws2 = wb.Worksheets(sheet_name)
            # Get max row count of given sheet object
            max_source_row_count = sheet_obj.UsedRange.Rows.Count
            # Remove any filter in given sheet object
            if sheet_obj.FilterMode:
                sheet_obj.ShowAllData()
            # In case no specific column list provided then copy whole sheet (sheet_obj) to new "sheet_name"
            if len(req_cols) == 0:
                max_col_in_source = get_excel_col_index(sheet_obj.UsedRange.Columns.Count)
                sheet_obj.Range(f"A1:{max_col_in_source}{sheet_obj.UsedRange.Rows.Count}").Copy(
                    ws2.Range(f"A1:{max_col_in_source}{sheet_obj.UsedRange.Rows.Count}"))
            else:
                # In case specific column list provided then copy specified columns from sheet_obj to new "sheet_name"
                # taking counter as destination sheet will have selective columns
                destination_col_index = 1
                for col_index in req_cols:
                    source_row_range=f"{col_index}1:{col_index}{max_source_row_count}"
                    destination_row_range=f"{get_excel_col_index(destination_col_index)}1:{get_excel_col_index(destination_col_index)}{max_source_row_count}"
                    sheet_obj.Range(source_row_range).Copy(ws2.Range(destination_row_range))
                    destination_col_index += 1
        return wb, ws2

    def copy_excel(self, source_file, source_sheet, destination_file="", destination_sheet="", req_cols=[],exclude_columns=[]):
        # Copy data from one sheet to another sheet
        # In case no seperate file provided the treat source as destination
        if destination_file == "":
            destination_file = source_file
            # In case no seperate sheet name provided the treat current sheet +" copy" as destination sheetname
            if destination_sheet == "":
                destination_sheet = source_sheet + " copy"
        # Get data from source sheet
        old_wb, old_ws = self.read_excel(source_file, source_sheet)
        # If both required and exclude column has values the exit as both cannot have values
        if len(req_cols)>0 and len(exclude_columns)>0:
            print("Please provide either required cols for pivot or exclude columns from pivot")
            sys.exit(1)
        else:
            # Putting values_list and columns_list (Refer config file) inside required column list so the appear in copied sheet
            req_cols_index = [get_excel_col_index(i) for i in range(1, old_ws.UsedRange.Columns.Count) if
                              old_ws.Cells(1, i).Value in self.groupby_col+self.group_on_cols]
        # Appending excel column name in previously prepared sheet
        if len(req_cols)>0:
            # Include columns from req_cols
            req_cols_index.extend([get_excel_col_index(i) for i in range(1, old_ws.UsedRange.Columns.Count) if
                          old_ws.Cells(1, i).Value in req_cols])
        elif len(exclude_columns)>0:
            # exclude all columns from exclude_columns and include rest of the columns
            req_cols_index.extend([get_excel_col_index(i) for i in range(1, old_ws.UsedRange.Columns.Count) if
                              old_ws.Cells(1, i).Value not in exclude_columns])
        wb_new, ws_new = self.write_excel(destination_file, sheet_name=destination_sheet, sheet_obj=old_ws,
                                          req_cols=req_cols_index)
        return old_wb, wb_new, ws_new

    def get_unique_column_values(self, ws: object, column_count: int):
        # Get unique values from specified columns
        return set(ws.Range(
            f"{get_excel_col_index(column_count)}2:{get_excel_col_index(column_count)}{ws.UsedRange.Rows.Count}").Value)

