'''
Utils containing descritive specific functions
'''
from utils.commonUtils import get_excel_col_index, deleteFile,create_folder
from utils.excelUtils import excelPivot

class descritivePivot(excelPivot):
    def __init__(self, config_filename):
        super().__init__(config_filename)
        # setting output variables
        self.output = self.config["input_output_sheet_data"]["output"]["filename"]
        self.pivot_table_sheet_name = self.config["input_output_sheet_data"]["output"]["sheetname"]
        self.remove_output_if_exist = self.config["input_output_sheet_data"]["output"]["remove_output_if_exist"]
        self.input = self.config["input_output_sheet_data"]["input"]

    def create_descritive_pivot(self):
        ######################################################################################
        # Delete old file if already exist if remove_output_if_exist is True
        ######################################################################################
        if self.remove_output_if_exist:
            deleteFile(self.output)

        ######################################################################################
        # Creating output sheet and formatting it
        ######################################################################################
        wb, ws_pvt = self.write_excel(self.output, self.pivot_table_sheet_name)
        # Format pivot sheet
        ws_pvt.Activate()
        # Freeze 2nd row
        ws_pvt.Range("A2").Select()
        self.excel.ActiveWindow.FreezePanes = True
        ######################################################################################
        # Creating dictionary containing all input file/ sheet data
        ######################################################################################
        input_dict = dict()
        pivot_start_col = 1
        # Getting each input from config
        for in_data in self.input:
            # coping data from input to output sheet
            wb_old, wb, ws = self.copy_excel(in_data["filename"], in_data["sheetname"], destination_file=wb,
                                             destination_sheet=in_data["alias"], req_cols=self.include_columns,exclude_columns=self.exclude_columns)
            # Closing opened input file
            wb_old.Close(True)
            # Getting headers from input file
            headers = []
            for i in range(1, ws.UsedRange.Columns.Count + 1):
                headers.append(ws.Cells(1, i).Value)
            # Getting pivot table width for the input file
            pivot_table_width = 2
            for groupby_col in self.groupby_col:
                pivot_table_width += len(self.get_unique_column_values(ws, headers.index(groupby_col) + 1))
            # Updating
            input_dict.update({
                in_data["alias"]: {
                    "sheet_object": ws,
                    "header": headers,
                    "pivot_start_row": 3,
                    "pivot_table_width": pivot_table_width
                }
            })
            ws_pvt.Cells(1, pivot_start_col).Value = f'{in_data["alias"]} Data Pivot'
            ws_pvt.Range(
                f"{get_excel_col_index(pivot_start_col)}1:{get_excel_col_index(pivot_start_col+pivot_table_width-1)}1").Interior.Color = int(
                self.config["pivot_table_conf"]["custom_header"]["color"], 16)
            ws_pvt.Range(f"{get_excel_col_index(pivot_start_col)}1").Font.Bold = \
            self.config["pivot_table_conf"]["custom_header"]["isBold"]
            pivot_start_col += self.config["pivot_table_conf"]["gap_between_two_pivots"]["vertical"]
        # print(list(input_dict))
        list_of_columns=[ col for col in input_dict[list(input_dict)[0]]["header"] if col not in self.group_on_cols+self.groupby_col ]
        # print(list_of_columns)
        for pivot_cols in list_of_columns:
            pivot_start_col = 1
            yes_cords = list()
            for alias, sheet_data in input_dict.items():
                self.pivot_table(wb, sheet_data["sheet_object"], ws_pvt,
                                 f"{alias}:{pivot_cols} Analysis",[pivot_cols], sheet_data["pivot_start_row"], pivot_start_col)
                unique_values = self.get_unique_column_values(sheet_data["sheet_object"],
                                                              sheet_data["header"].index(pivot_cols) + 1)
                if self.config["pivot_table_conf"]["show_yes_percent"]:
                    yes_cords.append([sheet_data["pivot_table_width"],self.add_yes_percent(ws_pvt, unique_values, sheet_data["pivot_start_row"] + 1,
                                                          pivot_start_col + sheet_data["pivot_table_width"])])
                sheet_data["pivot_start_row"] += len(unique_values) + \
                                                 self.config["pivot_table_conf"]["gap_between_two_pivots"]["horizontal"]
                pivot_start_col += self.config["pivot_table_conf"]["gap_between_two_pivots"]["vertical"]
            if self.config["pivot_table_conf"]["show_yes_percent"]:
                if "format_diff" in self.config["pivot_table_conf"] and len(yes_cords)>1:
                    for cord_count in range(0, len(yes_cords)-1):
                        self.format_yes_percent(ws_pvt, yes_cords[cord_count], yes_cords[cord_count + 1])
        wb.Worksheets("Sheet1").Delete()
        create_folder(str(self.output))
        wb.SaveAs(str(self.output))
        wb.Close(True)

    def add_yes_percent(self, ws, unique_values, row, col):
        ws.Cells(row, col).Value = "Yes %"
        ws.Range(f"{get_excel_col_index(col)}{row}").Interior.Color = int(
            self.config["pivot_table_conf"]["custom_header"]["color"], 16)
        ws.Range(f"{get_excel_col_index(col)}{row}").Font.Bold = self.config["pivot_table_conf"]["custom_header"][
            "isBold"]
        pivot_yes_percent_cord = list()
        for i in range(1, len(unique_values) + 2):
            formula = f'=ROUND(({get_excel_col_index(col - 2)}{i + row}/{get_excel_col_index(col - 1)}{i + row})*100,1) & "%"'
            ws.Cells(i + row, col).Value = formula
            pivot_yes_percent_cord.append((i + row, col))
        return pivot_yes_percent_cord

    def format_yes_percent(self, ws, cord_old, cord_new):
        count = 0
        cord_old_pvt_tab_width=cord_old[0]
        cord_new_pvt_tab_width = cord_new[0]
        if len(cord_old[1])==len(cord_new[1]):
            # print(cord_old, cord_new)
            for i in cord_old[1]:
                old_value = float(ws.Cells(i[0], i[1]).Value.replace("%", ""))
                new_value = float(ws.Cells(cord_new[1][count][0], cord_new[1][count][1]).Value.replace("%", ""))
                diff = new_value-old_value
                if abs(diff) > self.config["pivot_table_conf"]["format_diff"]["tolerance"]:
                    if diff > 0:
                        ws.Range(f"{get_excel_col_index(cord_new[1][count][1])}{cord_new[1][count][0]}").Interior.Color = int(
                            self.config["pivot_table_conf"]["format_diff"]["color"]["increase"], 16)
                    elif diff < 0:
                        ws.Range(f"{get_excel_col_index(cord_new[1][count][1])}{cord_new[1][count][0]}").Interior.Color = int(
                            self.config["pivot_table_conf"]["format_diff"]["color"]["decrease"], 16)
                count += 1
        else:
            old_pivot_row=cord_old[1][0][0]
            old_pivot_col = cord_old[1][0][1]
            new_pivot_row = cord_new[1][0][0]
            new_pivot_col = cord_new[1][0][1]
            # print(f"in else: old_pivot_row:{old_pivot_row},old_pivot_col:{old_pivot_col}")
            # print(f"in else: new_pivot_row:{new_pivot_row},new_pivot_col:{new_pivot_col}")
            ws.Cells(old_pivot_row-3,old_pivot_col-cord_old_pvt_tab_width).Value="Mismatch in unique values"
            ws.Range(
                f"{get_excel_col_index(old_pivot_col - cord_old_pvt_tab_width)}{old_pivot_row - 3}:{get_excel_col_index(old_pivot_col)}{old_pivot_row - 3}").Interior.Color = int(
                self.config["pivot_table_conf"]["format_diff"]["color"]["decrease"], 16)
            ws.Cells(new_pivot_row - 3,
                         new_pivot_col - cord_new_pvt_tab_width).Value = "Mismatch in unique values"
            ws.Range(
                    f"{get_excel_col_index(new_pivot_col - cord_new_pvt_tab_width)}{new_pivot_row - 3}:{get_excel_col_index(new_pivot_col)}{new_pivot_row - 3}").Interior.Color = int(
                    self.config["pivot_table_conf"]["format_diff"]["color"]["decrease"], 16)
