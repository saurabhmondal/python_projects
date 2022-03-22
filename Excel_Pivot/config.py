groupby_col="DECISION"
group_on_cols="EVENT_ID"
exclude_from_pivot=['EVENT_ID','RECV_DT','EVENT_CREATED_DT','DECISION']
list_of_columns = ['AMOUNT','ACCIDENT']

output=r"C:\Users\Ssaurabh\PycharmProjects\PivotTableTest\Comparison.xlsx"
 
old_file_name = r"C:\Users\Ssaurabh\PycharmProjects\PivotTableTest\learning_docvrs_old.xlsx"
new_file_name =r"C:\Users\Ssaurabh\PycharmProjects\PivotTableTest\learning_docvrs_new.xlsx"

new_data_sheet_name = "Sheet1"
old_data_sheet_name = "Sheet1"
pivot_table_sheet_name = 'pivot_table'
remove_output_if_exist=True




