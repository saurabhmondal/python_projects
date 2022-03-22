import pandas as pd

def create_pivot(filename,sheet_index,groupby_col,group_on_cols,exclude_from_pivot):
    # Reading Excel file
    xl=pd.ExcelFile(filename)
    # LOading data into data frame
    df=pd.read_excel(filename,sheet_name=xl.sheet_names[sheet_index],engine='openpyxl')
    df_lis_w_yes_perct=dict()
    df_lis_wo_yes_perct = dict()
    for pivot_cols in set(df.columns)-set(exclude_from_pivot):
        # Creating initial pivot table
        df_pvt=pd.pivot_table(df,index=[pivot_cols],columns=groupby_col,values=[group_on_cols],aggfunc='count')
        df_pvt.fillna(0,inplace=True)
        # Storing row wise sum in another column
        df_pvt["Total Result"]=df_pvt.sum(axis = 1, skipna = True)
        # Storing current index
        org_index=df_pvt.index.tolist()
        # Appending "Total Result" in index for future use
        org_index.append("Total Result")
        # Getting colwise sum
        horizental_sum=df_pvt.sum(axis = 0, skipna = True)
        # Adding col wise sum to last row
        df_pvt=df_pvt.append(horizental_sum,ignore_index=True)
        # Renaming column levels to match reporting
        df_pvt=df_pvt.rename(columns={group_on_cols:pivot_cols}, level=0)
        df_pvt.columns = df_pvt.columns.rename(f"Count of {group_on_cols}", level=0)
        # Renaming Index to to refect Actual Decisions and Total Result
        df_pvt=df_pvt.rename(index=dict(zip(df_pvt.index.tolist(),org_index)))
        try:
            # Adding 'Yes %' column
            df_pvt["Yes %"]= round(df_pvt[(pivot_cols,  'Yes')]*100/df_pvt[("Total Result","")],2)
            df_lis_w_yes_perct.update({pivot_cols:df_pvt})
        except KeyError:
            df_lis_wo_yes_perct.update({pivot_cols:df_pvt})
    for key,value in df_lis_w_yes_perct.items():
        print("*"*20+f" {key} "+"*"*20)
        print(value)

if __name__=="__main__":
    groupby_col="DECISION"
    group_on_cols="EVENT_ID"
    exclude_from_pivot=['EVENT_ID','RECV_DT','EVENT_CREATED_DT','DECISION']
    create_pivot(r"C:\Users\Ssaurabh\Documents\Release_New_Small.xlsx",0,groupby_col,group_on_cols,exclude_from_pivot)
