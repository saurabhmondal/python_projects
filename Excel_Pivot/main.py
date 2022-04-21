def get_excel_col_index(num):
    threshold=(ord("Z")-ord("A")+1)
    if int(num)<=threshold:
        return chr(num + ord("A")-1)
    else:
        return get_excel_col_index(int(num/threshold))+get_excel_col_index(num%threshold)

print(get_excel_col_index(51))

