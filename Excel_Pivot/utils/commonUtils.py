import json
import os,sys


def get_excel_col_index(num):
    '''Excepts any positive integer as input and return excel column name equivalent
    input and output mapping as below
    1-->A, 2-->B,25-->Z,26-->AA
    '''
    threshold=(ord("Z")-ord("A")+1)
    rem=num%threshold
    if num<=threshold:
        if rem==0:
            return "Z"
        else:
            return chr(ord("A")+rem-1)
    div=int(num/threshold)
    if div>=1:
        if rem==0:
            return get_excel_col_index(div-1)+get_excel_col_index(rem)
        else:
            return get_excel_col_index(div) + get_excel_col_index(rem)

def create_folder(file_path):
    if sys.platform=="win32":
        sep="\\"
    else:
        sep="/"
    folder_name = sep.join(file_path.split(sep)[:-1])
    print(f"Creating folder:{folder_name}")
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

def readFlatFile(file_path):
    if not os.path.exists(file_path):
        return None
    else:
        ext=file_path.split(".")[-1]
        with open(file_path, "r") as f:
            if ext=="json":
                return json.load(f)
            else:
                return f.read()

def deleteFile(filepath):
    if os.path.isfile(filepath):
        try:
            os.remove(filepath)
            print(f"{filepath} removed successfully")
        except OSError as error:
            print(error)
            print(f"File :{filepath} can not be removed")