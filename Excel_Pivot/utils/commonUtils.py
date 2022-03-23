import json
import os,sys

def get_excel_col_index(num):
    threshold=(ord("Z")-ord("A")+1)
    if int(num)<=threshold:
        return chr(num + ord("A")-1)
    else:
        return get_excel_col_index(int(num/threshold))+get_excel_col_index(num%threshold)

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