import json
import logging
import logging.config
import os

common_logger = logging.getLogger()


def create_folder(file_path):
    folder_name = "/".join(file_path.split("/")[:-1])
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
def write_to_file(filename,content):
    print(content)
    with open(filename,"w+") as f:
        if filename.split('.')[-1]=='json':
            # json.dump(content,f,indent=2)
            pass
        else:
            f.write(content)


def set_logger(log_conf_filename, log_file_name):
    logging.basicConfig()
    with open(log_conf_filename, "r") as f:
        log_conf = json.load(f)
    create_folder(log_file_name)
    log_conf["handlers"]["file"]["filename"] = log_file_name
    logging.config.dictConfig(log_conf)
