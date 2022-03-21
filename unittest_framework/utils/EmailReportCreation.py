import jinja2
import datetime,os,json,re
from .common_utils import create_folder

def get_normalized_time(x,miliseconds_round_upto=2):
    if isinstance(x,datetime.datetime):
        time_delta=datetime.datetime.strftime(x,"%Y-%m-%d %H:%M:%S.%f")
    elif isinstance(x,datetime.timedelta):
        hrs=int(round(x.seconds/3600,0))
        mins=int(round((x.seconds -(hrs*3600))/60,0))
        secs=round((x.seconds-(hrs*3600+mins*60))+x.microseconds/1000000,miliseconds_round_upto)
        time_delta=f"{hrs} hr {mins} min {secs} sec"
    if isinstance(x,datetime.datetime):
        time_delta_list=time_delta.split(":")
        time_delta_list[-1]=str(round(float(time_delta_list[-1]),miliseconds_round_upto))
        tz=" "+datetime.datetime.now(datetime.timezone.utc).astimezone().tzname()
        return ":".join(time_delta_list)+tz
    elif isinstance(x,datetime.timedelta):
        return time_delta

def HtmlReport(overallData,header,tableData,filename):
    ''' Fetching out the results of Test Execution '''
    loader=jinja2.FileSystemLoader(overallData["report_template"])
    env=jinja2.Environment(loader=loader)
    template=env.get_template('')
    required_funcs={
        'round':round,
        'strftime':get_normalized_time
    }
    with open(overallData["env_config"],'r') as f:
        env_config=json.load(f)
    overallData["Environment"]=env_config["EnvironmentType"]
    overallData["Branch"]= "develop" if "dev" in env_config["EnvironmentType"].lower() else "release"
    overallData["TestScriptLocation"]='"'+env_config["githubRepo"]+'"'
    for row in tableData:
        row[0]="TC"+re.sub('_+','_',row[0]).split("_")[1]
        row[-1]=f'{round(float(row[-1].split(" ")[0]),2)} seconds'
    tableData.sort()
    create_folder(filename)
    parsed_template=template.render(overallData=overallData,header=header,TableData=tableData,**required_funcs)
    with open(filename,"w+") as f:
        f.write(parsed_template)