from genericpath import exists
import subprocess
import os
import pathlib
import datetime
import shutil
import tkinter
from tkinter import messagebox, Tk
import win32com.client
import datetime

def addtask():
    computer_name = "" #leave all blank for current computer, current user
    computer_username = ""
    computer_userdomain = ""
    computer_password = ""
    action_id = "Test Task" #arbitrary action ID
    action_path = os.path.join(os.getcwd(),'makebackup.exe'). #executable path (could be python.exe)
    action_arguments = r'' #arguments (could be something.py)
    action_workdir = os.getcwd() #working directory for action executable
    author = "RS_Backup" #so that end users know who you are
    description = "testing task" #so that end users can identify the task
    task_id = "RevitServer_Backup"
    task_hidden = False #set this to True to hide the task in the interface
    username = ""
    password = ""
    run_flags = "TASK_RUN_NO_FLAGS" #see dict below, use in combo with username/password



    #define constants
    TASK_TRIGGER_DAILY = 2
    TASK_CREATE = 2
    TASK_CREATE_OR_UPDATE = 6
    TASK_ACTION_EXEC = 0
    IID_ITask = "{148BD524-A2AB-11CE-B11F-00AA00530503}"
    RUNFLAGSENUM = {
        "TASK_RUN_NO_FLAGS"              : 0,
        "TASK_RUN_AS_SELF"               : 1,
        "TASK_RUN_IGNORE_CONSTRAINTS"    : 2,
        "TASK_RUN_USE_SESSION_ID"        : 4,
        "TASK_RUN_USER_SID"              : 8 
    }

    #connect to the scheduler (Vista/Server 2008 and above only)
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect(computer_name or None, computer_username or None, computer_userdomain or None, computer_password or None)
    rootFolder = scheduler.GetFolder("\\")

    #(re)define the task
    taskDef = scheduler.NewTask(0)
    colTriggers = taskDef.Triggers
    trigger = colTriggers.Create(TASK_TRIGGER_DAILY)
    trigger.DaysInterval = 1
    trigger.StartBoundary = "2021-11-20T20:00:00-00:00" #never start
    trigger.Enabled = True

    colActions = taskDef.Actions
    action = colActions.Create(TASK_ACTION_EXEC)
    action.ID = action_id
    action.Path = action_path
    action.WorkingDirectory = action_workdir
    action.Arguments = action_arguments

    info = taskDef.RegistrationInfo
    info.Author = author
    info.Description = description

    settings = taskDef.Settings
    settings.Enabled = False
    settings.Hidden = task_hidden

    #register the task (create or update, just keep the task name the same)
    result = rootFolder.RegisterTaskDefinition(task_id, taskDef, TASK_CREATE_OR_UPDATE, "", "", RUNFLAGSENUM[run_flags] ) #username, password

    #run the task once
    task = rootFolder.GetTask(task_id)
    task.Enabled = True
    runningTask = task.Run("")
    task.Enabled = True

def rvtnames(path): #все файлы и их пути на RevitServer
    for root, dirs, files in os.walk(path): 
        for file in files: 
            if(file.endswith(".rvt")): 
                filedir = os.path.join(root,file)
                fullfolderpath = os.path.split(filedir)[0]
                parts = pathlib.Path(fullfolderpath).parts[1:]
                folderpath=os.path.join(*parts)
                pathlist.append(folderpath)

def make_new_dir(path): #создание новой директории если ее нет
    isExist = os.path.exists(path)
    if not isExist:
        os.makedirs(path)

def del_old(path,prefix,older_day): #удаление старейшего архива
    trim = prefix+'__'
    archive = os.listdir(path)
    for i in archive:
        if i.split(trim)[1] < str(older_day):
            del_path = os.path.join(path,i) 
            shutil.rmtree(del_path)
            
#словарь из файла переменных

varpath = os.path.join(os.getcwd(), 'var.txt')
vars = open(varpath, 'r').read().splitlines()
d = dict(list(i.split(' = ') for i in vars))

#начальные переменные
try:
    days_ago = int(d['days_ago'])
    rs_path = d['rs_path']
    localpath = d['localpath']
    prefix = d['prefix']
    srv_ip=d['srv_ip']
except:
    window = Tk()
    window.withdraw()    
    messagebox.showerror("Ошибка чтения файла", "Проверьте правильность данных с помощью Configurator")
#создание путей
today = datetime.date.today()
delta = datetime.timedelta(days=days_ago)
older_day = today-delta
xpath = os.path.join(localpath,prefix)
today_path = xpath+'__'+str(today)
old_path = xpath+'__'+str(older_day)

#листы переменных
folderlist = []
pathlist = []


#основной код
#del_old(old_path)
del_old(localpath,prefix,older_day)
rvtnames(rs_path)
for path in pathlist:
    rs_rvtpath  = path
    finpath = os.path.join(today_path,path)
    fin_directory = os.path.splitext(finpath)[0]
    directory = fullfolderpath = os.path.split(fin_directory)[0]
    make_new_dir(directory)
    
    command1 = 'cd C:/Program Files/Autodesk/Revit 2019/RevitServerToolCommand'
    command2 = f'RevitServerTool createLocalRVT "{rs_rvtpath}" -s {srv_ip} -d "{fin_directory}"'
    print(rs_rvtpath)
    try:
        res = subprocess.call(command1+'&&'+command2, shell = True) #the method returns the exit code print("Returned Value: ", res)
        print('Yes')
    except:
        pass
        print('No')
