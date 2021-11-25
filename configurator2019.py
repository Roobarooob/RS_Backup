import os
from pickle import STOP
from tkinter import*
from tkinter.font import Font
from tkinter import messagebox

import win32com.client
import datetime



def addtask():
    time_zone = datetime.datetime.now().astimezone()
    utc_offset = time_zone.utcoffset() // datetime.timedelta(hours = 1)
    t_hour = r_h.get()
    t_min = r_m.get()
    t_dayint = r_d.get()
    revit_vers = w_srv_vers
    if t_hour.isnumeric() and t_min.isnumeric() and t_dayint.isnumeric():
        check = True
    else:
        check = False
    if  check and int(t_hour)<=23 and int(t_hour)>=0 and int(t_min)<=60 and int(t_min)>=0 and int(t_dayint)>=0:
        correct_hour = str(int(t_hour)-int(utc_offset))
        if int(correct_hour)==24:
            correct_hour=str(0)
        elif int(correct_hour)>24:
            correct_hour=str(int(correct_hour)-24)
        elif int(correct_hour)<0:
            correct_hour=str(int(correct_hour)+24)   
        computer_name = "" #leave all blank for current computer, current user
        computer_username = ""
        computer_userdomain = ""
        computer_password = ""
        action_id = "RevitServer Backup Action" #arbitrary action ID
        action_path = os.path.join(os.getcwd(),'makebackup.exe') #executable path (could be python.exe)
        action_arguments = r'' #arguments (could be something.py)
        action_workdir = os.getcwd() #working directory for action executable
        author = "RS_Backup" #so that end users know who you are
        description = "Резервное копирование с RevitServer в определённое время" #so that end users can identify the task
        task_id = f"RevitServer {revit_vers} Backup "
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
        trigger.DaysInterval = t_dayint
        trigger.StartBoundary = f"2021-11-20T{correct_hour}:{t_min}:00-00:00"
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
        #runningTask = task.Run("")
        task.Enabled = True
        window = Tk()
        window.withdraw()    
        messagebox.showinfo('Успех',f"Задача {task_id} успешно создана!")  
    else:        
        window = Tk()
        window.withdraw()    
        messagebox.showerror('Ошибка',"Задайте корректные параметры задачи планировщика")  
varpath = os.path.join(os.getcwd(), 'var.txt')

vars = open(varpath).read().splitlines()
d = dict(list(i.split(' = ') for i in vars))

r_days_ago = d['days_ago']
r_rs_path = d['rs_path']
r_localpath = d['localpath']
r_prefix = d['prefix']
r_srv_ip = d['srv_ip']
w_srv_vers = '2019'
def ok():
    w_days_ago = e1.get()
    w_rs_path = e2.get()
    w_localpath = e3.get()
    w_prefix = e4.get()
    w_srv_ip = e5.get()
    text = f'days_ago = {w_days_ago}\nrs_path = {w_rs_path}\nlocalpath = {w_localpath}\nprefix = {w_prefix}\nsrv_ip = {w_srv_ip}\nsrv_vers = {w_srv_vers}'
    try:
        open(varpath, 'w').write(text)
        window = Tk()
        window.withdraw()    
        messagebox.showinfo('Успех',"Задача данные сохранены!")
    except:
        window = Tk()
        window.withdraw()    
        messagebox.showerror("Ошибка записи", "Откройте программу от имени администратора")    
def close():
    root.destroy()

  
root = Tk()
root.title(f'Configurator {w_srv_vers}')
root.resizable(False, False)
root.protocol('WM_DELETE_WINDOW', close)

exfont = Font(size =12)

Label(root,text='Количество дней хранения', font = exfont).grid(column=0, row=0)
Label(root,text='Путь к Revit Server до папки "Projects"',font = exfont).grid(column=0, row=1)
Label(root,text='Путь к Архиву',font = exfont).grid(column=0, row=2)
Label(root,text='Префикс Архивной папки',font = exfont).grid(column=0, row=3)
Label(root,text='Ip адрес RevitServer',font = exfont).grid(column=0, row=4)


e1 = Entry(root, font = exfont)
e1.grid(column=1, row=0,sticky=NSEW)
e1.insert(END,r_days_ago)
e2 = Entry(root,font = exfont)
e2.insert(END,r_rs_path)
e2.grid(column=1, row=1,sticky=NSEW)
e3 = Entry(root,font = exfont)
e3.grid(column=1, row=2,sticky=NSEW)
e3.insert(END,r_localpath)
e4 = Entry(root,font = exfont)
e4.grid(column=1, row=3,sticky=NSEW)
e4.insert(END,r_prefix)
e5 = Entry(root, font = exfont)
e5.grid(column=1, row=4,sticky=NSEW)
e5.insert(END,r_srv_ip)
e6 = Entry(root, font = exfont)


Button(root,text='Записать',command = ok,font = exfont,width=30).grid(column=0, columnspan=2, row=6,sticky=NSEW, padx=3,pady=3)

Label(root, text='Настройки планировщика Windows').grid(column=0,row=7,columnspan=2)
Label(root, text = 'Время запуска',font = exfont).grid(column=0,row=8)
time_frame = Frame(root)
time_frame.grid(column=1 ,row=8,sticky=NSEW, rowspan=2)
r_h = Entry(time_frame,font = exfont, width=5)
r_h.grid(column = 0, row = 0)
Label(time_frame, text = ':',font = exfont).grid(column=1,row=0)
r_m = Entry(time_frame,font = exfont, width=5)
r_m.grid(column = 2, row = 0)
Label(root, text='Интервал повторения - каждые', font=exfont,wraplength=300).grid(column=0,row=9)
r_d = Entry(time_frame,font = exfont, width=5)
r_d.grid(column = 0, row = 1,sticky=W)
Label(time_frame, text = 'дней',font = exfont).grid(column=1,row=1, columnspan=2)
Button(root,text='Создать задачу в планировщике Windows',command = addtask,font = exfont,).grid(column=0, columnspan=2, row=10 ,sticky=NSEW,padx=3,pady=3)
root.mainloop()
