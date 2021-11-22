from genericpath import exists
import subprocess
import os
import pathlib
import datetime
import shutil
import tkinter
from tkinter import messagebox, Tk


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
