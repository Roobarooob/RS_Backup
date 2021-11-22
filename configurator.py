import os
import tkinter
from tkinter import*
from tkinter.font import Font
from tkinter import messagebox

varpath = os.path.join(os.getcwd(), 'var.txt')

vars = open(varpath).read().splitlines()
d = dict(list(i.split(' = ') for i in vars))

r_days_ago = d['days_ago']
r_rs_path = d['rs_path']
r_localpath = d['localpath']
r_prefix = d['prefix']
r_srv_ip = d['srv_ip']
def ok():
    w_days_ago = e1.get()
    w_rs_path = e2.get()
    w_localpath = e3.get()
    w_prefix = e4.get()
    w_srv_ip = e5.get()
    text = f'days_ago = {w_days_ago}\nrs_path = {w_rs_path}\nlocalpath = {w_localpath}\nprefix = {w_prefix}\nsrv_ip = {w_srv_ip}'
    try:
        open(varpath, 'w').write(text)
    except:
        window = Tk()
        window.withdraw()    
        messagebox.showerror("Ошибка записи", "Откройте программу от имени администратора")    
    root.quit()
    root.destroy()
def close():
    root.destroy()

  
root = Tk()
root.title('Configurator')
root.resizable(False, False)

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

Button(root,text='Записать',command = ok,font = exfont,width=30).grid(column=0, row=5,sticky=NSEW, padx=3,pady=3)
Button(root,text='Отмена',command = close,font = exfont,width=30).grid(column=1, row=5,sticky=NSEW,padx=3,pady=3)
root.mainloop()
