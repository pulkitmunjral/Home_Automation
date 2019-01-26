import xlwt
import xlrd
from tkinter import*
import requests as r
wb=xlrd.open_workbook('remote_access.xls')
s=wb.sheet_by_index(0)
ret=s.col_values(3)
o=['OFF','ON']
c=["black","orange"]

def login(a,b):
    wb=xlrd.open_workbook('remote_access.xls')
    s=wb.sheet_by_index(0)
    user=s.col_values(0)
    pas=s.col_values(1)
    up=s.col_values(2)
    ret=s.col_values(3)
    for q in range(1,len(user)):
        if(a==user[q]):
            j=1
            break
        else:
            j=0
    global g
    g=q
    if(j==1):
        p=0
        for p in range(0,1):
            if(int(b)==int(pas[q])):
                e='http://indianiotcloud.com/retrieve.php?id='+ret[q]
                p=r.get(e)
                d=p.json()
                i=int(d['result'][0]['field1'])
                j=int(d['result'][0]['field2'])
                k=int(d['result'][0]['field3'])
                l=int(d['result'][0]['field4'])
                c=["black","orange"]

                Label(F2,text=o[i],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=7,column=2)
                Button(F2,text='',bg=c[i],width='3',height='1').grid(row=7,column=3)
                Label(F2,text=o[j],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=9,column=2)
                Button(F2,text='',bg=c[j],width='3',height='1').grid(row=9,column=3)
                Label(F2,text=o[k],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=11,column=2)
                Button(F2,text='',bg=c[k],width='3',height='1').grid(row=11,column=3)
                Label(F2,text=o[l],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=13,column=2)
                Button(F2,text='',bg=c[l],width='3',height='1').grid(row=13,column=3)
                new_window(F2)
    else:
        new_window(F1)

    
def do(i,j,k,l):
    a='http://indianiotcloud.com/update1.php?id='+ret[g]
    b='&field1='+str(i)
    c='&field2='+str(j)
    d='&field3='+str(k)
    e='&field4='+str(l)
    p=a+b+c+d+e
    r.get(p)
    c=["black","orange"]
    Label(F2,text=o[i],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=7,column=2)
    Button(F2,text='',bg=c[i],width='3',height='1').grid(row=7,column=3)
    Label(F2,text=o[j],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=9,column=2)
    Button(F2,text='',bg=c[j],width='3',height='1').grid(row=9,column=3)
    Label(F2,text=o[k],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=11,column=2)
    Button(F2,text='',bg=c[k],width='3',height='1').grid(row=11,column=3)
    Label(F2,text=o[l],font='Helvatica 20 italic ',fg="black",width='8',height='1').grid(row=13,column=2)
    Button(F2,text='',bg=c[l],width='3',height='1').grid(row=13,column=3)

def bub1():
    e='http://indianiotcloud.com/retrieve.php?id='+ret[g]
    p=r.get(e)
    d=p.json()
    i=int(d['result'][0]['field1'])
    j=int(d['result'][0]['field2'])
    k=int(d['result'][0]['field3'])
    l=int(d['result'][0]['field4'])
    if(i==1):
        i=0
    elif(i==0):
        i=1
    do(i,j,k,l)

    
def bub2():
    e='http://indianiotcloud.com/retrieve.php?id='+ret[g]
    p=r.get(e)
    d=p.json()
    i=int(d['result'][0]['field1'])
    j=int(d['result'][0]['field2'])
    k=int(d['result'][0]['field3'])
    l=int(d['result'][0]['field4'])
    if(j==1):
        j=0
    elif(j==0):
        j=1
    do(i,j,k,l)

    
def bub3():
    e='http://indianiotcloud.com/retrieve.php?id='+ret[g]
    p=r.get(e)
    d=p.json()
    i=int(d['result'][0]['field1'])
    j=int(d['result'][0]['field2'])
    k=int(d['result'][0]['field3'])
    l=int(d['result'][0]['field4'])
    if(k==1):
        k=0
    elif(k==0):
        k=1
    do(i,j,k,l)

    
def bub4():
    e='http://indianiotcloud.com/retrieve.php?id='+ret[g]
    p=r.get(e)
    d=p.json()
    i=int(d['result'][0]['field1'])
    j=int(d['result'][0]['field2'])
    k=int(d['result'][0]['field3'])
    l=int(d['result'][0]['field4'])
    if(l==1):
        l=0
    elif(l==0):
        l=1
    do(i,j,k,l)
    

def new_window(fr):
    fr.tkraise()
    for f in (F1,F2,F3):
        f.grid(row=0,column=0,sticky="news")
        




root=Tk()
root.title('Welcome')
root.geometry('900x700')
F1=Frame(root)
F3=Frame(root)
F2=Frame(root)

new_window(F1)
Label(F1,text='').grid(row=1,column=1)
Label(F1,text='Welcome to Home Automation',font='Helvatica 25 bold italic underline',bg="red",fg="black").grid(row=2,column=2)
Label(F1,text='').grid(row=3,column=1)
Label(F1,text='').grid(row=4,column=1)
Label(F1,text='').grid(row=5,column=1)
Label(F1,text='Username:-',font='Helvatica 20 italic ',fg="black").grid(row=6,column=1)
user=StringVar()
Entry(F1,textvariable=user).grid(row=6,column=2)
Label(F1,text='').grid(row=7,column=1)
Label(F1,text='Password:-',font='Helvatica 20 italic ',fg="black").grid(row=8,column=1)
pas=StringVar()
Entry(F1,textvariable=pas).grid(row=8,column=2)
Label(F1,text='').grid(row=9,column=1)
Label(F1,text='').grid(row=10,column=1)
Button(F1,text='Submit',font='Helvatica 20 bold italic',bg="red",fg="black",command=lambda:login(user.get(),pas.get())).grid(row=11,column=2)


Label(F2,text='\t\t\t\t').grid(row=1,column=1)
Label(F2,text='One Touch Solution',font='Helvatica 25 bold italic underline',bg="red",fg="black").grid(row=2,column=2)
Label(F2,text='').grid(row=3,column=1)
Label(F2,text='').grid(row=4,column=1)
Label(F2,text='Appliances',font='Helvatica 20 italic ',fg="black").grid(row=5,column=1)
Label(F2,text='Postion',font='Helvatica 20 italic ',fg="black").grid(row=5,column=2)
Label(F2,text='').grid(row=6,column=1)
Button(F2,text='Bulb',font='Helvatica 20 bold italic',bg="black",fg="red",width='8',height='1',command=bub1).grid(row=7,column=1)
Label(F2,text='').grid(row=8,column=1)
Button(F2,text='Fan',font='Helvatica 20 bold italic',bg="black",fg="red",width='8',height='1',command=bub2).grid(row=9,column=1)
Label(F2,text='').grid(row=10,column=1)
Button(F2,text='Tv',font='Helvatica 20 bold italic',bg="black",fg="red",width='8',height='1',command=bub3).grid(row=11,column=1)
Label(F2,text='').grid(row=12,column=1)
Button(F2,text='Light',font='Helvatica 20 bold italic',bg="black",fg="red",width='8',height='1',command=bub4).grid(row=13,column=1)


Label(F2,text=' \n \n').grid(row=14,column=1)
Button(F2,text='Refresh',font='Helvatica 20 bold italic',bg="orange",fg="black",width='8',height='1',command=lambda:login(user.get(),pas.get())).grid(row=15,column=2)
Button(F2,text='Exit',font='Helvatica 20 bold italic',bg="orange",fg="black",width='8',height='1',command=lambda:new_window(F1)).grid(row=15,column=3)
