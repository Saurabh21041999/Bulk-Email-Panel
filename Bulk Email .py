#!/usr/bin/env python
# coding: utf-8

# In[3]:


from tkinter import *
from PIL import ImageTk
from tkinter import messagebox,filedialog
import os
import pandas as pd
import time
import sys
sys.path.append('C:\\Users\\admin\\Bulk mail')
import email_function

class Bulk_Email:
    def __init__(self,root):
        self.root=root
        self.root.title("Bulk Email Application")
        self.root.geometry("1000x550+200+50")
        self.root.resizable(False,False)
        self.root.config(background="White")
        
        
        #============================Variables================================#
        
        self.var_choice=StringVar()
        self.var_choice.set("single")
        
        #=============================Icons===============================#
        
        self.email_icon=ImageTk.PhotoImage(file="C:\\Users\\admin\\message.png")
        self.setting_icon=ImageTk.PhotoImage(file="C:\\Users\\admin\\setting.png")
        
        #=============================Title===============================#
        
        title=Label(self.root,text="Bulk Email Send Panel",image=self.email_icon,compound=LEFT,font=("Goudy old style",48,"bold"),padx=10,bg="#222A35",fg="white",anchor="w")
        title.place(x=0,y=0,relwidth=1)
        desc=Label(self.root,text="Use Excel File to Send the Bulk Email at once, with just one click.Ensure Email Column Name must be Email",font=("calibr (body)",14),bg="#FFD966",fg="#262626")
        desc.place(x=0,y=80,relwidth=1)
        
        #=============================Buttons===============================#
        
        btn_setting=Button(self.root,image=self.setting_icon,bd=0,activebackground="#222A35",bg="#222A35",cursor="hand2",command=self.setting_window)
        btn_setting.place(x=930,y=6)
        
        single=Radiobutton(self.root,text="Single",value="single",command=self.check_single_or_bulk,variable=self.var_choice,activebackground="white",font=("times new roman",30,"bold"),bg="white",fg="#262626")
        single.place(x=50,y=150)
        bulk=Radiobutton(self.root,text="Bulk",value="bulk",command=self.check_single_or_bulk,variable=self.var_choice,activebackground="white",font=("times new roman",30,"bold"),bg="white",fg="#262626")
        bulk.place(x=250,y=150)
        
        btn_clear=Button(self.root,text="CLEAR",command=self.clear1,font=("times new roman",18,"bold"),activebackground="#262626",activeforeground="white",bg="#262626",fg="white",cursor="hand2")
        btn_clear.place(x=700,y=490,width=120,height=30)
        btn_send=Button(self.root,text="SEND",command=self.send_email,font=("times new roman",18,"bold"),activebackground="#00B0F0",activeforeground="white",bg="#00B0F0",fg="white",cursor="hand2")
        btn_send.place(x=830,y=490,width=120,height=30)
        
        self.btn_browser=Button(self.root,text="BROWSER",command=self.browser_file,font=("times new roman",15,"bold"),activebackground="#8FAADC",activeforeground="#262626",bg="#8FAADC",fg="#262626",cursor="hand2",state=DISABLED)
        self.btn_browser.place(x=670,y=250,width=120,height=30)
        
        
        #===============================Main=====================================#
        
        to=Label(self.root,text="To (Email Address)",font=("times new roman",18),bg="white")
        to.place(x=50,y=250)
        self.txt_to=Entry(self.root,width=15,font=("times new roman",15),bg="lightyellow")
        self.txt_to.place(x=300,y=250,width=350,height=30)
        
        subject=Label(self.root,text="Subject",font=("times new roman",18),bg="white")
        subject.place(x=50,y=300)
        self.txt_subject=Entry(self.root,width=15,font=("times new roman",15),bg="lightyellow")
        self.txt_subject.place(x=300,y=300,width=450,height=30)
        
        message=Label(self.root,text="Message",font=("times new roman",18),bg="white")
        message.place(x=50,y=350)
        self.txt_message=Text(self.root,width=15,font=("times new roman",12),bg="lightyellow")
        self.txt_message.place(x=300,y=350,width=650,height=120)
        
        #===============================Status=====================================#
        self.lbl_total=Label(self.root,font=("times new roman",18),bg="white")
        self.lbl_total.place(x=50,y=490)
        
        self.lbl_sent=Label(self.root,font=("times new roman",18),bg="white",fg="green")
        self.lbl_sent.place(x=300,y=490)
        
        self.lbl_left=Label(self.root,font=("times new roman",18),bg="white",fg="orange")
        self.lbl_left.place(x=420,y=490)
        
        self.lbl_failed=Label(self.root,font=("times new roman",18),bg="white",fg="red")
        self.lbl_failed.place(x=550,y=490)
        
        self.check_file_exist()
    
                
    
    def browser_file(self):
        op=filedialog.askopenfile(initialdir='/',title="Select Excel file for emails",filetype=(("All Files","*.*"),("Excel file",".xlsx"))) 
        if op!=None:
            data=pd.read_excel(op.name)
            
            if 'Email' in data.columns:
                self.emails=list(data['Email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split("/")[-1]))
                    self.txt_to.config(state='readonly')
                    self.lbl_total.config(text="Total: "+str(len(self.emails)))
                    self.lbl_sent.config(text="SENT: ")
                    self.lbl_left.config(text="LEFT: ")
                    self.lbl_failed.config(text="FAILED: ")
                else:
                     messagebox.showerror("Error","This file doesn't have any Email",parent=self.root)
            else:
                 messagebox.showerror("Error","Please select file which has Email columns",parent=self.root)
            
            
    
    
    
    def send_email(self):
        x=len(self.txt_message.get(1.0,END))
        if  self.txt_to.get()=="" or self.txt_subject.get()=="" or x==1:
            messagebox.showerror("Error","All field are required",parent=self.root)
        else:
            if self.var_choice.get()=="single":
                status=email_function.email_send_fun(self.txt_to.get(),self.txt_subject.get(),self.txt_message.get('1.0',END),self.from_,self.pass_)
                if status=="s":
                    messagebox.showinfo("Success","Email has been sent",parent=self.root)
                if status=="f":    
                    messagebox.showerror("Error","Email not sent,Try Again",parent=self.root)
            if self.var_choice.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send_fun(x,self.txt_subject.get(),self.txt_message.get('1.0',END),self.from_,self.pass_)
                    if status=="s":
                        self.s_count+=1
                    if status=="f":
                        self.f_count+=1
                    self.status_bar()
                    
                messagebox.showinfo("Success","Email has been sent,Please check status",parent=self.root)
    
    
    def status_bar(self):
        self.lbl_total.config(text="STATUS: "+str(len(self.emails))+"=>>")
        self.lbl_sent.config(text="SENT: "+str(self.s_count))
        self.lbl_left.config(text="LEFT: "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="FAILED: "+str(self.f_count))
        self.lbl_total.update()          
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()
    
    
    def check_single_or_bulk(self):
        
        if self.var_choice.get()=="single":
            #messagebox.showinfo("success","Single",parent=self.root)
            self.btn_browser.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.clear1() 
            
        if self.var_choice.get()=="bulk":
            #messagebox.showinfo("success","Bulk",parent=self.root)
            self.btn_browser.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.txt_to.config(state='readonly')
            
    
    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_subject.delete(0,END)
        self.txt_message.delete('1.0',END)
        self.var_choice.set("single")
        self.btn_browser.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_left.config(text="")
        self.lbl_failed.config(text="")
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
 #====================================================Setting Window===================================================#    



    def setting_window(self):
        self.root2=Toplevel()
        self.root2.title("Setting")
        self.root2.geometry("700x350+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(background="White")
        self.check_file_exist()
        
        #=============================Title===============================#
        
        title2=Label(self.root2,text="Credential Setting",image=self.setting_icon,compound=LEFT,font=("Goudy old style",48,"bold"),padx=10,bg="#222A35",fg="white",anchor="w")
        title2.place(x=0,y=0,relwidth=1)
        desc=Label(self.root2,text="Enter the Email address and password from which to send the all emails",font=("calibr (body)",14),bg="#FFD966",fg="#262626")
        desc.place(x=0,y=80,relwidth=1)
        
        #=========================Main=====================================#
        
        from_=Label(self.root2,text="Email Address",font=("times new roman",18),bg="white")
        from_.place(x=50,y=150)
        self.txt_from=Entry(self.root2,width=15,font=("times new roman",15),bg="lightyellow")
        self.txt_from.place(x=250,y=150,width=330,height=30)
        
        pass_=Label(self.root2,text="Password",font=("times new roman",18),bg="white")
        pass_.place(x=50,y=200)
        self.txt_pass=Entry(self.root2,width=15,font=("times new roman",15),bg="lightyellow",show="*")
        self.txt_pass.place(x=250,y=200,width=330,height=30)
        
        #==========================Buttons=====================================#

        btn_clear1=Button(self.root2,text="CLEAR",command=self.clear2,font=("times new roman",18,"bold"),activebackground="#262626",activeforeground="white",bg="#262626",fg="white",cursor="hand2")
        btn_clear1.place(x=300,y=260,width=120,height=30)
        btn_save=Button(self.root2,text="SAVE",command=self.save_email,font=("times new roman",18,"bold"),activebackground="#00B0F0",activeforeground="white",bg="#00B0F0",fg="white",cursor="hand2")
        btn_save.place(x=430,y=260,width=120,height=30)
        
        self.txt_from.insert(0,self.from_)
        self.txt_pass.insert(0,self.pass_)
        
    def clear2(self):
        self.txt_from.delete(0,END)
        self.txt_pass.delete(0,END)
        
        
        
    def check_file_exist(self):  
        if os.path.exists("important.txt")==False:
            f=open('important.txt','w')
            f.write(",")
            f.close()
        f2=open('important.txt','r')
        self.credentials=[]
        for i in f2:
            self.credentials.append( [i.split(",")[0] , i.split(",")[1]] )
            
        self.from_=self.credentials[0][0]
        self.pass_=self.credentials[0][1]
           
    def save_email(self):
        if  self.txt_from.get()=="" or self.txt_pass.get()=="" :
            messagebox.showerror("Error","All field are required",parent=self.root2)
        else:
            f=open('important.txt','w')
            f.write(self.txt_from.get()+","+self.txt_pass.get())
            f.close()
            messagebox.showinfo("Success","Saved Succesfully",parent=self.root2)
            self.check_file_exist()
root=Tk()
obj=Bulk_Email(root)
root.mainloop()


# In[8]:


print("iwbevvaisvqxiqwy")


# In[ ]:




