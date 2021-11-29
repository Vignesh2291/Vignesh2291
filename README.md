- 👋 Hi, I’m @Vignesh2291
- 👀 I’m interested in ...Data
- 🌱 I’m currently learning ...Python,SQL,Power BI
- 💞️ I’m looking to collaborate on ...Data
- 📫 How to reach me ...LinkedIn

<!---
Vignesh2291/Vignesh2291 is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
https://youtu.be/i3CSD7bMMbg
https://www.youtube.com/watch?v=5X5LWcLtkzg
import pandas as pd
import getpass 
import tkinter as Tk
import sqlite3
import datetime as dt
from ttkwidgets import autocomplete
import win32gui
import babel.numbers
import openpyxl
import xlsxwriter
import xlwings as xw
 
from typing import Awaitable
from os import strerror,startfile,walk
from tkcalendar import DateEntry
from ttkwidgets.autocomplete import AutocompleteCombobox, autocompletecombobox
from tkinter import *
from tkinter import ttk
from tkinter.ttk import Combobox, Treeview
from tkinter import messagebox
from datetime import datetime
from tkinter import filedialog as fd
#import matplotlib.pyplot as plt

########################################################################################################################################################
###############################__________________________Basic Needs_________________________________###################################################
########################################################################################################################################################
Icon_Image=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\EQ_HD.ico'
Data_Base_Support=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn\Support.db'
mypath=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'
DB_Name = []
for (dirpath, dirnames, filenames) in walk(mypath):
    DB_Name.extend(filenames)
    break
DB_Name = [val for val in DB_Name if val.endswith(".db")]
DB_Name.remove("Support.db")

def WindowCheck():
    WindowsList = []
    def winEnumHandler( hwnd, ctx ):
        if win32gui.IsWindowVisible( hwnd ):
            WindowsList.append(win32gui.GetWindowText( hwnd ))

    win32gui.EnumWindows( winEnumHandler, None )
    return WindowsList
#---Select Query---
def SelectQueryfun(query, record,Data_Base):
    try:
        conn = sqlite3.connect(Data_Base,timeout=30)
        c = conn.cursor()
        c.execute(query, record)
        data = []
        for row in c.fetchall():
            data.append(row)
        conn.close()
        return data
    except:
        conn.close()
#---View Query---
def ViewQueryfun(query,Data_Base):
    try:
        conn = sqlite3.connect(Data_Base,timeout=30)
        c = conn.cursor()
        c.execute(query)
        data = []
        for row in c.fetchall():
            data.append(row)
        conn.close()
        return data
    except:
        conn.close()
#----Insert Q----
def InsertQ(ColumQ,valQ,Data_Base):
    try:
        global Transactions
        conn = sqlite3.connect(Data_Base,timeout=30)
        cursor=conn.cursor()
        cursor.execute(ColumQ,valQ)
        conn.commit()
        conn.close()
    except:
        conn.close()
#----Combodropdown1----
def Combodrop1(Condition2,F_Que,Data_Base):
    try:
        conn = sqlite3.connect(Data_Base,timeout=30)
        cursor = conn.cursor()    
        cursor.execute(Condition2, (F_Que,))
        result = []
        for row in cursor.fetchall():
            result.append(row[1])
        conn.close()
        return result
    except:
        conn.close()          

#----Combodropdown----
def Combodrop(Condition,Data_Base):
    try:
        conn = sqlite3.connect(Data_Base,timeout=30)
        cursor=conn.cursor()
        cursor.execute(Condition)
        result = []
        for row in cursor.fetchall():
            result.append(row[0])
        conn.close()
        return result
    except:
        conn.close()
#---User Details ---
UserID = getpass.getuser().lower()
S_Q="select * from TblUser_data WHERE User_ID = (?) "
S_L=[UserID]
rows=SelectQueryfun(S_Q,S_L,Data_Base_Support)
if (len(rows))==0:
    Access_Offinn=Tk()
    Access_Offinn.title("Office_Inn Access")
    Access_Offinn.geometry("400x400") 
    Access_Offinn.resizable(0,0)
    Access_Offinn.config(bg="white")
    Access_Offinn.iconbitmap(Icon_Image)
    messagebox.showerror("Data Missing","User Details missing in Data table, Reach Innovation team to get access")
    quit()
TName1=str()
for row in rows:
    U_Name=row[1]
    T_Name=row[2]
    U_Access=row[3]
    A_Access=row[4]
    
    com_query1 = ('Select distinct(Team_Name) as class from TblTeam_data')
    rows=ViewQueryfun(com_query1,Data_Base_Support)
    Team_list = []
    for i in rows:
        Team_list.append(i[0])
    
    if T_Name not in Team_list:
        Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn\Other.db'
    else:
        TName=T_Name.replace(" ", "_")
        Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+TName+'.db'
 
if A_Access!="Yes":
    messagebox.showerror("Waring","Access denied. Reach Innovation team to get Admin access")
    quit()
    
#---Team Name--
query1 = ('Select distinct(Team_Name) as class from TblTeam_data')
rows=ViewQueryfun(query1,Data_Base_Support)
TeamName = []
for i in rows:
    TeamName.append(i[0])

 
########################################################################################################################################################
#----------------Home Page---------------------------
########################################################################################################################################################
def Main_Function():
    Officeinn=Tk()
    Officeinn.title("Office_Inn_Report ")
    Officeinn.geometry("1300x800+300+250") 
    Officeinn.resizable(0,0)
    Officeinn.config(bg="white")
    Officeinn.iconbitmap(Icon_Image)
    frame1=Frame(Officeinn,bg="white",width=850,height=550,relief="solid", borderwidth=0,highlightcolor="white")
    frame1.pack(fill=BOTH, padx=1,pady=1,expand=True)
    frame1.propagate(False)
    
    Label_UserName=Label(frame1,bg='white',font=("Calibri",13,"bold"),borderwidth=0, highlightbackground=None,highlightcolor=None,text='Welcome '+U_Name)
    Label_UserName.place(relx=0.04,rely=0.04)

    sep1 = Frame(frame1,bg="white",relief="solid", borderwidth=0,highlightcolor="blue", highlightthickness=2,highlightbackground="#41729c")
    sep1.place(y=56,width=1500,height=1)
    #sep2 = Frame(frame1,bg="white",relief="solid", borderwidth=0,highlightcolor="blue", highlightthickness=2,highlightbackground="#41729c")
    #sep2.place(y=156,width=1500,height=1)
    sep3 = Frame(frame1,bg="white",relief="solid", borderwidth=0,highlightcolor="blue", highlightthickness=2,highlightbackground="#41729c")
    #sep3.place(y=456-75,width=1500,height=1)
    
    def Team_change(a):
        for item in Tree1.get_children():
            Tree1.delete(item)
        Ent_1.config(text="")
        Ent_2.config(text="")
        Ent_3.config(text="")
        Ent_4.config(text="")
        Ent_5.config(text="")
        Ent_6.config(text="")
        Ent_7.config(text="")
        Ent_8.config(text="")
        Ent_9.config(text="")
        Ent_10.config(text="")
        Ent_11.config(text="")
        Ent_12.config(text="")
        Ent_13.config(text="")
        Ent_14.config(text="") 
        
    
    L_TeamName= Label(frame1,text="Team Name",fg="#515056",bg="white",bd=0, font=("Calibri", 13, "bold"))
    L_TeamName.place(x=50,y=115-30)
    C_TeamName = AutocompleteCombobox(frame1 ,font=("Calibri",12),foreground="#515056",completevalues=TeamName)
    C_TeamName.place(x=50,y=145-30,width=260)
    C_TeamName.bind("<<ComboboxSelected>>",Team_change)
    C_TeamName.insert(0,T_Name)
    
    L_Date=Label(frame1,text="Date",fg="#515056",bg="white",bd=0, font=("Calibri", 13, "bold"))
    L_Date.place(x=400+50,y=115-30)
    E_Date=DateEntry(frame1, background= "black", foreground= "white",bd=2,font=('Calibri',11),date_pattern='dd-MM-yyyy')
    E_Date.place(x=400+50,y=145-30,width=110, height=20+5)
    E_Date.bind('<ButtonRelease-1>',Team_change)
    
    #----------------------------------------------------------------------------------------------------------------------------------------------#
    
    D_Count = Label(frame1, width=10,text="Daily Count", fg='#FA8072',bg="white",font=("Calibri",12,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    D_Count.place(x=185,y=155)
    MTD_Count = Label(frame1, width=10,text="MTD",fg='#FA8072', bg="white",font=("Calibri",12,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    MTD_Count.place(x=300,y=155)
    
    D_Count1 = Label(frame1, width=10,text="Daily Count", fg='#FA8072',bg="white",font=("Calibri",12,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    D_Count1.place(x=535+50,y=155)
    MTD_Count1 = Label(frame1, width=10,text="MTD",fg='#FA8072', bg="white",font=("Calibri",12,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    MTD_Count1.place(x=650+50,y=155)
    
    
    
    
    L_P_No=Label(frame1,text="Production Count:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_P_No.place(x=50,y=185)
    Ent_1 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_1.place(x=185,y=185)
    Ent_2 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_2.place(x=300,y=185)
    
    L_QC_No=Label(frame1,text="QC Count:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_QC_No.place(x=50,y=185+50)
    Ent_3 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_3.place(x=185,y=185+50)
    Ent_4 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_4.place(x=300,y=185+50)
    
    L_IT_Checked=Label(frame1,text="Items Checked:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_IT_Checked.place(x=50,y=185+100)
    Ent_5 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_5.place(x=185,y=185+100)
    Ent_6 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_6.place(x=300,y=185+100)
    
    L_Err=Label(frame1,text="Error:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_Err.place(x=50,y=185+150)
    Ent_7 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_7.place(x=185,y=185+150)
    Ent_8 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_8.place(x=300,y=185+150)
    
    
    L_Pro=Label(frame1,text="Productivity %:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_Pro.place(x=400+50,y=185)
    Ent_9 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_9.place(x=535+50,y=185)
    Ent_10 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_10.place(x=650+50,y=185)
    
    L_Utiliz=Label(frame1,text="Utilisation %:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_Utiliz.place(x=400+50,y=185+50)
    Ent_11 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_11.place(x=535+50,y=185+50)
    Ent_12 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_12.place(x=650+50,y=185+50)
    
    
    L_QC_Per=Label(frame1,text="Quality %:",fg="#515056",bg="white",bd=0, font=("Calibri", 11, "bold"))
    L_QC_Per.place(x=400+50,y=185+100)
    Ent_13 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_13.place(x=535+50,y=185+100)
    Ent_14 = Label(frame1, width=10,text="", bg="white",font=("Calibri",10,"bold"),highlightbackground=None,highlightcolor=None,justify='center')
    Ent_14.place(x=650+50,y=185+100)
    
    Tree1=Treeview(frame1,columns=(1,2,3),show='headings',height = 17, selectmode = "extended")
    Tree1.column("# 1",anchor=W, stretch=NO,width=400 )
    Tree1.heading(1, text="Transaction",anchor=W)
    Tree1.column("# 2",anchor=CENTER, stretch=NO, width=150)
    Tree1.heading(2, text="Production",anchor=CENTER)
    Tree1.column("# 3",anchor=CENTER, stretch=NO, width=160)
    Tree1.heading(3, text="QC",anchor=CENTER)

    Tree1.place(x=50,y=420-30)
    style = ttk.Style()
    style.theme_use("vista")
    style.configure('Treeview.Heading',background='#9e9d9d',foreground='Black',font=("Calibri", 12, "bold"),)
    vsb = ttk.Scrollbar(frame1, orient="vertical", command=Tree1.yview)
    vsb.place(x=600+145 , y=422-30 , height=363)
    Tree1.configure(yscrollcommand=vsb.set)
   
    
    
    
    def View_data(): 
        Submit_Button.config(state='disabled')
        #--Tree Update--
        if C_TeamName.get()=="" or E_Date.get()==""  :
                 messagebox.showerror("Warning", "Select required fields to view data")
                 return False
        for item in Tree1.get_children():
            Tree1.delete(item)
         
        P_Date_R= E_Date.get()
        T_Name_R=C_TeamName.get()
        Status="Diarised"
        Date1="%-"+P_Date_R[3:-5]+"-%"
         
        A_Production = pd.DataFrame()
        MTD_Production=pd.DataFrame()
        QC_Count_Df=pd.DataFrame()
        ER_Count_Df=pd.DataFrame()
        QC_Count_MTD_Df=pd.DataFrame()
        ER_Count_MTD_Df=pd.DataFrame()
        
        
        for DB in DB_Name:
            Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+DB
            try:
                conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                query1="SELECT Transactions,SUM(case when P_Type = 'Production' then A_Count ELSE 0 end) as Pro,SUM(case when P_Type = 'QC' then A_Count ELSE 0 end) as QC FROM TblProduction where TRANSACTION_TEAM='"+ T_Name_R +"' and P_Date='"+ P_Date_R +"'and A_STATUS!= '"+ Status +"' GROUP BY Transactions"
                query2="SELECT Transactions,SUM(case when P_Type = 'Production' then A_Count ELSE 0 end),SUM(case when P_Type = 'QC' then A_Count ELSE 0 end)FROM TblProduction where TRANSACTION_TEAM='"+ T_Name_R +"' and P_Date like  '"+ Date1 +"' and A_STATUS!= '"+ Status +"' GROUP BY Transactions"
                query3="select sum(A_Count) from TblQCProducation where WIT_ProcessDate='"+ P_Date_R +"' and TRANSACTION_TEAM='"+ T_Name_R +"'"
                query4="select sum(A_Count) from TblQCProducation where WIT_ProcessDate='"+ P_Date_R +"' and TRANSACTION_TEAM='"+ T_Name_R +"' and Completed_Correctly='No'"
                query5="select sum(A_Count) from TblQCProducation where WIT_ProcessDate like '"+ Date1 +"' and TRANSACTION_TEAM='"+ T_Name_R +"'"
                query6="select sum(A_Count) from TblQCProducation where WIT_ProcessDate like '"+ Date1 +"' and TRANSACTION_TEAM='"+ T_Name_R +"' and Completed_Correctly='No'"
                
                
                df1 = pd.read_sql_query(query1, conn)
                A_Production = A_Production.append(df1)
                
                df2 = pd.read_sql_query(query2, conn)
                MTD_Production = MTD_Production.append(df2)
                
                df3 = pd.read_sql_query(query3, conn)
                QC_Count_Df = QC_Count_Df.append(df3)
                
                df4 = pd.read_sql_query(query4, conn)
                ER_Count_Df = ER_Count_Df.append(df4)
                
                df5 = pd.read_sql_query(query5, conn)
                QC_Count_MTD_Df = QC_Count_MTD_Df.append(df5)
                
                df6 = pd.read_sql_query(query6, conn)
                ER_Count_MTD_Df = ER_Count_MTD_Df.append(df6)
                
                conn.close()  
                
                
            except:
                conn.close()
                messagebox.showerror("Database Crash","Office Inn crashed, Please reopen the tool")
                return False
            
        
         
        A_Production=A_Production.groupby(['TRANSACTIONS']).agg({'Pro':'sum','QC':'sum'}).reset_index()
        
        #print(A_Production)
         
        
        
        A_Production1 = A_Production.values.tolist()
        
    
        for row in A_Production1:
            Tree1.insert("", END, values=row)
            
        Pro_Count=A_Production.sum(axis=0).values[1]
        QC_Count=A_Production.sum(axis=0).values[2]
        It_Checked=QC_Count_Df.sum(axis=0).values[0]
        Err_Checked=ER_Count_Df.sum(axis=0).values[0]
        It_Checked_MTD=QC_Count_MTD_Df.sum(axis=0).values[0]
        Err_Checked_MTD=ER_Count_MTD_Df.sum(axis=0).values[0]
        
        if Err_Checked==0 or It_Checked==0:
            Error=100
        else:
            Error=(1-(Err_Checked/It_Checked))*100
        #print(round(Error,2) )
        
        if Err_Checked_MTD==0 or It_Checked_MTD==0:
            Error_MTD=100
        else:
            Error_MTD=(1-(Err_Checked_MTD/It_Checked_MTD))*100
        #print(round(Error_MTD,2) )
           
        round(Error_MTD,2)    
       
        Pro_Count_MTD=MTD_Production.sum(axis=0).values[1]
        QC_Count_MTD=MTD_Production.sum(axis=0).values[2]
        
        #--Count_Update_Screen--
        Ent_1.config(text=int(Pro_Count))
        Ent_2.config(text=int(Pro_Count_MTD))
        Ent_3.config(text=int(QC_Count))
        Ent_4.config(text=int(QC_Count_MTD))
        Ent_5.config(text=int(It_Checked))
        Ent_6.config(text=int(It_Checked_MTD))
        Ent_7.config(text=int(Err_Checked))
        Ent_8.config(text=int(Err_Checked_MTD))
        Ent_9.config(text="TBD")
        Ent_10.config(text="TBD")
        Ent_11.config(text="TBD")
        Ent_12.config(text="TBD")
        Ent_13.config(text=str(round(Error,2))+"%")
        Ent_14.config(text=str(round(Error_MTD,2))+"%") 
        
        
        #--Graph--
        #print(A_Production)
        #A_Production.plot(x ='TRANSACTIONS', y=['Pro','QC'], kind = 'line')
        
        #plt.show()
        
        
        
        
        Submit_Button.config(state='normal')
        
    
    Submit_Button = Button(frame1,text=">",width=12,height=1,bg="#00728F",foreground="white", font=("Calibri", 11,"bold"),borderwidth=3,  cursor="hand2",command=View_data )
    Submit_Button.place(x=600,y=105)
########################################################################################################################################################
#----------------Hr Producation----------------------
########################################################################################################################################################
    def H_Producation():
        global P_selection
        Clock_Button.config(state='disabled',bg="#98d4e3",font=("Calibri", 11,"bold"))
         
      
        H_Pro_w=Frame(Officeinn,bg="white",width=1300,height=900,relief="solid", borderwidth=0,highlightcolor="white")
        H_Pro_w.place(x=1,y=58)
         

        #---Team Combo---
        def T_Check(a):
            
            if len(H_Pro_C_Teamname.get())>0 :
                if H_Pro_C_Teamname.get() not in TeamName:
                    messagebox.showerror("Incorrect Team Name", "Enter valid Team name")
                    H_Pro_C_Teamname.focus_set()
                    H_Pro_C_Teamname.delete(0, END)
                else:
                    pass
            
            for item in Hr_Count.get_children():
                Hr_Count.delete(item)
            

        H_Pro_L_Teamname = Label(H_Pro_w,text="Team Name",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        H_Pro_L_Teamname.place(x=40,y=30)
        H_Pro_C_Teamname = AutocompleteCombobox(H_Pro_w ,font=("Calibri",11),foreground="#515056",completevalues=TeamName)
        H_Pro_C_Teamname.place(x=40,y=50,width=180)
        #Admin_C_Teamname.bind('<FocusOut>',T_Check)
        H_Pro_C_Teamname.bind('<<ComboboxSelected>>',T_Check)

        #---ProcessDate---
        def d_check(a):
            for item in Hr_Count.get_children():
                Hr_Count.delete(item)
             
        H_Pro_L_Date=Label(H_Pro_w,text="Date",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        H_Pro_L_Date.place(x=300,y=30)
        H_Pro_E_Date=DateEntry(H_Pro_w, background= "black", foreground= "white",bd=2,font=('Calibri',11),date_pattern='dd-MM-yyyy')
        H_Pro_E_Date.place(x=300,y=50,width=110, height=20+5)
         
        #Admin_E_Date.bind('<FocusOut>',d_check)
        H_Pro_E_Date.bind('<ButtonRelease-1>',d_check)
        #H_Pro_E_Date.bind('<<ComboboxSelected>>',d_check)

        #--Pro or QC--
        P_selection=1
        def p_type():
            global P_selection
            P_selection=v.get()
            for item in Hr_Count.get_children():
                Hr_Count.delete(item)

        v = IntVar()
        H_Pro_L_PType=Label(H_Pro_w,text="Production Type",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        H_Pro_L_PType.place(x=480,y=30)

        H_Pro_R_PType=Radiobutton(H_Pro_w,text="Production",variable=v,value=1,bg="white",command=p_type)
        H_Pro_R_PType.place(x=480,y=50)

        H_Pro_R_QType=Radiobutton(H_Pro_w,text="QC",variable=v,value=2,bg="white",command=p_type)
        H_Pro_R_QType.place(x=580,y=50)

        v.set(1)
        
        #--Submit--

        def view_Data():
            
            global P_selection
            P_Date_H_Pro= H_Pro_E_Date.get()
            T_Name_H_Pro=H_Pro_C_Teamname.get()
            
            #if P_Date_H_Pro==dt.datetime.today().strftime('%d-%m-%Y'):
            #    DB_Name.remove("Office_Inn_DB.db")

            if len(P_Date_H_Pro)<1 or len(T_Name_H_Pro)<1:
                
                messagebox.showerror("Data Missing", "Select Team Name and Process Date")
                H_Pro_w.focus_force()
                return False

            for item in Hr_Count.get_children():
                Hr_Count.delete(item)
            A_Production = pd.DataFrame()
            if P_selection==1:
                for DB in DB_Name:
                 
                    Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+DB
                    try:
                        conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                        query1=""" select  TRANSACTIONS,USERID,  
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) < 8 then A_Count ELSE 0 end) as "Less than 8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 8 then A_Count ELSE 0 end) as "8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 9 then A_Count ELSE 0 end) as "9",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 10 then A_Count ELSE 0 end) as "10",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 11 then A_Count ELSE 0 end) as "11",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 12 then A_Count ELSE 0 end) as "12",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 11 then A_Count ELSE 0 end) as "13",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 14 then A_Count ELSE 0 end) as "14",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 15 then A_Count ELSE 0 end) as "15",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 16 then A_Count ELSE 0 end) as "16",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 17 then A_Count ELSE 0 end) as "17",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) > 17  then A_Count ELSE 0 end) as "Greater than 17",
                            sum(case when END_TIME is Null then A_Count ELSE 0 end) as "Correction",
                            sum(A_count) as "Total Count"
                            from TblProduction where P_Date= '"""+ P_Date_H_Pro +"""' and TRANSACTION_TEAM= '"""+ T_Name_H_Pro +"""' and A_STATUS<> "Diarised" and P_TYPE="Production" group by TRANSACTIONS,USERID ORDER BY TRANSACTIONS ASC , USERID ASC
                        """
                        df = pd.read_sql_query(query1, conn)
                        A_Production = A_Production.append(df)
                        conn.close()  
                    except:
                        conn.close()
                        messagebox.showerror("Database Crash","Office Inn crashed, Please reopen the tool")
                        #H_Pro_w.destroy()
                        return False
                        

            elif P_selection==2:
                for DB in DB_Name:
                    Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+DB
                    try:
                        conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                        query1=""" select  TRANSACTIONS,USERID,  
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) < 8 then A_Count ELSE 0 end) as "Less than 8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 8 then A_Count ELSE 0 end) as "8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 9 then A_Count ELSE 0 end) as "9",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 10 then A_Count ELSE 0 end) as "10",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 11 then A_Count ELSE 0 end) as "11",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 12 then A_Count ELSE 0 end) as "12",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 13 then A_Count ELSE 0 end) as "13",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 14 then A_Count ELSE 0 end) as "14",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 15 then A_Count ELSE 0 end) as "15",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 16 then A_Count ELSE 0 end) as "16",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 17 then A_Count ELSE 0 end) as "17",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) > 17  then A_Count ELSE 0 end) as "Greater than 17",
                            sum(case when END_TIME is Null then A_Count ELSE 0 end) as "Correction",
                            sum(A_count) as "Total Count"
                            from TblProduction where P_Date= '"""+ P_Date_H_Pro +"""' and TRANSACTION_TEAM= '"""+ T_Name_H_Pro +"""'  and P_TYPE="QC" group by TRANSACTIONS,USERID ORDER BY TRANSACTIONS ASC 
                        """
                        df = pd.read_sql_query(query1, conn)
                        A_Production = A_Production.append(df)
                        conn.close()
                    except:
                        conn.close()
                        messagebox.showerror("Database Issue","Office Inn crashed, Please reopen the tool")
                        #H_Pro_w.destroy()
                        return False  
                     
                  
            if len(A_Production)==0:
             messagebox.showerror("Alert","No Data found")
             H_Pro_w.focus_force()
             return False
         
            A_Production = A_Production.values.tolist() 
            for row in A_Production:
                Hr_Count.insert("", END, values=row) 
            
        def Export_Data():
            global P_selection
            P_Date_H_Pro= H_Pro_E_Date.get()
            T_Name_H_Pro=H_Pro_C_Teamname.get()

            if len(P_Date_H_Pro)<1 or len(T_Name_H_Pro)<1:
                messagebox.showerror("Data Missing", "Select Team Name and Process Date")
                H_Pro_w.focus_force()
                return False
            A_Production1 = pd.DataFrame()
            if P_selection==1:
                for DB in DB_Name:
                    Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+DB
                    try:
                        conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                        query1=""" select  TRANSACTIONS,USERID,  
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) < 8 then A_Count ELSE 0 end) as "Less than 8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 8 then A_Count ELSE 0 end) as "8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 9 then A_Count ELSE 0 end) as "9",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 10 then A_Count ELSE 0 end) as "10",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 11 then A_Count ELSE 0 end) as "11",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 12 then A_Count ELSE 0 end) as "12",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 13 then A_Count ELSE 0 end) as "13",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 14 then A_Count ELSE 0 end) as "14",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 15 then A_Count ELSE 0 end) as "15",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 16 then A_Count ELSE 0 end) as "16",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 17 then A_Count ELSE 0 end) as "17",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) > 17  then A_Count ELSE 0 end) as "Greater than 17",
                            sum(case when END_TIME is Null then A_Count ELSE 0 end) as "Correction",
                            sum(A_count) as "Total Count"
                            from TblProduction where P_Date= '"""+ P_Date_H_Pro +"""' and TRANSACTION_TEAM= '"""+ T_Name_H_Pro +"""' and A_STATUS<> "Diarised" and P_TYPE="Production" group by TRANSACTIONS,USERID ORDER BY TRANSACTIONS ASC 
                        """
                        df = pd.read_sql_query(query1, conn)
                        A_Production1 = A_Production1.append(df)
                        conn.close()
                    except:
                        conn.close()
                        messagebox.showerror("System Crash","Office Inn crashed, Please reopen the tool")
                        #H_Pro_w.destroy()
                        return False
            elif P_selection==2:
                for DB in DB_Name:
                    Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+DB
                    try:
                        conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                        query1=""" select  TRANSACTIONS,USERID,  
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) < 8 then A_Count ELSE 0 end) as "Less than 8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 8 then A_Count ELSE 0 end) as "8",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 9 then A_Count ELSE 0 end) as "9",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 10 then A_Count ELSE 0 end) as "10",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 11 then A_Count ELSE 0 end) as "11",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 12 then A_Count ELSE 0 end) as "12",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 13 then A_Count ELSE 0 end) as "13",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 14 then A_Count ELSE 0 end) as "14",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 15 then A_Count ELSE 0 end) as "15",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 16 then A_Count ELSE 0 end) as "16",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) = 17 then A_Count ELSE 0 end) as "17",
                            sum(case when cast(substr(END_TIME, 11, 3) as decimal) > 17  then A_Count ELSE 0 end) as "Greater than 17",
                            sum(case when END_TIME is Null then A_Count ELSE 0 end) as "Correction",
                            sum(A_count) as "Total Count"
                            from TblProduction where P_Date= '"""+ P_Date_H_Pro +"""' and TRANSACTION_TEAM= '"""+ T_Name_H_Pro +"""'  and P_TYPE="QC" group by TRANSACTIONS,USERID ORDER BY TRANSACTIONS ASC 
                        """
                        df = pd.read_sql_query(query1, conn)
                        A_Production1 = A_Production1.append(df)
                        conn.close()
                    except:
                        conn.close()
                        messagebox.showerror("System Crash","Office Inn crashed, Please reopen the tool")
                        #H_Pro_w.destroy()
                        return False

            if len(rows)==0:
                messagebox.showerror("Alert","No Data found")
                H_Pro_w.focus_force()
                return False
            else:
                #df1 = pd.read_sql_query(query, conn)
                messagebox.showinfo('Completed','Report Created!')
                H_Pro_w.focus_force()
                xw.view(A_Production1, table=False)
            conn.close()
            
            

        S_button=Button(H_Pro_w,text=">", command=view_Data,font=("Bauhaus 93", 10,"bold"),bg='#00728F',fg='white',cursor="hand2")
        S_button.place(x=670,y=50,width=30)

        Export_button=Button(H_Pro_w,text="Export Report", command=Export_Data,width=12,height=1,bg="#00728F",foreground="white", font=("Calibri", 11,"bold"),borderwidth=3,cursor="hand2" )
        Export_button.place(x=750,y=45)

        #---Data in Tree--
        Hr_Count=Treeview(H_Pro_w,columns=(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16),show='headings',height = 30)
        Hr_Count.column("# 1",anchor=W, stretch=NO,width=150+20+20 )
        Hr_Count.heading(1, text="Transactions",anchor=W)
        

        Hr_Count.column("# 2",anchor=W, stretch=NO, width=70 )
        Hr_Count.heading(2, text="User ID",anchor=W)

        Hr_Count.column("# 3",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(3, text="<8:00",anchor=CENTER)
        
        Hr_Count.column("# 4",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(4, text="8:00",anchor=CENTER)

        Hr_Count.column("# 5",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(5, text="9:00",anchor=CENTER)

        Hr_Count.column("# 6",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(6, text="10:00",anchor=CENTER)

        Hr_Count.column("# 7",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(7, text="11:00",anchor=CENTER)

        Hr_Count.column("# 8",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(8, text="12:00",anchor=CENTER)

        Hr_Count.column("#9",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(9, text="13:00",anchor=CENTER)

        Hr_Count.column("# 10",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(10, text="14:00",anchor=CENTER)

        Hr_Count.column("# 11",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(11, text="15:00",anchor=CENTER)

        Hr_Count.column("# 12",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(12, text="16:00",anchor=CENTER)

        Hr_Count.column("# 13",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(13, text="17:00",anchor=CENTER)

        Hr_Count.column("# 14",anchor=CENTER, stretch=NO, width=70)
        Hr_Count.heading(14, text=">17:00",anchor=CENTER)

        Hr_Count.column("# 15",anchor=CENTER, stretch=NO, width=80)
        Hr_Count.heading(15, text="Correction",anchor=CENTER)

        Hr_Count.column("# 16",anchor=CENTER, stretch=NO, width=115)
        Hr_Count.heading(16, text="Total Count",anchor=W)
        
        Hr_Count.place(x=0,y=105)
        style = ttk.Style()
        style.theme_use("vista")
        style.configure('Treeview.Heading',background='#9e9d9d',foreground='Black',font=("Calibri", 11, "bold"))
        vsb1 = ttk.Scrollbar(Hr_Count, orient="vertical", command=Hr_Count.yview)
        vsb1.place(x=562+697+20 , y=2, height=343+280)
        Hr_Count.configure(yscrollcommand=vsb1.set)
        
        #columns=('Transactions',)
        
        # def treeview_sort_column(tv, col, reverse):
        #     l = [(tv.set(k, col), k) for k in tv.get_children('')]
        #     l.sort(key=lambda t: int(t[0]), reverse=reverse)
        # Hr_Count.heading(columns, text=columns,command=lambda c=columns: treeview_sort_column(Hr_Count, c, False))
        

        def Dash_avoid_H_Pro():
            #H_Pro_w.destroy()
            H_Pro_w.destroy()
            Hme_Button.place_forget()
            Officeinn.focus_force()
            
            Clock_Button.config(state='normal',bg="white",font=("Calibri", 11,"bold"))

        Hme_Button = Button(frame1, text="⌂",background='white',cursor="hand2",command=Dash_avoid_H_Pro,font=("Calibri", 11,"bold") )
        Hme_Button.place(relx=0.01,rely=0.04)

########################################################################################################################################################
#----------------Admin-------------------------------
########################################################################################################################################################
    def Admin_P():
            
        Admin_Button.config(state='disabled',bg="#98d4e3",font=("Calibri", 11,"bold"))
        global F_Count
        Admin_w= Toplevel(Officeinn)
        Admin_w.geometry("600x530+700+378")
        Admin_w.resizable(0,0)
        Admin_w.title("Admin")
        Admin_w.config(bg='white')
        Admin_w.iconbitmap(Icon_Image)

        L_top=Frame(Admin_w,width=590,height=2,bg='#00728F')
        #L_top.place(x=5,y=100)
        L_Mid=Frame(Admin_w,width=590,height=2,bg='#00728F')
        #L_Mid.place(x=5,y=335)
        L_Bot=Frame(Admin_w,width=590,height=2,bg='#00728F')
        L_Bot.place(x=5,y=470-5)

        #---Team Combo---
        def T_Check(a):
            
            for item in Trans1.get_children():
                Trans1.delete(item)
            #Admin_C_UserID.delete(0, END)
            Admin_E_Trans.config(state='normal') 
            Admin_E_Trans.delete(0,END)
            Admin_E_Trans.config(state='disabled') 
            Admin_E_P_Type.config(state='normal') 
            Admin_E_P_Type.delete(0,END)
            Admin_E_P_Type.config(state='disabled')
            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,0)
            Admin_E_P_Count.config(state='disabled')
            Admin_L_UpCount.place_forget()
            P_Count_Admin=0
            
            if len(Admin_C_Teamname.get())>0 :
                if Admin_C_Teamname.get() not in TeamName:
                    messagebox.showerror("Incorrect Team Name", "Enter valid Team name")
                    Admin_C_Teamname.focus_set()
                    Admin_C_Teamname.delete(0, END)
                else:
                    pass
        
        
        Admin_L_Teamname = Label(Admin_w,text="Team Name",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_Teamname.place(x=40,y=30)
        Admin_C_Teamname = AutocompleteCombobox(Admin_w ,font=("Calibri",11),foreground="#515056",completevalues=TeamName)
        Admin_C_Teamname.place(x=40,y=50,width=180)
        #Admin_C_Teamname.bind('<FocusOut>',T_Check)
        Admin_C_Teamname.bind('<<ComboboxSelected>>',T_Check)

        #---U_ID---
        com_query2 = ('Select * from TblUser_data')
        U_ID_S=ViewQueryfun(com_query2,Data_Base_Support)
        U_List = []
        for i in U_ID_S:
            U_List.append(i[0])

        def UID_Check(a):
            for item in Trans1.get_children():
                Trans1.delete(item)
            #Admin_C_UserID.delete(0, END)
            Admin_E_Trans.config(state='normal') 
            Admin_E_Trans.delete(0,END)
            Admin_E_Trans.config(state='disabled') 
            Admin_E_P_Type.config(state='normal') 
            Admin_E_P_Type.delete(0,END)
            Admin_E_P_Type.config(state='disabled')
            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,0)
            Admin_E_P_Count.config(state='disabled')
            Admin_L_UpCount.place_forget()
            
            if len(Admin_C_UserID.get())>0 :
                if Admin_C_UserID.get() not in U_List:
                    messagebox.showerror("Warning", "Enter valid User ID")
                    Admin_C_UserID.focus_set()
                    Admin_C_UserID.delete(0, END)
                else:
                    pass
        
        Admin_L_UserID = Label(Admin_w,text="User ID",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_UserID.place(x=250+20,y=30)
        Admin_C_UserID = AutocompleteCombobox(Admin_w ,font=("Calibri",11),completevalues=U_List,foreground="#515056")
        Admin_C_UserID.place(x=250+20,y=50,width=100)
        Admin_C_UserID.bind('<<ComboboxSelected>>',UID_Check)

        #---ProcessDate---
        def d_check(a):
            for item in Trans1.get_children():
                Trans1.delete(item)
            
            #Admin_C_UserID.delete(0, END)
            Admin_E_Trans.config(state='normal') 
            Admin_E_Trans.delete(0,END)
            Admin_E_Trans.config(state='disabled') 
            Admin_E_P_Type.config(state='normal') 
            Admin_E_P_Type.delete(0,END)
            Admin_E_P_Type.config(state='disabled')
            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,0)
            Admin_E_P_Count.config(state='disabled')
            Admin_L_UpCount.place_forget()
            
        Admin_L_Date=Label(Admin_w,text="Date",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_Date.place(x=420,y=30)
        Admin_E_Date=DateEntry(Admin_w, background= "black", foreground= "white",bd=2,font=('Calibri',11),date_pattern='dd-MM-yyyy')
        Admin_E_Date.place(x=420,y=50,width=110, height=20+5)
        today = dt.datetime.today()
        offset = max(1, (today.weekday() + 6) % 7 - 3)
        timedelta = dt.timedelta(offset)
        most_recent = today - timedelta
        P_Date=most_recent.strftime('%d-%m-%Y')
        Admin_E_Date._set_text(P_Date)
        Admin_E_Date.bind('<ButtonRelease-1>',d_check)
        
        #--Enter Button
        def Get_data():

            if Admin_C_Teamname.get()=="" or Admin_C_UserID.get()=="" or Admin_E_Date.get()=="":
                    messagebox.showerror("Warning", "Select required fields to view data")
                    return False

            for item in Trans1.get_children():
                Trans1.delete(item)
            User_ID_Admin=Admin_C_UserID.get()
            P_Date_Admin= Admin_E_Date.get()
            T_Name_Admin=Admin_C_Teamname.get()

            A_Production = pd.DataFrame()
            for DB in DB_Name:
                Data_Base=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+DB
                try:
                    conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                    query1="SELECT Transactions,A_STATUS,SUM(case when P_Type = 'Production' then A_Count ELSE 0 end)FROM TblProduction where USERID='"+ User_ID_Admin +"' and P_Date='"+ P_Date_Admin +"' and TRANSACTION_TEAM='"+ T_Name_Admin +"' and P_Type = 'Production'  GROUP BY Transactions,A_STATUS ORDER BY Transactions and A_STATUS ASC"
                    df = pd.read_sql_query(query1, conn)
                    A_Production = A_Production.append(df)
                    conn.close()  
                except:
                    conn.close()
                    messagebox.showerror("Database Crash","Office Inn crashed, Please reopen the tool")
                    return False
            
            A_Production = A_Production.values.tolist()
            
            if len(A_Production)==0:
                messagebox.showerror("Alert","No Data found")
                Admin_w.focus_force()
                return False
            for row in A_Production:
                Trans1.insert("", END, values=row) 
                

                
        SSD = Button(Admin_w,text="↵",width=2,height=1,bg="#00728F",foreground="white", font=("Calibri", 10,"bold"),borderwidth=2,command=Get_data,cursor="hand2" )
        SSD.place(x=550,y=50)

        #--Process Line items
        Trans1=Treeview(Admin_w,columns=(1,2,3),show='headings',height = 10)
        Trans1.column("# 1",anchor=W, stretch=NO,width=195+38 )
        Trans1.heading(1, text="Transactions",anchor=W)
        Trans1.column("# 2",anchor=W, stretch=NO, width=145+38)
        Trans1.heading(2, text="Process Type",anchor=W)
        Trans1.column("# 3",anchor=CENTER, stretch=NO, width=125+38)
        Trans1.heading(3, text="Production Count",anchor=CENTER)
        Trans1.place(x=8,y=105)
        style = ttk.Style()
        style.theme_use("vista")
        style.configure('Treeview.Heading',background='#9e9d9d',foreground='Black',font=("Calibri", 11, "bold"))
        vsb1 = ttk.Scrollbar(Trans1, orient="vertical", command=Trans1.yview)
        vsb1.place(x=562 , y=2, height=223)
        Trans1.configure(yscrollcommand=vsb1.set)

        def selectItem(a):
            global P_Type_Admin,P_Count_Admin,Trans_Admin
            curItem = Trans1.focus()
            treedict = (Trans1.item(curItem))
            Trans_Admin = treedict['values'][0]
            P_Type_Admin = treedict['values'][1]
            P_Count_Admin=treedict['values'][2]
        
            Admin_E_Trans.config(state='normal') 
            Admin_E_Trans.delete(0,END)
            Admin_E_Trans.insert(0,Trans_Admin)
            Admin_E_Trans.config(state='disabled') 

            Admin_E_P_Type.config(state='normal') 
            Admin_E_P_Type.delete(0,END)
            Admin_E_P_Type.insert(0,P_Type_Admin)
            Admin_E_P_Type.config(state='disabled')

            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,0)
            Admin_E_P_Count.config(state='disabled')
            Admin_L_UpCount.place_forget()

        Trans1.bind('<ButtonRelease-1>', selectItem)

        #--Selected Tranaction--
        Admin_L_Trans=Label(Admin_w,text="Selected Tranaction",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_Trans.place(x=50,y=350)
        Admin_E_Trans=Entry(Admin_w, background= "white", foreground= "black",bd=2,font=('Calibri',10),state='disabled')
        Admin_E_Trans.place(x=50,y=370,width=180, height=25)

        #--Selected P Type--
        Admin_L_P_Type=Label(Admin_w,text="Selected P_Type",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_P_Type.place(x=380,y=350)
        Admin_E_P_Type=Entry(Admin_w, background= "white", foreground= "black",bd=2,font=('Calibri',10),state='disabled')
        Admin_E_P_Type.place(x=380,y=370,width=180, height=25)
        #--0 Count--
        Admin_L_P_Count=Label(Admin_w,text="Count Regulator",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_P_Count.place(x=50,y=400+10)
        Admin_E_P_Count=Entry(Admin_w, background= "white", foreground= "black",bd=2,font=('Calibri',10,"bold"),justify='center')
        Admin_E_P_Count.place(x=90,y=420+10,width=29, height=25)
        Admin_E_P_Count.insert(0,0)
        Admin_E_P_Count.config(state='disabled')
        #--Updated Count--
        Admin_L_UpCount=Label(Admin_w,text="Count",fg="#515056",bg="white",bd=0, font=("Calibri", 12, "bold"))

        #-- + & - Buttons--
        def adding():
            global P_Count_Admin,F_Count

            if len(Trans1.selection())==0:
                messagebox.showerror("Warning","Select the transaction to Update")
                return False
            
            Add = int(Admin_E_P_Count.get())+1
            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,Add)
            Admin_E_P_Count.config(state='disabled')
            F_Count=int(Admin_E_P_Count.get())+P_Count_Admin
            Admin_L_UpCount.config(text="Final Count : "+ str(F_Count))
            Admin_L_UpCount.place(x=240+10,y=430)
        
        def sub():
            global P_Count_Admin,F_Count
            if len(Trans1.selection())==0:
                messagebox.showerror("Warning","Select the transaction to update")
                return False
            sub1 = int(Admin_E_P_Count.get())-1
            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,sub1)
            Admin_E_P_Count.config(state='disabled')
            F_Count=int(Admin_E_P_Count.get())+P_Count_Admin
            if F_Count<0:
                messagebox.showerror("Warning","Negative values not allowed")
                Admin_L_UpCount.config(text="Final Count : "+ str(P_Count_Admin))
                Admin_L_UpCount.place(x=240+10,y=430)
                Admin_E_P_Count.config(state='normal')
                Admin_E_P_Count.delete(0,END)
                Admin_E_P_Count.insert(0,0)
                Admin_E_P_Count.config(state='disabled')
                return False
            Admin_L_UpCount.config(text="Final Count : "+ str(F_Count))
            Admin_L_UpCount.place(x=240+10,y=430)

        Sub_Butt = Button(Admin_w,text="-",width=2,height=1,bg="#00728F",foreground="white", font=("Calibri", 10,"bold"),borderwidth=2,command=sub, activebackground='#fc453f',cursor="hand2" )
        Sub_Butt.place(x=50,y=420+10)

        Add_Butt = Button(Admin_w,text="+",width=2,height=1,bg="#00728F",foreground="white", font=("Calibri", 10,"bold"),borderwidth=2,command=adding,activebackground="#00ab1a",cursor="hand2" )
        Add_Butt.place(x=135,y=420+10)

        def Dash_avoid():
            Admin_w.destroy()
            Officeinn.focus_force()
            Admin_Button.config(state='normal',bg="#00728F",font=("Calibri", 11,"bold"))


        Admin_w.grab_set() 
        Admin_w.protocol("WM_DELETE_WINDOW",Dash_avoid)

        #--Admin Submit--
        def admin_Submit():

            global P_Count_Admin,P_Count_Admin
            
            if Admin_C_Teamname.get()=="" or Admin_C_UserID.get()=="":
                    messagebox.showerror("Warning","Select team & User Details to update")
                    return False

            if len(Trans1.selection())==0:
                messagebox.showerror("Warning","Select the transaction to update")
                return False
            F_Count=int(Admin_E_P_Count.get())+P_Count_Admin
            if F_Count == P_Count_Admin:
                    messagebox.showerror("Warning","No changes made!!")
                    return False
            
            U_Date=datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            Admin_PDate=Admin_E_Date.get()
            Admin_UserID=Admin_C_UserID.get()
            
            Admin_T_Name=Admin_E_Trans.get()
            Admin_User_Team=Admin_C_Teamname.get()
            Admin_A_Count=Admin_E_P_Count.get()
            Admin_P_Type="Production"
            Admin_A_Status=Admin_E_P_Type.get()
            Admin_comments= f'Updated by {U_Name} on {U_Date}'
            
            
            S_Q="""select * from TblTeam_data where TRANSACTION_N = ? and TEAM_NAME = ? """
            S_L=[Admin_T_Name,Admin_User_Team]
            rows1=SelectQueryfun(S_Q,S_L,Data_Base_Support)
        
            for i in rows1:
                S_Team_N1=i[2]

            if S_Team_N1==Admin_User_Team:
                S_Team_N=""
            else:
                S_Team_N=S_Team_N1
        
            S_Q="select * from TblUser_data WHERE User_ID = (?) "
            S_L=[Admin_UserID]
            #print(Admin_UserID)
            rows=SelectQueryfun(S_Q,S_L,Data_Base_Support)
            for row in rows:
                U_Name1=row[1]
                T_Name1=row[2]
            User_Team=T_Name1
            #print(T_Name1)
            if T_Name1 not in Team_list:
                    Data_Base1=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn\Other.db'
            else:
                # if T_Name1=="CEC-UK" or T_Name1=="CEC-US":
                #     T_Name1="Customer Experience Centre"
                T_Name1=User_Team.replace(" ", "_")
                #print(T_Name1)
                Data_Base1=r'\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Team Database\Office Inn'+'\\'+T_Name1+'.db'
            
                #print(T_Name1)

            ColumQ="INSERT INTO TBlProduction (P_DATE,USERID,USER_TEAM,TRANSACTIONS,Sub_Team_Name,TRANSACTION_TEAM,A_COUNT,P_TYPE,A_STATUS,COMMENTS) VALUES (?,?,?,?,?,?,?,?,?,?);"
            ValQ=[Admin_PDate,Admin_UserID,User_Team,Admin_T_Name,S_Team_N,Admin_User_Team,Admin_A_Count,Admin_P_Type,Admin_A_Status,Admin_comments]
            InsertQ(ColumQ,ValQ,Data_Base1)

            Admin_E_Trans.config(state='normal') 
            Admin_E_Trans.delete(0,END)
            Admin_E_Trans.config(state='disabled') 
            Admin_E_P_Type.config(state='normal') 
            Admin_E_P_Type.delete(0,END)
            Admin_E_P_Type.config(state='disabled')
            Admin_E_P_Count.config(state='normal')
            Admin_E_P_Count.delete(0,END)
            Admin_E_P_Count.insert(0,0)
            Admin_E_P_Count.config(state='disabled')
            Admin_L_UpCount.place_forget()
            Get_data()
            messagebox.showinfo("Alert","Transaction Updated")
                

        Update_Admin = Button(Admin_w, text="Update",bg="#00728F",foreground="white",font=("Calibri", 11,"bold"),command=admin_Submit,cursor="hand2")
        Update_Admin.place(x=380+70,y=420 )

        #----------------------------Raw Report---------------------------
        #--Start Date--
        Admin_L_SDate=Label(Admin_w,text="Start Date",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_SDate.place(x=50,y=490-20)
        Admin_E_SDate=DateEntry(Admin_w, background= "black", foreground= "white",bd=2,font=('Calibri',11),date_pattern='dd-MM-yyyy')
        Admin_E_SDate.place(x=50,y=510-20,width=110, height=20+5)

        Admin_L_EDate=Label(Admin_w,text="End Date",fg="#515056",bg="white",bd=0, font=("Calibri", 10, "bold"))
        Admin_L_EDate.place(x=250,y=490-20)
        Admin_E_EDate=DateEntry(Admin_w, background= "black", foreground= "white",bd=2,font=('Calibri',11),date_pattern='dd-MM-yyyy')
        Admin_E_EDate.place(x=250,y=510-20,width=110, height=20+5)

        def Raw_Report():

            if Admin_C_Teamname.get() == '':
                messagebox.showerror("Team Name Missing", "Select Team Name")
                Admin_w.focus_force()
                return False
            
            if len(Admin_E_SDate.get())<1 or len(Admin_E_EDate.get())<1:
                messagebox.showerror("Data Missing", "Select Start and End Date")
                Admin_w.focus_force()
                return False
            
            if Admin_E_SDate.get() > Admin_E_EDate.get() :
                messagebox.showerror("Data Missing", "Incorrect Start and End Date")
                Admin_w.focus_force()
                return False
                
            ReportPath = fd.askdirectory(title='Select Folder Path')
            if ReportPath != '':
                F_Date=Admin_E_SDate.get()
                T_Date=Admin_E_EDate.get()
                
                TempPath = ReportPath + r"/Office Inn Raw Data Report " + datetime.now().strftime('%d-%m-%Y %H%M%S') + ".xlsx"
                
                query1 = "select * from TblProduction where   P_Date >= '"+ F_Date +"' and P_Date <= '"+ T_Date +"' and (USER_TEAM = '" + Admin_C_Teamname.get() +"' OR TRANSACTION_TEAM = '" + Admin_C_Teamname.get() +"')"
                query2 = "select * from TblNon_Production where   P_DATE >= '"+ F_Date +"' and P_DATE <= '"+ T_Date +"' and USER_TEAM = '" + Admin_C_Teamname.get() +"'"
                query3 = "select * from TblQCProducation where WIT_ProcessDate between '"+ F_Date +"' AND '"+ T_Date +"' or Uniq_ID in (SELECT Uniq_ID from TblQCProducation WHERE QC_Date >= '"+ F_Date +"' AND QC_Date <= '"+ T_Date +"')"

                
                Excelsheets = ['Production', 'Non_Production','QC']
                writer = pd.ExcelWriter(TempPath, engine = 'xlsxwriter')
                
                Production = pd.DataFrame()
                Non_Production = pd.DataFrame()
                QC = pd.DataFrame()

                for DB in DB_Name:
                    Data_Base=mypath+'\\'+DB
                    try:
                        conn = sqlite3.connect(Data_Base,timeout=45,uri=True)
                        
                        df1 = pd.read_sql_query(query1, conn)
                        Production = Production.append(df1)
                        
                        df2 = pd.read_sql_query(query2, conn)
                        Non_Production = Non_Production.append(df2)
                        
                        df3 = pd.read_sql_query(query3, conn)
                        QC = QC.append(df3)
                        conn.close()  
                    except:
                        conn.close()
                        messagebox.showerror("Data Crash","Office Inn crashed, Please reopen the tool")
                        return False
                
                QC=QC[QC.TRANSACTION_TEAM == str(Admin_C_Teamname.get())] 
                
                Dfs=[Production,Non_Production,QC]
                i=0
                for q in Dfs:
                    q.to_excel(writer, sheet_name=Excelsheets[i], index=False)
                    i=i+1
                conn.close()
                writer.save()
                
                messagebox.showinfo('Office Inn Raw Data','Office Inn Raw Data Report Exported!')
                startfile(TempPath)
            else:
                messagebox.showerror('Folder Validation','Folder Path Not Selected!')



        Raw_Report_Button = Button(Admin_w, text="Raw Report",bg="#00728F",foreground="white",font=("Calibri", 11,"bold"),command=Raw_Report,cursor="hand2" )
        Raw_Report_Button.place(x=380+70,y=510-30 )
##############################################################################################################################################################
    Admin_Button = Button(frame1,text="Admin",width=12,height=1,bg="#00728F",foreground="white", font=("Calibri", 11,"bold"),borderwidth=3,command= Admin_P,cursor="hand2" )
    Admin_Button.place(x=550+50,y=9)

    ClockIcon = PhotoImage(file = r"\\ltsbr\data\common\Equiniti India Ops\Innovation Projects\Office Inn\clock.png")
    Clock_Button = Button(frame1, image = ClockIcon,background='white',command= H_Producation,cursor="hand2" )
    Clock_Button.place(x=490,y=13)
    Officeinn.mainloop()
    
    
    
                               
    
    
CheckWindow = WindowCheck()
if "Office_Inn 1.4.1" in CheckWindow:
    messagebox.showerror('Already Opened','This application is already opened! Please check')
else:
    Main_Function()
