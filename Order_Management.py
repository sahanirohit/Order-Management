from tkinter import *
import pymysql
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import DateEntry

class Main:
    def __init__(self, root):
        self.root = root
        self.root.state("zoomed")
        self.root.title("Order Management System")
        self.root.minsize(1370, 720)   
        self.root.config(bg="#FDEFF4")
        # self.order_form()
        # self.allData()
        # self.table()
        # self.fetch_orders()
        

        # Login Frame
        self.username_var = StringVar()
        self.password_var = StringVar()
        self.username_var.set("admin")
        self.password_var.set("admin")
        self.login = Frame(self.root, bg="white",highlightbackground="black",highlightthickness=2)
        self.login.place(x=480, y=200, width=500, height=300)
        title = Label(self.login, text="Login Here",font=("times new roman",20,"bold"),bg="white",fg="#D82148").place(x=180,y=22)

        username = Label(self.login, text="Username",font=("times new roman",14),bg="white",fg="black").place(x=100,y=110)
        self.txt_username = Entry(self.login,font=("times new roman",14), textvariable=self.username_var, bg="#EEEEEE",bd=0)
        self.txt_username.place(x=200,y=110)

        password = Label(self.login, text="Password",font=("times new roman",14),bg="white",fg="black").place(x=100,y=150)
        self.txt_password = Entry(self.login,font=("times new roman",14), textvariable=self.password_var, bg="#EEEEEE",bd=0)
        self.txt_password.place(x=200,y=150)

        login_btn = Button(self.login,text="Login",bd=0,command=self.lg,font=("times new roman",14,"bold"),fg="#141E27",bg="#42C2FF").place(x=240,y=200,width=100, height=25)
        # self.login.destroy()
        # self.order.destroy()
        # self.data.destroy()

    # Order Form

    def order_form(self):

        # ======== All Variables ========= #
        self.date_var = StringVar()
        self.partyname_var = StringVar()
        self.orderno_var = StringVar()
        self.designno_var = StringVar()
        self.pick_var = StringVar()
        self.mtr_var = StringVar()
        self.wq_var = StringVar()
        self.panno_var = StringVar()
        self.status_var = StringVar()
        self.d1_var = StringVar()
        self.q1_var = StringVar()
        self.c1_var = StringVar()
        self.d2_var = StringVar()
        self.q2_var = StringVar()
        self.c2_var = StringVar()
        self.d3_var = StringVar()
        self.q3_var = StringVar()
        self.c3_var = StringVar()
        self.d4_var = StringVar()
        self.q4_var = StringVar()
        self.c4_var = StringVar()
        self.d5_var = StringVar()
        self.q5_var = StringVar()
        self.c5_var = StringVar()
        self.d6_var = StringVar()
        self.q6_var = StringVar()
        self.c6_var = StringVar()
        self.d7_var = StringVar()
        self.q7_var = StringVar()
        self.c7_var = StringVar()
        self.d8_var = StringVar()
        self.q8_var = StringVar()
        self.c8_var = StringVar()
        self.search_by_var = StringVar()
        self.search_txt_var = StringVar()
        self.txt_globe_var = StringVar()

        class CustomDateEntry(DateEntry):

            def _select(self, event=None):
                date = self._calendar.selection_get()
                if date is not None:
                    self._set_text(date.strftime('%d/%m/%Y'))
                    self.event_generate('<<DateEntrySelected>>')
                self._top_cal.withdraw()
                if 'readonly' not in self.state():
                    self.focus_set()

        # entry = CustomDateEntry(root)
        # entry._set_text(entry._date.strftime('%d/%m/%Y'))
        # entry.pack()
        
        self.order = Frame(self.root,bg="white", highlightbackground="black", highlightthickness=2)
        self.order.place(x=60,y=40,width=590,height=710)
        title = Label(self.order, text="Order Form",font=("times new roman",16,"bold"),bg="yellow").pack(fill=X)
            
        # Row 1
        date = Label(self.order,text="Date",font=("times new roman",12,"normal"),bg="white").place(x=20,y=40)
        self.txt_date = CustomDateEntry(self.order, width= 15, textvariable=self.date_var, bg="#EEEEEE",bd=0)
        self.txt_date._set_text(self.txt_date._date.strftime('%d/%m/%Y'))
        self.txt_date.place(x=20,y=70)

        partyname = Label(self.order,text="Party Name",font=("times new roman",12,"normal"),bg="white").place(x=160,y=40)
        self.txt_partyname = Entry(self.order, textvariable=self.partyname_var, width= 20, background= "#EEEEEE",bd=0)
        self.txt_partyname.place(x=160,y=70,height=20)

        orderno = Label(self.order,text="Order No.",font=("times new roman",12,"normal"),bg="white").place(x=300,y=40)
        self.txt_orderno = Entry(self.order, textvariable=self.orderno_var,width= 20, background= "#EEEEEE",bd=0)
        self.txt_orderno.place(x=300,y=70,height=20)

        designno = Label(self.order,text="Design No.",font=("times new roman",12,"normal"),bg="white").place(x=440,y=40)
        self.txt_designno = Entry(self.order, textvariable=self.designno_var,width= 20, background= "#EEEEEE",bd=0)
        self.txt_designno.place(x=440,y=70,height=20)

        # Row 2
        pick = Label(self.order,text="Pick",font=("times new roman",12,"normal"),bg="white").place(x=20,y=100)
        self.txt_pick = Entry(self.order, textvariable=self.pick_var,width= 20, background= "#EEEEEE",bd=0)
        self.txt_pick.place(x=20,y=130,height=20)

        meter = Label(self.order,text="Meter",font=("times new roman",12,"normal"),bg="white").place(x=160,y=100)
        self.txt_meter = Entry(self.order, textvariable=self.mtr_var, width= 20, background= "#EEEEEE",bd=0)
        self.txt_meter.place(x=160,y=130,height=20)

        wq = Label(self.order,text="Warp Quality",font=("times new roman",12,"normal"),bg="white").place(x=300,y=100)
        self.txt_wq = Entry(self.order, textvariable=self.wq_var, width= 20, background= "#EEEEEE",bd=0)
        self.txt_wq.place(x=300,y=130,height=20)

        panno = Label(self.order,text="Panno",font=("times new roman",12,"normal"),bg="white").place(x=440,y=100)
        self.txt_panno = Entry(self.order, textvariable=self.panno_var, width= 20, background= "#EEEEEE",bd=0)
            
        self.txt_panno.place(x=440,y=130,height=20)

        # Row 3
        status = Label(self.order,text="Status",font=("times new roman",12,"normal"),bg="white").place(x=20,y=160)
        self.txt_status = Entry(self.order, textvariable=self.status_var, width= 20, background= "#EEEEEE",bd=0)
        self.status_var.set("PENDING")
        self.txt_status.place(x=20,y=190,height=20)

        deniyar1 = Label(self.order,text="Deniyar 1",font=("times new roman",12,"normal"),bg="white").place(x=160,y=160)
        self.txt_deniyar1 = Entry(self.order, textvariable=self.d1_var, width= 20, background= "#EEEEEE",bd=0)
        self.txt_deniyar1.place(x=160,y=190,height=20)

        quality1 = Label(self.order,text="Quality 1",font=("times new roman",12,"normal"),bg="white").place(x=300,y=160)
        self.txt_quality1 = Entry(self.order, textvariable=self.q1_var, width= 20, background= "#EEEEEE",bd=0)
        self.txt_quality1.place(x=300,y=190,height=20)

        color1 = Label(self.order,text="Color 1",font=("times new roman",12,"normal"),bg="white").place(x=440,y=160)
        self.txt_color1 = Entry(self.order, textvariable=self.c1_var, width= 20, background= "#EEEEEE",bd=0)
        self.txt_color1.place(x=440,y=190,height=20)


        # Row 4
        deniyar2 = Label(self.order,text="Deniyar 2",font=("times new roman",12,"normal"),bg="white").place(x=20,y=220)
        self.txt_deniyar2 = Entry(self.order, textvariable=self.d2_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar2.place(x=20,y=250,height=20)

        quality2 = Label(self.order,text="Quality 2",font=("times new roman",12,"normal"),bg="white").place(x=210,y=220)
        self.txt_quality2 = Entry(self.order, textvariable=self.q2_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality2.place(x=210,y=250,height=20)

        color2 = Label(self.order,text="Color 2",font=("times new roman",12,"normal"),bg="white").place(x=400,y=220)
        self.txt_color2 = Entry(self.order, textvariable=self.c2_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color2.place(x=400,y=250,height=20)

        # Row 5
        deniyar3 = Label(self.order,text="Deniyar 3",font=("times new roman",12,"normal"),bg="white").place(x=20,y=280)
        self.txt_deniyar3 = Entry(self.order, textvariable=self.d3_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar3.place(x=20,y=310,height=20)

        quality3 = Label(self.order,text="Quality 3",font=("times new roman",12,"normal"),bg="white").place(x=210,y=280)
        self.txt_quality3 = Entry(self.order, textvariable=self.q3_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality3.place(x=210,y=310,height=20)

        color3 = Label(self.order,text="Color 3",font=("times new roman",12,"normal"),bg="white").place(x=400,y=280)
        self.txt_color3 = Entry(self.order, textvariable=self.c3_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color3.place(x=400,y=310,height=20)

        # Row 6
        deniyar4 = Label(self.order,text="Deniyar 4",font=("times new roman",12,"normal"),bg="white").place(x=20,y=340)
        self.txt_deniyar4 = Entry(self.order, textvariable=self.d4_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar4.place(x=20,y=370,height=20)

        quality4 = Label(self.order,text="Quality 4",font=("times new roman",12,"normal"),bg="white").place(x=210,y=340)
        self.txt_quality4 = Entry(self.order, textvariable=self.q4_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality4.place(x=210,y=370,height=20)

        color4 = Label(self.order,text="Color 4",font=("times new roman",12,"normal"),bg="white").place(x=400,y=340)
        self.txt_color4 = Entry(self.order, textvariable=self.c4_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color4.place(x=400,y=370,height=20)

        # Row 7
        deniyar5 = Label(self.order,text="Deniyar 5",font=("times new roman",12,"normal"),bg="white").place(x=20,y=400)
        self.txt_deniyar5 = Entry(self.order, textvariable=self.d5_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar5.place(x=20,y=430,height=20)

        quality5 = Label(self.order,text="Quality 5",font=("times new roman",12,"normal"),bg="white").place(x=210,y=400)
        self.txt_quality5 = Entry(self.order, textvariable=self.q5_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality5.place(x=210,y=430,height=20)

        color5 = Label(self.order,text="Color 5",font=("times new roman",12,"normal"),bg="white").place(x=400,y=400)
        self.txt_color5 = Entry(self.order, textvariable=self.c5_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color5.place(x=400,y=430,height=20)

         # Row 8
        deniyar6 = Label(self.order,text="Deniyar 6",font=("times new roman",12,"normal"),bg="white").place(x=20,y=460)
        self.txt_deniyar6 = Entry(self.order, textvariable=self.d6_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar6.place(x=20,y=490,height=20)

        quality6 = Label(self.order,text="Quality 6",font=("times new roman",12,"normal"),bg="white").place(x=210,y=460)
        self.txt_quality6 = Entry(self.order, textvariable=self.q6_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality6.place(x=210,y=490,height=20)

        color6 = Label(self.order,text="Color 6",font=("times new roman",12,"normal"),bg="white").place(x=400,y=460)
        self.txt_color6 = Entry(self.order, textvariable=self.c6_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color6.place(x=400,y=490,height=20)

        # Row 9
        deniyar7 = Label(self.order,text="Deniyar 7",font=("times new roman",12,"normal"),bg="white").place(x=20,y=520)
        self.txt_deniyar7 = Entry(self.order, textvariable=self.d7_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar7.place(x=20,y=550,height=20)

        quality7 = Label(self.order,text="Quality 7",font=("times new roman",12,"normal"),bg="white").place(x=210,y=520)
        self.txt_quality7 = Entry(self.order, textvariable=self.q7_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality7.place(x=210,y=550,height=20)

        color7 = Label(self.order,text="Color 7",font=("times new roman",12,"normal"),bg="white").place(x=400,y=520)
        self.txt_color7 = Entry(self.order, textvariable=self.c7_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color7.place(x=400,y=550,height=20)

        # Row 10
        deniyar8 = Label(self.order,text="Deniyar 8",font=("times new roman",12,"normal"),bg="white").place(x=20,y=580)
        self.txt_deniyar8 = Entry(self.order, textvariable=self.d8_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_deniyar8.place(x=20,y=610,height=20)

        quality8 = Label(self.order,text="Quality 8",font=("times new roman",12,"normal"),bg="white").place(x=210,y=580)
        self.txt_quality8 = Entry(self.order, textvariable=self.q8_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_quality8.place(x=210,y=610,height=20)

        color8 = Label(self.order,text="Color 8",font=("times new roman",12,"normal"),bg="white").place(x=400,y=580)
        self.txt_color8 = Entry(self.order, textvariable=self.c8_var, width= 27, background= "#EEEEEE",bd=0)
        self.txt_color8.place(x=400,y=610,height=20)

        # Row 11
        btn_insert = Button(self.order,text="INSERT", command=self.insert_order, font=("times new roman",8,"normal"),bd=0).place(x=20,y=660,width=80, height=30)
            
        btn_update = Button(self.order,text="UPDATE", command=self.update_order, font=("times new roman",8,"normal"),bd=0).place(x=135,y=660,width=80, height=30)
            
        btn_delete = Button(self.order,text="DELETE", command=self.delete_orders, font=("times new roman",8,"normal"),bd=0).place(x=245,y=660,width=80, height=30)
            
        btn_print = Button(self.order,text="PRINT",font=("times new roman",8,"normal"),bd=0).place(x=360,y=660,width=80, height=30)

        btn_clear = Button(self.order,text="CLEAR", command=self.clear_orders, font=("times new roman",8,"normal"),bd=0).place(x=485,y=660,width=80, height=30)
    
    # Completed Orders Data
    def completed_orders(self):

        self.order.destroy()
        self.data.destroy()
        # self.table.destroy()
        self.completed_frame = Frame(root, bg="white", width=1430, height=700, highlightbackground="black", highlightthickness=2)
        self.completed_frame.place(x=50, y=40)
        title = Label(self.completed_frame,text="All Completed Orders",font=("times new roman",14,"bold"),bg="white",fg="red").place(x=600, y=5)


        # completed Table Frame
        self.completed_table = Frame(self.completed_frame, bg="white",highlightbackground="black",highlightthickness=1)
        self.completed_table.place(x=8,y=40, width=1410, height=650)

        scroll_x = Scrollbar(self.completed_table,orient=HORIZONTAL)
        scroll_y = Scrollbar(self.completed_table,orient=VERTICAL)
        self.completed_order_table = ttk.Treeview(self.completed_table,columns=("date","partyname","status","orderno","designno","pick","mtr","wq","panno","d1","q1","c1","d2","q2","c2","d3","q3","c3","d4","q4","c4","d5","q5","c5","d6","q6","c6","d7","q7","c7","d8","q8","c8"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)
        scroll_x.config(command=self.completed_order_table.xview)
        scroll_y.config(command=self.completed_order_table.yview)

        self.completed_order_table.heading("date",text="Date")
        self.completed_order_table.column("date",width=70,anchor=CENTER)
        self.completed_order_table.heading("partyname",text="Party Name")
        self.completed_order_table.column("partyname",width=90,anchor=CENTER)
        self.completed_order_table.heading("orderno",text="Order No.")
        self.completed_order_table.column("status",width=80,anchor=CENTER)
        self.completed_order_table.heading("status",text="Status")
        self.completed_order_table.column("orderno",width=70,anchor=CENTER)
        self.completed_order_table.heading("designno",text="Design No.")
        self.completed_order_table.column("designno",width=80,anchor=CENTER)
        self.completed_order_table.heading("pick",text="Pick")
        self.completed_order_table.column("pick",width=50,anchor=CENTER)
        self.completed_order_table.heading("mtr",text="Meter")
        self.completed_order_table.column("mtr",width=40,anchor=CENTER)
        self.completed_order_table.heading("wq",text="Warp Quality")
        self.completed_order_table.column("wq",width=100,anchor=CENTER)
        self.completed_order_table.heading("panno",text="Panno")
        self.completed_order_table.column("panno",width=50,anchor=CENTER)
        self.completed_order_table.heading("d1",text="Den 1")
        self.completed_order_table.column("d1",width=70,anchor=CENTER)
        self.completed_order_table.heading("q1",text="Qual. 1")
        self.completed_order_table.column("q1",width=70,anchor=CENTER)
        self.completed_order_table.heading("c1",text="Color 1")
        self.completed_order_table.column("c1",width=70,anchor=CENTER)
        self.completed_order_table.heading("d2",text="Den 2")
        self.completed_order_table.column("d2",width=70,anchor=CENTER)
        self.completed_order_table.heading("q2",text="Qual. 2")
        self.completed_order_table.column("q2",width=70,anchor=CENTER)
        self.completed_order_table.heading("c2",text="Color 2")
        self.completed_order_table.column("c2",width=70,anchor=CENTER)
        self.completed_order_table.heading("d3",text="Den 3")
        self.completed_order_table.column("d3",width=70,anchor=CENTER)
        self.completed_order_table.heading("q3",text="Qual. 3")
        self.completed_order_table.column("q3",width=70,anchor=CENTER)
        self.completed_order_table.heading("c3",text="Color 3")
        self.completed_order_table.column("c3",width=70,anchor=CENTER)
        self.completed_order_table.heading("d4",text="Den 4")
        self.completed_order_table.column("d4",width=70,anchor=CENTER)
        self.completed_order_table.heading("q4",text="Qual. 4")
        self.completed_order_table.column("q4",width=70,anchor=CENTER)
        self.completed_order_table.heading("c4",text="Color 4")
        self.completed_order_table.column("c4",width=70,anchor=CENTER)
        self.completed_order_table.heading("d5",text="Den 5")
        self.completed_order_table.column("d5",width=70,anchor=CENTER)
        self.completed_order_table.heading("q5",text="Qual. 5")
        self.completed_order_table.column("q5",width=70,anchor=CENTER)
        self.completed_order_table.heading("c5",text="Color 5")
        self.completed_order_table.column("c5",width=70,anchor=CENTER)
        self.completed_order_table.heading("d6",text="Den 6")
        self.completed_order_table.column("d6",width=70,anchor=CENTER)
        self.completed_order_table.heading("q6",text="Qual. 6")
        self.completed_order_table.column("q6",width=70,anchor=CENTER)
        self.completed_order_table.heading("c6",text="Color 6")
        self.completed_order_table.column("c6",width=70,anchor=CENTER)
        self.completed_order_table.heading("d7",text="Den 7")
        self.completed_order_table.column("d7",width=70,anchor=CENTER)
        self.completed_order_table.heading("q7",text="Qual. 7")
        self.completed_order_table.column("q7",width=70,anchor=CENTER)
        self.completed_order_table.heading("c7",text="Color 7")
        self.completed_order_table.column("c7",width=70,anchor=CENTER)
        self.completed_order_table.heading("d8",text="Den 8")
        self.completed_order_table.column("d8",width=70,anchor=CENTER)
        self.completed_order_table.heading("q8",text="Qual. 8")
        self.completed_order_table.column("q8",width=70,anchor=CENTER)
        self.completed_order_table.heading("c8",text="Color 8")
        self.completed_order_table.column("c8",width=70,anchor=CENTER)
        self.completed_order_table["show"]="headings"
        self.completed_order_table.pack(fill=BOTH,expand=1)

        con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
        cur = con.cursor()
        cur.execute("select * from orders where status LIKE '%COMPLETED%'")
        rows = cur.fetchall()
        if len(rows)!=0:
            self.completed_order_table.delete(*self.completed_order_table.get_children())
            for row in rows:
                self.completed_order_table.insert('',END,values=row)
            con.commit()
        con.close()

    # Menubar
    def menu(self):
        mymenu = Menu(self.root)
        mymenu.add_command(label="Home", command=self.home)
        mymenu.add_command(label="Pending Orders")
        mymenu.add_command(label="Completed Orders", command=self.completed_orders)
        self.root.config(menu=mymenu)

    # Data Frame
    def allData(self):
        self.data = Frame(root,bg="white", highlightbackground="black", highlightthickness=2)
        self.data.place(x=700, y=40, width=770, height=710)

        lbl_search = Label(self.data,text="Search By", bg="white",font=("times new roma",10,"normal")).place(x=20, y=20)

        combo_search = ttk.Combobox(self.data, textvariable=self.search_by_var, width=15, font=("times new roman",10,"normal"),state="readonly")
        combo_search['values']=("Select","PartyName","DesignNo","OrderNo","status")
        combo_search.current(0)
        combo_search.place(x=130,y=20)

        txt_search = Entry(self.data, textvariable=self.search_txt_var, width=27, bg="#EEEEEE",bd=1).place(x=290,y=20,height=20)
        btn_search = Button(self.data, text="Search", command=self.search_orders, bd=0).place(x=500,y=20,height=20,width=80)
        btn_showall = Button(self.data, text="Show All", command=self.fetch_orders, bd=0).place(x=620,y=20,height=20,width=80)
    
    # Table Frame
    def tableFrame(self):
        self.table = Frame(self.data, bg="white",highlightbackground="black",highlightthickness=1)
        self.table.place(x=8,y=60, width=750, height=640)

        scroll_x = Scrollbar(self.table,orient=HORIZONTAL)
        scroll_y = Scrollbar(self.table,orient=VERTICAL)
        self.order_table = ttk.Treeview(self.table,columns=("date","partyname","status","orderno","designno","pick","mtr","wq","panno","d1","q1","c1","d2","q2","c2","d3","q3","c3","d4","q4","c4","d5","q5","c5","d6","q6","c6","d7","q7","c7","d8","q8","c8"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)
        scroll_x.config(command=self.order_table.xview)
        scroll_y.config(command=self.order_table.yview)

        self.order_table.heading("date",text="Date")
        self.order_table.column("date",width=70,anchor=CENTER)
        self.order_table.heading("partyname",text="Party Name")
        self.order_table.column("partyname",width=90,anchor=CENTER)
        self.order_table.heading("orderno",text="Order No.")
        self.order_table.column("status",width=80,anchor=CENTER)
        self.order_table.heading("status",text="Status")
        self.order_table.column("orderno",width=70,anchor=CENTER)
        self.order_table.heading("designno",text="Design No.")
        self.order_table.column("designno",width=80,anchor=CENTER)
        self.order_table.heading("pick",text="Pick")
        self.order_table.column("pick",width=50,anchor=CENTER)
        self.order_table.heading("mtr",text="Meter")
        self.order_table.column("mtr",width=40,anchor=CENTER)
        self.order_table.heading("wq",text="Warp Quality")
        self.order_table.column("wq",width=100,anchor=CENTER)
        self.order_table.heading("panno",text="Panno")
        self.order_table.column("panno",width=50,anchor=CENTER)
        self.order_table.heading("d1",text="Den 1")
        self.order_table.column("d1",width=70,anchor=CENTER)
        self.order_table.heading("q1",text="Qual. 1")
        self.order_table.column("q1",width=70,anchor=CENTER)
        self.order_table.heading("c1",text="Color 1")
        self.order_table.column("c1",width=70,anchor=CENTER)
        self.order_table.heading("d2",text="Den 2")
        self.order_table.column("d2",width=70,anchor=CENTER)
        self.order_table.heading("q2",text="Qual. 2")
        self.order_table.column("q2",width=70,anchor=CENTER)
        self.order_table.heading("c2",text="Color 2")
        self.order_table.column("c2",width=70,anchor=CENTER)
        self.order_table.heading("d3",text="Den 3")
        self.order_table.column("d3",width=70,anchor=CENTER)
        self.order_table.heading("q3",text="Qual. 3")
        self.order_table.column("q3",width=70,anchor=CENTER)
        self.order_table.heading("c3",text="Color 3")
        self.order_table.column("c3",width=70,anchor=CENTER)
        self.order_table.heading("d4",text="Den 4")
        self.order_table.column("d4",width=70,anchor=CENTER)
        self.order_table.heading("q4",text="Qual. 4")
        self.order_table.column("q4",width=70,anchor=CENTER)
        self.order_table.heading("c4",text="Color 4")
        self.order_table.column("c4",width=70,anchor=CENTER)
        self.order_table.heading("d5",text="Den 5")
        self.order_table.column("d5",width=70,anchor=CENTER)
        self.order_table.heading("q5",text="Qual. 5")
        self.order_table.column("q5",width=70,anchor=CENTER)
        self.order_table.heading("c5",text="Color 5")
        self.order_table.column("c5",width=70,anchor=CENTER)
        self.order_table.heading("d6",text="Den 6")
        self.order_table.column("d6",width=70,anchor=CENTER)
        self.order_table.heading("q6",text="Qual. 6")
        self.order_table.column("q6",width=70,anchor=CENTER)
        self.order_table.heading("c6",text="Color 6")
        self.order_table.column("c6",width=70,anchor=CENTER)
        self.order_table.heading("d7",text="Den 7")
        self.order_table.column("d7",width=70,anchor=CENTER)
        self.order_table.heading("q7",text="Qual. 7")
        self.order_table.column("q7",width=70,anchor=CENTER)
        self.order_table.heading("c7",text="Color 7")
        self.order_table.column("c7",width=70,anchor=CENTER)
        self.order_table.heading("d8",text="Den 8")
        self.order_table.column("d8",width=70,anchor=CENTER)
        self.order_table.heading("q8",text="Qual. 8")
        self.order_table.column("q8",width=70,anchor=CENTER)
        self.order_table.heading("c8",text="Color 8")
        self.order_table.column("c8",width=70,anchor=CENTER)
        self.order_table["show"]="headings"
        self.order_table.pack(fill=BOTH,expand=1)
        self.order_table.bind("<ButtonRelease-1>",self.get_cursor)
        

    # Functions

    def home(self):
        self.completed_frame.destroy()
        self.order_form()
        self.allData()
        self.tableFrame()
        self.fetch_orders()


    def lg(self):
            if self.txt_username.get() == "" or self.txt_password.get() == "":
                messagebox.showerror("Error","All fields are required !!")
            else:
                try:
                    con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
                    cur = con.cursor()
                    cur.execute("select * from user where username=%s and password=%s",(self.txt_username.get(),self.txt_password.get()))
                    row=cur.fetchone()
                    if row==None:
                        messagebox.showerror("Error","Invalid Username & Password")
                    else:
                        messagebox.showinfo("Success","Login Successfully")
                        con.close()
                        self.login.destroy()
                        self.order_form()
                        self.allData()
                        self.tableFrame()
                        self.fetch_orders()
                        self.menu()
                except Exception as es:
                    messagebox.showerror("Error",f"Error due to: {str(es)}",parent=self.root)

    # All MySQL Queries

    # Fetch Orders Data
    def fetch_orders(self):
        try:
            con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
            cur = con.cursor()
            cur.execute("select * from orders")
            rows = cur.fetchall()
            if len(rows)!=0:
                self.order_table.delete(*self.order_table.get_children())
                for row in rows:
                    self.order_table.insert('',END,values=row)
                con.commit()
            con.close()
        except Exception as e:
            messagebox.showwarning("No Data","You have no Data to show")
            
    

    # Update Data to Database
    def update_order(self):
        
        con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
        cur = con.cursor()
        cur.execute("update orders set date=%s,partyname=%s,status=%s, designno=%s,pick=%s,mtr=%s,wq=%s,panno=%s,d1=%s,q1=%s,c1=%s,d2=%s,q2=%s,c2=%s,d3=%s,q3=%s,c3=%s,d4=%s,q4=%s,c4=%s,d5=%s,q5=%s,c5=%s,d6=%s,q6=%s,c6=%s,d7=%s,q7=%s,c7=%s,d8=%s,q8=%s,c8=%s where orderno=%s",(
            self.txt_date.get(),
            self.txt_partyname.get(),
            self.txt_status.get(),
            self.txt_designno.get(),
            self.txt_pick.get(),
            self.txt_meter.get(),
            self.txt_wq.get(),
            self.txt_panno.get(),
            self.txt_deniyar1.get(),
            self.txt_quality1.get(),
            self.txt_color1.get(),
            self.txt_deniyar2.get(),
            self.txt_quality2.get(),
            self.txt_color2.get(),
            self.txt_deniyar3.get(),
            self.txt_quality3.get(),
            self.txt_color3.get(),
            self.txt_deniyar4.get(),
            self.txt_quality4.get(),
            self.txt_color4.get(),
            self.txt_deniyar5.get(),
            self.txt_quality5.get(),
            self.txt_color5.get(),
            self.txt_deniyar6.get(),
            self.txt_quality6.get(),
            self.txt_color6.get(),
            self.txt_deniyar7.get(),
            self.txt_quality7.get(),
            self.txt_color7.get(),
            self.txt_deniyar8.get(),
            self.txt_quality8.get(),
            self.txt_color8.get(),
            self.txt_orderno.get()
            ))
        con.commit()
        messagebox.showinfo("Success","Your order Updated Successfully")
        self.fetch_orders()
        con.close()

    # Delete Data from Database
    def delete_orders(self):
        con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
        cur = con.cursor()
        cur.execute("delete from orders where orderno=%s",self.orderno_var.get())
        con.commit()
        con.close()
        messagebox.showinfo("Success","Your Order Deleted Successfully")
        self.fetch_orders()

    # Insert Data to Database
    def insert_order(self):
        if self.txt_date.get() == "" or self.txt_partyname.get() == "" or self.txt_status.get() == "" or self.txt_orderno.get() == "":
            messagebox.showerror("Error","Please fill all required details...")
        else:
            try:
                con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
                cur = con.cursor()
                cur.execute("select * from orders where orderno=%s ",(self.txt_orderno.get()))
                row=cur.fetchone()
                if row==None:
                    con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
                    cur = con.cursor()
                    cur.execute("insert into orders values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(
                        self.txt_date.get(),
                        self.txt_partyname.get(),
                        self.txt_status.get(),
                        self.txt_orderno.get(),
                        self.txt_designno.get(),
                        self.txt_pick.get(),
                        self.txt_meter.get(),
                        self.txt_wq.get(),
                        self.txt_panno.get(),
                        self.txt_deniyar1.get(),
                        self.txt_quality1.get(),
                        self.txt_color1.get(),
                        self.txt_deniyar2.get(),
                        self.txt_quality2.get(),
                        self.txt_color2.get(),
                        self.txt_deniyar3.get(),
                        self.txt_quality3.get(),
                        self.txt_color3.get(),
                        self.txt_deniyar4.get(),
                        self.txt_quality4.get(),
                        self.txt_color4.get(),
                        self.txt_deniyar5.get(),
                        self.txt_quality5.get(),
                        self.txt_color5.get(),
                        self.txt_deniyar6.get(),
                        self.txt_quality6.get(),
                        self.txt_color6.get(),
                        self.txt_deniyar7.get(),
                        self.txt_quality7.get(),
                        self.txt_color7.get(),
                        self.txt_deniyar8.get(),
                        self.txt_quality8.get(),
                        self.txt_color8.get()
                            
                    ))
                    con.commit()
                    messagebox.showinfo("Success","Your order Inserted Successfully")
                    self.fetch_orders()
                    con.close()
                else:
                    messagebox.showwarning("Warning","This Order No already exists.")
                    con.close()
            except Exception as es:
                messagebox.showerror("Error",f"Error due to: {str(es)}",parent=self.root)
            
    def search_orders(self):
        # print(self.search_by.get())
        con=pymysql.connect(host="localhost",user="root",password="",database="manage_orders")
        cur = con.cursor()
        cur.execute("select * from orders where "+str(self.search_by_var.get())+" LIKE '%"+str(self.search_txt_var.get())+"%'")
        rows = cur.fetchall()
        if len(rows)!=0:
            self.order_table.delete(*self.order_table.get_children())
            for row in rows:
                self.order_table.insert('',END,values=row)
            con.commit()
        con.close()

    # Clear All Fields
    def clear_orders(self):        
        self.date_var.set("")
        self.partyname_var.set("")
        self.orderno_var.set("")
        self.designno_var.set("")
        self.pick_var.set("")
        self.mtr_var.set("")
        self.wq_var.set("")
        self.panno_var.set("")
        self.d1_var.set("")
        self.q1_var.set("")
        self.c1_var.set("")
        self.d2_var.set("")
        self.q2_var.set("")
        self.c2_var.set("")
        self.d3_var.set("")
        self.q3_var.set("")
        self.c3_var.set("")
        self.d4_var.set("")
        self.q4_var.set("")
        self.c4_var.set("")
        self.d5_var.set("")
        self.q5_var.set("")
        self.c5_var.set("")
        self.d6_var.set("")
        self.q6_var.set("")
        self.c6_var.set("")
        self.d7_var.set("")
        self.q7_var.set("")
        self.c7_var.set("")
        self.d8_var.set("")
        self.q8_var.set("")
        self.c8_var.set("")

    # Edit Data
    def get_cursor(self,event):
        cursor_row = self.order_table.focus()
        contents = self.order_table.item(cursor_row)
        row = contents['values']
                   
        self.date_var.set(row[0])
        self.partyname_var.set(row[1])
        self.status_var.set(row[2])
        self.orderno_var.set(row[3])
        self.designno_var.set(row[4])
        self.pick_var.set(row[5])
        self.mtr_var.set(row[6])
        self.wq_var.set(row[7])
        self.panno_var.set(row[8])
        self.d1_var.set(row[9])
        self.q1_var.set(row[10])
        self.c1_var.set(row[11])
        self.d2_var.set(row[12])
        self.q2_var.set(row[13])
        self.c2_var.set(row[14])
        self.d3_var.set(row[15])
        self.q3_var.set(row[16])
        self.c3_var.set(row[17])
        self.d4_var.set(row[18])
        self.q4_var.set(row[19])
        self.c4_var.set(row[20])
        self.d5_var.set(row[21])
        self.q5_var.set(row[22])
        self.c5_var.set(row[23])
        self.d6_var.set(row[24])
        self.q6_var.set(row[25])
        self.c6_var.set(row[26])
        self.d7_var.set(row[27])
        self.q7_var.set(row[28])
        self.c7_var.set(row[29])
        self.d8_var.set(row[30])
        self.q8_var.set(row[31])
        self.c8_var.set(row[32])

root = Tk()
obj = Main(root)
root.mainloop()