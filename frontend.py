import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from pymongo import MongoClient
from tkinter import simpledialog
# from tkcalendar import DateEntry
import datetime
import pandas as pd
from tkinter import filedialog
import os
from bson import ObjectId
import ttkbootstrap as ttk
from ttkbootstrap import Style


client = MongoClient("mongodb://localhost:27017/")

class MonthYearEntry(ttk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # Create a Combobox for selecting the year
        self.year_var = tk.StringVar()
        self.year_var.trace_add('write', self._on_select_year)
        self.year_combo = ttk.Combobox(self, textvariable=self.year_var, width=6, 
                                       values=[str(y) for y in range(2010, 2101)])
        self.year_combo.set('2022')
        self.year_combo.pack(side='right', padx=(0, 5))
        
        # Create a Combobox for selecting the month
        self.month_var = tk.StringVar()
        self.month_var.trace_add('write', self._on_select_month)
        self.month_combo = ttk.Combobox(self, textvariable=self.month_var, width=10, 
                                        values=('January', 'February', 'March', 'April', 
                                                'May', 'June', 'July', 'August', 
                                                'September', 'October', 'November', 'December'))
        self.month_combo.set('January')
        self.month_combo.pack(side='left', padx=(5, 0))

    def _on_select_month(self, *args):
        self.event_generate('<<MonthYearChanged>>')

    def _on_select_year(self, *args):
        self.event_generate('<<MonthYearChanged>>')

    def get(self):
        month = self.month_combo.current() + 1
        year = int(self.year_combo.get())
        return f'{year}-{month:02}'

    def delete(self):
        self.month_combo.set('')
        self.year_combo.set('')
    
    def insert(self, date):
        year = date[:4]
        month = int(date[5:7])
        values=['January', 'February', 'March', 'April', 
                'May', 'June', 'July', 'August', 
                'September', 'October', 'November', 'December']
        self.month_combo.set(values[month-1])
        self.year_combo.set(year)

class App:
    def __init__(self, master):
        super().__init__()
        # Main Setup
        self.master = master
        self.style = Style()
        self.master.title("Shri Shiddivinayaka Industries")
        self.master.geometry('1000x600')
        self.master.minsize(1000,600)
        os.popen("mongod")
        
        # Widgets
        if hasattr(self, 'sidebar'):
            self.master.sidebar.destroy()
        self.master.sidebar = Sidebar(self.master)

class Sidebar(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.frame = tk.Frame(self)
        self.frame.pack(expand=True, fill="both")
        self.place(x=0, y=0, relheight=1, relwidth=.2)
        self.create_w(self.frame)
        if hasattr(self, 'main'):
            self.main.destroy()
        self.main = Main(parent)    

    def create_w(self, frame):
        self.currentdb=  ttk.Label(frame, text="Welcome", font=('Calibri',25, 'bold'))
        refresh = ttk.Button(frame,text='Refresh',bootstyle='warning', command=self.add_data)
        delete = ttk.Button(frame, text="Delete",bootstyle='danger', command=self.getval)
        add = ttk.Button(frame, text="Add", command=self.newdb)
        search_box = tk.LabelFrame(frame)
        self.search_entry = ttk.Entry(search_box)
        search_button = ttk.Button(search_box, text="Search", bootstyle='success', command=self.search)

        self.dblist = tk.Listbox(frame, height=10)
        self.dblist.bind_all("<Escape>", self.clear_sel)
        self.dblist.bind("<Double-Button-1>", self.on_double)
        self.add_data()
        

        frame.columnconfigure((0,1,2), weight=1, uniform='a')
        #frame.rowconfigure((0,1,2,3,4), weight=1, uniform='a')

        self.bind('<Configure>', self.update_s)
        search_box.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.search_entry.pack(side='left')
        search_button.pack(side='right')
        self.currentdb.grid(row=0, column=0, sticky='w', padx=5, pady=5, columnspan=3)
        refresh.grid(row=1, column=0,sticky='nsew', pady=5, padx=5)
        delete.grid(row=1, column=1,sticky='nsew', pady=5, padx=5)
        add.grid(row=1, column=2, sticky='nsew',pady=5, padx=5)
        self.dblist.grid(row=3, column=0, columnspan=3, sticky='nsew', padx=5, pady=5)
    


    def currdb(self):
        self.main.draw(self.currentdb.cget('text'))

    def on_double(self,event):
        i = self.dblist.curselection()
        curdb = self.dblist.get(i)
        self.currentdb.config(text=curdb)
        self.currdb()

    def update_s(self, event):
        x = self.winfo_width()
        if x==200:
            self.search_entry.configure(width=19)
        else:
            self.search_entry.configure(width=int(17+x/19))
    
    def add_data(self):
        self.dblist.delete(0, 'end')
        database_names = client.list_database_names()
        for name in database_names:
            if name not in("admin", "local", "config"):
                self.dblist.insert(tk.END, name)

    def getval(self):
        try:
            index = self.dblist.curselection()
            condition = messagebox.askyesno('Warning',f"Are you sure you want to delete {self.dblist.get(index)}")
            if condition is True:
                con = messagebox.askyesno("Sure", 'Sure?')
                if con is True:
                    dbname = f'{self.dblist.get(index)}'
                    db = client[dbname]
                    client.drop_database(dbname)
                    messagebox.showinfo("Deleted",f"{self.dblist.get(index)} has been deleted")
                    self.add_data()
        except Exception as e:
            messagebox.showerror("Error", "Nothing is selected")

    def newdb(self):
        name = simpledialog.askstring("Name", "Name of the Database:")
        db = client[name]
        db.create_collection("RCNutPurchase", validator={
        '$jsonSchema': {'bsonType': 'object','required': 
                        ['BillNo','Date','Rate','Bags','Quintal','KGs','SGST','CGST','IGST'],
        'properties': {
            'BillNo': {'bsonType': 'int'},
            'Date': {'bsonType': 'date'},
            'Rate': {'bsonType': ['double','int']},
            'Bags': {'bsonType': ['int','double']},
            'Quintal': {'bsonType': ['double','int']},
            'KGs': {'bsonType': ['double','int']},
            'SGST': {'bsonType': ['double','int']},
            'CGST': {'bsonType': ['double','int']},
            'IGST': {'bsonType': ['double','int']}
        }
        }
        })
        col = db["RCNutPurchase"]
        col.create_index([('BillNo',1)], unique=True)
        db.create_collection("CKDespatch", validator={
            '$jsonSchema': {
                'bsonType': 'object',
                'required': [
                    'BillNo',
                    'Date',
                    'Qty',
                    'WRate',
                    'Wholes',
                    'PRate',
                    'Pieces',
                    'CGST',
                    'SGST'
                ],
                'properties': {
                    'BillNo': {
                        'bsonType': 'int',
                        'description': 'Unique bill number'
                    },
                    'Date': {
                        'bsonType': 'date',
                        'description': 'Date of despatch'
                    },
                    'Qty': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'Quantity of product despatched'
                    },
                    'WRate': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'Weighted rate of product despatched'
                    },
                    'Wholes': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'Weight of wholes despatched'
                    },
                    'PRate': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'Price rate of product despatched'
                    },
                    'Pieces': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'Number of pieces despatched'
                    },
                    'CGST': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'CGST charged'
                    },
                    'SGST': {
                        'bsonType': [
                            'int',
                            'double'
                        ],
                        'description': 'SGST charged'
                    }
                }
            }
        })
        col = db["CKDespatch"]
        col.create_index([('BillNo',1)], unique=True)

        db.create_collection('HDR', validator={
            "$jsonSchema": {
                "bsonType": "object",
                "required": [
                    "Sl",
                    "Date",
                    "HRate",
                    "Husk",
                    "HuskAmt",
                    "DRate",
                    "Dust",
                    "DustAmt",
                    "RRate",
                    "Rejection",
                    "RejectionAmt",
                    "GRate",
                    "Gbags",
                    "GbagAmt"
                ],
                "properties": {
                    "Sl": {
                        "bsonType": "int"
                    },
                    "Date": {
                        "bsonType": "date"
                    },
                    "HRate": {
                        "bsonType": ["int", "double"]
                    },
                    "Husk": {
                        "bsonType": ["int", "double"]
                    },
                    "HuskAmt": {
                        "bsonType": ["int", "double"]
                    },
                    "DRate": {
                        "bsonType": ["int", "double"]
                    },
                    "Dust": {
                        "bsonType": ["int", "double"]
                    },
                    "DustAmt": {
                        "bsonType": ["int", "double"]
                    },
                    "RRate": {
                        "bsonType": ["int", "double"]
                    },
                    "Rejection": {
                        "bsonType": ["int", "double"]
                    },
                    "RejectionAmt": {
                        "bsonType": ["int", "double"]
                    },
                    "GRate": {
                        "bsonType": ["int", "double"]
                    },
                    "Gbags": {
                        "bsonType": ["int", "double"]
                    },
                    "GbagAmt": {
                        "bsonType": ["int", "double"]
                    }
                }
            }
        }
    )
        col = db["HDR"]
        col.create_index([('Sl',1)], unique=True)
        db.create_collection('JEChart', validator={
            "$jsonSchema": {
                "bsonType": "object",
                "required": [
                "Date",
                "Sales",
                "Payment",
                "TCS_TDS",
                "Purchase",
                "Receipt"
                ],
                "properties": {
                "Date": {
                    "bsonType": "date"
                },
                "Sales": {
                    "bsonType": [
                    "double",
                    "int"
                    ],
                    "description": "must be a double or int"
                },
                "Payment": {
                    "bsonType": [
                    "double",
                    "int"
                    ],
                    "description": "must be a double or int"
                },
                "TCS_TDS": {
                    "bsonType": [
                    "double",
                    "int"
                    ],
                    "description": "must be a double or int"
                },
                "Purchase": {
                    "bsonType": [
                    "double",
                    "int"
                    ],
                    "description": "must be a double or int"
                },
                "Receipt": {
                    "bsonType": [
                    "double",
                    "int"
                    ],
                    "description": "must be a double or int"
                }
                }
            }
            }
            )
        col = db["JEChart"]
        col.create_index([('Date',1)], unique=True)
        db.create_collection('KernelList', validator={
            "$jsonSchema": {
                "bsonType": "object",
                "required": [
                "LotNo",
                "TotalWholes",
                "TotalPieces",
                "DispWholes",
                "DispPieces",
                "DispRejection",
                "Party",
                "PartyWholes",
                "PartyPieces"
                ],
                "properties": {
                "LotNo": {
                    "bsonType": "int",
                    "description": "must be an integer and is required"
                },
                "TotalWholes": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                },
                "TotalPieces": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                },
                "DispWholes": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                },
                "DispPieces": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                },
                "DispRejection": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                },
                "Party": {
                    "bsonType": [
                    "string"
                    ],
                    "description": "must be a string and is required"
                },
                "PartyWholes": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                },
                "PartyPieces": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "must be a number and is required"
                }
                }
            }
            }
            )
        col = db["KernelList"]
        col.create_index([('LotNo',1)], unique=True)
        db.create_collection('SDespatch',validator={
            "$jsonSchema": {
                "bsonType": "object",
                "required": [
                "BillNo",
                "Date",
                "Rate",
                "Bags",
                "KGs",
                "Amount",
                "CGST",
                "SGST",
                "ActualQty"
                ],
                "properties": {
                "BillNo": {
                    "bsonType": "int",
                    "description": "unique bill number for the despatch"
                },
                "Date": {
                    "bsonType": "date",
                    "description": "date of the despatch"
                },
                "Rate": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "rate per KG"
                },
                "Bags": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "number of bags"
                },
                "KGs": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "total weight in KGs"
                },
                "Amount": {
                    "bsonType": [
                    "int",
                    "double"
                    ]
                },
                "CGST": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "CGST"
                },
                "SGST": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "CGST"
                },
                "ActualQty": {
                    "bsonType": [
                    "int",
                    "double"
                    ],
                    "description": "actual quantity of nuts in KGs"
                }
                }
            }
            }
            )
        col = db["SDespatch"]
        col.create_index([('BillNo',1)], unique=True)
        self.add_data()
        messagebox.showinfo("Success", f"{name} Database has been created")

    def search(self):
        string = self.search_entry.get()
        database_names = client.list_database_names()
        if string == "":
            return
        for i in database_names:
            if string in i:
                self.dblist.delete(0, 'end')
                self.dblist.insert(tk.END, string)
                return
        messagebox.showerror("Error", "Not Found")
        self.add_data()

    def clear_sel(self,event):
        self.dblist.selection_clear(0, tk.END)

class Main(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.frame = tk.Frame(self)
        self.frame.pack(expand=True, fill="both")
        self.place(relx=.20, y=0, relwidth=.8, relheight=1)
        self.frame.configure(bg='#F9FBFA')

    def draw(self, text):
        if hasattr(self, 'a'):
            self.a.destroy()
        self.a = A(self,text)
    
class A(ttk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = client[db]
        self.place(x=0, y=0, relwidth=1, relheight=1)

        self.book = ttk.Notebook(self)
        self.book.pack(fill="both", expand=True)

        self.tab_list1 = []
        self.tab_list2 = []
        col = self.db.list_collection_names()
        for tab in self.book.tabs():
            self.book.forget(tab)
        self.tab_list1.clear()
        otherthanA = ['Bills', 'Container','Stock', 'admin','config', 'local', 'GST' ]
        if db not in otherthanA:
            # for i in col:
            #     tab = tk.Frame(self.book)
            #     self.book.add(tab, text=i)
            #     self.tab_list1.append(tab)

            tab1 = tk.Frame(self.book)
            self.book.add(tab1, text="RCNutPurchase")
            tab2 = tk.Frame(self.book)
            self.book.add(tab2, text="CKDespatch")
            tab3 = tk.Frame(self.book)
            self.book.add(tab3, text="SDespatch")
            tab4 = tk.Frame(self.book)
            self.book.add(tab4, text="HDR")
            tab5 = tk.Frame(self.book)
            self.book.add(tab5, text="KernelList")
            tab6 = tk.Frame(self.book)
            self.book.add(tab6, text="JEChart")

            if hasattr(self, "rc"):
                self.rc.destroy()
            if hasattr(self, "sdd"):
                self.sdd.destroy()
            if hasattr(self, "ckn"):
                self.ckn.destroy()
            if hasattr(self, "hdr"):
                self.hdr.destroy()
            if hasattr(self, "chart"):
                self.chart.destroy()
            if hasattr(self, "ker"):
                self.ker.destroy()
            
        
            self.rc = rcn(tab1, db, "RCNutPurchase")
            self.sdd = sdd(tab3, db, "SDespatch")
            self.ckn = ckn(tab2, db, "CKDespatch")
            self.hdr = hdr(tab4, db, "HDR") 
            self.chart = chart(tab6, db, "JEChart")
            self.ker = ker(tab5, db, "KernelList")


        if db =="Stock":
            tab1 = tk.Frame(self.book)
            self.book.add(tab1, text="Sales")
            tab2 = tk.Frame(self.book)
            self.book.add(tab2, text="Purchase")
            tab3 = tk.Frame(self.book)
            self.book.add(tab3, text="Production")
            tab4 = tk.Frame(self.book)
            self.book.add(tab4, text="RCN")
            tab5 = tk.Frame(self.book)
            self.book.add(tab5, text="CKN")
            if hasattr(self, "sa"):
                self.sa.destroy()
            if hasattr(self, "pa"):
                self.pa.destroy()
            if hasattr(self, "pu"):
                self.pu.destroy()
            if hasattr(self, "r"):
                self.r.destroy()
            if hasattr(self, "ck"):
                self.ck.destroy()
            self.ss = ss(tab1, "Stock", "Sales")
            self.ps = ps(tab2, "Stock", "Purchase")
            self.pa = pa(tab3, "Stock", "Production")
            self.srcn = srcn(tab4, "Stock", "RCN")
            self.sckn = sckn(tab5, "Stock", "CKN")

        if db == "Bills":
            if hasattr(self, "psd"):
                self.psd.destroy()
            if hasattr(self, "sad"):
                self.sad.destroy()
            if hasattr(self, "gstapp"):
                self.gstapp.destroy()
            if hasattr(self, "cash"):
                self.cash.destroy()

            t1 = tk.Frame(self.book)
            self.book.add(t1, text="PSdetails")
            self.psd = psd(t1, "Bills", "PSdetails")

            t2 = tk.Frame(self.book)
            self.book.add(t2, text="SalesDetail")
            self.sad = sad(t2, "Bills", "Sdetails")
            
            t3 = tk.Frame(self.book)
            self.book.add(t3, text="GSTAppear")
            self.gstapp = gstapp(t3, "Bills", "GSTApp")
            
            t4 = tk.Frame(self.book)
            self.book.add(t4, text='CashBill')
            self.cash = cash(t4, "Bills", "Cash")
        
        if db == "GST":
            if hasattr(self, "gstpsd"):
                self.gstpsd.destroy()
            if hasattr(self, "gstsad"):
                self.gstsad.destroy()

            ti1 = tk.Frame(self.book)
            self.book.add(ti1, text="Purchase")
            self.gstpsd = gstpsd(ti1, "GST", "Purchase")

            ti2 = tk.Frame(self.book)
            self.book.add(ti2, text="Sales")
            self.gstsad = gstsad(ti2, "GST", "Sales")
        
        if db == "Container":
            if hasattr(self, "conw"):
                self.conw.destroy()
            if hasattr(self, "conp"):
                self.conp.destroy()
            if hasattr(self, "cont"):
                self.cont.destroy()
            tj1 = tk.Frame(self.book)
            self.book.add(tj1, text="Wholes")
            self.conw = conw(tj1, "Container", "Wholes")
            tj2 = tk.Frame(self.book)
            self.book.add(tj2, text="Pieces")
            self.conp = conp(tj2, "Container", "Pieces")
            tj3 = tk.Frame(self.book)
            self.book.add(tj3, text="Total")
            self.cont = cont(tj3, "Container", "Total")

class rcn(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        # treeview 
        # data = self.collection.find().sort("BillNo",1)
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'BillNo', 'Date', 'Rate', 'Bags', 'Quintal', 'KGs', \
                    'Amount' ,'SGST', 'CGST','IGST', 'Total')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("BillNo", width=80, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Rate", width=80, minwidth=25)
        self.tree.column("Bags", width=80, minwidth=25)
        self.tree.column("Quintal", width=80, minwidth=25)
        self.tree.column("KGs", width=80, minwidth=25)
        self.tree.column("Amount", width=100, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('BillNo', text='Bill NO')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Rate', text='Rate')
        self.tree.heading('Bags', text='Bags')
        self.tree.heading('Quintal', text='Quintal')
        self.tree.heading('KGs', text='KGs')
        self.tree.heading('Amount', text='Amount')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.BillNo_label = ttk.Label(self.data_Entry, text="BillNO:")
        self.BillNo_label.grid(row=0, column=0, sticky=tk.W)
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.Rate_label = ttk.Label(self.data_Entry, text="Rate:")
        self.Rate_label.grid(row=2, column=0, sticky=tk.W)

        self.Bags_label = ttk.Label(self.data_Entry, text="Bags:")
        self.Bags_label.grid(row=3, column=0, sticky=tk.W)

        self.Quintal_label = ttk.Label(self.data_Entry, text="Quintal:")
        self.Quintal_label.grid(row=4, column=0, sticky=tk.W)

        self.Quantity_label = ttk.Label(self.data_Entry, text="KGs:")
        self.Quantity_label.grid(row=5, column=0, sticky=tk.W)

        self.options = [
            "Bags", 
            "Kgs",
            "Quintals"
        ]
        self.clicked = tk.StringVar()
        self.clicked.set("Kgs")

        self.drop = tk.OptionMenu(self.data_Entry, self.clicked, *self.options)
        self.drop.grid(row=6, column=2, sticky=tk.W, padx=4, pady=4)

        self.amount_label = ttk.Label(self.data_Entry, text = "Amount:")
        self.amount_label.grid(row=6,column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=7, column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=8, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=9, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.BillNo_entry = ttk.Entry(self.data_Entry)
        self.BillNo_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.Rate_entry = ttk.Entry(self.data_Entry)
        self.Rate_entry.grid(row=2, column=1, sticky=tk.W)

        self.Bags_entry = ttk.Entry(self.data_Entry)
        self.Bags_entry.grid(row=3, column=1, sticky=tk.W)

        self.Quintal_entry = ttk.Entry(self.data_Entry)
        self.Quintal_entry.grid(row=4, column=1, sticky=tk.W)

        self.KGs_entry = ttk.Entry(self.data_Entry)
        self.KGs_entry.grid(row=5, column=1, sticky=tk.W)

        self.amount_entry = ttk.Entry(self.data_Entry)
        self.amount_entry.grid(row=6, column=1, sticky=tk.W)

        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=7, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=9, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry, width=12,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=10, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=10, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=10, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=10, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.calAmt_button = ttk.Button(self.data_Entry, text="Calculate", command=self.cal)
        self.calAmt_button.grid(row=6, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset", command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)
        
        self.reset_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.reset_button.grid(column=4, row=10, sticky=tk.NSEW, padx=3, pady=3)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("BillNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Rate'], i['Bags'], \
                                        i['Quintal'], i['KGs'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                            i['Total']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalRate": { "$sum": "$Rate" }, \
                                                        "totalBags":{"$sum":"$Bags" }, "totalQ": { "$sum": \
                                                            "$Quintal" }, "totalkg":{"$sum":"$KGs" }, \
                                                                "totalAmt": { "$sum": "$Amount" },\
                                                        "totalSGST":{"$sum":"$SGST" },"totalCGST": \
                                                            { "$sum": "$CGST" }, \
                                                            "totalIGST":{"$sum":"$IGST"}, "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalRate'], j['totalBags'], \
                                        j['totalQ'], j['totalkg'], j['totalAmt'], j['totalSGST'], j['totalCGST'],\
                                           j['totalIGST'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.BillNo_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.Rate_entry.delete(0, tk.END)
            self.Bags_entry.delete(0, tk.END)
            self.Quintal_entry.delete(0, tk.END)
            self.KGs_entry.delete(0, tk.END)
            self.amount_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            
            self.BillNo_entry.insert(0, values[1])
            self.Date_entry.entry.insert(0, values[2])
            self.Rate_entry.insert(0, values[3])
            self.Bags_entry.insert(0, values[4])
            self.Quintal_entry.insert(0, values[5])
            self.KGs_entry.insert(0, values[6])
            self.amount_entry.insert(0, values[7])
            self.SGST_entry.insert(0, values[8])
            self.CGST_entry.insert(0, values[9])
            self.IGST_entry.insert(0, values[10])

    def cal(self):
        value = self.clicked.get()
        self.amount_entry.delete(0, tk.END)
        try:
            if value=='Kgs':
                amount = float(self.Rate_entry.get())*float(self.KGs_entry.get())
                self.amount_entry.insert(0, amount)
            if value=='Quintals':
                amount = float(self.Rate_entry.get())*float(self.Quintal_entry.get())
                self.amount_entry.insert(0, amount)
            if value=='Bags':
                amount = float(self.Rate_entry.get())*float(self.Bags_entry.get())
                self.amount_entry.insert(0, amount)
        except Exception as e:
            messagebox.showerror('Error', e)

    def add(self):
        try:
            bill = int(self.BillNo_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            rate = float(self.Rate_entry.get())
            bag = float(self.Bags_entry.get())
            q = float(self.Quintal_entry.get())
            kg = float(self.KGs_entry.get())
            amt = float(self.amount_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            data = {'BillNo': bill, 'Date':date, 'Rate':rate, 'Bags':bag, "Quintal":q\
                    , "KGs":kg, 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }
            
            self.collection.insert_one(data)
            self.collection.update_many(filter={'Total': {'$exists': False}},update=\
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST', '$Amount']}}}])
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.BillNo_entry.delete(0, tk.END)
        self.Rate_entry.delete(0, tk.END)
        self.Bags_entry.delete(0, tk.END)
        self.Quintal_entry.delete(0, tk.END)
        self.KGs_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)

        self.Rate_entry.insert(0, 0)
        self.Bags_entry.insert(0, 0)
        self.Quintal_entry.insert(0, 0)
        self.KGs_entry.insert(0, 0)
        self.amount_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)

    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("BillNo",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalRate": { "$sum": "$Rate" }, \
                                                        "totalBags":{"$sum":"$Bags" }, "totalQ": { "$sum": \
                                                            "$Quintal" }, "totalkg":{"$sum":"$KGs" }, \
                                                                "totalAmt": { "$sum": "$Amount" },\
                                                        "totalSGST":{"$sum":"$SGST" },"totalCGST": \
                                                            { "$sum": "$CGST" }, "totalIGST":{"$sum":"$IGST"},\
                                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        counter = 1

        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Rate'], i['Bags'], \
                                        i['Quintal'], i['KGs'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                            i['Total']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalRate'], j['totalBags'], \
                                        j['totalQ'], j['totalkg'], j['totalAmt'],\
                                            j['totalSGST'], j['totalCGST'],j['totalIGST'],\
                                            j['total']), tags= 'total')

    def update(self):
        bill = int(self.BillNo_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        rate = float(self.Rate_entry.get())
        bag = float(self.Bags_entry.get())
        q = float(self.Quintal_entry.get())
        kg = float(self.KGs_entry.get())
        amt = float(self.amount_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get())
        igst = float(self.IGST_entry.get())
        self.collection.update_one({"BillNo": bill}, {'$set':{'BillNo': bill, \
                                        'Date':date, 'Rate':rate, 'Bags':bag, "Quintal":q\
                    , "KGs":kg, 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst}}) 
        self.collection.update_many({"BillNo": bill}, update=\
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST', '$IGST','$Amount']}}}])
        self.showall()  

    def delete(self):
        bill = int(self.BillNo_entry.get())
        self.collection.delete_one({'BillNo':bill})
        self.reset()
        self.showall()

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'BillNo', 
                    'Date', 'Rate', 'Bags', 'Quintal', 'KGs', 
                    'Amount' ,'SGST', 'CGST','IGST', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

class ckn(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'BillNo', 'Date', 'Qty', 'WRate', 'Wholes', 'PRate', \
                    'Pieces' ,'Amount', 'SGST', 'CGST','IGST', 'Total')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("BillNo", width=80, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Qty", width=80, minwidth=25)
        self.tree.column("WRate", width=80, minwidth=25)
        self.tree.column("Wholes", width=80, minwidth=25)
        self.tree.column("PRate", width=80, minwidth=25)
        self.tree.column("Pieces", width=100, minwidth=25)
        self.tree.column("Amount", width=80, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('BillNo', text='Bill NO')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Qty', text='QTY')
        self.tree.heading('WRate', text='WRate')
        self.tree.heading('Wholes', text='Wholes')
        self.tree.heading('PRate', text='PRate')
        self.tree.heading('Pieces', text='Pieces')
        self.tree.heading('Amount', text='Amount')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.BillNo_label = ttk.Label(self.data_Entry, text="BillNO:")
        self.BillNo_label.grid(row=0, column=0, sticky=tk.W)
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.Qty_label = ttk.Label(self.data_Entry, text="Quantity:")
        self.Qty_label.grid(row=2, column=0, sticky=tk.W)

        self.WRate_label = ttk.Label(self.data_Entry, text="WRate:")
        self.WRate_label.grid(row=3, column=0, sticky=tk.W)

        self.Wholes_label = ttk.Label(self.data_Entry, text="Wholes:")
        self.Wholes_label.grid(row=4, column=0, sticky=tk.W)

        self.PRate_label = ttk.Label(self.data_Entry, text="PRate:")
        self.PRate_label.grid(row=5, column=0, sticky=tk.W)

        self.Pieces_label = ttk.Label(self.data_Entry, text="Pieces:")
        self.Pieces_label.grid(row=6, column=0, sticky=tk.W)

        self.amount_label = ttk.Label(self.data_Entry, text = "Amount:")
        self.amount_label.grid(row=7,column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=8, column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=9, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=10, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.BillNo_entry = ttk.Entry(self.data_Entry)
        self.BillNo_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.Qty_entry = ttk.Entry(self.data_Entry)
        self.Qty_entry.grid(row=2, column=1, sticky=tk.W)

        self.WRate_entry = ttk.Entry(self.data_Entry)
        self.WRate_entry.grid(row=3, column=1, sticky=tk.W)

        self.Wholes_entry = ttk.Entry(self.data_Entry)
        self.Wholes_entry.grid(row=4, column=1, sticky=tk.W)

        self.PRate_entry = ttk.Entry(self.data_Entry)
        self.PRate_entry.grid(row=5, column=1, sticky=tk.W)

        self.Pieces_entry = ttk.Entry(self.data_Entry)
        self.Pieces_entry.grid(row=6, column=1, sticky=tk.W)

        self.amount_entry = ttk.Entry(self.data_Entry)
        self.amount_entry.grid(row=7, column=1, sticky=tk.W)
        
        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=9, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=10, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.calAmt_button = ttk.Button(self.data_Entry, text="Calculate", command=self.cal)
        self.calAmt_button.grid(row=7, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

        self.reset_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.reset_button.grid(column=4, row=11, padx=3, pady=3)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("BillNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Qty'], i['WRate'], \
                                        i['Wholes'], i['PRate'], i['Pieces'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                            i['Total']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, \
                                                        "totalWR":{"$sum":"$WRate" }, "totalW": { "$sum": \
                                                            "$Wholes" }, "totalPR":{"$sum":"$PRate" }, \
                                                            "totalP":{"$sum":"$Pieces"},\
                                                                "totalAmt": { "$sum": "$Amount" },\
                                                        "totalSGST":{"$sum":"$SGST" },"totalCGST": \
                                                            { "$sum": "$CGST" }, \
                                                            "totalIGST":{"$sum":"$IGST"}, "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalQty'], j['totalWR'], \
                                        j['totalW'], j['totalPR'], j['totalP'],j['totalAmt'], j['totalSGST'], j['totalCGST'],\
                                           j['totalIGST'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.BillNo_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.Qty_entry.delete(0, tk.END)
            self.WRate_entry.delete(0, tk.END)
            self.Wholes_entry.delete(0, tk.END)
            self.PRate_entry.delete(0, tk.END)
            self.Pieces_entry.delete(0, tk.END)
            self.amount_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            
            self.BillNo_entry.insert(0, values[1])
            self.Date_entry.entry.insert(0, values[2])
            self.Qty_entry.insert(0, values[3])
            self.WRate_entry.insert(0, values[4])
            self.Wholes_entry.insert(0, values[5])
            self.PRate_entry.insert(0, values[6])
            self.Pieces_entry.insert(0, values[7])
            self.amount_entry.insert(0, values[8])
            self.SGST_entry.insert(0, values[9])
            self.CGST_entry.insert(0, values[10])
            self.IGST_entry.insert(0, values[11])

    def cal(self):
        self.amount_entry.delete(0, tk.END)
        try:
            amount = float(self.PRate_entry.get())*float(self.Pieces_entry.get())+float(self.WRate_entry.get())*float(self.Wholes_entry.get())
            self.amount_entry.insert(0, amount)
           
        except Exception as e:
            messagebox.showerror('Error', e)

    def add(self):
        try:
            bill = int(self.BillNo_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            qty = float(self.Qty_entry.get())
            prate = float(self.PRate_entry.get())
            p = float(self.Pieces_entry.get())
            wrate = float(self.WRate_entry.get())
            w = float(self.Wholes_entry.get())
            amt = float(self.amount_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            data = {'BillNo': bill, 'Date':date, 'Qty':qty, 'WRate':wrate,'Wholes':w,\
                    'PRate':prate, 'Pieces':p\
                    , 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }
            #data1 = {[{$addFields: {Total: { $add: ["$SGST", "$CGST", "$Amount"] }}}]}
            
            self.collection.insert_one(data)
            self.collection.update_many(filter={'Total': {'$exists': False}},update=\
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST', '$Amount']}}}])
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.BillNo_entry.delete(0, tk.END)
        self.Qty_entry.delete(0, tk.END)
        self.WRate_entry.delete(0, tk.END)
        self.Wholes_entry.delete(0, tk.END)
        self.PRate_entry.delete(0, tk.END)
        self.Pieces_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)

        self.Qty_entry.insert(0, 0)
        self.WRate_entry.insert(0, 0)
        self.Wholes_entry.insert(0, 0)
        self.PRate_entry.insert(0, 0)
        self.Pieces_entry.insert(0, 0)
        self.amount_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)

    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("BillNo",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, \
                                                        "totalWR":{"$sum":"$WRate" }, "totalW": { "$sum": \
                                                            "$Wholes" }, "totalPR":{"$sum":"$PRate" }, \
                                                            "totalP":{"$sum":"$Pieces"},\
                                                                "totalAmt": { "$sum": "$Amount" },\
                                                        "totalSGST":{"$sum":"$SGST" },"totalCGST": \
                                                            { "$sum": "$CGST" }, \
                                                            "totalIGST":{"$sum":"$IGST"}, "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Qty'], i['WRate'], \
                                        i['Wholes'], i['PRate'], i['Pieces'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                            i['Total']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalQty'], j['totalWR'], \
                                        j['totalW'], j['totalPR'], j['totalP'],j['totalAmt'], j['totalSGST'], j['totalCGST'],\
                                           j['totalIGST'], j['total']), tags= 'total')

    def update(self):
        bill = int(self.BillNo_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        qty = float(self.Qty_entry.get())
        prate = float(self.PRate_entry.get())
        p = float(self.Pieces_entry.get())
        wrate = float(self.WRate_entry.get())
        w = float(self.Wholes_entry.get())
        amt = float(self.amount_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get()) 
        igst = float(self.IGST_entry.get())
        self.collection.update_one({"BillNo": bill}, {'$set':{'BillNo': bill, 'Date':date, 'Qty':qty, 'WRate':wrate,'Wholes':w,\
                    'PRate':prate, 'Pieces':p\
                    , 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }}) 
        self.collection.update_many({"BillNo": bill}, update=\
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST', '$IGST','$Amount']}}}])
        self.showall()  

    def delete(self):
        bill = int(self.BillNo_entry.get())
        self.collection.delete_one({'BillNo':bill})
        self.reset()
        self.showall()

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'BillNo', 'Date', 'Qty', 'WRate', 'Wholes', 'PRate', \
                    'Pieces' ,'Amount', 'SGST', 'CGST','IGST', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

class sdd(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        # treeview 
        # data = self.collection.find().sort("BillNo",1)
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'BillNo', 'Date', 'Rate', 'Bags', 'KGs', \
                    'Amount' ,'SGST', 'CGST','IGST','Total', 'ActualQty')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("BillNo", width=80, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Rate", width=80, minwidth=25)
        self.tree.column("Bags", width=80, minwidth=25)
        self.tree.column("KGs", width=80, minwidth=25)
        self.tree.column("Amount", width=100, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column("ActualQty", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('BillNo', text='Bill NO')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Rate', text='Rate')
        self.tree.heading('Bags', text='Bags')
        self.tree.heading('KGs', text='KGs')
        self.tree.heading('Amount', text='Amount')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('Total', text='Total')
        self.tree.heading('ActualQty', text='ActualQty')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.BillNo_label = ttk.Label(self.data_Entry, text="BillNO:")
        self.BillNo_label.grid(row=0, column=0, sticky=tk.W)
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.Rate_label = ttk.Label(self.data_Entry, text="Rate:")
        self.Rate_label.grid(row=2, column=0, sticky=tk.W)

        self.Bags_label = ttk.Label(self.data_Entry, text="Bags:")
        self.Bags_label.grid(row=3, column=0, sticky=tk.W)

        self.Quantity_label = ttk.Label(self.data_Entry, text="KGs:")
        self.Quantity_label.grid(row=4, column=0, sticky=tk.W)

        self.options = [
            "Bags", 
            "Kgs"
        ]
        self.clicked = tk.StringVar()
        self.clicked.set("Kgs")

        self.drop = tk.OptionMenu(self.data_Entry, self.clicked, *self.options)
        self.drop.grid(row=5, column=2, sticky=tk.W)

        self.amount_label = ttk.Label(self.data_Entry, text = "Amount:")
        self.amount_label.grid(row=5,column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=6, column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=7, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=8, column=0, sticky=tk.W)

        self.Actual_label = ttk.Label(self.data_Entry, text="Actual Qty:")
        self.Actual_label.grid(row=9, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.BillNo_entry = ttk.Entry(self.data_Entry)
        self.BillNo_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.Rate_entry = ttk.Entry(self.data_Entry)
        self.Rate_entry.grid(row=2, column=1, sticky=tk.W)

        self.Bags_entry = ttk.Entry(self.data_Entry)
        self.Bags_entry.grid(row=3, column=1, sticky=tk.W)


        self.KGs_entry = ttk.Entry(self.data_Entry)
        self.KGs_entry.grid(row=4, column=1, sticky=tk.W)

        self.amount_entry = ttk.Entry(self.data_Entry)
        self.amount_entry.grid(row=5, column=1, sticky=tk.W)

        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=6, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=7, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.Actual_entry = ttk.Entry(self.data_Entry)
        self.Actual_entry.grid(row=9, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        # #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=10, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=10, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=10, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=10, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.calAmt_button = ttk.Button(self.data_Entry, text="Calculate", command=self.cal)
        self.calAmt_button.grid(row=5, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset", command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.showex_button = ttk.Button(parent, text="Show Excess-Less table", command=self.exls)
        self.showex_button.place(relx=.8, rely=.2)

        self.reset_button = ttk.Button(self.data_Entry, text="Export", command=self.save1)
        self.reset_button.grid(column=4, row=10, sticky=tk.NSEW, padx=3, pady=3)

    def save1(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'BillNo', 'Date', 'Rate', 'Bags', 'KGs', \
                    'Amount' ,'SGST', 'CGST','IGST','Total', 'ActualQty'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)
    
    def save2(self):
        data = []
        for item in self.tree1.get_children():
            values = self.tree1.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'BillNo', 'Date', 'Rate', 'Bags', 'KGs', \
                    'ActualQty', 'Excess', 'Less', 'ExAmount', 'LsAmount'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def exls(self):
        frame = tk.Toplevel()
        frame.geometry('1000x600')
        self.from_entry1 = ttk.DateEntry(frame,
                            dateformat='%Y-%m-%d')
        self.from_entry1.grid(row=0, column=1, sticky=tk.W)

        self.to_entry1 = ttk.DateEntry(frame,
                            dateformat='%Y-%m-%d')
        self.to_entry1.grid(row=0, column=3, sticky=tk.W)
        self.from_label1 = ttk.Label(frame, text="From:")
        self.from_label1.grid(row=0, column= 0, sticky=tk.W)

        self.to_label1 = ttk.Label(frame, text="To:")
        self.to_label1.grid(row=0, column=2,sticky=tk.W)

        self.findRange_button1 = ttk.Button(frame, text="Find Range", command=self.rangeexl)
        self.findRange_button1.grid(row=1, column=0, sticky=tk.W, padx=3, pady=3)

        self.reset_button = ttk.Button(frame, text="Export", command=self.save2)
        self.reset_button.grid(column=1, row=1, sticky=tk.W, padx=3, pady=3)


        self.tree1 = ttk.Treeview(frame)
        self.tree1['columns'] = ('SL.No', 'BillNo', 'Date', 'Rate', 'Bags', 'KGs', \
                    'ActualQty', 'Excess', 'Less', 'ExAmount', 'LsAmount')
        self.tree1.column("SL.No", width=50, minwidth=25)
        self.tree1.column("BillNo", width=80, minwidth=25)
        self.tree1.column("Date", width=130, minwidth=25)
        self.tree1.column("Rate", width=80, minwidth=25)
        self.tree1.column("Bags", width=80, minwidth=25)
        self.tree1.column("KGs", width=80, minwidth=25)
        self.tree1.column("ActualQty", width=80, minwidth=25)
        self.tree1.column("Excess", width=100, minwidth=25)
        self.tree1.column("Less", width=80, minwidth=25)
        self.tree1.column("ExAmount", width=80, minwidth=25)
        self.tree1.column("LsAmount", width=80, minwidth=25)
    
        self.tree1.column('#0',width=0, stretch=tk.NO )
        self.tree1.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree1.heading('BillNo', text='Bill NO')
        self.tree1.heading('Date', text='Date')
        self.tree1.heading('Rate', text='Rate')
        self.tree1.heading('Bags', text='Bags')
        self.tree1.heading('KGs', text='KGs')
        self.tree1.heading('ActualQty', text='ActualQty')
        self.tree1.heading('Excess', text='Excess')
        self.tree1.heading('Less', text='Less')
        self.tree1.heading('ExAmount', text='ExAmount')
        self.tree1.heading('LsAmount', text='LsAmount')
        self.tree1.place(relwidth=1, rely=.1, relheight=1)
        self.tree1.bind("<Double-1>", self.click)
        self.tree1.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        data = self.collection.find().sort("BillNo",1)
        counter = 1
        for i in data:
            self.tree1.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Rate'], i['Bags'], \
                                         i['KGs'], i['ActualQty'], i['Excess'], i['Less'], i['ExAmount'], i['LsAmount']))
            counter+=1
        
        pipeline = [{ "$group": { "_id": None, "totalRate": { "$sum": "$Rate" }, \
                                                        "totalBags":{"$sum":"$Bags" }, "totalkg":{"$sum":"$KGs" }, \
                                                                "totalEx": { "$sum": "$Excess" },\
                                                        "totalLs":{"$sum":"$Less" },"totalExamt": \
                                                            { "$sum": "$ExAmount" }, \
                                                            "totalLsamt":{"$sum":"$LsAmount"}\
                                                                , "totalA": { "$sum": "$ActualQty" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree1.insert('', 'end',values=('Total',  0, 0, j['totalRate'], j['totalBags'], \
                                         j['totalkg'], j['totalA'], j['totalEx'], j['totalLs'], j['totalExamt'],\
                                           j['totalLsamt']), tags= 'total')
    def rangeexl(self):
        self.tree1.delete(*self.tree1.get_children())
        fromdate = self.from_entry1.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry1.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("BillNo",1)
        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},{ "$group": { "_id": None, "totalRate": { "$sum": "$Rate" }, \
                                                        "totalBags":{"$sum":"$Bags" }, "totalkg":{"$sum":"$KGs" }, \
                                                                "totalEx": { "$sum": "$Excess" },\
                                                        "totalLs":{"$sum":"$Less" },"totalExamt": \
                                                            { "$sum": "$ExAmount" }, \
                                                            "totalLsamt":{"$sum":"$LsAmount"}\
                                                                , "totalA": { "$sum": "$ActualQty" }}}]
        result = self.collection.aggregate(pipeline)
        counter = 1
        for i in data:
            self.tree1.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Rate'], i['Bags'], \
                                         i['KGs'], i['ActualQty'], i['Excess'], i['Less'], i['ExAmount'], i['LsAmount']))
            counter+=1
        for j in result:
            self.tree1.insert('', 'end',values=('Total',  0, 0, j['totalRate'], j['totalBags'], \
                                         j['totalkg'], j['totalA'], j['totalEx'], j['totalLs'], j['totalExamt'],\
                                           j['totalLsamt']), tags= 'total')

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("BillNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Rate'], i['Bags'], \
                                         i['KGs'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                            i['Total'],i['ActualQty']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalRate": { "$sum": "$Rate" }, \
                                                        "totalBags":{"$sum":"$Bags" }, "totalkg":{"$sum":"$KGs" }, \
                                                                "totalAmt": { "$sum": "$Amount" },\
                                                        "totalSGST":{"$sum":"$SGST" },"totalCGST": \
                                                            { "$sum": "$CGST" }, \
                                                            "totalIGST":{"$sum":"$IGST"}, "total":{"$sum":"$Total" }\
                                                                , "totalA": { "$sum": "$ActualQty" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalRate'], j['totalBags'], \
                                         j['totalkg'], j['totalAmt'], j['totalSGST'], j['totalCGST'],\
                                           j['totalIGST'], j['total'],j['totalA']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.BillNo_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.Rate_entry.delete(0, tk.END)
            self.Bags_entry.delete(0, tk.END)
            self.KGs_entry.delete(0, tk.END)
            self.amount_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            self.Actual_entry.delete(0, tk.END)
            
            self.BillNo_entry.insert(0, values[1])
            self.Date_entry.entry.insert(0, values[2])
            self.Rate_entry.insert(0, values[3])
            self.Bags_entry.insert(0, values[4])
            self.KGs_entry.insert(0, values[5])
            self.amount_entry.insert(0, values[6])
            self.SGST_entry.insert(0, values[7])
            self.CGST_entry.insert(0, values[8])
            self.IGST_entry.insert(0, values[9])
            self.Actual_entry.insert(0, values[11])

    def cal(self):
        value = self.clicked.get()
        self.amount_entry.delete(0, tk.END)
        try:
            if value=='Kgs':
                amount = float(self.Rate_entry.get())*float(self.KGs_entry.get())
                self.amount_entry.insert(0, amount)
            if value=='Bags':
                amount = float(self.Rate_entry.get())*float(self.Bags_entry.get())
                self.amount_entry.insert(0, amount)
        except Exception as e:
            messagebox.showerror('Error', e)

    def add(self):
        try:
            bill = int(self.BillNo_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            rate = float(self.Rate_entry.get())
            bag = float(self.Bags_entry.get())
            kg = float(self.KGs_entry.get())
            amt = float(self.amount_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            a = float(self.Actual_entry.get())
            data = {'BillNo': bill, 'Date':date, 'Rate':rate, 'Bags':bag\
                    , "KGs":kg, 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst, "ActualQty":a }
            
            self.collection.insert_one(data)
            self.collection.update_one({"BillNo": bill}, {'$set':{'BillNo': bill, \
                                        'Date':date, 'Rate':rate, 'Bags':bag\
                    , "KGs":kg, 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst, "ActualQty":a}}) 
            self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST','$Amount']}}}]
            )
            self.collection.update_one({"BillNo": bill},
                                [{'$addFields':  {'Excess': { '$max': [0, { '$subtract': ['$ActualQty','$KGs'] }] }}}]
            )
            self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'Less': { '$max': [0, { "$subtract": ['$KGs', '$ActualQty'] }] }}}]
            )
            self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'ExAmount': { '$max': [0, {"$multiply" :[ {"$subtract": ['$ActualQty', '$KGs']},"$Rate" ]}] }}}]
            )
            self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'LsAmount': { '$max': [0, {"$multiply" :[ {"$subtract": ['$KGs', '$ActualQty']},"$Rate" ]}] }}}]
            )
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.BillNo_entry.delete(0, tk.END)
        self.Rate_entry.delete(0, tk.END)
        self.Bags_entry.delete(0, tk.END)
        self.KGs_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)
        self.Actual_entry.delete(0, tk.END)

        self.Rate_entry.insert(0, 0)
        self.Bags_entry.insert(0, 0)
        self.KGs_entry.insert(0, 0)
        self.amount_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)
        self.Actual_entry.insert(0, 0)

    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("BillNo",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalRate": { "$sum": "$Rate" }, \
                                                        "totalBags":{"$sum":"$Bags" }, "totalkg":{"$sum":"$KGs" }, \
                                                                "totalAmt": { "$sum": "$Amount" },\
                                                        "totalSGST":{"$sum":"$SGST" },"totalCGST": \
                                                            { "$sum": "$CGST" }, \
                                                            "totalIGST":{"$sum":"$IGST"}, "total":{"$sum":"$Total" }\
                                                                , "totalA": { "$sum": "$ActualQty" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        counter = 1

        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Rate'], i['Bags'], \
                                         i['KGs'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                            i['Total'],i['ActualQty']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalRate'], j['totalBags'], \
                                         j['totalkg'], j['totalAmt'], j['totalSGST'], j['totalCGST'],\
                                           j['totalIGST'], j['total'],j['totalA']), tags= 'total')

    def update(self):
        bill = int(self.BillNo_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        rate = float(self.Rate_entry.get())
        bag = float(self.Bags_entry.get())
        a = float(self.Actual_entry.get())
        kg = float(self.KGs_entry.get())
        amt = float(self.amount_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get())
        igst = float(self.IGST_entry.get())
        self.collection.update_one({"BillNo": bill}, {'$set':{'BillNo': bill, \
                                        'Date':date, 'Rate':rate, 'Bags':bag\
                    , "KGs":kg, 'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst, "ActualQty":a}}) 
        self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST','$Amount']}}}]
            )
        self.collection.update_one({"BillNo": bill},
                                [{'$addFields':  {'Excess': { '$max': [0, { '$subtract': ['$ActualQty','$KGs'] }] }}}]
            )
        self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'Less': { '$max': [0, { "$subtract": ['$KGs', '$ActualQty'] }] }}}]
            )
        self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'ExAmount': { '$max': [0, {"$multiply" :[ {"$subtract": ['$ActualQty', '$KGs']},"$Rate" ]}] }}}]
            )
        self.collection.update_one({"BillNo": bill},
                                [{'$addFields': {'LsAmount': { '$max': [0, {"$multiply" :[ {"$subtract": ['$KGs', '$ActualQty']},"$Rate" ]}] }}}]
            )
        self.showall()  

    def delete(self):
        bill = int(self.BillNo_entry.get())
        self.collection.delete_one({'BillNo':bill})
        self.reset()
        self.showall()

class hdr(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Date', 'HRate', 'Husk', 'HuskAmt', 'DRate', \
                    'Dust' ,'DustAmt', 'RRate', 'Rejection','RejectionAmt', 'GRate', \
                        'Gbags', 'GbagAmt', 'Total')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("HRate", width=80, minwidth=25)
        self.tree.column("Husk", width=80, minwidth=25)
        self.tree.column("HuskAmt", width=80, minwidth=25)
        self.tree.column("DRate", width=80, minwidth=25)
        self.tree.column("Dust", width=80, minwidth=25)
        self.tree.column("DustAmt", width=100, minwidth=25)
        self.tree.column("RRate", width=80, minwidth=25)
        self.tree.column("Rejection", width=80, minwidth=25)
        self.tree.column("RejectionAmt", width=80, minwidth=25)
        self.tree.column("GRate", width=80, minwidth=25)
        self.tree.column("Gbags", width=80, minwidth=25)
        self.tree.column("GbagAmt", width=80, minwidth=25)
        self.tree.column("Total", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('HRate', text='HRate')
        self.tree.heading('Husk', text='Husk')
        self.tree.heading('HuskAmt', text='HuskAmt')
        self.tree.heading('DRate', text='DRate')
        self.tree.heading('Dust', text='Dust')
        self.tree.heading('DustAmt', text='DustAmt')
        self.tree.heading('RRate', text='RRate')
        self.tree.heading('Rejection', text='Rejection')
        self.tree.heading('RejectionAmt', text='RejectionAmt')
        self.tree.heading('GRate', text='GRate')
        self.tree.heading('Gbags', text='Gbags')
        self.tree.heading('GbagAmt', text='GbagAmt')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.sl_label = ttk.Label(self.data_Entry, text="Sl.no:")
        self.sl_label.grid(row=0, column=2, sticky='nsew')

        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=0, column=0, sticky=tk.W)
        
        self.HRate_label = ttk.Label(self.data_Entry, text="Husk Rate:")
        self.HRate_label.grid(row=1, column=0, sticky=tk.W)

        self.Husk_label = ttk.Label(self.data_Entry, text="Husk:")
        self.Husk_label.grid(row=2, column=0, sticky=tk.W)

        self.HuskAmt_label = ttk.Label(self.data_Entry, text="Husk Amount:")
        self.HuskAmt_label.grid(row=3, column=0, sticky=tk.W)

        self.DRate_label = ttk.Label(self.data_Entry, text="Dust Rate:")
        self.DRate_label.grid(row=4, column=0, sticky=tk.W)

        self.Dust_label = ttk.Label(self.data_Entry, text="Dust:")
        self.Dust_label.grid(row=5, column=0, sticky=tk.W)

        self.DustAmt_label = ttk.Label(self.data_Entry, text="Dust Amount")
        self.DustAmt_label.grid(row=6, column=0, sticky=tk.W)

        self.RRate_label = ttk.Label(self.data_Entry, text = "Rejection Rate:")
        self.RRate_label.grid(row=7,column=0, sticky=tk.W)

        self.Rej_label = ttk.Label(self.data_Entry, text="Rejection:")
        self.Rej_label.grid(row=8, column=0, sticky=tk.W)

        self.RejAmt_label = ttk.Label(self.data_Entry, text="Rejection Amount:")
        self.RejAmt_label.grid(row=9, column=0, sticky=tk.W)

        self.GRate_label = ttk.Label(self.data_Entry, text="Gunny Rate:")
        self.GRate_label.grid(row=10, column=0, sticky=tk.W)

        self.Gbags_label = ttk.Label(self.data_Entry, text="Gunny Bags:")
        self.Gbags_label.grid(row=11, column=0, sticky=tk.W)

        self.GbagAmt_label = ttk.Label(self.data_Entry, text="Gunny Amount:")
        self.GbagAmt_label.grid(row=12, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.sl_entry = ttk.Entry(self.data_Entry)
        self.sl_entry.grid(row=0, column=3, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=0, column=1, sticky=tk.W)

        self.HRate_entry = ttk.Entry(self.data_Entry)
        self.HRate_entry.grid(row=1, column=1, sticky=tk.W)

        self.Husk_entry = ttk.Entry(self.data_Entry)
        self.Husk_entry.grid(row=2, column=1, sticky=tk.W)

        self.HuskAmt_entry = ttk.Entry(self.data_Entry)
        self.HuskAmt_entry.grid(row=3, column=1, sticky=tk.W)

        self.DRate_entry = ttk.Entry(self.data_Entry)
        self.DRate_entry.grid(row=4, column=1, sticky=tk.W)

        self.Dust_entry = ttk.Entry(self.data_Entry)
        self.Dust_entry.grid(row=5, column=1, sticky=tk.W)

        self.DustAmt_entry = ttk.Entry(self.data_Entry)
        self.DustAmt_entry.grid(row=6, column=1, sticky=tk.W)

        self.RRate_entry = ttk.Entry(self.data_Entry)
        self.RRate_entry.grid(row=7, column=1, sticky=tk.W)
        
        self.Rej_entry = ttk.Entry(self.data_Entry)
        self.Rej_entry.grid(row=8, column=1, sticky=tk.W)

        self.RejAmt_entry = ttk.Entry(self.data_Entry)
        self.RejAmt_entry.grid(row=9, column=1, sticky=tk.W)

        self.GRate_entry = ttk.Entry(self.data_Entry)
        self.GRate_entry.grid(row=10, column=1, sticky=tk.W)

        self.Gbags_entry = ttk.Entry(self.data_Entry)
        self.Gbags_entry.grid(row=11, column=1, sticky=tk.W)

        self.GbagAmt_entry = ttk.Entry(self.data_Entry)
        self.GbagAmt_entry.grid(row=12, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.W, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Date', 'HRate', 'Husk', 
                                'HuskAmt', 'DRate',
                                'Dust' ,'DustAmt', 'RRate', 
                                'Rejection','RejectionAmt', 'GRate',
                                'Gbags', 'GbagAmt', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Sl",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(i['Sl'],  i['Date'], i['HRate'], i['Husk'], i['HuskAmt'], \
                                        i['DRate'], i['Dust'], i['DustAmt'], i['RRate'], i['Rejection'], i['RejectionAmt'],\
                                            i['GRate'],i['Gbags'], i['GbagAmt'],\
                                            i['Total']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalHR": { "$sum": "$HRate" }, \
                                                        "totalH":{"$sum":"$Husk" }, "totalHA": { "$sum": \
                                                            "$HuskAmt" }, "totalDR":{"$sum":"$DRate" }, \
                                                            "totalD":{"$sum":"$Dust"},\
                                                                "totalDA": { "$sum": "$DustAmt" },\
                                                        "totalRR":{"$sum":"$RRate" },"totalR": \
                                                            { "$sum": "$Rejection" }, \
                                                            "totalRA":{"$sum":"$RejectionAmt"},\
                                                                 "totalGR":{"$sum":"$GRate"},\
                                                                     "totalG":{"$sum":"$Gbags"},\
                                                                     "totalGA":{"$sum":"$GbagAmt"},\
                                                                          "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totalHR'], j['totalH'], \
                                        j['totalHA'], j['totalDR'], j['totalD'],j['totalDA'], j['totalRR'], j['totalR'],\
                                           j['totalRA'], j['totalGR'],j['totalG'],j['totalGA'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.sl_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.HRate_entry.delete(0, tk.END)
            self.Husk_entry.delete(0, tk.END)
            self.HuskAmt_entry.delete(0, tk.END)
            self.DRate_entry.delete(0, tk.END)
            self.Dust_entry.delete(0, tk.END)
            self.DustAmt_entry.delete(0, tk.END)
            self.RRate_entry.delete(0, tk.END)
            self.Rej_entry.delete(0, tk.END)
            self.RejAmt_entry.delete(0, tk.END)
            self.GRate_entry.delete(0, tk.END)
            self.Gbags_entry.delete(0, tk.END)
            self.GbagAmt_entry.delete(0, tk.END)
            
            self.sl_entry.insert(0, values[0])
            self.Date_entry.entry.insert(0, values[1])
            self.HRate_entry.insert(0, values[2])
            self.Husk_entry.insert(0, values[3])
            self.HuskAmt_entry.insert(0, values[4])
            self.DRate_entry.insert(0, values[5])
            self.Dust_entry.insert(0, values[6])
            self.DustAmt_entry.insert(0, values[7])
            self.RRate_entry.insert(0, values[8])
            self.Rej_entry.insert(0, values[9])
            self.RejAmt_entry.insert(0, values[10])
            self.GRate_entry.insert(0, values[11])
            self.Gbags_entry.insert(0, values[12])
            self.GbagAmt_entry.insert(0, values[13])

    def cal(self):
        self.amount_entry.delete(0, tk.END)
        try:
            amount = float(self.PRate_entry.get())*float(self.Pieces_entry.get())+float(self.WRate_entry.get())*float(self.Wholes_entry.get())
            self.amount_entry.insert(0, amount)
           
        except Exception as e:
            messagebox.showerror('Error', e)

    def add(self):
        try:
            sl = int(self.sl_entry.get())
            hr = int(self.HRate_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            h = float(self.Husk_entry.get())
            ha = float(self.HuskAmt_entry.get())
            dr = float(self.DRate_entry.get())
            d = float(self.Dust_entry.get())
            da = float(self.DustAmt_entry.get())
            rr = float(self.RRate_entry.get())
            r = float(self.Rej_entry.get())
            ra = float(self.RejAmt_entry.get()) 
            gr = float(self.GRate_entry.get())
            g = float(self.Gbags_entry.get())
            ga = float(self.GbagAmt_entry.get())
            data = {    
                        "Sl": sl,
                        "Date": date,
                        "HRate": hr,
                        "Husk": h,
                        "HuskAmt": ha,
                        "DRate": dr,
                        "Dust": d,
                        "DustAmt": da,
                        "RRate": rr,
                        "Rejection": r,
                        "RejectionAmt": ra,
                        "GRate": gr,
                        "Gbags": g,
                        "GbagAmt": ga,
                    }
            
            self.collection.insert_one(data)
            self.collection.update_one(filter={'Sl': sl},update=\
                                [{'$addFields': {'Total': {'$add': ['$HuskAmt', '$DustAmt','$RejectionAmt', '$GbagAmt']}}}])
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.sl_entry.delete(0, tk.END)
        self.HRate_entry.delete(0, tk.END)
        self.Husk_entry.delete(0, tk.END)
        self.HuskAmt_entry.delete(0, tk.END)
        self.DRate_entry.delete(0, tk.END)
        self.Dust_entry.delete(0, tk.END)
        self.DustAmt_entry.delete(0, tk.END)
        self.RRate_entry.delete(0, tk.END)
        self.Rej_entry.delete(0, tk.END)
        self.RejAmt_entry.delete(0, tk.END)
        self.GRate_entry.delete(0, tk.END)
        self.Gbags_entry.delete(0, tk.END)
        self.GbagAmt_entry.delete(0, tk.END)
            
        self.sl_entry.insert(0, 0)
        self.HRate_entry.insert(0, 0)
        self.Husk_entry.insert(0, 0)
        self.HuskAmt_entry.insert(0, 0)
        self.DRate_entry.insert(0, 0)
        self.Dust_entry.insert(0, 0)
        self.DustAmt_entry.insert(0, 0)
        self.RRate_entry.insert(0, 0)
        self.Rej_entry.insert(0, 0)
        self.RejAmt_entry.insert(0, 0)
        self.GRate_entry.insert(0, 0)
        self.Gbags_entry.insert(0, 0)
        self.GbagAmt_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Sl",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalHR": { "$sum": "$HRate" }, \
                                                        "totalH":{"$sum":"$Husk" }, "totalHA": { "$sum": \
                                                            "$HuskAmt" }, "totalDR":{"$sum":"$DRate" }, \
                                                            "totalD":{"$sum":"$Dust"},\
                                                                "totalDA": { "$sum": "$DustAmt" },\
                                                        "totalRR":{"$sum":"$RRate" },"totalR": \
                                                            { "$sum": "$Rejection" }, \
                                                            "totalRA":{"$sum":"$RejectionAmt"},\
                                                                 "totalGR":{"$sum":"$GRate"},\
                                                                     "totalG":{"$sum":"$Gbags"},\
                                                                     "totalGA":{"$sum":"$GbagAmt"},\
                                                                          "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(i['Sl'],  i['Date'], i['HRate'], i['Husk'], i['HuskAmt'], \
                                        i['DRate'], i['Dust'], i['DustAmt'], i['RRate'], i['Rejection'], i['RejectionAmt'],\
                                            i['GRate'],i['Gbags'], i['GbagAmt'],\
                                            i['Total']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totalHR'], j['totalH'], \
                                        j['totalHA'], j['totalDR'], j['totalD'],j['totalDA'], j['totalRR'], j['totalR'],\
                                           j['totalRA'], j['totalGR'],j['totalG'],j['totalGA'], j['total']), tags= 'total')

    def update(self):
        sl = int(self.sl_entry.get())
        hr = int(self.HRate_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        h = float(self.Husk_entry.get())
        ha = float(self.HuskAmt_entry.get())
        dr = float(self.DRate_entry.get())
        d = float(self.Dust_entry.get())
        da = float(self.DustAmt_entry.get())
        rr = float(self.RRate_entry.get())
        r = float(self.Rej_entry.get())
        ra = float(self.RejAmt_entry.get()) 
        gr = float(self.GRate_entry.get())
        g = float(self.Gbags_entry.get())
        ga = float(self.GbagAmt_entry.get())
        data = {    
                    "Sl": sl,
                    "Date": date,
                    "HRate": hr,
                    "Husk": h,
                    "HuskAmt": ha,
                    "DRate": dr,
                    "Dust": d,
                    "DustAmt": da,
                    "RRate": rr,
                    "Rejection": r,
                    "RejectionAmt": ra,
                    "GRate": gr,
                    "Gbags": g,
                    "GbagAmt": ga,
                    }
            
        self.collection.update_one({"Sl": sl}, {'$set':{    
                    "Sl": sl,
                    "Date": date,
                    "HRate": hr,
                    "Husk": h,
                    "HuskAmt": ha,
                    "DRate": dr,
                    "Dust": d,
                    "DustAmt": da,
                    "RRate": rr,
                    "Rejection": r,
                    "RejectionAmt": ra,
                    "GRate": gr,
                    "Gbags": g,
                    "GbagAmt": ga,
                    }})
        self.collection.update_one({'Sl': sl},
                            [{'$set': {'Total': {'$add': ['$HuskAmt', '$DustAmt','$RejectionAmt', '$GbagAmt']}}}])
        self.showall()  

    def delete(self):
        sl = int(self.sl_entry.get())
        self.collection.delete_one({'Sl':sl})
        self.reset()
        self.showall()

class chart(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Month', 'Sales', 'Payment', 'TCS_TDS', 'Purchase','Receipt')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Month", width=50, minwidth=25)
        self.tree.column("Sales", width=80, minwidth=25)
        self.tree.column("Payment", width=130, minwidth=25)
        self.tree.column("TCS_TDS", width=80, minwidth=25)
        self.tree.column("Purchase", width=80, minwidth=25)
        self.tree.column("Receipt", width=80, minwidth=25)

        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Month', text='Month')
        self.tree.heading('Sales', text='Sales')
        self.tree.heading('Payment', text='Payment')
        self.tree.heading('TCS_TDS', text='TCS/TDS')
        self.tree.heading('Purchase', text='Purchase')
        self.tree.heading('Receipt', text='Reciept')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=0, column=0, sticky=tk.W)

        self.Sales_label = ttk.Label(self.data_Entry, text="Sales:")
        self.Sales_label.grid(row=2, column=0, sticky=tk.W)

        self.Payment_label = ttk.Label(self.data_Entry, text="Payment:")
        self.Payment_label.grid(row=3, column=0, sticky=tk.W)

        self.tds_label = ttk.Label(self.data_Entry, text="TCS/TDS:")
        self.tds_label.grid(row=4, column=0, sticky=tk.W)

        self.Purchase_label = ttk.Label(self.data_Entry, text="Purchase:")
        self.Purchase_label.grid(row=5, column=0, sticky=tk.W)

        self.Reciept_label = ttk.Label(self.data_Entry, text="Receipt:")
        self.Reciept_label.grid(row=6, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.monthyear = MonthYearEntry(self.data_Entry)
        self.monthyear.grid(row=0, column=1)


        self.Sales_entry = ttk.Entry(self.data_Entry)
        self.Sales_entry.grid(row=2, column=1, sticky=tk.W)

        self.Payment_entry = ttk.Entry(self.data_Entry)
        self.Payment_entry.grid(row=3, column=1, sticky=tk.W)

        self.tcs_entry = ttk.Entry(self.data_Entry)
        self.tcs_entry.grid(row=4, column=1, sticky=tk.W)

        self.Purchase_entry = ttk.Entry(self.data_Entry)
        self.Purchase_entry.grid(row=5, column=1, sticky=tk.W)

        self.Receipt_entry = ttk.Entry(self.data_Entry)
        self.Receipt_entry.grid(row=6, column=1, sticky=tk.W)

        self.from_entry = MonthYearEntry(self.range_entry)
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = MonthYearEntry(self.range_entry)
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=11, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Month', 'Sales', 'Payment',
                                          'TCS_TDS','Purchase',
                                          'Receipt'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Date",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['Date'], i['Sales'], i['Payment'], i['TCS_TDS'], \
                                        i['Purchase'], i['Receipt']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalS": { "$sum": "$Sales" }, \
                                                        "totalP":{"$sum":"$Payment" }, "totalT": { "$sum": \
                                                            "$TCS_TSD" }, "totalPR":{"$sum":"$Purchase" }, \
                                                            "totalR":{"$sum":"$Receipt"},\
                                                            }}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0,  j['totalS'], j['totalP'], \
                                        j['totalT'], j['totalPR'], j['totalR']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            #self.monthyear.delete()
            self.Sales_entry.delete(0, tk.END)
            self.Payment_entry.delete(0, tk.END)
            self.tcs_entry.delete(0, tk.END)
            self.Purchase_entry.delete(0, tk.END)
            self.Receipt_entry.delete(0, tk.END)
        
            self.monthyear.insert(values[1])
            self.Sales_entry.insert(0, values[2])
            self.Payment_entry.insert(0, values[3])
            self.tcs_entry.insert(0, values[4])
            self.Purchase_entry.insert(0, values[5])
            self.Receipt_entry.insert(0, values[6])


    def add(self):
        try:
            datestr = self.monthyear.get()
            if len(datestr)<= 7:
                datestr = datestr+"-01 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            sales = float(self.Sales_entry.get())
            p = float(self.Payment_entry.get())
            t = float(self.tcs_entry.get())
            pr= float(self.Purchase_entry.get())
            r = float(self.Receipt_entry.get())
            data = {'Date':date, 'Sales':sales, 'Payment':p,'TCS_TDS':t,\
                    'Purchase':pr, 'Receipt':r}
            #data1 = {[{$addFields: {Total: { $add: ["$SGST", "$CGST", "$Amount"] }}}]}
            
            self.collection.insert_one(data)
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.Sales_entry.delete(0, tk.END)
        self.Payment_entry.delete(0, tk.END)
        self.tcs_entry.delete(0, tk.END)
        self.Purchase_entry.delete(0, tk.END)
        self.Receipt_entry.delete(0, tk.END)
        

        self.Sales_entry.insert(0, 0)
        self.Payment_entry.insert(0, 0)
        self.tcs_entry.insert(0, 0)
        self.Purchase_entry.insert(0, 0)
        self.Receipt_entry.insert(0, 0)

    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+"-01 00:00:00"
        todate = self.to_entry.get()
        todate = todate+"-01 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Date",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalS": { "$sum": "$Sales" }, \
                                                        "totalP":{"$sum":"$Payment" }, "totalT": { "$sum": \
                                                            "$TCS_TSD" }, "totalPR":{"$sum":"$Purchase" }, \
                                                            "totalR":{"$sum":"$Receipt"},\
                                                            }}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['Date'], i['Sales'], i['Payment'], i['TCS_TDS'], \
                                        i['Purchase'], i['Receipt']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0,  j['totalS'], j['totalP'], \
                                        j['totalT'], j['totalPR'], j['totalR']), tags= 'total')

    def update(self):
        datestr = self.monthyear.get()
        if len(datestr)<= 7:
            datestr = datestr+"-01 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        sales = float(self.Sales_entry.get())
        p = float(self.Payment_entry.get())
        t = float(self.tcs_entry.get())
        pr= float(self.Purchase_entry.get())
        r = float(self.Receipt_entry.get())
        data = {'Date':date, 'Sales':sales, 'Payment':p,'TCS_TDS':t,\
                'Purchase':pr, 'Receipt':r}
        self.collection.update_one({"Date": date}, {'$set':{'Date':date, 'Sales':sales, 'Payment':p,'TCS_TDS':t,\
                    'Purchase':pr, 'Receipt':r}}) 
        self.showall()  

    def delete(self):
        datestr = self.monthyear.get()
        if len(datestr)<= 7:
            datestr = datestr+"-01 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        self.collection.delete_one({'Date':date})
        self.reset()
        self.showall()

class ker(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'LotNo', 'TotalWholes', 'TotalPieces', 'TotalTotal', 'DispWholes', 'DispPieces', \
                    'DispRejection' ,'DispTotal', 'Party', 'PartyWholes','PartyPieces', 'PartyTotal')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("LotNo", width=80, minwidth=25)
        self.tree.column("TotalWholes", width=80, minwidth=25)
        self.tree.column("TotalPieces", width=80, minwidth=25)
        self.tree.column("TotalTotal", width=80, minwidth=25)
        self.tree.column("DispWholes", width=80, minwidth=25)
        self.tree.column("DispPieces", width=80, minwidth=25)
        self.tree.column("DispRejection", width=80, minwidth=25)
        self.tree.column("DispTotal", width=80, minwidth=25)
        self.tree.column("Party", width=80, minwidth=25)
        self.tree.column("PartyWholes", width=80, minwidth=25)
        self.tree.column("PartyPieces", width=80, minwidth=25)
        self.tree.column("PartyTotal", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('LotNo', text='LotNo')
        self.tree.heading('TotalWholes', text='TotalWholes')
        self.tree.heading('TotalPieces', text='TotalPieces')
        self.tree.heading('TotalTotal', text='TotalTotal')
        self.tree.heading('DispWholes', text='DispWholes')
        self.tree.heading('DispPieces', text='DispPieces')
        self.tree.heading('DispRejection', text='DispRejection')
        self.tree.heading('DispTotal', text='DispTotal')
        self.tree.heading('Party', text='Party')
        self.tree.heading('PartyWholes', text='PartyWholes')
        self.tree.heading('PartyPieces', text='PartyPieces')
        self.tree.heading('PartyTotal', text='PartyTotal')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        # self.range_entry = ttk.LabelFrame(parent, text="Search")
        # self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.Lot_label = ttk.Label(self.data_Entry, text="LotNo:")
        self.Lot_label.grid(row=0, column=0, sticky=tk.W)
        
        self.tw_label = ttk.Label(self.data_Entry, text="Total Wholes:")
        self.tw_label.grid(row=1, column=0, sticky=tk.W)

        self.tp_label = ttk.Label(self.data_Entry, text="Total Pieces:")
        self.tp_label.grid(row=2, column=0, sticky=tk.W)

        self.dw_label = ttk.Label(self.data_Entry, text="Dispatched Wholes:")
        self.dw_label.grid(row=3, column=0, sticky=tk.W)

        self.dp_label = ttk.Label(self.data_Entry, text="Dispatched Pieces:")
        self.dp_label.grid(row=4, column=0, sticky=tk.W)

        self.dr_label = ttk.Label(self.data_Entry, text="Dispatched Rejection:")
        self.dr_label.grid(row=5, column=0, sticky=tk.W)

        self.p_label = ttk.Label(self.data_Entry, text="Party:")
        self.p_label.grid(row=6, column=0, sticky=tk.W)

        self.pw_label = ttk.Label(self.data_Entry, text="Party Wholes:")
        self.pw_label.grid(row=7, column=0, sticky=tk.W)

        self.pp_label = ttk.Label(self.data_Entry, text="Party Pieces:")
        self.pp_label.grid(row=8, column=0, sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.lot_entry = ttk.Entry(self.data_Entry)
        self.lot_entry.grid(row=0, column=1, sticky=tk.W)

        self.tw_entry = ttk.Entry(self.data_Entry)
        self.tw_entry.grid(row=1, column=1, sticky=tk.W)

        self.tp_entry = ttk.Entry(self.data_Entry)
        self.tp_entry.grid(row=2, column=1, sticky=tk.W)

        self.dw_entry = ttk.Entry(self.data_Entry)
        self.dw_entry.grid(row=3, column=1, sticky=tk.W)

        self.dp_entry = ttk.Entry(self.data_Entry)
        self.dp_entry.grid(row=4, column=1, sticky=tk.W)

        self.dr_entry = ttk.Entry(self.data_Entry)
        self.dr_entry.grid(row=5, column=1, sticky=tk.W)

        self.p_entry = ttk.Entry(self.data_Entry)
        self.p_entry.grid(row=6, column=1, sticky=tk.W)
        
        self.pw_entry = ttk.Entry(self.data_Entry)
        self.pw_entry.grid(row=7, column=1, sticky=tk.W)

        self.pp_entry = ttk.Entry(self.data_Entry)
        self.pp_entry.grid(row=8, column=1, sticky=tk.W)

        

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=11, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'LotNo', 'TotalWholes', 
                                         'TotalPieces', 'TotalTotal', 'DispWholes', 
                                         'DispPieces', 
                                        'DispRejection' ,'DispTotal', 'Party', 'PartyWholes',
                                        'PartyPieces', 'PartyTotal'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("LotNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['LotNo'], i['TotalWholes'], i['TotalPieces'], i['TotalTotal'], \
                                        i['DispWholes'], i['DispPieces'], i['DispRejection'], i['DispTotal'],\
                                              i['Party'],i['PartyWholes'], i['PartyPieces'],\
                                            i['PartyTotal']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totaltw": { "$sum": "$TotalWholes" }, \
                                        "totaltp":{"$sum":"$TotalPieces" },\
                                        "totaltt":{ "$sum":"$TotalTotal" }, \
                                        "totaldw":{"$sum":"$DispWholes" }, \
                                        "totaldp":{"$sum":"$DispPieces"},\
                                        "totaldr":{ "$sum": "$DispRejection" },\
                                        "totaldt":{"$sum":"$DispTotal" }, \
                                        "totalpw":{"$sum":"$PiecesWholes"},\
                                        "totalpp":{"$sum":"$PartyPieces" }, \
                                        "totalpt":{"$sum":"$PartyTotal"}}}]
        
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totaltw'], j['totaltp'], \
                                        j['totaltt'], j['totaldw'], j['totaldp'],j['totaldr'], j['totaldt'],0, j['totalpw'],\
                                           j['totalpp'], j['totalpt']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.lot_entry.delete(0, tk.END)
            self.tw_entry.delete(0, tk.END)
            self.tp_entry.delete(0, tk.END)
            self.dw_entry.delete(0, tk.END)
            self.dp_entry.delete(0, tk.END)
            self.dr_entry.delete(0, tk.END)
            self.p_entry.delete(0, tk.END)
            self.pw_entry.delete(0, tk.END)
            self.pp_entry.delete(0, tk.END)
            
            self.lot_entry.insert(0, values[1])
            self.tw_entry.insert(0, values[2])
            self.tp_entry.insert(0, values[3])
            self.dw_entry.insert(0, values[5])
            self.dp_entry.insert(0, values[6])
            self.dr_entry.insert(0, values[7])
            self.p_entry.insert(0, values[9])
            self.pw_entry.insert(0, values[10])
            self.pp_entry.insert(0, values[11])

    def add(self):
        try:
            lot = int(self.lot_entry.get())
            tw = float(self.tw_entry.get())
            tp = float(self.tp_entry.get())
            dw = float(self.dw_entry.get())
            dp = float(self.dp_entry.get())
            dr = float(self.dr_entry.get())
            p = self.p_entry.get()
            pw = float(self.pw_entry.get())
            pp = float(self.pp_entry.get())

            data = {'LotNo': lot, 'TotalWholes':tw, 'TotalPieces':tp, 'DispWholes':dw,'DispPieces':dp,\
                    'DispRejection':dr, 'Party':p\
                    , 'PartyWholes':pw, "PartyPieces":pp}
            
            self.collection.insert_one(data)
            self.collection.update_one({'LotNo': lot},
                                [{'$addFields': {'TotalTotal': {'$add': ['$TotalWholes', '$TotalPieces']}}}])
            self.collection.update_one({'LotNo': lot},
                                [{'$addFields': {'DispTotal': {'$add': ['$DispWholes', '$DispPieces', '$DispRejection']}}}])
            self.collection.update_one({'LotNo': lot},
                                [{'$addFields': {'PartyTotal': {'$add': ['$PartyWholes', '$PartyPieces']}}}])
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.lot_entry.delete(0, tk.END)
        self.tw_entry.delete(0, tk.END)
        self.tp_entry.delete(0, tk.END)
        self.dw_entry.delete(0, tk.END)
        self.dp_entry.delete(0, tk.END)
        self.dr_entry.delete(0, tk.END)
        self.p_entry.delete(0, tk.END)
        self.pw_entry.delete(0, tk.END)
        self.pp_entry.delete(0, tk.END)
            
        self.tw_entry.insert(0, 0)
        self.tp_entry.insert(0, 0)
        self.dw_entry.insert(0, 0)
        self.dp_entry.insert(0, 0)
        self.dr_entry.insert(0, 0)
        self.pw_entry.insert(0, 0)
        self.pp_entry.insert(0, 0)


    def update(self):
        lot = int(self.lot_entry.get())
        tw = float(self.tw_entry.get())
        tp = float(self.tp_entry.get())
        dw = float(self.dw_entry.get())
        dp = float(self.dp_entry.get())
        dr = float(self.dr_entry.get())
        p = self.p_entry.get()
        pw = float(self.pw_entry.get())
        pp = float(self.pp_entry.get())
        self.collection.update_one({"LotNo": lot}, {'$set':{'LotNo': lot, 'TotalWholes':tw, 'TotalPieces':tp, 'DispWholes':dw,'DispPieces':dp,\
                    'DispRejection':dr, 'Party':p\
                    , 'PartyWholes':pw, "PartyPieces":pp}}) 
        self.collection.update_one({'LotNo': lot},
                                [{'$addFields': {'TotalTotal': {'$add': ['$TotalWholes', '$TotalPieces']}}}])
        self.collection.update_one({'LotNo': lot},
                                [{'$addFields': {'DispTotal': {'$add': ['$DispWholes', '$DispPieces', '$DispRejection']}}}])
        self.collection.update_one({'LotNo': lot},
                                [{'$addFields': {'PartyTotal': {'$add': ['$PartyWholes', '$PartyPieces']}}}])
        self.showall()  

    def delete(self):
        lot = int(self.lot_entry.get())
        self.collection.delete_one({'LotNo':lot})
        self.reset()
        self.showall()

class ss(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'BillNo', 'Date', 'Kernels', 'Shells', 'RCN', \
                    'Husk' ,'Dust')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("BillNo", width=80, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Kernels", width=80, minwidth=25)
        self.tree.column("Shells", width=80, minwidth=25)
        self.tree.column("RCN", width=80, minwidth=25)
        self.tree.column("Husk", width=80, minwidth=25)
        self.tree.column("Dust", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('BillNo', text='BillNo')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Kernels', text='Kernels')
        self.tree.heading('Shells', text='Shells')
        self.tree.heading('RCN', text='RCN')
        self.tree.heading('Husk', text='Husk')
        self.tree.heading('Dust', text='Dust')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.bill_label = ttk.Label(self.data_Entry, text="Bill NO:")
        self.bill_label.grid(row=0, column=0, sticky=tk.W)
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.ker_label = ttk.Label(self.data_Entry, text="Kernels:")
        self.ker_label.grid(row=2, column=0, sticky=tk.W)

        self.sh_label = ttk.Label(self.data_Entry, text="Shells:")
        self.sh_label.grid(row=3, column=0, sticky=tk.W)

        self.rcn_label = ttk.Label(self.data_Entry, text="R.C.Nuts:")
        self.rcn_label.grid(row=4, column=0, sticky=tk.W)

        self.husk_label = ttk.Label(self.data_Entry, text="Husk:")
        self.husk_label.grid(row=5, column=0, sticky=tk.W)

        self.dust_label = ttk.Label(self.data_Entry, text="Dust:")
        self.dust_label.grid(row=6, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.bill_entry = ttk.Entry(self.data_Entry)
        self.bill_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.ker_entry = ttk.Entry(self.data_Entry)
        self.ker_entry.grid(row=2, column=1, sticky=tk.W)

        self.sh_entry = ttk.Entry(self.data_Entry)
        self.sh_entry.grid(row=3, column=1, sticky=tk.W)

        self.rcn_entry = ttk.Entry(self.data_Entry)
        self.rcn_entry.grid(row=4, column=1, sticky=tk.W)

        self.husk_entry = ttk.Entry(self.data_Entry)
        self.husk_entry.grid(row=5, column=1, sticky=tk.W)

        self.dust_entry = ttk.Entry(self.data_Entry)
        self.dust_entry.grid(row=6, column=1, sticky=tk.W)


        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'BillNo', 'Date', 'Kernels', 'Shells', 'RCN',
                    'Husk' ,'Dust'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("BillNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Kernels'], i['Shells'], \
                                        i['RCN'], i['Husk'], i['Dust']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalk": { "$sum": "$Kernels" }, \
                                                        "totals":{"$sum":"$Shells" }, "totalr": { "$sum": \
                                                            "$RCN" }, "totalh":{"$sum":"$Husk" }, \
                                                            "totald":{"$sum":"$Dust"}}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalk'], j['totals'], \
                                        j['totalr'], j['totalh'], j['totald']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.Date_entry.entry.delete(0, tk.END)
            self.bill_entry.delete(0, tk.END)
            self.husk_entry.delete(0, tk.END)
            self.dust_entry.delete(0, tk.END)
            self.ker_entry.delete(0, tk.END)
            self.sh_entry.delete(0, tk.END)
            self.rcn_entry.delete(0, tk.END)
            
            
            self.bill_entry.insert(0, values[1])
            self.Date_entry.entry.insert(0, values[2])
            self.ker_entry.insert(0, values[3])
            self.sh_entry.insert(0, values[4])
            self.rcn_entry.insert(0, values[5])
            self.husk_entry.insert(0, values[6])
            self.dust_entry.insert(0, values[7])

    def add(self):
        try:
            bill = int(self.bill_entry.get())
            k = int(self.ker_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            h = float(self.husk_entry.get())
            d = float(self.dust_entry.get())
            r = float(self.rcn_entry.get())
            s = float(self.sh_entry.get()) 
            
            data = {    
                        "BillNo": bill,
                        "Date": date,
                        "Kernels": k,
                        "Shells": s,
                        "RCN": r,
                        "Husk": h,
                        "Dust": d
                    }
            
            self.collection.insert_one(data)
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.bill_entry.delete(0, tk.END)
        self.husk_entry.delete(0, tk.END)
        self.dust_entry.delete(0, tk.END)
        self.rcn_entry.delete(0, tk.END)
        self.ker_entry.delete(0, tk.END)
        self.sh_entry.delete(0, tk.END)
            
            
        self.bill_entry.insert(0, 0)
        self.ker_entry.insert(0, 0)
        self.sh_entry.insert(0, 0)
        self.rcn_entry.insert(0, 0)
        self.husk_entry.insert(0, 0)
        self.dust_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Sl",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalk": { "$sum": "$Kernels" }, \
                                                        "totals":{"$sum":"$Shells" }, "totalr": { "$sum": \
                                                            "$RCN" }, "totalh":{"$sum":"$Husk" }, \
                                                            "totald":{"$sum":"$Dust"}}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Kernels'], i['Shells'], \
                                        i['RCN'], i['Husk'], i['Dust']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalk'], j['totals'], \
                                        j['totalr'], j['totalh'], j['totald']), tags= 'total')

    def update(self):
        bill = int(self.bill_entry.get())
        k = int(self.ker_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        h = float(self.husk_entry.get())
        d = float(self.dust_entry.get())
        r = float(self.rcn_entry.get())
        s = float(self.sh_entry.get()) 
            
        self.collection.update_one({"BillNo": bill}, {'$set':{    
                        "BillNo": bill,
                        "Date": date,
                        "Kernels": k,
                        "Shells": s,
                        "RCN": r,
                        "Husk": h,
                        "Dust": d
                    }})
        self.showall()  

    def delete(self):
        bill = int(self.bill_entry.get())
        self.collection.delete_one({'BillNo':bill})
        self.reset()
        self.showall()

class ps(tk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'BillNo', 'Date', 'Kernels', 'Shells', 'RCN', \
                    'Husk' ,'Dust')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("BillNo", width=80, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Kernels", width=80, minwidth=25)
        self.tree.column("Shells", width=80, minwidth=25)
        self.tree.column("RCN", width=80, minwidth=25)
        self.tree.column("Husk", width=80, minwidth=25)
        self.tree.column("Dust", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('BillNo', text='BillNo')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Kernels', text='Kernels')
        self.tree.heading('Shells', text='Shells')
        self.tree.heading('RCN', text='RCN')
        self.tree.heading('Husk', text='Husk')
        self.tree.heading('Dust', text='Dust')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.bill_label = ttk.Label(self.data_Entry, text="Bill NO:")
        self.bill_label.grid(row=0, column=0, sticky=tk.W)
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.ker_label = ttk.Label(self.data_Entry, text="Kernels:")
        self.ker_label.grid(row=2, column=0, sticky=tk.W)

        self.sh_label = ttk.Label(self.data_Entry, text="Shells:")
        self.sh_label.grid(row=3, column=0, sticky=tk.W)

        self.rcn_label = ttk.Label(self.data_Entry, text="R.C.Nuts:")
        self.rcn_label.grid(row=4, column=0, sticky=tk.W)

        self.husk_label = ttk.Label(self.data_Entry, text="Husk:")
        self.husk_label.grid(row=5, column=0, sticky=tk.W)

        self.dust_label = ttk.Label(self.data_Entry, text="Dust:")
        self.dust_label.grid(row=6, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.bill_entry = ttk.Entry(self.data_Entry)
        self.bill_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.ker_entry = ttk.Entry(self.data_Entry)
        self.ker_entry.grid(row=2, column=1, sticky=tk.W)

        self.sh_entry = ttk.Entry(self.data_Entry)
        self.sh_entry.grid(row=3, column=1, sticky=tk.W)

        self.rcn_entry = ttk.Entry(self.data_Entry)
        self.rcn_entry.grid(row=4, column=1, sticky=tk.W)

        self.husk_entry = ttk.Entry(self.data_Entry)
        self.husk_entry.grid(row=5, column=1, sticky=tk.W)

        self.dust_entry = ttk.Entry(self.data_Entry)
        self.dust_entry.grid(row=6, column=1, sticky=tk.W)


        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'BillNo', 'Date', 'Kernels', 'Shells', 'RCN', \
                    'Husk' ,'Dust'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("BillNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Kernels'], i['Shells'], \
                                        i['RCN'], i['Husk'], i['Dust']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalk": { "$sum": "$Kernels" }, \
                                                        "totals":{"$sum":"$Shells" }, "totalr": { "$sum": \
                                                            "$RCN" }, "totalh":{"$sum":"$Husk" }, \
                                                            "totald":{"$sum":"$Dust"}}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalk'], j['totals'], \
                                        j['totalr'], j['totalh'], j['totald']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.Date_entry.entry.delete(0, tk.END)
            self.bill_entry.delete(0, tk.END)
            self.husk_entry.delete(0, tk.END)
            self.dust_entry.delete(0, tk.END)
            self.ker_entry.delete(0, tk.END)
            self.sh_entry.delete(0, tk.END)
            self.rcn_entry.delete(0, tk.END)
            
            
            self.bill_entry.insert(0, values[1])
            self.Date_entry.entry.insert(0, values[2])
            self.ker_entry.insert(0, values[3])
            self.sh_entry.insert(0, values[4])
            self.rcn_entry.insert(0, values[5])
            self.husk_entry.insert(0, values[6])
            self.dust_entry.insert(0, values[7])

    def add(self):
        try:
            bill = int(self.bill_entry.get())
            k = int(self.ker_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            h = float(self.husk_entry.get())
            d = float(self.dust_entry.get())
            r = float(self.rcn_entry.get())
            s = float(self.sh_entry.get()) 
            
            data = {    
                        "BillNo": bill,
                        "Date": date,
                        "Kernels": k,
                        "Shells": s,
                        "RCN": r,
                        "Husk": h,
                        "Dust": d
                    }
            
            self.collection.insert_one(data)
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.bill_entry.delete(0, tk.END)
        self.husk_entry.delete(0, tk.END)
        self.dust_entry.delete(0, tk.END)
        self.ker_entry.delete(0, tk.END)
        self.sh_entry.delete(0, tk.END)
        self.rcn_entry.delete(0, tk.END)
            
            
        self.bill_entry.insert(0, 0)
        self.ker_entry.insert(0, 0)
        self.sh_entry.insert(0, 0)
        self.rcn_entry.insert(0, 0)
        self.husk_entry.insert(0, 0)
        self.dust_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Sl",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalk": { "$sum": "$Kernels" }, \
                                                        "totals":{"$sum":"$Shells" }, "totalr": { "$sum": \
                                                            "$RCN" }, "totalh":{"$sum":"$Husk" }, \
                                                            "totald":{"$sum":"$Dust"}}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['BillNo'], i['Date'], i['Kernels'], i['Shells'], \
                                        i['RCN'], i['Husk'], i['Dust']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalk'], j['totals'], \
                                        j['totalr'], j['totalh'], j['totald']), tags= 'total')

    def update(self):
        bill = int(self.bill_entry.get())
        k = int(self.ker_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        h = float(self.husk_entry.get())
        d = float(self.dust_entry.get())
        r = float(self.rcn_entry.get())
        s = float(self.sh_entry.get()) 
            
        self.collection.update_one({"BillNo": bill}, {'$set':{    
                        "BillNo": bill,
                        "Date": date,
                        "Kernels": k,
                        "Shells": s,
                        "RCN": r,
                        "Husk": h,
                        "Dust": d
                    }})
        self.showall()  

    def delete(self):
        bill = int(self.bill_entry.get())
        self.collection.delete_one({'BillNo':bill})
        self.reset()
        self.showall()

class pa(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Lot', 'Date', 'RCNroast')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Lot", width=80, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("RCNroast", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Lot', text='Lot')
        self.tree.heading('Date', text='Date')
        self.tree.heading('RCNroast', text='RCN Roast')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.bill_label = ttk.Label(self.data_Entry, text="Lot NO:")
        self.bill_label.grid(row=0, column=0, sticky=tk.W)
        
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.ker_label = ttk.Label(self.data_Entry, text="RCN Roast:")
        self.ker_label.grid(row=2, column=0, sticky=tk.W)


        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 2, column=1, sticky=tk.W)

        #entry
        self.lot_entry = ttk.Entry(self.data_Entry)
        self.lot_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.ros_entry = ttk.Entry(self.data_Entry)
        self.ros_entry.grid(row=2, column=1, sticky=tk.W)


        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Lot', 'Date', 'RCNroast'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("BillNo",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['Lot'], i['Date'], i['RCNroast']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalr": { "$sum": "$RCNroast" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalr']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.Date_entry.entry.delete(0, tk.END)
            self.lot_entry.delete(0, tk.END)
            self.ros_entry.delete(0, tk.END)
            
            self.lot_entry.insert(0, values[1])
            self.Date_entry.entry.insert(0, values[2])
            self.ros_entry.insert(0, values[3])

    def add(self):
        try:
            lot = int(self.lot_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            r = float(self.ros_entry.get())
            
            data = {    
                        "Lot": lot,
                        "Date": date,
                        "RCNroast": r
                    }
            
            self.collection.insert_one(data)
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.lot_entry.delete(0, tk.END)
        self.ros_entry.delete(0, tk.END)
            
        self.lot_entry.insert(0, 0)
        self.ros_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Lot",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalr": { "$sum": "$RCNroast" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter,  i['Lot'], i['Date'], i['RCNroast']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalr']), tags= 'total')

    def update(self):
        lot = int(self.lot_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        r = float(self.ros_entry.get())
            
        self.collection.update_one({"Lot": lot}, {'$set':{    
                        "Lot": lot,
                        "Date": date,
                        "RCNroast": r
                    }})
        self.showall()  

    def delete(self):
        lot = int(self.bill_entry.get())
        self.collection.delete_one({'Lot':lot})
        self.reset()
        self.showall()

class srcn(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Date', 'OSBStock', 'Purchase', 'Sales', \
                    'Production' ,'CSStock')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("OSBStock", width=80, minwidth=25)
        self.tree.column("Purchase", width=80, minwidth=25)
        self.tree.column("Sales", width=80, minwidth=25)
        self.tree.column("Production", width=80, minwidth=25)
        self.tree.column("CSStock", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('OSBStock', text='OSBStock')
        self.tree.heading('Purchase', text='Purchase')
        self.tree.heading('Sales', text='Sales')
        self.tree.heading('Production', text='Production')
        self.tree.heading('CSStock', text='CSStock')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.osb_label = ttk.Label(self.data_Entry, text="OSBStock:")
        self.osb_label.grid(row=2, column=0, sticky=tk.W)

        self.pu_label = ttk.Label(self.data_Entry, text="Purchase:")
        self.pu_label.grid(row=3, column=0, sticky=tk.W)

        self.s_label = ttk.Label(self.data_Entry, text="Sales:")
        self.s_label.grid(row=4, column=0, sticky=tk.W)

        self.pr_label = ttk.Label(self.data_Entry, text="Production:")
        self.pr_label.grid(row=5, column=0, sticky=tk.W)

        self.css_label = ttk.Label(self.data_Entry, text="CSStock:")
        self.css_label.grid(row=6, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.osb_entry = ttk.Entry(self.data_Entry)
        self.osb_entry.grid(row=2, column=1, sticky=tk.W)

        self.pu_entry = ttk.Entry(self.data_Entry)
        self.pu_entry.grid(row=3, column=1, sticky=tk.W)

        self.s_entry = ttk.Entry(self.data_Entry)
        self.s_entry.grid(row=4, column=1, sticky=tk.W)

        self.pr_entry = ttk.Entry(self.data_Entry)
        self.pr_entry.grid(row=5, column=1, sticky=tk.W)

        self.css_entry = ttk.Entry(self.data_Entry)
        self.css_entry.grid(row=6, column=1, sticky=tk.W)


        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Date', 'OSBStock', 'Purchase', 'Sales', \
                    'Production' ,'CSStock'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Month",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter, i['Month'], i['OSBStock'], i['Purchase'], \
                                        i['Sales'], i['Production'], i['CSStock']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalo": { "$sum": "$OSBStock" }, \
                                                        "totalpu":{"$sum":"$Purchase" }, "totals": { "$sum": \
                                                            "$Sales" }, "totalpr":{"$sum":"$Production" }, \
                                                            "totalc":{"$sum":"$CSStock"}}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totalo'], j['totalpu'], \
                                        j['totals'], j['totalpr'], j['totalc']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.Date_entry.entry.delete(0, tk.END)
            self.osb_entry.delete(0, tk.END)
            self.pu_entry.delete(0, tk.END)
            self.s_entry.delete(0, tk.END)
            self.pr_entry.delete(0, tk.END)
            self.css_entry.delete(0, tk.END)
            
            self.Date_entry.entry.insert(0, values[1])
            self.osb_entry.insert(0, values[2])
            self.pu_entry.insert(0, values[3])
            self.s_entry.insert(0, values[4])
            self.pr_entry.insert(0, values[5])
            self.css_entry.insert(0, values[6])

    def add(self):
        try:
            o = int(self.osb_entry.get())
            pu = int(self.pu_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            s = float(self.s_entry.get())
            pr = float(self.pr_entry.get())
            c = float(self.css_entry.get())
            
            data = {    
                        "Month": date,
                        "OSBStock": o,
                        "Purchase": pu,
                        "Sales": s,
                        "Production": pr,
                        "CSStock": c
                    }
            
            self.collection.insert_one(data)
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.osb_entry.delete(0, tk.END)
        self.pu_entry.delete(0, tk.END)
        self.s_entry.delete(0, tk.END)
        self.pr_entry.delete(0, tk.END)
        self.css_entry.delete(0, tk.END)
            
        self.osb_entry.insert(0, 0)
        self.pu_entry.insert(0, 0)
        self.s_entry.insert(0, 0)
        self.pr_entry.insert(0, 0)
        self.css_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Month":{'$gte':fd, '$lte':td}}).sort("Month",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Month":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalo": { "$sum": "$OSBStock" }, \
                                                        "totalpu":{"$sum":"$Purchase" }, "totals": { "$sum": \
                                                            "$Sales" }, "totalpr":{"$sum":"$Production" }, \
                                                            "totalc":{"$sum":"$CSStock"}}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter, i['Month'], i['OSBStock'], i['Purchase'], \
                                        i['Sales'], i['Production'], i['CSStock']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totalo'], j['totalpu'], \
                                        j['totals'], j['totalpr'], j['totalc']), tags= 'total')

    def update(self):
        o = int(self.osb_entry.get())
        pu = int(self.pu_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        s = float(self.s_entry.get())
        pr = float(self.pr_entry.get())
        c = float(self.css_entry.get())
            
        self.collection.update_one({"Month": date}, {'$set':{    
                        "Month": date,
                        "OSBStock": o,
                        "Purchase": pu,
                        "Sales": s,
                        "Production": pr,
                        "CSStock": c
                    }})
        self.showall()  

    def delete(self):
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        self.collection.delete_one({'Month':date})
        self.reset()
        self.showall()

class sckn(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Date', 'OSBStock', 'Purchase', 'Sales', \
                    'Production' ,'CSStock')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("OSBStock", width=80, minwidth=25)
        self.tree.column("Purchase", width=80, minwidth=25)
        self.tree.column("Sales", width=80, minwidth=25)
        self.tree.column("Production", width=80, minwidth=25)
        self.tree.column("CSStock", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('OSBStock', text='OSBStock')
        self.tree.heading('Purchase', text='Purchase')
        self.tree.heading('Sales', text='Sales')
        self.tree.heading('Production', text='Production')
        self.tree.heading('CSStock', text='CSStock')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)

        self.osb_label = ttk.Label(self.data_Entry, text="OSBStock:")
        self.osb_label.grid(row=2, column=0, sticky=tk.W)

        self.pu_label = ttk.Label(self.data_Entry, text="Purchase:")
        self.pu_label.grid(row=3, column=0, sticky=tk.W)

        self.s_label = ttk.Label(self.data_Entry, text="Sales:")
        self.s_label.grid(row=4, column=0, sticky=tk.W)

        self.pr_label = ttk.Label(self.data_Entry, text="Production:")
        self.pr_label.grid(row=5, column=0, sticky=tk.W)

        self.css_label = ttk.Label(self.data_Entry, text="CSStock:")
        self.css_label.grid(row=6, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)

        self.osb_entry = ttk.Entry(self.data_Entry)
        self.osb_entry.grid(row=2, column=1, sticky=tk.W)

        self.pu_entry = ttk.Entry(self.data_Entry)
        self.pu_entry.grid(row=3, column=1, sticky=tk.W)

        self.s_entry = ttk.Entry(self.data_Entry)
        self.s_entry.grid(row=4, column=1, sticky=tk.W)

        self.pr_entry = ttk.Entry(self.data_Entry)
        self.pr_entry.grid(row=5, column=1, sticky=tk.W)

        self.css_entry = ttk.Entry(self.data_Entry)
        self.css_entry.grid(row=6, column=1, sticky=tk.W)


        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)
        
        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.W, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Date', 'OSBStock', 'Purchase', 'Sales', \
                    'Production' ,'CSStock'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Month",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter, i['Month'], i['OSBStock'], i['Purchase'], \
                                        i['Sales'], i['Production'], i['CSStock']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalo": { "$sum": "$OSBStock" }, \
                                                        "totalpu":{"$sum":"$Purchase" }, "totals": { "$sum": \
                                                            "$Sales" }, "totalpr":{"$sum":"$Production" }, \
                                                            "totalc":{"$sum":"$CSStock"}}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totalo'], j['totalpu'], \
                                        j['totals'], j['totalpr'], j['totalc']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.Date_entry.entry.delete(0, tk.END)
            self.osb_entry.delete(0, tk.END)
            self.pu_entry.delete(0, tk.END)
            self.s_entry.delete(0, tk.END)
            self.pr_entry.delete(0, tk.END)
            self.css_entry.delete(0, tk.END)
            
            self.Date_entry.entry.insert(0, values[1])
            self.osb_entry.insert(0, values[2])
            self.pu_entry.insert(0, values[3])
            self.s_entry.insert(0, values[4])
            self.pr_entry.insert(0, values[5])
            self.css_entry.insert(0, values[6])

    def add(self):
        try:
            o = int(self.osb_entry.get())
            pu = int(self.pu_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            s = float(self.s_entry.get())
            pr = float(self.pr_entry.get())
            c = float(self.css_entry.get())
            
            data = {    
                        "Month": date,
                        "OSBStock": o,
                        "Purchase": pu,
                        "Sales": s,
                        "Production": pr,
                        "CSStock": c
                    }
            
            self.collection.insert_one(data)
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.osb_entry.delete(0, tk.END)
        self.pu_entry.delete(0, tk.END)
        self.s_entry.delete(0, tk.END)
        self.pr_entry.delete(0, tk.END)
        self.css_entry.delete(0, tk.END)
            
        self.osb_entry.insert(0, 0)
        self.pu_entry.insert(0, 0)
        self.s_entry.insert(0, 0)
        self.pr_entry.insert(0, 0)
        self.css_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Month":{'$gte':fd, '$lte':td}}).sort("Month",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Month":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalo": { "$sum": "$OSBStock" }, \
                                                        "totalpu":{"$sum":"$Purchase" }, "totals": { "$sum": \
                                                            "$Sales" }, "totalpr":{"$sum":"$Production" }, \
                                                            "totalc":{"$sum":"$CSStock"}}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter, i['Month'], i['OSBStock'], i['Purchase'], \
                                        i['Sales'], i['Production'], i['CSStock']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, j['totalo'], j['totalpu'], \
                                        j['totals'], j['totalpr'], j['totalc']), tags= 'total')

    def update(self):
        o = int(self.osb_entry.get())
        pu = int(self.pu_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        s = float(self.s_entry.get())
        pr = float(self.pr_entry.get())
        c = float(self.css_entry.get())
            
        self.collection.update_one({"Month": date}, {'$set':{    
                        "Month": date,
                        "OSBStock": o,
                        "Purchase": pu,
                        "Sales": s,
                        "Production": pr,
                        "CSStock": c
                    }})
        self.showall()  

    def delete(self):
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        self.collection.delete_one({'Month':date})
        self.reset()
        self.showall()

class psd(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Id', 'Date','Party', 'INVNo', 'Qty', 'QtyType',\
                                'Amount', 'SGST', 'CGST','IGST', 'Total')
        self.tree.column("Id", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Party", width=80, minwidth=25)
        self.tree.column("INVNo", width=80, minwidth=25)
        self.tree.column("Qty", width=80, minwidth=25)
        self.tree.column("QtyType", width=80, minwidth=25)
        self.tree.column("Amount", width=80, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Id', text='Id', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('Party', text='Party')
        self.tree.heading('INVNo', text='INV.NO')
        self.tree.heading('Qty', text='QTY')
        self.tree.heading('QtyType', text='QTY Type')
        self.tree.heading('Amount', text='Amount')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.id_label = ttk.Label(self.data_Entry, text="Id:")
        self.id_label.grid(row=0, column=0, sticky=tk.W)

        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)
        
        self.p_label = ttk.Label(self.data_Entry, text="Party:")
        self.p_label.grid(row=2, column=0, sticky=tk.W)

        self.inv_label = ttk.Label(self.data_Entry, text="INV.NO:")
        self.inv_label.grid(row=3, column=0, sticky=tk.W)

        self.Qty_label = ttk.Label(self.data_Entry, text="Quantity:")
        self.Qty_label.grid(row=4, column=0, sticky=tk.W)

        self.qtyt_label = ttk.Label(self.data_Entry, text="Qty Type:")
        self.qtyt_label.grid(row=5, column=0, sticky=tk.W)

        self.amount_label = ttk.Label(self.data_Entry, text = "Amount:")
        self.amount_label.grid(row=6,column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=7, column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=8, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=9, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.id_entry = ttk.Entry(self.data_Entry)
        self.id_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)
        
        self.p_entry = ttk.Entry(self.data_Entry)
        self.p_entry.grid(row=2, column=1, sticky=tk.W)

        self.inv_entry = ttk.Entry(self.data_Entry)
        self.inv_entry.grid(row=3, column=1, sticky=tk.W)

        self.Qty_entry = ttk.Entry(self.data_Entry)
        self.Qty_entry.grid(row=4, column=1, sticky=tk.W)

        self.qtyt_entry = ttk.Entry(self.data_Entry)
        self.qtyt_entry.grid(row=5, column=1, sticky=tk.W)

        self.amount_entry = ttk.Entry(self.data_Entry)
        self.amount_entry.grid(row=6, column=1, sticky=tk.W)
        
        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=7, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=9, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=11, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['Id', 'Date','Party', 'INVNo', 'Qty', 'QtyType',\
                                'Amount', 'SGST', 'CGST','IGST', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Date",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], i['Date'], i['INVNo'], i['Party'], i['Qty'],\
                                    i['QtyType'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                    i['Total']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalAmt": { "$sum": "$Amount" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, j['totalQty'],
                                        0,j['totalAmt'], j['totalSGST'], j['totalCGST'],
                                        j['totalIGST'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.id_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.p_entry.delete(0, tk.END)
            self.inv_entry.delete(0, tk.END)
            self.Qty_entry.delete(0, tk.END)
            self.qtyt_entry.delete(0, tk.END)
            self.amount_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            
            self.id_entry.insert(0, values[0])
            self.Date_entry.entry.insert(0, values[1])
            self.p_entry.insert(0, values[2])
            self.inv_entry.insert(0, values[3])
            self.Qty_entry.insert(0, values[4])
            self.qtyt_entry.insert(0, values[5])
            self.amount_entry.insert(0, values[6])
            self.SGST_entry.insert(0, values[7])
            self.CGST_entry.insert(0, values[8])
            self.IGST_entry.insert(0, values[9])


    def add(self):
        try:
            id = int(self.id_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            i = self.inv_entry.get()
            qty = float(self.Qty_entry.get())
            qtyt = self.qtyt_entry.get()
            p = self.p_entry.get()
            amt = float(self.amount_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            data = {'Id':id, 'Date':date, 'Party': p, 'INVNo':i,'Qty':qty, 'QtyType':qtyt,
                    'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }
            
            self.collection.insert_one(data)
            self.collection.update_one({'Id': id},
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST', '$Amount']}}}])
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.id_entry.delete(0, tk.END)
        self.p_entry.delete(0, tk.END)
        self.inv_entry.delete(0, tk.END)
        self.Qty_entry.delete(0, tk.END)
        self.qtyt_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)
            
        self.p_entry.insert(0, 0)
        self.inv_entry.insert(0, 0)
        self.Qty_entry.insert(0, 0)
        self.qtyt_entry.insert(0, 0)
        self.amount_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Date",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalAmt": { "$sum": "$Amount" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], i['Date'], i['INVNo'], i['Party'], i['Qty'],\
                                    i['QtyType'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                    i['Total']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, j['totalQty'],
                                        0,j['totalAmt'], j['totalSGST'], j['totalCGST'],
                                        j['totalIGST'], j['total']), tags= 'total')

    def update(self):
        id = int(self.id_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        i = self.inv_entry.get()
        qty = float(self.Qty_entry.get())
        qtyt = self.qtyt_entry.get()
        p = self.p_entry.get()
        amt = float(self.amount_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get()) 
        igst = float(self.IGST_entry.get())
        self.collection.update_one({"Id": id}, {'$set':{'Date':date, 'Party': p, 'INVNo':i,'Qty':qty, 'QtyType':qtyt,
                    'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }}) 
        self.collection.update_one({'Date': date},
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST', '$Amount']}}}])
        self.showall()  

    def delete(self):
        id = int(self.id_entry.get())
        self.collection.delete_one({'Id':id})
        self.reset()
        self.showall()

class sad(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Id', 'Date','Party', 'INVNo', 'Qty', 'QtyType',\
                                'Amount', 'SGST', 'CGST','IGST', 'Total')
        self.tree.column("Id", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Party", width=80, minwidth=25)
        self.tree.column("INVNo", width=80, minwidth=25)
        self.tree.column("Qty", width=80, minwidth=25)
        self.tree.column("QtyType", width=80, minwidth=25)
        self.tree.column("Amount", width=80, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Id', text='Id', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('Party', text='Party')
        self.tree.heading('INVNo', text='INV.NO')
        self.tree.heading('Qty', text='QTY')
        self.tree.heading('QtyType', text='QTY Type')
        self.tree.heading('Amount', text='Amount')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.id_label = ttk.Label(self.data_Entry, text="Id:")
        self.id_label.grid(row=0, column=0, sticky=tk.W)

        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)
        
        self.p_label = ttk.Label(self.data_Entry, text="Party:")
        self.p_label.grid(row=2, column=0, sticky=tk.W)

        self.inv_label = ttk.Label(self.data_Entry, text="INV.NO:")
        self.inv_label.grid(row=3, column=0, sticky=tk.W)

        self.Qty_label = ttk.Label(self.data_Entry, text="Quantity:")
        self.Qty_label.grid(row=4, column=0, sticky=tk.W)

        self.qtyt_label = ttk.Label(self.data_Entry, text="Qty Type:")
        self.qtyt_label.grid(row=5, column=0, sticky=tk.W)

        self.amount_label = ttk.Label(self.data_Entry, text = "Amount:")
        self.amount_label.grid(row=6,column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=7, column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=8, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=9, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.id_entry = ttk.Entry(self.data_Entry)
        self.id_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)
        
        self.p_entry = ttk.Entry(self.data_Entry)
        self.p_entry.grid(row=2, column=1, sticky=tk.W)

        self.inv_entry = ttk.Entry(self.data_Entry)
        self.inv_entry.grid(row=3, column=1, sticky=tk.W)

        self.Qty_entry = ttk.Entry(self.data_Entry)
        self.Qty_entry.grid(row=4, column=1, sticky=tk.W)

        self.qtyt_entry = ttk.Entry(self.data_Entry)
        self.qtyt_entry.grid(row=5, column=1, sticky=tk.W)

        self.amount_entry = ttk.Entry(self.data_Entry)
        self.amount_entry.grid(row=6, column=1, sticky=tk.W)
        
        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=7, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=9, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=11, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['Id', 'Date','Party', 'INVNo', 'Qty', 'QtyType',\
                                'Amount', 'SGST', 'CGST','IGST', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Date",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], i['Date'], i['INVNo'], i['Party'], i['Qty'],\
                                    i['QtyType'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                    i['Total']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalAmt": { "$sum": "$Amount" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, j['totalQty'],
                                        0,j['totalAmt'], j['totalSGST'], j['totalCGST'],
                                        j['totalIGST'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.id_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.p_entry.delete(0, tk.END)
            self.inv_entry.delete(0, tk.END)
            self.Qty_entry.delete(0, tk.END)
            self.qtyt_entry.delete(0, tk.END)
            self.amount_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            
            self.id_entry.insert(0, values[0])
            self.Date_entry.entry.insert(0, values[1])
            self.p_entry.insert(0, values[2])
            self.inv_entry.insert(0, values[3])
            self.Qty_entry.insert(0, values[4])
            self.qtyt_entry.insert(0, values[5])
            self.amount_entry.insert(0, values[6])
            self.SGST_entry.insert(0, values[7])
            self.CGST_entry.insert(0, values[8])
            self.IGST_entry.insert(0, values[9])


    def add(self):
        try:
            id = int(self.id_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            i = self.inv_entry.get()
            qty = float(self.Qty_entry.get())
            qtyt = self.qtyt_entry.get()
            p = self.p_entry.get()
            amt = float(self.amount_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            data = {'Id':id, 'Date':date, 'Party': p, 'INVNo':i,'Qty':qty, 'QtyType':qtyt,
                    'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }
            
            self.collection.insert_one(data)
            self.collection.update_one({'Id': id},
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST', '$Amount']}}}])
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.id_entry.delete(0, tk.END)
        self.p_entry.delete(0, tk.END)
        self.inv_entry.delete(0, tk.END)
        self.Qty_entry.delete(0, tk.END)
        self.qtyt_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)
            
        self.p_entry.insert(0, 0)
        self.inv_entry.insert(0, 0)
        self.Qty_entry.insert(0, 0)
        self.qtyt_entry.insert(0, 0)
        self.amount_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Date",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalAmt": { "$sum": "$Amount" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], i['Date'], i['INVNo'], i['Party'], i['Qty'],\
                                    i['QtyType'], i['Amount'], i['SGST'], i['CGST'],i['IGST'],\
                                    i['Total']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, j['totalQty'],
                                        0,j['totalAmt'], j['totalSGST'], j['totalCGST'],
                                        j['totalIGST'], j['total']), tags= 'total')

    def update(self):
        id = int(self.id_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        i = self.inv_entry.get()
        qty = float(self.Qty_entry.get())
        qtyt = self.qtyt_entry.get()
        p = self.p_entry.get()
        amt = float(self.amount_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get()) 
        igst = float(self.IGST_entry.get())
        self.collection.update_one({"Id": id}, {'$set':{'Date':date, 'Party': p, 'INVNo':i,'Qty':qty, 'QtyType':qtyt,
                    'Amount':amt, "SGST":sgst, "CGST": cgst, "IGST":igst }}) 
        self.collection.update_one({'Date': date},
                                [{'$addFields': {'Total': {'$add': ['$SGST', '$CGST','$IGST', '$Amount']}}}])
        self.showall()  

    def delete(self):
        id = int(self.id_entry.get())
        self.collection.delete_one({'Id':id})
        self.reset()
        self.showall()

class gstapp(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Date','scgst', 'ssgst', 'sigst', 'pcgst',\
                                'psgst', 'pigst', 'Paid','PDate')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("scgst", width=80, minwidth=25)
        self.tree.column("ssgst", width=80, minwidth=25)
        self.tree.column("sigst", width=80, minwidth=25)
        self.tree.column("pcgst", width=80, minwidth=25)
        self.tree.column("psgst", width=80, minwidth=25)
        self.tree.column("pigst", width=80, minwidth=25)
        self.tree.column("Paid", width=80, minwidth=25)
        self.tree.column("PDate", width=80, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('scgst', text='SalesCGST')
        self.tree.heading('ssgst', text='SalesSGST')
        self.tree.heading('sigst', text='SalesIGST')
        self.tree.heading('pcgst', text='PurchaseCGST')
        self.tree.heading('psgst', text='PurchaseSGST')
        self.tree.heading('pigst', text='PurchaseIGST')
        self.tree.heading('Paid', text='Paid')
        self.tree.heading('PDate', text='Paid Date')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=0, column=0, sticky=tk.W)
        
        self.scgst_label = ttk.Label(self.data_Entry, text="Sales CGST:")
        self.scgst_label.grid(row=1, column=0, sticky=tk.W)

        self.ssgst_label = ttk.Label(self.data_Entry, text="Sales SGST:")
        self.ssgst_label.grid(row=2, column=0, sticky=tk.W)

        self.sigst_label = ttk.Label(self.data_Entry, text="Sales IGST:")
        self.sigst_label.grid(row=3, column=0, sticky=tk.W)

        self.pcgst_label = ttk.Label(self.data_Entry, text="Purchase CGST:")
        self.pcgst_label.grid(row=4, column=0, sticky=tk.W)

        self.psgst_label = ttk.Label(self.data_Entry, text = "Purchase SGST:")
        self.psgst_label.grid(row=5,column=0, sticky=tk.W)

        self.pigst_label = ttk.Label(self.data_Entry, text="Purchase IGST:")
        self.pigst_label.grid(row=6, column=0, sticky=tk.W)

        self.Paid_label = ttk.Label(self.data_Entry, text="Paid:")
        self.Paid_label.grid(row=7, column=0, sticky=tk.W)

        self.pdate_label = ttk.Label(self.data_Entry, text="Paid Date:")
        self.pdate_label.grid(row=8, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=0, column=1, sticky=tk.W)
        
        self.scgst_entry = ttk.Entry(self.data_Entry)
        self.scgst_entry.grid(row=1, column=1, sticky=tk.W)

        self.ssgst_entry = ttk.Entry(self.data_Entry)
        self.ssgst_entry.grid(row=2, column=1, sticky=tk.W)

        self.sigst_entry = ttk.Entry(self.data_Entry)
        self.sigst_entry.grid(row=3, column=1, sticky=tk.W)

        self.pcgst_entry = ttk.Entry(self.data_Entry)
        self.pcgst_entry.grid(row=4, column=1, sticky=tk.W)

        self.psgst_entry = ttk.Entry(self.data_Entry)
        self.psgst_entry.grid(row=5, column=1, sticky=tk.W)
        
        self.pigst_entry = ttk.Entry(self.data_Entry)
        self.pigst_entry.grid(row=6, column=1, sticky=tk.W)

        self.Paid_entry = ttk.Entry(self.data_Entry)
        self.Paid_entry.grid(row=7, column=1, sticky=tk.W)

        self.pd_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.pd_entry.grid(row=8, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=11, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Date','scgst', 'ssgst', 'sigst', 'pcgst',
                                'psgst', 'pigst', 'Paid','PDate'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Date",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter, i['Date'], i['scgst'], i['ssgst'], i['sigst'],\
                                    i['pcgst'], i['psgst'], i['pigst'], i['Paid'],i['PDate']))
            counter+=1

        pipeline = [{ "$group": { "_id": None, "totalsc": { "$sum": "$scgst" }, 
                                            "totalss": { "$sum": "$ssgst" },
                                            "totalsi":{"$sum":"$sigst" },
                                            "totalpc":{ "$sum": "$pcgst" }, 
                                            "totalps":{"$sum":"$psgst"}, 
                                            "totalpi":{"$sum":"$pigst"},
                                            "totalp":{"$sum":"$Paid"}
                                            }}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total', 0, j['totalsc'],
                                        j['totalss'], 
                                        j['totalsi'], 
                                        j['totalpc'],
                                        j['totalps'], 
                                        j['totalpi'],
                                        j['totalp'],0,
                                        ), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.Date_entry.entry.delete(0, tk.END)
            self.scgst_entry.delete(0, tk.END)
            self.ssgst_entry.delete(0, tk.END)
            self.sigst_entry.delete(0, tk.END)
            self.pcgst_entry.delete(0, tk.END)
            self.psgst_entry.delete(0, tk.END)
            self.pigst_entry.delete(0, tk.END)
            self.Paid_entry.delete(0, tk.END)
            self.pd_entry.delete(0, tk.END)
            
            self.Date_entry.entry.insert(0, values[1])
            self.scgst_entry.insert(0, values[2])
            self.ssgst_entry.insert(0, values[3])
            self.sigst_entry.insert(0, values[4])
            self.pcgst_entry.insert(0, values[5])
            self.psgst_entry.insert(0, values[6])
            self.pigst_entry.insert(0, values[7])
            self.Paid_entry.insert(0, values[8])
            self.pd_entry.insert(0, values[9])

    def add(self):
        try:
            datestr1 = self.Date_entry.entry.get()
            if len(datestr1)<= 10:
                datestr1 = datestr1+" 00:00:00"
            date1 = datetime.datetime.strptime(datestr1, '%Y-%m-%d 00:00:00')
            sc = float(self.scgst_entry.get())
            ss = float(self.ssgst_entry.get())
            si = float(self.sigst_entry.get())
            pc = float(self.pcgst_entry.get())
            ps = float(self.psgst_entry.get())
            pi = float(self.pigst_entry.get()) 
            p = float(self.Paid_entry.get())
            datestr2 = self.pd_entry.get()
            if len(datestr2)<= 10:
                datestr2 = datestr2+" 00:00:00"
            date2 = datetime.datetime.strptime(datestr2, '%Y-%m-%d 00:00:00')
            data = {'Date':date1, 'scgst': sc, 'ssgst':ss,'sigst':si, 'pcgst':pc,
                    'psgst':ps, "pigst":pi, "Paid": p, "PDate":date1 }
            
            self.collection.insert_one(data)
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.scgst_entry.delete(0, tk.END)
        self.ssgst_entry.delete(0, tk.END)
        self.sigst_entry.delete(0, tk.END)
        self.pcgst_entry.delete(0, tk.END)
        self.psgst_entry.delete(0, tk.END)
        self.pigst_entry.delete(0, tk.END)
        self.Paid_entry.delete(0, tk.END)
            
        self.scgst_entry.insert(0, 0)
        self.ssgst_entry.insert(0, 0)
        self.sigst_entry.insert(0, 0)
        self.pcgst_entry.insert(0, 0)
        self.psgst_entry.insert(0, 0)
        self.pigst_entry.insert(0, 0)
        self.Paid_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Date",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalsc": { "$sum": "$scgst" }, 
                                            "totalss": { "$sum": "$ssgst" },
                                            "totalsi":{"$sum":"$sigst" },
                                            "totalpc":{ "$sum": "$pcgst" }, 
                                            "totalps":{"$sum":"$psgst"}, 
                                            "totalpi":{"$sum":"$pigst"},
                                            "totalp":{"$sum":"$Paid"}
                                            }}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        counter = 1
        for i in data:
            self.tree.insert('', 'end',values=(counter, i['Date'], i['scgst'], i['ssgst'], i['sigst'],\
                                    i['pcgst'], i['psgst'], i['pigst'], i['Paid'],i['PDate']))
            counter+=1
        
        for j in result:
            self.tree.insert('', 'end',values=('Total', j['totalsc'],
                                        j['totalss'], 
                                        j['totalsi'], 
                                        j['totalpc'],
                                        j['totalps'], 
                                        j['totalpi'],
                                        j['totalp'],0,
                                        ), tags= 'total')

    def update(self):
        datestr1 = self.Date_entry.entry.get()
        if len(datestr1)<= 10:
            datestr1 = datestr1+" 00:00:00"
        date1 = datetime.datetime.strptime(datestr1, '%Y-%m-%d 00:00:00')
        sc = float(self.scgst_entry.get())
        ss = float(self.ssgst_entry.get())
        si = float(self.sigst_entry.get())
        pc = float(self.pcgst_entry.get())
        ps = float(self.psgst_entry.get())
        pi = float(self.pigst_entry.get()) 
        p = float(self.Paid_entry.get())
        datestr2 = self.pd_entry.get()
        if len(datestr2)<= 10:
            datestr2 = datestr2+" 00:00:00"
        date2 = datetime.datetime.strptime(datestr2, '%Y-%m-%d 00:00:00')
        self.collection.update_one({"Date": date1}, {'$set':{'Date':date1, 'scgst': sc, 'ssgst':ss,'sigst':si, 'pcgst':pc,
                    'psgst':ps, "pigst":pi, "Paid": p, "PDate":date2 }}) 
        self.showall()  

    def delete(self):
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        self.collection.delete_one({'Date':date})
        self.reset()
        self.showall()

class cash(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('SL.No', 'Date', 'Party', 'Rate', 'Qty', 'QtyType',
                                'Total')
        self.tree.column("SL.No", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Party", width=80, minwidth=25)
        self.tree.column("Rate", width=80, minwidth=25)
        self.tree.column("Qty", width=80, minwidth=25)
        self.tree.column("QtyType", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('SL.No', text='SL.NO', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('Party', text='Party')
        self.tree.heading('Rate', text='Rate')
        self.tree.heading('Qty', text='QTY')
        self.tree.heading('QtyType', text='QTY Type')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.46, relwidth=1, relheight=.55)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.sl_label = ttk.Label(self.data_Entry, text="Sl.NO:")
        self.sl_label.grid(row=0, column=0, sticky=tk.W)

        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)
        
        self.p_label = ttk.Label(self.data_Entry, text="Party:")
        self.p_label.grid(row=2, column=0, sticky=tk.W)

        self.r_label = ttk.Label(self.data_Entry, text="Rate:")
        self.r_label.grid(row=3, column=0, sticky=tk.W)

        self.Qty_label = ttk.Label(self.data_Entry, text="Quantity:")
        self.Qty_label.grid(row=4, column=0, sticky=tk.W)

        self.qtyt_label = ttk.Label(self.data_Entry, text="Qty Type:")
        self.qtyt_label.grid(row=5, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.sl_entry = ttk.Entry(self.data_Entry)
        self.sl_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)
        
        self.p_entry = ttk.Entry(self.data_Entry)
        self.p_entry.grid(row=2, column=1, sticky=tk.W)

        self.r_entry = ttk.Entry(self.data_Entry)
        self.r_entry.grid(row=3, column=1, sticky=tk.W)

        self.Qty_entry = ttk.Entry(self.data_Entry)
        self.Qty_entry.grid(row=4, column=1, sticky=tk.W)

        self.qtyt_entry = ttk.Entry(self.data_Entry)
        self.qtyt_entry.grid(row=5, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=11, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=11, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=11, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=11, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=11, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['SL.No', 'Date', 'Party', 'Rate', 'Qty', 'QtyType',
                                'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Sl",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        for i in data:
            self.tree.insert('', 'end',values=(i['Sl'], i['Date'], i['Party'],i['Rate'], i['Qty'],
                                    i['QtyType'], i['Total']))

        pipeline = [{ "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalR": { "$sum": "$Rate" }, 
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalR'],
                                        j['totalQty'], 0, j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.sl_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.p_entry.delete(0, tk.END)
            self.r_entry.delete(0, tk.END)
            self.Qty_entry.delete(0, tk.END)
            self.qtyt_entry.delete(0, tk.END)
            
            self.sl_entry.insert(0, values[0])
            self.Date_entry.entry.insert(0, values[1])
            self.p_entry.insert(0, values[2])
            self.r_entry.insert(0, values[3])
            self.Qty_entry.insert(0, values[4])
            self.qtyt_entry.insert(0, values[5])

    def add(self):
        try:
            sl = int(self.sl_entry.get())
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            r = float(self.r_entry.get())
            qty = float(self.Qty_entry.get())
            qtyt = self.qtyt_entry.get()
            p = self.p_entry.get()

            data = {'Sl':sl,'Date':date, 'Party': p, 'Rate':r,'Qty':qty, 'QtyType':qtyt }
            
            self.collection.insert_one(data)
            self.collection.update_one({'Sl': sl},
                                [{'$addFields': {'Total': {'$multiply': ['$Rate', '$Qty']}}}])
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.sl_entry.delete(0, tk.END)
        self.p_entry.delete(0, tk.END)
        self.r_entry.delete(0, tk.END)
        self.Qty_entry.delete(0, tk.END)
        self.qtyt_entry.delete(0, tk.END)
            
        self.sl_entry.insert(0, 0)
        self.p_entry.insert(0, 0)
        self.r_entry.insert(0, 0)
        self.Qty_entry.insert(0, 0)
        self.qtyt_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Sl",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalR": { "$sum": "$Rate" }, 
                                            "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Sl'], i['Date'], i['Party'],i['Rate'], i['Qty'],
                                    i['QtyType'], i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['totalR'],
                                        j['totalQty'], 0, j['total']), tags= 'total')

    def update(self):
        sl = int(self.sl_entry.get())
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        r = float(self.r_entry.get())
        qty = float(self.Qty_entry.get())
        qtyt = self.qtyt_entry.get()
        p = self.p_entry.get()
        self.collection.update_one({"Sl": sl}, {'$set':{'Sl':sl,'Date':date, 
                                            'Party': p, 'Rate':r,'Qty':qty, 'QtyType':qtyt }}) 
        self.collection.update_one({'Sl': sl},
                                [{'$addFields': {'Total': {'$multiply': ['$Rate', '$Qty']}}}])
        self.showall()  

    def delete(self):
        sl = self.sl_entry.get()
        self.collection.delete_one({'Sl':sl})
        self.reset()
        self.showall()

class gstpsd(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Id', 'Date','Party', 'VNo', 'GSTIN','Qty', 'Unit',\
                                'Purchase', 'CGST', 'SGST','IGST','RoundOff', 'Total')
        self.tree.column("Id", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Party", width=80, minwidth=25)
        self.tree.column("VNo", width=80, minwidth=25)
        self.tree.column("GSTIN", width=80, minwidth=25)
        self.tree.column("Qty", width=80, minwidth=25)
        self.tree.column("Unit", width=80, minwidth=25)
        self.tree.column("Purchase", width=80, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("RoundOff", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Id', text='Id', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('Party', text='Party')
        self.tree.heading('VNo', text='VoucherNo')
        self.tree.heading('GSTIN', text='GSTIN/UIN')
        self.tree.heading('Qty', text='QTY')
        self.tree.heading('Unit', text='Unit')
        self.tree.heading('Purchase', text='Purchase')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('RoundOff', text='RoundOff')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.50)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.i_label = ttk.Label(self.data_Entry, text="ID:")
        self.i_label.grid(row=0, column=0, sticky=tk.W)

        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)
        
        self.p_label = ttk.Label(self.data_Entry, text="Party:")
        self.p_label.grid(row=2, column=0, sticky=tk.W)

        self.v_label = ttk.Label(self.data_Entry, text="Voucher No:")
        self.v_label.grid(row=3, column=0, sticky=tk.W)

        self.g_label = ttk.Label(self.data_Entry, text="GSTIN/UIN:")
        self.g_label.grid(row=4, column=0, sticky=tk.W)

        self.Qty_label = ttk.Label(self.data_Entry, text="Quantity:")
        self.Qty_label.grid(row=5, column=0, sticky=tk.W)

        self.u_label = ttk.Label(self.data_Entry, text="Unit:")
        self.u_label.grid(row=6, column=0, sticky=tk.W)

        self.pur_label = ttk.Label(self.data_Entry, text = "Purchase:")
        self.pur_label.grid(row=7,column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=8, column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=9, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=10, column=0, sticky=tk.W)

        self.r_label = ttk.Label(self.data_Entry, text="Round Off:")
        self.r_label.grid(row=11, column=0, sticky=tk.W)

        self.t_label = ttk.Label(self.data_Entry, text="Gross Total:")
        self.t_label.grid(row=12, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.i_entry = ttk.Entry(self.data_Entry)
        self.i_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)
        
        self.p_entry = ttk.Entry(self.data_Entry)
        self.p_entry.grid(row=2, column=1, sticky=tk.W)

        self.v_entry = ttk.Entry(self.data_Entry)
        self.v_entry.grid(row=3, column=1, sticky=tk.W)
        
        self.g_entry = ttk.Entry(self.data_Entry)
        self.g_entry.grid(row=4, column=1, sticky=tk.W)

        self.Qty_entry = ttk.Entry(self.data_Entry)
        self.Qty_entry.grid(row=5, column=1, sticky=tk.W)

        self.u_entry = ttk.Entry(self.data_Entry)
        self.u_entry.grid(row=6, column=1, sticky=tk.W)

        self.pur_entry = ttk.Entry(self.data_Entry)
        self.pur_entry.grid(row=7, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=9, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=10, column=1, sticky=tk.W)

        self.r_entry = ttk.Entry(self.data_Entry)
        self.r_entry.grid(row=11, column=1, sticky=tk.W)

        self.t_entry = ttk.Entry(self.data_Entry)
        self.t_entry.grid(row=12, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['Id', 'Date','Party', 'VNo', 'GSTIN','Qty', 'Unit',
                                'Purchase', 'CGST', 'SGST','IGST','RoundOff', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Date",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], 
                                               i['Date'], 
                                               i['Party'],
                                               i['VoucherNo'], 
                                               i['GSTIN_UIN'], 
                                               i['Qty'],
                                               i['Unit'], 
                                               i['Purchase'], 
                                               i['CGST'], 
                                               i['SGST'],
                                               i['IGST'],
                                               i['RoundOff'],
                                               i['Total']))

        pipeline = [{ "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalp": { "$sum": "$Purchase" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "totalr":{"$sum":"$RoundOff"},
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, 0, j['totalQty'],
                                        0,j['totalp'], j['totalCGST'], j['totalSGST'],
                                        j['totalIGST'], j['totalr'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.i_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.p_entry.delete(0, tk.END)
            self.v_entry.delete(0, tk.END)
            self.g_entry.delete(0, tk.END)
            self.Qty_entry.delete(0, tk.END)
            self.u_entry.delete(0, tk.END)
            self.pur_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            self.r_entry.delete(0, tk.END)
            self.t_entry.delete(0, tk.END)
            
            self.i_entry.insert(0, values[0])
            self.Date_entry.entry.insert(0, values[1])
            self.p_entry.insert(0, values[2])
            self.v_entry.insert(0, values[3])
            self.g_entry.insert(0, values[4])
            self.Qty_entry.insert(0, values[5])
            self.u_entry.insert(0, values[6])
            self.pur_entry.insert(0, values[7])
            self.CGST_entry.insert(0, values[8])
            self.SGST_entry.insert(0, values[9])
            self.IGST_entry.insert(0, values[10])
            self.r_entry.insert(0, values[11])
            self.t_entry.insert(0, values[12])


    def add(self):
        try:
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            i = int(self.i_entry.get())
            qty = float(self.Qty_entry.get())
            u = self.u_entry.get()
            v = self.v_entry.get()
            g = self.g_entry.get()
            p = self.p_entry.get()
            pur = float(self.pur_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            r = float(self.r_entry.get())
            t = float(self.t_entry.get())
            data = {'Id':i,'Date':date, 'Party': p, 'VoucherNo':v,'GSTIN_UIN':g,
                    'Qty':qty, 'Unit':u,
                    'Purchase':pur, "CGST":cgst, "SGST": sgst, "IGST":igst,
                     'RoundOff':r, 'Total':t }
            
            self.collection.insert_one(data)
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.i_entry.delete(0, tk.END)
        self.p_entry.delete(0, tk.END)
        self.v_entry.delete(0, tk.END)
        self.g_entry.delete(0, tk.END)
        self.Qty_entry.delete(0, tk.END)
        self.u_entry.delete(0, tk.END)
        self.pur_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)
        self.r_entry.delete(0, tk.END)
        self.t_entry.delete(0, tk.END)
            
        self.i_entry.insert(0, 0)
        self.p_entry.insert(0, 0)
        self.v_entry.insert(0, 0)
        self.g_entry.insert(0, 0)
        self.Qty_entry.insert(0, 0)
        self.u_entry.insert(0, 0)
        self.pur_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)
        self.r_entry.insert(0, 0)
        self.t_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Date",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totalp": { "$sum": "$Purchase" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "totalr":{"$sum":"$RoundOff"},
                                            "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], 
                                               i['Date'], 
                                               i['Party'],
                                               i['VoucherNo'], 
                                               i['GSTIN_UIN'], 
                                               i['Qty'],
                                               i['Unit'], 
                                               i['Purchase'], 
                                               i['CGST'], 
                                               i['SGST'],
                                               i['IGST'],
                                               i['RoundOff'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, 0, j['totalQty'],
                                        0,j['totalp'], j['totalCGST'], j['totalSGST'],
                                        j['totalIGST'], j['totalr'], j['total']), tags= 'total')

    def update(self):
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        i = int(self.i_entry.get())
        qty = float(self.Qty_entry.get())
        u = self.u_entry.get()
        v = self.v_entry.get()
        g = self.g_entry.get()
        p = self.p_entry.get()
        pur = float(self.pur_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get()) 
        igst = float(self.IGST_entry.get())
        r = float(self.r_entry.get())
        t = float(self.t_entry.get())
        self.collection.update_one({"Id": i}, {'$set':{'Id':i,
                    'Date':date, 'Party': p, 
                    'VoucherNo':v,'GSTIN_UIN':g,
                    'Qty':qty, 'Unit':u,
                    'Purchase':pur, "CGST":cgst, 
                    "SGST": sgst, "IGST":igst,
                    'RoundOff':r, 'Total':t }}) 
        self.showall()  

    def delete(self):
        i = int(self.i_entry.get())
        self.collection.delete_one({'Id':i})
        self.reset()
        self.showall()

class gstsad(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Id', 'Date','Party', 'VNo', 'GSTIN','Qty', 'Unit',\
                                'Sales', 'CGST', 'SGST','IGST','RoundOff', 'Total')
        self.tree.column("Id", width=50, minwidth=25)
        self.tree.column("Date", width=130, minwidth=25)
        self.tree.column("Party", width=80, minwidth=25)
        self.tree.column("VNo", width=80, minwidth=25)
        self.tree.column("GSTIN", width=80, minwidth=25)
        self.tree.column("Qty", width=80, minwidth=25)
        self.tree.column("Unit", width=80, minwidth=25)
        self.tree.column("Sales", width=80, minwidth=25)
        self.tree.column("SGST", width=80, minwidth=25)
        self.tree.column("CGST", width=80, minwidth=25)
        self.tree.column("IGST", width=80, minwidth=25)
        self.tree.column("RoundOff", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Id', text='Id', anchor=tk.CENTER)
        self.tree.heading('Date', text='Date')
        self.tree.heading('Party', text='Party')
        self.tree.heading('VNo', text='VoucherNo')
        self.tree.heading('GSTIN', text='GSTIN/UIN')
        self.tree.heading('Qty', text='QTY')
        self.tree.heading('Unit', text='Unit')
        self.tree.heading('Sales', text='Sales')
        self.tree.heading('SGST', text='SGST')
        self.tree.heading('CGST', text='CGST')
        self.tree.heading('IGST', text='IGST')
        self.tree.heading('RoundOff', text='RoundOff')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.50)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        #labels
        self.i_label = ttk.Label(self.data_Entry, text="ID:")
        self.i_label.grid(row=0, column=0, sticky=tk.W)

        self.Date_label = ttk.Label(self.data_Entry, text="Date:")
        self.Date_label.grid(row=1, column=0, sticky=tk.W)
        
        self.p_label = ttk.Label(self.data_Entry, text="Party:")
        self.p_label.grid(row=2, column=0, sticky=tk.W)

        self.v_label = ttk.Label(self.data_Entry, text="Voucher No:")
        self.v_label.grid(row=3, column=0, sticky=tk.W)

        self.g_label = ttk.Label(self.data_Entry, text="GSTIN/UIN:")
        self.g_label.grid(row=4, column=0, sticky=tk.W)

        self.Qty_label = ttk.Label(self.data_Entry, text="Quantity:")
        self.Qty_label.grid(row=5, column=0, sticky=tk.W)

        self.u_label = ttk.Label(self.data_Entry, text="Unit:")
        self.u_label.grid(row=6, column=0, sticky=tk.W)

        self.pur_label = ttk.Label(self.data_Entry, text = "Sales:")
        self.pur_label.grid(row=7,column=0, sticky=tk.W)

        self.CGST_label = ttk.Label(self.data_Entry, text="CGST:")
        self.CGST_label.grid(row=8, column=0, sticky=tk.W)

        self.SGST_label = ttk.Label(self.data_Entry, text="SGST:")
        self.SGST_label.grid(row=9, column=0, sticky=tk.W)

        self.IGST_label = ttk.Label(self.data_Entry, text="IGST:")
        self.IGST_label.grid(row=10, column=0, sticky=tk.W)

        self.r_label = ttk.Label(self.data_Entry, text="Round Off:")
        self.r_label.grid(row=11, column=0, sticky=tk.W)

        self.t_label = ttk.Label(self.data_Entry, text="Gross Total:")
        self.t_label.grid(row=12, column=0, sticky=tk.W)

        self.from_label = ttk.Label(self.range_entry, text="From:")
        self.from_label.grid(row=0, column= 4, sticky=tk.W)

        self.to_label = ttk.Label(self.range_entry, text="To:")
        self.to_label.grid(row=0, column=6,sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.i_entry = ttk.Entry(self.data_Entry)
        self.i_entry.grid(row=0, column=1, sticky=tk.W)

        self.Date_entry = ttk.DateEntry(self.data_Entry,
                            dateformat='%Y-%m-%d')
        self.Date_entry.grid(row=1, column=1, sticky=tk.W)
        
        self.p_entry = ttk.Entry(self.data_Entry)
        self.p_entry.grid(row=2, column=1, sticky=tk.W)

        self.v_entry = ttk.Entry(self.data_Entry)
        self.v_entry.grid(row=3, column=1, sticky=tk.W)
        
        self.g_entry = ttk.Entry(self.data_Entry)
        self.g_entry.grid(row=4, column=1, sticky=tk.W)

        self.Qty_entry = ttk.Entry(self.data_Entry)
        self.Qty_entry.grid(row=5, column=1, sticky=tk.W)

        self.u_entry = ttk.Entry(self.data_Entry)
        self.u_entry.grid(row=6, column=1, sticky=tk.W)

        self.pur_entry = ttk.Entry(self.data_Entry)
        self.pur_entry.grid(row=7, column=1, sticky=tk.W)

        self.CGST_entry = ttk.Entry(self.data_Entry)
        self.CGST_entry.grid(row=8, column=1, sticky=tk.W)

        self.SGST_entry = ttk.Entry(self.data_Entry)
        self.SGST_entry.grid(row=9, column=1, sticky=tk.W)

        self.IGST_entry = ttk.Entry(self.data_Entry)
        self.IGST_entry.grid(row=10, column=1, sticky=tk.W)

        self.r_entry = ttk.Entry(self.data_Entry)
        self.r_entry.grid(row=11, column=1, sticky=tk.W)

        self.t_entry = ttk.Entry(self.data_Entry)
        self.t_entry.grid(row=12, column=1, sticky=tk.W)

        self.from_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.from_entry.grid(row=0, column=5, sticky=tk.W)

        self.to_entry = ttk.DateEntry(self.range_entry,
                            dateformat='%Y-%m-%d')
        self.to_entry.grid(row=0, column=7, sticky=tk.W)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.findRange_button = ttk.Button(self.range_entry, text="Find Range", command=self.range)
        self.findRange_button.grid(row=1, column=4, sticky=tk.NSEW, padx=3, pady=3)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        df = pd.DataFrame(data, columns=['Id', 'Date','Party', 'VNo', 'GSTIN','Qty', 'Unit',
                                'Sales', 'CGST', 'SGST','IGST','RoundOff', 'Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Date",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], 
                                               i['Date'], 
                                               i['Party'],
                                               i['VoucherNo'], 
                                               i['GSTIN_UIN'], 
                                               i['Qty'],
                                               i['Unit'], 
                                               i['Sales'], 
                                               i['CGST'], 
                                               i['SGST'],
                                               i['IGST'],
                                               i['RoundOff'],
                                               i['Total']))

        pipeline = [{ "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totals": { "$sum": "$Sales" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "totalr":{"$sum":"$RoundOff"},
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, 0, j['totalQty'],
                                        0,j['totals'], j['totalCGST'], j['totalSGST'],
                                        j['totalIGST'], j['totalr'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.i_entry.delete(0, tk.END)
            self.Date_entry.entry.delete(0, tk.END)
            self.p_entry.delete(0, tk.END)
            self.v_entry.delete(0, tk.END)
            self.g_entry.delete(0, tk.END)
            self.Qty_entry.delete(0, tk.END)
            self.u_entry.delete(0, tk.END)
            self.pur_entry.delete(0, tk.END)
            self.SGST_entry.delete(0, tk.END)
            self.CGST_entry.delete(0, tk.END)
            self.IGST_entry.delete(0, tk.END)
            self.r_entry.delete(0, tk.END)
            self.t_entry.delete(0, tk.END)
            
            self.i_entry.insert(0, values[0])
            self.Date_entry.entry.insert(0, values[1])
            self.p_entry.insert(0, values[2])
            self.v_entry.insert(0, values[3])
            self.g_entry.insert(0, values[4])
            self.Qty_entry.insert(0, values[5])
            self.u_entry.insert(0, values[6])
            self.pur_entry.insert(0, values[7])
            self.CGST_entry.insert(0, values[8])
            self.SGST_entry.insert(0, values[9])
            self.IGST_entry.insert(0, values[10])
            self.r_entry.insert(0, values[11])
            self.t_entry.insert(0, values[12])


    def add(self):
        try:
            datestr = self.Date_entry.entry.get()
            if len(datestr)<= 10:
                datestr = datestr+" 00:00:00"
            date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
            i = int(self.i_entry.get())
            qty = float(self.Qty_entry.get())
            u = self.u_entry.get()
            v = self.v_entry.get()
            g = self.g_entry.get()
            p = self.p_entry.get()
            pur = float(self.pur_entry.get())
            sgst = float(self.SGST_entry.get())
            cgst = float(self.CGST_entry.get()) 
            igst = float(self.IGST_entry.get())
            r = float(self.r_entry.get())
            t = float(self.t_entry.get())
            data = {'Id':i,'Date':date, 'Party': p, 'VoucherNo':v,'GSTIN_UIN':g,
                    'Qty':qty, 'Unit':u,
                    'Sales':pur, "CGST":cgst, "SGST": sgst, "IGST":igst,
                     'RoundOff':r, 'Total':t }
            
            self.collection.insert_one(data)
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)
    
    def reset(self):
        self.i_entry.delete(0, tk.END)
        self.p_entry.delete(0, tk.END)
        self.v_entry.delete(0, tk.END)
        self.g_entry.delete(0, tk.END)
        self.Qty_entry.delete(0, tk.END)
        self.u_entry.delete(0, tk.END)
        self.pur_entry.delete(0, tk.END)
        self.SGST_entry.delete(0, tk.END)
        self.CGST_entry.delete(0, tk.END)
        self.IGST_entry.delete(0, tk.END)
        self.r_entry.delete(0, tk.END)
        self.t_entry.delete(0, tk.END)
            
        self.i_entry.insert(0, 0)
        self.p_entry.insert(0, 0)
        self.v_entry.insert(0, 0)
        self.g_entry.insert(0, 0)
        self.Qty_entry.insert(0, 0)
        self.u_entry.insert(0, 0)
        self.pur_entry.insert(0, 0)
        self.CGST_entry.insert(0, 0)
        self.SGST_entry.insert(0, 0)
        self.IGST_entry.insert(0, 0)
        self.r_entry.insert(0, 0)
        self.t_entry.insert(0, 0)
        
    def range(self):
        self.tree.delete(*self.tree.get_children())
        fromdate = self.from_entry.get()
        fromdate = fromdate+" 00:00:00"
        todate = self.to_entry.get()
        todate = todate+" 00:00:00"
        fd = datetime.datetime.strptime(fromdate, '%Y-%m-%d 00:00:00')
        td = datetime.datetime.strptime(todate, '%Y-%m-%d 00:00:00')
        data = self.collection.find({"Date":{'$gte':fd, '$lte':td}}).sort("Date",1)
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Date":{'$gte':fd, '$lte':td}}},\
                    { "$group": { "_id": None, "totalQty": { "$sum": "$Qty" }, 
                                            "totals": { "$sum": "$Sales" },
                                            "totalSGST":{"$sum":"$SGST" },
                                            "totalCGST":{ "$sum": "$CGST" }, 
                                            "totalIGST":{"$sum":"$IGST"}, 
                                            "totalr":{"$sum":"$RoundOff"},
                                            "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Id'], 
                                               i['Date'], 
                                               i['Party'],
                                               i['VoucherNo'], 
                                               i['GSTIN_UIN'], 
                                               i['Qty'],
                                               i['Unit'], 
                                               i['Sales'], 
                                               i['CGST'], 
                                               i['SGST'],
                                               i['IGST'],
                                               i['RoundOff'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, 0, 0, j['totalQty'],
                                        0,j['totals'], j['totalCGST'], j['totalSGST'],
                                        j['totalIGST'], j['totalr'], j['total']), tags= 'total')

    def update(self):
        datestr = self.Date_entry.entry.get()
        if len(datestr)<= 10:
            datestr = datestr+" 00:00:00"
        date = datetime.datetime.strptime(datestr, '%Y-%m-%d 00:00:00')
        i = int(self.i_entry.get())
        qty = float(self.Qty_entry.get())
        u = self.u_entry.get()
        v = self.v_entry.get()
        g = self.g_entry.get()
        p = self.p_entry.get()
        pur = float(self.pur_entry.get())
        sgst = float(self.SGST_entry.get())
        cgst = float(self.CGST_entry.get()) 
        igst = float(self.IGST_entry.get())
        r = float(self.r_entry.get())
        t = float(self.t_entry.get())
        self.collection.update_one({"Id": i}, {'$set':{'Id':i,
                    'Date':date, 'Party': p, 
                    'VoucherNo':v,'GSTIN_UIN':g,
                    'Qty':qty, 'Unit':u,
                    'Sales':pur, "CGST":cgst, 
                    "SGST": sgst, "IGST":igst,
                    'RoundOff':r, 'Total':t }}) 
        self.showall()  

    def delete(self):
        i = int(self.i_entry.get())
        self.collection.delete_one({'Id':i})
        self.reset()
        self.showall()

class conw(tk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Container', 'Grade','GradeD', 'Trip1', 'Trip2','Trip3', 'Trip4',
                                'Trip5', 'Trip6', 'Trip7', 'Total')
        
        self.tree.column("Container", width=50, minwidth=25)
        self.tree.column("Grade", width=130, minwidth=25)
        self.tree.column("GradeD", width=80, minwidth=25)
        self.tree.column("Trip1", width=80, minwidth=25)
        self.tree.column("Trip2", width=80, minwidth=25)
        self.tree.column("Trip3", width=80, minwidth=25)
        self.tree.column("Trip4", width=80, minwidth=25)
        self.tree.column("Trip5", width=80, minwidth=25)
        self.tree.column("Trip6", width=80, minwidth=25)
        self.tree.column("Trip7", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Container', text='Container', anchor=tk.CENTER)
        self.tree.heading('Grade', text='Grade')
        self.tree.heading('GradeD', text='Subgrade')
        self.tree.heading('Trip1', text='Trip1')
        self.tree.heading('Trip2', text='Trip2')
        self.tree.heading('Trip3', text='Trip3')
        self.tree.heading('Trip4', text='Trip4')
        self.tree.heading('Trip5', text='Trip5')
        self.tree.heading('Trip6', text='Trip6')
        self.tree.heading('Trip7', text='Trip7')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.50)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        self.tt = ttk.LabelFrame(parent, text ="Trip Date details")
        self.tt.grid(row=0, column=2, sticky='nw')

        self.text = tk.Text(self.tt, height=10, width=40)
        self.text.grid(row=0, column=0, sticky='nwew')

        #labels
        self.i_label = ttk.Label(self.data_Entry, text="Container No.:")
        self.i_label.grid(row=0, column=0, sticky=tk.W)

        self.g_label = ttk.Label(self.data_Entry, text="Grade:")
        self.g_label.grid(row=1, column=0, sticky=tk.W)
        
        self.gd_label = ttk.Label(self.data_Entry, text="Sub Grade:")
        self.gd_label.grid(row=2, column=0, sticky=tk.W)

        self.t1_label = ttk.Label(self.data_Entry, text="Trip1:")
        self.t1_label.grid(row=3, column=0, sticky=tk.W)

        self.t2_label = ttk.Label(self.data_Entry, text="Trip2:")
        self.t2_label.grid(row=4, column=0, sticky=tk.W)

        self.t3_label = ttk.Label(self.data_Entry, text="Trip3:")
        self.t3_label.grid(row=5, column=0, sticky=tk.W)

        self.t4_label = ttk.Label(self.data_Entry, text="Trip4:")
        self.t4_label.grid(row=6, column=0, sticky=tk.W)

        self.t5_label = ttk.Label(self.data_Entry, text="Trip5:")
        self.t5_label.grid(row=7, column=0, sticky=tk.W)

        self.t6_label = ttk.Label(self.data_Entry, text="Trip6:")
        self.t6_label.grid(row=8, column=0, sticky=tk.W)

        self.t7_label = ttk.Label(self.data_Entry, text="Trip7:")
        self.t7_label.grid(row=9, column=0, sticky=tk.W)

        self.t_label = ttk.Label(self.data_Entry, text="Gross Total:")
        self.t_label.grid(row=10, column=0, sticky=tk.W)

        self.con_label = ttk.Label(self.range_entry ,text = "Container:")
        self.con_label.grid(row=0, column=0, sticky=tk.W)

        self.s_label = ttk.Label(self.range_entry, text="Grade:")
        self.s_label.grid(row=1, column=0, sticky=tk.W)

        self.s_label = ttk.Label(self.range_entry, text="Sub Grade:")
        self.s_label.grid(row=2, column=0, sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.i_entry = ttk.Entry(self.data_Entry)
        self.i_entry.grid(row=0, column=1, sticky=tk.W)
        
        self.g_entry = ttk.Entry(self.data_Entry)
        self.g_entry.grid(row=1, column=1, sticky=tk.W)

        self.gd_entry = ttk.Entry(self.data_Entry)
        self.gd_entry.grid(row=2, column=1, sticky=tk.W)
        
        self.t1_entry = ttk.Entry(self.data_Entry)
        self.t1_entry.grid(row=3, column=1, sticky=tk.W)

        self.t2_entry = ttk.Entry(self.data_Entry)
        self.t2_entry.grid(row=4, column=1, sticky=tk.W)

        self.t3_entry = ttk.Entry(self.data_Entry)
        self.t3_entry.grid(row=5, column=1, sticky=tk.W)

        self.t4_entry = ttk.Entry(self.data_Entry)
        self.t4_entry.grid(row=6, column=1, sticky=tk.W)

        self.t5_entry = ttk.Entry(self.data_Entry)
        self.t5_entry.grid(row=7, column=1, sticky=tk.W)

        self.t6_entry = ttk.Entry(self.data_Entry)
        self.t6_entry.grid(row=8, column=1, sticky=tk.W)

        self.t7_entry = ttk.Entry(self.data_Entry)
        self.t7_entry.grid(row=9, column=1, sticky=tk.W)

        self.t_entry = ttk.Entry(self.data_Entry)
        self.t_entry.grid(row=10, column=1, sticky=tk.W)
        
        self.con_entry = ttk.Entry(self.range_entry)
        self.con_entry.grid(row=0, column=1)
        
        self.findg_entry = ttk.Entry(self.range_entry)
        self.findg_entry.grid(row=1, column=1)

        self.findsg_entry = ttk.Entry(self.range_entry)
        self.findsg_entry.grid(row=2, column=1)


        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        # self.ad_button = ttk.Button(self.tt, text="Add", command=self.ad)
        # self.ad_button.grid(row=1, column=0, sticky=tk.NSEW)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.calAmt_button = ttk.Button(self.data_Entry, text="Calculate", command=self.cal)
        self.calAmt_button.grid(row=10, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.find_button = ttk.Button(self.range_entry, text = "Find", command=self.find)
        self.find_button.grid(row=0, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.findRange_button = ttk.Button(self.range_entry, text="Find", command=self.range)
        self.findRange_button.grid(row=1, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.findRange_button = ttk.Button(self.range_entry, text="Find", command=self.srange)
        self.findRange_button.grid(row=2, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.reset()

    def cal(self):
        self.t_entry.delete(0, tk.END)
        try:
            amount = float(self.t1_entry.get())+float(self.t2_entry.get())+float(self.t3_entry.get())\
                +float(self.t4_entry.get())+float(self.t5_entry.get())+float(self.t6_entry.get())\
                +float(self.t7_entry.get())
            self.t_entry.insert(0, amount)
           
        except Exception as e:
            messagebox.showerror('Error', e)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        
        df = pd.DataFrame(data, columns=['Con', 'Grade','GradeD', 'Trip1', 'Trip2','Trip3', 'Trip4',
                                'Trip5', 'Trip6', 'Trip7','Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Con",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))

        pipeline = [{ "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                            "total2": { "$sum": "$Trip2" },
                                            "total3": { "$sum": "$Trip3" },
                                            "total4":{"$sum":"$Trip4" },
                                            "total5":{ "$sum": "$Trip5" }, 
                                            "total6":{"$sum":"$Trip6"}, 
                                            "total7":{"$sum":"$Trip7"},
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.i_entry.delete(0, tk.END)
            self.g_entry.delete(0, tk.END)
            self.gd_entry.delete(0, tk.END)
            self.t1_entry.delete(0, tk.END)
            self.t2_entry.delete(0, tk.END)
            self.t3_entry.delete(0, tk.END)
            self.t4_entry.delete(0, tk.END)
            self.t5_entry.delete(0, tk.END)
            self.t6_entry.delete(0, tk.END)
            self.t7_entry.delete(0, tk.END)
            self.t_entry.delete(0, tk.END)
            self.text.delete('1.0', tk.END)
            
            self.i_entry.insert(0, values[0])
            self.g_entry.insert(0, values[1])
            self.gd_entry.insert(0, values[2])
            self.t1_entry.insert(0, values[3])
            self.t2_entry.insert(0, values[4])
            self.t3_entry.insert(0, values[5])
            self.t4_entry.insert(0, values[6])
            self.t5_entry.insert(0, values[7])
            self.t6_entry.insert(0, values[8])
            self.t7_entry.insert(0, values[9])
            self.t_entry.insert(0, values[10])

            query = {'GradeD':values[2]}
            projection = {"Details":1}
            data = self.collection.find_one(query,projection )
            self.text.insert("1.0", data['Details'])
            
    def add(self):
        try:
            i = int(self.i_entry.get())
            g = self.g_entry.get()
            gd = self.gd_entry.get()
            t1 = float(self.t1_entry.get())
            t2 = float(self.t2_entry.get())
            t3 = float(self.t3_entry.get())
            t4 = float(self.t4_entry.get())
            t5 = float(self.t5_entry.get())
            t6 = float(self.t6_entry.get())
            t7 = float(self.t7_entry.get())
            t = float(self.t_entry.get())

            tt = self.text.get("1.0", tk.END)
            da= {"Details":tt}
            data = {'Con': i,'Grade':g, 'GradeD': gd, 'Trip1':t1,'Trip2':t2,
                    'Trip3':t3, 'Trip4':t4,
                    'Trip5':t5, "Trip6":t6, "Trip7": t7,
                    'Total':t, 'Details':tt }
            
            self.collection.insert_one(data)
            
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)

    def ad(self):
        try:
            i = int(self.i_entry.get())
            t = self.text.get("1.0", tk.END)
            data = {"Details":t}
            
            self.collection.update_one({'Con':i}, {'$set':data})
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)


    def reset(self):
        self.i_entry.delete(0, tk.END)
        self.g_entry.delete(0, tk.END)
        self.gd_entry.delete(0, tk.END)
        self.t1_entry.delete(0, tk.END)
        self.t2_entry.delete(0, tk.END)
        self.t3_entry.delete(0, tk.END)
        self.t4_entry.delete(0, tk.END)
        self.t5_entry.delete(0, tk.END)
        self.t6_entry.delete(0, tk.END)
        self.t7_entry.delete(0, tk.END)
        self.t_entry.delete(0, tk.END)
        self.text.delete('1.0', tk.END)
            
        self.i_entry.insert(0, 0)
        self.g_entry.insert(0, 0)
        self.gd_entry.insert(0, 0)
        self.t1_entry.insert(0, 0)
        self.t2_entry.insert(0, 0)
        self.t3_entry.insert(0, 0)
        self.t4_entry.insert(0, 0)
        self.t5_entry.insert(0, 0)
        self.t6_entry.insert(0, 0)
        self.t7_entry.insert(0, 0)
        self.t_entry.insert(0, 0)
        
    def find(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        g = self.findg_entry.get()
        data = self.collection.find({"Con":i})
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Con":i}},
                    { "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                                "total2": { "$sum": "$Trip2" },
                                                "total3": { "$sum": "$Trip3" },
                                                "total4":{"$sum":"$Trip4" },
                                                "total5":{ "$sum": "$Trip5"}, 
                                                "total6":{"$sum":"$Trip6"}, 
                                                "total7":{"$sum":"$Trip7"},
                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=("Total",  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')

    def range(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        g = self.findg_entry.get()
        data = self.collection.find({"Con":i, "Grade": g})
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Grade":g ,"Con":i}},
                    { "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                                "total2": { "$sum": "$Trip2" },
                                                "total3": { "$sum": "$Trip3" },
                                                "total4":{"$sum":"$Trip4" },
                                                "total5":{ "$sum": "$Trip5"}, 
                                                "total6":{"$sum":"$Trip6"}, 
                                                "total7":{"$sum":"$Trip7"},
                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')

    def srange(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        g = self.findsg_entry.get()
        data = self.collection.find({"Con":i, "GradeD": g})
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"GradeD":g, "Con":i}},
                    { "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                                "total2": { "$sum": "$Trip2" },
                                                "total3": { "$sum": "$Trip3" },
                                                "total4":{"$sum":"$Trip4" },
                                                "total5":{ "$sum": "$Trip5"}, 
                                                "total6":{"$sum":"$Trip6"}, 
                                                "total7":{"$sum":"$Trip7"},
                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')

    def update(self):
        i = int(self.i_entry.get())
        g = self.g_entry.get()
        gd = self.gd_entry.get()
        t1 = float(self.t1_entry.get())
        t2 = float(self.t2_entry.get())
        t3 = float(self.t3_entry.get())
        t4 = float(self.t4_entry.get())
        t5 = float(self.t5_entry.get())
        t6 = float(self.t6_entry.get())
        t7 = float(self.t7_entry.get())
        t = float(self.t_entry.get())

        tt = self.text.get("1.0", tk.END)
        self.collection.update_one({"GradeD":gd}, {'$set':{'Con':i,'Grade':g, 'GradeD': gd, 'Trip1':t1,'Trip2':t2,
                    'Trip3':t3, 'Trip4':t4,
                    'Trip5':t5, "Trip6":t6, "Trip7": t7,
                    'Total':t , 'Details':tt}})  

    def delete(self):
        gd = self.gd_entry.get()
        self.collection.delete_one({'GradeD':gd})
        self.reset()
        self.showall()

class conp(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Container', 'Grade','GradeD', 'Trip1', 'Trip2','Trip3', 'Trip4',
                                'Trip5', 'Trip6', 'Trip7', 'Total')
        self.tree.column("Container", width=50, minwidth=25)
        self.tree.column("Grade", width=130, minwidth=25)
        self.tree.column("GradeD", width=80, minwidth=25)
        self.tree.column("Trip1", width=80, minwidth=25)
        self.tree.column("Trip2", width=80, minwidth=25)
        self.tree.column("Trip3", width=80, minwidth=25)
        self.tree.column("Trip4", width=80, minwidth=25)
        self.tree.column("Trip5", width=80, minwidth=25)
        self.tree.column("Trip6", width=80, minwidth=25)
        self.tree.column("Trip7", width=80, minwidth=25)
        self.tree.column("Total", width=100, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Container', text='Container', anchor=tk.CENTER)
        self.tree.heading('Grade', text='Grade')
        self.tree.heading('GradeD', text='Subgrade')
        self.tree.heading('Trip1', text='Trip1')
        self.tree.heading('Trip2', text='Trip2')
        self.tree.heading('Trip3', text='Trip3')
        self.tree.heading('Trip4', text='Trip4')
        self.tree.heading('Trip5', text='Trip5')
        self.tree.heading('Trip6', text='Trip6')
        self.tree.heading('Trip7', text='Trip7')
        self.tree.heading('Total', text='Total')
        self.tree.place(x=0, rely=.50, relwidth=1, relheight=.50)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        self.tt = ttk.LabelFrame(parent, text ="Trip Date details")
        self.tt.grid(row=0, column=2, sticky='nw')

        self.text = tk.Text(self.tt, height=10, width=40)
        self.text.grid(row=0, column=0, sticky='nwew')

        #labels
        self.i_label = ttk.Label(self.data_Entry, text="Container No.:")
        self.i_label.grid(row=0, column=0, sticky=tk.W)

        self.g_label = ttk.Label(self.data_Entry, text="Grade:")
        self.g_label.grid(row=1, column=0, sticky=tk.W)
        
        self.gd_label = ttk.Label(self.data_Entry, text="Sub Grade:")
        self.gd_label.grid(row=2, column=0, sticky=tk.W)

        self.t1_label = ttk.Label(self.data_Entry, text="Trip1:")
        self.t1_label.grid(row=3, column=0, sticky=tk.W)

        self.t2_label = ttk.Label(self.data_Entry, text="Trip2:")
        self.t2_label.grid(row=4, column=0, sticky=tk.W)

        self.t3_label = ttk.Label(self.data_Entry, text="Trip3:")
        self.t3_label.grid(row=5, column=0, sticky=tk.W)

        self.t4_label = ttk.Label(self.data_Entry, text="Trip4:")
        self.t4_label.grid(row=6, column=0, sticky=tk.W)

        self.t5_label = ttk.Label(self.data_Entry, text="Trip5:")
        self.t5_label.grid(row=7, column=0, sticky=tk.W)

        self.t6_label = ttk.Label(self.data_Entry, text="Trip6:")
        self.t6_label.grid(row=8, column=0, sticky=tk.W)

        self.t7_label = ttk.Label(self.data_Entry, text="Trip7:")
        self.t7_label.grid(row=9, column=0, sticky=tk.W)

        self.t_label = ttk.Label(self.data_Entry, text="Gross Total:")
        self.t_label.grid(row=10, column=0, sticky=tk.W)

        self.con_label = ttk.Label(self.range_entry ,text = "Container:")
        self.con_label.grid(row=0, column=0, sticky=tk.W)

        self.s_label = ttk.Label(self.range_entry, text="Grade:")
        self.s_label.grid(row=1, column=0, sticky=tk.W)

        self.s_label = ttk.Label(self.range_entry, text="Sub Grade:")
        self.s_label.grid(row=2, column=0, sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.i_entry = ttk.Entry(self.data_Entry)
        self.i_entry.grid(row=0, column=1, sticky=tk.W)
        
        self.g_entry = ttk.Entry(self.data_Entry)
        self.g_entry.grid(row=1, column=1, sticky=tk.W)

        self.gd_entry = ttk.Entry(self.data_Entry)
        self.gd_entry.grid(row=2, column=1, sticky=tk.W)
        
        self.t1_entry = ttk.Entry(self.data_Entry)
        self.t1_entry.grid(row=3, column=1, sticky=tk.W)

        self.t2_entry = ttk.Entry(self.data_Entry)
        self.t2_entry.grid(row=4, column=1, sticky=tk.W)

        self.t3_entry = ttk.Entry(self.data_Entry)
        self.t3_entry.grid(row=5, column=1, sticky=tk.W)

        self.t4_entry = ttk.Entry(self.data_Entry)
        self.t4_entry.grid(row=6, column=1, sticky=tk.W)

        self.t5_entry = ttk.Entry(self.data_Entry)
        self.t5_entry.grid(row=7, column=1, sticky=tk.W)

        self.t6_entry = ttk.Entry(self.data_Entry)
        self.t6_entry.grid(row=8, column=1, sticky=tk.W)

        self.t7_entry = ttk.Entry(self.data_Entry)
        self.t7_entry.grid(row=9, column=1, sticky=tk.W)

        self.t_entry = ttk.Entry(self.data_Entry)
        self.t_entry.grid(row=10, column=1, sticky=tk.W)
        
        self.con_entry = ttk.Entry(self.range_entry)
        self.con_entry.grid(row=0, column=1)
        
        self.findg_entry = ttk.Entry(self.range_entry)
        self.findg_entry.grid(row=1, column=1)

        self.findsg_entry = ttk.Entry(self.range_entry)
        self.findsg_entry.grid(row=2, column=1)


        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=13, column=0, sticky=tk.NSEW, padx=3, pady=3)

        # self.ad_button = ttk.Button(self.tt, text="Add", command=self.ad)
        # self.ad_button.grid(row=1, column=0, sticky=tk.NSEW)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=13, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=13, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=13, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.save_button = ttk.Button(self.data_Entry, text="Export", command=self.save)
        self.save_button.grid(row=13, column=4, sticky=tk.NSEW, padx=3, pady=3)

        self.calAmt_button = ttk.Button(self.data_Entry, text="Calculate", command=self.cal)
        self.calAmt_button.grid(row=10, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.find_button = ttk.Button(self.range_entry, text = "Find", command=self.find)
        self.find_button.grid(row=0, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.findRange_button = ttk.Button(self.range_entry, text="Find", command=self.range)
        self.findRange_button.grid(row=1, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.findRange_button = ttk.Button(self.range_entry, text="Find", command=self.srange)
        self.findRange_button.grid(row=2, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.reset()

    def cal(self):
        self.t_entry.delete(0, tk.END)
        try:
            amount = float(self.t1_entry.get())+float(self.t2_entry.get())+float(self.t3_entry.get())\
                +float(self.t4_entry.get())+float(self.t5_entry.get())+float(self.t6_entry.get())\
                +float(self.t7_entry.get())
            self.t_entry.insert(0, amount)
           
        except Exception as e:
            messagebox.showerror('Error', e)

    def save(self):
        data = []
        for item in self.tree.get_children():
            values = self.tree.item(item)["values"]
            data.append(values)
        
        df = pd.DataFrame(data, columns=['Con', 'Grade','GradeD', 'Trip1', 'Trip2','Trip3', 'Trip4',
                                'Trip5', 'Trip6', 'Trip7','Total'])
        filename = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if filename:
            df.to_excel(filename, index=False)

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Con",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))

        pipeline = [{ "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                            "total2": { "$sum": "$Trip2" },
                                            "total3": { "$sum": "$Trip3" },
                                            "total4":{"$sum":"$Trip4" },
                                            "total5":{ "$sum": "$Trip5" }, 
                                            "total6":{"$sum":"$Trip6"}, 
                                            "total7":{"$sum":"$Trip7"},
                                            "total":{"$sum":"$Total" }}}]
        result = self.collection.aggregate(pipeline)
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')
        
    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.i_entry.delete(0, tk.END)
            self.g_entry.delete(0, tk.END)
            self.gd_entry.delete(0, tk.END)
            self.t1_entry.delete(0, tk.END)
            self.t2_entry.delete(0, tk.END)
            self.t3_entry.delete(0, tk.END)
            self.t4_entry.delete(0, tk.END)
            self.t5_entry.delete(0, tk.END)
            self.t6_entry.delete(0, tk.END)
            self.t7_entry.delete(0, tk.END)
            self.t_entry.delete(0, tk.END)
            self.text.delete('1.0', tk.END)
            
            self.i_entry.insert(0, values[0])
            self.g_entry.insert(0, values[1])
            self.gd_entry.insert(0, values[2])
            self.t1_entry.insert(0, values[3])
            self.t2_entry.insert(0, values[4])
            self.t3_entry.insert(0, values[5])
            self.t4_entry.insert(0, values[6])
            self.t5_entry.insert(0, values[7])
            self.t6_entry.insert(0, values[8])
            self.t7_entry.insert(0, values[9])
            self.t_entry.insert(0, values[10])

            query = {'GradeD':values[2]}
            projection = {"Details":1}
            data = self.collection.find_one(query,projection )
            self.text.insert("1.0", data['Details'])
            
    def add(self):
        try:
            i = int(self.i_entry.get())
            g = self.g_entry.get()
            gd = self.gd_entry.get()
            t1 = float(self.t1_entry.get())
            t2 = float(self.t2_entry.get())
            t3 = float(self.t3_entry.get())
            t4 = float(self.t4_entry.get())
            t5 = float(self.t5_entry.get())
            t6 = float(self.t6_entry.get())
            t7 = float(self.t7_entry.get())
            t = float(self.t_entry.get())

            tt = self.text.get("1.0", tk.END)
            da= {"Details":tt}
            data = {'Con': i,'Grade':g, 'GradeD': gd, 'Trip1':t1,'Trip2':t2,
                    'Trip3':t3, 'Trip4':t4,
                    'Trip5':t5, "Trip6":t6, "Trip7": t7,
                    'Total':t, 'Details':tt }
            
            self.collection.insert_one(data)
            
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)

    def ad(self):
        try:
            i = int(self.i_entry.get())
            t = self.text.get("1.0", tk.END)
            data = {"Details":t}
            
            self.collection.update_one({'Con':i}, {'$set':data})
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)


    def reset(self):
        self.i_entry.delete(0, tk.END)
        self.g_entry.delete(0, tk.END)
        self.gd_entry.delete(0, tk.END)
        self.t1_entry.delete(0, tk.END)
        self.t2_entry.delete(0, tk.END)
        self.t3_entry.delete(0, tk.END)
        self.t4_entry.delete(0, tk.END)
        self.t5_entry.delete(0, tk.END)
        self.t6_entry.delete(0, tk.END)
        self.t7_entry.delete(0, tk.END)
        self.t_entry.delete(0, tk.END)
        self.text.delete('1.0', tk.END)
            
        self.i_entry.insert(0, 0)
        self.g_entry.insert(0, 0)
        self.gd_entry.insert(0, 0)
        self.t1_entry.insert(0, 0)
        self.t2_entry.insert(0, 0)
        self.t3_entry.insert(0, 0)
        self.t4_entry.insert(0, 0)
        self.t5_entry.insert(0, 0)
        self.t6_entry.insert(0, 0)
        self.t7_entry.insert(0, 0)
        self.t_entry.insert(0, 0)
        
    def find(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        g = self.findg_entry.get()
        data = self.collection.find({"Con":i})
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Con":i}},
                    { "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                                "total2": { "$sum": "$Trip2" },
                                                "total3": { "$sum": "$Trip3" },
                                                "total4":{"$sum":"$Trip4" },
                                                "total5":{ "$sum": "$Trip5"}, 
                                                "total6":{"$sum":"$Trip6"}, 
                                                "total7":{"$sum":"$Trip7"},
                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=("Total",  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')

    def range(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        g = self.findg_entry.get()
        data = self.collection.find({"Con":i, "Grade": g})
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"Grade":g ,"Con":i}},
                    { "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                                "total2": { "$sum": "$Trip2" },
                                                "total3": { "$sum": "$Trip3" },
                                                "total4":{"$sum":"$Trip4" },
                                                "total5":{ "$sum": "$Trip5"}, 
                                                "total6":{"$sum":"$Trip6"}, 
                                                "total7":{"$sum":"$Trip7"},
                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')

    def srange(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        g = self.findsg_entry.get()
        data = self.collection.find({"Con":i, "GradeD": g})
        self.currnet_label.config(text="Showing Range")

        pipeline = [{"$match":{"GradeD":g, "Con":i}},
                    { "$group": { "_id": None, "total1": { "$sum": "$Trip1" }, 
                                                "total2": { "$sum": "$Trip2" },
                                                "total3": { "$sum": "$Trip3" },
                                                "total4":{"$sum":"$Trip4" },
                                                "total5":{ "$sum": "$Trip5"}, 
                                                "total6":{"$sum":"$Trip6"}, 
                                                "total7":{"$sum":"$Trip7"},
                                                "total":{"$sum":"$Total" }}}]
        result =  self.collection.aggregate(pipeline)

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con'],
                                               i['Grade'], 
                                               i['GradeD'],
                                               i['Trip1'], 
                                               i['Trip2'], 
                                               i['Trip3'],
                                               i['Trip4'], 
                                               i['Trip5'], 
                                               i['Trip6'], 
                                               i['Trip7'],
                                               i['Total']))
        
        for j in result:
            self.tree.insert('', 'end',values=('Total',  0, 0, j['total1'],
                                        j['total2'], j['total3'], j['total4'],
                                        j['total5'], j['total6'],j['total7'], j['total']), tags= 'total')

    def update(self):
        i = int(self.i_entry.get())
        g = self.g_entry.get()
        gd = self.gd_entry.get()
        t1 = float(self.t1_entry.get())
        t2 = float(self.t2_entry.get())
        t3 = float(self.t3_entry.get())
        t4 = float(self.t4_entry.get())
        t5 = float(self.t5_entry.get())
        t6 = float(self.t6_entry.get())
        t7 = float(self.t7_entry.get())
        t = float(self.t_entry.get())

        tt = self.text.get("1.0", tk.END)
        self.collection.update_one({"GradeD":gd}, {'$set':{'Con':i,'Grade':g, 'GradeD': gd, 'Trip1':t1,'Trip2':t2,
                    'Trip3':t3, 'Trip4':t4,
                    'Trip5':t5, "Trip6":t6, "Trip7": t7,
                    'Total':t , 'Details':tt}}) 

    def delete(self):
        gd = self.gd_entry.get()
        self.collection.delete_one({'GradeD':gd})
        self.reset()
        self.showall()

class cont(ttk.Frame):
    def __init__(self, parent, db, col):
        super().__init__(parent)
        self.db = client[db]
        self.collection = self.db[col]
        self.create_w(parent)

    def create_w(self, parent):
        self.tree = ttk.Treeview(parent, bootstyle='info')
        self.tree['columns'] = ('Container')
        self.tree.column("Container", width=50, minwidth=25)
        self.tree.column('#0',width=0, stretch=tk.NO )
        self.tree.heading('Container', text='Container', anchor=tk.CENTER)
        self.tree.place(x=0, rely=.60, relwidth=1, relheight=.40)
        self.tree.bind("<Double-1>", self.click)
        self.tree.update()

        #label frames
        self.data_Entry = ttk.LabelFrame(parent, text="Data Entry")
        self.data_Entry.grid(row=0, column=0, sticky=tk.W)

        self.range_entry = ttk.LabelFrame(parent, text="Search")
        self.range_entry.grid(row=0, column=1, sticky=tk.NW)

        self.text = tk.Text(self.data_Entry, height=15, width=60)
        self.text.grid(row=1, column=1, sticky='nsew')

        #labels
        self.i_label = ttk.Label(self.data_Entry, text="Container No.:")
        self.i_label.grid(row=0, column=0, sticky=tk.W)

        self.i_label = ttk.Label(self.data_Entry, text="Details:")
        self.i_label.grid(row=1, column=0, sticky=tk.W)

        self.con_label = ttk.Label(self.range_entry ,text = "Container:")
        self.con_label.grid(row=0, column=0, sticky=tk.W)

        self.currnet_label = ttk.Label(parent, text="Showing All")
        self.currnet_label.grid(row = 0, column=1, sticky=tk.W)

        #entry
        self.i_entry = ttk.Entry(self.data_Entry)
        self.i_entry.grid(row=0, column=1, sticky=tk.NSEW)
        
        self.con_entry = ttk.Entry(self.range_entry)
        self.con_entry.grid(row=0, column=1)

        #buttons
        self.add_button = ttk.Button(self.data_Entry, text="Add", command=self.add)
        self.add_button.grid(row=3, column=0, sticky=tk.NSEW, padx=3, pady=3)

        # self.ad_button = ttk.Button(self.tt, text="Add", command=self.ad)
        # self.ad_button.grid(row=1, column=0, sticky=tk.NSEW)

        self.update_button = ttk.Button(self.data_Entry, text="Update", command=self.update)
        self.update_button.grid(row=3, column=1, sticky=tk.NSEW, padx=3, pady=3)

        self.delete_button = ttk.Button(self.data_Entry, text="Delete", command=self.delete)
        self.delete_button.grid(row=3, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.show_button = ttk.Button(self.data_Entry, text="Show All", command=self.showall)
        self.show_button.grid(row=3, column=3, sticky=tk.NSEW, padx=3, pady=3)

        self.reset_button = ttk.Button(parent, text="Reset",command=self.reset)
        self.reset_button.grid(row=0, column=4, sticky=tk.NE, padx=20)

        self.find_button = ttk.Button(self.range_entry, text = "Find", command=self.find)
        self.find_button.grid(row=0, column=2, sticky=tk.NSEW, padx=3, pady=3)

        self.reset()

    def showall(self):
        self.tree.delete(*self.tree.get_children())
        self.currnet_label.config(text="Showing All")
        data = self.collection.find().sort("Con",1)
        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))

        for i in data:
            self.tree.insert('', 'end',values=(i['Con']))

    def click(self, event):
        item = self.tree.selection()[0]
        values = self.tree.item(item, "values")
        if(values):
            self.i_entry.delete(0, tk.END)
            self.text.delete('1.0', tk.END)
            
            self.i_entry.insert(0, values[0])
    
            query = {'Con':int(values[0])}
            projection = {"Details":1}
            data = self.collection.find_one(query,projection )
            self.text.insert("1.0", data['Details'])
            
    def add(self):
        try:
            i = int(self.i_entry.get())
            t = self.text.get("1.0", tk.END)
            data = {'Con': i,'Details':t}
            
            self.collection.insert_one(data)
            
            self.reset()
            self.showall()
        except Exception as e:
            messagebox.showerror('Error', e)

    def reset(self):
        self.i_entry.delete(0, tk.END)
        self.text.delete('1.0', tk.END)
            
        self.i_entry.insert(0, 0)

    def find(self):
        self.tree.delete(*self.tree.get_children())
        i = int(self.con_entry.get())
        data = self.collection.find({"Con":i})
        self.currnet_label.config(text="Showing Range")

        self.tree.tag_configure('total',  background='#29B6F6', font=('Calibri', 12, 'bold'))
        
        for i in data:
            self.tree.insert('', 'end',values=(i['Con']))

    def update(self):
        i = int(self.i_entry.get())
        t = self.text.get("1.0", tk.END)
        data = {"Details":t}
            
        self.collection.update_one({'Con':i}, {'$set':data})
        
        self.showall()  

    def delete(self):
        i = int(self.i_entry.get())
        self.collection.delete_one({'Con':i})
        self.reset()
        self.showall()


root = ttk.Window()
app = App(root)
root.mainloop()