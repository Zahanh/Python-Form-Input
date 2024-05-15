import tkinter as tk
from tkinter import ttk,messagebox
from tkcalendar import DateEntry
from datetime import datetime
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from src.document import new_document
from src.excel_output import Excel
from src.labels import label
import os,shutil
from pathlib import Path

## pyinstaller -F --hidden-import "babel.numbers" --icon=icon.ico --noconsole display.py --onefile


class display(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.tk.call('source','src/azure.tcl')
        self.tk.call('set_theme','light')
        self.resizable(width=False,height=False)

        self.geometry('675x700')
        self.title('Digital Order Form')
        tabControl = ttk.Notebook(self) 
  
        self.tab1 = ttk.Frame(tabControl) 
        self.tab2 = ttk.Frame(tabControl) 
        self.tab3 = ttk.Frame(tabControl) 
        self.tab4 = ttk.Frame(tabControl) 
        self.tab5 = ttk.Frame(tabControl)
        
        tabControl.add(self.tab1, text ='File Information') 
        tabControl.add(self.tab2, text ='Client Contact Information') 
        tabControl.add(self.tab3, text ='Property Information') 
        tabControl.add(self.tab4, text ='Certifications & Notes') 
        tabControl.add(self.tab5, text ='Submit')
        tabControl.pack(expand = 1, fill ="both")

    def File_Information(self):
        """
        phone number, email, mailing address, file number, cost estimate
        """
        padx = 10
        pady = 10
        title = ttk.Label(master=self.tab1,text='File Information Input\n',justify='center')
        title.grid(column=1,row=0)

        l1 = ttk.Label(master=self.tab1,text='Phone number',justify='left')
        l1.grid(column=0,row=1,padx=padx,pady=pady)
        e1 = ttk.Entry(master=self.tab1,justify='left',width=15)
        e1.insert(0,'(111) 111-1111')
        e1.grid(column=1,row=1,padx=padx,pady=pady,sticky='ew')

        l2 = ttk.Label(master=self.tab1,text='Email',justify='left')
        l2.grid(column=0,row=2,padx=padx,pady=pady)
        e2 = ttk.Entry(master=self.tab1,justify='left',width=15)
        e2.insert(0,'anEmail@gmail.com')
        e2.grid(column=1,row=2,padx=padx,pady=pady,sticky='ew')

        l3 = ttk.Label(master=self.tab1,text='Mailing Address',justify='left')
        l3.grid(column=0,row=3,padx=padx,pady=pady)
        e3 = ttk.Entry(master=self.tab1,justify='left',width=35)
        e3.insert(0,'123 Made Up Lane')
        e3.grid(column=1,row=3,padx=padx,pady=pady,sticky='ew')

        l4 = ttk.Label(master=self.tab1,text=f'File Number ',justify='left')
        l4.grid(column=0,row=4,padx=padx,pady=pady)
        e4 = ttk.Entry(master=self.tab1,justify='left',width=15)
        e4.insert(0,f'{str(datetime.now().year)[-2:]} - ')
        e4.grid(column=1,row=4,padx=padx,pady=pady,sticky='ew')

        l5 = ttk.Label(master=self.tab1,text='Cost Estimate ($)',justify='left')
        l5.grid(column=0,row=5,padx=padx,pady=pady)
        e5 = ttk.Entry(master=self.tab1,justify='left',width=15)
        e5.grid(column=1,row=5,padx=padx,pady=pady,sticky='ew')


        self.tab1.rowconfigure(5)
        self.tab1.columnconfigure(2)

        return [e1,e2,e3,e4,e5]

    def Contact_Information(self):
        """
        requested by, customer, phone number, email, date requested, delever by, closing date
        """
        padx    = 10
        pady    = 10
        title   = ttk.Label(master=self.tab2,text='Client Information Input\n',justify='center')
        title.grid(column=1,row=0)

        l1 = ttk.Label(master=self.tab2,text='Requested By',justify='left')
        l1.grid(column=0,row=1,padx=padx,pady=pady)
        e1 = ttk.Entry(master=self.tab2,justify='left',width=15)
        e1.grid(column=1,row=1,padx=padx,pady=pady,sticky='ew')

        l2 = ttk.Label(master=self.tab2,text='Representing/Customer',justify='left')
        l2.grid(column=0,row=2,padx=padx,pady=pady)
        e2 = ttk.Entry(master=self.tab2,justify='left',width=15)
        e2.grid(column=1,row=2,padx=padx,pady=pady,sticky='ew')

        l3 = ttk.Label(master=self.tab2,text='Phone Number',justify='left')
        l3.grid(column=0,row=3,padx=padx,pady=pady)
        e3 = ttk.Entry(master=self.tab2,justify='left',width=35)
        e3.grid(column=1,row=3,padx=padx,pady=pady,sticky='ew')

        l4 = ttk.Label(master=self.tab2,text=f'Email ',justify='left')
        l4.grid(column=0,row=4,padx=padx,pady=pady)
        e4 = ttk.Entry(master=self.tab2,justify='left',width=15)
        e4.grid(column=1,row=4,padx=padx,pady=pady,sticky='ew')
        
        l5 = ttk.Label(master=self.tab2,text='Date requested',justify='left')
        l5.grid(column=0,row=5,padx=padx,pady=pady)
        cal1 = DateEntry(self.tab2)
        cal1.grid(column=1,row=5,padx=padx,pady=pady,sticky='ew')

        l6 = ttk.Label(master=self.tab2,text='Deliver By',justify='left')
        l6.grid(column=0,row=6,padx=padx,pady=pady)
        cal2 = DateEntry(self.tab2)
        cal2.grid(column=1,row=6,padx=padx,pady=pady,sticky='ew')

        l7 = ttk.Label(master=self.tab2,text='Closing Date',justify='left')
        l7.grid(column=0,row=7,padx=padx,pady=pady)
        cal3 = DateEntry(self.tab2)
        cal3.grid(column=1,row=7,padx=padx,pady=pady,sticky='ew')


        self.tab2.rowconfigure(7)
        self.tab2.columnconfigure(2)
        return [e1,e2,e3,e4,cal1,cal2,cal3]

    def Property_Information(self):
        """
        parcel #, address, subdivision, legal, lot&blk,tax maps, acres
        """
        padx    = 10
        pady    = 10
        master  = self.tab3
        title   = ttk.Label(master=master,text='Property Information Input\n',justify='center')
        title.grid(column=1,row=0)

        l1 = ttk.Label(master=master,text='Parcel #',justify='left')
        l1.grid(column=0,row=1,padx=padx,pady=pady)
        e1 = ttk.Entry(master=master,justify='left',width=15)
        e1.grid(column=1,row=1,padx=padx,pady=pady,sticky='ew')

        l2 = ttk.Label(master=master,text='Address',justify='left')
        l2.grid(column=0,row=2,padx=padx,pady=pady)
        e2 = ttk.Entry(master=master,justify='left',width=15)
        e2.grid(column=1,row=2,padx=padx,pady=pady,sticky='ew')

        l3 = ttk.Label(master=master,text='Subdivision',justify='left')
        l3.grid(column=0,row=3,padx=padx,pady=pady)
        e3 = ttk.Entry(master=master,justify='left',width=35)
        e3.grid(column=1,row=3,padx=padx,pady=pady,sticky='ew')

        l4 = ttk.Label(master=master,text=f'Legal ',justify='left')
        l4.grid(column=0,row=4,padx=padx,pady=pady)
        e4 = ttk.Entry(master=master,justify='left',width=15)
        e4.grid(column=1,row=4,padx=padx,pady=pady,sticky='ew')
        
        l5 = ttk.Label(master=master,text='Lot & Blk',justify='left')
        l5.grid(column=0,row=5,padx=padx,pady=pady)
        e5 = ttk.Entry(master=master,justify='left',width=15)
        e5.grid(column=1,row=5,padx=padx,pady=pady,sticky='ew')

        l6 = ttk.Label(master=master,text='Tax Maps',justify='left')
        l6.grid(column=0,row=6,padx=padx,pady=pady)
        e6 = ttk.Entry(master=master,justify='left',width=15)
        e6.grid(column=1,row=6,padx=padx,pady=pady,sticky='ew')

        l7 = ttk.Label(master=master,text='Acres',justify='left')
        l7.grid(column=0,row=7,padx=padx,pady=pady)
        e7 = ttk.Entry(master=master,justify='left',width=15)
        e7.grid(column=1,row=7,padx=padx,pady=pady,sticky='ew')



        master.rowconfigure(7)
        master.columnconfigure(2)

        return [e1,e2,e3,e4,e5,e6,e7]

    def Certification_Information(self):
        """
        buyer,lender,title,title ins company,flood zone info, general notes
        """
        padx    = 10
        pady    = 10
        master  = self.tab4
        title   = ttk.Label(master=master,text='Certification Information Input\n',justify='center')
        title.grid(column=1,row=0)

        l1 = ttk.Label(master=master,text='Buyer',justify='left')
        l1.grid(column=0,row=1,padx=padx,pady=pady)
        e1 = ttk.Entry(master=master,justify='left',width=15)
        e1.grid(column=1,row=1,padx=padx,pady=pady,sticky='ew')

        l2 = ttk.Label(master=master,text='Lender',justify='left')
        l2.grid(column=0,row=2,padx=padx,pady=pady)
        e2 = ttk.Entry(master=master,justify='left',width=15)
        e2.grid(column=1,row=2,padx=padx,pady=pady,sticky='ew')

        l3 = ttk.Label(master=master,text='Title',justify='left')
        l3.grid(column=0,row=3,padx=padx,pady=pady)
        e3 = ttk.Entry(master=master,justify='left',width=35)
        e3.grid(column=1,row=3,padx=padx,pady=pady,sticky='ew')

        l4 = ttk.Label(master=master,text=f'Title Insurance Company',justify='left')
        l4.grid(column=0,row=4,padx=padx,pady=pady)
        e4 = ttk.Entry(master=master,justify='left',width=15)
        e4.grid(column=1,row=4,padx=padx,pady=pady,sticky='ew')
        
        l5 = ttk.Label(master=master,text='Flood Zone Information',justify='left')
        l5.grid(column=0,row=5,padx=padx,pady=pady)
        e5 = ttk.Entry(master=master,justify='left',width=15)
        e5.insert(0,'X FEMA')
        e5.grid(column=1,row=5,padx=padx,pady=pady,sticky='ew')

        l6 = ttk.Label(master=master,text='General Notes',justify='left')
        l6.grid(column=0,row=6,padx=padx,pady=pady)
        e6 = tk.Text(master=master,width=15,height=20)
        e6.grid(column=1,row=6,padx=padx,pady=pady,sticky='ew')



        master.rowconfigure(6)
        master.columnconfigure(2)

        return [e1,e2,e3,e4,e5],e6

    def run(self):
        fi      = self.File_Information()
        ci      = self.Contact_Information()
        pi      = self.Property_Information()
        certs,cert_t   = self.Certification_Information()

        master  = self.tab5
        padx    = 10
        pady    = 10      

        l       = ttk.Label(master=master,text='Select Excel file')
        l.grid(column=0,row=0,padx=padx,pady=pady)
        file_name = ttk.Entry(master=master,width=35)
        file_name.grid(column=1,row=1,padx=padx,pady=pady)

        def select_file():
            file_name.insert(0,askopenfilename())

        btn = ttk.Button(master=master,command=select_file,text='Select File')
        btn.grid(column=0,row=1,padx=padx,pady=pady)

        d = ttk.Label(master=master,text='Select Folder to Save Files')
        d.grid(column=0,row=2,padx=padx,pady=pady)
        dir_name = ttk.Entry(master=master,width=35)
        dir_name.grid(column=1,row=3,padx=padx,pady=pady)

        def select_dir():
            dir_name.insert(0,askdirectory())

        btn = ttk.Button(master=master,command=select_dir,text='Select Folder')
        btn.grid(column=0,row=3,padx=padx,pady=pady)


        def submit():
            '''
            Submitting button which should handle getting the information from fi,ci,pi,certs and inserting it into excel and word doc
            '''
            ## fi,ci,pi,certs
            directory = Path(dir_name.get())
            ## file_name, dir_name
            if os.path.exists(directory):
                FI      = [x.get() for x in fi]         # phone number, email, mailing address, file number, cost estimate
                PI      = [x.get() for x in pi]         # parcel #, address, subdivision, legal, lot&blk, tax maps, acres
                CERT    = [x.get() for x in certs]      # buyer, lender, title, title ins company, flood zone info, general notes
                NOTES   = cert_t.get('1.0','end-1c')
                CI      = [x.get() for x in ci]         # requested by, customer, phone number, email, date requested, delever by, closing date
                docx_path = os.path.join(directory,'File ' + str(FI[3]))
                                    
                # create/save word doc
                try:
                    os.mkdir(docx_path)
                except BaseException:
                    pass
                
                new_document(doc_path=docx_path,doc_name=f'File - {FI[3]}.docx').input(b_address=FI[2],b_email=FI[1],b_phone=FI[0],f_num=FI[3],estimate=FI[4],
                                            request=CI[0],cust=CI[1],c_phone=CI[2],c_email=CI[3],request_date=CI[4],deliver=CI[5],closing_date=CI[6],
                                            parcel=PI[0],p_address=PI[1],subdivision=PI[2],legal=PI[3],buyer=CERT[0],lender=CERT[1],title=CERT[2],title_ins=CERT[3],flood_zone=CERT[4],
                                            notes=NOTES)
                # update excel
                if file_name.get() == '' or os.path.splitext(file_name.get())[1] != '.xlsx':
                    messagebox.showinfo('Info', 'Excel file not chosen, creating a file and saving data at directory.')
                    excel_file = f'File - {FI[3]}.xlsx'
                else:
                    excel_file = file_name.get()

                Excel(excel_file).add_to_excel(job_num=FI[3],parcel=PI[0],address=PI[1],customer=CI[1],subdivision=PI[2],lot=PI[4],tax_map=PI[5],acres=PI[len(PI)-1],price=FI[4],dueDate=CI[5],dateOrdered=CI[4],contactEmail=CI[3],contactPhone=CI[2],Notes=NOTES)

                # create/save label
                label(docx_path).create_label(job_num=FI[3],subdivision=PI[2],lot=PI[4],customer=CI[1])

            else:
                messagebox.showerror(title='Error',message='Invalid saving directory. Please select a valid folder to save the files.')

        
        t = '_______________________________________________________'
        l2 = ttk.Label(master=master,text=t)
        l2.grid(column=0,row=4)
        l3 = ttk.Label(master=master,text=t)
        l3.grid(column=1,row=4)
        sub = ttk.Button(master=master,command=submit,text='Submit')
        sub.grid(column=0,row=5,padx=padx,pady=pady)


        ### If the file is not selected, we should output into a formatted excel doc
            ## else output to the given doc.
        master.rowconfigure(4)
        master.columnconfigure(2)
        self.mainloop()


if __name__ == '__main__':
    app = display()
    app.run()
