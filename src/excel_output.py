import pandas as pd
## requires odfpy for ods files, openpyxl for excel
# document = pd.read_excel('Excel Example.ods', engine='odf')


class __funcs:
    def __init__(self) -> None:
        pass

class Excel(__funcs):
    def __init__(self,document:str) -> None:
        super().__init__()
        self.doc = document

    def add_to_excel(self,job_num,parcel,address,customer,subdivision,lot,tax_map,acres,dateOrdered,price,dueDate,contactEmail,contactPhone,Notes):
        vals    = [job_num,parcel,address,customer,subdivision,lot,tax_map,acres,dateOrdered,price,dueDate,contactEmail,contactPhone,Notes]
        df      = pd.DataFrame(columns=['JOB NUMBER','PARCEL','ADDRESS','CUSTOMER','SUBDIVISION NAME','LOT & BLK','TAX MAPS','ACRES','DATE ORDERED','PRICE','DUE DATE','CONTACT EMAIL','CONTACT PHONE','NOTES'])

        df.loc[len(df)] = vals
        try:
            d = pd.read_excel(self.doc,sheet_name='Sheet1')
            if len(d) > 0:
                d = d.drop(d[d['JOB NUMBER'] == job_num].index)
            data = pd.concat([df,d])
            data.to_excel(self.doc,sheet_name='Sheet1',index=False)
        except BaseException:
            df.to_excel(self.doc,sheet_name='Sheet1',index=False)



