
from PyQt5.QtCore import QThread,pyqtSignal
from createPDF import *

class Create_Invoices(QThread):
   
    def __init__ (self,invoices_in_page,input_file,output_file,company_name,phone,codes,logo,signature,month1,year):
        super().__init__()
        self.input_file=input_file
        self.month=month1
        self.year=year
        self.company_name=company_name
        self.logo=logo
        self.signature=signature
        self.phone=phone
        self.output_file=output_file
        self.invoices_in_page=invoices_in_page
        self.codes=codes
    
    value=0
    value_changed = pyqtSignal(float)
    error=pyqtSignal(str)
    info=pyqtSignal(str)
    
    def run(self):
            try:
                if(self.invoices_in_page==3):
                    createPDF_3(self,self.input_file,self.output_file,self.company_name,self.phone,self.codes,self.logo,self.signature,self.month,self.year)
                elif(self.invoices_in_page==4):
                    createPDF_4(self,self.input_file,self.output_file,self.company_name,self.phone,self.codes,self.logo,self.signature,self.month,self.year)
                return
            except Exception as e:
                self.error.emit(str(e)) 
                return   


    