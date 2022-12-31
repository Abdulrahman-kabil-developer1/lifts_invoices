import os
import sys
import pandas as pd
import arabic_reshaper
from bidi.algorithm import get_display
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import QtWidgets
import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from UI import Ui_MainWindow as ui
pdfmetrics.registerFont(TTFont('Arabic', 'Janna LT Bold.ttf'))

"""
    Abdulrahman Ragab Abdullah 
    Cairo-EG
    ********Contact********
    abdulrahman.ragab.kabil@gmail.com
    (+20) 1280059456 - 1149312512
    https://github.com/Abdulrahman-Kabil-developer1
    https://www.linkedin.com/in/abdulrahman-kabil-5729621a2/
    
"""
############ convert int to text ############

def first_to_text(first):
    if first==1:
        return " و واحد"
    elif first==2:
        return " و اثنان"
    elif first==3:
        return " و ثلاثة"
    elif first==4:
        return " و أربعة"
    elif first==5:
        return " و خمسة"
    elif first==6:
        return " و ستة"
    elif first==7:
        return " و سبعة"
    elif first==8:
        return " و ثمانية"
    elif first==9:
        return " و تسعة"
    elif first==0:
        return ""
def second_to_text(second):
    if(second==1):
        return " و عشر"
    elif(second==2):
        return " و عشرون"
    elif(second==3):
        return " و ثلاثون"
    elif(second==4):
        return " و اربعون"
    elif(second==5):
        return " و خمسون"
    elif(second==6):
        return " و ستون"
    elif(second==7):
        return " و سبعون"
    elif(second==8):
        return " و ثمانون"
    elif(second==9):
        return " و تسعون"
    elif(second==0):
        return ""  
def third_to_text(third):
    if(third==1):
        return " و مائة"
    elif(third==2):
        return " و مائتان"
    elif(third==3):
        return " و ثلاثمائة"
    elif(third==4):
        return " و اربعمائة"
    elif(third==5):
        return " و خمسمائة"
    elif(third==6):
        return " و ستمائة"
    elif(third==7):
        return " و سبعمائة"
    elif(third==8):
        return " و ثمانمائة"
    elif(third==9):
        return " و تسعمائة"
    elif(third==0):
        return  ""
def fourth_to_text(fourth):
    if(fourth==1):
        return("ألف")
    elif(fourth==2):
        return("ألفان")
    elif(fourth==3):
        return("ثلاثة ألاف")
    elif(fourth==4):
        return("اربعة ألاف")
    elif(fourth==5):
        return("خمسة ألاف")
    elif(fourth==6):
        return("ستة ألاف")
    elif(fourth==7):
        return("سبعة ألاف")
    elif(fourth==8):
        return("ثمانية ألاف")
    elif(fourth==9):
        return("تسعة ألاف")
    elif(fourth==0):
        return ""
def check_result(result):
    #edit convert result
    if result[0]==" "and result[1]=="و" and result[2]==" ":
        return result[3:]
    else:
        return result
def int_to_text(num):
    """convert int to text

    Args:
        num (string): number to convert beetwen 1 to 9999

    Returns:
        string: number in Arabic text
    """
    num=int(num)
    first=num%10
    second=int(num/10%10)
    third=int(num/100%10)
    fourth=int(num/1000%10)
    if len(str(num))==1:
        result=str( first_to_text(first))+" جنية لاغير"
        result= check_result(result)
        return "فقط "+result
    elif len(str(num))==2:
        result=str( first_to_text(first)) +str( second_to_text(second))+" جنية لاغير"
        result= check_result(result)
        return "فقط "+result
    elif len(str(num))==3:
        result=str( third_to_text(third)) +str( first_to_text(first))+str( second_to_text(second))+" جنية لاغير"
        result= check_result(result)
        return "فقط "+result
    elif len(str(num))==4:
        result=str( fourth_to_text(fourth))+str( third_to_text(third))+str( first_to_text(first))+str( second_to_text(second))+" جنية فقط لاغير"
        result= check_result(result)
        return "فقط "+result 
    else :
        result= "خطأ"
        return result

#############################################
#check if month 30 or 31 days   
def checkMonth(month):
    """check if month 30 or 31 days

    Args:
        month (string): month number

    Returns:
        int: num of days of (month)
    """
    if month==1 or month==3 or month==5 or month==7 or month==8 or month==10 or month==12:
        return 31
    elif month==4 or month==6 or month==9 or month==11:
        return 30
    else:
        return 28

def arabic_text(text):
    """convert inverse text to readabel text

    Args:
        text (string): text to convert

    Returns:
        string: text in arabic readabel text
    """
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    return bidi_text

class create_month_Thread(QThread):
    excel_file=""
    month=""
    year=""
    company_name=""
    logo=""
    phone=""
    output_file=""
    num_per_page=""
    codes=""
    value=0
    value_changed = pyqtSignal(float)
    error=pyqtSignal(str)
    info=pyqtSignal(str)
    num_columns=7
    def validate_file(self,dataFrame):
        if (len(dataFrame.columns)!=self.num_columns):
            self.error.emit("خطأ في ملف البيانات")
            return False
        dataTypes=["int64","object","int64","object","object","object","object"]
        for i in range(self.num_columns):
            if str(dataFrame.dtypes[i])!=dataTypes[i]:
                self.error.emit(" خطأ في ملف البيانات العمود رقم "+str(i+1)+" يجب ان يكون من النوع"+dataTypes[i])
                return False
        return True

    def createPDF_3(self,excel_file,output_file,company_name,phone,codes,logo,month1,year):
        """
            create invoices pdf file from excel file using "reportlab" library (3 invoices per page)
            
            Args:
                excel_file (str): excel file path
                output_file (str): pdf file path (result)
                company_name (str): company name
                phone (str): company phone
                logo (str): company logo
                month1 (str): month
                year (str): year
        
        """
        year=str(year)
        process=pd.read_excel(excel_file)
        if (self.validate_file(process))==False:
            return
        invoice_width = 9.9*cm
        # my_path='Doc1.pdf'# file path
        c = canvas.Canvas(output_file,bottomup=1,pagesize=A4)
        count=0 #count num of created invoice in current page
        count2=0 #count num of created invoices till now used to set progress bar value
        dont_save=1
        c.setFont('Arabic', 14)
        c.translate(cm,cm) # make unite cm
        # c.setStrokeColorRGB(1,0,0) # red colour of line
        
        # create page structure
        c.setLineWidth(1.5)#width of the line
        c.setLineCap(1)
        c.setDash(3,6)#dashed line
        c.line(-1*cm,18.8*cm,22*cm,18.8*cm)
        c.line(-1*cm,8.9*cm,22*cm,8.9*cm)
        c.line(7*cm,-1*cm,7*cm,29.7*cm)
        
        if codes != "":
            codes_list=codes.split("-")
            try:
                codes_list=[int(item) for item in codes_list]
                if len(codes_list)>0:
                    found_no_print=process.loc[(process.serial.isin(codes_list)) & (process['print or no']=='لا')]
                    process=process.loc[(process.serial.isin(codes_list)) & (process['print or no']=='نعم') ]
                    num_process=len(process)
                    not_found_codes=list(set(codes_list)-(set(process.serial).union(set(found_no_print.serial))))
                    message=''
                    if num_process>0:
                        message='هذة الاكواد سيتم طباعتها \n'+str(list(process.serial))
                    if len(found_no_print)>0:
                        message+="\nهذة الاكواد تم العثور عليها في الملف ولكنها لم تمنح اذن الطباعة \n"+str(list(found_no_print.serial))
                    if len(not_found_codes)>0:
                        message+="\n هذة الاكواد لم يتم العثور عليها في الملف \n"+str(list(not_found_codes))+"\n"
                    self.info.emit(message)      
                if num_process==0:
                    self.error.emit("لا يوجد اكواد يمكن طباعتها")
                    return
            except:
                self.info.emit(" خطأ في رقم الفاتورة يجب ان يكون ارقام صحيحة فقط وليست حروف")
                return
        else:
            process=process[process['print or no'] == 'نعم']
            num_process=len(process)
            
        for proc in range (0,len(process)):
            proc_serial=str(process.iloc[proc,0])
            proc_name=str(process.iloc[proc].values[1])
            proc_price=str(process.iloc[proc].values[2])
            proc_status=str(process.iloc[proc].values[3])
            
            if (logo!=""):
                #draw logo
                c.drawImage(logo, 7.5*cm, 19.5*cm-(count*invoice_width), width=12*cm, height=8.5*cm,preserveAspectRatio=True, mask='auto')
            if(proc_status=="مؤخر"): 
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            else:
                month=(int(month1)+1)
                if(month==13):
                    month=1
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            cur_r=28.25  #current Row to write in 
            ############################################
            text=arabic_text('إيصال صيانة')
            c.drawCentredString(13.5*cm,(cur_r-0.1)*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text('نسخة إيصال')
            c.drawCentredString(3*cm,(cur_r-0.1)*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            if (company_name!=""):
                #if company name is not empty
                text=arabic_text(company_name)
                c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
                c.drawCentredString(3*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=1
            
            ############################################
            text=arabic_text("رقم : ")
            c.drawRightString(17.5*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            #invoice number
            c.drawRightString(16.2*cm,cur_r*cm-(count*invoice_width),proc_serial)
            c.drawRightString(6*cm,(cur_r-1)*cm-(count*invoice_width),proc_serial)
            ############################################
            text=arabic_text("المبلغ : ")
            c.drawRightString(12*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("جــــــــ")
            c.setFont('Arabic', 13)
            c.drawString(9.5*cm,(cur_r+0.5)*cm-(count*invoice_width),text)
            c.setFont('Arabic', 14)
            c.drawRightString(10.3*cm,cur_r*cm-(count*invoice_width),proc_price)
            c.drawRightString(1*cm,(cur_r-1)*cm-(count*invoice_width),proc_price)
            cur_r-=1
            ############################################
            text=arabic_text("وصلنا من السيد/ ")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text(proc_name)
            c.drawRightString(16.3*cm,cur_r*cm-(count*invoice_width),text)
            if(len(proc_name)>30):
                c.drawRightString(6.8*cm,(cur_r+1)*cm-(count*invoice_width),text)
            else:
                c.drawCentredString(3*cm,(cur_r+1)*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            text=arabic_text("مبلغ و قدرة : ")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text(int_to_text(proc_price))
            c.drawRightString(17*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            text=arabic_text("وذلك قيمة : صيانة المصعد من 1 - "+days+" / "+month+" / "+year)
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            text=arabic_text("تحريراً في : "+days+" / "+month+" / "+year)
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            text=arabic_text("توقيع المستلم -------------------------")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text) 
            cur_r-=1
            ############################################
            if (company_name==""):
                cur_r-=1
            if (phone!=""):
                text=arabic_text("للتواصل : "+phone)
                c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            else:
                cur_r-=1
            ############################################
            text=arabic_text("تاريخ السداد : "+"   "+" / "+"  "+" / "+"    20  ")
            c.drawRightString(6*cm,24.25*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("ملاحظات")
            c.drawCentredString(3*cm,23.25*cm-(count*invoice_width),text)
            count=(count+1)%3
            count2+=1
            self.value=(count2/num_process)*100
            self.value_changed.emit(self.value)
            if (count)%3==0 and proc+1!=len(process):   #if it is the last invoice in the page
                c.showPage()    #create new page
                c.setFont('Arabic', 14)
                c.translate(cm,cm) 
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1)
                c.setDash(3,6)
                c.line(-1*cm,18.8*cm,22*cm,18.8*cm)
                c.line(-1*cm,8.9*cm,22*cm,8.9*cm)
                c.line(7*cm,-1*cm,7*cm,29.7*cm)
            dont_save=0
            
        if dont_save==0:
            c.save()
    
    def createPDF_4(self,input_file,output_file,company_name,phone,codes,logo,month1,year):
        """
            create invoices pdf file from excel file using "reportlab" library (4 invoices per page)
            
            Args:
                excel_file (str): excel file path
                output_file (str): pdf file path (result)
                company_name (str): company name
                phone (str): company phone
                codes (str) : if user want only create some process by serial code
                logo (str): company logo
                month1 (str): month
                year (str): year
        
        """
        year=str(year)
        process=pd.read_excel(input_file)
        if (self.validate_file(process))==False:
            return
        num_process=(process['print or no'] == 'نعم').sum()
        invoice_width = 7.425*cm
        # my_path='Doc1.pdf'# file path
        c = canvas.Canvas(output_file,bottomup=1,pagesize=A4)
        count=0 #count num of created invoice in current page
        count2=0 #count num of created invoices till now used to set progress bar value
        dont_save=1
        c.setFont('Arabic', 14)
        c.translate(cm,cm) #starting point of coordinate to one inch
        # c.setStrokeColorRGB(1,0,0) # red colour of line
        c.setLineWidth(1.5)#width of the line
        c.setLineCap(1)
        c.setDash(3,6)#dashed line
        c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
        c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
        c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
        c.line(7*cm,-1*cm,7*cm,29.7*cm)
        if codes != "":
            codes_list=codes.split("-")
            try:
                codes_list=[int(item) for item in codes_list]
                if len(codes_list)>0:
                    found_no_print=process.loc[(process.serial.isin(codes_list)) & (process['print or no']=='لا')]
                    process=process.loc[(process.serial.isin(codes_list)) & (process['print or no']=='نعم') ]
                    num_process=len(process)
                    not_found_codes=list(set(codes_list)-(set(process.serial).union(set(found_no_print.serial))))
                    message=''
                    if num_process>0:
                        message='هذة الاكواد سيتم طباعتها \n'+str(list(process.serial))
                    if len(found_no_print)>0:
                        message+="\nهذة الاكواد تم العثور عليها في الملف ولكنها لم تمنح اذن الطباعة \n"+str(list(found_no_print.serial))
                    if len(not_found_codes)>0:
                        message+="\n هذة الاكواد لم يتم العثور عليها في الملف \n"+str(list(not_found_codes))+"\n"
                    self.info.emit(message)   
                if num_process==0:
                    self.error.emit("لا يوجد اكواد يمكن طباعتها")
                    return
            except:
                self.info.emit(" خطأ في رقم الفاتورة يجب ان يكون ارقام صحيحة فقط وليست حروف")
                return
        else:
            process=process[process['print or no'] == 'نعم']
            num_process=len(process)
        for proc in range (0,len(process)):
            proc_serial=str(process.iloc[proc,0])
            proc_name=str(process.iloc[proc].values[1])
            proc_price=str(process.iloc[proc].values[2])
            proc_status=str(process.iloc[proc].values[3])
            if logo!="":
                c.drawImage(logo, 7.5*cm, 22*cm-(count*invoice_width), width=12*cm, height=6.025*cm,preserveAspectRatio=True, mask='auto')
            if(proc_status=="مؤخر"):
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            else:
                month=(int(month1)+1)
                if(month==13):
                    month=1
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            cur_r=28.25    #current Row 
            ############################################
            text=arabic_text('إيصال صيانة')
            c.drawCentredString(13.5*cm,(cur_r-0.1)*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text('نسخة إيصال')
            c.drawCentredString(3*cm,(cur_r-0.1)*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            if (company_name!=""):
                text=arabic_text(company_name)
                c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
                c.drawCentredString(3*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=0.7
            
            ############################################
            text=arabic_text("رقم : ")
            c.drawRightString(17.5*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            #invoice number
            c.drawRightString(16.2*cm,cur_r*cm-(count*invoice_width),proc_serial)
            #draw string with border
            #c.drawBoundary(10,10*cm,18*cm,3*cm,3*cm)
            c.drawRightString(6*cm,(cur_r-1)*cm-(count*invoice_width),proc_serial)
            ############################################
            text=arabic_text("المبلغ : ")
            c.drawRightString(12*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("جــــــــ")
            c.setFont('Arabic', 13)
            c.drawString(9.5*cm,(cur_r+0.5)*cm-(count*invoice_width),text)
            c.setFont('Arabic', 14)
            c.drawRightString(10.3*cm,cur_r*cm-(count*invoice_width),proc_price)
            c.drawRightString(1*cm,(cur_r-1)*cm-(count*invoice_width),proc_price)
            cur_r-=0.7
            ############################################
            text=arabic_text("وصلنا من السيد/ ")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text(proc_name)
            c.drawRightString(16.3*cm,cur_r*cm-(count*invoice_width),text)
            if(len(proc_name)>30):
                c.drawRightString(6.8*cm,(cur_r+0.7)*cm-(count*invoice_width),text)
            else:
                c.drawCentredString(3*cm,(cur_r+0.7)*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("مبلغ و قدرة : ")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text(int_to_text(proc_price))
            c.drawRightString(17*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("وذلك قيمة : صيانة المصعد من 1 - "+days+" / "+month+" / "+year)
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("تحريراً في : "+days+" / "+month+" / "+year)
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("توقيع المستلم -------------------------")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text) 
            cur_r-=0.7
            ############################################
            if (company_name==""):
                cur_r-=0.7
            text=arabic_text("للتواصل : "+phone)
            c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("تاريخ السداد : "+"   "+" / "+"  "+" / "+"    20  ")
            c.drawRightString(6*cm,24.25*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("ملحوظات")
            c.drawCentredString(3*cm,23.25*cm-(count*invoice_width),text)
            
            count=(count+1)%4
            count2+=1
            self.value=(count2/num_process)*100
            self.value_changed.emit(self.value)
            if (count)%4==0 and proc+1!=len(process):
                #create new page
                c.showPage()
                c.setFont('Arabic', 14)
                c.translate(cm,cm) #starting point of coordinate to one inch
                # c.setStrokeColorRGB(1,0,0) # red colour of line
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1)
                c.setDash(3,6)#dashed line
                c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
                c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
                c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
                c.line(7*cm,-1*cm,7*cm,29.7*cm)
                #create image as watermark
            dont_save=0
            
        if dont_save==0:
            c.save()

    def run(self):
        #run self.createPDF_3
        try:
            if(self.num_per_page==3):
                self.createPDF_3(self.excel_file,self.output_file,self.company_name,self.phone,self.codes,self.logo,self.month,self.year)
            elif(self.num_per_page==4):
                self.createPDF_4(self.excel_file,self.output_file,self.company_name,self.phone,self.codes,self.logo,self.month,self.year)
            return
        except Exception as e:
            self.error.emit(str(e)) 
            return   
####################################################################
def update_last_dir(dir):
    with open("last_dir.txt", 'w',encoding='utf-8') as f:
        f.write(dir)
def get_last_dir():
    try:
        with open("last_dir.txt",'r',encoding='utf-8') as f:
            last_dir=f.read()
            if os.path.exists(last_dir):
                return last_dir
            else:
                return os.getcwd()
    except:
        return os.getcwd()
        
class Main(QMainWindow, ui):
    """a class for the main window"""
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handel_buttons()
        self.UI_changes()
        self.setFixedSize(self.size())#stop the maximize button
        #self.setWindowIcon(QIcon("img/receipt.png"))
    current_month=datetime.datetime.now().strftime("%m")
    current_year=datetime.datetime.now().strftime("%Y")
    def UI_changes(self):
        """changes in UI like hiding the title bar
        """
        self.comboBox.setCurrentIndex(int(self.current_month)-1)
        self.lineEdit_4.setText(self.current_year)
    def handel_buttons(self):
        """
            connect buttons in GUI with methods
        """
        self.pushButton.clicked.connect(self.choose_file_excel)
        self.pushButton_13.clicked.connect(self.choose_file_image)
        self.pushButton_4.clicked.connect(self.choose_save)
        self.pushButton_2.clicked.connect(self.create_month_receipts)
    def clear_create_month_receipts(self):
        """clear the create month receipts tab data
        """
        self.lineEdit.setText("")
        self.lineEdit_2.setText("")
        self.lineEdit_11.setText("")
        self.lineEdit_13.setText("")
        self.lineEdit_27.setText("")
        self.comboBox.setCurrentIndex(int(self.current_month)-1)
        self.comboBox_5.setCurrentIndex(0)
        self.lineEdit_4.setText(self.current_year)
    def choose_file_excel(self):
        """open a file dialog to choose the excel file
        """
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=get_last_dir(),filter="Excel (*.xlsx)")
        if file:
            self.lineEdit.setText(file)
            filename=file.split("/")[-1]
            update_last_dir(file.replace(filename,""))
        if file==None:
            QMessageBox.warning(self,"Error","يجب إختيار ملف")
            return   
    def choose_file_image(self):
        """open a file dialog to choose the image file"""
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=get_last_dir(),filter="Image (*.png *.jpg)")
        if file:
            self.lineEdit_27.setText(file)
        if file==None:
            QMessageBox.warning(self,"Error","يجب إختيار ملف")
            return             
    def choose_save(self):
        """open a file dialog to choose the save directory"""
        if self.lineEdit.text()=='': #if no file selected
            QMessageBox.warning(self,"Error","يجب إختيار ملف العمارات اولاً")
            return
        if self.lineEdit_4.text()=='': #if no year
            QMessageBox.warning(self,"Error","يجب إدخال السنة")
            return
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        newFileName="ايصالات صيانة المصاعد"+self.comboBox.currentText()+"-"+self.lineEdit_4.text()
        fileName, _ = QFileDialog.getSaveFileName(self,"Save File",directory=get_last_dir()+"\\"+newFileName,filter="PDF Files (*.pdf)", options=options)
        if fileName:
            self.lineEdit_2.setText(fileName)
            filename=fileName.split("/")[-1]
            update_last_dir(fileName.replace(filename,""))
        else:  
            return         
    def create_month_receipts(self):
        if self.comboBox_5.currentIndex()==0:
            invoices_in_page=3
        else:
            invoices_in_page=4
        if (self.lineEdit.text()==''):
            QMessageBox. warning(self, "ERROR", "يجب اختيار ملف العمارات اولاً!")
            return
        if (self.lineEdit_4.text()==''):
            QMessageBox. warning(self, "ERROR", "يجب ادخال السنة!")
            return
        if (self.lineEdit_13.text()==''):
            #create qustion Message no phone "لا " or "اضافة هاتف"
            msg=QMessageBox()
            msg.setText("هل تريد إضافة رقم الهاتف للايصالات؟")
            msg.setWindowTitle('تنبيه')
            msg.setIcon(QMessageBox.Question)
            msg.addButton("لا", QMessageBox.NoRole)
            msg.addButton("اضافة هاتف", QMessageBox.YesRole)
            msg.setWindowIcon(QIcon("receipt.png"))
            replay=msg.exec_()
            if replay==1:
                return
        if (self.lineEdit_2.text()==''):
            QMessageBox. warning(self, "ERROR", "يجب اختيار مكان حفظ الملف!")
            return
        logo=self.lineEdit_27.text()
        if (logo!=''):
            if not os.path.exists(logo):
                QMessageBox. warning(self, "ERROR", "الصورة غير موجودة!")
                return
        codes=self.lineEdit_14.text()
        input_file=self.lineEdit.text()
        if not os.path.exists(input_file):
            QMessageBox. warning(self, "ERROR", "ملف العمارات غير موجود!")
            return
        company_name=self.lineEdit_11.text()
        year=self.lineEdit_4.text()
        month= self.comboBox.currentText()
        phone=self.lineEdit_13.text()
        output_file=self.lineEdit_2.text()+".pdf"
        try:
            self.calc=create_month_Thread()
            self.calc.excel_file=input_file
            self.calc.year=year
            self.calc.month=month
            self.calc.company_name=company_name
            self.calc.phone=phone
            self.calc.logo=logo
            self.calc.output_file=output_file
            self.calc.codes=codes
            self.calc.value_changed.connect(self.update_progress)
            self.calc.error.connect(self.show_error)
            self.calc.info.connect(self.show_info)
            if (invoices_in_page==3):
                self.calc.num_per_page=3
            else:
                self.calc.num_per_page=4
            self.calc.start()
            self.pushButton_2.setEnabled(False)
            # createPDF_4(input_file,output_file,company_name,phone,logo,month,year)
            #create_all_receipts(input_file,output_file,company_name,month,year)
        except Exception as e:
            self.pushButton_2.setEnabled(True)
            self.show_error(e)
            self.update_progress(0)
            return
    def show_error(self,msg):
        QMessageBox.warning(self, "ERROR","لقد وجدنا هذة الاخطاء:\n"+str(msg))
        self.clear_create_month_receipts()
        self.pushButton_2.setEnabled(True)
        self.update_progress(0)
        return
    def show_info(self,msg):
        self.pushButton_2.setEnabled(True)
        QMessageBox.information(self, "info","نود اعلامك بانة:\n"+str(msg))   
    def update_progress(self,value):
        try:
            self.progressBar.setValue(int(value))
            self.progressBar.setFormat(("%.02f %%" % value))
            if value==100:
                QMessageBox.information(self, "Success", "تم إنشاء الملف بنجاح!")
                self.progressBar.setValue(0)
                self.lineEdit_2.setText("")
                self.pushButton_2.setEnabled(True)
                return
            #print error in qmessage
        except Exception as e:
            self.pushButton_2.setEnabled(True)
            QMessageBox.warning(self, "ERROR", "لقد وجدنا هذة الاخطاء:\n"+str(e))
            self.update_progress(0)
            return
            
        
def main():
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
