import os
import sys
import pandas as pd
from fpdf import FPDF
import arabic_reshaper
from bidi.algorithm import get_display
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.uic import loadUiType
from PyQt5 import QtWidgets
import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('Arabic', 'Janna LT Bold.ttf'))

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

def create_all_receipts(input_file,output_file,company_name,month1,year):
    """create all the receipts for the given month and year using FPDF library

    Args:
        input_file (string): excel file dir
        output_file (string): pdf file dir
        company_name (string): company name that will appear in the receipts
        month1 (string):   month number
        year (string): year number
    """
    year=str(year)
    process=pd.read_excel(input_file)
    pdf=FPDF()
    pdf=FPDF('P','mm','A4')
    pdf.add_page()
    pdfW=210
    pdfH=297
    pdf.add_font("Janna", "", "Janna LT Bold.ttf", uni=True)
    pdf.set_font('Janna','',10)
    
    
    ###############################################################################
    i=1
    pdf.set_y(0)
    for proc in range(1,len(process)+1):
        proc_name=str(process.iloc[proc-1].values[1])
        proc_price=str(process.iloc[proc-1].values[2])
        proc_status=str(process.iloc[proc-1].values[3])
        if(proc_status=="مؤخر"):
            month=int(month1)
            if(len(str(month))==1):
                month="0"+str(month)
            month=str(month)
            proc_serial=str(process.iloc[proc-1].values[0])+month+"/"+year[2:4]
            days=str(checkMonth(int(month)))
        else:
            month=int(month1)+1
            if(len(str(month))==1):
                month="0"+str(month)
            month=str(month)
            proc_serial=str(process.iloc[proc-1].values[0])+month+"/"+year[2:4]
            days=str(checkMonth(int(month)))
        pdf.dashed_line(pdfW/2+3,0,pdfW/2+3,pdfH,2,1)
        
        pdf.set_font('Janna','U',10)
        text=arabic_text("ايصال صيانة")
        pdf.cell(88,7,text,0,0,'C')
        text=arabic_text("نسخة ايصال صيانة")
        pdf.cell(120,7,text,0,1,'C')
        pdf.set_font('Janna','',10)
        text=arabic_text("    ")
        pdf.cell(10,7,text,0,0,'R')
        ###############################################################################
        if (company_name==""):
            text=arabic_text(" ")
            pdf.cell(10,7,text,0,1,'R')
        else:
            pdf.set_font('Janna','',10)
            text=arabic_text(company_name)
            pdf.cell(70,7,text,0,0,'C')
            text=arabic_text(company_name)
            pdf.cell(137,7,text,0,1,'C')
            pdf.set_font('Janna','',10)
        ###############################################################################
        pdf.set_font('Janna','U',10)
        text=arabic_text(str(proc_price)+" جنية")
        pdf.cell(20,7,text,0,0,'R')
        
        pdf.set_font('Janna','',10)
        text=arabic_text("المبلغ: ")
        pdf.cell(14,7,text,0,0,'R')
        
        pdf.set_font('Janna','U',10)
        text=str(proc_serial)
        pdf.cell(30,7,text,0,0,'R')
        
        pdf.set_font('Janna','',10)
        text=arabic_text("رقم الايصال: ")
        pdf.cell(24,7,text,0,0,'R')
        
        pdf.set_font('Janna','U',10)
        text=arabic_text(str(proc_price)+" جنية")
        pdf.cell(36,7,text,0,0,'R')
        
        pdf.set_font('Janna','',10)
        text=arabic_text("المبلغ: ")
        pdf.cell(14,7, text,0,0,'R')
        
        pdf.set_font('Janna','U',10)
        text=proc_serial
        pdf.cell(32,7,text,0,0,'R')
        
        pdf.set_font('Janna','',10)
        text=arabic_text("رقم الايصال: ")
        pdf.cell(24,7, text,0,1,'R')
        ###############################################################################
        pdf.set_font('Janna','U',10)
        text=arabic_text(proc_name)
        pdf.cell(68,7, text,0,0,'R')
        pdf.set_font('Janna','',10)
        text=arabic_text("وصلنا من السيد/")
        pdf.cell(30,7, text,0,0)
        pdf.set_font('Janna','U',10)
        text=arabic_text(proc_name)
        pdf.cell(70,7, text,0,0,'R')
        pdf.set_font('Janna','',10)
        text=arabic_text("وصلنا من السيد/" )
        pdf.cell(30,7, text,0,1)
        ###############################################################################
        pdf.set_font('Janna','U',10)
        text=arabic_text(int_to_text(proc_price))
        pdf.cell(75,7, text,0,0,'R')
        pdf.set_font('Janna','',10)
        text=arabic_text("مبلغ و قدرة: " )
        pdf.cell(20,7, text,0,0,'R')
        pdf.set_font('Janna','U',10)
        text=arabic_text(int_to_text(proc_price))
        pdf.cell(80,7, text,0,0,'R')
        pdf.set_font('Janna','',10)
        text=arabic_text("مبلغ و قدرة: " )
        pdf.cell(20,7, text,0,1,'R')
        ###############################################################################
        pdf.set_font('Janna','',10)
        text=arabic_text(" وذلك قيمة:  صيانة المصعد من 1 الى "+days + " شهر "+month+" سنة "+year)
        pdf.cell(97,7, text,0,0,'R')
        text=text=arabic_text(" وذلك قيمة:  صيانة المصعد من 1 الى "+days+ " شهر "+month+" سنة "+year)
        pdf.cell(100,7, text,0,1,'R')
        ###############################################################################
        pdf.set_font('Janna','U',10)
        text=days+"-"+month+"-"+year
        if proc_status == "مقدم":
            text="01"+"-"+month+"-"+year 
        pdf.cell(75,7,text,0,0,'R')
        pdf.set_font('Janna','',10)
        text=arabic_text("تحريراً في: ")
        pdf.cell(20,7, text,0,0,'R')
        pdf.set_font('Janna','U',10)
        text=days+"-"+month+"-"+year
        if proc_status == "مقدم":
            text="01"+"-"+month+"-"+year 
        pdf.cell(75,7,text,0,0,'R')
        pdf.set_font('Janna','',10)
        text=arabic_text("تحريراً في: ")
        pdf.cell(20,7, text,0,1,'R')
        ###############################################################################
        text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
        pdf.cell(88,4, text,0,0,'C')
        text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
        pdf.cell(120,4, text,0,1,'C')
        ###############################################################################
        pdf.set_font('Janna','',10)
        pdf.cell(70,7,"------------",0,0,'R')
        text=arabic_text("توقيع المستلم")
        pdf.cell(25,7, text,0,0,'C')
        pdf.cell(70,7,"-----------",0,0,'R')
        text=arabic_text("توقيع المستلم")
        pdf.cell(25,7, text,0,1,'C')
        ###############################################################################
        end=70.5
        pdf.dashed_line(0,i*(end),pdfW,i*(end),2,1)
        ###################################################################################
        pdf.ln(12)
        i=i+1
        ###################################################################################
        if proc%4==0:
            page_num=proc/4
            # pdf.set_y(-18)
            # pdf.set_font('Janna','',5)
            # print("pdf.y = ",pdf.get_y())
            # pdf.cell(0, 0, 'Page ' + str(int(page_num)), 0, 0, 'C')
            # print("pdf.y = ",pdf.get_y())
            # print ("page num = ",page_num)
            if (page_num!=len(process)/4):
                pdf.add_page()
            pdf.set_y(0)
            
            i=1

    pdf.output(output_file, 'F')
    pdf.close()

def create_receipt(input_file,output_file,company_name,code,month1,year):
    """create one receipt for the given process code in specifc month and year using FPDF library

    Args:
        input_file (string): excel file dir
        output_file (string): pdf file dir
        company_name (string): company name that will appear in the receipts
        code: process code
        month1 (string):   month number
        year (string): year number
    Returns:
        int: 0 if found process code and created receipt, 1 if not found process code
    """
    year=str(year)
    process=pd.read_excel(input_file)
    pdf=FPDF()
    pdf=FPDF('P','mm','A4')
    pdf.add_page()
    pdfW=210
    pdfH=297
    pdf.add_font("Janna", "", "Janna LT Bold.ttf", uni=True)
    pdf.set_font('Janna','',10)
    h=70
    found=0
    
    ###############################################################################
    i=1
    pdf.set_y(0)
    for proc in range(1,len(process)+1):
        proc_serial=str(process.iloc[proc-1].values[0])
        if proc_serial == code:
            proc_name=str(process.iloc[proc-1].values[1])
            proc_price=str(process.iloc[proc-1].values[2])
            proc_status=str(process.iloc[proc-1].values[3])
            if(proc_status=="مؤخر"):
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc-1].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            else:
                month=int(month1)+1
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc-1].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            pdf.dashed_line(pdfW/2+3,0,pdfW/2+3,pdfH,2,1)
            
            pdf.set_font('Janna','U',10)
            text=arabic_text("ايصال صيانة")
            pdf.cell(88,7,text,0,0,'C')
            text=arabic_text("نسخة ايصال صيانة")
            pdf.cell(120,7,text,0,1,'C')
            pdf.set_font('Janna','',10)
            text=arabic_text("    ")
            pdf.cell(10,7,text,0,0,'R')
            ###############################################################################
            if (company_name==""):
                text=arabic_text(" ")
                pdf.cell(10,7,text,0,1,'R')
            else:
                pdf.set_font('Janna','',10)
                text=arabic_text(company_name)
                pdf.cell(70,7,text,0,0,'C')
                text=arabic_text(company_name)
                pdf.cell(137,7,text,0,1,'C')
                pdf.set_font('Janna','',10)
            ###############################################################################
            pdf.set_font('Janna','U',10)
            text=arabic_text(str(proc_price)+" جنية")
            pdf.cell(20,7,text,0,0,'R')
            
            pdf.set_font('Janna','',10)
            text=arabic_text("المبلغ: ")
            pdf.cell(14,7,text,0,0,'R')
            
            pdf.set_font('Janna','U',10)
            text=str(proc_serial)
            pdf.cell(30,7,text,0,0,'R')
            
            pdf.set_font('Janna','',10)
            text=arabic_text("رقم الايصال: ")
            pdf.cell(24,7,text,0,0,'R')
            
            pdf.set_font('Janna','U',10)
            text=arabic_text(str(proc_price)+" جنية")
            pdf.cell(36,7,text,0,0,'R')
            
            pdf.set_font('Janna','',10)
            text=arabic_text("المبلغ: ")
            pdf.cell(14,7, text,0,0,'R')
            
            pdf.set_font('Janna','U',10)
            text=proc_serial
            pdf.cell(32,7,text,0,0,'R')
            
            pdf.set_font('Janna','',10)
            text=arabic_text("رقم الايصال: ")
            pdf.cell(24,7, text,0,1,'R')
            ###############################################################################
            pdf.set_font('Janna','U',10)
            text=arabic_text(proc_name)
            pdf.cell(68,7, text,0,0,'R')
            pdf.set_font('Janna','',10)
            text=arabic_text("وصلنا من السيد/")
            pdf.cell(30,7, text,0,0)
            pdf.set_font('Janna','U',10)
            text=arabic_text(proc_name)
            pdf.cell(70,7, text,0,0,'R')
            pdf.set_font('Janna','',10)
            text=arabic_text("وصلنا من السيد/" )
            pdf.cell(30,7, text,0,1)
            ###############################################################################
            pdf.set_font('Janna','U',10)
            text=arabic_text(int_to_text(proc_price))
            pdf.cell(75,7, text,0,0,'R')
            pdf.set_font('Janna','',10)
            text=arabic_text("مبلغ و قدرة: " )
            pdf.cell(20,7, text,0,0,'R')
            pdf.set_font('Janna','U',10)
            text=arabic_text(int_to_text(proc_price))
            pdf.cell(80,7, text,0,0,'R')
            pdf.set_font('Janna','',10)
            text=arabic_text("مبلغ و قدرة: " )
            pdf.cell(20,7, text,0,1,'R')
            ###############################################################################
            pdf.set_font('Janna','',10)
            text=arabic_text(" وذلك قيمة:  صيانة المصعد من 1 الى "+days + " شهر "+month+" سنة "+year)
            pdf.cell(97,7, text,0,0,'R')
            text=text=arabic_text(" وذلك قيمة:  صيانة المصعد من 1 الى "+days+ " شهر "+month+" سنة "+year)
            pdf.cell(100,7, text,0,1,'R')
            ###############################################################################
            pdf.set_font('Janna','U',10)
            text=days+"-"+month+"-"+year
            if proc_status == "مقدم":
                text="01"+"-"+month+"-"+year 
            pdf.cell(75,7,text,0,0,'R')
            pdf.set_font('Janna','',10)
            text=arabic_text("تحريراً في: ")
            pdf.cell(20,7, text,0,0,'R')
            pdf.set_font('Janna','U',10)
            text=days+"-"+month+"-"+year
            if proc_status == "مقدم":
                text="01"+"-"+month+"-"+year 
            pdf.cell(75,7,text,0,0,'R')
            pdf.set_font('Janna','',10)
            text=arabic_text("تحريراً في: ")
            pdf.cell(20,7, text,0,1,'R')
            ###############################################################################
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            pdf.cell(88,4, text,0,0,'C')
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            pdf.cell(120,4, text,0,1,'C')
            ###############################################################################
            pdf.set_font('Janna','',10)
            pdf.cell(70,7,"------------",0,0,'R')
            text=arabic_text("توقيع المستلم")
            pdf.cell(25,7, text,0,0,'C')
            pdf.cell(70,7,"-----------",0,0,'R')
            text=arabic_text("توقيع المستلم")
            pdf.cell(25,7, text,0,1,'C')
            ###############################################################################
            end=70.5
            pdf.dashed_line(0,i*(end),pdfW,i*(end),2,1)
            ###################################################################################
            pdf.ln(12)
            i=i+1
            ###################################################################################
            if proc%4==0:
                page_num=proc/4
                # pdf.set_y(-18)
                # pdf.set_font('Janna','',5)
                # print("pdf.y = ",pdf.get_y())
                # pdf.cell(0, 0, 'Page ' + str(int(page_num)), 0, 0, 'C')
                # print("pdf.y = ",pdf.get_y())
                # print ("page num = ",page_num)
                if (page_num!=len(process)/4):
                    pdf.add_page()
                pdf.set_y(0)
                
                i=1
            ###################################################################################
            found=1
            break
        else:
            continue
    if found==0:
        return 0
    else:
        pdf.output(output_file, 'F')
        pdf.close()
        return 1  

#####################################################################

def createPDF_3(excel_file,output_file,company_name,phone,logo,month1,year):
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
    invoice_width = 9.9*cm
    # my_path='Doc1.pdf'# file path
    c = canvas.Canvas(output_file,bottomup=1,pagesize=A4)
    count=0
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
    
    for proc in range (0,len(process)):
        proc_name=str(process.iloc[proc].values[1])
        proc_price=str(process.iloc[proc].values[2])
        proc_status=str(process.iloc[proc].values[3])
        print_permition=str(process.iloc[proc].values[4])
        
        if print_permition=="لا": 
            continue # skip this process if print_permition is "no"
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
            month=int(month1)+1
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
        text=arabic_text("للتواصل : "+phone)
        c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
        ############################################
        text=arabic_text("تاريخ السداد : "+"   "+" / "+"  "+" / "+"    20  ")
        c.drawRightString(6*cm,24.25*cm-(count*invoice_width),text)
        ############################################
        text=arabic_text("ملحوظات")
        c.drawCentredString(3*cm,23.25*cm-(count*invoice_width),text)
        
        count=(count+1)%3
        
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
    c.save()

def createPDF_4(input_file,output_file,company_name,phone,logo,month1,year):
    """
        create invoices pdf file from excel file using "reportlab" library (4 invoices per page)
        
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
    process=pd.read_excel(input_file)
    invoice_width = 7.425*cm
    # my_path='Doc1.pdf'# file path
    c = canvas.Canvas(output_file,bottomup=1,pagesize=A4)
    count=0
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
    cur_p=0
    for proc in range (0,len(process)):
        proc_name=str(process.iloc[proc].values[1])
        proc_price=str(process.iloc[proc].values[2])
        proc_status=str(process.iloc[proc].values[3])
        print_permition=str(process.iloc[proc].values[4])
        
        if print_permition=="لا":
            continue
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
            month=int(month1)+1
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
    c.save()

def createPDF_3_1(excel_file,output_file,company_name,phone,logo,code,month1,year):
    """create invoice of process of this code from excel file using "reportlab" library (3 invoices per page)
        
    Args:
        excel_file (str): excel file path
        output_file (str): pdf file path (result)
        company_name (str): company name
        phone (str): company phone
        logo (str): company logo
        code (str): process code
        month1 (str): month
        year (str): year
    Returns:
        int: 0 if process not found, 1 if process found , 2 if process found but don't have print flag
        
        """
    
    year=str(year)
    process=pd.read_excel(excel_file)
    invoice_width = 9.9*cm
    # my_path='Doc1.pdf'# file path
    c = canvas.Canvas(output_file,bottomup=1,pagesize=A4)
    count=0
    c.setFont('Arabic', 14)
    c.translate(cm,cm)
    c.setLineWidth(1.5)#width of the line
    c.setLineCap(1)
    c.setDash(3,6)#dashed line
    c.line(-1*cm,18.8*cm,22*cm,18.8*cm)
    c.line(-1*cm,8.9*cm,22*cm,8.9*cm)
    c.line(7*cm,-1*cm,7*cm,29.7*cm)
    found=0
    for proc in range (0,len(process)):
        proc_serial=str(process.iloc[proc,0])
        print_permition=str(process.iloc[proc].values[4])
        
        
        if proc_serial!=code: #check if process code is correct
            continue
        else :
            if print_permition=="لا": #found but don't have print flag
                found=2
                continue
            found=1
            proc_name=str(process.iloc[proc].values[1])
            proc_price=str(process.iloc[proc].values[2])
            proc_status=str(process.iloc[proc].values[3])
            if (logo!=""):
                c.drawImage(logo, 7.5*cm, 19.5*cm-(count*invoice_width), width=12*cm, height=8.5*cm,preserveAspectRatio=True, mask='auto')
            
            if(proc_status=="مؤخر"):
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(process.iloc[proc].values[0])+month+"/"+year[2:4]
                days=str(checkMonth(int(month)))
            else:
                month=int(month1)+1
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
            cur_r-=1
            ############################################
            if (company_name!=""):
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
            text=arabic_text("للتواصل : "+phone)
            c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("تاريخ السداد : "+"   "+" / "+"  "+" / "+"    20  ")
            c.drawRightString(6*cm,24.25*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("ملحوظات")
            c.drawCentredString(3*cm,23.25*cm-(count*invoice_width),text)
            
            count=(count+1)%3
            
            if (count)%3==0 and proc+1!=len(process):
                #create new page
                c.showPage()
                c.setFont('Arabic', 14)
                c.translate(cm,cm) 
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1) #round line
                c.setDash(3,6)#dashed line
                c.line(-1*cm,18.8*cm,22*cm,18.8*cm)
                c.line(-1*cm,8.9*cm,22*cm,8.9*cm)
                c.line(7*cm,-1*cm,7*cm,29.7*cm)
            
    if found==0:
        return 0
    if found==2:
        return 2
    else:
        c.save()
        return 1
    
def createPDF_4_1(input_file,output_file,company_name,phone,logo,code,month1,year):
    """create invoice of process of this code from excel file using "reportlab" library (4 invoices per page)
        
    Args:
        excel_file (str): excel file path
        output_file (str): pdf file path (result)
        company_name (str): company name
        phone (str): company phone
        logo (str): company logo
        code (str): process code
        month1 (str): month
        year (str): year
    Returns:
        int: 0 if process not found, 1 if process found , 2 if process found but don't have print flag
        
        """
    year=str(year)
    process=pd.read_excel(input_file)
    invoice_width = 7.425*cm
    # my_path='Doc1.pdf'# file path
    c = canvas.Canvas(output_file,bottomup=1,pagesize=A4)
    count=0
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
    cur_p=0
    found=0
    for proc in range (0,len(process)):
        proc_serial=str(process.iloc[proc,0])
        print_permition=str(process.iloc[proc].values[4])
        if proc_serial!=code:
            continue
        else :
            if print_permition=="لا":
                found=2
                continue
            found=1
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
                month=int(month1)+1
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
            
            if (count)%4==0 and proc+1!=len(process):
                #create new page
                c.showPage()
                c.setFont('Arabic', 14)
                c.translate(cm,cm)
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1)
                c.setDash(3,6)#dashed line
                c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
                c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
                c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
                c.line(7*cm,-1*cm,7*cm,29.7*cm)
    if found==0:
        return 0
    if found==2:
        return 2
    else:
        c.save()
        return 1

MainUI,_ = loadUiType('UI.ui')
class Main(QMainWindow, MainUI):
    """a class for the main window"""
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handel_buttons()
        self.UI_changes()
        self.setFixedSize(self.size())    #stop the maximize button
    current_month=datetime.datetime.now().strftime("%m")
    current_year=datetime.datetime.now().strftime("%Y")
    last_dir=os.getcwd()
    def UI_changes(self):
        """changes in UI like hiding the title bar
        """
        self.comboBox.setCurrentIndex(int(self.current_month)-1)
        self.comboBox_2.setCurrentIndex(int(int(self.current_month)-1))
        self.lineEdit_4.setText(self.current_year)
        self.lineEdit_8.setText(self.current_year)
    def handel_buttons(self):
        """
            connect buttons in GUI with methods
        """
        #self.pushButton_2.clicked.connect(self.create_month_receipt)
        self.pushButton.clicked.connect(self.choose_file_excel)
        self.pushButton_3.clicked.connect(self.choose_file_excel)
        self.pushButton_13.clicked.connect(self.choose_file_image)
        self.pushButton_14.clicked.connect(self.choose_file_image)
        self.pushButton_4.clicked.connect(self.choose_save)
        self.pushButton_5.clicked.connect(self.choose_save)
        self.pushButton_2.clicked.connect(self.create_month_receipts)
        self.pushButton_6.clicked.connect(self.create_receipt)
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
    def clear_create_one_receipt(self):
        """clear the create one receipt tab data
        """
        self.lineEdit_5.setText("")
        self.lineEdit_6.setText("")
        self.lineEdit_9.setText("")
        self.lineEdit_10.setText("")
        self.comboBox_2.setCurrentIndex(int(self.current_month)-1)
        self.comboBox_6.setCurrentIndex(0)
        self.lineEdit_8.setText(self.current_year)
        self.lineEdit_12.setText("")
        self.lineEdit_28.setText("")
    def choose_file_excel(self):
        """open a file dialog to choose the excel file
        """
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=self.last_dir,filter="Excel (*.xlsx)")
        if file:
            if(self.tabWidget.currentIndex()==0):
                self.lineEdit.setText(file)
            else:
                self.lineEdit_5.setText(file)
            filename=file.split("/")[-1]
            self.last_dir=file.replace(filename,"")
        if file==None:
            QMessageBox.warning(self,"Error","يجب إختيار ملف")
            return   
    def choose_file_image(self):
        """open a file dialog to choose the image file"""
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=self.last_dir,filter="Image (*.png *.jpg)")
        if file:
            if(self.tabWidget.currentIndex()==0):
                self.lineEdit_27.setText(file)
            else:
                self.lineEdit_28.setText(file)
        if file==None:
            QMessageBox.warning(self,"Error","يجب إختيار ملف")
            return   
            
    def choose_save(self):
        """open a file dialog to choose the save directory"""
        if self.tabWidget.currentIndex()==0:
            if self.lineEdit.text()=='': #if no file selected
                QMessageBox.warning(self,"Error","يجب إختيار ملف العمارات اولاً")
                return
            if self.lineEdit_4.text()=='': #if no year
                QMessageBox.warning(self,"Error","يجب إدخال السنة")
                return
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            newFileName="ايصالات صيانة المصاعد"+self.comboBox.currentText()+"-"+self.lineEdit_4.text()
            fileName, _ = QFileDialog.getSaveFileName(self,"Save File",directory=self.last_dir+"\\"+newFileName,filter="PDF Files (*.pdf)", options=options)
            if fileName:
                self.lineEdit_2.setText(fileName)
            else:  
                return
            
        else:
            if self.lineEdit_5.text()=='': # file excel is empty
                QMessageBox. warning(self, "ERROR", "يجب اختيار ملف العمارات اولاً!")
                return
            if (self.lineEdit_9.text()==''): #code empty
                QMessageBox. warning(self, "ERROR", "يجب ادخال كود العمارة!")
                return
            if (self.lineEdit_8.text()==''): #year empty
                QMessageBox. warning(self, "ERROR", "يجب ادخال السنة!")
                return
            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            newFileName=" ايصال مصعد"+ self.lineEdit_9.text()+" - "+self.comboBox_2.currentText()+"-"+self.lineEdit_8.text()
            fileName, _ = QFileDialog.getSaveFileName(self,"Save File",directory=self.last_dir+"\\"+newFileName,filter="PDF Files (*.pdf)", options=options)
            if fileName:
                self.lineEdit_6.setText(fileName)
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
            QMessageBox. warning(self, "ERROR", "يجب ادخال رقم الشركة!")
            return
        if (self.lineEdit_2.text()==''):
            QMessageBox. warning(self, "ERROR", "يجب اختيار مكان حفظ الملف!")
            return
        logo=self.lineEdit_27.text()
        if (logo!=''):
            if not os.path.exists(logo):
                QMessageBox. warning(self, "ERROR", "الصورة غير موجودة!")
                return
        
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
            if (invoices_in_page==3):
                createPDF_3(input_file,output_file,company_name,phone,logo,month,year)
            else:
                createPDF_4(input_file,output_file,company_name,phone,logo,month,year)
            #create_all_receipts(input_file,output_file,company_name,month,year)
        except Exception as e:
            QMessageBox.warning(self, "ERROR","لقد وجدنا هذه الاخطاء :" +str(e))
            self.clear_create_month_receipts()
            return
        QMessageBox.information(self, "Success", "تم إنشاء الملف بنجاح!")
        self.lineEdit_2.setText("")
        return
    
    def create_receipt(self):
        if self.lineEdit_5.text()=='':
            QMessageBox. warning(self, "ERROR", "يجب اختيار ملف العمارات اولاً!")
            return
        if self.lineEdit_9.text()=='':
            QMessageBox. warning(self, "ERROR", "يجب ادخال كود العمارة!")
            return
        if self.lineEdit_8.text()=='':
            QMessageBox. warning(self, "ERROR", "يجب ادخال السنة!")
            return
        if self.lineEdit_6.text()=='':
            QMessageBox. warning(self, "ERROR", "يجب اختيار مكان حفظ الملف!")
            return
        
        input_file=self.lineEdit_5.text()
        if not os.path.exists(input_file):
            QMessageBox. warning(self, "ERROR", "ملف العمارات غير موجود!")
            return
        logo=self.lineEdit_28.text()
        if logo!="":
            if not os.path.exists(logo):
                QMessageBox. warning(self, "ERROR", "الصورة غير موجودة!")
                return
        company_name=self.lineEdit_10.text()
        code=self.lineEdit_9.text()
        year=self.lineEdit_8.text()
        phone=self.lineEdit_12.text()
        month=self.comboBox_2.currentText()
        output_file=self.lineEdit_6.text()+".pdf"
        try:
            if self.comboBox_6.currentIndex()==0:
                result=createPDF_3_1(input_file,output_file,company_name,phone,logo,code,month,year)
            else:
                result=createPDF_4_1(input_file,output_file,company_name,phone,logo,code,month,year)
        except Exception as e:
            QMessageBox.warning(self, "ERROR","لقد وجدنا هذه الاخطاء :" +str(e))
            self.clear_one_receipt()
            return
        if result==1:
            QMessageBox.information(self, "Success", "تم إنشاء الملف بنجاح!")
            self.lineEdit_6.setText("")
            return
        if result==2:
            QMessageBox.information(self, "INFO", "العمارة موجودة ولكن لا تحتوي علي اذن الطباعة!")
            self.lineEdit_6.setText("")
            return
        else:
            QMessageBox.warning(self, "ERROR","لم يتم العثور على العمارة!")
            self.clear_create_one_receipt()
            return
        
def main():
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
