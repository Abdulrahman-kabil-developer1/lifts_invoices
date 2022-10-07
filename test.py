from cgitb import text
from gettext import textdomain
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
from PyQt5 import QtCore, QtGui, QtWidgets
import datetime
################int to text##################
def firstToText(first):
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
def secondToText(second):
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
def thirdToText(third):
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
def fourthToText(fourth):
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
def checkResult(result):
    if result[0]==" "and result[1]=="و" and result[2]==" ":
        return result[3:]
    else:
        return result
def intToText1(num):
    num=int(num)
    first=num%10
    second=int(num/10%10)
    third=int(num/100%10)
    fourth=int(num/1000%10)
    if len(str(num))==1:
        result= str( firstToText(first))+" جنية فقط لاغير"
        result= checkResult(result)
        return result
    elif len(str(num))==2:
        result= str( firstToText(first)) +str( secondToText(second))+" جنية فقط لاغير"
        result= checkResult(result)
        return result
    elif len(str(num))==3:
        result= str( thirdToText(third)) +str( firstToText(first))+str( secondToText(second))+" جنية فقط لاغير"
        result= checkResult(result)
        return result
    elif len(str(num))==4:
        result= str( fourthToText(fourth))+str( thirdToText(third))+str( firstToText(first))+str( secondToText(second))+" جنية فقط لاغير"
        result= checkResult(result)
        return result
    else :
        result= "خطأ"
        return result
#############################################
#check if month 30 or 31 days   
def checkMonth(month):
    if month==1 or month==3 or month==5 or month==7 or month==8 or month==10 or month==12:
        return 31
    elif month==4 or month==6 or month==9 or month==11:
        return 30
    else:
        return 28

def arabic_text(text):
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    return bidi_text

def create_all(input_file,output_file,company_name,phone,month1,year):
    year=str(year)
    process=pd.read_excel(input_file)
    pdf=FPDF()
    pdf=FPDF('P','mm','A4')
    pdf.add_page()
    pdfW=210
    pdfH=297
    pdf.add_font("Janna", "", "Janna LT Bold.ttf", uni=True)
    pdf.set_font('Janna',''   ,12)
    
    
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
        #pdf.dashed_line(pdfW/2+3,0,pdfW/2+3,pdfH,2,1)
        
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text("نسخة ايصال صيانة")
        pdf.cell(90,7,text,0,0,'C')
        text=arabic_text("ايصــال صيـانـة")
        pdf.cell(120,7,text,0,1,'C')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("    ")
        pdf.cell(10,7,text,0,0,'R')
        pdf.ln(2.5)
        ###############################################################################
        if (company_name==""):
            text=arabic_text(" ")
            pdf.cell(10,7,text,0,1,'R')
        else:
            # pdf.set_font('Janna',''   ,12)
            # text=arabic_text(company_name)
            # pdf.cell(70,7,text,0,0,'C')
            text=arabic_text(company_name)
            pdf.cell(137,7,text,0,1,'C')
            pdf.set_font('Janna',''   ,12)
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text(str(proc_price)+" جنية")
        pdf.cell(20,7,text,0,0,'R')
        
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("المبلغ: ")
        pdf.cell(14,7,text,0,0,'R')
        
        pdf.set_font('Janna','U'   ,12)
        text=str(proc_serial)
        pdf.cell(30,7,text,0,0,'R')
        
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("رقم الايصال: ")
        pdf.cell(24,7,text,0,0,'R')
        
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text(str(proc_price)+" جنية")
        pdf.cell(36,7,text,0,0,'R')
        
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("المبلغ: ")
        pdf.cell(14,7, text,0,0,'R')
        
        pdf.set_font('Janna','U'   ,12)
        text=proc_serial
        pdf.cell(32,7,text,0,0,'R')
        
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("رقم الايصال: ")
        pdf.cell(24,7, text,0,1,'R')
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text(proc_name)
        pdf.cell(68,7, text,0,0,'R')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("وصلنا من السيد/")
        pdf.cell(30,7, text,0,0)
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text(proc_name)
        pdf.cell(70,7, text,0,0,'R')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("وصلنا من السيد/" )
        pdf.cell(30,7, text,0,1)
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text(intToText1(proc_price))
        pdf.cell(75,7, text,0,0,'R')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("مبلغ و قدرة: " )
        pdf.cell(20,7, text,0,0,'R')
        pdf.set_font('Janna','U'   ,12)
        text=arabic_text(intToText1(proc_price))
        pdf.cell(80,7, text,0,0,'R')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("مبلغ و قدرة: " )
        pdf.cell(20,7, text,0,1,'R')
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna',''   ,12)
        text=arabic_text(" وذلك قيمة:  صيانة المصعد من 1- "+days + " / "+month+" / "+year)
        pdf.cell(97,7, text,0,0,'R')
        text=arabic_text(" وذلك قيمة:  صيانة المصعد من 1- "+days + " / "+month+" / "+year)
        pdf.cell(100,7, text,0,1,'R')
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna','U'   ,12)
        text=days+"-"+month+"-"+year
        if proc_status == "مقدم":
            text="01"+"-"+month+"-"+year 
        pdf.cell(75,7,text,0,0,'R')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("تحريراً في: ")
        pdf.cell(20,7, text,0,0,'R')
        pdf.set_font('Janna','U'   ,12)
        text=days+"-"+month+"-"+year
        if proc_status == "مقدم":
            text="01"+"-"+month+"-"+year 
        pdf.cell(75,7,text,0,0,'R')
        pdf.set_font('Janna',''   ,12)
        text=arabic_text("تحريراً في: ")
        pdf.cell(20,7, text,0,1,'R')
        pdf.ln(2.5)
        ###############################################################################
        text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
        pdf.cell(88,4, text,0,0,'C')
        text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
        pdf.cell(120,4, text,0,1,'C')
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna',''   ,12)
        pdf.cell(70,7,"------------",0,0,'R')
        text=arabic_text("توقيع المستلم")
        pdf.cell(25,7, text,0,0,'C')
        pdf.cell(70,7,"-----------",0,0,'R')
        text=arabic_text("توقيع المستلم")
        pdf.cell(25,7, text,0,1,'C')
        pdf.ln(2.5)
        ###############################################################################
        pdf.set_font('Janna',''   ,12)
        text=arabic_text(" للتواصل :"+phone)
        pdf.cell(70,7,text,0,0,'C')
        text=arabic_text(" للتواصل :"+phone)
        pdf.cell(137,7,text,0,1,'C')
        pdf.set_font('Janna',''   ,12)
        pdf.ln(2.5)
        ###############################################################################
        end=70.5
        #pdf.dashed_line(0,i*(end),pdfW,i*(end),2,1)
        ###################################################################################
        #pdf.ln(12)
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
            
            # if (page_num!=len(process)/4):
            #     pdf.add_page()
            pdf.set_y(0)
            
            i=1

    pdf.output(output_file, 'F')
    pdf.close()
    
create_all("2.xlsx","2.pdf","ار ليفت للمصاعد","01280059456","9","2022")