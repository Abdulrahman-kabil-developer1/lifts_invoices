from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import letter,A4
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
pdfmetrics.registerFont(TTFont('Arabic', 'Janna LT Bold.ttf'))
nums=["200309/22","200309/22","200309/22","200109/22","200209/22","200309/22",""]
price=["500","500","500","300","400","500","554"]
process=["عمارة 9400 شارع كريم بنونة (الحاج حسن)","عمارة 9400 شارع كريم بنونة (الحاج حسن)","عمارة 1 ب 54 المقطم","عمارة 9400 شارع كريم بنونة (الحاج حسن)","عمارة 9400 شارع كريم بنونة (الحاج حسن)","عمارة 1 ب 54 المقطم",""]
companyName="شركة ال قابيل للمصاعد"
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
        result=str( firstToText(first))+" جنية لاغير"
        result= checkResult(result)
        return "فقط "+result
    elif len(str(num))==2:
        result=str( firstToText(first)) +str( secondToText(second))+" جنية لاغير"
        result= checkResult(result)
        return "فقط "+result
    elif len(str(num))==3:
        result=str( thirdToText(third)) +str( firstToText(first))+str( secondToText(second))+" جنية لاغير"
        result= checkResult(result)
        return "فقط "+result
    elif len(str(num))==4:
        result=str( fourthToText(fourth))+str( thirdToText(third))+str( firstToText(first))+str( secondToText(second))+" جنية فقط لاغير"
        result= checkResult(result)
        return "فقط "+result 
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

def createPDF_4(input_file,output_file,company_name,phone,logo,month1,year):
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
            print ("count",count)
            print("proc",proc)
            print("cur_p",cur_p)
            print("dddd")
            continue
        if logo!="":
            c.drawImage(logo, 7.5*cm, 22*cm-(count*invoice_width), width=12*cm, height=6.025*cm,preserveAspectRatio=True, mask='auto')
        print ("count",count)
        print("proc",proc)
        print("cur_p",cur_p)
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
        text=arabic_text(intToText1(proc_price))
        c.drawRightString(16.5*cm,cur_r*cm-(count*invoice_width),text)
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
            print(len(process))
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
createPDF_4("1.xlsx",'Doc4.pdf',"شركة ار ليفت للمصاعد","01001041961","2.png",'9','2022')