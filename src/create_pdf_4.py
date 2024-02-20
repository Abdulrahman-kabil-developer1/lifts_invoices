import tempfile
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from src.assistant import generate_QR_code, int_to_text,check_month,arabic_text

def create_qr_4(c,data):
    count=0
    for i in range(len(data)):
        img = generate_QR_code(data[i])
        temp_image_path = tempfile.NamedTemporaryFile(suffix=".png").name
        img.save(temp_image_path)
        text1=arabic_text("تم انشاء هذا الايصال الكترونيا بواسطة شركة Yodix Solutions")
        text2=arabic_text(" لأنظمة الادارة للتواصل 1149312512 (20+) - yodix@mail.com")
        c.setFont('Arabic', 8)
        if i%2==0:
            c.drawImage(temp_image_path, 2.8*cm, 24.5*cm-(count*7.425)*cm, width=3*cm, height=3*cm,preserveAspectRatio=True, mask='auto')
            c.drawCentredString(4.3*cm,23.8*cm-(count*7.425*cm),text1)
            c.drawCentredString(4.3*cm,23.3*cm-(count*7.425*cm),text2)
        else:
            c.drawImage(temp_image_path, 13.3*cm, 24.5*cm-(count*7.425)*cm, width=3*cm, height=3*cm,preserveAspectRatio=True, mask='auto')
            c.drawCentredString(14.8*cm,23.8*cm-(count*7.425*cm),text1)
            c.drawCentredString(14.8*cm,23.3*cm-(count*7.425*cm),text2)
            count+=1
            
        c.setFont('Arabic', 14)
            
def create_pdf_4(self,input_file,output_file,company_name,phone,codes,logo,signature,month1,year):
        """
            create invoices pdf file from excel file using "reportlab" library (3 invoices per page)
            
            Args:
                input_file (str): excel file path
                output_file (str): pdf file path (result)
                company_name (str): company name
                phone (str): company phone
                logo (str): company logo
                signature (str): manager signature
                month1 (str): month
                year (str): year
        
        """
        year=str(year)
        try:
            df=pd.read_excel(input_file, engine='openpyxl')
            df=df[['serial','full_name','price','status','print_or_no']]
        except Exception as e :
            self.error.emit("خطأ في ملف البيانات\n"+str(e)+"\n يجب ان يحتوي الملف علي اعمدة من نوع \n [int64,object,int64,object,object]\n و اسمائها \n [serial,full_name,price,status,print_or_no]")
            return  
        invoice_width = 7.425*cm
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
        c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
        c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
        c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
        c.line(9.5*cm,-1*cm,9.5*cm,29.7*cm)
        
        if codes != "":
            codes_list=codes.split("-")
            try:
                codes_list=[int(item) for item in codes_list]
                if len(codes_list)>0:
                    found_no_print=df.loc[(df.serial.isin(codes_list)) & (df.print_or_no=='لا')]
                    df=df.loc[(df.serial.isin(codes_list)) & (df.print_or_no=='نعم') ]
                    num_df=len(df)
                    not_found_codes=list(set(codes_list)-(set(df.serial).union(set(found_no_print.serial))))
                    message=''
                    if num_df>0:
                        message='هذة الاكواد سيتم طباعتها \n'+str(list(df.serial))
                    if len(found_no_print)>0:
                        message+="\nهذة الاكواد تم العثور عليها في الملف ولكنها لم تمنح اذن الطباعة \n"+str(list(found_no_print.serial))
                    if len(not_found_codes)>0:
                        message+="\n هذة الاكواد لم يتم العثور عليها في الملف \n"+str(list(not_found_codes))+"\n"
                    self.info.emit(message)      
                if num_df==0:
                    self.error.emit("لا يوجد اكواد يمكن طباعتها")
                    return
            except:
                self.info.emit(" خطأ في رقم الفاتورة يجب ان يكون ارقام صحيحة فقط وليست حروف")
                return
        else:
            df=df[df.print_or_no == 'نعم']
            num_df=len(df)
        
        data_for_qr = []
        for proc in range (0,len(df)):
            process={}
            proc_serial=str(int(df.iloc[proc].serial))
            proc_name=str(df.iloc[proc].full_name)
            proc_price=str(int(df.iloc[proc].price))
            proc_status=str(df.iloc[proc].status)
            if(company_name!=''):
                process['الشركة'] = company_name
            if(phone!=''):
                process['موبايل'] = phone
            process['الكود'] = proc_serial
            process['الاسم'] = proc_name
            process['المبلغ'] = proc_price
            process['طريقة الدفع'] = proc_status
            if (logo!=""):
                #draw logo
                if proc%2==0:
                    c.drawImage(logo, 9*cm, 22*cm-(count*invoice_width), width=12*cm, height=6.025*cm,preserveAspectRatio=True, mask='auto')
                else:
                    c.drawImage(logo, -1.5*cm, 22*cm-(count*invoice_width), width=12*cm, height=6.025*cm,preserveAspectRatio=True, mask='auto')
            if(proc_status=="مؤخر"): 
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(proc_serial)+month+"/"+year[2:4]
                days=str(check_month(int(month),int(year)))
            else:
                month=(int(month1)+1)
                if(month==13):
                    month=1
                    year = int(year)+1
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                year = str(year)
                proc_serial=str(df.iloc[proc].values[0])+month+"/"+year[2:4]
                days=str(check_month(int(month),int(year)))
            cur_r=28.25  #current Row to write in 
            ############################################
            if proc%2==0:
                text=arabic_text('إيصال صيانة')
                c.drawCentredString(14.7*cm,(cur_r-0.1)*cm-(count*invoice_width),text)
            else:
                text=arabic_text('إيصال صيانة')
                c.drawCentredString(4.2*cm,(cur_r-0.1)*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            
            if (company_name!=""):
                text=arabic_text(company_name)
                if proc%2==0:
                    c.drawCentredString(14.7*cm,cur_r*cm-(count*invoice_width),text)
                else:
                    c.drawCentredString(4.7*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=0.7
            
            ############################################
            text=arabic_text("رقم : ")
            if proc%2==0:
                c.drawRightString(19*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(8.5*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            #invoice number
            if proc%2==0:
                c.drawRightString(17.7*cm,cur_r*cm-(count*invoice_width),proc_serial)
            else:
                c.drawRightString(7.2*cm,cur_r*cm-(count*invoice_width),proc_serial)
            ############################################
            text=arabic_text("المبلغ : ")
            if proc%2==0:
                c.drawRightString(14*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(3.5*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text("جــــــــ")
            c.setFont('Arabic', 13)
            if proc%2==0:
                c.drawString(11.5*cm,(cur_r+0.5)*cm-(count*invoice_width),text)
            else:
                c.drawString(1*cm,(cur_r+0.5)*cm-(count*invoice_width),text)
            c.setFont('Arabic', 14)
            if proc%2==0:
                c.drawRightString(12.3*cm,cur_r*cm-(count*invoice_width),proc_price)
            else:
                c.drawRightString(1.8*cm,cur_r*cm-(count*invoice_width),proc_price)
            cur_r-=0.7
            ############################################
            text=arabic_text("وصلنا من السيد/ ")
            if proc%2==0:
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(9.3*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text(proc_name)
            if len(proc_name)>23 :
                c.setFont('Arabic', 12)
            if len(proc_name)>31:
                c.setFont('Arabic', 11)
            if proc%2==0:
                c.drawRightString(16.3*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(5.8*cm,cur_r*cm-(count*invoice_width),text)
            c.setFont('Arabic', 14)
            cur_r-=0.7
            ############################################
            text=arabic_text("مبلغ و قدرة : ")
            if proc%2==0:
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(9.3*cm,cur_r*cm-(count*invoice_width),text)
            ############################################
            text=arabic_text(int_to_text(proc_price))
            if proc%2==0:
                c.drawRightString(17*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(6.5*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("وذلك قيمة : صيانة المصعد من 1 - "+days+" / "+month+" / "+year)
            if  proc%2==0:
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawRightString(9.3*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            text=str("صيانة المصعد من 1 - "+days+" / "+month+" / "+year)
            process["مدة الصيانة"]= text
            ############################################
            if(proc_status=="مؤخر"):
                text=arabic_text("تحريراً في : "+days+" / "+month+" / "+year)
                if proc%2==0:
                    c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
                else:
                    c.drawRightString(9.3*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=0.7
                process['تاريخ الاستخراج']= "تحريراً في : "+days+" / "+month+" / "+year
                
            else:
                text=arabic_text(f"تحريراً في : 01 / {month} / {year}")
                if proc%2==0:
                    c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
                else:
                    c.drawRightString(9.3*cm,cur_r*cm-(count*invoice_width),text)
                process['تاريخ الاستخراج']=f"تحريراً في : 01 / {month} / {year}"
                cur_r-=0.7
                if (month=='01'):
                    year = int(year)-1
                    year = str(year)
                
            process['البرنامج'] = 'تم انشاء هذا الايصال الكترونيا بواسطة شركة Yodix Solutions لأنظمة الأدارة للتواصل +201149312512'
            data_for_qr.append(process)
            ############################################
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            if proc%2==0:
                c.drawCentredString(14.7*cm,cur_r*cm-(count*invoice_width),text)
            else:
                c.drawCentredString(4.2*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("توقيع المستلم -------------------------")
            if proc%2==0:
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text) 
            else:
                c.drawRightString(9.3*cm,cur_r*cm-(count*invoice_width),text) 
            if (signature!=""):
                #draw signature
                if proc%2==0:
                    c.drawImage(signature, 12*cm, (cur_r-0.85)*cm-(count*invoice_width), width=5*cm, height=2*cm,preserveAspectRatio=True, mask='auto')
                else:
                    c.drawImage(signature, 1.5*cm, (cur_r-0.85)*cm-(count*invoice_width), width=5*cm, height=2*cm,preserveAspectRatio=True, mask='auto')
            cur_r-=0.7
            ############################################
            if (company_name==""):
                cur_r-=0.7
            if (phone!=""):
                text=arabic_text("للتواصل : "+phone)
                if proc%2==0:
                    c.drawCentredString(14.7*cm,cur_r*cm-(count*invoice_width),text)
                else:
                    c.drawCentredString(4.2*cm,cur_r*cm-(count*invoice_width),text)
            else:
                cur_r-=0.7
            ############################################
            
            if (proc%2!=0):
                count= count+1
            count=(count)%4
            count2+=1
            self.value=(count2/num_df)*100
            self.value_changed.emit(self.value)
            if (proc+1)%8==0 and proc+1!=len(df):   #if it is the last invoice in the page
                c.showPage()    #create new page
                c.setFont('Arabic', 14)
                c.translate(cm,cm) 
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1)
                c.setDash(3,6)
                c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
                c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
                c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
                c.line(9.5*cm,-1*cm,9.5*cm,29.7*cm)
                create_qr_4(c,data_for_qr)
                data_for_qr=[]
                c.showPage()    #create new page
                c.setFont('Arabic', 14)
                c.translate(cm,cm) 
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1)
                c.setDash(3,6)
                c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
                c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
                c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
                c.line(9.5*cm,-1*cm,9.5*cm,29.7*cm)
                
        
            if proc+1==len(df) and (proc+1)%8!=0:   #if it is the last invoice in the page
                c.showPage()    #create new page
                c.setFont('Arabic', 14)
                c.translate(cm,cm) 
                c.setLineWidth(1.5)#width of the line
                c.setLineCap(1)
                c.setDash(3,6)
                c.line(-1*cm,21.275*cm,22*cm,21.275*cm)
                c.line(-1*cm,13.85*cm,22*cm,13.85*cm)
                c.line(-1*cm,6.425*cm,22*cm,6.425*cm)
                c.line(9.5*cm,-1*cm,9.5*cm,29.7*cm)
                create_qr_4(c,data_for_qr)
                data_for_qr=[]
                
            dont_save=0
            
            
        if dont_save==0:
            c.save()


# createPDF_3('self','العمارات 2.xlsx','out.pdf',"",'01001041961','','final logo.png','sign1.png',
#             1,2024)
