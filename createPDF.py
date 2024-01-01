import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('Arabic', 'Janna LT Bold.ttf'))
from assistant import int_to_text,check_month,arabic_text

def createPDF_3(self,input_file,output_file,company_name,phone,codes,logo,signature,month1,year):
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
            
        for proc in range (0,len(df)):
            proc_serial=str(df.iloc[proc].serial)
            proc_name=str(df.iloc[proc].full_name)
            proc_price=str(df.iloc[proc].price)
            proc_status=str(df.iloc[proc].status)
            
            if (logo!=""):
                #draw logo
                c.drawImage(logo, 7.5*cm, 19.5*cm-(count*invoice_width), width=12*cm, height=8.5*cm,preserveAspectRatio=True, mask='auto')
            if(proc_status=="مؤخر"): 
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(df.iloc[proc].values[0])+month+"/"+year[2:4]
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
            if(proc_status=="مؤخر"):
                text=arabic_text("تحريراً في : "+days+" / "+month+" / "+year)
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=1
            else:
                text=arabic_text(f"تحريراً في : 01 / {month} / {year}")
                year = int(year)-1
                year = str(year)
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=1
            ############################################
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=1
            ############################################
            text=arabic_text("توقيع المستلم -------------------------")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text) 
            if (signature!=""):
                #draw signature
                c.drawImage(signature, 12*cm, (cur_r-0.85)*cm-(count*invoice_width), width=5*cm, height=2*cm,preserveAspectRatio=True, mask='auto')
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
            self.value=(count2/num_df)*100
            self.value_changed.emit(self.value)
            if (count)%3==0 and proc+1!=len(df):   #if it is the last invoice in the page
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

def createPDF_4(self,input_file,output_file,company_name,phone,codes,logo,signature,month1,year):
        """
            create invoices pdf file from excel file using "reportlab" library (4 invoices per page)
            
            Args:
                input_file (str): excel file path
                output_file (str): pdf file path (result)
                company_name (str): company name
                phone (str): company phone
                codes (str) : if user want only create some df by serial code
                logo (str): company logo
                signature (str): manager signature
                month1 (str): month
                year (str): year
        
        """
        year=str(year)
        
        try:
            df=pd.read_excel(input_file)
            df=df[['serial','full_name','price','status','print_or_no']]
        except Exception as e :
            self.error.emit("خطأ في ملف البيانات\n"+str(e)+"\n يجب ان يحتوي الملف علي اعمدة من نوع \n [int64,object,int64,object,object]\n و اسمائها \n [serial,full_name,price,status,print_or_no]")
            return  
        
        num_df=(df.print_or_no == 'نعم').sum()
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
        for proc in range (0,len(df)):
            proc_serial=str(df.iloc[proc].serial)
            proc_name=str(df.iloc[proc].full_name)
            proc_price=str(df.iloc[proc].price)
            proc_status=str(df.iloc[proc].status)
            
            if logo!="":
                c.drawImage(logo, 7.5*cm, 22*cm-(count*invoice_width), width=12*cm, height=6.025*cm,preserveAspectRatio=True, mask='auto')
            if(proc_status=="مؤخر"):
                month=int(month1)
                if(len(str(month))==1):
                    month="0"+str(month)
                month=str(month)
                proc_serial=str(df.iloc[proc].values[0])+month+"/"+year[2:4]
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
            if(proc_status=="مؤخر"):
                text=arabic_text("تحريراً في : "+days+" / "+month+" / "+year)
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=0.7
            else:
                text=arabic_text(f"تحريراً في : 01 / {month} / {year}")
                year = int(year)-1
                year = str(year)
                c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text)
                cur_r-=0.7
            ############################################
            text=arabic_text("وتحرر هذا ايصالاً بالاستلام")
            c.drawCentredString(13.5*cm,cur_r*cm-(count*invoice_width),text)
            cur_r-=0.7
            ############################################
            text=arabic_text("توقيع المستلم -------------------------")
            c.drawRightString(19.8*cm,cur_r*cm-(count*invoice_width),text) 
            if (signature!=""):
                #draw signature
                c.drawImage(signature, 12*cm, (cur_r-1)*cm-(count*invoice_width), width=5.3*cm, height=2*cm,preserveAspectRatio=True, mask='auto')
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
            self.value=(count2/num_df)*100
            self.value_changed.emit(self.value)
            if (count)%4==0 and proc+1!=len(df):
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
