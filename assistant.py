import arabic_reshaper
from bidi.algorithm import get_display
import os
import calendar
import datetime

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
        result=str( fourth_to_text(fourth))+str( third_to_text(third))+str( first_to_text(first))+str( second_to_text(second))+" جنية لاغير"
        result= check_result(result)
        return "فقط "+result 
    else :
        result= "خطأ"
        return result


def check_month(month,year):
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
    elif calendar.isleap(year):
        return 29
    else:
        return 28

def get_cur_month_year():
    cur_month=datetime.datetime.now().strftime("%m")
    cur_year=datetime.datetime.now().strftime("%Y")
    return cur_month,cur_year
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
        
def check_path_exists(path):
    if os.path.exists(path):
        return True
    return False