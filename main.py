import sys
from PyQt5.QtGui import QIcon
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow,QMessageBox,QApplication
from UI import Ui_MainWindow as ui
from assistant import check_path_exists,get_cur_month_year,update_last_dir,get_last_dir
from createInvoices import Create_Invoices

"""
    Abdulrahman Ragab Abdullah 
    Cairo-EG
    ********Contact********
    abdulrahman.ragab.kabil@gmail.com
    (+20) 1280059456 - 1149312512
    https://github.com/Abdulrahman-Kabil-developer1
    https://www.linkedin.com/in/abdulrahman-kabil-5729621a2/
    
"""

class Main(QMainWindow, ui):
    """a class for the main window"""
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.handel_buttons()
        self.UI_changes()
        self.setFixedSize(self.size())#stop the maximize button
        
    current_month,current_year=get_cur_month_year()
    
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
        self.pushButton_13.clicked.connect(self.choose_file_logo)
        self.pushButton_14.clicked.connect(self.choose_file_signature)
        self.pushButton_4.clicked.connect(self.choose_save)
        self.pushButton_2.clicked.connect(self.create_month_receipts)
    
    def choose_file_excel(self):
        """open a file dialog to choose the excel file
        """
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=get_last_dir(),filter="Excel (*.xlsx)")
        if file:
            self.lineEdit.setText(file)
            filename=file.split("/")[-1]
            update_last_dir(file.replace(filename,""))
        if file==None:
            self.show_error( "يجب إختيار ملف")
            return   
        
    def choose_file_logo(self):
        """open a file dialog to choose the image file"""
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=get_last_dir(),filter="Image (*.png *.jpg)")
        if file:
            self.lineEdit_27.setText(file)
        if file==None:
            self.show_error("يجب إختيار ملف")
            return   
        
    def choose_file_signature(self):
        """open a file dialog to choose the image file"""
        file, _ = QtWidgets.QFileDialog.getOpenFileName(None,directory=get_last_dir(),filter="Image (*.png *.jpg)")
        if file:
            self.lineEdit_28.setText(file)
        if file==None:
            self.show_error( "يجب إختيار ملف")
            return  
                    
    def choose_save(self):
        """open a file dialog to choose the save directory"""
        if self.lineEdit.text()=='': #if no file selected
            self.show_error( "يجب إختيار ملف العمارات اولاً")
            return
        if self.lineEdit_4.text()=='': #if no year
            self.show_error( "يجب إدخال السنة")
            return
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        newFileName="ايصالات صيانة المصاعد"+self.comboBox.currentText()+"-"+self.lineEdit_4.text()
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self,"Save File",directory=get_last_dir()+"\\"+newFileName,filter="PDF Files (*.pdf)", options=options)
        if fileName:
            self.lineEdit_2.setText(fileName)
            filename=fileName.split("/")[-1]
            update_last_dir(fileName.replace(filename,""))
        else:  
            return         
    
    def clear_create_month_receipts(self):
        """clear the create month receipts tab data
        """
        self.lineEdit.setText("")
        self.lineEdit_2.setText("")
        self.lineEdit_11.setText("")
        self.lineEdit_13.setText("")
        self.lineEdit_27.setText("")
        self.lineEdit_28.setText("")
        self.comboBox.setCurrentIndex(int(self.current_month)-1)
        self.comboBox_5.setCurrentIndex(0)
        self.lineEdit_4.setText(self.current_year)
        

    def create_month_receipts(self):
        if self.comboBox_5.currentIndex()==0:
            invoices_in_page=3
        else:
            invoices_in_page=4
        if (self.lineEdit.text()==''):
            self.show_error("يجب اختيار ملف العمارات اولاً!")
            return
        if (self.lineEdit_4.text()==''):
            self.show_error("يجب ادخال السنة!")
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
            self.show_error( "يجب اختيار مكان حفظ الملف!")
            return
        logo=self.lineEdit_27.text()
        signature=self.lineEdit_28.text()
        if (logo!=''):
            if not check_path_exists(logo):
                self.show_error( "صورة الشعار غير موجودة!")
                return
        if (signature!=''):
            if not check_path_exists(signature):
                self.show_error( "صورة التوقيع غير موجودة!")
                return
        
        codes=self.lineEdit_14.text()
        input_file=self.lineEdit.text()
        if not check_path_exists(input_file):
            self.show_error( "ملف العمارات غير موجود!")
            return
        company_name=self.lineEdit_11.text()
        year=self.lineEdit_4.text()
        month= self.comboBox.currentText()
        phone=self.lineEdit_13.text()
        output_file=self.lineEdit_2.text()+".pdf"
        try:
            self.Creator=Create_Invoices(invoices_in_page,input_file,output_file,company_name,phone,codes,logo,signature,month,year)
            self.Creator.value_changed.connect(self.update_progress)
            self.Creator.error.connect(self.show_error)
            self.Creator.info.connect(self.show_info)
            self.Creator.start()
            self.pushButton_2.setEnabled(False)
            
        except Exception as e:
            self.pushButton_2.setEnabled(True)
            self.show_error(e)
            self.update_progress(0)
            return
    
    def show_error(self,msg):
        QMessageBox.warning(self, "ERROR","لقد وجدنا هذة الاخطاء:\n"+str(msg))
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
                self.show_info("تم إنشاء الملف بنجاح!")
                self.progressBar.setValue(0)
                self.lineEdit_2.setText("")
                self.pushButton_2.setEnabled(True)
                return
        except Exception as e:
            self.pushButton_2.setEnabled(True)
            self.show_error(e)
            self.update_progress(0)
            return
            
        
def main():
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
