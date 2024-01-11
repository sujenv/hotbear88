import os
import sys
import logging
import pyodbc
import win32com.client as win32
from PyQt5 import uic
from PyQt5.QtCore import Qt, QDateTime, pyqtSignal, QTimer
from PyQt5.QtWidgets import QApplication, QMainWindow, QMdiArea, QDialog, QLineEdit, QMessageBox, QFileDialog
from dialogs import *
#for non_ui version-------------------------
#from main_ui import Ui_MainWindow
#from login_ui import Ui_LoginDialog

# Initialize current_datetime as a global variable
current_datetime = QDateTime.currentDateTime().toString(Qt.DefaultLocaleLongDate)

class LoginWindow(QDialog):
#for non_ui version--------------------------
#class LoginWindow(QDialog, Ui_LoginDialog):

    # Custom signal to pass username to pass username to MainWindow
    login_successful = pyqtSignal(str)  
    user_credentials = userlist
    
    def __init__(self): 
        super().__init__()
        
        # Load the main window UI
        uic.loadUi("login.ui", self)
        
        #for non_ui version-------------------------
        #self.setupUi(self)

        self.password_entry.setEchoMode(QLineEdit.Password)
        self.login_button.clicked.connect(self.check_login)

        # Initialize login attempts counter
        self.login_attempts = 0  

    def check_login(self):
        username = self.username_entry.text()
        password = self.password_entry.text()

        # Replace with your authentication logic
        if username in LoginWindow.user_credentials and LoginWindow.user_credentials[username] == password:
         
            self.status_label.setText("Login Successful")
            self.login_successful.emit(username)  # Emit the signal to pass username to MainWindow
            self.authenticate_user(username)
        
        else:
            # Increment login attempts
            self.login_attempts += 1  
            self.status_label.setText("Login Failed: " + str(self.login_attempts) + " times / Please check your username and password")
        
            self.username_entry.setText("")
            self.password_entry.setText("")

            # Check login attempts count and close if >= 3
            if self.login_attempts >= 3:
                self.close()

    def authenticate_user(self, username):
        QMessageBox.information(self, "Login Successful", f" 환영합니다!, {username} 님!")
    
        # Log the access information
        logging.info(f"User {username} logged in successfully.")

        # Close the login window or navigate to the main application window.
        self.accept()  


class MainWindow(QMainWindow, SubWindowBase):
#for non_ui version-------------------------
#class MainWindow(QMainWindow, Ui_MainWindow, SubWindowBase):

    def __init__(self, username):
        super().__init__()
        
        # Load the main window UI
        uic.loadUi("main.ui", self)
        
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Create a variable to keep track of the currently open dialog
        self.open_dialog = None
                
        # Find and Access the QMdiArea from the UI file
        self.mdi_area = self.findChild(QMdiArea, "mdiArea")
        # Base information
        self.sel_customer.triggered.connect(self.call_customer)
        self.sel_customercar.triggered.connect(self.call_customercar)
        self.sel_cuscorpno.triggered.connect(self.call_customercorpno)
        self.sel_customeraddress.triggered.connect(self.call_customeraddress)
        self.sel_customerbkaccno.triggered.connect(self.call_customerbkaccno)
        self.sel_employee.triggered.connect(self.call_employee)
        self.sel_employeecar.triggered.connect(self.call_employeecar)
        self.sel_delivery.triggered.connect(self.call_delivery)
        self.sel_nr_salesitem.triggered.connect(self.call_nr_salesitem)
        self.sel_nr_salesprice.triggered.connect(self.call_nr_salesprice)
        self.sel_ap_product.triggered.connect(self.call_ap_product)
        self.sel_ap_cost.triggered.connect(self.call_ap_cost)
        self.sel_regno.triggered.connect(self.call_regno)
        self.sel_phoneno.triggered.connect(self.call_phoneno)
        self.sel_address.triggered.connect(self.call_empaddress)
        self.sel_cusphoneno.triggered.connect(self.call_cusphoneno)
        self.sel_bankaccno.triggered.connect(self.call_bankaccno)
        self.sel_genbkaccproduct.triggered.connect(self.call_genbankaccno)
        self.sel_genbkacccheck.triggered.connect(self.call_genbankacccheck)
        self.sel_paymentproduct.triggered.connect(self.call_paymentproduct)
        self.sel_gpaiinfo.triggered.connect(self.call_gpaiinfo)
        self.sel_gpaigen.triggered.connect(self.call_gpaigen)
        self.sel_gpaiview.triggered.connect(self.call_gpaiview)
        self.sel_oilusage.triggered.connect(self.call_oilusage)
        self.sel_oilusageclosing.triggered.connect(self.call_oilusageclosing)
        self.sel_oilusagecom.triggered.connect(self.call_oilusagecom)
        self.sel_oilusageemp.triggered.connect(self.call_oilusageemp)
        self.sel_supportaptinfo.triggered.connect(self.call_supportaptinfo)
        self.sel_supportaptgen.triggered.connect(self.call_supportaptgen)
        self.sel_supportaptview.triggered.connect(self.call_supportaptview)

        self.DB_update.triggered.connect(self.updatedbs)

        #self.sel_move_contents.triggered.connect(self.ini_move_data)
        
        # AR AP Parts
        self.sel_arfileopen.triggered.connect(self.excel_araps_open)
        self.sel_artoap.triggered.connect(self.excel_araps_open)
        self.sel_arclosing.triggered.connect(self.excel_araps_open)
        self.sel_apfileopen.triggered.connect(self.excel_araps_open)
        self.sel_apcheck.triggered.connect(self.excel_araps_open)
        self.sel_apslip.triggered.connect(self.excel_araps_open)
        self.sel_apcheckttl.triggered.connect(self.excel_araps_open)
               
        # Consumables handling
        self.sel_consumable_product.triggered.connect(self.call_consumable_product)
        self.sel_consumable_inprice.triggered.connect(self.call_consumable_in_price)
        self.sel_consumable_outprice.triggered.connect(self.call_consumable_out_price)
        self.sel_conversion.triggered.connect(self.call_conversion)
        self.sel_consumable_receipt.triggered.connect(self.call_consumable_receipt)
        self.sel_consumable_sales.triggered.connect(self.call_consumable_sales)
        self.sel_consumable_inventory.triggered.connect(self.call_inventory)
        self.sel_closing_receipt.triggered.connect(self.call_consumable_closing_receipt)
        self.sel_closing_sales.triggered.connect(self.call_consumable_closing_sales)

        # Salary Handling
        self.sel_salary_calmaster.triggered.connect(self.call_calmaster)
        self.sel_salary_inh_ot.triggered.connect(self.call_inh_ot)
        self.sel_salary_outh_ot.triggered.connect(self.call_outh_ot)
        self.sel_salary_inh_basic.triggered.connect(self.call_inh_basic)
        self.sel_salary_outh_basic.triggered.connect(self.call_outh_basic)    
        self.sel_salary_pension.triggered.connect(self.call_pension)
        self.sel_salary_employeeinfo.triggered.connect(self.call_employeeinfo)
        self.sel_absenteeism.triggered.connect(self.call_absenteeism)
        
        # Advance Payment handling
        self.sel_advancepay.triggered.connect(self.call_advancepay)

        # Apt management
        self.sel_apt_master.triggered.connect(self.call_apt_master)
        self.sel_apt_contact.triggered.connect(self.call_apt_contact)
        self.sel_apt_address.triggered.connect(self.call_apt_address)
        self.sel_apt_contract.triggered.connect(self.call_apt_contract)
        self.sel_apt_contractpic.triggered.connect(self.call_apt_contractpic)        
        self.sel_apt_report.triggered.connect(self.call_apt_report)
        self.sel_reinfo.triggered.connect(self.call_reinfo)
        self.sel_aptcloth_btb.triggered.connect(self.call_aptcloth_btb)
        

        # Store references to open sub-windows to avoid duplication of sub windows
        self.open_sub_windows = []

        # Arrange the sub-windows as cascade manner
        self.action_cascade.triggered.connect(self.cascade_sub_windows)
   
        # Create a QTimer to update the status bar with the current date and time
        self.datetime_timer = QTimer(self)
        self.datetime_timer.timeout.connect(self.update_datetime)
        self.datetime_timer.start(1000)  # Update every 1 second

        # Initial value for current_datetime
        self.current_datetime = QDateTime.currentDateTime()

        # Call set_username to update the status bar
        self.set_username(username)

        # Store the current username as an instance variable
        self.current_username = username
       
        # Make log file
        self.make_logfiles("access_main.log")    

    def updatedbs(self):
        # Move employee and customer table from araps to consumables
        QMessageBox.about(self, "DB Update", "데이터베이스를 최신 정보로 업데이트 할 것입니다.\n 확인을 누르고 데이터베이스가 업데이트 될 때까지 잠시 기다려주십시오!")
        self.ini_move_data()
        QMessageBox.about(self, "DB Updated", "데이터베이스가 최신 정보로 업데이트 되었습니다.")

    def update_datetime(self):
        # Update the current_datetime with the current date and time
        self.current_datetime = QDateTime.currentDateTime()
        self.set_username(self.current_username)

    def set_username(self, username):
        # Format the current_datetime in the desired format
        formatted_datetime = self.current_datetime.toString("yyyy년 MM월 dd일 a hh:mm:ss")        
        self.statusBar().showMessage(f"Logged in as {username} | Access Date-Time: {formatted_datetime}")

    # Function to handle menu item actions to avoid duplication of sub windows
    def handle_menu_item(self, menu_item, dialog_class):
        for sub_window in self.open_sub_windows:
            if isinstance(sub_window.widget(), type(dialog_class)):
                # If the dialog is already open, bring it to the front
                self.mdi_area.setActiveSubWindow(sub_window)
                return

        # If the dialog is not open, create a new one
        self.make_dialog(dialog_class)

        # Update the status bar here (you can call set_username or any other appropriate method)
        self.set_username(self.current_username)

    # Call make_dialog --------------------------------------------------------
    def call_dialog(self, dialog_class, sel):
        dialog = dialog_class(self.current_username, current_datetime)
        self.handle_menu_item(sel, dialog)
        logging.info(f"User {self.current_username} called {dialog_class.__name__}.")

    # First db ----------------------------------------------------------------------
    def call_customer(self):
        self.call_dialog(CustomerDialog, self.sel_customer)
    def call_customercar(self):
        self.call_dialog(CustomerCarDialog, self.sel_customercar)
    def call_customercorpno(self):
        self.call_dialog(CustomerRegNoDialog, self.sel_cuscorpno)
    def call_customeraddress(self):
        self.call_dialog(CustomerAddressDialog, self.sel_customeraddress)
    def call_customerbkaccno(self):
        self.call_dialog(CustomerBkAccDialog, self.sel_customerbkaccno)
    def call_employee(self):
        self.call_dialog(EmployeeDialog, self.sel_employee)
    def call_employeecar(self):
        self.call_dialog(EmployeeCarDialog, self.sel_employeecar)
    def call_delivery(self):
        self.call_dialog(DeliveryDialog, self.sel_delivery)
    def call_nr_salesitem(self):
        self.call_dialog(SalesItemDialog, self.sel_nr_salesitem)
    def call_nr_salesprice(self):
        self.call_dialog(SalesPriceDialog, self.sel_nr_salesprice)
    def call_ap_product(self):
        self.call_dialog(ApProductDialog, self.sel_ap_product)
    def call_ap_cost(self):
        self.call_dialog(CostDialog, self.sel_ap_cost)
    def call_advancepay(self):
        self.call_dialog(AdvancePayDialog, self.sel_advancepay)
    def call_regno(self):
        self.call_dialog(EmployeeRegNoDialog, self.sel_regno)
    def call_phoneno(self):
        self.call_dialog(EmployeePhoneNoDialog, self.sel_phoneno)
    def call_empaddress(self):
        self.call_dialog(EmployeeAddressDialog, self.sel_address)
    def call_cusphoneno(self):
        self.call_dialog(CustomerPhoneNoDialog, self.sel_cusphoneno)
    def call_bankaccno(self):
        self.call_dialog(BankAccProductInfoDialog, self.sel_bankaccno)
    def call_genbankaccno(self):
        self.call_dialog(BankAccProductGenDialog, self.sel_genbkaccproduct)
    def call_genbankacccheck(self):
        self.call_dialog(BankAccProductChkDialog, self.sel_genbkacccheck)
    def call_paymentproduct(self):
        self.call_dialog(PaymentProductDialog, self.sel_paymentproduct)
    def call_gpaiinfo(self):
        self.call_dialog(GPAIInfoDialog, self.sel_gpaiinfo)
    def call_gpaigen(self):
        self.call_dialog(GPAIGenDialog, self.sel_gpaigen)
    def call_gpaiview(self):
        self.call_dialog(GPAIViewDialog, self.sel_gpaiview)
    def call_oilusage(self):
        self.call_dialog(OilUsageInfoDialog, self.sel_oilusage)
    def call_oilusageclosing(self):
        self.call_dialog(OilUsageGenDialog, self.sel_oilusageclosing)
    def call_oilusagecom(self):
        self.call_dialog(OilUsageViewDialog, self.sel_oilusagecom)
    def call_oilusageemp(self):
        self.call_dialog(OilUsageEmpViewDialog, self.sel_oilusageemp)
    def call_supportaptinfo(self):
        self.call_dialog(SupportAptInfoDialog, self.sel_supportaptinfo)
    def call_supportaptgen(self):
        self.call_dialog(SupportAptGenDialog, self.sel_supportaptgen)
    def call_supportaptview(self):
        self.call_dialog(SupportAptEmpViewDialog, self.sel_supportaptview)

    # Second db ----------------------------------------------------------------------
    def call_consumable_product(self):
        self.call_dialog(ConsumableProductDialog, self.sel_consumable_product)
    def call_consumable_in_price(self):
        self.call_dialog(ConsumableInPriceDialog, self.sel_consumable_inprice)
    def call_consumable_out_price(self):
        self.call_dialog(ConsumableOutPriceDialog, self.sel_consumable_outprice)
    def call_conversion(self):
        self.call_dialog(ConversionDialog, self.sel_conversion)
    def call_consumable_receipt(self):
        self.call_dialog(ConsumableReceiptDialog, self.sel_consumable_receipt)
    def call_consumable_sales(self):
        self.call_dialog(ConsumableSalesDialog, self.sel_consumable_sales)
    def call_inventory(self):
        self.call_dialog(CurrentInventoryDialog, self.sel_consumable_inventory)
    def call_consumable_closing_receipt(self):
        self.call_dialog(ConsumableInClosingDialog, self.sel_closing_receipt)
    def call_consumable_closing_sales(self):
        self.call_dialog(ConsumableOutClosingDialog, self.sel_closing_sales)

    # Third db ----------------------------------------------------------------------
    def call_calmaster(self):
        self.call_dialog(CalMasterDialog, self.sel_salary_calmaster)
    def call_inh_ot(self):
        self.call_dialog(CalWorkingHrInhOtDialog, self.sel_salary_inh_ot)
    def call_outh_ot(self):
        self.call_dialog(CalWorkingHrOuthOtDialog, self.sel_salary_outh_ot)
    def call_inh_basic(self):
        self.call_dialog(CalSalaryInhDialog, self.sel_salary_inh_basic)
    def call_outh_basic(self):
        self.call_dialog(CalSalaryOuthDialog, self.sel_salary_outh_basic)
    def call_pension(self):
        self.call_dialog(CalPensionDialog, self.sel_salary_pension)
    def call_employeeinfo(self):
        self.call_dialog(CalEmployeeInfoDialog, self.sel_salary_employeeinfo)
    def call_absenteeism(self):
        self.call_dialog(CalAbsenteeismDialog, self.sel_absenteeism)

    # Foutth db ---------------------------------------------------------------------
    def call_apt_master(self):
        self.call_dialog(AptMasterDialog, self.sel_apt_master)
    def call_apt_contact(self):
        self.call_dialog(AptContactDialog, self.sel_apt_contact)
    def call_apt_address(self):
        self.call_dialog(AptAddressDialog, self.sel_apt_address)        
    def call_apt_contract(self):
        self.call_dialog(AptContractDialog, self.sel_apt_contract)
    def call_apt_contractpic(self):
        self.call_dialog(AptContractPicDialog, self.sel_apt_contractpic)
    def call_apt_report(self):
        self.call_dialog(AptReportDialog, self.sel_apt_report)
    def call_reinfo(self):
        self.call_dialog(RecyclingDialog, self.sel_reinfo)
    def call_aptcloth_btb(self):
        self.call_dialog(AptClothBtbDialog, self.sel_aptcloth_btb)

    # Arrange the child windows in a cascading manner----------------------------------
    def cascade_sub_windows(self):
        self.mdi_area.cascadeSubWindows()

    # Open Excel Files Folder ---------------------------------------------------------
    def excel_araps_open(self):
        file_filter = "Excel Files (*.xls *.xlsx *.xlsm)"
        # in case of excel files are located under the program file folder named arap
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open Excel File', './dbs', file_filter) 

        if file_path:
            try:
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = True  
                excel.Workbooks.Open(file_path)
            except Exception as e:
                # Handle any errors that might occur during COM Automation.
                print(f"Error opening Excel file: {str(e)}")

    def close_connection(self, cursor, connection):
        cursor.close()
        connection.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    login_window = LoginWindow()
    
    if login_window.exec() == QDialog.Accepted:
        username = login_window.username_entry.text()  # Get the username from the login window
        main_window = MainWindow(username)
        main_window.show()
        sys.exit(app.exec())    