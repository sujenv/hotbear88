import sys
import re
import math
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QInputDialog, QWidget, QMenu, QShortcut, QLineEdit, QComboBox
from datetime import datetime
from commonmd import *
from cal import CalendarView
#<--for non_ui version-->
#from aptclothbtb_ui import UI_AptClothBtbDialog

# Dialog and Import common modules -----------------------------------------------------
class AptClothBtbDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class AptClothBtbDialog(QDialog, UI_AptClothBtbDialog, SubWindowBase):

    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database4()

        # Load ui file
        uic.loadUi("aptclothbtb.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)

        # Create tv_aptclothbtb and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "", "", "numeric", "numeric", "", "numeric", "", "numeric", "", "", "", "numeric", "numeric", "numeric", "numeric", "", "numeric", "", "", "", "", "", ""]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_aptclothbtb)
        self.tv_aptclothbtb.setItemDelegate(delegate)
        self.tv_aptclothbtb.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_aptclothbtb.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_aptclothbtb.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_aptclothbtb.verticalHeader().setVisible(False)

        # While selecting row in tv_aptclothbtb, each cell values to displayed to designated widgets
        self.tv_aptclothbtb.clicked.connect(self.show_selected_data)

        # Fill combobox items when the application starts
        self.get_combobox_contents()

        # Initiate display of data
        self.make_data()
        self.conn_button_to_method()
        self.connect_signal_slot()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Create context menus ----------------------------------------------------------------------------------
        self.context_menu1 = self.create_context_menu(self.entry_aptclothbtb_cefffrom)
        self.context_menu2 = self.create_context_menu(self.entry_aptclothbtb_ceffthru)
        self.context_menu3 = self.create_context_menu(self.entry_aptclothbtbchange_cefffrom)
        self.context_menu4 = self.create_context_menu(self.entry_aptclothbtbchange_ceffthru)

        self.entry_aptclothbtb_cefffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptclothbtb_cefffrom.customContextMenuRequested.connect(self.show_context_menu1)

        self.entry_aptclothbtb_ceffthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptclothbtb_ceffthru.customContextMenuRequested.connect(self.show_context_menu2)

        self.entry_aptclothbtbchange_cefffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptclothbtbchange_cefffrom.customContextMenuRequested.connect(self.show_context_menu3)

        self.entry_aptclothbtbchange_ceffthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptclothbtbchange_ceffthru.customContextMenuRequested.connect(self.show_context_menu4)
        #-----------------------------------------------------------------------------------------------------------------

        self.entry_stylesheet_as_is()
        self.hide_aptclothbtb_change_widget()

        # Make log file
        self.make_logfiles("access_AptClothBtb.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_aptclothbtb, partial(self.copy_cells, self.tv_aptclothbtb))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_aptclothbtb, partial(self.paste_cells, self.tv_aptclothbtb))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_aptclothbtb, partial(self.handle_return_key, self.tv_aptclothbtb))

    # Change styles for the selected widgets 1
    def entry_stylesheet_as_is(self):
        entry_prefix = 'entry_aptclothbtbchange_'
        cb_prefix = 'cb_aptclothbtbchange_'

        entry_widgets = ['customercode', 'concode', 'cefffrom', 'ceffthru', 'ucost', 'contractvalue',
                        'paytycode', 'issuevoucherdt', 'vouchercleardt', 'remark']

        cb_widgets = ['customername', 'conname', 'paytyname', 'issuevoucher', 'paymentclear']

        for widget_name in entry_widgets:
            widget = getattr(self, f"{entry_prefix}{widget_name}")
            widget.setStyleSheet('color:black;background:rgb(0,0,0)')

        for widget_name in cb_widgets:
            widget = getattr(self, f"{cb_prefix}{widget_name}")
            widget.setStyleSheet('color:black;background:rgb(0,0,0)')
            
    # Change styles for the selected widgets 2
    def entry_stylesheet_to_be(self):
        entry_prefix = 'entry_aptclothbtbchange_'
        cb_prefix = 'cb_aptclothbtbchange_'

        entry_widgets = ['customercode', 'concode', 'cefffrom', 'ceffthru', 'ucost', 'contractvalue',
                        'paytycode', 'issuevoucherdt', 'vouchercleardt', 'remark']

        cb_widgets = ['customername', 'conname', 'paytyname', 'issuevoucher', 'paymentclear']

        for widget_name in entry_widgets:
            widget = getattr(self, f"{entry_prefix}{widget_name}")
            widget.setStyleSheet('color:black;background:rgb(255,255,255)')

        for widget_name in cb_widgets:
            widget = getattr(self, f"{cb_prefix}{widget_name}")
            widget.setStyleSheet('color:black;background:rgb(255,255,0)')        

        self.entry_aptclothbtbchange_cefffrom.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptclothbtbchange_ceffthru.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptclothbtbchange_ucost.setStyleSheet('color:black;background:rgb(255,0,0)')
        self.entry_aptclothbtbchange_vouchercleardt.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.cb_aptclothbtbchange_paymentclear.setStyleSheet('color:white;background:rgb(255,0,0)')

    # Show widgets for the cost change parts 
    def show_aptclothbtb_change_widget(self):
        ddt, endofdate = self.display_eff_date()
        self.pb_aptclothbtbchange_update_insert.setVisible(True)

        entry_prefix = 'entry_aptclothbtbchange_'
        cb_prefix = 'cb_aptclothbtbchange_'

        entry_widgets = ['customercode', 'concode', 'cefffrom', 'ceffthru', 'ucost', 'contractvalue',
                        'paytycode', 'issuevoucherdt', 'vouchercleardt', 'remark']

        cb_widgets = ['customername', 'conname', 'paytyname', 'issuevoucher', 'paymentclear']

        for widget_name in entry_widgets + cb_widgets:
            widget = getattr(self, f"{entry_prefix if widget_name in entry_widgets else cb_prefix}{widget_name}")

            # Check the type of the widget
            if isinstance(widget, QLineEdit):  # Assuming QLineEdit for entry widgets
                widget.setReadOnly(False)
            elif isinstance(widget, QComboBox):  # Assuming QComboBox for combobox widgets
                widget.setEnabled(True) 
                
        self.entry_aptclothbtbchange_ceffthru.setText(endofdate)
        self.entry_stylesheet_to_be()

    # Hide widgets for the cost change parts 
    def hide_aptclothbtb_change_widget(self):
        self.pb_aptclothbtbchange_update_insert.setVisible(False)
        
        entry_prefix = 'entry_aptclothbtbchange_'
        cb_prefix = 'cb_aptclothbtbchange_'

        entry_widgets = ['customercode', 'concode', 'cefffrom', 'ceffthru', 'ucost', 'contractvalue',
                        'paytycode', 'issuevoucherdt', 'vouchercleardt', 'remark']

        cb_widgets = ['customername', 'conname', 'paytyname', 'issuevoucher', 'paymentclear']

        for widget_name in entry_widgets + cb_widgets:
            widget = getattr(self, f"{entry_prefix if widget_name in entry_widgets else cb_prefix}{widget_name}")
            
            # Check the type of the widget
            if isinstance(widget, QLineEdit):  # Assuming QLineEdit for entry widgets
                widget.setReadOnly(True)
            elif isinstance(widget, QComboBox):  # Assuming QComboBox for combobox widgets
                widget.setEnabled(False) 

        self.entry_stylesheet_as_is()

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu

    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_aptclothbtb_cefffrom.mapToGlobal(pos))
    def show_context_menu2(self, pos):
        self.context_menu2.exec_(self.entry_aptclothbtb_ceffthru.mapToGlobal(pos))
    def show_context_menu3(self, pos):
        self.context_menu3.exec_(self.entry_aptclothbtbchange_cefffrom.mapToGlobal(pos))
    def show_context_menu4(self, pos):
        self.context_menu4.exec_(self.entry_aptclothbtbchange_ceffthru.mapToGlobal(pos))

    # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_aptclothbtb_cefffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptclothbtb_ceffthru:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptclothbtbchange_cefffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptclothbtbchange_ceffthru:
            target_lineedit.setText(date)            
    #-----------------------------------------------------------------------------------------------------------------            

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_aptclothbtb
        self.process_key_event(event, tv_widget)

    # Display current date only
    def display_eff_date(self):
        now = datetime.now()
        curr_date = now.strftime("%Y/%m/%d")
        ddt = f"{curr_date}"         
        endofdate = "2050/12/31"

        return ddt, endofdate
    
    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_aptclothbtb_aname, "SELECT DISTINCT adesc FROM apt_master ORDER BY adesc")
        self.insert_combobox_initiate(self.cb_aptclothbtb_apttypename, "SELECT DISTINCT description FROM apt_type")
        self.insert_combobox_initiate(self.cb_aptclothbtb_customername, "SELECT DISTINCT cname FROM customer")
        self.insert_combobox_initiate(self.cb_aptclothbtb_conname, "SELECT DISTINCT contydescr FROM cloth_contracttype ORDER BY contydescr")
        self.insert_combobox_initiate(self.cb_aptclothbtb_paytyname, "SELECT DISTINCT paytydescr FROM cloth_paymenttype ORDER BY paytydescr")
        self.insert_combobox_initiate(self.cb_aptclothbtb_paymentclear, "SELECT DISTINCT class2 FROM employee ORDER BY class2 DESC")
        self.insert_combobox_initiate(self.cb_aptclothbtb_issuevoucher, "SELECT DISTINCT class2 FROM employee ORDER BY class2 DESC")

        self.insert_combobox_initiate(self.cb_aptclothbtbchange_customername, "SELECT DISTINCT cname FROM customer")
        self.insert_combobox_initiate(self.cb_aptclothbtbchange_conname, "SELECT DISTINCT contydescr FROM cloth_contracttype ORDER BY contydescr")
        self.insert_combobox_initiate(self.cb_aptclothbtbchange_paytyname, "SELECT DISTINCT paytydescr FROM cloth_paymenttype ORDER BY paytydescr")
        self.insert_combobox_initiate(self.cb_aptclothbtbchange_paymentclear, "SELECT DISTINCT class2 FROM employee ORDER BY class2 DESC")
        self.insert_combobox_initiate(self.cb_aptclothbtbchange_issuevoucher, "SELECT DISTINCT class2 FROM employee ORDER BY class2 DESC")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) # using common module
        self.lbl_aptclothbtb_id.setText("")
        self.entry_aptclothbtb_acode.setText("")
        self.cb_aptclothbtb_aname.setCurrentIndex(0) 
        self.cb_aptclothbtb_apttypename.setCurrentIndex(0)
        self.cb_aptclothbtb_customername.setCurrentIndex(0)
        self.cb_aptclothbtb_conname.setCurrentIndex(0)
        self.cb_aptclothbtb_paytyname.setCurrentIndex(0)
        self.cb_aptclothbtb_paymentclear.setCurrentIndex(0)
        self.cb_aptclothbtb_issuevoucher.setCurrentIndex(0)

        self.cb_aptclothbtbchange_customername.setCurrentIndex(0)
        self.cb_aptclothbtbchange_conname.setCurrentIndex(0)
        self.cb_aptclothbtbchange_paytyname.setCurrentIndex(0)
        self.cb_aptclothbtbchange_paymentclear.setCurrentIndex(0)
        self.cb_aptclothbtbchange_issuevoucher.setCurrentIndex(0)

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_aptclothbtb_show.clicked.connect(self.make_data)
        self.pb_aptclothbtb_show_con.clicked.connect(self.make_data_con)
        self.pb_aptclothbtb_search.clicked.connect(self.search_data)
        self.pb_aptclothbtb_clear_data.clicked.connect(self.clear_data)
        self.pb_aptclothbtb_close.clicked.connect(self.close_dialog)
        self.pb_aptclothbtb_insert.clicked.connect(self.tv_insert)
        self.pb_aptclothbtb_update.clicked.connect(self.SelectionMessageBox)
        self.pb_aptclothbtb_delete.clicked.connect(self.tv_delete)
        self.pb_aptclothbtbchange_update_insert.clicked.connect(self.reflect_contract_change)


    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_aptclothbtb_aname.activated.connect(self.cb_aptclothbtb_aname_changed)
        self.cb_aptclothbtb_apttypename.activated.connect(self.cb_aptclothbtb_apttypename_changed)        
        self.cb_aptclothbtb_customername.activated.connect(self.cb_aptclothbtb_customername_changed)
        self.cb_aptclothbtb_conname.activated.connect(self.cb_aptclothbtb_conname_changed)
        self.cb_aptclothbtb_paytyname.activated.connect(self.cb_aptclothbtb_paytyname_changed)
        self.cb_aptclothbtb_issuevoucher.activated.connect(self.cb_aptclothbtb_issuevoucher_change)
        self.entry_aptclothbtb_cefffrom.editingFinished.connect(self.entry_aptclothbtb_cefffrom_change)
        self.entry_aptclothbtb_contractvalue.editingFinished.connect(self.entry_aptclothbtb_contractvalue_change)
        
        self.cb_aptclothbtbchange_customername.activated.connect(self.cb_aptclothbtbchange_customername_changed)
        self.cb_aptclothbtbchange_conname.activated.connect(self.cb_aptclothbtbchange_conname_changed)
        self.entry_aptclothbtbchange_cefffrom.editingFinished.connect(self.entry_aptclothbtbchange_cefffrom_change)
        self.entry_aptclothbtbchange_ucost.editingFinished.connect(self.entry_aptclothbtbchange_ucost_change)
        self.cb_aptclothbtbchange_paytyname.activated.connect(self.cb_aptclothbtbchange_paytyname_changed)

    # Tab order for sub window
    def set_tab_order(self):
        widgets = [self.pb_aptclothbtb_show, self.cb_aptclothbtb_aname, self.cb_aptclothbtb_apttypename,
            self.cb_aptclothbtb_customername, self.cb_aptclothbtb_conname, self.entry_aptclothbtb_cefffrom, 
            self.entry_aptclothbtb_ceffthru, self.entry_aptclothbtb_contractvalue, self.cb_aptclothbtb_paytyname,
            self.cb_aptclothbtb_issuevoucher, self.entry_aptclothbtb_issuevoucherdt, self.entry_aptclothbtb_vouchercleardt,
            self.cb_aptclothbtb_paymentclear, self.entry_aptclothbtb_remark,

            self.cb_aptclothbtbchange_customername, self.cb_aptclothbtbchange_conname, self.entry_aptclothbtbchange_cefffrom,
            self.entry_aptclothbtbchange_ceffthru, self.entry_aptclothbtbchange_ucost, self.cb_aptclothbtbchange_paytyname,
            self.cb_aptclothbtbchange_issuevoucher, self.entry_aptclothbtbchange_issuevoucherdt, self.entry_aptclothbtbchange_vouchercleardt,
            self.cb_aptclothbtbchange_paymentclear, self.entry_aptclothbtbchange_remark,

            self.pb_aptclothbtb_show_con, self.pb_aptclothbtb_search, self.pb_aptclothbtb_clear_data,
            self.pb_aptclothbtb_close, self.pb_aptclothbtb_insert, self.pb_aptclothbtb_update, 
            self.pb_aptclothbtb_delete]

        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_aptclothbtb
        self.cursor.execute("Select * From vw_clothbtb WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]    

        sql_query = "Select * From vw_clothbtb order by id"
        column_widths = [80, 100, 250, 80, 80, 80, 80, 120, 80, 80, 100, 100, 80, 100, 80, 80, 80, 80, 80, 80, 100, 100, 80, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # Make table data
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Make table data
    def make_data_con(self):
        ddt, ddt_1 = disply_date_info()
        query = f"Select * From vw_clothbtb Where 계약시작일 <= #{ddt}# and #{ddt}# <= 계약종료일 order by id"
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_aptclothbtb_input(self):
        acode = str(self.entry_aptclothbtb_acode.text())
        ccode = int(self.entry_aptclothbtb_customercode.text())
        noh = int(self.entry_aptclothbtb_noh.text())
        ctype = int(self.entry_aptclothbtb_apttypecode.text())

        date_format_pattern = r'\d{4}/\d{2}/\d{2}'
        srtdt = str(self.entry_aptclothbtb_cefffrom.text())
        efffrom = srtdt if re.match(date_format_pattern, srtdt) else "2024/01/01"
        enddt = str(self.entry_aptclothbtb_ceffthru.text())
        effthru = enddt if re.match(date_format_pattern, srtdt) else "2024/12/31"

        input_val = self.entry_aptclothbtb_contractvalue.text()
        if input_val.replace(".", "").replace("-", "").isdigit():
            conval = float(input_val)
        else:
            conval = 0
            
        ptype = int(self.entry_aptclothbtb_paytycode.text())
        isuvou = str(self.cb_aptclothbtb_issuevoucher.currentText())
        voupr = str(self.entry_aptclothbtb_issuevoucherdt.text())
        paydt = str(self.entry_aptclothbtb_vouchercleardt.text())
        paycomp = str(self.cb_aptclothbtb_paymentclear.currentText())
        remark = str(self.entry_aptclothbtb_remark.text())
        
        return acode, ccode, noh, ctype, efffrom, effthru, conval, ptype, isuvou, voupr, paydt, paycomp, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()
        return username, user_id, formatted_datetime

    # insert new product data to MySQL table
    def tv_insert(self):
        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            idx = self.max_row_id("apt_clothbtb")
            username, user_id, formatted_datetime = self.common_values_set()
            acode, ccode, noh, ctype, efffrom, effthru, conval, ptype, isuvou, voupr, paydt, paycomp, remark = self.get_aptclothbtb_input() 

            if (idx>0 and conval>0 ) and all(len(var) > 0 for var in (acode, efffrom, effthru, voupr)):
                self.cursor.execute('''INSERT INTO apt_clothbtb 
                            (id, acode, ccode, nofreg, contracttype, efffrom, effthru, contractvalue, paymenttype, issuevoucher, voucherprinted, paymentdate, paymentcompleted, trxdate, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, acode, ccode, noh, ctype, efffrom, effthru, conval, ptype, isuvou, voupr, paydt, paycomp, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data()
                logging.info(f"User {username} inserted {idx} row to the apt clothbtb table.")
            else:
                self.show_missing_message("입력 이상")
                return
        else:
            self.show_cancel_message("데이터 추가 취소")
            return    

    # update 조건에 따라 분기 할 것
    def SelectionMessageBox(self):

        # If there's no selection
        if len(self.lbl_aptclothbtb_id.text()) == 0:
            self.show_missing_message_update("입력 확인")

        # In case of row selection

        conA = '''계약 내용 중 오류 수정 - 현재 행을 수정, 추가 행을 만들지 않음!'''
        conB = '''계약 내용의 변경 또는 갱신 - 현재 행은 변경 없음, 변경된 내용으로 추가 행을 만듦!'''
        
        conditions = [conA, conB]
        condition, okPressed = QInputDialog.getItem(self, "입력 조건 선택", "어떤 작업을 진행하시겠습니까?:", conditions, 0, False)

        if okPressed:
            if condition == conA:
                self.fix_typo()     
            elif condition == conB:
                self.show_aptclothbtb_change_widget()
            else:
                return

    # Fix typing error in the contract with customer
    def fix_typo(self):
        
        confirm_dialog = self.show_update_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            
            username, user_id, formatted_datetime = self.common_values_set()
            acode, ccode, noh, ctype, efffrom, effthru, conval, ptype, isuvou, voupr, paydt, paycomp, remark = self.get_aptclothbtb_input() 
            idx = int(self.lbl_aptclothbtb_id.text())
            
            if (idx>0 and conval>0 ) and all(len(var) > 0 for var in (acode, efffrom, effthru, voupr)):
                self.cursor.execute('''UPDATE apt_clothbtb SET 
                        acode=?, ccode=?, nofreg=?, contracttype=?, efffrom=?, effthru=?, contractvalue=?, paymenttype=?, issuevoucher=?, voucherprinted=?, paymentdate=?, paymentcompleted=?, trxdate=?, userid=?, remark=? WHERE id=?'''
                        , (acode, ccode, noh, ctype, efffrom, effthru, conval, ptype, isuvou, voupr, paydt, paycomp, formatted_datetime, user_id, remark, idx))
            else:
                self.show_missing_message("입력 이상")
                return

            self.conn.commit()
            self.show_update_success_message()
            self.refresh_data()
            logging.info(f"User {username} updated row number {idx} in the apt clothbtb table.")
        else:
            self.show_cancel_message("데이터 변경 취소")
            return
        
    def execute_aptclothbtb_contract_change(self):
        ccode2 = str(self.entry_aptclothbtbchange_customercode.text())
        ctype2 = int(self.entry_aptclothbtbchange_concode.text())
        efffrom2 = str(self.entry_aptclothbtbchange_cefffrom.text())
        effthru2 = str(self.entry_aptclothbtbchange_ceffthru.text())
        conval2 = float(self.entry_aptclothbtbchange_contractvalue.text())
        ptype2 = str(self.entry_aptclothbtbchange_paytycode.text())
        isuvou2 = str(self.cb_aptclothbtbchange_issuevoucher.currentText())
        voupr2 = str(self.entry_aptclothbtbchange_issuevoucherdt.text())
        paydt2 = str(self.entry_aptclothbtbchange_vouchercleardt.text())
        paycomp2 = str(self.cb_aptclothbtbchange_paymentclear.currentText())
        remark2 = str(self.entry_aptclothbtbchange_remark.text())

        return ccode2, ctype2, conval2, efffrom2, effthru2, ptype2, isuvou2, voupr2, paydt2, paycomp2, remark2

    def reflect_contract_change(self):
        kaptcode = str(self.entry_aptclothbtb_acode.text())
        username, user_id, formatted_datetime = self.common_values_set()
        acode, ccode, noh, ctype, efffrom, effthru, conval, ptype, isuvou, voupr, paydt, paycomp, remark = self.get_aptclothbtb_input()
        ccode2, ctype2, conval2, efffrom2, effthru2, ptype2, isuvou2, voupr2, paydt2, paycomp2, remark2  = self.execute_aptclothbtb_contract_change()

        idx = int(self.max_row_id("apt_clothbtb"))        
                
        if (idx>0 and ctype2>0 and abs(conval2)>0) and all(len(var) > 0 for var in (ccode2, efffrom2, effthru2)):

            org_id = str(self.lbl_aptclothbtb_id.text())
            effthru1 = str(self.entry_aptclothbtb_ceffthru.text()) # 다시 불러와야 함..
            # 기존 id의 유효종료일을 변경유효시작일 -1 일로 수정
            self.cursor.execute('''UPDATE apt_clothbtb SET effthru=? WHERE id=?''', (effthru1, org_id))
            
            #변경된 내용을 신규로 추가
            self.cursor.execute('''INSERT INTO apt_clothbtb 
                    (id, acode, ccode, nofreg, contracttype, efffrom, effthru, contractvalue, paymenttype, issuevoucher, voucherprinted, paymentdate, paymentcompleted, trxdate, userid, remark) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                    , (idx, acode, ccode2, noh, ctype2, efffrom2, effthru2, conval2, ptype2, isuvou2, voupr2, paydt2, paycomp2, formatted_datetime, user_id, remark2))
            
            self.conn.commit()
            self.show_update_success_message()
            self.refresh_data()
            logging.info(f"User {username} updated row number {idx} in the apt contract table.")
            
        else:
            self.show_missing_message("입력 이상")
            return

    # delete row according to id selected
    def tv_delete(self):
        confirm_dialog = self.show_delete_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = self.lbl_aptclothbtb_id.text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM apt_clothbtb WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()
            logging.info(f"User {username} deleted {idx} row to the apt contract cloth table.")  
        else:
            self.show_cancel_message("데이터 삭제 취소")
            return

    # Search data
    def search_data(self):
        aptname = self.cb_aptclothbtb_aname.currentText()
        cusname = self.cb_aptclothbtb_customername.currentText()
        conname= self.cb_aptclothbtb_conname.currentText()
        efffrom = self.entry_aptclothbtb_cefffrom.text()
        effthru = self.entry_aptclothbtb_ceffthru.text()

        conditions = {'v01': (aptname, "아파트명 like '%{}%'"), 
                    'v02': (cusname, "거래처명 like  '%{}%'"), 
                    'v03': (conname, "계약내용 like '%{}%'"),
                    'v04': (efffrom, "계약시작일 <= #{}#"),
                    'V05': (effthru, "계약종료일 >= #{}#"),
                    }
        
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_clothbtb WHERE {' AND '.join(selected_conditions)} ORDER BY id"

        QMessageBox.about(self, "검색 조건 확인", f"아파트명: {aptname} \n거래처명: {cusname} \n계약형태: {conname} \n계약시작일: {efffrom} \n계약종료일: {effthru} \n\n위 조건으로 검색을 수행합니다!")

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Combobox apt name index changed
    def cb_aptclothbtb_aname_changed(self):
        self.entry_aptclothbtb_acode.clear()
        self.entry_aptclothbtb_noh.clear()
        self.entry_aptclothbtb_apttypecode.clear()

        selected_item = self.cb_aptclothbtb_aname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode, nohousehold, typeofapt From vw_apt_info WHERE adesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtb_acode, self.entry_aptclothbtb_noh,
                                self.entry_aptclothbtb_apttypecode,]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
                self.entry_aptclothbtb_apttypecode_change()
            else:
                pass

    # Combobox apt type index changed
    def entry_aptclothbtb_apttypecode_change(self):
        self.cb_aptclothbtb_apttypename.setCurrentIndex(0)
        tycode = int(self.entry_aptclothbtb_apttypecode.text())

        if tycode >= 0:
            query = f"SELECT DISTINCT description From vw_apt_info WHERE typeofapt ={tycode}"

        self.cursor.execute(query)
        result = self.cursor.fetchone()

        if result:
            item01 = str(result[0])
            self.cb_aptclothbtb_apttypename.setCurrentText(item01)

    # Combobox apt type index changed
    def cb_aptclothbtb_apttypename_changed(self):
        self.entry_aptclothbtb_apttypecode.clear()
        selected_item = self.cb_aptclothbtb_apttypename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT code From apt_type WHERE description ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtb_apttypecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox customer index changed
    def cb_aptclothbtb_customername_changed(self):
        self.entry_aptclothbtb_customercode.clear()
        selected_item = self.cb_aptclothbtb_customername.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From customer WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtb_customercode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox contract discription index changed
    def cb_aptclothbtb_conname_changed(self):
        self.entry_aptclothbtb_concode.clear()
        selected_item = self.cb_aptclothbtb_conname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT contycode From cloth_contracttype WHERE contydescr ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtb_concode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox pay type description index changed
    def cb_aptclothbtb_paytyname_changed(self):
        self.entry_aptclothbtb_paytycode.clear()
        selected_item = self.cb_aptclothbtb_paytyname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT paytycode From cloth_paymenttype WHERE paytydescr ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtb_paytycode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # while contract change efffrom date, change effthru date
    def entry_aptclothbtb_cefffrom_change(self):
        chg_date_str = self.entry_aptclothbtb_cefffrom.text()

        try:
            chg_date = parse_date(chg_date_str)                             # 날짜 문자열을 날짜 객체로 변환
            org_date = chg_date + timedelta(days=365)                       # 날짜에서 365일을 더해줍니다.
            org_date_str = org_date.strftime('%Y/%m/%d')                    # 결과를 문자열로 변환
            self.entry_aptclothbtb_ceffthru.setText(org_date_str)           # 변경된 effthru 날짜를 표시

            start_date = self.entry_aptclothbtb_cefffrom.text()
            end_date = self.entry_aptclothbtb_ceffthru.text()
            months = self.month_difference(start_date, end_date)

            self.entry_aptclothbtb_term.setText(str(months))  # 결과를 문자열로 변환하여 표시

        except ValueError as e:
            QMessageBox.critical(self, "날짜 형식 오류", str(e))             # 날짜 형식이 잘못된 경우 사용자에게 알림


    def month_difference(self, start_date, end_date):
        try:
            start_date = datetime.strptime(start_date, "%Y-%m-%d")
            end_date = datetime.strptime(end_date, "%Y-%m-%d")
            end_date = end_date + timedelta(days=1)     # 그냥 계산하면 12개월을 11개월로 표기함으로 end_date에 1을 더해서 12로 표기

        except ValueError:
            try:
                start_date = datetime.strptime(start_date, "%Y/%m/%d")
                end_date = datetime.strptime(end_date, "%Y/%m/%d")
                end_date = end_date + timedelta(days=1)

            except ValueError:
                raise ValueError("Invalid date format. Please use either %Y-%m-%d or %Y/%m/%d.")

        month_diff = (end_date.year - start_date.year) * 12 + end_date.month - start_date.month

        return month_diff

    # Contract value changed
    def entry_aptclothbtb_contractvalue_change(self):
        self.entry_aptclothbtb_contractvat.clear()
        self.entry_aptclothbtb_contractttlval.clear()
        sval = float(self.entry_aptclothbtb_contractvalue.text())
        chkvat = self.cb_aptclothbtb_issuevoucher.currentText()

        if chkvat == "Y":
            vat =  sval * 0.1
            vat = math.floor(vat/1)*1
            
            ttlval = sval * 1.1
            ttlval = math.floor(ttlval/1)*1

        else:
            vat =  sval * 0
            ttlval = sval

        self.entry_aptclothbtb_contractvat.setText(str(vat))
        self.entry_aptclothbtb_contractttlval.setText(str(ttlval))

        text_noh = self.entry_aptclothbtb_noh.text()
        if text_noh is None or text_noh == "":
            noh = 1
        else:
            noh = float(text_noh)

        text_mths = self.entry_aptclothbtb_term.text()
        if text_mths is None or text_mths == "":
            mths = 12
        else:
            mths = float(text_mths)

        if noh == 1:
            r_ucost = 0
        else:
            r_ucost = sval / noh / mths
            r_ucost = math.floor(r_ucost/1)*1
        self.entry_aptclothbtb_ucost.setText(str(r_ucost))
    
        ddt, endofdate = self.display_eff_date()
        
        basedate = self.entry_aptclothbtb_cefffrom.text()

        if basedate is None or basedate == "":
            try:
                chgdate = datetime.strptime(basedate, "%Y/%m/%d")
            except ValueError:
                try:
                    chgdate = datetime.strptime(basedate, "%Y-%m-%d")
                except ValueError:
                    print("날짜 형식이 맞지 않습니다.")
            
            plustendt = chgdate + timedelta(days=7)

        else:
            try:
                chgdate = datetime.strptime(basedate, "%Y/%m/%d")
            except ValueError:
                try:
                    chgdate = datetime.strptime(basedate, "%Y-%m-%d")
                except ValueError:
                    print("날짜 형식이 맞지 않습니다.")

            plustendt = chgdate + timedelta(days=7)
        
        # Format the date as "YYYY/MM/DD"
        formatted_plustendt = plustendt.strftime("%Y/%m/%d")

        self.entry_aptclothbtb_issuevoucherdt.setText(str(formatted_plustendt))
        self.entry_aptclothbtb_vouchercleardt.setText(str(endofdate))

    # Status of Issue Voucher
    def cb_aptclothbtb_issuevoucher_change(self):
        chkv = self.cb_aptclothbtb_issuevoucher.currentText()
        if chkv == "N" or chkv == "":
            ddt, endofdate = self.display_eff_date()
            self.entry_aptclothbtb_issuevoucherdt.setText(str(endofdate))
            # 세액 = 0, 총액 = 공급가
            self.entry_aptclothbtb_contractvat.setText("0")
            sval = float(self.entry_aptclothbtb_contractvalue.text())
            sval = math.floor(sval/1)*1
            self.entry_aptclothbtb_contractttlval.setText(str(sval))

        else:
            return
        
    # Combobox customer index changed for contract revision
    def cb_aptclothbtbchange_customername_changed(self):
        self.entry_aptclothbtbchange_customercode.clear()
        selected_item = self.cb_aptclothbtbchange_customername.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From customer WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtbchange_customercode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox contract discription index changed for contract revision
    def cb_aptclothbtbchange_conname_changed(self):
        self.entry_aptclothbtbchange_concode.clear()
        selected_item = self.cb_aptclothbtbchange_conname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT contycode From cloth_contracttype WHERE contydescr ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtbchange_concode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # while contract change efffrom date, change effthru date for contract revision
    def entry_aptclothbtbchange_cefffrom_change(self):
        chg_date_str = self.entry_aptclothbtbchange_cefffrom.text()

        try:
            chg_date = parse_date(chg_date_str)                             # 날짜 문자열을 날짜 객체로 변환
            org_date = chg_date - timedelta(days=1)                         # 날짜에서 1일을 빼줍니다.
            org_date_str = org_date.strftime('%Y-%m-%d')                    # 결과를 문자열로 변환
            self.entry_aptclothbtb_ceffthru.setText(org_date_str)           # 변경된 effthru 날짜를 표시
        
            start_date = self.entry_aptclothbtbchange_cefffrom.text()
            chg_date = parse_date(start_date)                               # 날짜 문자열을 날짜 객체로 변환
            org_date = chg_date + timedelta(days=365)                       # 날짜에서 365일을 더해줍니다.
            org_date_str = org_date.strftime('%Y/%m/%d')                    # 결과를 문자열로 변환
            self.entry_aptclothbtbchange_ceffthru.setText(org_date_str)  

            ddt, endofdate = self.display_eff_date()
            self.entry_aptclothbtbchange_issuevoucherdt.setText(endofdate)
            self.entry_aptclothbtbchange_vouchercleardt.setText(endofdate)

        except ValueError as e:
            QMessageBox.critical(self, "날짜 형식 오류", str(e))             # 날짜 형식이 잘못된 경우 사용자에게 알림

    # Unit cost change for contract revision
    def entry_aptclothbtbchange_ucost_change(self):
        
        # Number of residence
        noh = int(self.entry_aptclothbtb_noh.text())
        
        # Months count
        start_date = self.entry_aptclothbtbchange_cefffrom.text()
        end_date = self.entry_aptclothbtbchange_ceffthru.text()
        
        start_date = datetime.strptime(start_date, "%Y/%m/%d")
        end_date = datetime.strptime(end_date, "%Y/%m/%d")
        end_date = end_date + timedelta(days=1)

        term = (end_date.year - start_date.year) * 12 + end_date.month - start_date.month

        # unit cost
        ucost = int(self.entry_aptclothbtbchange_ucost.text())

        # supplied value
        convalamt = noh * term * ucost
        self.entry_aptclothbtbchange_contractvalue.setText(str(convalamt))


    # Payment type change for the contract revision
    def cb_aptclothbtbchange_paytyname_changed(self):
        self.entry_aptclothbtbchange_paytycode.clear()
        selected_item = self.cb_aptclothbtbchange_paytyname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT paytycode From cloth_paymenttype WHERE paytydescr ='{selected_item}'"
            line_edit_widgets = [self.entry_aptclothbtbchange_paytycode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # clear input field entry
    def clear_data(self):
        self.lbl_aptclothbtb_id.setText("")
        clear_widget_data(self)

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(24):  # 24columns
            cell_text = self.tv_aptclothbtb.model().item(row_index, column_index).text()
            cell_values.append(cell_text)
            
        # Populate the input widgets with the data from the selected row
        self.lbl_aptclothbtb_id.setText(cell_values[0])
        self.entry_aptclothbtb_acode.setText(cell_values[1])
        self.cb_aptclothbtb_aname.setCurrentText(cell_values[2])
        self.entry_aptclothbtb_noh.setText(cell_values[3])
        self.entry_aptclothbtb_apttypecode.setText(cell_values[4])
        self.cb_aptclothbtb_apttypename.setCurrentText(cell_values[5])
        self.entry_aptclothbtb_customercode.setText(cell_values[6])
        self.cb_aptclothbtb_customername.setCurrentText(cell_values[7])
        self.entry_aptclothbtb_concode.setText(cell_values[8])
        self.cb_aptclothbtb_conname.setCurrentText(cell_values[9])
        self.entry_aptclothbtb_cefffrom.setText(cell_values[10])
        self.entry_aptclothbtb_ceffthru.setText(cell_values[11])
        self.entry_aptclothbtb_term.setText(cell_values[12])
        self.entry_aptclothbtb_contractvalue.setText(cell_values[13])
        self.entry_aptclothbtb_contractvat.setText(cell_values[14])
        self.entry_aptclothbtb_contractttlval.setText(cell_values[15])
        self.entry_aptclothbtb_ucost.setText(cell_values[16])
        self.entry_aptclothbtb_paytycode.setText(cell_values[17])
        self.cb_aptclothbtb_paytyname.setCurrentText(cell_values[18])
        self.cb_aptclothbtb_issuevoucher.setCurrentText(cell_values[19])
        self.entry_aptclothbtb_issuevoucherdt.setText(cell_values[20])
        self.entry_aptclothbtb_vouchercleardt.setText(cell_values[21])
        self.cb_aptclothbtb_paymentclear.setCurrentText(cell_values[22])
        self.entry_aptclothbtb_remark.setText(cell_values[23])


    def refresh_data(self):
        self.clear_data()
        self.make_data()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_AptClothBtb.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = AptClothBtbDialog()
    dialog.show()
    sys.exit(app.exec())