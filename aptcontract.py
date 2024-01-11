import sys
import re
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QInputDialog, QWidget, QMenu, QShortcut
from datetime import datetime
from commonmd import *
from cal import CalendarView
#<--for non_ui version-->
#from aptcontract_ui import UI_AptContractDialog

# Dialog and Import common modules -----------------------------------------------------
class AptContractDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class AptContractDialog(QDialog, UI_AptContractDialog, SubWindowBase):

    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database4()

        # Load ui file
        uic.loadUi("aptcontract.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)

        # Create tv_aptcontract and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "", "", "", "", "", "", "", "", "", "", "", "", "", "",]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_aptcontract)
        self.tv_aptcontract.setItemDelegate(delegate)
        self.tv_aptcontract.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_aptcontract.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_aptcontract.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_aptcontract.verticalHeader().setVisible(False)

        # While selecting row in tv_aptcontract, each cell values to displayed to designated widgets
        self.tv_aptcontract.clicked.connect(self.show_selected_data)

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

        # Automatically input current date
        self.display_currentdate()

        # Create context menus ----------------------------------------------------------------------------------
        self.context_menu1 = self.create_context_menu(self.entry_aptcontract_cefffrom)
        self.context_menu2 = self.create_context_menu(self.entry_aptcontract_ceffthru)
        self.context_menu3 = self.create_context_menu(self.entry_aptcontractchange_cefffrom)
        self.context_menu4 = self.create_context_menu(self.entry_aptcontractchange_ceffthru)

        self.entry_aptcontract_cefffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontract_cefffrom.customContextMenuRequested.connect(self.show_context_menu1)

        self.entry_aptcontract_ceffthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontract_ceffthru.customContextMenuRequested.connect(self.show_context_menu2)

        self.entry_aptcontractchange_cefffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontractchange_cefffrom.customContextMenuRequested.connect(self.show_context_menu3)

        self.entry_aptcontractchange_ceffthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontractchange_ceffthru.customContextMenuRequested.connect(self.show_context_menu4)
        #-----------------------------------------------------------------------------------------------------------------

        self.entry_stylesheet_as_is()
        self.hide_aptcontract_change_widget()

        # Make log file
        self.make_logfiles("access_AptContract.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_aptcontract, partial(self.copy_cells, self.tv_aptcontract))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_aptcontract, partial(self.paste_cells, self.tv_aptcontract))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_aptcontract, partial(self.handle_return_key, self.tv_aptcontract))

    # Change styles for the selected widgets 1
    def entry_stylesheet_as_is(self):
        self.entry_aptcontractchange_customercode.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.cb_aptcontractchange_customername.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractchange_sjcode.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.cb_aptcontractchange_sjname.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractchange_contractvalue.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractchange_cefffrom.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractchange_ceffthru.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractchange_remark.setStyleSheet('color:black;background:rgb(255,255,255)')

    # Change styles for the selected widgets 2
    def entry_stylesheet_to_be(self):
        self.entry_aptcontractchange_customercode.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.cb_aptcontractchange_customername.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractchange_sjcode.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.cb_aptcontractchange_sjname.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractchange_contractvalue.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractchange_cefffrom.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractchange_ceffthru.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractchange_remark.setStyleSheet('color:white;background:rgb(255,0,0)')

    # Show widgets for the cost change parts 
    def show_aptcontract_change_widget(self):
        ddt, endofdate = self.display_eff_date()
        self.pb_aptcontractchange_update_insert.setVisible(True)
        self.entry_aptcontractchange_customercode.setReadOnly(False)
        self.cb_aptcontractchange_customername.setEnabled(True)
        self.entry_aptcontractchange_sjcode.setReadOnly(False)
        self.cb_aptcontractchange_sjname.setEnabled(True)
        self.entry_aptcontractchange_contractvalue.setReadOnly(False)
        self.entry_aptcontractchange_cefffrom.setReadOnly(False)
        self.entry_aptcontractchange_ceffthru.setReadOnly(False)
        self.entry_aptcontractchange_remark.setReadOnly(False)

        self.entry_aptcontractchange_ceffthru.setText(endofdate)
        self.entry_stylesheet_to_be()

    # Hide widgets for the cost change parts 
    def hide_aptcontract_change_widget(self):
        self.pb_aptcontractchange_update_insert.setVisible(False)
        self.entry_aptcontractchange_customercode.setReadOnly(True)
        self.cb_aptcontractchange_customername.setEnabled(False)
        self.entry_aptcontractchange_sjcode.setReadOnly(True)
        self.cb_aptcontractchange_sjname.setEnabled(True)
        self.entry_aptcontractchange_contractvalue.setReadOnly(True)      
        self.entry_aptcontractchange_cefffrom.setReadOnly(True)
        self.entry_aptcontractchange_ceffthru.setReadOnly(True)
        self.entry_aptcontractchange_remark.setReadOnly(True)
        self.entry_stylesheet_as_is()

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu

    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_aptcontract_cefffrom.mapToGlobal(pos))
    def show_context_menu2(self, pos):
        self.context_menu2.exec_(self.entry_aptcontract_ceffthru.mapToGlobal(pos))
    def show_context_menu3(self, pos):
        self.context_menu3.exec_(self.entry_aptcontractchange_cefffrom.mapToGlobal(pos))
    def show_context_menu4(self, pos):
        self.context_menu4.exec_(self.entry_aptcontractchange_ceffthru.mapToGlobal(pos))

    # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_aptcontract_cefffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptcontract_ceffthru:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptcontractchange_cefffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptcontractchange_ceffthru:
            target_lineedit.setText(date)            
    #-----------------------------------------------------------------------------------------------------------------            

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_aptcontract
        self.process_key_event(event, tv_widget)

    # Display current date only
    def display_eff_date(self):
        now = datetime.now()
        curr_date = now.strftime("%Y/%m/%d")
        ddt = f"{curr_date}"         
        endofdate = "2050/12/31"

        return ddt, endofdate
    
    # Display current date only
    def display_currentdate(self):
        ddt, ddt_1 = disply_date_info()
        self.entry_aptcontract_cefffrom.setText(ddt)
        self.entry_aptcontract_ceffthru.setText(ddt_1)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_aptcontract_aname, "SELECT DISTINCT adesc FROM apt_master ORDER BY adesc")
        self.insert_combobox_initiate(self.cb_aptcontract_apttypename, "SELECT DISTINCT description FROM apt_type")
        self.insert_combobox_initiate(self.cb_aptcontract_customername, "SELECT DISTINCT cdescription FROM apt_customer")
        self.insert_combobox_initiate(self.cb_aptcontract_sjname, "SELECT DISTINCT cdesc FROM apt_cic_master ORDER BY cdesc")

        self.insert_combobox_initiate(self.cb_aptcontractchange_customername, "SELECT DISTINCT cdescription FROM apt_customer")
        self.insert_combobox_initiate(self.cb_aptcontractchange_sjname, "SELECT DISTINCT cdesc FROM apt_cic_master ORDER BY cdesc")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) # using common module
        self.lbl_aptcontract_id.setText("")
        self.entry_aptcontract_acode.setText("")
        self.cb_aptcontract_aname.setCurrentIndex(0) 
        self.cb_aptcontract_apttypename.setCurrentIndex(0)
        self.cb_aptcontract_customername.setCurrentIndex(0)
        self.cb_aptcontract_sjname.setCurrentIndex(0)

        self.cb_aptcontractchange_customername.setCurrentIndex(0)
        self.cb_aptcontractchange_sjname.setCurrentIndex(0)

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_aptcontract_show.clicked.connect(self.make_data)
        self.pb_aptcontract_show_con.clicked.connect(self.make_data_con)
        self.pb_aptcontract_search.clicked.connect(self.search_data)
        self.pb_aptcontract_clear_data.clicked.connect(self.clear_data)
        self.pb_aptcontract_close.clicked.connect(self.close_dialog)
        self.pb_aptcontract_insert.clicked.connect(self.tb_insert)
        self.pb_aptcontract_update.clicked.connect(self.SelectionMessageBox)
        self.pb_aptcontract_delete.clicked.connect(self.tb_delete)
        self.pb_aptcontractchange_update_insert.clicked.connect(self.reflect_contract_change)


    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_aptcontract_aname.activated.connect(self.cb_aptcontract_aname_changed)
        self.cb_aptcontract_apttypename.activated.connect(self.cb_aptcontract_apttypename_changed)        
        self.cb_aptcontract_customername.activated.connect(self.cb_aptcontract_customername_changed)
        self.cb_aptcontract_sjname.activated.connect(self.cb_aptcontract_sjname_changed)

        self.cb_aptcontractchange_customername.activated.connect(self.cb_aptcontractchange_customername_changed)
        self.cb_aptcontractchange_sjname.activated.connect(self.cb_aptcontractchange_sjname_changed)
        self.entry_aptcontractchange_cefffrom.editingFinished.connect(self.aptcontract_change_cfffrom_change)

    # Tab order for sub window
    def set_tab_order(self):
        widgets = [self.pb_aptcontract_show, self.entry_aptcontract_acode, self.cb_aptcontract_aname,
            self.entry_aptcontract_noh, self.entry_aptcontract_apttypecode, self.cb_aptcontract_apttypename,
            self.entry_aptcontract_ciccode, self.entry_aptcontract_customercode, self.cb_aptcontract_customername,
            self.entry_aptcontract_sjcode, self.cb_aptcontract_sjname, self.entry_aptcontract_contractvalue,
            self.entry_aptcontract_cefffrom, self.entry_aptcontract_ceffthru, self.entry_aptcontract_remark,
            self.pb_aptcontract_show_con, self.pb_aptcontract_search, self.pb_aptcontract_clear_data,
            self.pb_aptcontract_close, self.pb_aptcontract_insert, self.pb_aptcontract_update, 
            self.pb_aptcontract_delete]

        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_aptcontract
        self.cursor.execute("Select * From vw_apt_contract WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]    

        sql_query = "Select * From vw_apt_contract order by id"
        column_widths = [80, 100, 250, 100, 100, 100, 100, 120, 120, 120, 120, 120]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # Make table data
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Make table data
    def make_data_con(self):
        query = "Select * From vw_apt_contract Where id IS NOT NULL order by id"
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_aptcontract_input(self):
        kaptcode = str(self.entry_aptcontract_acode.text())
        cuscode = str(self.entry_aptcontract_customercode.text()) if len(self.entry_aptcontract_customercode.text()) > 0 else 'N/A'
        mycode = 0 if (self.entry_aptcontract_sjcode.text() == 'None' or self.entry_aptcontract_sjcode.text() == '')  else int(self.entry_aptcontract_sjcode.text())

        input_val = self.entry_aptcontract_contractvalue.text()
        if input_val.replace(".", "").replace("-", "").isdigit():
            conval = float(input_val)
        else:
            conval = 0
            
        date_format_pattern = r'\d{4}/\d{2}/\d{2}'
        srtdt = str(self.entry_aptcontract_cefffrom.text())
        efffrom = srtdt if re.match(date_format_pattern, srtdt) else "2023/01/01"
        enddt = str(self.entry_aptcontract_ceffthru.text())
        effthru = enddt if re.match(date_format_pattern, srtdt) else "2023/12/31"

        remark = str(self.entry_aptcontract_remark.text())
        
        return kaptcode, cuscode, mycode, conval, efffrom, effthru, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()
        return username, user_id, formatted_datetime

    # insert new product data to MySQL table
    def tb_insert(self):
        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            idx = self.max_row_id("apt_contract")
            username, user_id, formatted_datetime = self.common_values_set()
            kaptcode, cuscode, mycode, conval, efffrom, effthru, remark = self.get_aptcontract_input() 

            if (idx>0 and mycode>0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
                self.cursor.execute('''INSERT INTO apt_contract (id, acode, customercode, mycode, contracted_value, efffrom, effthru, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, kaptcode, cuscode, mycode, conval, efffrom, effthru, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data()
                logging.info(f"User {username} inserted {idx} row to the apt contract table.")
            else:
                self.show_missing_message("입력 이상")
                return
        else:
            self.show_cancel_message("데이터 추가 취소")
            return    

    # update 조건에 따라 분기 할 것
    def SelectionMessageBox(self):

        # If there's no selection
        if len(self.lbl_aptcontract_id.text()) == 0:
            self.show_missing_message_update("입력 확인")
            return

        # In case of row selection
        username, user_id, formatted_datetime = self.common_values_set()
        kaptcode, cuscode, mycode, conval, efffrom, effthru, remark = self.get_aptcontract_input() 
       
        if self.lbl_aptcontract_id.text() == 'None':
            idx = int(self.max_row_id("apt_contract"))
                
            if (idx>0 and mycode>0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
                self.cursor.execute('''INSERT INTO apt_contract (id, acode, customercode, mycode, contracted_value, efffrom, effthru, trx_date, userid, remark) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                        , (idx, kaptcode, cuscode, mycode, conval, efffrom, effthru, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()
                logging.info(f"User {username} updated row number {idx} in the apt contract table.")
                
            else:
                self.show_missing_message("입력 이상")
                return

        else:
            conA = '''계약 내용 중 오류 수정 - 현재 행을 수정, 추가 행을 만들지 않음!'''
            conB = '''계약 내용의 변경 또는 갱신 - 현재 행은 변경 없음, 변경된 내용으로 추가 행을 만듦!'''
            
            conditions = [conA, conB]
            condition, okPressed = QInputDialog.getItem(self, "입력 조건 선택", "어떤 작업을 진행하시겠습니까?:", conditions, 0, False)

            if okPressed:
                if condition == conA:
                    self.fix_typo()     
                elif condition == conB:
                    self.show_aptcontract_change_widget()
                else:
                    return

    # Fix typing error in the contract with customer
    def fix_typo(self):
        
        confirm_dialog = self.show_update_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            
            username, user_id, formatted_datetime = self.common_values_set()
            kaptcode, cuscode, mycode, conval, efffrom, effthru, remark = self.get_aptcontract_input() 
            idx = int(self.lbl_aptcontract_id.text())
            
            if (idx>0 and mycode>0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
                self.cursor.execute('''UPDATE apt_contract SET 
                        acode=?, customercode=?, mycode=?, contracted_value=?, efffrom=?, effthru=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                        , (kaptcode, cuscode, mycode, conval, efffrom, effthru, formatted_datetime, user_id, remark, idx))
            else:
                self.show_missing_message("입력 이상")
                return

            self.conn.commit()
            self.show_update_success_message()
            self.refresh_data()
            logging.info(f"User {username} updated row number {idx} in the apt contract table.")
        else:
            self.show_cancel_message("데이터 변경 취소")
            return

    def execute_contract_change(self):
        cuscode = str(self.entry_aptcontractchange_customercode.text())
        mycode = int(self.entry_aptcontractchange_sjcode.text())
        conval = float(self.entry_aptcontractchange_contractvalue.text())
        efffrom = str(self.entry_aptcontractchange_cefffrom.text())
        effthru = str(self.entry_aptcontractchange_ceffthru.text())
        remark = str(self.entry_aptcontractchange_remark.text())

        return cuscode, mycode, conval, efffrom, effthru, remark

    def reflect_contract_change(self):
        kaptcode = str(self.entry_aptcontract_acode.text())
        username, user_id, formatted_datetime = self.common_values_set()
        idx = int(self.max_row_id("apt_contract"))        
        cuscode, mycode, conval, efffrom, effthru, remark  = self.execute_contract_change()
        
        if (idx>0 and mycode>0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
            # 기존 id의 유효종료일을 변경유효시작일 -1 일로 수정
            org_id = str(self.lbl_aptcontract_id.text())
            effthru1 = str(self.entry_aptcontract_ceffthru.text()) # 다시 불러와야 함..
            self.cursor.execute('''UPDATE apt_contract SET 
                    acode=?, customercode=?, mycode=?, contracted_value=?, efffrom=?, effthru=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                    , (kaptcode, cuscode, mycode, conval, efffrom, effthru1, formatted_datetime, user_id, remark, org_id))
            
            #변경된 내용을 신규로 추가
            self.cursor.execute('''INSERT INTO apt_contract (id, acode, customercode, mycode, contracted_value, efffrom, effthru, trx_date, userid, remark) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                    , (idx, kaptcode, cuscode, mycode, conval, efffrom, effthru, formatted_datetime, user_id, remark))
            
            self.conn.commit()
            self.show_update_success_message()
            self.refresh_data()
            logging.info(f"User {username} updated row number {idx} in the apt contract table.")
            
        else:
            self.show_missing_message("입력 이상")
            return

    # delete row according to id selected
    def tb_delete(self):
        confirm_dialog = self.show_delete_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = self.lbl_aptcontract_id.text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM apt_contract WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()
            logging.info(f"User {username} deleted {idx} row to the apt contract table.")  
        else:
            self.show_cancel_message("데이터 삭제 취소")
            return

    # Search data
    def search_data(self):
        aptname = self.cb_aptcontract_aname.currentText()
        cusname = self.cb_aptcontract_customername.currentText()
        sjname= self.cb_aptcontract_sjname.currentText()

        conditions = {'v01': (aptname, "adesc like '%{}%'"), 'v02': (cusname, "cdescription='{}'"), 'v03': (sjname, "cdesc='{}'"),}
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_apt_contract WHERE {' AND '.join(selected_conditions)} ORDER BY adesc"

        QMessageBox.about(self, "검색 조건 확인", f"아파트명: {aptname} \n거래처이름: {cusname} \n수제담당회사명: {sjname} \n\n위 조건으로 검색을 수행합니다!")

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    def cb_aptcontract_aname_changed(self):
        self.entry_aptcontract_acode.clear()
        selected_item = self.cb_aptcontract_aname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode From apt_master WHERE adesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontract_acode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox apt type index changed
    def cb_aptcontract_apttypename_changed(self):
        self.entry_aptcontract_apttypecode.clear()
        selected_item = self.cb_aptcontract_apttypename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT code From apt_type WHERE description ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontract_apttypecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox customer index changed
    def cb_aptcontract_customername_changed(self):
        self.entry_aptcontract_customercode.clear()
        selected_item = self.cb_aptcontract_customername.currentText()

        if selected_item:
            query = f"SELECT DISTINCT taxid From apt_customer WHERE cdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontract_customercode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox suje index changed
    def cb_aptcontract_sjname_changed(self):
        self.entry_aptcontract_sjcode.clear()
        selected_item = self.cb_aptcontract_sjname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT code From apt_cic_master WHERE cdesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontract_sjcode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox customer index changed
    def cb_aptcontractchange_customername_changed(self):
        self.entry_aptcontractchange_customercode.clear()
        selected_item = self.cb_aptcontractchange_customername.currentText()

        if selected_item:
            query = f"SELECT DISTINCT taxid From apt_customer WHERE cdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontractchange_customercode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox suje index changed
    def cb_aptcontractchange_sjname_changed(self):
        self.entry_aptcontractchange_sjcode.clear()
        selected_item = self.cb_aptcontractchange_sjname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT code From apt_cic_master WHERE cdesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontractchange_sjcode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # while cost change efffrom date, change cost effthru date
    def aptcontract_change_cfffrom_change(self):
        chg_date_str = self.entry_aptcontract_change_cefffrom.text()

        try:
            chg_date = parse_date(chg_date_str)                         # 날짜 문자열을 날짜 객체로 변환
            org_date = chg_date - timedelta(days=1)                     # 날짜에서 1일을 빼줍니다.
            org_date_str = org_date.strftime('%Y-%m-%d')                # 결과를 문자열로 변환
            self.entry_aptcontract_ceffthru.setText(org_date_str)       # 변경된 effthru 날짜를 표시
        
        except ValueError as e:
            QMessageBox.critical(self, "날짜 형식 오류", str(e)) # 날짜 형식이 잘못된 경우 사용자에게 알림

    # clear input field entry
    def clear_data(self):
        self.lbl_aptcontract_id.setText("")
        clear_widget_data(self)
        
        self.display_currentdate()
        self.entry_aptcontract_contractvalue.setText("0")

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(15):  # 15columns
            cell_text = self.tv_aptcontract.model().item(row_index, column_index).text()
            cell_values.append(cell_text)
            
        # Populate the input widgets with the data from the selected row
        self.lbl_aptcontract_id.setText(cell_values[0])
        self.entry_aptcontract_acode.setText(cell_values[1])
        self.cb_aptcontract_aname.setCurrentText(cell_values[2])
        self.entry_aptcontract_noh.setText(cell_values[3])
        self.entry_aptcontract_apttypecode.setText(cell_values[4])
        self.cb_aptcontract_apttypename.setCurrentText(cell_values[5])
        self.entry_aptcontract_ciccode.setText(cell_values[6])
        self.entry_aptcontract_customercode.setText(cell_values[7])
        #self.entry_aptcontractchange_customercode.setText(cell_values[7]) # added
        self.cb_aptcontract_customername.setCurrentText(cell_values[8])
        #self.cb_aptcontractchange_customername.setCurrentText(cell_values[8]) # added
        self.entry_aptcontract_sjcode.setText(cell_values[9])
        #self.entry_aptcontractchange_sjcode.setText(cell_values[9]) # added
        self.cb_aptcontract_sjname.setCurrentText(cell_values[10])
        #self.cb_aptcontractchange_sjname.setCurrentText(cell_values[10]) # added
        self.entry_aptcontract_contractvalue.setText(cell_values[11])
        #self.entry_aptcontractchange_contractvalue.setText(cell_values[11]) # added

        ddt, ddt_1 = disply_date_info()
        if cell_values[12] == 'None':
            self.entry_aptcontract_cefffrom.setText(ddt)
            #self.entry_aptcontractchange_cefffrom.setText(ddt)
        else:
            self.entry_aptcontract_cefffrom.setText(cell_values[12])
            #self.entry_aptcontractchange_cefffrom.setText(cell_values[12])

        if cell_values[13] == 'None':
            self.entry_aptcontract_ceffthru.setText(ddt_1)
            #self.entry_aptcontractchange_ceffthru.setText(ddt_1)
        else:
            self.entry_aptcontract_ceffthru.setText(cell_values[13])
            #self.entry_aptcontractchange_ceffthru.setText(cell_values[13])
        
        self.entry_aptcontract_remark.setText(cell_values[14])
        #self.entry_aptcontractchange_remark.setText(cell_values[14])

    def refresh_data(self):
        self.clear_data()
        self.make_data()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_AptContract.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = AptContractDialog()
    dialog.show()
    sys.exit(app.exec())