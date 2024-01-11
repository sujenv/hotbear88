import sys
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
#from aptcontractpic_ui import UI_AptContractPicDialog

#  Dialog and Import common modules -----------------------------------------------------
class AptContractPicDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class AptContractPicDialog(QDialog, UI_AptContractPicDialog, SubWindowBase):

    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database4()

        # Load ui file
        uic.loadUi("aptcontractpic.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_aptcontractpicpic and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "", "", "", "", "", "numeric", "", "", "", "", "numeric", "", "", "",]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_aptcontractpic)
        self.tv_aptcontractpic.setItemDelegate(delegate)
        self.tv_aptcontractpic.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_aptcontractpic.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_aptcontractpic.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_aptcontractpic.verticalHeader().setVisible(False)

        # While selecting row in tv_aptcontractpic, each cell values to displayed to designated widgets
        self.tv_aptcontractpic.clicked.connect(self.show_selected_data)

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
        self.context_menu1 = self.create_context_menu(self.entry_aptcontractpic_eefffrom)
        self.context_menu2 = self.create_context_menu(self.entry_aptcontractpic_eeffthru)
        self.context_menu3 = self.create_context_menu(self.entry_aptcontractpic_change_eefffrom)
        self.context_menu4 = self.create_context_menu(self.entry_aptcontractpic_change_eeffthru)

        self.entry_aptcontractpic_eefffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontractpic_eefffrom.customContextMenuRequested.connect(self.show_context_menu1)

        self.entry_aptcontractpic_eeffthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontractpic_eeffthru.customContextMenuRequested.connect(self.show_context_menu2)

        self.entry_aptcontractpic_change_eefffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontractpic_change_eefffrom.customContextMenuRequested.connect(self.show_context_menu3)

        self.entry_aptcontractpic_change_eeffthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_aptcontractpic_change_eeffthru.customContextMenuRequested.connect(self.show_context_menu4)

        self.entry_stylesheet_as_is()
        self.hide_aptcontractpic_change_widget()

        # Make log file
        self.make_logfiles("access_AptContractPic.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_aptcontractpic, partial(self.copy_cells, self.tv_aptcontractpic))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_aptcontractpic, partial(self.paste_cells, self.tv_aptcontractpic))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_aptcontractpic, partial(self.handle_return_key, self.tv_aptcontractpic))

    # Change styles for the selected widgets 1
    def entry_stylesheet_as_is(self):
        self.entry_aptcontractpic_change_ecode.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.cb_aptcontractpic_change_ename.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractpic_change_payment.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractpic_change_eefffrom.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractpic_change_eeffthru.setStyleSheet('color:black;background:rgb(255,255,255)')
        self.entry_aptcontractpic_change_remark.setStyleSheet('color:black;background:rgb(255,255,255)')

    # Change styles for the selected widgets 2
    def entry_stylesheet_to_be(self):
        self.entry_aptcontractpic_change_ecode.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.cb_aptcontractpic_change_ename.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractpic_change_payment.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractpic_change_eefffrom.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractpic_change_eeffthru.setStyleSheet('color:white;background:rgb(255,0,0)')
        self.entry_aptcontractpic_change_remark.setStyleSheet('color:white;background:rgb(255,0,0)')

    # Show widgets for the cost change parts 
    def show_aptcontractpic_change_widget(self):
        ddt, endofdate = self.display_eff_date()
        self.pb_aptcontractpic_change_update_insert.setVisible(True)
        self.entry_aptcontractpic_change_ecode.setReadOnly(False)
        self.cb_aptcontractpic_change_ename.setEnabled(True)
        self.entry_aptcontractpic_change_payment.setReadOnly(False)
        self.entry_aptcontractpic_change_eefffrom.setReadOnly(False)
        self.entry_aptcontractpic_change_eeffthru.setReadOnly(False)
        self.entry_aptcontractpic_change_remark.setReadOnly(False)

        self.entry_aptcontractpic_change_eeffthru.setText(endofdate)
        self.entry_stylesheet_to_be()

    # Hide widgets for the cost change parts 
    def hide_aptcontractpic_change_widget(self):
        self.pb_aptcontractpic_change_update_insert.setVisible(False)
        self.entry_aptcontractpic_change_ecode.setReadOnly(True)
        self.cb_aptcontractpic_change_ename.setEnabled(False)
        self.entry_aptcontractpic_change_payment.setReadOnly(True)
        self.entry_aptcontractpic_change_eefffrom.setReadOnly(True)      
        self.entry_aptcontractpic_change_eeffthru.setReadOnly(True)
        self.entry_aptcontractpic_change_remark.setReadOnly(True)
        self.entry_stylesheet_as_is()

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu

    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_aptcontractpic_eefffrom.mapToGlobal(pos))
    def show_context_menu2(self, pos):
        self.context_menu2.exec_(self.entry_aptcontractpic_eeffthru.mapToGlobal(pos))    
    def show_context_menu3(self, pos):
        self.context_menu3.exec_(self.entry_aptcontractpic_change_eefffrom.mapToGlobal(pos))
    def show_context_menu4(self, pos):
        self.context_menu4.exec_(self.entry_aptcontractpic_change_eeffthru.mapToGlobal(pos))

    # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_aptcontractpic_eefffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptcontractpic_eeffthru:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptcontractpic_change_eefffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_aptcontractpic_change_eeffthru:
            target_lineedit.setText(date)               

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_aptcontractpic
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
        self.entry_aptcontractpic_eefffrom.setText(ddt)
        self.entry_aptcontractpic_eeffthru.setText(ddt_1)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_aptcontractpic_aname, "SELECT DISTINCT adesc FROM apt_master ORDER BY adesc")
        self.insert_combobox_initiate(self.cb_aptcontractpic_customername, "SELECT DISTINCT cdescription FROM apt_customer")
        self.insert_combobox_initiate(self.cb_aptcontractpic_ename, "SELECT ename FROM employee WHERE class1 <> 'r' ORDER BY ename")
        self.insert_combobox_initiate(self.cb_aptcontractpic_change_ename, "SELECT ename FROM employee WHERE class1 <> 'r' ORDER BY ename")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) # using common module
        self.lbl_aptcontractpic_id.setText("")
        self.entry_aptcontractpic_acode.setText("")
        self.cb_aptcontractpic_aname.setCurrentIndex(0) 
        self.cb_aptcontractpic_customername.setCurrentIndex(0)
        self.cb_aptcontractpic_ename.setCurrentIndex(0)
        self.cb_aptcontractpic_change_ename.setCurrentIndex(0)

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_aptcontractpic_show.clicked.connect(self.make_data)
        self.pb_aptcontractpic_show_con.clicked.connect(self.make_data_con)
        self.pb_aptcontractpic_search.clicked.connect(self.search_data)
        self.pb_aptcontractpic_clear_data.clicked.connect(self.clear_data)
        self.pb_aptcontractpic_close.clicked.connect(self.close_dialog)

        self.pb_aptcontractpic_insert.clicked.connect(self.tb_insert)
        self.pb_aptcontractpic_update.clicked.connect(self.SelectionMessageBox)
        self.pb_aptcontractpic_delete.clicked.connect(self.tb_delete)
        self.pb_aptcontractpic_change_update_insert.clicked.connect(self.reflect_contract_change)

    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_aptcontractpic_aname.activated.connect(self.cb_aptcontractpic_aname_changed)        
        self.cb_aptcontractpic_customername.activated.connect(self.cb_aptcontractpic_customername_changed)
        self.cb_aptcontractpic_ename.activated.connect(self.cb_aptcontractpic_ename_changed)
        self.cb_aptcontractpic_change_ename.activated.connect(self.cb_aptcontractpic_change_ename_changed)
        self.entry_aptcontractpic_change_eefffrom.editingFinished.connect(self.aptcontractpic_change_efffrom_change)
    
    # Tab order for sub window
    def set_tab_order(self):
        widgets = [self.pb_aptcontractpic_show, self.entry_aptcontractpic_acode, self.cb_aptcontractpic_aname,
            self.entry_aptcontractpic_noh, self.entry_aptcontractpic_customercode, self.cb_aptcontractpic_customername,
            self.entry_aptcontractpic_contractvalue, self.entry_aptcontractpic_cefffrom, self.entry_aptcontractpic_ceffthru, 
            self.entry_aptcontractpic_ecode, self.cb_aptcontractpic_ename, self.entry_aptcontractpic_payment, 
            self.entry_aptcontractpic_eefffrom, self.entry_aptcontractpic_eeffthru, self.entry_aptcontractpic_remark, 
            self.pb_aptcontractpic_show_con, self.pb_aptcontractpic_search, self.pb_aptcontractpic_clear_data, 
            self.pb_aptcontractpic_close, self.pb_aptcontractpic_insert, self.pb_aptcontractpic_update, 
            self.pb_aptcontractpic_delete]

        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_aptcontractpic
        self.cursor.execute("Select * From vw_apt_contract_pic WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]    

        sql_query = "Select * From vw_apt_contract_pic order by id"
        column_widths = [80, 100, 250, 100, 100, 100, 100, 120, 120, 120, 120, 120]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # Make table data
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Make table data
    def make_data_con(self):
        query = "Select * From vw_apt_contract_pic Where id IS NOT NULL order by id"

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_aptcontractpic_input(self):
        kaptcode = str(self.entry_aptcontractpic_acode.text())
        ecode = int(self.entry_aptcontractpic_ecode.text())
        efffrom = str(self.entry_aptcontractpic_eefffrom.text())
        effthru = str(self.entry_aptcontractpic_eeffthru.text())
        #payment = int(self.entry_aptcontractpic_payment.text())
        #payment = float(self.entry_aptcontractpic_payment.text()) if self.entry_aptcontractpic_payment.text().replace(".", "").isdigit() else 0
        input_val = self.entry_aptcontractpic_payment.text()
        if input_val.replace(".", "").replace("-", "").isdigit():
            payment = float(input_val)
        else:
            payment = 0        
        
        remark = str(self.entry_aptcontractpic_remark.text())

        return kaptcode, ecode, efffrom, effthru, payment, remark

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
            idx = self.max_row_id("apt_contract_pic")
            username, user_id, formatted_datetime = self.common_values_set()
            kaptcode, ecode, efffrom, effthru, payment, remark = self.get_aptcontractpic_input() 

            if (idx>0 and ecode>0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
                self.cursor.execute('''INSERT INTO apt_contract_pic (id, acode, ecode, efffrom, effthru, payment, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, kaptcode, ecode, efffrom, effthru, payment, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data() 
                logging.info(f"User {username} inserted {idx} row to the apt contract pic table.")
            else:
                self.show_missing_message("입력 이상")
                return
        else:
            self.show_cancel_message("데이터 추가 취소")
            return    

    # update 조건에 따라 분기 할 것
    def SelectionMessageBox(self):

        # If there's no selection
        if len(self.lbl_aptcontractpic_id.text()) == 0:
            self.show_missing_message_update("입력 확인")
            return

        # In case of row selection
        username, user_id, formatted_datetime = self.common_values_set()
        kaptcode, ecode, efffrom, effthru, payment, remark = self.get_aptcontractpic_input() 

        if self.lbl_aptcontractpic_id.text() == 'None':
            idx = int(self.max_row_id("apt_contract_pic"))
                
            if (idx>0 and ecode>0 and abs(payment)>0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
                self.cursor.execute('''INSERT INTO apt_contract_pic (id, acode, ecode, efffrom, effthru, payment, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, kaptcode, ecode, efffrom, effthru, payment, formatted_datetime, user_id, remark))  
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()
                logging.info(f"User {username} updated row number {idx} in the apt contract pic table.")
                
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
                    self.show_aptcontractpic_change_widget()
                else:
                    return

    # Fix typing error in the contract with customer
    def fix_typo(self):
        
        confirm_dialog = self.show_update_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            
            username, user_id, formatted_datetime = self.common_values_set()
            kaptcode, ecode, efffrom, effthru, payment, remark = self.get_aptcontractpic_input()  
            idx = int(self.lbl_aptcontractpic_id.text())
            
            if (idx>0 and ecode>0 and abs(payment)>=0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
                self.cursor.execute('''UPDATE apt_contract_pic SET 
                        acode=?, ecode=?, efffrom=?, effthru=?, payment=?, trx_date=?, userid=?, remark=? WHERE id=?) '''
                        , (kaptcode, ecode, efffrom, effthru, payment, formatted_datetime, user_id, remark, idx))
            else:
                self.show_missing_message("입력 이상")
                return

            self.conn.commit()
            self.show_update_success_message()
            self.refresh_data()
            logging.info(f"User {username} updated row number {idx} in the apt contract pic table.")
        else:
            self.show_cancel_message("데이터 변경 취소")
            return
        
    def execute_contract_change(self):
        ecode = str(self.entry_aptcontractpic_change_ecode.text())
        input_text = self.entry_aptcontractpic_change_payment.text()

        if input_text is not None and input_text.strip() != "":
            payment = float(input_text)
        else:
            # Set a default value for payment when the input is None or empty
            payment = 0.0  # You can replace 0.0 with any default value you prefer
        #payment = float(self.entry_aptcontractpic_change_payment.text())
        
        efffrom = str(self.entry_aptcontractpic_change_eefffrom.text())
        effthru = str(self.entry_aptcontractpic_change_eeffthru.text())
        remark = str(self.entry_aptcontractpic_change_remark.text())

        return ecode, payment, efffrom, effthru, remark

    def reflect_contract_change(self):
        idx = int(self.max_row_id("apt_contract_pic"))
        kaptcode = str(self.entry_aptcontractpic_acode.text())
        username, user_id, formatted_datetime = self.common_values_set()
        ecode, payment, efffrom, effthru, remark = self.execute_contract_change()
        
        if (idx>0 and ecode>0 and abs(payment)>=0) and all(len(var) > 0 for var in (kaptcode, efffrom, effthru)):
            # 기존 id의 유효종료일을 변경유효시작일 -1 일로 수정
            org_id = str(self.lbl_aptcontractpic_id.text())
            effthru1 = str(self.entry_aptcontractpic_eeffthru.text()) # 다시 불러와야 함..
            self.cursor.execute('''UPDATE apt_contract_pic SET 
                    acode=?, ecode=?, efffrom=?, effthru=?, payment=?, trx_date=?, userid=?, remark=? WHERE id=?) '''
                    , (kaptcode, ecode, efffrom, effthru1, payment, formatted_datetime, user_id, remark, org_id)) 
            
            #변경된 내용을 신규로 추가
            self.cursor.execute('''INSERT INTO apt_contract_pic (id, acode, ecode, efffrom, effthru, payment, trx_date, userid, remark) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                        , (idx, kaptcode, ecode, efffrom, effthru, payment, formatted_datetime, user_id, remark))  
        
            self.conn.commit()
            self.show_update_success_message()
            self.refresh_data()
            logging.info(f"User {username} updated row number {idx} in the apt contract pic table.")
            
        else:
            self.show_missing_message("입력 이상")
            return

    # delete row according to id selected
    def tb_delete(self):
        confirm_dialog = self.show_delete_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = self.lbl_aptcontractpic_id.text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM apt_contract_pic WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()
            logging.info(f"User {username} deleted {idx} row to the apt contract pic table.")
        else:
            self.show_cancel_message("데이터 삭제 취소")
            return

    # Search data
    def search_data(self):
        aptname = self.cb_aptcontractpic_aname.currentText()
        cusname = self.cb_aptcontractpic_customername.currentText()
        pic = self.cb_aptcontractpic_ename.currentText()

        conditions = {'v01': (aptname, "adesc like '%{}%'"), 'v02': (cusname, "cdescription='{}'"), 'v03': (pic, "ename='{}'"),}

        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_apt_contract_pic WHERE {' AND '.join(selected_conditions)} ORDER BY adesc"

        QMessageBox.about(self, "검색 조건 확인", f"아파트명: {aptname} \n거래처이름: {cusname} \n 담당자명: {pic} \n\n위 조건으로 검색을 수행합니다!")

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Combobox apt name index changed
    def cb_aptcontractpic_aname_changed(self):
        self.entry_aptcontractpic_acode.clear()
        selected_item = self.cb_aptcontractpic_aname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode From apt_master WHERE adesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontractpic_acode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox customer index changed
    def cb_aptcontractpic_customername_changed(self):
        self.entry_aptcontractpic_customercode.clear()
        selected_item = self.cb_aptcontractpic_customername.currentText()

        if selected_item:
            query = f"SELECT DISTINCT taxid From apt_customer WHERE cdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontractpic_customercode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox suje index changed
    def cb_aptcontractpic_ename_changed(self):
        self.entry_aptcontractpic_ecode.clear()
        selected_item = self.cb_aptcontractpic_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontractpic_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox suje index changed
    def cb_aptcontractpic_change_ename_changed(self):
        self.entry_aptcontractpic_change_ecode.clear()
        selected_item = self.cb_aptcontractpic_change_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_aptcontractpic_change_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # while cost change efffrom date, change cost effthru date
    def aptcontractpic_change_efffrom_change(self):
        chg_date_str = self.entry_aptcontractpic_change_eefffrom.text()

        try:
            chg_date = parse_date(chg_date_str)                         # 날짜 문자열을 날짜 객체로 변환
            org_date = chg_date - timedelta(days=1)                     # 날짜에서 1일을 빼줍니다.
            org_date_str = org_date.strftime('%Y-%m-%d')                # 결과를 문자열로 변환
            self.entry_aptcontractpic_eeffthru.setText(org_date_str)    # 변경된 cost effthru 날짜를 표시
        
        except ValueError as e:
            QMessageBox.critical(self, "날짜 형식 오류", str(e)) # 날짜 형식이 잘못된 경우 사용자에게 알림


    # clear input field entry
    def clear_data(self):
        self.lbl_aptcontractpic_id.setText("")
        clear_widget_data(self)

        self.display_currentdate()                
        self.entry_aptcontractpic_contractvalue.setText("0")
        self.entry_aptcontractpic_payment.setText("0")

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(15):  # 15columns
            cell_text = self.tv_aptcontractpic.model().item(row_index, column_index).text()
            cell_values.append(cell_text)
            
        # Populate the input widgets with the data from the selected row
        self.lbl_aptcontractpic_id.setText(cell_values[0])
        self.entry_aptcontractpic_acode.setText(cell_values[1])
        self.cb_aptcontractpic_aname.setCurrentText(cell_values[2])
        self.entry_aptcontractpic_noh.setText(cell_values[3])
        self.entry_aptcontractpic_customercode.setText(cell_values[4])
        self.cb_aptcontractpic_customername.setCurrentText(cell_values[5])
        self.entry_aptcontractpic_contractvalue.setText(cell_values[6])
        self.entry_aptcontractpic_cefffrom.setText(cell_values[7])
        self.entry_aptcontractpic_ceffthru.setText(cell_values[8])
        self.entry_aptcontractpic_ecode.setText(cell_values[9])
        self.cb_aptcontractpic_ename.setCurrentText(cell_values[10])
        self.entry_aptcontractpic_payment.setText(cell_values[11])
        self.entry_aptcontractpic_eefffrom.setText(cell_values[12])
        self.entry_aptcontractpic_eeffthru.setText(cell_values[13])
        self.entry_aptcontractpic_remark.setText(cell_values[14])

        #print(cell_values[10], cell_values[11])

    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_AptContractPic.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = AptContractPicDialog()
    dialog.show()
    sys.exit(app.exec())