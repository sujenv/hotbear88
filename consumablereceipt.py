import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QMenu, QShortcut
from PyQt5.QtCore import Qt
from datetime import datetime
from commonmd import *
#for non_ui version-------------------------
#from consumablereceipt_ui import Ui_ConsumableReceiptDialog

# Consumable Item Receipt -----------------------------------------------------
class ConsumableReceiptDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class ConsumableReceiptDialog(QDialog, Ui_ConsumableReceiptDialog, SubWindowBase): 
    
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database2()

        uic.loadUi("consumablereceipt.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_consumablereceipt and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "numeric", "", "numeric", "", "", "numeric", "numeric", "numeric", "", "", "", "", ]
        
        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_consumablereceipt)
        self.tv_consumablereceipt.setItemDelegate(delegate)
        self.tv_consumablereceipt.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_consumablereceipt.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_consumablereceipt.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_consumablereceipt.verticalHeader().setVisible(False)

        # While selecting row in tv_consumablereceipt, each cell values to displayed to designated widgets
        self.tv_consumablereceipt.clicked.connect(self.show_selected_data)

        # Fill combobox items when the application starts
        self.get_combobox_contents()
        
        # Initial Display of data
        self.make_data()
        self.conn_button_to_method()
        self.conn_signal_to_slot()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Display Current Date
        self.display_currentdate()

        # Make log file
        self.make_logfiles("access_ConsumableReceipt.log")

    # Display current date only
    def display_currentdate(self):
        now = datetime.now()
        curr_date = now.strftime("%Y/%m/%d")
        ddt = f"{curr_date}"         
        
        self.entry_consumablereceipt_trxdt.setText(ddt)

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_consumablereceipt, partial(self.copy_cells, self.tv_consumablereceipt))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_consumablereceipt, partial(self.paste_cells, self.tv_consumablereceipt))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_consumablereceipt, partial(self.handle_return_key, self.tv_consumablereceipt))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_consumablereceipt
        self.process_key_event(event, tv_widget)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_consumablereceipt_cname, "SELECT DISTINCT cname FROM customer where type01='con'")
        self.insert_combobox_initiate(self.cb_consumablereceipt_ename, "SELECT DISTINCT ename FROM employee where class1='r' ORDER BY ename")
        self.insert_combobox_initiate(self.cb_consumablereceipt_pname, "SELECT DISTINCT pdescription FROM product ORDER BY pdescription")
        self.insert_combobox_initiate(self.cb_consumablereceipt_status_descr, "SELECT DISTINCT sdescription FROM status where id = 1 ORDER BY sdescription")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.entry_consumablereceipt_unitprice.setText("0") # 단가 초기화
        self.entry_consumablereceipt_status.setText("1")
        self.cb_consumablereceipt_status_descr.setCurrentIndex(1) # 첫번째 item 보여주기, 나머지는 blank

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_consumablereceipt_show.clicked.connect(self.make_data)
        self.pb_consumablereceipt_cancel.clicked.connect(self.close_dialog)
        self.pb_consumablereceipt_clear.clicked.connect(self.clear_data)
        self.pb_consumablereceipt_search.clicked.connect(self.search_data)

        self.pb_consumablereceipt_insert.clicked.connect(self.tb_insert)
        self.pb_consumablereceipt_update.clicked.connect(self.tb_update)
        self.pb_consumablereceipt_delete.clicked.connect(self.tb_delete)

    # Connect signal to method    
    def conn_signal_to_slot(self):
        self.cb_consumablereceipt_cname.activated.connect(self.customer_description_changed)
        self.cb_consumablereceipt_ename.activated.connect(self.employee_name_changed)
        self.cb_consumablereceipt_pname.activated.connect(self.product_name_changed)
        self.entry_consumablereceipt_qty.textEdited.connect(self.qty_changed)
        self.cb_consumablereceipt_status_descr.activated.connect(self.cb_consumablereceipt_status_descr_changed)

    # tab order for receiptproduct window
    def set_tab_order(self):
        widgets = [self.pb_consumablereceipt_show, self.entry_consumablereceipt_ccode, self.cb_consumablereceipt_cname,
            self.entry_consumablereceipt_ecode, self.cb_consumablereceipt_ename, self.entry_consumablereceipt_pcode,
            self.cb_consumablereceipt_pname, self.entry_consumablereceipt_um, self.entry_consumablereceipt_qty,
            self.entry_consumablereceipt_unitprice, self.entry_consumablereceipt_amount, self.entry_consumablereceipt_trxdt,
            self.entry_consumablereceipt_status, self.cb_consumablereceipt_status_descr, self.entry_consumablereceipt_remark,
            self.pb_consumablereceipt_clear, self.pb_consumablereceipt_search, self.pb_consumablereceipt_insert,
            self.pb_consumablereceipt_update, self.pb_consumablereceipt_delete, self.pb_consumablereceipt_cancel]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_consumablereceipt
        self.cursor.execute("SELECT * FROM vw_aplist_02 where 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_aplist_02 where status <> '9' order by id desc"
        column_widths = [80, 100, 100, 100, 100, 100]

        return sql_query, tv_widget, column_info, column_names, column_widths 
    
    # show receiptproduct table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_aplist_input(self):
        ccode = int(self.entry_consumablereceipt_ccode.text())
        ecode = int(self.entry_consumablereceipt_ecode.text())
        pcode = int(self.entry_consumablereceipt_pcode.text())
        qty = int(self.entry_consumablereceipt_qty.text())
        trx_date = str(self.entry_consumablereceipt_trxdt.text())
        status = str(self.entry_consumablereceipt_status.text())
        remark = str(self.entry_consumablereceipt_remark.text())

        return ccode, ecode, pcode, qty, trx_date, status, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime

    # insert new receiptproduct data to MySQL table
    def tb_insert(self):
        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:            
            idx = self.max_row_id("aplist")
            username, user_id, formatted_datetime = self.common_values_set()
            ccode, ecode, pcode, qty, trx_date, status,  remark = self.get_aplist_input()               # Get the value of other variables
            
            if (idx>0 and ccode>0 and ecode>0 and pcode>0 and qty>0) and all(len(var) > 0 for var in (trx_date, status)):            
                self.cursor.execute('''INSERT INTO aplist (id, ccode, ecode, pcode, qty, trx_date, status, userid, up_date, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, ccode, ecode, pcode, qty, trx_date, status, user_id, formatted_datetime, remark))
                self.conn.commit()

                # 생수18L의 경우 자체 소진으로 triggering--------------------------------------------------------------
                if pcode == 1005:
                    self.trigger_issue_product()        

                self.show_insert_success_message()
                self.refresh_data() 
                logging.info(f"User {username} inserted {idx} row to the APlist table.")        
            else:
                self.show_missing_message("입력 이상")
                return
        else:
            self.show_cancel_message("데이터 추가 취소")
            return

    # In case of the product code no 1005, receipt => issue
    def trigger_issue_product(self):

        self.cursor.execute("SELECT MAX(id) FROM arlist")
        row = self.cursor.fetchone()
        max_id = row[0]
        
        id = max_id + 1
        ccode = 80000002    # 수제산업
        ecode = 20301003    # 자체소진
        pcode = 1005        # 생수18L
        qty = int(self.entry_consumablereceipt_qty.text()) 
        trx_date = str(self.entry_consumablereceipt_trxdt.text())
        status = '2'
        username = self.current_username   
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()
        remark = str(self.entry_consumablereceipt_remark.text())
        
        self.cursor.execute('''INSERT INTO arlist (id, ccode, ecode, pcode, qty, trx_date, status, userid, up_date, remark) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                    , (id, ccode, ecode, pcode, qty, trx_date, status, user_id, formatted_datetime, remark))
        self.conn.commit()

    # revise the values in the selected row
    def tb_update(self):

        confirm_dialog = self.show_update_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = int(self.lbl_consumablereceipt_id.text())
            username, user_id, formatted_datetime = self.common_values_set()
            ccode, ecode, pcode, qty, trx_date, status, remark = self.get_aplist_input() 

            if (idx>0 and ccode>0 and ecode>0 and pcode>0 and qty>0) and all(len(var) > 0 for var in (trx_date, status)): 
                self.cursor.execute('''UPDATE aplist SET 
                            ccode=?, ecode=?, pcode=?, qty=?, trx_date=?, status=?, userid=?, up_date=?, remark=? WHERE id=?'''
                            , (ccode, ecode, pcode, qty, trx_date, status, user_id, formatted_datetime, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()  
                logging.info(f"User {username} updated row number {idx} in the AP List table.")
            else:
                self.show_missing_message("입력 이상")
                return
        else:
            self.show_cancel_message("데이터 변경 취소")
            return
        
    # delete row according to id selected
    def tb_delete(self):
        confirm_dialog = self.show_delete_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = self.lbl_consumablereceipt_id.text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM aplist WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()  
            logging.info(f"User {username} deleted {idx} row to the AP List table.")                   
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # Search Data
    def search_data(self):
        cname = self.cb_consumablereceipt_cname.currentText()
        pname = self.cb_consumablereceipt_pname.currentText()
        trxdt = self.entry_consumablereceipt_trxdt.text()

        conditions = {'v01': (cname, "cname like '%{}%'"),
                      'v02': (pname, "pdescription like '%{}%'"),
                      'v03': (trxdt, "trx_date like '%{}%'"),}

        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_aplist_02 WHERE {' AND '.join(selected_conditions)} ORDER BY id desc"

        QMessageBox.about(self, "검색 조건 확인", f"거래처명: {cname}\n 제품명: {pname}\n 날짜: {trxdt} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # customer code editing finished and connect
    def customer_description_changed(self):
        self.entry_consumablereceipt_ccode.clear()
        selected_item = self.cb_consumablereceipt_cname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From customer WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablereceipt_ccode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass
    
    # status code editing finished and connect
    def cb_consumablereceipt_status_descr_changed(self):
        self.entry_consumablereceipt_status.clear()
        selected_item = self.cb_consumablereceipt_status_descr.currentText()

        if selected_item:
            query = f"SELECT DISTINCT scode From status WHERE sdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablereceipt_status]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # employee name editing finished and connect
    def employee_name_changed(self):
        self.entry_consumablereceipt_ecode.clear()
        selected_item = self.cb_consumablereceipt_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablereceipt_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # product name editing finished and connect
    def product_name_changed(self):
        self.entry_consumablereceipt_pcode.clear()
        self.entry_consumablereceipt_unitprice.clear()
        self.entry_consumablereceipt_um.clear()

        selected_item = self.cb_consumablereceipt_pname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT pcode, um, unitprice From vw_price_receipt WHERE pdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablereceipt_pcode, self.entry_consumablereceipt_um, self.entry_consumablereceipt_unitprice]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # quantity changed and connect
    def qty_changed(self):
        self.entry_consumablereceipt_amount.clear()
        u_price = int(self.entry_consumablereceipt_unitprice.text())
        #qtyadd = int(self.entry_consumablereceipt_qty.text())
        qty_text = self.entry_consumablereceipt_qty.text()
        qtyadd = 0 if qty_text == '' else int(qty_text)

        if u_price >= 0 and qtyadd > 0:
            
            amount = u_price * qtyadd
            self.entry_consumablereceipt_amount.setText(str(amount))       

    # clear input field entry
    def clear_data(self):
        self.lbl_consumablereceipt_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

        self.cb_consumablereceipt_cname.setCurrentIndex(0)
        self.cb_consumablereceipt_ename.setCurrentIndex(0)
        self.cb_consumablereceipt_pname.setCurrentIndex(0)        
        self.entry_consumablereceipt_qty.setText("0")
        self.entry_consumablereceipt_unitprice.setText("0")
        self.entry_consumablereceipt_amount.setText("0")
        self.entry_consumablereceipt_status.setText("1")
        self.cb_consumablereceipt_status_descr.setCurrentIndex(1)
        self.display_currentdate() # 현재 날짜 기본입력
 

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(15):  # 15 columns
            cell_text = self.tv_consumablereceipt.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_consumablereceipt_id.setText(cell_values[0])
        self.entry_consumablereceipt_ccode.setText(cell_values[1])
        self.cb_consumablereceipt_cname.setCurrentText(cell_values[2])
        self.entry_consumablereceipt_ecode.setText(cell_values[3])
        self.cb_consumablereceipt_ename.setCurrentText(cell_values[4])
        self.entry_consumablereceipt_pcode.setText(cell_values[5])
        self.cb_consumablereceipt_pname.setCurrentText(cell_values[6])
        self.entry_consumablereceipt_um.setText(cell_values[7])
        self.entry_consumablereceipt_qty.setText(cell_values[8])
        self.entry_consumablereceipt_unitprice.setText(cell_values[9])
        self.entry_consumablereceipt_amount.setText(cell_values[10])
        self.entry_consumablereceipt_trxdt.setText(cell_values[11])
        self.entry_consumablereceipt_status.setText(cell_values[12])
        self.cb_consumablereceipt_status_descr.setCurrentText(cell_values[13])
        self.entry_consumablereceipt_remark.setText(cell_values[14])
    
    # Clear all entry and combo boxes
    def refresh_data(self):
        self.clear_data()
        self.make_data()
        
        self.display_currentdate() # 현재 날짜 기본입력
        self.cb_consumablereceipt_status_descr.setCurrentIndex(1) # 입고 기본 입력

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_ConsumableReceipt.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = ConsumableReceiptDialog()
    dialog.show()
    sys.exit(app.exec())