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
#from consumableslaes_ui import Ui_ConsumableSalesDialog

# Consumable Item Sales contents -----------------------------------------------------
class ConsumableSalesDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class ConsumableSalesDialog(QDialog, Ui_ConsumableSalesDialog, SubWindowBase): 
    
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database2()

        uic.loadUi("consumablesales.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_consumablesales and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "numeric", "", "numeric", "", "", "numeric", "numeric", "numeric", "", "", "", "", ]
        
        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_consumablesales)
        self.tv_consumablesales.setItemDelegate(delegate)
        self.tv_consumablesales.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_consumablesales.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_consumablesales.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_consumablesales.verticalHeader().setVisible(False)

        # While selecting row in tv_consumablesales, each cell values to displayed to designated widgets
        self.tv_consumablesales.clicked.connect(self.show_selected_data)

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

        self.display_currentdate()

        # Make log file
        self.make_logfiles("access_ConsumableSales.log")

    # Display current date only
    def display_currentdate(self):
        now = datetime.now()
        curr_date = now.strftime("%Y/%m/%d")
        ddt = f"{curr_date}"         
        
        self.entry_consumablesales_trxdt.setText(ddt)

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_consumablesales, partial(self.copy_cells, self.tv_consumablesales))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_consumablesales, partial(self.paste_cells, self.tv_consumablesales))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_consumablesales, partial(self.handle_return_key, self.tv_consumablesales))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_consumablesales
        self.process_key_event(event, tv_widget)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_consumablesales_cname, "SELECT DISTINCT cname FROM customer where type01='s'")
        self.insert_combobox_initiate(self.cb_consumablesales_ename, "SELECT DISTINCT ename FROM employee where class1='e' ORDER BY ename")
        self.insert_combobox_initiate(self.cb_consumablesales_pname, "SELECT DISTINCT pdescription FROM product ORDER BY pdescription")
        self.insert_combobox_initiate(self.cb_consumablesales_status_descr, "SELECT DISTINCT sdescription FROM status where id = 2")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.entry_consumablesales_unitprice.setText("0") # 단가 초기화
        self.entry_consumablesales_status.setText("2")
        self.entry_consumablesales_amount.setText("0")
        self.cb_consumablesales_status_descr.setCurrentIndex(1) # 얘만 두번째 item 보여주기, 나머지는 blank

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_consumablesales_show.clicked.connect(self.make_data)
        self.pb_consumablesales_cancel.clicked.connect(self.close_dialog)
        self.pb_consumablesales_clear.clicked.connect(self.clear_data)
        self.pb_consumablesales_search.clicked.connect(self.search_data)

        self.pb_consumablesales_insert.clicked.connect(self.tb_insert)
        self.pb_consumablesales_update.clicked.connect(self.tb_update)
        self.pb_consumablesales_delete.clicked.connect(self.tb_delete)

    # Connect signal to method    
    def conn_signal_to_slot(self):
        self.cb_consumablesales_cname.activated.connect(self.customer_description_changed)
        self.cb_consumablesales_ename.activated.connect(self.employee_name_changed)
        self.cb_consumablesales_pname.activated.connect(self.product_name_changed)
        self.entry_consumablesales_qty.textEdited.connect(self.qty_changed)
        self.cb_consumablesales_status_descr.activated.connect(self.cb_consumablesales_status_descr_changed)

    # tab order for issueproduct window
    def set_tab_order(self):
        widgets = [self.pb_consumablesales_show, self.entry_consumablesales_ccode, self.cb_consumablesales_cname,
            self.entry_consumablesales_ecode, self.cb_consumablesales_ename, self.entry_consumablesales_pcode,
            self.cb_consumablesales_pname, self.entry_consumablesales_um, self.entry_consumablesales_qty,
            self.entry_consumablesales_unitprice, self.entry_consumablesales_amount, self.entry_consumablesales_trxdt,
            self.entry_consumablesales_status, self.cb_consumablesales_status_descr, self.entry_consumablesales_remark,
            self.pb_consumablesales_search, self.pb_consumablesales_clear, self.pb_consumablesales_insert, 
            self.pb_consumablesales_update, self.pb_consumablesales_delete, self.pb_consumablesales_cancel]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_consumablesales

        self.cursor.execute("SELECT * FROM vw_arlist_02 WHERE 1=0")
        column_info = self.cursor.description
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_arlist_02 where scode <> '9' order by id desc"
        column_widths = [80, 100, 100, 100, 100, 100]
    
        return sql_query, tv_widget, column_info, column_names, column_widths 
    
    # show issueproduct table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_arlist_input(self):
        ccode = int(self.entry_consumablesales_ccode.text())
        ecode = int(self.entry_consumablesales_ecode.text())
        pcode = int(self.entry_consumablesales_pcode.text())
        qty = int(self.entry_consumablesales_qty.text()) 
        trx_date = str(self.entry_consumablesales_trxdt.text())
        status = str(self.entry_consumablesales_status.text())
        remark = str(self.entry_consumablesales_remark.text())

        return ccode, ecode, pcode, qty, trx_date, status, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime
    
    # insert new issueproduct data to MySQL table
    def tb_insert(self):
        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            
            idx = self.max_row_id("arlist")
            username, user_id, formatted_datetime = self.common_values_set()
            ccode, ecode, pcode, qty, trx_date, status, remark = self.get_arlist_input() 
            
            if (idx>0 and ccode>0 and pcode>0 and qty>0) and all(len(var) > 0 for var in (trx_date, status)):
                self.cursor.execute('''INSERT INTO arlist (id, ccode, ecode, pcode, qty, trx_date, status, userid, up_date, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, ccode, ecode, pcode, qty, trx_date, status, user_id, formatted_datetime, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data() 
                logging.info(f"User {username} inserted {idx} row to the AR List table.")        
            else:
                self.show_missing_message("입력 이상")
                return
        else:
            self.show_cancel_message("데이터 추가 취소")
            return

    # revise the values in the selected row
    def tb_update(self):
        confirm_dialog = self.show_update_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = int(self.lbl_consumablesales_id.text())
            username, user_id, formatted_datetime = self.common_values_set()  
            ccode, ecode, pcode, qty, trx_date, status, remark = self.get_arlist_input() 
            
            if (idx>0 and ccode>0 and pcode>0 and qty>0) and all(len(var) > 0 for var in (trx_date, status)):
                self.cursor.execute('''UPDATE arlist SET 
                            ccode=?, ecode=?, pcode=?, qty=?, trx_date=?, status=?, userid=?, up_date=?, remark=? WHERE id=?'''
                            , (ccode, ecode, pcode, qty, trx_date, status, user_id, formatted_datetime, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()  
                logging.info(f"User {username} updated row number {idx} in the AR List table.")
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
            idx = self.lbl_consumablesales_id.text()
            username, user_id, formatted_datetime = self.common_values_set()

            self.cursor.execute("DELETE FROM arlist WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()  
            logging.info(f"User {username} deleted {idx} row to the AR List table.")                
        else:
            self.show_cancel_message("데이터 삭제 취소")
            
    # Search data
    def search_data(self):
        ename = self.cb_consumablesales_ename.currentText()
        pname = self.cb_consumablesales_pname.currentText()
        trxdt= self.entry_consumablesales_trxdt.text()

        conditions = {'v01': (ename, "ename like '%{}%'"),
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
        query = f"SELECT * FROM vw_arlist_02 WHERE {' AND '.join(selected_conditions)} ORDER BY id desc"

        QMessageBox.about(self, "검색 조건 확인", f"직원명: {ename}\n제품명: {pname}\n날짜: {trxdt} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # customer code editing finished and connect
    def customer_description_changed(self):
        self.entry_consumablesales_ccode.clear()
        selected_item = self.cb_consumablesales_cname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From customer WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablesales_ccode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # status code editing finished and connect
    def cb_consumablesales_status_descr_changed(self):
        self.entry_consumablesales_status.clear()
        selected_item = self.cb_consumablesales_status_descr.currentText()

        if selected_item:
            query = f"SELECT DISTINCT scode From status WHERE sdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablesales_status]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # employee name editing finished and connect
    def employee_name_changed(self):

        self.entry_consumablesales_ecode.clear()
        selected_item = self.cb_consumablesales_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablesales_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # product name editing finished and connect
    def product_name_changed(self):
        self.entry_consumablesales_pcode.clear()
        self.entry_consumablesales_um.clear()
        self.entry_consumablesales_unitprice.clear()

        selected_item = self.cb_consumablesales_pname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT pcode, um, salesprice From vw_price_sales WHERE pdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_consumablesales_pcode, self.entry_consumablesales_um, self.entry_consumablesales_unitprice]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # quantity changed and connect
    def qty_changed(self):
        self.entry_consumablesales_amount.clear()
        
        u_price = int(self.entry_consumablesales_unitprice.text())
        #qtyadd = int(self.entry_consumablesales_qty.text())
        #lenqty = len(self.entry_consumablesales_qty.text())
        qty_text = self.entry_consumablesales_qty.text()
        qtyadd = 0 if qty_text == '' else int(qty_text)
        lenqty = len(qty_text)

        if u_price > 0 and lenqty > 0:
            amount = u_price * qtyadd
            self.entry_consumablesales_amount.setText(str(amount))       
        else:
            return

    # clear input field entry
    def clear_data(self):
        self.lbl_consumablesales_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

        self.cb_consumablesales_cname.setCurrentIndex(0)
        self.cb_consumablesales_ename.setCurrentIndex(0)
        self.cb_consumablesales_pname.setCurrentIndex(0)
        self.entry_consumablesales_qty.setText("0")
        self.entry_consumablesales_unitprice.setText("0")
        self.entry_consumablesales_amount.setText("0")
        self.entry_consumablesales_status.setText("2")
        self.cb_consumablesales_status_descr.setCurrentIndex(1)
        self.display_currentdate()

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(15):  # 15 columns
            cell_text = self.tv_consumablesales.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_consumablesales_id.setText(cell_values[0])
        self.entry_consumablesales_ccode.setText(cell_values[1])
        self.cb_consumablesales_cname.setCurrentText(cell_values[2])
        self.entry_consumablesales_ecode.setText(cell_values[3])
        self.cb_consumablesales_ename.setCurrentText(cell_values[4])
        self.entry_consumablesales_pcode.setText(cell_values[5])
        self.cb_consumablesales_pname.setCurrentText(cell_values[6])
        self.entry_consumablesales_um.setText(cell_values[7])
        self.entry_consumablesales_qty.setText(cell_values[8])
        self.entry_consumablesales_unitprice.setText(cell_values[9])
        self.entry_consumablesales_amount.setText(cell_values[10])
        self.entry_consumablesales_trxdt.setText(cell_values[11])
        self.entry_consumablesales_status.setText(cell_values[12])
        self.cb_consumablesales_status_descr.setCurrentText(cell_values[13])
        self.entry_consumablesales_remark.setText(cell_values[14])
    
    # Clear all entry and combo boxes
    def refresh_data(self):
        self.clear_data()
        self.make_data()
        
        self.display_currentdate() # 현재 날짜 기본입력
        self.cb_consumablesales_status_descr.setCurrentIndex(1) # 출고 기본 입력

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_ConsumableSales.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = ConsumableSalesDialog()
    dialog.show()
    sys.exit(app.exec())