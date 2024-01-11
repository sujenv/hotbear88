import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QMenu, QInputDialog, QShortcut
from PyQt5.QtCore import Qt
from datetime import datetime
from commonmd import *
#for non_ui version-------------------------
#from consumableoutprice_ui import Ui_ConsumableOutPriceDialog

# Consumable Sales Price -----------------------------------------------------
class ConsumableOutPriceDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class ConsumableOutPriceDialog(QDialog, Ui_ConsumableOutPriceDialog, SubWindowBase): 
    
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database2()

        uic.loadUi("consumableoutprice.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_consumableoutprice and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)        
        
        # Define the column types
        column_types = ["numeric", "numeric", "", "numeric", "", "", "numeric", "numeric", "", "", "", ]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_consumableoutprice)
        self.tv_consumableoutprice.setItemDelegate(delegate)
        self.tv_consumableoutprice.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_consumableoutprice.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_consumableoutprice.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_consumableoutprice.verticalHeader().setVisible(False)

        # While selecting row in tv_consumableoutprice, each cell values to displayed to designated widgets
        self.tv_consumableoutprice.clicked.connect(self.show_selected_data)

        # Fill combobox items when the application starts
        self.get_combobox_contents()

        # Initial Display of data
        self.make_data()
        self.connect_btn_method()
        self.conn_signal_to_slot()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        self.display_eff_date()
        self.entry_consumableoutprice_qty.setText("1")

        # Make log file
        self.make_logfiles("access_ConsumableOutPrice.log")

    # Display current date only
    def display_eff_date(self):
        now = datetime.now()
        curr_date = now.strftime("%Y/%m/%d")
        ddt = f"{curr_date}"         
        endofdate = "2050/12/31"

        self.entry_consumableoutprice_efffrom.setText(ddt)
        self.entry_consumableoutprice_effthru.setText(endofdate)

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_consumableoutprice, partial(self.copy_cells, self.tv_consumableoutprice))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_consumableoutprice, partial(self.paste_cells, self.tv_consumableoutprice))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_consumableoutprice, partial(self.handle_return_key, self.tv_consumableoutprice))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_consumableoutprice
        self.process_key_event(event, tv_widget)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_consumableoutprice_cname, "SELECT DISTINCT cname FROM customer where type01='s'")
        self.insert_combobox_initiate(self.cb_consumableoutprice_pname, "SELECT DISTINCT pdescription FROM product ORDER BY pdescription")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):    
        self.combobox_initializing(combo_box, sql_query) 
        self.cb_consumableoutprice_cname.setCurrentIndex(0) 
        self.cb_consumableoutprice_pname.setCurrentIndex(0)

    # Connect button to method
    def connect_btn_method(self):
        self.pb_consumableoutprice_show.clicked.connect(self.make_data)
        self.pb_consumableoutprice_cancel.clicked.connect(self.close_dialog)
        self.pb_consumableoutprice_clear.clicked.connect(self.clear_data)
        
        self.pb_consumableoutprice_insert.clicked.connect(self.tb_insert)
        self.pb_consumableoutprice_update.clicked.connect(self.tb_update)
        self.pb_consumableoutprice_delete.clicked.connect(self.tb_delete)

    # Connect signal to method    
    def conn_signal_to_slot(self):
        self.cb_consumableoutprice_cname.activated.connect(self.customer_description_changed)
        self.cb_consumableoutprice_pname.activated.connect(self.product_name_changed)

    # tab order for salesprice window
    def set_tab_order(self):
        
        widgets = [self.pb_consumableoutprice_show, self.entry_consumableoutprice_ccode, self.cb_consumableoutprice_cname,
            self.entry_consumableoutprice_pcode, self.cb_consumableoutprice_pname, self.entry_consumableoutprice_um,
            self.entry_consumableoutprice_qty, self.entry_consumableoutprice_unitprice, self.entry_consumableoutprice_efffrom,
            self.entry_consumableoutprice_effthru, self.entry_consumableoutprice_remark, self.pb_consumableoutprice_clear,
            self.pb_consumableoutprice_insert, self.pb_consumableoutprice_update, self.pb_consumableoutprice_delete,
            self.pb_consumableoutprice_cancel]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_consumableoutprice

        self.cursor.execute("SELECT * FROM vw_price_sales where 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]   

        sql_query = "Select * From vw_price_sales"
        column_widths = [80, 100, 100, 100, 100, 100]
        
        return sql_query, tv_widget, column_info, column_names, column_widths
    
    # show salesprice table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_sales_price_input(self):
        ccode = int(self.entry_consumableoutprice_ccode.text())
        pcode = int(self.entry_consumableoutprice_pcode.text())
        uprice = float(self.entry_consumableoutprice_unitprice.text())
        efffrom = str(self.entry_consumableoutprice_efffrom.text())
        effthru = str(self.entry_consumableoutprice_effthru.text())
        remark = str(self.entry_consumableoutprice_remark.text())

        return ccode, pcode, uprice, efffrom, effthru, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime

    # insert new salesprice data to MySQL table
    def tb_insert(self):
        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:

            idx = self.max_row_id2("salesprice")                                                                  # Get the max id 
            username, user_id, formatted_datetime = self.common_values_set()
            ccode, pcode, uprice, efffrom, effthru, remark = self.get_sales_price_input()           # Get the value of other variables

            if (idx>0 and ccode>0 and pcode>0 and abs(uprice)>=0) and all(len(var) > 0 for var in (efffrom, effthru)):
                self.cursor.execute('''INSERT INTO salesprice (id, ccode, pcode, salesprice, efffrom, effthru, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, ccode, pcode, uprice, efffrom, effthru, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data() 
                logging.info(f"User {username} inserted {idx} row to the Sales Price table.")
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
            
            idx = int(self.lbl_consumableoutprice_id.text())
            username, user_id, formatted_datetime = self.common_values_set()    
            ccode, pcode, uprice, efffrom, effthru, remark = self.get_sales_price_input()         # Get the value of other variables       

            if (idx>0 and ccode>0 and pcode>0 and abs(uprice)>=0) and all(len(var) > 0 for var in (efffrom, effthru)):
                self.cursor.execute('''UPDATE salesprice SET 
                            ccode=?, pcode=?, salesprice=?, efffrom=?, effthru=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                            , (ccode, pcode, uprice, efffrom, effthru, formatted_datetime, user_id, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()  
                logging.info(f"User {username} updated row number {idx} in the Sales Price table.")
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
            idx = self.lbl_consumableoutprice_id.text()
            username, user_id, formatted_datetime = self.common_values_set()            
            self.cursor.execute("DELETE FROM salesprice WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()  
            logging.info(f"User {username} deleted {idx} row to the Sales Priece table.")            
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # customer description editing finished and connect
    def customer_description_changed(self):
        self.entry_consumableoutprice_ccode.clear()
        selected_item = self.cb_consumableoutprice_cname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From customer WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_consumableoutprice_ccode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                return

    # product name editing finished and connect
    def product_name_changed(self):
        self.entry_consumableoutprice_pcode.clear()

        selected_item = self.cb_consumableoutprice_pname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT pcode, um From product WHERE pdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_consumableoutprice_pcode, self.entry_consumableoutprice_um]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # clear input field entry
    def clear_data(self):
        self.lbl_consumableoutprice_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()
        
        self.display_eff_date()            
        self.entry_consumableoutprice_qty.setText("1")
        self.cb_consumableoutprice_cname.setCurrentIndex(0)
        self.cb_consumableoutprice_pname.setCurrentIndex(0)

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(11):  # 11columns
            cell_text = self.tv_consumableoutprice.model().item(row_index, column_index).text()
            cell_values.append(cell_text)                                        

        # Populate the input widgets with the data from the selected row
        self.lbl_consumableoutprice_id.setText(cell_values[0])
        self.entry_consumableoutprice_ccode.setText(cell_values[1])
        self.cb_consumableoutprice_cname.setCurrentText(cell_values[2])
        self.entry_consumableoutprice_pcode.setText(cell_values[3])
        self.cb_consumableoutprice_pname.setCurrentText(cell_values[4])
        self.entry_consumableoutprice_um.setText(cell_values[5])
        self.entry_consumableoutprice_qty.setText(cell_values[6])
        self.entry_consumableoutprice_unitprice.setText(cell_values[7])
        self.entry_consumableoutprice_efffrom.setText(cell_values[8])
        self.entry_consumableoutprice_effthru.setText(cell_values[9])
        self.entry_consumableoutprice_remark.setText(cell_values[10])
    
    def refresh_data(self):
        self.clear_data()
        self.make_data()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_salesprice.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = ConsumableOutPriceDialog()
    dialog.show()
    sys.exit(app.exec())