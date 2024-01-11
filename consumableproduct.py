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
#from consumableproduct_ui import Ui_ConsumableProductDialog

# product table contents -----------------------------------------------------
class ConsumableProductDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class ConsumableProductDialog(QDialog, Ui_ConsumableProductDialog, SubWindowBase): 
    
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database2()

        uic.loadUi("consumableproduct.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_consumableproduct and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "", "", "",]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_consumableproduct)
        self.tv_consumableproduct.setItemDelegate(delegate)
        self.tv_consumableproduct.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_consumableproduct.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_consumableproduct.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_consumableproduct.verticalHeader().setVisible(False)

        # While selecting row in tv_consumableproduct, each cell values to displayed to designated widgets
        self.tv_consumableproduct.clicked.connect(self.show_selected_data)

        # Initial Display of data
        self.make_data()
        self.connect_btn_method()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Make log file
        self.make_logfiles("access_consumableproduct.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_consumableproduct, partial(self.copy_cells, self.tv_consumableproduct))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_consumableproduct, partial(self.paste_cells, self.tv_consumableproduct))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_consumableproduct, partial(self.handle_return_key, self.tv_consumableproduct))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_consumableproduct
        self.process_key_event(event, tv_widget)

    # Connect button to method
    def connect_btn_method(self):
        self.pb_consumableproduct_show.clicked.connect(self.make_data)
        self.pb_consumableproduct_cancel.clicked.connect(self.close_dialog)
        self.pb_consumableproduct_clear.clicked.connect(self.clear_data)
        
        self.pb_consumableproduct_insert.clicked.connect(self.tb_insert)
        self.pb_consumableproduct_update.clicked.connect(self.tb_update)
        self.pb_consumableproduct_delete.clicked.connect(self.tb_delete)

    # tab order for product window
    def set_tab_order(self):
        widgets = [self.pb_consumableproduct_show, self.entry_consumableproduct_code, self.entry_consumableproduct_name,
            self.entry_consumableproduct_um, self.entry_consumableproduct_active, self.entry_consumableproduct_remark,
            self.pb_consumableproduct_clear, self.pb_consumableproduct_insert, self.pb_consumableproduct_update,
            self.pb_consumableproduct_delete, self.pb_consumableproduct_cancel]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_consumableproduct

        self.cursor.execute("SELECT id, pcode, pdescription, um, active, remark FROM product WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]    

        sql_query = "Select id, pcode, pdescription, um, active, remark From product Order By id desc"
        column_widths = [80, 100, 100, 50, 50, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths

    # show product table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_product_input(self):
        pcode = int(self.entry_consumableproduct_code.text())
        pname = str(self.entry_consumableproduct_name.text())
        um = str(self.entry_consumableproduct_um.text())
        active = str(self.entry_consumableproduct_active.text())
        remark = str(self.entry_consumableproduct_remark.text())

        return pcode, pname, um, active, remark

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
            
            idx = self.max_row_id("product")                                    # Get the max id 
            username, user_id, formatted_datetime = self.common_values_set()
            pcode, pname, um, active, remark = self.get_product_input()         # Get the value of other variables

            if (idx>0 and pcode>0) and all(len(var) > 0 for var in (pname, um, active)):

                self.cursor.execute('''INSERT INTO product (id, pcode, pdescription, um, active, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, pcode, pname, um, active, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data() 
                logging.info(f"User {username} inserted {idx} row to the product table.")
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
            
            idx = int(self.lbl_consumableproduct_id.text())
            username, user_id, formatted_datetime = self.common_values_set() 
            pcode, pname, um, active, remark = self.get_product_input()  # Get the value of other variables      

            if (idx>0 and pcode>0) and all(len(var) > 0 for var in (pname, um, active)):
                self.cursor.execute('''UPDATE product SET 
                            pcode=?, pdescription=?, um=?, active=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                            , (pcode, pname, um, active, formatted_datetime, user_id, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()  
                logging.info(f"User {username} updated row number {idx} in the product table.")
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
            idx = self.lbl_consumableproduct_id.text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM product WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()  
            logging.info(f"User {username} deleted {idx} row to the product table.")                
            
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # clear input field entry
    def clear_data(self):
        self.lbl_consumableproduct_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(6):  
            cell_text = self.tv_consumableproduct.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_consumableproduct_id.setText(cell_values[0])
        self.entry_consumableproduct_code.setText(cell_values[1])
        self.entry_consumableproduct_name.setText(cell_values[2])
        self.entry_consumableproduct_um.setText(cell_values[3])
        self.entry_consumableproduct_active.setText(cell_values[4])
        self.entry_consumableproduct_remark.setText(cell_values[5])
    
    def refresh_data(self):
        self.clear_data()
        self.make_data()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_ConsumableProduct.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = ConsumableProductDialog()
    dialog.show()
    sys.exit(app.exec())