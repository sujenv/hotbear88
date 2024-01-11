import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QShortcut
from PyQt5.QtCore import Qt
from datetime import datetime
from commonmd import *
#for non_ui version-------------------------
#from conversion_ui import Ui_ConversionDialog

# conversion table contents -----------------------------------------------------
class ConversionDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class ConversionDialog(QDialog, Ui_ConversionDialog, SubWindowBase): 
    
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database2()

        uic.loadUi("conversion.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)
        
        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_conversion and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "", "numeric", "", ""]
        
        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_conversion)
        self.tv_conversion.setItemDelegate(delegate)
        self.tv_conversion.setModel(self.proxy_model)

        # Enable sorting
        self.tv_conversion.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_conversion.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_conversion.verticalHeader().setVisible(False)

        # While selecting row in tv_conversion, each cell values to displayed to designated widgets
        self.tv_conversion.clicked.connect(self.show_selected_data)

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

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_conversion, partial(self.copy_cells, self.tv_conversion))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_conversion, partial(self.paste_cells, self.tv_conversion))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_conversion, partial(self.handle_return_key, self.tv_conversion))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_conversion
        self.process_key_event(event, tv_widget)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_conversion_pname, "SELECT DISTINCT pdescription FROM product ORDER BY pdescription")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):    
        self.combobox_initializing(combo_box, sql_query) # using common module
        self.cb_conversion_pname.setCurrentIndex(0)

    # Connect button to method
    def connect_btn_method(self):
        self.pb_conversion_show.clicked.connect(self.make_data)
        self.pb_conversion_cancel.clicked.connect(self.close_dialog)
        self.pb_conversion_clearinput.clicked.connect(self.clear_data)
        
        self.pb_conversion_insert.clicked.connect(self.tb_insert)
        self.pb_conversion_update.clicked.connect(self.tb_update)
        self.pb_conversion_delete.clicked.connect(self.tb_delete)

    # Connect signal to method    
    def conn_signal_to_slot(self):
        self.cb_conversion_pname.activated.connect(self.product_name_changed)

    # tab order for conversion window
    def set_tab_order(self):
        widgets = [self.pb_conversion_show, self.entry_conversion_code, self.cb_conversion_pname,
            self.entry_conversion_um1, self.entry_conversion_cf, self.entry_conversion_um2,
            self.entry_conversion_remark, self.pb_conversion_clearinput, self.pb_conversion_insert,
            self.pb_conversion_update, self.pb_conversion_delete, self.pb_conversion_cancel]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])        

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_conversion

        sql_query = '''Select conversion.id, conversion.pcode, product.pdescription, conversion.um1, conversion.cf, conversion.um2, conversion.remark From conversion 
                        Left join product on (conversion.pcode = product.pcode and conversion.um1 = product.um)
                        Order By conversion.id desc'''
        
        query = "Select id, pcode, 0 as pdescription, um1, cf, um2, remark From conversion WHERE 1=0"
        self.cursor.execute(query)
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        column_widths = [80, 100, 100, 50, 50, 50, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show conversion table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)
  
    # Get the value of other variables
    def get_conversion_input(self):
        pcode = int(self.entry_conversion_code.text())
        um01 = str(self.entry_conversion_um1.text())
        cf = int(self.entry_conversion_cf.text())
        um02 = str(self.entry_conversion_um2.text())
        remark = str(self.entry_conversion_remark.text())

        return pcode, um01, cf, um02, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime

    # insert new conversion data to MySQL table
    def tb_insert(self):

        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:
            
            idx = self.max_row_id("conversion")                                  # Get the max id 
            username, user_id, formatted_datetime = self.common_values_set()
            pcode, um01, cf, um02, remark = self.get_conversion_input()         # Get the value of other variables

            if (idx>0 and pcode>0 and cf>0) and all(len(var) > 0 for var in (um01, um02)):

                self.cursor.execute('''INSERT INTO conversion (id, pcode, um1, cf, um2, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, pcode, um01, cf, um02, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data() 
                logging.info(f"User {username} inserted {idx} row to the conversion table.")
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
            idx = int(self.lbl_conversion_id.text())
            username, user_id, formatted_datetime = self.common_values_set()         
            pcode, um01, cf, um02, remark = self.get_conversion_input()         # Get the value of other variables

            if (idx>0 and pcode>0 and cf>0) and all(len(var) > 0 for var in (um01, um02)):
                self.cursor.execute('''UPDATE conversion SET 
                            pcode=?, um1=?, cf=?, um2=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                            , (pcode, um01, cf, um02, formatted_datetime, user_id, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()  
                logging.info(f"User {username} updated row number {idx} in the conversion table.")
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

            idx = self.lbl_conversion_id.text()
            username, user_id, formatted_datetime = self.common_values_set()

            self.cursor.execute("DELETE FROM conversion WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()  
            logging.info(f"User {username} deleted {idx} row to the conversion table.")                
            
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # product name editing finished and connect
    def product_name_changed(self):
        self.entry_conversion_code.clear()
        selected_item = self.cb_conversion_pname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT pcode From product WHERE pdescription ='{selected_item}'"
            line_edit_widgets = [self.entry_conversion_code]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # clear input field entry
    def clear_data(self):
        self.lbl_conversion_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

        self.cb_conversion_pname.setCurrentIndex(0)

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(7):  # 7columns
            cell_text = self.tv_conversion.model().item(row_index, column_index).text()
            cell_values.append(cell_text)         

        # Populate the input widgets with the data from the selected row
        self.lbl_conversion_id.setText(cell_values[0])
        self.entry_conversion_code.setText(cell_values[1])
        self.cb_conversion_pname.setCurrentText(cell_values[2])
        self.entry_conversion_um1.setText(cell_values[3])
        self.entry_conversion_cf.setText(cell_values[4])
        self.entry_conversion_um2.setText(cell_values[5])
        self.entry_conversion_remark.setText(cell_values[6])
    
    def refresh_data(self):
        self.clear_data()
        self.make_data()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_Conversion.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = ConversionDialog()
    dialog.show()
    sys.exit(app.exec())