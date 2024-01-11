import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QShortcut
from PyQt5.QtCore import Qt,QTimer
from datetime import datetime
from commonmd import *
#for non_ui version-------------------------
#from inventory_ui import Ui_DialogInventory

# inventory table contents -----------------------------------------------------
class CurrentInventoryDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class CurrentInventoryDialog(QDialog, Ui_DialogInventory, SubWindowBase): 
    
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database2()

        uic.loadUi("inventory.ui", self)

        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_inventory and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "", "numeric","numeric","numeric","numeric",]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_inventory)
        self.tv_inventory.setItemDelegate(delegate)
        self.tv_inventory.setModel(self.proxy_model)

        # Enable sorting
        self.tv_inventory.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_inventory.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_inventory.verticalHeader().setVisible(False)

        # While selecting row in tv_inventory, each cell values to displayed to designated widgets
        self.tv_inventory.clicked.connect(self.show_selected_data)

        # Initiate combo boxes
        self.combobox_initiation()

        # Initial Display of data
        self.make_data()
        self.conn_button_to_method()
        self.entry_inventory_datetime.setText("Initializing...")

        # Set up a QTimer to update the datetime label every second
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_datetime)
        # Update every 1000 milliseconds (1 second)
        self.update_timer.start(1000)  

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

    # Refresh datetime every second
    def update_datetime(self):
        now = datetime.now()
        curr_date = now.strftime("%Y-%m-%d")
        curr_time = now.strftime("%H:%M:%S")
        dtime = f"{curr_date} {curr_time}"
        self.entry_inventory_datetime.setText(dtime)

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_inventory, partial(self.copy_cells, self.tv_inventory))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_inventory, partial(self.paste_cells, self.tv_inventory))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_inventory, partial(self.handle_return_key, self.tv_inventory))
    
    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_inventory
        self.process_key_event(event, tv_widget)

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_inventory_show.clicked.connect(self.make_data)
        self.pb_inventory_cancel.clicked.connect(self.close_dialog)
        self.pb_inventory_clearinput.clicked.connect(self.clear_data)
        self.pb_inventory_asis.clicked.connect(self.make_on_hand)

    # tab order for inventory window
    def set_tab_order(self):
        widgets = [self.pb_inventory_show, self.entry_inventory_name, 
                   self.entry_inventory_qty, self.pb_inventory_clearinput]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_inventory
        self.cursor.execute("SELECT * FROM vw_inventory_now  WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]
    
        sql_query = f"Select * From vw_inventory_now"
        column_widths = [80, 100, 100]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show inventory table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Initiate Combo_Box 
    def combobox_initiation(self):
        self.cb_inventory_year.clear() # Clear existing items
        self.cb_inventory_year.addItem("")  # Add a blank item as the first option
        self.cb_inventory_month.clear()
        self.cb_inventory_month.addItem("")
        self.cb_inventory_day.clear()
        self.cb_inventory_day.addItem("")

        for year in range(2001, 2101):  # Range includes 2001 but not 2101
            self.cb_inventory_year.addItem(str(year))

        for month in range(1, 13):
            self.cb_inventory_month.addItem(str(month))

        for day in range(1, 32):
            self.cb_inventory_day.addItem(str(day))

        current_date = datetime.now()
        current_year = current_date.year
        current_month = current_date.month
        current_day = current_date.day

        self.cb_inventory_year.setCurrentText(str(current_year))
        self.cb_inventory_month.setCurrentText(str(current_month))
        self.cb_inventory_day.setCurrentText(str(current_day))

    # Make On Hand
    def make_on_hand(self):
        
        year = int(self.cb_inventory_year.currentText())
        month = int(self.cb_inventory_month.currentText())
        day = int(self.cb_inventory_day.currentText())
        desired_date = date(year, month, day)

        sql_query = f''' SELECT pcode, '' as pdescription, 0 as ini_Q, 0 as rec_Q, 0 as sale_Q, 0 as On_hand FROM vw_inventory_01'''
        self.cursor.execute(sql_query)
        column_info1 = self.cursor.description
        column_names1 = [col[0] for col in column_info1]
        
        query = f'''
        SELECT vw_inventory_01.pcode, product.pdescription, sum(vw_inventory_01.iniQ) as ini_Q, sum(vw_inventory_01.recQ) as rec_Q, sum(vw_inventory_01.salQ) as sale_Q,
        sum(vw_inventory_01.iniQ) + sum(vw_inventory_01.recQ) - sum(vw_inventory_01.salQ) as On_hand
        FROM vw_inventory_01
        LEFT JOIN product ON vw_inventory_01.pcode = product.pcode
        WHERE vw_inventory_01.trx_date <= #{desired_date}#
        Group BY vw_inventory_01.pcode, product.pdescription
        ORDER BY vw_inventory_01.pcode;'''    
        column_widths1 = [80, 100, 100, 100, 100, 100]

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info1, column_names1, column_widths1)

    # clear input field entry
    def clear_data(self):
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()        
       
    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(3):  # 3 columns
            cell_text = self.tv_inventory.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.entry_inventory_id.setText(cell_values[0])
        self.entry_inventory_name.setText(cell_values[1])
        self.entry_inventory_qty.setText(cell_values[2])

    def refresh_data(self):
        self.clear_data()
        self.make_data()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    dialog = CurrentInventoryDialog()
    dialog.show()
    sys.exit(app.exec())