import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QShortcut, QMenu, QInputDialog
from PyQt5.QtCore import Qt
from datetime import datetime
from cal import CalendarView
from commonmd import *
#for non_ui version-------------------------
#from gpai_payment_view_ui import Ui_GPAIViewDialog

# Table contents -----------------------------------------------------
class GPAIViewDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class GPAIViewDialog(QDialog, Ui_GPAIViewDialog, SubWindowBase):
     
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()    
  
        self.conn, self.cursor = connect_to_database3()

        # load ui file
        uic.loadUi("gpai_payment_view.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_gpaiview and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "", "numeric", "", "", "numeric"]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_gpaiview)
        self.tv_gpaiview.setItemDelegate(delegate)
        self.tv_gpaiview.setModel(self.proxy_model)

        # Enable sorting
        self.tv_gpaiview.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_gpaiview.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_gpaiview.verticalHeader().setVisible(False)

        # While selecting row in tv_gpaiview, each cell values to displayed to designated widgets
        self.tv_gpaiview.clicked.connect(self.show_selected_data)

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

        # Make log file
        self.make_logfiles("access_GPAIViewDialog.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_gpaiview, partial(self.copy_cells, self.tv_gpaiview))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_gpaiview, partial(self.paste_cells, self.tv_gpaiview))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_gpaiview, partial(self.handle_return_key, self.tv_gpaiview))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_gpaiview
        self.process_key_event(event, tv_widget)

    
    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_gpaiview_ename, "SELECT DISTINCT ename FROM employee ORDER BY ename")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.lbl_gpaiview_id.setText("")
        self.entry_gpaiview_ecode.setText("")
        self.cb_gpaiview_ename.setCurrentIndex(0) 

    # Connect button to method
    def conn_button_to_method(self):
        self.pb_gpaiview_show.clicked.connect(self.make_data)
        self.pb_gpaiview_search.clicked.connect(self.search_data)        
        self.pb_gpaiview_close.clicked.connect(self.close_dialog)
        self.pb_gpaiview_clear.clicked.connect(self.clear_data)
        
    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_gpaiview_ename.activated.connect(self.cb_gpaiview_ename_changed)

    # tab order for employee window
    def set_tab_order(self):
       
        widgets = [self.pb_gpaiview_show, self.entry_gpaiview_ecode, self.cb_gpaiview_ename,
            self.entry_gpaiview_class1, self.entry_gpaiview_indval, self.entry_gpaiview_gdate,
            self.entry_gpaiview_remark, 
            self.pb_gpaiview_search, self.pb_gpaiview_clear, self.pb_gpaiview_close, ]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_gpaiview
        self.cursor.execute("SELECT * FROM vw_gpai_gen_view WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_gpai_gen_view Order By ename"
        column_widths = [80, 100, 100, 100, 80, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show employee table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime    


    # Search data
    def search_data(self):
      
        ename = self.cb_gpaiview_ename.currentText()
        class1 = self.entry_gpaiview_class1.text()
        gdate = self.entry_gpaiview_gdate.text()

        conditions = {'v01': (ename, "ename like '%{}%'"),
                      'v02': (class1, "class1 like '%{}%'"),
                      'v03': (gdate, "gdate = #{}#"),
                      }
        
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_gpai_gen_view WHERE {' AND '.join(selected_conditions)} ORDER BY ename"

        QMessageBox.about(self, "검색 조건 확인", f"지급예정일: {gdate} \n직원명: {ename} \n구분: {class1} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)


    # Employee Name Index Changed
    def cb_gpaiview_ename_changed(self):
        self.entry_gpaiview_ecode.clear()
        selected_item = self.cb_gpaiview_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_gpaiview_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass
        
    # clear input field entry
    def clear_data(self):
        self.lbl_gpaiview_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(6):  # 6columns
            cell_text = self.tv_gpaiview.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_gpaiview_id.setText(cell_values[0])
        self.entry_gpaiview_gdate.setText(cell_values[1])
        self.entry_gpaiview_ecode.setText(cell_values[2])
        self.cb_gpaiview_ename.setCurrentText(cell_values[3])
        self.entry_gpaiview_class1.setText(cell_values[4])
        self.entry_gpaiview_indval.setText(cell_values[5])

    
    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_GPAIViewDialog.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
        
    app = QtWidgets.QApplication(sys.argv)
    dialog = GPAIViewDialog()
    dialog.show()
    sys.exit(app.exec())