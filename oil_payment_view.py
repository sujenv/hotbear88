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
#from oil_payment_view_ui import Ui_OilUsageViewDialog

# Table contents -----------------------------------------------------
class OilUsageViewDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class OilUsageViewDialog(QDialog, Ui_OilUsageViewDialog, SubWindowBase):
     
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()    
  
        self.conn, self.cursor = connect_to_database3()

        # load ui file
        uic.loadUi("oil_payment_view.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_oilusageview and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["", "numeric", "", "numeric",]
        column_types1 = ["numeric", "", "", "", "numeric",] 
        
        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_oilusageview)
        self.tv_oilusageview.setItemDelegate(delegate)
        self.tv_oilusageview.setModel(self.proxy_model)

        delegate1 = NumericDelegate(column_types1, self.tv_oilusageviewdetail)
        self.tv_oilusageviewdetail.setItemDelegate(delegate1)
        self.tv_oilusageviewdetail.setModel(self.proxy_model)

        # Enable sorting
        self.tv_oilusageview.setSortingEnabled(True)
        self.tv_oilusageviewdetail.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_oilusageview.setAlternatingRowColors(True)  
        self.tv_oilusageviewdetail.setAlternatingRowColors(True)  

        # Hide the first index column
        self.tv_oilusageview.verticalHeader().setVisible(False)
        self.tv_oilusageviewdetail.verticalHeader().setVisible(False)

        # While selecting row in tv_oilusageview, each cell values to displayed to designated widgets
        self.tv_oilusageview.clicked.connect(self.show_selected_data)
        self.tv_oilusageviewdetail.clicked.connect(self.show_selected_data_detail)

        # Fill combobox items when the application starts
        self.get_combobox_contents()

        # Initiate display of data
        self.make_data()
        self.make_data1()         
        self.conn_button_to_method()
        self.connect_signal_slot()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Create context menus ----------------------------------------------------------------------------------
        self.context_menu1 = self.create_context_menu(self.entry_oilusageview_paydt)

        self.entry_oilusageview_paydt.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_oilusageview_paydt.customContextMenuRequested.connect(self.show_context_menu1)

        # Make log file
        self.make_logfiles("access_OilUsageViewDialog.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_oilusageview, partial(self.copy_cells, self.tv_oilusageview))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_oilusageview, partial(self.paste_cells, self.tv_oilusageview))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_oilusageview, partial(self.handle_return_key, self.tv_oilusageview))

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu
    
    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_oilusageview_paydt.mapToGlobal(pos))
 
     # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_oilusageview_paydt:
            target_lineedit.setText(date)

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_oilusageview
        self.process_key_event(event, tv_widget)

    
    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_oilusageview_cname, "SELECT DISTINCT cname FROM company ORDER BY cname")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.entry_oilusageview_ccode.setText("")
        self.cb_oilusageview_cname.setCurrentIndex(0) 

    # Connect button to method
    def conn_button_to_method(self):
        self.pb_oilusageview_show.clicked.connect(self.make_data)
        self.pb_oilusageview_search.clicked.connect(self.search_data)        
        self.pb_oilusageview_close.clicked.connect(self.close_dialog)
        self.pb_oilusageview_clear.clicked.connect(self.clear_data)
        self.pb_oilusageview_delete.clicked.connect(self.tb_delete)
        
    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_oilusageview_cname.activated.connect(self.cb_oilusageview_cname_changed)

    # tab order for employee window
    def set_tab_order(self):
       
        widgets = [self.pb_oilusageview_show, self.entry_oilusageview_ccode, self.cb_oilusageview_cname,
            self.entry_oilusageview_payval, self.entry_oilusageview_paydt, self.entry_oilusageview_getdt,
            self.pb_oilusageview_search, self.pb_oilusageview_clear, self.pb_oilusageview_close, ]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_oilusageview
        self.cursor.execute("SELECT * FROM vw_paymentoilusagecom WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_paymentoilusagecom Order By paydt"
        column_widths = [100, 100, 100, 100]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show employee table data inside of the MDI
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)


    # To reduce duplications
    def common_query_statement1(self):
        tv_widget = self.tv_oilusageviewdetail
        self.cursor.execute("SELECT * FROM vw_paymentoilusagecomdetail WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_paymentoilusagecomdetail Order By paydt"
        column_widths = [100, 100, 100, 100, 100, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show employee table data inside of the MDI
    def make_data1(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement1() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime    

    # delete row according to id selected
    def tb_delete(self):
        confirm_dialog = self.show_delete_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            
            paydt = self.entry_oilusageview_paydt.text()
            username, user_id, formatted_datetime = self.common_values_set()

            if len(paydt) > 0: 
                self.cursor.execute("DELETE FROM paymentoilusagcom WHERE paydt=?", (paydt,))
                self.conn.commit()
                self.show_delete_success_message()
                self.refresh_data()  
                logging.info(f"User {username} deleted paydt {paydt}, at the paymentoilusagcom table.")
            else:
                return    
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # Search data
    def search_data(self):
      
        cname = self.cb_oilusageview_cname.currentText()
        paydt = self.entry_oilusageview_paydt.text()

        conditions = {'v01': (cname, "cname like '%{}%'"),
                      'v02': (paydt, "paydt = #{}#"),
                      }
        
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_paymentoilusagecom WHERE {' AND '.join(selected_conditions)} ORDER BY paydt"

        QMessageBox.about(self, "검색 조건 확인", f"지급예정일: {paydt} \n지급사명: {cname} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)


    # Employee Name Index Changed
    def cb_oilusageview_cname_changed(self):
        self.entry_oilusageview_ccode.clear()
        selected_item = self.cb_oilusageview_cname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From company WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_oilusageview_ccode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass
        
    # clear input field entry
    def clear_data(self):
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

    # table widget cell double click
    def show_selected_data(self, item):
        self.clear_data()
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(4):  # 4columns
            cell_text = self.tv_oilusageview.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.entry_oilusageview_paydt.setText(cell_values[0])
        self.entry_oilusageview_ccode.setText(cell_values[1])
        self.cb_oilusageview_cname.setCurrentText(cell_values[2])
        self.entry_oilusageview_payval.setText(cell_values[3])

    def show_selected_data_detail(self, item):

        self.clear_data()
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(5):  # 6columns
            cell_text = self.tv_oilusageviewdetail.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.entry_oilusageview_ccode.setText(cell_values[0])
        self.cb_oilusageview_cname.setCurrentText(cell_values[1])
        self.entry_oilusageview_paydt.setText(cell_values[2])
        self.entry_oilusageview_getdt.setText(cell_values[3])
        self.entry_oilusageview_payval.setText(cell_values[4])
        
    def refresh_data(self):
        self.clear_data()
        self.make_data()
        self.make_data1()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_OilUsageViewDialog.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
        
    app = QtWidgets.QApplication(sys.argv)
    dialog = OilUsageViewDialog()
    dialog.show()
    sys.exit(app.exec())