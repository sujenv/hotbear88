import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QShortcut, QMenu
from PyQt5.QtCore import Qt
from datetime import datetime
from cal import CalendarView
from commonmd import *
#for non_ui version-------------------------
#from bankacc_product_chk_ui import Ui_BankAccProductChkDialog

# Bank Account Product table contents -----------------------------------------------------
class BankAccProductChkDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class BankAccProductChkDialog(QDialog, Ui_BankAccProductchkDialog, SubWindowBase):
     
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()    
  
        self.conn, self.cursor = connect_to_database3()

        # load ui file
        uic.loadUi("bankacc_product_chk.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_bankproductchkpay and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "numeric", "", "", "", "", "", "", "", "", "numeric", ""]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_bankproductchkpay)
        self.tv_bankproductchkpay.setItemDelegate(delegate)
        self.tv_bankproductchkpay.setModel(self.proxy_model)

        # Enable sorting
        self.tv_bankproductchkpay.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_bankproductchkpay.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_bankproductchkpay.verticalHeader().setVisible(False)

        # While selecting row in tv_bankproductchkpay, each cell values to displayed to designated widgets
        self.tv_bankproductchkpay.clicked.connect(self.show_selected_data)

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

        # Create context menus ----------------------------------------------------------------------------------
        self.context_menu1 = self.create_context_menu(self.entry_bankproductchk_paydate)

        self.entry_bankproductchk_paydate.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_bankproductchk_paydate.customContextMenuRequested.connect(self.show_context_menu1)

        # Make log file
        self.make_logfiles("access_BankAccProductChkDialog.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_bankproductchkpay, partial(self.copy_cells, self.tv_bankproductchkpay))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_bankproductchkpay, partial(self.paste_cells, self.tv_bankproductchkpay))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_bankproductchkpay, partial(self.handle_return_key, self.tv_bankproductchkpay))

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu

    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_bankproductchk_paydate.mapToGlobal(pos))

    # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_bankproductchk_paydate:
            target_lineedit.setText(date)

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_bankproductchkpay
        self.process_key_event(event, tv_widget)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_bankproductchk_cname, "SELECT DISTINCT cname FROM company ORDER BY cname")
        self.insert_combobox_initiate(self.cb_bankproductchk_ename, "SELECT DISTINCT ename FROM employee ORDER BY ename")
        self.insert_combobox_initiate(self.cb_bankproductchk_bkname, "SELECT DISTINCT bname FROM bankid ORDER BY bname")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.lbl_bankproductchk_id.setText("")
        self.entry_bankproductchk_ccode.setText("")
        self.entry_bankproductchk_ecode.setText("")
        self.cb_bankproductchk_cname.setCurrentIndex(0) 
        self.cb_bankproductchk_ename.setCurrentIndex(0) 


    # Connect button to method
    def conn_button_to_method(self):
        self.pb_bankproductchk_show.clicked.connect(self.make_data)
        self.pb_bankproductchk_search.clicked.connect(self.search_data)        
        self.pb_bankproductchk_close.clicked.connect(self.close_dialog)
        self.pb_bankproductchk_clear.clicked.connect(self.clear_data)

    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_bankproductchk_cname.activated.connect(self.cb_bankproductchk_cname_changed)
        self.cb_bankproductchk_ename.activated.connect(self.cb_bankproductchk_ename_changed)
        self.cb_bankproductchk_bkname.activated.connect(self.cb_bankproductchk_bkname_changed)
        self.entry_bankproductchk_sdt.editingFinished.connect(self.sdt_changed)

    # tab order for employee window
    def set_tab_order(self):
       
        widgets = [self.pb_bankproductchk_show, self.entry_bankproductchk_ccode, self.cb_bankproductchk_cname,
            self.entry_bankproductchk_ecode, self.cb_bankproductchk_ename, 
            self.entry_bankproductchk_baowner, self.entry_bankproductchk_remark, 
            self.pb_bankproductchk_search, self.pb_bankproductchk_clear, self.pb_bankproductchk_close]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_bankproductchkpay
        self.cursor.execute("SELECT * FROM vw_productpayment WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_productpayment Order By id"
        column_widths = [80, 100, 100, 100, 100, 100, 100, 100, 120, 50, 100, 150, 100, 150]

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
        
        cname = self.cb_bankproductchk_cname.currentText()
        ename = self.cb_bankproductchk_ename.currentText()
        bname = self.cb_bankproductchk_bkname.currentText()
        
        conditions = {'v01': (cname, "cname like '%{}%'"),
                      'v02': (ename, "ename like '%{}%'"),
                      'v03': (bname, "bname like '%{}%'"),
                      }
        
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_productpayment WHERE {' AND '.join(selected_conditions)} ORDER BY cname, ename"

        QMessageBox.about(self, "검색 조건 확인", f"업체명: {cname} \n직원명:{ename} \n은행명:{bname} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Combobox apt type index changed
    def cb_bankproductchk_cname_changed(self):
        self.entry_bankproductchk_ccode.clear()
        selected_item = self.cb_bankproductchk_cname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ccode From company WHERE cname ='{selected_item}'"
            line_edit_widgets = [self.entry_bankproductchk_ccode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    def cb_bankproductchk_ename_changed(self):
        self.entry_bankproductchk_ecode.clear()
        selected_item = self.cb_bankproductchk_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_bankproductchk_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Bank Name Index Changed
    def cb_bankproductchk_bkname_changed(self):
        self.entry_bankproductchk_bankid.clear()
        selected_item = self.cb_bankproductchk_bkname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT bcode From bankid WHERE bname ='{selected_item}'"
            line_edit_widgets = [self.entry_bankproductchk_bankid]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Effective Date Index Changed
    def sdt_changed(self):
        # inputed string type date
        date_string = self.entry_bankproductchk_sdt.text()
        # convert string type date to date format
        startdt = datetime.strptime(date_string, "%Y/%m/%d")
        # 날짜 형식을 원하는 형식으로 출력
        #date_object = startdt.strftime("%Y-%m-%d")

        # Find the last day of the month for the given date
        _, last_day = calendar.monthrange(startdt.year, startdt.month)
        last_day_of_month = datetime(startdt.year, startdt.month, last_day)
        
        # Calculate the end date, which is one day before the last day of the month
        #enddt = last_day_of_month - timedelta(days=1)
        enddt = last_day_of_month - timedelta(days=0)
        paydt = last_day_of_month + timedelta(days=15)
        
        # Format the end date as a string and set it to the desired widget
        enddt = enddt.strftime("%Y/%m/%d")
        paydt = paydt.strftime("%Y/%m/%d")
        
        self.entry_bankproductchk_edt.setText(enddt)
        self.entry_bankproductchk_paydate.setText(paydt)

    # clear input field entry
    def clear_data(self):
        self.lbl_bankproductchk_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(14):  # 14 columns
            cell_text = self.tv_bankproductchkpay.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_bankproductchk_id.setText(cell_values[0])
        self.entry_bankproductchk_ecode.setText(cell_values[1])
        self.cb_bankproductchk_ename.setCurrentText(cell_values[2])
        self.entry_bankproductchk_ccode.setText(cell_values[3])
        self.cb_bankproductchk_cname.setCurrentText(cell_values[4])
        self.entry_bankproductchk_paydate.setText(cell_values[5])
        self.entry_bankproductchk_sdt.setText(cell_values[6])
        self.entry_bankproductchk_edt.setText(cell_values[7])
        self.entry_bankproductchk_baowner.setText(cell_values[8])
        self.entry_bankproductchk_bankid.setText(cell_values[9])
        self.cb_bankproductchk_bkname.setCurrentText(cell_values[10])
        self.entry_bankproductchk_bankaccno.setText(cell_values[11])
        self.entry_bankproductchk_payval.setText(cell_values[12])
        self.entry_bankproductchk_remark.setText(cell_values[13])
    
    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_BankAccProductChkDialog.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
        
    app = QtWidgets.QApplication(sys.argv)
    dialog = BankAccProductChkDialog()
    dialog.show()
    sys.exit(app.exec())