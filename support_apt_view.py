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
#from oil_paymentemp_view_ui import Ui_OilUsageEmpViewDialog

# Table contents -----------------------------------------------------
class SupportAptEmpViewDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class OilUsageEmpViewDialog(QDialog, Ui_OilUsageEmpViewDialog, SubWindowBase):
     
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()    
  
        self.conn, self.cursor = connect_to_database3()

        # load ui file
        uic.loadUi("support_apt_view.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_supportaptmonth and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types1 = ["numeric", "numeric", "", "", "", "", "numeric", "", "", "",]
        column_types2 = ["numeric", "", "numeric", "", "", "", "",] 
        
        # Set the custom delegate for the specific column
        delegate1 = NumericDelegate(column_types1, self.tv_supportapt)
        self.tv_supportapt.setItemDelegate(delegate1)
        self.tv_supportapt.setModel(self.proxy_model)

        delegate2 = NumericDelegate(column_types2, self.tv_supportaptmonth)
        self.tv_supportaptmonth.setItemDelegate(delegate2)
        self.tv_supportaptmonth.setModel(self.proxy_model)

        # Enable sorting
        self.tv_supportapt.setSortingEnabled(True)
        self.tv_supportaptmonth.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_supportapt.setAlternatingRowColors(True)  
        self.tv_supportaptmonth.setAlternatingRowColors(True)  

        # Hide the first index column
        self.tv_supportapt.verticalHeader().setVisible(False)
        self.tv_supportaptmonth.verticalHeader().setVisible(False)


        # While selecting row in tv_supportaptmonth, each cell values to displayed to designated widgets
        self.tv_supportapt.clicked.connect(self.show_selected_data_1)
        self.tv_supportaptmonth.clicked.connect(self.show_selected_data_2)


        # Fill combobox items when the application starts
        self.get_combobox_contents()

        # Initiate display of data
        self.make_data1()
        self.make_data2()
        self.conn_button_to_method()
        self.connect_signal_slot()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Create context menus ----------------------------------------------------------------------------------
        self.context_menu1 = self.create_context_menu(self.entry_supportaptview_paydt)

        self.entry_supportaptview_paydt.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_supportaptview_paydt.customContextMenuRequested.connect(self.show_context_menu1)

        # Make log file
        self.make_logfiles("access_SupportAptEmpViewDialog.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_supportaptmonth, partial(self.copy_cells, self.tv_supportaptmonth))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_supportaptmonth, partial(self.paste_cells, self.tv_supportaptmonth))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_supportaptmonth, partial(self.handle_return_key, self.tv_supportaptmonth))

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu
    
    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_supportaptview_paydt.mapToGlobal(pos))
 
     # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_supportaptview_paydt:
            target_lineedit.setText(date)

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_supportaptmonth
        self.process_key_event(event, tv_widget)

    
    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_supportaptview_ename, "SELECT DISTINCT ename FROM employee ORDER BY ename")
        self.insert_combobox_initiate(self.cb_supportaptview_aname, "SELECT DISTINCT aname FROM aptmaster ORDER BY aname")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.entry_supportaptview_ecode.setText("")
        self.cb_supportaptview_ename.setCurrentIndex(0) 
        self.cb_supportaptview_aname.setCurrentIndex(0)         

    # Connect button to method
    def conn_button_to_method(self):
        self.pb_supportaptview_show.clicked.connect(self.make_data)
        self.pb_supportaptview_search.clicked.connect(self.search_data)
        self.pb_supportaptview_close.clicked.connect(self.close_dialog)
        self.pb_supportaptview_clear.clicked.connect(self.clear_data)
        self.pb_supportaptview_delete.clicked.connect(self.tb_delete)
        
    def make_data(self):
        self.make_data1()
        self.make_data2()
        
    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_supportaptview_ename.activated.connect(self.cb_supportaptview_ename_changed)

    # tab order for employee window
    def set_tab_order(self):
       
        widgets = [self.pb_supportaptview_show, self.entry_supportaptview_ecode, self.cb_supportaptview_ename,
            self.entry_supportaptview_payval, self.entry_supportaptview_paydt, self.entry_supportaptview_srtdt,
            self.pb_supportaptview_search, self.pb_supportaptview_clear, self.pb_supportaptview_close, ]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement1(self):
        tv_widget = self.tv_supportapt
        self.cursor.execute("SELECT * FROM vw_supportapt WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_supportapt Order By ename"
        column_widths = [80, 100, 100, 100, 200, 150, 100, 100, 100, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show employee table data inside of the MDI
    def make_data1(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement1() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)


    # To reduce duplications
    def common_query_statement2(self):
        tv_widget = self.tv_supportaptmonth
        self.cursor.execute("SELECT * FROM vw_supportaptmonth WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_supportaptmonth Order By paydt, ename"
        column_widths = [100, 100, 100, 100, 100, 100, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show employee table data inside of the MDI
    def make_data2(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement2() 
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
            
            paydt = self.entry_supportaptview_paydt.text()
            username, user_id, formatted_datetime = self.common_values_set()

            if len(paydt) > 0: 
                self.cursor.execute("DELETE FROM supportaptmonth WHERE paydt=?", (paydt,))
                self.conn.commit()
                self.show_delete_success_message()
                self.refresh_data()  
                logging.info(f"User {username} deleted paydt {paydt}, at the support apt table.")
            else:
                return    
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # Search data
    def search_data(self):
      
        cname = self.cb_supportaptview_ename.currentText()
        paydt = self.entry_supportaptview_paydt.text()

        conditions = {'v01': (cname, "ename like '%{}%'"),
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
        query = f"SELECT * FROM vw_supportaptmonth WHERE {' AND '.join(selected_conditions)} ORDER BY paydt"

        QMessageBox.about(self, "검색 조건 확인", f"지급예정일: {paydt} \n이름: {cname} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement2()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)


    # Employee Name Index Changed
    def cb_supportaptview_ename_changed(self):
        self.entry_supportaptview_ecode.clear()
        selected_item = self.cb_supportaptview_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_supportaptview_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Employee Name Index Changed
    def cb_supportaptview_aname_changed(self):
        self.entry_supportaptview_acode.clear()
        selected_item = self.cb_supportaptview_aname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode From aptmaster WHERE aname ='{selected_item}'"
            line_edit_widgets = [self.entry_supportaptview_acode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # clear input field entry
    def clear_data(self):
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()


    def show_selected_data_1(self, item):

        self.clear_data()
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(10):  # 10columns
            cell_text = self.tv_supportapt.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_supportaptview_id.setText(cell_values[0])
        self.entry_supportaptview_ecode.setText(cell_values[1])
        self.cb_supportaptview_ename.setCurrentText(cell_values[2])
        self.entry_supportaptview_acode.setText(cell_values[3])
        self.cb_supportaptview_aname.setCurrentText(cell_values[4]) 
        self.entry_supportaptview_descr.setText(cell_values[5])  
        self.entry_supportaptview_payval.setText(cell_values[6])        
        self.entry_supportaptview_srtdt.setText(cell_values[7])
        self.entry_supportaptview_enddt.setText(cell_values[8])
        self.entry_supportaptview_remark.setText(cell_values[9])
        

    # table widget cell double click
    def show_selected_data_2(self, item):
        self.clear_data()
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(7):  # 7columns
            cell_text = self.tv_supportaptmonth.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.entry_supportaptview_ecode.setText(cell_values[0])
        self.cb_supportaptview_ename.setCurrentText(cell_values[1])
        self.entry_supportaptview_payval.setText(cell_values[2])
        self.entry_supportaptview_srtdt.setText(cell_values[3])
        self.entry_supportaptview_enddt.setText(cell_values[4])
        self.entry_supportaptview_paydt.setText(cell_values[5])
        self.entry_supportaptview_remark.setText(cell_values[6])


    def refresh_data(self):
        self.clear_data()
        self.make_data1()
        self.make_data2()

if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_SupportAptEmpViewDialog.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
        
    app = QtWidgets.QApplication(sys.argv)
    dialog = SupportAptEmpViewDialog()
    dialog.show()
    sys.exit(app.exec())