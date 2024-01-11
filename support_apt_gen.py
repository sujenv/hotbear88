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
#from support_apt_gen_ui import Ui_SupportAptGenDialog

# Table contents -----------------------------------------------------
class SupportAptGenDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class SupportAptGenDialog(QDialog, Ui_SupportAptGenDialog, SubWindowBase):
     
    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()    
  
        self.conn, self.cursor = connect_to_database3()

        # load ui file
        uic.loadUi("support_apt_gen.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_supportaptgen and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "", "", "", "numeric", "", "", "",]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_supportaptgen)
        self.tv_supportaptgen.setItemDelegate(delegate)
        self.tv_supportaptgen.setModel(self.proxy_model)

        # Enable sorting
        self.tv_supportaptgen.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_supportaptgen.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_supportaptgen.verticalHeader().setVisible(False)

        # While selecting row in tv_supportaptgen, each cell values to displayed to designated widgets
        self.tv_supportaptgen.clicked.connect(self.show_selected_data)

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
        self.context_menu1 = self.create_context_menu(self.entry_supportaptgen_gdate)
        self.context_menu2 = self.create_context_menu(self.entry_supportaptgen_efffrom)
        self.context_menu3 = self.create_context_menu(self.entry_supportaptgen_effthru)
        
        self.entry_supportaptgen_gdate.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_supportaptgen_gdate.customContextMenuRequested.connect(self.show_context_menu1)

        self.entry_supportaptgen_efffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_supportaptgen_efffrom.customContextMenuRequested.connect(self.show_context_menu2)

        self.entry_supportaptgen_effthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_supportaptgen_effthru.customContextMenuRequested.connect(self.show_context_menu3)

        # Make log file
        self.make_logfiles("access_SupportAptGenDialog.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_supportaptgen, partial(self.copy_cells, self.tv_supportaptgen))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_supportaptgen, partial(self.paste_cells, self.tv_supportaptgen))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_supportaptgen, partial(self.handle_return_key, self.tv_supportaptgen))


    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu

    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_supportaptgen_gdate.mapToGlobal(pos))
    def show_context_menu2(self, pos):
        self.context_menu2.exec_(self.entry_supportaptgen_efffrom.mapToGlobal(pos))
    def show_context_menu3(self, pos):
        self.context_menu3.exec_(self.entry_supportaptgen_effthru.mapToGlobal(pos))

    # Show Calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Show selected date to the select Qlineedit
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_supportaptgen_gdate:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_supportaptgen_efffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_supportaptgen_effthru:
            target_lineedit.setText(date)


    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_supportaptgen
        self.process_key_event(event, tv_widget)

    # Display end of date only
    def display_eff_date(self):
        endofdate = "2050/12/31"

        return endofdate
    
    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_supportaptgen_ename, "SELECT DISTINCT ename FROM employee ORDER BY ename")
        self.insert_combobox_initiate(self.cb_supportaptgen_aname, "SELECT DISTINCT aname FROM aptmaster ORDER BY aname")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) 
        self.lbl_supportaptgen_id.setText("")
        self.entry_supportaptgen_ecode.setText("")
        self.cb_supportaptgen_ename.setCurrentIndex(0) 

    # Connect button to method
    def conn_button_to_method(self):
        self.pb_supportaptgen_show.clicked.connect(self.make_data)
        self.pb_supportaptgen_search.clicked.connect(self.search_data)        
        self.pb_supportaptgen_close.clicked.connect(self.close_dialog)
        self.pb_supportaptgen_clear.clicked.connect(self.clear_data)
        self.pb_supportaptgen_export_emp.clicked.connect(self.export_data_to_access_emp)
        
    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_supportaptgen_ename.activated.connect(self.cb_supportaptgen_ename_changed)
        self.cb_supportaptgen_aname.activated.connect(self.cb_supportaptgen_aname_changed)

    # tab order for employee window
    def set_tab_order(self):
       
        widgets = [self.pb_supportaptgen_show, self.entry_supportaptgen_ecode, self.cb_supportaptgen_ename,
            self.entry_supportaptgen_acode, self.entry_supportaptgen_gdate, self.entry_supportaptgen_descr,
            self.entry_supportaptgen_sval, 
            self.entry_supportaptgen_efffrom, self.entry_supportaptgen_effthru, self.entry_supportaptgen_gdate,
            self.entry_supportaptgen_remark, 
            self.pb_supportaptgen_search, self.pb_supportaptgen_clear, self.pb_supportaptgen_close, ]
        
        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_supportaptgen
        self.cursor.execute("SELECT * FROM vw_supportapt WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select * From vw_supportapt Order By ename, aname"
        column_widths = [80, 100, 100, 100, 250, 200, 100, 100, 100, 150]

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
      
        ename = self.cb_supportaptgen_ename.currentText()
        aname = self.cb_supportaptgen_aname.currentText()
        

        conditions = {'v01': (ename, "ename like '%{}%'"),
                      'v02': (aname, "aname like '%{}%'"),
                      }
        
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_supportapt WHERE {' AND '.join(selected_conditions)} ORDER BY ename, aname"

        QMessageBox.about(self, "검색 조건 확인", f"직원명: {ename} \n아파트명: {aname} \n\n위 조건으로 검색을 수행합니다!")
        
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)


    # Employee Name Index Changed
    def cb_supportaptgen_ename_changed(self):
        self.entry_supportaptgen_ecode.clear()
        selected_item = self.cb_supportaptgen_ename.currentText()

        if selected_item:
            query = f"SELECT DISTINCT ecode From employee WHERE ename ='{selected_item}'"
            line_edit_widgets = [self.entry_supportaptgen_ecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Apt Name Index Changed
    def cb_supportaptgen_aname_changed(self):
        self.entry_supportaptgen_acode.clear()
        selected_item = self.cb_supportaptgen_aname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode From aptmaster WHERE aname ='{selected_item}'"
            line_edit_widgets = [self.entry_supportaptgen_acode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    def export_data_to_access_emp(self):

        paydt = self.entry_supportaptgen_gdate.text()
        srtdt = self.entry_supportaptgen_efffrom.text()
        enddt = self.entry_supportaptgen_effthru.text()

        if not paydt or not srtdt or not enddt:
            # At least one of the date values is empty
            empty_fields = []
            if not paydt:
                empty_fields.append("정산예정일")
            if not srtdt:
                empty_fields.append("정산시작일")
            if not enddt:
                empty_fields.append("정산종료일")

            # Display a warning message with the names of the empty fields
            warning_message = f"다음 날짜 필드가 채워져 있는지 확인해 주세요!: \n\n {', '.join(empty_fields)}"
            QMessageBox.warning(self, "주의", warning_message)

        else:
            
            self.search_data()

            result = QMessageBox.question(
                self,
                "월별 개인별 아파트 지원금 데이터 추가",
                f"정산시작일: {srtdt} \n정산종료일: {enddt} \n지급예정일: {paydt}\n\n"
                "위 기준으로 개인별 아파트 지원금 정산 데이터를 추가하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            # User clicked Yes, continue with the code
            if result == QMessageBox.Yes:
        
                dt1 = len(paydt)

                if dt1>0:

                    # set the name of tableview widget
                    table_widget = self.tv_supportaptgen

                    # Get the model associated with the QTableView
                    model = table_widget.model()
                    
                    # Check if there is data to insert
                    if model.rowCount() == 0:
                        QMessageBox.about(self, "데이터 확인", "테이블뷰에 추가할 데이터가 없습니다!")
                        return

                    # Extract data from the model and prepare for insertion
                    data_to_insert = []
                    
                    for row in range(model.rowCount()):
                        # Extracting data for each column in a row
                        id = model.item(row, 0).text()
                        ecode = model.item(row, 1).text()
                        ename = model.item(row, 2).text()
                        acode = model.item(row, 3).text()
                        aname = model.item(row, 4).text()
                        descr = model.item(row, 5).text()
                        sval = model.item(row, 6).text()
                        efffrom = model.item(row, 7).text()
                        effthru = model.item(row, 8).text()
                        remark = model.item(row, 9).text()

                        # Create a tuple with the extracted values
                        data = (ecode, sval, srtdt, enddt, paydt, remark)
                        
                        # Append the tuple to the data_to_insert list
                        data_to_insert.append(data)

                    try:
                        # Set up the connection to the MS Access database
                        relative_dbs_folder = 'dbs'
                        db_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
                        filename = 'payroll.accdb'
                        db_path = os.path.join(relative_dbs_folder, filename)
                        conn = pyodbc.connect(rf'DRIVER={db_driver};' rf'DBQ={db_path};')
                        cursor = conn.cursor()        

                        # Insert data into the MonthlySalary table using executemany
                        cursor.executemany("""
                            INSERT INTO supportaptmonth (ecode, sval, efffrom, effthru, paydt, remark)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, data_to_insert)
                    
                        # Commit the changes and close the connection
                        conn.commit()
                        conn.close()

                        QMessageBox.about(self, "데이터 추가", "데이터가 성공적으로 추가되었습니다!")

                    except Exception as e:
                        QMessageBox.about(self, "에러 발생", f"데이터 추가 중 에러가 발생했습니다: {str(e)}")
                else:
                    QMessageBox.critical(self, "필수 데이터 입력 안됨!", "지급 예정일을 확인해 주세요!")
            
            else:
                # User clicked No, do nothing or handle as needed
                return
            

    # clear input field entry
    def clear_data(self):
        self.lbl_supportaptgen_id.setText("")
        for line_edit in self.findChildren(QtWidgets.QLineEdit):
            line_edit.clear()

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(10):  # 10columns
            cell_text = self.tv_supportaptgen.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_supportaptgen_id.setText(cell_values[0])
        self.entry_supportaptgen_ecode.setText(cell_values[1])
        self.cb_supportaptgen_ename.setCurrentText(cell_values[2])
        self.entry_supportaptgen_acode.setText(cell_values[3])
        self.cb_supportaptgen_aname.setCurrentText(cell_values[4])
        self.entry_supportaptgen_descr.setText(cell_values[5])
        self.entry_supportaptgen_sval.setText(cell_values[6])
        self.entry_supportaptgen_efffrom.setText(cell_values[7])
        self.entry_supportaptgen_effthru.setText(cell_values[8])
        self.entry_supportaptgen_remark.setText(cell_values[9])
    
    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_SupportAptGenDialog.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
        
    app = QtWidgets.QApplication(sys.argv)
    dialog = SupportAptGenDialog()
    dialog.show()
    sys.exit(app.exec())