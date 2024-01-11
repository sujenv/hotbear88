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
#<--for non_ui version-->
#from aptmaster_ui import UI_AptMasterDialog

# customer table contents -----------------------------------------------------
class AptMasterDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class AptMasterDialog(QDialog, UI_AptMasterDialog, SubWindowBase):

    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database4()

        # Load ui file
        uic.loadUi("aptmaster.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_aptmaster and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", "numeric", "", "", "numeric", "", ""]
        
        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_aptmaster)
        self.tv_aptmaster.setItemDelegate(delegate)
        self.tv_aptmaster.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_aptmaster.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_aptmaster.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_aptmaster.verticalHeader().setVisible(False)

        # While selecting row in tv_aptmaster, each cell values to displayed to designated widgets
        self.tv_aptmaster.clicked.connect(self.show_selected_data)

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
        self.make_logfiles("access_AptMaster.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_aptmaster, partial(self.copy_cells, self.tv_aptmaster))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_aptmaster, partial(self.paste_cells, self.tv_aptmaster))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_aptmaster, partial(self.handle_return_key, self.tv_aptmaster))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_aptmaster
        self.process_key_event(event, tv_widget)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_aptmaster_adesc, "SELECT DISTINCT adesc FROM apt_master ORDER BY adesc")
        self.insert_combobox_initiate(self.cb_aptmaster_typedesc, "SELECT DISTINCT description FROM apt_type ORDER BY description")
        self.insert_combobox_initiate(self.cb_aptmaster_cicdesc, "SELECT DISTINCT cdesc FROM apt_cic_master ORDER BY cdesc")

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) # using common module
        
        self.lbl_aptmaster_id.setText("")
        self.entry_aptmaster_code.setText("")
        self.cb_aptmaster_adesc.setCurrentIndex(0) 
        self.cb_aptmaster_typedesc.setCurrentIndex(0) 
        self.cb_aptmaster_cicdesc.setCurrentIndex(0)

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_aptmaster_show.clicked.connect(self.make_data)
        self.pb_aptmaster_search.clicked.connect(self.search_data)
        self.pb_aptmaster_clear_data.clicked.connect(self.clear_data)
        self.pb_aptmaster_close.clicked.connect(self.close_dialog)

        self.pb_aptmaster_insert.clicked.connect(self.tb_insert)
        self.pb_aptmaster_update.clicked.connect(self.tb_update)
        self.pb_aptmaster_delete.clicked.connect(self.tb_delete)
 
     # Connect Signal to Slot
    def connect_signal_slot(self):
        self.cb_aptmaster_adesc.activated.connect(self.cb_aptmaster_adesc_changed)        
        self.cb_aptmaster_typedesc.activated.connect(self.cb_aptmaster_typedesc_changed)
        self.cb_aptmaster_cicdesc.activated.connect(self.cb_aptmaster_cicdesc_changed)

    # Tab order for sub window
    def set_tab_order(self):
        widgets = [self.pb_aptmaster_show, self.entry_aptmaster_code, self.cb_aptmaster_adesc,
            self.entry_aptmaster_noh, self.entry_aptmaster_typecode, self.cb_aptmaster_typedesc,
            self.entry_aptmaster_ciccode, self.cb_aptmaster_cicdesc, self.entry_aptmaster_remark,
            self.pb_aptmaster_search, self.pb_aptmaster_clear_data, self.pb_aptmaster_close,
            self.pb_aptmaster_insert, self.pb_aptmaster_update , self.pb_aptmaster_delete]

        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_aptmaster
        self.cursor.execute("Select id, acode, adesc, nohousehold, typeofapt, description, cic, cdesc, remark  From vw_apt_master WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select id, acode, adesc, nohousehold, typeofapt, description, cic, cdesc, remark  From vw_apt_master order by id"
        column_widths = [80, 100, 250, 100, 80, 80, 100, 100, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # Make table data
    def make_data(self): 
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_aptmaster_input(self):
        kaptcode = str(self.entry_aptmaster_code.text())
        aname = str(self.cb_aptmaster_adesc.currentText())
        noh = int(self.entry_aptmaster_noh.text())
        tycode = int(self.entry_aptmaster_typecode.text())
        ciccode = int(self.entry_aptmaster_ciccode.text()) 
        remark = str(self.entry_aptmaster_remark.text())

        return kaptcode, aname, noh, tycode, ciccode, remark

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
            idx = self.max_row_id("aptmaster") # Get the max id
            username, user_id, formatted_datetime = self.common_values_set()
            kaptcode, aname, noh, tycode, ciccode, remark = self.get_aptmaster_input()  # Get the value of other variables

            if (idx>0 and noh>=0 and tycode and ciccode) and all(len(var) > 0 for var in (kaptcode)):
                self.cursor.execute('''INSERT INTO apt_master (id, acode, adesc, nohousehold, typeofapt, cic, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, kaptcode, aname, noh, tycode, ciccode, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data()
                logging.info(f"User {username} inserted {idx} row to the apt master table.")
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
            idx = int(self.lbl_aptmaster_id.text())
            username, user_id, formatted_datetime = self.common_values_set()     
            kaptcode, aname, noh, tycode, ciccode, remark = self.get_aptmaster_input()  # Get the value of other variables

            if (idx>0 and noh>=0 and tycode and ciccode) and all(len(var) > 0 for var in (kaptcode)):
                self.cursor.execute('''UPDATE apt_master SET 
                            acode=?, adesc=?, nohousehold=?, typeofapt=?, cic=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                            , (kaptcode, aname, noh, tycode, ciccode, formatted_datetime, user_id, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()
                logging.info(f"User {username} updated row number {idx} in the apt master table.")
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
            idx = self.lbl_aptmaster_id.text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM apt_master WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()
            logging.info(f"User {username} deleted {idx} row to the apt master table.")
        else:
            self.show_cancel_message("데이터 삭제 취소")
            return
        
    # Search data
    def search_data(self):

        aptname = self.cb_aptmaster_adesc.currentText()
        apttype = self.cb_aptmaster_typedesc.currentText()
        cname= self.cb_aptmaster_cicdesc.currentText()

        conditions = {'v01': (aptname, "adesc like '%{}%'"), 'v02': (apttype, "description='{}'"), 'v03': (cname, "cdesc='{}'"),}
        selected_conditions = []

        for key, (value, condition_format) in conditions.items():
            if len(value) > 0:
                selected_conditions.append(condition_format.format(value))

        if not selected_conditions:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 비어 있습니다!")
            return

        # Join the selected conditions to form the SQL query
        query = f"SELECT * FROM vw_apt_master WHERE {' AND '.join(selected_conditions)} ORDER BY adesc"

        QMessageBox.about(self, "검색 조건 확인", f"아파트명: {aptname} \n타입: {apttype} \n수제담당회사명: {cname} \n\n위 조건으로 검색을 수행합니다!")

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Combobox apt name index changed
    def cb_aptmaster_adesc_changed(self):
        self.entry_aptmaster_code.clear()
        selected_item = self.cb_aptmaster_adesc.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode From apt_master WHERE adesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptmaster_code]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox customer index changed
    def cb_aptmaster_typedesc_changed(self):
        self.entry_aptmaster_typecode.clear()
        selected_item = self.cb_aptmaster_typedesc.currentText()

        if selected_item:
            query = f"SELECT DISTINCT code From apt_type WHERE description ='{selected_item}'"
            line_edit_widgets = [self.entry_aptmaster_typecode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox suje index changed
    def cb_aptmaster_cicdesc_changed(self):
        self.entry_aptmaster_ciccode.clear()
        selected_item = self.cb_aptmaster_cicdesc.currentText()

        if selected_item:
            query = f"SELECT DISTINCT code From apt_cic_master WHERE cdesc ='{selected_item}'"
            line_edit_widgets = [self.entry_aptmaster_ciccode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass


    # clear input field entry
    def clear_data(self):
        self.lbl_aptmaster_id.setText("")
        clear_widget_data(self)

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(9):  # Assuming you have 9 columns
            cell_text = self.tv_aptmaster.model().item(row_index, column_index).text()
            cell_values.append(cell_text)
            
        # Populate the input widgets with the data from the selected row
        self.lbl_aptmaster_id.setText(cell_values[0])
        self.entry_aptmaster_code.setText(cell_values[1])
        self.cb_aptmaster_adesc.setCurrentText(cell_values[2])
        self.entry_aptmaster_noh.setText(cell_values[3])
        self.entry_aptmaster_typecode.setText(cell_values[4])
        self.cb_aptmaster_typedesc.setCurrentText(cell_values[5])
        self.entry_aptmaster_ciccode.setText(cell_values[6])
        self.cb_aptmaster_cicdesc.setCurrentText(cell_values[7])
        self.entry_aptmaster_remark.setText(cell_values[8])

    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_AptMaster.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    
    app = QtWidgets.QApplication(sys.argv)
    dialog = AptMasterDialog()
    dialog.show()
    sys.exit(app.exec())