import sys
import logging
from functools import partial
from openpyxl.styles import Font
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QStandardItemModel, QKeySequence
from PyQt5.QtWidgets import QDialog, QMessageBox, QMenu, QShortcut
from PyQt5.QtCore import Qt, QTimer
from datetime import datetime
from commonmd import *
from cal import CalendarView
#<--for non_ui version-->
#from reinfo_ui import Ui_RecyclingDialog 

# Dialog and Import common modules
class RecyclingDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class RecyclingDialog(QDialog, Ui_RecyclingDialog, SubWindowBase):

    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()

        self.conn, self.cursor = connect_to_database4()

        # Initialize a list to store the states and IDs of the last 11 checkboxes
        self.checkbox_states = {"chk_01": False, "chk_02": False, "chk_03": False,
                                "chk_04": False, "chk_05": False, "chk_06": False,
                                "chk_07": False, "chk_08": False, "chk_09": False,
                                "chk_10": False, "chk_11": False}        
        
        #for ui version-------------------------
        uic.loadUi("reinfo.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Keep going date time
        self.update_timer()

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_recycling and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "", "", "", "", "numeric", "", "", "",]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types, self.tv_recycling)
        self.tv_recycling.setItemDelegate(delegate)
        self.tv_recycling.setModel(self.proxy_model)
       
        # Enable sorting
        self.tv_recycling.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_recycling.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_recycling.verticalHeader().setVisible(False)

        # While selecting row in tv_recycling, each cell values to displayed to designated widgets
        self.tv_recycling.clicked.connect(self.show_selected_data)

        # Fill combobox items when the application starts
        self.get_combobox_contents()
        
        # Initiate display of data
        self.make_data()

        # Connect the button to the method
        self.connect_btn_method()
        self.connect_signal_slot()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Create context menus ----------------------------------------------------------------------------------
        self.context_menu1 = self.create_context_menu(self.entry_recycling_efffrom)
        self.context_menu2 = self.create_context_menu(self.entry_recycling_effthru)

        self.entry_recycling_efffrom.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_recycling_efffrom.customContextMenuRequested.connect(self.show_context_menu1)

        self.entry_recycling_effthru.setContextMenuPolicy(Qt.CustomContextMenu)
        self.entry_recycling_effthru.customContextMenuRequested.connect(self.show_context_menu2)
        #--------------------------------------------------------------------------------------------------------

        # Make log file
        self.make_logfiles("access_re_master.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_recycling, partial(self.copy_cells, self.tv_recycling))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_recycling, partial(self.paste_cells, self.tv_recycling))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_recycling, partial(self.handle_return_key, self.tv_recycling))

    # Mouse Right click, show "달력보기" menu ---------------------------------------------------------------------
    def create_context_menu(self, target_lineedit):
        context_menu = QMenu()
        custom_action = context_menu.addAction("달력보기")
        custom_action.triggered.connect(lambda: self.show_calendar(target_lineedit))
        return context_menu

    def show_context_menu1(self, pos):
        self.context_menu1.exec_(self.entry_recycling_efffrom.mapToGlobal(pos))
    def show_context_menu2(self, pos):
        self.context_menu2.exec_(self.entry_recycling_effthru.mapToGlobal(pos))

    # Populate calendar
    def show_calendar(self, target_lineedit):
        calendar_dialog = CalendarView()
        calendar_dialog.selected_date_changed.connect(lambda date: self.set_selected_date(date, target_lineedit))
        calendar_dialog.exec()

    # Display selected date to specific widget
    def set_selected_date(self, date, target_lineedit):
        if target_lineedit == self.entry_recycling_efffrom:
            target_lineedit.setText(date)
        elif target_lineedit == self.entry_recycling_effthru:
            target_lineedit.setText(date)

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_recycling
        self.process_key_event(event, tv_widget)

    # Update timer every one second
    def update_timer(self):
        self.update_timer = QTimer(self) # Set up a QTimer to update the datetime label every second
        self.update_timer.timeout.connect(self.updateDDT)
        self.update_timer.start(1000)  # Update every 1000 milliseconds (1 second)

    # date time update
    def updateDDT(self):
        now = datetime.now()
        curr_date = now.strftime("%Y-%m-%d")
        curr_time = now.strftime("%H:%M:%S")
        
        # Added a space between date and time
        dtime = f"{curr_date} {curr_time}" 
        
        self.lbl_recycling_ddt.setText(dtime)

    # Pass combobox info and sql to next method
    def get_combobox_contents(self):
        self.insert_combobox_initiate(self.cb_recycling_aptname, "SELECT DISTINCT aname FROM vw_re ORDER BY aname")
        self.insert_combobox_initiate(self.cb_recycling_class1, "SELECT DISTINCT class1 FROM vw_re ORDER BY class1")
        self.add_items_to_cb_box(self.cb_recycling_check)

    # Add items to combobox cb_recycling_check
    def add_items_to_cb_box(self, combobox):
        combobox.addItem("")
        combobox.addItem("0")
        combobox.addItem("2")
        combobox.setCurrentIndex(0)

    # Initiate Combo_Box 
    def insert_combobox_initiate(self, combo_box, sql_query):
        self.combobox_initializing(combo_box, sql_query) # using common module
        self.lbl_recycling_aptid.setText("")
        self.entry_recycling_aptcode.setText("")
        self.cb_recycling_aptname.setCurrentIndex(0) 
        self.cb_recycling_class1.setCurrentIndex(0) 

    # Connect button to method
    def connect_btn_method(self):
        self.pb_recycling_aptshow.clicked.connect(self.make_data)
        self.pb_recycling_aptsearch.clicked.connect(self.search_data)
        self.pb_recycling_aptreflect.clicked.connect(self.reflect_data)

        self.pb_recycling_insert.clicked.connect(self.tb_insert)
        self.pb_recycling_update.clicked.connect(self.tb_update)
        self.pb_recycling_update_all.clicked.connect(self.tb_update_all)
        self.pb_recycling_delete.clicked.connect(self.tb_delete)
        self.pb_recycling_clear_data.clicked.connect(self.clear_data)
        self.pb_recycling_close.clicked.connect(self.close_dialog)

    # Connect Signal to Slot
    def connect_signal_slot(self):
        self.connect_checkbox(self.chk_01, self.lbl_01, "chk_01")
        self.connect_checkbox(self.chk_02, self.lbl_02, "chk_02")
        self.connect_checkbox(self.chk_03, self.lbl_03, "chk_03")
        self.connect_checkbox(self.chk_04, self.lbl_04, "chk_04")
        self.connect_checkbox(self.chk_05, self.lbl_05, "chk_05")
        self.connect_checkbox(self.chk_06, self.lbl_06, "chk_06")
        self.connect_checkbox(self.chk_07, self.lbl_07, "chk_07")
        self.connect_checkbox(self.chk_08, self.lbl_08, "chk_08")
        self.connect_checkbox(self.chk_09, self.lbl_09, "chk_09")
        self.connect_checkbox(self.chk_10, self.lbl_10, "chk_10")
        self.connect_checkbox(self.chk_11, self.lbl_11, "chk_11")
        self.cb_recycling_class1.activated.connect(self.cb_recycling_class1_changed)
        self.cb_recycling_aptname.activated.connect(self.cb_recycling_aptname_changed)

    # Check box state change and connected to method
    def connect_checkbox(self, checkbox, label_widget, checkbox_name):
        checkbox.stateChanged.connect(lambda state, name=checkbox_name, checkbox=checkbox, label_widget=label_widget: self.checkboxStateChanged(name, checkbox, label_widget))

    # Define a common function for checkboxStateChanged
    def checkboxStateChanged(self, checkbox_name, checkbox, label_widget):
        if checkbox and label_widget:
            checkbox_status = "2" if checkbox.isChecked() else "0"

        id_text = label_widget.text()
        if id_text != "":
            id_number = int(id_text)
        else:
            id_number = 0
        
        self.checkbox_states[checkbox_name] = checkbox.isChecked()
        info = f"{checkbox_name}: {checkbox_status}, id: {id_number}"
        
        # Check if any checkbox is checked
        if any(self.checkbox_states.values()):
            info = f"{checkbox_name}: {checkbox_status}, id: {id_number}"
            #print(info)

    # Tab order for sub window
    def set_tab_order(self):
        widgets = [self.pb_recycling_aptshow, self.entry_recycling_aptcode, self.cb_recycling_aptname,
            self.cb_recycling_class1, self.entry_recycling_class2, self.cb_recycling_check,
            self.entry_recycling_efffrom, self.entry_recycling_effthru, self.entry_recycling_remark, 
            self.pb_recycling_aptsearch, self.pb_recycling_aptreflect, self.pb_recycling_clear_data, 
            self.pb_recycling_insert, self.pb_recycling_update, self.pb_recycling_delete, 
            self.pb_recycling_close]

        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_recycling
        self.cursor.execute("Select id, acode, aname, class1, class2, check, efffrom, effthru, remark  From vw_re WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]        
        
        sql_query = "Select id, acode, aname, class1, class2, check, efffrom, effthru, remark  From vw_re order by id"
        column_widths = [80, 100, 200, 100, 80, 80, 100, 100, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # Make table data
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Search Data
    def search_data(self):
        self.lbl_recycling_aptid.setText("")
        self.cb_recycling_class1.setCurrentIndex(0)
        self.entry_recycling_class2.setText("")
        self.cb_recycling_check.setCurrentIndex(0)
        self.entry_recycling_remark.setText("")

        aptname = self.cb_recycling_aptname.currentText()
        v01 = len(aptname)

        if v01>0 :
            query = f"SELECT id, acode, aname, class1, class2, check, efffrom, effthru, remark FROM vw_re where aname='{aptname}' order by id"
        else:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 모두 비어 있습니다!")
            return

        QMessageBox.about(self, "검색 조건 확인", f"전체 목록에서  {aptname} \n\n아파트에 대하여 검색을 수행합니다!")

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

        self.reset_checkboxes()
        self.update_checkbox_state()

    # Clear the checked checkbox as unchecked
    def reset_checkboxes(self):
        # Assuming you have a list of checkboxes or a dictionary mapping them
        checkboxes = [self.chk_01, self.chk_02, self.chk_03, self.chk_04, self.chk_05, self.chk_06, self.chk_07, self.chk_08, self.chk_09, self.chk_10, self.chk_11]

        # Reset all checkboxes to unchecked
        for checkbox in checkboxes:
            checkbox.blockSignals(True)   # Disable signals
            checkbox.setChecked(False)    # Reset the checkbox state
            checkbox.blockSignals(False)  # Enable signals again

    # 체크박스 상태 자동 업데이트 및 ID 삽입 --------------------------------------------------------------
    def update_checkbox_state(self):
        # Assuming you have a list of checkboxes in the desired order
        checkboxes = [self.chk_01, self.chk_02, self.chk_03, self.chk_04, self.chk_05, self.chk_06, self.chk_07, self.chk_08, self.chk_09, self.chk_10, self.chk_11]
        
        checkbox_column_index = 5 # the Checkbox status is in the 6th column (index 5)
        id_column_index = 0 # the ID number is in the 0th column (index 0)

        # Loop through the rows in the QTableWidget and update checkbox states
        for row in range(self.tv_recycling.model().rowCount()):
            item_checkbox = self.tv_recycling.model().item(row, checkbox_column_index)
            item_id = self.tv_recycling.model().item(row, id_column_index)

            if item_checkbox is not None and item_id is not None:
                checkbox_status = int(item_checkbox.text())  # Convert the text to an integer
                id_number = item_id.text()

                # Ensure the row index is within the range of checkboxes
                if 0 <= row < len(checkboxes):
                    
                    # Set the ID number into the corresponding label
                    label_index = row  # Assuming row and label indices match
                    label_name = f"lbl_{label_index + 1:02d}"  # Format the label name with leading zeros
                    label_widget = getattr(self, label_name, None)

                    if label_widget is not None:
                        label_widget.setText(id_number)
                    
                    # Check the checkbox if status is 2
                    checkboxes[row].setChecked(checkbox_status == 2)  
    # ---------------------------------------------------------------------------------------------------

    # Check the state of the checkboxes and id numbers to get the info
    def reflect_data(self):

        for checkbox_number in range(1, 12):
            checkbox_name = f"chk_{checkbox_number:02d}"
            checkbox = getattr(self, checkbox_name)
            checkbox_status = "2" if checkbox.isChecked() else "0"
            
            # Get the corresponding label name
            label_name = f"lbl_{checkbox_number:02d}"
            label_widget = getattr(self, label_name)
            
            # Extract the id value from the label text
            if label_widget:
                id_text = label_widget.text()
                if id_text.strip().isdigit():
                    id_number = int(id_text)
                else:
                    id_number = 0  # Set a default value if the label doesn't contain a valid number
            else:
                id_number = 0  # Set a default value if the label widget is not found
            
            self.reflect_chkbox_state_tb_update(id_number, checkbox_status)      

    # Reflect the state of the check box to the db table
    def reflect_chkbox_state_tb_update(self, id_number, checkbox_status):
        
        # Execute the SQL update statement using id_number and checkbox_status
        self.cursor.execute('''UPDATE re_master SET check=? WHERE id=?''', (checkbox_status, id_number))
        self.conn.commit()
        
        # Update tv_recycling display
        aptname = self.cb_recycling_aptname.currentText()
        query = f"SELECT id, acode, aname, class1, class2, check, efffrom, effthru, remark FROM vw_re where aname='{aptname}'"

        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        tv_widget.clearContents()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)

    # Combobox class1 index changed
    def cb_recycling_class1_changed(self):
        self.entry_recycling_class2.clear()
        selected_item = self.cb_recycling_class1.currentText()

        if selected_item:
            query = f"SELECT DISTINCT class2 From vw_re WHERE class1 ='{selected_item}'"
            line_edit_widgets = [self.entry_recycling_class2]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Combobox aptname index changed
    def cb_recycling_aptname_changed(self):
        self.entry_recycling_aptcode.clear()
        selected_item = self.cb_recycling_aptname.currentText()

        if selected_item:
            query = f"SELECT DISTINCT acode From vw_re WHERE aname ='{selected_item}'"
            line_edit_widgets = [self.entry_recycling_aptcode]
        
            # Check if any line edit widgets are provided
            if line_edit_widgets:
                self.lineEdit_contents(line_edit_widgets, query)
            else:
                pass

    # Clear entry input fields
    def clear_data(self):
        self.lbl_recycling_aptid.setText("")
        clear_widget_data(self)

    # Get the value of other variables
    def get_recycling_input(self):
        acode = str(self.entry_recycling_aptcode.text())
        class2 = str(self.entry_recycling_class2.text())
        acheck = int(self.cb_recycling_check.currentText())
        efffrom = str(self.entry_recycling_efffrom.text())
        effthru = str(self.entry_recycling_effthru.text())
        remark = str(self.entry_recycling_remark.text())

        return acode, class2, acheck, efffrom, effthru, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()

        return username, user_id, formatted_datetime
        
    # insert new customer data to MySQL table
    def tb_insert(self):
        confirm_dialog = self.show_insert_confirmation_dialog()
        
        if confirm_dialog == QMessageBox.Yes:      

            idx = self.max_row_id("recycling") 
            username, user_id, formatted_datetime = self.common_values_set()
            acode, class2, acheck, efffrom, effthru, remark = self.get_recycling_input()  # Get the value of other variables

            if (idx>0 and acheck>=0) and all(len(var) > 0 for var in (acode, class2)):

                self.cursor.execute('''INSERT INTO re_master (id, acode, class2, check, efffrom, effthru, trx_date, userid, remark) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'''
                            , (idx, acode, class2, acheck, efffrom, effthru, formatted_datetime, user_id, remark))
                self.conn.commit()
                self.show_insert_success_message()
                self.refresh_data()
                logging.info(f"User {username} inserted {idx} row to the re_master table.")
            else:
                self.show_missing_message("입력 이상")
        else:
            self.show_cancel_message("데이터 추가 취소")

    def clear_update_input(self):
        self.lbl_recycling_aptid.setText("")
        self.cb_recycling_class1.setCurrentIndex(0)
        self.entry_recycling_class2.setText("")
        self.cb_recycling_check.setCurrentIndex(0)
        self.entry_recycling_remark.setText("")
    
    # revise the values in the selected row
    def tb_update(self):
        confirm_dialog = self.show_update_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            
            idx = int(self.lbl_recycling_aptid.text())
            username, user_id, formatted_datetime = self.common_values_set()     
            acode, class2, acheck, efffrom, effthru, remark = self.get_recycling_input()  # Get the value of other variables
            
            if (idx>0 and acheck>=0) and all(len(var) > 0 for var in (acode, class2)):
                self.cursor.execute('''UPDATE re_master SET acode=?, class2=?, check=?, efffrom=?, effthru=?, trx_date=?, userid=?, remark=? WHERE id=?'''
                            , (acode, class2, acheck, efffrom, effthru, formatted_datetime, user_id, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.clear_update_input()
                self.keep_search_data()
                logging.info(f"User {username} updated row number {idx} in the re_master table.")
            else:
                self.show_missing_message("입력 이상")
                pass       
        else:
            self.show_cancel_message("데이터 변경 취소")

    def keep_search_data(self):
        aptname = self.cb_recycling_aptname.currentText()
        v01 = len(aptname)

        if v01>0 :
            query = f"SELECT id, acode, aname, class1, class2, check, efffrom, effthru, remark FROM vw_re where aname='{aptname}' order by id"
        else:
            QMessageBox.about(self, "검색 조건 확인", "검색 조건이 모두 비어 있습니다!")
            return
      
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        tv_widget.clearContents()
        self.populate_dialog(self.cursor, query, tv_widget, column_info, column_names,column_widths)
        self.reset_checkboxes()
        self.update_checkbox_state()

    # revise the values in the selected row
    def tb_update_all(self):
        confirm_dialog = self.show_update_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            
            username, user_id, formatted_datetime = self.common_values_set()     
            acode, class2, acheck, efffrom, effthru, remark = self.get_recycling_input()  # Get the value of other variables
            
            if (acheck>=0) and all(len(var) > 0 for var in (acode, class2, efffrom, effthru)):
                self.cursor.execute('''UPDATE re_master SET efffrom=?, effthru=?, trx_date=?, userid=?, remark=? WHERE acode=? and check=2'''
                            , (efffrom, effthru, formatted_datetime, user_id, remark, acode))
                self.conn.commit()
                self.show_update_success_message()
                self.clear_update_input()
                self.keep_search_data()
                logging.info(f"User {username} updated the effective_from and effective_thru dates for aptcode {acode} in the re_master table.")
            else:
                self.show_missing_message("입력 이상")
                pass       
        else:
            self.show_cancel_message("데이터 변경 취소")

    # delete row according to id selected
    def tb_delete(self):
        confirm_dialog = self.show_delete_confirmation_dialog()
        if confirm_dialog == QMessageBox.Yes:
            idx = self.lbl_recycling_aptid.text()
            username, user_id, formatted_datetime = self.common_values_set()

            self.cursor.execute("DELETE FROM re_master WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()
            logging.info(f"User {username} deleted {idx} row to the re_master table.")
        else:
            self.show_cancel_message("데이터 삭제 취소")

    # table widget cell double click
    def show_selected_data(self, item):
        
        row_index = item.row()  # Get the row index of the clicked item
        cell_values = []    # Initialize a list to store the cell values

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(9):  # have 9 columns
            cell_text = self.tv_recycling.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_recycling_aptid.setText(cell_values[0])
        self.entry_recycling_aptcode.setText(cell_values[1])
        self.cb_recycling_aptname.setCurrentText(cell_values[2])
        self.cb_recycling_class1.setCurrentText(cell_values[3])
        self.entry_recycling_class2.setText(cell_values[4])
        self.cb_recycling_check.setCurrentText(cell_values[5])
        self.entry_recycling_efffrom.setText(cell_values[6])
        self.entry_recycling_effthru.setText(cell_values[7])
        self.entry_recycling_remark.setText(cell_values[8])
    
    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_re_master.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = RecyclingDialog()
    dialog.show()
    sys.exit(app.exec())