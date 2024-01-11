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
#from delivery_ui import Ui_DeliveryDialog

# Delivery Type List-----------------------------------------------------------------
class DeliveryDialog(QDialog, SubWindowBase):
#for non_ui version-------------------------
#class DeliveryDialog(QDialog, Ui_DeliveryDialog, SubWindowBase):

    def __init__(self, current_username = None, current_datetime = None):
        super().__init__()  

        self.conn, self.cursor = connect_to_database1()

        # Load ui file          
        uic.loadUi("delivery.ui", self)
        #for non_ui version-------------------------
        #self.setupUi(self)

        # Add window flags
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)  

        # Initialize current_username and current_datetime directly
        self.current_username, self.current_datetime = initialize_username_and_datetime(current_username, current_datetime)

        # Enable automatic deletion on close
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)  

        # Create tv_delivery and QSortFilterProxyModel
        self.model = QStandardItemModel()
        self.proxy_model = NumericStringSortModel(self.model)
        self.proxy_model.setSourceModel(self.model)

        # Define the column types
        column_types = ["numeric", "numeric", "", ""]

        # Set the custom delegate for the specific column
        delegate = NumericDelegate(column_types,self.tv_delivery)
        self.tv_delivery.setItemDelegate(delegate)
        self.tv_delivery.setModel(self.proxy_model)

        # Enable sorting
        self.tv_delivery.setSortingEnabled(True)

        # Enable alternating row colors
        self.tv_delivery.setAlternatingRowColors(True)  
        
        # Hide the first index column
        self.tv_delivery.verticalHeader().setVisible(False)

        # While selecting row in tv_delivery, each cell values to displayed to designated widgets
        self.tv_delivery.clicked.connect(self.show_selected_data)

        # Initiate display of data
        self.make_data()
        self.conn_button_to_method()

        # Set tab order for input widgets
        self.set_tab_order()

        # Initiate CTRL+C, CTRL+V and ENTER
        self.setup_shortcuts()

        # Make log file
        self.make_logfiles("access_delivery.log")

    # Create a shortcut for CTRL+C/CTRL+V/ Return key
    def setup_shortcuts(self):
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.tv_delivery, partial(self.copy_cells, self.tv_delivery))
        self.paste_shortcut = QShortcut(QKeySequence.Paste, self.tv_delivery, partial(self.paste_cells, self.tv_delivery))
        self.return_shortcut = QShortcut(Qt.Key_Return, self.tv_delivery, partial(self.handle_return_key, self.tv_delivery))

    # Call process_key_event and pass the event and your QTableWidget instance
    def keyPressEvent(self, event):
        tv_widget = self.tv_delivery
        self.process_key_event(event, tv_widget)

    # Connect the button to the method
    def conn_button_to_method(self):
        self.pb_delivery_show.clicked.connect(self.make_data)
        self.pb_delivery_clear.clicked.connect(self.clear_data)        
        self.pb_delivery_close.clicked.connect(self.close_dialog)
        
        self.pb_delivery_insert.clicked.connect(self.tb_insert)
        self.pb_delivery_update.clicked.connect(self.tb_update)
        self.pb_delivery_delete.clicked.connect(self.tb_delete)

    # Tab order for sub window
    def set_tab_order(self):
        widgets = [self.pb_delivery_show, self.entry_delivery_code, self.entry_delivery_name
                , self.entry_delivery_remark, self.pb_delivery_clear, self.pb_delivery_insert
                , self.pb_delivery_update, self.pb_delivery_delete]

        for i in range(len(widgets) - 1):
            self.setTabOrder(widgets[i], widgets[i + 1])

    # To reduce duplications
    def common_query_statement(self):
        tv_widget = self.tv_delivery
        self.cursor.execute("Select id, dcode, dname, remark From delivery WHERE 1=0")
        column_info = self.cursor.description
        column_names = [col[0] for col in column_info]

        sql_query = "Select id, dcode, dname, remark From delivery order by id"
        column_widths = [80, 100, 100, 150]

        return sql_query, tv_widget, column_info, column_names, column_widths 

    # show delivery type tablewidget
    def make_data(self):
        sql_query, tv_widget, column_info, column_names, column_widths = self.common_query_statement() 
        self.populate_dialog(self.cursor, sql_query, tv_widget, column_info, column_names,column_widths)

    # Get the value of other variables
    def get_delivery_input(self):
        dcode = int(self.entry_delivery_code.text())
        dname = str(self.entry_delivery_name.text())
        remark = str(self.entry_delivery_remark.text())
        return dcode, dname, remark

    # Make Common values set
    def common_values_set(self):
        username = self.current_username
        user_id = self.userID_gen(username)
        formatted_datetime = self.dt_time_info()
        return username, user_id, formatted_datetime

    # insert new data to MySQL table
    def tb_insert(self):
        currentid = self.lbl_delivery_id.text()
        if not currentid:
            confirm_dialog = self.show_insert_confirmation_dialog()
            if confirm_dialog == QMessageBox.Yes:
                idx = self.max_row_id("delivery")
                username, user_id, formatted_datetime = self.common_values_set()
                dcode, dname, remark = self.get_delivery_input() 

                if (idx>0) and all(len(var) > 0 for var in (dname)):
                    self.cursor.execute('''INSERT INTO delivery (id, dcode, dname, trxdate, userid, remark) 
                                VALUES (?, ?, ?, ?, ?, ?, ?)'''
                                , (idx, dcode, dname, formatted_datetime, user_id, remark))
                    self.conn.commit()
                    self.show_insert_success_message()
                    self.refresh_data() 
                    logging.info(f"User {username} inserted {idx} row to the delivery table.")
                else:
                    self.show_missing_message("입력 이상")
                    return
            else:
                self.show_cancel_message("데이터 추가 취소")
                return    
        else:
            QMessageBox.information(self, "Input Error", "기존 id가 선택된 상태에서는 신규 입력이 불가합니다!")
            return
        
    # update values in the selected row
    def tb_update(self):
        confirm_dialog = self.show_update_confirmation_dialog()

        if confirm_dialog == QMessageBox.Yes:
            idx = int(self.lbl_delivery_id.text())
            username, user_id, formatted_datetime = self.common_values_set()
            dcode, dname, remark = self.get_delivery_input() 

            if (idx>0) and all(len(var) > 0 for var in (dname)):
                self.cursor.execute('''UPDATE delivery SET 
                            dcode=?, dname=?, trxdate=?, userid=?, remark=? WHERE id=?'''
                            , (dcode, dname, formatted_datetime, user_id, remark, idx))
                self.conn.commit()
                self.show_update_success_message()
                self.refresh_data()
                logging.info(f"User {username} updated row number {idx} in the delivery table.")
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
            idx = self.self.lbl_delivery_id.text().text()
            username, user_id, formatted_datetime = self.common_values_set()
            self.cursor.execute("DELETE FROM delivery WHERE id=?", (idx,))
            self.conn.commit()
            self.show_delete_success_message()
            self.refresh_data()
            logging.info(f"User {username} deleted {idx} row to the delivery table.")
        else:
            self.show_cancel_message("데이터 삭제 취소")
            return

    # clear input field entry
    def clear_data(self):
        self.lbl_delivery_id.setText("")
        clear_widget_data(self)

    # table widget cell double click
    def show_selected_data(self, item):
        # Get the row index of the clicked item
        row_index = item.row()

        # Initialize a list to store the cell values
        cell_values = []

        # Loop through the columns and retrieve the text from each cell
        for column_index in range(4):  # 4 columns
            cell_text = self.tv_delivery.model().item(row_index, column_index).text()
            cell_values.append(cell_text)

        # Populate the input widgets with the data from the selected row
        self.lbl_delivery_id.setText(cell_values[0])
        self.entry_delivery_code.setText(cell_values[1])
        self.entry_delivery_name.setText(cell_values[2])
        self.entry_delivery_remark.setText(cell_values[3])

    def refresh_data(self):
        self.clear_data()
        self.make_data()


if __name__ == "__main__":

    log_subfolder = "logs"
    os.makedirs(log_subfolder, exist_ok=True)
    log_file_path = os.path.join(log_subfolder, "access_delivery.log")

    logging.basicConfig(
        filename=log_file_path,  
        level=logging.INFO,    
        format="%(asctime)s [%(levelname)s] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    app = QtWidgets.QApplication(sys.argv)
    dialog = DeliveryDialog()
    dialog.show()
    sys.exit(app.exec())