import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtGui import QTextCharFormat, QColor
from PyQt5.QtCore import QDate, QDateTime, pyqtSignal
from lunardate import LunarDate

class CalendarView(QDialog):

    selected_date_changed = pyqtSignal(str)  # Signal to emit selected date
        
    def __init__(self, sender=None):
        super().__init__()
        uic.loadUi("cal.ui", self)
        self.sender = sender

        self.korean_holidays = {}       
        self.events = {}
        
        # Set the color of holidays to red for the current year
        self.set_holiday_colors()
       
        self.calendarWidget.selectionChanged.connect(self.on_calendar_selection_changed)
        self.calendarWidget.currentPageChanged.connect(self.on_calendar_page_changed)

        self.pb_cal_close.clicked.connect(self.close_cal)

        # Display the current date and time when initializing
        current_datetime = QDateTime.currentDateTime()
        self.label.setText(f"Selected Date: {current_datetime.toString('yyyy/MM/dd')}\n")

        # Emit the signal with the current date and time
        self.selected_date_changed.emit(current_datetime.toString("yyyy/MM/dd"))

        #self.resize(400, 300)
        self.show()

    def close_cal(self):
         self.close()  

    def set_holiday_colors(self):
        red_text_format = QTextCharFormat()
        red_text_format.setForeground(QColor("red"))

        # Clear existing holiday colors
        for date_str in self.korean_holidays.keys():
            holiday_date = QDate.fromString(date_str, "yyyy/MM/dd")
            self.calendarWidget.setDateTextFormat(holiday_date, QTextCharFormat())

        # Get the selected year from the calendar's current page
        year = self.calendarWidget.yearShown()

        # Update holidays based on the selected year
        self.add_kor_holidays(year)
        self.add_lunar_holidays(year)

        # Apply red text color to the updated holidays
        for date_str in self.korean_holidays.keys():
            holiday_date = QDate.fromString(date_str, "yyyy/MM/dd")
            self.calendarWidget.setDateTextFormat(holiday_date, red_text_format)

    def on_calendar_page_changed(self):
        # Update holiday colors for the selected year
        self.set_holiday_colors()

    def on_calendar_selection_changed(self):
        selected_date = self.calendarWidget.selectedDate()
        date_str = selected_date.toString("yyyy/MM/dd")
        korean_holiday = self.korean_holidays.get(date_str, "")

        self.label.setText(f"Selected Date: {date_str}\n"
                           f"Korean Holiday: {korean_holiday}\n")
        
        # Emit the signal with the selected date
        self.selected_date_changed.emit(date_str)

    def add_kor_holidays(self, year):
        newyear_date = (year, 1, 1)
        indmov_date = (year, 3, 1)
        labor_date = (year, 5, 1)
        childrens_date = (year, 5, 5)
        memorial_date = (year, 6, 6)
        independence_date = (year, 8, 15)
        foundation_date = (year, 10, 3)
        hangul_date = (year, 10, 9)
        christmas_date = (year, 12, 25)

        # Convert the tuple to string format
        newyear_date_str = QDate(*newyear_date).toString("yyyy/MM/dd")
        indmov_date_str = QDate(*indmov_date).toString("yyyy/MM/dd")
        labor_date_str = QDate(*labor_date).toString("yyyy/MM/dd")
        childrens_date_str = QDate(*childrens_date).toString("yyyy/MM/dd")
        memorial_date_str = QDate(*memorial_date).toString("yyyy/MM/dd")
        independence_date_str = QDate(*independence_date).toString("yyyy/MM/dd")
        foundation_date_str = QDate(*foundation_date).toString("yyyy/MM/dd")
        hangul_date_str = QDate(*hangul_date).toString("yyyy/MM/dd")
        christmas_date_str = QDate(*christmas_date).toString("yyyy/MM/dd")

        # Add Korean holiday to the dictionary
        self.korean_holidays[newyear_date_str] = "신정 (Korean New Year)"
        self.korean_holidays[indmov_date_str] = "삼일절 (Independence Movement Day)"
        self.korean_holidays[labor_date_str] = "근로자의날 Labor Day"
        self.korean_holidays[childrens_date_str] = "어린이날 (Children's Day)"
        self.korean_holidays[memorial_date_str] = "현충일 (Memorial Day)"
        self.korean_holidays[independence_date_str] = "광복절 (Independence Day)"
        self.korean_holidays[foundation_date_str] = "개천절 (National Foundation Day)"
        self.korean_holidays[hangul_date_str] = "한글날 (Hangul Day)"
        self.korean_holidays[christmas_date_str] = "성탄절 (Christmas Day)"

    def add_lunar_holidays(self, year):

        seollal_date = LunarDate(year, 1, 1).toSolarDate()
        buddha_date = LunarDate(year, 4, 8).toSolarDate()
        chuseok_date = LunarDate(year, 8, 15).toSolarDate()

        # Convert lunar dates to string format
        seollal_date_str = seollal_date.strftime("%Y-%m-%d")
        buddha_date_str = buddha_date.strftime("%Y-%m-%d")
        chuseok_date_str = chuseok_date.strftime("%Y-%m-%d")

        # Add Korean lunar variable holidays to the dictionary
        self.korean_holidays[seollal_date_str] = "설날 (Seollal)"
        self.korean_holidays[buddha_date_str] = "부처님오신날 (Buddha's Coming)"
        self.korean_holidays[chuseok_date_str] = "추석 (Chuseok)"

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = CalendarView()
    sys.exit(app.exec_())