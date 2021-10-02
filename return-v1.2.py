import sys
import time
from datetime import datetime, timedelta
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Get KIWOOM OpenAPI+
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # KIWOOM OpenAPI+ Login
        #self.kiwoom.dynamicCall("CommConnect()")

        # KIWOOM OpenAPI+ Event
        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)

        # Create PyQt5 window
        self.setWindowTitle("KIWOOM")
        self.setGeometry(500, 100, 400, 600)


        ##########################
        #     Start Settings     #
        ##########################
        self.text_edit = QTextEdit(self)
        self.text_edit.setEnabled(True)
        self.text_edit.setGeometry(30, 260, 340, 320)
        self.text_edit.append("Please Wait Login...")

        self.get_date()
        self.create_btn("Quit", 140, 20, QCoreApplication.instance().quit)


        ###########################
        #     Label : Account     #
        ###########################
        self.create_label("Account : ", 40, 60)
        self.create_btn("Check", 140, 60, self.btn_account)
        self.text_account = QLineEdit(self)
        self.text_account.move(260, 60)


        ############################
        #     Label : Password     #
        ############################
        self.create_label("Password : ", 40, 100)
        self.create_btn("Submit", 140, 100, self.btn_password)


        ##############################
        #     Label : Date Start     #
        ##############################
        self.create_label("Date Start : ", 40, 140)
        #self.create_btn("Apply", 260, 140, self.btn_dateStart)

        self.dateStartIns = QDateEdit(self)
        self.set_dateInstance(self.dateStartIns, self.dateStart, 140, 140)
        self.dateStartIns.dateChanged.connect(self.update_dateStart)


        ############################
        #     Label : Date End     #
        ############################
        self.create_label("Date End : ", 40, 180)
        #self.create_btn("Apply", 260, 180, self.btn_dateEnd)

        self.dateEndIns = QDateEdit(self)
        self.set_dateInstance(self.dateEndIns, self.dateEnd, 140, 180)
        self.dateEndIns.dateChanged.connect(self.update_dateEnd)


        ############################
        #     Calculate Return     #
        ############################
        self.create_label("Return : ", 40, 220)
        self.create_btn("Calculate", 140, 220, self.btn_return)


    def tt(self, newDate):
        print("TEST" + newDate.toString())

    #########################
    #     Ftn : Connect     #
    #########################
    def event_connect(self, err_code):
        if err_code == 0:
            self.get_started()
            self.text_edit.append("Login Complete.")


    ########################
    #     Ftn : Create     #
    ########################
    def create_label(self, label_name, px, py):
        name = QLabel(label_name, self)
        name.move(px, py)

    def create_btn(self, btn_name, px, py, ftn):
        name = QPushButton(btn_name, self)
        name.move(px, py)
        name.resize(100, 30)
        name.clicked.connect(ftn)


    ##########################
    #     Ftn : Get Info     #
    ##########################
    def get_started(self):
        self.get_account()
        self.pwState = 1
        self.print_date(self.text_edit, self.dateStartStr, "Start")
        self.text_edit.append("Date start : " + self.dateStartStr[0] + "-" + self.dateStartStr[1] + "-" + self.dateStartStr[2])
        self.text_edit.append("Date end : "   + self.dateEndStr[0]   + "-" + self.dateEndStr[1]   + "-" + self.dateEndStr[2])

    def get_account(self):
        self.account = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).rstrip(';')
        self.account = self.account[:4] + "-" + self.account[4:8] + "-" + self.account[8:]
        self.text_account.setText(self.account)

    def get_password(self):
        self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
        self.pwState = 1

    def get_date(self):
        dateToday = [datetime.today().year, datetime.today().month, datetime.today().day]
        self.dateUpdated = 0
        self.dateStart = [dateToday[0], dateToday[1], 1]
        self.dateEnd   = [dateToday[0], dateToday[1], dateToday[2]]
        self.dateStartStr = self.set_dateString(self.dateStart)
        self.dateEndStr   = self.set_dateString(self.dateEnd)


    ########################
    #     Ftn : Button     #
    ########################
    def btn_account(self):
        self.get_account()
        self.text_edit.append("Complete account check.")

    def btn_password(self):
        self.get_password()
        self.text_edit.append("Complete password submit.")

    def btn_dateStart(self):
        sYear  = self.set_dateStart.date().toString("yyyy")
        sMonth = self.set_dateStart.date().toString("MM")
        sDay   = self.set_dateStart.date().toString("dd")
        self.dateStart = [int(sYear), int(sMonth), int(sDay)]
        self.dateStartStr = [sYear, sMonth, sDay]
        self.text_edit.append("Date Start : " + sYear + "-" + sMonth + "-" + sDay)


    def btn_dateEnd(self):
        sYear  = self.set_dateEnd.date().toString("yyyy")
        sMonth = self.set_dateEnd.date().toString("MM")
        sDay   = self.set_dateEnd.date().toString("dd")
        self.dateEnd = [int(sYear), int(sMonth), int(sDay)]
        self.dateEndStr = [sYear, sMonth, sDay]
        self.text_edit.append("Date End : " + sYear + "-" + sMonth + "-" + sDay)


    def btn_return(self):
        account = self.account
        password_state = self.pwState
        password_type  = "00"
        date_start = self.dateStartStr[0] + self.dateStartStr[1] + self.dateStartStr[2]
        date_end = self.dateEndStr[0] + self.dateEndStr[1] + self.dateEndStr[2]

        if password_state == 0:
            self.text_edit.append("Please submit your password.")

        else:
            self.text_edit.append("\n*************************")
            self.text_edit.append("    " + date_start + " - " + date_end)

            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "계좌번호", account)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호", "")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가시작일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가종료일", date_end)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", password_type)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "OPW00016_REQ", "OPW00016", 0, "0101")


    def btn_get_code(self):
        code = self.code_edit.text()
        self.text_edit.append("종목코드 : " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "OPT10001_REQ", "OPT10001", 0, "0101")


    #########################
    #     Ftn : TR Data     #
    #########################
    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "OPT10001_REQ":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")
            volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "거래량")

            self.text_edit.append("종목명: " + name.strip())
            self.text_edit.append("거래량: " + volume.strip())
        
        elif rqname == "OPW00016_REQ":
            yieldResult = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, 0, "수익률")
            yieldPercent = float(yieldResult) / 100
            
            self.text_edit.append("    Return : {:.2f} %".format(yieldPercent))
            self.text_edit.append("*************************\n")


    ########################
    #     Ftn : Others     #
    ########################
    def print_date(self, target, dateStr, dateType):
        target.append("Date " + dateType + " : "+ dateStr[0] + "-" + dateStr[1] + "-" + dateStr[2])

    def set_dateString(self, dateArr):
        tmpDate = datetime(dateArr[0], dateArr[1], dateArr[2])
        sYear   = tmpDate.strftime("%Y")
        sMonth  = tmpDate.strftime("%m")
        sDay    = tmpDate.strftime("%d")
        
        return [sYear, sMonth, sDay]

    def set_dateInstance(self, dateIns, dateArr, px, py):
        dateIns.setCalendarPopup(True)
        dateIns.setDateRange(QDate(2000, 1, 1), QDate(2099, 12, 31))
        dateIns.setDate(QDate(dateArr[0], dateArr[1], dateArr[2]))
        dateIns.move(px, py)

    def update_dateStart(self, newDate):
        if newDate > self.dateEndIns.date():
            self.dateUpdated += 1
            self.dateStartIns.setDate(self.dateEndIns.date())

            # Date Start > Date End
            if self.dateUpdated == 2:
                self.dateUpdated = 0
                self.text_edit.append("\nStart date cannot exceed end date.")
                self.print_date(self.text_edit, self.dateStartStr, "Start")

        else:
            nYear  = newDate.toString("yyyy")
            nMonth = newDate.toString("MM")
            nDay   = newDate.toString("dd")

            self.dateStart = [int(nYear), int(nMonth), int(nDay)]
            self.dateStartStr = self.set_dateString(self.dateStart)

            # Date Start <= Date End
            if self.dateUpdated == 0:
                self.print_date(self.text_edit, self.dateStartStr, "Start")
            
    def update_dateEnd(self, newDate):
        if newDate < self.dateStartIns.date():
            self.dateUpdated += 1
            self.dateEndIns.setDate(self.dateStartIns.date())

            # Date End < Date Start
            if self.dateUpdated == 2:
                self.dateUpdated = 0
                self.text_edit.append("\nEnd date cannot precede start date.")
                self.print_date(self.text_edit, self.dateEndStr, "End")

        elif newDate > QDate.currentDate():
            self.dateUpdated += 1
            self.dateEndIns.setDate(QDate.currentDate())

            # Date End > Today
            if self.dateUpdated == 2:
                self.dateUpdated = 0
                self.text_edit.append("\nEnd date cannot exceed today.")
                self.print_date(self.text_edit, self.dateEndStr, "End")

        else:
            nYear  = newDate.toString("yyyy")
            nMonth = newDate.toString("MM")
            nDay   = newDate.toString("dd")

            self.dateEnd = [int(nYear), int(nMonth), int(nDay)]
            self.dateEndStr = self.set_dateString(self.dateEnd)
            
            # Date End <= Today
            if self.dateUpdated == 0:
                self.print_date(self.text_edit, self.dateEndStr, "End")





if __name__ == "__main__":
    print("\n===== kiwoom.py =====\n")

    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()
