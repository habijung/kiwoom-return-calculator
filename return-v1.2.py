import sys
import time
from datetime import datetime, timedelta
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
from PyQt5.QtGui import *



class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Get KIWOOM OpenAPI+
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # KIWOOM OpenAPI+ Login
        self.kiwoom.dynamicCall("CommConnect()")

        # KIWOOM OpenAPI+ Event
        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)

        # Create PyQt5 window
        self.setWindowTitle("KIWOOM Return Calculator")
        self.setGeometry(500, 100, 400, 600)


        ##########################
        #     Start Settings     #
        ##########################
        self.text_edit = QTextEdit(self)
        self.text_edit.setDisabled(True)
        self.text_edit.setGeometry(30, 260, 340, 320)
        self.text_edit.append("************************************")
        self.text_edit.append("*     KIWOOM Return Calculator")
        self.text_edit.append("*     Please Wait for Login...")

        self.btnQuit = self.create_btn("Quit", 140, 20)
        self.btnQuit.setDisabled(True)
        self.btnQuit.clicked.connect(QCoreApplication.instance().quit)

        self.get_date()


        ###########################
        #     Label : Account     #
        ###########################
        self.create_label("Account : ", 40, 60)

        self.btnAccount = self.create_btn("Check", 140, 60)
        self.btnAccount.setDisabled(True)
        self.btnAccount.clicked.connect(self.btn_account_click)

        self.text_account = QLineEdit(self)
        self.text_account.move(260, 60)
        self.text_account.setDisabled(True)


        ############################
        #     Label : Password     #
        ############################
        self.create_label("Password : ", 40, 100)

        self.btnPassword = self.create_btn("Submit", 140, 100)
        self.btnPassword.setDisabled(True)
        self.btnPassword.clicked.connect(self.btn_password_click)


        ##############################
        #     Label : Date Start     #
        ##############################
        self.create_label("Date Start : ", 40, 140)

        self.dateStartIns = QDateEdit(self)
        self.set_dateInstance(self.dateStartIns, self.dateStart, 140, 140)
        self.dateStartIns.setDisabled(True)
        self.dateStartIns.dateChanged.connect(self.update_dateStart)


        ############################
        #     Label : Date End     #
        ############################
        self.create_label("Date End : ", 40, 180)

        self.dateEndIns = QDateEdit(self)
        self.set_dateInstance(self.dateEndIns, self.dateEnd, 140, 180)
        self.dateEndIns.setDisabled(True)
        self.dateEndIns.dateChanged.connect(self.update_dateEnd)


        ############################
        #     Calculate Return     #
        ############################
        self.create_label("Return : ", 40, 220)

        self.btnReturn = self.create_btn("Calculate", 140, 220)
        self.btnReturn.setDisabled(True)
        self.btnReturn.clicked.connect(self.btn_return_click)



    ###############################
    #     Ftn : Program Start     #
    ###############################
    def event_connect(self, err_code):
        if err_code == 0:
            self.text_edit.append("*     Login Complete.")
            self.text_edit.append("************************************\n")
            self.get_started()

    def get_started(self):
        self.set_edit_on()
        self.get_account()
        self.pwState = 1
        self.print_date(self.text_edit, self.dateStart, "Start")
        self.print_date(self.text_edit, self.dateEnd, "End")

    def get_account(self):
        self.account = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).rstrip(';')
        accountStr = self.account[:4] + "-" + self.account[4:8] + "-" + self.account[8:]
        self.text_account.setText(accountStr)

    def get_password(self):
        self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
        self.pwState = 1

    def get_date(self):
        dateToday = [datetime.today().year, datetime.today().month, datetime.today().day]
        self.dateUpdated = 0
        self.dateStart = QDate(dateToday[0], dateToday[1], 1)
        self.dateEnd   = QDate(dateToday[0], dateToday[1], dateToday[2])

    def set_edit_on(self):
        self.text_edit.setEnabled(True)

        self.btnQuit.setEnabled(True)
        self.btnAccount.setEnabled(True)
        self.btnPassword.setEnabled(True)
        self.btnReturn.setEnabled(True)

        self.dateStartIns.setEnabled(True)
        self.dateEndIns.setEnabled(True)


    ########################
    #     Ftn : Create     #
    ########################
    def create_label(self, label_name, px, py):
        name = QLabel(label_name, self)
        name.move(px, py)

    def create_btn(self, btn_name, px, py):
        name = QPushButton(btn_name, self)
        name.move(px, py)
        name.resize(100, 30)

        return name


    ########################
    #     Ftn : Button     #
    ########################
    def btn_get_code(self):
        code = self.code_edit.text()
        self.text_edit.append("종목코드 : " + code)

        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "OPT10001_REQ", "OPT10001", 0, "0101")

    def btn_account_click(self):
        self.get_account()
        self.text_edit.append("Complete account check.")

    def btn_password_click(self):
        self.get_password()
        self.text_edit.append("Complete password submit.")

    def btn_return_click(self):
        self.text_edit.append("\n************************************")
        self.dateTmp = self.dateStart

        if self.dateTmp == self.dateEnd:
            self.request_return(self.dateTmp, self.dateEnd, 0)

        else:
            while self.dateTmp != self.dateEnd:
                self.request_return(self.dateTmp, self.dateEnd, 1)

            self.request_return(self.dateStart, self.dateEnd, 2)


    ##################################
    #     Ftn : Calculate Return     #
    ##################################
    def request_return(self, qStart, qEnd, sig):
        if sig == 0:
            date_start = self.set_dateString(qStart)

            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "계좌번호", self.account)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호", "")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가시작일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가종료일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", "00")
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "req0", "OPW00016", 0, "0101")

        elif sig == 1:
            date_start = self.set_dateString(qStart.addDays(1))

            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "계좌번호", self.account)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호", "")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가시작일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가종료일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", "00")
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "req0", "OPW00016", 0, "0101")

        else:
            date_start = self.set_dateString(qStart.addDays(1))
            date_end   = self.set_dateString(qEnd)

            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "계좌번호", self.account)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호", "")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가시작일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가종료일", date_end)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", "00")
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "req1", "OPW00016", 0, "0102")


    #########################
    #     Ftn : TR Data     #
    #########################
    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "OPT10001_REQ":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")
            volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "거래량")

            self.text_edit.append("종목명: " + name.strip())
            self.text_edit.append("거래량: " + volume.strip())
        
        elif rqname == "req0":
            ret = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, 0, "수익률")
            per = float(ret) / 100

            str1 = self.set_dateString(self.dateTmp)
            str2 = self.dateTmp.toString("ddd")
            str3 = str1[:4] + "-" + str1[4:6] + "-" + str1[6:] + " (" + str2 + ")"
            str4 = str3 + " : {:.2f} %".format(per)

            self.text_edit.append(str4)
            self.dateTmp = self.dateTmp.addDays(1)
            time.sleep(0.2)

        elif rqname == "req1":
            self.text_edit.append("Print Req1")


    ##############################
    #     Ftn : Date Control     #
    ##############################
    def print_date(self, target, qdate, dateType):
        dateStr = self.set_dateString(qdate)
        target.append("Date " + dateType + " : " + dateStr[:4] + "-" + dateStr[4:6] + "-" + dateStr[6:])

    def set_dateString(self, qdate):
        sYear  = qdate.toString("yyyy")
        sMonth = qdate.toString("MM")
        sDay   = qdate.toString("dd")
        sDate  = sYear + sMonth + sDay

        return sDate

    def set_dateInstance(self, dateIns, dateTmp, px, py):
        dateIns.setCalendarPopup(True)
        dateIns.setDateRange(QDate(2000, 1, 1), QDate(2099, 12, 31))
        dateIns.setDate(dateTmp)
        dateIns.move(px, py)

    def update_dateStart(self, newDate):
        if newDate > self.dateEndIns.date():
            self.dateUpdated += 1
            self.dateStartIns.setDate(self.dateEndIns.date())

            # Date Start > Date End
            if self.dateUpdated == 2:
                self.dateUpdated = 0
                self.text_edit.append("Start date cannot exceed end date.")
                self.print_date(self.text_edit, self.dateStart, "Start")

        else:
            self.dateStart = newDate

            # Date Start <= Date End
            if self.dateUpdated == 0:
                self.print_date(self.text_edit, self.dateStart, "Start")
            
    def update_dateEnd(self, newDate):
        if newDate < self.dateStartIns.date():
            self.dateUpdated += 1
            self.dateEndIns.setDate(self.dateStartIns.date())

            # Date End < Date Start
            if self.dateUpdated == 2:
                self.dateUpdated = 0
                self.text_edit.append("End date cannot precede start date.")
                self.print_date(self.text_edit, self.dateEnd, "End")

        elif newDate > QDate.currentDate():
            self.dateUpdated += 1
            self.dateEndIns.setDate(QDate.currentDate())

            # Date End > Today
            if self.dateUpdated == 2:
                self.dateUpdated = 0
                self.text_edit.append("End date cannot exceed today.")
                self.print_date(self.text_edit, self.dateEnd, "End")

        else:
            self.dateEnd = newDate
            
            # Date End <= Today
            if self.dateUpdated == 0:
                self.print_date(self.text_edit, self.dateEnd, "End")





if __name__ == "__main__":
    print("\n===== kiwoom.py =====\n")

    app = QApplication(sys.argv)

    mywindow = MyWindow()
    mywindow.show()

    app.exec_()
