import sys
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
        self.kiwoom.dynamicCall("CommConnect()")

        # KIWOOM OpenAPI+ Event
        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)

        # Create PyQt5 window
        self.setWindowTitle("KIWOOM")
        self.setGeometry(300, 100, 300, 800)

        # Create Log Message Box
        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(10, 260, 280, 500)
        self.text_edit.setEnabled(False)
        self.text_edit.append("Please Wait Login...")

        # Program Quit
        btn_quit = QPushButton("Quit", self)
        btn_quit.move(80, 20)
        btn_quit.clicked.connect(QCoreApplication.instance().quit)


        ###########################
        #     Label : Account     #
        ###########################
        label_account = QLabel('Account : ', self)
        label_account.move(20, 60)

        self.text_account = QLineEdit(self)
        self.text_account.move(80, 60)

        btn_account = QPushButton("Check", self)
        btn_account.move(190, 60)
        btn_account.clicked.connect(self.btn_account_click)


        ############################
        #     Label : Password     #
        ############################
        label_password = QLabel('PW : ', self)
        label_password.move(20, 100)

        btn_password = QPushButton("Submit", self)
        btn_password.move(190, 100)
        btn_password.clicked.connect(self.btn_password_click)


        ##############################
        #     Label : Date Start     #
        ##############################
        label_dateStart = QLabel('Date Start : ', self)
        label_dateStart.move(20, 140)

        self.dateStart = [datetime.today().year, datetime.today().month, 1]
        self.set_dateStart = QDateEdit(self)
        self.set_dateStart.setCalendarPopup(True)
        self.set_dateStart.setDateRange(QDate(1999, 1, 1), QDate(2099, 12, 31))
        self.set_dateStart.setDate(QDate(self.dateStart[0], self.dateStart[1], self.dateStart[2]))
        self.set_dateStart.move(80, 140)

        btn_dateStart = QPushButton("Apply", self)
        btn_dateStart.move(190, 140)
        btn_dateStart.clicked.connect(self.btn_dateStart_click)


        ############################
        #     Label : Date End     #
        ############################
        label_dateStart = QLabel('Date End : ', self)
        label_dateStart.move(20, 180)

        self.dateEnd = [datetime.today().year, datetime.today().month, datetime.today().day]
        self.set_dateEnd = QDateEdit(self)
        self.set_dateEnd.setCalendarPopup(True)
        self.set_dateEnd.setDateRange(QDate(1999, 1, 1), QDate(2099, 12, 31))
        self.set_dateEnd.setDate(QDate(self.dateEnd[0], self.dateEnd[1], self.dateEnd[2]))
        self.set_dateEnd.move(80, 180)

        btn_dateEnd = QPushButton("Apply", self)
        btn_dateEnd.move(190, 180)
        btn_dateEnd.clicked.connect(self.btn_dateEnd_click)


        ###########################
        #     Calculate Yield     #
        ###########################
        label_yield = QLabel('Yield : ', self)
        label_yield.move(20, 220)

        btn_yield = QPushButton("Calculate", self)
        btn_yield.move(190, 220)
        btn_yield.clicked.connect(self.btn_yield_click)


    ####################
    #     Function     #
    ####################
    def event_connect(self, err_code):
        if err_code == 0:
            self.get_started()
            self.text_edit.append("Login Complete.")


    def get_started(self):
        self.get_account()
        self.pwState      = 0
        self.dateStartStr = [self.set_dateStart.date().toString("yyyy"),
                             self.set_dateStart.date().toString("MM"),
                             self.set_dateStart.date().toString("dd")]
        self.dateEndStr   = [self.set_dateEnd.date().toString("yyyy"),
                             self.set_dateEnd.date().toString("MM"),
                             self.set_dateEnd.date().toString("dd")]
        

    def get_account(self):
        self.account = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).rstrip(';')
        self.accountStr = self.account[:4] + "-" + self.account[4:8] + "-" + self.account[8:]
        self.text_account.setText(self.accountStr)

    
    def get_password(self):
        self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", "ShowAccountWindow", "")
        self.pwState = 1


    def btn_account_click(self):
        self.get_account()
        self.text_edit.append("Account Check.")

    
    def btn_password_click(self):
        self.get_password()
        self.text_edit.append("Submit Complete.")


    def btn_dateStart_click(self):
        sYear  = self.set_dateStart.date().toString("yyyy")
        sMonth = self.set_dateStart.date().toString("MM")
        sDay   = self.set_dateStart.date().toString("dd")
        self.dateStart = [int(sYear), int(sMonth), int(sDay)]
        self.dateStartStr = [sYear, sMonth, sDay]
        self.text_edit.append("Date Start : " + sYear + "-" + sMonth + "-" + sDay)


    def btn_dateEnd_click(self):
        sYear  = self.set_dateEnd.date().toString("yyyy")
        sMonth = self.set_dateEnd.date().toString("MM")
        sDay   = self.set_dateEnd.date().toString("dd")
        self.dateEnd = [int(sYear), int(sMonth), int(sDay)]
        self.dateEndStr = [sYear, sMonth, sDay]
        self.text_edit.append("Date End : " + sYear + "-" + sMonth + "-" + sDay)


    def btn_yield_click(self):
        account = self.account
        password_state = self.pwState
        date_start = self.dateStartStr[0] + self.dateStartStr[1] + self.dateStartStr[2]
        date_end = self.dateEndStr[0] + self.dateEndStr[1] + self.dateEndStr[2]
        pw_type = "00"

        if password_state == 0:
            self.text_edit.append("Submit your password.")

        else:
            self.text_edit.append("\n*************************")
            self.text_edit.append("    " + date_start + " - " + date_end)

            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "계좌번호", account)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호", "")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가시작일", date_start)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "평가종료일", date_end)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString", "비밀번호입력매체구분", pw_type)
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "OPW00016_REQ", "OPW00016", 0, "0101")


    def btn_get_code(self):
        code = self.code_edit.text()
        self.text_edit.append("종목코드 : " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "OPT10001_REQ", "OPT10001", 0, "0101")


    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        if rqname == "OPT10001_REQ":
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")
            volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "거래량")

            self.text_edit.append("종목명: " + name.strip())
            self.text_edit.append("거래량: " + volume.strip())
        
        elif rqname == "OPW00016_REQ":
            yieldResult = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, 0, "수익률")
            yieldPercent = float(yieldResult) / 100
            
            self.text_edit.append("    Yield : {:.2f} %".format(yieldPercent))
            self.text_edit.append("*************************\n")



if __name__ == "__main__":
    print("\n===== kiwoom.py =====\n")

    app = QApplication(sys.argv)
    mywindow = MyWindow()
    mywindow.show()
    app.exec_()
