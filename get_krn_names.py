from config import *

print("get_krn_names")

file = open("tradebot_codes.txt", "r")
codes = file.read().splitlines()

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self.setup()

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def setup(self):
        self.OnEventConnect.connect(self._event_connect)

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        app = application.Application().connect(path=r"C:/OpenAPI/opstarter.exe")
        title = "Open API Login"
        dlg = timings.wait_until_passes(20, 0.5, lambda: app.window(title=title))
        pass_ctrl = dlg.Edit2
        pass_ctrl.set_focus()
        pass_ctrl.type_keys(PASS) # pass
        cert_ctrl = dlg.Edit3
        cert_ctrl.set_focus()
        cert_ctrl.type_keys(PASSID) # pass id
        btn_ctrl = dlg.Button0
        btn_ctrl.click()
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()



app = QApplication(sys.argv)
kiwoom = Kiwoom()
kiwoom.comm_connect()

df_krn_names = pd.DataFrame(columns = ['종목코드', '종목명'])
#print(df_krn_names)

for num in codes:
    #print(num)
    krn_name = kiwoom.GetMasterCodeName(num)
    #print(krn_name)
    df_krn_names = df_krn_names.append({'종목코드': num, '종목명': krn_name}, ignore_index=True)
    #print(df_krn_names)
