#virtualenv -p ~/anaconda3/envs/solina_bot_32/python.exe

# 이 스크립트는 koa studio 에 접속하여 주식 정보를 다운로드한다
# 0178 10047 체결강도추이일별요청

from config import *
print("opt10047")

OPT_NUM = "opt10047"
OPT_NUM_J_NUM = '10047'
SCRN_NUM = "0178"
TR_REQ_TIME_INTERVAL = 0.2
SHEET_NAME = '10047'
outputs = output_list[OPT_NUM_J_NUM]

df_final = pd.DataFrame(columns=outputs)


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self.setup()

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def setup(self):
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self._receive_tr_data)

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        app = application.Application().connect(path=r"C:/OpenAPI/opstarter.exe")
        title = "Open API Login"
        dlg = timings.wait_until_passes(20, 0.5, lambda: app.window(title=title))
        pass_ctrl = dlg.Edit2
        pass_ctrl.set_focus()
        send_keys(PASS) # pass
        cert_ctrl = dlg.Edit3
        cert_ctrl.set_focus()
        send_keys(PASSID) # pass id
        send_keys("{ENTER}")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()


    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString", code,
                               real_type, field_name, index, item_name)
        return ret.strip()

    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2':
            self.remained_data = True
        else:
            self.remained_data = False

        self._opt10062(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def _opt10062(self, rqname, trcode):
        #print("=======================NEW FUNCTION CALL========================")
        data_cnt = self._get_repeat_cnt(trcode, rqname)
        global outputs
        global df_final
        temp_list = []

        for i in range(data_cnt):
            if df_final.shape[0] == 500: # max output data is 3000 rows
                self.remained_data = False
                break
            temp_dict = {}
            for output in outputs:
                if df_final.shape[0] == 500:
                    self.remained_data = False
                    break
                data = self._comm_get_data(trcode, "", rqname, i, output)
                temp_dict[output] = data
            temp_list.append(temp_dict)
            temp_df = pd.DataFrame([temp_dict])
            df_final = df_final.append(temp_df, ignore_index=True)


app = QApplication(sys.argv)
kiwoom = Kiwoom()
kiwoom.comm_connect()

# 10047 TR 요청
today = datetime.now()
start_date = today
today = today.strftime("%Y%m%d")
start_date = start_date.strftime("%Y%m%d")

for num in codes:
    print(num)
    kiwoom.set_input_value("종목코드", num)
    kiwoom.set_input_value("틱구분", "1") # 수량으로 구분
    kiwoom.set_input_value("채결강도구분","1") # 순매수 (2 는 순매도)
    kiwoom.comm_rq_data(OPT_NUM, OPT_NUM, 0, SCRN_NUM)

    while kiwoom.remained_data == True:
        time.sleep(TR_REQ_TIME_INTERVAL)
        kiwoom.set_input_value("종목코드", num)
        kiwoom.set_input_value("틱구분", "1") # 수량으로 구분
        kiwoom.set_input_value("채결강도구분","1") # 순매수 (2 는 순매도)
        kiwoom.comm_rq_data(OPT_NUM, OPT_NUM, 2, SCRN_NUM)

    # generating directory if DNE
    path = 'C:/OpenAPI/kiwoom_tradebot/solina_bot_data/' + today
    if not os.path.exists(path):
            os.makedirs(path)
            print("new directory generated")
    
    # generating excel sheet if DNE
    path_file = path + '/data.xlsx'
    if not os.path.exists(path_file):
        writer = pd.ExcelWriter(path_file, engine='openpyxl', mode='w')     
    # adding sheet to existing excel file
    else:
        writer = pd.ExcelWriter(path_file, engine='openpyxl', mode='a') 
    
    sheet_name_temp = SHEET_NAME + '_' + num
    df_final.to_excel(writer, sheet_name = sheet_name_temp, na_rep = 'NA', index = False, encoding = "utf-8-sig", engine = 'openpyxl')
    
    writer.save()
    writer.close()

    # emptying global df_final for iteration with next stock code
    df_final = df_final[0:0]
