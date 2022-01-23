#virtualenv -p ~/anaconda3/envs/solina_bot_32/python.exe

# 이 스크립트는 번개3을 통해 오픈 api 버전 업데이트를 자동으로 처리한다
import sys
import warnings
warnings.simplefilter("ignore", UserWarning)
sys.coinit_flags = 2
from pywinauto import application
from pywinauto import timings
import time
import os
import sys

app = application.Application()
app.start("C:/KiwoomFlash3/bin/nkministarter.exe")

title = "번개3 Login"
dlg = timings.WaitUntilPasses(20, 0.5, lambda: app.window(title=title))

# id 는 기억됨

pass_ctrl = dlg.Edit2
pass_ctrl.SetFocus()
pass_ctrl.TypeKeys('') # pass

cert_ctrl = dlg.Edit3
cert_ctrl.SetFocus()
cert_ctrl.TypeKeys('') # pass id

btn_ctrl = dlg.Button0
btn_ctrl.click()

time.sleep(50)
os.system("taskkill /f /im nkmini.exe")
print("버전처리 성공")
