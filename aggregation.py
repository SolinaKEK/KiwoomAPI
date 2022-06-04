#virtualenv -p ~/anaconda3/envs/solina_bot_32/python.exe

from config import *
from get_krn_names import df_krn_names

print("aggregation")
# reading stocks code list
file = open("tradebot_codes.txt", "r")
codes = file.read().splitlines()

today = datetime.now()
today = today.strftime("%Y%m%d")
source_file = 'C:/OpenAPI/kiwoom_tradebot/solina_bot_data/' + today + '/data.xlsx'
destination = "C:/OpenAPI/kiwoom_tradebot/solina_bot_data/final_sheets/"
origin = "C:/OpenAPI/kiwoom_tradebot/reference_file"

# reading column headers from reference file
df = pd.read_excel(origin + '/reference_file.xlsx', sheet_name = 'RAWDATA', engine = 'openpyxl', index_col=None)
df = df[0:0]

# 여기서부터 종목 코드 별 진행
for c in codes:
    # reference file 복사
    row_index = df_krn_names.index[df_krn_names['종목코드'] == c].tolist()
    row_index = int(row_index[0])
    krn_name = df_krn_names.iat[row_index,1]


    file_name = today + '_' + c + '.xlsx'
    file_path = destination + file_name
    final_path = str(destination + today + '_' + krn_name + '.xlsx')


    files = os.listdir(origin)
    if not os.path.isdir(destination):
        os.makedirs(destination)
    for file in files:
        if not os.path.exists(file_path):
            shutil.copy(origin + '/reference_file.xlsx', file_path)
            print("success: reference file copied")

    wb = xw.Book(file_path)
    ws = wb.sheets["RAWDATA"]
    ws2 = wb.sheets["창문"]

    sheet_name = '10059_' + c
    df1 = pd.read_excel(source_file, sheet_name=sheet_name, engine = 'openpyxl', index_col=None)
    sheet_name = '10047_' + c
    df2 = pd.read_excel(source_file, sheet_name=sheet_name, engine = 'openpyxl', index_col=None)

    del df1["대비기호"]
    del df1["등락률"]
    del df1["누적거래대금"]

    df1.insert(3, "체결강도", df2["체결강도"], True)
    df1.insert(8, "", 0, True)

    # 엑셀 시트에 맞춰서 데이터 형태 바꾸는 중...
    myList = df1["현재가"].tolist()
    df1["현재가"] = list(map(abs, myList))
    myList2 = df1["체결강도"].tolist()
    df1["체결강도"] = list(map(lambda x: x/100, myList2))

    ws.range('A1').options(pd.DataFrame, index=False).value = df1
    ws2.range('B2').value = krn_name

    wb.save(file_path)
    wb.close()

    os.rename(file_path, final_path)

os.system("taskkill /f /im Excel.exe")
print("success. complete.")



