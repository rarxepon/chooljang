import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

'''
def browse_file(entry):
    file_path = filedialog.askopenfilename(title="파일 선택")
    entry.delete(0, tk.END)
    entry.insert(0, file_path)


def add_file_path_fields():
    file_path_frame = tk.Frame(root)
    file_path_frame.grid(row=0, column=0, pady=5)

    labels = ["파일경로(출장신청목록)", "파일경로(여비상세)"]
    entries = {}

    for i, label in enumerate(labels):
        tk.Label(file_path_frame, text=label).grid(row=0, column=i, padx=5, sticky="e")
        entry = tk.Entry(file_path_frame,width=50)
        browse_button = tk.Button(file_path_frame, text="찾아보기", command=lambda entry=entry: browse_file(entry))
        entry.grid(row=1, column=i, padx=70, sticky="w")
        browse_button.grid(row=1, column=i + 1, padx=5, sticky="w")
        entries[label] = entry

    entry_frames.append((file_path_frame, entries))
'''
#
def browse_file(entry, is_input_path):
    global 입력경로, 출력경로
    file_path = filedialog.askopenfilename(title="파일 선택")
    entry.delete(0, tk.END)
    entry.insert(0, file_path)
    
    # is_input_path에 따라 입력경로 또는 출력경로를 업데이트
    if is_input_path:
        입력경로 = file_path
    else:
        출력경로 = file_path

def add_file_path_fields():
    global 입력경로, 출력경로
    file_path_frame = tk.Frame(root)
    file_path_frame.grid(row=0, column=0, pady=5)

    labels = ["파일경로(출장신청목록)", "파일경로(여비상세)"]
    entries = {}

    for i, label in enumerate(labels):
        tk.Label(file_path_frame, text=label).grid(row=0, column=i, padx=5, sticky="e")
        entry = tk.Entry(file_path_frame, width=50)
        browse_button = tk.Button(file_path_frame, text="찾아보기", command=lambda entry=entry, is_input=i==0: browse_file(entry, is_input))
        entry.grid(row=1, column=i, padx=70, sticky="w")
        browse_button.grid(row=1, column=i + 1, padx=5, sticky="w")
        entries[label] = entry

    entry_frames.append((file_path_frame, entries))
#

def add_entry_fields():
    entry_frame = tk.Frame(root)
    entry_frame.grid(row=len(entry_frames) + 1, column=0, pady=5)

    labels = ["이름", "직급", "거래처구분(선택)", "생년월일", "입금유형(선택)",  "은행(선택)", "계좌번호(숫자만)", "등급(선택)",
              "출발지(선택)", "도착지(선택)", "교통편(선택)", "정산유형(선택)", "일비(원)"]
    entries = {}

    for i, label in enumerate(labels):
        tk.Label(entry_frame, text=label).grid(row=0, column=i, padx=5, sticky="e")

        if label == "거래처구분(선택)":
            거래처_options = ["10:법인사업자", "20:개인사업자", "30:개인", "40:기타"]
            entry = ttk.Combobox(entry_frame, values=거래처_options, width=15)
        elif label == "입금유형(선택)":
            입금유형_options = ["10:계좌이체", "20:대량이체", "30:원천징수", "40:고지서", "50:CMS", "60:수표", "99:현금"]
            entry = ttk.Combobox(entry_frame, values=입금유형_options, width=15)
        elif label == "은행(선택)":
            은행_options = ["001:한국은행", "002:산업은행", "003:기업은행", "004:국민은행", "007:수협은행",
                            "008:수출입은행", "011:농협은행", "012:지역 농축협", "020:우리은행", "023:SC제일은행",
                            "027:한국씨티은행", "031:대구은행","032:부산은행", "034:광주은행", "035:제주은행",
                            "037:전북은행", "039:경남은행", "045:새마을금고중앙회", "048:신협", "050:상호저축은행",
                            "051:기타 외국계은행(중국 교통은행 등)", "052:모간스탠리은행", "054:HSBC은행", "055:도이치은행",
                            "057:제이피모간체이스은행", "058:미즈호은행", "059:미쓰비시도쿄UFJ은행", "060:BOA은행",
                            "061:비엔피파리바은행", "062:중국공상은행", "063:중국은행", "064:산림조합중앙회", "065:대화은행",
                            "067:중국건설은행", "071:우체국","076:신용보증기금","077:기술보증기금", "081:KEB하나은행",
                            "088:신한은행", "089:케이뱅크", "090:카카오뱅크", "092:토스뱅크","093:한국주택금융공사",
                            "094:서울보증보험", "209:유안타증권","218:KB증권", "221:상상인증권", "223:리딩투자증권",
                            "224:BNK투자증권", "225:IBK투자증권", "226:KB증권","227:KTB투자증권", "238:미래에셋대우",
                            "240:삼성증권", "243:한국투자증권", "247:NH투자증권", "261:교보증권", "262:하이투자증권",
                            "263:현대차투자증권", "264:키움증권", "265:이베스트투자증권", "266:SK증권", "267:대신증권",
                            "268:메리츠종합금융증권", "269:한화투자증권", "270:하나금융투자", "271:토스증권","278:신한금융투자",
                            "279:DB금융투자", "280:유진투자증권", "287:메리츠종합금융증권", "288:카카오페이증권",
                            "289:NH투자증권","290:부국증권", "291:신영증권", "292:케이프투자증권","294:한국포스증권"]
            entry = ttk.Combobox(entry_frame, values=은행_options, width=15)
        elif label == "등급(선택)":
            등급_options = ["1등급", "2등급"]
            entry = ttk.Combobox(entry_frame, values=등급_options, width=15)
        elif label in ["출발지(선택)", "도착지(선택)"]:
            지역_options = ["동두천시청", "직접입력"]
            entry = ttk.Combobox(entry_frame, values=지역_options, width=15)
        elif label == "교통편(선택)":
            교통편_options = ["001:없음","002:자가","003:버스","004:철도","005:항공","006:선박"]
            entry = ttk.Combobox(entry_frame, values=교통편_options, width=15)
        elif label == "정산유형(선택)":
            정산유형_options = ["001:할인정액","002:실비","003:상한액1/2추가","004:상한액3/10추가"]
            entry = ttk.Combobox(entry_frame, values=정산유형_options, width=15)
        elif label == "계좌번호(숫자만)":
            entry = tk.Entry(entry_frame, width=15)
        else:
            entry = tk.Entry(entry_frame, width=9)

        entry.grid(row=1, column=i, padx=5, sticky="w")
        entries[label] = entry

    entry_frames.append((entry_frame, entries))
    

def display_entries():
    global 인적데이터
    '''
    0:이름 1:직급 2:거래처구분(선택) 3:생년월일 4:입금유형(선택) 5:은행(선택) 6:계좌번호(숫자만) 7:등급(선택)
    8:출발지(선택) 9:도착지(선택) 10:교통편(선택) 11:정산유형(선택) 12:일비(원)"
    '''
    입력정보 = []
    for frame, entries in entry_frames:
        entry_data = []
        for label, entry in entries.items():
            entry_data.append(entry.get())
        입력정보.append(entry_data)

    인적데이터 = 입력정보


def 검증():
    print("사용자 입력 데이터:", 인적데이터)
    print("입력경로:", 입력경로)
    print("출력경로:", 출력경로)



'''
def input_data_to_excel(workbook_path, sheet_name, start_cell, data_array):
    # 엑셀 파일 열기
    workbook = openpyxl.load_workbook(workbook_path)
    
    # 시트 선택
    sheet = workbook[sheet_name]
    
    # 시작 셀 좌표 추출
    start_row, start_col = openpyxl.utils.coordinate_from_string(start_cell)
    start_row = int(start_row)
    start_col = openpyxl.utils.column_index_from_string(start_col)
    
    # 데이터 입력
    for i in range(len(data_array)):
        for j in range(len(data_array[i])):
            sheet.cell(row=start_row + i, column=start_col + j, value=data_array[i][j])
    
    # 변경사항 저장
    workbook.save(workbook_path)
    print("데이터가 엑셀에 입력되었습니다.")

# 예제 사용법
workbook_path = 'your_excel_file.xlsx'
sheet_name = 'john'
start_cell = 'B1'


input_data_to_excel(workbook_path, sheet_name, start_cell, 인적데이터)
'''

if __name__ == "__main__":
    root = tk.Tk()
    root.title("출장 정보 입력")

    entry_frames = []

    add_file_path_fields()

    add_button = tk.Button(root, text="입력 추가", command=add_entry_fields)
    add_button.grid(row=1, column=0, pady=10)

    display_button = tk.Button(root, text="입력완료", command=display_entries)
    display_button.grid(row=len(entry_frames) + 20, column=0, pady=10)

    display_button = tk.Button(root, text="데이터검증", command=검증)
    display_button.grid(row=len(entry_frames) + 40, column=0, pady=10)

    root.mainloop()
