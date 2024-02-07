import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

root = tk.Tk()
root.title("출장 정보 입력")

entry_frames = []

def browse_file(entry, is_input_path):
    global 입력경로, 출력경로
    file_path = filedialog.askopenfilename(title="파일 선택")
    entry.delete(0, tk.END)
    entry.insert(0, file_path)
    
    if is_input_path:
        입력경로 = file_path
    else:
        출력경로 = file_path

def add_file_path_fields():
    file_path_frame = tk.Frame(root)
    file_path_frame.grid(row=0, column=0, pady=5)

    labels = ["파일경로(출장신청목록)", "파일경로(여비상세)", "파일경로(공통항목)"]  # 추가된 부분
    entries = {}

    for i, label in enumerate(labels):
        tk.Label(file_path_frame, text=label).grid(row=i, column=0, padx=5, sticky="e")  # 수정된 부분
        entry = tk.Entry(file_path_frame, width=50)
        browse_button = tk.Button(file_path_frame, text="찾아보기", command=lambda entry=entry, is_input=i<2: browse_file(entry, is_input))  # 수정된 부분
        entry.grid(row=i, column=1, padx=70, sticky="w")  # 수정된 부분
        browse_button.grid(row=i, column=2, padx=5, sticky="w")  # 수정된 부분
        entries[label] = entry

    entry_frames.append((file_path_frame, entries))

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
            entry = ttk.Combobox(entry_frame, width=15)
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
    global 인적데이터, 경로데이터, 입력정보
    '''
    0:이름 1:직급 2:거래처구분(선택) 3:생년월일 4:입금유형(선택) 5:은행(선택) 6:계좌번호(숫자만) 7:등급(선택)
    8:출발지(선택) 9:도착지(선택) 10:교통편(선택) 11:정산유형(선택) 12:일비(원)
    '''

    입력정보 = []
    for frame, entries in entry_frames:
        entry_data = []
        for label, entry in entries.items():
            entry_data.append(entry.get())
        입력정보.append(entry_data)
 
    인적데이터 = []
    경로데이터 = []

    인적데이터 = 입력정보[1:]
    경로데이터 = 입력정보[0][0]

    bank_array(경로데이터[2])

def bank_array(공통항목_path):
    workbook = openpyxl.load_workbook(공통항목_path)
    sheet_은행코드 = workbook['은행코드_참고자료']
    
    global 은행_options 
    은행_options = []
    
    for row_num in range(2, 255):
        은행_options.append(sheet_은행코드.cell(row=row_num, column=1).value)

def 검증():
    print("\n-\n")
    print("입력정보:\n", 입력정보)
    print("\n-\n")
    print("경로데이터:\n", 경로데이터)
    print("\n-\n")
    print("인적데이터:\n", 인적데이터)
    print("\n-\n")
    print("입력경로:\n", 입력경로)
    print("\n-\n")
    print("출력경로:\n", 출력경로)


add_file_path_fields()

add_button = tk.Button(root, text="입력 추가", command=add_entry_fields)
add_button.grid(row=1, column=0, pady=10)

display_button = tk.Button(root, text="입력완료", command=display_entries)
display_button.grid(row=len(entry_frames) + 20, column=0, pady=10)

display_button = tk.Button(root, text="데이터검증", command=검증)
display_button.grid(row=len(entry_frames) + 40, column=0, pady=10)

root.mainloop()
