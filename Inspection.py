# This Python file uses the following encoding: utf-8
# -*- coding: utf-8 -*-

'''
네이버 지식인 진료 분야별 건강상담 크롤링 - 크롤링 데이터 정제
'''

#%%
import os
import winsound
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd

# 엑셀 파일 오픈
def openfile():
    global file_path, df, curr_index

    file_path = filedialog.askopenfilename(parent = root, filetypes = [('Xlsx files', '*.xlsx')])
    if (file_path != ''):
        df = pd.read_excel(file_path)

        file_name_var.set(os.path.split(file_path)[-1])
        curr_index = 1
        number_var.set(f'{curr_index}')
        show_data()

# 입력한 행 번호에 해당하는 DB 내용 출력
def show_data(event=None):
    global curr_index

    if (not number_var.get().isdigit()):
        messagebox.showinfo('크롤링', '정확한 인덱스 값이 아닙니다.')
        number_var.set(f'{curr_index}')         # 이전 인덱스로 복구
        return

    num = int(number_var.get()) - 1
    if (0 <= num < len(df)):
        A_title_var.set(df.loc[num][0])
        B_question.delete(1.0, END)
        B_question.insert(1.0, df.loc[num][1])
        C_answer.delete(1.0, END)
        C_answer.insert(1.0, df.loc[num][2])
        D_url_var.set(df.loc[num][3])
        progress_var.set('{}/{}'.format(num+1, len(df)))
        curr_index = int(number_var.get())
    else:
        messagebox.showinfo('크롤링', '인덱스 범위를 벗어났습니다.')
        number_var.set(f'{curr_index}')

# 수정된 DB 내용 저장
def save_data(event=None):
    if (not number_var.get().isdigit()):
        messagebox.showinfo('크롤링', '정확한 인덱스 값이 아닙니다.')
        number_var.set(f'{curr_index}')         # 이전 인덱스로 복구
        return

    num = int(number_var.get()) - 1
    if (0 <= num < len(df)):
        # 수정한 값으로 저장
        df.loc[num][0] = A_title_var.get()
        df.loc[num][1] = B_question.get(1.0, END)
        df.loc[num][2] = C_answer.get(1.0, END)
        df.loc[num][3] = D_url_var.get()
        df.to_excel(file_path, index=False)

        winsound.PlaySound('SystemExit', winsound.SND_ALIAS | winsound.SND_ASYNC)

# 삭제
def delete_data():
    if (not number_var.get().isdigit()):
        messagebox.showinfo('크롤링', '정확한 인덱스 값이 아닙니다.')
        number_var.set(f'{curr_index}')         # 이전 인덱스로 복구
        return

    num = int(number_var.get()) - 1
    if (0 <= num < len(df)):
        # 지정한 행 자료 삭제 후 저장, 행 번호는 재설정
        df.drop(num, inplace=True)
        df.reset_index(drop=True, inplace=True)
        df.to_excel(file_path, index=False)

        if (len(df) == 0):          # 전체가 삭제된 경우는 화면 초기화
            # 화면 초기화
            number_var.set('')
            A_title_var.set('')
            B_question.delete(1.0, END)
            C_answer.delete(1.0, END)
            D_url_var.set('')
            progress_var.set('{}/{}'.format(0, len(df)))
        else:
            if (num >= len(df)):    # 맨 뒤 자료가 삭제된 경우는 앞으로 이동
                pre_data()
            else:
                show_data()

        winsound.PlaySound('SystemHand', winsound.SND_ALIAS | winsound.SND_ASYNC)

# 이전
def pre_data():
    if (not number_var.get().isdigit()):
        messagebox.showinfo('크롤링', '정확한 인덱스 값이 아닙니다.')
        number_var.set(f'{curr_index}')         # 이전 인덱스로 복구
        return

    num = int(number_var.get()) - 1
    if (0 < num):
        number_var.set(f'{num}')
        show_data()
    else:
        messagebox.showinfo('크롤링', '처음 자료입니다.')

# 다음
def next_data():
    if (not number_var.get().isdigit()):
        messagebox.showinfo('크롤링', '정확한 인덱스 값이 아닙니다.')
        number_var.set(f'{curr_index}')         # 이전 인덱스로 복구
        return

    num = int(number_var.get()) - 1
    if (num < len(df)-1):
        number_var.set(f'{num+2}')
        show_data()
    else:
        messagebox.showinfo('크롤링', '마지막 자료입니다.')


root = Tk()
root.resizable(False, False)
root.title('건강 관련 지식인 크롤링')

file_path = ''
df = None
curr_index = 1                      # 1-based

file_name_var = StringVar()
number_var = StringVar()
progress_var = StringVar()
progress_var.set('0/0')

A_title_var = StringVar()           # 제목
D_url_var = StringVar()             # Q&A url

Button(root, text=' Xlsx DB 열기 ', command=openfile).grid(row=0, column=0, padx=10, pady=10, sticky='ewsn')
Label(root, textvariable=file_name_var, width=20).grid(row=0, column=1, padx=10, pady=10, sticky='ewsn')

entry = Entry(root, textvariable=number_var)
entry.grid(row=1, column=0, padx=10, pady=10, sticky='ewsn')
entry.bind('<Return>', show_data)
Label(root, textvariable=progress_var).grid(row=1, column=1, padx=10, pady=10, sticky='ewsn')

Entry(root, textvariable=A_title_var, width=40).grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky='ewsn')
Entry(root, textvariable=D_url_var, width=80).grid(row=2, column=2, columnspan=4, padx=10, pady=10, sticky='ewsn')

Button(root, text=' ◀(이전) ', command=pre_data).grid(row=0, column=2, rowspan=2, padx=10, pady=10, sticky='ewsn')
Button(root, text=' ▶(다음) ', command=next_data).grid(row=0, column=3, rowspan=2, padx=10, pady=10, sticky='ewsn')

Button(root, text=' 저장 ', command=save_data).grid(row=0, column=4, rowspan=2, padx=10, pady=10, sticky='ewsn')
root.bind('<Control-s>', save_data)
Button(root, text=' 삭제 ', command=delete_data).grid(row=0, column=5, rowspan=2, padx=10, pady=10, sticky='ewsn')

B_question = Text(root, height=15)  # 질문
B_question.grid(row=3, column=0, columnspan=6, padx=10, pady=10, sticky='ewsn')
C_answer = Text(root, height=20)    # 답변
C_answer.grid(row=4, column=0, columnspan=6, padx=10, pady=10, sticky='ewsn')

root.mainloop()
