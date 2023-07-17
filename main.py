# 라이브러리 임포트
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import random as rd


class Student:
    def __init__(self):
        self.num = 0
        self.name = "g"
        self.sex = 0
        self.room = 0
        self.seat = [0, 0]


# 엑셀 입력 및 정리
student_xlsx = pd.read_excel('./student.xlsx')
freshman_num = int(student_xlsx["1학년 전체 수"][0])
junior_num = int(student_xlsx["2학년 전체 수"][0])
senior_num = int(student_xlsx["3학년 전체 수"][0])

# 사람, 방 리스트 생성
freshman_list = []
junior_list = []
senior_list = []
freshman_m_room = []
junior_m_room = []
senior_m_room = []
freshman_f_room = []
junior_f_room = []
senior_f_room = []
seat_list = []


for i in range(freshman_num):
    a = Student()
    a.num = int(student_xlsx["학번1"][i])
    a.name = student_xlsx["이름1"][i]
    if student_xlsx["성별1"][i] == "남":
        a.sex = 0
    else:
        a.sex = 1
    freshman_list.append(a)
for i in range(junior_num):
    a = Student()
    a.num = int(student_xlsx["학번2"][i])
    a.name = student_xlsx["이름2"][i]
    if student_xlsx["성별2"][i] == "남":
        a.sex = 0
    else:
        a.sex = 1
    junior_list.append(a)
for i in range(senior_num):
    a = Student()
    a.num = int(student_xlsx["학번3"][i])
    a.name = student_xlsx["이름3"][i]
    if student_xlsx["성별3"][i] == "남":
        a.sex = 0
    else:
        a.sex = 1
    senior_list.append(a)

for i in range(31):
    if int(student_xlsx["학년1"][i]) == 1:
        b = [201 + i, int(student_xlsx["인원수1"][i])]
        freshman_m_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 2:
        b = [201 + i, int(student_xlsx["인원수1"][i])]
        junior_m_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 3:
        b = [201 + i, int(student_xlsx["인원수1"][i])]
        senior_m_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 4:
        b = [201 + i, int(student_xlsx["인원수1"][i])]
        freshman_f_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 5:
        b = [201 + i, int(student_xlsx["인원수1"][i])]
        junior_f_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 6:
        b = [201 + i, int(student_xlsx["인원수1"][i])]
        senior_f_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 7:
        b = [201 + i, 1]
        freshman_m_room.append(b)
        b = [201 + i, 1]
        junior_m_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 8:
        b = [201 + i, 1]
        senior_m_room.append(b)
        b = [201 + i, 1]
        junior_m_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 9:
        b = [201 + i, 1]
        freshman_m_room.append(b)
        b = [201 + i, 1]
        senior_m_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 10:
        b = [201 + i, 1]
        freshman_f_room.append(b)
        b = [201 + i, 1]
        junior_f_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 11:
        b = [201 + i, 1]
        senior_f_room.append(b)
        b = [201 + i, 1]
        junior_f_room.append(b)
    elif int(student_xlsx["학년1"][i]) == 12:
        b = [201 + i, 1]
        freshman_f_room.append(b)
        b = [201 + i, 1]
        senior_f_room.append(b)
for i in range(31):
    if int(student_xlsx["학년2"][i]) == 1:
        b = [201 + i, int(student_xlsx["인원수2"][i])]
        freshman_m_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 2:
        b = [301 + i, int(student_xlsx["인원수2"][i])]
        junior_m_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 3:
        b = [301 + i, int(student_xlsx["인원수2"][i])]
        senior_m_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 4:
        b = [301 + i, int(student_xlsx["인원수2"][i])]
        freshman_f_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 5:
        b = [301 + i, int(student_xlsx["인원수2"][i])]
        junior_f_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 6:
        b = [301 + i, int(student_xlsx["인원수2"][i])]
        senior_f_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 7:
        b = [301 + i, 1]
        freshman_m_room.append(b)
        b = [301 + i, 1]
        junior_m_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 8:
        b = [301 + i, 1]
        senior_m_room.append(b)
        b = [301 + i, 1]
        junior_m_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 9:
        b = [301 + i, 1]
        freshman_m_room.append(b)
        b = [301 + i, 1]
        senior_m_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 10:
        b = [301 + i, 1]
        freshman_f_room.append(b)
        b = [301 + i, 1]
        junior_f_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 11:
        b = [301 + i, 1]
        senior_f_room.append(b)
        b = [301 + i, 1]
        junior_f_room.append(b)
    elif int(student_xlsx["학년2"][i]) == 12:
        b = [301 + i, 1]
        freshman_f_room.append(b)
        b = [301 + i, 1]
        senior_f_room.append(b)
for i in range(26):
    if int(student_xlsx["학년3"][i]) == 1:
        b = [401 + i, int(student_xlsx["인원수3"][i])]
        freshman_m_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 2:
        b = [401 + i, int(student_xlsx["인원수3"][i])]
        junior_m_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 3:
        b = [401 + i, int(student_xlsx["인원수3"][i])]
        senior_m_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 4:
        b = [401 + i, int(student_xlsx["인원수3"][i])]
        freshman_f_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 5:
        b = [401 + i, int(student_xlsx["인원수3"][i])]
        junior_f_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 6:
        b = [401 + i, int(student_xlsx["인원수3"][i])]
        senior_f_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 7:
        b = [401 + i, 1]
        freshman_m_room.append(b)
        b = [401 + i, 1]
        junior_m_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 8:
        b = [401 + i, 1]
        senior_m_room.append(b)
        b = [401 + i, 1]
        junior_m_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 9:
        b = [401 + i, 1]
        freshman_m_room.append(b)
        b = [401 + i, 1]
        senior_m_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 10:
        b = [401 + i, 1]
        freshman_f_room.append(b)
        b = [401 + i, 1]
        junior_f_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 11:
        b = [401 + i, 1]
        senior_f_room.append(b)
        b = [401 + i, 1]
        junior_f_room.append(b)
    elif int(student_xlsx["학년3"][i]) == 12:
        b = [401 + i, 1]
        freshman_f_room.append(b)
        b = [401 + i, 1]
        senior_f_room.append(b)

# 대망의 셔플
rd.shuffle(freshman_list)
rd.shuffle(junior_list)
rd.shuffle(senior_list)

# 셔플한 방 배분
for j in range(freshman_num):
    student = freshman_list[j]
    if student.sex == 0:
        student.room = freshman_m_room[0][0]
        freshman_m_room[0][1] -= 1
        if freshman_m_room[0][1] == 0: freshman_m_room.pop(0)
    else:
        student.room = freshman_f_room[0][0]
        freshman_f_room[0][1] -= 1
        if freshman_f_room[0][1] == 0: freshman_f_room.pop(0)
for j in range(junior_num):
    student = junior_list[j]
    if student.sex == 0:
        student.room = junior_m_room[0][0]
        junior_m_room[0][1] -= 1
        if junior_m_room[0][1] == 0: junior_m_room.pop(0)
    else:
        student.room = junior_f_room[0][0]
        junior_f_room[0][1] -= 1
        if junior_f_room[0][1] == 0: junior_f_room.pop(0)
for j in range(senior_num):
    student = senior_list[j]
    if student.sex == 0:
        student.room = senior_m_room[0][0]
        senior_m_room[0][1] -= 1
        if senior_m_room[0][1] == 0: senior_m_room.pop(0)
    else:
        student.room = senior_f_room[0][0]
        senior_f_room[0][1] -= 1
        if senior_f_room[0][1] == 0: senior_f_room.pop(0)

# 한번 더 셔플
rd.shuffle(freshman_list)
rd.shuffle(junior_list)
rd.shuffle(senior_list)

# 독서실 자리 배치
gg_1 = 0
gg_2 = 0
gg_3 = 0
for i in range(15):
    k = "북쪽라인" + str(i)
    a = student_xlsx[k]
    n_1 = int(a[0])
    n_2 = int(a[1])
    for m in range(15):
        if a[2+m] == 1:
            # 더 채워넣어야 할 것.
            # 마지막 독서실 자리만 넣으면 됨.

# UI 디자인
a = sorted(freshman_list, key=lambda x: x.room)
b = sorted(junior_list, key=lambda x: x.room)
c = sorted(senior_list, key=lambda x: x.room)
a = sorted(a + b + c,  key=lambda x: x.room)

for i in a:
    print(i.name, i.room)


# 배치 내보내기
