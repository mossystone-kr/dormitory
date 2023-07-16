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


# 엑셀 파일로 학생 정보+자습실 배치, 방 배치 등 입력
student_csv = pd.read_csv("./student.csv")          # 엑셀 입력 및 정리
junior_num = int(student_csv["1학년 전체 수"][0])
senior_num = int(student_csv["2학년 전체 수"][0])
sophister_num = int(student_csv["3학년 전체 수"][0])

junior_list = []                                    # 사람 리스트 생성
senior_list = []
sophister_list = []


for i in range(junior_num):
    a = Student()
    a.num = int(student_csv["학번1"][i])
    a.name = student_csv["이름1"][i]
    if student_csv["성별1"][i] == "남":
        a.sex = 0
    else:
        a.sex = 1
    junior_list.append(a)
for i in range(senior_num):
    a = Student()
    a.num = int(student_csv["학번2"][i])
    a.name = student_csv["이름2"][i]
    if student_csv["성별2"][i] == "남":
        a.sex = 0
    else:
        a.sex = 1
    senior_list.append(a)
for i in range(sophister_num):
    a = Student()
    a.num = int(student_csv["학번3"][i])
    a.name = student_csv["이름3"][i]
    if student_csv["성별3"][i] == "남":
        a.sex = 0
    else:
        a.sex = 1
    sophister_list.append(a)

# 배치
rd.shuffle(junior_list)                             # 대망의 셔플
rd.shuffle(senior_list)
rd.shuffle(sophister_list)


# UI 디자인


# 배치 내보내기
