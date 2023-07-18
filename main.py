# 라이브러리 임포트
import sys
import pandas as pd
import random as rd
import uitest as ui


class Student:
    def __init__(self):
        self.num = 0
        self.name = "g"
        self.sex = 0
        self.room = 0
        self.seat = [0, 0]


# 엑셀 입력 및 정리
student_xlsx = pd.read_excel('./student.xlsx')  # 이 부분을 studentFileName을 이용해서 바꿔주면 됨
freshman_num = int(student_xlsx["1학년 전체 수"][0])
junior_num = int(student_xlsx["2학년 전체 수"][0])
senior_num = int(student_xlsx["3학년 전체 수"][0])

# 사람, 방 리스트 생성
freshman_list = []
junior_list = []
senior_list = []
freshman_m_list = []
junior_m_list = []
senior_m_list = []
freshman_f_list = []
junior_f_list = []
senior_f_list = []
freshman_m_room = []
junior_m_room = []
senior_m_room = []
freshman_f_room = []
junior_f_room = []
senior_f_room = []
seat_list = []
list_to_fix_room = []
list_to_fix_seat = []

for i in range(freshman_num):
    a = Student()
    a.num = int(student_xlsx["학번1"][i])
    a.name = student_xlsx["이름1"][i]
    if student_xlsx["성별1"][i] == "남":
        a.sex = 0
        freshman_m_list.append(a)
    else:
        a.sex = 1
        freshman_f_list.append(a)
for i in range(junior_num):
    a = Student()
    a.num = int(student_xlsx["학번2"][i])
    a.name = student_xlsx["이름2"][i]
    if student_xlsx["성별2"][i] == "남":
        a.sex = 0
        junior_m_list.append(a)
    else:
        a.sex = 1
        junior_f_list.append(a)
    junior_list.append(a)
for i in range(senior_num):
    a = Student()
    a.num = int(student_xlsx["학번3"][i])
    a.name = student_xlsx["이름3"][i]
    if student_xlsx["성별3"][i] == "남":
        a.sex = 0
        senior_m_list.append(a)
    else:
        a.sex = 1
        senior_f_list.append(a)
    senior_list.append(a)

freshman_list = freshman_m_list + freshman_f_list
junior_list = junior_m_list + junior_f_list
senior_list = senior_m_list + senior_f_list

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

# 고정 인원 추출
for i in range(int(student_xlsx["고정학번1"][0])):


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
rd.shuffle(freshman_m_list)
rd.shuffle(freshman_f_list)
rd.shuffle(junior_m_list)
rd.shuffle(junior_f_list)
rd.shuffle(senior_m_list)
rd.shuffle(senior_f_list)

# 독서실 자리 배치
gg_1 = 0
gg_2 = 0
gg_3 = 0
for i in range(15):
    k = "북쪽라인" + str(i+1)
    a = student_xlsx[k]
    n_1 = int(a[0])
    n_2 = int(a[1])
    list_1 = []
    for m in range(15):
        if a[2 + m] == 0:
            list_1.append(0)
        elif a[2 + m] == 1:
            list_1.append(freshman_m_list[0])
            freshman_m_list.pop(0)
        elif a[2 + m] == 4:
            list_1.append(freshman_f_list[0])
            freshman_f_list.pop(0)
        elif a[2 + m] == 2:
            list_1.append(junior_m_list[0])
            junior_m_list.pop(0)
        elif a[2 + m] == 5:
            list_1.append(junior_f_list[0])
            junior_f_list.pop(0)
        elif a[2 + m] == 3:
            list_1.append(senior_m_list[0])
            senior_m_list.pop(0)
        elif a[2 + m] == 6:
            list_1.append(senior_f_list[0])
            senior_f_list.pop(0)
    seat_list.append(list_1)

# 데이터 정리
a = sorted(freshman_list, key=lambda x: x.num)
b = sorted(junior_list, key=lambda x: x.num)
c = sorted(senior_list, key=lambda x: x.num)
total = sorted(a + b + c,  key=lambda x: x.room)

# ui 불러오기


# 배치 내보내기
