import datetime
import glob
import logging
from logging.handlers import RotatingFileHandler
import os
import re
import math
import time
import numpy as np
import openpyxl as xl
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QDoubleValidator, QStandardItemModel, QIcon, QStandardItem, QIntValidator, QFont
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QProgressBar, QPlainTextEdit, QWidget, QGridLayout, QGroupBox, QLineEdit, QSizePolicy, QToolButton, QLabel, QFrame, QListView, QMenuBar, QStatusBar, QPushButton, QCalendarWidget, QVBoxLayout, QFileDialog, QComboBox
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QObject, QThread, QRect, QSize, QDate, QThreadPool
import pandas as pd
import cx_Oracle
import random
from collections import OrderedDict
from collections import defaultdict 
from multipledispatch import dispatch

class ThreadClass(QObject):
    OtherReturnError = pyqtSignal(Exception)
    OtherReturnInfo = pyqtSignal(str)
    OtherReturnWarning = pyqtSignal(str)
    OtherReturnEnd = pyqtSignal(bool)

    def __init__(self, 
                debugFlag,
                debugDate,
                cb_main,
                list_masterFile, 
                maxCnt,
                maxCnt_1, 
                emgHoldList):
        super().__init__()
        self.isDebug = debugFlag
        self.debugDate = debugDate
        self.cb_main = cb_main
        self.list_masterFile = list_masterFile
        self.maxCnt = maxCnt
        self.maxCnt_1 = maxCnt_1
        self.emgHoldList = emgHoldList

    #워킹데이 체크 내부함수
    def checkWorkDay(self, df, today, compDate):
        dtToday = pd.to_datetime(datetime.datetime.strptime(today, '%Y%m%d'))
        dtComp = pd.to_datetime(compDate, unit='s')
        workDay = 0
        #함수_수정-start
        index = int(df.index[(df['Date']==dtComp)].tolist()[0])
        while dtToday > pd.to_datetime(df['Date'][index], unit='s'):
            if df['WorkingDay'][index] == 1:
                workDay -= 1    
            index += 1
        #함수_수정-end
        for i in df.index:
            dt = pd.to_datetime(df['Date'][i], unit='s')
            if dtToday < dt and dt <= dtComp:
                if df['WorkingDay'][i] == 1:
                    workDay += 1
        return workDay

    #콤마 삭제용 내부함수
    def delComma(self, value):
        return str(value).split('.')[0]

    #하이픈 삭제
    def delHypen(self, value):
        return str(value).split('-')[0]
    

    #디비 불러오기 공통내부함수
    def readDB(self, ip, port, sid, userName, password, sql):
        location = r'C:\\instantclient_21_6'
        os.environ["PATH"] = location + ";" + os.environ["PATH"]
        dsn = cx_Oracle.makedsn(ip, port, sid)
        db = cx_Oracle.connect(userName, password, dsn)
        cursor= db.cursor()
        cursor.execute(sql)
        out_data = cursor.fetchall()
        df_oracle = pd.DataFrame(out_data)
        col_names = [row[0] for row in cursor.description]
        df_oracle.columns = col_names
        return df_oracle

    #생산시간 합계용 내부함수
    def getSec(self, time_str):
        time_str = re.sub(r'[^0-9:]', '', str(time_str))
        if len(time_str) > 0:
            h, m, s = time_str.split(':')
            return int(h) * 3600 + int(m) * 60 + int(s)
        else:
            return 0
    #백슬래쉬 삭제용 내부함수
    def delBackslash(self, value):
        value = re.sub(r"\\c", "", str(value))
        return value

    #알람 상세 누적 기록용 내부함수
    def concatAlarmDetail(self, df_target, no, category, df_data, index, smtAssy, shortageCnt,Gr):
        """
        Args:
            df_target(DataFrame)    : 알람상세내역 DataFrame
            no(int)                 : 알람 번호
            category(str)           : 알람 분류
            df_data(DataFrame)      : 원본 DataFrame
            index(int)              : 원본 DataFrame의 인덱스
            smtAssy(str)            : Smt Assy 이름
            shortageCnt(int)        : 부족 수량
            Gr(str)                 : 기종분류표 그룹
        Return:
            return(DataFrame)       : 알람상세 Merge결과 DataFrame 
        """
        if category == '1':
            return pd.concat([df_target, 
                                pd.DataFrame.from_records([{"No.":no,
                                                            "분류" : category,
                                                            "L/N" : df_data['Linkage Number'][index],
                                                            "MS CODE" : df_data['MS Code'][index], 
                                                            "SMT ASSY" : smtAssy, 
                                                            "수주수량" : df_data['미착공수주잔'][index], 
                                                            "부족수량" : shortageCnt, 
                                                            "검사호기" : '-', 
                                                            # "대상 검사시간(초)" : 0, 
                                                            # "필요시간(초)" : 0, 
                                                            "완성예정일" : df_data['Planned Prod. Completion date'][index]}
                                                            ])])
        elif category == '2': #분류2 - 부족수량은 부족한 카운트 맥스 값, 검사호기 - 맥스 그룹명
            return pd.concat([df_target, 
                                pd.DataFrame.from_records([{"No.":no,
                                                            "분류" : category,
                                                            "L/N" : df_data['Linkage Number'][index],
                                                            "MS CODE" : df_data['MS Code'][index], 
                                                            "SMT ASSY" : '-', 
                                                            "수주수량" : df_data['미착공수주잔'][index], 
                                                            "부족수량" : shortageCnt, 
                                                            "검사호기" : Gr, 
                                                            # "대상 검사시간(초)" : df_data['TotalTime'][index], 
                                                            # "필요시간(초)" : (df_data['미착공수주잔'][index] - df_data['설비능력반영_착공량'][index]) * df_data['TotalTime'][index], 
                                                            "완성예정일" : df_data['Planned Prod. Completion date'][index]}
                                                            ])])
        elif category == '기타1':
            return pd.concat([df_target, 
                                pd.DataFrame.from_records([{"No.":no,
                                                            "분류" : category,
                                                            "L/N" : df_data['Linkage Number'][index],
                                                            "MS CODE" : df_data['MS Code'][index], 
                                                            "SMT ASSY" : '미등록', 
                                                            "수주수량" : df_data['미착공수주잔'][index], 
                                                            "부족수량" : 0, 
                                                            "검사호기" : '-', 
                                                            # "대상 검사시간(초)" : 0, 
                                                            # "필요시간(초)" : 0, 
                                                            "완성예정일" : df_data['Planned Prod. Completion date'][index]}
                                                            ])])
        elif category == '기타2': #기타2, 부족수량은 최대착공량 부족한 수량
            return pd.concat([df_target, 
                                pd.DataFrame.from_records([{"No.":no,
                                                            "분류" : category,
                                                            "L/N" : df_data['Linkage Number'][index],
                                                            "MS CODE" : df_data['MS Code'][index], 
                                                            "SMT ASSY" : '-', 
                                                            "수주수량" : df_data['미착공수주잔'][index], 
                                                            "부족수량" : shortageCnt, 
                                                            "검사호기" : '-', 
                                                            # "대상 검사시간(초)" : 0, 
                                                            # "필요시간(초)" : 0, 
                                                            "완성예정일" : df_data['Planned Prod. Completion date'][index]}
                                                            ])])
        elif category == '기타3': #기타3, SMT ASSY에 part no가 나오도록한다.
            return pd.concat([df_target, 
                                pd.DataFrame.from_records([{"No.":no,
                                                            "분류" : category,
                                                            "L/N" : df_data['Linkage Number'][index],
                                                            "MS CODE" : df_data['MS Code'][index], 
                                                            "SMT ASSY" : smtAssy, 
                                                            "수주수량" : df_data['미착공수주잔'][index], 
                                                            "부족수량" : shortageCnt, 
                                                            "검사호기" : '-', 
                                                            # "대상 검사시간(초)" : 0, 
                                                            # "필요시간(초)" : 0, 
                                                            "완성예정일" : df_data['Planned Prod. Completion date'][index]}
                                                            ])])
                            
    def smtReflectInst(self, 
                        df_input, 
                        isRemain, 
                        dict_smtCnt, 
                        alarmDetailNo, 
                        df_alarmDetail, 
                        rowNo):
        """
        Args:
            df_input(DataFrame)         : 입력 DataFrame 
            isRemain(Bool)              : 잔여착공 여부 Flag
            dict_smtCnt(Dict)           : Smt잔여량 Dict
            alarmDetailNo(int)          : 알람 번호
            df_alarmDetail(DataFrame)   : 알람 상세 기록용 DataFrame
            rowNo(int)                  : 사용 Smt Assy 갯수
        Return:
            return(List)                 
                df_input(DataFrame)         : 입력 DataFrame (갱신 후)
                dict_smtCnt(Dict)           : Smt잔여량 Dict (갱신 후)
                alarmDetailNo(int)          : 알람 번호
                df_alarmDetail(DataFrame)   : 알람 상세 기록용 DataFrame (갱신 후)
        """
        instCol = '평준화_적용_착공량'
        resultCol = 'SMT반영_착공량'
        if isRemain:
            instCol = '잔여_착공량'
            resultCol = 'SMT반영_착공량_잔여'
        for i in df_input.index:
            # if df_input['PRODUCT_TYPE'][i] == 'MAIN' and 'CT' not in df_input['MS Code'][i]:
                for j in range(1,rowNo):
                    if j == 1:
                        rowCnt = 1
                    if str(df_input[f'ROW{str(j)}'][i]) != '' and str(df_input[f'ROW{str(j)}'][i]) != 'None':
                    # if str(df_input[f'ROW{str(j)}'][i]) != '' and str(df_input[f'ROW{str(j)}'][i]) != 'nan':
                        rowCnt = j
                    else:
                        break
                minCnt = 9999
                for j in range(1,rowCnt+1):
                        smtAssyName = str(df_input[f'ROW{str(j)}'][i])
                        if df_input['MS Code'][i] != 'nan' and df_input['MS Code'][i] != 'None' and df_input['MS Code'][i] != '':
                            if smtAssyName != '' and smtAssyName != 'None'and smtAssyName != 'nan':
                            # if smtAssyName != '' and smtAssyName != 'nan':
                                if df_input['긴급오더'][i] == '대상' or df_input['당일착공'][i] == '대상':
                                    if dict_smtCnt[smtAssyName] < 0:
                                        diffCnt = df_input['미착공수주잔'][i]
                                        if dict_smtCnt[smtAssyName] + df_input['미착공수주잔'][i] > 0:
                                            diffCnt = 0 - dict_smtCnt[smtAssyName]
                                        if not isRemain:
                                            df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                                    alarmDetailNo,
                                                                                    '1', 
                                                                                    df_input,
                                                                                    i, 
                                                                                    smtAssyName, 
                                                                                    diffCnt,
                                                                                    '-')
                                            alarmDetailNo += 1
                                else:
                                    if smtAssyName in dict_smtCnt:                                    
                                        if dict_smtCnt[smtAssyName] >= df_input[instCol][i]:
                                            if minCnt > df_input[instCol][i]:
                                                minCnt = df_input[instCol][i]
                                        else: 
                                            if dict_smtCnt[smtAssyName] > 0:
                                                if minCnt > dict_smtCnt[smtAssyName]:
                                                    minCnt = dict_smtCnt[smtAssyName]

                                            else:
                                                minCnt = 0
                                            if not isRemain:
                                                if dict_smtCnt[smtAssyName] > 0:
                                                    df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                                            alarmDetailNo,
                                                                                            '1', 
                                                                                            df_input,
                                                                                            i, 
                                                                                            smtAssyName, 
                                                                                            df_input[instCol][i] - dict_smtCnt[smtAssyName],
                                                                                            '-')
                                                    alarmDetailNo += 1

                                    else: #smt assy inven에 값이 없어서 df_smtCnt에 값이 없을때
                                         #기타3 smt assy inven - part no 없을때
                                        minCnt = 0
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                                alarmDetailNo,
                                                                                '기타3', 
                                                                                df_input,
                                                                                i, 
                                                                                smtAssyName, 
                                                                                0,
                                                                                '-')
                                        alarmDetailNo += 1
                                            

                        else:
                                df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                        alarmDetailNo,
                                                                        '기타1', 
                                                                        df_input,
                                                                        i, 
                                                                        '미등록', 
                                                                        0,
                                                                        '-')
                                alarmDetailNo += 1
                if minCnt != 9999:
                    df_input[resultCol][i] = minCnt
                else:
                    df_input[resultCol][i] = df_input[instCol][i]

                for j in range(1,rowCnt+1):
                    if smtAssyName != '' and smtAssyName != 'None' and smtAssyName != 'nan':
                    # if smtAssyName != '' and smtAssyName != 'nan':
                        smtAssyName = str(df_input[f'ROW{str(j)}'][i])
                        if smtAssyName in dict_smtCnt:
                            dict_smtCnt[smtAssyName] -= df_input[resultCol][i]
        return [df_input, dict_smtCnt, alarmDetailNo, df_alarmDetail]
    
    def count_emg(self,fix,max1,max2,undo,worktime,module_type,Gr): 
        """
        긴급오더: 대상
            1, 2차 max, 최대 착공량, 확정수량에서 설비능력 반영수량을 빼준다.
                Args:
                    fix: 설비능력 반영수량
                    max1: 1차 max
                    max2:  2차 max
                    undo: 미착공 수주량
                    worktime: 공수
                    module_type: 최대착공수량
                    Gr: 1차, 2차 유무
                Return:
                    fix: 설비능력 반영수량(갱신 후)
                    max1: 1차 max(갱신 후)
                    max2: 2ck max(갱신 후)
                    count_max: 최대착공수량(갱신 후)

        """
        count_max = module_type
        if str(Gr)  == '2':
            max2 -= undo
            max1 -= undo
            count_max -= undo * worktime
            fix = undo
            return(fix,max1,max2,count_max)
        elif str(Gr) == '1':
            max1 -= undo
            count_max -= undo * worktime
            fix = undo
            return(fix,max1,max2,count_max)
    
    def func_nonemg(self,fix,max1,max2,smt,worktime,module_type,Gr): 
        """
        긴급오더: 비대상
            1, 2차 max, 최대 착공량에서 설비능력 반영수량을 빼준다.
                Args:
                    fix: 설비능력 반영수량, 설비능력 반영수량_잔여
                    max1: 1차 max
                    max2:  2차 max
                    smt: smt 반영 착공수량
                    worktime: 공수
                    module_type: 최대착공수량
                    Gr: 1차, 2차 유무
                Return:
                    fix: 설비능력 반영수량, 설비능력 반영수량_잔여(갱신 후)
                    max1: 1차 max(갱신 후)
                    max2: 2차 max(갱신 후)
                    count_max: 최대착공수량(갱신 후)
        """
        count_max = module_type #외부 딕셔너리 value를 가져와서 함수 내부에서 계산하기 위한 목적
        if str(Gr) == '1':
            if max1 > smt:
                if count_max >= smt * worktime:
                    fix = smt
                else:
                    fix = math.floor(count_max / worktime)           
            else:
                if count_max >= max1 * worktime:
                    fix = max1
                else:
                    fix = math.floor(count_max / worktime)
            max1 -= fix
            count_max -= fix * worktime
            return(fix,max1,count_max)
        elif str(Gr) == '2':
            if max1 > smt:
                if count_max >= smt * worktime:
                    fix = smt
                else:
                    fix = math.floor(count_max / worktime)
            else:
                if count_max >= max1 * worktime:
                    fix = max1
                else:
                    fix = math.floor(count_max / worktime)
            max2 -= fix
            max1 -= fix
            count_max -= fix * worktime #외부 딕셔너리 계산할 경우 별도의 변수로 계산 필요할것
            return(fix,max1,max2,count_max) #fix : 설비능력 반영수량, max1 : 1차 max, module type : 최대 착공수량
    
    def func_nonemg2(self,fix,smt,worktime,module_type,Gr): #비 긴급오더, 기종분류표 비대상 or NC모듈인 경우 중복 구분 처리
        """
        긴급오더: 비대상
            최대 착공량에서 설비능력 반영수량(설비능력 반영수량_잔여)을 빼준다.
                Args:
                    fix: 설비능력 반영수량, 설비능력 반영수량_잔여
                    smt: smt 반영 착공수량
                    worktime: 공수
                    module_type: 최대착공수량
                    Gr: 1차, 2차 유무
                Return:
                    fix: 설비능력 반영수량, 설비능력 반영수량_잔여(갱신 후)
                    count_max: 최대착공수량(갱신 후)
        """
        count_max = module_type #외부 딕셔너리 value를 가져와서 함수 내부에서 계산하기 위한 목적   #딕셔너리 잘못 가져옴 수정 필요
        if str(Gr) == '2':
            if count_max >= smt:
                fix = smt
            else:
                fix = math.floor(count_max / worktime)
            count_max -= fix * worktime
            return(fix,count_max)
        elif str(Gr) == '1':
            if count_max >= smt:
                fix = smt
            else:
                fix = count_max
            count_max -= fix
            return(fix,count_max)
    def func_emg(self,fix,module_type,undo): #alarm detail 설정할 것
        """
        긴급오더: 대상,
            최대 착공량에서 미착공수주잔을 빼준다.
                Args:
                    fix: 설비능력 반영수량
                    module_type: 최대착공수량
                    undo: 미착공수량
                Return:
                    count_max: 최대착공수량(갱신 후)
                    fix: 설비능력 반영수량
                    alarmx: 알람 유무 카운트
        """
        alarmx = 0
        count_max = module_type
        count_max -= undo
        fix = undo
        if count_max - undo < 0:
            
            alarmx = 1
        return(count_max,fix,alarmx)
    #기타 함수 추가
    
    def run(self):
        #pandas 경고없애기 옵션 적용
        pd.set_option('mode.chained_assignment', None)
        try:
            #긴급오더, 홀딩오더 불러오기
            #사용자 입력값 불러오기, self.max_cnt
            module_loading = self.maxCnt
            nonmodule_loading = self.maxCnt_1
            emgLinkage = self.emgHoldList[0]
            emgmscode = self.emgHoldList[1]
            holdLinkage = self.emgHoldList[2]
            holdmscode = self.emgHoldList[3]
            #긴급오더, 홀딩오더 데이터프레임화
            df_emgLinkage = pd.DataFrame({'Linkage Number':emgLinkage})
            df_emgmscode = pd.DataFrame({'MS Code':emgmscode})
            df_holdLinkage = pd.DataFrame({'Linkage Number':holdLinkage})
            df_holdmscode = pd.DataFrame({'MS Code':holdmscode})
            #각 Linkage Number 컬럼의 타입을 일치시킴
            df_emgLinkage['Linkage Number'] = df_emgLinkage['Linkage Number'].astype(str)
            df_holdLinkage['Linkage Number'] = df_holdLinkage['Linkage Number'].astype(str)
            #긴급오더, 홍딩오더 Join 전 컬럼 추가
            df_emgLinkage['긴급오더'] = '대상'
            df_emgmscode['긴급오더'] = '대상'
            df_holdLinkage['홀딩오더'] = '대상'
            df_holdmscode['홀딩오더'] = '대상'
            #레벨링 리스트 불러오기
            # df_levelingMain = pd.read_excel(self.list_masterFile[1]) #레벨링 리스트 수정
            df_levelingSp = pd.read_excel(self.list_masterFile[2])
            

            #미착공 대상만 추출(특수_모듈)
            df_levelingSpDropSEQ = df_levelingSp[df_levelingSp['Sequence No'].isnull()]
            df_levelingSpUndepSeq = df_levelingSp[df_levelingSp['Sequence No']=='Undep']
            df_levelingSpUncorSeq = df_levelingSp[df_levelingSp['Sequence No']=='Uncor']
            df_levelingSp = pd.concat([df_levelingSpDropSEQ, df_levelingSpUndepSeq, df_levelingSpUncorSeq])
            df_levelingSp['모듈 구분'] = '모듈'
            df_levelingSp = df_levelingSp.reset_index(drop=True)
            df_levelingSp['미착공수주잔'] = df_levelingSp.groupby('Linkage Number')['Linkage Number'].transform('size')
            # df_levelingSp['미착공수량'] = df_levelingSp.groupby('Linkage Number')['Linkage Number'].transform('size')

            #비모듈 레벨링 리스트 불러오기 - 경로에 파일이 있으면 불러올것
            date_nonSp = datetime.datetime.today().strftime('%Y%m%d')
            if self.isDebug:
                date_nonSp = self.debugDate
                NonSp_BLFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date_nonSp +r'\\BL.xlsx'
                NonSp_TerminalFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date_nonSp +r'\\TERMINAL.xlsx'
            if os.path.exists(NonSp_BLFilePath): #파일이 없으면 빈 스트링
                df_levelingNonSp_BL = pd.read_excel(self.list_masterFile[9])
                df_leveling_NonSpDropSEQ_BL = df_levelingNonSp_BL[df_levelingNonSp_BL['Sequence No'].isnull()]
                df_leveling_NonSpUndepSeq_BL = df_levelingNonSp_BL[df_levelingNonSp_BL['Sequence No']=='Undep']
                df_leveling_NonSpUncorSeq_BL = df_levelingNonSp_BL[df_levelingNonSp_BL['Sequence No']=='Uncor']
                df_levelingNonSp_BL = pd.concat([df_leveling_NonSpDropSEQ_BL, df_leveling_NonSpUndepSeq_BL, df_leveling_NonSpUncorSeq_BL])
                df_levelingNonSp_BL['모듈 구분'] = '비모듈'
                df_levelingNonSp_BL = df_levelingNonSp_BL.reset_index(drop=True)
                df_levelingNonSp_BL['미착공수주잔'] = df_levelingNonSp_BL.groupby('Linkage Number')['Linkage Number'].transform('size')
                df_levelingSp = pd.concat([df_levelingSp, df_levelingNonSp_BL])
            
            if os.path.exists(NonSp_TerminalFilePath): #파일이 없으면 빈 스트링
                df_levelingNonSp_Terminal = pd.read_excel(self.list_masterFile[10])
                df_leveling_NonSpDropSEQ_Terminal = df_levelingNonSp_Terminal[df_levelingNonSp_Terminal['Sequence No'].isnull()]
                df_leveling_NonSpUndepSeq_Terminal = df_levelingNonSp_Terminal[df_levelingNonSp_Terminal['Sequence No']=='Undep']
                df_leveling_NonSpUncorSeq_Terminal = df_levelingNonSp_Terminal[df_levelingNonSp_Terminal['Sequence No']=='Uncor']
                df_levelingNonSp_Terminal = pd.concat([df_leveling_NonSpDropSEQ_Terminal, df_leveling_NonSpUndepSeq_Terminal, df_leveling_NonSpUncorSeq_Terminal])
                df_levelingNonSp_Terminal['모듈 구분'] = '비모듈'
                df_levelingNonSp_Terminal = df_levelingNonSp_Terminal.reset_index(drop=True)
                df_levelingNonSp_Terminal['미착공수주잔'] = df_levelingNonSp_Terminal.groupby('Linkage Number')['Linkage Number'].transform('size')
                df_levelingSp = pd.concat([df_levelingSp, df_levelingNonSp_Terminal])
                
            if self.isDebug:
                df_levelingSp.to_excel(r'd:\\FAM3_Leveling-1\\verification\\flow1.xlsx')

            df_sosFile = pd.read_excel(self.list_masterFile[0])
            # df_sosFile['Linkage Number'] = df_sosFile['Linkage Number'].astype(str)
            df_sosFile['Linkage Number'] = df_sosFile['Linkage Number'].astype(str)
            df_levelingSp['Linkage Number'] = df_levelingSp['Linkage Number'].astype(str)
            # if self.isDebug:
                # df_sosFile.to_excel('.\\debug\\flow2.xlsx')

            for i in df_sosFile.index:
                if df_sosFile['Material'][i] == 'S9307UF':
                    self.OtherReturnWarning.emit(f'SWITCH(S9307UF)의 수주잔이 확인되었습니다. 확인바랍니다.')
            
            #착공 대상 외 모델 삭제
            df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('ZOTHER')].index)
            df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('YZ')].index)
            df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('SF')].index)
            df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('KM')].index)
            df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('TA80')].index)
            df_sosFile = df_sosFile.drop(df_sosFile[df_sosFile['MS Code'].str.contains('CT')].index)

            if self.isDebug:
                df_sosFile.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow3.xlsx')

            #워킹데이 캘린더 불러오기
            dfCalendar = pd.read_excel(self.list_masterFile[4])
            today = datetime.datetime.today().strftime('%Y%m%d')
            if self.isDebug:
                today = self.debugDate

            # df_levelingSp.to_excel(r'd:\\FAM3_Leveling-1\\verification\\df_levelingSp.xlsx')
            print(type(df_sosFile['Linkage Number'][1]))
            print(type(df_levelingSp['Linkage Number'][1]))
            #진척 파일 - SOS2파일 Join
            # df_sosFileMerge = pd.merge(df_sosFile, df_progressFile, left_on='Linkage Number', right_on='LINKAGE NO', how='left').drop_duplicates(['Linkage Number'])
            df_sosFileMerge = pd.merge(df_sosFile, df_levelingSp).drop_duplicates(['Linkage Number'])
            df_sosFileMerge = df_sosFileMerge[['Linkage Number',
                                            'MS Code',
                                            'Planned Prod. Completion date',
                                            'Order Quantity',
                                            '미착공수주잔',
                                            '모듈 구분']]
            #위 파일을 완성지정일 기준 오름차순 정렬 및 인덱스 재설정
            df_sosFileMerge = df_sosFileMerge.sort_values(by=['Planned Prod. Completion date'],
                                                            ascending=[True])
            df_sosFileMerge = df_sosFileMerge.reset_index(drop=True)
            
            #대표모델 Column 생성
            df_sosFileMerge['대표모델'] = df_sosFileMerge['MS Code'].str[:9]
            #남은 워킹데이 Column 생성
            df_sosFileMerge['남은 워킹데이'] = 0
            #긴급오더, 홀딩오더 Linkage Number Column 타입 일치
            df_emgLinkage['Linkage Number'] = df_emgLinkage['Linkage Number'].astype(str)
            df_holdLinkage['Linkage Number'] = df_holdLinkage['Linkage Number'].astype(str)
            #긴급오더, 홀딩오더와 위 Sos파일을 Join
            
            # df_sosFileMerge.to_excel(r'd:\\FAM3_Leveling-1\\verification\\df_sosFileMerge.xlsx')
            # df_emgLinkage.to_excel(r'd:\\FAM3_Leveling-1\\verification\\df_emgLinkage.xlsx')
            df_MergeLink = pd.merge(df_sosFileMerge, df_emgLinkage, on='Linkage Number', how='left')
            dfMergemscode = pd.merge(df_sosFileMerge, df_emgmscode, on='MS Code', how='left')
            df_MergeLink = pd.merge(df_MergeLink, df_holdLinkage, on='Linkage Number', how='left')
            dfMergemscode = pd.merge(dfMergemscode, df_holdmscode, on='MS Code', how='left')
            df_MergeLink['긴급오더'] = df_MergeLink['긴급오더'].combine_first(dfMergemscode['긴급오더'])
            df_MergeLink['홀딩오더'] = df_MergeLink['홀딩오더'].combine_first(dfMergemscode['홀딩오더'])
            df_MergeLink['당일착공'] = ''

            # df_MergeLink.to_excel(r'd:\\FAM3_Leveling-1\\flow4-2.xlsx')
            for i in df_MergeLink.index:
                df_MergeLink['남은 워킹데이'][i] = self.checkWorkDay(dfCalendar, today, df_MergeLink['Planned Prod. Completion date'][i])
                if df_MergeLink['남은 워킹데이'][i] < 0:
                    df_MergeLink['긴급오더'][i] = '대상'
                elif df_MergeLink['남은 워킹데이'][i] == 0:
                    df_MergeLink['당일착공'][i] = '대상'
            df_MergeLink = df_MergeLink[df_MergeLink['미착공수주잔'] != 0]
            df_MergeLink['Linkage Number'] = df_MergeLink['Linkage Number'].astype(str)
            #MODEL 만들기
            df_MergeLink['MODEL'] = df_MergeLink['MS Code'].str[:7]
            df_MergeLink['MODEL'] = df_MergeLink['MODEL'].astype(str).apply(self.delHypen)
            
            # if self.isDebug:
            #     df_MergeLink.to_excel(r'd:\\FAM3_Leveling-1\\verification\\flow1_emg_test.xlsx')

            yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y%m%d')
            if self.isDebug:
                yesterday = (datetime.datetime.strptime(self.debugDate,'%Y%m%d') - datetime.timedelta(days=1)).strftime('%Y%m%d')

            #smt asssy 수량 불러오기 - 알람 test용도
            df_SmtAssyInven = self.readDB('10.36.15.42',
                                    1521,
                                    'NEURON',
                                    'ymi_user',
                                    'ymi123!',
                                    "SELECT INV_D, PARTS_NO, CURRENT_INV_QTY FROM pdsg0040 where INV_D = TO_DATE("+ str(yesterday) +",'YYYYMMDD')")
            # df_SmtAssyInven.columns = ['INV_D','PARTS_NO','CURRENT_INV_QTY']
            df_SmtAssyInven['현재수량'] = 0
            
            #알람 test용도
            df_joinSmt = df_SmtAssyInven.copy()
            
            dict_smtCnt = {}
            for i in df_joinSmt.index:
                dict_smtCnt[df_joinSmt['PARTS_NO'][i]] = df_joinSmt['CURRENT_INV_QTY'][i]
            
            #PB01: S9221DS, TA40: S9091BU 재고량 미확인 모델 dict_smtCnt 추가
            df_smtUnCheck = pd.read_excel(self.list_masterFile[8])
            for i in df_smtUnCheck.index:
                if df_smtUnCheck['SMT_ASSY'][i] != '' and 'nan' and 'None':
                    dict_smtCnt[df_smtUnCheck['SMT_ASSY'][i]] = df_smtUnCheck['수량'][i]
            
            df_productTime = self.readDB('ymzn-bdv19az029-rds.cgbtxsdj6fjy.ap-northeast-1.rds.amazonaws.com',
                                    1521,
                                    'TPROD',
                                    'TEST_SCM',
                                    'test_scm',
                                    'SELECT * FROM FAM3_PRODUCT_TIME_TB')
            df_productTime['TotalTime'] = df_productTime['COMPONENT_SET'].apply(self.getSec) + df_productTime['MAEDZUKE'].apply(self.getSec) + df_productTime['MAUNT'].apply(self.getSec) + df_productTime['LEAD_CUTTING'].apply(self.getSec) + df_productTime['VISUAL_EXAMINATION'].apply(self.getSec) + df_productTime['PICKUP'].apply(self.getSec) + df_productTime['ASSAMBLY'].apply(self.getSec) + df_productTime['M_FUNCTION_CHECK'].apply(self.getSec) + df_productTime['A_FUNCTION_CHECK'].apply(self.getSec) + df_productTime['PERSON_EXAMINE'].apply(self.getSec)
            df_productTime['대표모델'] = df_productTime['MODEL'].str[:9]
            df_productTime = df_productTime.drop_duplicates(['대표모델'])
            df_productTime = df_productTime.reset_index(drop=True)
            # df_productTime.to_excel(r'd:\\FAM3_Leveling-1\\flow7.xlsx')
            # print(df_productTime.columns)

            #MSCode_ASSY DB불러오기
            df_pdbs = self.readDB('10.36.15.42',
                                1521,
                                'neuron',
                                'ymfk_user',
                                'ymfk_user',
                                "SELECT SMT_MS_CODE, SMT_SMT_ASSY, SMT_CRP_GR_NO FROM sap.pdbs0010 WHERE SMT_CRP_GR_NO = '100L1304' or SMT_CRP_GR_NO = '100L1318' or SMT_CRP_GR_NO = '100L1331' or SMT_CRP_GR_NO = '100L1312' or SMT_CRP_GR_NO = '100L1303'") 
            
            # df_pdbs = df_pdbs.reset_index(drop=True)
            df_pdbs = df_pdbs[~df_pdbs['SMT_MS_CODE'].str.contains('AST')]
            df_pdbs = df_pdbs[~df_pdbs['SMT_MS_CODE'].str.contains('BMS')]
            df_pdbs = df_pdbs[~df_pdbs['SMT_MS_CODE'].str.contains('WEB')]
            # df_pdbs = df_pdbs[~df_del]

            gb = df_pdbs.groupby('SMT_MS_CODE')
            df_temp = pd.DataFrame([df_pdbs.loc[gb.groups[n], 'SMT_SMT_ASSY'].values for n in gb.groups], index=gb.groups.keys())
            df_temp.columns = ['ROW'+ str(i +1) for i in df_temp.columns]
            rowNo = len(df_temp.columns)
            df_temp = df_temp.reset_index()
            df_temp.rename(columns={'index':'MS Code'}, inplace=True)
            # df_temp.rename(columns={'index':'SMT_SMT_ASSY'}, inplace=True)
            # df_temp.rename(columns={'SMT_SMT_ASSY':'MS Code'}, inplace=True)
            # df_temp.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_temp_3_flow3.xlsx')
            # df_MergeLink.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_MergeLink_flow3.xlsx')
            # 컬럼명 변경필요

            df_addSmtAssy = pd.merge(df_MergeLink, df_temp, on='MS Code')
            # df_addSmtAssy = pd.merge(df_MergeLink, df_temp, left_on='MS Code',right_on='SMT_MS_CODE', how='left')
            # df_addSmtAssy = pd.merge(df_MergeLink, df_temp, on='SMT_MS_CODE', how='left')
            df_addSmtAssy = df_addSmtAssy.drop_duplicates(['Linkage Number'])
            df_addSmtAssy = df_addSmtAssy.reset_index(drop=True)
            

            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow9_thread_1.xlsx')

            df_addSmtAssy['대표모델별_최소착공필요량_per_일'] = 0
            dict_integCnt = {}
            dict_minContCnt = {}

            for i in df_addSmtAssy.index:
                if df_addSmtAssy['대표모델'][i] in dict_integCnt:
                    dict_integCnt[df_addSmtAssy['대표모델'][i]] += int(df_addSmtAssy['미착공수주잔'][i])
                else:
                    dict_integCnt[df_addSmtAssy['대표모델'][i]] = int(df_addSmtAssy['미착공수주잔'][i])

                if df_addSmtAssy['남은 워킹데이'][i] == 0:
                    workDay = 1
                else:
                    workDay = df_addSmtAssy['남은 워킹데이'][i]
                
                if len(dict_minContCnt) > 0:
                    if df_addSmtAssy['대표모델'][i] in dict_minContCnt:
                        if dict_minContCnt[df_addSmtAssy['대표모델'][i]][0] < math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay):
                            dict_minContCnt[df_addSmtAssy['대표모델'][i]][0] = math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay)
                            dict_minContCnt[df_addSmtAssy['대표모델'][i]][1] = df_addSmtAssy['Planned Prod. Completion date'][i]
                    else:
                        dict_minContCnt[df_addSmtAssy['대표모델'][i]] = [math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay),
                                                                    df_addSmtAssy['Planned Prod. Completion date'][i]]
                else:
                    dict_minContCnt[df_addSmtAssy['대표모델'][i]] = [math.ceil(dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay),
                                                                    df_addSmtAssy['Planned Prod. Completion date'][i]]
                if workDay <= 0:
                    workDay = 1
                df_addSmtAssy['대표모델별_최소착공필요량_per_일'][i] = dict_integCnt[df_addSmtAssy['대표모델'][i]]/workDay

            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow10.xlsx')
            
            dict_minContCopy = dict_minContCnt.copy()
            
            #평준화 적용
            df_addSmtAssy['평준화_적용_착공량'] = 0
            for i in df_addSmtAssy.index:
                if df_addSmtAssy['대표모델'][i] in dict_minContCopy:
                    if dict_minContCopy[df_addSmtAssy['대표모델'][i]][0] >= int(df_addSmtAssy['미착공수주잔'][i]):
                        df_addSmtAssy['평준화_적용_착공량'][i] = int(df_addSmtAssy['미착공수주잔'][i])
                        dict_minContCopy[df_addSmtAssy['대표모델'][i]][0] -= int(df_addSmtAssy['미착공수주잔'][i])
                    else:
                        df_addSmtAssy['평준화_적용_착공량'][i] = dict_minContCopy[df_addSmtAssy['대표모델'][i]][0]
                        dict_minContCopy[df_addSmtAssy['대표모델'][i]][0] = 0
            df_addSmtAssy['잔여_착공량'] = df_addSmtAssy['미착공수주잔'] - df_addSmtAssy['평준화_적용_착공량']

            df_smtCopy = pd.DataFrame(columns=df_addSmtAssy.columns)
            #당일착공 추가-1129
            df_addSmtAssy = df_addSmtAssy.sort_values(by=['긴급오더',
                                                            '당일착공',
                                                            'Planned Prod. Completion date',
                                                            '평준화_적용_착공량'],
                                                            ascending=[False,
                                                                        False,
                                                                        True,
                                                                        False])
            
            
            rowCnt = 0
            df_addSmtAssy = df_addSmtAssy[df_addSmtAssy['미착공수주잔'] != 0]
            df_addSmtAssy = df_addSmtAssy.reset_index(drop=True)
            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\zflow11_thread.xlsx')

            #긴급오더test용 data
            # df_addSmtAssy = pd.read_excel(r'd:\\FAM3_Leveling-1\\flow11_thread.xlsx')
            #SMT 잔여수량 적용
            df_addSmtAssy['SMT반영_착공량'] = 0
            df_alarmDetail = pd.DataFrame(columns=["No.", "분류", "L/N", "MS CODE", "SMT ASSY", "수주수량", "부족수량", "검사호기", "부족 MAX값", "부족 최대 착공량", "대상 검사시간(초)", "필요시간(초)", "완성예정일"])
            alarmDetailNo = 1
            df_addSmtAssy, dict_smtCnt, alarmDetailNo, df_alarmDetail = self.smtReflectInst(df_addSmtAssy, 
                                                                                            False, 
                                                                                            dict_smtCnt, 
                                                                                            alarmDetailNo,
                                                                                            df_alarmDetail,
                                                                                            rowNo)

            df_addSmtAssy['SMT반영_착공량_잔여'] = 0
            df_addSmtAssy, dict_smtCnt, alarmDetailNo, df_alarmDetail = self.smtReflectInst(df_addSmtAssy, 
                                                                                            True, 
                                                                                            dict_smtCnt, 
                                                                                            alarmDetailNo,
                                                                                            df_alarmDetail,
                                                                                            rowNo)
            
            # df_alarmDetail.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_alarmDetail_기타3.xlsx')
            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_addSmtAssy_기타3.xlsx')
            #특수 기종분류표 반영 착공 로직 start
            df_condition = pd.read_excel(self.list_masterFile[7])
            df_condition['No'] = df_condition['No'].fillna(method='ffill')
            df_condition['1차_MAX_그룹'] = df_condition['1차_MAX_그룹'].fillna(method='ffill')
            df_condition['2차_MAX_그룹'] = df_condition['2차_MAX_그룹'].fillna(method='ffill')
            df_condition['1차_MAX'] = df_condition['1차_MAX'].fillna(method='ffill')
            df_condition['2차_MAX'] = df_condition['2차_MAX'].fillna(method='ffill')

            
            # module_loading = float(self.spOrderinput.text())
            # module_loading = self.maxCnt
            # nonmodule_loading = float(self.spOrderinput_1.text())
            dict_capableCnt = defaultdict(list) #모델 : 공수
            dict_firstMaxCnt = defaultdict(list) #1차_MAX_그룹 : 1차_MAX
            dict_secondMaxCnt = defaultdict(list) #2차_MAX_그룹 : 2차_MAX
            dict_module = defaultdict(list) #모델 : 구분 
            dict_modelFirstGr = defaultdict(list) # 모델 : 1차_MAX_그룹
            dict_modelSecondGr = defaultdict(list) # 모델 : 2차_MAX_그룹
            dict_ModuleMax = defaultdict(list) # '모듈' : 최대 착공 수량, '비모듈' : 최대 착공 수량
            dict_ModuleMax = {'모듈':module_loading, '비모듈':nonmodule_loading, '비대상':999999} # '모듈' : 최대 착공 수량, '비모듈' : 최대 착공 수량  '비대상'이 필요한가?


            #딕셔너리 설정
            for i in df_condition.index:
                dict_capableCnt[df_condition['MODEL'][i]] = df_condition['공수'][i]
                dict_module[df_condition['MODEL'][i]] = df_condition['구분'][i] #모듈, 비모듈 구분 용도
                if df_condition['2차_MAX_그룹'][i] != '-':
                    dict_firstMaxCnt[df_condition['1차_MAX_그룹'][i]] = df_condition['1차_MAX'][i]
                    dict_modelFirstGr[df_condition['MODEL'][i]] = df_condition['1차_MAX_그룹'][i] #모델 : 1차 그룹 매칭 용도
                    dict_secondMaxCnt[df_condition['2차_MAX_그룹'][i]] = df_condition['2차_MAX'][i]
                    dict_modelSecondGr[df_condition['MODEL'][i]] = df_condition['2차_MAX_그룹'][i] #모델 : 2차 그룹 매칭 용도
                elif df_condition['1차_MAX_그룹'][i] != '-':
                    dict_firstMaxCnt[df_condition['1차_MAX_그룹'][i]] = df_condition['1차_MAX'][i]
                    dict_modelFirstGr[df_condition['MODEL'][i]] = df_condition['1차_MAX_그룹'][i]
                else:
                    continue
            

            df_addSmtAssy = df_addSmtAssy.reset_index(drop=True) #index 리셋
            df_addSmtAssy['설비능력 반영수량'] = 0
            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\zflow12_thread.xlsx')
            #r기종분류표 기준 착공확정수량 계산 시~작
            for i in df_addSmtAssy.index:
                if df_addSmtAssy['긴급오더'][i] == '대상':
                    if df_addSmtAssy['MODEL'][i] in dict_module.keys(): #기종분류표에 model이 있는가
                        if df_addSmtAssy['MODEL'][i] in dict_modelSecondGr.keys(): #2차 max값 유무
                            df_addSmtAssy['설비능력 반영수량'][i],dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.count_emg(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                                                                                                                                                    dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                    dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                    df_addSmtAssy['미착공수주잔'][i],
                                                                                                                                                                                                                                                    dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                                                                                                                                    dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                                                                                                                                    2)
                            if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] > df_addSmtAssy['미착공수주잔'][i] * dict_capableCnt[df_addSmtAssy['MODEL'][i]]: #최대 착공량 > 미착공 수주량 * 공수
                                if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] - df_addSmtAssy['미착공수주잔'][i] >= 0: #2차 max - 미착공 수주량 >= 0
                                    if dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] > 0: #1차 max > 0
                                        continue
                                    else: #1차 max
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                    dict_modelFirstGr[df_addSmtAssy['MODEL'][i]])
                                        alarmDetailNo += 1
                                else:
                                    #분류2 - 2차 max 알람처리
                                    df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],
                                                                    dict_modelSecondGr[df_addSmtAssy['MODEL'][i]])
                                    alarmDetailNo += 1                                 
                            else:
                                if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] - df_addSmtAssy['미착공수주잔'][i] >= 0: #2차 max - 미착공 수주량 >= 0
                                    if dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] > 0: #1차 max > 0
                                        #기타2 - 최대 착공량 알람 처리
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '기타2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                    '-')
                                        alarmDetailNo += 1  
                                    else:
                                        #기타2 - 최대 착공량, 분류2 - 1차 max 알람 처리
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '기타2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                    '-')
                                        alarmDetailNo += 1  
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                    '-')
                                        alarmDetailNo += 1  
                                else:
                                    #기타2 - 최대 착공량, 분류2 - 2차 max 알람 처리
                                    df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '기타2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                    '-')
                                    alarmDetailNo += 1                                  
                                    df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                alarmDetailNo,
                                                                '2', 
                                                                df_addSmtAssy,
                                                                i, 
                                                                '-', 
                                                                df_addSmtAssy['미착공수주잔'][i] - dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],
                                                                dict_modelSecondGr[df_addSmtAssy['MODEL'][i]])
                                    alarmDetailNo += 1                              
                        else:
                            if df_addSmtAssy['MODEL'][i] in dict_modelFirstGr.keys(): #1차 max값 유무
                                df_addSmtAssy['설비능력 반영수량'][i],dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.count_emg(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                                                                                                                                                    dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                    dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                    df_addSmtAssy['미착공수주잔'][i],
                                                                                                                                                                                                                                                    dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                                                                                                                                    dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                                                                                                                                    1)
                                if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] - df_addSmtAssy['미착공수주잔'][i] * dict_capableCnt[df_addSmtAssy['MODEL'][i]] > 0: #최대 착공량 - 미착공 수주량 * 공수 
                                    if dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] - df_addSmtAssy['미착공수주잔'][i] > 0: #1차 max - 미착공 수량 > 0
                                        continue
                                    else:
                                        #분류2 - 1차 max 알람처리
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                    dict_modelFirstGr[df_addSmtAssy['MODEL'][i]])
                                        alarmDetailNo += 1                              
                                else:
                                    if dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] - df_addSmtAssy['미착공수주잔'][i] > 0: #1차 max - 미착공 수량 > 0
                                        #기타2 - 최대 착공량 알람 처리
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '기타2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                    '-')
                                        alarmDetailNo += 1                              
                                    else:
                                        #기타2 - 최대 착공량, 분류2 - 1차 max 알람 처리
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '기타2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                    '-')
                                        alarmDetailNo += 1                              
                                        df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                    dict_modelFirstGr[df_addSmtAssy['MODEL'][i]])
                                        alarmDetailNo += 1  
                            else: #1차 max값이 없는 경우
                                dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],df_addSmtAssy['설비능력 반영수량'][i],alarmx = self.func_emg(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                        dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                        df_addSmtAssy['미착공수주잔'][i])
                                if alarmx == 1:
                                    df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                    alarmDetailNo,
                                                                    '기타2', 
                                                                    df_addSmtAssy,
                                                                    i, 
                                                                    '-', 
                                                                    df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                    '-')
                                    alarmDetailNo += 1

                    else:
                        dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],df_addSmtAssy['설비능력 반영수량'][i],alarmx = self.func_emg(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                        dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                        df_addSmtAssy['미착공수주잔'][i])
                        if alarmx == 1:
                                df_alarmDetail = self.concatAlarmDetail(df_alarmDetail,
                                                                alarmDetailNo,
                                                                '기타2', 
                                                                df_addSmtAssy,
                                                                i, 
                                                                '-', 
                                                                df_addSmtAssy['미착공수주잔'][i] - dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                '-')
                                alarmDetailNo += 1 
            
                else:# 긴급 오더가 아닌 경우
                    if df_addSmtAssy['MODEL'][i] in dict_module.keys(): #기종분류표에 model이 있는가
                        if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] > 0: #최대 착공 수량 > 0
                            if df_addSmtAssy['MODEL'][i] in dict_modelSecondGr.keys(): #2차 max값이 있는가
                                if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] > 0: #2차 max > 0?
                                    if dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] > 0: #1차 max > 0?
                                        if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] >= df_addSmtAssy['SMT반영_착공량'][i]: #2차 max >= smt 반영 착공량[i]
                                            # print(dict_secondMaxCnt)
                                            # print(dict_firstMaxCnt)
                                            df_addSmtAssy['설비능력 반영수량'][i],dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                                                                                                                                                                    dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                                    dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                                    df_addSmtAssy['SMT반영_착공량'][i],
                                                                                                                                                                                                                                                                    dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                                                                                                                                                    dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                                                                                                                                                    2)#2
                                            # print(dict_firstMaxCnt)
                                            # print(dict_secondMaxCnt)
                                        else: # 2차 max가 smt 반영 착공량[i] 보다 작을때
                                            if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] > dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]]: #2차 MAX > 1차  MAX
                                                count_temp = dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]]
                                            else:
                                                count_temp = dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]]
                                            if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] >= count_temp * dict_capableCnt[df_addSmtAssy['MODEL'][i]]: #최대 착공수량 >= temp * 공수
                                                df_addSmtAssy['설비능력 반영수량'][i] = count_temp
                                            else:
                                                df_addSmtAssy['설비능력 반영수량'][i] = math.floor(dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]]/dict_capableCnt[df_addSmtAssy['MODEL'][i]]) #설비능력 반영수량 = 최대착공수량/공수
                                            dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['설비능력 반영수량'][i] #2차 max -= 착공확정수량
                                            dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['설비능력 반영수량'][i] #1차 max -= 착공확정수량
                                            dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] -= df_addSmtAssy['설비능력 반영수량'][i] * dict_capableCnt[df_addSmtAssy['MODEL'][i]] #최대 착공 수량 -= 설비능력 반영수량 * 공수
                                        if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] == 0: #모듈 최대 착공 수량이 0이면 for문 종료
                                                break
                                        else:
                                            continue  
                                    else:
                                        continue # 1차 max 부족으로 다음 i로
                                else:
                                    continue # 2차 max 부족으로 다음 i로
                            else:
                                # print(type(df_addSmtAssy['설비능력 반영수량'][i])) #numpy.int64
                                # # print(type(dict_firstMaxCnt['F3RP6']))
                                # # print(dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]])
                                # print(type(0)) #int
                                # print(type(df_addSmtAssy['SMT반영_착공량'][i])) #numpy.int64
                                # print(type(dict_capableCnt[df_addSmtAssy['MODEL'][i]])) #numpy.float64 공수에 따라서 변경될것 같은데
                                # print(type(dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]])) #numpy.float64
                                if df_addSmtAssy['MODEL'][i] in dict_modelFirstGr.keys(): #1차 max값이 있는가
                                    # print(dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]])
                                    df_addSmtAssy['설비능력 반영수량'][i],dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                                                                                            dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                            0,
                                                                                                                                                                                            df_addSmtAssy['SMT반영_착공량'][i],
                                                                                                                                                                                            dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                                                                            dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                                                                            1) #2차 max 빈값으로 넣기, 1
                                else:
                                    df_addSmtAssy['설비능력 반영수량'][i],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg2(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                                    df_addSmtAssy['SMT반영_착공량'][i], 
                                                                                                                                    dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                    dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                    2)
                    else: #기종분류표에 모델이 없는 경우
                        df_addSmtAssy['설비능력 반영수량'][i],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg2(df_addSmtAssy['설비능력 반영수량'][i],
                                                                                                                        df_addSmtAssy['SMT반영_착공량'][i],
                                                                                                                        dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                        dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                        1)
                        if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] == 0: #모듈 최대 착공 수량이 0이면 for문 종료
                                break
                        else:
                            continue 
                    
                if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] == 0:
                    break
                else:
                    continue
            
            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\overloading_test.xlsx')
            
            #=========='SMT반영_착공량_잔여'에도 설비능력 적용
            # 설비능력 반영수량 >> 설비능력 반영수량_잔여 
            df_addSmtAssy['설비능력 반영수량_잔여'] = 0
            for i in df_addSmtAssy.index:
                if df_addSmtAssy['긴급오더'][i] == '대상':
                    continue
                else:
                    if df_addSmtAssy['MODEL'][i] in dict_module.keys(): #기종분류표에 model이 있는가
                        if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] > 0: #최대 착공 수량 > 0
                            if df_addSmtAssy['MODEL'][i] in dict_modelSecondGr.keys(): #2차 max값이 있는가
                                if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] > 0: #2차 max > 0?
                                    if dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] > 0: #1차 max > 0?
                                        if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] >= df_addSmtAssy['SMT반영_착공량_잔여'][i]: #2차 max >= smt 반영 착공량[i]
                                            # print(dict_secondMaxCnt)
                                            # print(dict_firstMaxCnt)
                                            df_addSmtAssy['설비능력 반영수량_잔여'][i],dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg(df_addSmtAssy['설비능력 반영수량_잔여'][i],
                                                                                                                                                                                                                                                                    dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                                    dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                                                                                                    df_addSmtAssy['SMT반영_착공량_잔여'][i],
                                                                                                                                                                                                                                                                    dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                                                                                                                                                    dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                                                                                                                                                    2)#2
                                            # print(dict_firstMaxCnt)
                                            # print(dict_secondMaxCnt)
                                        else: # 2차 max가 smt 반영 착공량[i] 보다 작을때
                                            if dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] > dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]]: #2차 MAX > 1차  MAX
                                                count_temp = dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]]
                                            else:
                                                count_temp = dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]]
                                            if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] >= count_temp * dict_capableCnt[df_addSmtAssy['MODEL'][i]]: #최대 착공수량 >= temp * 공수
                                                df_addSmtAssy['설비능력 반영수량_잔여'][i] = count_temp
                                            else:
                                                df_addSmtAssy['설비능력 반영수량_잔여'][i] = math.floor(dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]]/dict_capableCnt[df_addSmtAssy['MODEL'][i]]) #설비능력 반영수량_잔여 = 최대착공수량/공수
                                            dict_secondMaxCnt[dict_modelSecondGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['설비능력 반영수량_잔여'][i] #2차 max -= 착공확정수량
                                            dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]] -= df_addSmtAssy['설비능력 반영수량_잔여'][i] #1차 max -= 착공확정수량
                                            dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] -= df_addSmtAssy['설비능력 반영수량_잔여'][i] * dict_capableCnt[df_addSmtAssy['MODEL'][i]] #최대 착공 수량 -= 설비능력 반영수량_잔여 * 공수
                                        if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] == 0: #모듈 최대 착공 수량이 0이면 for문 종료
                                                break
                                        else:
                                            continue  
                                    else:
                                        continue # 1차 max 부족으로 다음 i로
                                else:
                                    continue # 2차 max 부족으로 다음 i로
                            else:
                                if df_addSmtAssy['MODEL'][i] in dict_modelFirstGr.keys(): #1차 max값이 있는가
                                    df_addSmtAssy['설비능력 반영수량_잔여'][i],dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg(df_addSmtAssy['설비능력 반영수량_잔여'][i],
                                                                                                                                                                                            dict_firstMaxCnt[dict_modelFirstGr[df_addSmtAssy['MODEL'][i]]],
                                                                                                                                                                                            0,
                                                                                                                                                                                            df_addSmtAssy['SMT반영_착공량_잔여'][i],
                                                                                                                                                                                            dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                                                                            dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                                                                            1) #2차 max 빈값으로 넣기, 1
                                else:
                                    df_addSmtAssy['설비능력 반영수량_잔여'][i],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg2(df_addSmtAssy['설비능력 반영수량_잔여'][i],
                                                                                                                                    df_addSmtAssy['SMT반영_착공량_잔여'][i], 
                                                                                                                                    dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                                    dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                                    2)
                    else: #기종분류표에 모델이 없는 경우
                        df_addSmtAssy['설비능력 반영수량_잔여'][i],dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] = self.func_nonemg2(df_addSmtAssy['설비능력 반영수량_잔여'][i],
                                                                                                                        df_addSmtAssy['SMT반영_착공량_잔여'][i],
                                                                                                                        dict_capableCnt[df_addSmtAssy['MODEL'][i]],
                                                                                                                        dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]],
                                                                                                                        1)
                        if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] == 0: #모듈 최대 착공 수량이 0이면 for문 종료
                                break
                        else:
                            continue 
                    
                if dict_ModuleMax[df_addSmtAssy['모듈 구분'][i]] == 0:
                    break
                else:
                    continue
            
            #==============================알람 처리 시~작!========================================================
            df_addSmtAssy = df_addSmtAssy.reset_index(drop=True)
            # if self.isDebug:
            #     df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow13.xlsx')
            #     df_alarmDetail = df_alarmDetail.reset_index(drop=True)
            #     df_alarmDetail.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_alarmDetail.xlsx')

            df_firstAlarm = df_alarmDetail[df_alarmDetail['분류'] == '1']
            df_firstAlarmSummary = df_firstAlarm.groupby("SMT ASSY")['부족수량'].sum()
            df_firstAlarmSummary = df_firstAlarmSummary.reset_index()
            df_firstAlarmSummary['부족수량'] = df_firstAlarmSummary['부족수량']
            # df_firstAlarmSummary['부족수량'] = df_firstAlarmSummary['부족수량']
            df_firstAlarmSummary['분류'] = '1'
            df_firstAlarmSummary['MS CODE'] = '-'
            df_firstAlarmSummary['검사호기(그룹)'] = '-'
            df_firstAlarmSummary['부족 시간'] = '-'
            df_firstAlarmSummary['Message'] = '[SMT ASSY : '+ df_firstAlarmSummary["SMT ASSY"]+ ']가 부족합니다. SMT ASSY 제작을 지시해주세요.'
            # del df_firstAlarmSummary['부족수량']

            df_secAlarmSummary = df_alarmDetail[df_alarmDetail['분류'] == '2']
            # df_secAlarmSummary = df_secAlarm.groupby("검사호기")['필요시간(초)'].sum()
            df_secAlarmSummary = df_secAlarmSummary.reset_index()
            df_secAlarmSummary['부족 시간'] = '-'
            df_secAlarmSummary['분류'] = '2'
            df_secAlarmSummary['MS CODE'] = '-'
            df_secAlarmSummary['SMT ASSY'] = '-'
            df_secAlarmSummary['검사호기(그룹)'] = df_secAlarmSummary['검사호기']
            df_secAlarmSummary['부족수량'] = df_secAlarmSummary['부족수량']
            df_secAlarmSummary['Message'] = 'MAX값이 부족합니다. 생산 가능여부를 확인해 주세요.'
            del df_secAlarmSummary['필요시간(초)']

            df_alarmSummary = pd.concat([df_firstAlarmSummary, df_secAlarmSummary])

            df_etcList = df_alarmDetail[(df_alarmDetail['분류'] == '기타1') | (df_alarmDetail['분류'] == '기타2') | (df_alarmDetail['분류'] == '기타3')]
            df_etcList = df_etcList.drop_duplicates(['MS CODE'])

            df_etcList = df_etcList.reset_index()
            # df_etcList.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_etcList.xlsx')
            for i in df_etcList.index:
                if df_etcList['분류'][i] == '기타1':
                    df_alarmSummary = pd.concat([df_alarmSummary, 
                                                pd.DataFrame.from_records([{"분류" : df_etcList['분류'][i],
                                                                            "MS CODE" : df_etcList['MS CODE'][i],
                                                                            "SMT ASSY" : '-', 
                                                                            "부족수량" : 0,
                                                                            "검사호기" : '-', 
                                                                            "부족 시간" : 0, 
                                                                            "Message" : '해당 MS CODE에서 사용되는 SMT ASSY가 등록되지 않았습니다. 등록 후 다시 실행해주세요.'}
                                                                            ])])
                elif df_etcList['분류'][i] == '기타2':
                    df_alarmSummary = pd.concat([df_alarmSummary, 
                                                pd.DataFrame.from_records([{"분류" : df_etcList['분류'][i],
                                                                            "MS CODE" : df_etcList['MS CODE'][i],
                                                                            "SMT ASSY" : '-', 
                                                                            "부족수량" : df_etcList['부족수량'][i],
                                                                            "검사호기" : '-', 
                                                                            "부족 시간" : 0, 
                                                                            "Message" : '긴급오더 및 당일착공 대상의 총 착공량이 입력한 최대착공량보다 큽니다. 최대착공량을 확인해주세요.'}
                                                                            ])])            
                elif df_etcList['분류'][i] == '기타3':
                    df_alarmSummary = pd.concat([df_alarmSummary, 
                                                pd.DataFrame.from_records([{"분류" : df_etcList['분류'][i],
                                                                            "MS CODE" : df_etcList['MS CODE'][i],
                                                                            "SMT ASSY" : df_etcList['SMT ASSY'][i], 
                                                                            "수량" : 0,
                                                                            "검사호기" : '-', 
                                                                            "부족 시간" : 0, 
                                                                            "Message" : 'SMT ASSY 정보가 등록되지 않아 재고를 확인할 수 없습니다. 등록 후 다시 실행해주세요.'}
                                                                            ])])
         
            df_alarmSummary = df_alarmSummary.reset_index(drop=True)
            df_alarmSummary = df_alarmSummary[['분류', 'MS CODE', 'SMT ASSY', '부족수량', '검사호기(그룹)', '부족 시간', 'Message']]
            if self.isDebug:
                df_alarmSummary.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_alarmSummary.xlsx')
                df_alarmDetail.to_excel(r'd:\\FAM3_Leveling-1\\test\\df_alarmDetail.xlsx')
            with pd.ExcelWriter(r'd:\\FAM3_Leveling-1\\Output\\Alarm\\FAM3_AlarmList_'+today+'.xlsx') as writer:
                df_alarmSummary.to_excel(writer, sheet_name='정리', index=True)
                df_alarmDetail.to_excel(writer, sheet_name='상세', index=True)   
            
            #================================총착공량 계산=====================================
            df_addSmtAssy['총착공량'] = df_addSmtAssy['설비능력 반영수량'] + df_addSmtAssy['설비능력 반영수량_잔여']
            df_addSmtAssy = df_addSmtAssy[df_addSmtAssy['총착공량'] != 0]

            # df_addSmtAssy.to_excel(r'd:\\FAM3_Leveling-1\\test\\zflow14.xlsx')
            #최대착공량만큼 착공 못했을 경우, 메시지 출력
            if dict_ModuleMax['모듈'] > 0 :
                remain = dict_ModuleMax['모듈']
                self.OtherReturnWarning.emit(f'아직 착공하지 못한 특수(모듈)이 [{int(remain)}대] 남았습니다. 설비능력 부족이 예상됩니다. 확인해주세요.')
            if dict_ModuleMax['비모듈'] > 0 :
                remain = dict_ModuleMax['비모듈']
                self.OtherReturnWarning.emit(f'아직 착공하지 못한 특수(비모듈)이 [{int(remain)}대] 남았습니다. 설비능력 부족이 예상됩니다. 확인해주세요.')

            #=====================================================사이클 로직 수정 해야함===============================================================
            #레벨링 리스트와 병합
            df_addSmtAssy = df_addSmtAssy.astype({'Linkage Number':'str'})
            df_levelingSp = df_levelingSp.astype({'Linkage Number':'str'})
            df_mergeOrder = pd.merge(df_addSmtAssy, df_levelingSp, on='Linkage Number', how='left')
            # if self.isDebug:
            #     df_mergeOrder.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow15.xlsx')
            df_mergeOrderResult = pd.DataFrame().reindex_like(df_mergeOrder)
            df_mergeOrderResult = df_mergeOrderResult[0:0]
            #총착공량 만큼 개별화
            for i in df_addSmtAssy.index:
                for j in df_mergeOrder.index:
                    if df_addSmtAssy['Linkage Number'][i] == df_mergeOrder['Linkage Number'][j]:
                        if j > 0:
                            if df_mergeOrder['Linkage Number'][j] != df_mergeOrder['Linkage Number'][j-1]:
                                orderCnt = int(df_addSmtAssy['총착공량'][i])
                        else:
                            orderCnt = int(df_addSmtAssy['총착공량'][i])
                        if orderCnt > 0:
                            df_mergeOrderResult = df_mergeOrderResult.append(df_mergeOrder.iloc[j])
                            # df_mergeOrderResult = pd.concat([df_mergeOrderResult, df_mergeOrder.iloc[j]])
                            orderCnt -= 1

            #사이클링을 위해 검사설비별로 정리
            df_mergeOrderResult = df_mergeOrderResult.sort_values(by=['대표모델'],
                                                                    ascending=[False])
            df_mergeOrderResult = df_mergeOrderResult.reset_index(drop=True)
            # if self.isDebug:
            #     df_mergeOrderResult.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow15-3.xlsx')
            #긴급오더 제외하고 사이클 대상만 식별하여 검사장치별로 갯수 체크
            df_cycleCopy = df_mergeOrderResult[df_mergeOrderResult['긴급오더'].isnull()]
            df_cycleCopy['검사장치Cnt'] = df_cycleCopy.groupby('대표모델')['대표모델'].transform('size')
            # print(df_cycleCopy['검사장치Cnt'])
            df_cycleCopy = df_cycleCopy.sort_values(by=['검사장치Cnt'],
                                                    ascending=[False])
            df_cycleCopy = df_cycleCopy.reset_index(drop=True)
            #긴급오더 포함한 Df와 병합
            df_mergeOrderResult = pd.merge(df_mergeOrderResult, df_cycleCopy[['Planned Order', '검사장치Cnt']], on='Planned Order', how='left')
            df_mergeOrderResult = df_mergeOrderResult.sort_values(by=['검사장치Cnt'],
                                                                    ascending=[False])
            df_mergeOrderResult = df_mergeOrderResult.reset_index(drop=True)
            # if self.isDebug:
            #     df_mergeOrderResult.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow15-4.xlsx')
            
            #=============사이클 test
            #최대 사이클 번호 체크
            maxCycle = float(df_cycleCopy['검사장치Cnt'][0])
            cycleGr = 1.0
            df_mergeOrderResult['사이클그룹'] = 0
            #각 검사장치별로 사이클 그룹을 작성하고, 최대 사이클과 비교하여 각 사이클그룹에서 배수처리
            for i in df_mergeOrderResult.index:
                if df_mergeOrderResult['긴급오더'][i] != '대상':
                    multiCnt = maxCycle/df_mergeOrderResult['검사장치Cnt'][i] #KJ_정수처리가 아니라 소수로도 처리가 가능하도록 한다
                    if i == 0:
                        df_mergeOrderResult['사이클그룹'][i] = cycleGr
                    else:
                        if df_mergeOrderResult['대표모델'][i] != df_mergeOrderResult['대표모델'][i-1]:
                            if i == 2:
                                cycleGr = 2.0 
                            else:
                                cycleGr = 1.0 
                        df_mergeOrderResult['사이클그룹'][i] = cycleGr * multiCnt
                    cycleGr += 1.0
                if cycleGr >= maxCycle:
                    cycleGr = 1.0
            #배정된 사이클 그룹 순으로 정렬
            df_mergeOrderResult = df_mergeOrderResult.sort_values(by=['사이클그룹'],
                                                                        ascending=[True])
            df_mergeOrderResult = df_mergeOrderResult.reset_index(drop=True)
            #불필요
            maxCycleNo = int(df_mergeOrderResult['사이클그룹'][len(df_mergeOrderResult)-1])
            # if self.isDebug:
            #     df_mergeOrderResult.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow16_cycle.xlsx')
            #===========사이클 test
            df_mergeOrderResult = df_mergeOrderResult.reset_index()
            for i in df_mergeOrderResult.index:
                if df_mergeOrderResult['긴급오더'][i] != '대상':
                    if i != 0 and (df_mergeOrderResult['대표모델'][i] == df_mergeOrderResult['대표모델'][i-1]):
                        for j in df_mergeOrderResult.index:
                            if (j != 0 and j < len(df_mergeOrderResult)-1) and (df_mergeOrderResult['대표모델'][i] != df_mergeOrderResult['대표모델'][j + 1]) and (df_mergeOrderResult['대표모델'][i] != df_mergeOrderResult['대표모델'][j]):
                                df_mergeOrderResult['index'][i] = (float(df_mergeOrderResult['index'][j]) + float(df_mergeOrderResult['index'][j+1]))/2
                                df_mergeOrderResult = df_mergeOrderResult.sort_values(by=['index'],
                                                                                      ascending=[True])
                                df_mergeOrderResult = df_mergeOrderResult.reset_index(drop=True)
                                break
            df_mergeOrderResult = df_mergeOrderResult.reset_index()
            for i in df_mergeOrderResult.index:
                if df_mergeOrderResult['긴급오더'][i] != '대상':
                    if i != 0 and (df_mergeOrderResult['대표모델'][i] == df_mergeOrderResult['대표모델'][i-1]):
                        for j in df_mergeOrderResult.index:
                            if (j != 0 and j < len(df_mergeOrderResult)-1) and (df_mergeOrderResult['대표모델'][i] != df_mergeOrderResult['대표모델'][j + 1]) and (df_mergeOrderResult['대표모델'][i] != df_mergeOrderResult['대표모델'][j]):
                                df_mergeOrderResult['index'][i] = (float(df_mergeOrderResult['index'][j]) + float(df_mergeOrderResult['index'][j+1]))/2
                                df_mergeOrderResult = df_mergeOrderResult.sort_values(by=['index'],
                                                                                      ascending=[True])
                                df_mergeOrderResult = df_mergeOrderResult.reset_index(drop=True)
                                break
            
            

            if self.isDebug:
                df_mergeOrderResult.to_excel(r'd:\\FAM3_Leveling-1\\test\\flow17-5.xlsx')
            
            df_mergeOrderResult['No (*)'] = (df_mergeOrderResult.index.astype(int) + 1) * 10
            df_mergeOrderResult['Planned Order'] = df_mergeOrderResult['Planned Order'].astype(int).astype(str).str.zfill(10)
            df_mergeOrderResult['Scheduled End Date'] = df_mergeOrderResult['Scheduled End Date'].astype(str).str.zfill(10)
            df_mergeOrderResult['Specified Start Date'] = df_mergeOrderResult['Specified Start Date'].astype(str).str.zfill(10)
            df_mergeOrderResult['Specified End Date'] = df_mergeOrderResult['Specified End Date'].astype(str).str.zfill(10)
            df_mergeOrderResult['Spec Freeze Date'] = df_mergeOrderResult['Spec Freeze Date'].astype(str).str.zfill(10)
            df_mergeOrderResult['Component Number'] = df_mergeOrderResult['Component Number'].astype(int).astype(str).str.zfill(4)

            df_mergeOrderResult = df_mergeOrderResult[['No (*)', 
                                                        'Sequence No', 
                                                        'Production Order', 
                                                        'Planned Order', 
                                                        'Manual', 
                                                        'Scheduled Start Date (*)', 
                                                        'Scheduled End Date', 
                                                        'Specified Start Date', 
                                                        'Specified End Date', 
                                                        'Demand destination country', 
                                                        'MS-CODE', 
                                                        'Allocate', 
                                                        'Spec Freeze Date', 
                                                        'Linkage Number', 
                                                        'Order Number', 
                                                        'Order Item', 
                                                        'Combination flag', 
                                                        'Project Definition', 
                                                        'Error message', 
                                                        'Leveling Group', 
                                                        'Leveling Class', 
                                                        'Planning Plant', 
                                                        'Component Number', 
                                                        'Serial Number']]

            outputFile = '.\\Output\\Result\\'+ today +'_Other.xlsx'

            self.OtherReturnEnd.emit(True)
            self.thread().quit()
        except Exception as e:
            self.OtherReturnError.emit(e)
            self.thread().quit()
            return
        
class CustomFormatter(logging.Formatter):
    FORMATS = {
        logging.ERROR:   ('[%(asctime)s] %(levelname)s:%(message)s','yellow'),
        logging.DEBUG:   ('[%(asctime)s] %(levelname)s:%(message)s','white'),
        logging.INFO:    ('[%(asctime)s] %(levelname)s:%(message)s','white'),
        logging.WARNING: ('[%(asctime)s] %(levelname)s:%(message)s', 'yellow')
    }
    def format( self, record ):
        last_fmt = self._style._fmt
        opt = CustomFormatter.FORMATS.get(record.levelno)
        if opt:
            fmt, color = opt
            self._style._fmt = "<font color=\"{}\">{}</font>".format(QtGui.QColor(color).name(),fmt)
        res = logging.Formatter.format( self, record )
        self._style._fmt = last_fmt
        return res

class QTextEditLogger(logging.Handler):
    def __init__(self, parent=None):
        super().__init__()
        self.widget = QPlainTextEdit(parent)
        self.widget.setReadOnly(True)    
        self.widget.setGeometry(QRect(10, 260, 661, 161))
        self.widget.setStyleSheet('background-color: rgb(53, 53, 53);\ncolor: rgb(255, 255, 255);')
        self.widget.setObjectName('logBrowser')
        font = QFont()
        font.setFamily('Nanum Gothic')
        font.setBold(False)
        font.setPointSize(9)
        self.widget.setFont(font)
    def emit(self, record):
        msg = self.format(record)
        self.widget.appendHtml(msg) 
        # move scrollbar
        scrollbar = self.widget.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

class CalendarWindow(QWidget):
    submitClicked = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.initUI()
    def initUI(self):
        cal = QCalendarWidget(self)
        cal.setGridVisible(True)
        cal.clicked[QDate].connect(self.showDate)
        self.lb = QLabel(self)
        date = cal.selectedDate()
        self.lb.setText(date.toString("yyyy-MM-dd"))
        vbox = QVBoxLayout()
        vbox.addWidget(cal)
        vbox.addWidget(self.lb)
        self.submitBtn = QToolButton(self)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, 
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.submitBtn.setSizePolicy(sizePolicy)
        self.submitBtn.setMinimumSize(QSize(0, 35))
        self.submitBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.submitBtn.setObjectName('submitBtn')
        self.submitBtn.setText('착공지정일 결정')
        self.submitBtn.clicked.connect(self.confirm)
        vbox.addWidget(self.submitBtn)
        self.setLayout(vbox)
        self.setWindowTitle('캘린더')
        self.setGeometry(500,500,500,400)
        self.show()
    def showDate(self, date):
        self.lb.setText(date.toString("yyyy-MM-dd"))
    @pyqtSlot()
    def confirm(self):
        self.submitClicked.emit(self.lb.text())
        self.close()
class UISubWindow(QMainWindow):
    submitClicked = pyqtSignal(list)
    status = ''
    def __init__(self):
        super().__init__()
        self.setupUi()
    def setupUi(self):
        self.setObjectName('SubWindow')
        self.resize(600, 600)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.linkageInput = QLineEdit(self.groupBox)
        self.linkageInput.setMinimumSize(QSize(0, 25))
        self.linkageInput.setObjectName('linkageInput')
        self.linkageInput.setValidator(QDoubleValidator(self))
        self.gridLayout3.addWidget(self.linkageInput, 0, 1, 1, 3)
        self.linkageInputBtn = QPushButton(self.groupBox)
        self.linkageInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.linkageInputBtn, 0, 4, 1, 2)
        self.linkageAddExcelBtn = QPushButton(self.groupBox)
        self.linkageAddExcelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.linkageAddExcelBtn, 0, 6, 1, 2)
        self.mscodeInput = QLineEdit(self.groupBox)
        self.mscodeInput.setMinimumSize(QSize(0, 25))
        self.mscodeInput.setObjectName('mscodeInput')
        self.mscodeInputBtn = QPushButton(self.groupBox)
        self.mscodeInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.mscodeInput, 1, 1, 1, 3)
        self.gridLayout3.addWidget(self.mscodeInputBtn, 1, 4, 1, 2)
        self.mscodeAddExcelBtn = QPushButton(self.groupBox)
        self.mscodeAddExcelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.mscodeAddExcelBtn, 1, 6, 1, 2)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored,
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        self.submitBtn = QToolButton(self.groupBox)
        sizePolicy.setHeightForWidth(self.submitBtn.sizePolicy().hasHeightForWidth())
        self.submitBtn.setSizePolicy(sizePolicy)
        self.submitBtn.setMinimumSize(QSize(100, 35))
        self.submitBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.submitBtn.setObjectName('submitBtn')
        self.gridLayout3.addWidget(self.submitBtn, 3, 5, 1, 2)
        
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label2 = QLabel(self.groupBox)
        self.label2.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label2.setObjectName('label2')
        self.gridLayout3.addWidget(self.label2, 1, 0, 1, 1)
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.gridLayout3.addWidget(self.line, 2, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        listViewModelLinkage = QStandardItemModel()
        self.listViewLinkage = QListView(self.groupBox2)
        self.listViewLinkage.setModel(listViewModelLinkage)
        self.gridLayout5.addWidget(self.listViewLinkage, 1, 0, 1, 1)
        self.label3 = QLabel(self.groupBox2)
        self.label3.setAlignment(Qt.AlignLeft | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label3.setObjectName('label3')
        self.gridLayout5.addWidget(self.label3, 0, 0, 1, 1)
        self.vline = QFrame(self.groupBox2)
        self.vline.setFrameShape(QFrame.VLine)
        self.vline.setFrameShadow(QFrame.Sunken)
        self.vline.setObjectName('vline')
        self.gridLayout5.addWidget(self.vline, 1, 1, 1, 1)
        listViewModelmscode = QStandardItemModel()
        self.listViewmscode = QListView(self.groupBox2)
        self.listViewmscode.setModel(listViewModelmscode)
        self.gridLayout5.addWidget(self.listViewmscode, 1, 2, 1, 1)
        self.label4 = QLabel(self.groupBox2)
        self.label4.setAlignment(Qt.AlignLeft | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label4.setObjectName('label4')
        self.gridLayout5.addWidget(self.label4, 0, 2, 1, 1)
        self.label5 = QLabel(self.groupBox2)
        self.label5.setAlignment(Qt.AlignLeft | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label5.setObjectName('label5')       
        self.gridLayout5.addWidget(self.label5, 0, 3, 1, 1) 
        self.linkageDelBtn = QPushButton(self.groupBox2)
        self.linkageDelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout5.addWidget(self.linkageDelBtn, 2, 0, 1, 1)
        self.mscodeDelBtn = QPushButton(self.groupBox2)
        self.mscodeDelBtn.setMinimumSize(QSize(0,25))
        self.gridLayout5.addWidget(self.mscodeDelBtn, 2, 2, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)
        self.mscodeInput.returnPressed.connect(self.addmscode)
        self.linkageInput.returnPressed.connect(self.addLinkage)
        self.linkageInputBtn.clicked.connect(self.addLinkage)
        self.mscodeInputBtn.clicked.connect(self.addmscode)
        self.linkageDelBtn.clicked.connect(self.delLinkage)
        self.mscodeDelBtn.clicked.connect(self.delmscode)
        self.submitBtn.clicked.connect(self.confirm)
        self.linkageAddExcelBtn.clicked.connect(self.addLinkageExcel)
        self.mscodeAddExcelBtn.clicked.connect(self.addmscodeExcel)
        self.retranslateUi(self)
        self.show()
    
    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('SubWindow', '긴급/홀딩오더 입력'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('SubWindow', 'Linkage No 입력 :'))
        self.linkageInputBtn.setText(_translate('SubWindow', '추가'))
        self.label2.setText(_translate('SubWindow', 'MS-CODE 입력 :'))
        self.mscodeInputBtn.setText(_translate('SubWindow', '추가'))
        self.submitBtn.setText(_translate('SubWindow','추가 완료'))
        self.label3.setText(_translate('SubWindow', 'Linkage No List'))
        self.label4.setText(_translate('SubWindow', 'MS-Code List'))
        self.linkageDelBtn.setText(_translate('SubWindow', '삭제'))
        self.mscodeDelBtn.setText(_translate('SubWindow', '삭제'))
        self.linkageAddExcelBtn.setText(_translate('SubWindow', '엑셀 입력'))
        self.mscodeAddExcelBtn.setText(_translate('SubWindow', '엑셀 입력'))
    @pyqtSlot()
    def addLinkage(self):
        linkageNo = self.linkageInput.text()
        if len(linkageNo) == 16:
            if linkageNo.isdigit():
                model = self.listViewLinkage.model()
                linkageItem = QStandardItem()
                linkageItemModel = QStandardItemModel()
                dupFlag = False
                for i in range(model.rowCount()):
                    index = model.index(i,0)
                    item = model.data(index)
                    if item == linkageNo:
                        dupFlag = True
                    linkageItem = QStandardItem(item)
                    linkageItemModel.appendRow(linkageItem)
                if not dupFlag:
                    linkageItem = QStandardItem(linkageNo)
                    linkageItemModel.appendRow(linkageItem)
                    self.listViewLinkage.setModel(linkageItemModel)
                else:
                    QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
            else:
                QMessageBox.information(self, 'Error', '숫자만 입력해주세요.')
        elif len(linkageNo) == 0: 
            QMessageBox.information(self, 'Error', 'Linkage Number 데이터가 입력되지 않았습니다.')
        else:
            QMessageBox.information(self, 'Error', '16자리의 Linkage Number를 입력해주세요.')
    
    @pyqtSlot()
    def delLinkage(self):
        model = self.listViewLinkage.model()
        linkageItem = QStandardItem()
        linkageItemModel = QStandardItemModel()
        for index in self.listViewLinkage.selectedIndexes():
            selected_item = self.listViewLinkage.model().data(index)
            for i in range(model.rowCount()):
                index = model.index(i,0)
                item = model.data(index)
                linkageItem = QStandardItem(item)
                if selected_item != item:
                    linkageItemModel.appendRow(linkageItem)
            self.listViewLinkage.setModel(linkageItemModel)
    @pyqtSlot()
    def addmscode(self):
        mscode = self.mscodeInput.text()
        if len(mscode) > 0:
            model = self.listViewmscode.model()
            mscodeItem = QStandardItem()
            mscodeItemModel = QStandardItemModel()
            dupFlag = False
            for i in range(model.rowCount()):
                index = model.index(i,0)
                item = model.data(index)
                if item == mscode:
                    dupFlag = True
                mscodeItem = QStandardItem(item)
                mscodeItemModel.appendRow(mscodeItem)
            if not dupFlag:
                mscodeItem = QStandardItem(mscode)
                mscodeItemModel.appendRow(mscodeItem)
                self.listViewmscode.setModel(mscodeItemModel)
            else:
                QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
        else: 
            QMessageBox.information(self, 'Error', 'MS-CODE 데이터가 입력되지 않았습니다.')
    @pyqtSlot()
    def delmscode(self):
        model = self.listViewmscode.model()
        mscodeItem = QStandardItem()
        mscodeItemModel = QStandardItemModel()
        for index in self.listViewmscode.selectedIndexes():
            selected_item = self.listViewmscode.model().data(index)
            for i in range(model.rowCount()):
                index = model.index(i,0)
                item = model.data(index)
                mscodeItem = QStandardItem(item)
                if selected_item != item:
                    mscodeItemModel.appendRow(mscodeItem)
            self.listViewmscode.setModel(mscodeItemModel)
    @pyqtSlot()
    def addLinkageExcel(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                df = pd.read_excel(fileName)
                for i in df.index:
                    linkageNo = str(df[df.columns[0]][i])
                    if len(linkageNo) == 16:
                        if linkageNo.isdigit():
                            model = self.listViewLinkage.model()
                            linkageItem = QStandardItem()
                            linkageItemModel = QStandardItemModel()
                            dupFlag = False
                            for i in range(model.rowCount()):
                                index = model.index(i,0)
                                item = model.data(index)
                                if item == linkageNo:
                                    dupFlag = True
                                linkageItem = QStandardItem(item)
                                linkageItemModel.appendRow(linkageItem)
                            if not dupFlag:
                                linkageItem = QStandardItem(linkageNo)
                                linkageItemModel.appendRow(linkageItem)
                                self.listViewLinkage.setModel(linkageItemModel)
                            else:
                                QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
                        else:
                            QMessageBox.information(self, 'Error', '숫자만 입력해주세요.')
                    elif len(linkageNo) == 0: 
                        QMessageBox.information(self, 'Error', 'Linkage Number 데이터가 입력되지 않았습니다.')
                    else:
                        QMessageBox.information(self, 'Error', '16자리의 Linkage Number를 입력해주세요.')
        except Exception as e:
            QMessageBox.information(self, 'Error', '에러발생 : ' + e)
    @pyqtSlot()
    def addmscodeExcel(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                df = pd.read_excel(fileName)
                for i in df.index:
                    mscode = str(df[df.columns[0]][i])
                    if len(mscode) > 0:
                        model = self.listViewmscode.model()
                        mscodeItem = QStandardItem()
                        mscodeItemModel = QStandardItemModel()
                        dupFlag = False
                        for i in range(model.rowCount()):
                            index = model.index(i,0)
                            item = model.data(index)
                            if item == mscode:
                                dupFlag = True
                            mscodeItem = QStandardItem(item)
                            mscodeItemModel.appendRow(mscodeItem)
                        if not dupFlag:
                            mscodeItem = QStandardItem(mscode)
                            mscodeItemModel.appendRow(mscodeItem)
                            self.listViewmscode.setModel(mscodeItemModel)
                        else:
                            QMessageBox.information(self, 'Error', '중복된 데이터가 있습니다.')
                    else: 
                        QMessageBox.information(self, 'Error', 'MS-CODE 데이터가 입력되지 않았습니다.')
        except Exception as e:
            QMessageBox.information(self, 'Error', '에러발생 : ' + e)
    @pyqtSlot()
    def confirm(self):
        self.submitClicked.emit([self.listViewLinkage.model(), self.listViewmscode.model()])
        self.close()
class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()
        
    def setupUi(self):
        logger = logging.getLogger(__name__)
        rfh = RotatingFileHandler(filename='./Log.log', 
                                    mode='a',
                                    maxBytes=5*1024*1024,
                                    backupCount=2,
                                    encoding=None,
                                    delay=0
                                    )
        logging.basicConfig(level=logging.DEBUG, 
                            format = '%(asctime)s:%(levelname)s:%(message)s', 
                            datefmt = '%m/%d/%Y %H:%M:%S',
                            handlers=[rfh])
        self.setObjectName('MainWindow')
        self.resize(900, 1000)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.mainOrderinput = QLineEdit(self.groupBox)
        self.mainOrderinput.setMinimumSize(QSize(0, 25))
        self.mainOrderinput.setObjectName('mainOrderinput')
        self.mainOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.mainOrderinput, 0, 1, 1, 1)
        self.spOrderinput = QLineEdit(self.groupBox)
        self.spOrderinput.setMinimumSize(QSize(0, 25))
        self.spOrderinput.setObjectName('spOrderinput')
        self.spOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.spOrderinput, 1, 1, 1, 1)
        self.spOrderinput_1 = QLineEdit(self.groupBox)
        self.spOrderinput_1.setMinimumSize(QSize(0, 25))
        self.spOrderinput_1.setObjectName('spOrderinput_1')
        self.spOrderinput_1.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.spOrderinput_1, 2, 1, 1, 1)
        self.powerOrderinput = QLineEdit(self.groupBox)
        self.powerOrderinput.setMinimumSize(QSize(0, 25))
        self.powerOrderinput.setObjectName('powerOrderinput')
        self.powerOrderinput.setValidator(QIntValidator(self))
        self.gridLayout3.addWidget(self.powerOrderinput, 3, 1, 1, 1)
        self.dateBtn = QToolButton(self.groupBox)
        self.dateBtn.setMinimumSize(QSize(0,25))
        self.dateBtn.setObjectName('dateBtn')
        self.gridLayout3.addWidget(self.dateBtn, 4, 1, 1, 1)
        self.emgFileInputBtn = QPushButton(self.groupBox)
        self.emgFileInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.emgFileInputBtn, 5, 1, 1, 1)
        self.holdFileInputBtn = QPushButton(self.groupBox)
        self.holdFileInputBtn.setMinimumSize(QSize(0,25))
        self.gridLayout3.addWidget(self.holdFileInputBtn, 8, 1, 1, 1)
        self.label4 = QLabel(self.groupBox)
        self.label4.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label4.setObjectName('label4')
        self.gridLayout3.addWidget(self.label4, 6, 1, 1, 1)
        self.label5 = QLabel(self.groupBox)
        self.label5.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label5.setObjectName('label5')
        self.gridLayout3.addWidget(self.label5, 6, 2, 1, 1)
        self.label6 = QLabel(self.groupBox)
        self.label6.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label6.setObjectName('label6')
        self.gridLayout3.addWidget(self.label6, 9, 1, 1, 1)
        self.label7 = QLabel(self.groupBox)
        self.label7.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label7.setObjectName('label7')
        self.gridLayout3.addWidget(self.label7, 9, 2, 1, 1)
        listViewModelEmgLinkage = QStandardItemModel()
        self.listViewEmgLinkage = QListView(self.groupBox)
        self.listViewEmgLinkage.setModel(listViewModelEmgLinkage)
        self.gridLayout3.addWidget(self.listViewEmgLinkage, 7, 1, 1, 1)
        listViewModelEmgmscode = QStandardItemModel()
        self.listViewEmgmscode = QListView(self.groupBox)
        self.listViewEmgmscode.setModel(listViewModelEmgmscode)
        self.gridLayout3.addWidget(self.listViewEmgmscode, 7, 2, 1, 1)
        listViewModelHoldLinkage = QStandardItemModel()
        self.listViewHoldLinkage = QListView(self.groupBox)
        self.listViewHoldLinkage.setModel(listViewModelHoldLinkage)
        self.gridLayout3.addWidget(self.listViewHoldLinkage, 10, 1, 1, 1)
        listViewModelHoldmscode = QStandardItemModel()
        self.listViewHoldmscode = QListView(self.groupBox)
        self.listViewHoldmscode.setModel(listViewModelHoldmscode)
        self.gridLayout3.addWidget(self.listViewHoldmscode, 10, 2, 1, 1)
        self.labelBlank = QLabel(self.groupBox)
        self.labelBlank.setObjectName('labelBlank')
        self.gridLayout3.addWidget(self.labelBlank, 3, 4, 1, 1)
        self.progressbar = QProgressBar(self.groupBox)
        self.progressbar.setObjectName('progressbar')
        self.gridLayout3.addWidget(self.progressbar, 12, 1, 1, 2)
        self.runBtn = QToolButton(self.groupBox)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, 
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.runBtn.sizePolicy().hasHeightForWidth())
        self.runBtn.setSizePolicy(sizePolicy)
        self.runBtn.setMinimumSize(QSize(0, 35))
        self.runBtn.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.runBtn.setObjectName('runBtn')
        self.gridLayout3.addWidget(self.runBtn, 12, 3, 1, 2)
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label9 = QLabel(self.groupBox)
        self.label9.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label9.setObjectName('label9')
        self.gridLayout3.addWidget(self.label9, 1, 0, 1, 1)
        self.label10 = QLabel(self.groupBox)
        self.label10.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label10.setObjectName('label10')
        self.gridLayout3.addWidget(self.label10, 3, 0, 1, 1)
        self.label19 = QLabel(self.groupBox)
        self.label19.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label19.setObjectName('label19')
        self.gridLayout3.addWidget(self.label19, 2, 0, 1, 1)
        self.label11 = QLabel(self.groupBox)
        self.label11.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label11.setObjectName('label11')
        self.gridLayout3.addWidget(self.label11, 0, 2, 1, 1)
        # self.label12 = QLabel(self.groupBox)
        # self.label12.setAlignment(Qt.AlignRight | 
        #                         Qt.AlignTrailing | 
        #                         Qt.AlignVCenter)
        # self.label12.setObjectName('label12')
        # self.gridLayout3.addWidget(self.label12, 1, 2, 1, 1)
        # self.label13 = QLabel(self.groupBox)
        # self.label13.setAlignment(Qt.AlignRight | 
        #                         Qt.AlignTrailing | 
        #                         Qt.AlignVCenter)
        # self.label13.setObjectName('label13')
        # self.gridLayout3.addWidget(self.label13, 2, 2, 1, 1)
        self.label8 = QLabel(self.groupBox)
        self.label8.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label8.setObjectName('label8')
        self.gridLayout3.addWidget(self.label8, 4, 0, 1, 1) 
        self.labelDate = QLabel(self.groupBox)
        self.labelDate.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.labelDate.setObjectName('labelDate')
        self.gridLayout3.addWidget(self.labelDate, 3, 2, 3, 1) 
        self.label2 = QLabel(self.groupBox)
        self.label2.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label2.setObjectName('label2')
        self.gridLayout3.addWidget(self.label2, 5, 0, 1, 1)
        self.label3 = QLabel(self.groupBox)
        self.label3.setAlignment(Qt.AlignRight | 
                                    Qt.AlignTrailing | 
                                    Qt.AlignVCenter)
        self.label3.setObjectName('label3')
        self.gridLayout3.addWidget(self.label3, 8, 0, 1, 1)
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.cb_main = QComboBox(self.groupBox)
        self.gridLayout3.addWidget(self.cb_main, 0, 3, 1, 1)
        # self.cb_sp = QComboBox(self.groupBox)
        # self.gridLayout3.addWidget(self.cb_sp, 1, 3, 1, 1)
        # self.cb_power = QComboBox(self.groupBox)
        # self.gridLayout3.addWidget(self.cb_power, 2, 3, 1, 1)
        self.gridLayout3.addWidget(self.line, 11, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        self.logBrowser = QTextEditLogger(self.groupBox2)
        # self.logBrowser.setFormatter(
        #                             logging.Formatter('[%(asctime)s] %(levelname)s:%(message)s', 
        #                                                 datefmt='%Y-%m-%d %H:%M:%S')
        #                             )
        self.logBrowser.setFormatter(CustomFormatter())
        logging.getLogger().addHandler(self.logBrowser)
        logging.getLogger().setLevel(logging.INFO)
        self.gridLayout5.addWidget(self.logBrowser.widget, 0, 0, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)
        self.dateBtn.clicked.connect(self.selectStartDate)
        self.emgFileInputBtn.clicked.connect(self.emgWindow)
        self.holdFileInputBtn.clicked.connect(self.holdWindow)
        self.runBtn.clicked.connect(self.mainStartLeveling)
        #디버그용 플래그
        self.isDebug = True
        self.isFileReady = False

        if self.isDebug:
            self.debugDate = QLineEdit(self.groupBox)
            self.debugDate.setObjectName('debugDate')
            self.gridLayout3.addWidget(self.debugDate, 12, 0, 1, 1)
            self.debugDate.setPlaceholderText('디버그용 날짜입력')
        self.thread = QThread()
        self.thread.setTerminationEnabled(True)
        self.show()

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('MainWindow', 'FA-M3 착공 평준화 자동화 프로그램 Rev0.00'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('MainWindow', '메인 생산대수:'))
        self.label9.setText(_translate('MainWindow', '특수(모듈) 생산대수:'))
        self.label10.setText(_translate('MainWindow', '전원 생산대수:'))
        self.label11.setText(_translate('MainWindow', '메인 잔업시간:'))
        self.label19.setText(_translate('MainWindow', '특수(비모듈) 생산대수:'))
        # self.label12.setText(_translate('MainWindow', '특수 잔업시간:'))
        # self.label13.setText(_translate('MainWindow', '전원 잔업시간:'))       
        self.runBtn.setText(_translate('MainWindow', '실행'))
        self.label2.setText(_translate('MainWindow', '긴급오더 입력 :'))
        self.label3.setText(_translate('MainWindow', '홀딩오더 입력 :'))
        self.label4.setText(_translate('MainWindow', 'Linkage No List'))
        self.label5.setText(_translate('MainWindow', 'MSCode List'))
        self.label6.setText(_translate('MainWindow', 'Linkage No List'))
        self.label7.setText(_translate('MainWindow', 'MSCode List'))
        self.label8.setText(_translate('MainWndow', '착공지정일 입력 :'))
        self.labelDate.setText(_translate('MainWndow', '미선택'))
        self.dateBtn.setText(_translate('MainWindow', ' 착공지정일 선택 '))
        self.emgFileInputBtn.setText(_translate('MainWindow', '리스트 입력'))
        self.holdFileInputBtn.setText(_translate('MainWindow', '리스트 입력'))
        self.labelBlank.setText(_translate('MainWindow', '            '))
        self.cb_main.addItems(['잔업없음','1시간','2시간','3시간','4시간'])
        maxOrderInputFilePath = r'.\\착공량입력.xlsx'
        if not os.path.exists(maxOrderInputFilePath):
            logging.error('%s 파일이 없습니다. 착공량을 수동으로 입력해주세요.', maxOrderInputFilePath)
        else:
            df_orderInput = pd.read_excel(maxOrderInputFilePath)
            self.mainOrderinput.setText(str(df_orderInput['착공량'][0]))
            self.spOrderinput.setText(str(df_orderInput['착공량'][1]))
            self.powerOrderinput.setText(str(df_orderInput['착공량'][3]))

        logging.info('프로그램이 정상 기동했습니다')

    #착공지정일 캘린더 호출
    def selectStartDate(self):
        self.w = CalendarWindow()
        self.w.submitClicked.connect(self.getStartDate)
        self.w.show()
    
    #긴급오더 윈도우 호출
    @pyqtSlot()
    def emgWindow(self):
        self.w = UISubWindow()
        self.w.submitClicked.connect(self.getEmgListview)
        self.w.show()
    #홀딩오더 윈도우 호출
    @pyqtSlot()
    def holdWindow(self):
        self.w = UISubWindow()
        self.w.submitClicked.connect(self.getHoldListview)
        self.w.show()
    #긴급오더 리스트뷰 가져오기
    def getEmgListview(self, list):
        if len(list) > 0 :
            self.listViewEmgLinkage.setModel(list[0])
            self.listViewEmgmscode.setModel(list[1])
            logging.info('긴급오더 리스트를 정상적으로 불러왔습니다.')
        else:
            logging.error('긴급오더 리스트가 없습니다. 다시 한번 확인해주세요')
    
    #홀딩오더 리스트뷰 가져오기
    def getHoldListview(self, list):
        if len(list) > 0 :
            self.listViewHoldLinkage.setModel(list[0])
            self.listViewHoldmscode.setModel(list[1])
            logging.info('홀딩오더 리스트를 정상적으로 불러왔습니다.')
        else:
            logging.error('긴급오더 리스트가 없습니다. 다시 한번 확인해주세요')
    
    #프로그레스바 갱신
    def updateProgressbar(self, val):
        self.progressbar.setValue(val)
    #착공지정일 가져오기
    def getStartDate(self, date):
        if len(date) > 0 :
            self.labelDate.setText(date)
            logging.info('착공지정일이 %s 로 정상적으로 지정되었습니다.', date)
        else:
            logging.error('착공지정일이 선택되지 않았습니다.')

    def enableRunBtn(self):
        self.runBtn.setEnabled(True) 
        self.runBtn.setText('실행')
    def disableRunBtn(self):
        self.runBtn.setEnabled(False) 
        self.runBtn.setText('실행 중...')

    def mainShowError(self, str):
        logging.error(f'Main라인 에러 - {str}')
        self.enableRunBtn()

    def mainShowWarning(self, str):
        logging.warning(f'Main라인 경고 - {str}')

    def mainThreadEnd(self, isEnd):
        logging.info('착공이 완료되었습니다.')
        self.enableRunBtn()  

    def OtherShowError(self, str):
        logging.error(f'특수라인 에러 - {str}')
        self.enableRunBtn()

    def OtherShowWarning(self, str):
        logging.warning(f'특수라인 경고 - {str}')

    def OtherThreadEnd(self, isEnd):
        logging.info('착공이 완료되었습니다.')
        self.enableRunBtn()  

    @pyqtSlot()
    def mainStartLeveling(self):
        #마스터 데이터 불러오기 내부함수
        def loadMasterFile():
            self.isFileReady = True
            masterFileList = []
            date = datetime.datetime.today().strftime('%Y%m%d')
            if self.isDebug:
                date = self.debugDate.text()
            sosFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\SOS2.xlsx'    
            # progressFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\진척.xlsx'
            mainFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\MAIN.xlsx'
            spFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\OTHER.xlsx'
            powerFilePath = r'.\\input\\Master_File\\' + date +r'\\POWER.xlsx'
            calendarFilePath = r'd:\\FAM3_Leveling-1\\Input\\Calendar_File\\FY' + date[2:4] + '_Calendar.xlsx'
            smtAssyFilePath = r'd:\\FAM3_Leveling-1\\input\\DB\\MSCode_SMT_Assy.xlsx'
            # usedSmtAssyFilePath = r'.\\input\\DB\\MSCode_SMT_Assy.xlsx'
            # secMainListFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\100L1311('+date[4:8]+')MAIN_2차.xlsx'
            inspectFacFilePath = r'd:\\FAM3_Leveling-1\\input\\DB\\Inspect_Fac.xlsx'
            modelSpecFilePath = r'd:\\FAM3_Leveling-1\\input\\MSCODE_Table\\FAM3기종분류표.xlsx'
            smtAssyUnCheckFilePath = r'd:\\FAM3_Leveling-1\\input\\MSCODE_Table\\SMT수량_비관리대상.xlsx'
            NonSp_BLFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\BL.xlsx' #비모듈 - BL 레벨링 리스트
            if os.path.exists(NonSp_BLFilePath) == False:
                NonSp_BLFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File'
            NonSp_TerminalFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File\\' + date +r'\\TERMINAL.xlsx' #비모듈 - TERMINAL 레벨링 리스트
            if os.path.exists(NonSp_TerminalFilePath) == False:
                NonSp_TerminalFilePath = r'd:\\FAM3_Leveling-1\\input\\Master_File'

            pathList = [sosFilePath,  
                        mainFilePath, 
                        spFilePath, 
                        powerFilePath, 
                        calendarFilePath, 
                        smtAssyFilePath,  
                        inspectFacFilePath,
                        modelSpecFilePath,
                        smtAssyUnCheckFilePath,
                        NonSp_BLFilePath,
                        NonSp_TerminalFilePath]
            for path in pathList:
                if os.path.exists(path):
                    file = glob.glob(path)[0]
                    masterFileList.append(file)
                else:
                    logging.error('%s 파일이 없습니다. 확인해주세요.', path)
                    self.runBtn.setEnabled(True)
                    # if path != NonSpFilePath:
                    #     self.isFileReady = False
            if self.isFileReady :
                logging.info('마스터 파일 및 캘린더 파일을 정상적으로 불러왔습니다.')
            return masterFileList
        
        self.disableRunBtn()
        # try:
        list_masterFile = loadMasterFile()
        list_emgHold = []
        list_emgHold.append([str(self.listViewEmgLinkage.model().data(self.listViewEmgLinkage.model().index(x,0))) for x in range(self.listViewEmgLinkage.model().rowCount())])
        list_emgHold.append([self.listViewEmgmscode.model().data(self.listViewEmgmscode.model().index(x,0)) for x in range(self.listViewEmgmscode.model().rowCount())])
        list_emgHold.append([str(self.listViewHoldLinkage.model().data(self.listViewHoldLinkage.model().index(x,0))) for x in range(self.listViewHoldLinkage.model().rowCount())])
        list_emgHold.append([self.listViewHoldmscode.model().data(self.listViewHoldmscode.model().index(x,0)) for x in range(self.listViewHoldmscode.model().rowCount())])

        if self.isFileReady :
            if len(self.spOrderinput.text()) > 0:
                self.thread_Other = ThreadClass(self.isDebug,
                                                self.debugDate.text(), 
                                                self.cb_main.currentText(),
                                                list_masterFile,
                                                float(self.spOrderinput.text()),
                                                float(self.spOrderinput_1.text()), 
                                                list_emgHold)
                self.thread_Other.moveToThread(self.thread)
                self.thread.started.connect(self.thread_Other.run)
                self.thread_Other.OtherReturnError.connect(self.OtherShowError)
                self.thread_Other.OtherReturnEnd.connect(self.OtherThreadEnd)
                self.thread_Other.OtherReturnWarning.connect(self.OtherShowWarning)
                self.thread.start()
            else:
                logging.info('메인기종 착공량이 입력되지 않아 메인기종 착공은 미실시 됩니다.')
        else:
            logging.warning('필수 파일이 없어 더이상 진행할 수 없습니다.')

if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
    sys.exit(app.exec_())