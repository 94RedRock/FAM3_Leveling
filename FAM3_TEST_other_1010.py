import pandas as pd

#df_addSmtAssy
data = pd.read_excel('flow9.xlsx')


# ate별 최대 생산 대수
ate_no_max = {'F': 50, 'FV': 100, 'V': 50, 'S': 30}

data['확정수량'] = 0

# 특수 모듈일 경우에
data_other = data[data.PRODUCT_TYPE == 'OTHER']

for x in data_other.index:
    # 설비별 최대 생산대수 값(MAX)이 0 이상일 때 MAX - 미착공수주잔을 계산하여 remained에 저장
    if ate_no_max[data.loc[x, 'ATE_NO']] > 0:
        remained = ate_no_max[data.loc[x, 'ATE_NO']] - data.loc[x, '미착공수주잔']
        #미착공 수량 계산
        if remained >= 0:
            data.loc[x, '확정수량'] = data.loc[x, '미착공수주잔']
            ate_no_max[data.loc[x, 'ATE_NO']] = remained
        # remained 0 미만이면 확정수량은 MAX가 되고, MAX는 0이 된다.
        else:
            data.loc[x, '확정수량'] = ate_no_max[data.loc[x, 'ATE_NO']]
            ate_no_max[data.loc[x, 'ATE_NO']] = 0


#data.to_excel('result_test.xlsx', index=False)

