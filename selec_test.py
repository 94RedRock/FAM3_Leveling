import pandas as pd

#리스트 컴프리헨션

#원본 합치기
#df_leveling_list = pd.read_excel(list_masterFile[2])

df_levelingSp = pd.read_excel(r'd:\\python_test\\MAIN.xlsx')

df_levelingSpDropSEQ = df_levelingSp[df_levelingSp['Sequence No'].isnull()]
df_levelingSpUndepSeq = df_levelingSp[df_levelingSp['Sequence No']=='Undep']
df_levelingSpUncorSeq = df_levelingSp[df_levelingSp['Sequence No']=='Uncor']
df_levelingSp = pd.concat([df_levelingSpDropSEQ, df_levelingSpUndepSeq, df_levelingSpUncorSeq])
df_levelingSp = df_levelingSp.reset_index(drop=True)

#특수 라인 빼기
df_sosAddMainModel = pd.read_excel(r'd:\\python_test\\flow9.xlsx')
df_OtherDropSEQ = df_sosAddMainModel[df_sosAddMainModel['PRODUCT_TYPE']=='OTHER']
df_Other = df_OtherDropSEQ.reset_index(drop=True)

f_max = int(50)
v_max = int(50)
fv_max = int(100)
s_max = int(30)

f_max_cnt
#커밋 test
for i in df_Other.index:
    if df_Other['ATE_NO'][i] == 'F':
        leftover_cnt = int(df_Other['미착공수주잔'][i])
        if f_max - leftover_cnt < 0:
            df_Other['확정수량'][i] == leftover_cnt
            f_max =
            #max 값을 설정하려면 변수를 조건절 밖으로 꺼내야 한다.
        else:
            df_Other['확정수량'][i] == f_max
            break

        

    
if 'product_type' == other:

