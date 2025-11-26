import pandas as pd
import xlsxwriter
import os
import sys

# 1. 파일 경로 설정 (오직 input.xlsx만 찾습니다)
current_folder = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(current_folder, "input.xlsx")

print(f"▶ 작업 폴더: {current_folder}")
print(f"▶ 찾고 있는 파일: {target_file}")

# 2. 파일 존재 여부 확인
if not os.path.exists(target_file):
    print("\n[!!! 비상 !!!] 'input.xlsx' 파일이 없습니다!")
    print("1. 엑셀 파일을 'input.xlsx' 이름으로 저장하셨나요?")
    print("2. 이 파이썬 파일과 같은 폴더에 넣으셨나요?")
    print("파일 목록을 확인해보세요:")
    print(os.listdir(current_folder))
    sys.exit()

# 3. 데이터 불러오기 (엑셀 전용)
try:
    # engine='openpyxl'을 명시해서 확실하게 엽니다
    df = pd.read_excel(target_file, engine='openpyxl')
    print("▶ [성공] input.xlsx 파일을 정상적으로 읽어왔습니다!")
except Exception as e:
    print(f"\n[치명적 오류] 엑셀 파일을 읽는 도중 에러 발생: {e}")
    print("혹시 파일이 열려 있나요? 엑셀을 끄고 다시 실행해주세요.")
    sys.exit()

# 4. 데이터 클리닝 및 변환
try:
    # 컬럼명 변경 (수식 오류 방지)
    rename_map = {
        df.columns[1]: 'YouTube_Freq', df.columns[2]: 'YouTube_Trust',
        df.columns[3]: 'Comm_Freq', df.columns[4]: 'Comm_Trust',
        df.columns[5]: 'News_Freq', df.columns[6]: 'News_Trust',
        df.columns[7]: 'App_Freq', df.columns[8]: 'App_Trust',
        df.columns[9]: 'Report_Freq', df.columns[10]: 'Report_Trust',
        df.columns[11]: 'SNS_Freq', df.columns[12]: 'SNS_Trust',
        df.columns[13]: 'Trading_Freq',
        df.columns[14]: 'Risk_1', df.columns[15]: 'Risk_2', df.columns[16]: 'Risk_3',
        df.columns[18]: 'Quiz_1', df.columns[19]: 'Quiz_2', df.columns[20]: 'Quiz_3',
        df.columns[21]: 'Quiz_4', df.columns[22]: 'Quiz_5'
    }
    df = df.rename(columns=rename_map)

    # 매핑 규칙
    freq_map = {'이용안함': 1, '월 1회 미만': 2, '월 1~2회': 2, '주 1~2회': 3, '주 1~3회': 3, '주 3~5회': 4, '거의 매일 (주 5회 이상)': 5}
    trust_map = {'전혀 신뢰 안 함': 1, '신뢰하지 않는 편': 2, '보통': 3, '신뢰하는 편': 4, '매우 신뢰함': 5}
    agree_map = {'전혀 동의 안 함': 1, '동의하지 않는 편': 2, '보통': 3, '동의하는 편': 4, '매우 동의함': 5}
    trading_map = {'월 1회 미만': 1, '월 1회 ~ 3회': 2, '주 1회 ~ 2회': 3, '주 3회 ~ 5회': 4, '거의 매일 (주 5회 이상)': 5}

    # 매핑 적용
    cols_maps = [
        (['YouTube_Freq', 'Comm_Freq', 'News_Freq', 'App_Freq', 'Report_Freq', 'SNS_Freq'], freq_map),
        (['YouTube_Trust', 'Comm_Trust', 'News_Trust', 'App_Trust', 'Report_Trust', 'SNS_Trust'], trust_map),
        (['Risk_1', 'Risk_2', 'Risk_3'], agree_map)
    ]

    for cols, mapper in cols_maps:
        for col in cols:
            if col in df.columns:
                df[f'{col}_Num'] = df[col].astype(str).str.strip().map(mapper).fillna(0)

    if 'Trading_Freq' in df.columns:
        df['Trading_Freq_Num'] = df['Trading_Freq'].astype(str).str.strip().map(trading_map).fillna(0)

    # 금융 지식 점수
    if 'Quiz_1' in df.columns:
        df['Score_1'] = (df['Quiz_1'].str.contains('하락', na=False)).astype(int)
        df['Score_2'] = (df['Quiz_2'].str.contains('저평가', na=False)).astype(int)
        df['Score_3'] = (df['Quiz_3'].str.contains('감소', na=False)).astype(int)
        df['Score_4'] = (df['Quiz_4'].str.contains('ETF', na=False)).astype(int)
        df['Score_5'] = (df['Quiz_5'].str.contains('주주', na=False)).astype(int)
        df['Fin_Lit_Total'] = df[['Score_1', 'Score_2', 'Score_3', 'Score_4', 'Score_5']].sum(axis=1)

    # 숫자형 컬럼 리스트 확보
    num_cols = [c for c in df.columns if c.endswith('_Num')] + ['Fin_Lit_Total']

except Exception as e:
    print(f"\n[오류] 데이터 처리 중 문제가 발생했습니다: {e}")
    sys.exit()

# 5. 결과 파일 생성
output_file = 'Final_Submission.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

# [시트 1] Data
df.to_excel(writer, sheet_name='Data', index=False)
worksheet1 = writer.sheets['Data']
worksheet1.set_column('A:Z', 15)

# [시트 2] Stats (값 직접 계산)
stats_df = df[num_cols].describe().T[['mean', '50%', 'std', 'min', 'max']]
stats_df.columns = ['평균', '중위값', '표준편차', '최소값', '최대값']
stats_df.to_excel(writer, sheet_name='Stats')
worksheet2 = writer.sheets['Stats']
worksheet2.set_column('A:A', 25)
worksheet2.set_column('B:F', 12)

# [시트 3] Reliability (값 복사 + 수식)
worksheet3 = workbook.add_worksheet('Reliability')
trust_cols = ['YouTube_Trust_Num', 'Comm_Trust_Num', 'News_Trust_Num', 'App_Trust_Num', 'Report_Trust_Num', 'SNS_Trust_Num']
headers = ['YouTube', 'Community', 'News', 'App', 'Report', 'SNS', 'Total_Score']

for i, h in enumerate(headers):
    worksheet3.write(0, i, h)

data_rows = len(df)
for row_idx, row in df.iterrows():
    for col_idx, col_name in enumerate(trust_cols):
        val = row[col_name] if col_name in df.columns else 0
        worksheet3.write(row_idx + 1, col_idx, val)
    # 총점 수식
    worksheet3.write_formula(row_idx + 1, 6, f"=SUM(A{row_idx+2}:F{row_idx+2})")

# 하단 수식
start_footer = data_rows + 2
worksheet3.write(start_footer, 0, "항목 분산(Variance)")
for i in range(7):
    col_char = chr(65 + i)
    worksheet3.write_formula(start_footer, i, f"=VAR.S({col_char}2:{col_char}{data_rows+1})")

calc_row = start_footer + 3
worksheet3.write(calc_row, 0, "▶ 크론바흐 알파 계산 과정")
worksheet3.write(calc_row+1, 0, "1. 문항 수 (K)")
worksheet3.write(calc_row+1, 1, 6)
worksheet3.write(calc_row+2, 0, "2. 항목 분산의 합 (Sum Var)")
var_row_idx = start_footer + 1
worksheet3.write_formula(calc_row+2, 1, f"=SUM(A{var_row_idx}:F{var_row_idx})")
worksheet3.write(calc_row+3, 0, "3. 총점의 분산 (Total Var)")
worksheet3.write_formula(calc_row+3, 1, f"=G{var_row_idx}")
worksheet3.write(calc_row+4, 0, "4. 최종 결과 (Alpha)")
k_cell = f"B{calc_row+2}"
sum_var_cell = f"B{calc_row+3}"
tot_var_cell = f"B{calc_row+4}"
worksheet3.write_formula(calc_row+4, 1, f"=({k_cell}/({k_cell}-1)) * (1 - ({sum_var_cell}/{tot_var_cell}))")

writer.close()
print(f"\n▶ [완료] 형님, 고생하셨습니다. '{output_file}' 파일이 생성되었습니다.")