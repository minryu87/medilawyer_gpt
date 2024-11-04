#%%
import openpyxl
import pandas as pd
import requests
from datetime import datetime
from tqdm import tqdm
import math
import pickle
import os

# API 설정
GPT4V_KEY = "0df571534be34d3fa5ac6dc8d9c2aa9e"
GPT4V_ENDPOINT = "https://ohkimsai.openai.azure.com/openai/deployments/gpt-4o-ryumin/chat/completions?api-version=2024-02-15-preview"
headers = {
    "Content-Type": "application/json",
    "api-key": GPT4V_KEY,
}

# 경로 설정
folder_path = "C:\\medilawyer_gpt"
prompt_file = f"{folder_path}\\medilawyer_prompt.xlsx"

# I. 프롬프트 불러오기
prompt_df = pd.read_excel(prompt_file, sheet_name='prompt')
prompt_TB = prompt_df.copy()

# 사용 가능한 버전 목록 제공
print(prompt_TB[['Version', 'Created_Date']])
now_version = input("어떤 버전의 프롬프트를 실행하시겠습니까? ")

# Version 컬럼을 문자열로 변환
prompt_TB['Version'] = prompt_TB['Version'].astype(str)

# 입력한 버전의 존재 여부 확인
if now_version not in prompt_TB['Version'].values:
    print(f"입력하신 버전 '{now_version}'이 존재하지 않습니다. 유효한 버전을 입력해주세요.")
    exit()

# 선택된 버전의 row 가져오기 및 유효성 검사
now_prompt_row = prompt_TB[prompt_TB['Version'] == now_version].iloc[0]
missing_columns = [col for col in ['Prompt', 'Temperature', 'Top_p'] if pd.isnull(now_prompt_row[col])]
if missing_columns:
    missing_values = ', '.join(missing_columns)
    print(f"{missing_values} 값이 비어있습니다. medilawyer_prompt.xlsx 엑셀 파일을 확인해주세요.")
    exit()

# 모든 값이 존재하는 경우 변수에 값 저장
now_prompt = now_prompt_row['Prompt']
now_temperature = now_prompt_row['Temperature']
now_topp = now_prompt_row['Top_p']
print(f"프롬프트 버전 {now_version}을 성공적으로 로드하였습니다")

# II. 데이터 불러오기
data_file = f"{folder_path}\\medilawyer_data.xlsx"
data_df = pd.read_excel(data_file, sheet_name='data')
medilawyer_data_TB = data_df.copy()
now_data_count = len(medilawyer_data_TB) - 1

# 예상 시간 계산
min_est_time = round(now_data_count * 5 / 60)
max_est_time = round(now_data_count * 8 / 60)
print(f"리뷰 데이터 {now_data_count}개를 성공적으로 로드하였습니다. 처리에 소요되는 예상 시간은 {min_est_time}~{max_est_time}분입니다.")

# 진행 상태 불러오기
pickle_file = f"{folder_path}\\progress_{now_version}.pkl"
if os.path.exists(pickle_file):
    with open(pickle_file, 'rb') as f:
        saved_results, start_index = pickle.load(f)
    print(f"이전에 저장된 진행 상태를 불러왔습니다. {start_index}번째 항목부터 다시 시작합니다.")
else:
    saved_results = []
    start_index = 0

# GPT 함수 정의
def gpt_generate_response(now_contents):
    payload = {
        "messages": [
            {"role": "system", "content": f"너는 주어진 리뷰의 명예훼손죄, 모욕죄 해당 여부를 판단하는 봇이다. 현재 판단 대상의 문장은 {now_contents}이다."},
            {"role": "user", "content": now_prompt}
        ],
        "temperature": now_temperature,
        "top_p": now_topp,
        "max_tokens": 1000
    }
    
    try:
        response = requests.post(GPT4V_ENDPOINT, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        response_data = response.json()
        result = response_data['choices'][0]['message']['content'].strip().split('|')
        return result if len(result) == 11 else ["error"] * 11
    except Exception as e:
        print(f"Error during GPT request: {e}")
        return ["error"] * 11

# III. GPT 실행 및 결과 수집
for i, row in tqdm(medilawyer_data_TB.iloc[start_index:].iterrows(), total=now_data_count - start_index, initial=start_index):
    contents = row['Contents']
    for _ in range(3):  # 최대 3회 시도
        response = gpt_generate_response(contents)
        if response[0] != "error":
            break
    if response[0] == "error":
        saved_results.append([row['Idx'], contents, row['Um_Def'], 'error', 'False', row['Um_Ins'], 'error', 'False', 
            'error', '', 'error', '', 'error', '', 'error', '', 'error', '', 'error', '', 'error', '', 
            'error', '', 'error', '', 'error', ''])
    else:
        def_tf = 'True' if str(row['Um_Def']).strip().lower() == str(response[0]).strip().lower() else 'False'
        ins_tf = 'True' if str(row['Um_Ins']).strip().lower() == str(response[6]).strip().lower() else 'False'
        saved_results.append([
            row['Idx'], contents, row['Um_Def'], response[0], def_tf, row['Um_Ins'], response[6], ins_tf,
            response[0], '', response[1], '', response[2], '', response[3], '', response[4], '', response[5], '', 
            response[6], '', response[7], '', response[8], '', response[9], ''
        ])
    
    # 중간 진행 상태 저장
    with open(pickle_file, 'wb') as f:
        pickle.dump((saved_results, i + 1), f)

# IV. 결과 저장 및 최종 상태 삭제
columns = [
    'Idx', 'Contents', 'Um_Def', 'GPT_Def', 'Def_TF', 'Um_Ins', 'GPT_Ins', 'Ins_TF', 
    '명예훼손죄 여부', 'Remarks1', '명예훼손죄 판단 요약', 'Remarks2', '공연성 판단 근거', 'Remarks3', 
    '사실의 적시 판단 근거', 'Remarks4', '비방의 목적 판단 근거', 'Remarks5', '피해자의 특정 판단 근거', 
    'Remarks6', '모욕죄 여부', 'Remarks7', '모욕죄 판단 요약', 'Remarks8', '공연성 판단 근거', 
    'Remarks9', '사람에 대한 모욕 판단 근거', 'Remarks10'
]
medilawyer_result_TB = pd.DataFrame(saved_results, columns=columns)
result_file = f"{folder_path}\\medilawyer_result.xlsx"
sheet_name = f"result_{now_version}_{datetime.now().strftime('%m%d_%H%M')}"

with pd.ExcelWriter(result_file, mode='a', if_sheet_exists='new') as writer:
    medilawyer_result_TB.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"결과가 {result_file} 파일의 {sheet_name} 시트에 저장되었습니다.")

if os.path.exists(pickle_file):
    os.remove(pickle_file)

# 성능 계산하기
def calculate_metrics(df, column1, column2):
    df_filtered = df[
        (df[column1].astype(str).str.strip().str.lower().isin(['true', 'false'])) &
        (df[column2].astype(str).str.strip().str.lower().isin(['true', 'false']))
    ]
    tp = len(df_filtered[(df_filtered[column1].astype(str).str.strip().str.lower() == 'true') & 
                         (df_filtered[column2].astype(str).str.strip().str.lower() == 'true')])
    tn = len(df_filtered[(df_filtered[column1].astype(str).str.strip().str.lower() == 'false') & 
                         (df_filtered[column2].astype(str).str.strip().str.lower() == 'false')])
    fp = len(df_filtered[(df_filtered[column1].astype(str).str.strip().str.lower() == 'false') & 
                         (df_filtered[column2].astype(str).str.strip().str.lower() == 'true')])
    fn = len(df_filtered[(df_filtered[column1].astype(str).str.strip().str.lower() == 'true') & 
                         (df_filtered[column2].astype(str).str.strip().str.lower() == 'false')])
    f1_score = 2 * tp / (2 * tp + fp + fn) if (2 * tp + fp + fn) > 0 else None
    return tp, tn, fp, fn, f1_score

Def_TP, Def_TN, Def_FP, Def_FN, Def_F1 = calculate_metrics(medilawyer_result_TB, 'Um_Def', 'GPT_Def')
Ins_TP, Ins_TN, Ins_FP, Ins_FN, Ins_F1 = calculate_metrics(medilawyer_result_TB, 'Um_Ins', 'GPT_Ins')

prompt_TB.loc[prompt_TB['Version'] == now_version, 
    ['Def_F1', 'Def_TP', 'Def_TN', 'Def_FP', 'Def_FN', 'Ins_F1', 'Ins_TP', 'Ins_TN', 'Ins_FP', 'Ins_FN']] = [
    Def_F1, Def_TP, Def_TN, Def_FP, Def_FN, Ins_F1, Ins_TP, Ins_TN, Ins_FP, Ins_FN
]
with pd.ExcelWriter(prompt_file, mode='a', if_sheet_exists='overlay') as writer:
    prompt_TB.to_excel(writer, sheet_name='prompt', index=False)

print(f"프롬프트 {now_version} 버전의 성능 결과를 medilawyer_prompt.xlsx 파일에 저장하였습니다. 모든 작업이 완료되었습니다.")
# %%
