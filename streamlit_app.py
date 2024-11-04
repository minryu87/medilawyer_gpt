import os
import streamlit as st
import pandas as pd
from datetime import datetime
import requests
import pickle
from pathlib import Path

# 기본 경로 설정
folder_path = Path("C:/medilawyer_gpt")
prompt_file = folder_path / "medilawyer_prompt.xlsx"
data_file = folder_path / "medilawyer_data.xlsx"
result_file_template = folder_path / "medilawyer_result.xlsx"

# API 설정
GPT4V_KEY = "0df571534be34d3fa5ac6dc8d9c2aa9e"
GPT4V_ENDPOINT = "https://ohkimsai.openai.azure.com/openai/deployments/gpt-4o-ryumin/chat/completions?api-version=2024-02-15-preview"
headers = {
    "Content-Type": "application/json",
    "api-key": GPT4V_KEY,
}

# Streamlit App
st.title("Medilawyer GPT Processor")
st.write("This application processes data using GPT and stores results in `C:/medilawyer_gpt/`")

# 1. 프롬프트 파일 로드 및 버전 선택
if not prompt_file.exists():
    st.error("프롬프트 파일이 존재하지 않습니다. `C:/medilawyer_gpt/` 폴더에 프롬프트 파일을 저장해주세요.")
else:
    # 프롬프트 불러오기
    prompt_df = pd.read_excel(prompt_file, sheet_name="prompt")
    prompt_TB = prompt_df.copy()
    
    # 선택할 수 있는 Version 목록 표시
    version_list = prompt_TB['Version'].astype(str).tolist()
    now_version = st.selectbox("어떤 버전의 프롬프트를 실행하시겠습니까?", version_list)
    
    if st.button("Run Processing"):
        # 선택된 Version의 프롬프트 로드
        now_prompt_row = prompt_TB[prompt_TB['Version'] == now_version].iloc[0]
        now_prompt = now_prompt_row['Prompt']
        now_temperature = now_prompt_row['Temperature']
        now_topp = now_prompt_row['Top_p']
        st.write(f"프롬프트 버전 {now_version}을 성공적으로 로드하였습니다.")
        
        # 2. 데이터 파일 로드
        if not data_file.exists():
            st.error("데이터 파일이 존재하지 않습니다. `C:/medilawyer_gpt/` 폴더에 데이터 파일을 저장해주세요.")
        else:
            data_df = pd.read_excel(data_file, sheet_name="data")
            medilawyer_data_TB = data_df.copy()
            now_data_count = len(medilawyer_data_TB) - 1
            st.write(f"리뷰 데이터 {now_data_count}개를 성공적으로 로드하였습니다.")
            
            # GPT 응답 함수
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
                    st.write(f"Error during GPT request: {e}")
                    return ["error"] * 11

            # GPT 실행 및 결과 수집
            results = []
            for _, row in medilawyer_data_TB.iterrows():
                contents = row['Contents']
                for _ in range(3):
                    response = gpt_generate_response(contents)
                    if response[0] != "error":
                        break
                if response[0] == "error":
                    results.append([row['Idx'], contents, row['Um_Def'], 'error', 'False', row['Um_Ins'], 'error', 'False'] + ["error"] * 11)
                else:
                    def_tf = 'True' if str(row['Um_Def']).strip().lower() == str(response[0]).strip().lower() else 'False'
                    ins_tf = 'True' if str(row['Um_Ins']).strip().lower() == str(response[6]).strip().lower() else 'False'
                    results.append([row['Idx'], contents, row['Um_Def'], response[0], def_tf, row['Um_Ins'], response[6], ins_tf] + response[:11])

            # 결과를 DataFrame으로 생성
            columns = [
                'Idx', 'Contents', 'Um_Def', 'GPT_Def', 'Def_TF', 'Um_Ins', 'GPT_Ins', 'Ins_TF',
                '명예훼손죄 여부', 'Remarks1', '명예훼손죄 판단 요약', 'Remarks2', '공연성 판단 근거', 'Remarks3',
                '사실의 적시 판단 근거', 'Remarks4', '비방의 목적 판단 근거', 'Remarks5', '피해자의 특정 판단 근거',
                'Remarks6', '모욕죄 여부', 'Remarks7', '모욕죄 판단 요약', 'Remarks8', '공연성 판단 근거',
                'Remarks9', '사람에 대한 모욕 판단 근거', 'Remarks10'
            ]
            medilawyer_result_TB = pd.DataFrame(results, columns=columns)

            # 결과 저장
            timestamp = datetime.now().strftime("%m%d_%H%M")
            sheet_name = f"result_{now_version}_{timestamp}"
            result_file = result_file_template
            with pd.ExcelWriter(result_file, mode='a', if_sheet_exists='new') as writer:
                medilawyer_result_TB.to_excel(writer, sheet_name=sheet_name, index=False)

            st.write(f"결과가 {result_file} 파일의 {sheet_name} 시트에 저장되었습니다.")
