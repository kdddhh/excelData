import pandas as pd
from openai import OpenAI

import os
from dotenv import load_dotenv

load_dotenv()

# 2. 불러온 환경 변수 중에서 'OPENAI_API_KEY'라는 이름의 값을 꺼내옵니다.
my_key = os.getenv("OPENAI_API_KEY")

# 1. 엑셀 데이터 로드 
df1 = pd.read_excel(r'C:\Users\kados\OneDrive\바탕 화면\클레임 엑셀 파일\정기 시험 보고서\p_60a0_20251203p_60a0_rpa 라인 원본 데이터.xls', sheet_name='라인 출고 데이터 원본')
df2 = pd.read_excel(r'C:\Users\kados\OneDrive\바탕 화면\클레임 엑셀 파일\정기 시험 보고서\p_60a0_20251203p_60a0_rpa 라인 원본 데이터.xls', sheet_name='P_60A0_20251203P_60A0_1 (3)')
df3 = pd.read_excel(r'C:\Users\kados\OneDrive\바탕 화면\클레임 엑셀 파일\정기 시험 보고서\p_60a0_20251203p_60a0_rpa 라인 원본 데이터.xls', sheet_name='P_60A0_20251203P_60A0_1 (4)')

# 2. OpenAI 클라이언트 설정
client = OpenAI(api_key=my_key)

# 3. AI에게 알려줄 데이터 뼈대 정보
df1_info = f"라인 출고 원본 columns: {list(df1.columns)}"
df2_info = f"P06A0 측정 추가 columns: {list(df2.columns)}"
df3_info = f"구분 추가 columns: {list(df3.columns)}"

# 4. 프롬프트 
prompt = f"""
당신은 파이썬 데이터 분석 전문가입니다.
현재 메모리에는 3개의 데이터프레임(df1, df2, df3)이 로드되어 있습니다.

- df1 컬럼 정보: {df1_info}
- df2 컬럼 정보: {df2_info}
- df3 컬럼 정보: {df3_info}

[최종 목표]
df3의 '구분' 컬럼에 정상적인 값이 있는 바코드 데이터만 추출하여 df1, df2와 병합한 뒤, 아래 규칙을 적용해 최종 `result_df`를 생성하세요. 헤더는 파이썬 코드를 통해 동적으로 MultiIndex(대분류-소분류)로 묶어내야 합니다.

[데이터 가공 및 연산 규칙 (절대 준수)]
1. 필수 데이터 필터링: 먼저 df3의 '구분' 컬럼값이 존재하는(결측치 NaN이나 빈 문자열이 아닌) 행만 남기도록 필터링하세요.
2. 데이터 병합: 필터링된 df3를 마스터로 삼아 연속으로 Left Join 하세요. 
   - 1차 병합: df3와 df1 병합 (on='바코드리드 데이터', suffixes=('_df3', '_df1'))
   - 2차 병합: 1차 결과와 df2 병합 (on='바코드리드 데이터', suffixes=('', '_df2'))
   반드시 각 병합 단계마다 suffixes를 위와 같이 다르게 지정하여 Duplicate columns 에러를 방지하세요.
3. 체결력 0값 처리: 컬럼명에 '체결력'이 포함된 모든 열을 찾아, 값이 0인 데이터만 `NaN`으로 변환하세요.
4. 온도센서 동적 계산: 병합된 데이터 중 컬럼명에 '기준 온도(24°C)'가 포함된 컬럼을 변수 `a`로 두고, `a / (a + 2) * 5` 공식을 계산하여 '온도센서'라는 신규 컬럼을 만드세요.

[MultiIndex 동적 생성 알고리즘 규칙]
병합과 연산이 끝난 후, 단일 컬럼명들을 순회하며 아래 키워드 패턴에 따라 (대분류, 소분류) 튜플을 생성해 MultiIndex 헤더로 덮어씌우세요.

1. 공정 데이터: 컬럼명에 'A0작업 결과'가 포함되거나, 해당 공정에 속하는 측정값(Judge, 체결력, Angle, Torque 등)인 경우, 가장 최근에 등장한 'XXA0작업 결과'를 대분류로 삼으세요.
2. 전압/저항 라인: 'BATT', 'INV', 'Main Fuse', '고전압', '릴레이' 키워드가 포함된 컬럼은 대분류를 '전압 및 통전테스트'로 묶으세요.
3. 전류 라인: '0A', '100A', '-100A' 키워드가 포함된 컬럼은 대분류를 '전류측정결과'로 묶으세요.
4. 신규 연산 라인: 새로 생성한 '온도센서' 및 MIN/MAX/표준 등은 대분류를 '측정결과'로 묶고, 표준은 소수점 셋 째자리까지 표현합니다.
5. 불필요 컬럼 절삭: 최종 결과 컬럼 목록 중에서 '-100A / 릴레이 작동_Ave' 라는 이름이 포함된 컬럼의 위치(Index)를 찾은 뒤, 그 컬럼의 뒤에 있는 모든 컬럼은 result_df에서 완전히 삭제(Drop)하세요.
6. Unnamed 공백 처리 (추가됨): 컬럼명에 'Unnamed'라는 단어가 포함된 경우, 대분류와 소분류 모두 빈 문자열('')로 지정하여 엑셀 출력 시 헤더가 완전히 공백으로 보이도록 처리하세요.

[주의사항]
- 데이터를 새로 불러오는 코드는 절대 작성하지 마세요. (df1, df2, df3는 이미 존재함)
- 마크다운 기호를 제외하고 실행 가능한 순수 파이썬 코드만 도출하세요.
"""

response = client.chat.completions.create(
    model="gpt-5.2", 
    messages=[{"role": "user", "content": prompt}],
    temperature=0.1 
)

generated_code = response.choices[0].message.content.replace('```python', '').replace('```', '').strip()

print("--- 📊 이번 작업 토큰 사용량 ---")
print(f"입력(프롬프트) 토큰: {response.usage.prompt_tokens} 개")
print(f"출력(생성된 코드) 토큰: {response.usage.completion_tokens} 개")
print(f"총 소모 토큰: {response.usage.total_tokens} 개")
print("--------------------------------")

print("--- 🤖 AI가 작성한 코드 ---")
print(generated_code)
print("--------------------------")

# 5. 안전한 실행 환경 구축 및 실행
try:
    # df3를 반드시 추가해야 합니다.
    env = {"pd": pd, "df1": df1, "df2": df2, "df3": df3}
    
    exec(generated_code, env)
    
    if 'result_df' in env:
        # MultiIndex 구조가 잘 저장되도록 index=False 유지
        env['result_df'].index.name = None
        env['result_df'].to_excel('결과.xlsx', index=True)
        print("✅ 성공! '결과.xlsx' 파일이 생성되었습니다.")
    else:
        print("❌ AI가 코드를 실행했지만, 'result_df'라는 변수를 만들지 않았습니다.")
        
except Exception as e:
    print(f"❌ AI가 작성한 코드에서 에러 발생: {e}")