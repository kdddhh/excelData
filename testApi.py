import requests
import os

# API 서버 주소 (GET 방식으로 호출)
API_URL = os.getenv("API_URL")

print("서버에 데이터 처리를 요청하고 다운로드를 기다리는 중입니다...")

# 서버에 파일 없이 요청만 보냄
response = requests.get(API_URL)

if response.status_code == 200:
    # 📝 호출자가 이 파일을 저장하고 싶은 본인 컴퓨터의 경로 지정
    
    with open(os.getenv("SAVE_PATH"), 'wb') as f:
        f.write(response.content)
    
    print(f"🎉 성공! 엑셀 파일이 {os.getenv("SAVE_PATH")} 에 저장되었습니다.")
else:
    print("❌ 에러 발생:", response.text)