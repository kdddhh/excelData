import requests

# API 서버 주소 (GET 방식으로 호출)
API_URL = "http://192.168.200.201:8402/download-excel"

print("서버에 데이터 처리를 요청하고 다운로드를 기다리는 중입니다...")

# 서버에 파일 없이 요청만 보냄
response = requests.get(API_URL)

if response.status_code == 200:
    # 📝 호출자가 이 파일을 저장하고 싶은 본인 컴퓨터의 경로 지정
    save_path = r"C:\Users\kados\OneDrive\바탕 화면\클레임 엑셀 파일\서버에서_가져온_결과.xlsx"
    
    with open(save_path, 'wb') as f:
        f.write(response.content)
    
    print(f"🎉 성공! 엑셀 파일이 {save_path} 에 저장되었습니다.")
else:
    print("❌ 에러 발생:", response.text)