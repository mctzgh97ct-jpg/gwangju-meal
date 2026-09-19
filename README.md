# 광주대학교 오늘의 학식

광주대학교 공식 `진월광장 > 식당메뉴` 게시판의 최신 주간 엑셀 식단표를 읽어 웹에서 보기 쉽게 보여주는 Flask 앱입니다.

## 현재 기능
- 최신 식단 게시글 자동 탐색
- 첨부 엑셀 자동 다운로드
- 날짜별 학생정식 표시
- 날짜별 푸드단품 표시
- 학생정식 / 푸드단품 탭 전환
- 오늘 또는 가장 가까운 식단 날짜 자동 선택
- 모바일 대응

## 로컬 실행
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

브라우저에서 `http://127.0.0.1:5000` 접속.

## Render
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`

학교 엑셀 양식이 변경되면 `app.py`의 메뉴 구역 인식 규칙을 조정해야 할 수 있습니다.
