# Overtime 초과근무 기록 시스템

Google Sheets를 데이터 저장소로 사용하는 초과근무 기록용 Vercel 웹앱입니다.

## 주요 기능

- PIN 비밀번호 보호: 기본값 `0992`
- 월별 초과근무 조회 및 저장
- 날짜별 메모 저장: 월 시트의 `H`열 사용
- 지금 시작 / 지금 종료 빠른 저장
- 방과후 요일 설정 저장
- 월별 CSV 다운로드 및 PDF 인쇄
- 수정 이력 기록: `수정이력` 시트 자동 생성
- PWA 설치 지원
- 새 월 시트 생성 전 형식, 중복, 템플릿 시트 검증

## Vercel 환경변수

필수:

- `GOOGLE_CLIENT_EMAIL`: Google 서비스 계정 이메일
- `GOOGLE_PRIVATE_KEY`: Google 서비스 계정 private key
- `GOOGLE_SHEET_ID`: 연결할 Google Spreadsheet ID

선택:

- `APP_PIN`: 앱 비밀번호. 설정하지 않으면 `0992`
- `TEMPLATE_SHEET_NAME`: 기본값 `양식`
- `DB_SHEET_NAME`: 기본값 `통합DB`
- `SETTINGS_SHEET_NAME`: 기본값 `설정`
- `HISTORY_SHEET_NAME`: 기본값 `수정이력`

## Google Sheet 구조

월별 시트는 다음 열을 사용합니다.

| 열 | 용도 |
| --- | --- |
| A | 날짜 |
| C | 시작 시간 |
| D | 종료 시간 |
| E | 아침 초과 |
| F | 오후 초과 |
| G | 일계 |
| H | 메모 |
| J | 요약 값 |
| K | 방과후 요일 |
| L | 방과후 선택 여부 |

`양식` 시트가 새 월 시트의 템플릿입니다. 새 월 이름은 `26-5` 또는 `2026-5` 형식으로 입력합니다.
