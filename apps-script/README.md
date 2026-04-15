# Apps Script Backend

이 Apps Script는 두 가지를 처리합니다.

- 출석 대상 스프레드시트 읽기/쓰기
- 매칭 기록 및 미매칭 큐 저장

## 저장용 스프레드시트

기본 저장 대상 스프레드시트는 아래 ID로 코드에 반영되어 있습니다.

- `1uFP6fsP37jxplPzk3-M2ONsuXDxJ2P1b2Fr-gGMDeC4`
- 링크: [backend spreadsheet](https://docs.google.com/spreadsheets/d/1uFP6fsP37jxplPzk3-M2ONsuXDxJ2P1b2Fr-gGMDeC4/edit?gid=0#gid=0)

이 스프레드시트에는 아래 시트가 자동 생성됩니다.

- `MatchingHistory`
- `UnmatchedQueue`

원하면 스크립트 속성 `MATCH_BACKEND_SPREADSHEET_ID`로 다른 저장용 스프레드시트로 바꿀 수 있습니다.

## 배포 방법

1. 새 standalone Google Apps Script 프로젝트를 만듭니다.
2. [Code.gs](./Code.gs)와 [appsscript.json](./appsscript.json)을 넣습니다.
3. 스크립트 속성에 아래 값을 추가합니다.
   - `MATCH_BACKEND_TOKEN`
   - `MATCH_BACKEND_SPREADSHEET_ID` (선택)
4. `Deploy > New deployment > Web app`으로 배포합니다.
5. 배포 URL을 `KAKAO_CHECK_APPS_SCRIPT_URL`에 넣습니다.

이 방식이면 기존 출석 시트에 바인딩된 Apps Script를 건드릴 필요가 없습니다.

## Apps Script가 받는 액션

- `listSheetTitles`
  - 출석 대상 스프레드시트의 탭 목록 조회
- `loadRosterRows`
  - 특정 탭의 명단 로드
- `writeUpdates`
  - 출석 상태 반영
- `listPending`
  - 현재 열린 미매칭 목록 조회
- `syncRun`
  - 실행 기록 저장 + 미매칭 큐 갱신

## Vercel 환경변수

- `KAKAO_CHECK_SPREADSHEET_ID`
  - 실제 출석 대상 스프레드시트 ID
- `KAKAO_CHECK_APPS_SCRIPT_URL`
  - Apps Script Web App URL
- `KAKAO_CHECK_APPS_SCRIPT_TOKEN`
  - Apps Script 호출 토큰

## 참고

- Apps Script는 토큰만 검증하고, 사용자 로그인/시트 권한은 Vercel API 쪽에서 처리합니다.
- 미리보기 모드에서는 `syncRun`이 호출되지 않습니다.
