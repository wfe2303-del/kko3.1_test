# Apps Script Match Backend

이 템플릿은 현재 프론트 기능을 유지한 채로 아래 2가지를 추가하기 위한 백엔드입니다.

- 매 실행의 매칭 기록 저장
- 미매칭 인원 큐 저장 및 다음 실행에서 재매칭 처리

## 저장 시트

Apps Script가 연결된 스프레드시트에 아래 시트를 자동으로 만듭니다.

- `MatchingHistory`
  - 실행 시각, 실행자, 시트명, 로그 요약, 미매칭 수, 재매칭 수 등을 저장
- `UnmatchedQueue`
  - 아직 명단에 연결되지 않은 사람을 `OPEN`
  - 다음 실행에서 해소되면 `RESOLVED`

기본 저장 대상 스프레드시트는 아래 파일에 이미 반영되어 있습니다.

- [`Code.gs`](./Code.gs) 기본값: `1uFP6fsP37jxplPzk3-M2ONsuXDxJ2P1b2Fr-gGMDeC4`
- 시트 링크: [backend spreadsheet](https://docs.google.com/spreadsheets/d/1uFP6fsP37jxplPzk3-M2ONsuXDxJ2P1b2Fr-gGMDeC4/edit?gid=0#gid=0)

## Apps Script 배포

1. 새 Google Apps Script 프로젝트를 만듭니다.
2. [`Code.gs`](./Code.gs)와 [`appsscript.json`](./appsscript.json)을 사용합니다.
3. 스크립트 속성에 아래 값을 넣습니다.
   - `MATCH_BACKEND_TOKEN`
   - `MATCH_BACKEND_SPREADSHEET_ID` (선택)
4. `Deploy > New deployment > Web app` 으로 배포합니다.
5. 실행 권한은 본인, 접근은 배포 환경에 맞게 설정합니다.

`MATCH_BACKEND_SPREADSHEET_ID`를 비워두면 위 기본 스프레드시트를 사용합니다.

즉, 기존 출석 시트에 바인딩된 Apps Script를 수정할 필요 없이 이 `apps-script` 폴더를 기준으로 별도 standalone 프로젝트를 만들어 배포하면 됩니다.

## Vercel 환경변수

프론트는 Apps Script를 직접 호출하지 않고 Vercel API를 통해 프록시합니다.

- `KAKAO_CHECK_APPS_SCRIPT_URL`
- `KAKAO_CHECK_APPS_SCRIPT_TOKEN`

기존 값도 그대로 필요합니다.

- `KAKAO_CHECK_SPREADSHEET_ID`
- `KAKAO_CHECK_GOOGLE_CLIENT_ID`
- `KAKAO_CHECK_ALLOWED_ORIGINS`
- `KAKAO_CHECK_ALLOWED_EMAIL_DOMAINS`
- `KAKAO_CHECK_ALLOWED_EMAILS`

## 동작 흐름

1. 사용자가 로그 파일로 출석 체크를 실행합니다.
2. 프론트가 기존 `OPEN` 미매칭 목록을 읽습니다.
3. 이번 로그와 이전 미매칭을 함께 비교해서 재매칭 가능한 건을 찾습니다.
4. Google Sheet 출석 반영 후 실행 결과를 `MatchingHistory`에 저장합니다.
5. 이번 실행에서 아직 해소되지 않은 사람은 `UnmatchedQueue`에 누적합니다.

## 참고

- `미리보기` 모드에서는 기록을 저장하지 않습니다.
- Apps Script 백엔드가 없으면 기존 출석 체크 기능은 그대로 동작하고, 기록 저장만 건너뜁니다.
