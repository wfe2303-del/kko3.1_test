# Kakao Check 3.0

카카오톡 입장 로그 파일을 업로드해서 출석을 체크하는 웹앱입니다.

현재 구조:

- 로그인: 웹앱 전용 `아이디 / 비밀번호`
- 세션: same-origin HttpOnly 쿠키
- 시트 읽기/쓰기: Vercel API -> Apps Script Web App
- 미매칭/실행 기록 저장: Apps Script 백엔드

즉, 사용자는 Google 로그인을 하지 않아도 되고, 웹앱 계정 권한으로 출석 체크를 수행합니다.

## 환경변수

기본 예시는 [.env.example](/C:/Users/user/Downloads/kakaocheck_3.0-main/.env.example:1)에 있습니다.
로컬에서는 `.env.local`을 사용하면 됩니다.

- `KAKAO_CHECK_SPREADSHEET_ID`
  - 실제 출석 대상 스프레드시트 ID
- `KAKAO_CHECK_ALLOWED_ORIGINS`
  - 허용할 배포 주소 목록
  - `https://*.vercel.app` 형태의 와일드카드도 지원
- `KAKAO_CHECK_SESSION_SECRET`
  - 세션 서명용 긴 랜덤 문자열
- `KAKAO_CHECK_USERS_JSON`
  - 로그인 가능한 계정 목록
- `KAKAO_CHECK_APPS_SCRIPT_URL`
  - Apps Script Web App 배포 URL
- `KAKAO_CHECK_APPS_SCRIPT_TOKEN`
  - Apps Script 백엔드 토큰

## 사용자 계정 형식

`KAKAO_CHECK_USERS_JSON` 예시:

```json
[
  {
    "username": "manager",
    "password": "change-me",
    "displayName": "운영자",
    "allowedSheets": ["*"],
    "canWrite": true
  },
  {
    "username": "staff-a",
    "password": "change-me-too",
    "displayName": "스태프 A",
    "allowedSheets": ["4월 1주", "4월 2주"],
    "canWrite": true
  }
]
```

설명:

- `allowedSheets`
  - `["*"]` 이면 모든 시트 접근 가능
  - 배열로 넣으면 지정한 탭만 접근 가능
- `canWrite`
  - `false`면 읽기만 가능

원하면 `password` 대신 `passwordSha256`도 사용할 수 있습니다.

## Vercel 배포

이 프로젝트는 별도 빌드 없이 정적 파일 + `api/` 서버리스 함수 조합으로 배포됩니다.

### 1. Vercel 프로젝트 연결

1. 저장소를 Vercel에 연결합니다.
2. Framework Preset은 자동 감지 또는 `Other`로 둬도 됩니다.
3. Build Command는 비워둬도 됩니다.

`vercel.json`은 이미 추가되어 있어 `api/**/*.js` 함수 타임아웃만 지정해둔 상태입니다. [vercel.json](/C:/Users/user/Downloads/kakaocheck_3.0-main/vercel.json:1)

### 2. Vercel 환경변수 추가

아래 6개를 프로젝트 환경변수에 넣습니다.

- `KAKAO_CHECK_SPREADSHEET_ID`
- `KAKAO_CHECK_ALLOWED_ORIGINS`
- `KAKAO_CHECK_SESSION_SECRET`
- `KAKAO_CHECK_USERS_JSON`
- `KAKAO_CHECK_APPS_SCRIPT_URL`
- `KAKAO_CHECK_APPS_SCRIPT_TOKEN`

현재 사용자 환경 기준으로는 `KAKAO_CHECK_ALLOWED_ORIGINS`를 최소 이렇게 두는 걸 추천합니다.

```txt
https://your-production-domain.com,https://*.vercel.app
```

커스텀 도메인이 없으면 예를 들어:

```txt
https://your-project-name.vercel.app,https://*.vercel.app
```

### 3. 주의할 점

- Preview URL도 쓰려면 `https://*.vercel.app`를 넣는 게 편합니다.
- 세션 쿠키는 HTTPS에서 `Secure`로 붙습니다. Vercel 배포 환경에서는 정상 동작합니다.
- `KAKAO_CHECK_USERS_JSON`은 민감 정보라 Vercel 환경변수에만 넣고 저장소에는 넣지 않습니다.

## 로컬 개발

로컬은 `vercel dev` 기준으로 맞춰두었습니다.

```bash
npm install
npm run dev
```

또는 Vercel CLI가 이미 있으면 바로 `vercel dev`로 실행해도 됩니다.

## Apps Script 백엔드

Apps Script는 두 역할을 합니다.

- 출석 대상 스프레드시트에서 탭 목록 조회, 명단 로드, 출석 상태 반영
- `MatchingHistory`, `UnmatchedQueue` 저장

자세한 배포 방법은 [apps-script/README.md](/C:/Users/user/Downloads/kakaocheck_3.0-main/apps-script/README.md:1)를 보면 됩니다.

## 체크 포인트

1. Apps Script Web App이 배포되어 있는지 확인
2. Vercel 환경변수 6개가 모두 들어갔는지 확인
3. `KAKAO_CHECK_ALLOWED_ORIGINS`에 현재 배포 주소가 포함됐는지 확인
4. 로그인 후 시트 목록이 보이는지 확인
5. 로그 파일 업로드 후 미리보기/실행 테스트

## 주의

- 브라우저가 Google Sheets API를 직접 호출하지 않습니다.
- 사용자 권한은 Google 계정이 아니라 `KAKAO_CHECK_USERS_JSON`에 정의된 웹앱 계정 기준입니다.
- 운영 환경에서는 `password` 대신 `passwordSha256` 사용을 권장합니다.
