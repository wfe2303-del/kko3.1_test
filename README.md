타이탄 출석부 v7 - Public 저장소 업로드 안전판

핵심 변경
- 실제 spreadsheetId / clientId 를 repo에서 제거
- 로컬에서는 `js/config.runtime.local.js`(비추적)로 실행 가능
- Vercel 배포에서는 `/api/runtime-config`가 환경변수 값을 내려주도록 추가
- 허용 origin 검사 유지
- CSP(Content-Security-Policy) 메타 유지
- fetch 요청 no-store / no-referrer / credentials omit 적용 유지
- access token 만료 시 자동 로그아웃 처리 유지
- pagehide 시 메모리 토큰 정리 유지

왜 이렇게 바꿨나
- public 저장소에 실제 운영값이 남지 않도록 하기 위함
- 브라우저 직결 구조 특성상 배포된 앱에서는 clientId / spreadsheetId 를 완전히 숨길 수는 없지만, 최소한 repo에는 남기지 않도록 분리

파일 구조
- `js/config-loader.js` : 로컬 비추적 설정 파일과 서버 런타임 설정을 불러옴
- `js/config.runtime.local.example.js` : 로컬 실행용 예시 파일
- `api/runtime-config.js` : Vercel 환경변수 → 브라우저 설정 전달
- `.env.example` : 환경변수 예시

public 저장소 업로드 전제
- 이 repo 자체에는 실제 운영값을 넣지 않음
- 실제 값은 아래 둘 중 하나로만 주입
  1. 로컬 테스트: `js/config.runtime.local.example.js` 를 복사해서 `js/config.runtime.local.js` 생성 (Git 커밋 금지)
  2. Vercel 배포: Project Settings → Environment Variables 에 값 등록

로컬 테스트
1. `js/config.runtime.local.example.js` 를 복사해서 `js/config.runtime.local.js` 생성
2. 아래 값 입력
   - spreadsheetId
   - clientId
   - allowedOrigins
   - allowedEmailDomains / allowedEmails
3. 정적 서버로 실행
4. 브라우저에서 로그인 후 탭 조회 / 데모 / 실제 기록 순서로 검증

Vercel 배포
1. Vercel 프로젝트 생성
2. 아래 환경변수 등록
   - `KAKAO_CHECK_SPREADSHEET_ID`
   - `KAKAO_CHECK_GOOGLE_CLIENT_ID`
   - `KAKAO_CHECK_ALLOWED_ORIGINS`
   - `KAKAO_CHECK_ALLOWED_EMAIL_DOMAINS`
   - `KAKAO_CHECK_ALLOWED_EMAILS`
   - `KAKAO_CHECK_GOOGLE_LOGIN_HINT`
   - `KAKAO_CHECK_GOOGLE_HOSTED_DOMAIN_HINT`
3. Google Cloud Console 에서 OAuth Authorized JavaScript origins 를 실제 배포 주소와 동일하게 등록
4. 대상 스프레드시트 공유 권한을 링크 공개가 아니라 필요한 사용자/도메인으로만 제한

중요
- public repo 안전판이지, 배포된 브라우저 앱에서 값을 완전히 숨기는 구조는 아님
- 더 강한 보안이 필요하면 쓰기 작업을 Apps Script Web App 또는 별도 백엔드로 이동 권장
