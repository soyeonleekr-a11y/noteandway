# TRAVELOG DASHBOARD — Firebase + Netlify 배포 가이드

## 1. Firebase 프로젝트 만들기
1. https://console.firebase.google.com 접속 → "프로젝트 추가" → 이름 입력 (예: travelog-dashboard) → 생성
2. 왼쪽 메뉴 **Firestore Database** → "데이터베이스 만들기" → 프로덕션 모드 → 리전은 asia-northeast3(서울) 추천
3. 왼쪽 메뉴 **Authentication** → "시작하기" → 로그인 방법 탭 → **익명(Anonymous)** 활성화
4. Firestore Database → 규칙 탭 → 이 프로젝트의 `firestore.rules` 내용을 그대로 붙여넣고 "게시"
5. 프로젝트 설정(톱니바퀴) → 일반 탭 → 맨 아래 "내 앱" → 웹 앱 아이콘(`</>`) 클릭 → 앱 등록
6. 나오는 `firebaseConfig` 객체 값을 복사해서 `src/firebase.js`의 `firebaseConfig`에 그대로 붙여넣기

## 2. 로컬에서 확인 (선택)
Node.js가 설치되어 있다면:
```bash
npm install
npm run dev
```
브라우저에서 `http://localhost:5173` 접속해서 정상 작동하는지 확인하세요.

## 3. GitHub에 올리기
1. https://github.com 에서 새 저장소 생성 (예: travelog-dashboard)
2. 이 폴더 전체를 그 저장소에 업로드 (GitHub 웹의 "Add file → Upload files"로 드래그 앤 드롭 가능. `node_modules`, `dist` 폴더는 올리지 않아도 돼요 — `.gitignore`에 이미 제외되어 있어요)

## 4. Netlify 배포
1. https://app.netlify.com 가입/로그인 → "Add new site" → "Import an existing project" → GitHub 저장소 선택
2. Build command: `npm run build` / Publish directory: `dist` (netlify.toml에 이미 지정되어 있어서 자동으로 채워질 거예요)
3. "Deploy site" 클릭 → 몇 분 뒤 `https://랜덤이름.netlify.app` 주소로 접속 가능
4. Site settings → Domain management에서 원하는 서브도메인으로 변경 가능

## 5. (선택) 컨설팅 버튼 활성화하기
컨설팅(GATE 4) 버튼은 Anthropic API를 호출해요. Claude.ai 안에서는 무료로 동작했지만,
독립된 웹사이트에서는 본인 명의의 API 키가 필요하고 **호출할 때마다 아주 소액의 비용**이 들어요
(웹 검색 포함해도 보통 1회 호출에 몇 센트 수준이에요. 매달 한두 번만 누른다면 커피 한 잔 값도 안 돼요).

1. https://console.anthropic.com 에서 API 키 발급 (결제 수단 등록 필요)
2. Netlify → Site settings → Environment variables → `ANTHROPIC_API_KEY` 추가
3. Deploys 탭 → "Trigger deploy"로 재배포

이 단계를 건너뛰면 컨설팅 버튼만 빼고 나머지 기능(매출/게시글 관리, 인사이트 분석, 데이터 저장)은
**전부 완전 무료**로 계속 사용할 수 있어요.

## 데이터는 어디에 저장되나요?
모든 입력 데이터(플랫폼 매출, 게시글, 블로그 이름, 컨설팅 기록)는 Firebase Firestore의
`dashboardData` 컬렉션에 저장돼요. 이 규모(월 몇 건의 엑셀 업로드 + 개인 조회)는
Firebase 무료(Spark) 플랜과 Netlify 무료 플랜 한도에 비하면 아주 작아서, 계속 무료로 운영할 수 있어요.
