# TRAVELOG DASHBOARD: Claude Code 작업 메모

이 저장소는 여행 블로그 '요니의 여행수첩'(요니)과 서브 블로그(짠순)의 **제휴 수익·포스팅 성과 대시보드**야.
운영자는 이하연(하연)이고, 처음 코드는 동생(soyeonleekr-a11y)이 만들었어.
- 배포 주소: https://letsgonoteandway.netlify.app/
- 원본 저장소: soyeonleekr-a11y/noteandway (하연이는 협업자, Write 권한)
- 하연이 포크: yony9751/noteandway

## 사용자에 대해
- 하연이는 **웹퍼블리셔 10년 차**야. HTML/CSS/UI는 전문가 수준이지만 React, 빌드 도구, 서버 함수, Firebase는 아직 익숙하지 않아.
  → 새 개념(React 상태, Netlify Functions, Firestore 규칙 등)이 나오면 **짧고 쉽게 설명**하면서 진행해줘.
- 대화는 친근한 한국어 반말. 화면 문구(UI 카피)도 지금처럼 친근한 '~해요' 체로 맞춰줘.
- 디자인 완성도(여백, 타이포, 정보 위계)를 중요하게 생각해. UI를 바꿀 땐 기존 다크 네이비 + 골드 톤을 유지해줘.

## 구조
```
CLAUDE.md                       ← 이 파일
.claude/settings.json           ← Claude Code 설정 (훅 등록)
.claude/hooks/guard.mjs         ← 위험한 git/gh 명령과 비밀값 커밋을 막는 훅 (아래 '금지사항')
travelog-dashboard (1)/         ← 실제 앱 폴더 (폴더 이름에 공백과 괄호가 있으니 경로를 따옴표로 감싸줘)
  package.json                  ← "type": "module" (ES 모듈)
  netlify.toml                  ← build: npm run build, publish: dist, functions: netlify/functions
  firestore.rules               ← Firestore 보안 규칙 (콘솔에 직접 게시해야 적용됨)
  src/App.jsx                   ← 앱 전체 (약 1,450줄, 컴포넌트·스타일이 한 파일에 있음)
  src/firebase.js               ← Firebase 설정 (apiKey는 공개 식별자라 노출돼도 괜찮음)
  netlify/functions/consult.js  ← 컨설팅용 Anthropic API 호출 (ANTHROPIC_API_KEY 환경변수 필요)
```
- 스택: React 18 + Vite 5, recharts(차트), xlsx/papaparse(엑셀·CSV 파싱), lucide-react(아이콘), Firebase(Firestore + 익명 Auth)
- 탭: 개요(OverviewTab) / 플랫폼 매출(RevenueTab) / 블로그 운영(PostsTab) / 인사이트(InsightTab) / 컨설팅(ConsultingTab)
- 데이터: Firestore `dashboardData` 컬렉션에 문서 하나당 키 하나 (`safeGet`/`safeSet`)
  - `platform-data`: {플랫폼: {YYYY-MM: {total, products:[{name, qty, unitRevenue, totalRevenue}]}}}
  - `posts-data`, `post-records`, `blog-names`, `consulting-history`
- 첫 화면의 강아지 잠금(PatternGate)은 **클라이언트에서만 동작하는 화면 가림**이야. 보안 장치가 아니야.

## 자주 쓰는 명령 (앱 폴더 안에서)
```
cd "travelog-dashboard (1)"
npm install
npm run dev      # http://localhost:5173 에서 확인
npm run build    # 배포 전에 빌드 오류가 없는지 꼭 확인
```
컨설팅 함수까지 로컬에서 테스트하려면 Netlify CLI(`npx netlify dev`)와 ANTHROPIC_API_KEY가 필요해.

## 작업 규칙
- **main 브랜치에 직접 푸시하지 마.** `hayeon/작업이름` 브랜치를 만들어서 작업하고, PR을 만든 뒤 하연이가 확인하고 머지해.
- 커밋 전에 `npm run build`가 통과하는지 확인해.
- Firestore에 이미 저장된 데이터 구조를 바꿀 때는, 예전 형태도 읽을 수 있게 변환 코드를 넣어줘 (`normalizeProduct`처럼).
- 실제 매출 데이터를 지우거나 덮어쓰는 코드는 실행하기 전에 하연이한테 먼저 물어봐.

## 금지사항
아래 ❌ 항목은 `.claude/hooks/guard.mjs` 훅이 Bash 명령 실행 직전에 검사해서 **자동으로 막아** (표시: 🔒). 훅이 막으면 이유를 읽고 우회하지 말고 하연이한테 물어봐.

- ❌ 🔒 main/master에 직접 푸시 금지. main 브랜치 위에서 바로 커밋하는 것도 금지야.
- ❌ 🔒 force 푸시(`-f`, `--force`), `git push --all` 금지. (`--force-with-lease`는 작업 브랜치에서만 가능)
- ❌ 🔒 PR 머지(`gh pr merge`) 금지. 머지는 하연이가 확인하고 직접 해.
- ❌ 🔒 `.env`, `CLAUDE.local.md`, 정산 엑셀/CSV, 인증서·비밀키 커밋 금지. `git add -f`로 .gitignore를 우회하는 것도 금지야.
- ❌ 🔒 API 키(`sk-ant-...` 등), 매출 금액(100만 원 이상), 개인 이메일을 코드·문서·커밋 메시지에 넣지 마. 저장소가 공개야. 필요한 건 `CLAUDE.local.md`에 적어.
- ❌ Firestore 실제 데이터(`dashboardData`, `allowedUsers`)를 삭제하거나 덮어쓰지 마. 코드든 일회성 스크립트든 하연이가 OK 하기 전에는 실행하지 않아.
- ❌ `firestore.rules`를 느슨하게 바꾸지 마. (`if true`, 로그인만 확인하는 `request.auth != null` 등)
- ❌ `ANTHROPIC_API_KEY`를 프론트(`src/`)에서 쓰거나 코드에 적지 마. Netlify 함수에서만 환경변수로 읽어.

훅의 한계: Claude Code 안에서 실행하는 명령만 검사해. 터미널에서 직접 입력하는 git 명령은 막지 못하니까 하연이도 같은 규칙을 지켜줘.

## 공개 저장소 주의
- 2026-11-06부터 **저장소가 공개(public)야.** (Netlify 무료 플랜에서 하연이 커밋도 빌드되게 하려고 바꿈)
- 매출 숫자, 이메일 같은 개인 정보는 커밋하지 말고 `CLAUDE.local.md`(커밋 안 됨)에 적어.
- API 키는 Netlify 환경변수에만 둬. 코드나 커밋에 넣지 마.

## 배포 상태 (2026-11-06 기준)
- main에 머지하면 Netlify가 **자동으로 배포돼.** (동생이 GitHub 연결 + 저장소 공개로 해결. 배포 파일이 `index-Dce6VQrB.js` → `index-Kte3X-8h.js`로 바뀐 걸 확인했어.)
- 안 바뀌면 Netlify > Site configuration > Build & deploy에서 **Base directory가 `travelog-dashboard (1)`인지** 확인해.
- 컨설팅 버튼이 실제로 동작하는지는 아직 확인 못 했어. (Netlify 환경변수 `ANTHROPIC_API_KEY` 포함)
- Netlify 권한을 받으면 `npx netlify login` → `npx netlify link`로 연결해서 배포 로그를 직접 확인해줘.

## 할 일 (우선순위 순)
1. ~~배포 연결 해결~~ **완료** (자동 배포 확인). 남은 것: 컨설팅 버튼이 실제로 동작하는지, Netlify 환경변수에 `ANTHROPIC_API_KEY`가 있는지 확인.
2. **보안 (중요)**
   - ✅ 완료: Google 로그인(강아지 잠금 교체), `allowedUsers/{이메일}` 허용 목록, `firestore.rules` 게시 (하연이가 콘솔 설정을 마쳤다고 알려줌)
   - ✅ 코드 완료 (PR: hayeon/consult-auth): 컨설팅 함수(`netlify/functions/consult.js`)가 로그인 토큰(Authorization: Bearer)을 받아서, 그 토큰으로 Firestore `allowedUsers/{이메일}`을 읽을 수 있을 때만 Anthropic을 호출해. 확인이 안 되면 항상 막는 쪽으로 처리하고, 프롬프트는 10,000자까지만 받아. 허용 목록은 Firestore 규칙을 그대로 따라가.
   - ⬜ 남음: 배포 후 **실제 로그인한 상태로 컨설팅 버튼이 되는지** 확인 (진짜 토큰으로는 아직 못 해봤어). Anthropic 콘솔에서 월 사용 한도도 걸어두면 안전해.
   - ⬜ 확인: Firebase Authentication에서 익명 로그인이 꺼져 있는지.
3. **지역 분류 개선** (`DESTINATIONS`, `matchDestination`) — PR: hayeon/region-classify
   - 매출 대부분이 '기타'로 빠져서 인사이트 '종합 TOP 10' 1위가 '기타'로 나왔어.
   - 사전에 없는 지명: 제천, 설악산, 청송, 다카마쓰, 나오시마, 마쓰야마, 나고야, LA, 라스베가스, 그랜드캐년, 샌프란시스코, 요세미티, 인스파이어(인천), 북한산(서울) 등
   - 버그: `find()`가 배열 순서로 첫 매칭을 고르는 방식이라 "보코 서울 명동 … 싱가포르 뷔페"가 싱가포르로 분류돼. → **텍스트에서 가장 먼저 나오는 지명**을 고르게 바꾸거나, 국내 지명을 우선 매칭해.
   - 분류가 안 된 항목은 화면에서 직접 지역을 지정할 수 있게 (수동 태깅, Firestore에 저장)
4. **상품 합치기 개선**: 마이리얼트립 상하이 우전 수향마을 야경투어가 문구만 다른 이름으로 3~4줄로 쪼개져 있어. 정산 엑셀에 상품 ID 컬럼이 있으면 ID 기준으로 합치고, 없으면 대괄호·★ 문구를 뺀 핵심 이름으로 묶는 방식을 검토해.
5. **이심 카테고리**: 마이리얼트립 상품명 "중국", "일본"(각 230건대)은 마이심 이심이야. '이심' 상품유형으로 분류해.
6. **링크프라이스 추가**: PLATFORMS에 추가하되 하위 머천트(아고다, 트립닷컴, 호텔스닷컴, 익스피디아, 에어알로)별로 나눠서 집계해. 확정/미확정(취소) 구분도 필요해. 하연이가 엑셀 샘플을 주면 컬럼 매핑부터 설계해.
7. **UX**: 첫 화면 기본 월을 '데이터가 있는 가장 최근 달'로 바꿔 (지금은 데이터가 없는 이번 달이라 ₩0으로 보여). 추이 그래프에서 데이터가 없는 달은 0 대신 비워두거나 점선으로 표시해.
8. **나중에 (마케팅 기능)**
   - 글별 수익: 포스팅마다 추적값(서브 ID 등)을 달아 정산 엑셀과 매칭 → 조회수 1,000당 수익
   - 커미션 상향 기간을 매출 그래프 위에 표시
   - 상금·원고료를 커미션과 따로 집계
   - 조회수는 높은데 수익이 낮은 글, 시즌이 다가오는 글을 '다시 손볼 글'로 추천
   - 컨설팅 프롬프트를 '여행지 추천'에서 '다음 달 포스팅 리스트(플랫폼·코드 포함)'로 개선

## 참고 데이터
- 매출·조회수 같은 실제 숫자는 `CLAUDE.local.md`에 있어.
