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
  - `platform-data`: {플랫폼: {YYYY-MM: {total, products:[{name, qty, unitRevenue, totalRevenue, category?, salesAmount?}], source?}}}
    - `source`: 업로드 때 읽은 열 이름 {name, revenue, qty, sales, category, rows, uploadedAt}. 없으면 열 이름 지정 전에 올린 예전 데이터야.
    - `salesAmount`: 판매금액 합계 (판매금액 열을 지정한 플랫폼만). **수수료율 = 정산금(totalRevenue) ÷ 판매금액 × 100** (`commissionRate`)
    - `category`: 엑셀의 '상품군' (마이리얼트립만, 있을 때만 저장. Firestore는 undefined를 못 저장해)
  - `region-tags`: {태그키: {region, destination}}. 상품 키는 `product::플랫폼::핵심이름`(productGroupKey), 예전 키 `product::플랫폼::원래이름`도 계속 읽어.
  - `region-rules`: [{keyword, destination, region, createdAt}]. 하연이가 인사이트 탭 '자동 분류 규칙'에서 직접 추가하는 "단어 → 지역" 규칙 (`matchRule`).
- **지역을 정하는 순서** (`classifyText`): ① 직접 지정(`region-tags`) → ② 하연이 규칙(`region-rules`) → ③ 기본 사전(`DESTINATIONS` + `PLACE_ALIASES`)
- **마이리얼트립 정산 엑셀 열**: 예약일 · 마케팅 링크 ID · 예약 ID · 주문 구분 · 상품군 · 상품명 · 정산 기준일 · 판매 수량 · 판매 금액 · 정산 기준 금액 · 수익 금액 · 제세공과금 · 정산 예상 금액
  - **수익 = `정산 예상 금액`** (하연이가 정함). `PLATFORM_COLUMNS`에 열 이름을 정확히 적어두고 그 이름으로만 찾아.
  - 예전 코드는 단어가 들어간 첫 열을 골라서 상품명 대신 `상품군`, 수익 대신 `판매 금액`을 읽었어. 2026-11-06 이전에 올린 마이리얼트립 달은 다시 올려야 해 (매출 탭에 안내가 떠).
  - '주문 구분'이 **환불완료**인 줄은 정산 예상 금액이 마이너스로 들어와서, 전부 더하면 환불이 자동으로 빠져. (판매 수량은 확인 필요)
  - 하연이가 다시 올려서 총매출이 엑셀과 맞는 걸 확인했어 (2026-11-06).
- **여기어때 정산 엑셀**: **수익 = `정산금액`** (하연이 확인). 왼쪽에 있는 `실적 인정 금액`(판매금액)을 수익으로 읽으면 안 돼.
  - 상품명·수량 열 이름은 아직 몰라서 단어로 찾아. 정확한 열 이름을 받으면 `PLATFORM_COLUMNS`의 여기어때에 `name`, `qty`를 추가해.
  - **판매금액 = `실적 인정 금액`** (하연이 확인, 실제 엑셀 열 이름). 혹시 몰라 `총 실적 인정 금액`도 두 번째 후보로 찾아. 수수료율 계산에만 써.
  - `PLATFORM_COLUMNS`의 열 이름은 문자열 하나 또는 후보 배열(앞쪽 우선)로 쓸 수 있어. 비교할 땐 띄어쓰기·눈에 안 보이는 글자를 무시해.
- **수수료율 화면** (PR: hayeon/commission-rate)
  - 매출 탭: 총매출 위에 판매금액·정산금·수수료율 카드 + 분류별(상품 유형 배지와 같은 분류) 표. 상품 표에 판매금액·수수료율 열과 정렬.
  - 인사이트 탭 '플랫폼별 수수료율 비교': 분류 × 플랫폼 표, 분류별로 가장 높은 곳에 '유리'. '전체' 줄엔 '유리'를 안 붙여 (상품 구성에 따라 달라져서). 아래 '같은 상품 비교'는 핵심 이름이 똑같은 상품만 잡혀.
- **클룩**: 열 이름 미확인. 예전 방식(단어가 든 첫 열) 그대로야.
- 매출 탭 '다시 올려주세요' 안내: 수익 열을 지정한 플랫폼에서, 저장된 `source.revenue`가 지정한 열과 다르면(또는 기록이 없으면) 떠.
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
아래 ❌ 항목은 `.claude/hooks/guard.mjs` 훅이 Bash·PowerShell 명령 실행 직전에 검사해서 **자동으로 막아** (표시: 🔒). 훅이 막으면 이유를 읽고 우회하지 말고 하연이한테 물어봐.

- ❌ 🔒 main/master에 직접 푸시 금지. main 브랜치 위에서 바로 커밋하는 것도 금지야.
- ❌ 🔒 force 푸시(`-f`, `--force`), `git push --all` 금지. (`--force-with-lease`는 작업 브랜치에서만 가능)
- ❌ 🔒 저장소를 바꾸는 gh 명령 금지: `gh repo delete/edit/rename`, `gh api`의 DELETE·PUT·PATCH 요청, `gh pr merge --admin`
- ❌ 하연이가 채팅에서 **"머지해"라고 하기 전에는** PR을 머지하지 마. 머지하면 바로 실제 사이트에 배포돼.
  - 🔐 `gh pr merge`는 훅이 막지 않고 **터미널 확인창**을 띄워. 하연이가 승인해야 실행돼.
- ❌ 🔒 `.env`, `CLAUDE.local.md`, 정산 엑셀/CSV, 인증서·비밀키 커밋 금지. `git add -f`로 .gitignore를 우회하는 것도 금지야.
- ❌ 🔒 API 키(`sk-ant-...` 등), 매출 금액(100만 원 이상), 개인 이메일을 코드·문서·커밋 메시지에 넣지 마. 저장소가 공개야. 필요한 건 `CLAUDE.local.md`에 적어.
- ❌ Firestore 실제 데이터(`dashboardData`, `allowedUsers`)를 삭제하거나 덮어쓰지 마. 코드든 일회성 스크립트든 하연이가 OK 하기 전에는 실행하지 않아.
- ❌ `firestore.rules`를 느슨하게 바꾸지 마. (`if true`, 로그인만 확인하는 `request.auth != null` 등)
- ❌ `ANTHROPIC_API_KEY`를 프론트(`src/`)에서 쓰거나 코드에 적지 마. Netlify 함수에서만 환경변수로 읽어.

훅의 한계: Claude Code 안에서 실행하는 명령만 검사해. 터미널에서 직접 입력하는 git 명령은 막지 못하니까 하연이도 같은 규칙을 지켜줘.

## PR 만들기·머지 (GitHub CLI)
- `gh`가 설치돼 있고 `yony9751` 계정으로 로그인돼 있어 (원본 저장소 Write 권한). Claude Code 세션에서 `gh`가 안 잡히면 `"/c/Program Files/GitHub CLI/gh.exe"`로 실행해.
- 흐름: 작업 브랜치 푸시 → `gh pr create --repo soyeonleekr-a11y/noteandway --base main` → 바뀐 내용을 하연이한테 요약 → 하연이가 "머지해" → `gh pr merge <번호> --merge --delete-branch` → 확인창 승인 → Netlify 자동 배포
- 여러 PR이 서로 쌓여 있으면(한 브랜치가 다른 브랜치 위에서 시작) 아래쪽 PR부터 머지해.

## 공개 저장소 주의
- 2026-11-06부터 **저장소가 공개(public)야.** (Netlify 무료 플랜에서 하연이 커밋도 빌드되게 하려고 바꿈)
- 매출 숫자, 이메일 같은 개인 정보는 커밋하지 말고 `CLAUDE.local.md`(커밋 안 됨)에 적어.
- API 키는 Netlify 환경변수에만 둬. 코드나 커밋에 넣지 마.

## 배포 상태 (2026-11-06 기준)
- main에 머지하면 Netlify가 **자동으로 배포돼.** (동생이 GitHub 연결 + 저장소 공개로 해결. 배포 파일이 `index-Dce6VQrB.js` → `index-Kte3X-8h.js`로 바뀐 걸 확인했어.)
- 안 바뀌면 Netlify > Site configuration > Build & deploy에서 **Base directory가 `travelog-dashboard (1)`인지** 확인해.
- 컨설팅 버튼이 실제 사이트에서 정상 동작하는 걸 확인했어. (Netlify 환경변수 `ANTHROPIC_API_KEY` 등록 + 재배포 완료, Anthropic은 선불 크레딧 사용)
- Netlify 권한을 받으면 `npx netlify login` → `npx netlify link`로 연결해서 배포 로그를 직접 확인해줘.

## 할 일 (우선순위 순)
1. ~~배포 연결 해결~~ **완료** (자동 배포 + 컨설팅 동작 확인)
2. **보안 (중요)**
   - ✅ 완료: Google 로그인(강아지 잠금 교체), `allowedUsers/{이메일}` 허용 목록, `firestore.rules` 게시 (하연이가 콘솔 설정을 마쳤다고 알려줌)
   - ✅ 완료 (PR: hayeon/consult-auth, 실제 로그인 상태로 동작 확인): 컨설팅 함수(`netlify/functions/consult.js`)가 로그인 토큰(Authorization: Bearer)을 받아서, 그 토큰으로 Firestore `allowedUsers/{이메일}`을 읽을 수 있을 때만 Anthropic을 호출해. 확인이 안 되면 항상 막는 쪽으로 처리하고, 프롬프트는 10,000자까지만 받아. 허용 목록은 Firestore 규칙을 그대로 따라가.
   - ⬜ 확인: Firebase Authentication에서 익명 로그인이 꺼져 있는지.
3. ~~지역 분류 개선~~ **완료** (PR: hayeon/region-classify)
   - `matchDestination`: 텍스트에서 **먼저 나오는 지명**을 고르고, 나라 이름(`country: true`)은 도시 이름이 없을 때만 써. `MATCH_IGNORE`로 "아로마→로마" 같은 오분류를 막아.
   - 블로그 글은 화면에 보여줄 때마다 다시 분류해 (예전에 저장된 지역값은 안 씀)
   - 인사이트 탭 '지역 직접 지정하기' 패널에서 수동 태깅 → Firestore `region-tags`
   - **이제 하연이가 규칙을 화면에서 직접 추가해** (인사이트 탭 > 지역 직접 지정하기 > 📌 자동 분류 규칙, PR: hayeon/region-rules). 입력하는 동안 "미분류 N건이 바로 분류돼요", "이미 분류된 N건도 바뀌어요"를 미리 보여줘. 두 글자 이상만 받아.
   - 코드에 넣는 기본 별칭은 `PLACE_ALIASES` (누구에게나 기본으로 적용할 것만). 지역 이름은 `DESTINATIONS`의 이름과 같아야 하고, 없는 지역이면 먼저 `D(...)`로 추가해.
     - 지금까지: 페더데일·블루마운틴→시드니, 할레아칼라·거북이스노쿨링+해양5종→하와이, 타마린드 빌리지→치앙마이, 게이세이·스카이라이너→도쿄, 고와쿠엔·도고온천→마쓰야마
     - "거북이 스노쿨링"만으로는 하와이로 보지 않아 (괌·세부에도 있어서). 규칙은 구체적인 이름으로 받아.
   - 지명을 못 찾으면 띄어쓰기를 지우고 한 번 더 찾아 (4글자 이상 키워드만. 짧은 지명은 "필로 마사지"→로마처럼 우연히 생겨서 제외)
4. ~~상품 합치기~~ **완료** (PR: hayeon/mrt-columns): 정산 엑셀에 상품 ID가 없어서 핵심 이름(`productGroupKey`)으로 묶어. 대괄호류·★…★ 문구·이모지·기호·띄어쓰기는 빼고, 소괄호 (1인) 같은 건 상품 구분 정보라 남겨. 저장된 원본은 그대로 두고 화면에서만 묶어(`groupProducts`). 대표 이름은 수익이 가장 큰 이름이야.
   - ✅ 확인: 마이리얼트립 환불완료 줄은 금액이 마이너스라 그대로 더하면 돼.
5. ~~이심 카테고리~~ **완료** (PR: hayeon/esim-latest-month): 마이리얼트립 상품명이 나라 이름 하나뿐인 줄("중국", "일본" 등, `ESIM_PLAIN_NAMES`)을 '이심'으로 분류해. 플랫폼 상품에도 유형(`classifyProductType`)을 붙여서 매출 탭·인사이트 표에 배지로 보여줘.
6. **링크프라이스 추가**: PLATFORMS에 추가하되 하위 머천트(아고다, 트립닷컴, 호텔스닷컴, 익스피디아, 에어알로)별로 나눠서 집계해. 확정/미확정(취소) 구분도 필요해. 하연이가 엑셀 샘플을 주면 컬럼 매핑부터 설계해.
7. **UX**
   - ✅ 완료 (PR: hayeon/esim-latest-month): 개요 탭 기본 월 = 데이터가 있는 가장 최근 달 (`latestDataMonth`). 매출·블로그 탭의 월 선택은 '업로드할 달'이라 이번 달 그대로 둬.
   - ✅ 완료 (PR: hayeon/chart-gaps): 추이 그래프에서 데이터가 없는 달은 0 대신 빈칸(null). 실선은 끊기고 흐린 점선으로 이어줘. 선으로 안 이어지는 달(시작·끝·외톨이)은 동그라미로 표시하고, 툴팁은 "데이터 없음". 매출 탭 6개월 비교 막대그래프도 같은 규칙이야. (`monthValueOrNull`, `edgeDot`)
8. **나중에 (마케팅 기능)**
   - 글별 수익: 포스팅마다 추적값(서브 ID 등)을 달아 정산 엑셀과 매칭 → 조회수 1,000당 수익. 마이리얼트립은 정산 엑셀에 **`마케팅 링크 ID`** 열이 있어서 이걸로 글과 이을 수 있어.
   - 커미션 상향 기간을 매출 그래프 위에 표시
   - 상금·원고료를 커미션과 따로 집계
   - 조회수는 높은데 수익이 낮은 글, 시즌이 다가오는 글을 '다시 손볼 글'로 추천
   - 컨설팅 프롬프트를 '여행지 추천'에서 '다음 달 포스팅 리스트(플랫폼·코드 포함)'로 개선

## 참고 데이터
- 매출·조회수 같은 실제 숫자는 `CLAUDE.local.md`에 있어.
