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

## 배포 상태 (2026-11-06 기준, 확인 필요)
- 컨설팅 502 수정(PR: hayeon/fix-consult)은 머지됐지만 **사이트에 반영되지 않았어.** 배포 파일 이름이 `assets/index-Dce6VQrB.js`로 그대로야.
- 가장 유력한 원인: Netlify가 GitHub와 연결되지 않은 수동(드래그 앤 드롭) 배포.
  → Netlify > Site configuration > Build & deploy에서 저장소 연결 여부를 확인하고, **Base directory를 `travelog-dashboard (1)`로** 설정해야 해.
- 하연이가 Netlify 접근 권한을 받는 중이야. 받으면 `npx netlify login` → `npx netlify link`로 연결해서 배포 로그를 직접 확인해줘.

## 할 일 (우선순위 순)
1. **배포 연결 해결** (위 내용) → 컨설팅 버튼이 실제로 동작하는지 확인. Netlify 환경변수에 `ANTHROPIC_API_KEY`가 있는지도 확인해.
2. **보안 (중요)**: 지금 `firestore.rules`는 `request.auth != null`만 검사하는데 로그인이 익명 자동 로그인이야. 그래서 **사이트 주소를 아는 누구나 데이터를 읽고 쓰고 지울 수 있어.**
   - Firebase Auth에서 Google 로그인을 켜고, 앱의 PatternGate를 Google 로그인으로 교체
   - 규칙을 허용 이메일 목록으로 제한 (예: `request.auth.token.email in ['leehayaon@gmail.com', '동생 이메일']` + `email_verified`)
   - 규칙은 Firebase 콘솔에 직접 게시해야 해. 하연이한테 단계별로 안내해줘.
3. **지역 분류 개선** (`DESTINATIONS`, `matchDestination`)
   - 지금은 매출 대부분이 '기타'로 빠져서 인사이트 '종합 TOP 10' 1위가 "기타 ₩26,628,227"로 나와.
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

## 참고 데이터 (2026년 10월)
- 총 매출 ₩11,425,863 (마이리얼트립 ₩7,927,118), 발행 122건, 조회수 131,456
- 글당 조회수: 요니 1,547 (59건) / 짠순 638 (63건)
- 요니 글당 조회수: 9월 2,779 → 10월 1,547 (약 −44%)
