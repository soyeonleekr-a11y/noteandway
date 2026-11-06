import React, { useState, useEffect, useMemo } from "react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, PieChart, Pie, Cell,
} from "recharts";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  Plane, TrendingUp, Upload, Sparkles, FileSpreadsheet, RefreshCw,
  AlertCircle, CheckCircle2, Link2, Loader2, Trash2, Compass, MapPin,
  Trophy, BarChart3, LogOut,
} from "lucide-react";
import { doc, getDoc, setDoc } from "firebase/firestore";
import {
  GoogleAuthProvider, signInWithPopup, signInWithRedirect, onAuthStateChanged, signOut,
} from "firebase/auth";
import { db, auth } from "./firebase.js";

const PLATFORMS = ["여기어때", "마이리얼트립", "클룩"];
const PLATFORM_COLOR = { "여기어때": "#E8A33D", "마이리얼트립": "#5B8DEF", "클룩": "#EF5B5B" };
const TABS = [
  { id: "overview", label: "개요", gate: "GATE 1", icon: Plane },
  { id: "revenue", label: "플랫폼 매출", gate: "GATE 2", icon: TrendingUp },
  { id: "posts", label: "블로그 운영", gate: "GATE 3", icon: FileSpreadsheet },
  { id: "insight", label: "인사이트", gate: "GATE 4", icon: BarChart3 },
  { id: "consulting", label: "컨설팅", gate: "GATE 5", icon: Sparkles },
];

// 글 제목/상품명에서 나라·지역을 추정하기 위한 키워드 사전
// - D(이름, 국내/해외, [같은 곳으로 묶을 다른 표기·명소], 나라 단위 여부)
// - 나라 단위(일본, 미국…)는 같은 텍스트에 도시 이름이 없을 때만 쓰여요. "일본 오사카" → 오사카
// - 도시끼리는 텍스트에서 먼저 나오는 쪽이 이겨요. "보코 서울 명동 … 싱가포르 뷔페" → 서울
const D = (name, region, keywords = [], country = false) => ({ name, region, keywords: [name, ...keywords], country });
const DESTINATIONS = [
  // 일본
  D("도쿄", "해외"), D("오사카", "해외", ["USJ", "유니버설 스튜디오 재팬"]), D("교토", "해외"),
  D("후쿠오카", "해외"), D("삿포로", "해외"), D("오키나와", "해외"), D("나고야", "해외"),
  D("다카마쓰", "해외", ["다카마츠", "타카마츠", "타카마쓰"]), D("나오시마", "해외"),
  D("마쓰야마", "해외", ["마츠야마"]),
  D("일본", "해외", [], true),
  // 동남아
  D("다낭", "해외"), D("하노이", "해외"), D("호치민", "해외"), D("푸꾸옥", "해외"), D("나트랑", "해외"),
  D("베트남", "해외", [], true),
  D("방콕", "해외"), D("푸켓", "해외"), D("치앙마이", "해외"), D("파타야", "해외"), D("태국", "해외", [], true),
  D("세부", "해외"), D("보라카이", "해외"), D("마닐라", "해외"), D("필리핀", "해외", [], true),
  D("싱가포르", "해외"),
  D("발리", "해외"), D("인도네시아", "해외", [], true),
  D("라오스", "해외", [], true), D("캄보디아", "해외", [], true), D("미얀마", "해외", [], true),
  // 중화권
  D("타이베이", "해외"), D("대만", "해외", [], true),
  D("홍콩", "해외"), D("마카오", "해외"),
  D("상하이", "해외"), D("베이징", "해외"), D("중국", "해외", [], true),
  // 휴양지
  D("괌", "해외"), D("사이판", "해외"), D("몰디브", "해외"), D("하와이", "해외"),
  // 유럽
  D("파리", "해외"), D("프랑스", "해외", [], true),
  D("로마", "해외"), D("밀라노", "해외"), D("이탈리아", "해외", [], true),
  D("바르셀로나", "해외"), D("스페인", "해외", [], true),
  D("스위스", "해외", [], true), D("독일", "해외", [], true),
  D("런던", "해외"), D("영국", "해외", [], true),
  D("프라하", "해외"), D("체코", "해외", [], true),
  D("크로아티아", "해외", [], true), D("산토리니", "해외"), D("그리스", "해외", [], true),
  D("튀르키예", "해외", ["터키"], true),
  // 미주·대양주·기타
  D("뉴욕", "해외"),
  D("LA", "해외", ["로스앤젤레스", "로스엔젤레스", "엘에이"]),
  D("라스베가스", "해외", ["라스베이거스"]),
  D("그랜드캐년", "해외", ["그랜드캐니언", "그랜드 캐년", "그랜드 캐니언"]),
  D("샌프란시스코", "해외"), D("요세미티", "해외"),
  D("미국", "해외", [], true), D("캐나다", "해외", [], true),
  D("시드니", "해외"), D("호주", "해외", [], true), D("뉴질랜드", "해외", [], true),
  D("몽골", "해외", [], true), D("두바이", "해외"),
  // 국내
  D("서울", "국내", ["북한산", "명동", "홍대", "강남", "잠실", "성수", "여의도"]),
  D("인천", "국내", ["인스파이어", "영종도", "송도"]),
  D("제주", "국내", ["서귀포", "애월"]),
  D("부산", "국내", ["해운대", "광안리"]),
  D("강릉", "국내"), D("속초", "국내"), D("양양", "국내"), D("설악산", "국내"),
  D("평창", "국내"), D("춘천", "국내"), D("가평", "국내"),
  D("제천", "국내"), D("단양", "국내"), D("청송", "국내"), D("안동", "국내"),
  D("경주", "국내"), D("포항", "국내"), D("대구", "국내"), D("울산", "국내"),
  D("여수", "국내"), D("순천", "국내"), D("담양", "국내"), D("보성", "국내"), D("목포", "국내"),
  D("전주", "국내"), D("군산", "국내"), D("대전", "국내"), D("태안", "국내"), D("수원", "국내"),
  D("남해", "국내"), D("통영", "국내"), D("거제", "국내"),
  D("강원", "국내", [], true),
];
// 지명이 다른 단어 속에 들어 있어서 잘못 잡히는 표현 ("아로마 마사지" → 로마, "제주항공 오사카" → 제주)
const MATCH_IGNORE = ["제주항공", "아로마", "세부사항", "세부 사항", "상세"];
// 상품 유형 사전: [화면에 보일 유형 이름, ...찾을 키워드]. 위에 있는 유형부터 먼저 찾아요.
const PRODUCT_TYPES = [
  ["이심", "이심", "esim", "e-sim"], ["유심", "유심", "usim"], ["와이파이", "와이파이", "포켓와이파이"],
  ["항공권", "항공권"], ["풀빌라", "풀빌라"], ["펜션", "펜션"], ["호텔", "호텔"], ["숙소", "숙소"],
  ["렌터카", "렌터카", "렌트카"], ["크루즈", "크루즈"], ["골프", "골프"], ["스파", "스파"],
  ["입장권", "입장권"], ["티켓", "티켓"], ["액티비티", "액티비티"], ["투어", "투어"],
  ["여행자보험", "여행자보험"],
];
// 마이리얼트립 정산 엑셀에서 상품명이 나라 이름 하나뿐인 줄("중국", "일본")은 마이심 이심 상품이에요.
// 나라 단위 지명 + 이심을 따로 파는 도시형 국가(홍콩·마카오·싱가포르)를 기준으로 봐요.
const ESIM_PLAIN_NAMES = new Set([
  ...DESTINATIONS.filter((d) => d.country && d.region === "해외").flatMap((d) => d.keywords),
  "홍콩", "마카오", "싱가포르",
]);
// 영문 키워드(LA, USJ)는 "PLAZA"처럼 다른 단어 속에 들어간 경우를 빼고 찾아요.
function findKeyword(text, kw) {
  if (/^[A-Za-z]+$/.test(kw)) {
    const m = new RegExp(`(^|[^A-Za-z])${kw}(?![A-Za-z])`).exec(text);
    return m ? m.index + m[1].length : -1;
  }
  return text.indexOf(kw);
}
function matchDestination(text) {
  if (!text) return null;
  let src = String(text);
  // 지울 표현은 같은 길이의 공백으로 바꿔서, 다른 지명의 위치(먼저 나온 순서)는 그대로 둬요.
  MATCH_IGNORE.forEach((w) => { src = src.split(w).join(" ".repeat(w.length)); });
  let best = null;
  DESTINATIONS.forEach((d) => {
    d.keywords.forEach((kw) => {
      const idx = findKeyword(src, kw);
      if (idx < 0) return;
      const better = !best
        || (best.d.country && !d.country)
        || (best.d.country === d.country && (idx < best.idx || (idx === best.idx && kw.length > best.len)));
      if (better) best = { d, idx, len: kw.length };
    });
  });
  return best ? { name: best.d.name, region: best.d.region } : null;
}
// 수동 태그는 Firestore "region-tags"에 { 태그키: { region, destination } } 형태로 저장돼요.
// 상품 태그 키는 '핵심 이름'(productGroupKey) 기준이라, 문구만 다른 이름들이 같은 태그를 같이 써요.
function productTagKey(platform, name) { return `product::${platform}::${productGroupKey(name)}`; }
// 상품 합치기 전에는 원래 이름 그대로 태그를 저장했어요. 그때 지정한 태그도 계속 읽을 수 있게 같이 찾아봐요.
function legacyProductTagKey(platform, name) { return `product::${platform}::${normalizeNameKey(name)}`; }
function productTagKeys(platform, group) {
  const keys = [productTagKey(platform, group.name), ...group.variants.map((v) => legacyProductTagKey(platform, v))];
  return [...new Set(keys)];
}
// 묶인 상품(group)의 지역: 새 키 → 예전 키 순서로 직접 지정한 태그를 찾고, 없으면 사전으로 추정해요.
function classifyProduct(platform, group, tags) {
  const keys = productTagKeys(platform, group);
  const hit = keys.find((k) => tags?.[k]);
  return classifyText(group.name, tags, hit || keys[0]);
}
function postTagKey(title) { return `post::${normalizeNameKey(title)}`; }
// 직접 지정한 지역이 있으면 그걸 쓰고, 없으면 사전으로 추정해요.
function classifyText(text, tags, tagKey) {
  const tag = tags?.[tagKey];
  if (tag) return { region: tag.region, destination: tag.destination, manual: true };
  const d = matchDestination(text);
  return d ? { region: d.region, destination: d.name, manual: false } : { region: "기타", destination: "기타", manual: false };
}
function matchProductType(text) {
  if (!text) return null;
  const lower = String(text).toLowerCase();
  const found = PRODUCT_TYPES.find(([, ...keywords]) => keywords.some((k) => lower.includes(k.toLowerCase())));
  return found ? found[0] : null;
}
// 플랫폼 매출 상품의 유형. 이름에 키워드가 없어도 플랫폼별 규칙으로 알아낼 수 있는 건 여기서 처리해요.
// category: 엑셀의 '상품군' 열(마이리얼트립). 키워드로 못 알아낸 상품은 상품군을 그대로 보여줘요.
function classifyProductType(platform, name, category) {
  if (platform === "마이리얼트립" && ESIM_PLAIN_NAMES.has(normalizeNameKey(name))) return "이심";
  return matchProductType(name) || (category ? String(category).trim() : null);
}
// 상품명 옆에 유형 배지를 붙여서 보여줘요. 여러 이름이 하나로 묶였으면 '묶음 N' 배지도 붙여요.
function ProductName({ name, type, variants }) {
  const merged = variants && variants.length > 1;
  return (
    <span className="product-name">
      {name}
      {type && <span className="type-badge">{type}</span>}
      {merged && (
        <span className="group-badge" title={`하나로 묶은 이름 ${variants.length}개\n${variants.map((v) => `· ${v}`).join("\n")}`}>
          묶음 {variants.length}
        </span>
      )}
    </span>
  );
}
// 매출이나 블로그 데이터가 하나라도 있는 가장 최근 달. 하나도 없으면 이번 달이에요.
function latestDataMonth(platformData, postsData) {
  const now = currentMonth();
  const found = [platformData, postsData]
    .flatMap((byKey) => Object.values(byKey || {}))
    .flatMap((byMonth) => Object.keys(byMonth || {}))
    .filter((m) => /^\d{4}-\d{2}$/.test(m) && m <= now)
    .sort();
  return found.length ? found[found.length - 1] : now;
}
// 예전 방식(엑셀 한 줄 = 상품 하나, revenue 필드만 존재)으로 저장된 데이터도
// 새 형태({name, qty, unitRevenue, totalRevenue})로 안전하게 변환해서 읽어요.
function normalizeProduct(prod) {
  const totalRevenue = prod.totalRevenue != null ? prod.totalRevenue : (prod.revenue || 0);
  const qty = prod.qty != null ? prod.qty : 1;
  const unitRevenue = prod.unitRevenue != null ? prod.unitRevenue : (qty ? totalRevenue / qty : totalRevenue);
  // salesAmount: 판매금액. 판매금액 열 없이 올린 예전 데이터는 null 이에요.
  const salesAmount = typeof prod.salesAmount === "number" ? prod.salesAmount : null;
  return { name: prod.name, qty, unitRevenue, totalRevenue, category: prod.category || null, salesAmount };
}

// ───────── 상품 합치기 ─────────
// 비교용 '핵심 이름': 대괄호 등 괄호 속 홍보 문구, ★…★ 문구, 이모지·기호, 띄어쓰기를 빼요.
// 예) "[단독특가] 상하이 우전 수향마을 야경투어 ★한국어 가이드★" 와 "상하이 우전수향마을 야경 투어" → 같은 상품
// 소괄호 ( )는 "(1인)", "(왕복)"처럼 상품을 구분하는 정보일 때가 많아서 남겨둬요.
function productGroupKey(name) {
  const core = String(name || "").normalize("NFKC")
    .replace(/\[[^\]]*\]|【[^】]*】|「[^」]*」|『[^』]*』|〔[^〕]*〕|〈[^〉]*〉|《[^》]*》/g, " ")
    .replace(/[★☆][^★☆]*[★☆]/g, " ")
    .replace(/\p{Extended_Pictographic}|[★☆♥♡❤♪♬※●○◆◇■□▶►▷▲△▼▽✔✓·•|~～!！]/gu, " ")
    .replace(/\s+/g, "")
    .toLowerCase();
  return core || normalizeNameKey(name).toLowerCase(); // 전부 지워지면(이름이 홍보 문구뿐이면) 원래 이름으로 비교해요
}
// 같은 핵심 이름인 상품을 한 줄로 합쳐요. 대표 이름은 수익이 가장 큰 이름이고, 묶인 이름들은 variants 에 남겨요.
// 저장된 원본 데이터는 그대로 두고 화면에서만 합쳐서, 잘못 묶이면 productGroupKey 만 고치면 돼요.
function groupProducts(products) {
  const map = {};
  (products || []).forEach((raw) => {
    const p = normalizeProduct(raw);
    const key = productGroupKey(p.name);
    if (!map[key]) map[key] = { qty: 0, totalRevenue: 0, category: null, byName: {}, salesAmount: null };
    const g = map[key];
    g.qty += p.qty;
    g.totalRevenue += p.totalRevenue;
    if (p.salesAmount != null) g.salesAmount = (g.salesAmount || 0) + p.salesAmount;
    if (p.category && !g.category) g.category = p.category;
    const nk = normalizeNameKey(p.name);
    if (!g.byName[nk]) g.byName[nk] = { name: p.name, totalRevenue: 0, qty: 0 };
    g.byName[nk].totalRevenue += p.totalRevenue;
    g.byName[nk].qty += p.qty;
  });
  return Object.values(map).map((g) => {
    const names = Object.values(g.byName).sort((a, b) => b.totalRevenue - a.totalRevenue || b.qty - a.qty);
    return {
      name: names[0].name,
      variants: names.map((n) => n.name),
      qty: g.qty,
      totalRevenue: g.totalRevenue,
      unitRevenue: g.qty ? g.totalRevenue / g.qty : g.totalRevenue,
      category: g.category,
      salesAmount: g.salesAmount,
      rate: commissionRate(g.totalRevenue, g.salesAmount),
    };
  });
}

// ───────── 플랫폼별 엑셀 열 이름 ─────────
// 열 이름을 정확히 아는 플랫폼은 여기에 적어두고, 그 이름(띄어쓰기 무시)으로만 찾아요.
// 비슷한 단어로 찾으면 '판매 금액'(고객 결제액)을 수익으로 잘못 읽을 수 있어서예요.
// 여기에 없는 플랫폼은 예전처럼 '상품'·'수익/매출/금액'·'수량' 단어가 든 첫 열을 써요.
// 일부 열만 아는 플랫폼은 아는 열만 적어요. 안 적은 열(name/qty)은 예전처럼 단어로 찾아요.
const PLATFORM_COLUMNS = {
  // sales: 고객이 결제한 판매금액. 수수료율(정산금 ÷ 판매금액)을 계산할 때만 써요.
  "마이리얼트립": { name: "상품명", revenue: "정산 예상 금액", qty: "판매 수량", sales: "판매 금액", category: "상품군" },
  // 수익은 '정산금액', 판매금액은 '총 실적 인정 금액' (하연이 확인)
  "여기어때": { revenue: "정산금액", sales: "총 실적 인정 금액" },
};
const COLUMN_LABELS = { name: "상품명", revenue: "수익", qty: "수량", sales: "판매금액", category: "상품군" };

// ───────── 수수료율 ─────────
// 수수료율 = 정산금(수익) ÷ 판매금액 × 100. 판매금액이 없거나 0이면 계산하지 않아요(null).
function commissionRate(settle, sales) {
  return sales ? (settle / sales) * 100 : null;
}
function formatRate(rate) {
  return rate == null ? "-" : `${rate.toFixed(1)}%`;
}
// 묶인 상품 목록(groupProducts 결과)을 분류(투어·호텔·티켓…)별로 모아 판매금액·정산금·수수료율을 계산해요.
// 판매금액 열 없이 올라간 상품은 빼요. 분류는 상품 유형 배지와 같은 규칙이고, 못 정하면 '기타'예요.
function rateByCategory(platform, groups) {
  const map = {};
  groups.forEach((g) => {
    if (g.salesAmount == null) return;
    const category = classifyProductType(platform, g.name, g.category) || "기타";
    if (!map[category]) map[category] = { category, sales: 0, settle: 0, products: 0 };
    map[category].sales += g.salesAmount;
    map[category].settle += g.totalRevenue;
    map[category].products += 1;
  });
  return Object.values(map)
    .map((r) => ({ ...r, rate: commissionRate(r.settle, r.sales) }))
    .sort((a, b) => b.sales - a.sales);
}
// 판매금액이 있는 상품만 합친 전체 요약. 판매금액이 하나도 없으면 null.
function rateSummary(groups) {
  const withSales = groups.filter((g) => g.salesAmount != null);
  if (!withSales.length) return null;
  const sales = withSales.reduce((s, g) => s + g.salesAmount, 0);
  const settle = withSales.reduce((s, g) => s + g.totalRevenue, 0);
  return { sales, settle, rate: commissionRate(settle, sales), products: withSales.length, missing: groups.length - withSales.length };
}
function sameHeader(a, b) {
  return String(a ?? "").replace(/\s+/g, "") === String(b ?? "").replace(/\s+/g, "");
}
function exactHeader(row, header) {
  const want = String(header).replace(/\s+/g, "");
  return Object.keys(row || {}).find((k) => String(k).replace(/\s+/g, "") === want) || null;
}
// 상품명이 눈으로는 같아 보여도 공백/특수문자 차이로 다른 문자열일 수 있어서,
// 합치는 기준(키)으로 쓸 때는 이 함수로 정규화해요. 화면에는 원래 이름을 그대로 보여줘요.
function normalizeNameKey(name) {
  return String(name || "").normalize("NFKC").replace(/\s+/g, " ").trim();
}

function pad2(n) { return String(n).padStart(2, "0"); }
function currentMonth() { const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`; }
function monthList(count = 12) {
  const arr = []; const d = new Date();
  for (let i = 0; i < count; i++) {
    arr.push(`${d.getFullYear()}-${pad2(d.getMonth() + 1)}`);
    d.setMonth(d.getMonth() - 1);
  }
  return arr;
}
function formatMonthLabel(m) { if (!m) return "-"; const [y, mo] = m.split("-"); return `${y}년 ${parseInt(mo, 10)}월`; }
function formatWon(n) { const v = Math.round(Number(n) || 0); return "₩" + v.toLocaleString("ko-KR"); }
function parseNumber(v) {
  if (typeof v === "number") return v;
  if (v === null || v === undefined) return 0;
  const s = String(v).replace(/[^0-9.\-]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function findKey(row, keywords) {
  const keys = Object.keys(row || {});
  return keys.find((k) => keywords.some((kw) => k.includes(kw)));
}
function normalizeMonth(raw) {
  if (!raw) return null;
  const s = String(raw);
  const m = s.match(/(\d{4})[.\-\/년\s]*\s*(\d{1,2})/);
  if (!m) return null;
  return `${m[1]}-${pad2(parseInt(m[2], 10))}`;
}
// 엑셀의 게시 날짜 셀(Date 객체 또는 "2026-07-15" / "2026.7.15" 등 문자열)을 'YYYY-MM-DD'로 통일해요.
function parseDateValue(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (raw instanceof Date && !isNaN(raw)) {
    return `${raw.getFullYear()}-${pad2(raw.getMonth() + 1)}-${pad2(raw.getDate())}`;
  }
  const s = String(raw).trim();
  const m = s.match(/(\d{4})[.\-\/년]\s*(\d{1,2})[.\-\/월]\s*(\d{1,2})/);
  if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  const d = new Date(s);
  return isNaN(d) ? null : `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
// 게시일로부터 오늘까지 경과일(게시 당일도 1일로 계산)과 일평균 조회수를 계산해요.
// 갓 올라온 글은 조회수가 낮아도 경과일이 짧으니 일평균으로 보면 공정하게 비교돼요.
function daysSincePublish(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d)) return null;
  const diff = Math.floor((Date.now() - d.getTime()) / 86400000) + 1;
  return diff > 0 ? diff : 1;
}
function dailyAvgViews(views, dateStr) {
  const days = daysSincePublish(dateStr);
  if (!days) return null;
  return views / days;
}

// Firestore의 "dashboardData" 컬렉션에 문서 하나당 값 하나씩 저장 (JSON 문자열 변환 불필요)
async function safeGet(key) {
  try {
    const snap = await getDoc(doc(db, "dashboardData", key));
    return snap.exists() ? snap.data().value : null;
  } catch (e) { console.error(`[firestore] get(${key}) 실패`, e); return null; }
}
async function safeSet(key, value) {
  try {
    await setDoc(doc(db, "dashboardData", key), { value, updatedAt: new Date().toISOString() });
    return true;
  } catch (e) { console.error(`[firestore] set(${key}) 실패`, e); return false; }
}

function SplitNumber({ value, size = "lg" }) {
  return (
    <span className={`split-number ${size === "lg" ? "split-lg" : "split-sm"}`}>{value}</span>
  );
}

function EmptyState({ icon: Icon, title, desc }) {
  return (
    <div className="empty-state">
      <Icon size={28} strokeWidth={1.5} />
      <div className="empty-title">{title}</div>
      <div className="empty-desc">{desc}</div>
    </div>
  );
}

// Firebase 로그인 오류 코드를 쉬운 안내 문구로 바꿔요.
const AUTH_ERROR_TEXT = {
  "auth/unauthorized-domain": "이 주소는 아직 로그인 허용 도메인에 등록되지 않았어요. (Firebase > Authentication > 설정 > 승인된 도메인)",
  "auth/operation-not-allowed": "Google 로그인이 아직 켜져 있지 않아요. (Firebase > Authentication > 로그인 방법)",
  "auth/network-request-failed": "네트워크 연결을 확인하고 다시 시도해주세요.",
};

function GoogleMark() {
  return (
    <svg width="16" height="16" viewBox="0 0 48 48" aria-hidden="true">
      <path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3C33.7 32.7 29.2 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.4-.4-3.5z" />
      <path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.7 15.1 19 12 24 12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z" />
      <path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2C29.2 35.1 26.7 36 24 36c-5.2 0-9.6-3.3-11.3-8l-6.5 5C9.5 39.6 16.2 44 24 44z" />
      <path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.2-2.2 4.2-4.1 5.6l6.2 5.2C37 39.2 44 34 44 24c0-1.3-.1-2.4-.4-3.5z" />
    </svg>
  );
}

function LoginScreen() {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function handleLogin() {
    setBusy(true); setError("");
    const provider = new GoogleAuthProvider();
    provider.setCustomParameters({ prompt: "select_account" });
    try {
      await signInWithPopup(auth, provider);
    } catch (e) {
      if (e.code === "auth/popup-blocked") {
        await signInWithRedirect(auth, provider); // 팝업이 막힌 브라우저는 페이지 이동 방식으로
        return;
      }
      if (e.code !== "auth/popup-closed-by-user" && e.code !== "auth/cancelled-popup-request") {
        console.error("[firebase auth] Google 로그인 실패", e);
        setError(AUTH_ERROR_TEXT[e.code] || "로그인에 실패했어요. 잠시 후 다시 시도해주세요.");
      }
    }
    setBusy(false);
  }

  return (
    <div className="app-shell login-shell">
      <Style />
      <div className="login-card">
        <div className="brand-mark"><Plane size={18} /></div>
        <div className="login-title">TRAVELOG DASHBOARD</div>
        <div className="login-sub">등록된 Google 계정으로 로그인해주세요.</div>
        <button className="google-btn" onClick={handleLogin} disabled={busy} type="button">
          {busy ? <Loader2 size={16} className="spin" /> : <GoogleMark />}
          <span>Google 계정으로 로그인</span>
        </button>
        {error && <div className="error-box login-error"><AlertCircle size={14} /> {error}</div>}
      </div>
    </div>
  );
}

function DeniedScreen({ email, onSignOut }) {
  return (
    <div className="app-shell login-shell">
      <Style />
      <div className="login-card">
        <div className="brand-mark denied"><AlertCircle size={18} /></div>
        <div className="login-title">접근 권한이 없는 계정이에요</div>
        <div className="login-sub">
          <b>{email}</b><br />이 계정은 대시보드에 등록되어 있지 않아요.<br />다른 계정으로 로그인하거나 관리자에게 등록을 요청해주세요.
        </div>
        <button className="ghost-btn" onClick={onSignOut} type="button"><LogOut size={13} /> 다른 계정으로 로그인</button>
      </div>
    </div>
  );
}

export default function App() {
  // user: undefined = 로그인 상태 확인 중, null = 로그아웃, 객체 = 로그인됨
  const [user, setUser] = useState(undefined);
  // access: checking = 허용 목록 확인 중, allowed = 통과, denied = 등록 안 된 계정
  const [access, setAccess] = useState("checking");
  const [loaded, setLoaded] = useState(false);
  const [activeTab, setActiveTab] = useState("overview");
  const [platformData, setPlatformData] = useState({});
  const [postsData, setPostsData] = useState({});
  const [postRecords, setPostRecords] = useState({});
  const [blogNames, setBlogNames] = useState(["해외여행 블로그", "국내여행 블로그"]);
  const [consultingHistory, setConsultingHistory] = useState([]);
  const [regionTags, setRegionTags] = useState({});
  const [toast, setToast] = useState(null);

  // Firebase가 로그인 상태를 기억해서, 새로고침해도 다시 로그인할 필요가 없어요.
  useEffect(() => {
    return onAuthStateChanged(auth, async (u) => {
      // 예전 강아지 잠금 시절의 익명 로그인이 남아 있으면 정리해요.
      if (u && u.isAnonymous) { await signOut(auth); return; }
      setUser(u);
      if (!u) { setAccess("checking"); return; }
      setAccess("checking");
      // allowedUsers/{이메일} 문서가 있는지 확인해요. (실제 차단은 Firestore 규칙이 해요)
      try {
        const snap = await getDoc(doc(db, "allowedUsers", u.email));
        setAccess(snap.exists() ? "allowed" : "denied");
      } catch (e) {
        console.error("[firestore] 허용 목록 확인 실패", e);
        setAccess("denied");
      }
    });
  }, []);

  async function handleSignOut() {
    await signOut(auth);
    // 다른 계정으로 다시 로그인할 때 이전 데이터가 잠깐 보이지 않게 비워요.
    setLoaded(false);
    setPlatformData({}); setPostsData({}); setPostRecords({}); setConsultingHistory([]); setRegionTags({});
  }

  useEffect(() => {
    if (access !== "allowed") return;
    (async () => {
      const [pd, bd, pr, bn, ch, rt] = await Promise.all([
        safeGet("platform-data"), safeGet("posts-data"), safeGet("post-records"),
        safeGet("blog-names"), safeGet("consulting-history"), safeGet("region-tags"),
      ]);
      if (rt) setRegionTags(rt);
      if (pd) setPlatformData(pd);
      if (bd) setPostsData(bd);
      if (pr) setPostRecords(pr);
      if (bn) setBlogNames(bn);
      if (ch) setConsultingHistory(ch);
      // 첫 화면은 이번 달이 아니라 데이터가 있는 가장 최근 달을 보여줘요. (이번 달은 아직 정산 전이라 ₩0으로 보이니까)
      setSelectedMonth(latestDataMonth(pd, bd));
      setLoaded(true);
    })();
  }, [access]);

  function showToast(msg, kind = "ok") {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 3200);
  }

  async function updatePlatformData(next) {
    setPlatformData(next);
    const ok = await safeSet("platform-data", next);
    if (!ok) showToast("저장에 실패했어요. 다시 시도해주세요.", "err");
  }
  async function updatePostsData(next) {
    setPostsData(next);
    const ok = await safeSet("posts-data", next);
    if (!ok) showToast("저장에 실패했어요. 다시 시도해주세요.", "err");
  }
  async function updatePostRecords(next) {
    setPostRecords(next);
    const ok = await safeSet("post-records", next);
    if (!ok) showToast("저장에 실패했어요. 다시 시도해주세요.", "err");
  }
  async function updateBlogNames(next) {
    setBlogNames(next);
    await safeSet("blog-names", next);
  }
  // tag가 null이면 직접 지정을 해제하고 사전 자동 분류로 돌아가요.
  // cleanupKeys: 같은 상품의 예전 방식 태그 키들. 새로 저장하거나 해제할 때 같이 지워서 남지 않게 해요.
  async function updateRegionTag(key, tag, cleanupKeys = []) {
    const next = { ...regionTags };
    cleanupKeys.forEach((k) => { if (k !== key) delete next[k]; });
    if (tag) next[key] = tag; else delete next[key];
    setRegionTags(next);
    const ok = await safeSet("region-tags", next);
    if (!ok) showToast("저장에 실패했어요. 다시 시도해주세요.", "err");
    else showToast(tag ? `'${tag.destination}'(으)로 지정했어요` : "직접 지정을 해제했어요");
  }

  const months = useMemo(() => monthList(12), []);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth());

  const monthTotal = (m) => PLATFORMS.reduce((s, p) => s + (platformData[p]?.[m]?.total || 0), 0);
  const monthPosts = (m) => blogNames.reduce((s, b) => s + (postsData[b]?.[m]?.posts || 0), 0);
  const monthViews = (m) => blogNames.reduce((s, b) => s + (postsData[b]?.[m]?.views || 0), 0);

  // 데이터가 없는 달은 0이 아니라 null 로 둬요. 그래야 그래프가 ₩0으로 뚝 떨어지지 않고 그 달을 비워둬요.
  const trendData = useMemo(() => {
    return [...months].reverse().map((m) => {
      const row = { month: formatMonthLabel(m).replace("년 ", ".").replace("월", "") };
      PLATFORMS.forEach((p) => { row[p] = monthValueOrNull(platformData[p]?.[m]); });
      row["총합"] = monthTotal(m);
      return row;
    });
  }, [months, platformData, blogNames, postsData]);

  const pieData = PLATFORMS.map((p) => ({ name: p, value: platformData[p]?.[selectedMonth]?.total || 0 }));

  if (user === null) return <LoginScreen />;
  if (user && access === "denied") return <DeniedScreen email={user.email} onSignOut={handleSignOut} />;

  if (!user || access === "checking" || !loaded) {
    return (
      <div className="app-shell loading-shell">
        <Loader2 className="spin" size={22} /> <span style={{ marginLeft: 10 }}>불러오는 중...</span>
        <Style />
      </div>
    );
  }

  return (
    <div className="app-shell">
      <Style />
      <header className="header">
        <div className="brand">
          <div className="brand-mark"><Plane size={18} /></div>
          <div>
            <div className="brand-title">TRAVELOG DASHBOARD</div>
            <div className="brand-sub">블로그 운영 · 어필리에이트 매출 관제탑</div>
          </div>
          <button className="logout-btn" onClick={handleSignOut} title={`로그아웃 (${user.email})`}><LogOut size={14} /></button>
        </div>
        <div className="boarding-strip">
          <span>DEPARTURE</span><span className="dot">•</span>
          <span>{blogNames.join(" / ")}</span>
        </div>
      </header>

      <nav className="tabs">
        {TABS.map((t) => {
          const Icon = t.icon;
          const activeCls = activeTab === t.id ? "tab active" : "tab";
          return (
            <button key={t.id} className={activeCls} onClick={() => setActiveTab(t.id)}>
              <span className="tab-gate">{t.gate}</span>
              <Icon size={16} />
              <span>{t.label}</span>
            </button>
          );
        })}
      </nav>

      <main className="main">
        {activeTab === "overview" && (
          <OverviewTab
            months={months} selectedMonth={selectedMonth} setSelectedMonth={setSelectedMonth}
            platformData={platformData} postsData={postsData} blogNames={blogNames}
            monthTotal={monthTotal} monthPosts={monthPosts} monthViews={monthViews}
            trendData={trendData} pieData={pieData}
          />
        )}
        {activeTab === "revenue" && (
          <RevenueTab
            months={months} platformData={platformData} updatePlatformData={updatePlatformData}
            showToast={showToast}
          />
        )}
        {activeTab === "posts" && (
          <PostsTab
            months={months} postsData={postsData} updatePostsData={updatePostsData}
            postRecords={postRecords} updatePostRecords={updatePostRecords}
            blogNames={blogNames} updateBlogNames={updateBlogNames} showToast={showToast}
            regionTags={regionTags}
          />
        )}
        {activeTab === "consulting" && (
          <ConsultingTab
            months={months} platformData={platformData} postsData={postsData} blogNames={blogNames}
            consultingHistory={consultingHistory} setConsultingHistory={setConsultingHistory}
            showToast={showToast}
          />
        )}
        {activeTab === "insight" && (
          <InsightTab
            months={months} platformData={platformData} postRecords={postRecords} blogNames={blogNames}
            regionTags={regionTags} updateRegionTag={updateRegionTag}
          />
        )}
      </main>

      {toast && (
        <div className={`toast ${toast.kind === "err" ? "toast-err" : "toast-ok"}`}>
          {toast.kind === "err" ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
          <span>{toast.msg}</span>
        </div>
      )}
    </div>
  );
}

// ───────── 추이 그래프 도우미 ─────────
// 그 달 데이터가 아예 없으면 null(빈칸), 있으면 총액. (총액이 0원인 달은 0으로 그대로 보여줘요)
function monthValueOrNull(monthData) {
  return monthData ? monthData.total || 0 : null;
}
function formatWonOrEmpty(v) {
  return v == null ? "데이터 없음" : formatWon(v);
}
// 앞뒤 달 중 하나라도 비어서 실선으로 이어지지 않는 점(시작·끝·외톨이 달)에만 동그라미를 그려요.
function edgeDot(data, key, color) {
  return function renderEdgeDot(props) {
    const { cx, cy, index } = props;
    const v = data[index]?.[key];
    const linkedBoth = data[index - 1]?.[key] != null && data[index + 1]?.[key] != null;
    if (v == null || linkedBoth || cx == null || cy == null) return <g key={`${key}-dot-${index}`} />;
    return <circle key={`${key}-dot-${index}`} cx={cx} cy={cy} r={3} fill="#1C324F" stroke={color} strokeWidth={2} />;
  };
}
// 첫 데이터와 마지막 데이터 사이에 빈 달이 끼어 있는 플랫폼이 하나라도 있는지 (점선 안내문용)
function hasTrendGap(data) {
  return PLATFORMS.some((p) => {
    const idx = data.map((row, i) => (row[p] != null ? i : -1)).filter((i) => i >= 0);
    return idx.length > 1 && idx[idx.length - 1] - idx[0] + 1 > idx.length;
  });
}

function OverviewTab({ months, selectedMonth, setSelectedMonth, platformData, blogNames, postsData, monthTotal, monthPosts, monthViews, trendData, pieData }) {
  const total = monthTotal(selectedMonth);
  const posts = monthPosts(selectedMonth);
  const views = monthViews(selectedMonth);
  const hasAnyRevenue = PLATFORMS.some((p) => platformData[p]?.[selectedMonth]);

  return (
    <div className="tab-body">
      <div className="row-between">
        <h2 className="section-title"><Compass size={18} /> 이번 달 관제 현황</h2>
        <select className="month-select" value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)}>
          {/* 최근 데이터가 12개월보다 오래됐으면 목록에 없어서, 그 달도 선택지로 넣어줘요 */}
          {!months.includes(selectedMonth) && <option value={selectedMonth}>{formatMonthLabel(selectedMonth)}</option>}
          {months.map((m) => <option key={m} value={m}>{formatMonthLabel(m)}</option>)}
        </select>
      </div>

      <div className="stat-grid">
        <div className="stat-card stub">
          <div className="stat-label">총 매출</div>
          <SplitNumber value={formatWon(total)} />
        </div>
        <div className="stat-card stub">
          <div className="stat-label">발행 글 수</div>
          <SplitNumber value={`${posts.toLocaleString("ko-KR")}건`} size="sm" />
        </div>
        <div className="stat-card stub">
          <div className="stat-label">총 조회수</div>
          <SplitNumber value={views.toLocaleString("ko-KR")} size="sm" />
        </div>
      </div>

      <div className="grid-2">
        <div className="panel">
          <div className="panel-title">최근 12개월 매출 추이</div>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={trendData}>
              <CartesianGrid stroke="#2A3F5C" strokeDasharray="3 3" />
              <XAxis dataKey="month" tick={{ fill: "#93A2BC", fontSize: 11 }} />
              <YAxis tick={{ fill: "#93A2BC", fontSize: 11 }} tickFormatter={(v) => (v >= 10000 ? `${Math.round(v / 10000)}만` : v)} />
              <Tooltip contentStyle={{ background: "#16273F", border: "1px solid #2A3F5C", color: "#F4F1E8" }} formatter={formatWonOrEmpty} />
              <Legend wrapperStyle={{ fontSize: 12, color: "#93A2BC" }} />
              {/* 1) 아래층: 데이터가 없는 달을 건너뛰어 이어주는 흐린 점선 (범례·툴팁에는 안 나와요) */}
              {PLATFORMS.map((p) => (
                <Line
                  key={`${p}-bridge`} type="monotone" dataKey={p} stroke={PLATFORM_COLOR[p]} strokeWidth={1.5}
                  strokeDasharray="4 4" strokeOpacity={0.45} dot={false} activeDot={false}
                  connectNulls legendType="none" tooltipType="none" isAnimationActive={false}
                />
              ))}
              {/* 2) 위층: 데이터가 이어지는 구간만 실선. 선으로 안 이어지는 달은 동그라미로 표시해요 */}
              {PLATFORMS.map((p) => (
                <Line
                  key={p} type="monotone" dataKey={p} stroke={PLATFORM_COLOR[p]} strokeWidth={2}
                  dot={edgeDot(trendData, p, PLATFORM_COLOR[p])}
                />
              ))}
            </LineChart>
          </ResponsiveContainer>
          {hasTrendGap(trendData) && (
            <div className="hint">점선은 데이터가 없는 달을 건너뛰어 이은 구간이에요.</div>
          )}
        </div>

        <div className="panel">
          <div className="panel-title">{formatMonthLabel(selectedMonth)} 플랫폼별 비중</div>
          {hasAnyRevenue ? (
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie data={pieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={3}>
                  {pieData.map((entry, i) => <Cell key={i} fill={PLATFORM_COLOR[entry.name]} />)}
                </Pie>
                <Tooltip contentStyle={{ background: "#16273F", border: "1px solid #2A3F5C", color: "#F4F1E8" }} formatter={(v) => formatWon(v)} />
                <Legend wrapperStyle={{ fontSize: 12, color: "#93A2BC" }} />
              </PieChart>
            </ResponsiveContainer>
          ) : (
            <EmptyState icon={Upload} title="아직 매출 데이터가 없어요" desc="플랫폼 매출 탭에서 엑셀을 업로드해주세요." />
          )}
        </div>
      </div>

      <div className="panel">
        <div className="panel-title">블로그별 운영 현황 ({formatMonthLabel(selectedMonth)})</div>
        <div className="blog-cards">
          {blogNames.map((b) => {
            const d = postsData[b]?.[selectedMonth];
            return (
              <div key={b} className="blog-card">
                <div className="blog-card-name"><MapPin size={14} /> {b}</div>
                <div className="blog-card-row"><span>발행 글</span><b>{d ? `${d.posts}건` : "-"}</b></div>
                <div className="blog-card-row"><span>조회수</span><b>{d ? d.views.toLocaleString("ko-KR") : "-"}</b></div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// 판매금액 대비 정산금(수수료율) 요약: 위는 지표 카드 3개, 아래는 분류별 표 + 수수료율 크기를 보여주는 얇은 막대.
// 막대는 한 가지 색(골드)이고, 정확한 숫자는 옆 글자와 마우스 툴팁이 알려줘요.
function RatePanel({ summary, rows }) {
  const maxRate = Math.max(0.0001, ...rows.map((r) => Math.max(r.rate ?? 0, 0)));
  return (
    <div className="rate-panel">
      <div className="kpi-row">
        <div className="kpi">
          <div className="kpi-label">판매금액</div>
          <div className="kpi-value">{formatWon(summary.sales)}</div>
          <div className="kpi-sub">고객이 결제한 금액</div>
        </div>
        <div className="kpi">
          <div className="kpi-label">정산금 (수익)</div>
          <div className="kpi-value">{formatWon(summary.settle)}</div>
          <div className="kpi-sub">판매금액이 있는 상품 {summary.products}개 기준</div>
        </div>
        <div className="kpi kpi-accent">
          <div className="kpi-label">수수료율</div>
          <div className="kpi-value kpi-hero">{formatRate(summary.rate)}</div>
          <div className="kpi-sub">판매금액 대비 정산금</div>
        </div>
      </div>
      {rows.length > 0 && (
        <div className="table-scroll">
          <table className="table rate-table">
            <thead><tr><th>분류</th><th>상품</th><th>판매금액</th><th>정산금</th><th className="rate-col">수수료율</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.category}>
                  <td><span className="type-badge">{r.category}</span></td>
                  <td>{r.products}개</td>
                  <td className="num-cell">{formatWon(r.sales)}</td>
                  <td className="num-cell">{formatWon(r.settle)}</td>
                  <td className="rate-col">
                    <div className="rate-bar" title={`${r.category}: 판매금액 ${formatWon(r.sales)} 중 정산금 ${formatWon(r.settle)} (${formatRate(r.rate)})`}>
                      <span className="rate-bar-track">
                        <span className="rate-bar-fill" style={{ width: `${r.rate != null && r.rate > 0 ? Math.max(3, (r.rate / maxRate) * 100) : 0}%` }} />
                      </span>
                      <span className="rate-bar-value">{formatRate(r.rate)}</span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {summary.missing > 0 && <div className="hint">판매금액이 없는 상품 {summary.missing}개는 수수료율 계산에서 빠졌어요.</div>}
    </div>
  );
}

function RevenueTab({ months, platformData, updatePlatformData, showToast }) {
  const [platform, setPlatform] = useState(PLATFORMS[0]);
  const [month, setMonth] = useState(currentMonth());
  const [parsing, setParsing] = useState(false);
  const [sortField, setSortField] = useState("totalRevenue");
  const [sortDir, setSortDir] = useState("desc");
  const [showAll, setShowAll] = useState(false);
  const PAGE_SIZE = 30;

  const current = platformData[platform]?.[month];

  useEffect(() => { setShowAll(false); }, [platform, month]);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setParsing(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!rows.length) throw new Error("empty");
      const profile = PLATFORM_COLUMNS[platform] || {};
      // 열 이름을 지정해둔 열은 그 이름으로만 찾고(없으면 엉뚱한 열을 읽지 않게 멈춰요),
      // 지정 안 한 열은 예전처럼 단어가 들어간 첫 열을 써요.
      const nameKey = profile.name ? exactHeader(rows[0], profile.name) : findKey(rows[0], ["상품", "품목", "product", "Product"]);
      const revKey = profile.revenue ? exactHeader(rows[0], profile.revenue) : findKey(rows[0], ["수익", "매출", "금액", "revenue", "Revenue"]);
      const qtyKey = profile.qty ? exactHeader(rows[0], profile.qty) : findKey(rows[0], ["수량", "판매", "qty", "Qty"]);
      const categoryKey = profile.category ? exactHeader(rows[0], profile.category) : null;
      // 판매금액은 수수료율용이라 없어도 업로드는 돼요. (지정한 플랫폼만 읽어요)
      const salesKey = profile.sales ? exactHeader(rows[0], profile.sales) : null;
      const missing = [[nameKey, profile.name], [revKey, profile.revenue]].filter(([k, h]) => h && !k).map(([, h]) => `'${h}'`);
      if (missing.length) throw new Error(`${platform} 엑셀에서 ${missing.join(", ")} 열을 찾지 못했어요. 엑셀 첫 줄의 열 이름을 확인해주세요.`);
      if (!nameKey || !revKey) throw new Error("컬럼을 찾을 수 없어요 (상품명 / 수익금 열이 필요해요)");

      // 엑셀은 보통 한 줄 = 거래 1건이라 수량 열이 따로 없을 수 있어요.
      // 그래서 같은 상품명은 한 줄로 합치고, 수량 열이 없으면 등장 횟수를 판매수량으로 계산해요.
      // (공백/특수문자 차이로 다른 문자열처럼 보이는 경우까지 대비해 정규화한 키로 합쳐요)
      // 문구만 다른 이름끼리 묶는 건 화면에서 해요(groupProducts). 저장할 땐 원래 이름을 살려둬요.
      const grouped = {};
      rows.forEach((r) => {
        const name = String(r[nameKey]).trim();
        if (!name) return;
        const key = normalizeNameKey(name);
        const revenue = parseNumber(r[revKey]);
        const qty = qtyKey ? parseNumber(r[qtyKey]) || 1 : 1;
        if (!grouped[key]) grouped[key] = { name, qty: 0, totalRevenue: 0 };
        if (salesKey) grouped[key].salesAmount = (grouped[key].salesAmount || 0) + parseNumber(r[salesKey]);
        grouped[key].qty += qty;
        grouped[key].totalRevenue += revenue;
        const category = categoryKey ? String(r[categoryKey]).trim() : "";
        if (category && !grouped[key].category) grouped[key].category = category; // Firestore는 undefined를 못 저장해서, 있을 때만 넣어요
      });
      const products = Object.values(grouped).map((p) => ({
        ...p,
        unitRevenue: p.qty ? p.totalRevenue / p.qty : p.totalRevenue,
      }));
      if (!products.length) throw new Error("읽을 수 있는 상품 데이터가 없어요");
      const total = products.reduce((s, p) => s + p.totalRevenue, 0);

      // 나중에 "이 숫자 어느 열에서 왔지?"를 확인할 수 있게, 읽은 열 이름을 같이 저장해요.
      const source = { name: nameKey, revenue: revKey, qty: qtyKey || null, sales: salesKey || null, category: categoryKey || null, rows: rows.length, uploadedAt: new Date().toISOString() };
      const next = { ...platformData, [platform]: { ...(platformData[platform] || {}), [month]: { products, total, source } } };
      await updatePlatformData(next);
      const salesNote = profile.sales && !salesKey ? ` · '${profile.sales}' 열이 없어서 수수료율은 빠졌어요` : "";
      showToast(`${platform} · ${formatMonthLabel(month)} 매출 상품 ${products.length}종 반영 완료 (수익: '${revKey}' 열)${salesNote}`);
    } catch (err) {
      showToast(err.message || "엑셀 파일을 읽는 데 실패했어요", "err");
    } finally {
      setParsing(false);
      e.target.value = "";
    }
  }

  async function handleClear() {
    if (!current) return;
    const platObj = { ...(platformData[platform] || {}) };
    delete platObj[month];
    const next = { ...platformData, [platform]: platObj };
    await updatePlatformData(next);
    showToast("삭제했어요");
  }

  const compareData = [...months].slice(0, 6).reverse().map((m) => {
    const row = { month: formatMonthLabel(m).replace("년 ", ".").replace("월", "") };
    PLATFORMS.forEach((p) => { row[p] = monthValueOrNull(platformData[p]?.[m]); });
    return row;
  });

  const sortedProducts = useMemo(() => {
    if (!current) return [];
    // 문구만 다른 같은 상품(예: [특가]·★ 문구 차이)은 핵심 이름 기준으로 한 줄로 합쳐서 보여줘요.
    const arr = groupProducts(current.products);
    // 값이 없는 상품(예: 수수료율 '-')은 정렬 방향과 상관없이 항상 맨 뒤로 보내요.
    arr.sort((a, b) => {
      const va = a[sortField], vb = b[sortField];
      if (va == null && vb == null) return 0;
      if (va == null) return 1;
      if (vb == null) return -1;
      return sortDir === "desc" ? vb - va : va - vb;
    });
    return arr;
  }, [current, sortField, sortDir]);
  const mergedCount = sortedProducts.filter((p) => p.variants.length > 1).length;
  // 수수료율: 판매금액 열과 함께 올린 달만 계산돼요.
  const summary = useMemo(() => rateSummary(sortedProducts), [sortedProducts]);
  const categoryRates = useMemo(() => rateByCategory(platform, sortedProducts), [platform, sortedProducts]);
  const profile = PLATFORM_COLUMNS[platform];
  const profileFields = profile ? Object.entries(profile) : []; // [["revenue", "정산금액"], ...]
  // 수익 열을 지정한 플랫폼인데, 이 달이 그 열이 아닌 다른 열로 계산돼 있으면 다시 올려달라고 안내해요.
  // (기록이 아예 없는 예전 데이터 + 열 지정 전에 다른 열로 올린 데이터 둘 다 해당)
  const needsReupload = !!(current && profile?.revenue && !sameHeader(current.source?.revenue, profile.revenue));

  const visibleProducts = showAll ? sortedProducts : sortedProducts.slice(0, PAGE_SIZE);
  const hiddenCount = sortedProducts.length - visibleProducts.length;

  return (
    <div className="tab-body">
      <h2 className="section-title"><TrendingUp size={18} /> 플랫폼 매출 업로드</h2>

      <div className="panel">
        <div className="upload-row">
          <select className="select" value={platform} onChange={(e) => setPlatform(e.target.value)}>
            {PLATFORMS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          <select className="select" value={month} onChange={(e) => setMonth(e.target.value)}>
            {months.map((m) => <option key={m} value={m}>{formatMonthLabel(m)}</option>)}
          </select>
          <label className="upload-btn">
            <Upload size={15} />
            {parsing ? "읽는 중..." : "엑셀 업로드 (.xlsx/.csv)"}
            <input type="file" accept=".xlsx,.xls,.csv" hidden onChange={handleFile} disabled={parsing} />
          </label>
          {current && (
            <button className="ghost-btn" onClick={handleClear}><Trash2 size={14} /> 이 달 데이터 삭제</button>
          )}
        </div>
        {profile ? (
          <div className="hint">
            {platform} 엑셀은{" "}
            {profileFields.map(([field, header], i) => (
              <React.Fragment key={field}>{i > 0 && ", "}<b>{header}</b>({COLUMN_LABELS[field]})</React.Fragment>
            ))}{" "}
            열을 이름 그대로 찾아요.{(!profile.name || !profile.qty) && " 상품명·수량 열은 단어로 찾으니 업로드 후 '읽은 열'을 확인해주세요."} 문구만 다른 같은 상품은 한 줄로 묶어서 보여줘요.
          </div>
        ) : (
          <div className="hint">헤더에 <b>상품명</b>과 <b>수익금(매출)</b> 열이 있으면 자동 인식해요. 같은 상품명은 자동으로 합쳐지고, 수량 열이 없으면 거래 건수를 판매수량으로 계산해요.</div>
        )}

        {current ? (
          <div className="rev-summary">
            {needsReupload && (
              <div className="warn-box">
                <AlertCircle size={15} />
                <div>
                  <b>이 달은 정산 엑셀을 다시 올려주세요.</b><br />
                  {current.source?.revenue
                    ? <>수익이 '{current.source.revenue}' 열로 계산돼 있어요. '{profile.revenue}' 열로 다시 계산하려면 같은 달을 다시 올려주세요. (덮어써져요)</>
                    : <>예전 방식으로 올린 데이터라 수익이 '{profile.revenue}'가 아닌 다른 열로 계산됐을 수 있어요. 같은 달을 다시 올리면 덮어써져요.</>}
                </div>
              </div>
            )}
            {summary ? (
              <RatePanel summary={summary} rows={categoryRates} />
            ) : (
              <div className="hint rate-missing">
                {profile?.sales
                  ? <>이 달은 판매금액 없이 올라가 있어서 수수료율을 계산할 수 없어요. 엑셀을 다시 올리면 분류별 수수료율이 보여요.</>
                  : <>{platform}은 판매금액 열이 아직 지정되지 않아서 수수료율을 계산할 수 없어요.</>}
              </div>
            )}
            <div className="row-between" style={{ marginBottom: 10 }}>
              <div className="rev-total">{platform} · {formatMonthLabel(month)} 총매출 <SplitNumber value={formatWon(current.total)} size="sm" /></div>
              <select
                className="select"
                value={`${sortField}-${sortDir}`}
                onChange={(e) => { const [f, d] = e.target.value.split("-"); setSortField(f); setSortDir(d); }}
              >
                <option value="totalRevenue-desc">총 수익금 높은순</option>
                <option value="totalRevenue-asc">총 수익금 낮은순</option>
                <option value="qty-desc">판매수량 많은순</option>
                <option value="qty-asc">판매수량 적은순</option>
                {summary && <option value="rate-desc">수수료율 높은순</option>}
                {summary && <option value="rate-asc">수수료율 낮은순</option>}
              </select>
            </div>
            {(current.source || mergedCount > 0) && (
              <div className="hint source-hint">
                {current.source && (
                  <>읽은 열: 상품명 <b>{current.source.name}</b> · 수익 <b>{current.source.revenue}</b>{current.source.qty && <> · 수량 <b>{current.source.qty}</b></>}</>
                )}
                {current.source && mergedCount > 0 && " · "}
                {mergedCount > 0 && <>문구만 다른 이름을 묶은 상품 <b>{mergedCount}개</b> ('묶음' 배지에 마우스를 올리면 원래 이름이 보여요)</>}
              </div>
            )}
            <table className="table">
              <thead>
                <tr>
                  <th>순위</th><th>상품명</th><th>판매수량</th>
                  {summary && <th>판매금액</th>}
                  <th>개별 수익금</th><th>총 수익금</th>
                  {summary && <th>수수료율</th>}
                </tr>
              </thead>
              <tbody>
                {visibleProducts.map((p, i) => (
                  <tr key={`${p.name}-${i}`} className={i === 0 ? "top-row" : ""}>
                    <td>{i + 1}</td>
                    <td><ProductName name={p.name} type={classifyProductType(platform, p.name, p.category)} variants={p.variants} /></td>
                    <td>{p.qty}</td>
                    {summary && <td className="num-cell">{p.salesAmount != null ? formatWon(p.salesAmount) : "-"}</td>}
                    <td>{formatWon(p.unitRevenue)}</td>
                    <td>{formatWon(p.totalRevenue)}</td>
                    {summary && <td className="num-cell rate-cell">{formatRate(p.rate)}</td>}
                  </tr>
                ))}
              </tbody>
            </table>
            {sortedProducts.length > PAGE_SIZE && (
              <button className="ghost-btn" style={{ marginTop: 10 }} onClick={() => setShowAll(!showAll)}>
                {showAll ? "접기" : `나머지 ${hiddenCount}개 더보기`}
              </button>
            )}
          </div>
        ) : (
          <EmptyState icon={FileSpreadsheet} title="이 달 데이터가 아직 없어요" desc="플랫폼에서 받은 정산 엑셀을 업로드하면 상품별 수익이 정리돼요." />
        )}
      </div>

      <div className="panel">
        <div className="panel-title">최근 6개월 플랫폼 비교</div>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={compareData}>
            <CartesianGrid stroke="#2A3F5C" strokeDasharray="3 3" />
            <XAxis dataKey="month" tick={{ fill: "#93A2BC", fontSize: 11 }} />
            <YAxis tick={{ fill: "#93A2BC", fontSize: 11 }} tickFormatter={(v) => (v >= 10000 ? `${Math.round(v / 10000)}만` : v)} />
            <Tooltip contentStyle={{ background: "#16273F", border: "1px solid #2A3F5C", color: "#F4F1E8" }} formatter={formatWonOrEmpty} />
            <Legend wrapperStyle={{ fontSize: 12, color: "#93A2BC" }} />
            {PLATFORMS.map((p) => <Bar key={p} dataKey={p} fill={PLATFORM_COLOR[p]} radius={[3, 3, 0, 0]} />)}
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function PostsTab({ months, postsData, updatePostsData, postRecords, updatePostRecords, blogNames, updateBlogNames, showToast, regionTags }) {
  const [blog, setBlog] = useState(blogNames[0]);
  const [month, setMonth] = useState(currentMonth());
  const [postsInput, setPostsInput] = useState("");
  const [viewsInput, setViewsInput] = useState("");
  const [parsing, setParsing] = useState(false);
  const [showAllPosts, setShowAllPosts] = useState(false);
  const POST_PAGE_SIZE = 30;

  useEffect(() => { setBlog(blogNames[0]); }, [blogNames.join("|")]);

  useEffect(() => { setShowAllPosts(false); }, [blog, month]);

  useEffect(() => {
    const d = postsData[blog]?.[month];
    setPostsInput(d ? String(d.posts) : "");
    setViewsInput(d ? String(d.views) : "");
  }, [blog, month, postsData]);

  async function saveManual() {
    const next = {
      ...postsData,
      [blog]: {
        ...(postsData[blog] || {}),
        [month]: { posts: parseNumber(postsInput), views: parseNumber(viewsInput) },
      },
    };
    await updatePostsData(next);
    showToast(`${blog} · ${formatMonthLabel(month)} 저장 완료`);
  }

  async function handleExcelUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setParsing(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!rows.length) throw new Error("빈 파일이에요");
      const titleKey = findKey(rows[0], ["제목", "타이틀", "title", "Title"]);
      const viewsKey = findKey(rows[0], ["조회", "views", "Views"]);
      const platformKey = findKey(rows[0], ["매체", "플랫폼", "platform", "Platform"]);
      const purposeKey = findKey(rows[0], ["목적", "purpose", "Purpose"]);
      const dateKey = findKey(rows[0], ["게시일", "게시날짜", "발행일", "발행날짜", "날짜", "date", "Date"]);
      if (!titleKey || !viewsKey) throw new Error("글 제목 / 조회수 열을 찾을 수 없어요");

      const records = rows
        .map((r) => {
          const title = String(r[titleKey]).trim();
          const dest = matchDestination(title);
          return {
            title,
            views: parseNumber(r[viewsKey]),
            platform: platformKey ? String(r[platformKey]).trim() : "",
            purpose: purposeKey ? String(r[purposeKey]).trim() : "",
            publishDate: dateKey ? parseDateValue(r[dateKey]) : null,
            region: dest ? dest.region : "기타",
            destination: dest ? dest.name : "기타",
            productType: matchProductType(title) || "-",
          };
        })
        .filter((r) => r.title);
      if (!records.length) throw new Error("읽을 수 있는 행이 없어요");

      const nextRecords = { ...postRecords, [blog]: { ...(postRecords[blog] || {}), [month]: records } };
      await updatePostRecords(nextRecords);

      const totalViews = records.reduce((s, r) => s + r.views, 0);
      const nextPosts = { ...postsData, [blog]: { ...(postsData[blog] || {}), [month]: { posts: records.length, views: totalViews } } };
      await updatePostsData(nextPosts);

      showToast(`${blog} · ${formatMonthLabel(month)} 글 ${records.length}건 분석 완료`);
    } catch (err) {
      showToast(err.message || "엑셀 파일을 읽는 데 실패했어요", "err");
    } finally {
      setParsing(false);
      e.target.value = "";
    }
  }

  const history = Object.entries(postsData[blog] || {}).sort((a, b) => (a[0] < b[0] ? 1 : -1)).slice(0, 12);
  const chartData = [...history].reverse().map(([m, d]) => ({ month: formatMonthLabel(m).replace("년 ", ".").replace("월", ""), 조회수: d.views, 발행글: d.posts }));

  // 업로드할 때 저장된 지역 대신, 최신 사전 + 직접 지정한 지역으로 다시 분류해요. (예전에 올린 글도 바로 반영)
  const records = useMemo(
    () => (postRecords[blog]?.[month] || []).map((r) => ({
      ...r,
      ...classifyText(r.title, regionTags, postTagKey(r.title)),
      productType: matchProductType(r.title) || "-", // 예전에 "esim"으로 저장된 글도 "이심"으로 보여요
    })),
    [postRecords, blog, month, regionTags],
  );
  const destAnalysis = useMemo(() => {
    const map = {};
    records.forEach((r) => {
      const key = `${r.region}-${r.destination}`;
      if (!map[key]) map[key] = { region: r.region, destination: r.destination, count: 0, views: 0, productTypes: new Set() };
      map[key].count += 1;
      map[key].views += r.views;
      if (r.productType && r.productType !== "-") map[key].productTypes.add(r.productType);
    });
    return Object.values(map)
      .map((v) => ({ ...v, productTypes: [...v.productTypes].join(", ") || "-" }))
      .sort((a, b) => b.count - a.count);
  }, [records]);

  // 게시일이 있는 글은 경과일 대비 일평균 조회수로 정렬해서, 갓 올라온 글도 공정하게 비교돼요.
  const postPerformance = useMemo(() => {
    return records
      .map((r) => {
        const days = daysSincePublish(r.publishDate);
        const avg = dailyAvgViews(r.views, r.publishDate);
        return { ...r, days, avg };
      })
      .sort((a, b) => {
        if (a.avg == null && b.avg == null) return b.views - a.views;
        if (a.avg == null) return 1;
        if (b.avg == null) return -1;
        return b.avg - a.avg;
      });
  }, [records]);
  const undatedCount = postPerformance.filter((r) => r.days == null).length;
  const visiblePosts = showAllPosts ? postPerformance : postPerformance.slice(0, POST_PAGE_SIZE);
  const hiddenPostsCount = postPerformance.length - visiblePosts.length;

  return (
    <div className="tab-body">
      <h2 className="section-title"><FileSpreadsheet size={18} /> 블로그 운영 현황</h2>

      <div className="panel">
        <div className="panel-title">블로그 이름 설정</div>
        <div className="blog-name-row">
          {blogNames.map((b, i) => (
            <input
              key={i} className="text-input" value={b}
              onChange={(e) => { const arr = [...blogNames]; arr[i] = e.target.value; updateBlogNames(arr); }}
            />
          ))}
        </div>
      </div>

      <div className="panel">
        <div className="upload-row">
          <select className="select" value={blog} onChange={(e) => setBlog(e.target.value)}>
            {blogNames.map((b) => <option key={b} value={b}>{b}</option>)}
          </select>
          <select className="select" value={month} onChange={(e) => setMonth(e.target.value)}>
            {months.map((m) => <option key={m} value={m}>{formatMonthLabel(m)}</option>)}
          </select>
          <label className="upload-btn">
            <Upload size={15} />
            {parsing ? "읽는 중..." : "게시글 엑셀 업로드"}
            <input type="file" accept=".xlsx,.xls,.csv" hidden onChange={handleExcelUpload} disabled={parsing} />
          </label>
        </div>
        <div className="hint">헤더에 <b>글 제목</b> · <b>조회수</b>는 꼭 있어야 하고, <b>소개한 매체</b>(여기어때/마이리얼트립/클룩) · <b>목적</b>(매출/조회수) · <b>게시 날짜</b>는 없어도 괜찮아요(선택). 있으면 이 달 발행 글 수·조회수가 자동 계산되고, 아래에서 나라·상품 분석과 게시일 기준 성과도 볼 수 있어요. 게시 날짜를 넣으면 갓 올라온 글도 조회수가 낮다고 불리하게 평가되지 않고 일평균 조회수로 공정하게 비교돼요.</div>

        <div className="manual-row">
          <div className="field"><label>발행 글 수 (직접 입력)</label><input className="text-input" value={postsInput} onChange={(e) => setPostsInput(e.target.value)} placeholder="예: 45" /></div>
          <div className="field"><label>조회수 (직접 입력)</label><input className="text-input" value={viewsInput} onChange={(e) => setViewsInput(e.target.value)} placeholder="예: 12000" /></div>
          <button className="primary-btn" onClick={saveManual}>이 달 저장</button>
        </div>
      </div>

      <div className="panel">
        <div className="panel-title">{formatMonthLabel(month)} 발행 글 나라·상품 분석 ({blog})</div>
        {destAnalysis.length ? (
          <table className="table">
            <thead><tr><th>지역</th><th>나라/지역명</th><th>포스팅 수</th><th>조회수 합계</th><th>주요 상품유형</th></tr></thead>
            <tbody>
              {destAnalysis.map((d, i) => (
                <tr key={i} className={i === 0 ? "top-row" : ""}>
                  <td><span className={`reco-badge ${d.region === "해외" ? "badge-intl" : d.region === "국내" ? "badge-domestic" : "badge-etc"}`}>{d.region}</span></td>
                  <td>{d.destination}</td>
                  <td>{d.count}건</td>
                  <td>{d.views.toLocaleString("ko-KR")}</td>
                  <td>{d.productTypes}</td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <EmptyState icon={BarChart3} title="이 달 게시글 분석 데이터가 없어요" desc="위에서 게시글 엑셀을 업로드하면 어떤 나라·상품 포스팅이 많았는지 정리돼요." />
        )}
      </div>

      <div className="panel">
        <div className="panel-title">{formatMonthLabel(month)} 게시글별 성과 — 일평균 조회수 기준 ({blog})</div>
        {postPerformance.length ? (
          <>
            <div className="hint" style={{ marginTop: 0 }}>
              게시일 기준 경과일로 나눈 일평균 조회수로 정렬했어요. 최근에 올라온 글도 조회수가 아직 낮다는 이유로 순위에서 불리해지지 않아요.
              {undatedCount > 0 && ` (게시 날짜가 없는 글 ${undatedCount}건은 뒤로 정렬돼요)`}
            </div>
            <table className="table">
              <thead><tr><th>글 제목</th><th>나라/지역</th><th>게시일</th><th>경과일</th><th>조회수</th><th>일평균 조회수</th></tr></thead>
              <tbody>
                {visiblePosts.map((r, i) => (
                  <tr key={`${r.title}-${i}`} className={i === 0 ? "top-row" : ""}>
                    <td>{r.title}</td>
                    <td>{r.destination}</td>
                    <td>{r.publishDate || "-"}</td>
                    <td>{r.days != null ? `${r.days}일` : "-"}</td>
                    <td>{r.views.toLocaleString("ko-KR")}</td>
                    <td>{r.avg != null ? r.avg.toFixed(1) : "-"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {postPerformance.length > POST_PAGE_SIZE && (
              <button className="ghost-btn" style={{ marginTop: 10 }} onClick={() => setShowAllPosts(!showAllPosts)}>
                {showAllPosts ? "접기" : `나머지 ${hiddenPostsCount}개 더보기`}
              </button>
            )}
          </>
        ) : (
          <EmptyState icon={BarChart3} title="이 달 게시글 데이터가 없어요" desc="게시 날짜 열까지 넣어 업로드하면 발행 효율을 공정하게 비교할 수 있어요." />
        )}
      </div>

      <div className="panel">
        <div className="panel-title">{blog} · 최근 추이</div>
        {history.length ? (
          <>
            <ResponsiveContainer width="100%" height={220}>
              <LineChart data={chartData}>
                <CartesianGrid stroke="#2A3F5C" strokeDasharray="3 3" />
                <XAxis dataKey="month" tick={{ fill: "#93A2BC", fontSize: 11 }} />
                <YAxis tick={{ fill: "#93A2BC", fontSize: 11 }} />
                <Tooltip contentStyle={{ background: "#16273F", border: "1px solid #2A3F5C", color: "#F4F1E8" }} />
                <Legend wrapperStyle={{ fontSize: 12, color: "#93A2BC" }} />
                <Line type="monotone" dataKey="조회수" stroke="#5B8DEF" strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="발행글" stroke="#E8A33D" strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
            <table className="table" style={{ marginTop: 12 }}>
              <thead><tr><th>월</th><th>발행 글 수</th><th>조회수</th></tr></thead>
              <tbody>{history.map(([m, d]) => (<tr key={m}><td>{formatMonthLabel(m)}</td><td>{d.posts}건</td><td>{d.views.toLocaleString("ko-KR")}</td></tr>))}</tbody>
            </table>
          </>
        ) : (
          <EmptyState icon={FileSpreadsheet} title="아직 기록이 없어요" desc="위에서 이 달 데이터를 입력하거나 게시글 엑셀을 업로드해주세요." />
        )}
      </div>
    </div>
  );
}

function ConsultingTab({ platformData, postsData, blogNames, consultingHistory, setConsultingHistory, showToast }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  function mostRecentDataMonth() {
    const all = new Set();
    PLATFORMS.forEach((p) => Object.keys(platformData[p] || {}).forEach((m) => all.add(m)));
    if (all.size === 0) return null;
    return [...all].sort().reverse()[0];
  }

  async function runConsulting() {
    const targetMonth = mostRecentDataMonth();
    if (!targetMonth) { setError("먼저 플랫폼 매출 데이터를 업로드해주세요."); return; }
    setLoading(true); setError(null);
    try {
      const revSummary = PLATFORMS.map((p) => {
        const d = platformData[p]?.[targetMonth];
        return `${p}: 총매출 ${d ? formatWon(d.total) : "데이터 없음"}, 베스트셀러 ${d && d.products[0] ? d.products[0].name : "-"}`;
      }).join("\n");
      const postsSummary = blogNames.map((b) => {
        const d = postsData[b]?.[targetMonth];
        return `${b}: 발행글 ${d ? d.posts : 0}건, 조회수 ${d ? d.views : 0}`;
      }).join("\n");

      const prompt = `너는 여행 블로그 어필리에이트 수익화 컨설턴트야. 아래는 ${formatMonthLabel(targetMonth)} 실적 데이터야.

[플랫폼별 매출]
${revSummary}

[블로그별 발행/조회]
${postsSummary}

이 데이터를 참고하고, 최신 여행 트렌드를 웹 검색해서 다음달에 뜰만한 여행지를 분석해줘.
아래 JSON 형식으로만 답해. 다른 설명, 코드블록 표시 없이 순수 JSON 객체 하나만 출력해.

{
  "summary": "지난달 성과에 대한 2~3문장 요약과 개선 포인트",
  "recommendations": [
    {"destination": "여행지명", "region": "해외 또는 국내", "reason": "다음달에 뜰 이유 1~2문장", "contentIdea": "블로그 콘텐츠 아이디어 1문장"}
  ]
}
recommendations는 해외 2곳, 국내 2곳 총 4개로 만들어줘.`;

      // 프론트에서 Anthropic API를 직접 호출하지 않고, API 키를 안전하게 보관하는
      // Netlify Function(netlify/functions/consult.js)을 통해 호출해요.
      // 함수는 "로그인한 허용 계정"만 받아주기 때문에, 로그인 증명서(ID 토큰)를 같이 보내요.
      // getIdToken()은 만료가 가까우면 알아서 새로 받아와요.
      const token = await auth.currentUser?.getIdToken();
      if (!token) throw new Error("로그인 정보를 확인할 수 없어요. 로그아웃 후 다시 로그인해주세요.");
      const response = await fetch("/.netlify/functions/consult", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ prompt }),
      });
      if (!response.ok) {
        const bodyText = await response.text().catch(() => "");
        let detail = "";
        try {
          detail = JSON.parse(bodyText).error || "";
        } catch (parseErr) {
          // 서버가 JSON이 아닌 응답(예: 함수를 못 찾아서 나오는 404 HTML 페이지)을 준 경우
          detail = response.status === 404
            ? "컨설팅 함수(consult)를 찾을 수 없어요. Netlify에 netlify/functions 폴더가 제대로 배포됐는지 확인해주세요."
            : "";
        }
        throw new Error(detail || `컨설팅 서버 호출에 실패했어요 (상태 코드 ${response.status})`);
      }
      const data = await response.json();
      // 웹 검색을 쓰면 답변이 인용 단위로 여러 text 블록에 나뉘어 와요. 줄바꿈 없이 이어 붙여야 JSON이 깨지지 않아요.
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("");
      const clean = text.replace(/```json|```/g, "").trim();
      const jsonStart = clean.indexOf("{");
      const jsonEnd = clean.lastIndexOf("}");
      const parsed = JSON.parse(clean.slice(jsonStart, jsonEnd + 1));

      const entry = { month: targetMonth, generatedAt: new Date().toISOString(), ...parsed };
      const nextHistory = [entry, ...consultingHistory].slice(0, 12);
      setConsultingHistory(nextHistory);
      await safeSet("consulting-history", nextHistory);
      showToast("컨설팅 리포트가 도착했어요");
    } catch (err) {
      setError(err.message || "컨설팅 생성 중 문제가 발생했어요. 잠시 후 다시 시도해주세요.");
    } finally {
      setLoading(false);
    }
  }

  const latest = consultingHistory[0];

  return (
    <div className="tab-body">
      <h2 className="section-title"><Sparkles size={18} /> 다음달 컨설팅</h2>

      <div className="panel consulting-panel">
        <div className="consulting-intro">
          <p>가장 최근 매출/운영 데이터를 바탕으로 지난달 성과를 요약하고, 최신 여행 트렌드를 검색해 다음달에 밀어볼 여행지 4곳을 제안해요.</p>
        </div>
        <button className="primary-btn big" onClick={runConsulting} disabled={loading}>
          {loading ? <><Loader2 size={16} className="spin" /> 분석 중...</> : <><Sparkles size={16} /> 이번 달 컨설팅 받기</>}
        </button>
        {error && <div className="error-box"><AlertCircle size={14} /> {error}</div>}
      </div>

      {latest && (
        <div className="panel">
          <div className="panel-title">{formatMonthLabel(latest.month)} 실적 기준 리포트</div>
          <p className="consult-summary">{latest.summary}</p>
          <div className="reco-grid">
            {(latest.recommendations || []).map((r, i) => (
              <div key={i} className="reco-card">
                <div className={`reco-badge ${r.region === "해외" ? "badge-intl" : "badge-domestic"}`}>{r.region}</div>
                <div className="reco-name"><MapPin size={14} /> {r.destination}</div>
                <div className="reco-reason">{r.reason}</div>
                <div className="reco-idea">💡 {r.contentIdea}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {consultingHistory.length > 1 && (
        <div className="panel">
          <div className="panel-title">지난 리포트</div>
          <div className="history-list">
            {consultingHistory.slice(1).map((h, i) => (
              <details key={i} className="history-item">
                <summary>{formatMonthLabel(h.month)} 리포트</summary>
                <p className="consult-summary">{h.summary}</p>
                <div className="reco-grid">
                  {(h.recommendations || []).map((r, j) => (
                    <div key={j} className="reco-card small">
                      <div className={`reco-badge ${r.region === "해외" ? "badge-intl" : "badge-domestic"}`}>{r.region}</div>
                      <div className="reco-name"><MapPin size={13} /> {r.destination}</div>
                      <div className="reco-reason">{r.reason}</div>
                    </div>
                  ))}
                </div>
              </details>
            ))}
          </div>
        </div>
      )}

      {!latest && <EmptyState icon={Sparkles} title="아직 받은 컨설팅이 없어요" desc="버튼을 눌러 첫 리포트를 받아보세요." />}
    </div>
  );
}

function InsightTab({ months, platformData, postRecords, blogNames, regionTags, updateRegionTag }) {
  const [scope, setScope] = useState("all");

  // 같은 상품이 여러 달에 걸쳐, 또는 문구만 다른 이름으로 여러 줄 나와도 플랫폼별로 하나로 합쳐서 집계해요.
  const allProducts = useMemo(() => {
    const out = [];
    PLATFORMS.forEach((p) => {
      const list = [];
      Object.entries(platformData[p] || {}).forEach(([m, data]) => {
        if (scope !== "all" && m !== scope) return;
        list.push(...(data.products || []));
      });
      groupProducts(list).forEach((g) => {
        const c = classifyProduct(p, g, regionTags);
        out.push({ ...g, platform: p, region: c.region, destination: c.destination, productType: classifyProductType(p, g.name, g.category) });
      });
    });
    return out;
  }, [platformData, scope, regionTags]);

  const allPosts = useMemo(() => {
    const list = [];
    blogNames.forEach((b) => {
      const monthsObj = postRecords[b] || {};
      Object.entries(monthsObj).forEach(([m, records]) => {
        if (scope !== "all" && m !== scope) return;
        records.forEach((r) => list.push({ ...r, ...classifyText(r.title, regionTags, postTagKey(r.title)), blog: b, month: m }));
      });
    });
    return list;
  }, [postRecords, blogNames, scope, regionTags]);

  // 지역 직접 지정용 목록: 기간 선택과 상관없이 전체 기간의 상품·글을 하나씩만 모아요.
  const taggableItems = useMemo(() => {
    const map = {};
    PLATFORMS.forEach((p) => {
      const list = Object.values(platformData[p] || {}).flatMap((data) => data.products || []);
      groupProducts(list).forEach((g) => {
        const key = productTagKey(p, g.name);
        // cleanupKeys: 저장·해제할 때 예전 방식(원래 이름) 태그도 같이 정리하려고 들고 있어요.
        map[key] = { key, cleanupKeys: productTagKeys(p, g), kind: "상품", name: g.name, variants: g.variants, source: p, revenue: g.totalRevenue, views: 0, ...classifyProduct(p, g, regionTags) };
      });
    });
    blogNames.forEach((b) => {
      Object.values(postRecords[b] || {}).forEach((records) => {
        records.forEach((r) => {
          const key = postTagKey(r.title);
          if (!map[key]) map[key] = { key, kind: "글", name: r.title, source: b, revenue: 0, views: 0, ...classifyText(r.title, regionTags, key) };
          map[key].views += r.views || 0;
        });
      });
    });
    return Object.values(map).sort((a, b) => b.revenue - a.revenue || b.views - a.views);
  }, [platformData, postRecords, blogNames, regionTags]);

  const revenueTop10 = useMemo(() => [...allProducts].sort((a, b) => b.totalRevenue - a.totalRevenue).slice(0, 10), [allProducts]);
  const qtyTop10 = useMemo(() => [...allProducts].sort((a, b) => b.qty - a.qty).slice(0, 10), [allProducts]);
  const viewsTop10 = useMemo(() => [...allPosts].sort((a, b) => b.views - a.views).slice(0, 10), [allPosts]);
  // 발행 효율: 게시일이 있는 글만 대상으로 일평균 조회수 상위 10개. 최근에 올라온 글도 공정하게 비교돼요.
  const efficiencyTop10 = useMemo(() => {
    return allPosts
      .map((r) => ({ ...r, days: daysSincePublish(r.publishDate), avg: dailyAvgViews(r.views, r.publishDate) }))
      .filter((r) => r.avg != null)
      .sort((a, b) => b.avg - a.avg)
      .slice(0, 10);
  }, [allPosts]);

  const compositeTop10 = useMemo(() => {
    const map = {};
    allProducts.forEach((p) => {
      const key = `${p.region}-${p.destination}`;
      if (!map[key]) map[key] = { region: p.region, destination: p.destination, revenue: 0, views: 0, productCount: 0, postCount: 0 };
      map[key].revenue += p.totalRevenue;
      map[key].productCount += 1;
    });
    allPosts.forEach((r) => {
      const key = `${r.region}-${r.destination}`;
      if (!map[key]) map[key] = { region: r.region, destination: r.destination, revenue: 0, views: 0, productCount: 0, postCount: 0 };
      map[key].views += r.views;
      map[key].postCount += 1;
    });
    const rows = Object.values(map);
    const maxRevenue = Math.max(1, ...rows.map((r) => r.revenue));
    const maxViews = Math.max(1, ...rows.map((r) => r.views));
    return rows
      .map((r) => ({ ...r, score: (r.revenue / maxRevenue) * 0.5 + (r.views / maxViews) * 0.5 }))
      .sort((a, b) => b.score - a.score)
      .slice(0, 10);
  }, [allProducts, allPosts]);

  const hasData = allProducts.length > 0 || allPosts.length > 0;
  const regionBadge = (region) => `reco-badge ${region === "해외" ? "badge-intl" : region === "국내" ? "badge-domestic" : "badge-etc"}`;

  return (
    <div className="tab-body">
      <div className="row-between">
        <h2 className="section-title"><Trophy size={18} /> 인사이트 요약</h2>
        <select className="month-select" value={scope} onChange={(e) => setScope(e.target.value)}>
          <option value="all">전체 기간</option>
          {months.map((m) => <option key={m} value={m}>{formatMonthLabel(m)}</option>)}
        </select>
      </div>

      {!hasData ? (
        <div className="panel">
          <EmptyState icon={Trophy} title="분석할 데이터가 아직 없어요" desc="플랫폼 매출 엑셀과 블로그 운영 게시글 엑셀을 업로드하면 인사이트가 만들어져요." />
        </div>
      ) : (
        <>
          <div className="panel">
            <div className="panel-title">💰 팔린 금액 TOP 10 — 나라/상품</div>
            {revenueTop10.length ? (
              <table className="table">
                <thead><tr><th>순위</th><th>지역</th><th>나라/지역</th><th>상품명</th><th>플랫폼</th><th>총 수익금</th></tr></thead>
                <tbody>
                  {revenueTop10.map((p, i) => (
                    <tr key={`${p.platform}-${p.name}`} className={i === 0 ? "top-row" : ""}>
                      <td>{i + 1}</td>
                      <td><span className={regionBadge(p.region)}>{p.region}</span></td>
                      <td>{p.destination}</td>
                      <td><ProductName name={p.name} type={p.productType} variants={p.variants} /></td>
                      <td>{p.platform}</td>
                      <td>{formatWon(p.totalRevenue)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : <EmptyState icon={TrendingUp} title="매출 데이터가 없어요" desc="플랫폼 매출 탭에서 엑셀을 업로드해주세요." />}
          </div>

          <div className="panel">
            <div className="panel-title">📦 팔린 개수 TOP 10 — 나라/상품</div>
            {qtyTop10.length ? (
              <table className="table">
                <thead><tr><th>순위</th><th>지역</th><th>나라/지역</th><th>상품명</th><th>플랫폼</th><th>판매수량</th></tr></thead>
                <tbody>
                  {qtyTop10.map((p, i) => (
                    <tr key={`${p.platform}-${p.name}`} className={i === 0 ? "top-row" : ""}>
                      <td>{i + 1}</td>
                      <td><span className={regionBadge(p.region)}>{p.region}</span></td>
                      <td>{p.destination}</td>
                      <td><ProductName name={p.name} type={p.productType} variants={p.variants} /></td>
                      <td>{p.platform}</td>
                      <td>{p.qty.toLocaleString("ko-KR")}건</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : <EmptyState icon={TrendingUp} title="매출 데이터가 없어요" desc="플랫폼 매출 탭에서 엑셀을 업로드해주세요." />}
          </div>

          <div className="panel">
            <div className="panel-title">👀 조회수 TOP 10 — 나라/상품</div>
            {viewsTop10.length ? (
              <table className="table">
                <thead><tr><th>순위</th><th>지역</th><th>나라/지역</th><th>글 제목</th><th>소개 매체</th><th>조회수</th></tr></thead>
                <tbody>
                  {viewsTop10.map((r, i) => (
                    <tr key={i} className={i === 0 ? "top-row" : ""}>
                      <td>{i + 1}</td>
                      <td><span className={regionBadge(r.region)}>{r.region}</span></td>
                      <td>{r.destination}</td>
                      <td>{r.title}</td>
                      <td>{r.platform || "-"}</td>
                      <td>{r.views.toLocaleString("ko-KR")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : <EmptyState icon={FileSpreadsheet} title="게시글 데이터가 없어요" desc="블로그 운영 탭에서 게시글 엑셀을 업로드해주세요." />}
          </div>

          <div className="panel">
            <div className="panel-title">⚡ 발행 효율 TOP 10 — 일평균 조회수</div>
            <div className="hint" style={{ marginTop: 0, marginBottom: 12 }}>게시일 대비 일평균 조회수 기준이라, 최근에 올린 글도 조회수가 아직 낮다는 이유로 밀리지 않아요. (게시 날짜를 넣어 업로드한 글만 집계돼요)</div>
            {efficiencyTop10.length ? (
              <table className="table">
                <thead><tr><th>순위</th><th>지역</th><th>나라/지역</th><th>글 제목</th><th>게시일</th><th>경과일</th><th>일평균 조회수</th></tr></thead>
                <tbody>
                  {efficiencyTop10.map((r, i) => (
                    <tr key={i} className={i === 0 ? "top-row" : ""}>
                      <td>{i + 1}</td>
                      <td><span className={regionBadge(r.region)}>{r.region}</span></td>
                      <td>{r.destination}</td>
                      <td>{r.title}</td>
                      <td>{r.publishDate}</td>
                      <td>{r.days}일</td>
                      <td>{r.avg.toFixed(1)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : <EmptyState icon={FileSpreadsheet} title="게시 날짜가 있는 글이 없어요" desc="블로그 운영 탭에서 게시 날짜 열을 포함해 엑셀을 업로드해주세요." />}
          </div>

          <div className="panel">
            <div className="panel-title">🏆 종합 TOP 10 — 나라/지역/상품 분석</div>
            <div className="hint" style={{ marginTop: 0, marginBottom: 12 }}>매출 비중과 조회수 비중을 절반씩 반영한 종합 인기 순위예요.</div>
            {compositeTop10.length ? (
              <table className="table">
                <thead><tr><th>순위</th><th>지역</th><th>나라/지역</th><th>총매출</th><th>총조회수</th><th>관련 상품</th><th>관련 게시글</th></tr></thead>
                <tbody>
                  {compositeTop10.map((r, i) => (
                    <tr key={i} className={i === 0 ? "top-row" : ""}>
                      <td>{i + 1}</td>
                      <td><span className={regionBadge(r.region)}>{r.region}</span></td>
                      <td>{r.destination}</td>
                      <td>{formatWon(r.revenue)}</td>
                      <td>{r.views.toLocaleString("ko-KR")}</td>
                      <td>{r.productCount}건</td>
                      <td>{r.postCount}건</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : <EmptyState icon={Trophy} title="종합 분석할 데이터가 없어요" desc="매출과 게시글 데이터를 함께 업로드하면 종합 순위가 만들어져요." />}
          </div>

          <RateComparePanel platformData={platformData} scope={scope} />

          <RegionTagPanel items={taggableItems} regionTags={regionTags} onSave={updateRegionTag} />
        </>
      )}
    </div>
  );
}

// 플랫폼별 수수료율을 분류별로 나란히 비교해요. 같은 분류에서 수수료율이 가장 높은 플랫폼에 '유리' 표시를 붙여요.
// (색만으로 구분하지 않게 글자 표시를 같이 써요) 아래에는 이름이 같은 상품끼리의 비교도 보여줘요.
function buildRateCompare(platformData, scope) {
  const perPlatform = {};
  const byProduct = {}; // 핵심 이름 → { 플랫폼: 묶인 상품 }
  PLATFORMS.forEach((p) => {
    const list = [];
    Object.entries(platformData[p] || {}).forEach(([m, data]) => {
      if (scope !== "all" && m !== scope) return;
      list.push(...(data.products || []));
    });
    const groups = groupProducts(list);
    const summary = rateSummary(groups);
    if (!summary) return; // 판매금액이 하나도 없는 플랫폼은 비교에서 빠져요
    perPlatform[p] = { summary, rows: rateByCategory(p, groups) };
    groups.forEach((g) => {
      if (g.salesAmount == null) return;
      const k = productGroupKey(g.name);
      if (!byProduct[k]) byProduct[k] = {};
      byProduct[k][p] = g;
    });
  });
  const platforms = PLATFORMS.filter((p) => perPlatform[p]);
  const bestOf = (cells) => {
    const rated = platforms.filter((p) => cells[p]?.rate != null);
    if (rated.length < 2) return null; // 비교할 상대가 없으면 '유리' 표시 안 해요
    return rated.reduce((a, b) => (cells[b].rate > cells[a].rate ? b : a));
  };
  const categories = [...new Set(platforms.flatMap((p) => perPlatform[p].rows.map((r) => r.category)))];
  const rows = categories
    .map((category) => {
      const cells = Object.fromEntries(platforms.map((p) => [p, perPlatform[p].rows.find((r) => r.category === category) || null]));
      const sales = platforms.reduce((s, p) => s + (cells[p]?.sales || 0), 0);
      return { category, cells, best: bestOf(cells), sales };
    })
    .sort((a, b) => b.sales - a.sales);
  const totalCells = Object.fromEntries(platforms.map((p) => [p, perPlatform[p].summary]));
  const sameProducts = Object.values(byProduct)
    .filter((cells) => Object.keys(cells).length > 1)
    .map((cells) => {
      const first = Object.values(cells).sort((a, b) => b.salesAmount - a.salesAmount)[0];
      return { name: first.name, cells, best: bestOf(cells), sales: Object.values(cells).reduce((s, g) => s + g.salesAmount, 0) };
    })
    .sort((a, b) => b.sales - a.sales);
  return { platforms, rows, total: { cells: totalCells }, sameProducts };
}

function RateCell({ cell, isBest }) {
  if (!cell || cell.rate == null) return <td className="rate-compare-cell empty">-</td>;
  const settle = cell.settle;
  return (
    <td className={`rate-compare-cell ${isBest ? "best" : ""}`} title={`판매금액 ${formatWon(cell.sales)} · 정산금 ${formatWon(settle)}`}>
      <div className="rate-compare-value">
        {formatRate(cell.rate)}
        {isBest && <span className="best-chip">유리</span>}
      </div>
      <div className="rate-compare-sub">판매 {formatWon(cell.sales)}</div>
    </td>
  );
}

function RateComparePanel({ platformData, scope }) {
  const data = useMemo(() => buildRateCompare(platformData, scope), [platformData, scope]);
  const { platforms, rows, total, sameProducts } = data;
  const header = (
    <tr>
      <th>분류</th>
      {platforms.map((p) => (
        <th key={p}><span className="platform-swatch" style={{ background: PLATFORM_COLOR[p] }} />{p}</th>
      ))}
    </tr>
  );
  return (
    <div className="panel">
      <div className="panel-title">🧾 플랫폼별 수수료율 비교 — 분류별</div>
      <div className="hint" style={{ marginTop: 0, marginBottom: 12 }}>
        수수료율은 판매금액 대비 정산금 비율이에요. 같은 분류에서 수수료율이 가장 높은 플랫폼에 <b>유리</b> 표시가 붙어요. 판매금액 열과 함께 올린 달만 계산돼요.
      </div>
      {platforms.length === 0 ? (
        <EmptyState icon={BarChart3} title="수수료율을 계산할 데이터가 없어요" desc="판매금액 열이 있는 정산 엑셀을 플랫폼 매출 탭에서 올려주세요." />
      ) : (
        <>
          {platforms.length === 1 && (
            <div className="hint" style={{ marginTop: 0, marginBottom: 10 }}>
              지금은 <b>{platforms[0]}</b>만 판매금액이 있어서 나란히 비교할 플랫폼이 없어요. 다른 플랫폼도 판매금액 열이 지정되면 옆에 같이 나와요.
            </div>
          )}
          <div className="table-scroll">
            <table className="table rate-compare-table">
              <thead>{header}</thead>
              <tbody>
                {/* 전체 줄은 상품 구성(이심 비중 등)에 따라 달라져서 '유리' 표시를 안 붙여요. 비교는 분류별 줄로 해요. */}
                <tr className="rate-compare-total">
                  <td>전체</td>
                  {platforms.map((p) => <RateCell key={p} cell={total.cells[p]} isBest={false} />)}
                </tr>
                {rows.map((r) => (
                  <tr key={r.category}>
                    <td><span className="type-badge">{r.category}</span></td>
                    {platforms.map((p) => <RateCell key={p} cell={r.cells[p]} isBest={r.best === p} />)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {platforms.length > 1 && (
            <>
              <div className="panel-subtitle">같은 상품 비교</div>
              {sameProducts.length ? (
                <div className="table-scroll">
                  <table className="table rate-compare-table">
                    <thead>
                      <tr>
                        <th>상품명</th>
                        {platforms.map((p) => <th key={p}><span className="platform-swatch" style={{ background: PLATFORM_COLOR[p] }} />{p}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {sameProducts.map((s) => (
                        <tr key={s.name}>
                          <td className="tag-name">{s.name}</td>
                          {platforms.map((p) => <RateCell key={p} cell={s.cells[p] ? { rate: s.cells[p].rate, sales: s.cells[p].salesAmount, settle: s.cells[p].totalRevenue } : null} isBest={s.best === p} />)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="hint" style={{ marginTop: 0 }}>
                  두 플랫폼에서 이름이 같은 상품은 아직 없어요. 플랫폼마다 상품명 표기가 달라서, 같은 상품인데 여기 안 잡힐 수 있어요. 그럴 땐 위의 분류별 비교를 참고해주세요.
                </div>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
}

const TAG_PAGE_SIZE = 15;
function RegionTagPanel({ items, regionTags, onSave }) {
  const [mode, setMode] = useState("untagged"); // untagged: 미분류 / manual: 직접 지정한 항목
  const [query, setQuery] = useState("");
  const [limit, setLimit] = useState(TAG_PAGE_SIZE);

  const untagged = items.filter((i) => i.region === "기타");
  const manual = items.filter((i) => i.manual);
  const q = query.trim();
  const list = q ? items.filter((i) => i.name.includes(q)) : mode === "manual" ? manual : untagged;
  const untaggedRevenue = untagged.reduce((s, i) => s + i.revenue, 0);

  // 자동완성 후보: 사전 지명 + 직접 입력했던 지명
  const destOptions = useMemo(() => {
    const set = new Set(DESTINATIONS.map((d) => d.name));
    Object.values(regionTags || {}).forEach((t) => set.add(t.destination));
    return [...set];
  }, [regionTags]);

  function switchMode(m) { setMode(m); setQuery(""); setLimit(TAG_PAGE_SIZE); }

  return (
    <div className="panel">
      <div className="panel-title">🏷️ 지역 직접 지정하기</div>
      <div className="hint" style={{ marginTop: 0, marginBottom: 12 }}>
        지역을 찾지 못해 '기타'로 빠진 항목이에요. 지역을 지정하면 위의 모든 표와 블로그 운영 탭에 바로 반영돼요.
        {untagged.length > 0 && <> 지금 미분류: <b>{untagged.length}건 · {formatWon(untaggedRevenue)}</b></>}
      </div>
      <div className="tag-toolbar">
        <div className="seg">
          <button className={`ghost-btn ${!q && mode === "untagged" ? "active" : ""}`} onClick={() => switchMode("untagged")}>미분류 {untagged.length}</button>
          <button className={`ghost-btn ${!q && mode === "manual" ? "active" : ""}`} onClick={() => switchMode("manual")}>직접 지정함 {manual.length}</button>
        </div>
        <input
          className="text-input tag-search" value={query} placeholder="상품명·글 제목 검색 (잘못 분류된 항목 고치기)"
          onChange={(e) => { setQuery(e.target.value); setLimit(TAG_PAGE_SIZE); }}
        />
      </div>

      {list.length ? (
        <>
          <div className="table-scroll">
            <table className="table">
              <thead><tr><th>구분</th><th>상품명 / 글 제목</th><th>매출 · 조회수</th><th>지역 지정</th></tr></thead>
              <tbody>
                {list.slice(0, limit).map((item) => (
                  <TagRow key={`${item.key}-${item.manual}`} item={item} onSave={onSave} />
                ))}
              </tbody>
            </table>
          </div>
          {list.length > limit && (
            <button className="ghost-btn tag-more" onClick={() => setLimit(limit + TAG_PAGE_SIZE)}>
              더 보기 ({list.length - limit}건 남음)
            </button>
          )}
        </>
      ) : (
        <EmptyState
          icon={MapPin}
          title={q ? "검색 결과가 없어요" : mode === "manual" ? "직접 지정한 항목이 아직 없어요" : "미분류 항목이 없어요 🎉"}
          desc={q ? "다른 단어로 검색해보세요." : mode === "manual" ? "미분류 목록에서 지역을 지정하면 여기에 모여요." : "모든 상품과 글의 지역을 찾았어요."}
        />
      )}
      <datalist id="destination-options">
        {destOptions.map((d) => <option key={d} value={d} />)}
      </datalist>
    </div>
  );
}

function TagRow({ item, onSave }) {
  const known = item.region !== "기타";
  const [destination, setDestination] = useState(known ? item.destination : "");
  const [region, setRegion] = useState(known ? item.region : "해외");
  const [saving, setSaving] = useState(false);

  // 사전에 있는 지명을 고르면 국내/해외도 자동으로 맞춰요.
  function changeDestination(v) {
    setDestination(v);
    const d = DESTINATIONS.find((x) => x.name === v.trim());
    if (d) setRegion(d.region);
  }
  async function run(tag) { setSaving(true); await onSave(item.key, tag, item.cleanupKeys || []); setSaving(false); }

  const value = destination.trim();
  const unchanged = known && value === item.destination && region === item.region;
  return (
    <tr>
      <td><span className="reco-badge badge-etc kind-badge">{item.kind}</span></td>
      <td className="tag-name">
        {item.name}
        <div className="tag-sub">
          {item.source} · 지금: {item.destination}{item.manual ? " (직접 지정)" : known ? " (자동)" : ""}
          {item.variants?.length > 1 && ` · 비슷한 이름 ${item.variants.length}개를 묶었어요`}
        </div>
      </td>
      <td className="num-cell">{item.kind === "상품" ? formatWon(item.revenue) : `${item.views.toLocaleString("ko-KR")}회`}</td>
      <td>
        <div className="tag-controls">
          <select className="select" value={region} onChange={(e) => setRegion(e.target.value)}>
            <option value="해외">해외</option>
            <option value="국내">국내</option>
          </select>
          <input
            className="text-input" list="destination-options" value={destination} placeholder="예: 나고야"
            onChange={(e) => changeDestination(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && value && !unchanged) run({ region, destination: value }); }}
          />
          <button className="primary-btn small" disabled={!value || unchanged || saving} onClick={() => run({ region, destination: value })}>
            {saving ? <Loader2 size={13} className="spin" /> : "저장"}
          </button>
          {item.manual && (
            <button className="ghost-btn" disabled={saving} onClick={() => run(null)}>해제</button>
          )}
        </div>
      </td>
    </tr>
  );
}

function Style() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');

      * { box-sizing: border-box; }
      .app-shell {
        background: #0F1B2D;
        color: #F4F1E8;
        font-family: 'Inter', -apple-system, sans-serif;
        min-height: 100%;
        padding: 0 0 40px 0;
      }
      .loading-shell { display:flex; align-items:center; justify-content:center; height: 300px; color:#93A2BC; }
      .login-shell { display:flex; align-items:center; justify-content:center; min-height: 100vh; }
      .login-card {
        background: #16273F; border: 1px solid #2A3F5C; border-radius: 12px; padding: 32px 28px;
        width: 320px; display: flex; flex-direction: column; gap: 12px; align-items: center;
      }
      .login-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 16px; letter-spacing: 0.5px; }
      .login-sub { font-size: 12px; color: #93A2BC; text-align: center; line-height: 1.6; margin-bottom: 8px; }
      .login-sub b { color: #F4F1E8; font-weight: 600; }
      .login-card .text-input { width: 100%; min-width: 0; }
      .login-card .brand-mark.denied { background: rgba(239,91,91,0.15); color: #EF5B5B; }
      .google-btn {
        width: 100%; display: flex; align-items: center; justify-content: center; gap: 10px;
        background: #F4F1E8; color: #0F1B2D; border: none; border-radius: 6px;
        padding: 11px 16px; font-size: 13px; font-weight: 600; font-family: 'Inter', sans-serif; cursor: pointer;
        transition: background .15s;
      }
      .google-btn:hover { background: #fff; }
      .google-btn:disabled { opacity: 0.7; cursor: wait; }
      .login-error { align-items: flex-start; line-height: 1.5; text-align: left; }
      @media (max-width: 380px) { .login-card { width: calc(100vw - 32px); padding: 28px 20px; } }
      .logout-btn {
        margin-left: auto; background: transparent; border: 1px solid #2A3F5C; color: #93A2BC;
        width: 30px; height: 30px; border-radius: 6px; display: flex; align-items: center; justify-content: center; cursor: pointer;
      }
      .logout-btn:hover { color: #F4F1E8; border-color: #56698c; }
      .spin { animation: spin 1s linear infinite; }
      @keyframes spin { to { transform: rotate(360deg); } }

      .header {
        padding: 22px 24px 16px 24px;
        border-bottom: 1px solid #2A3F5C;
        display: flex; flex-direction: column; gap: 12px;
      }
      .brand { display: flex; align-items: center; gap: 12px; }
      .brand-mark {
        width: 36px; height: 36px; border-radius: 8px;
        background: linear-gradient(135deg, #E8A33D, #c9822a);
        display: flex; align-items: center; justify-content: center; color: #0F1B2D;
        transform: rotate(-8deg);
      }
      .brand-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 18px; letter-spacing: 0.5px; }
      .brand-sub { font-size: 12px; color: #93A2BC; margin-top: 2px; }
      .boarding-strip {
        font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: #E8A33D;
        letter-spacing: 1.5px; text-transform: uppercase; display: flex; gap: 8px; align-items:center;
      }
      .boarding-strip .dot { color: #2A3F5C; }

      .tabs { display: flex; gap: 4px; padding: 12px 24px 0 24px; flex-wrap: wrap; }
      .tab {
        display: flex; align-items: center; gap: 7px; padding: 10px 14px;
        background: transparent; border: 1px solid transparent; border-bottom: none;
        color: #93A2BC; font-size: 13px; font-weight: 500; cursor: pointer;
        border-radius: 8px 8px 0 0; font-family: 'Inter', sans-serif;
      }
      .tab-gate { font-family: 'IBM Plex Mono', monospace; font-size: 9px; color: #56698c; letter-spacing: 1px; }
      .tab.active { background: #16273F; border-color: #2A3F5C; color: #F4F1E8; }
      .tab.active .tab-gate { color: #E8A33D; }
      .tab:hover:not(.active) { color: #F4F1E8; }

      .main { padding: 0 24px; }
      .tab-body {
        background: #16273F; border: 1px solid #2A3F5C; border-radius: 0 10px 10px 10px;
        padding: 22px; display: flex; flex-direction: column; gap: 18px;
      }

      .section-title { display: flex; align-items: center; gap: 8px; font-family: 'Space Grotesk', sans-serif; font-size: 16px; font-weight: 700; margin: 0; }
      .row-between { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; }

      .month-select, .select {
        background: #1C324F; border: 1px solid #2A3F5C; color: #F4F1E8;
        padding: 8px 10px; border-radius: 6px; font-size: 13px; font-family: 'Inter', sans-serif;
      }

      .stat-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 14px; }
      .stat-card {
        background: #1C324F; border: 1px solid #2A3F5C; border-radius: 10px; padding: 16px 18px;
        position: relative;
      }
      .stat-card.stub::before {
        content: ""; position: absolute; left: -6px; top: 50%; width: 12px; height: 12px;
        background: #0F1B2D; border-radius: 50%; transform: translateY(-50%);
      }
      .stat-label { font-size: 12px; color: #93A2BC; margin-bottom: 8px; }
      .split-number {
        font-family: 'IBM Plex Mono', monospace; color: #E8A33D; letter-spacing: 1px;
        background: repeating-linear-gradient(180deg, rgba(232,163,61,0.05) 0px, rgba(232,163,61,0.05) 1px, transparent 1px, transparent 4px);
      }
      .split-lg { font-size: 26px; font-weight: 600; }
      .split-sm { font-size: 20px; font-weight: 600; }

      .grid-2 { display: grid; grid-template-columns: 1.3fr 1fr; gap: 14px; }
      @media (max-width: 800px) { .grid-2 { grid-template-columns: 1fr; } }

      .panel { background: #1C324F; border: 1px solid #2A3F5C; border-radius: 10px; padding: 16px; }
      .panel-title { font-size: 13px; font-weight: 600; color: #F4F1E8; margin-bottom: 12px; }

      .blog-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; }
      .blog-card { background: #16273F; border: 1px dashed #2A3F5C; border-radius: 8px; padding: 12px 14px; }
      .blog-card-name { display: flex; align-items: center; gap: 6px; font-weight: 600; font-size: 13px; margin-bottom: 8px; color: #E8A33D; }
      .blog-card-row { display: flex; justify-content: space-between; font-size: 13px; color: #93A2BC; padding: 3px 0; }
      .blog-card-row b { color: #F4F1E8; }

      .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 8px; padding: 36px 12px; color: #56698c; text-align: center; }
      .empty-title { font-size: 13px; font-weight: 600; color: #93A2BC; }
      .empty-desc { font-size: 12px; }

      .upload-row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
      .upload-btn {
        display: flex; align-items: center; gap: 7px; background: #E8A33D; color: #0F1B2D;
        padding: 9px 14px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer;
      }
      .ghost-btn {
        display: flex; align-items: center; gap: 6px; background: transparent; border: 1px solid #2A3F5C;
        color: #93A2BC; padding: 8px 12px; border-radius: 6px; font-size: 12px; cursor: pointer;
      }
      .ghost-btn:hover { color: #F4F1E8; border-color: #56698c; }
      .hint { font-size: 12px; color: #56698c; margin-top: 10px; }
      .hint b { color: #93A2BC; }

      .rev-summary { margin-top: 16px; }
      .rev-total { font-size: 13px; color: #93A2BC; margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }

      .table { width: 100%; border-collapse: collapse; font-size: 13px; }
      .table th { text-align: left; color: #56698c; font-weight: 500; padding: 8px 10px; border-bottom: 1px solid #2A3F5C; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; }
      .table td { padding: 8px 10px; border-bottom: 1px solid #22344d; color: #F4F1E8; }
      .table .top-row td { color: #E8A33D; font-weight: 600; }

      .manual-row { display: flex; gap: 14px; align-items: flex-end; flex-wrap: wrap; margin-top: 14px; }
      .field { display: flex; flex-direction: column; gap: 5px; }
      .field label { font-size: 11px; color: #56698c; }
      .text-input {
        background: #16273F; border: 1px solid #2A3F5C; color: #F4F1E8;
        padding: 9px 10px; border-radius: 6px; font-size: 13px; font-family: 'Inter', sans-serif; min-width: 140px;
      }
      .primary-btn {
        display: flex; align-items: center; gap: 7px; background: #E8A33D; color: #0F1B2D; border: none;
        padding: 10px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer;
      }
      .primary-btn.big { padding: 13px 20px; font-size: 14px; }
      .primary-btn:disabled { opacity: 0.6; cursor: not-allowed; }

      .blog-name-row { display: flex; gap: 10px; flex-wrap: wrap; }
      .sheet-fetch { display: flex; align-items: center; gap: 8px; flex: 1; min-width: 260px; }
      .sheet-fetch .text-input { flex: 1; }

      .toast {
        position: fixed; bottom: 22px; left: 50%; transform: translateX(-50%);
        background: #1C324F; border: 1px solid #2A3F5C; padding: 10px 16px; border-radius: 8px;
        display: flex; align-items: center; gap: 8px; font-size: 13px; z-index: 50;
      }
      .toast-ok { color: #7CD992; }
      .toast-err { color: #EF5B5B; }

      .consulting-panel { display: flex; flex-direction: column; gap: 12px; align-items: flex-start; }
      .consulting-intro p { font-size: 13px; color: #93A2BC; margin: 0; line-height: 1.6; max-width: 560px; }
      .error-box { display: flex; align-items: center; gap: 6px; color: #EF5B5B; font-size: 12px; }

      .consult-summary { font-size: 13px; color: #d7dced; line-height: 1.7; margin: 0 0 14px 0; }
      .reco-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; }
      .reco-card { background: #16273F; border: 1px solid #2A3F5C; border-radius: 8px; padding: 14px; position: relative; }
      .reco-card.small { padding: 10px 12px; }
      .reco-badge {
        display: inline-block; font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 10px;
        margin-bottom: 8px; letter-spacing: 0.5px;
      }
      .badge-intl { background: rgba(91,141,239,0.18); color: #5B8DEF; }
      .badge-domestic { background: rgba(232,163,61,0.18); color: #E8A33D; }
      .badge-etc { background: rgba(147,162,188,0.18); color: #93A2BC; }
      .product-name { display: inline-flex; align-items: center; flex-wrap: wrap; gap: 6px; }
      .type-badge {
        display: inline-block; font-size: 10px; font-weight: 600; line-height: 1; padding: 3px 6px;
        border-radius: 4px; border: 1px solid rgba(232,163,61,0.45); color: #E8A33D;
        letter-spacing: 0.3px; white-space: nowrap;
      }
      .table .top-row .type-badge { font-weight: 700; }
      .group-badge {
        display: inline-block; font-size: 10px; font-weight: 500; line-height: 1; padding: 3px 6px;
        border-radius: 4px; background: rgba(147,162,188,0.14); color: #93A2BC; white-space: nowrap; cursor: help;
      }
      .warn-box {
        display: flex; gap: 10px; align-items: flex-start; margin-bottom: 14px; padding: 12px 14px;
        border-radius: 8px; border: 1px solid rgba(232,163,61,0.45); background: rgba(232,163,61,0.08);
        color: #F4F1E8; font-size: 12px; line-height: 1.6;
      }
      .warn-box svg { color: #E8A33D; flex-shrink: 0; margin-top: 2px; }
      .warn-box b { color: #E8A33D; }
      .source-hint { margin-top: 0; margin-bottom: 10px; }

      /* 수수료율: 지표 카드 + 분류별 표 */
      .rate-panel { margin-bottom: 18px; padding-bottom: 16px; border-bottom: 1px solid #2A3F5C; }
      .kpi-row { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; margin-bottom: 12px; }
      @media (max-width: 640px) { .kpi-row { grid-template-columns: 1fr; } }
      .kpi { background: #16273F; border: 1px solid #2A3F5C; border-radius: 8px; padding: 12px 14px; min-width: 0; }
      .kpi-label { font-size: 11px; color: #93A2BC; margin-bottom: 6px; }
      .kpi-value { font-family: 'IBM Plex Mono', monospace; font-size: 17px; font-weight: 600; color: #F4F1E8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .kpi-sub { font-size: 11px; color: #56698c; margin-top: 4px; }
      .kpi-accent { border-color: rgba(232,163,61,0.55); background: linear-gradient(180deg, rgba(232,163,61,0.10), rgba(232,163,61,0.02)); }
      .kpi-accent .kpi-label { color: #E8A33D; font-weight: 600; }
      .kpi-hero { font-size: 26px; color: #E8A33D; letter-spacing: 0.5px; }
      .rate-table td, .rate-table th { white-space: nowrap; }
      .rate-col { width: 34%; min-width: 180px; }
      .rate-bar { display: flex; align-items: center; gap: 10px; cursor: default; }
      .rate-bar-track { flex: 1; height: 6px; border-radius: 3px; background: #22344d; overflow: hidden; }
      .rate-bar-fill { display: block; height: 100%; border-radius: 3px; background: #E8A33D; }
      .rate-bar-value { font-family: 'IBM Plex Mono', monospace; font-size: 12px; font-weight: 600; color: #F4F1E8; min-width: 46px; text-align: right; }
      .rate-cell { color: #F4F1E8; font-weight: 600; }
      .rate-missing { margin-top: 0; margin-bottom: 14px; }

      /* 플랫폼별 수수료율 비교 */
      .panel-subtitle { font-size: 12px; font-weight: 600; color: #93A2BC; margin: 18px 0 8px; }
      .platform-swatch { display: inline-block; width: 8px; height: 8px; border-radius: 2px; margin-right: 6px; vertical-align: 1px; }
      .rate-compare-table th, .rate-compare-table td { white-space: nowrap; }
      .rate-compare-cell { min-width: 130px; }
      .rate-compare-cell.empty { color: #56698c; }
      .rate-compare-value { display: flex; align-items: center; gap: 6px; font-family: 'IBM Plex Mono', monospace; font-size: 14px; font-weight: 600; color: #F4F1E8; }
      .rate-compare-sub { font-size: 11px; color: #56698c; margin-top: 2px; }
      .rate-compare-cell.best { background: rgba(232,163,61,0.08); box-shadow: inset 2px 0 0 #E8A33D; }
      .best-chip { font-family: 'Inter', sans-serif; font-size: 10px; font-weight: 700; line-height: 1; padding: 3px 6px; border-radius: 4px; background: #E8A33D; color: #0F1B2D; }
      .rate-compare-total td { border-bottom: 1px solid #2A3F5C; }
      .rate-compare-total td:first-child { font-weight: 600; color: #93A2BC; }

      .tag-toolbar { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-bottom: 10px; }
      .seg { display: flex; gap: 6px; }
      .ghost-btn.active { color: #E8A33D; border-color: #E8A33D; background: rgba(232,163,61,0.08); }
      .ghost-btn:disabled { opacity: 0.5; cursor: not-allowed; }
      .tag-search { flex: 1; min-width: 220px; }
      .table-scroll { overflow-x: auto; }
      .tag-name { min-width: 220px; max-width: 420px; line-height: 1.45; }
      .tag-sub { font-size: 11px; color: #56698c; margin-top: 3px; }
      .kind-badge { margin-bottom: 0; white-space: nowrap; }
      .num-cell { font-family: 'IBM Plex Mono', monospace; white-space: nowrap; color: #93A2BC; }
      .tag-controls { display: flex; gap: 6px; align-items: center; }
      .tag-controls .select { padding: 7px 8px; }
      .tag-controls .text-input { min-width: 0; width: 130px; padding: 7px 9px; }
      .primary-btn.small { padding: 7px 12px; font-size: 12px; white-space: nowrap; }
      .tag-more { margin: 12px auto 0; }
      .reco-name { display: flex; align-items: center; gap: 6px; font-weight: 700; font-size: 14px; margin-bottom: 6px; }
      .reco-reason { font-size: 12px; color: #93A2BC; line-height: 1.5; margin-bottom: 8px; }
      .reco-idea { font-size: 12px; color: #d7dced; background: #1C324F; padding: 6px 8px; border-radius: 6px; }

      .history-list { display: flex; flex-direction: column; gap: 8px; }
      .history-item summary { cursor: pointer; font-size: 13px; color: #93A2BC; padding: 8px 0; }
      .history-item[open] summary { color: #E8A33D; }
    `}</style>
  );
}
