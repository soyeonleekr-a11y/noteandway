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
  Trophy, BarChart3, LogOut, Dog,
} from "lucide-react";
import { doc, getDoc, setDoc } from "firebase/firestore";
import { signInAnonymously } from "firebase/auth";
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

// 글 제목/상품명에서 나라·지역을 추정하기 위한 키워드 사전 (구체적인 지명부터 먼저 매칭)
const DESTINATIONS = [
  { name: "도쿄", region: "해외" }, { name: "오사카", region: "해외" }, { name: "교토", region: "해외" },
  { name: "후쿠오카", region: "해외" }, { name: "삿포로", region: "해외" }, { name: "오키나와", region: "해외" },
  { name: "일본", region: "해외" },
  { name: "다낭", region: "해외" }, { name: "하노이", region: "해외" }, { name: "호치민", region: "해외" },
  { name: "푸꾸옥", region: "해외" }, { name: "나트랑", region: "해외" }, { name: "베트남", region: "해외" },
  { name: "방콕", region: "해외" }, { name: "푸켓", region: "해외" }, { name: "치앙마이", region: "해외" },
  { name: "파타야", region: "해외" }, { name: "태국", region: "해외" },
  { name: "타이베이", region: "해외" }, { name: "대만", region: "해외" },
  { name: "세부", region: "해외" }, { name: "보라카이", region: "해외" }, { name: "마닐라", region: "해외" },
  { name: "필리핀", region: "해외" },
  { name: "싱가포르", region: "해외" },
  { name: "홍콩", region: "해외" }, { name: "마카오", region: "해외" },
  { name: "상하이", region: "해외" }, { name: "베이징", region: "해외" }, { name: "중국", region: "해외" },
  { name: "괌", region: "해외" }, { name: "사이판", region: "해외" },
  { name: "몰디브", region: "해외" },
  { name: "발리", region: "해외" }, { name: "인도네시아", region: "해외" },
  { name: "파리", region: "해외" }, { name: "프랑스", region: "해외" },
  { name: "로마", region: "해외" }, { name: "밀라노", region: "해외" }, { name: "이탈리아", region: "해외" },
  { name: "바르셀로나", region: "해외" }, { name: "스페인", region: "해외" },
  { name: "스위스", region: "해외" }, { name: "독일", region: "해외" },
  { name: "런던", region: "해외" }, { name: "영국", region: "해외" },
  { name: "프라하", region: "해외" }, { name: "체코", region: "해외" },
  { name: "크로아티아", region: "해외" }, { name: "산토리니", region: "해외" }, { name: "그리스", region: "해외" },
  { name: "튀르키예", region: "해외" }, { name: "터키", region: "해외" },
  { name: "하와이", region: "해외" }, { name: "뉴욕", region: "해외" }, { name: "미국", region: "해외" },
  { name: "캐나다", region: "해외" },
  { name: "시드니", region: "해외" }, { name: "호주", region: "해외" }, { name: "뉴질랜드", region: "해외" },
  { name: "몽골", region: "해외" }, { name: "두바이", region: "해외" },
  { name: "라오스", region: "해외" }, { name: "캄보디아", region: "해외" }, { name: "미얀마", region: "해외" },
  { name: "제주", region: "국내" }, { name: "부산", region: "국내" }, { name: "강릉", region: "국내" },
  { name: "여수", region: "국내" }, { name: "전주", region: "국내" }, { name: "경주", region: "국내" },
  { name: "속초", region: "국내" }, { name: "남해", region: "국내" }, { name: "통영", region: "국내" },
  { name: "거제", region: "국내" }, { name: "가평", region: "국내" }, { name: "양양", region: "국내" },
  { name: "포항", region: "국내" }, { name: "담양", region: "국내" }, { name: "순천", region: "국내" },
  { name: "목포", region: "국내" }, { name: "안동", region: "국내" }, { name: "보성", region: "국내" },
  { name: "인천", region: "국내" }, { name: "서울", region: "국내" }, { name: "강원", region: "국내" },
];
const PRODUCT_TYPES = [
  "항공권", "호텔", "풀빌라", "펜션", "숙소", "렌터카", "투어", "액티비티", "티켓",
  "입장권", "스파", "크루즈", "유심", "esim", "와이파이", "여행자보험", "골프",
];
function matchDestination(text) {
  if (!text) return null;
  const found = DESTINATIONS.find((d) => text.includes(d.name));
  return found || null;
}
function matchProductType(text) {
  if (!text) return null;
  const found = PRODUCT_TYPES.find((p) => text.toLowerCase().includes(p.toLowerCase()));
  return found || null;
}
// 예전 방식(엑셀 한 줄 = 상품 하나, revenue 필드만 존재)으로 저장된 데이터도
// 새 형태({name, qty, unitRevenue, totalRevenue})로 안전하게 변환해서 읽어요.
function normalizeProduct(prod) {
  const totalRevenue = prod.totalRevenue != null ? prod.totalRevenue : (prod.revenue || 0);
  const qty = prod.qty != null ? prod.qty : 1;
  const unitRevenue = prod.unitRevenue != null ? prod.unitRevenue : (qty ? totalRevenue / qty : totalRevenue);
  return { name: prod.name, qty, unitRevenue, totalRevenue };
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

const DOG_BREEDS = [
  { id: "spotted", label: "바둑이" },
  { id: "welsh", label: "웰시" },
  { id: "retriever", label: "리트리버" },
  { id: "pomeranian", label: "포메라니안" },
];

function PatternGate({ onUnlock }) {
  const [count, setCount] = useState(0);

  function handleClick(id) {
    if (id === "spotted") {
      const next = count + 1;
      setCount(next);
      if (next >= 3) onUnlock();
    } else {
      setCount(0);
    }
  }

  return (
    <div className="app-shell login-shell">
      <Style />
      <div className="login-card gate-card">
        <div className="login-title">TRAVELOG DASHBOARD</div>
        <div className="login-sub">마음에 드는 강아지를 골라보세요 🐾</div>
        <div className="dog-row">
          {DOG_BREEDS.map((b) => (
            <button key={b.id} className="dog-btn" onClick={() => handleClick(b.id)} type="button">
              <span className={`dog-icon dog-${b.id}`}><Dog size={26} /></span>
              <span className="dog-label">{b.label}</span>
            </button>
          ))}
        </div>
        <div className="gate-dots">
          {[0, 1, 2].map((i) => <span key={i} className={`gate-dot ${i < count ? "filled" : ""}`} />)}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [unlocked, setUnlocked] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [activeTab, setActiveTab] = useState("overview");
  const [platformData, setPlatformData] = useState({});
  const [postsData, setPostsData] = useState({});
  const [postRecords, setPostRecords] = useState({});
  const [blogNames, setBlogNames] = useState(["해외여행 블로그", "국내여행 블로그"]);
  const [consultingHistory, setConsultingHistory] = useState([]);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    if (!unlocked) return;
    (async () => {
      try { await signInAnonymously(auth); } catch (e) { console.error("[firebase auth] 익명 로그인 실패", e); }
      const [pd, bd, pr, bn, ch] = await Promise.all([
        safeGet("platform-data"), safeGet("posts-data"), safeGet("post-records"),
        safeGet("blog-names"), safeGet("consulting-history"),
      ]);
      if (pd) setPlatformData(pd);
      if (bd) setPostsData(bd);
      if (pr) setPostRecords(pr);
      if (bn) setBlogNames(bn);
      if (ch) setConsultingHistory(ch);
      setLoaded(true);
    })();
  }, [unlocked]);

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

  const months = useMemo(() => monthList(12), []);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth());

  const monthTotal = (m) => PLATFORMS.reduce((s, p) => s + (platformData[p]?.[m]?.total || 0), 0);
  const monthPosts = (m) => blogNames.reduce((s, b) => s + (postsData[b]?.[m]?.posts || 0), 0);
  const monthViews = (m) => blogNames.reduce((s, b) => s + (postsData[b]?.[m]?.views || 0), 0);

  const trendData = useMemo(() => {
    return [...months].reverse().map((m) => {
      const row = { month: formatMonthLabel(m).replace("년 ", ".").replace("월", "") };
      PLATFORMS.forEach((p) => { row[p] = platformData[p]?.[m]?.total || 0; });
      row["총합"] = monthTotal(m);
      return row;
    });
  }, [months, platformData, blogNames, postsData]);

  const pieData = PLATFORMS.map((p) => ({ name: p, value: platformData[p]?.[selectedMonth]?.total || 0 }));

  if (!unlocked) {
    return <PatternGate onUnlock={() => setUnlocked(true)} />;
  }

  if (!loaded) {
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
          <button className="logout-btn" onClick={() => setUnlocked(false)} title="잠그기"><LogOut size={14} /></button>
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
              <Tooltip contentStyle={{ background: "#16273F", border: "1px solid #2A3F5C", color: "#F4F1E8" }} formatter={(v) => formatWon(v)} />
              <Legend wrapperStyle={{ fontSize: 12, color: "#93A2BC" }} />
              {PLATFORMS.map((p) => (
                <Line key={p} type="monotone" dataKey={p} stroke={PLATFORM_COLOR[p]} strokeWidth={2} dot={false} />
              ))}
            </LineChart>
          </ResponsiveContainer>
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
      const nameKey = findKey(rows[0], ["상품", "품목", "product", "Product"]);
      const revKey = findKey(rows[0], ["수익", "매출", "금액", "revenue", "Revenue"]);
      const qtyKey = findKey(rows[0], ["수량", "판매", "qty", "Qty"]);
      if (!nameKey || !revKey) throw new Error("컬럼을 찾을 수 없어요 (상품명 / 수익금 열이 필요해요)");

      // 엑셀은 보통 한 줄 = 거래 1건이라 수량 열이 따로 없을 수 있어요.
      // 그래서 같은 상품명은 한 줄로 합치고, 수량 열이 없으면 등장 횟수를 판매수량으로 계산해요.
      // (공백/특수문자 차이로 다른 문자열처럼 보이는 경우까지 대비해 정규화한 키로 합쳐요)
      const grouped = {};
      rows.forEach((r) => {
        const name = String(r[nameKey]).trim();
        if (!name) return;
        const key = normalizeNameKey(name);
        const revenue = parseNumber(r[revKey]);
        const qty = qtyKey ? parseNumber(r[qtyKey]) || 1 : 1;
        if (!grouped[key]) grouped[key] = { name, qty: 0, totalRevenue: 0 };
        grouped[key].qty += qty;
        grouped[key].totalRevenue += revenue;
      });
      const products = Object.values(grouped).map((p) => ({
        ...p,
        unitRevenue: p.qty ? p.totalRevenue / p.qty : p.totalRevenue,
      }));
      if (!products.length) throw new Error("읽을 수 있는 상품 데이터가 없어요");
      const total = products.reduce((s, p) => s + p.totalRevenue, 0);

      const next = { ...platformData, [platform]: { ...(platformData[platform] || {}), [month]: { products, total } } };
      await updatePlatformData(next);
      showToast(`${platform} · ${formatMonthLabel(month)} 매출 상품 ${products.length}종 반영 완료`);
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
    PLATFORMS.forEach((p) => { row[p] = platformData[p]?.[m]?.total || 0; });
    return row;
  });

  const sortedProducts = useMemo(() => {
    if (!current) return [];
    // 저장된 데이터에 같은 상품명이 여러 줄로 남아있을 수 있어서(예전 업로드 등),
    // 화면에 보여줄 때도 한 번 더 정규화한 이름 기준으로 합쳐서 표시해요.
    const grouped = {};
    current.products.forEach((raw) => {
      const p = normalizeProduct(raw);
      const key = normalizeNameKey(p.name);
      if (!grouped[key]) grouped[key] = { name: p.name, qty: 0, totalRevenue: 0 };
      grouped[key].qty += p.qty;
      grouped[key].totalRevenue += p.totalRevenue;
    });
    const arr = Object.values(grouped).map((p) => ({
      ...p,
      unitRevenue: p.qty ? p.totalRevenue / p.qty : p.totalRevenue,
    }));
    arr.sort((a, b) => (sortDir === "desc" ? b[sortField] - a[sortField] : a[sortField] - b[sortField]));
    return arr;
  }, [current, sortField, sortDir]);

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
        <div className="hint">헤더에 <b>상품명</b>과 <b>수익금(매출)</b> 열이 있으면 자동 인식해요. 같은 상품명은 자동으로 합쳐지고, 수량 열이 없으면 거래 건수를 판매수량으로 계산해요.</div>

        {current ? (
          <div className="rev-summary">
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
              </select>
            </div>
            <table className="table">
              <thead><tr><th>순위</th><th>상품명</th><th>판매수량</th><th>개별 수익금</th><th>총 수익금</th></tr></thead>
              <tbody>
                {visibleProducts.map((p, i) => (
                  <tr key={`${p.name}-${i}`} className={i === 0 ? "top-row" : ""}>
                    <td>{i + 1}</td>
                    <td>{p.name}</td>
                    <td>{p.qty}</td>
                    <td>{formatWon(p.unitRevenue)}</td>
                    <td>{formatWon(p.totalRevenue)}</td>
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
            <Tooltip contentStyle={{ background: "#16273F", border: "1px solid #2A3F5C", color: "#F4F1E8" }} formatter={(v) => formatWon(v)} />
            <Legend wrapperStyle={{ fontSize: 12, color: "#93A2BC" }} />
            {PLATFORMS.map((p) => <Bar key={p} dataKey={p} fill={PLATFORM_COLOR[p]} radius={[3, 3, 0, 0]} />)}
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function PostsTab({ months, postsData, updatePostsData, postRecords, updatePostRecords, blogNames, updateBlogNames, showToast }) {
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

  const records = postRecords[blog]?.[month] || [];
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
      const response = await fetch("/.netlify/functions/consult", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
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
      const text = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
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

function InsightTab({ months, platformData, postRecords, blogNames }) {
  const [scope, setScope] = useState("all");

  // 같은 상품(플랫폼+상품명)이 여러 달에 걸쳐 등장해도 하나로 합쳐서 중복 없이 집계해요.
  const allProducts = useMemo(() => {
    const map = {};
    PLATFORMS.forEach((p) => {
      const monthsObj = platformData[p] || {};
      Object.entries(monthsObj).forEach(([m, data]) => {
        if (scope !== "all" && m !== scope) return;
        (data.products || []).forEach((raw) => {
          const prod = normalizeProduct(raw);
          const key = `${p}::${normalizeNameKey(prod.name)}`;
          if (!map[key]) {
            const dest = matchDestination(prod.name);
            map[key] = {
              name: prod.name, platform: p, qty: 0, totalRevenue: 0,
              region: dest ? dest.region : "기타", destination: dest ? dest.name : "기타",
            };
          }
          map[key].qty += prod.qty;
          map[key].totalRevenue += prod.totalRevenue;
        });
      });
    });
    return Object.values(map);
  }, [platformData, scope]);

  const allPosts = useMemo(() => {
    const list = [];
    blogNames.forEach((b) => {
      const monthsObj = postRecords[b] || {};
      Object.entries(monthsObj).forEach(([m, records]) => {
        if (scope !== "all" && m !== scope) return;
        records.forEach((r) => list.push({ ...r, blog: b, month: m }));
      });
    });
    return list;
  }, [postRecords, blogNames, scope]);

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
                      <td>{p.name}</td>
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
                      <td>{p.name}</td>
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
        </>
      )}
    </div>
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
      .login-sub { font-size: 12px; color: #93A2BC; text-align: center; margin-bottom: 8px; }
      .login-card .text-input { width: 100%; min-width: 0; }
      .dog-row { display: flex; gap: 16px; margin-top: 4px; }
      .dog-btn {
        display: flex; flex-direction: column; align-items: center; gap: 6px;
        background: none; border: none; cursor: pointer; padding: 0;
      }
      .dog-icon {
        width: 56px; height: 56px; border-radius: 50%; display: flex; align-items: center; justify-content: center;
        transition: transform .12s ease;
      }
      .dog-btn:active .dog-icon { transform: scale(0.9); }
      .dog-label { font-size: 11px; color: #93A2BC; }
      .dog-spotted {
        background: radial-gradient(circle at 32% 34%, #14202f 11%, transparent 12%),
          radial-gradient(circle at 68% 60%, #14202f 14%, transparent 15%),
          radial-gradient(circle at 76% 26%, #14202f 8%, transparent 9%), #F4F1E8;
        color: #14202f;
      }
      .dog-welsh { background: linear-gradient(180deg, #E8A33D 55%, #F4F1E8 55%); color: #0F1B2D; }
      .dog-retriever { background: #c98a3e; color: #2b1a08; }
      .dog-pomeranian { background: #f0d9a8; color: #7a5a28; }
      .gate-dots { display: flex; gap: 8px; margin-top: 10px; }
      .gate-dot { width: 8px; height: 8px; border-radius: 50%; background: #2A3F5C; transition: background .15s; }
      .gate-dot.filled { background: #E8A33D; }
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
      .reco-name { display: flex; align-items: center; gap: 6px; font-weight: 700; font-size: 14px; margin-bottom: 6px; }
      .reco-reason { font-size: 12px; color: #93A2BC; line-height: 1.5; margin-bottom: 8px; }
      .reco-idea { font-size: 12px; color: #d7dced; background: #1C324F; padding: 6px 8px; border-radius: 6px; }

      .history-list { display: flex; flex-direction: column; gap: 8px; }
      .history-item summary { cursor: pointer; font-size: 13px; color: #93A2BC; padding: 8px 0; }
      .history-item[open] summary { color: #E8A33D; }
    `}</style>
  );
}
