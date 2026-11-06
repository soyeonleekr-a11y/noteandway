#!/usr/bin/env node
// Claude Code PreToolUse(Bash) 가드.
// Claude가 Bash 명령을 실행하기 직전에 호출돼요. 막아야 하는 명령이면 exit 2 + 이유(stderr)로 돌려보내고,
// 그 이유는 Claude에게 그대로 전달돼요. 아무 문제 없으면 조용히 exit 0 해요.
//
// 막는 것 (CLAUDE.md '금지사항'과 같은 내용):
//   1. main/master 로 푸시, force 푸시(-f/--force), main 에서 바로 커밋
//   2. PR 머지(gh pr merge) — 머지는 하연이가 직접
//   3. 비밀값/개인정보 커밋: .env, CLAUDE.local.md, 엑셀·CSV, 인증서, API 키, 매출 숫자, 개인 이메일
//
// 한계: Claude Code 안에서 실행되는 명령만 검사해요. 터미널에서 직접 입력한 git 명령은 막지 못해요.
// 훅 자체가 고장 나면(파싱 오류 등) 작업을 막지 않고 통과시켜요.
import { execFileSync } from "node:child_process";

const PROTECTED = ["main", "master"];

// ───────── 입력 읽기 ─────────
let raw = "";
process.stdin.setEncoding("utf8");
for await (const chunk of process.stdin) raw += chunk;

let event;
try { event = JSON.parse(raw); } catch { process.exit(0); }
const cmd = String(event?.tool_input?.command ?? "");
const cwd = event?.cwd || process.cwd();
if (!cmd.trim()) process.exit(0);

function deny(reason) {
  process.stderr.write(
    `🚫 가드 훅이 막았어요: ${reason}\n` +
    `→ 우회하지 말고, 꼭 필요한 작업이면 하연이한테 먼저 물어봐 주세요. (규칙: CLAUDE.md '금지사항')\n`,
  );
  process.exit(2);
}

function git(args) {
  try {
    return execFileSync("git", ["-c", "core.quotepath=off", ...args], {
      cwd, encoding: "utf8", stdio: ["ignore", "pipe", "ignore"], maxBuffer: 64 * 1024 * 1024,
    });
  } catch { return ""; }
}
const currentBranch = () => git(["rev-parse", "--abbrev-ref", "HEAD"]).trim();

// ───────── 명령어를 (따옴표·heredoc 고려해서) 명령 단위로 쪼개기 ─────────
// echo "git push origin main" 처럼 따옴표 안에 든 글자나 커밋 메시지 본문은 명령으로 보지 않아요.
function tokenize(src) {
  const segs = [[]];
  const pendingHeredocs = [];
  let cur = null;
  let i = 0;
  const pushTok = () => { if (cur !== null) { segs.at(-1).push(cur); cur = null; } };
  const breakSeg = () => { pushTok(); if (segs.at(-1).length) segs.push([]); };

  while (i < src.length) {
    const c = src[i];
    if (c === "'") {
      const j = src.indexOf("'", i + 1);
      const end = j < 0 ? src.length : j;
      cur = (cur ?? "") + src.slice(i + 1, end);
      i = end + 1; continue;
    }
    if (c === '"') {
      let j = i + 1; let s = "";
      while (j < src.length && src[j] !== '"') {
        if (src[j] === "\\" && j + 1 < src.length) { s += src[j + 1]; j += 2; } else s += src[j++];
      }
      cur = (cur ?? "") + s;
      i = j + 1; continue;
    }
    if (c === "\\" && i + 1 < src.length) { cur = (cur ?? "") + src[i + 1]; i += 2; continue; }
    if (c === "<" && src[i + 1] === "<" && src[i + 2] !== "<") {
      const m = /^<<-?\s*(['"]?)([A-Za-z_]\w*)\1/.exec(src.slice(i));
      if (m) { pushTok(); pendingHeredocs.push(m[2]); i += m[0].length; continue; }
    }
    if (c === "\n") {
      i++;
      // heredoc 본문은 통째로 건너뛰어요.
      while (pendingHeredocs.length) {
        const delim = pendingHeredocs.shift();
        const re = new RegExp(`(^|\\n)[ \\t]*${delim}[ \\t]*(?=\\n|$)`);
        const m = re.exec(src.slice(i));
        i = m ? i + m.index + m[0].length : src.length;
      }
      breakSeg(); continue;
    }
    if (c === ";" || c === "|" || c === "&") {
      breakSeg();
      while (src[i] === c) i++;
      continue;
    }
    if (c === " " || c === "\t" || c === "\r") { pushTok(); i++; continue; }
    cur = (cur ?? "") + c; i++;
  }
  pushTok();
  return segs.filter((s) => s.length);
}

// 앞쪽의 FOO=bar 같은 환경변수 지정과 `command` 를 건너뛰고 실제 명령부터 돌려줘요.
function stripPrefix(tokens) {
  let k = 0;
  while (k < tokens.length && (/^[A-Za-z_]\w*=/.test(tokens[k]) || tokens[k] === "command")) k++;
  return tokens.slice(k);
}

function parseGit(tokens) {
  const base = tokens[0]?.split(/[\\/]/).pop();
  if (base !== "git" && base !== "git.exe") return null;
  let i = 1;
  while (i < tokens.length && tokens[i].startsWith("-")) i += (tokens[i] === "-C" || tokens[i] === "-c") ? 2 : 1;
  return { sub: tokens[i], args: tokens.slice(i + 1) };
}

// ───────── 검사 규칙 ─────────
function sensitiveReason(path) {
  const p = path.replace(/\\/g, "/");
  const base = p.split("/").pop();
  if (/^\.env(\..+)?$/i.test(base) && base.toLowerCase() !== ".env.example") return ".env 파일(비밀값)";
  if (base.toLowerCase() === "claude.local.md") return "CLAUDE.local.md(개인 메모·매출 숫자)";
  if (/\.(xlsx|xls|csv)$/i.test(base)) return "엑셀/CSV(정산·매출 원본 데이터)";
  if (/^(service-?account)[^/]*\.json$/i.test(base)) return "서비스 계정 키";
  if (/\.(pem|p12|pfx)$/i.test(base)) return "인증서/비밀키";
  return null;
}

const SECRET_PATTERNS = [
  [/sk-ant-[A-Za-z0-9_-]{20,}/, "Anthropic API 키"],
  [/-----BEGIN [A-Z ]*PRIVATE KEY-----/, "비밀키"],
  [/AKIA[0-9A-Z]{16}/, "AWS 액세스 키"],
];
const PRIVATE_DATA_PATTERNS = [
  [/₩\s?\d{1,3}(,\d{3}){2,}/, "매출 금액(100만 원 이상)"],
  [/\d{1,3}(,\d{3}){2,}\s?원/, "매출 금액(100만 원 이상)"],
  [/[\w.+-]+@(gmail|naver|daum|hanmail|kakao|nate|outlook|hotmail)\.(com|net)\b/i, "개인 이메일 주소"],
];

function checkPush(args) {
  const flags = args.filter((a) => a.startsWith("-"));
  const positional = args.filter((a) => !a.startsWith("-"));
  const force = flags.some((f) => f === "--force" || (/^-[a-z]+$/i.test(f) && f.includes("f")));
  if (force) deny("force 푸시(-f / --force)는 쓸 수 없어요. 기록을 덮어써서 되돌리기 어려워요.");
  if (flags.includes("--all") || flags.includes("--mirror")) deny("--all / --mirror 푸시는 main까지 같이 올려서 쓸 수 없어요.");

  const refspecs = positional.slice(1); // 첫 번째는 리모트 이름(origin)
  const dests = refspecs.map((r) => r.replace(/^\+/, "").split(":").pop().replace(/^refs\/heads\//, ""));
  if (dests.some((d) => PROTECTED.includes(d))) deny("main/master 브랜치로는 직접 푸시할 수 없어요. hayeon/작업이름 브랜치에 푸시하고 PR을 만들어주세요.");
  if (refspecs.length === 0 || dests.includes("HEAD")) {
    const branch = currentBranch();
    if (PROTECTED.includes(branch)) deny(`지금 ${branch} 브랜치에 있어서, 이 푸시는 main으로 올라가요. hayeon/작업이름 브랜치를 만들어서 작업해주세요.`);
  }
}

function checkAdd(args) {
  if (args.some((a) => a === "-f" || a === "--force")) deny("git add -f(강제 추가)는 .gitignore 보호를 무시해서 쓸 수 없어요.");
  for (const a of args.filter((x) => !x.startsWith("-"))) {
    const why = sensitiveReason(a);
    if (why) deny(`커밋하면 안 되는 파일이에요: ${a} → ${why}. (저장소가 공개예요)`);
  }
}

function checkCommit() {
  const branch = currentBranch();
  if (PROTECTED.includes(branch)) deny(`${branch} 브랜치에서 바로 커밋할 수 없어요. hayeon/작업이름 브랜치를 먼저 만들어주세요.`);

  const staged = git(["diff", "--cached", "--name-only"]).split("\n").filter(Boolean);
  for (const f of staged) {
    const why = sensitiveReason(f);
    if (why) deny(`스테이징된 파일은 커밋하면 안 돼요: ${f} → ${why}. (git restore --staged 로 빼주세요)`);
  }

  const addedByFile = new Map();
  let file = "";
  for (const line of git(["diff", "--cached", "-U0", "--no-color"]).split("\n")) {
    if (line.startsWith("+++ ")) { file = line.slice(4).replace(/^b\//, ""); continue; }
    if (line.startsWith("+") && !line.startsWith("+++")) {
      if (!addedByFile.has(file)) addedByFile.set(file, []);
      addedByFile.get(file).push(line.slice(1));
    }
  }
  for (const [f, lines] of addedByFile) {
    if (/(^|\/)package-lock\.json$/.test(f) || f === ".claude/hooks/guard.mjs") continue;
    const text = lines.join("\n");
    for (const [re, label] of SECRET_PATTERNS) if (re.test(text)) deny(`${f} 의 추가된 내용에 ${label}가 들어 있어요. 커밋하지 마세요.`);
    for (const [re, label] of PRIVATE_DATA_PATTERNS) if (re.test(text)) deny(`${f} 의 추가된 내용에 ${label}가 들어 있어요. 저장소가 공개라서 CLAUDE.local.md에만 적어주세요.`);
  }
  // 커밋 메시지도 공개되니까 같이 검사해요.
  for (const [re, label] of [...SECRET_PATTERNS, ...PRIVATE_DATA_PATTERNS]) {
    if (re.test(cmd)) deny(`커밋 메시지에 ${label}가 들어 있어요. 빼고 다시 시도해주세요.`);
  }
}

// ───────── 실행 ─────────
try {
  for (const seg of tokenize(cmd)) {
    const tokens = stripPrefix(seg);
    if (!tokens.length) continue;

    if (tokens[0] === "gh" && tokens[1] === "pr" && tokens[2] === "merge") {
      deny("PR 머지는 하연이가 확인하고 직접 해요. (gh pr merge 금지)");
    }
    const g = parseGit(tokens);
    if (!g) continue;
    if (g.sub === "push") checkPush(g.args);
    else if (g.sub === "add") checkAdd(g.args);
    else if (g.sub === "commit") checkCommit();
  }
} catch (e) {
  // 가드가 고장 나도 작업을 막지는 않아요.
  process.stderr.write(`[guard] 내부 오류로 검사를 건너뛰었어요: ${e?.message ?? e}\n`);
  process.exit(0);
}
process.exit(0);
