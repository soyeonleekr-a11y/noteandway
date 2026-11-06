#!/usr/bin/env node
// Claude Code PreToolUse(Bash) 가드.
// Claude가 Bash 명령을 실행하기 직전에 호출돼요. 막아야 하는 명령이면 exit 2 + 이유(stderr)로 돌려보내고,
// 그 이유는 Claude에게 그대로 전달돼요. 아무 문제 없으면 조용히 exit 0 해요.
//
// 막는 것 (CLAUDE.md '금지사항'과 같은 내용):
//   1. main/master 로 푸시, force 푸시(-f/--force), main 에서 바로 커밋
//   2. 저장소를 망가뜨릴 수 있는 gh 명령 (저장소 삭제·설정 변경, gh api 삭제 요청, 관리자 권한 머지)
//   3. 비밀값/개인정보 커밋: .env, CLAUDE.local.md, 엑셀·CSV, 인증서, API 키, 매출 숫자, 개인 이메일
//
// 확인창을 띄우는 것:
//   - PR 머지(gh pr merge): 머지하면 바로 실제 사이트에 배포돼서, 하연이가 터미널 확인창에서 승인해야 실행돼요.
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
// 역슬래시: Bash는 따옴표 밖에선 다음 글자를 그대로, 큰따옴표 안에선 $ ` " \ 앞에서만 특수 문자로 써요.
//           PowerShell은 역슬래시가 그냥 경로 구분자예요. (C:\Program Files\...)
const IS_POWERSHELL = event?.tool_name === "PowerShell";
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
        if (!IS_POWERSHELL && src[j] === "\\" && j + 1 < src.length && '$`"\\\n'.includes(src[j + 1])) { s += src[j + 1]; j += 2; } else s += src[j++];
      }
      cur = (cur ?? "") + s;
      i = j + 1; continue;
    }
    if (!IS_POWERSHELL && c === "\\" && i + 1 < src.length) { cur = (cur ?? "") + src[i + 1]; i += 2; continue; }
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
    // ; | & 와 괄호·중괄호도 명령의 경계로 봐요. (Bash: ( … ) $( … ) { …; }  /  PowerShell: if ($?) { … })
    if (c === ";" || c === "|" || c === "&" || c === "(" || c === ")" || c === "{" || c === "}") {
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

// 앞쪽의 FOO=bar 같은 환경변수 지정, `command`, 그리고 if/then/do 같은 문법 단어를 건너뛰고 실제 명령부터 돌려줘요.
const PREFIX_WORDS = new Set(["command", "then", "else", "elif", "do", "!", "$", "time", "exec"]);
function stripPrefix(tokens) {
  let k = 0;
  while (k < tokens.length && (/^[A-Za-z_]\w*=/.test(tokens[k]) || PREFIX_WORDS.has(tokens[k]))) k++;
  return tokens.slice(k);
}

// "C:\Program Files\GitHub CLI\gh.exe" 처럼 전체 경로로 실행해도 같은 명령으로 알아봐요.
function commandName(token) {
  return String(token ?? "").split(/[\\/]/).pop().toLowerCase().replace(/\.exe$/, "");
}

function parseGit(tokens) {
  if (commandName(tokens[0]) !== "git") return null;
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

// gh 명령 검사. 확인창이 필요하면 이유 문자열을, 아니면 null 을 돌려줘요. 위험한 명령은 바로 막아요.
function checkGh(args) {
  const [group, action] = args;
  const rest = args.slice(2);
  if (group === "repo" && ["delete", "rename", "edit", "archive", "unarchive"].includes(action)) {
    deny(`gh repo ${action} 는 저장소 자체를 바꾸는 명령이라 쓸 수 없어요.`);
  }
  if (group === "release" && action === "delete") deny("gh release delete 는 쓸 수 없어요.");
  if (group === "api") {
    const i = args.findIndex((a) => a === "-X" || a === "--method");
    const method = (i >= 0 ? args[i + 1] : (args.find((a) => /^(-X|--method=)/.test(a)) || "").replace(/^(-X|--method=)/, "")).toUpperCase();
    if (["DELETE", "PUT", "PATCH"].includes(method)) deny(`gh api ${method} 요청은 저장소 설정·데이터를 바꿀 수 있어서 쓸 수 없어요.`);
  }
  if (group === "pr" && action === "merge") {
    if (rest.includes("--admin")) deny("gh pr merge --admin 은 보호 규칙을 무시하는 관리자 머지라서 쓸 수 없어요.");
    const target = rest.find((a) => !a.startsWith("-")) || "(현재 브랜치의 PR)";
    return `PR ${target} 을(를) main에 머지하려고 해요. 머지하면 바로 실제 사이트에 배포돼요. 하연이가 승인해야 진행돼요.`;
  }
  return null;
}

// ───────── 실행 ─────────
let askReason = null;
try {
  for (const seg of tokenize(cmd)) {
    const tokens = stripPrefix(seg);
    if (!tokens.length) continue;

    const name = commandName(tokens[0]);
    if (name === "gh") {
      askReason = checkGh(tokens.slice(1)) || askReason;
      continue;
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

// 막을 건 없지만 하연이 승인이 필요한 명령이면, Claude Code가 터미널에 확인창을 띄우게 해요.
if (askReason) {
  process.stdout.write(JSON.stringify({
    hookSpecificOutput: {
      hookEventName: "PreToolUse",
      permissionDecision: "ask",
      permissionDecisionReason: `🔐 ${askReason}`,
    },
  }));
}
process.exit(0);
