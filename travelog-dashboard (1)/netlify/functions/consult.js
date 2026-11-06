// 컨설팅 함수: 로그인한 "허용 계정"만 Anthropic API를 호출할 수 있어요.
//
// 동작 순서
//  1. 브라우저가 로그인 증명서(Firebase ID 토큰)를 Authorization 헤더에 실어서 보내요.
//  2. 이 함수가 그 토큰으로 Firestore의 allowedUsers/{내 이메일} 문서를 읽어봐요.
//     - 토큰이 가짜이거나 만료됐으면 Google이 거절해요 (401)
//     - 허용 목록에 없는 계정이면 firestore.rules가 거절하거나 문서가 없어요 (403)
//     - 문서를 읽을 수 있을 때만 통과예요. 확인이 안 되면 항상 "막는 쪽"으로 처리해요.
//  3. 통과한 요청만 서버에 보관된 ANTHROPIC_API_KEY로 Anthropic API를 호출해요.
//     API 키는 브라우저에 절대 노출되지 않아요.
//
// 허용 목록을 이 파일에 따로 적지 않고 Firestore 규칙을 그대로 따르기 때문에,
// Firebase 콘솔에서 allowedUsers 문서를 추가/삭제하면 이 함수에도 바로 반영돼요.

// 프로젝트 ID는 공개돼도 되는 값이에요 (src/firebase.js에도 있어요).
const PROJECT_ID = process.env.FIREBASE_PROJECT_ID || "noteandway-9ae5b";
const MAX_PROMPT_CHARS = 10000; // 허용 계정이라도 너무 긴 요청은 막아서 요금 폭탄을 방지해요.

const reply = (statusCode, body) => ({
  statusCode,
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(body),
});

// 토큰 가운데 부분(payload)에서 이메일 등을 꺼내요. 이 값만으로는 믿지 않고,
// 아래 Firestore 확인이 성공했을 때(= Google이 토큰을 검증했을 때)만 사용해요.
function decodePayload(token) {
  try {
    return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8"));
  } catch {
    return null;
  }
}

// "ok" | "unauthenticated"(토큰 문제) | "forbidden"(허용 안 된 계정) | "unavailable"(확인 서버 문제)
async function checkAccess(token) {
  const payload = decodePayload(token);
  const email = payload?.email;
  if (typeof email !== "string" || !email) return "unauthenticated";

  const url =
    `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/` +
    `allowedUsers/${encodeURIComponent(email)}?mask.fieldPaths=name`;
  let res;
  try {
    res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  } catch (e) {
    console.error("[consult] 허용 목록 확인 중 네트워크 오류", e);
    return "unavailable";
  }
  if (res.status === 200) return payload.email_verified === true ? "ok" : "forbidden";
  if (res.status === 401) return "unauthenticated";
  if (res.status === 403 || res.status === 404) return "forbidden";
  console.error("[consult] 허용 목록 확인 실패", res.status);
  return "unavailable";
}

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return reply(405, { error: "허용되지 않은 메서드예요" });
  }

  // 1) 로그인 확인 (API 키 설정 여부 같은 정보는 로그인한 사람에게만 알려줘요)
  const authHeader = event.headers?.authorization || event.headers?.Authorization || "";
  const match = /^Bearer\s+(\S+)$/i.exec(authHeader);
  if (!match) {
    return reply(401, { error: "로그인이 필요해요. 다시 로그인해주세요." });
  }
  const access = await checkAccess(match[1]);
  if (access === "unauthenticated") {
    return reply(401, { error: "로그인이 만료됐거나 올바르지 않아요. 로그아웃 후 다시 로그인해주세요." });
  }
  if (access === "forbidden") {
    return reply(403, { error: "이 계정은 컨설팅을 사용할 권한이 없어요." });
  }
  if (access !== "ok") {
    return reply(503, { error: "로그인 확인 서버에 연결하지 못했어요. 잠시 후 다시 시도해주세요." });
  }

  // 2) 서버 설정 확인
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return reply(500, { error: "ANTHROPIC_API_KEY가 설정되지 않았어요. Netlify 사이트 설정 > Environment variables에서 등록해주세요." });
  }

  // 3) 요청 내용 확인
  let prompt;
  try {
    ({ prompt } = JSON.parse(event.body || "{}"));
  } catch (e) {
    return reply(400, { error: "요청 형식이 올바르지 않아요" });
  }
  if (typeof prompt !== "string" || !prompt.trim()) {
    return reply(400, { error: "prompt가 필요해요" });
  }
  if (prompt.length > MAX_PROMPT_CHARS) {
    return reply(400, { error: `요청이 너무 길어요 (최대 ${MAX_PROMPT_CHARS.toLocaleString("ko-KR")}자)` });
  }

  // 4) Anthropic 호출
  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5-5",
        max_tokens: 4000,
        messages: [{ role: "user", content: prompt }],
        tools: [{ type: "web_search_20250305", name: "web_search" }],
      }),
    });

    const data = await res.json();
    if (!res.ok) {
      return reply(res.status, { error: data?.error?.message || "Anthropic API 호출에 실패했어요" });
    }

    return reply(200, data);
  } catch (e) {
    return reply(500, { error: e.message || "서버 오류가 발생했어요" });
  }
};
