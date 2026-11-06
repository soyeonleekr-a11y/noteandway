// Netlify 환경변수 ANTHROPIC_API_KEY 를 서버 쪽에서만 사용해 Anthropic API를 호출해요.
// 프론트엔드(App.jsx)는 이 함수(/.netlify/functions/consult)만 호출하고,
// API 키는 브라우저에 절대 노출되지 않아요.
export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "허용되지 않은 메서드예요" }) };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "ANTHROPIC_API_KEY가 설정되지 않았어요. Netlify 사이트 설정 > Environment variables에서 등록해주세요." }),
    };
  }

  let prompt;
  try {
    ({ prompt } = JSON.parse(event.body || "{}"));
  } catch (e) {
    return { statusCode: 400, body: JSON.stringify({ error: "요청 형식이 올바르지 않아요" }) };
  }
  if (!prompt) {
    return { statusCode: 400, body: JSON.stringify({ error: "prompt가 필요해요" }) };
  }

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
      return {
        statusCode: res.status,
        body: JSON.stringify({ error: data?.error?.message || "Anthropic API 호출에 실패했어요" }),
      };
    }

    return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) };
  } catch (e) {
    return { statusCode: 500, body: JSON.stringify({ error: e.message || "서버 오류가 발생했어요" }) };
  }
};
