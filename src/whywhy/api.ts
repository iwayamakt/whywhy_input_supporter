
// src/whywhy/api.ts
export type WhyWhyResponse = {
  ok: boolean;
  judgement: string;
  issues: string[];          // できれば配列に統一
  proposal: string;
  raw?: string;             // ボットの生文字列(JSON文字列)
  rawResponse?: string;     // フローの生レスポンス（デバッグ用）
};

const FLOW_ENDPOINT =
  "https://defaultd6eabf0ae5744f8bae6acbf6b315c6.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f1705aeeac964e8993b348cfdca2939e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=G7UI75jdDXRIIc97Jma3cEHxHBWUNLpSS8cmNSH-NAQ";

// 共通：フェッチ（タイムアウト + 退避リトライ）
async function fetchWithRetry(
  input: RequestInfo,
  init: RequestInit,
  retry = 3,
  timeoutMs = 30000
): Promise<Response> {
  for (let i = 0; i <= retry; i++) {
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const res = await fetch(input, { ...init, signal: controller.signal });
      clearTimeout(id);

      // 429/5xx はリトライ
      if ([429, 502, 503].includes(res.status)) {
        if (i === retry) return res;
        await new Promise((r) => setTimeout(r, 1000 * Math.pow(2, i)));
        continue;
      }
      return res;
    } catch (e) {
      clearTimeout(id);
      if (i === retry) throw e;
      await new Promise((r) => setTimeout(r, 1000 * Math.pow(2, i)));
    }
  }
  throw new Error("unreachable");
}


export async function callFlow(input: unknown): Promise<WhyWhyResponse> {
  const res = await fetchWithRetry(
    FLOW_ENDPOINT,
    {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      body: JSON.stringify(input),
    },
    3,
    300000 // Reasoning想定なら伸ばす
  );

  const rawText = await res.text();
  if (!res.ok) throw new Error(`Flow error ${res.status}: ${rawText}`);

  if (!rawText) {
    return { ok: false, judgement: "", issues: [], proposal: "", raw: "", rawResponse: "" };
  }

  let body: any;
  try {
    body = JSON.parse(rawText);
    // Responseが入れ子で返る場合への保険
    body = body?.body ?? body?.inputs?.body ?? body;
  } catch {
    // JSONで返らないなら、そのまま返す
    return { ok: false, judgement: "", issues: [], proposal: "", raw: "", rawResponse: rawText };
  }

  const judgement = String(body?.judgement ?? "");
  const ok =
    body?.ok === true ||
    judgement.toUpperCase() === "OK";

  const issues =
    Array.isArray(body?.issues) ? body.issues.map(String)
      : (typeof body?.issues === "string" ? safeJsonArray(body.issues) : []);

  return {
    ok,
    judgement,
    issues,
    proposal: String(body?.proposal ?? ""),
    raw: typeof body?.raw === "string" ? body.raw : "",
    rawResponse: rawText,
  };
}

function safeJsonArray(s: string): string[] {
  try {
    const v = JSON.parse(s);
    return Array.isArray(v) ? v.map(String) : [];
  } catch {
    return [];
  }
}

