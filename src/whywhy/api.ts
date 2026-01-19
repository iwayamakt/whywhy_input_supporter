
// src/whywhy/api.ts
export type WhyWhyResponse = {
  ok: boolean;
  judgement: string;
  issues: string[] | string;
  proposal: string;
  raw?: string;
};

const FLOW_ENDPOINT =
  (process.env.WHYWHY_FLOW_ENDPOINT as string) ||
  "<Flow の HTTP 受信 URL をここに貼る>"; // ?code=... を含む

const FLOW_AUTH_KEY =
  (process.env.WHYWHY_FLOW_KEY as string) || ""; // 使わなければ空

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
      if ([429, 502, 503, 504].includes(res.status)) {
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
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (FLOW_AUTH_KEY) headers["x-functions-key"] = FLOW_AUTH_KEY; // 使う場合のみ

  const res = await fetchWithRetry(
    FLOW_ENDPOINT,
    {
      method: "POST",
      headers,
      body: JSON.stringify(input),
    },
    3,
    60000 // 最大 60 秒
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Flow error ${res.status}: ${text}`);
  }
  return (await res.json()) as WhyWhyResponse;
}
