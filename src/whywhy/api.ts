
// src/whywhy/api.ts
export type WhyWhyResponse = {
  ok: boolean;
  judgement: string;
  issues: string[] | string;
  proposal: string;
  raw?: string;
};

const FLOW_ENDPOINT =
  "https://defaultd6eabf0ae5744f8bae6acbf6b315c6.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f1705aeeac964e8993b348cfdca2939e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=G7UI75jdDXRIIc97Jma3cEHxHBWUNLpSS8cmNSH-NAQ";

const FLOW_AUTH_KEY = ""; // 使わなければ空

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

  const text = await res.text();
  if (!res.ok) {
    throw new Error(`Flow error ${res.status}: ${text}`);
  }

  if (!text) {
    return { ok: false, judgement: "", issues: "", proposal: "", raw: "" };
  }

  try {
    const parsed = JSON.parse(text) as any;
    const body = parsed?.inputs?.body ?? parsed?.body ?? parsed;
    const base = (body?.raw && typeof body.raw === "string") ? JSON.parse(body.raw) : body;
    const normalizedIssues =
      typeof base?.issues === "string" ? (JSON.parse(base.issues) as string[]) : base?.issues;
    return {
      ok: Boolean(base?.ok ?? base?.judgement),
      judgement: base?.judgement ?? "",
      issues: normalizedIssues ?? "",
      proposal: base?.proposal ?? "",
      raw: text,
    };
  } catch {
    return { ok: false, judgement: "", issues: "", proposal: "", raw: text };
  }
}
