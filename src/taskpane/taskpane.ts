
// src/taskpane/taskpane.ts
import { readTemplate, writeTemplate } from "../whywhy/excel";
import type { WhyWhyData, ValidationIssue } from "../whywhy/types";
import { callFlow } from "../whywhy/api";

const $ = (id: string) => document.getElementById(id) as HTMLElement;

const toggleSection = (id: string, show: boolean) => {
  const el = $(id);
  el.classList.toggle("hidden", !show);
};

const setProgress = (step: "send" | "analyze" | "done" | "idle") => {
  const steps = Array.from(document.querySelectorAll(".progress .step")) as HTMLElement[];
  const order = ["send", "analyze", "done"];
  steps.forEach((el) => {
    const s = el.dataset.step ?? "";
    el.classList.remove("active", "done");
    if (step === "idle") return;
    if (s === step) el.classList.add("active");
    if (order.indexOf(s) !== -1 && order.indexOf(s) < order.indexOf(step)) {
      el.classList.add("done");
    }
  });
};

const setLoading = (show: boolean, step: "send" | "analyze" | "done") => {
  toggleSection("loading", show);
  const steps = Array.from(document.querySelectorAll(".loading-steps .lstep")) as HTMLElement[];
  const order = ["send", "analyze", "done"];
  steps.forEach((el) => {
    const s = el.dataset.step ?? "";
    el.classList.remove("active", "done");
    if (s === step) el.classList.add("active");
    if (order.indexOf(s) !== -1 && order.indexOf(s) < order.indexOf(step)) {
      el.classList.add("done");
    }
  });
};

const log = (msg: string) => {
  const el = $("log") as HTMLPreElement;
  el.textContent += msg + "\n";
};

const clearLog = () => {
  ( $("log") as HTMLPreElement ).textContent = "";
};

const setIssues = (issues: ValidationIssue[]) => {
  const el = $("issues");
  el.innerHTML = "";
  if (!issues.length) {
    el.innerHTML = `<div class="issue"><span class="sev-warn">指摘なし</span></div>`;
    toggleSection("issuesSection", true);
    return;
  }
  for (const it of issues) {
    const sevClass = it.severity === "error" ? "sev-error" : "sev-warn";
    const qs = (it.questions ?? []).slice(0, 2).map(q => `<div class="q">Q: ${q}</div>`).join("");
    el.innerHTML += `
      <div class="issue">
        <div class="issue-head">
          <span class="${sevClass}">指摘あり</span>
          <span class="issue-field">${it.field}</span>
          <span class="issue-msg">${it.message}</span>
        </div>
        ${qs}
      </div>
    `;
  }
  toggleSection("issuesSection", true);
};

const clearIssuesUI = () => {
  $("issues").innerHTML = "";
  toggleSection("issuesSection", false);
};

const setAiIssues = (issues: string[] | string, judgement?: string) => {
  const list = Array.isArray(issues) ? issues : issues ? [issues] : [];
  if (!list.length) {
    setIssues([]);
    return;
  }
  const sev = judgement === "NG" ? "error" : "warn";
  setIssues(list.map((message) => ({ field: "", severity: sev, message })));
};

const setProposal = (proposal?: string) => {
  const el = $("proposal");
  if (!proposal?.trim()) {
    el.textContent = "（提案なし）";
    toggleSection("proposalSection", true);
    return;
  }
  const normalized = proposal.replace(/\\\\n/g, "\n").replace(/\\n/g, "\n");
  el.textContent = normalized;
  toggleSection("proposalSection", true);
};

const fields = [
  { id: "phenomenon", label: "現象" },
  { id: "cause", label: "原因" },
  { id: "actionImmediate", label: "対処" },
  { id: "countermeasure", label: "対策" },
  { id: "problem", label: "問題" },
  { id: "why1", label: "なぜ①" },
  { id: "why2", label: "なぜ②" },
  { id: "why3", label: "なぜ③" },
  { id: "why4", label: "なぜ④" },
  { id: "why5", label: "なぜ⑤" },
];

const updateMissing = () => {
  const missing: string[] = [];
  for (const f of fields) {
    const el = document.getElementById(f.id) as HTMLInputElement | HTMLTextAreaElement | null;
    const val = el?.value?.trim() ?? "";
    if (val === "") missing.push(f.label);
  }
  const list = $("missingList");
  list.innerHTML = "";
  if (!missing.length) {
    list.innerHTML = `<span class="chip ok">不足なし</span>`;
    return;
  }
  for (const item of missing) {
    const target = fields.find((f) => f.label === item)?.id ?? "";
    list.innerHTML += `<button class="chip chip-btn" type="button" data-target="${target}">${item}</button>`;
  }
};

const clearProposalUI = () => {
  $("proposal").textContent = "";
  toggleSection("proposalSection", false);
};

const getFormData = (): WhyWhyData => {
  const v = (id: string) => (document.getElementById(id) as HTMLInputElement | HTMLTextAreaElement).value ?? "";
  return {
    meta: {
      product: v("metaProduct"),
      date: v("metaDate"),
      part: v("metaPart"),
      owner: v("metaOwner"),
    },
    content: {
      phenomenon: v("phenomenon"),
      problem: v("problem"),
      why1: v("why1"),
      cause: v("cause"),
      why2: v("why2"),
      actionImmediate: v("actionImmediate"),
      why3: v("why3"),
      why4: v("why4"),
      countermeasure: v("countermeasure"),
      why5: v("why5"),
    },
  };
};

const FIXED_QUESTION = `あなたは不具合解析（なぜなぜ分析）のレビューア兼、対策立案の技術コンサルタント。
以下の入力（現象・問題・原因・なぜ①〜⑤・対処（暫定）・対策（恒久））について、
矛盾／論理飛躍／情報不足／用語不一致／検証不足がないかを徹底的にレビューし、
「指摘」と「次にやること（提案）」まで実行可能な形で出力せよ。

【最重要ルール（指摘の出し方）】
- 指摘（issues）は上限を設けない。見つかる限り全て列挙する。
- 「大きい問題」だけでなく「小さな曖昧さ」「前提不足」「測定不足」「条件未定義」「用語の揺れ」も漏らさず指摘する。
- 途中でまとめに入らず、指摘が尽きるまで列挙を続ける（打ち切り禁止）。
- 少なくとも10件以上を目標に指摘し、10件未満の場合は「指摘が少ない理由（入力が十分である/情報が少なすぎる等）」を明示する。

【チェック対象（必須）】
A. 現象：
  - 発生タイミング、再現率、条件依存（温度/負荷/時間/ロット/シーケンス等）、観測方法・計測点が明確か
B. 原因：
  - 現象／なぜ①との技術因果が説明されているか（メカニズム、物理・電気的根拠、前提条件）
C. なぜ①〜⑤：
  - 各段が「一つ上の段」を直接説明しているか（因果の方向が逆転していないか、飛躍がないか）
  - 途中でプロセス要因（レビュー不足等）へ逃げて技術因果が断絶していないか
D. 対処（暫定）：
  - 現象抑制に直結し、手順・条件が具体か（恒久対策と混同していないか）
E. 対策（恒久）：
  - 最深部（なぜ⑤相当）に作用し再発防止になるか
  - 検証計画（合否基準・再現条件・範囲）が明確か
F. 整合性（横串）：
  - 現象↔原因↔なぜ①〜⑤↔対処↔対策で、用語・範囲・責任・因果のつながりが一貫しているか

【必須出力（この見出し・順序で）】
1) judgement：OK/NG（根拠を1行）
2) issues：
   - 箇条書きで上限なしに列挙
   - 各項目の先頭に必ず【現象/原因/なぜ①/なぜ②/なぜ③/なぜ④/なぜ⑤/対処/対策/整合性】のいずれかを付ける
   - 可能なら「不足している具体情報（例：波形、温度、閾値、条件、ロット差）」も併記する
3) proposal（P1/P2/P3）：
   - 箇条書きで上限なしに列挙
   - 追加確認・測定・試験・解析・是正を、優先度 P1（最優先）/P2/P3 で提示
   - 各提案に必ず「目的」「何を/どこで/どう測る」「条件」「判定指標」「結果A/Bの次アクション」を含める
4) actionPlan：
   - 1週間でできること / 1ヶ月でやること / 恒久対策（技術＋プロセス）を3層で提示
5) alternatives：
   - 競合する代替仮説を2案提示し、識別に必要な追加確認観点を列挙

【制約】
- 入力に無い事実は断定しない。推測は「仮説」と明記する。
- 用語は入力の表記を優先し、Why表記は使わず「なぜ①〜⑤」で表す。
- 不足情報が多く結論が出ない場合は、その旨を明記し「結論に必要な最小データ（Minimum Data Set）」をP1として提示する。
`;

const setFormData = (d: WhyWhyData) => {
  const set = (id: string, val: string) => ((document.getElementById(id) as HTMLInputElement | HTMLTextAreaElement).value = val ?? "");
  set("metaProduct", d.meta.product);
  set("metaDate", d.meta.date);
  set("metaPart", d.meta.part);
  set("metaOwner", d.meta.owner);

  set("phenomenon", d.content.phenomenon);
  set("problem", d.content.problem);
  set("why1", d.content.why1);
  set("cause", d.content.cause);
  set("why2", d.content.why2);
  set("actionImmediate", d.content.actionImmediate);
  set("why3", d.content.why3);
  set("why4", d.content.why4);
  set("countermeasure", d.content.countermeasure);
  set("why5", d.content.why5);
};

async function onRead() {
  clearLog();
  log("読み取り開始…");
  const data = await readTemplate();
  setFormData(data);
  log("読み取り完了");
  clearIssuesUI();
  clearProposalUI();
  updateMissing();
}

async function onValidate() {
  return onFlow();
}

async function onWrite() {
  clearLog();
  const onlyIfEmpty = (document.getElementById("onlyIfEmpty") as HTMLInputElement).checked;
  log(`書き戻し開始…（空欄のみ上書き=${onlyIfEmpty}）`);

  const data = getFormData();
  await writeTemplate(data, { onlyIfEmpty });

  log("書き戻し完了");
}

async function onFlow() {
  clearLog();
  log("AI確認 送信中…");
  setProgress("send");
  setLoading(true, "send");

  const data = getFormData();
  const question = FIXED_QUESTION;

  const payload = {
    excelData: data,
    question,
  };

  try {
    const res = await callFlow(payload);
    setProgress("analyze");
    setLoading(true, "analyze");

    log(`OK: ${res.ok}`);
    log(`判定: ${res.judgement}`);
    if (res.raw && (res.judgement == null || res.judgement === "")) {
      log("RAW:");
      log(res.raw);
    }
    if (res.issues) {
      if (Array.isArray(res.issues)) {
        log("指摘:");
        for (const it of res.issues) log(`- ${it}`);
      } else {
        log(`指摘: ${res.issues}`);
      }
    }
    if (res.proposal) log(`提案: ${res.proposal}`);
    if (res.issues) {
      setAiIssues(res.issues, res.judgement);
    } else {
      setAiIssues([], res.judgement);
    }
    setProposal(res.proposal);
    updateMissing();
    setProgress("done");
    setLoading(true, "done");
    setTimeout(() => setLoading(false, "done"), 600);
    $("issuesSection").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (e) {
    setLoading(false, "send");
    setProgress("idle");
    throw e;
  }
}

Office.onReady(() => {
  ($("btnRead") as HTMLButtonElement).onclick = () => onRead().catch(e => log("ERROR: " + (e?.message ?? e)));
  ($("btnValidate") as HTMLButtonElement).onclick = () => onValidate().catch(e => log("ERROR: " + (e?.message ?? e)));
  ($("btnWrite") as HTMLButtonElement).onclick = () => onWrite().catch(e => log("ERROR: " + (e?.message ?? e)));
  let logVisible = false;
  ($("btnToggleLog") as HTMLButtonElement).onclick = () => {
    logVisible = !logVisible;
    $("log").classList.toggle("hidden", !logVisible);
    ($("btnToggleLog") as HTMLButtonElement).textContent = logVisible ? "非表示" : "表示";
  };
  $("missingList").addEventListener("click", (e) => {
    const target = (e.target as HTMLElement)?.closest<HTMLButtonElement>("button[data-target]");
    if (!target) return;
    const id = target.dataset.target;
    if (!id) return;
    document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "center" });
    (document.getElementById(id) as HTMLInputElement | HTMLTextAreaElement | null)?.focus();
  });
  for (const f of fields) {
    const el = document.getElementById(f.id);
    el?.addEventListener("input", updateMissing);
  }
  updateMissing();
  setProgress("idle");
  log("準備完了");
});
