
// src/taskpane/taskpane.ts
import { readTemplate, writeTemplate } from "../whywhy/excel";
import { validateTemplate } from "../whywhy/validate";
import type { WhyWhyData } from "../whywhy/types";

const $ = (id: string) => document.getElementById(id) as HTMLElement;

const log = (msg: string) => {
  const el = $("log") as HTMLPreElement;
  el.textContent += msg + "\n";
};

const clearLog = () => {
  ( $("log") as HTMLPreElement ).textContent = "";
};

const setIssues = (issues: ReturnType<typeof validateTemplate>) => {
  const el = $("issues");
  el.innerHTML = "";
  if (!issues.length) {
    el.innerHTML = `<div class="issue"><span class="sev-warn">OK</span> 指摘なし</div>`;
    return;
  }
  for (const it of issues) {
    const sevClass = it.severity === "error" ? "sev-error" : "sev-warn";
    const qs = (it.questions ?? []).slice(0, 2).map(q => `<div class="q">Q: ${q}</div>`).join("");
    el.innerHTML += `
      <div class="issue">
        <div><span class="${sevClass}">${it.severity.toUpperCase()}</span> [${it.field}] ${it.message}</div>
        ${qs}
      </div>
    `;
  }
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
  setIssues([]); // 一旦クリア
}

async function onValidate() {
  clearLog();
  log("品質チェック…");
  const data = getFormData();
  const issues = validateTemplate(data);
  setIssues(issues);
  log(`指摘件数: ${issues.length}`);
}

async function onWrite() {
  clearLog();
  const onlyIfEmpty = (document.getElementById("onlyIfEmpty") as HTMLInputElement).checked;
  log(`書き戻し開始…（空欄のみ上書き=${onlyIfEmpty}）`);

  const data = getFormData();
  await writeTemplate(data, { onlyIfEmpty });

  log("書き戻し完了");
}

Office.onReady(() => {
  ($("btnRead") as HTMLButtonElement).onclick = () => onRead().catch(e => log("ERROR: " + (e?.message ?? e)));
  ($("btnValidate") as HTMLButtonElement).onclick = () => onValidate().catch(e => log("ERROR: " + (e?.message ?? e)));
  ($("btnWrite") as HTMLButtonElement).onclick = () => onWrite().catch(e => log("ERROR: " + (e?.message ?? e)));
  log("準備完了");
});
