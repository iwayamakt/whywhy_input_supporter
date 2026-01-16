// src/whywhy/validate.ts
import type { WhyWhyData, ValidationIssue } from "./types";

const hasGuessWords = (s: string) => /(たぶん|多分|と思う|かもしれない)/.test(s);
const isTooShort = (s: string, n = 20) => s.trim().length < n;
const isJustCareful = (s: string) => {
  const t = s.trim();
  if (t === "") return false;
  // 「確認する」「注意する」だけで終わる、または同義の抽象で止まるパターン
  return /^(確認(する|した)?|注意(する|した)?|気を付け(る|た)?|気をつけ(る|た)?|よく見る|見直す)$/.test(t)
    || /(注意|気を付け|確認)(だけ|するようにする|徹底)/.test(t) && t.length < 30;
};

const includesProcessTargetDeviationResult = (s: string) => {
  // 厳密判定ではなく「4要素を引き出すための雑なフラグ」
  // 工程っぽい語 / 対象っぽい語 / 逸脱っぽい語 / 結果っぽい語 がそれぞれ含まれるか
  const process = /(工程|加工|作業|段取り|セット|組立|検図|設計|出図)/.test(s);
  const target = /(ワーク|入子|図面|セット図|データ|部品|穴|コアピン|型|金型)/.test(s);
  const deviation = /(間違|誤|逆|抜け|漏れ|ズレ|未|違う|反対)/.test(s);
  const result = /(不良|仕損|作り替え|手戻り|追加工|修正|損害|再発)/.test(s);
  return { process, target, deviation, result };
};

export function validateTemplate(d: WhyWhyData): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // 現象
  const ph = d.content.phenomenon ?? "";
  if (ph.trim() === "") {
    issues.push({
      field: "現象",
      severity: "error",
      message: "未入力",
      questions: ["どの工程で発生したか？", "対象（部品/入子/図面など）は何か？"],
    });
  } else {
    const f = includesProcessTargetDeviationResult(ph);
    const missing = Object.entries(f).filter(([, v]) => !v).map(([k]) => k);
    if (missing.length >= 2) {
      issues.push({
        field: "現象",
        severity: "warn",
        message: "工程/対象/逸脱/結果の要素が不足",
        questions: ["どこで（工程/設備）？", "何がどう外れた（指示/ルールとの差）？"],
      });
    }
    if (hasGuessWords(ph)) {
      issues.push({ field: "現象", severity: "warn", message: "推測語が多い。事実と推測を分離" });
    }
    if (isTooShort(ph, 30)) {
      issues.push({ field: "現象", severity: "warn", message: "短すぎる。事実の情報量不足" });
    }
  }

  // 問題
  const pr = d.content.problem ?? "";
  if (pr.trim() === "") {
    issues.push({
      field: "問題",
      severity: "error",
      message: "未入力",
      questions: ["影響は何か（品質/納期/安全/コスト）？", "何が本質的にまずいか？"],
    });
  } else {
    // 現象の言い換え判定（簡易）
    if (pr.trim() === ph.trim()) {
      issues.push({
        field: "問題",
        severity: "warn",
        message: "現象の言い換えになっている可能性",
        questions: ["影響（品質/納期/安全/コスト）で1行定義にする", "検知できない仕組みの問題に落とす"],
      });
    }
    if (isTooShort(pr, 8)) {
      issues.push({ field: "問題", severity: "warn", message: "短すぎる。影響の観点が不足" });
    }
  }

  // なぜ①/②/③/④/⑤ + 原因
  const chain: { key: keyof WhyWhyData["content"]; label: string }[] = [
    { key: "why1", label: "なぜ①" },
    { key: "cause", label: "原因" },
    { key: "why2", label: "なぜ②" },
    { key: "why3", label: "なぜ③" },
    { key: "why4", label: "なぜ④" },
    { key: "why5", label: "なぜ⑤" },
  ];

  for (const c of chain) {
    const text = (d.content[c.key] ?? "").trim();
    if (text === "") continue; // 未入力はまず許容。後で必須化も可能
    if (isJustCareful(text)) {
      issues.push({
        field: c.label,
        severity: "warn",
        message: "抽象的（注意/確認止まり）。仕組み化へ",
        questions: ["チェック方法（いつ/誰が/何を見る）？", "手順・治具・承認・記録で再発防止に落とす？"],
      });
    }
    if (hasGuessWords(text)) {
      issues.push({ field: c.label, severity: "warn", message: "推測語が多い。根拠を明記" });
    }
  }

  // 対処 / 対策
  const im = (d.content.actionImmediate ?? "").trim();
  const cm = (d.content.countermeasure ?? "").trim();
  if (im === "") {
    issues.push({ field: "対処", severity: "warn", message: "未入力（止血対応の記録が不足）" });
  }
  if (cm === "") {
    issues.push({ field: "対策", severity: "warn", message: "未入力（再発防止が未定義）" });
  }
  if (isJustCareful(cm)) {
    issues.push({
      field: "対策",
      severity: "warn",
      message: "抽象的（注意/確認止まり）。具体化が必要",
      questions: ["手順書/チェックリスト化？", "工程設計側のルール明文化＋承認ポイント追加？"],
    });
  }

  return issues;
}
