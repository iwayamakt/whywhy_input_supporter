// src/whywhy/types.ts
export type WhyWhyData = {
  meta: {
    product: string;
    date: string;   // UI表示用（Excel Dateでも文字列化する）
    part: string;
    owner: string;
  };
  content: {
    phenomenon: string;
    problem: string;
    why1: string;
    cause: string;
    why2: string;
    actionImmediate: string;
    why3: string;
    why4: string;
    countermeasure: string;
    why5: string;
  };
};

export type ValidationIssue = {
  field: string;          // 例: "現象", "問題", "なぜ①"
  severity: "error" | "warn";
  message: string;        // 指摘理由
  questions?: string[];   // 不足質問（最大2想定）
};
