// src/whywhy/excel.ts
import { mapping } from "../whywhy_mapping_fixed_firstsheet";
import type { WhyWhyData } from "./types";

const sheetIndex = mapping.workbook.sheetIndex;

const excelSerialToDateString = (n: number): string => {
  const excelEpoch = Date.UTC(1899, 11, 30);
  const ms = excelEpoch + n * 24 * 60 * 60 * 1000;
  return new Date(ms).toLocaleDateString("ja-JP");
};

const toStr = (v: unknown): string => {
  if (v == null) return "";
  // Office.jsのvaluesは Date が来ることがある
  if (v instanceof Date) return v.toLocaleDateString("ja-JP");
  if (typeof v === "number") return excelSerialToDateString(v);
  if (typeof v === "string" && /^[0-9]+(\.[0-9]+)?$/.test(v)) {
    const n = Number(v);
    if (!Number.isNaN(n)) return excelSerialToDateString(n);
  }
  return String(v);
};

export async function readTemplate(): Promise<WhyWhyData> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getFirst();

    const r = (addr: string) => ws.getRange(addr);

    const ranges = {
      product: r(mapping.meta.productName),
      date: r(mapping.meta.announceDate),
      part: r(mapping.meta.partName),
      owner: r(mapping.meta.owner),

      phenomenon: r(mapping.content.phenomenon),
      problem: r(mapping.content.problem),
      why1: r(mapping.content.why1),
      cause: r(mapping.content.cause),
      why2: r(mapping.content.why2),
      actionImmediate: r(mapping.content.actionImmediate),
      why3: r(mapping.content.why3),
      why4: r(mapping.content.why4),
      countermeasure: r(mapping.content.countermeasure),
      why5: r(mapping.content.why5),
    };

    Object.values(ranges).forEach((x) => x.load("values"));
    await ctx.sync();

    const v = (x: Excel.Range) => toStr(x.values?.[0]?.[0]);

    return {
      meta: {
        product: v(ranges.product),
        date: v(ranges.date),
        part: v(ranges.part),
        owner: v(ranges.owner),
      },
      content: {
        phenomenon: v(ranges.phenomenon),
        problem: v(ranges.problem),
        why1: v(ranges.why1),
        cause: v(ranges.cause),
        why2: v(ranges.why2),
        actionImmediate: v(ranges.actionImmediate),
        why3: v(ranges.why3),
        why4: v(ranges.why4),
        countermeasure: v(ranges.countermeasure),
        why5: v(ranges.why5),
      },
    };
  });
}

type WriteOptions = {
  onlyIfEmpty?: boolean; // trueなら、既存セルが空のときだけ上書き
};

export async function writeTemplate(patch: Partial<WhyWhyData>, opt: WriteOptions = {}): Promise<void> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getFirst();

    const setIf = async (addr: string, text?: string) => {
      if (text == null) return;

      const cell = ws.getRange(addr);
      if (opt.onlyIfEmpty) {
        cell.load("values");
        await ctx.sync();
        const current = toStr(cell.values?.[0]?.[0]).trim();
        if (current !== "") return; // 既に入っているなら書かない
      }

      cell.values = [[text]];
    };

    await setIf(mapping.meta.productName, patch.meta?.product);
    // 発表日は運用があるので、ここはUIで扱うならセット。不要ならコメントアウトでOK。
    // await setIf(mapping.meta.announceDate, patch.meta?.date);

    await setIf(mapping.meta.partName, patch.meta?.part);
    await setIf(mapping.meta.owner, patch.meta?.owner);

    await setIf(mapping.content.phenomenon, patch.content?.phenomenon);
    await setIf(mapping.content.problem, patch.content?.problem);
    await setIf(mapping.content.why1, patch.content?.why1);
    await setIf(mapping.content.cause, patch.content?.cause);
    await setIf(mapping.content.why2, patch.content?.why2);
    await setIf(mapping.content.actionImmediate, patch.content?.actionImmediate);
    await setIf(mapping.content.why3, patch.content?.why3);
    await setIf(mapping.content.why4, patch.content?.why4);
    await setIf(mapping.content.countermeasure, patch.content?.countermeasure);
    await setIf(mapping.content.why5, patch.content?.why5);

    await ctx.sync();
  });
}
``
