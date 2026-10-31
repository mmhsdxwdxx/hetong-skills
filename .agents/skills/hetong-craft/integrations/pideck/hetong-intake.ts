/** Project-scoped PiDeck bridge. Questions and receipts belong to hetong_intake.py. */
import { spawn } from "node:child_process";
import { realpathSync } from "node:fs";
import { mkdtemp, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, isAbsolute, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import type { ExtensionAPI, ExtensionToolContext } from "@earendil-works/pi-coding-agent";
import { StringEnum } from "@earendil-works/pi-ai";
import { Type } from "typebox";

type JsonObject = Record<string, unknown>;
type Phase = "auto" | "categories";
type Progress = (result: { content: { type: "text"; text: string }[]; details: JsonObject }) => void;

const coreScript = resolve(dirname(realpathSync(fileURLToPath(import.meta.url))), "../../scripts/hetong_intake.py");
const knownFields = new Set(["perspective", "stage", "action", "delivery", "categories", "author"]);
const parameters = Type.Object({
  receipt_path: Type.String({ description: "本次合同受理单的绝对路径；同一合同继续使用原路径。" }),
  sources: Type.Optional(Type.Array(Type.String(), { description: "本次原件路径；复用已有受理单时可省略。" })),
  known: Type.Optional(Type.Object({
    perspective: Type.Optional(StringEnum(["buyer", "seller", "other"])),
    stage: Type.Optional(StringEnum(["unsigned", "signed", "performing", "unknown"])),
    action: Type.Optional(StringEnum(["review", "comment", "revise"])),
    delivery: Type.Optional(StringEnum(["tracked", "clean", "both"])),
    categories: Type.Optional(Type.Array(Type.String())),
    author: Type.Optional(Type.String()),
  }, { additionalProperties: false, description: "仅记录用户已明确的受理信息；取值由 core 校验，不包含商业参数。" })),
  known_evidence: Type.Optional(Type.String({ description: "known 非空时必填：用户实际指令的来源文字，不能填推断或默认值。" })),
}, { additionalProperties: false });

function isObject(value: unknown): value is JsonObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function abortError(): Error {
  const error = new Error("受理交互已中止；未自动选择任何选项。");
  error.name = "AbortError";
  return error;
}

function toolResult(value: JsonObject, receiptPath: string) {
  const details = { ...value, receipt_path: receiptPath };
  return { content: [{ type: "text" as const, text: JSON.stringify(details, null, 2) }], details };
}

/** PiDeck's successful input response omits cancelled; only its explicit submission is normalized. */
export function normalizePiDeckResponse(raw: unknown): JsonObject {
  if (raw === undefined || raw === null || (typeof raw === "string" && !raw.trim())) {
    return { cancelled: true, answers: [] };
  }
  if (typeof raw !== "string") throw new Error("PiDeck 回执不是字符串，受理单未接受该结果。");
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error("PiDeck 回执不是有效 JSON，受理单未接受该结果。");
  }
  if (!isObject(parsed)) throw new Error("PiDeck 回执格式无效，受理单未接受该结果。");
  if (["afkTimeoutMs", "followUp", "annotations"].some((key) => Object.prototype.hasOwnProperty.call(parsed, key))) {
    throw new Error("PiDeck 回执含超时、后续输入或注释状态，不能作为明确完整提交。");
  }
  if (parsed.cancelled === true) return { cancelled: true, answers: [] };
  if ((parsed.cancelled !== undefined && parsed.cancelled !== false) || !Array.isArray(parsed.answers)) {
    throw new Error("PiDeck 回执没有完整提交结果，受理单未接受该结果。");
  }
  // Drop unrelated envelope fields; core validates IDs, types, values, completeness, and plan identity.
  return { cancelled: false, answers: parsed.answers };
}

async function runCore(
  args: string[],
  options: { cwd: string; signal?: AbortSignal; progress?: Progress; receiptPath: string; allowPending?: boolean },
): Promise<JsonObject> {
  if (options.signal?.aborted) throw abortError();
  return new Promise((resolveResult, reject) => {
    const child = spawn("python3", [coreScript, ...args], {
      cwd: options.cwd,
      env: { ...process.env, PYTHONIOENCODING: "utf-8", PYTHONDONTWRITEBYTECODE: "1" },
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    let pendingLine = "";
    let aborted = false;
    let finished = false;
    let killTimer: ReturnType<typeof setTimeout> | undefined;
    const cleanup = () => {
      options.signal?.removeEventListener("abort", onAbort);
      if (killTimer !== undefined) clearTimeout(killTimer);
    };
    const onAbort = () => {
      aborted = true;
      child.kill("SIGTERM");
      killTimer = setTimeout(() => child.kill("SIGKILL"), 1000);
      killTimer.unref();
    };
    options.signal?.addEventListener("abort", onAbort, { once: true });
    if (options.signal?.aborted) onAbort();

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk: string) => {
      stdout += chunk;
      pendingLine += chunk;
      let newline: number;
      while ((newline = pendingLine.indexOf("\n")) >= 0) {
        const line = pendingLine.slice(0, newline);
        pendingLine = pendingLine.slice(newline + 1);
        try {
          const event: unknown = JSON.parse(line);
          if (isObject(event) && event.event === "form_opened" && typeof event.url === "string") {
            options.progress?.({
              content: [{ type: "text", text: `请在本机表单完成选择：[打开受理表单](${event.url})` }],
              details: { status: "pending", event: "form_opened", url: event.url, receipt_path: options.receiptPath },
            });
          }
        } catch {
          // Final JSON is pretty-printed; only complete event lines are progress updates.
        }
      }
    });
    child.stderr.on("data", (chunk: string) => { stderr = (stderr + chunk).slice(-8000); });
    child.once("error", (error) => {
      if (finished) return;
      finished = true;
      cleanup();
      reject(error);
    });
    child.once("close", (code) => {
      if (finished) return;
      finished = true;
      cleanup();
      if (aborted) { reject(abortError()); return; }
      if (code !== 0 && !(options.allowPending && code === 3)) {
        reject(new Error(`受理 core 执行失败（${code ?? "signal"}）：${stderr.trim() || "未返回结果"}`));
        return;
      }
      const resultText = stdout.split("\n").filter((line) => {
        try {
          const value: unknown = JSON.parse(line);
          return !(isObject(value) && value.event === "form_opened");
        } catch { return true; }
      }).join("\n").trim();
      try {
        const value: unknown = JSON.parse(resultText);
        if (!isObject(value)) throw new Error("invalid result");
        resolveResult(value);
      } catch {
        reject(new Error("受理 core 未返回有效 JSON 对象，未生成成功回执。"));
      }
    });
  });
}

async function executeIntake(
  toolCallId: string,
  params: { receipt_path: string; sources?: string[]; known?: JsonObject; known_evidence?: string },
  phase: Phase,
  signal: AbortSignal | undefined,
  onUpdate: Progress | undefined,
  ctx: ExtensionToolContext,
) {
  const receiptPath = params.receipt_path;
  let temporaryDir: string | undefined;
  try {
    if (!isAbsolute(receiptPath)) throw new Error("receipt_path 必须是本次任务受理单的绝对路径。");
    const known = params.known ?? {};
    if (!isObject(known) || Object.keys(known).some((key) => !knownFields.has(key))) {
      throw new Error("known 只能包含受理字段，不能包含商业参数或其他字段。");
    }
    const hasKnown = Object.keys(known).length > 0;
    if (hasKnown && !params.known_evidence?.trim()) throw new Error("known 非空时必须提供用户实际指令的 known_evidence。");
    const sources = params.sources ?? [];
    if (hasKnown && !sources.length) throw new Error("记录 known 时请同时提供 sources；省略 sources 只复用已有受理单。");
    temporaryDir = await mkdtemp(join(tmpdir(), "hetong-intake-"));
    const invoke = (args: string[], allowPending = false, activeSignal = signal) => runCore(args, {
      cwd: ctx.cwd, signal: activeSignal, progress: onUpdate, receiptPath, allowPending,
    });
    if (sources.length) {
      const args = ["init", receiptPath, ...sources.flatMap((source) => ["--source", resolve(ctx.cwd, source)])];
      if (hasKnown) {
        const knownPath = join(temporaryDir, "known.json");
        await writeFile(knownPath, JSON.stringify(known), { encoding: "utf8", mode: 0o600 });
        args.push("--known", knownPath, "--evidence", params.known_evidence!);
      }
      await invoke(args);
    } else {
      if (!(await stat(receiptPath)).isFile()) throw new Error("已有受理单不是文件。");
      await invoke(["status", receiptPath]);
    }

    const nativePiDeck = ctx.mode === "rpc" && ctx.hasUI && Object.prototype.hasOwnProperty.call(process.env, "PIDECK_AUTO_SESSION_TITLE");
    for (let round = 0; round < 3; round++) {
      const planPath = join(temporaryDir, `plan-${round}.json`);
      const plan = await invoke(["plan", receiptPath, "--host", "pideck", "--phase", phase, "--out", planPath]);
      if (!Array.isArray(plan.questions)) throw new Error("受理 core 没有返回问题数组。");
      if (!plan.questions.length) return toolResult(await invoke(["status", receiptPath]), receiptPath);
      if (!nativePiDeck) {
        const result = await invoke(["collect", receiptPath, "--phase", phase], true);
        return toolResult(result, receiptPath);
      }
      if (!isObject(plan.arguments) || !Array.isArray(plan.arguments.questions)) {
        throw new Error("受理 core 没有返回 PiDeck 问卷参数。");
      }
      const envelope = JSON.stringify({ ...plan.arguments, __piDeckBatchAsk: 1, review: true });
      onUpdate?.(toolResult({ status: "pending", interaction: "pideck", request_id: plan.request_id }, receiptPath));
      let raw: unknown;
      try {
        raw = await ctx.ui.input(envelope, "__piDeckBatchAsk__", { signal });
      } catch (error) {
        if (!signal?.aborted) throw error;
      }
      const response = signal?.aborted ? { cancelled: true, answers: [] } : normalizePiDeckResponse(raw);
      const responsePath = join(temporaryDir, `response-${round}.json`);
      await writeFile(responsePath, JSON.stringify(response), { encoding: "utf8", mode: 0o600 });
      // A genuine cancel is recorded even when the tool's AbortSignal has just fired.
      const result = await runCore([
        "accept", receiptPath, "--plan", planPath, "--response", responsePath, "--host", "pideck", "--tool-call-id", toolCallId,
      ], { cwd: ctx.cwd, signal: response.cancelled === true ? undefined : signal, receiptPath });
      if (result.status !== "pending" || phase !== "auto") return toolResult(result, receiptPath);
    }
    const result = await runCore(["status", receiptPath], { cwd: ctx.cwd, signal, receiptPath });
    return toolResult({ ...result, reason: "已达到本次受理的三轮上限；未自动选择任何选项。" }, receiptPath);
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") {
      try {
        const result = await runCore(["status", receiptPath], { cwd: ctx.cwd, receiptPath });
        return toolResult({ ...result, interaction_status: "aborted" }, receiptPath);
      } catch { /* Initialization may not have created a receipt yet. */ }
    }
    const result = toolResult({ status: "error", error: error instanceof Error ? error.message : String(error), commercial_conditions_confirmed: false }, receiptPath);
    return { ...result, isError: true };
  } finally {
    if (temporaryDir) await rm(temporaryDir, { recursive: true, force: true });
  }
}

export default function hetongIntake(pi: ExtensionAPI) {
  for (const [name, phase, description] of [
    ["hetong_intake", "auto", "收集本次合同尚未明确的立场、阶段、处理方式、交付样式与适用的预设类别。"],
    ["hetong_buyer_categories", "categories", "在购买方立场已明确后，收集本次甲方预设条款类别的多选范围。"],
    ["hetong_clause_categories", "categories", "在买卖立场已明确后，收集相应甲方或乙方预设条款类别。"],
  ] as const) {
    pi.registerTool({
      name,
      label: name === "hetong_intake" ? "合同受理" : name === "hetong_buyer_categories" ? "甲方条款类别" : "买卖条款类别",
      description,
      promptSnippet: description,
      promptGuidelines: [
        "只询问当前合同尚未明确的受理信息，用户已给出的答案应结合实际来源传入 known。",
        "类别选择不确认商业条件、不授权写入；金额、付款、期限、担保等重要商业条件必须用普通文字询问并等待明确答复。",
        "取消或未提交不构成答复；返回 pending/cancelled 时不得默认选择，也不得立即重复提问。",
      ],
      parameters,
      exposure: "model-only",
      executionMode: "sequential",
      execute: (toolCallId, params, signal, onUpdate, ctx) => executeIntake(toolCallId, params, phase, signal, onUpdate, ctx),
    });
  }
}
