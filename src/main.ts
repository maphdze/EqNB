import "./styles.css";
import {
  insertEquationLine,
  insertEquationReference,
  loadEquationReferences,
  refreshAllFields,
  type EquationReference
} from "./wordService";
import type { NumberingMode } from "./ooxml";

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <section class="shell">
    <header class="topbar">
      <div>
        <p class="eyebrow">Word OMath</p>
        <h1>公式编号</h1>
      </div>
      <span class="status" id="office-status">加载中</span>
    </header>

    <section class="panel">
      <h2>编号格式</h2>
      <label class="option">
        <input type="radio" name="numbering-mode" value="plain" checked />
        <span>
          <strong>纯流水号</strong>
          <small>(1), (2), (3)</small>
        </span>
      </label>
      <label class="option">
        <input type="radio" name="numbering-mode" value="chapter" />
        <span>
          <strong>章节-流水号</strong>
          <small>(2-3), (2.3)</small>
        </span>
      </label>
      <label class="field">
        <span>分隔符</span>
        <select id="separator">
          <option value="-">连字符 -</option>
          <option value=".">句点 .</option>
          <option value=":">冒号 :</option>
        </select>
      </label>
    </section>

    <section class="actions">
      <button id="insert-equation" type="button">插入公式行</button>
      <button id="refresh-captions" type="button" class="secondary">刷新编号与引用</button>
    </section>

    <section class="panel">
      <div class="section-title">
        <h2>正文引用</h2>
        <button id="load-references" type="button" class="ghost">读取</button>
      </div>
      <select id="reference-list" size="6" aria-label="公式引用列表"></select>
      <button id="insert-reference" type="button" class="secondary full">插入选中引用</button>
    </section>

    <p class="hint" id="message">请在 Word 中打开任务窗格后使用。</p>
  </section>
`;

Office.onReady((info) => {
  const status = document.querySelector<HTMLSpanElement>("#office-status")!;
  status.textContent = info.host === Office.HostType.Word ? "Word 已连接" : "非 Word 环境";
  bindActions();
});

function getMessage(): HTMLParagraphElement {
  return document.querySelector<HTMLParagraphElement>("#message")!;
}

function setMessage(message: string, kind: "info" | "error" = "info"): void {
  const target = getMessage();
  target.textContent = message;
  target.dataset.kind = kind;
}

function getSelectedMode(): NumberingMode {
  return document.querySelector<HTMLInputElement>('input[name="numbering-mode"]:checked')!.value as NumberingMode;
}

function getSeparator(): string {
  return document.querySelector<HTMLSelectElement>("#separator")!.value;
}

function getSelectedReference(): string | null {
  const list = document.querySelector<HTMLSelectElement>("#reference-list")!;
  return list.value || null;
}

function renderReferences(references: EquationReference[]): void {
  const list = document.querySelector<HTMLSelectElement>("#reference-list")!;
  list.replaceChildren(
    ...references.map((item) => {
      const option = document.createElement("option");
      option.value = item.bookmarkName;
      option.textContent = item.label;
      return option;
    })
  );
}

async function runAction(label: string, action: () => Promise<void>): Promise<void> {
  setMessage(`${label}中...`);
  try {
    await action();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setMessage(message, "error");
  }
}

function bindActions(): void {
  document.querySelector<HTMLButtonElement>("#insert-equation")!.addEventListener("click", () => {
    void runAction("插入公式行", async () => {
      const bookmarkName = await insertEquationLine({
        mode: getSelectedMode(),
        separator: getSeparator()
      });
      setMessage(`已插入公式行，引用书签：${bookmarkName}`);
    });
  });

  document.querySelector<HTMLButtonElement>("#refresh-captions")!.addEventListener("click", () => {
    void runAction("刷新编号与引用", async () => {
      const count = await refreshAllFields();
      setMessage(`已刷新 ${count} 个域。`);
    });
  });

  document.querySelector<HTMLButtonElement>("#load-references")!.addEventListener("click", () => {
    void runAction("读取公式引用", async () => {
      const references = await loadEquationReferences();
      renderReferences(references);
      setMessage(references.length ? `已读取 ${references.length} 条公式。` : "未找到由本插件插入的公式。");
    });
  });

  document.querySelector<HTMLButtonElement>("#insert-reference")!.addEventListener("click", () => {
    void runAction("插入正文引用", async () => {
      const bookmarkName = getSelectedReference();
      if (!bookmarkName) {
        throw new Error("请先读取并选择一个公式编号。");
      }
      await insertEquationReference(bookmarkName);
      setMessage("已插入正文交叉引用。");
    });
  });
}
