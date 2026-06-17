import "./styles.css";

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
});
