# Word 公式编号插件初版

这是按照 `word公式自动编号_技术路线.md` 做的 Office.js 任务窗格插件原型。

## 已实现

- 任务窗格中选择编号格式：纯流水号、章节-流水号。
- 在光标处插入一行 Word 原生 OMML 空公式。
- 编号使用 Word 域：`SEQ Equation`，章节模式额外使用 `STYLEREF 1 \s`。
- 每条公式编号建立隐藏书签，正文引用使用 `REF bookmark \h`。
- 提供读取公式列表、插入正文引用、刷新全部域按钮。

## 本地运行

```powershell
npm install
npm run dev
```

然后在 Word 桌面版中旁加载 `manifest.xml`。任务窗格地址为：

```text
https://localhost:5173/index.html
```

首次打开本地 HTTPS 地址时，浏览器或 Office 可能需要信任开发证书。

## 注意

- 章节编号依赖文档中“标题 1”已经使用 Word 自动编号样式。
- 插入公式后，编号和引用需要点击“刷新编号与引用”更新。
- 初版通过 OOXML 插入空 OMML 公式，后续还需要在真实 Word 环境里验证不同 Word 版本对公式编辑态和字段更新的兼容性。
