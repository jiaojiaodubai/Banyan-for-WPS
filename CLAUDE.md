# Banyan for WPS - 开发指南

## 项目定位

Banyan for WPS 是一个基于 WPSJS API 的 WPS Office 加载项，与 Banyan for Zotero 配合使用，用于在 WPS 文档中管理引注与参考文献。它本质上是运行在 Chromium 环境中的网页应用，但宿主会额外注入 WPSJS API；相关术语中，“加载项”“插件”“集成”在文档里经常混用。

面向用户的交互入口只有三类：[Ribbon 菜单](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/customize-ribbon/overview)、[侧边栏](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/TaskPane/task-pane-overview)、[对话框](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/Application/member/ShowDialog)。部分交互会与 Banyan for Zotero 后端通信，再据返回结果修改文档或插件设置。具体行为以 [docs/FEATURES.md](docs/FEATURES.md) 为准。

## 术语与数据约定

- 域（Field）：Word 术语，部分文档也称“字段”。
- 引注（Citation）：标识“此处引用了文献”的对象；脚注场景下同时包含正文中的脚注引用和页脚中的脚注文本。
- 参考文献表 / 文献列表（Bibliography）：集中列出引用文献的列表，可位于文档末尾或各章末尾。
- 文献列表题录（Bibliography Entry）：文献列表中的单项；在复合顺序引注风格中，一项可能对应多条文献，例如“1.b”。
- 条目（Item）：Zotero 中的一条文献记录，包含作者、标题、时间、来源等元数据。

### 域协议

域代码只用于标识，不影响域结果。

```plaintext
BANYAN_CITATION {uuid}
BANYAN_BIBLIOGRAPHY_TITLE
BANYAN_BIBLIOGRAPHY_ENTRY {id}
BANYAN_CHAPTER_BREAK
```

- `BANYAN_BIBLIOGRAPHY_ENTRY` 的 `id` 用于实现“点击引注跳转到对应题录”。
- `BANYAN_CHAPTER_BREAK` 用于分割章节，同时存储章节级设置。
- 所有域的 `Data` 必须存储 JSON，结构见 `src/typings/style.d.ts` 中 `RenderUnit` 及其子类型。
- 域结果由数据中的 `Token` 指定。
- `Token` 可能包含 Zotero 规定的富文本标记：`<i>`、`<b>`、`<sub>`、`<sup>`、`<span style="font-variant:small-caps;">`、`<span class="nocase">`。
- `Token.link` 的含义：`http://` / `https://` 为外部链接；`banyan://entry/{id}` 跳转到题录书签 `banyan_bib_{entryId}`。

### 配置落点

- 插件设置存放在 `Application.PluginStorage`，跨窗口共享并由 WPS 持久化。
- Banyan 后端优先使用 Zotero HTTP Server 默认端口 `23119`，连接失败时回退到调试端口 `23124`；不读取共享配置文件或 Token。

- 文档全局配置存放在文档自定义属性（`CustomDocumentProperties`）的 `BANYAN_PREF` 中，值为 JSON 字符串。
- 第一章的章节配置与全局配置一同存放在 `BANYAN_PREF`；插入章节分隔符后，各章的章节配置存放在其前一个 `BANYAN_CHAPTER_BREAK` 域的 Data 中。

## 工程约束

### 安全与稳健性

- 所有外部数据都必须验证，包括域数据、HTTP 响应和用户输入。
- 单个域失败不能影响其他域；网络错误不能导致整体崩溃。
- 定稿前必须自动备份，命名格式为 `原文件名-MM-DD HH-MM-SS.扩展名`，并采用原子替换。
- 只允许修改 `BANYAN_` 开头的域。

### 性能与兼容性

- 刷新时使用批量请求，并倒序遍历域集合。
- 超过 1 秒的操作必须使用 `withProgress`。
- 缓存 Ribbon UI 对象和参考文献表位置。
- 仅支持简体中文（`msoLanguageIDSimplifiedChinese`、`msoLanguageIDChineseSingapore`）和英语（`msoLanguageIDEnglishUS`）；其他语言回退到英语。
- 禁止修改后端接口类型文件 `src/typings/http.d.ts`、`src/typings/item.d.ts`、`src/typings/style.d.ts`；如需扩展，只能使用 TypeScript 类型扩展语法。

## 流程要求

### 修改代码前后

- 修改源代码后，必须对照 [docs/FEATURES.md](docs/FEATURES.md) 检查是否破坏既定行为。
- 若修改会破坏 `FEATURES.md` 中定义的行为，必须先获得开发者授权。
- 若功能行为确实变更，必须同步更新 [docs/FEATURES.md](docs/FEATURES.md)。
- 会话结束前必须运行 `npm run lint:check` 和 `npm run build`，且结果不得包含错误或警告。

### 调试原则

- 禁止在没有实证的情况下猜测问题原因。
- 若问题无法仅靠阅读代码确定原因，使用 `src/utils/log.ts` 输出日志到桌面进行观察。
- 遇到困惑时主动与开发者确认。
- 修复前先说明问题原因和修复手段，获得确认后再编辑代码。
- 修复后按上述检查流程验证，并清理为调试加入的代码与桌面残留日志。

### 文档要求

- 所有新增或修改的 Markdown 文档都必须通过 markdown-lint。
- 修改 [docs/FEATURES.md](docs/FEATURES.md) 前需要开发者确认；其他文档修改后也应检查是否需要同步更新 [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) 和 [docs/USER_MANUAL.md](docs/USER_MANUAL.md)。

### 常用命令

```shell
npm run dev
npm run build
npx tsc --noEmit
```

## 已知限制

以下限制来自宿主或外部工具，不能在本项目内解决：

1. HTML 中的脚本必须使用 `type=text/javascript`，不能使用 `type=module`，否则窗口打开明显变慢。
2. Ribbon 菜单图片必须使用相对路径，不能 `import`，否则图标不显示。
3. `wpsjs publish` 总会生成 `jspluginonline` 在线节点，且 macOS 暂不支持通过 publish 部署在线插件；因此本项目使用自定义 `build.js` 构建，并通过 `dev/install.js` 注册为离线插件。
4. `Application.ShowDialog` 的焦点行为受 WPS 宿主限制：部分环境下打开时不会立即获得焦点，关闭后还可能切回上一次激活的窗口；该问题已通过最小对照实验确认，插件侧只能保持 `isModal=true` 以减轻影响，无法从网页脚本层彻底修复。

## 参考

- [加载项开发流程](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/addin-overview)
- [WPSJS API](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wpsjs-api/wpsjs-api-overview)
