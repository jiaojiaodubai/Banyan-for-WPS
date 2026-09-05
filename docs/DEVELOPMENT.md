# Banyan WPS Add-in - 开发文档

本文档面向开发者，提供架构说明、API 使用指南、调试技巧等信息。

## 架构概览

### 模块划分

```text
src/
├── main.ts              # 入口模块，初始化配置和注册回调
├── moulds/              # 业务模块
│   ├── ribbon.ts        # Ribbon UI 管理
│   ├── citation.ts      # 引注管理
│   ├── bibliography.ts  # 参考文献管理
│   ├── refresh.ts       # 刷新功能
│   ├── finalize.ts      # 文档定稿
│   ├── preference.ts    # 设置管理
│   └── taskpane.ts      # 任务窗格（主模块）
├── ui/                  # UI 组件
│   ├── taskpane.ts      # 任务窗格（UI 逻辑）
│   ├── preference.ts    # 设置对话框
│   └── components/
│       └── banyan-checkbox.ts  # 自定义复选框
└── utils/               # 工具函数
    ├── field.ts         # 域操作工具
    ├── http.ts          # HTTP 通信
    ├── config.ts        # 配置管理
    ├── i10n.ts          # 国际化
  ├── log.ts           # 日志工具
  ├── operation-lock.ts # 操作锁
    ├── progress.ts      # 进度提示
    └── window.ts        # 窗口管理
```

### 数据流

```text
用户操作 → Ribbon 按钮
    ↓
业务模块（moulds/）
    ↓
工具函数（utils/）
    ↓
WPS API / HTTP API
    ↓
更新文档 / 更新 UI
```

## 核心模块详解

### 1. 入口模块 (main.ts)

**职责**：

- 加载全局配置
- 注册 Ribbon 回调函数

**初始化流程**：

1. 初始化 `Application.PluginStorage`
2. 调用 `registerRibbonCallbacks()` 注册回调

### 2. Ribbon 模块 (ribbon.ts)

**职责**：

- 管理 Ribbon UI 的 8 个按钮
- 处理按钮点击事件
- 管理主题切换

**按钮列表**：

- `btnCitation` - 插入/编辑引注
- `btnChapterBreak` - 插入章节分隔符
- `btnBibliography` - 插入/编辑参考文献表
- `btnCitationPane` - 打开引注窗格
- `btnRefresh` - 刷新
- `btnUnlink` - 定稿
- `btnSettings` - 设置
- `btnDarkTheme` - 暗色主题切换

**关键函数**：

- `OnAddinLoad(ribbonUI)` - 缓存 ribbonUI 对象
- `OnAction(control)` - 处理按钮点击
- `GetImage(control)` - 返回按钮图标路径
- `OnGetEnabled(control)` - 返回按钮启用状态
- `OnGetPressed(control)` - 返回按钮按下状态（toggleButton）

**主题切换**：

- 通过 BroadcastChannel 同步主题状态
- 主题模式持久化到 `Application.PluginStorage`（由 `utils/config.ts` 管理）
- 调用 `refreshRibbonIcons()` 刷新所有按钮图标

### 3. 引注模块 (citation.ts)

**职责**：

- 插入和编辑引注
- 创建和更新域代码
- 渲染引注内容

**关键函数**：

- `onCitationEvent()` - 处理插入/编辑引注事件
- `addIntextCitation()` / `editIntextCitation()` - 插入/编辑正文引注
- `addNoteCitation()` / `editNoteCitation()` - 插入/编辑脚注引注

**工作流程**：

1. 检查光标位置（`wps.Selection.StoryType === wdMainTextStory`）
2. 检查引用样式（`getPreference().style`）
3. 判断操作模式（添加 vs 编辑）
4. 调用后端 API（`request("citation")`）
5. 创建或更新域（`range.Fields.Add()` 或 `field.Data = ...`）
6. 渲染内容（`renderStyledField()`）：写入文本 → 应用引注 Word 样式 → 应用 TextUnit 局部样式

新增引注会先创建红色占位域（`{ INTEXT_CITATION }` 或 `{ NOTE_CITATION }`），如果后续刷新或渲染未能完成，用户能在文档中看到明显的异常反馈。

### 4. 参考文献模块 (bibliography.ts)

**职责**：

- 插入和编辑参考文献表
- 收集引注上下文
- 应用标题/条目样式
- 为条目添加跳转书签

**关键函数**：

- `onBibliographyEvent()` - 处理插入/编辑参考文献表事件
- `insertBibliography()` - 插入新参考文献表
- `editBibliographyEntry()` - 编辑参考文献条目
- `collectCitationContexts()` - 收集引注上下文

**工作流程**：

1. 检查光标位置
2. 检查引用样式
3. 收集当前章节的所有引注上下文
4. 插入红色临时占位域（`{ BIBLIOGRAPHY }`），提供插入阶段及异常残留时的视觉反馈
5. 调用后端 API（新增书目走 `request("refresh")`，编辑条目走 `request("bibliography")`）
6. 删除当前章节已有的书目域（`deleteExistingBibliography()`）
7. 插入新的参考文献表（`insertBibliography()`）
8. 使用 `renderStyledField()` 渲染：写入文本 → 应用标题/条目 Word 样式 → 应用 TextUnit 局部样式
9. 为书目条目添加书签（`addBookmarkToField()`）

### 5. 刷新模块 (refresh.ts)

**职责**：

- 刷新引注和参考文献
- 批量处理所有引注
- 更新参考文献表

**关键函数**：

- `onRefreshEvent()` - 处理刷新事件
- `refresh(range?, syncItems?)` - 刷新指定章节范围
- `refreshAll(syncItems?)` - 遍历刷新所有章节

**工作流程**：

1. `onRefreshEvent()` 读取 `refreshAll` 设置
2. 若 `refreshAll=true`，调用 `refreshAll()`；否则调用 `refresh()` 仅刷新当前章节
3. `refreshAll()` 逐章节计算更新范围（`getUpdateRange()`）并调用 `refresh()`
4. `refresh()` 收集该章节引注并批量调用后端 API（`request("refresh")`）
5. 更新每个引注的域数据和渲染结果
6. 若该章节存在书目，则刷新书目并为条目添加书签

**性能优化**：

- 批量请求而非逐个请求
- 缓存参考文献表位置
- 倒序删除域，避免索引变化

### 6. 定稿模块 (finalize.ts)

**职责**：

- 将域代码转换为纯文本
- 备份文档
- 解除域链接

**关键函数**：

- `onFinalizeEvent()` - 处理定稿事件
- `backupDocument()` - 备份文档
- `unlinkFields()` - 解除域链接

**工作流程**：

1. 显示确认对话框
2. 备份文档（`backupDocument()`）
3. 遍历所有 Banyan 域
4. 解除引注域和参考文献域链接（`field.Unlink()`）
5. 删除章节分隔符域（`field.Delete()`）
6. 显示处理统计信息

**安全措施**：

- 备份文件命名：`原文件名-MM-DD HH-MM-SS.扩展名`
- 使用原子性文件替换
- 备份失败则中止操作

### 7. 设置模块 (preference.ts)

**职责**：

- 管理全局设置和章节设置
- 管理章节分隔符
- 提供设置对话框

**关键函数**：

- `onPreferenceEvent()` - 显示设置对话框
- `onInsertChapterBreakEvent()` - 插入章节分隔符
- `getPreference()` - 获取当前章节设置
- `setPreference()` - 设置当前章节设置
- `removePreference()` - 删除文档中现有设置

**对话框行为**：

- 设置对话框通过 `Application.ShowDialog` 以 `isModal=true` 打开
- 对话框脚本为 classic script（`type="text/javascript"`），避免 WPS 对 ESM 入口的启动迟滞
- WPS 宿主在部分环境存在窗口焦点缺陷：对话框关闭后可能切回之前激活窗口

**设置层级**：

- **全局设置**：存储在文档自定义属性（`CustomDocumentProperties`）的 `BANYAN_PREF` 中（JSON 编码）
  - `syncItems`: 刷新时是否同步条目元数据
  - `refreshAll`: 是否刷新所有章节
- **章节设置**：第一章存放在 `BANYAN_PREF`；插入章节分隔符后，各章的章节设置存放在其前一个章节分隔符域的 Data 中
  - `style`: 引用样式
  - `extraSource`: 手动添加的未引用文献
  - `bibliographyTitleStyle`: 参考文献标题样式名
  - `bibliographyEntryStyle`: 参考文献条目样式名

### 8. 任务窗格模块 (taskpane.ts)

**职责**：

- 管理任务窗格的创建和显示
- 提供引注列表界面
- 支持章节导航

**关键函数**：

- `toggleTaskpaneVisibility()` - 切换任务窗格显示
- `notifyTaskpaneCitationsRefreshed()` - 通知任务窗格刷新

**UI 逻辑** (ui/taskpane.ts)：

- `loadCitations()` - 加载引注列表
- `renderCitationList()` - 渲染引注列表
- `jumpToCitation()` - 跳转到引注位置
- `showItemInLibrary()` - 在 Zotero 中定位条目

### 9. 域操作工具 (field.ts)

**职责**：

- 读写域数据
- 渲染富文本
- 应用样式
- 管理书签

**关键函数**：

- `readFieldData<T>(field)` - 从域读取 JSON 数据
- `renderField(field, tokens)` - 渲染域内容
- `renderStyledField(field, applyFieldStyle, tokens)` - 按“写入文本 → 应用字段样式 → 应用 TextUnit 样式”的顺序渲染带 Word 样式的域
- `renderRange(range, tokens)` - 渲染范围内容
- `applyStyleToField(field, styleName)` - 应用样式
- `applyCitationStyle(field)` - 应用引注样式
- `addBookmarkToField(field, bookmarkName)` - 添加书签
- `getBibliographyBookmarkName(entryId)` - 获取参考文献书签名

**Token 渲染**：

- 带 Word 样式的 Field 必须使用 `renderStyledField()`，确保 TextUnit 的颜色、粗斜体、上下标、链接等局部样式最后应用，覆盖 Word 样式默认值
- 支持粗体、斜体、上标、下标
- 支持颜色、背景色
- 支持超链接（URL 和书签）
- 支持 HTML 标签解析

**占位文本**：

- 引注占位：`{ INTEXT_CITATION }` / `{ NOTE_CITATION }`
- 参考文献表占位：`{ BIBLIOGRAPHY }`
- 占位文本统一使用红色（`FIELD_PLACEHOLDER_COLOR`），用于在插入、刷新或渲染未完成时给用户明显反馈

### 10. HTTP 通信模块 (http.ts)

**职责**：

- 与 Zotero 后端通信
- 处理请求和响应
- 错误处理

**关键函数**：

- `request<T>(endpoint, data)` - 发送 HTTP 请求

**端点**：

- `POST /banyan/citation` - 获取/编辑引注
- `POST /banyan/bibliography` - 编辑参考文献条目
- `POST /banyan/refresh` - 刷新引注与题录
- `POST /banyan/style` - 获取当前样式

**错误处理**：

- 网络错误 → 提示启动 Zotero
- `cancelled` → 静默返回 null
- 其他错误 → 显示错误信息

### 11. 配置管理模块 (config.ts)

**职责**：

- 管理插件配置
- 通过 `Application.PluginStorage` 持久化插件设置

**关键函数**：

- `loadConfig()` - 初始化配置存储
- `getConfig(key, defaultValue)` - 获取配置项
- `setConfig(key, value, options)` - 设置配置项

**配置存储**：

- `Application.PluginStorage`（由 WPS 持久化）

### 12. 国际化模块 (i10n.ts)

**职责**：

- 提供多语言支持
- 根据 WPS 语言自动选择

**关键函数**：

- `useI10n(localeData)` - 创建翻译函数

**使用示例**：

```typescript
const t = useI10n({
  [wps.Enum.msoLanguageIDSimplifiedChinese]: { key: "中文" },
  [wps.Enum.msoLanguageIDEnglishUS]: { key: "English" }
})
t("key") // 根据 WPS 语言自动选择
```

## WPS API 使用指南

### 1. 域操作

#### 创建域

```typescript
const field = range.Fields.Add(
  range,
  wps.Enum.wdFieldAddin,  // 域类型
  "BANYAN_CITATION {id}",  // 域代码
  false  // 不保留格式
)
```

#### 域数据存储

```typescript
// 存储
field.Data = JSON.stringify(data)

// 读取
const data = JSON.parse(field.Data)
```

#### 域结果操作

```typescript
// 设置显示文本
field.Result.Text = "引注文本"

// 设置格式
field.Result.Font.Bold = true
field.Result.Font.Italic = true
field.Result.Font.Superscript = true
```

#### 解除域链接

```typescript
// 保留渲染结果，删除域代码
field.Unlink()

// 删除域和内容
field.Delete()
```

### 2. 范围操作

#### 复制范围

```typescript
// 避免修改原范围
const copy = range.Duplicate
```

#### 折叠范围

```typescript
// 折叠到开始
range.Collapse(wps.Enum.wdCollapseStart)

// 折叠到结束
range.Collapse(wps.Enum.wdCollapseEnd)
```

#### 设置范围

```typescript
// 设置起止位置
range.SetRange(start, end)
```

#### 移动范围

```typescript
// 移动到下一个字符
range.Move(wps.Enum.wdCharacter, 1)

// 移动到下一个段落
range.Move(wps.Enum.wdParagraph, 1)
```

### 3. 脚注操作

#### 添加脚注

```typescript
const note = range.Footnotes.Add(range)
```

#### 访问脚注内容

```typescript
const noteRange = note.Range
const field = noteRange.Fields.Item(1)
```

### 4. 样式操作

#### 应用样式

```typescript
// 应用自定义样式
range.Style = "样式名"

// 应用内置样式
range.Style = wps.Enum.wdStyleFootnoteText
```

#### 创建样式

```typescript
const style = wps.ActiveDocument.Styles.Add(
  "样式名",
  wps.Enum.wdStyleTypeCharacter  // 字符样式
)
style.BaseStyle = wps.Enum.wdStyleDefaultParagraphFont
```

### 5. 书签操作

#### 添加书签

```typescript
wps.ActiveDocument.Bookmarks.Add("书签名", range)
```

#### 超链接到书签

```typescript
wps.ActiveDocument.Hyperlinks.Add(
  range,
  undefined,  // 地址
  "书签名",   // 子地址
  undefined,  // 提示
  undefined   // 显示文本
)
```

### 6. 文档自定义属性 (CustomDocumentProperties)

文档级设置（`syncItems`、`refreshAll` 等全局项以及第一章的章节设置）保存在文档自定义属性 `BANYAN_PREF` 中，值为 JSON 编码的字符串。相关实现见 `src/modules/preference.ts`。

#### 读取属性

```typescript
const props = wps.ActiveDocument.CustomDocumentProperties
const raw = props.Item("BANYAN_PREF").Value // JSON 字符串
const pref = JSON.parse(raw)
```

#### 写入属性

```typescript
const props = wps.ActiveDocument.CustomDocumentProperties
const value = JSON.stringify(pref)
// 属性已存在时直接更新
props.Item("BANYAN_PREF").Value = value
// 属性不存在时新增（link=false，type=4 表示文本类型）
props.Add("BANYAN_PREF", false, 4, value)
```

#### 删除属性

```typescript
wps.ActiveDocument.CustomDocumentProperties.Item("BANYAN_PREF").Delete()
```

## 调试技巧

### 1. 控制台日志

```typescript
console.log("调试信息")
console.warn("警告信息")
console.error("错误信息")
```

在 WPS 开发者工具中查看（F12）

### 2. 域代码查看

在 WPS 中按 `Alt+F9` 切换域代码显示

### 3. 网络请求

在浏览器开发者工具的 Network 标签中查看 HTTP 请求

### 4. 断点调试

在浏览器开发者工具中设置断点：

1. 打开 Sources 标签
2. 找到对应的源文件
3. 点击行号设置断点

### 5. 桌面日志

项目内置了日志模块 `src/utils/log.ts`，可将关键调试日志写入桌面文件：

- 日志文件名：`Banyan-for-WPS-debug.log`
- 适用于定位宿主行为差异、网络异常和字段遍历问题
- 结论确认后应清理临时埋点，避免长期调试噪声

## 构建和部署

### 开发模式

```bash
npm run dev
```

- 支持热更新
- 在 WPS 中加载 `http://localhost:3889/index.html`

### 构建生产版本

```bash
npm run build
```

- 输出到 `dist/` 目录
- 自动复制 `manifest.xml` 和 `ribbon.xml`
- 处理 UI HTML 文件路径
- 复制静态资源（assets）
- 设置对话框采用 classic script 产物 `dist/ui/preference.js`

### 部署

1. 将 `release/Banyan_<version>/` 目录复制到 Zotero 插件的 `chrome/content/` 目录
2. 使用 `dev/install.js` 或 `dev/install.ts` 部署到 WPS jsaddons 目录
3. 重启 WPS

## 性能优化

### 1. 批量操作

- 刷新时批量请求所有引注，而非逐个请求
- 倒序遍历域集合，避免删除时索引变化

```typescript
// 正确：倒序遍历
for (let i = fields.Count; i >= 1; i--) {
  const field = fields.Item(i)
  field.Delete()
}

// 错误：正序遍历（删除后索引变化）
for (let i = 1; i <= fields.Count; i++) {
  const field = fields.Item(i)
  field.Delete()  // 删除后，后续域的索引会变化
}
```

### 2. 缓存

- 缓存参考文献表位置，避免重复查找
- 缓存 Ribbon UI 对象，避免重复获取

```typescript
// 缓存 ribbonUI
const app = Application as WPSMergedRoot & { ribbonUI?: WPS.RibbonUi }
if (!app.ribbonUI) {
  app.ribbonUI = ribbonUI
}
```

### 3. 懒加载

- Token 验证采用部分验证（至少一个有效）而非全部验证
- 域数据按需解析，不预加载

### 4. 进度提示

使用 `withProgress` 包裹长时间操作：

```typescript
await withProgress("正在刷新引注...", async () => {
  // 长时间操作
})
```

## 常见问题

### 问题：无法连接到 Zotero

**原因**：

- Zotero 未启动
- Banyan 插件未启用
- Zotero HTTP Server 未使用标准端口 23119 或调试端口 23124

**解决**：

1. 启动 Zotero
2. 在 Zotero 中启用 Banyan 插件
3. 确认 Zotero HTTP Server 使用标准端口 23119 或调试端口 23124

### 问题：引注显示为 { BANYAN_CITATION }

**原因**：

- 刷新失败
- 网络错误
- 样式数据错误

**解决**：

1. 点击"刷新"按钮
2. 检查网络连接
3. 重新选择引用样式

### 问题：定稿失败

**原因**：

- 文档未保存
- 文件系统权限不足
- 磁盘空间不足

**解决**：

1. 保存文档
2. 检查文件权限
3. 清理磁盘空间

### 问题：主题不同步

**原因**：

- BroadcastChannel 不支持

**解决**：

1. 手动刷新窗口
2. 重新打开任务窗格

### 问题：设置对话框焦点行为异常

**现象**：

- 对话框打开后不一定立即获得窗口焦点
- 对话框关闭后可能切换到之前激活的窗口

**说明**：

- 该行为已通过最小化对照实验确认属于 WPS 宿主问题
- 当前策略是保持 `ShowDialog(..., true)` 以降低交互干扰

## 代码规范

### 命名规范

- 函数：驼峰命名法（`onCitationEvent`）
- 常量：大写下划线命名法（`THEME_SYNC_CHANNEL_NAME`）
- 类型：帕斯卡命名法（`CitationSource`）

### 类型安全

- 使用 TypeScript 严格模式
- 所有函数必须有明确的返回类型
- 使用类型守卫验证运行时数据

### 错误处理

- 对外部 I/O（HTTP、文件系统、WPS 宿主 API）应在调用层或边界层做异常处理
- 错误信息必须清晰明确
- 用户取消操作不显示错误

## FEATURES 对齐状态（基于当前代码）

当前已完成一次对齐修复，涉及：

1. 刷新按钮按 `refreshAll` 设置切换“当前章节 / 全文档”。
2. 章节分隔符插入增加“段落开头”校验。
3. 新增参考文献表路径补齐条目书签。
4. 引注与书目插入路径增加红色占位域视觉反馈。

后续对齐（2026-09-05）：

1. 文档级设置存储从 Custom XML Part 迁移为文档自定义属性 `BANYAN_PREF`（JSON 编码），相关文档已同步。
2. 保存设置（设置对话框、编辑参考文献条目）时保留章节已有的 `extraSource`，避免未显式修改时误清除手动添加的文献。

### 注释规范

- 复杂逻辑必须添加注释
- 公共 API 必须添加 JSDoc 注释
- 避免无意义的注释

## 测试

### 手动测试清单

- [ ] 插入正文引注
- [ ] 插入脚注引注
- [ ] 编辑引注
- [ ] 插入参考文献表
- [ ] 编辑参考文献条目
- [ ] 刷新（手动按钮）
- [ ] 刷新（插入/编辑引注后的自动刷新）
- [ ] 插入章节分隔符
- [ ] 定稿
- [ ] 打开引注窗格
- [ ] 切换主题
- [ ] 修改设置

### 测试环境

- WPS Office 2019
- WPS Office 2021
- Windows 10
- Windows 11

### 性能测试

- 小文档（< 10 个引注）
- 中等文档（10-50 个引注）
- 大文档（50-200 个引注）
- 超大文档（> 200 个引注）
