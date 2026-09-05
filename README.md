# Banyan WPS Add-in

Banyan WPS Add-in 是一个用于 WPS Office 的插件，与 Zotero Banyan 插件配合使用，为用户提供在 WPS 文档中管理学术引用和参考文献的完整解决方案。

## 快速开始

### 前置条件

- WPS Office 2019 及以上版本
- Zotero 6.0 及以上版本
- Zotero Banyan 插件

### 安装

1. 下载最新版本的 Banyan WPS Add-in
2. 将插件文件复制到 WPS jsaddons 目录
3. 重启 WPS Office
4. 在 WPS 中启用 Banyan 插件

### 使用

详细使用说明请参阅 [用户手册](docs/USER_MANUAL.md)。

## 核心功能

- ✅ 插入和编辑引注（正文引注和脚注引注）
- ✅ 插入和编辑参考文献表
- ✅ 刷新引注和参考文献
- ✅ 文档定稿（将域代码转换为纯文本）
- ✅ 章节分隔符管理
- ✅ 引注窗格（可视化管理引注）
- ✅ 暗色主题支持
- ✅ 多语言支持（简体中文、英语）

## 文档

- **[用户手册](docs/USER_MANUAL.md)** - 面向最终用户的使用指南
- **[特性目标](docs/FEATURES.md)** - 必须实现的功能约束
- **[开发文档](docs/DEVELOPMENT.md)** - 架构、API 使用、调试技巧
- **[开发指南](CLAUDE.md)** - 代码约束和参考

## 开发

### 环境要求

- Node.js 16+
- npm 或 yarn

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run dev
```

在 WPS 中加载 `http://localhost:3889/index.html`

### 构建

```bash
npm run build
```

构建输出在 `release/Banyan_<version>/` 目录。
