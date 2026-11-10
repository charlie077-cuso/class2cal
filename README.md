# 课表转日历 · Class2Cal

把学校教务系统导出的 Excel 课表（`.xls` / `.xlsx`）转换成可以导入手机日历的 `.ics` 文件。纯前端、无后端，文件只在浏览器本地解析，不会上传。

> **在线使用**：<https://charlie077-cuso.github.io/class2cal/>（手机浏览器直接打开，支持添加到主屏幕/离线使用）。

## 功能

- **四步向导**：上传课表 → 核对课程 → 校对上课时间 → 生成 `.ics`
- **自动解析**：识别星期表头、节次列、课程名、教师、地点、周次（含 `1-16`、`5-7,9-19`、单周/双周、跳周）
- **手动校正**：解析结果可逐门编辑、删除、手动添加；支持多工作表切换
- **时间校对**：询问「学期第 1 周周一、第 1 节开始时间、每节时长、课间休息、午休」，自动推算每个时段，也可逐行修改；若文件内自带时间会优先预填
- **标准 ICS**：`VTIMEZONE` + `TZID`、每周重复（`RRULE` + `EXDATE`）、单双周/跳周、可选上课提醒、75 字节安全折行、CRLF、稳定 UID
- **手机优先**：响应式布局，iPhone / Android 浏览器可用；支持 PWA（可添加到主屏幕、离线打开）
- **隐私**：不上传文件、不依赖服务器、无追踪

## 使用步骤

1. 登录学校教务系统，在「我的课表」页面导出 Excel（`.xls` 或 `.xlsx`）。
2. 打开本工具，选择或拖入课表文件。
3. 核对解析出的课程：点课程卡片可修改名称、教师、地点、星期、节次、周次。
4. 填写学校的真实作息：学期第 1 周周一的日期、第 1 节开始时间、每节课时长、课间休息、午休时长，逐行确认时段表。
5. 选择时区、提醒方式和重复模式，点击「生成并下载 .ics」。
6. 在手机上打开下载的 `.ics` 文件，选择「日历」导入即可。

### 导入说明

- **iPhone / iPad**：Safari 下载后会自动提示用「日历」打开；若没有，去「文件」App 点一下下载的文件。
- **Android**：在通知栏或「文件管理」中打开 `.ics`，选择系统日历导入。
- **Google 日历**：网页版 → 设置 → 导入和导出 → 上传 `.ics`。

## 本地运行

无需构建步骤，classic script 直接加载：

```bash
npm install          # 仅安装开发/测试依赖（浏览器端依赖已放在 vendor/）
node tools/serve.js  # http://127.0.0.1:8123
```

直接双击 `index.html` 也能打开，但建议用本地服务器（Service Worker、文件读取行为更稳定）。

## 测试

```bash
node tests/parser.test.js   # 解析模块（真实样例 + 合成格式样例）
node tests/ics.test.js      # ICS 生成（node-ical 回读校验）
npm run test:e2e            # Playwright E2E（使用本机 Edge，含移动端 390×844）
```

E2E 说明：本机沙箱禁止子进程使用管道，`chromium.launch()` 会 `spawn EPERM`；`tests/e2e.mjs` 通过 `tools/cdp-browser.js` 手动启动 Edge 的调试端口再 `connectOverCDP`。

## 项目结构

```
index.html              页面（四步向导）
css/styles.css          样式（移动优先、深色模式）
js/parser.js            Excel 课表解析（UMD，浏览器/Node 通用）
js/ics.js               ICS 生成（UMD，浏览器/Node 通用）
js/app.js               交互逻辑
vendor/xlsx.full.min.js SheetJS（本地引入，无 CDN 依赖）
icons/                  PWA 图标
manifest.webmanifest    PWA 清单
sw.js                   Service Worker（离线缓存）
tools/serve.js          本地静态服务器
tools/gen-fixtures.js   生成合成格式测试课表
tools/cdp-browser.js    沙箱环境下的 Edge + Playwright 连接工具
tools/deploy-pages.js   一键部署到 GitHub Pages
tests/                  解析 / ICS / E2E 测试与截图
docs/CONTRACT.md        数据模型与模块契约（开发文档）
fixtures/               真实与合成样例课表
```

## 支持的课表格式

当前针对常见的「星期 × 节次」网格型课表：

- 表头行包含 `星期一`…`星期日`（也识别 `周一`、`一` 等写法）
- 首列是节次/时段（`0102节`、`第1-2节`、`1-2` 等）
- 单元格内是课程块：课程名 / 教师 / 周次 / 地点，多门课用空行分隔

不同学校格式差异很大。解析失败时可以用「手动创建课表」逐门录入，或把样例格式反馈给我们以扩展解析器。

## 部署到 GitHub Pages

需要一枚 GitHub Personal Access Token（classic，至少 `repo` 权限；若仓库已存在并公开，用 fine-grained token 需 `Contents: Read and write` 与 `Pages: Read and write`）。

```bash
# PowerShell
$env:GITHUB_TOKEN = 'ghp_xxx'
$env:GITHUB_REPOSITORY = 'your-name/class2cal'
node tools/deploy-pages.js --create   # 首次：自动创建仓库并开启 Pages
```

也可以手动：把本目录（不含 `node_modules/`、`.npm-cache/`、`.tmp/`、`tests/`、`tools/`）推到仓库默认分支，在仓库 Settings → Pages 中选择「Deploy from a branch」→ `/(root)`。

部署完成后地址形如 `https://your-name.github.io/class2cal/`，手机浏览器直接访问即可。

## 隐私

- 课表文件只在浏览器内存中解析，**不会**发送到网络。
- 没有账号、没有后端、没有统计脚本。
- 设置（作息时间、时区、提醒）保存在浏览器 `localStorage`，清空浏览器数据即删除。

## License

MIT
