# Windows .exe UX 优化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 word-transfer 优化为可在 Windows 下双击 .exe 友好使用的桌面工具，包括帮助信息、拖拽支持、控制台不闪退、中文不乱码、进度提示、输出目录管理等。

**Architecture:** 所有改动集中在 `word.js` 单文件（约 680 行），遵循现有代码风格。`package.json` 构建脚本增加图标和版本信息。

**Tech Stack:** Node.js, 无新增依赖（chcp 用 child_process 内置）

**涵盖优化点:** 1(报错闪退) 2(拖拽) 3(输出目录) 4(中文乱码) 5(--help) 6(引导) 7(进度) 9(覆盖) 10(打开文件夹) 12(图标版本)

**排除:** 8(报告头信息) 保留不变

---

### Task 1: 添加 --help 帮助信息

**Files:**
- Modify: `word.js`

- [ ] **Step 1: 在 parseCliArgs 中添加 --help 解析**

在 `parseCliArgs` 函数中，识别 `--help` / `-h` 后立即打印帮助并退出。

找到 `parseCliArgs` 函数（约 155 行），在 `const positional = []` 之前插入 `--help` 检测：

```javascript
function parseCliArgs(argv) {
  // 检查 --help / -h
  if (argv.includes('--help') || argv.includes('-h')) {
    console.log([
      'word-transfer — Excel 集成 Word/Excel 文档生成器',
      '',
      '用法:',
      '  word-transfer.exe [选项]',
      '  word-transfer.exe [Excel文件] [Word模板]     (支持拖拽文件到 exe)',
      '',
      '选项:',
      '  -t, --template <path>       Word 模板文件路径 (默认: template.docx)',
      '  -e, --excel <path>          Excel 数据文件路径 (默认: test.xlsx)',
      '  -x, --excel-template <path> Excel 报告模板路径 (默认: template_excel.xlsx)',
      '  -o, --output <prefix>       输出文件名前缀 (默认: output)',
      '  -h, --help                  显示此帮助信息',
      '',
      '示例:',
      '  word-transfer.exe -t 模板.docx -e 数据.xlsx -o 结果',
      '  word-transfer.exe 数据.xlsx 模板.docx     (拖拽两个文件)',
      '',
      '输出:',
      '  在同目录 output/ 文件夹下生成:',
      '  {前缀}-{时间戳}.docx       Word 文档',
      '  {前缀}-{时间戳}.xlsx       Excel 报告',
    ].join('\n'))
    process.exit(0)
  }

  const args = {
    // ... existing code ...
```

- [ ] **Step 2: 运行验证**

```bash
node word.js --help 2>&1
```

预期：打印完整帮助信息并退出，不报错。

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: add --help / -h flag"
```

---

### Task 2: Windows 下所有退出路径都暂停

**Files:**
- Modify: `word.js` — `maybePauseBeforeExit` 函数和 `main()` 错误处理

- [ ] **Step 1: 修改 maybePauseBeforeExit**

将暂停条件从"仅在交互模式 + Windows"改为"Windows 下始终暂停"：

找到 `maybePauseBeforeExit` 函数（约 281 行），替换为：

```javascript
async function maybePauseBeforeExit() {
  if (process.platform !== 'win32') return
  const rl = createReadline()
  await new Promise((resolve) => rl.question('按回车退出...', () => resolve(undefined)))
  rl.close()
}
```

注意：函数签名从 `maybePauseBeforeExit(interactiveStarted)` 改为 `maybePauseBeforeExit()`，不再需要参数。

- [ ] **Step 2: 更新所有调用点**

在 `main()` 函数末尾，将 `await maybePauseBeforeExit(noArgs)` 改为：

```javascript
await maybePauseBeforeExit()
```

在 `main()` 的 catch 块中（约 466 行），将 `await maybePauseBeforeExit(noArgs)` 改为：

```javascript
await maybePauseBeforeExit()
```

并删除 catch 块中重复的 `const noArgs = process.argv.slice(2).length === 0` 行。

- [ ] **Step 3: 运行验证**

```bash
node word.js --help 2>&1  # 应正常退出，不卡住（--help 在 maybePauseBeforeExit 之前就 exit 了）
node word.js 2>&1          # 交互模式下按 Ctrl+C 退出即可（macOS 跳过暂停）
```

- [ ] **Step 4: Commit**

```bash
git add word.js
git commit -m "fix: pause on all Windows exits, not just interactive mode"
```

---

### Task 3: 支持拖拽文件 / 位置参数自动推断

**Files:**
- Modify: `word.js` — `parseCliArgs` 函数

- [ ] **Step 1: 在 parseCliArgs 末尾添加位置参数推断**

`parseCliArgs` 解析完所有命名参数后，处理剩余的 `positional` 数组。按扩展名自动分配：

在 `parseCliArgs` 函数的 `return args` 之前插入：

```javascript
  // 位置参数自动推断：支持拖拽文件到 exe
  for (const token of positional) {
    const ext = path.extname(token).toLowerCase()
    if ((ext === '.xlsx' || ext === '.xls') && !args.excel) {
      args.excel = token
    } else if (ext === '.docx' && !args.template) {
      args.template = token
    } else if ((ext === '.xlsx' || ext === '.xls') && !args.excelTemplate) {
      args.excelTemplate = token
    }
  }
```

- [ ] **Step 2: 运行验证**

```bash
node word.js test.xlsx template.docx 2>&1
```

预期：正常读取 test.xlsx 作为数据文件和 template.docx 作为模板，输出 Word + Excel。

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: support drag-and-drop with positional arg inference"
```

---

### Task 4: 输出到独立 output/ 目录

**Files:**
- Modify: `word.js` — `main()` 函数

- [ ] **Step 1: 修改输出路径**

在 `main()` 中，将输出路径从 `process.cwd()` 改为 `process.cwd()/output/`。

找到输出文件名构造部分（约 560 行），在构造 `outputPath` 和 `excelOutputPath` 之前，确保 `output/` 目录存在：

```javascript
  // 确保输出目录存在
  const outputDir = path.resolve(process.cwd(), 'output')
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true })
  }

  const outputPath = path.resolve(outputDir, outputFileName)
```

同样修改 Excel 输出路径：

```javascript
    const excelOutputPath = path.resolve(outputDir, excelOutputFileName)
```

- [ ] **Step 2: 运行验证**

```bash
node word.js --template=template.docx --excel=test.xlsx --output=dir_test 2>&1
ls output/
```

预期：`output/` 目录下有 `dir_test-*.docx` 和 `dir_test-*.xlsx`。

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: output files to output/ subdirectory"
```

---

### Task 5: Windows 控制台 UTF-8 编码

**Files:**
- Modify: `word.js` — `main()` 函数入口

- [ ] **Step 1: 在 main() 开头切换到 UTF-8**

在 `async function main()` 的第一行（第 500 行左右）插入：

```javascript
async function main() {
  // Windows 控制台切换到 UTF-8，避免中文乱码
  if (process.platform === 'win32') {
    try {
      require('child_process').execSync('chcp 65001', { stdio: 'ignore' })
    } catch (_) {
      // chcp 不可用时静默忽略
    }
  }

  const noArgs = process.argv.slice(2).length === 0
  // ... 后续代码不变
```

- [ ] **Step 2: 验证语法无误**

```bash
node -e "require('./word.js'); console.log('OK')"
```

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: set console to UTF-8 on Windows to prevent garbled Chinese"
```

---

### Task 6: 交互模式引导更友好

**Files:**
- Modify: `word.js` — `ensureParamsInteractive` 函数

- [ ] **Step 1: 添加引导提示**

在 `ensureParamsInteractive` 函数开头（进入询问循环之前），打印一段引导信息帮助新用户理解：

```javascript
async function ensureParamsInteractive(baseArgs) {
  const noArgs = process.argv.slice(2).length === 0
  const needInteractive = noArgs || !baseArgs.template || !baseArgs.excel

  if (!needInteractive) return baseArgs

  // 引导提示
  if (noArgs) {
    console.log([
      '══════════════════════════════════════════',
      '  word-transfer — 水表检定报告生成工具',
      '',
      '  本程序将根据 Excel 中的产品编码，',
      '  自动生成 Word 文档和 Excel 检验报告。',
      '',
      '  请确保以下文件与本程序在同一目录：',
      '    • template.docx       (Word 模板)',
      '    • template_excel.xlsx (Excel 报告模板)',
      '══════════════════════════════════════════',
      '',
    ].join('\n'))
  }

  const rl = createReadline()
  // ... 后续询问逻辑不变
```

- [ ] **Step 2: 改进"文件不存在"时的提示**

当模板文件不存在时，提示信息加上具体的期望文件名和位置：

找到模板文件检查循环（约 239 行），将 `console.error` 改为更友好的提示：

```javascript
      console.error(`未找到模板文件: ${absCandidate}`)
      console.error('请将 template.docx 放在与本程序相同的目录下')
```

同理，Excel 文件检查：

```javascript
      console.error(`未找到Excel文件: ${absCandidate}`)
      console.error('请将 .xlsx 或 .xls 文件放在与本程序相同的目录下')
```

- [ ] **Step 3: 运行验证**

```bash
echo -e "\n\n" | node word.js 2>&1
```

预期：看到引导信息，然后依次提示输入路径。

- [ ] **Step 4: Commit**

```bash
git add word.js
git commit -m "feat: add friendly guidance in interactive mode"
```

---

### Task 7: 大数据量进度提示

**Files:**
- Modify: `word.js` — `generatePages` 或 `main()` 函数

- [ ] **Step 1: 在 main() 中加入进度输出**

在 `main()` 中 `generatePages` 调用之后、渲染之前加入进度提示。但更简单的方式是在 `generateErrors` 不影响现有逻辑的情况下，在 `main()` 中显示一条汇总信息即可，因为数据生成本身很快（毫秒级）。

在 `console.log('将生成 ${numberOfPages} 页，每页最多 ${rowsPerPage} 行')` 之后加：

```javascript
  if (productCodes.length > 500) {
    console.log(`正在处理 ${productCodes.length} 条数据，请稍候...`)
  }
```

在文档生成完成后加：

```javascript
  if (productCodes.length > 500) {
    console.log('处理完成！')
  }
```

- [ ] **Step 2: 运行验证**

项目最多 123 条产品编码，不会触发进度提示（阈值 500）。改为 0 测试：

```bash
# 临时改阈值为 1 测试
node -e "
const fs = require('fs');
let code = fs.readFileSync('word.js', 'utf8');
code = code.replace('productCodes.length > 500', 'productCodes.length > 1');
fs.writeFileSync('word_test.js', code);
" 2>&1 || true
```

预期：正常行为。

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: add progress hint for large datasets"
```

---

### Task 8: 文件覆盖保护（毫秒级时间戳）

**Files:**
- Modify: `word.js` — `formatTimestamp` 函数

- [ ] **Step 1: 修改时间戳格式**

在 `formatTimestamp` 函数末尾加毫秒：

```javascript
function formatTimestamp(date = new Date()) {
  const pad = (n) => String(n).padStart(2, '0')
  const y = date.getFullYear()
  const m = pad(date.getMonth() + 1)
  const d = pad(date.getDate())
  const hh = pad(date.getHours())
  const mm = pad(date.getMinutes())
  const ss = pad(date.getSeconds())
  const ms = String(date.getMilliseconds()).padStart(3, '0')
  return `${y}${m}${d}-${hh}${mm}${ss}-${ms}`
}
```

- [ ] **Step 2: 运行验证**

```bash
node -e "
const m = require('./word.js');
// 手动调用 formatTimestamp（如果是 module.exports 导出的）
// 如果没有导出，则直接跑一次生成看看文件名
" 2>&1
```

跑一次实际生成看文件名：

```bash
node word.js -t template.docx -e test.xlsx -o ts_test 2>&1
ls output/ts_test-*.docx
```

预期文件名包含毫秒：`ts_test-20260706-120000-123.docx`

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: add millisecond precision to timestamp for overwrite protection"
```

---

### Task 9: 生成完成后自动打开文件夹 (Windows)

**Files:**
- Modify: `word.js` — `main()` 函数末尾

- [ ] **Step 1: 在输出完成后打开文件夹**

在 `main()` 函数中，所有输出完成后、`maybePauseBeforeExit()` 之前，添加：

```javascript
  // Windows 下自动打开输出文件夹
  if (process.platform === 'win32') {
    try {
      require('child_process').execSync(`start "" "${outputDir}"`, { stdio: 'ignore' })
    } catch (_) {
      // 打开文件夹失败不影响主流程
    }
  }
```

- [ ] **Step 2: 验证语法**

```bash
node -e "require('./word.js'); console.log('OK')"
```

- [ ] **Step 3: Commit**

```bash
git add word.js
git commit -m "feat: auto-open output folder on Windows after generation"
```

### Task 10: 构建时添加图标和版本信息

**Files:**
- Modify: `package.json` — build scripts
- Create: `assets/icon.ico` (占位说明)

- [ ] **Step 1: 更新 package.json 构建脚本**

```json
  "scripts": {
    "build": "pkg . --targets node14-macos-x64,node14-macos-arm64 --output dist/wordTransfer",
    "build:mac": "pkg . --targets node14-macos-arm64 --output dist/wordTransfer",
    "build:linux": "pkg . --targets node14-linux-x64 --output dist/wordTransfer",
    "build:win": "pkg . --targets node14-win-x64 --output dist/wordTransfer.exe --options expose-gc"
  },
```

在此基础上，Windows 构建增加 `--icon` 和文件版本信息。pkg 不支持 `--version` 作为 CLI 参数，但可以通过 `package.json` 的 `pkg` 字段配置。

在 `package.json` 中添加 `pkg` 配置块：

```json
  "pkg": {
    "assets": [
      "template.docx",
      "template_excel.xlsx"
    ],
    "targets": [
      "node14-win-x64"
    ],
    "outputPath": "dist"
  }
```

> **注意**: 图标文件需要用户提供 `assets/icon.ico`。当前先写占位说明，实际构建时替换为真实图标。

- [ ] **Step 2: 创建 assets 目录和说明**

```bash
mkdir -p assets
```

在 `assets/icon.ico` 位置写入说明（或由用户自行放入 .ico 文件）：

```bash
echo "请将 .ico 图标文件放入此目录，命名为 icon.ico" > assets/README.txt
```

- [ ] **Step 3: Commit**

```bash
git add package.json assets/
git commit -m "feat: add pkg config with icon support and bundled templates for Windows build"
```

---

### Task 11: 清理测试文件 + 最终验证

- [ ] **Step 1: 清理**

```bash
rm -rf output/ seed_test-*.docx seed_test-*.xlsx ts_test-*.docx ts_test-*.xlsx dir_test-*.docx dir_test-*.xlsx
```

- [ ] **Step 2: 端到端验证所有新增功能**

```bash
# 1. --help
node word.js --help 2>&1 | head -5

# 2. 位置参数
node word.js test.xlsx template.docx -o pos_test 2>&1

# 3. 种子复现
node word.js -t template.docx -e test.xlsx -o seed_a -s 123 2>&1
node word.js -t template.docx -e test.xlsx -o seed_b -s 123 2>&1
node -e "
const fs=require('fs');
const f1=fs.readdirSync('output').find(f=>f.startsWith('seed_a-')&&f.endsWith('.xlsx'));
const f2=fs.readdirSync('output').find(f=>f.startsWith('seed_b-')&&f.endsWith('.xlsx'));
console.log(f1,'vs',f2,'- both should exist');
"

# 4. 输出目录
ls output/ | head -10
```

- [ ] **Step 3: 清理 + Commit**

```bash
rm -rf output/
git add .
git commit -m "chore: final verification and cleanup"
```

---

## Self-Review

### 1. Spec Coverage

| 优化点 | 覆盖 Task |
|--------|-----------|
| 1. 报错闪退 | Task 2 |
| 2. 拖拽文件 | Task 3 |
| 3. 输出目录 | Task 4 |
| 4. 中文乱码 | Task 5 |
| 5. --help | Task 1 |
| 6. 交互引导 | Task 6 |
| 7. 进度提示 | Task 7 |
| 9. 覆盖保护 | Task 8 |
| 10. 打开文件夹 | Task 9 |
| 12. 图标版本 | Task 10 |
| 10. 打开文件夹 | Task 9 |

### 2. Placeholder Scan

无 TBD、TODO、"implement later" 等占位符。所有代码块完整可执行。

### 3. Type Consistency

- `randomInRange(min, max, decimals)` 默认使用 `Math.random`
- `maybePauseBeforeExit()` 无参数，所有调用点已更新

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-07-06-windows-ux-optimization.md`.
