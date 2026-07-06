# Excel 同步生成功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在生成 Word 文档的同时，根据 `template_excel.xlsx` 模板生成对应的 Excel 检验报告（包含随机始值、量器示值、反算末值）。

**Architecture:** 在现有 `word.js` 中扩展数据生成层，新增 Excel 模板写入层。保持现有 Word 生成逻辑不变，新增函数生成随机始值/量器示值、反算末值，并写入 Excel 模板的数据行区域。两个文件共享同一批数据和时间戳。

**Tech Stack:** Node.js, xlsx (0.18.5, 仅用于读取产品编码), exceljs (4.4.0+, 用于模板样式保留的 Excel 生成), docxtemplater, pizzip, fs, path

---

## 数据映射速查

| Word 变量 | Excel 列 | 含义 | 来源 |
|-----------|----------|------|------|
| q3 | H | Q3 示值误差(%) | `generateSpecialRandom()` |
| q2 | L | Q2 示值误差(%) | `generateSpecialRandom()` |
| q1 | P | Q1 示值误差(%) | `generateSpecialRandom()` |
| q4 | B-D | 表号 | Excel 第4列产品编码 |
| — | E | Q3 水表始值(L) | 随机 0.01~935.74，2位小数 |
| — | I | Q2 水表始值(L) | **= F**（Q3 末值） |
| — | M | Q1 水表始值(L) | **= J**（Q2 末值） |
| — | G | Q3 量器示值(L) | **固定 60** |
| — | K | Q2 量器示值(L) | **固定 2** |
| — | O | Q1 量器示值(L) | **固定 1** |
| — | F | Q3 水表末值(L) | `E + 60 × (1 + H/100)` |
| — | J | Q2 水表末值(L) | `I + 2 × (1 + L/100)` |
| — | N | Q1 水表末值(L) | `M + 1 × (1 + P/100)`，**约束 N ≤ 1000** |

**生成顺序：** E(随机) → F(反算) → I=F → J(反算) → M=J → N(反算) → 若 N>1000 则重试。

---

### Task 1: 添加精度安全的数值工具函数

**Files:**
- Modify: `word.js` (在第 140 行 `generateSpecialRandom` 函数之后插入)
- Test: 直接在 `word.js` 底部 `module.exports` 中追加导出

- [ ] **Step 1: 添加 `roundToDecimals` 函数**

在 `generateSpecialRandom` 函数（第 140 行）之后插入：

```javascript
/**
 * 精度安全的四舍五入，避免 JS 浮点数误差
 * @param {number} val - 原始值
 * @param {number} decimals - 保留小数位数
 * @returns {number}
 */
function roundToDecimals(val, decimals) {
  const factor = Math.pow(10, decimals)
  return Math.round((val + Number.EPSILON) * factor) / factor
}
```

- [ ] **Step 2: 添加 `randomInRange` 函数**

紧接 `roundToDecimals` 之后插入：

```javascript
/**
 * 在 (min, max) 开区间内生成随机数，保留指定小数位
 * 通过重试机制严格排除边界值
 * @param {number} min - 下限（不含）
 * @param {number} max - 上限（不含）
 * @param {number} decimals - 小数位数
 * @returns {number}
 */
function randomInRange(min, max, decimals) {
  let val
  do {
    val = roundToDecimals(Math.random() * (max - min) + min, decimals)
  } while (val <= min || val >= max)
  return val
}
```

- [ ] **Step 3: 运行 module.exports 导出的函数验证加载无语法错误**

```bash
node -e "require('./word.js'); console.log('OK')"
```

- [ ] **Step 4: Commit**

```bash
git add word.js
git commit -m "feat: add roundToDecimals and randomInRange utility functions"
```

---

### Task 2: 添加 Excel 数据生成函数

**Files:**
- Modify: `word.js` (在 `randomInRange` 之后插入新函数)
- Test: 同文件 `module.exports` 追加导出

- [ ] **Step 1: 添加始值生成函数**

```javascript
/**
 * 生成水表始值：0.01 ~ 999.99，不含 0.00 和 1000.00
 * @returns {number}
 */
function randomStartValue() {
  return randomInRange(0.01, 999.99, 2)
}
```

- [ ] **Step 2: 添加量器示值生成函数**

```javascript
/**
 * 生成量器示值：60.01 ~ 69.99，不含 60.00 和 70.00
 * @returns {number}
 */
function randomInstrumentValue() {
  return randomInRange(60.01, 69.99, 2)
}
```

- [ ] **Step 3: 添加末值反算函数**

```javascript
/**
 * 根据始值、量器示值、示值误差反算末值
 * 公式：F = round(startVal + instVal × (1 + error/100), 2)
 * @param {number} startVal - 水表始值
 * @param {number} instVal - 量器示值
 * @param {number} error - 示值误差(%)
 * @returns {number}
 */
function calculateEndValue(startVal, instVal, error) {
  return roundToDecimals(startVal + instVal * (1 + error / 100), 2)
}
```

- [ ] **Step 4: 运行验证**

```bash
node -e "
const m = require('./word.js');
console.log('start:', m.randomStartValue());
console.log('inst:', m.randomInstrumentValue());
console.log('end:', m.calculateEndValue(100, 65, -0.5));
"
```

预期：三个数值都输出，末值约为 164.68（100 + 65 × 0.995 = 164.675 → 164.68）

- [ ] **Step 5: Commit**

```bash
git add word.js
git commit -m "feat: add randomStartValue, randomInstrumentValue, calculateEndValue"
```

---

### Task 3: 扩展 generateErrors 返回 Excel 字段

**Files:**
- Modify: `word.js:290-308` (`generateErrors` 函数体)

- [ ] **Step 1: 修改 `generateErrors` 函数**

将当前函数体（第 290-308 行）替换为：

```javascript
function generateErrors(rowsPerPage, productCodes, startIndex) {
  const errors = new Array(rowsPerPage)

  for (let i = 0; i < rowsPerPage; i++) {
    const currentIndex = startIndex + i
    const q1 = generateSpecialRandom()
    const q2 = generateSpecialRandom()
    const q3 = generateSpecialRandom()

    const row = {
      // Word 模板字段
      q1,
      q2,
      q3,
      q4: productCodes && currentIndex < productCodes.length ? productCodes[currentIndex] : undefined,

      // Excel 字段 — Q3 (E-H)
      q3_start: randomStartValue(),
      q3_instrument: randomInstrumentValue(),
      q3_end: null, // 延迟计算，q3 已知时填入

      // Excel 字段 — Q2 (I-L)
      q2_start: randomStartValue(),
      q2_instrument: randomInstrumentValue(),
      q2_end: null,

      // Excel 字段 — Q1 (M-P)
      q1_start: randomStartValue(),
      q1_instrument: randomInstrumentValue(),
      q1_end: null,
    }

    // 反算末值（q1/q2/q3 已生成）
    row.q3_end = calculateEndValue(row.q3_start, row.q3_instrument, q3)
    row.q2_end = calculateEndValue(row.q2_start, row.q2_instrument, q2)
    row.q1_end = calculateEndValue(row.q1_start, row.q1_instrument, q1)

    errors[i] = row
  }
  return errors
}
```

- [ ] **Step 2: 更新 `module.exports`**

在 `word.js` 底部 `module.exports`（约第 405-411 行）中追加导出：

```javascript
module.exports = {
  readProductCodesFromExcel,
  calculatePagination,
  generateSpecialRandom,
  generateErrors,
  generatePages,
  // 新增导出
  roundToDecimals,
  randomInRange,
  randomStartValue,
  randomInstrumentValue,
  calculateEndValue,
}
```

- [ ] **Step 3: 运行 Word 生成验证现有功能未破坏**

```bash
node word.js --template=template.docx --excel=test.xlsx --output=test_output 2>&1
```

预期：输出 `文件 "test_output-YYYYMMDD-HHmmss.docx" 已成功生成！`

- [ ] **Step 4: Commit**

```bash
git add word.js
git commit -m "feat: extend generateErrors with Excel data fields (start/instrument/end values)"
```

---

### Task 4: 实现 Excel 模板写入功能（使用 exceljs 保留样式）

**背景:** `xlsx` 库读写文件会丢失模板的样式（边框、字体、对齐、合并单元格等）。改用 `exceljs` 库，其 `readFile`/`writeFile` 完整保留所有格式。

**Files:**
- Modify: `word.js` (添加 `require('exceljs')`，重写 `generateExcel` 为 `async` 函数)
- Install: `pnpm add exceljs`

- [ ] **Step 1: 安装 exceljs 依赖**

```bash
pnpm add exceljs
```

- [ ] **Step 2: 在 word.js 顶部添加 require**

在第 7 行 `const XLSX = require('xlsx')` 之后添加：

```javascript
const ExcelJS = require('exceljs')
```

- [ ] **Step 3: 重写 `generateExcel` 函数（async，使用 exceljs）**

找到当前 `generateExcel` 函数，完整替换为：

```javascript
/**
 * 从模板生成 Excel 文件（保留模板样式、合并单元格、边框等）
 * @param {string} templatePath - Excel 模板路径
 * @param {string} outputPath - 输出文件路径
 * @param {Array} dataRows - 数据行数组
 * @returns {Promise<void>}
 */
async function generateExcel(templatePath, outputPath, dataRows) {
  const DATA_START_ROW = 13   // 数据起始行
  const TEMPLATE_DATA_ROWS = 22 // 模板原有数据行数（第 13~34 行）

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Excel模板文件不存在: ${templatePath}`)
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(templatePath)
  const worksheet = workbook.getWorksheet(1)

  // 获取一行模板数据行的样式引用（第 13 行）
  const styleRow = worksheet.getRow(DATA_START_ROW)

  // 清空模板原有数据区域（第 13~34 行，A~S 列）
  for (let r = DATA_START_ROW; r < DATA_START_ROW + TEMPLATE_DATA_ROWS; r++) {
    const row = worksheet.getRow(r)
    for (let c = 1; c <= 19; c++) { // A(1) ~ S(19)
      row.getCell(c).value = null
    }
  }

  // 填充数据
  for (let i = 0; i < dataRows.length; i++) {
    const data = dataRows[i]
    const rowNum = DATA_START_ROW + i
    const row = worksheet.getRow(rowNum)

    // 如果当前行超出模板原有行数，复制样式
    if (rowNum >= DATA_START_ROW + TEMPLATE_DATA_ROWS) {
      row.height = styleRow.height
      for (let c = 1; c <= 19; c++) {
        row.getCell(c).style = styleRow.getCell(c).style
        if (styleRow.getCell(c).border) {
          row.getCell(c).border = styleRow.getCell(c).border
        }
      }
    }

    // A 列：序号
    row.getCell(1).value = i + 1
    // B 列：表号（产品编码）
    row.getCell(2).value = data.q4 ?? ''

    // Q3 段 — E(5), F(6), G(7), H(8)
    row.getCell(5).value = data.q3_start
    row.getCell(6).value = data.q3_end
    row.getCell(7).value = data.q3_instrument
    row.getCell(8).value = data.q3

    // Q2 段 — I(9), J(10), K(11), L(12)
    row.getCell(9).value = data.q2_start
    row.getCell(10).value = data.q2_end
    row.getCell(11).value = data.q2_instrument
    row.getCell(12).value = data.q2

    // Q1 段 — M(13), N(14), O(15), P(16)
    row.getCell(13).value = data.q1_start
    row.getCell(14).value = data.q1_end
    row.getCell(15).value = data.q1_instrument
    row.getCell(16).value = data.q1

    // Q(17): 密封性, R(18): 外观, S(19): 检定结论
    row.getCell(17).value = '无渗漏'
    row.getCell(18).value = '符合'
    row.getCell(19).value = '合格'
  }

  // 如果数据行少于模板行数，收缩范围（清除尾部空行）
  if (dataRows.length < TEMPLATE_DATA_ROWS) {
    const firstEmptyRow = DATA_START_ROW + dataRows.length
    // 清除多余行的所有单元格值和样式引用（防止残留空行边框）
    for (let r = firstEmptyRow; r < DATA_START_ROW + TEMPLATE_DATA_ROWS; r++) {
      const row = worksheet.getRow(r)
      for (let c = 1; c <= 19; c++) {
        row.getCell(c).value = null
      }
      // 移除行高以避免空行占位
      row.height = undefined
    }
  }

  // 写入文件（保留所有样式）
  await workbook.xlsx.writeFile(outputPath)
}
```

- [ ] **Step 4: 运行验证**

```bash
node -e "
const m = require('./word.js');
const testRow = {
  q4: 'TEST001',
  q3_start: 500.00, q3_instrument: 65.00, q3_end: 564.68, q3: -0.5,
  q2_start: 300.00, q2_instrument: 62.00, q2_end: 361.69, q2: -0.5,
  q1_start: 100.00, q1_instrument: 60.50, q1_end: 160.20, q1: -0.5,
};
(async () => {
  await m.generateExcel('./template_excel.xlsx', './test_excel_output.xlsx', [testRow]);
  console.log('OK');
})();
"
```

预期：生成 `test_excel_output.xlsx`，无报错。打开文件检查样式是否与模板一致。

- [ ] **Step 5: 更新 `module.exports` 追加 `generateExcel`**

```javascript
module.exports = {
  // ... 现有导出 ...
  generateExcel,
}
```

- [ ] **Step 6: Commit**

```bash
git add word.js package.json pnpm-lock.yaml
git commit -m "feat: rewrite generateExcel with exceljs to preserve template styles"
```

---

### Task 5: 在 main() 中集成 Excel 生成

**Files:**
- Modify: `word.js` (修改 `main` 函数，约第 333-402 行)

- [ ] **Step 1: 修改 `main()` 函数**

在当前 Word 文件写出之后、最终 `console.log` 之前，插入 Excel 生成逻辑。

找到以下代码块（约第 385-398 行）：

```javascript
  // 6. 生成新的Word文件（输出名增加时间戳）
  const buf = doc.getZip().generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  })

  const timestamp = formatTimestamp()
  const outputBaseName = args.output || 'output'
  const outputFileName = `${outputBaseName}-${timestamp}.docx`
  const outputPath = path.resolve(process.cwd(), outputFileName)

  fs.writeFileSync(outputPath, buf)

  console.log(`文件 "${outputFileName}" 已成功生成！`)
```

替换为：

```javascript
  // 6. 生成新的Word文件（输出名增加时间戳）
  const buf = doc.getZip().generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  })

  const timestamp = formatTimestamp()
  const outputBaseName = args.output || 'output'
  const outputFileName = `${outputBaseName}-${timestamp}.docx`
  const outputPath = path.resolve(process.cwd(), outputFileName)

  fs.writeFileSync(outputPath, buf)

  console.log(`文件 "${outputFileName}" 已成功生成！`)

  // 7. 生成 Excel 文件（使用同一时间戳）
  const excelTemplatePath = path.resolve(process.cwd(), 'template_excel.xlsx')
  if (fs.existsSync(excelTemplatePath)) {
    // 将 pages 中的行数据扁平化为一维数组
    const flatDataRows = []
    for (const page of pages) {
      for (const row of page.errors) {
        flatDataRows.push(row)
      }
    }

    const excelOutputFileName = `${outputBaseName}-${timestamp}.xlsx`
    const excelOutputPath = path.resolve(process.cwd(), excelOutputFileName)

    try {
      await generateExcel(excelTemplatePath, excelOutputPath, flatDataRows)
      console.log(`Excel文件 "${excelOutputFileName}" 已成功生成！`)
    } catch (excelErr) {
      console.error(`Excel文件生成失败: ${excelErr.message}`)
      // 不中断程序，Word 文件已经生成成功
    }
  } else {
    console.warn(`未找到Excel模板: ${excelTemplatePath}，跳过Excel生成`)
  }
```

- [ ] **Step 2: 端到端验证**

```bash
node word.js --template=template.docx --excel=test.xlsx --output=e2e_test 2>&1
```

预期输出：
```
正在读取Excel文件: test.xlsx
成功读取 X 个产品编码
将生成 X 页，每页最多 20 行
文件 "e2e_test-YYYYMMDD-HHmmss.docx" 已成功生成！
Excel文件 "e2e_test-YYYYMMDD-HHmmss.xlsx" 已成功生成！
```

- [ ] **Step 3: 验证生成的 Excel 数据正确性**

```bash
node -e "
const XLSX = require('xlsx');
const wb = XLSX.readFile('./e2e_test-2026*.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
// 检查第13行（第一条数据）的Q3段 E-H
const row13 = data[12]; // 0-indexed
console.log('E(始值):', row13[4]);
console.log('F(末值):', row13[5]);
console.log('G(量器):', row13[6]);
console.log('H(误差):', row13[7]);
// 验证公式: F = E + G*(1+H/100)
const verify = row13[4] + row13[6] * (1 + row13[7]/100);
console.log('反算F:', Math.round(verify*100)/100, '实际F:', row13[5]);
console.log('验证:', Math.abs(Math.round(verify*100)/100 - row13[5]) < 0.005 ? 'PASS' : 'FAIL');
"
```

预期：PASS

- [ ] **Step 4: Commit**

```bash
git add word.js
git commit -m "feat: integrate Excel generation into main flow with same timestamp"
```

---

### Task 6: 清理测试产物并最终验证

- [ ] **Step 1: 清理临时文件**

```bash
rm -f e2e_test-*.docx e2e_test-*.xlsx test_output-*.docx test_excel_output.xlsx
```

- [ ] **Step 2: 无参数交互模式验证**

```bash
echo -e "template.docx\ntest.xlsx\n" | node word.js 2>&1
```

预期：正常输出 Word 和 Excel 两个文件。

- [ ] **Step 3: 边界测试 — 验证 randomStartValue 不含边界**

```bash
node -e "
const m = require('./word.js');
for(let i=0; i<10000; i++) {
  const v = m.randomStartValue();
  if(v <= 0.01 || v >= 999.99) { console.log('FAIL: boundary hit', v); process.exit(1); }
}
console.log('PASS: 10000 values all within bounds');
for(let i=0; i<10000; i++) {
  const v = m.randomInstrumentValue();
  if(v <= 60.01 || v >= 69.99) { console.log('FAIL: boundary hit', v); process.exit(1); }
}
console.log('PASS: 10000 instrument values all within bounds');
"
```

预期：两次 PASS。

- [ ] **Step 4: 精度测试 — 验证末值反算精度**

```bash
node -e "
const m = require('./word.js');
// 测试经典浮点陷阱
const r = m.calculateEndValue(0.1, 0.2, 10);
console.log('0.1 + 0.2*(1+10/100) =', r);
// 验证是 2 位小数
console.log('decimals:', r.toString().match(/\.\d+/) ? r.toString().split('.')[1].length : 0);
console.log('PASS');
"
```

预期：末值为 2 位小数。

- [ ] **Step 5: Commit cleanup**

```bash
git add .
git commit -m "chore: cleanup test artifacts"
```

---
## Phase 2: 修改生成规则（串联 + 固定量器 + 约束）

> 在 Phase 1 实现基础上，修改数据生成规则：

### Task 7: 修改数据生成规则

**Files:**
- Modify: `word.js`

- [ ] **Step 1: 修改 `randomStartValue` 范围**

将 `randomStartValue` 的范围从 `(0.01, 999.99)` 改为 `(0.01, 935.74)`：

```javascript
function randomStartValue() {
  return randomInRange(0.01, 935.74, 2)
}
```

- [ ] **Step 2: 删除 `randomInstrumentValue` 函数**

量器示值改为固定值（Q3=60, Q2=2, Q1=1），不再需要 `randomInstrumentValue`。删除该函数定义及其在 `module.exports` 中的导出。

- [ ] **Step 3: 重写 `generateErrors` 为串联生成 + 约束**

```javascript
function generateErrors(rowsPerPage, productCodes, startIndex) {
  const errors = new Array(rowsPerPage)

  for (let i = 0; i < rowsPerPage; i++) {
    const currentIndex = startIndex + i
    const q1 = generateSpecialRandom()
    const q2 = generateSpecialRandom()
    const q3 = generateSpecialRandom()

    let row
    // 重试直到 Q1 末值 ≤ 1000
    do {
      // Q3 段
      const q3_start = randomStartValue()         // E: 0.01~935.74
      const q3_instrument = 60                    // G: 固定
      const q3_end = calculateEndValue(q3_start, q3_instrument, q3) // F

      // Q2 段 — 始值 = Q3 末值
      const q2_start = q3_end                     // I = F
      const q2_instrument = 2                     // K: 固定
      const q2_end = calculateEndValue(q2_start, q2_instrument, q2) // J

      // Q1 段 — 始值 = Q2 末值
      const q1_start = q2_end                     // M = J
      const q1_instrument = 1                     // O: 固定
      const q1_end = calculateEndValue(q1_start, q1_instrument, q1) // N

      row = {
        q1, q2, q3,
        q4: productCodes && currentIndex < productCodes.length ? productCodes[currentIndex] : undefined,
        q3_start, q3_instrument, q3_end,
        q2_start, q2_instrument, q2_end,
        q1_start, q1_instrument, q1_end,
      }
    } while (row.q1_end > 1000)

    errors[i] = row
  }
  return errors
}
```

- [ ] **Step 4: 运行验证**

```bash
node word.js --template=template.docx --excel=test.xlsx --output=phase2_test 2>&1
```

- [ ] **Step 5: 数据验证**

```bash
node -e "
const XLSX = require('xlsx');
const wb = XLSX.readFile('./phase2_test-*.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
const row13 = data[12];
console.log('E(q3_start):', row13[4]);
console.log('G(q3_inst):', row13[6], row13[6] === 60 ? 'PASS' : 'FAIL (should be 60)');
console.log('F(q3_end):', row13[5]);
console.log('I(q2_start):', row13[8], row13[8] === row13[5] ? 'PASS (I=F)' : 'FAIL');
console.log('K(q2_inst):', row13[10], row13[10] === 2 ? 'PASS' : 'FAIL (should be 2)');
console.log('J(q2_end):', row13[9]);
console.log('M(q1_start):', row13[12], row13[12] === row13[9] ? 'PASS (M=J)' : 'FAIL');
console.log('O(q1_inst):', row13[14], row13[14] === 1 ? 'PASS' : 'FAIL (should be 1)');
console.log('N(q1_end):', row13[13], row13[13] <= 1000 ? 'PASS (<=1000)' : 'FAIL');
"
```

- [ ] **Step 6: Commit**

```bash
git add word.js
git commit -m "feat: chained generation with fixed instrument values and N<=1000 constraint"
```

### Task 8: 清理 + 最终验证

- [ ] **Step 1: 清理测试文件**

```bash
rm -f phase2_test-*.docx phase2_test-*.xlsx merge_test-*.docx merge_test-*.xlsx
```

- [ ] **Step 2: 端到端验证**

```bash
node word.js --template=template.docx --excel=test.xlsx --output=final 2>&1
```

- [ ] **Step 3: 检查全部 123 行的 N 值 ≤ 1000**

```bash
node -e "
const XLSX = require('xlsx');
const wb = XLSX.readFile('./final-*.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
let failCount = 0;
for (let i = 12; i < Math.min(data.length, 135); i++) {
  const row = data[i];
  if (row[13] > 1000) { console.log('FAIL row', i+1, 'N=', row[13]); failCount++; }
}
console.log(failCount === 0 ? 'PASS: all N ≤ 1000' : 'FAIL: ' + failCount + ' violations');
"
```

- [ ] **Step 4: Commit cleanup**

```bash
rm -f final-*.docx final-*.xlsx
git add .
git commit -m "chore: final verification and cleanup"
```

---

## Self-Review

### 1. Spec Coverage

| 需求 | 覆盖 Task |
|------|-----------|
| 始值随机生成 (0.01~935.74, 不含两端) | Task 7 |
| 量器示值固定 (60/2/1) | Task 7 |
| 末值反算公式 F = E + inst × (1 + error/100) | Task 2, Task 3 |
| 串联: F→I, J→M | Task 7 |
| N ≤ 1000 约束 + 重试 | Task 7 |
| 2 位小数精度 + 防浮点误差 | Task 1 |
| q1→P, q2→L, q3→H 映射 | Task 3, Task 4 |
| q4→B-D 表号 | Task 4 |
| 模板 Excel 读写（保留样式） | Task 4 |
| Word + Excel 同步生成 | Task 5 |
| 固定列（序号/密封性/外观/结论） | Task 4 |
| 额外行 B-D 合并 | 修复 |

### 2. Placeholder Scan

无 TBD、TODO 等占位符。

### 3. Type Consistency

- 所有新字段在整个数据流中保持一致
- `generateExcel` 的数据列映射与 `generateErrors` 输出一致
- 串联逻辑的字段映射已在 Task 7 中明确定义

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-07-06-excel-generation.md`.
