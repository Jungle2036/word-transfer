# Excel集成Word/Excel文档生成器

一个基于Node.js的命令行工具，从Excel文件读取产品编码，生成含随机校验数据的Word文档和Excel检验报告。支持自动分页、数据串联、样式保留。

## 功能特性

- 📊 **Excel读取**: 从Excel文件第四列读取产品编码
- 📄 **Word文档生成**: 使用模板生成包含校验数据的Word文档
- 📈 **Excel报告生成**: 同步生成带样式的Excel检验报告（保留模板边框、合并单元格、字体等）
- 🔗 **数据串联**: Q3→Q2→Q1 水表数据逐级串联，末值约束 ≤1000
- 🔄 **自动分页**: 每页固定20行，自动计算总页数
- 🎲 **随机数据生成**: 80% 在 [-1.0, 1.0]，20% 在 [-1.5, -1.0) ∪ (1.0, 1.5]
- 💻 **跨平台**: 支持Windows、macOS和Linux
- 🔧 **灵活配置**: 支持命令行参数和交互式模式

## 安装

### 环境要求

- Node.js 14.0 或更高版本
- npm 或 pnpm

### 安装依赖

```bash
pnpm install
```

## 使用方法

### 命令行模式

```bash
# 基本用法（生成 Word + Excel）
node word.js --template=template.docx --excel=data.xlsx --output=result

# 使用短参数
node word.js -t template.docx -e data.xlsx -o result

# 指定 Excel 模板路径
node word.js -t template.docx -e data.xlsx -x my_excel_template.xlsx -o result
```

### 交互式模式

```bash
node word.js
```

程序会依次询问：
- Word 模板文件路径（默认：template.docx）
- Excel 数据文件路径（默认：test.xlsx）

## 参数说明

| 参数 | 短参数 | 描述 | 默认值 |
|------|--------|------|--------|
| `--template` | `-t` | Word 模板文件路径 | template.docx |
| `--excel` | `-e` | Excel 数据文件路径（读取产品编码） | test.xlsx |
| `--excel-template` | `-x` | Excel 报告模板文件路径 | template_excel.xlsx |
| `--output` | `-o` | 输出文件名前缀 | output |

## Excel 数据文件格式（输入）

用于读取产品编码的 Excel 文件要求：

1. **第一行为表头**（内容不限）
2. **第四列包含产品编码**（从第二行开始）
3. **支持的格式**：.xlsx 和 .xls

### 示例

| 列1 | 列2 | 列3 | 产品编码 |
|-----|-----|-----|----------|
| 数据1 | 数据2 | 数据3 | PROD001 |
| 数据1 | 数据2 | 数据3 | PROD002 |

## Excel 报告模板（输出模板）

`template_excel.xlsx` 作为 Excel 输出的样式模板，程序会在其数据行区域填入生成的校验数据，**完整保留模板的边框、字体、对齐、合并单元格等所有样式**。

模板要求：
- 第 1~12 行为报告头信息（可自定义）
- 第 13 行起为数据行（模板提供 22 行示例），列结构：

| 列 | 内容 | 来源 |
|----|------|------|
| A | 序号 | 自动递增 |
| B-D | 表号 | 产品编码 |
| E | Q3 水表始值 | 随机 0.01~935.74 |
| F | Q3 水表末值 | 公式反算 |
| G | Q3 量器示值 | 固定 60 |
| H | Q3 示值误差 | 随机误差 |
| I | Q2 水表始值 | = F（Q3末值） |
| J | Q2 水表末值 | 公式反算 |
| K | Q2 量器示值 | 固定 2 |
| L | Q2 示值误差 | 随机误差 |
| M | Q1 水表始值 | = J（Q2末值） |
| N | Q1 水表末值 | 公式反算（≤1000） |
| O | Q1 量器示值 | 固定 1 |
| P | Q1 示值误差 | 随机误差 |
| Q | 密封性实验 | 无渗漏 |
| R | 外观、标志和 | 符合 |
| S | 检定结论 | 合格 |

## 数据生成规则

### 串联关系

```
q3_start(E) → q3_end(F) → q2_start(I) → q2_end(J) → q1_start(M) → q1_end(N) ≤ 1000
```

### 反算公式

```
末值 = 始值 + 量器示值 × (1 + 误差/100)，保留 2 位小数
```

### 随机误差分布

- **80%** 概率：均匀分布在 [-1.0, 1.0]
- **20%** 概率：均匀分布在 [-1.5, -1.0) 或 (1.0, 1.5]

## Word 模板要求

### 单页模板（向后兼容）

```
{{#errors}}
Q1: {{q1}}, Q2: {{q2}}, Q3: {{q3}}, 产品编码: {{q4}}
{{/errors}}
```

### 多页模板（推荐）

```
{{#pages}}
第{{pageIndex}}页
{{#errors}}
Q1: {{q1}}, Q2: {{q2}}, Q3: {{q3}}, 产品编码: {{q4}}
{{/errors}}
{{/pages}}
```

## 输出文件

同时生成两个文件，使用同一时间戳：

```
{输出前缀}-{时间戳}.docx   ← Word 文档
{输出前缀}-{时间戳}.xlsx   ← Excel 报告
```

例如：`output-20240315-143022.docx` + `output-20240315-143022.xlsx`

## 构建可执行文件

```bash
# 构建所有平台
pnpm build

# 构建特定平台
pnpm build:win    # Windows
pnpm build:mac    # macOS
pnpm build:linux  # Linux
```

构建后的文件位于 `dist/` 目录。

## 技术栈

- **Node.js**: 运行环境
- **docxtemplater**: Word 文档模板处理
- **pizzip**: ZIP 文件处理
- **xlsx**: Excel 文件数据读取
- **exceljs**: Excel 报告模板写入（保留样式）
- **pkg**: 可执行文件打包

## 版本历史

### v1.1.0

- ✅ Excel 检验报告同步生成（保留模板样式）
- ✅ 数据串联生成（Q3→Q2→Q1）
- ✅ 固定量器示值（60/2/1）+ N≤1000 约束
- ✅ `--excel-template` 命令行参数
- ✅ 极端误差范围收窄至 ±1.5

### v1.0.0

- ✅ Excel 文件读取功能
- ✅ 自动分页计算
- ✅ Word 文档生成
- ✅ 命令行参数支持
- ✅ 交互式模式
- ✅ 错误处理

## 许可证

ISC License
