#!/usr/bin/env node
const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')
const fs = require('fs')
const path = require('path')
const readline = require('readline')
const XLSX = require('xlsx')

/**
 * 从Excel文件读取产品编码（优化版本）
 * @param {string} filePath - Excel文件路径
 * @returns {Promise<string[]>} 产品编码数组
 */
async function readProductCodesFromExcel(filePath) {
  try {
    // 检查文件是否存在
    if (!fs.existsSync(filePath)) {
      throw new Error(`Excel文件不存在: ${filePath}`)
    }

    // 检查文件扩展名
    const ext = path.extname(filePath).toLowerCase()
    if (ext !== '.xlsx' && ext !== '.xls') {
      throw new Error(`不支持的文件格式: ${ext}，请使用.xlsx或.xls格式`)
    }

    // 读取Excel文件
    const workbook = XLSX.readFile(filePath)

    // 获取第一个工作表
    const sheetNames = workbook.SheetNames
    if (sheetNames.length === 0) {
      throw new Error('Excel文件中没有工作表')
    }

    const worksheet = workbook.Sheets[sheetNames[0]]

    // 获取工作表范围
    if (!worksheet['!ref']) {
      throw new Error('Excel文件中没有数据')
    }

    const range = XLSX.utils.decode_range(worksheet['!ref'])

    // 检查是否有足够的数据行（至少需要表头和一行数据）
    if (range.e.r < 1) {
      throw new Error('Excel文件中没有足够的数据（至少需要表头和一行数据）')
    }

    // 检查是否有第四列
    if (range.e.c < 3) {
      throw new Error('Excel文件中没有第四列数据')
    }

    // 预分配数组，提高性能
    const estimatedRows = range.e.r - range.s.r
    const productCodes = []
    productCodes.length = 0 // 确保数组为空但保留预分配

    // 直接访问第四列的单元格，从第二行开始（跳过表头）
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: 3 }) // 第四列，索引为3
      const cell = worksheet[cellAddress]

      if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
        // 直接使用单元格值，减少类型转换
        const cellValue = cell.v
        let codeStr

        // 优化字符串处理
        if (typeof cellValue === 'string') {
          codeStr = cellValue.trim()
        } else {
          codeStr = String(cellValue).trim()
        }

        if (codeStr) {
          productCodes.push(codeStr)
        }
      }
    }

    // 检查是否读取到有效数据
    if (productCodes.length === 0) {
      throw new Error('Excel文件第四列中没有找到有效的产品编码数据')
    }

    return productCodes
  } catch (error) {
    // 重新抛出错误，保持原始错误信息
    if (error.message.includes('Excel文件') || error.message.includes('不支持的文件格式')) {
      throw error
    } else {
      throw new Error(`读取Excel文件时发生错误: ${error.message}`)
    }
  }
}

/**
 * 根据产品编码数量计算分页参数
 * @param {string[]} productCodes - 产品编码数组
 * @returns {Object} 分页参数 {pages, rowsPerPage, totalItems, lastPageRows}
 */
function calculatePagination(productCodes) {
  const totalItems = productCodes.length
  const rowsPerPage = 20 // 固定每页20行
  const pages = Math.ceil(totalItems / rowsPerPage)
  const lastPageRows = totalItems % rowsPerPage || rowsPerPage

  return {
    pages,
    rowsPerPage,
    totalItems,
    lastPageRows,
  }
}

/**
 * 生成一个随机数。
 * 80% 的概率在 [-1.0, 1.0] 之间。
 * 20% 的概率在 [-2.0, -1.0) 或 (1.0, 2.0] 之间。
 * @returns {number} 返回一个 -2.0 到 2.0 之间的一位小数
 */
function generateSpecialRandom() {
  const randomValue = Math.random()
  let result

  if (randomValue < 0.8) {
    // 80% 的情况: 值在 [-1.0, 1.0]
    result = Math.random() * 2 - 1
  } else {
    // 20% 的情况: 值在 [-2.0, -1.0) 或 (1.0, 2.0]
    result = Math.random() + 1 // 生成 [1.0, 2.0) 的数
    if (Math.random() < 0.5) {
      result = -result // 50% 的概率变负
    }
  }
  // 保留 1 位小数
  return parseFloat(result.toFixed(1))
}

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

/**
 * 在 (min, max) 开区间内生成随机数，保留指定小数位
 * 通过重试机制严格排除边界值
 * @param {number} min - 下限（不含）
 * @param {number} max - 上限（不含）
 * @param {number} decimals - 小数位数
 * @returns {number}
 */
function randomInRange(min, max, decimals) {
  const step = 1 / Math.pow(10, decimals)
  if (max - min <= 2 * step) {
    throw new RangeError(
      `randomInRange: range (${max - min}) too narrow for ${decimals} decimals`
    )
  }
  let val
  do {
    val = roundToDecimals(Math.random() * (max - min) + min, decimals)
  } while (val <= min || val >= max)
  return val
}

/**
 * 生成水表始值：0.01 ~ 999.99，不含 0.00 和 1000.00
 * @returns {number}
 */
function randomStartValue() {
  return randomInRange(0.01, 999.99, 2)
}

/**
 * 生成量器示值：60.01 ~ 69.99，不含 60.00 和 70.00
 * @returns {number}
 */
function randomInstrumentValue() {
  return randomInRange(60.01, 69.99, 2)
}

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

// --- 主程序 ---

function formatTimestamp(date = new Date()) {
  const pad = (n) => String(n).padStart(2, '0')
  const y = date.getFullYear()
  const m = pad(date.getMonth() + 1)
  const d = pad(date.getDate())
  const hh = pad(date.getHours())
  const mm = pad(date.getMinutes())
  const ss = pad(date.getSeconds())
  return `${y}${m}${d}-${hh}${mm}${ss}`
}

function parseCliArgs(argv) {
  const args = {
    template: undefined,
    excel: undefined,
    output: 'output',
  }

  const kv = /^--([^=]+)=(.*)$/
  const positional = []

  for (let i = 0; i < argv.length; i++) {
    const token = argv[i]
    const match = kv.exec(token)
    if (match) {
      const key = match[1]
      const value = match[2]
      switch (key) {
        case 'template':
          args.template = value
          break
        case 'excel':
          args.excel = value
          break
        case 'output':
          args.output = value
          break
      }
      continue
    }

    if (token === '--template' || token === '-t') {
      args.template = argv[++i]
      continue
    }
    if (token === '--excel' || token === '-e') {
      args.excel = argv[++i]
      continue
    }
    if (token === '--output' || token === '-o') {
      args.output = argv[++i]
      continue
    }
    if (token.startsWith('-')) {
      // 未知参数，忽略
      continue
    }
    positional.push(token)
  }

  // 不在此处强制校验，交由后续交互流程处理。
  return args
}

const cli = parseCliArgs(process.argv.slice(2))

function createReadline() {
  return readline.createInterface({ input: process.stdin, output: process.stdout })
}

function askQuestion(rl, question, defaultValue) {
  return new Promise((resolve) => {
    const q = defaultValue !== undefined ? `${question}（默认: ${defaultValue}）：` : `${question}：`
    rl.question(q, (answer) => {
      const trimmed = (answer || '').trim()
      if (!trimmed && defaultValue !== undefined) {
        resolve(String(defaultValue))
      } else {
        resolve(trimmed)
      }
    })
  })
}

async function ensureParamsInteractive(baseArgs) {
  const noArgs = process.argv.slice(2).length === 0
  const needInteractive = noArgs || !baseArgs.template || !baseArgs.excel

  if (!needInteractive) return baseArgs

  const rl = createReadline()
  try {
    // 询问模板路径（若不存在则反复询问）
    let templatePathInput = baseArgs.template || 'template.docx'
    // 仅在模板文件不存在时提示
    while (true) {
      const candidate = await askQuestion(rl, '请输入模板路径', templatePathInput)
      const absCandidate = path.resolve(process.cwd(), candidate)
      if (fs.existsSync(absCandidate)) {
        templatePathInput = candidate
        break
      }
      console.error(`未找到模板文件: ${absCandidate}`)
      templatePathInput = candidate
    }

    // 询问Excel文件路径（若不存在则反复询问）
    let excelPathInput = baseArgs.excel || 'test.xlsx'
    while (true) {
      const candidate = await askQuestion(rl, '请输入Excel文件路径', excelPathInput)
      const absCandidate = path.resolve(process.cwd(), candidate)
      if (fs.existsSync(absCandidate)) {
        // 检查文件扩展名
        const ext = path.extname(candidate).toLowerCase()
        if (ext === '.xlsx' || ext === '.xls') {
          excelPathInput = candidate
          break
        } else {
          console.error(`不支持的文件格式: ${ext}，请使用.xlsx或.xls格式`)
          excelPathInput = candidate
        }
      } else {
        console.error(`未找到Excel文件: ${absCandidate}`)
        excelPathInput = candidate
      }
    }

    return {
      template: templatePathInput,
      excel: excelPathInput,
      output: baseArgs.output || 'output',
    }
  } finally {
    rl.close()
  }
}

async function maybePauseBeforeExit(interactiveStarted) {
  const isWindows = process.platform === 'win32'
  if (!interactiveStarted || !isWindows) return
  const rl = createReadline()
  await new Promise((resolve) => rl.question('按回车退出...', () => resolve(undefined)))
  rl.close()
}

// 2. 数据构造工具函数：生成一页中的多行数据（优化版本）
function generateErrors(rowsPerPage, productCodes, startIndex) {
  // 预分配数组大小，提高性能
  const errors = new Array(rowsPerPage)

  for (let i = 0; i < rowsPerPage; i++) {
    const currentIndex = startIndex + i

    // 直接创建对象，减少属性赋值操作
    const row = {
      q1: generateSpecialRandom(),
      q2: generateSpecialRandom(),
      q3: generateSpecialRandom(),
      q4: productCodes && currentIndex < productCodes.length ? productCodes[currentIndex] : undefined,
    }

    errors[i] = row
  }
  return errors
}

// 3. 构造多页数据（每一页的数据彼此不同）（优化版本）
function generatePages(numberOfPages, rowsPerPage, productCodes) {
  // 预分配数组大小，提高性能
  const pages = new Array(numberOfPages)

  for (let pageIndex = 0; pageIndex < numberOfPages; pageIndex++) {
    const startIndex = pageIndex * rowsPerPage

    // 计算当前页的实际行数（最后一页可能少于rowsPerPage）
    let currentPageRows = rowsPerPage
    if (productCodes && startIndex + rowsPerPage > productCodes.length) {
      currentPageRows = productCodes.length - startIndex
    }

    // 直接创建页面对象
    pages[pageIndex] = {
      pageIndex: pageIndex + 1,
      errors: generateErrors(currentPageRows, productCodes, startIndex),
    }
  }
  return pages
}

async function main() {
  const noArgs = process.argv.slice(2).length === 0
  const args = await ensureParamsInteractive(cli)

  // 1. 加载模板文件
  const templatePath = path.resolve(process.cwd(), args.template || 'template.docx')
  if (!fs.existsSync(templatePath)) {
    console.error(`未找到模板文件: ${templatePath}`)
    process.exit(1)
  }
  const content = fs.readFileSync(templatePath, 'binary')

  const zip = new PizZip(content)

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  })

  // 2. 读取Excel文件中的产品编码
  let productCodes
  try {
    console.log(`正在读取Excel文件: ${args.excel}`)
    productCodes = await readProductCodesFromExcel(args.excel)
    console.log(`成功读取 ${productCodes.length} 个产品编码`)
  } catch (error) {
    console.error(`Excel文件读取失败: ${error.message}`)
    process.exit(1)
  }

  // 3. 根据产品编码数量自动计算分页参数
  const paginationParams = calculatePagination(productCodes)
  const numberOfPages = paginationParams.pages
  const rowsPerPage = paginationParams.rowsPerPage

  console.log(`将生成 ${numberOfPages} 页，每页最多 ${rowsPerPage} 行`)

  // 4. 组合模板所需数据结构
  // 兼容旧模板：继续提供单页的 `errors` 字段；
  // 新模板：使用 `pages` 数组进行整页循环。
  const pages = generatePages(numberOfPages, rowsPerPage, productCodes)
  const dataToInsert = {
    pages,
    // 第一页用于兼容旧模板
    errors: pages[0]?.errors || [],
    // 后续页，便于模板中在每一项之前插入分页符，且不会在末尾多出空白页
    pagesRest: pages.slice(1),
  }

  // 5. 渲染文档（用数据替换模板中的标签）
  doc.render(dataToInsert)

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

  // 双击运行时让窗口停留，便于查看结果
  await maybePauseBeforeExit(noArgs)
}

// Export functions for testing
module.exports = {
  readProductCodesFromExcel,
  calculatePagination,
  generateSpecialRandom,
  roundToDecimals,
  randomInRange,
  randomStartValue,
  randomInstrumentValue,
  calculateEndValue,
  generateErrors,
  generatePages,
}

// Only run main if this file is executed directly
if (require.main === module) {
  main().catch(async (err) => {
    // 提供更清晰的错误消息和用户指导
    if (err && err.message) {
      if (err.message.includes('Excel文件')) {
        console.error('\n❌ Excel文件处理错误:')
        console.error(`   ${err.message}`)
        console.error('\n💡 请检查:')
        console.error('   • Excel文件路径是否正确')
        console.error('   • 文件格式是否为.xlsx或.xls')
        console.error('   • 第四列是否包含产品编码数据')
        console.error('   • 文件是否被其他程序占用')
      } else if (err.message.includes('模板文件')) {
        console.error('\n❌ 模板文件错误:')
        console.error(`   ${err.message}`)
        console.error('\n💡 请检查模板文件路径是否正确')
      } else {
        console.error('\n❌ 程序执行错误:')
        console.error(`   ${err.message}`)
        if (err.stack) {
          console.error('\n详细错误信息:')
          console.error(err.stack)
        }
      }
    } else {
      console.error('\n❌ 发生未知错误')
      console.error(err)
    }

    // 保持Windows平台的暂停功能
    const noArgs = process.argv.slice(2).length === 0
    await maybePauseBeforeExit(noArgs)
    process.exit(1)
  })
}
