/**
 * 公式转换模块：LaTeX → MathML → OMML
 * Formula Converter: LaTeX → MathML → OMML
 */

// @ts-ignore - katex types
import katex from 'katex'
// @ts-ignore - xmldom types
import { DOMParser } from '@xmldom/xmldom'

/**
 * 步骤1: LaTeX → MathML
 * Step 1: LaTeX → MathML
 * @param latex - LaTeX formula string
 * @returns MathML XML string
 */
export function latexToMathML(latex: string): string {
	try {
		// 使用 KaTeX 渲染为 MathML
		const mathML = katex.renderToString(latex, {
			output: 'mathml',
			throwOnError: false,
			displayMode: true,
			strict: false
		})

		return mathML
	} catch (error) {
		console.error(`LaTeX conversion failed: ${(error as Error).message}`)
		throw error
	}
}

/**
 * 将普通字母转换为 Unicode 数学斜体字符
 * Convert regular letters to Unicode math italic characters
 */
function toMathItalic(text: string): string {
	// Unicode 数学斜体字母映射
	// 小写: a-z (U+1D44E - U+1D467)
	// 大写: A-Z (U+1D434 - U+1D44D)
	let result = ''
	for (let i = 0; i < text.length; i++) {
		const char = text[i]
		const code = char.charCodeAt(0)

		if (code >= 65 && code <= 90) {
			// A-Z -> 𝐴-𝑍 (U+1D434 - U+1D44D)
			result += String.fromCodePoint(0x1D434 + (code - 65))
		} else if (code >= 97 && code <= 122) {
			// a-z -> 𝑎-𝑧 (U+1D44E - U+1D467)
			result += String.fromCodePoint(0x1D44E + (code - 97))
		} else {
			// 其他字符保持不变
			result += char
		}
	}
	return result
}

/**
 * XML 转义
 * XML escape
 */
function escapeXml(text: string): string {
	if (!text) return ''
	return text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;')
}

/**
 * 获取节点的文本内容
 * Get text content of a node
 */
function getTextContent(node: Node): string {
	if (!node) return ''

	let text = ''
	if (node.childNodes) {
		for (let i = 0; i < node.childNodes.length; i++) {
			const child = node.childNodes[i]
			if (child.nodeType === 3) { // 文本节点
				text += child.nodeValue || ''
			} else {
				text += getTextContent(child)
			}
		}
	}
	return escapeXml(text)
}

/**
 * 转换所有子节点
 * Convert all child nodes
 */
function convertChildren(node: Node): string {
	if (!node.childNodes || node.childNodes.length === 0) {
		return ''
	}

	let result = ''
	for (let i = 0; i < node.childNodes.length; i++) {
		result += convertMathNode(node.childNodes[i])
	}
	return result
}

/**
 * 递归转换 MathML 节点为 OMML
 * Recursively convert MathML node to OMML
 */
function convertMathNode(node: Node): string {
	if (!node) return ''

	// 文本节点
	if (node.nodeType === 3) {
		return escapeXml(node.nodeValue || '')
	}

	const element = node as Element
	const tagName = element.localName || element.tagName

	switch (tagName) {
		case 'math':
			return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${convertChildren(node)}</m:oMath>`

		case 'mi': // 标识符 (变量) - 默认斜体
			const italicText = toMathItalic(getTextContent(node))
			return `<m:r><a:rPr lang="en-US" altLang="zh-CN" i="1" smtClean="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>${italicText}</m:t></m:r>`

		case 'mn': // 数字 - 不斜体
			return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr lang="en-US" altLang="zh-CN" i="0" smtClean="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>${getTextContent(node)}</m:t></m:r>`

		case 'mo': // 运算符 - 不斜体
			return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr lang="en-US" altLang="zh-CN" i="0" smtClean="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>${getTextContent(node)}</m:t></m:r>`

		case 'mtext': // 文本 - 不斜体
			return `<m:r><m:rPr><m:nor/></m:rPr><a:rPr lang="en-US" altLang="zh-CN" i="0" smtClean="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t>${getTextContent(node)}</m:t></m:r>`

		case 'mspace': // 空格
			return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr lang="en-US" altLang="zh-CN" i="0" smtClean="0"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><m:t> </m:t></m:r>`

		case 'msup': { // 上标 x^2
			const supBase = node.childNodes[0]
			const sup = node.childNodes[1]
			return `<m:sSup><m:e>${convertMathNode(supBase)}</m:e><m:sup>${convertMathNode(sup)}</m:sup></m:sSup>`
		}

		case 'msub': { // 下标 x_i
			const subBase = node.childNodes[0]
			const sub = node.childNodes[1]
			return `<m:sSub><m:e>${convertMathNode(subBase)}</m:e><m:sub>${convertMathNode(sub)}</m:sub></m:sSub>`
		}

		case 'msubsup': { // 上下标 x_i^2
			const subsupBase = node.childNodes[0]
			const subsupSub = node.childNodes[1]
			const subsupSup = node.childNodes[2]
			return `<m:sSubSup><m:e>${convertMathNode(subsupBase)}</m:e><m:sub>${convertMathNode(subsupSub)}</m:sub><m:sup>${convertMathNode(subsupSup)}</m:sup></m:sSubSup>`
		}

		case 'mfrac': { // 分数 a/b
			const numerator = node.childNodes[0]
			const denominator = node.childNodes[1]
			return `<m:f><m:num>${convertMathNode(numerator)}</m:num><m:den>${convertMathNode(denominator)}</m:den></m:f>`
		}

		case 'msqrt': // 平方根 √x
			return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${convertChildren(node)}</m:e></m:rad>`

		case 'mroot': { // n次根 ³√x
			const radicand = node.childNodes[0]
			const index = node.childNodes[1]
			return `<m:rad><m:deg>${convertMathNode(index)}</m:deg><m:e>${convertMathNode(radicand)}</m:e></m:rad>`
		}

		case 'munder': { // 下方 (如 lim 下标)
			const underBase = node.childNodes[0]
			const under = node.childNodes[1]
			return `<m:limLow><m:e>${convertMathNode(underBase)}</m:e><m:lim>${convertMathNode(under)}</m:lim></m:limLow>`
		}

		case 'mover': { // 上方 (如向量箭头)
			const overBase = node.childNodes[0]
			const over = node.childNodes[1]
			const accent = element.getAttribute('accent') === 'true'
			if (accent) {
				return `<m:acc><m:accPr><m:chr m:val="${getTextContent(over)}"/></m:accPr><m:e>${convertMathNode(overBase)}</m:e></m:acc>`
			}
			return `<m:limUpp><m:e>${convertMathNode(overBase)}</m:e><m:lim>${convertMathNode(over)}</m:lim></m:limUpp>`
		}

		case 'munderover': { // 上下方 (如求和符号)
			const underoverBase = node.childNodes[0]
			const underoverUnder = node.childNodes[1]
			const underoverOver = node.childNodes[2]
			return `<m:limLow><m:e><m:limUpp><m:e>${convertMathNode(underoverBase)}</m:e><m:lim>${convertMathNode(underoverOver)}</m:lim></m:limUpp></m:e><m:lim>${convertMathNode(underoverUnder)}</m:lim></m:limLow>`
		}

		case 'mrow': // 行/分组
			return convertChildren(node)

		case 'mfenced': { // 括号
			const open = element.getAttribute('open') || '('
			const close = element.getAttribute('close') || ')'
			return `<m:d><m:dPr><m:begChr m:val="${escapeXml(open)}"/><m:endChr m:val="${escapeXml(close)}"/></m:dPr><m:e>${convertChildren(node)}</m:e></m:d>`
		}

		case 'mtable': // 矩阵/表格
			return `<m:m>${convertChildren(node)}</m:m>`

		case 'mtr': // 矩阵行
			return `<m:mr>${convertChildren(node)}</m:mr>`

		case 'mtd': // 矩阵单元格
			return `<m:e>${convertChildren(node)}</m:e>`

		case 'mstyle': // 样式
			return convertChildren(node)

		case 'semantics': // 语义
			// 只处理第一个子节点 (presentation MathML)
			return node.childNodes[0] ? convertMathNode(node.childNodes[0]) : ''

		case 'annotation':
		case 'annotation-xml':
			// 忽略注释
			return ''

		default:
			// 未处理的标签，递归处理子节点
			return convertChildren(node)
	}
}

/**
 * 步骤2: MathML → OMML
 * Step 2: MathML → OMML
 * @param mathML - MathML XML string
 * @returns OMML XML string
 */
export function mathMLToOMML(mathML: string): string {
	try {
		const parser = new DOMParser()
		const doc = parser.parseFromString(mathML, 'text/xml')
		const mathNode = doc.documentElement

		const omml = convertMathNode(mathNode)
		return omml
	} catch (error) {
		console.error(`MathML conversion failed: ${(error as Error).message}`)
		throw error
	}
}

/**
 * 完整转换: LaTeX → OMML
 * Complete conversion: LaTeX → OMML
 * @param latex - LaTeX formula string
 * @returns OMML XML string
 */
export function latexToOMML(latex: string): string {
	const mathML = latexToMathML(latex)
	const omml = mathMLToOMML(mathML)
	return omml
}
