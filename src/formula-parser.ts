/**
 * LaTeX 公式解析器
 * 自动识别文本中的 $...$ (行内公式) 和 $$...$$ (块级公式)
 */

export interface ParsedSegment {
  type: 'text' | 'formula';
  content: string;
  isBlock?: boolean; // 是否为块级公式 ($$...$$)
}

/**
 * 解析文本，识别其中的 LaTeX 公式
 * @param text 输入文本
 * @returns 解析后的片段数组
 */
export function parseTextWithFormulas(text: string): ParsedSegment[] {
  const segments: ParsedSegment[] = [];
  let currentPos = 0;

  // 正则表达式：匹配 $$...$$ 或 $...$
  // 注意：$$...$$ 必须先匹配，因为它包含 $...$
  const formulaRegex = /\$\$([^\$]+)\$\$|\$([^\$]+)\$/g;

  let match: RegExpExecArray | null;

  while ((match = formulaRegex.exec(text)) !== null) {
    // 添加公式前的普通文本
    if (match.index > currentPos) {
      const textContent = text.substring(currentPos, match.index);
      if (textContent) {
        segments.push({
          type: 'text',
          content: textContent
        });
      }
    }

    // 添加公式
    // match[1] 是 $$...$$ 的内容，match[2] 是 $...$ 的内容
    const formulaContent = match[1] || match[2];
    const isBlock = !!match[1]; // 如果 match[1] 存在，说明是 $$...$$

    segments.push({
      type: 'formula',
      content: formulaContent,
      isBlock: isBlock
    });

    currentPos = match.index + match[0].length;
  }

  // 添加剩余的普通文本
  if (currentPos < text.length) {
    const textContent = text.substring(currentPos);
    if (textContent) {
      segments.push({
        type: 'text',
        content: textContent
      });
    }
  }

  // 如果没有找到任何公式，返回整个文本作为普通文本
  if (segments.length === 0) {
    segments.push({
      type: 'text',
      content: text
    });
  }

  return segments;
}

/**
 * 检查文本是否包含 LaTeX 公式
 * @param text 输入文本
 * @returns 是否包含公式
 */
export function hasFormula(text: string): boolean {
  return /\$\$[^\$]+\$\$|\$[^\$]+\$/.test(text);
}
