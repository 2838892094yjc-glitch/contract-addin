/**
 * AI 输出解析模块
 * 用于解析 AI 的 JSON 输出并转换为表单配置格式
 */

// ==================== 常量定义 ====================

// 有效的 formatFn 白名单（与 formatters.js 保持一致）
const VALID_FORMAT_FNS = [
  'none',
  'dateUnderline',
  'dateYearMonth',
  'chineseNumber',
  'chineseNumberWan',
  'amountWithChinese',
  'articleNumber',
  'percentageChinese',
];

// 有效的表单类型
const VALID_TYPES = [
  'text',
  'number',
  'date',
  'select',
  'radio',
  'textarea',
];

// 有效的 mode
const VALID_MODES = ['insert', 'paragraph'];

// ==================== 验证函数 ====================

/**
 * 验证单个变量的格式
 * @param {object} variable - AI 输出的变量对象
 * @returns {object} { valid: boolean, errors: string[] }
 */
function validateVariable(variable) {
  const errors = [];
  
  // 必填字段检查
  const requiredFields = ['context', 'placeholder', 'label', 'tag', 'type', 'formatFn', 'mode'];
  for (const field of requiredFields) {
    if (!(field in variable) || variable[field] === null || variable[field] === undefined || variable[field] === '') {
      errors.push(`缺少必填字段: ${field}`);
    }
  }
  
  // prefix 和 suffix 可以为空字符串，但必须存在
  if (!('prefix' in variable)) {
    errors.push(`缺少字段: prefix`);
  }
  if (!('suffix' in variable)) {
    errors.push(`缺少字段: suffix`);
  }
  
  // type 验证
  if (variable.type && !VALID_TYPES.includes(variable.type)) {
    errors.push(`无效的 type: ${variable.type}，有效值: ${VALID_TYPES.join(', ')}`);
  }
  
  // formatFn 验证
  if (variable.formatFn && !VALID_FORMAT_FNS.includes(variable.formatFn)) {
    errors.push(`无效的 formatFn: ${variable.formatFn}，有效值: ${VALID_FORMAT_FNS.join(', ')}`);
  }
  
  // mode 验证
  if (variable.mode && !VALID_MODES.includes(variable.mode)) {
    errors.push(`无效的 mode: ${variable.mode}，有效值: ${VALID_MODES.join(', ')}`);
  }
  
  // select/radio 必须有 options
  if ((variable.type === 'select' || variable.type === 'radio') && (!variable.options || !Array.isArray(variable.options) || variable.options.length === 0)) {
    errors.push(`type 为 ${variable.type} 时必须提供 options 数组`);
  }
  
  return {
    valid: errors.length === 0,
    errors
  };
}

/**
 * 验证 AI 输出的整体格式
 * @param {object} aiOutput - AI 返回的 JSON 对象
 * @returns {object} { valid: boolean, errors: string[] }
 */
function validateAIOutput(aiOutput) {
  const errors = [];
  
  // 检查是否有 variables 字段
  if (!aiOutput || typeof aiOutput !== 'object') {
    errors.push('AI 输出不是有效的 JSON 对象');
    return { valid: false, errors };
  }
  
  if (!('variables' in aiOutput)) {
    errors.push('AI 输出缺少 variables 字段');
    return { valid: false, errors };
  }
  
  if (!Array.isArray(aiOutput.variables)) {
    errors.push('variables 必须是数组');
    return { valid: false, errors };
  }
  
  // 验证每个变量
  aiOutput.variables.forEach((variable, index) => {
    const validation = validateVariable(variable);
    if (!validation.valid) {
      errors.push(`变量 ${index + 1} (${variable.label || '未知'}): ${validation.errors.join(', ')}`);
    }
  });
  
  return {
    valid: errors.length === 0,
    errors
  };
}

// ==================== 转换函数 ====================

/**
 * 生成唯一 ID
 * @param {string} tag - 拼音 tag
 * @returns {string} 小写的 ID
 */
function generateFieldId(tag) {
  return tag.charAt(0).toLowerCase() + tag.slice(1);
}

/**
 * 将单个 AI 变量转换为表单字段配置
 * @param {object} variable - AI 输出的变量
 * @returns {object} 表单字段配置
 */
function convertVariableToField(variable) {
  const field = {
    id: generateFieldId(variable.tag),
    label: variable.label,
    tag: variable.tag,
    type: variable.type,
    formatFn: variable.formatFn !== 'none' ? variable.formatFn : null,
    placeholder: variable.placeholder,
    
    // 保存原始的 context、prefix、suffix 用于埋点
    _aiContext: {
      context: variable.context,
      prefix: variable.prefix,
      placeholder: variable.placeholder,
      suffix: variable.suffix,
      layer: variable.layer || 1,
      confidence: variable.confidence || 'medium',
      reason: variable.reason || ''
    }
  };
  
  // 添加 options（如果有）
  if (variable.options && variable.options.length > 0) {
    field.options = variable.options;
  }
  
  // 处理 paragraph 模式
  if (variable.mode === 'paragraph') {
    field.hasParagraphToggle = true;
  }
  
  return field;
}

/**
 * 按 layer 和语义分组变量
 * @param {array} variables - AI 输出的变量数组
 * @returns {object} 分组后的变量 { sections: [...] }
 */
function groupVariablesBySection(variables) {
  const sections = new Map();
  
  // 默认分组：基础信息
  const defaultSectionId = 'ai_recognized_fields';
  const defaultSection = {
    id: defaultSectionId,
    header: { 
      label: 'AI 识别的变量', 
      tag: 'Section_AIRecognized' 
    },
    fields: []
  };
  
  // 将所有变量放入默认 section
  variables.forEach(variable => {
    const field = convertVariableToField(variable);
    defaultSection.fields.push(field);
  });
  
  return {
    sections: [defaultSection]
  };
}

/**
 * 解析 AI 输出并转换为表单配置
 * @param {object} aiOutput - AI 返回的 JSON 对象
 * @param {object} options - 选项 { validateOnly: boolean }
 * @returns {object} { success: boolean, config?: object, errors?: string[], warnings?: string[] }
 */
function parseAIOutput(aiOutput, options = {}) {
  const { validateOnly = false } = options;
  
  // 验证
  const validation = validateAIOutput(aiOutput);
  
  if (!validation.valid) {
    console.error('[AI Parser] 验证失败:', validation.errors);
    return {
      success: false,
      errors: validation.errors
    };
  }
  
  if (validateOnly) {
    return {
      success: true,
      errors: [],
      warnings: []
    };
  }
  
  // 转换
  try {
    const { sections } = groupVariablesBySection(aiOutput.variables);
    
    // 生成统计信息
    const stats = {
      total: aiOutput.variables.length,
      byLayer: {
        1: aiOutput.variables.filter(v => v.layer === 1).length,
        2: aiOutput.variables.filter(v => v.layer === 2).length,
        3: aiOutput.variables.filter(v => v.layer === 3).length,
      },
      byMode: {
        insert: aiOutput.variables.filter(v => v.mode === 'insert').length,
        paragraph: aiOutput.variables.filter(v => v.mode === 'paragraph').length,
      },
      byType: {}
    };
    
    VALID_TYPES.forEach(type => {
      stats.byType[type] = aiOutput.variables.filter(v => v.type === type).length;
    });
    
    console.log('[AI Parser] 解析成功，统计信息:', stats);
    
    return {
      success: true,
      config: sections,
      stats,
      warnings: []
    };
  } catch (error) {
    console.error('[AI Parser] 转换失败:', error);
    return {
      success: false,
      errors: [`转换失败: ${error.message}`]
    };
  }
}

/**
 * 记录未知格式到日志
 * @param {object} aiOutput - AI 输出
 */
function logUnknownFormats(aiOutput) {
  if (!aiOutput || !aiOutput.variables) return;
  
  const unknownFormats = [];
  
  aiOutput.variables.forEach((variable, index) => {
    if (variable.formatFn && !VALID_FORMAT_FNS.includes(variable.formatFn)) {
      unknownFormats.push({
        index: index + 1,
        label: variable.label,
        tag: variable.tag,
        formatFn: variable.formatFn,
        context: variable.context
      });
    }
  });
  
  if (unknownFormats.length > 0) {
    console.warn('[AI Parser] 发现未知的 formatFn:');
    console.table(unknownFormats);
    
    // 保存到 localStorage 供后续分析
    try {
      const existingLogs = JSON.parse(localStorage.getItem('ai_parser_unknown_formats') || '[]');
      existingLogs.push({
        timestamp: new Date().toISOString(),
        unknownFormats
      });
      // 只保留最近 100 条
      if (existingLogs.length > 100) {
        existingLogs.splice(0, existingLogs.length - 100);
      }
      localStorage.setItem('ai_parser_unknown_formats', JSON.stringify(existingLogs));
    } catch (e) {
      console.error('[AI Parser] 无法保存日志:', e);
    }
  }
}

/**
 * 从 AI 输出生成埋点信息
 * @param {object} aiOutput - AI 输出
 * @returns {array} 埋点信息数组
 */
function generateEmbedInfo(aiOutput) {
  if (!aiOutput || !aiOutput.variables) return [];
  
  return aiOutput.variables.map(variable => ({
    tag: variable.tag,
    label: variable.label,
    context: variable.context,
    prefix: variable.prefix,
    placeholder: variable.placeholder,
    suffix: variable.suffix,
    mode: variable.mode,
    type: variable.type,
    formatFn: variable.formatFn,
    layer: variable.layer,
    confidence: variable.confidence
  }));
}

// ==================== 导出 ====================

// 兼容浏览器环境
if (typeof window !== 'undefined') {
  window.AIParser = {
    parseAIOutput,
    validateAIOutput,
    validateVariable,
    logUnknownFormats,
    generateEmbedInfo,
    convertVariableToField,
    VALID_FORMAT_FNS,
    VALID_TYPES,
    VALID_MODES
  };
}

// 兼容 Node.js 环境
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseAIOutput,
    validateAIOutput,
    validateVariable,
    logUnknownFormats,
    generateEmbedInfo,
    convertVariableToField,
    VALID_FORMAT_FNS,
    VALID_TYPES,
    VALID_MODES
  };
}
