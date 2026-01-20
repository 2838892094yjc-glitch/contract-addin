/**
 * AI Skill 调用模块
 * 使用 Skill.md 作为提示词调用 AI API
 */

// ==================== 配置 ====================

// Skill.md 文件路径
const SKILL_MD_PATH = './contract-variable-skill/Skill.md';

// AI API 配置
const DOUBAO_API = {
    url: "https://ark.cn-beijing.volces.com/api/v3/bots/chat/completions",
    token: "25f59ab0-c652-4e2f-8b3f-f3fba0f2ee18",  // 请替换为你的 token
    model: "bot-20241216150614-gdg8f"
};

// 文档切片大小
const CHUNK_SIZE = 10000;  // 字符

// ==================== Skill 加载 ====================

let skillPrompt = null;

/**
 * 加载 Skill.md 文件内容
 * @returns {Promise<string>} Skill 提示词
 */
async function loadSkillPrompt() {
    if (skillPrompt) {
        return skillPrompt;
    }
    
    try {
        const response = await fetch(SKILL_MD_PATH);
        if (!response.ok) {
            throw new Error(`无法加载 Skill.md: ${response.status}`);
        }
        skillPrompt = await response.text();
        console.log('[AI Skill] Skill.md 加载成功，长度:', skillPrompt.length);
        return skillPrompt;
    } catch (error) {
        console.error('[AI Skill] 加载 Skill.md 失败:', error);
        // 返回一个简化版本的提示词
        return getDefaultPrompt();
    }
}

/**
 * 获取默认提示词（当 Skill.md 无法加载时使用）
 * @returns {string} 默认提示词
 */
function getDefaultPrompt() {
    return `你是一位资深合同起草专家。

你的任务是分析合同文本，识别其中的变量，并输出结构化 JSON。

输出格式：
{
  "variables": [
    {
      "context": "完整上下文",
      "prefix": "变量前固定文本",
      "placeholder": "需要埋点的部分",
      "suffix": "变量后固定文本",
      "label": "中文名称",
      "tag": "PinYinTag",
      "type": "text|number|date|select|radio|textarea",
      "options": ["选项1", "选项2"],
      "formatFn": "none",
      "mode": "insert|paragraph",
      "layer": 1|2|3
    }
  ]
}

只返回 JSON，不要其他任何内容。`;
}

// ==================== 文档切片 ====================

/**
 * 将长文档切分为多个片段
 * @param {string} text - 文档文本
 * @param {number} chunkSize - 每个片段的大小
 * @returns {array} 文档片段数组
 */
function chunkDocument(text, chunkSize = CHUNK_SIZE) {
    if (text.length <= chunkSize) {
        return [text];
    }
    
    const chunks = [];
    let start = 0;
    
    while (start < text.length) {
        let end = start + chunkSize;
        
        // 如果不是最后一块，尝试在句子边界处切分
        if (end < text.length) {
            // 向后查找句号、问号、感叹号等标点
            const punctuation = /[。！？；\n]/g;
            const substr = text.substring(start, end + 500);  // 多看 500 字符
            let lastPuncIndex = -1;
            let match;
            
            while ((match = punctuation.exec(substr)) !== null) {
                lastPuncIndex = start + match.index + 1;
            }
            
            if (lastPuncIndex > start && lastPuncIndex < end + 500) {
                end = lastPuncIndex;
            }
        }
        
        chunks.push(text.substring(start, end));
        start = end;
    }
    
    console.log(`[AI Skill] 文档切分为 ${chunks.length} 个片段`);
    return chunks;
}

// ==================== AI 调用 ====================

/**
 * 调用 AI API（单个片段）
 * @param {string} documentText - 文档文本
 * @param {number} chunkIndex - 当前片段索引
 * @param {number} totalChunks - 总片段数
 * @returns {Promise<object>} AI 返回的 JSON 对象
 */
async function callAIWithSkill(documentText, chunkIndex = 0, totalChunks = 1) {
    console.log(`[AI Skill] 开始调用 AI... (分块 ${chunkIndex + 1}/${totalChunks})`);
    console.log(`[AI Skill] 文档长度: ${documentText.length} 字符`);
    
    // 加载 Skill 提示词
    const skillContent = await loadSkillPrompt();
    
    // 构建完整提示词
    const chunkInfo = totalChunks > 1 
        ? `\n\n【注意】这是文档的第 ${chunkIndex + 1}/${totalChunks} 部分。` 
        : '';
    
    const fullPrompt = `${skillContent}

---

【合同文本】

${documentText}
${chunkInfo}

---

请按照上述 Skill 要求，分析合同文本并输出 JSON。
只返回 JSON，不要其他任何内容。`;
    
    try {
        const response = await fetch(DOUBAO_API.url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${DOUBAO_API.token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: DOUBAO_API.model,
                input: [{ 
                    role: "user", 
                    content: [{ type: "input_text", text: fullPrompt }] 
                }]
            })
        });
        
        if (!response.ok) {
            const errText = await response.text();
            throw new Error(`API 请求失败: ${response.status} - ${errText}`);
        }
        
        const result = await response.json();
        console.log("[AI Skill] API 响应:", result);
        
        // 解析响应内容
        let outputText = "";
        if (result.output && Array.isArray(result.output)) {
            const messageObj = result.output.find(o => o.type === 'message');
            
            if (messageObj && Array.isArray(messageObj.content)) {
                let textContent = messageObj.content.find(c => c.type === 'text' || c.type === 'output_text');
                
                if (textContent && textContent.text) {
                    outputText = textContent.text;
                }
            }
        } else if (result.output && result.output.content) {
            outputText = result.output.content;
        } else if (result.choices && result.choices[0] && result.choices[0].message) {
            outputText = result.choices[0].message.content;
        }
        
        console.log("[AI Skill] 原始输出:", outputText);
        
        // 提取 JSON
        const jsonMatch = outputText.match(/\{[\s\S]*"variables"[\s\S]*\[[\s\S]*\][\s\S]*\}/);
        if (!jsonMatch) {
            console.warn("[AI Skill] 无法从响应中提取 JSON");
            return { variables: [] };
        }
        
        try {
            const jsonData = JSON.parse(jsonMatch[0]);
            
            // 验证格式
            if (!window.AIParser) {
                console.warn('[AI Skill] AIParser 未加载，跳过验证');
                return jsonData;
            }
            
            const validation = window.AIParser.validateAIOutput(jsonData);
            if (!validation.valid) {
                console.error('[AI Skill] AI 输出验证失败:', validation.errors);
                // 记录未知格式
                window.AIParser.logUnknownFormats(jsonData);
            } else {
                console.log('[AI Skill] AI 输出验证通过');
            }
            
            return jsonData;
        } catch (parseError) {
            console.error('[AI Skill] JSON 解析失败:', parseError);
            console.error('[AI Skill] 原始文本:', jsonMatch[0].substring(0, 500));
            return { variables: [] };
        }
    } catch (error) {
        console.error('[AI Skill] API 调用失败:', error);
        throw error;
    }
}

/**
 * 分析完整文档（支持自动切片）
 * @param {string} fullText - 完整文档文本
 * @returns {Promise<object>} 合并后的 AI 输出
 */
async function analyzeDocument(fullText) {
    console.log('[AI Skill] 开始分析文档...');
    
    // 切分文档
    const chunks = chunkDocument(fullText, CHUNK_SIZE);
    
    if (chunks.length === 1) {
        // 单个片段，直接调用
        return await callAIWithSkill(chunks[0], 0, 1);
    }
    
    // 多个片段，并行调用
    console.log(`[AI Skill] 文档被切分为 ${chunks.length} 个片段，开始并行处理...`);
    
    try {
        const results = await Promise.all(
            chunks.map((chunk, index) => callAIWithSkill(chunk, index, chunks.length))
        );
        
        // 合并结果
        const mergedVariables = [];
        results.forEach(result => {
            if (result && result.variables) {
                mergedVariables.push(...result.variables);
            }
        });
        
        console.log(`[AI Skill] 合并完成，共识别 ${mergedVariables.length} 个变量`);
        
        // 去重（根据 tag）
        const uniqueVariables = deduplicateVariables(mergedVariables);
        
        return {
            variables: uniqueVariables
        };
    } catch (error) {
        console.error('[AI Skill] 文档分析失败:', error);
        throw error;
    }
}

/**
 * 去重变量（根据 tag）
 * @param {array} variables - 变量数组
 * @returns {array} 去重后的变量数组
 */
function deduplicateVariables(variables) {
    const seen = new Map();
    const unique = [];
    
    variables.forEach(variable => {
        if (!variable.tag) {
            unique.push(variable);
            return;
        }
        
        if (seen.has(variable.tag)) {
            // 已存在，根据 confidence 和 layer 决定是否替换
            const existing = seen.get(variable.tag);
            
            // 优先保留 confidence 更高的
            const confidencePriority = { high: 3, medium: 2, low: 1 };
            const existingConfidence = confidencePriority[existing.confidence || 'medium'];
            const newConfidence = confidencePriority[variable.confidence || 'medium'];
            
            if (newConfidence > existingConfidence || 
                (newConfidence === existingConfidence && (variable.layer || 1) < (existing.layer || 1))) {
                // 替换
                const index = unique.indexOf(existing);
                unique[index] = variable;
                seen.set(variable.tag, variable);
            }
        } else {
            seen.set(variable.tag, variable);
            unique.push(variable);
        }
    });
    
    const deduped = unique.length;
    const original = variables.length;
    if (deduped < original) {
        console.log(`[AI Skill] 去重：${original} → ${deduped} (移除 ${original - deduped} 个重复项)`);
    }
    
    return unique;
}

// ==================== 导出 ====================

// 兼容浏览器环境
if (typeof window !== 'undefined') {
    window.AISkill = {
        analyzeDocument,
        callAIWithSkill,
        loadSkillPrompt,
        chunkDocument,
        deduplicateVariables
    };
}

// 兼容 Node.js 环境
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        analyzeDocument,
        callAIWithSkill,
        loadSkillPrompt,
        chunkDocument,
        deduplicateVariables
    };
}
