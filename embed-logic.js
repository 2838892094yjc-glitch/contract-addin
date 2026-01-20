/**
 * 埋点逻辑模块
 * 基于 AI 输出精确定位 placeholder 并创建 Content Control
 */

// ==================== 精确搜索与定位 ====================

/**
 * 在文档中精确搜索文本
 * @param {Word.Document} doc - Word 文档对象
 * @param {string} searchText - 要搜索的文本
 * @param {string} context - 上下文（用于验证）
 * @returns {Promise<Word.RangeCollection>} 搜索结果
 */
async function searchInDocument(doc, searchText, context = null) {
    return await Word.run(async (wordContext) => {
        const searchResults = doc.body.search(searchText, {
            matchCase: false,
            matchWholeWord: false,
            matchWildcards: false
        });
        
        wordContext.load(searchResults, 'items');
        await wordContext.sync();
        
        console.log(`[Embed] 搜索 "${searchText.substring(0, 50)}..." 找到 ${searchResults.items.length} 处`);
        
        return searchResults;
    });
}

/**
 * 精确定位 placeholder（使用 context 验证）
 * @param {Word.Document} doc - Word 文档对象
 * @param {object} variable - AI 输出的变量对象
 * @returns {Promise<Word.Range|null>} 找到的 Range 对象
 */
async function locatePlaceholder(doc, variable) {
    const { context, prefix, placeholder, suffix } = variable;
    
    return await Word.run(async (wordContext) => {
        // 策略 1: 搜索完整 context
        let searchResults = doc.body.search(context, {
            matchCase: false,
            matchWholeWord: false
        });
        
        wordContext.load(searchResults, 'items');
        await wordContext.sync();
        
        if (searchResults.items.length === 0) {
            console.warn(`[Embed] 未找到 context: ${context.substring(0, 100)}`);
            
            // 策略 2: 尝试搜索 prefix + placeholder + suffix
            const fullText = prefix + placeholder + suffix;
            searchResults = doc.body.search(fullText, {
                matchCase: false,
                matchWholeWord: false
            });
            
            wordContext.load(searchResults, 'items');
            await wordContext.sync();
            
            if (searchResults.items.length === 0) {
                console.warn(`[Embed] 未找到完整文本: ${fullText.substring(0, 100)}`);
                
                // 策略 3: 尝试只搜索 placeholder（如果足够独特）
                if (placeholder.length > 3 && placeholder !== '____') {
                    searchResults = doc.body.search(placeholder, {
                        matchCase: false,
                        matchWholeWord: false
                    });
                    
                    wordContext.load(searchResults, 'items');
                    await wordContext.sync();
                    
                    if (searchResults.items.length > 0) {
                        console.log(`[Embed] 通过 placeholder 找到 ${searchResults.items.length} 处`);
                        return searchResults.items[0];
                    }
                }
                
                return null;
            }
        }
        
        // 找到了 context，现在定位 placeholder
        const contextRange = searchResults.items[0];
        
        // 计算 placeholder 在 context 中的位置
        const placeholderStart = prefix.length;
        const placeholderEnd = placeholderStart + placeholder.length;
        
        // 创建子 Range
        const placeholderRange = contextRange.getRange('Start');
        placeholderRange.moveStart('Character', placeholderStart);
        placeholderRange.moveEnd('Character', placeholderEnd - placeholderStart);
        
        await wordContext.sync();
        
        console.log(`[Embed] 成功定位 placeholder: ${variable.label}`);
        return placeholderRange;
    });
}

// ==================== Content Control 创建 ====================

/**
 * 在指定 Range 创建 Content Control
 * @param {Word.Range} range - 目标 Range
 * @param {object} variable - AI 输出的变量对象
 * @returns {Promise<Word.ContentControl>} 创建的 Content Control
 */
async function createContentControl(range, variable) {
    return await Word.run(async (wordContext) => {
        const contentControl = range.insertContentControl();
        
        // 设置属性
        contentControl.tag = variable.tag;
        contentControl.title = variable.label;
        contentControl.appearance = 'Tags';
        contentControl.color = '#D1E8FF';
        contentControl.cannotDelete = false;
        contentControl.cannotEdit = false;
        
        await wordContext.sync();
        
        console.log(`[Embed] 创建 Content Control: ${variable.label} (${variable.tag})`);
        return contentControl;
    });
}

/**
 * 基于 AI 输出创建单个 Content Control
 * @param {Word.Document} doc - Word 文档对象
 * @param {object} variable - AI 输出的变量对象
 * @returns {Promise<boolean>} 是否成功
 */
async function embedVariable(doc, variable) {
    try {
        // 1. 定位 placeholder
        const range = await locatePlaceholder(doc, variable);
        
        if (!range) {
            console.error(`[Embed] 无法定位变量: ${variable.label}`);
            return false;
        }
        
        // 2. 创建 Content Control
        await createContentControl(range, variable);
        
        return true;
    } catch (error) {
        console.error(`[Embed] 埋点失败 (${variable.label}):`, error);
        return false;
    }
}

/**
 * 批量埋点（基于 AI 输出）
 * @param {Word.Document} doc - Word 文档对象
 * @param {array} variables - AI 输出的变量数组
 * @param {object} options - 选项 { onProgress: function }
 * @returns {Promise<object>} { success: number, failed: number, errors: array }
 */
async function embedAllVariables(doc, variables, options = {}) {
    const { onProgress } = options;
    
    const results = {
        success: 0,
        failed: 0,
        errors: []
    };
    
    console.log(`[Embed] 开始批量埋点，共 ${variables.length} 个变量`);
    
    for (let i = 0; i < variables.length; i++) {
        const variable = variables[i];
        
        try {
            const success = await embedVariable(doc, variable);
            
            if (success) {
                results.success++;
            } else {
                results.failed++;
                results.errors.push({
                    variable: variable.label,
                    tag: variable.tag,
                    error: '定位失败'
                });
            }
            
            if (onProgress) {
                onProgress(i + 1, variables.length, variable.label);
            }
        } catch (error) {
            results.failed++;
            results.errors.push({
                variable: variable.label,
                tag: variable.tag,
                error: error.message
            });
        }
    }
    
    console.log(`[Embed] 埋点完成: 成功 ${results.success}, 失败 ${results.failed}`);
    
    if (results.errors.length > 0) {
        console.table(results.errors);
    }
    
    return results;
}

// ==================== 段落处理 ====================

/**
 * 处理可选段落（paragraph mode）
 * @param {Word.Document} doc - Word 文档对象
 * @param {object} variable - AI 输出的变量对象
 * @returns {Promise<boolean>} 是否成功
 */
async function embedParagraph(doc, variable) {
    return await Word.run(async (wordContext) => {
        // 搜索段落开头
        const searchResults = doc.body.search(variable.context, {
            matchCase: false,
            matchWholeWord: false
        });
        
        wordContext.load(searchResults, 'items');
        await wordContext.sync();
        
        if (searchResults.items.length === 0) {
            console.error(`[Embed] 未找到段落: ${variable.label}`);
            return false;
        }
        
        const paragraphRange = searchResults.items[0];
        
        // 扩展到整个段落
        const paragraph = paragraphRange.paragraphs.getFirst();
        wordContext.load(paragraph);
        await wordContext.sync();
        
        // 创建 Content Control 包裹整个段落
        const contentControl = paragraph.insertContentControl();
        contentControl.tag = variable.tag;
        contentControl.title = variable.label;
        contentControl.appearance = 'Tags';
        contentControl.color = '#FFE8D1';  // 使用不同颜色标识段落
        
        await wordContext.sync();
        
        console.log(`[Embed] 创建段落 Content Control: ${variable.label}`);
        return true;
    });
}

// ==================== 工具函数 ====================

/**
 * 检查 Content Control 是否已存在
 * @param {Word.Document} doc - Word 文档对象
 * @param {string} tag - Content Control tag
 * @returns {Promise<boolean>} 是否存在
 */
async function contentControlExists(doc, tag) {
    return await Word.run(async (wordContext) => {
        const contentControls = doc.contentControls.getByTag(tag);
        wordContext.load(contentControls, 'items');
        await wordContext.sync();
        
        return contentControls.items.length > 0;
    });
}

/**
 * 获取所有 Content Control
 * @param {Word.Document} doc - Word 文档对象
 * @returns {Promise<array>} Content Control 列表
 */
async function getAllContentControls(doc) {
    return await Word.run(async (wordContext) => {
        const contentControls = doc.contentControls;
        wordContext.load(contentControls, 'tag, title, text');
        await wordContext.sync();
        
        return contentControls.items.map(cc => ({
            tag: cc.tag,
            title: cc.title,
            text: cc.text
        }));
    });
}

// ==================== 导出 ====================

// 兼容浏览器环境
if (typeof window !== 'undefined') {
    window.EmbedLogic = {
        embedVariable,
        embedAllVariables,
        embedParagraph,
        locatePlaceholder,
        createContentControl,
        searchInDocument,
        contentControlExists,
        getAllContentControls
    };
}

// 兼容 Node.js 环境
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        embedVariable,
        embedAllVariables,
        embedParagraph,
        locatePlaceholder,
        createContentControl,
        searchInDocument,
        contentControlExists,
        getAllContentControls
    };
}
