# 合同变量识别 Skill

## 概述

这是一个基于 Claude AI 的合同变量智能识别系统，采用**三层识别架构**，能够自动识别合同中的变量并进行精确埋点。

## 核心特性

### 1. 三层识别架构

- **Layer 1**: 明显占位符扫描（`【】`、`____`、日期格式等）
- **Layer 2**: 隐式变量识别（具体公司名、金额、日期等）
- **Layer 3**: 可选段落分析（条件性条款、可选条款）

### 2. 精确埋点

- 精确定位 `placeholder`，不破坏周围固定文本（如单位、标签）
- 支持 `prefix`、`placeholder`、`suffix` 三段分离
- 使用 `context` 确保唯一定位

### 3. 格式化支持

只在需要转换格式的场景使用格式化函数：

| formatFn | 用途 |
|----------|------|
| `dateUnderline` | 日期转 `2024年01月15日` |
| `chineseNumber` | 数字转 `壹佰（100）` |
| `chineseNumberWan` | 数字转 `壹佰（100）万元` |
| `articleNumber` | 数字转 `第五条` |
| ...更多 | 见 `formatters.js` |

### 4. 多表单类型支持

- `text`: 单行文本
- `number`: 数字
- `date`: 日期选择器
- `select`: 下拉单选
- `radio`: 单选按钮
- `textarea`: 多行文本

## 文件结构

```
contract-variable-skill/
├── Skill.md                    # Claude Skill 定义（AI 提示词）
├── README.md                   # 本文件
├── INTEGRATION_TEST.md         # 集成测试指南
└── (未来可能添加)
    ├── examples/               # 示例合同
    └── test-results/           # 测试结果

contract-addin/
├── formatters.js               # 格式化函数模块
├── ai-parser.js                # AI 输出解析模块
├── ai-skill.js                 # AI Skill 调用模块
├── embed-logic.js              # 埋点逻辑模块
└── taskpane.js                 # 主程序（原有）
```

## 使用方法

### 方式 1: 在浏览器控制台中使用

1. 打开 Word 插件
2. 打开浏览器控制台（F12）
3. 运行以下代码：

```javascript
// 完整流程
async function autoRecognizeAndEmbed() {
    // 1. 读取文档
    let documentText;
    await Word.run(async (context) => {
        const body = context.document.body;
        context.load(body, 'text');
        await context.sync();
        documentText = body.text;
    });
    
    // 2. AI 识别
    const aiOutput = await window.AISkill.analyzeDocument(documentText);
    
    // 3. 验证
    const validation = window.AIParser.validateAIOutput(aiOutput);
    if (!validation.valid) {
        console.error('验证失败:', validation.errors);
        return;
    }
    
    // 4. 埋点
    await Word.run(async (context) => {
        const doc = context.document;
        const result = await window.EmbedLogic.embedAllVariables(doc, aiOutput.variables);
        console.log('埋点完成:', result);
    });
}

// 执行
autoRecognizeAndEmbed();
```

### 方式 2: 集成到 taskpane.js

在 `autoGenerateForm()` 函数中使用新模块：

```javascript
async function autoGenerateForm() {
    try {
        // 读取文档
        const documentText = await getDocumentText();
        
        // 调用 AI Skill
        const aiOutput = await window.AISkill.analyzeDocument(documentText);
        
        // 验证
        const validation = window.AIParser.validateAIOutput(aiOutput);
        if (!validation.valid) {
            showError('AI 识别失败: ' + validation.errors.join(', '));
            return;
        }
        
        // 解析
        const parseResult = window.AIParser.parseAIOutput(aiOutput);
        if (!parseResult.success) {
            showError('解析失败: ' + parseResult.errors.join(', '));
            return;
        }
        
        // 埋点
        await Word.run(async (context) => {
            const doc = context.document;
            const embedResult = await window.EmbedLogic.embedAllVariables(
                doc, 
                aiOutput.variables,
                {
                    onProgress: (current, total, label) => {
                        updateProgress(current, total, label);
                    }
                }
            );
            
            if (embedResult.failed > 0) {
                console.warn('部分埋点失败:', embedResult.errors);
            }
        });
        
        // 生成表单
        updateFormConfig(parseResult.config);
        buildForm();
        
    } catch (error) {
        console.error('自动生成失败:', error);
        showError('生成失败: ' + error.message);
    }
}
```

## API 参考

### window.AISkill

```javascript
// 分析完整文档
const aiOutput = await window.AISkill.analyzeDocument(documentText);

// 单独调用 AI（单个片段）
const result = await window.AISkill.callAIWithSkill(text, chunkIndex, totalChunks);

// 加载 Skill 提示词
const skillPrompt = await window.AISkill.loadSkillPrompt();
```

### window.AIParser

```javascript
// 解析 AI 输出
const parseResult = window.AIParser.parseAIOutput(aiOutput);
// { success: boolean, config?: object, errors?: string[] }

// 验证 AI 输出
const validation = window.AIParser.validateAIOutput(aiOutput);
// { valid: boolean, errors: string[] }

// 记录未知格式
window.AIParser.logUnknownFormats(aiOutput);

// 生成埋点信息
const embedInfo = window.AIParser.generateEmbedInfo(aiOutput);
```

### window.EmbedLogic

```javascript
// 批量埋点
const result = await window.EmbedLogic.embedAllVariables(doc, variables, {
    onProgress: (current, total, label) => { ... }
});
// { success: number, failed: number, errors: array }

// 单个变量埋点
const success = await window.EmbedLogic.embedVariable(doc, variable);

// 段落埋点
const success = await window.EmbedLogic.embedParagraph(doc, variable);

// 检查 Content Control 是否存在
const exists = await window.EmbedLogic.contentControlExists(doc, tag);

// 获取所有 Content Controls
const controls = await window.EmbedLogic.getAllContentControls(doc);
```

### window.Formatters

```javascript
// 应用格式化
const formatted = window.Formatters.applyFormat(value, formatFn);

// 验证 formatFn
const valid = window.Formatters.isValidFormatFn(formatFn);

// 获取所有有效的 formatFn
const allFormatFns = window.Formatters.getValidFormatFns();
```

## AI 输出格式

```json
{
  "variables": [
    {
      "context": "甲方（委托方）：____",
      "prefix": "甲方（委托方）：",
      "placeholder": "____",
      "suffix": "",
      "label": "甲方名称",
      "tag": "JiaFangMingCheng",
      "type": "text",
      "options": [],
      "formatFn": "none",
      "mode": "insert",
      "layer": 1,
      "confidence": "high",
      "reason": "明显的公司名称占位符"
    }
  ]
}
```

## 测试

参见 [INTEGRATION_TEST.md](./INTEGRATION_TEST.md) 了解如何测试完整流程。

## 常见问题

### Q: AI 识别不准确怎么办？

A: 可以通过以下方式改进：
1. 优化 `Skill.md` 中的提示词
2. 添加更多示例
3. 调整三层识别的规则

### Q: 埋点失败怎么办？

A: 检查：
1. `context` 是否足够精确
2. `placeholder` 是否与文档实际内容匹配
3. 文档中是否有多个相同的占位符

### Q: 如何添加新的格式化函数？

A: 
1. 在 `formatters.js` 中添加新函数
2. 更新 `FORMAT_FUNCTIONS` 映射表
3. 在 `Skill.md` 中添加说明
4. 更新 `ai-parser.js` 中的 `VALID_FORMAT_FNS`

## 版本历史

- **v1.0** (2026-01-20): 初始版本
  - 三层识别架构
  - 8 个格式化函数
  - 精确埋点逻辑
  - 完整的验证和错误处理

## 贡献指南

如果你想改进这个系统：

1. 测试：使用真实合同测试，记录问题
2. 优化：改进 Skill.md 或代码逻辑
3. 扩展：添加新的格式化函数或表单类型
4. 文档：更新 README 和测试文档

## 许可证

内部使用，保密。
