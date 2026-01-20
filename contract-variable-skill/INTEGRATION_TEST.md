# 合同变量识别 Skill - 集成测试指南

## 测试目标

验证完整的合同变量识别和埋点流程：
1. AI Skill 能够正确识别变量
2. AI Parser 能够正确解析输出
3. Embed Logic 能够精确埋点
4. Formatters 能够正确格式化

## 测试环境准备

### 1. 启动开发服务器

```bash
cd /Users/yangjingchi/Desktop/contract-addin
node server.js
```

### 2. 在 Word Online 中打开插件

1. 打开 Word Online
2. 插入 → 加载项 → 我的加载项
3. 选择"合同起草助手"

### 3. 准备测试文档

使用参考目录下的真实合同：
- `/Users/yangjingchi/Desktop/参考/委托加工合同-通用模板 20241205.docx`
- `/Users/yangjingchi/Desktop/参考/石榴汁原液采购协议_20251101（终版）.docx`

## 测试流程

### 测试 1: AI 识别测试

#### 步骤

1. 打开测试合同文档
2. 打开浏览器控制台（F12）
3. 运行以下代码：

```javascript
// 读取文档文本
await Word.run(async (context) => {
    const body = context.document.body;
    context.load(body, 'text');
    await context.sync();
    
    console.log('文档长度:', body.text.length);
    
    // 调用 AI Skill
    const aiOutput = await window.AISkill.analyzeDocument(body.text);
    
    console.log('AI 识别结果:', aiOutput);
    console.log('识别变量数:', aiOutput.variables.length);
    
    // 验证输出
    const validation = window.AIParser.validateAIOutput(aiOutput);
    console.log('验证结果:', validation);
    
    // 保存结果供后续使用
    window.testAIOutput = aiOutput;
});
```

#### 预期结果

- ✅ AI 成功返回 JSON 格式输出
- ✅ `variables` 数组不为空
- ✅ 验证通过（`validation.valid === true`）
- ✅ 所有变量包含必填字段（context, prefix, placeholder, suffix, label, tag, type, formatFn, mode）

#### 检查点

| 检查项 | 通过标准 | 实际结果 |
|--------|---------|---------|
| Layer 1 识别 | 识别出明显占位符（如 `【】`、`____`） | |
| Layer 2 识别 | 识别出隐式变量（如具体公司名、日期） | |
| Layer 3 识别 | 识别出可选段落（mode: paragraph） | |
| formatFn 准确性 | 大多数是 `none`，只在需要转换时使用其他函数 | |
| prefix/suffix 分离 | placeholder 不包含周围固定文字 | |

---

### 测试 2: 解析与转换测试

#### 步骤

```javascript
// 解析 AI 输出
const parseResult = window.AIParser.parseAIOutput(window.testAIOutput);

console.log('解析结果:', parseResult);
console.log('成功:', parseResult.success);
console.log('表单配置:', parseResult.config);
console.log('统计:', parseResult.stats);

// 保存配置供后续使用
window.testFormConfig = parseResult.config;
```

#### 预期结果

- ✅ `parseResult.success === true`
- ✅ `parseResult.config` 包含 sections 数组
- ✅ 每个 section 包含 fields 数组
- ✅ 统计信息正确（byLayer, byMode, byType）

---

### 测试 3: 埋点测试

#### 步骤

```javascript
await Word.run(async (context) => {
    const doc = context.document;
    
    // 选择要测试的几个变量
    const testVariables = window.testAIOutput.variables.slice(0, 5);
    
    console.log('测试埋点，变量数:', testVariables.length);
    
    // 批量埋点
    const embedResult = await window.EmbedLogic.embedAllVariables(
        doc, 
        testVariables,
        {
            onProgress: (current, total, label) => {
                console.log(`进度: ${current}/${total} - ${label}`);
            }
        }
    );
    
    console.log('埋点结果:', embedResult);
    
    // 验证 Content Controls
    const allControls = await window.EmbedLogic.getAllContentControls(doc);
    console.log('创建的 Content Controls:', allControls);
});
```

#### 预期结果

- ✅ 成功率 > 80%
- ✅ Content Controls 被创建在正确位置
- ✅ 没有破坏周围固定文本
- ✅ tag 和 title 正确设置

#### 检查点

| 检查项 | 通过标准 | 实际结果 |
|--------|---------|---------|
| 定位准确性 | placeholder 被精确定位，不包含 prefix/suffix | |
| 固定文本保留 | 周围的单位、标签等固定文字保持不变 | |
| Content Control 创建 | 每个变量都有对应的 CC | |
| 属性设置 | tag、title 正确 | |

---

### 测试 4: 格式化测试

#### 步骤

```javascript
// 测试各种格式化函数
const testCases = [
    { formatFn: 'none', input: '张三', expected: '张三' },
    { formatFn: 'dateUnderline', input: '2024-01-15', expected: '2024年01月15日' },
    { formatFn: 'chineseNumber', input: 100, expected: '壹佰（100）' },
    { formatFn: 'chineseNumberWan', input: 100, expected: '壹佰（100）万元' },
    { formatFn: 'articleNumber', input: 5, expected: '第五条' },
    { formatFn: 'percentageChinese', input: 10, expected: '百分之十' },
];

console.log('格式化测试:');
testCases.forEach(test => {
    const result = window.Formatters.applyFormat(test.input, test.formatFn);
    const passed = result === test.expected;
    console.log(`${passed ? '✅' : '❌'} ${test.formatFn}: ${test.input} → ${result} (期望: ${test.expected})`);
});
```

#### 预期结果

- ✅ 所有测试用例通过

---

### 测试 5: 完整流程测试

#### 步骤

将以上所有步骤合并，测试完整流程：

```javascript
async function testFullWorkflow() {
    console.log('=== 开始完整流程测试 ===');
    
    try {
        // 1. 读取文档
        let documentText;
        await Word.run(async (context) => {
            const body = context.document.body;
            context.load(body, 'text');
            await context.sync();
            documentText = body.text;
            console.log('✅ Step 1: 读取文档成功，长度:', documentText.length);
        });
        
        // 2. AI 识别
        const aiOutput = await window.AISkill.analyzeDocument(documentText);
        console.log('✅ Step 2: AI 识别完成，变量数:', aiOutput.variables.length);
        
        // 3. 验证输出
        const validation = window.AIParser.validateAIOutput(aiOutput);
        if (!validation.valid) {
            console.error('❌ Step 3: 验证失败:', validation.errors);
            return;
        }
        console.log('✅ Step 3: 验证通过');
        
        // 4. 解析转换
        const parseResult = window.AIParser.parseAIOutput(aiOutput);
        if (!parseResult.success) {
            console.error('❌ Step 4: 解析失败:', parseResult.errors);
            return;
        }
        console.log('✅ Step 4: 解析成功');
        console.table(parseResult.stats);
        
        // 5. 批量埋点
        await Word.run(async (context) => {
            const doc = context.document;
            
            const embedResult = await window.EmbedLogic.embedAllVariables(
                doc, 
                aiOutput.variables,
                {
                    onProgress: (current, total, label) => {
                        if (current % 10 === 0 || current === total) {
                            console.log(`埋点进度: ${current}/${total}`);
                        }
                    }
                }
            );
            
            console.log('✅ Step 5: 埋点完成');
            console.log(`  成功: ${embedResult.success}`);
            console.log(`  失败: ${embedResult.failed}`);
            
            if (embedResult.errors.length > 0) {
                console.warn('  错误详情:');
                console.table(embedResult.errors);
            }
        });
        
        console.log('=== 完整流程测试完成 ===');
        
    } catch (error) {
        console.error('❌ 测试失败:', error);
    }
}

// 运行测试
testFullWorkflow();
```

#### 预期结果

- ✅ 所有步骤成功完成
- ✅ 埋点成功率 > 80%
- ✅ 文档中的 Content Controls 正确创建
- ✅ 用户可以在表单中填写并更新文档

---

## 测试记录

### 测试执行记录表

| 测试编号 | 测试名称 | 执行时间 | 测试人 | 结果 | 备注 |
|---------|---------|---------|--------|------|------|
| T1 | AI 识别测试 | | | | |
| T2 | 解析与转换测试 | | | | |
| T3 | 埋点测试 | | | | |
| T4 | 格式化测试 | | | | |
| T5 | 完整流程测试 | | | | |

### 问题记录

| 问题编号 | 发现时间 | 问题描述 | 影响等级 | 状态 | 解决方案 |
|---------|---------|---------|---------|------|---------|
| | | | | | |

---

## 常见问题排查

### 问题 1: AI 返回空结果

**可能原因**：
- Skill.md 文件未正确加载
- API token 过期
- 文档文本为空

**排查步骤**：
1. 检查浏览器控制台是否有加载错误
2. 验证 API token
3. 确认文档有内容

### 问题 2: 验证失败

**可能原因**：
- AI 输出了未知的 formatFn
- 缺少必填字段
- JSON 格式错误

**排查步骤**：
1. 查看 `validation.errors` 详细信息
2. 检查 localStorage 中的 `ai_parser_unknown_formats`
3. 更新 Skill.md 添加更多说明

### 问题 3: 埋点失败

**可能原因**：
- context 不够精确，无法唯一定位
- placeholder 与实际文档内容不匹配
- 文档中存在多个相同的占位符

**排查步骤**：
1. 查看 `embedResult.errors` 详细信息
2. 手动搜索文档验证 context 是否存在
3. 优化 AI prompt 提高 context 质量

---

## 性能基准

| 指标 | 目标值 | 实际值 |
|------|--------|--------|
| AI 识别时间（1万字） | < 10s | |
| 变量识别准确率 | > 90% | |
| 埋点成功率 | > 80% | |
| formatFn 准确率 | > 95% | |

---

## 下一步改进

基于测试结果，可以考虑以下改进：

1. [ ] 优化 Skill.md 提示词
2. [ ] 增加更多格式化函数
3. [ ] 改进埋点算法（支持模糊匹配）
4. [ ] 添加用户反馈机制
5. [ ] 支持批量处理多个文档
