// 合同起草助手 - 恢复表单/AI/生成文档 + 云端同步
// 说明：
//  - 恢复完整表单与 AI 逻辑（含 contractConfig、buildForm、handleAIFill、updateContent 等）
//  - 保留云端同步模块（MSAL + Graph）
//  - [2025-12-16] 回归"预设固定坑位"模式，移除不稳定动态数组；优化 UI 样式。
//  - [2026-01-08] 表单配置外部化：从 form-config.json 动态加载，支持统一编辑/拖拽

/* ==================================================================
 * 表单配置动态加载系统
 * ================================================================== */

// 配置存储 Key
const FORM_CONFIG_KEY = "contract_addin:formConfig";
const FORM_CONFIG_VERSION_KEY = "contract_addin:formConfigVersion";
const CURRENT_CONFIG_VERSION = "v20260108e"; // 配置版本号，更新时修改

// 表单配置数组（动态加载）
let contractConfig = [];

// 默认配置（仅作为备用，正常从 form-config.json 加载）
const DEFAULT_CONTRACT_CONFIG = [
    // -------------------- 1. 所需文件 --------------------
    {
        id: "section_files",
        header: { label: "1. 所需文件", tag: "Section_Files" },
        fields: [
            // 特殊字段：此 Section 的内容将由 buildForm 函数动态注入 Cloud Sync UI
            { type: "html_placeholder", targetId: "cloud-sync-section" }
        ]
    },

    // -------------------- 2. 公司基本信息 --------------------
    {
        id: "section_company_info",
        header: { label: "2. 公司基本信息", tag: "Section_CompanyInfo" },
        fields: [
            // --- 0. 签订时间与地点 ---
            { id: "signingDate", label: "签订时间", tag: "SigningDate", type: "date", formatFn: "dateUnderline", placeholder: "选择日期", hasParagraphToggle: true },
            { id: "signingPlace", label: "签订地点", tag: "SigningPlace", type: "text", placeholder: "如：北京" },
            
            // --- 1. 律师代表 ---
            { id: "lawyerRep", label: "律师代表", tag: "LawyerRepresenting", type: "radio", options: ["公司", "投资方", "公司/投资方"] },
            
            // --- 2. 基础信息 ---
            { id: "projectShortName", label: "项目简称", tag: "ProjectShortName", type: "text" },
            { id: "companyName", label: "目标公司名称", tag: "CompanyName", type: "text" },
            { id: "companyBusiness", label: "主营业务", tag: "CompanyBusiness", type: "text" },
            { id: "companyCapital", label: "注册资本", tag: "CompanyCapital", type: "text" },
            { id: "companyCity", label: "所在城市", tag: "CompanyCity", type: "text" },
            { id: "regAddress", label: "注册地址", tag: "RegAddress", type: "text" },
            { id: "legalRep", label: "法定代表人姓名", tag: "LegalRepName", type: "text" },
            { id: "legalRepTitle", label: "法定代表人职务", tag: "LegalRepTitle", type: "select", options: ["董事长", "执行董事", "总经理"] },
            { id: "legalRepNationality", label: "法定代表人国籍", tag: "LegalRepNationality", type: "select", options: ["中国", "美国", "新加坡", "其他"] },
            { id: "businessDesc", label: "主营业务描述", tag: "BusinessDesc", type: "text" },
            { id: "currentDirectors", label: "现任董事姓名", tag: "CurrentDirectors", type: "text", placeholder: "多个请用逗号隔开" },
            { 
                id: "shareholderCount", 
                label: "股东总数", 
                tag: "ShareholderCount", 
                type: "number", 
                value: "1",
                autoCount: true, // 特殊标记：自动统计已启用的股东数量
                placeholder: "系统自动统计，也可手动修改"
            },

            // --- 3. 股东 1 (创始人/大股东) - 必填 ---
            { type: "divider", label: "股东 1 (创始人/大股东)" },
            { id: "sh1_name", label: "姓名/名称", tag: "SH1_Name", type: "text" },
            { id: "sh1_type", label: "类型", tag: "SH1_Type", type: "select", options: ["个人", "有限公司", "合伙企业"] },
            { id: "sh1_id", label: "证件号码", tag: "SH1_ID", type: "text" },
            { id: "sh1_nation", label: "国籍/所在地", tag: "SH1_Nation", type: "text" },
            { id: "sh1_address", label: "注册地址", tag: "SH1_Address", type: "text" },
            { id: "sh1_reg_cap", label: "认缴注册资本(万元)", tag: "SH1_RegCapital", type: "number" },
            { id: "sh1_paid_cap", label: "实缴注册资本(万元)", tag: "SH1_PaidCapital", type: "number" },
            { id: "sh1_ratio", label: "持股比例/出资比例(%)", tag: "SH1_Ratio", type: "number" },
            { id: "sh1_currency", label: "币种", tag: "SH1_Currency", type: "select", options: ["人民币", "美元"] },
            { type: "divider", label: "增资后" },
            { id: "sh1_post_reg_cap", label: "增资后注册资本(万元)", tag: "SH1_PostRegCapital", type: "number" },
            { id: "sh1_post_ratio", label: "增资后持股比例(%)", tag: "SH1_PostRatio", type: "number" }
        ]
    },

    // -------------------- 其他现有股东/历轮投资人 (可选段落) - 归属于 Section 2 --------------------
    {
        id: "section_existing_shareholders",
        type: "existing_shareholders",
        header: { label: "2.1 现有股东/历轮投资人", tag: "Section_ExistingShareholders" },
        shareholders: [
            // 创始股东
            { id: "sh2", label: "创始股东 2", tag: "SH2" },
            // 种子轮
            { id: "sh3", label: "种子轮投资人 1", tag: "SH3" },
            { id: "sh4", label: "种子轮投资人 2", tag: "SH4" },
            // 天使轮
            { id: "sh5", label: "天使轮投资人 1", tag: "SH5" },
            { id: "sh6", label: "天使轮投资人 2", tag: "SH6" },
            // Pre-A轮
            { id: "sh7", label: "Pre-A轮投资人 1", tag: "SH7" },
            { id: "sh8", label: "Pre-A轮投资人 2", tag: "SH8" },
            // A轮
            { id: "sh9", label: "A轮投资人 1", tag: "SH9" },
            { id: "sh10", label: "A轮投资人 2", tag: "SH10" },
            // B轮
            { id: "sh11", label: "B轮投资人 1", tag: "SH11" },
            { id: "sh12", label: "B轮投资人 2", tag: "SH12" }
        ],
        shareholderFields: [
            { id: "_name", label: "姓名/名称", tag: "_Name", type: "text" },
            { id: "_short", label: "简称", tag: "_Short", type: "text" },
            { id: "_round", label: "融资轮次", tag: "_Round", type: "select", options: ["创始", "种子轮", "天使轮", "Pre-A轮", "A轮", "B轮", "C轮", "其他"] },
            { id: "_type", label: "类型", tag: "_Type", type: "select", options: ["个人", "有限公司", "有限合伙"], triggerConditional: true },
            { id: "_id", label: "证件号码", tag: "_ID", type: "text" },
            { id: "_nation", label: "国籍/所在地", tag: "_Nation", type: "text" },
            { id: "_address", label: "注册地址", tag: "_Address", type: "text" },
            { 
                id: "_legalRep", 
                label: "法定代表人", 
                tag: "_LegalRep", 
                paraTag: "_LegalRepPara",
                type: "text", 
                showWhen: ["有限公司", "有限合伙"], 
                hasParagraphToggle: true 
            },
            { id: "_investAmount", label: "投资额(万元)", tag: "_InvestAmount", type: "number", showWhenRound: ["种子轮", "天使轮", "Pre-A轮", "A轮", "B轮", "C轮", "其他"] },
            { id: "_regCapital", label: "认缴注册资本(万元)", tag: "_RegCapital", type: "number" },
            { id: "_paidCapital", label: "实缴注册资本(万元)", tag: "_PaidCapital", type: "number" },
            { id: "_ratio", label: "持股比例/出资比例(%)", tag: "_Ratio", type: "number" },
            { id: "_currency", label: "币种", tag: "_Currency", type: "select", options: ["人民币", "美元"] },
            { id: "_postRegCapital", label: "增资后注册资本(万元)", tag: "_PostRegCapital", type: "number" },
            { id: "_postRatio", label: "增资后持股比例(%)", tag: "_PostRatio", type: "number" }
        ]
    },

    // -------------------- 3. 本轮融资信息 --------------------
    {
        id: "section_financing",
        header: { label: "3. 本轮融资信息", tag: "Section_Financing" },
        fields: [
            // --- A. 增资前股权调整 ---
            {
                id: "needEquityAdjust",
                label: "增资前是否需要调整股权",
                tag: "NeedEquityAdjust",
                type: "radio", 
                options: ["否", "是"],
                subFields: [
                    { type: "divider", label: "股权调整事项 1" },
                    { id: "adj1_type", label: "调整方式", tag: "Adj1_Type", type: "select", options: ["转出", "增资", "减资"] },
                    { id: "adj1_transferor", label: "出让方/增资方", tag: "Adj1_Transferor", type: "text" },
                    { id: "adj1_transferee", label: "受让方", tag: "Adj1_Transferee", type: "text" },
                    { id: "adj1_price", label: "价格(万元)", tag: "Adj1_Price", type: "number" },
                    
                    { type: "divider", label: "股权调整事项 2" },
                    { id: "adj2_type", label: "调整方式", tag: "Adj2_Type", type: "select", options: ["转出", "增资", "减资"] },
                    { id: "adj2_transferor", label: "出让方/增资方", tag: "Adj2_Transferor", type: "text" },
                    { id: "adj2_transferee", label: "受让方", tag: "Adj2_Transferee", type: "text" },
                    { id: "adj2_price", label: "价格(万元)", tag: "Adj2_Price", type: "number" }
                ]
            },

            // --- C. 本次增资信息 ---
            { type: "divider", label: "本次增资" },
            { id: "investmentAmount", label: "投资款总额(万元)", tag: "InvestmentAmount", type: "number" },
            { id: "capitalIncrease", label: "计入注册资本(万元)", tag: "CapitalIncrease", type: "number" },
            { id: "capitalReserve", label: "计入资本公积金", tag: "CapitalReserve", type: "text", value: "剩余部分", placeholder: "填'剩余部分'或具体数额" },
            { id: "postCapitalTotal", label: "增资后总注册资本(万元)", tag: "PostCapitalTotal", type: "number" },
            { id: "newEquityRatio", label: "本次取得股权比例(%)", tag: "NewEquityRatio", type: "number" },
            
            // --- D. 基础融资条款 ---
            { type: "divider", label: "基础条款" },
            { id: "paymentDeadline", label: "最晚缴纳时间", tag: "PaymentDeadline", type: "date" }
        ]
    },

    // -------------------- 本轮投资人 (Section 3 子项) --------------------
    {
        id: "section_current_investors",
        type: "current_investors",
        header: { label: "3.1 本轮投资人", tag: "Section_CurrentInvestors" },
        investors: [
            { id: "lead", label: "领投方", tag: "Inv_Lead" },
            { id: "follow1", label: "跟投方 1", tag: "Inv_Follow1" },
            { id: "follow2", label: "跟投方 2", tag: "Inv_Follow2" },
            { id: "follow3", label: "跟投方 3", tag: "Inv_Follow3" }
        ],
        investorFields: [
            { id: "_name", label: "名称/姓名", tag: "_Name", type: "text" },
            { id: "_short", label: "简称", tag: "_Short", type: "text" },
            { id: "_type", label: "类型", tag: "_Type", type: "select", options: ["有限公司", "有限合伙", "个人"], triggerConditional: true },
            { id: "_nation", label: "注册地/国籍", tag: "_Nation", type: "text" },
            { id: "_address", label: "注册地址", tag: "_Address", type: "text" },
            { id: "_id", label: "证件号码", tag: "_ID", type: "text" },
            { id: "_legalRep", label: "法定代表人", tag: "_LegalRep", paraTag: "_LegalRepPara", type: "text", showWhen: ["有限公司", "有限合伙"], hasParagraphToggle: true },
            { id: "_amount", label: "投资额(万元)", tag: "_Amount", type: "number" },
            { id: "_currency", label: "币种", tag: "_Currency", type: "select", options: ["人民币", "美元"] },
            { id: "_equityRatio", label: "本次取得股权比例(%)", tag: "_EquityRatio", type: "number" },
            { id: "_regCapital", label: "本次对应注册资本(万元)", tag: "_RegCapital", type: "number" },
            { id: "_postRegCapital", label: "增资后注册资本(万元)", tag: "_PostRegCapital", type: "number" },
            { id: "_postRatio", label: "增资后持股比例(%)", tag: "_PostRatio", type: "number" }
        ]
    },

    // -------------------- 4. 定义及其他签约方 --------------------
    {
        id: "section_definitions",
        header: { label: "4. 定义及其他签约方", tag: "Section_Definitions" },
        fields: [
            { id: "otherParties", label: "其他签约方信息", tag: "OtherParties", type: "text", placeholder: "如有其他方请在此备注" }
        ]
    },

    // -------------------- 5. 创始人、新董事会、核心员工 --------------------
    {
        id: "section_board",
        header: { label: "5. 创始人、新董事会、核心员工", tag: "Section_Board" },
        fields: [
            { id: "newBoardSize", label: "新董事会由几名董事组成", tag: "NewBoardSize", type: "number" },
            { id: "investorBoardSeats", label: "本轮投资方有权任命董事人数", tag: "InvestorBoardSeats", type: "number" },
            { id: "founderBoardSeats", label: "创始人有权任命董事人数", tag: "FounderBoardSeats", type: "number" },
            { id: "founderHasOutsideEquity", label: "创始人是否持有集团外公司股权", tag: "FounderHasOutsideEquity", type: "radio", options: ["是", "否"] },
            // "nonCompetePromise" 移至 Section 10
            { id: "coreStaffList", label: "核心员工名单 (姓名/职务)", tag: "CoreStaffList", type: "text" }
        ]
    },

    // -------------------- 6. 特殊赔偿、交易费用、争议解决 --------------------
    {
        id: "section_indemnity",
        header: { label: "6. 特殊赔偿及其他", tag: "Section_Indemnity" },
        fields: [
            // --- 特殊赔偿 (全部使用插入段落模式) ---
            { id: "indemnity_social", label: "1. 社保/公积金未足额缴纳", tag: "Indemnity_SocialSecurity", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_tax", label: "2. 未足额缴纳税款/滞纳金", tag: "Indemnity_Tax", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_penalty", label: "3. 行政处罚或责任", tag: "Indemnity_Penalty", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_license", label: "4. 业务牌照/资质缺失", tag: "Indemnity_License", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_equity", label: "5. 股权权属纠纷", tag: "Indemnity_Equity", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_ip", label: "6. 知识产权侵权/权属不完善", tag: "Indemnity_IP", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_litigation", label: "7. 未决诉讼/仲裁", tag: "Indemnity_Litigation", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            { id: "indemnity_noncompete", label: "8. 核心员工违反竞业/保密义务", tag: "Indemnity_NonCompete", type: "radio", options: ["适用", "不适用"], hasParagraphToggle: true },
            
            // --- 责任限制 ---
            { type: "divider", label: "责任限制" },
            { id: "liability_threshold", label: "免责门槛金额(万元)", tag: "Liability_Threshold", type: "number", placeholder: "如：50", formatFn: "chineseNumber" },
            { id: "warranty_valid_years", label: "声明保证有效期(年)", tag: "Warranty_ValidYears", type: "number", placeholder: "如：4", formatFn: "chineseNumber" },
            
            // --- 交易费用 ---
            { type: "divider", label: "交易费用" },
            { 
                id: "fee_success", 
                label: "交易成功 - 公司承担费用", 
                tag: "Fee_Success", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true,
                subFields: [
                    { id: "fee_cap", label: "费用上限金额(万元)", tag: "FeeCap", type: "number", placeholder: "如：50" }
                ]
            },
            { 
                id: "fee_fail", 
                label: "交易终止 - 各方自担费用", 
                tag: "Fee_Fail", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // --- 争议解决 ---
            { id: "arbitrationOrg", label: "仲裁机构", tag: "ArbitrationOrg", type: "text", value: "中国国际经济贸易仲裁委员会" },
            { id: "arbitrationPlace", label: "仲裁地", tag: "ArbitrationPlace", type: "text", value: "北京" },
            { id: "hasTS", label: "是否签署投资意向书", tag: "HasTS", type: "radio", options: ["是", "否"] },
            { id: "tsDate", label: "意向书签署日期", tag: "TSDate", type: "date" }
        ]
    },

    // -------------------- 7. 股权变动限制 --------------------
    {
        id: "section_preemptive",
        header: { label: "7. 股权变动限制", tag: "Section_Preemptive" },
        fields: [
            // --- 现有股东转让限制 ---
            { id: "transfer_restricted_party", label: "被限制转让的主体", tag: "TransferRestrictedParty", type: "text", value: "创始股东", placeholder: "例如：创始股东、现有股东" },
            { id: "transfer_consent", label: "转让股权需经谁同意", tag: "TransferConsentSubject", type: "text", value: "本轮投资方" },
            { id: "transfer_consent_type", label: "同意形式", tag: "TransferConsentType", type: "text", value: "书面同意" },
            
            // --- 投资人转股权 (新增) ---
            { id: "investorTransferRight", label: "投资人是否可自由转股", tag: "InvestorTransferRight", type: "radio", options: ["是", "否"], value: "是" },
            
            // --- 优先认购权 ---
            { id: "hasPreemptiveRight", label: "新股优先认购权", tag: "HasPreemptiveRight", type: "radio", options: ["是", "否"] },
            { id: "preemptiveHolder", label: "优先认购权人", tag: "PreemptiveHolder", type: "text", value: "本轮投资方" },
            { id: "hasSuperPreemptive", label: "是否享有超额认购权", tag: "HasSuperPreemptive", type: "radio", options: ["是", "否"] },

            // --- 优先购买权 & 共售权 ---
            { id: "hasRofr", label: "老股优先购买权", tag: "HasRofr", type: "radio", options: ["是", "否"] },
            { id: "hasCoSale", label: "共同出售权", tag: "HasCoSale", type: "radio", options: ["是", "否"] },
            { id: "rofrHolder", label: "权利享有方", tag: "RofrHolder", type: "text", value: "本轮投资方" },
            
            // --- 领售权 ---
            { id: "hasDragAlong", label: "领售权 (拖售权)", tag: "HasDragAlong", type: "radio", options: ["是", "否"] },
            { id: "dragAlongTrigger", label: "领售触发条件", tag: "DragAlongTrigger", type: "text", placeholder: "例如：交割后 5 年未上市" },
            { id: "dragAlongValuation", label: "领售最低估值 (亿元)", tag: "DragAlongValuation", type: "number" }
        ]
    },

    // -------------------- 8. 核心经济条款 --------------------
    {
        id: "section_economics",
        header: { label: "8. 核心经济条款", tag: "Section_Economics" },
        fields: [
            // --- 反稀释 ---
            { 
                id: "antiDilution", 
                label: "反稀释权条款", 
                tag: "HasAntiDilution", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { id: "antiDilutionHolder", label: "反稀释权人", tag: "AntiDilutionHolder", type: "text", value: "本轮投资方" },
            { id: "antiDilutionOrigPrice", label: "本轮原始认购价格(元/注册资本)", tag: "AntiDilutionOrigPrice", type: "number", placeholder: "例如：10" },
            { 
                id: "antiDilutionMethod", 
                label: "价格调整方式", 
                tag: "AntiDilutionMethod", 
                type: "select", 
                options: ["广义加权平均", "完全棘轮", "狭义加权平均"]
            },
            { 
                id: "antiDilutionFormula", 
                label: "计算公式", 
                tag: "AntiDilutionFormula", 
                type: "select", 
                options: ["广义加权平均", "完全棘轮", "狭义加权平均"],
                valueMap: {
                    "广义加权平均": `按照广义加权平均的方式调整其原始认购价格，使得调整后的认购价格等于按如下公式确定的价格：

P2 = P1 × (A + B) / (A + C)

为上述公式之目的，各字母的含义如下：

P2为调整后的认购价格；

P1为原始认购价格；

A为公司新融资之前的注册资本总额（在完全稀释的基础上）；

B为假设公司新融资采用P1作为新认购价格的情况下，所增加或发行的注册资本数额；

C为公司新融资中实际增加或发行的注册资本数额。`,
                    "完全棘轮": `按照完全棘轮的方式调整其原始认购价格，使得调整后的认购价格等于触发反稀释的新融资中新增股东的新认购价格：

P2 = 新认购价格

即反稀释权人的原始认购价格将被调整至与本次新融资中新增股东的认购价格相同。`,
                    "狭义加权平均": `按照狭义加权平均的方式调整其原始认购价格，使得调整后的认购价格等于按如下公式确定的价格：

P2 = P1 × (A + B) / (A + C)

为上述公式之目的，各字母的含义如下：

P2为调整后的认购价格；

P1为原始认购价格；

A为反稀释权人在新融资之前持有的公司注册资本数额；

B为假设公司新融资采用P1作为新认购价格的情况下，反稀释权人按其持股比例应认购的注册资本数额；

C为按反稀释权人持股比例计算的公司新融资中实际增加或发行的注册资本数额。`
                }
            },
            { id: "antiDilutionCompDays", label: "补偿期限(天)", tag: "AntiDilutionCompDays", type: "number", value: "30", formatFn: "chineseNumber" },
            { id: "preemptiveClauseRef", label: "优先认购权条款编号", tag: "PreemptiveClauseRef", type: "text", placeholder: "例如：第5.1条" },

            // --- 优先清算权 ---
            { id: "liquidationPref", label: "清算优先权", tag: "HasLiquidationPref", type: "radio", options: ["是", "否"] },
            { id: "liqRanking", label: "是否优于普通股", tag: "LiqRanking", type: "radio", options: ["是", "否"] },
            { id: "liqMultiple", label: "优先清算回报倍数 (X倍本金)", tag: "LiqMultiple", type: "number", value: "1" },
            { id: "liqInterest", label: "清算年化利率 (%)", tag: "LiqInterest", type: "number", value: "0" },
            { id: "participationType", label: "剩余财产分配方式", tag: "ParticipationType", type: "select", options: ["无参与权(Non-participating)", "完全参与(Full participating)", "附上限参与(Capped)"] }
        ]
    },

    // -------------------- 8.1 回购权 (独立 Section，可整体插入段落) --------------------
    {
        id: "section_redemption",
        header: { label: "8.1 回购权", tag: "Section_Redemption" },
        hasSectionToggle: true, // 标记整个 Section 可以作为"插入段落"
        fields: [
            // --- 回购权整体开关 ---
            { 
                id: "hasRedemptionRight", 
                label: "回购权条款", 
                tag: "Section_Redemption", // 使用 Section 的 tag，控制整个回购权段落
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true // 选择"适用"显示段落，"不适用"隐藏段落
            },
            
            // --- 回购触发事件 (每个都是插入段落) ---
            { type: "divider", label: "回购触发事件" },
            { 
                id: "redemptionEvent_IPO", 
                label: "事件1: 未上市/退出失败", 
                tag: "RedemptionEvent_IPO", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true,
                subFields: [
                    { id: "redemptionTriggerYears", label: "触发年限(年)", tag: "RedemptionTriggerYears", type: "number", value: "6", formatFn: "chineseNumber" }
                ]
            },
            { 
                id: "redemptionEvent_Breach", 
                label: "事件2: 严重违反协议", 
                tag: "RedemptionEvent_Breach", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { 
                id: "redemptionEvent_Law", 
                label: "事件3: 严重违反法律法规", 
                tag: "RedemptionEvent_Law", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { 
                id: "redemptionEvent_Policy", 
                label: "事件4: 法律政策变化", 
                tag: "RedemptionEvent_Policy", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { 
                id: "redemptionEvent_Founder", 
                label: "事件5: 创始人/核心人员问题", 
                tag: "RedemptionEvent_Founder", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { 
                id: "redemptionEvent_Control", 
                label: "事件6: 实际控制人变更", 
                tag: "RedemptionEvent_Control", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { 
                id: "redemptionEvent_Business", 
                label: "事件7: 主营业务变更/经营异常", 
                tag: "RedemptionEvent_Business", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // --- 回购主体 ---
            { type: "divider", label: "回购主体" },
            { id: "redemptionRightHolder", label: "回购权人", tag: "RedemptionRightHolder", type: "text", value: "本轮投资方与投资方" },
            { id: "redemptionObligor", label: "回购义务人", tag: "RedemptionObligor", type: "text", value: "公司与创始股东" },
            { id: "redemptionClauseRef", label: "回购价格条款编号", tag: "RedemptionClauseRef", type: "text", value: "第3.2条" },
            
            // --- 回购价格计算 ---
            { type: "divider", label: "回购价格计算" },
            { 
                id: "redemptionPriceMode", 
                label: "价格计算模式", 
                tag: "RedemptionPriceMode", 
                type: "select", 
                options: ["单利(成本+回报)", "复利(成本+回报)", "固定倍数", "两者孰高(单利vs公允)"],
                hasParagraphToggle: true,
                valueMap: {
                    "单利(成本+回报)": `拟回购股权的回购价格（"回购价格"）应当按照以下公式计算：

回购价格 ＝ I × (1 + R × N) + A

I 为回购权人为获得拟回购股权实际支付的成本总额；

R 为回购利率，即【RedemptionInterestRate】%；

N 是一个分数，其分子为交割日至回购义务人向回购权人足额支付全部回购价格之日（"回购日"）之间所经过的天数，分母为365；

A 为回购日之前公司已宣布分配但尚未向该回购权人实际支付的拟回购股权对应的全部分红或股息。`,
                    "复利(成本+回报)": `拟回购股权的回购价格（"回购价格"）应当按照以下公式计算：

回购价格 ＝ I × (1 + R)^N + A

I 为回购权人为获得拟回购股权实际支付的成本总额；

R 为回购利率，即【RedemptionCompoundRate】%；

N 为交割日至回购日之间所经过的年数（不满一年的部分按实际天数/365计算）；

A 为回购日之前公司已宣布分配但尚未向该回购权人实际支付的拟回购股权对应的全部分红或股息。`,
                    "固定倍数": `拟回购股权的回购价格（"回购价格"）应当按照以下公式计算：

回购价格 ＝ I × Y% + A

I 为回购权人为获得拟回购股权实际支付的成本总额；

Y 为回购倍数，即【RedemptionMultiple】%；

A 为回购日之前公司已宣布分配但尚未向该回购权人实际支付的拟回购股权对应的全部分红或股息。`,
                    "两者孰高(单利vs公允)": `拟回购股权的回购价格（"回购价格"）应取以下两者之较高者：

（一）按以下公式计算的金额：I × (1 + R × N) + A

I 为回购权人为获得拟回购股权实际支付的成本总额；

R 为回购利率，即【RedemptionInterestRate】%；

N 是一个分数，其分子为交割日至回购日之间所经过的天数，分母为365；

A 为已宣布但尚未支付的分红或股息。

（二）拟回购股权届时的公允市场价值或对应的公司净资产价值。`
                },
                subFields: [
                    { id: "redemptionInterestRate", label: "单利年化利率(%)", tag: "RedemptionInterestRate", type: "number", value: "8" },
                    { id: "redemptionCompoundRate", label: "复利年化利率(%)", tag: "RedemptionCompoundRate", type: "number", value: "10" },
                    { id: "redemptionMultiple", label: "回购倍数(%)", tag: "RedemptionMultiple", type: "number", value: "150", placeholder: "如150表示1.5倍" }
                ]
            },
            
            // --- 期限与违约 ---
            { type: "divider", label: "期限与违约" },
            { id: "redemptionNotifyDays", label: "通知其他回购权人期限(工作日)", tag: "RedemptionNotifyDays", type: "number", value: "3", formatFn: "chineseNumber" },
            { id: "redemptionPaymentDays", label: "回购支付期限(日)", tag: "RedemptionPaymentDays", type: "number", value: "40", formatFn: "chineseNumber" },
            { id: "redemptionPenaltyRate", label: "违约金利率(每日万分之)", tag: "RedemptionPenaltyRate", type: "number", value: "5" },
            { id: "redemptionAssetSaleDays", label: "资产变卖触发期限(日)", tag: "RedemptionAssetSaleDays", type: "number", value: "90", formatFn: "chineseNumber" },
            
            // --- 回购顺序 ---
            { type: "divider", label: "回购顺序" },
            { id: "redemptionPriorityHolder", label: "第一顺位(优先支付方)", tag: "RedemptionPriorityHolder", type: "text", value: "本轮投资方" },
            { id: "redemptionSecondaryHolder", label: "第二顺位", tag: "RedemptionSecondaryHolder", type: "text", value: "投资方" },
            
            // --- 特殊限制条款 ---
            { type: "divider", label: "特殊限制条款" },
            { 
                id: "redemptionCompanyFirst", 
                label: "公司优先回购条款", 
                tag: "RedemptionCompanyFirst", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true,
                subFields: [
                    { id: "redemptionCompanyFirstDays", label: "公司履约期限(日)", tag: "RedemptionCompanyFirstDays", type: "number", value: "120", formatFn: "chineseNumber" }
                ]
            },
            { 
                id: "redemptionFounderCap", 
                label: "创始股东责任上限条款", 
                tag: "RedemptionFounderCap", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            { id: "redemptionDirectorRight", label: "保留董事权利类型", tag: "RedemptionDirectorRight", type: "select", options: ["委派", "观察", "提名"] }
        ]
    },

    // -------------------- 9. 其他优先权 --------------------
    {
        id: "section_other_rights",
        header: { label: "9. 其他优先权", tag: "Section_OtherRights" },
        fields: [
            // --- IPO 自动转换 ---
            { id: "ipo_auto_convert", label: "IPO自动转股机制", tag: "IPOAutoConvert", type: "radio", options: ["是", "否"] },
            { id: "ipo_min_valuation", label: "合格IPO最低估值 (亿元)", tag: "IPOMinValuation", type: "number", value: "40" },
            { id: "ipo_min_proceeds", label: "合格IPO最低募资额 (亿元)", tag: "IPOMinProceeds", type: "number", value: "10" },

            // --- 信息权 ---
            { id: "hasInfoRights", label: "信息权", tag: "HasInfoRights", type: "radio", options: ["是", "否"] },
            { id: "report_annual", label: "年度财报提供期限 (年后x天)", tag: "ReportDays_Annual", type: "number", value: "45", formatFn: "chineseNumber" },
            { id: "report_quarterly", label: "季度财报提供期限 (季后x天)", tag: "ReportDays_Quarterly", type: "number", value: "30", formatFn: "chineseNumber" },
            { id: "report_monthly", label: "月度财报提供期限 (月后x天)", tag: "ReportDays_Monthly", type: "number", value: "15", formatFn: "chineseNumber" },
            { id: "report_budget", label: "年度预算提供期限 (年后x天)", tag: "ReportDays_Budget", type: "number", value: "45", formatFn: "chineseNumber" },

            // --- 其他条款 ---
            { id: "hasMFN", label: "最优惠条款 (MFN)", tag: "HasMFN", type: "radio", options: ["是", "否"] },
            { id: "hasNewProjectRight", label: "新项目投资权 (创始人再创业)", tag: "HasNewProjectRight", type: "radio", options: ["是", "否"] }
        ]
    },

    // -------------------- 10. 其他文件 --------------------
    {
        id: "section_other_docs",
        header: { label: "10. 其他文件", tag: "Section_OtherDocs" },
        fields: [
            { id: "nonCompetePromise", label: "该公司是否应出具不竞争承诺函", tag: "NonCompetePromise", type: "radio", options: ["是", "否"] },
            { id: "ipTransferAgreement", label: "知识产权转让协议", tag: "IPTransferAgreement", type: "radio", options: ["适用", "不适用"] },
            { id: "shareTransferConfirm", label: "历史转股确认函", tag: "ShareTransferConfirm", type: "radio", options: ["适用", "不适用"] },
            { id: "nomineeAgreement", label: "代持协议", tag: "NomineeAgreement", type: "radio", options: ["适用", "不适用"] }
        ]
    },

    // -------------------- 11. 过桥贷款 --------------------
    {
        id: "section_bridge_loan",
        header: { label: "11. 过桥贷款", tag: "Section_BridgeLoan" },
        fields: [
            { id: "hasBridgeLoan", label: "是否签署过桥贷款协议", tag: "HasBridgeLoan", type: "radio", options: ["是", "否"],
              subFields: [
                  { id: "loanDocName", label: "意向书/贷款协议名称", tag: "LoanDocName", type: "text" },
                  { id: "loanDate", label: "签署日期", tag: "LoanDate", type: "date" },
                  { id: "loanAmount", label: "贷款金额 (万元)", tag: "LoanAmount", type: "number" },
                  { id: "loanTerm", label: "贷款期限 (月)", tag: "LoanTerm", type: "number" },
                  { id: "loanInterest", label: "年化利率 (%)", tag: "LoanInterest", type: "number", value: "0" },
                  { id: "overduePenalty", label: "逾期滞纳金比例 (每日千分之)", tag: "OverduePenalty", type: "number", value: "2" },
                  { id: "loanRepayType", label: "偿还方式", tag: "LoanRepayType", type: "select", options: ["债转股 (转换本金)", "现金偿还"] },
                  // 加速还款事件
                  { id: "event_breach", label: "事件1: 违反本协议义务/承诺", tag: "Event_BreachAgreement", type: "radio", options: ["适用", "不适用"] },
                  { id: "event_ts_breach", label: "事件2: 违反投资意向书(TS)", tag: "Event_BreachTS", type: "radio", options: ["适用", "不适用"] },
                  { id: "event_insolvency", label: "事件3: 无力偿还到期债务/破产", tag: "Event_Insolvency", type: "radio", options: ["适用", "不适用"] }
              ]
            }
        ]
    },

    // -------------------- 12. 声明保证与赔偿限制 --------------------
    {
        id: "section_reps",
        header: { label: "12. 声明保证与赔偿限制", tag: "Section_RepsWarranties" },
        fields: [
            { id: "repsSubject", label: "声明保证主体", tag: "RepsSubject", type: "select", options: ["公司", "创始股东", "公司及创始股东连带"] },
            
            // --- 核心声明条款 ---
            { id: "rep_existence", label: "1. 公司合法设立且有效存续", tag: "Rep_ValidExistence", type: "radio", options: ["确认", "有例外"] },
            { id: "rep_no_debt", label: "2. 无未披露的隐性债务/担保", tag: "Rep_NoUndisclosedDebt", type: "radio", options: ["确认", "有例外"] },
            { id: "rep_subsidiaries", label: "3. 已完整披露子公司/分公司结构", tag: "Rep_DisclosureSubsidiaries", type: "radio", options: ["确认", "有例外"] },
            { id: "rep_tax", label: "4. 已按时足额申报/缴纳税款", tag: "Rep_TaxCompliance", type: "radio", options: ["确认", "有例外"] },
            { id: "rep_litigation", label: "5. 无重大未决诉讼或仲裁", tag: "Rep_NoLitigation", type: "radio", options: ["确认", "有例外"] },

            // --- 赔偿限制 ---
            { id: "indemnity_de_minimis", label: "起赔额/免赔额 (万元)", tag: "IndemnityDeMinimis", type: "number", value: "50" },
            { id: "indemnity_cap_amount", label: "赔偿上限金额 (万元)", tag: "IndemnityCapAmount", type: "number", value: "50" },
            { id: "indemnity_cap_ratio", label: "赔偿上限比例 (投资款的%)", tag: "IndemnityCapRatio", type: "number", value: "100" },
            { id: "indemnity_time_limit", label: "索赔时效 (交割日后x年)", tag: "IndemnityTimeLimit", type: "number", value: "4" }
        ]
    },

    // -------------------- 13. 交割先决条件 --------------------
    {
        id: "section_cps",
        header: { label: "13. 交割先决条件", tag: "Section_CPs" },
        fields: [
            // 1. 声明与保证 - 使用插入段落模式（保留原格式）
            { 
                id: "cp_warranties", 
                label: "1. 声明与保证真实准确完整", 
                tag: "CP_Warranties", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 2. 签署交易文件 - 使用插入段落模式
            { 
                id: "cp_docs", 
                label: "2. 签署交易文件(股东协议+新章程)", 
                tag: "CP_SignDocs", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true,  // 标记为插入段落模式
                subFields: [
                    { id: "cp_articles_date", label: "公司章程签订日期", tag: "CP_ArticlesDate", type: "date" },
                    { id: "cp_sha_date", label: "股东协议签订日期", tag: "CP_SHADate", type: "date" }
                ]
            },
            
            // 3. 股东会批准 - 使用插入段落模式
            { 
                id: "cp_approval", 
                label: "3. 股东会批准本次交易", 
                tag: "CP_Approval", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true,
                subFields: [
                    { id: "cp_board_size", label: "董事会总人数", tag: "CP_BoardSize", type: "number", placeholder: "如：5" },
                    { id: "cp_founder_directors", label: "创始股东委派董事数", tag: "CP_FounderDirectors", type: "number", placeholder: "如：2" }
                ]
            },
            
            // 4. 工商变更登记 - 插入段落模式
            { 
                id: "cp_aic", 
                label: "4. 完成工商变更登记", 
                tag: "CP_AIC", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 5. 关键人员全职加入 - 插入段落模式
            { 
                id: "cp_key_personnel", 
                label: "5. 关键人员全职加入", 
                tag: "CP_KeyPersonnel", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true,
                subFields: [
                    { id: "cp_labor_term", label: "劳动合同最低期限(年)", tag: "CP_LaborTerm", type: "number", placeholder: "如：4", formatFn: "chineseNumber" }
                ]
            },
            
            // 6. 无重大不利变化 - 插入段落模式
            { 
                id: "cp_no_mac", 
                label: "6. 无重大不利变化(MAC)", 
                tag: "CP_NoMAC", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 7. 汇款通知 - 插入段落模式
            { 
                id: "cp_remittance", 
                label: "7. 发出汇款通知", 
                tag: "CP_Remittance", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 8. 交割条件满足通知 - 插入段落模式
            { 
                id: "cp_closing_notice", 
                label: "8. 交割条件满足通知", 
                tag: "CP_ClosingNotice", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 9. 投资委员会批准 - 插入段落模式
            { 
                id: "cp_ic_approval", 
                label: "9. 投资委员会批准", 
                tag: "CP_ICApproval", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 10. 尽职调查完成 - 插入段落模式
            { 
                id: "cp_dd", 
                label: "10. 尽职调查完成", 
                tag: "CP_DD", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 11. 创始人持股公司承诺函 - 插入段落模式
            { 
                id: "cp_founder_holdco", 
                label: "11. 创始人持股公司承诺函", 
                tag: "CP_FounderHoldco", 
                type: "radio", 
                options: ["适用", "不适用"],
                hasParagraphToggle: true
            },
            
            // 付款天数
            { id: "cp_payment_days", label: "先决条件满足后付款天数", tag: "CP_PaymentDays", type: "number", value: "10" }
        ]
    },

    // -------------------- 14. 各方承诺 --------------------
    {
        id: "section_covenants",
        header: { label: "14. 各方承诺", tag: "Section_Promises" },
        fields: [
            // --- 期限类字段 ---
            { type: "divider", label: "各项承诺期限" },
            { id: "promise_labor_contract", label: "签署劳动合同期限(月)", tag: "Promise_LaborContractMonths", type: "number", placeholder: "如：1", formatFn: "chineseNumber" },
            { id: "promise_ip_transfer", label: "无形资产转让期限(月)", tag: "Promise_IPTransferMonths", type: "number", placeholder: "如：1", formatFn: "chineseNumber" },
            { id: "promise_trademark_apply", label: "商标申请期限(月)", tag: "Promise_TrademarkApplyMonths", type: "number", placeholder: "如：3", formatFn: "chineseNumber" },
            { id: "promise_trademark_reg", label: "商标注册期限(月)", tag: "Promise_TrademarkRegMonths", type: "number", placeholder: "如：6", formatFn: "chineseNumber" },
            { id: "promise_other_ip", label: "其他无形资产申请期限(月)", tag: "Promise_OtherIPMonths", type: "number", placeholder: "如：3", formatFn: "chineseNumber" },
            { id: "promise_license_delivery", label: "营业执照交付期限(工作日)", tag: "Promise_LicenseDeliveryDays", type: "number", placeholder: "如：30", formatFn: "chineseNumber" },
            { id: "promise_aic_change", label: "工商变更期限(工作日)", tag: "Promise_AICChangeDays", type: "number", placeholder: "如：30", formatFn: "chineseNumber" },
            
            // --- 商标名称 ---
            { type: "divider", label: "商标信息" },
            { id: "promise_trademark_name1", label: "商标名称1", tag: "Promise_TrademarkName1", type: "text", placeholder: "如：公司品牌名" },
            { id: "promise_trademark_name2", label: "商标名称2", tag: "Promise_TrademarkName2", type: "text", placeholder: "如：产品名" },
            { id: "promise_trademark_name3", label: "商标名称3", tag: "Promise_TrademarkName3", type: "text", placeholder: "如：Logo名" },
            
            // --- 最惠国条款 ---
            { type: "divider", label: "最惠国条款" },
            { id: "promise_mfn_timing", label: "最惠国条款适用时间", tag: "Promise_MFNTiming", type: "select", options: ["完成前及完成后", "仅完成后"] }
        ]
    },

    // -------------------- 15. 重大事项否决权 --------------------
    {
        id: "section_veto",
        header: { label: "15. 重大事项否决权", tag: "Section_Veto" },
        fields: [
            { id: "veto_subject", label: "拥有一票否决权的主体", tag: "VetoSubject", type: "text", value: "本轮投资方" },
            
            // --- 否决事项列表 ---
            { id: "veto_cap_inc", label: "1. 增加注册资本/发行新股", tag: "Veto_IncreaseCapital", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_cap_dec", label: "2. 减少注册资本/回购股权", tag: "Veto_DecreaseCapital", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_structure", label: "3. 修改融资方案/股权结构", tag: "Veto_Structure", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_rights", label: "4. 修改股东权利/优先权", tag: "Veto_AmendRights", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_articles", label: "5. 修改公司章程", tag: "Veto_AmendArticles", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_board", label: "6. 变更董事会人数/产生方式", tag: "Veto_ChangeBoard", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_senior", label: "7. 聘用/解聘高管(CEO/CFO等)", tag: "Veto_SeniorMgmt", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_assets", label: "8. 重大资产出售/收购/许可", tag: "Veto_DisposeAssets", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_guarantee", label: "9. 对外担保/借款", tag: "Veto_Guarantees", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_related", label: "10. 关联交易", tag: "Veto_RelatedTx", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_dividend", label: "11. 利润分配/分红", tag: "Veto_Dividends", type: "radio", options: ["适用", "不适用"] },
            { id: "veto_ipo_ma", label: "12. 上市(IPO)或并购(M&A)方案", tag: "Veto_IPO_MA", type: "radio", options: ["适用", "不适用"] }
        ]
    }
];

/**
 * 加载表单配置（优先从 LocalStorage，否则从 JSON 文件）
 */
async function loadFormConfig() {
    console.log("[FormConfig] 开始加载配置...");
    
    // 1. 尝试从 LocalStorage 加载
    try {
        const savedVersion = localStorage.getItem(FORM_CONFIG_VERSION_KEY);
        const savedConfig = localStorage.getItem(FORM_CONFIG_KEY);
        
        if (savedConfig && savedVersion === CURRENT_CONFIG_VERSION) {
            contractConfig = JSON.parse(savedConfig);
            console.log("[FormConfig] 从 LocalStorage 加载配置，共", contractConfig.length, "个 sections");
            return true;
        } else if (savedConfig && savedVersion !== CURRENT_CONFIG_VERSION) {
            console.log("[FormConfig] 配置版本不匹配，将重新加载默认配置");
        }
    } catch (e) {
        console.warn("[FormConfig] LocalStorage 读取失败:", e.message);
    }
    
    // 2. 尝试从 JSON 文件加载
    try {
        const response = await fetch('form-config.json?v=' + Date.now());
        if (response.ok) {
            contractConfig = await response.json();
            console.log("[FormConfig] 从 form-config.json 加载配置，共", contractConfig.length, "个 sections");
            // 保存到 LocalStorage
            saveFormConfig();
            return true;
        }
    } catch (e) {
        console.warn("[FormConfig] JSON 文件加载失败:", e.message);
    }
    
    // 3. 使用默认配置
    contractConfig = JSON.parse(JSON.stringify(DEFAULT_CONTRACT_CONFIG));
    console.log("[FormConfig] 使用默认配置，共", contractConfig.length, "个 sections");
    return true;
}

/**
 * 保存表单配置到 LocalStorage
 */
function saveFormConfig() {
    try {
        localStorage.setItem(FORM_CONFIG_KEY, JSON.stringify(contractConfig));
        localStorage.setItem(FORM_CONFIG_VERSION_KEY, CURRENT_CONFIG_VERSION);
        console.log("[FormConfig] 配置已保存到 LocalStorage");
    } catch (e) {
        console.warn("[FormConfig] 保存失败:", e.message);
    }
}

/**
 * 重置表单配置为默认值
 */
async function resetFormConfig() {
    console.log("[FormConfig] 重置为默认配置...");
    localStorage.removeItem(FORM_CONFIG_KEY);
    localStorage.removeItem(FORM_CONFIG_VERSION_KEY);
    await loadFormConfig();
    // 重新构建表单
    buildForm();
    showNotification("表单配置已重置为默认值", "success");
}

/**
 * 导出完整表单配置
 */
function exportFullFormConfig() {
    const data = JSON.stringify(contractConfig, null, 2);
    const blob = new Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement("a");
    a.href = url;
    a.download = `form-config-${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    
    URL.revokeObjectURL(url);
    showNotification("完整表单配置已导出", "success");
}

/**
 * 导入完整表单配置
 */
function importFullFormConfig(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const imported = JSON.parse(e.target.result);
            if (!Array.isArray(imported)) {
                throw new Error("无效的配置格式");
            }
            
            contractConfig = imported;
            saveFormConfig();
            
            // 重新构建表单
            buildForm();
            renderCustomFieldsPanel();
            
            showNotification(`已导入配置，共 ${contractConfig.length} 个模块`, "success");
        } catch (err) {
            showNotification("导入失败: " + err.message, "error");
        }
    };
    reader.readAsText(file);
}

// ---------------- 防抖 ----------------
function debounce(func, wait) {
    let timeout;
    return function () {
        const context = this;
        const args = arguments;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), wait);
    };
}

// ---------------- 查找字段定义 ----------------
function findFieldById(fieldId, fields = null) {
    if (!fieldId) return null;
    if (!fields) {
        for (const section of contractConfig) {
            const found = findFieldById(fieldId, section.fields);
            if (found) return found;
        }
        return null;
    }
    for (const field of fields) {
        if (field.id === fieldId) return field;
        if (field.subFields) {
            const found = findFieldById(fieldId, field.subFields);
            if (found) return found;
        }
    }
    return null;
}

// =====================================================================
// 全局 Word 操作队列 (解决 forceSaveFailed 核心方案)
// 确保所有对文档的物理操作都是串行且带有缓冲间隔的
// =====================================================================
const wordActionQueue = {
    _queue: Promise.resolve(),
    isRunning: false,
    
    /**
     * 添加一个任务到队列
     * @param {Function} task 返回 Promise 的异步函数
     * @returns {Promise}
     */
    add(task) {
        const wrappedTask = async () => {
            this.isRunning = true;
            try {
                await task();
            } catch (err) {
                console.error("[Queue] Task failed:", err);
            } finally {
                // 增加更长的缓冲（2000ms），给 Word Online 后端充足的喘息时间
                await new Promise(r => setTimeout(r, 2000));
                this.isRunning = false;
            }
        };
        this._queue = this._queue.then(wrappedTask);
        return this._queue;
    }
};

// ---------------- 投资轮次启用状态 ----------------
let enabledRounds = {
    seed: false,
    angel: false,
    preA: false,
    seriesA: false,
    seriesB: false
};

// ---------------- 本轮投资人启用状态 ----------------
let enabledCurrentInvestors = {
    lead: false,
    follow1: false,
    follow2: false,
    follow3: false
};

// ---------------- 其他现有股东启用状态 ----------------
let enabledExistingShareholders = {
    sh2: false,
    sh3: false,
    sh4: false,
    sh5: false,
    sh6: false,
    sh7: false,
    sh8: false,
    sh9: false,
    sh10: false,
    sh11: false,
    sh12: false
};

// 标记用户是否手动修改过股东总数（一旦手动修改，停止自动更新）
let shareholderCountUserModified = false;

// 【新增】用于存放初始化时从 XML 加载的表单快照数据
let lastLoadedFormData = {};

// =====================================================================
// LocalStorage 实时同步引擎 (双轨道之一)
// =====================================================================
const LS_FORM_STATE_KEY = "contract_addin:formState";
const LS_ENABLED_ROUNDS_KEY = "contract_addin:enabledRounds";
const LS_ENABLED_INVESTORS_KEY = "contract_addin:enabledCurrentInvestors";
const LS_ENABLED_SHAREHOLDERS_KEY = "contract_addin:enabledExistingShareholders";

/**
 * 将当前表单状态写入 LocalStorage (实时同步)
 */
function saveFormStateToLocalStorage(formData, roundsState, investorsState, shareholdersState) {
    try {
        localStorage.setItem(LS_FORM_STATE_KEY, JSON.stringify(formData || {}));
        localStorage.setItem(LS_ENABLED_ROUNDS_KEY, JSON.stringify(roundsState || enabledRounds));
        localStorage.setItem(LS_ENABLED_INVESTORS_KEY, JSON.stringify(investorsState || enabledCurrentInvestors));
        localStorage.setItem(LS_ENABLED_SHAREHOLDERS_KEY, JSON.stringify(shareholdersState || enabledExistingShareholders));
        console.log("[LS] Form state saved to LocalStorage");
    } catch (e) {
        console.warn("[LS] Failed to save form state:", e);
    }
}

/**
 * 从 LocalStorage 读取表单状态
 * @returns {{ formData: object, enabledRounds: object, enabledCurrentInvestors: object, enabledExistingShareholders: object } | null}
 */
function loadFormStateFromLocalStorage() {
    try {
        const formDataStr = localStorage.getItem(LS_FORM_STATE_KEY);
        const roundsStr = localStorage.getItem(LS_ENABLED_ROUNDS_KEY);
        const investorsStr = localStorage.getItem(LS_ENABLED_INVESTORS_KEY);
        const shareholdersStr = localStorage.getItem(LS_ENABLED_SHAREHOLDERS_KEY);
        
        if (formDataStr || roundsStr || investorsStr || shareholdersStr) {
            const result = {
                formData: formDataStr ? JSON.parse(formDataStr) : {},
                enabledRounds: roundsStr ? JSON.parse(roundsStr) : {},
                enabledCurrentInvestors: investorsStr ? JSON.parse(investorsStr) : {},
                enabledExistingShareholders: shareholdersStr ? JSON.parse(shareholdersStr) : {}
            };
            console.log("[LS] Form state loaded from LocalStorage:", Object.keys(result.formData).length, "fields");
            return result;
        }
    } catch (e) {
        console.warn("[LS] Failed to load form state:", e);
    }
    return null;
}

// ---------------- 云端无感自动同步 ----------------
const LS_AUTO_SYNC = "contract_addin:autoSyncEnabled";
const LS_CLOUD_FOLDER = "contract_addin:cloudFolderPath";
let autoSyncEnabled = true;
let autoSyncInProgress = false;
let autoSyncPending = false;
let lastAutoSyncFingerprint = "";

function buildAutoSyncFingerprint(formData, selectedFileIds) {
    const ids = (selectedFileIds || []).slice().sort();
    return JSON.stringify({ data: formData, ids });
}

const scheduleAutoSync = debounce(function () {
    if (!autoSyncEnabled) return;
    const checked = document.querySelectorAll(".file-checkbox:checked");
    if (!checked || checked.length === 0) return;
    batchSyncFiles({ silent: true, reason: "auto" });
}, 1500);

// ---------------- 占位符自动应用到当前文档 ----------------
async function applyPlaceholderToCurrentDoc(formData) {
    return wordActionQueue.add(async () => {
    console.log("auto-apply keys:", Object.keys(formData));
    if (typeof Word === 'undefined') {
        console.log("[Mock] Apply placeholder to current doc");
        return;
    }
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            for (const [key, valRaw] of Object.entries(formData)) {
                const val = valRaw ?? "";
                const patterns = [`【${key}】`, `[${key}]`];
                let replaced = false;
                for (const pat of patterns) {
                    const results = body.search(pat, { matchCase: false, matchWildcards: false });
                    context.load(results, "text, font/name, font/size, font/color, font/bold, font/italic, font/underline");
                    await context.sync();
                    if (results.items.length > 0) {
                        for (const r of results.items) {
                            // 【保留格式】保存字体属性
                            const savedFont = {
                                name: r.font.name,
                                size: r.font.size,
                                color: r.font.color,
                                bold: r.font.bold,
                                italic: r.font.italic,
                                underline: r.font.underline
                            };
                            // 插入新文本
                            r.insertText(String(val), "Replace");
                            await context.sync();
                            // 恢复字体属性（对插入后的范围）
                            if (savedFont.name) r.font.name = savedFont.name;
                            if (savedFont.size) r.font.size = savedFont.size;
                            if (savedFont.color) r.font.color = savedFont.color;
                            if (savedFont.bold !== undefined) r.font.bold = savedFont.bold;
                            if (savedFont.italic !== undefined) r.font.italic = savedFont.italic;
                            if (savedFont.underline) r.font.underline = savedFont.underline;
                        }
                        await context.sync();
                        replaced = true;
                        break;
                    }
                }
            }
            await context.sync();
        });
    } catch (err) {
        console.warn("Apply placeholder to current doc failed:", err);
    }
    });
}

// ---------------- 收集表单数据 (递归平铺) ----------------
function collectFormData(skipLocalStorageSave = false) {
    const container = document.getElementById("dynamic-form-container");
    if (!container) return {};

    const result = {};

    function collectRecursive(parentEl) {
        const inputs = parentEl.querySelectorAll("input, select");
        inputs.forEach(input => {
            const tag = input.dataset.tag;
            if (!tag) return;

            let val = null;
            if (input.type === "radio") {
                if (input.checked) val = input.value;
            } else {
                val = input.value;
            }

            // 如果有值，或者是被选中的radio
            if (val !== null) {
                // 如果已存在且当前是未选中的radio，不覆盖
                // 但 querySelectorAll 是顺序遍历，radio 组通常只有一个 checked
                if (input.type === "radio" && !input.checked) return;
                result[tag] = val;
            }
        });
    }
    
    collectRecursive(container);
    
    // 【双轨同步】每次收集后，立即存入 LocalStorage
    if (!skipLocalStorageSave) {
        saveFormStateToLocalStorage(result, enabledRounds, enabledCurrentInvestors, enabledExistingShareholders);
    }
    
    return result;
}

// 自动应用
function autoApplyToCurrentDoc() {
    const data = collectFormData();
    applyPlaceholderToCurrentDoc(data);
}

// ---------------- 页面内通知 (替代 alert) ----------------
function showNotification(message, type = "info") {
    // 移除已有通知
    const existingNotif = document.getElementById("app-notification");
    if (existingNotif) existingNotif.remove();
    
    const notif = document.createElement("div");
    notif.id = "app-notification";
    notif.style.cssText = `
        position: fixed;
        top: 20px;
        left: 50%;
        transform: translateX(-50%);
        padding: 12px 20px;
        border-radius: 8px;
        font-size: 13px;
        font-weight: 500;
        z-index: 9999;
        max-width: 80%;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        animation: slideDown 0.3s ease;
    `;
    
    if (type === "error") {
        notif.style.background = "#fde8e8";
        notif.style.color = "#c53030";
        notif.style.border = "1px solid #fc8181";
    } else if (type === "warning") {
        notif.style.background = "#fefcbf";
        notif.style.color = "#744210";
        notif.style.border = "1px solid #f6e05e";
    } else if (type === "success") {
        notif.style.background = "#c6f6d5";
        notif.style.color = "#22543d";
        notif.style.border = "1px solid #68d391";
    } else {
        notif.style.background = "#bee3f8";
        notif.style.color = "#2a4365";
        notif.style.border = "1px solid #63b3ed";
    }
    
    notif.textContent = message;
    document.body.appendChild(notif);
    
    // 5秒后自动消失
    setTimeout(() => {
        if (notif.parentNode) {
            notif.style.opacity = "0";
            notif.style.transition = "opacity 0.3s";
            setTimeout(() => notif.remove(), 300);
        }
    }, 5000);
}

// ---------------- 自定义确认对话框 (替代 window.confirm) ----------------
function showConfirmDialog(message, options = {}) {
    return new Promise((resolve) => {
        // 移除已有对话框
        const existingDialog = document.getElementById("app-confirm-dialog");
        if (existingDialog) existingDialog.remove();
        
        const title = options.title || "确认操作";
        const confirmText = options.confirmText || "确定";
        const cancelText = options.cancelText || "取消";
        
        // 创建遮罩层
        const overlay = document.createElement("div");
        overlay.id = "app-confirm-dialog";
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 10000;
            display: flex;
            align-items: center;
            justify-content: center;
            animation: fadeIn 0.2s ease;
        `;
        
        // 创建对话框
        const dialog = document.createElement("div");
        dialog.style.cssText = `
            background: white;
            border-radius: 12px;
            padding: 24px;
            max-width: 400px;
            width: 90%;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
            animation: slideUp 0.3s ease;
        `;
        
        // 标题
        const titleEl = document.createElement("h3");
        titleEl.style.cssText = `
            margin: 0 0 16px 0;
            font-size: 16px;
            font-weight: 600;
            color: #333;
        `;
        titleEl.textContent = title;
        
        // 消息内容
        const messageEl = document.createElement("div");
        messageEl.style.cssText = `
            font-size: 13px;
            color: #555;
            line-height: 1.6;
            margin-bottom: 20px;
            white-space: pre-wrap;
        `;
        messageEl.textContent = message;
        
        // 按钮容器
        const btnContainer = document.createElement("div");
        btnContainer.style.cssText = `
            display: flex;
            gap: 12px;
            justify-content: flex-end;
        `;
        
        // 取消按钮
        const cancelBtn = document.createElement("button");
        cancelBtn.style.cssText = `
            padding: 10px 20px;
            border: 1px solid #ddd;
            background: #f5f5f5;
            color: #666;
            border-radius: 6px;
            font-size: 13px;
            cursor: pointer;
            font-weight: 500;
        `;
        cancelBtn.textContent = cancelText;
        cancelBtn.onclick = () => {
            overlay.remove();
            resolve(false);
        };
        
        // 确认按钮
        const confirmBtn = document.createElement("button");
        confirmBtn.style.cssText = `
            padding: 10px 20px;
            border: none;
            background: #0f6cbd;
            color: white;
            border-radius: 6px;
            font-size: 13px;
            cursor: pointer;
            font-weight: 500;
        `;
        confirmBtn.textContent = confirmText;
        confirmBtn.onclick = () => {
            overlay.remove();
            resolve(true);
        };
        
        btnContainer.appendChild(cancelBtn);
        btnContainer.appendChild(confirmBtn);
        
        dialog.appendChild(titleEl);
        dialog.appendChild(messageEl);
        dialog.appendChild(btnContainer);
        overlay.appendChild(dialog);
        
        document.body.appendChild(overlay);
        
        // 聚焦确认按钮
        confirmBtn.focus();
    });
}

// ---------------- 构建表单 (UI 美化 & 交互修复) ----------------
function buildForm() {
    const container = document.getElementById("dynamic-form-container");
    if (!container) return;
    container.innerHTML = "";

    // 注入美化样式 (Modern LegalTech Style)
    const style = document.createElement("style");
    style.textContent = `
        .section-header-container { 
            margin-top: 24px; 
            margin-bottom: 20px; 
            padding: 0 8px; 
            background: transparent;
            border-radius: 0;
            border: none;
            box-shadow: none;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .section-header-static { 
            font-size: 18px; 
            font-weight: 700; 
            color: var(--primary-color); 
            margin: 0;
            letter-spacing: -0.02em;
            position: relative;
        }
        .section-header-static::before {
            content: '';
            display: inline-block;
            width: 4px;
            height: 18px;
            background: var(--accent-color);
            margin-right: 12px;
            border-radius: 4px;
            vertical-align: middle;
        }
        
        /* 双列布局 */
        .section-fields { 
            padding-left: 0; 
            display: grid; 
            grid-template-columns: repeat(2, 1fr); 
            gap: 20px;
        }
        .section-fields .divider-line,
        .section-fields .form-group.full-width { 
            grid-column: 1 / -1; 
        }
        @media (max-width: 680px) {
            .section-fields { grid-template-columns: 1fr; }
        }
        
        /* 表单卡片 - 极简白底，去边框 */
        .form-group { 
            background: var(--card-bg); 
            padding: 32px; 
            border-radius: var(--radius-lg); 
            border: none; 
            box-shadow: var(--shadow-sm); 
            width: 100% !important;
            box-sizing: border-box !important;
            display: block !important;
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
        }
        .form-group:hover {
            box-shadow: var(--shadow-md);
            transform: translateY(-2px);
        }
        
        /* Label 美化 */
        .label-row label {
            font-size: 13px;
            font-weight: 500;
            color: var(--text-secondary);
            margin-bottom: 8px;
            display: block;
        }
        
        /* 输入框美化 */
        input[type="text"], input[type="number"], input[type="date"], select, textarea {
            background: var(--input-bg);
            border: 1px solid transparent;
            border-radius: var(--radius-md);
            padding: 12px 16px;
            font-size: 14px;
            color: var(--text-main);
            transition: all 0.2s ease;
            width: 100%;
            outline: none;
        }
        input:focus, select:focus, textarea:focus {
            background: #fff;
            box-shadow: 0 0 0 2px var(--accent-color);
        }
        /* 字段编辑按钮 */
        .form-group .field-edit-btn {
            position: absolute;
            top: 8px;
            right: 8px;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            border: 1px solid #e2e8f0;
            background: #f8fafc;
            cursor: pointer;
            opacity: 0;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
            color: #64748b;
            z-index: 10;
        }
        .form-group:hover .field-edit-btn {
            opacity: 1;
        }
        .form-group .field-edit-btn:hover {
            background: #2563eb;
            color: #fff;
            border-color: #2563eb;
        }
        /* 拖拽状态 */
        .form-group[draggable="true"] {
            cursor: grab;
            position: relative;
        }
        .form-group[draggable="true"]:active {
            cursor: grabbing;
        }
        /* 拖拽把手区域（左侧） */
        .form-group[draggable="true"]::before {
            content: "⋮⋮";
            position: absolute;
            left: -20px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 14px;
            color: #cbd5e1;
            cursor: grab;
            opacity: 0;
            transition: opacity 0.2s;
        }
        .form-group[draggable="true"]:hover::before {
            opacity: 1;
        }
        .form-group.dragging {
            opacity: 0.5;
            border: 2px dashed #2563eb;
        }
        /* 整个字段可拖拽（通过事件处理）*/
        .form-group[draggable="true"] .label-row {
            cursor: grab;
        }
        /* 放置区样式 */
        .drop-zone {
            height: 8px;
            margin: 4px 0;
            border-radius: 4px;
            background: transparent;
            transition: all 0.2s ease;
            display: none;
        }
        .drop-zone.drag-over {
            height: 40px;
            background: linear-gradient(135deg, #dbeafe 0%, #eff6ff 100%);
            border: 2px dashed #2563eb;
        }
        body.dragging-field .drop-zone {
            display: block;
        }
        
        .label-row { 
            display: flex; 
            align-items: center; 
            margin-bottom: 12px; 
            justify-content: space-between; 
        }
        .label-row label { 
            font-size: 13px; 
            font-weight: 600; 
            color: #334155;
        }
        
        /* 插入按钮圆弧化 */
        .insert-btn { 
            font-size: 11px; 
            padding: 4px 12px; 
            background: linear-gradient(135deg, #dbeafe 0%, #eff6ff 100%); 
            color: #2563eb; 
            border: 1px solid #bfdbfe; 
            border-radius: 20px; 
            cursor: pointer;
            font-weight: 600;
            transition: all 0.25s ease;
        }
        .insert-btn:hover { 
            background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%); 
            color: #fff; 
            border-color: #2563eb;
            transform: translateY(-1px);
        }
        
        /* 输入控件圆弧化 */
        .input-field { 
            width: 100% !important; 
            box-sizing: border-box !important; 
            padding: 11px 14px; 
            border: 1.5px solid #e2e8f0; 
            border-radius: 10px; 
            font-size: 14px; 
            background: #f8fafc;
            display: block !important;
            transition: all 0.2s ease;
        }
        .input-field:focus {
            outline: none;
            border-color: #2563eb;
            background: #fff;
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        
        /* 单选框圆弧化标签 */
        .radio-group { 
            display: flex; 
            flex-wrap: wrap;
            gap: 10px; 
            margin-top: 8px; 
            width: 100% !important;
        }
        .radio-label { 
            display: inline-flex !important; 
            align-items: center !important; 
            justify-content: flex-start !important; 
            font-size: 13px; 
            color: #475569;
            cursor: pointer;
            padding: 8px 14px;
            background: #f8fafc;
            border: 1.5px solid #e2e8f0;
            border-radius: 20px;
            transition: all 0.2s ease;
            gap: 8px;
        }
        .radio-label:hover {
            border-color: #2563eb;
            background: #eff6ff;
        }
        .radio-label:has(input:checked) {
            background: linear-gradient(135deg, #dbeafe 0%, #eff6ff 100%);
            border-color: #2563eb;
            color: #1d4ed8;
            font-weight: 600;
        }
        .radio-label input { 
            margin: 0 !important; 
            width: 16px !important; 
            height: 16px !important; 
            cursor: pointer;
            flex-shrink: 0 !important;
            accent-color: #2563eb;
        }
        
        /* 下拉选择框圆弧化 */
        select.input-field {
            appearance: none;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%2364748b' d='M2 4l4 4 4-4'/%3E%3C/svg%3E");
            background-repeat: no-repeat;
            background-position: right 12px center;
            padding-right: 36px;
        }
        
        .sub-fields-container { 
            margin-top: 14px; 
            padding: 16px; 
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); 
            border-left: 4px solid #2563eb; 
            border-radius: 12px; 
            display: flex; 
            flex-direction: column; 
            gap: 14px; 
        }
        .divider-line { 
            display: flex; 
            align-items: center; 
            margin: 20px 0 14px 0; 
            color: #2563eb; 
            font-size: 13px; 
            font-weight: 700; 
        }
        .divider-line::after { 
            content: ""; 
            flex: 1; 
            height: 2px; 
            background: linear-gradient(90deg, #bfdbfe 0%, transparent 100%); 
            margin-left: 12px; 
            border-radius: 2px;
        }
        
        /* 投资轮次等特殊样式 */
        .round-wrapper, .investor-wrapper, .shareholder-wrapper {
            background: #fff;
            border-radius: 14px;
            border: 1px solid #e2e8f0;
            padding: 16px;
            margin-bottom: 12px;
        }
        .round-header, .investor-header, .shareholder-header {
            display: flex;
            align-items: center;
            gap: 12px;
            padding-bottom: 12px;
            border-bottom: 1px solid #f1f5f9;
            margin-bottom: 12px;
        }
        .round-toggle, .investor-toggle, .shareholder-toggle {
            width: 18px;
            height: 18px;
            accent-color: #2563eb;
            cursor: pointer;
        }
        .round-label, .investor-label, .shareholder-label {
            font-size: 14px;
            font-weight: 600;
            color: #334155;
            cursor: pointer;
        }
    `;
    container.appendChild(style);

    if (!contractConfig || contractConfig.length === 0) {
        container.innerHTML = "<p>未找到配置项。</p>";
        return;
    }
    
    // =====================================================================
    // 【双轨同步 - 优先级加载逻辑】
    // 1. LocalStorage (最高优先级，同浏览器即时同步)
    // 2. lastLoadedFormData (来自 Custom XML / Document Settings)
    // 3. 默认值
    // =====================================================================
    const lsState = loadFormStateFromLocalStorage();
    if (lsState) {
        // LocalStorage 有数据，使用它作为"真理源"
        if (lsState.formData && Object.keys(lsState.formData).length > 0) {
            lastLoadedFormData = { ...lastLoadedFormData, ...lsState.formData };
            console.log("[BuildForm] Merged LocalStorage formData:", Object.keys(lsState.formData).length, "fields");
        }
        if (lsState.enabledRounds) {
            enabledRounds = { ...enabledRounds, ...lsState.enabledRounds };
            console.log("[BuildForm] Merged LocalStorage enabledRounds:", lsState.enabledRounds);
        }
        if (lsState.enabledCurrentInvestors) {
            enabledCurrentInvestors = { ...enabledCurrentInvestors, ...lsState.enabledCurrentInvestors };
            console.log("[BuildForm] Merged LocalStorage enabledCurrentInvestors:", lsState.enabledCurrentInvestors);
        }
        if (lsState.enabledExistingShareholders) {
            enabledExistingShareholders = { ...enabledExistingShareholders, ...lsState.enabledExistingShareholders };
            console.log("[BuildForm] Merged LocalStorage enabledExistingShareholders:", lsState.enabledExistingShareholders);
        }
    }

    // 递归创建字段
    function createFields(fields, parent, sectionId) {
        fields.forEach((field, fieldIndex) => {
            // 分割线特殊处理
            if (field.type === "divider") {
                const div = document.createElement("div");
                div.className = "divider-line";
                div.textContent = field.label;
                parent.appendChild(div);
                return;
            }
            
            // 【新增】HTML 占位符处理 (用于 Cloud Sync)
            if (field.type === "html_placeholder") {
                const targetId = field.targetId;
                const targetEl = document.getElementById(targetId);
                if (targetEl) {
                    parent.appendChild(targetEl);
                    // 确保显示
                    targetEl.style.display = "block";
                }
                return;
            }

            const wrapper = document.createElement("div");
            wrapper.className = "form-group";
            wrapper.id = `form-group-${sectionId}-${field.id || fieldIndex}`;
            
            // 【新增】添加拖拽支持
            wrapper.draggable = true;
            wrapper.dataset.sectionId = sectionId;
            wrapper.dataset.fieldId = field.id;
            wrapper.dataset.fieldIndex = fieldIndex;
            
            // 【新增】编辑按钮（hover 时显示）
            const editBtn = document.createElement("button");
            editBtn.className = "field-edit-btn";
            editBtn.innerHTML = "⚙️";
            editBtn.title = "编辑此字段";
            editBtn.onclick = (e) => {
                e.stopPropagation();
                editFieldInSection(sectionId, field.id || fieldIndex);
            };
            wrapper.appendChild(editBtn);
            
            // 【新增】拖拽事件
            wrapper.addEventListener("dragstart", handleFieldDragStart);
            wrapper.addEventListener("dragend", handleFieldDragEnd);
            
            // Label Row
            const labelRow = document.createElement("div");
            labelRow.className = "label-row";
            const label = document.createElement("label");
            label.textContent = field.label;
            labelRow.appendChild(label);
            
            // Insert Button
            if (field.tag) {
                const insertBtn = document.createElement("button");
                insertBtn.className = "insert-btn";
                // 如果有 hasParagraphToggle，使用"插入段落"模式
                const isParagraphMode = field.hasParagraphToggle === true;
                insertBtn.textContent = isParagraphMode ? "插入段落" : "插入";
                insertBtn.title = isParagraphMode 
                    ? `将选中的段落包裹为 [${field.label}] (保留原内容和格式)`
                    : `在光标处插入 [${field.label}]`;
                insertBtn.onclick = () => insertControl(field.tag, field.label, isParagraphMode);
                labelRow.appendChild(insertBtn);
            }
            wrapper.appendChild(labelRow);

            // Controls
            if (field.type === "radio") {
                const radioGroup = document.createElement("div");
                radioGroup.className = "radio-group";
                const groupName = field.id + "_" + Math.random().toString(36).substr(2, 5);

                (field.options || []).forEach(opt => {
                    const rLabel = document.createElement("label");
                    rLabel.className = "radio-label";
                    const radio = document.createElement("input");
                    radio.type = "radio";
                    radio.name = groupName; 
                    radio.value = opt;
                    radio.dataset.tag = field.tag;
                    
                    // 【还原状态】使用加载的快照数据
                    if (lastLoadedFormData[field.tag] !== undefined) {
                        radio.checked = (radio.value === lastLoadedFormData[field.tag]);
                    }

                    radio.addEventListener("change", () => {
                        // 判断是否为"显示"选项（注意：必须排除"不适用"等否定选项）
                        const shouldShow = (opt === "是" || opt === "有" || opt === "适用" || opt === "确认");
                        
                        // 检查是否为插入段落模式
                        if (field.hasParagraphToggle) {
                            // 插入段落模式：选择"适用"恢复段落，选择"不适用"隐藏段落
                            toggleRoundVisibility(field.tag, shouldShow);
                        } else {
                            // 普通模式：检查是否有值映射 (valueMap)
                            const mappedValue = field.valueMap ? field.valueMap[opt] : opt;
                            updateContent(field.tag, mappedValue, field.label); // 实时更新当前文档
                        }
                        scheduleAutoSync(); // 触发云端同步
                        updateSectionProgress(); // 【新增】更新进度
                        
                        // 联动 SubFields
                        if (field.subFields) {
                            const subContainer = wrapper.querySelector(".sub-fields-container");
                            if (subContainer) {
                                subContainer.style.display = shouldShow ? "block" : "none";
                            }
                        }
                    });

                    rLabel.appendChild(radio);
                    rLabel.appendChild(document.createTextNode(opt));
                    radioGroup.appendChild(rLabel);
                });
                wrapper.appendChild(radioGroup);

            } else if (field.type === "select") {
                const select = document.createElement("select");
                select.className = "input-field";
                select.dataset.tag = field.tag;
                
                const defOpt = document.createElement("option");
                defOpt.value = "";
                defOpt.textContent = "请选择...";
                select.appendChild(defOpt);

                (field.options || []).forEach(opt => {
                    const option = document.createElement("option");
                    option.value = opt;
                    option.textContent = opt;
                    select.appendChild(option);
                });
                
                // 【还原状态】
                if (lastLoadedFormData[field.tag] !== undefined) {
                    select.value = lastLoadedFormData[field.tag];
                }
                
                select.addEventListener("change", () => {
                    // 检查是否有值映射 (valueMap)
                    const mappedValue = field.valueMap ? field.valueMap[select.value] : select.value;
                    updateContent(field.tag, mappedValue, field.label);
                    scheduleAutoSync();
                    updateSectionProgress(); // 【新增】更新进度
                });
                wrapper.appendChild(select);

            } else {
                // Text / Number / Date
                const input = document.createElement("input");
                input.type = field.type || "text";
                input.className = "input-field";
                input.id = field.id; // 添加 ID 以便于查找
                input.dataset.tag = field.tag;
                if (field.placeholder) input.placeholder = field.placeholder;
                
                // 【还原状态】优先使用快照数据，否则使用默认值
                const val = lastLoadedFormData[field.tag] !== undefined ? lastLoadedFormData[field.tag] : (field.value || "");
                input.value = val;

                input.addEventListener("input", () => {
                    // 【特殊处理】如果是自动计算字段，用户手动编辑后停止自动更新
                    if (field.autoCount) {
                        shareholderCountUserModified = true;
                        console.log("[ShareholderCount] 用户手动编辑，停止自动更新");
                    }
                    
                    debounce(() => {
                        let formattedValue = input.value;
                        if (field.formatFn === "dateUnderline") {
                            formattedValue = formatDateUnderline(input.value);
                        } else if (field.type === "date") {
                            formattedValue = formatDate(input.value);
                        } else if (field.formatFn === "chineseNumber") {
                            formattedValue = formatChineseNumber(input.value);
                        }
                        updateContent(field.tag, formattedValue, field.label);
                    }, 600)();
                    scheduleAutoSync();
                    updateSectionProgress();
                });
                wrapper.appendChild(input);
            }

            // SubFields Container (递归)
            if (field.subFields) {
                const subContainer = document.createElement("div");
                subContainer.className = "sub-fields-container";
                subContainer.style.display = "none"; // 默认隐藏
                createFields(field.subFields, subContainer, sectionId);
                wrapper.appendChild(subContainer);
            }

            parent.appendChild(wrapper);
        });
    }

    contractConfig.forEach((section, sectionIndex) => {
        const headerDiv = document.createElement("div");
        headerDiv.className = "section-header-container";
        headerDiv.id = `section-nav-${section.id}`; // 添加 ID 用于目录跳转
        const h3 = document.createElement("h3");
        h3.className = "section-header-static";
        h3.textContent = section.header.label;
        headerDiv.appendChild(h3);
        
        // 【新增】如果 Section 有 hasSectionToggle，添加整体"插入段落"按钮
        if (section.hasSectionToggle && section.header.tag) {
            const sectionInsertBtn = document.createElement("button");
            sectionInsertBtn.className = "insert-btn insert-wrapper-btn";
            sectionInsertBtn.textContent = "插入段落";
            sectionInsertBtn.title = `将选中的整个"${section.header.label}"段落包裹为可显示/隐藏的区块`;
            sectionInsertBtn.style.marginLeft = "10px";
            sectionInsertBtn.onclick = (e) => {
                    e.preventDefault();
                insertControl(section.header.tag, section.header.label, true);
            };
            headerDiv.appendChild(sectionInsertBtn);
        }
        
        container.appendChild(headerDiv);

        // 特殊处理：本轮投资人 section
        if (section.type === "current_investors") {
            // ==================== 本轮投资人处理 ====================
            const investorsContainer = document.createElement("div");
            investorsContainer.className = "current-investors-container";
            
            section.investors.forEach(investor => {
                const investorWrapper = document.createElement("div");
                investorWrapper.className = "investor-wrapper";
                investorWrapper.dataset.investorId = investor.id;
                
                // 投资人标题行（含复选框）
                const investorHeader = document.createElement("div");
                investorHeader.className = "round-header"; // 复用样式
                
                const checkbox = document.createElement("input");
                checkbox.type = "checkbox";
                checkbox.className = "investor-toggle";
                checkbox.id = `toggle_inv_${investor.id}`;
                checkbox.checked = enabledCurrentInvestors[investor.id] || false;
                
                const investorLabel = document.createElement("label");
                investorLabel.className = "round-label";
                investorLabel.htmlFor = `toggle_inv_${investor.id}`;
                investorLabel.textContent = investor.label;
                
                // 插入包裹 Content Control 的按钮
                const insertWrapperBtn = document.createElement("button");
                insertWrapperBtn.className = "insert-btn insert-wrapper-btn";
                insertWrapperBtn.textContent = "插入段落";
                insertWrapperBtn.title = `请先选中整段文字，再点击此按钮包裹为 ${investor.label}`;
                insertWrapperBtn.onclick = (e) => {
                    e.preventDefault();
                    insertControl(investor.tag, investor.label, true, investor.id);
                };
                
                investorHeader.appendChild(checkbox);
                investorHeader.appendChild(investorLabel);
                investorHeader.appendChild(insertWrapperBtn);
                investorWrapper.appendChild(investorHeader);
                
                // 子表单容器
                const subFormContainer = document.createElement("div");
                subFormContainer.className = "round-subform";
                subFormContainer.style.display = enabledCurrentInvestors[investor.id] ? "block" : "none";
                
                // 生成投资人字段
                const tagPrefix = investor.tag.replace("Inv_", "");
                section.investorFields.forEach(fieldTemplate => {
                    const field = {
                        ...fieldTemplate,
                        id: `${investor.id}${fieldTemplate.id}`,
                        tag: `${tagPrefix}${fieldTemplate.tag}`
                    };
                    
                    const wrapper = document.createElement("div");
                    wrapper.className = "form-group";
                    
                    // 处理条件显示字段
                    if (field.showWhen) {
                        wrapper.classList.add("conditional-field");
                        wrapper.dataset.showWhen = JSON.stringify(field.showWhen);
                        wrapper.dataset.triggerFieldId = `${investor.id}_type`;
                        wrapper.style.display = "none"; // 默认隐藏
                        
                        // 【新增】如果有段落切换功能，记录段落 tag
                        if (field.hasParagraphToggle) {
                            wrapper.dataset.paraTag = field.tag;
                        }
                    }
                    
                    const labelRow = document.createElement("div");
                    labelRow.className = "label-row";
                    
                    const label = document.createElement("label");
                    label.textContent = field.label;
                    label.htmlFor = field.id;
                    labelRow.appendChild(label);
                    
                    // Buttons...
                    if (fieldTemplate.hasParagraphToggle) {
                        const paraTag = fieldTemplate.paraTag 
                            ? `${tagPrefix}${fieldTemplate.paraTag}` 
                            : field.tag;
                        const insertParaBtn = document.createElement("button");
                        insertParaBtn.className = "insert-btn";
                        insertParaBtn.textContent = "插入段落";
                        insertParaBtn.title = `将选中的段落包裹为 [${field.label}段落] (保留原内容和格式)`;
                        insertParaBtn.onclick = (e) => {
                            e.preventDefault();
                            insertControl(paraTag, `${field.label}段落`, true);
                        };
                        labelRow.appendChild(insertParaBtn);
                    
                    const insertBtn = document.createElement("button");
                    insertBtn.className = "insert-btn";
                    insertBtn.textContent = "插入";
                        insertBtn.title = `在光标处插入 [${field.label}]`;
                    insertBtn.onclick = (e) => {
                        e.preventDefault();
                            insertControl(field.tag, field.label, false);
                    };
                    labelRow.appendChild(insertBtn);
                    } else {
                        const insertBtn = document.createElement("button");
                        insertBtn.className = "insert-btn";
                        insertBtn.textContent = "插入";
                        insertBtn.title = `在光标处插入 [${field.label}]`;
                        insertBtn.onclick = (e) => {
                            e.preventDefault();
                            insertControl(field.tag, field.label, false);
                        };
                        labelRow.appendChild(insertBtn);
                    }
                    
                    wrapper.appendChild(labelRow);
                    
                    // 创建输入控件
                    let input;
                    if (field.type === "select") {
                        input = document.createElement("select");
                        input.className = "input-field";
                        input.id = field.id;
                        input.name = field.tag;
                        
                        const defaultOpt = document.createElement("option");
                        defaultOpt.value = "";
                        defaultOpt.textContent = "请选择...";
                        input.appendChild(defaultOpt);
                        
                        field.options.forEach(opt => {
                            const optEl = document.createElement("option");
                            optEl.value = opt;
                            optEl.textContent = opt;
                            if (lastLoadedFormData[field.tag] === opt) {
                                optEl.selected = true;
                            }
                            input.appendChild(optEl);
                        });
                        
                        if (field.triggerConditional) {
                            input.addEventListener("change", () => {
                                const selectedValue = input.value;
                                const conditionalFields = subFormContainer.querySelectorAll(".conditional-field");
                                conditionalFields.forEach(cf => {
                                    const showWhen = JSON.parse(cf.dataset.showWhen || "[]");
                                    const shouldShow = showWhen.includes(selectedValue);
                                    cf.style.display = shouldShow ? "block" : "none";
                                
                                    const paraTag = cf.dataset.paraTag;
                                    if (paraTag) {
                                        toggleRoundVisibility(paraTag, shouldShow);
                                    }
                                });
                                scheduleAutoSync();
                            });
                        }
                        
                        input.addEventListener("change", () => {
                            updateContent(field.tag, input.value, field.label);
                            scheduleAutoSync();
                            updateSectionProgress(); // 【新增】
                        });
                    } else {
                        input = document.createElement("input");
                        input.type = field.type === "number" ? "number" : "text";
                        input.className = "input-field";
                        input.id = field.id;
                        input.name = field.tag;
                        if (field.placeholder) input.placeholder = field.placeholder;
                        if (lastLoadedFormData[field.tag]) {
                            input.value = lastLoadedFormData[field.tag];
                        }
                        
                        input.addEventListener("input", () => {
                            debounce(() => {
                                updateContent(field.tag, input.value, field.label);
                            }, 600)();
                            scheduleAutoSync();
                            updateSectionProgress(); // 【新增】
                        });
                    }
                    
                    wrapper.appendChild(input);
                    subFormContainer.appendChild(wrapper);
                });
                
                investorWrapper.appendChild(subFormContainer);
                
                checkbox.addEventListener("change", () => {
                    enabledCurrentInvestors[investor.id] = checkbox.checked;
                    subFormContainer.style.display = checkbox.checked ? "block" : "none";
                    saveFormStateToLocalStorage(collectFormData(true), enabledRounds, enabledCurrentInvestors, enabledExistingShareholders);
                    toggleRoundVisibility(investor.tag, checkbox.checked);
                    scheduleAutoSync();
                    updateSectionProgress();
                    updateShareholderCount(); // 自动更新股东总数
                });
                
                // Init Conditional Logic
                setTimeout(() => {
                    const tagPrefix = investor.tag.replace("Inv_", "");
                    const typeTag = `${tagPrefix}_Type`;
                    const typeInput = subFormContainer.querySelector(`select[name="${typeTag}"]`);
                    const currentValue = typeInput ? typeInput.value : "";
                    
                    const conditionalFields = subFormContainer.querySelectorAll(".conditional-field");
                    conditionalFields.forEach(cf => {
                        const showWhen = JSON.parse(cf.dataset.showWhen || "[]");
                        const shouldShow = currentValue && showWhen.includes(currentValue);
                        cf.style.display = shouldShow ? "block" : "none";
                    });
                }, 200);
                
                investorsContainer.appendChild(investorWrapper);
            });
            
            container.appendChild(investorsContainer);
        } else if (section.type === "existing_shareholders") {
            // ==================== 其他现有股东处理 ====================
            const shareholdersContainer = document.createElement("div");
            shareholdersContainer.className = "current-investors-container"; // 复用样式
            
            section.shareholders.forEach(shareholder => {
                const shWrapper = document.createElement("div");
                shWrapper.className = "investor-wrapper";
                shWrapper.dataset.shareholderId = shareholder.id;
                
                // 股东标题行（含复选框）
                const shHeader = document.createElement("div");
                shHeader.className = "round-header";
                
                const checkbox = document.createElement("input");
                checkbox.type = "checkbox";
                checkbox.className = "shareholder-toggle";
                checkbox.id = `toggle_sh_${shareholder.id}`;
                checkbox.checked = enabledExistingShareholders[shareholder.id] || false;
                
                const shLabel = document.createElement("label");
                shLabel.className = "round-label";
                shLabel.htmlFor = `toggle_sh_${shareholder.id}`;
                shLabel.textContent = shareholder.label;
                
                const noteSpan = document.createElement("span");
                noteSpan.style.fontSize = "10px";
                noteSpan.style.color = "#888";
                noteSpan.style.marginLeft = "auto";
                noteSpan.textContent = "(表格行)";
                
                shHeader.appendChild(checkbox);
                shHeader.appendChild(shLabel);
                shHeader.appendChild(noteSpan);
                shWrapper.appendChild(shHeader);
                
                // 子表单容器
                const subFormContainer = document.createElement("div");
                subFormContainer.className = "round-subform";
                subFormContainer.style.display = enabledExistingShareholders[shareholder.id] ? "block" : "none";
                
                // 生成股东字段
                const tagPrefix = shareholder.tag;
                section.shareholderFields.forEach(fieldTemplate => {
                    const field = {
                        ...fieldTemplate,
                        id: `${shareholder.id}${fieldTemplate.id}`,
                        tag: `${tagPrefix}${fieldTemplate.tag}`
                    };
                    
                    const wrapper = document.createElement("div");
                    wrapper.className = "form-group";
                    
                    // 【新增】处理条件显示字段 (showWhen - 根据类型)
                    if (fieldTemplate.showWhen) {
                        wrapper.classList.add("conditional-field");
                        wrapper.dataset.showWhen = JSON.stringify(fieldTemplate.showWhen);
                        wrapper.style.display = "none"; // 默认隐藏
                        if (fieldTemplate.hasParagraphToggle) {
                            const paraTag = fieldTemplate.paraTag 
                                ? `${tagPrefix}${fieldTemplate.paraTag}` 
                                : field.tag;
                            wrapper.dataset.paraTag = paraTag;
                        }
                    }
                    
                    // 【新增】处理条件显示字段 (showWhenRound - 根据融资轮次)
                    if (fieldTemplate.showWhenRound) {
                        wrapper.classList.add("conditional-round-field");
                        wrapper.dataset.showWhenRound = JSON.stringify(fieldTemplate.showWhenRound);
                        wrapper.style.display = "none"; // 默认隐藏
                    }
                    
                    const labelRow = document.createElement("div");
                    labelRow.className = "label-row";
                    const label = document.createElement("label");
                    label.textContent = field.label;
                    label.htmlFor = field.id;
                    labelRow.appendChild(label);
                    
                    // 【新增】根据 hasParagraphToggle 决定插入按钮类型
                    if (fieldTemplate.hasParagraphToggle) {
                        // 插入段落按钮
                        const paraTag = fieldTemplate.paraTag 
                            ? `${tagPrefix}${fieldTemplate.paraTag}` 
                            : field.tag;
                        const insertParaBtn = document.createElement("button");
                        insertParaBtn.className = "insert-btn";
                        insertParaBtn.textContent = "插入段落";
                        insertParaBtn.title = `将选中的段落包裹为 [${field.label}段落]`;
                        insertParaBtn.onclick = (e) => {
                            e.preventDefault();
                            insertControl(paraTag, `${field.label}段落`, true);
                        };
                        labelRow.appendChild(insertParaBtn);
                        
                        // 普通插入按钮
                        const insertBtn = document.createElement("button");
                        insertBtn.className = "insert-btn";
                        insertBtn.textContent = "插入";
                        insertBtn.onclick = (e) => {
                            e.preventDefault();
                            insertControl(field.tag, field.label, false);
                        };
                        labelRow.appendChild(insertBtn);
                    } else {
                    const insertBtn = document.createElement("button");
                    insertBtn.className = "insert-btn";
                    insertBtn.textContent = "插入";
                    insertBtn.onclick = (e) => {
                        e.preventDefault();
                        insertControl(field.tag, field.label);
                    };
                    labelRow.appendChild(insertBtn);
                    }
                    
                    wrapper.appendChild(labelRow);
                    
                    // Inputs...
                    let input;
                    if (field.type === "select") {
                        input = document.createElement("select");
                        input.className = "input-field";
                        input.id = field.id;
                        input.name = field.tag;
                        input.dataset.tag = field.tag;
                        
                        const defaultOpt = document.createElement("option");
                        defaultOpt.value = "";
                        defaultOpt.textContent = "请选择...";
                        input.appendChild(defaultOpt);
                        
                        field.options.forEach(opt => {
                            const optEl = document.createElement("option");
                            optEl.value = opt;
                            optEl.textContent = opt;
                            if (lastLoadedFormData[field.tag] === opt) {
                                optEl.selected = true;
                            }
                            input.appendChild(optEl);
                        });
                        
                        input.addEventListener("change", () => {
                            updateContent(field.tag, input.value, field.label);
                            
                            // 【新增】如果是类型字段 (triggerConditional)，联动条件字段
                            if (fieldTemplate.triggerConditional) {
                                const selectedValue = input.value;
                                const conditionalFields = subFormContainer.querySelectorAll(".conditional-field");
                                conditionalFields.forEach(cf => {
                                    const showWhen = JSON.parse(cf.dataset.showWhen || "[]");
                                    const shouldShow = showWhen.includes(selectedValue);
                                    cf.style.display = shouldShow ? "block" : "none";
                                    
                                    // 如果条件字段有段落切换功能，同时控制文档中的段落
                                    const paraTag = cf.dataset.paraTag;
                                    if (paraTag) {
                                        console.log(`[ConditionalField] ${paraTag} ${shouldShow ? '显示' : '隐藏'}`);
                                        toggleRoundVisibility(paraTag, shouldShow);
                                    }
                                });
                            }
                            
                            // 【新增】如果是融资轮次字段，联动轮次相关字段
                            if (field.id.endsWith("_round")) {
                                const selectedRound = input.value;
                                const roundConditionalFields = subFormContainer.querySelectorAll(".conditional-round-field");
                                roundConditionalFields.forEach(cf => {
                                    const showWhenRound = JSON.parse(cf.dataset.showWhenRound || "[]");
                                    const shouldShow = showWhenRound.includes(selectedRound);
                                    cf.style.display = shouldShow ? "block" : "none";
                                });
                            }
                            
                            scheduleAutoSync();
                            updateSectionProgress();
                        });
                    } else {
                        input = document.createElement("input");
                        input.type = field.type === "number" ? "number" : "text";
                        input.className = "input-field";
                        input.id = field.id;
                        input.name = field.tag;
                        input.dataset.tag = field.tag;
                        if (field.placeholder) input.placeholder = field.placeholder;
                        if (lastLoadedFormData[field.tag]) {
                            input.value = lastLoadedFormData[field.tag];
                        }
                        
                        input.addEventListener("input", () => {
                            debounce(() => {
                                updateContent(field.tag, input.value, field.label);
                            }, 600)();
                            scheduleAutoSync();
                            updateSectionProgress();
                        });
                    }
                    wrapper.appendChild(input);
                    subFormContainer.appendChild(wrapper);
                });
                
                // 【新增】初始化条件字段显示状态
                setTimeout(() => {
                    // 类型字段联动
                    const typeSelect = subFormContainer.querySelector(`select[data-tag="${tagPrefix}_Type"]`);
                    const currentTypeValue = typeSelect ? typeSelect.value : "";
                    const conditionalFields = subFormContainer.querySelectorAll(".conditional-field");
                    conditionalFields.forEach(cf => {
                        const showWhen = JSON.parse(cf.dataset.showWhen || "[]");
                        const shouldShow = currentTypeValue && showWhen.includes(currentTypeValue);
                        cf.style.display = shouldShow ? "block" : "none";
                    });
                    
                    // 融资轮次字段联动
                    const roundSelect = subFormContainer.querySelector(`select[data-tag="${tagPrefix}_Round"]`);
                    const currentRoundValue = roundSelect ? roundSelect.value : "";
                    const roundConditionalFields = subFormContainer.querySelectorAll(".conditional-round-field");
                    roundConditionalFields.forEach(cf => {
                        const showWhenRound = JSON.parse(cf.dataset.showWhenRound || "[]");
                        const shouldShow = currentRoundValue && showWhenRound.includes(currentRoundValue);
                        cf.style.display = shouldShow ? "block" : "none";
                    });
                }, 200);
                
                shWrapper.appendChild(subFormContainer);
                
                checkbox.addEventListener("change", async () => {
                    enabledExistingShareholders[shareholder.id] = checkbox.checked;
                    subFormContainer.style.display = checkbox.checked ? "block" : "none";
                    saveFormStateToLocalStorage(collectFormData(true), enabledRounds, enabledCurrentInvestors, enabledExistingShareholders);
                    await toggleShareholderFieldsVisibility(shareholder.tag, checkbox.checked);
                    scheduleAutoSync();
                    updateSectionProgress();
                    updateShareholderCount(); // 自动更新股东总数
                });
                
                shareholdersContainer.appendChild(shWrapper);
            });
            
            container.appendChild(shareholdersContainer);
        } else {
            // 普通 section
            const fieldsDiv = document.createElement("div");
            fieldsDiv.className = "section-fields";
            fieldsDiv.dataset.sectionId = section.id;
            createFields(section.fields, fieldsDiv, section.id);
            container.appendChild(fieldsDiv);
        }
    });
    
    // ========== 生成侧边进度条 ==========
    renderProgressBar();
    
    // ========== 初始化进度状态 ==========
    updateSectionProgress();
    
    // ========== 初始化股东总数 ==========
    updateShareholderCount();
}

// ---------------- 渲染进度侧边栏 (Step Timeline) ----------------
function renderProgressBar() {
    const stepList = document.getElementById("step-list");
    if (!stepList) return;
    
    stepList.innerHTML = "";
    
    contractConfig.forEach((section, index) => {
        // 主标题项
        const li = document.createElement("li");
        li.className = "step-item";
        li.dataset.sectionId = section.id;
        li.id = `step-nav-${section.id}`;
        
        const marker = document.createElement("div");
        marker.className = "step-marker";
        marker.innerHTML = index + 1;
        
        const content = document.createElement("div");
        content.className = "step-content";
        
        const title = document.createElement("div");
        title.className = "step-title";
        title.textContent = section.header.label;
        
        content.appendChild(title);
        li.appendChild(marker);
        li.appendChild(content);
        
        li.addEventListener("click", () => {
            const targetEl = document.getElementById(`section-nav-${section.id}`);
            if (targetEl) {
                targetEl.scrollIntoView({ behavior: "smooth", block: "start" });
                document.querySelectorAll(".step-item").forEach(item => item.classList.remove("active"));
                li.classList.add("active");
            }
        });
        
        stepList.appendChild(li);

        // 重要小标题层级展示（本轮投资人、现有股东）
        if (section.type === "current_investors" || section.type === "existing_shareholders") {
            const subItems = [];
            if (section.investors) section.investors.forEach(i => subItems.push({ label: i.label, tag: i.tag, id: i.id }));
            if (section.shareholders) section.shareholders.forEach(s => subItems.push({ label: s.label, tag: s.tag, id: s.id }));

            subItems.forEach(sub => {
                const subLi = document.createElement("li");
                subLi.className = "step-item sub-step";
                subLi.innerHTML = `<div class="step-marker"></div><div class="step-content"><div class="step-title">${sub.label}</div></div>`;
                subLi.addEventListener("click", (e) => {
                    e.stopPropagation();
                    const target = document.querySelector(`[data-round-id="${sub.id}"]`) || 
                                   document.querySelector(`[data-investor-id="${sub.id}"]`) ||
                                   document.querySelector(`[data-shareholder-id="${sub.id}"]`);
                    if (target) {
                        target.scrollIntoView({ behavior: "smooth", block: "center" });
                    }
                });
                stepList.appendChild(subLi);
            });
        }
    });
    
    // 添加"合同交付"项
    const finalizeLi = document.createElement("li");
    finalizeLi.className = "step-item";
    finalizeLi.dataset.sectionId = "finalize";
    finalizeLi.id = "step-nav-finalize";
    
    finalizeLi.innerHTML = `<div class="step-marker"><i class="ms-Icon ms-Icon--Package" aria-hidden="true"></i></div>
                            <div class="step-content"><div class="step-title">合同交付</div></div>`;
    
    finalizeLi.addEventListener("click", () => {
        const targetEl = document.getElementById("section-finalize");
        if (targetEl) {
            targetEl.scrollIntoView({ behavior: "smooth", block: "start" });
            document.querySelectorAll(".step-item").forEach(item => item.classList.remove("active"));
            finalizeLi.classList.add("active");
        }
    });
    
    stepList.appendChild(finalizeLi);
    
    // 滚动监听 - 自动高亮当前可见的 section
    let scrollTimeout;
    window.addEventListener("scroll", () => {
        clearTimeout(scrollTimeout);
        scrollTimeout = setTimeout(() => {
            let currentSection = null;
            contractConfig.forEach(section => {
                const el = document.getElementById(`section-nav-${section.id}`);
                if (el) {
                    const rect = el.getBoundingClientRect();
                    if (rect.top <= 150 && rect.bottom >= 150) {
                        currentSection = section.id;
                    }
                }
            });
            const finEl = document.getElementById("section-finalize");
            if (finEl && finEl.getBoundingClientRect().top <= 300) {
                currentSection = "finalize";
            }
            
            if (currentSection) {
                document.querySelectorAll(".step-item").forEach(item => {
                    if (item.dataset.sectionId === currentSection) {
                        item.classList.add("active");
                    } else {
                        item.classList.remove("active");
                    }
                });
            }
        }, 100);
    });
}

/**
 * 计算并更新股东总数
 * 统计：股东1 (固定) + 已勾选的其他股东 + 已勾选的本轮投资人
 */
function updateShareholderCount() {
    if (shareholderCountUserModified) {
        console.log("[ShareholderCount] 用户已手动修改，跳过自动更新");
        return;
    }
    
    // 股东1 始终算1个
    let count = 1;
    
    // 统计已启用的其他股东
    Object.keys(enabledExistingShareholders).forEach(key => {
        if (enabledExistingShareholders[key]) {
            count++;
        }
    });
    
    // 统计已启用的本轮投资人（他们也是增资后的股东）
    Object.keys(enabledCurrentInvestors).forEach(key => {
        if (enabledCurrentInvestors[key]) {
            count++;
        }
    });
    
    console.log(`[ShareholderCount] 自动计算: ${count} 个股东`);
    
    // 更新表单中的股东总数字段
    const countInput = document.getElementById("shareholderCount");
    if (countInput) {
        countInput.value = count;
        // 同时更新 Word 文档中的内容
        updateContent("ShareholderCount", String(count), "股东总数");
    }
}

function updateSectionProgress() {
    contractConfig.forEach((section, index) => {
        const stepItem = document.getElementById(`step-nav-${section.id}`);
        if (!stepItem) return;
        
        // 特殊处理 Section 1: 所需文件 (只要 #auth-connected-container 显示，就算完成)
        if (section.id === "section_files") {
            const connectedContainer = document.getElementById("auth-connected-container");
            const isConnected = connectedContainer && connectedContainer.style.display !== "none";
            if (isConnected) {
                stepItem.classList.add("completed");
                stepItem.querySelector(".step-marker").innerHTML = "✓";
            } else {
                stepItem.classList.remove("completed");
                stepItem.querySelector(".step-marker").innerHTML = index + 1;
            }
            return;
        }
        
        // 统计字段填写情况
        const sectionEl = document.getElementById(`section-nav-${section.id}`);
        if (!sectionEl) return;
        
        let contentEl = sectionEl.nextElementSibling;
        if (!contentEl) return;
        
        const allInputs = contentEl.querySelectorAll("input, select, textarea");
        let total = 0;
        let filled = 0;
        
        allInputs.forEach(input => {
            // 跳过隐藏的元素
            if (input.offsetParent === null) return;
            
            if (input.type === "checkbox" || input.type === "radio") {
                if (input.type === "radio") {
                    const name = input.name;
                    const groupRadios = contentEl.querySelectorAll(`input[name="${name}"]`);
                    if (input === groupRadios[0]) {
                        total++;
                        const isChecked = Array.from(groupRadios).some(r => r.checked);
                        if (isChecked) filled++;
                    }
                }
            } else {
                // Text / Select / Date
                total++;
                if (input.value && input.value.trim() !== "") {
                    filled++;
                }
            }
        });
        
        // 判定完成：填写率 > 80% 或者 total=0
        if (total > 0 && (filled / total) >= 0.8) {
            stepItem.classList.add("completed");
            stepItem.querySelector(".step-marker").innerHTML = "✓";
        } else {
            stepItem.classList.remove("completed");
            stepItem.querySelector(".step-marker").innerHTML = index + 1;
        }
    });
}

async function insertTextPreserveFormat(ctrl, text, context) {
    const isMultiLine = text && text.includes('\n');
    let savedFont = null;
    
    try {
        // 获取当前字体属性
        const range = ctrl.getRange();
        range.load("font/name, font/size, font/color, font/bold, font/italic, font/underline");
        await context.sync();

        // 保存字体属性
        savedFont = {
            name: range.font.name,
            size: range.font.size,
            color: range.font.color,
            bold: range.font.bold,
            italic: range.font.italic,
            underline: range.font.underline
        };
        
        console.log(`[InsertTextPreserveFormat] 保存格式:`, savedFont.name, savedFont.size, isMultiLine ? "(多行)" : "");
    } catch (err) {
        console.warn(`[InsertTextPreserveFormat] 获取格式失败:`, err.message);
    }

        // 插入新文本
    try {
        ctrl.insertText(text, "Replace");
        await context.sync();
        console.log(`[InsertTextPreserveFormat] 文本插入成功`);
    } catch (insertErr) {
        console.error(`[InsertTextPreserveFormat] 插入文本失败:`, insertErr.message);
        return;
    }

    // 尝试恢复格式（无论单行还是多行都尝试）
    if (savedFont && (savedFont.name || savedFont.size)) {
        try {
            // 对于多行文本，获取所有段落并设置格式
            if (isMultiLine) {
                const paragraphs = ctrl.paragraphs;
                paragraphs.load("items");
        await context.sync();
                
                for (const para of paragraphs.items) {
                    const paraRange = para.getRange();
                    if (savedFont.name && typeof savedFont.name === 'string') {
                        paraRange.font.name = savedFont.name;
                    }
                    if (savedFont.size && typeof savedFont.size === 'number' && savedFont.size > 0) {
                        paraRange.font.size = savedFont.size;
                    }
                    if (savedFont.color && typeof savedFont.color === 'string' && 
                        savedFont.color !== "null" && !savedFont.color.includes("auto")) {
                        paraRange.font.color = savedFont.color;
                    }
                    if (typeof savedFont.bold === 'boolean') {
                        paraRange.font.bold = savedFont.bold;
                    }
                    if (typeof savedFont.italic === 'boolean') {
                        paraRange.font.italic = savedFont.italic;
                    }
                }
        await context.sync();
                console.log(`[InsertTextPreserveFormat] 多行格式已恢复 (${paragraphs.items.length} 段)`);
            } else {
                // 单行文本，直接设置整个范围
                const newRange = ctrl.getRange();
                if (savedFont.name && typeof savedFont.name === 'string') {
                    newRange.font.name = savedFont.name;
                }
                if (savedFont.size && typeof savedFont.size === 'number' && savedFont.size > 0) {
                    newRange.font.size = savedFont.size;
                }
                if (savedFont.color && typeof savedFont.color === 'string' && 
                    savedFont.color !== "null" && !savedFont.color.includes("auto")) {
                    newRange.font.color = savedFont.color;
                }
                if (typeof savedFont.bold === 'boolean') {
                    newRange.font.bold = savedFont.bold;
                }
                if (typeof savedFont.italic === 'boolean') {
                    newRange.font.italic = savedFont.italic;
                }
                if (savedFont.underline && typeof savedFont.underline === 'string' && savedFont.underline !== "None") {
                    newRange.font.underline = savedFont.underline;
                }
                await context.sync();
                console.log(`[InsertTextPreserveFormat] 单行格式已恢复`);
            }
        } catch (formatErr) {
            console.warn(`[InsertTextPreserveFormat] 恢复格式失败，但文本已插入:`, formatErr.message);
        }
    }
}

// ---------------- 日期格式化 ----------------
function formatDate(dateStr) {
    if (!dateStr) return "";
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${year}年${month}月${day}日`;
}

// ---------------- 日期格式化（下划线格式）----------------
// 输出格式: _____年_____月_____日 (填入数字后为: 2026年01月04日)
function formatDateUnderline(dateStr) {
    if (!dateStr) return "_____年_____月_____日";
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return "_____年_____月_____日";
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}年${month}月${day}日`;
}

// ---------------- 数字转中文格式化 ----------------
// 将数字格式化为 "四（4）" 这样的格式
function formatChineseNumber(numStr) {
    if (!numStr || numStr === "") return "";
    const num = parseInt(numStr, 10);
    if (isNaN(num)) return numStr;
    
    const chineseDigits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"];
    let chinese = "";
    
    if (num >= 0 && num <= 10) {
        chinese = chineseDigits[num];
    } else if (num > 10 && num < 20) {
        chinese = "十" + (num % 10 === 0 ? "" : chineseDigits[num % 10]);
    } else if (num >= 20 && num < 100) {
        chinese = chineseDigits[Math.floor(num / 10)] + "十" + (num % 10 === 0 ? "" : chineseDigits[num % 10]);
    } else {
        chinese = String(num); // 超过99直接用数字
    }
    
    return `${chinese}（${num}）`;
}

// ---------------- 自动埋点：扫描【占位符】并创建 Content Control ----------------

/**
 * 确认并执行自动埋点
 */
async function confirmAutoEmbed() {
    const confirmed = await showConfirmDialog(
        "此操作将扫描文档中所有【占位符】格式的文本，并将其转换为可填充的埋点。\n\n已经是埋点的内容会自动跳过。",
        {
            title: "自动埋点确认",
            confirmText: "开始埋点",
            cancelText: "取消"
        }
    );
    
    if (confirmed) {
        showNotification("正在扫描文档...", "info");
        await autoCreateContentControls();
    }
}

/**
 * 自动扫描文档中的【xxx】格式占位符，将其转换为 Content Control
 * @returns {Promise<{success: number, skipped: number}>} 成功和跳过的数量
 */
async function autoCreateContentControls() {
    console.log("[AutoEmbed] 开始自动埋点...");
    
    if (typeof Word === 'undefined') {
        showNotification("Word API 不可用", "error");
        return { success: 0, skipped: 0 };
    }
    
    let successCount = 0;
    let skippedCount = 0;
    
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            
            // 搜索所有【xxx】格式的文本（使用通配符）
            // Word 的通配符语法：【*】 匹配【开头，任意字符，】结尾
            const searchResults = body.search("【*】", { matchWildcards: true });
            searchResults.load("items");
            await context.sync();
            
            console.log(`[AutoEmbed] 找到 ${searchResults.items.length} 个匹配项`);
            
            if (searchResults.items.length === 0) {
                showNotification("未找到任何【占位符】格式的文本", "info");
                return;
            }
            
            // 遍历每个匹配项
            for (let i = 0; i < searchResults.items.length; i++) {
                const range = searchResults.items[i];
                range.load("text, parentContentControlOrNullObject");
                await context.sync();
                
                // 检查是否已经在 Content Control 内（避免重复埋点）
                const parentCC = range.parentContentControlOrNullObject;
                parentCC.load("isNullObject, tag");
                await context.sync();
                
                if (!parentCC.isNullObject) {
                    // 已经是 Content Control，跳过
                    console.log(`[AutoEmbed] 跳过已埋点的: ${range.text}`);
                    skippedCount++;
                    continue;
                }
                
                // 提取占位符名称（去掉【】）
                const fullText = range.text;
                const chineseName = fullText.replace(/^【/, "").replace(/】$/, "").trim();
                
                if (!chineseName) {
                    console.log(`[AutoEmbed] 跳过空占位符: ${fullText}`);
                    skippedCount++;
                    continue;
                }
                
                // 将中文转换为驼峰拼音作为 tag
                // 使用 pinyin-pro 库
                let pinyinTag = chineseName;
                try {
                    if (typeof pinyinPro !== 'undefined' && pinyinPro.pinyin) {
                        // 获取拼音数组，每个字一个拼音
                        const pinyinArr = pinyinPro.pinyin(chineseName, { toneType: 'none', type: 'array' });
                        // 转换为驼峰格式：每个拼音首字母大写
                        pinyinTag = pinyinArr.map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join('');
                    }
                } catch (e) {
                    console.warn(`[AutoEmbed] 拼音转换失败，使用原文: ${chineseName}`, e);
                }
                
                // 创建 Content Control
                try {
                    const cc = range.insertContentControl("RichText");
                    cc.tag = pinyinTag;  // tag 使用驼峰拼音，如 "MuBiaoGongSiMingCheng"
                    cc.title = chineseName; // title 保持中文，如 "目标公司名称"
                    cc.appearance = "Tags"; // Tags 模式显示 title 标签
                    cc.color = "blue";
                    cc.cannotEdit = false;
                    cc.cannotDelete = false;
                    
                    await context.sync();
                    successCount++;
                    console.log(`[AutoEmbed] ✓ 成功埋点: ${chineseName} → tag="${pinyinTag}"`);
                } catch (err) {
                    console.warn(`[AutoEmbed] 埋点失败 (${chineseName}):`, err.message);
                    skippedCount++;
                }
            }
            
            await context.sync();
        });
        
        const message = `自动埋点完成！\n成功: ${successCount} 个\n跳过: ${skippedCount} 个`;
        showNotification(message, successCount > 0 ? "success" : "info", 5000);
        console.log(`[AutoEmbed] 完成: 成功 ${successCount}, 跳过 ${skippedCount}`);
        
    } catch (error) {
        console.error("[AutoEmbed] 自动埋点失败:", error);
        showNotification(`自动埋点失败: ${error.message}`, "error");
    }
    
    return { success: successCount, skipped: skippedCount };
}

// ================================================================================
// AI 智能识别变量模块
// ================================================================================

// 豆包 AI API 配置
const DOUBAO_API = {
    url: "https://ark.cn-beijing.volces.com/api/v3/responses",
    token: "669c58ca-3cfd-4912-a30d-f9727cbc6485",
    model: "doubao-seed-1-8-251228"
};

// AI 识别字段存储 Key
const AI_FIELDS_KEY = "ai_recognized_fields";

// 可用的分类 Section 列表
const AI_SECTION_CATEGORIES = [
    { id: "section_company_info", name: "公司基本信息", examples: "公司名称、注册资本、法定代表人、注册地址" },
    { id: "section_financing", name: "本轮融资信息", examples: "投资金额、估值、股比、融资轮次" },
    { id: "section_current_investors", name: "本轮投资人", examples: "投资人名称、投资额" },
    { id: "section_existing_shareholders", name: "现有股东", examples: "股东名称、持股比例" },
    { id: "section_board", name: "董事会/核心员工", examples: "董事姓名、创始人、高管" },
    { id: "section_economics", name: "核心经济条款", examples: "分红比例、清算优先权" },
    { id: "section_redemption", name: "回购权", examples: "回购价格、回购期限" },
    { id: "section_definitions", name: "定义及签约方", examples: "控股股东、关联方定义" },
    { id: "section_other", name: "其他", examples: "无法归类的变量" }
];

/**
 * 调用豆包 AI 分析文档
 * @param {string} documentText 文档文本内容
 * @param {number} chunkIndex 当前分块索引（可选）
 * @param {number} totalChunks 总分块数（可选）
 * @returns {Promise<Array>} 识别出的变量列表
 */
async function callDoubaoAI(documentText, chunkIndex = 0, totalChunks = 1) {
    console.log(`[AI] 开始调用豆包 AI... (分块 ${chunkIndex + 1}/${totalChunks})`);
    console.log(`[AI] 文档长度: ${documentText.length} 字符`);
    
    const chunkInfo = totalChunks > 1 ? `\n\n注意：这是文档的第 ${chunkIndex + 1}/${totalChunks} 部分。` : '';
    
    const prompt = `你是一位资深合同起草专家，正在制作合同模版供实习生使用。

【任务】
识别合同中的变量，挖空它们，保留固定条款。${chunkInfo}

【核心规则：text 必须能被精确搜索到】
⚠️ **关键**：text 字段必须是文档中**实际存在且可被精确搜索**的文本！

✅ 正确示例：
- 【xxx】格式 → text = "【xxx】" 或 "【】" (整个方括号)
- "逾期    个工作日" → text = "逾期    个工作日" (带上下文的完整短语)
- "有效期至    年    月    日" → text = "有效期至    年    月    日"
- "产品剩余有效期不低于    个月" → text = "不低于    个月"
- "收款人：" 后面有下划线 → text = "收款人：_________" (包含标签)

❌ 错误示例（会导致搜索失败）：
- ❌ text: "    " (纯空格无法搜索)
- ❌ text: "                     " (纯空格)
- ❌ text: "____" (单独的下划线可能有多处)

【占位符格式识别】
1. 【xxx】格式 → text 必须包含【和】
2. 如果是空白/下划线，text 必须包含前面的标签词（如"收款人：____"）
3. 日期格式如"    年    月    日" → 保留完整格式

【变量分类】

1. field（必填变量）
   - 值会变化，但**必须存在**于每份合同
   - 如：公司名称、签约日期、投资金额

2. paragraph（可选段落）
   - 整个部分**可能不出现**在某些合同中
   - 如：回购权条款、对赌条款
   - 通常包含"如有"、"若"、"可选"等词

【返回格式】
JSON 数组，每项包含：
- text: **只包含需要填写的部分**（空白/下划线/【占位符】，不要包含周围固定文字）
- label: 中文名称（用于表单显示，应包含上下文，如"产品剩余有效期月数"）
- sectionId: 分类ID
- type: text/date/number
- mode: "field" 或 "paragraph"
- alternativeTexts: 其他同义表述（可选）
- parentTag: 父级 label 拼音（可选）

【分类ID】
${AI_SECTION_CATEGORIES.map(s => `- ${s.id}: ${s.name}`).join('\n')}

【识别规则】
1. text 只包含空白/占位符本身，绝不包含周围固定文字
2. label 要有完整含义（因为 text 可能只是空白）
3. 同义词合并（如"甲方"/"投资方"）
4. paragraph 适用于可整段删除的内容

【示例返回】
[
  {
    "text": "甲方公司",
    "alternativeTexts": ["甲方", "投资方"],
    "label": "投资方名称",
    "sectionId": "section_current_investors",
    "type": "text",
    "mode": "field"
  },
  {
    "text": "如甲方要求回购",
    "label": "回购权条款",
    "sectionId": "section_redemption",
    "type": "text",
    "mode": "paragraph"
  },
  {
    "text": "回购价格",
    "label": "回购价格",
    "sectionId": "section_redemption",
    "type": "number",
    "mode": "field",
    "parentTag": "HuiGouQuanTiaoKuan"
  }
]

合同文本：
---
${documentText}
---

只返回 JSON 数组，不要其他任何内容。`;

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
                    content: [{ type: "input_text", text: prompt }] 
                }]
            })
        });
        
        if (!response.ok) {
            const errText = await response.text();
            throw new Error(`API 请求失败: ${response.status} - ${errText}`);
        }
        
        const result = await response.json();
        console.log("[AI] API 响应:", result);
        
        // 解析响应内容
        let outputText = "";
        if (result.output && Array.isArray(result.output)) {
            // 豆包 API 返回 output 是数组，找 type: "message"
            const messageObj = result.output.find(o => o.type === 'message');
            
            if (messageObj && Array.isArray(messageObj.content)) {
                // content 是数组，提取文本 - 支持多种类型格式
                let textContent = messageObj.content.find(c => c.type === 'text' || c.type === 'output_text');
                
                if (textContent && textContent.text) {
                    outputText = textContent.text;
                }
            }
        } else if (result.output && result.output.content) {
            outputText = result.output.content;
        } else if (result.choices && result.choices[0] && result.choices[0].message) {
            outputText = result.choices[0].message.content;
        } else if (typeof result === 'string') {
            outputText = result;
        } else {
            // 尝试查找 output 字段
            outputText = JSON.stringify(result);
        }
        
        console.log("[AI] 原始输出:", outputText);
        
        // 提取 JSON 数组
        const jsonMatch = outputText.match(/\[[\s\S]*\]/);
        if (!jsonMatch) {
            console.warn("[AI] 无法从响应中提取 JSON 数组");
            return [];
        }
        
        const variables = JSON.parse(jsonMatch[0]);
        console.log(`[AI] 解析出 ${variables.length} 个变量:`, variables);
        
        // 验证和过滤
        const validVariables = variables.filter(v => {
            if (!v.text || !v.label || !v.sectionId) {
                console.warn("[AI] 跳过无效变量:", v);
                return false;
            }
            // 确保 sectionId 有效
            const validSectionIds = AI_SECTION_CATEGORIES.map(s => s.id);
            if (!validSectionIds.includes(v.sectionId)) {
                v.sectionId = "section_other";
            }
            // 确保 mode 有效
            if (!v.mode || !['field', 'paragraph'].includes(v.mode)) {
                v.mode = 'field'; // 默认必填
            }
            // 确保 alternativeTexts 是数组
            if (v.alternativeTexts && !Array.isArray(v.alternativeTexts)) {
                v.alternativeTexts = [];
            }
            return true;
        });
        
        return validVariables;
        
    } catch (error) {
        console.error("[AI] 调用失败:", error);
        throw error;
    }
}

/**
 * 分段上传全文进行 AI 分析
 * @param {string} fullText 完整文档文本
 * @returns {Promise<Array>} 识别出的变量列表（已去重）
 */
async function analyzeFullDocument(fullText) {
    const CHUNK_SIZE = 10000; // 每段约 10000 字符
    console.log(`[AI] 文档总长度: ${fullText.length} 字符`);
    
    if (fullText.length <= CHUNK_SIZE) {
        // 文档够短，直接分析
        return await callDoubaoAI(fullText, 0, 1);
    }
    
    // 分段处理
    const chunks = [];
    for (let i = 0; i < fullText.length; i += CHUNK_SIZE) {
        chunks.push(fullText.substring(i, i + CHUNK_SIZE));
    }
    
    console.log(`[AI] 分为 ${chunks.length} 段进行分析`);
    let allVariables = [];
    
    for (let i = 0; i < chunks.length; i++) {
        showNotification(`AI 分析中... (${i + 1}/${chunks.length})`, "info", 3000);
        try {
            const result = await callDoubaoAI(chunks[i], i, chunks.length);
            allVariables = allVariables.concat(result);
            console.log(`[AI] 第 ${i + 1} 段识别出 ${result.length} 个变量`);
        } catch (error) {
            console.warn(`[AI] 第 ${i + 1} 段分析失败:`, error);
            // 继续处理下一段
        }
    }
    
    // 去重：合并同一变量
    const deduplicated = deduplicateVariables(allVariables);
    console.log(`[AI] 去重后: ${deduplicated.length} 个变量`);
    
    return deduplicated;
}

/**
 * 变量去重：合并同义词和重复变量
 */
function deduplicateVariables(variables) {
    const uniqueMap = new Map(); // key: tag, value: variable
    
    for (const v of variables) {
        const tag = generatePinyinTag(v.label);
        
        if (!uniqueMap.has(tag)) {
            // 第一次出现，直接加入
            uniqueMap.set(tag, v);
        } else {
            // 已存在，合并 alternativeTexts
            const existing = uniqueMap.get(tag);
            if (v.text !== existing.text) {
                // 不同的表述形式，加入 alternativeTexts
                if (!existing.alternativeTexts) {
                    existing.alternativeTexts = [];
                }
                if (!existing.alternativeTexts.includes(v.text)) {
                    existing.alternativeTexts.push(v.text);
                }
            }
            // 合并 alternativeTexts
            if (v.alternativeTexts && v.alternativeTexts.length > 0) {
                if (!existing.alternativeTexts) {
                    existing.alternativeTexts = [];
                }
                v.alternativeTexts.forEach(alt => {
                    if (!existing.alternativeTexts.includes(alt) && alt !== existing.text) {
                        existing.alternativeTexts.push(alt);
                    }
                });
            }
        }
    }
    
    return Array.from(uniqueMap.values());
}

/**
 * 内置常用汉字拼音映射表
 * 覆盖合同文档中最常见的汉字，提供可读的拼音 tag
 */
const PINYIN_MAP = {
    // 合同基础词汇
    '合': 'He', '同': 'Tong', '协': 'Xie', '议': 'Yi', '约': 'Yue', '定': 'Ding',
    '编': 'Bian', '号': 'Hao', '码': 'Ma', '签': 'Qian', '署': 'Shu', '订': 'Ding',
    '条': 'Tiao', '款': 'Kuan', '项': 'Xiang', '章': 'Zhang', '节': 'Jie', '附': 'Fu',
    '件': 'Jian', '录': 'Lu', '表': 'Biao', '单': 'Dan', '书': 'Shu', '证': 'Zheng',
    
    // 公司组织
    '公': 'Gong', '司': 'Si', '企': 'Qi', '业': 'Ye', '法': 'Fa', '人': 'Ren',
    '股': 'Gu', '东': 'Dong', '投': 'Tou', '资': 'Zi', '方': 'Fang', '代': 'Dai',
    '表': 'Biao', '理': 'Li', '机': 'Ji', '构': 'Gou', '组': 'Zu', '织': 'Zhi',
    '团': 'Tuan', '队': 'Dui', '部': 'Bu', '门': 'Men', '员': 'Yuan', '工': 'Gong',
    
    // 甲乙方及身份
    '甲': 'Jia', '乙': 'Yi', '丙': 'Bing', '丁': 'Ding', '委': 'Wei', '托': 'Tuo',
    '受': 'Shou', '授': 'Shou', '权': 'Quan', '负': 'Fu', '责': 'Ze', '主': 'Zhu',
    '管': 'Guan', '联': 'Lian', '系': 'Xi', '接': 'Jie', '收': 'Shou', '发': 'Fa',
    
    // 名称标识
    '名': 'Ming', '称': 'Cheng', '姓': 'Xing', '字': 'Zi', '号': 'Hao', '标': 'Biao',
    '题': 'Ti', '目': 'Mu', '品': 'Pin', '牌': 'Pai', '型': 'Xing', '式': 'Shi',
    
    // 地址位置
    '地': 'Di', '址': 'Zhi', '省': 'Sheng', '市': 'Shi', '区': 'Qu', '县': 'Xian',
    '街': 'Jie', '道': 'Dao', '路': 'Lu', '巷': 'Xiang', '弄': 'Long', '楼': 'Lou',
    '室': 'Shi', '所': 'Suo', '在': 'Zai', '处': 'Chu', '场': 'Chang', '位': 'Wei',
    
    // 联系方式
    '电': 'Dian', '话': 'Hua', '手': 'Shou', '机': 'Ji', '邮': 'You', '箱': 'Xiang',
    '网': 'Wang', '站': 'Zhan', '传': 'Chuan', '真': 'Zhen', '码': 'Ma', '信': 'Xin',
    
    // 金额财务
    '金': 'Jin', '额': 'E', '价': 'Jia', '格': 'Ge', '费': 'Fei', '用': 'Yong',
    '款': 'Kuan', '币': 'Bi', '元': 'Yuan', '角': 'Jiao', '分': 'Fen', '钱': 'Qian',
    '本': 'Ben', '息': 'Xi', '率': 'Lv', '税': 'Shui', '账': 'Zhang', '户': 'Hu',
    '行': 'Hang', '银': 'Yin', '支': 'Zhi', '付': 'Fu', '收': 'Shou', '入': 'Ru',
    '出': 'Chu', '转': 'Zhuan', '汇': 'Hui', '票': 'Piao', '据': 'Ju', '证': 'Zheng',
    
    // 日期时间
    '年': 'Nian', '月': 'Yue', '日': 'Ri', '时': 'Shi', '分': 'Fen', '秒': 'Miao',
    '期': 'Qi', '限': 'Xian', '间': 'Jian', '前': 'Qian', '后': 'Hou', '始': 'Shi',
    '终': 'Zhong', '止': 'Zhi', '内': 'Nei', '外': 'Wai', '至': 'Zhi', '到': 'Dao',
    
    // 数字量词
    '一': 'Yi', '二': 'Er', '三': 'San', '四': 'Si', '五': 'Wu',
    '六': 'Liu', '七': 'Qi', '八': 'Ba', '九': 'Jiu', '十': 'Shi',
    '百': 'Bai', '千': 'Qian', '万': 'Wan', '亿': 'Yi', '零': 'Ling',
    '个': 'Ge', '张': 'Zhang', '份': 'Fen', '件': 'Jian', '次': 'Ci',
    '批': 'Pi', '套': 'Tao', '组': 'Zu', '双': 'Shuang', '对': 'Dui',
    
    // 产品服务
    '产': 'Chan', '品': 'Pin', '货': 'Huo', '物': 'Wu', '服': 'Fu', '务': 'Wu',
    '项': 'Xiang', '目': 'Mu', '工': 'Gong', '程': 'Cheng', '作': 'Zuo', '业': 'Ye',
    '任': 'Ren', '务': 'Wu', '内': 'Nei', '容': 'Rong', '范': 'Fan', '围': 'Wei',
    
    // 质量标准
    '质': 'Zhi', '量': 'Liang', '标': 'Biao', '准': 'Zhun', '规': 'Gui', '格': 'Ge',
    '检': 'Jian', '验': 'Yan', '测': 'Ce', '试': 'Shi', '认': 'Ren', '证': 'Zheng',
    '合': 'He', '格': 'Ge', '良': 'Liang', '好': 'Hao', '优': 'You', '秀': 'Xiu',
    
    // 交付履行
    '交': 'Jiao', '付': 'Fu', '货': 'Huo', '运': 'Yun', '输': 'Shu', '送': 'Song',
    '提': 'Ti', '取': 'Qu', '履': 'Lv', '行': 'Xing', '完': 'Wan', '成': 'Cheng',
    '验': 'Yan', '收': 'Shou', '确': 'Que', '认': 'Ren', '批': 'Pi', '准': 'Zhun',
    
    // 权利义务
    '权': 'Quan', '利': 'Li', '义': 'Yi', '务': 'Wu', '责': 'Ze', '任': 'Ren',
    '享': 'Xiang', '有': 'You', '承': 'Cheng', '担': 'Dan', '保': 'Bao', '证': 'Zheng',
    '赔': 'Pei', '偿': 'Chang', '损': 'Sun', '失': 'Shi', '害': 'Hai', '违': 'Wei',
    
    // 变更终止
    '变': 'Bian', '更': 'Geng', '修': 'Xiu', '改': 'Gai', '补': 'Bu', '充': 'Chong',
    '终': 'Zhong', '止': 'Zhi', '解': 'Jie', '除': 'Chu', '续': 'Xu', '延': 'Yan',
    '期': 'Qi', '满': 'Man', '届': 'Jie', '到': 'Dao', '废': 'Fei', '销': 'Xiao',
    
    // 争议解决
    '争': 'Zheng', '议': 'Yi', '纠': 'Jiu', '纷': 'Fen', '仲': 'Zhong', '裁': 'Cai',
    '诉': 'Su', '讼': 'Song', '调': 'Tiao', '解': 'Jie', '协': 'Xie', '商': 'Shang',
    
    // 保密知识产权
    '密': 'Mi', '秘': 'Mi', '知': 'Zhi', '识': 'Shi', '专': 'Zhuan', '利': 'Li',
    '技': 'Ji', '术': 'Shu', '商': 'Shang', '标': 'Biao', '版': 'Ban', '著': 'Zhu',
    
    // 其他常用字
    '的': 'De', '和': 'He', '与': 'Yu', '及': 'Ji', '或': 'Huo', '者': 'Zhe',
    '为': 'Wei', '由': 'You', '从': 'Cong', '按': 'An', '照': 'Zhao', '依': 'Yi',
    '据': 'Ju', '根': 'Gen', '基': 'Ji', '于': 'Yu', '关': 'Guan', '于': 'Yu',
    '对': 'Dui', '向': 'Xiang', '经': 'Jing', '过': 'Guo', '通': 'Tong', '知': 'Zhi',
    '报': 'Bao', '告': 'Gao', '说': 'Shuo', '明': 'Ming', '示': 'Shi', '意': 'Yi',
    '见': 'Jian', '求': 'Qiu', '需': 'Xu', '要': 'Yao', '应': 'Ying', '该': 'Gai',
    '须': 'Xu', '必': 'Bi', '可': 'Ke', '能': 'Neng', '得': 'De', '将': 'Jiang',
    '被': 'Bei', '给': 'Gei', '让': 'Rang', '使': 'Shi', '令': 'Ling', '使': 'Shi',
    '原': 'Yuan', '料': 'Liao', '材': 'Cai', '包': 'Bao', '装': 'Zhuang', '箱': 'Xiang'
};

/**
 * 简单哈希函数：将字符串转换为唯一的 ASCII 标识符
 * 使用 djb2 算法变体，生成可读的字母数字组合
 */
function simpleHash(str) {
    let hash = 5381;
    for (let i = 0; i < str.length; i++) {
        hash = ((hash << 5) + hash) + str.charCodeAt(i);
        hash = hash & hash; // Convert to 32bit integer
    }
    // 转换为正数并使用 base36（0-9a-z）
    const positiveHash = Math.abs(hash);
    return positiveHash.toString(36);
}

/**
 * 使用内置映射表将汉字转换为拼音
 * 对于映射表中没有的字符，使用哈希
 */
function chineseToPinyinWithMap(text) {
    let result = '';
    let unknownChars = '';
    
    for (let char of text) {
        if (PINYIN_MAP[char]) {
            result += PINYIN_MAP[char];
        } else if (/[a-zA-Z0-9]/.test(char)) {
            // 保留英文数字
            result += char.toUpperCase();
        } else if (/[\u4e00-\u9fa5]/.test(char)) {
            // 中文字符但不在映射表中
            unknownChars += char;
        }
        // 其他字符（空格、符号等）忽略
    }
    
    // 如果有未知字符，用哈希补充
    if (unknownChars) {
        const hash = simpleHash(unknownChars);
        result += '_' + hash;
    }
    
    return result || `field_${simpleHash(text)}`;
}

/**
 * 生成拼音 tag（PascalCase）
 * 优先级：1. pinyin-pro CDN  2. 内置映射表  3. ASCII 哈希
 */
function generatePinyinTag(label) {
    // 尝试多种可能的全局变量名
    let pinyinLib = null;
    if (typeof pinyinPro !== 'undefined') pinyinLib = pinyinPro;
    else if (typeof window !== 'undefined' && typeof window.pinyinPro !== 'undefined') pinyinLib = window.pinyinPro;
    else if (typeof pinyin !== 'undefined') pinyinLib = pinyin;
    else if (typeof window !== 'undefined' && typeof window.pinyin !== 'undefined') pinyinLib = window.pinyin;
    
    // 方案 1: 使用 pinyin-pro（如果可用）
    if (pinyinLib) {
        try {
            const pinyinFunc = pinyinLib.pinyin || pinyinLib;
            const pinyinResult = pinyinFunc(label, { toneType: 'none', type: 'array' });
            
            // 将每个字的拼音首字母大写并连接
            const result = (Array.isArray(pinyinResult) ? pinyinResult : pinyinResult.split(' '))
                .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
                .join('');
            
            console.log(`[PinyinTag] CDN转换: "${label}" -> "${result}"`);
            return result;
        } catch (e) {
            console.warn(`[PinyinTag] CDN转换失败，使用内置映射:`, e.message);
        }
    }
    
    // 方案 2: 使用内置映射表
    const mapResult = chineseToPinyinWithMap(label);
    console.log(`[PinyinTag] 映射表转换: "${label}" -> "${mapResult}"`);
    return mapResult;
}

/**
 * 将 AI 识别的字段存储到文档 Settings
 */
async function saveAIFieldsToDocument(fields) {
    console.log(`[AI Storage] 保存 ${fields.length} 个 AI 字段到文档...`);
    
    await Word.run(async (context) => {
        const settings = context.document.settings;
        await saveToSettingsChunked(context, settings, AI_FIELDS_KEY, JSON.stringify(fields));
        await context.sync();
    });
    
    console.log("[AI Storage] 保存成功");
}

/**
 * 从文档 Settings 加载 AI 识别的字段
 */
async function loadAIFieldsFromDocument() {
    console.log("[AI Storage] 从文档加载 AI 字段...");
    
    try {
        return await Word.run(async (context) => {
            const settings = context.document.settings;
            settings.load("items");
            await context.sync();
            
            const data = await readFromSettingsChunked(context, settings, AI_FIELDS_KEY);
            if (data) {
                const fields = JSON.parse(data);
                console.log(`[AI Storage] 加载了 ${fields.length} 个 AI 字段`);
                return fields;
            }
            return [];
        });
    } catch (error) {
        console.warn("[AI Storage] 加载失败:", error);
        return [];
    }
}

/**
 * 在表单中渲染 AI 识别的字段
 */
function renderAIFieldsInForm(aiFields) {
    console.log(`[AI Form] 渲染 ${aiFields.length} 个 AI 字段到表单...`);
    
    // 清除之前的 AI 字段（如果有）
    document.querySelectorAll('.ai-field-wrapper').forEach(el => el.remove());
    document.querySelectorAll('.ai-fields-section').forEach(el => el.remove());
    
    if (aiFields.length === 0) return;
    
    // 去重：同一 tag 只保留第一个
    const uniqueFields = [];
    const seenTags = new Set();
    
    for (const field of aiFields) {
        const tag = generatePinyinTag(field.label);
        if (!seenTags.has(tag)) {
            seenTags.add(tag);
            uniqueFields.push(field);
        }
    }
    
    console.log(`[AI Form] 去重后: ${uniqueFields.length} 个字段`);
    
    // 按 sectionId 分组
    const grouped = {};
    uniqueFields.forEach(f => {
        if (!grouped[f.sectionId]) grouped[f.sectionId] = [];
        grouped[f.sectionId].push(f);
    });
    
    // 为每个 section 添加 AI 字段（直接混入，不创建单独区域）
    for (const [sectionId, fields] of Object.entries(grouped)) {
        // 查找对应的 section 容器
        const sectionHeader = document.querySelector(`[data-section-id="${sectionId}"]`);
        let targetContainer = null;
        
        if (sectionHeader) {
            // 找到 section 的字段容器
            targetContainer = sectionHeader.closest('.form-section')?.querySelector('.section-fields');
        }
        
        // 如果找不到特定 section，就放在表单容器末尾
        if (!targetContainer) {
            targetContainer = document.getElementById('dynamic-form-container');
        }
        
        if (!targetContainer) continue;
        
        // 分离父级和子级字段
        const parentFields = fields.filter(f => !f.parentTag);
        const childFields = fields.filter(f => f.parentTag);
        
        // 直接将 AI 字段插入到 section 中（不创建单独区域）
        parentFields.forEach(field => {
            const wrapper = createAIFieldElement(field);
            targetContainer.appendChild(wrapper);
            
            // 如果有子字段，渲染它们（缩进）
            const childrenOfThis = childFields.filter(c => 
                c.parentTag === generatePinyinTag(field.label)
            );
            childrenOfThis.forEach(child => {
                const childWrapper = createAIFieldElement(child, true);
                targetContainer.appendChild(childWrapper);
            });
        });
        
        // 渲染没有父级的子字段
        const orphanChildren = childFields.filter(c => {
            const hasParent = parentFields.some(p => 
                generatePinyinTag(p.label) === c.parentTag
            );
            return !hasParent;
        });
        orphanChildren.forEach(child => {
            const wrapper = createAIFieldElement(child);
            targetContainer.appendChild(wrapper);
        });
    }
    
    console.log("[AI Form] 渲染完成");
}

/**
 * 创建单个 AI 字段的表单元素
 * @param {Object} field 字段对象
 * @param {boolean} isNested 是否为嵌套字段（子字段需要缩进）
 */
function createAIFieldElement(field, isNested = false) {
    // 生成 tag（拼音驼峰）
    const pinyinTag = generatePinyinTag(field.label);
    const mode = field.mode || 'field';
    
    // 根据 mode 确定颜色
    const bgColor = mode === 'paragraph' ? 'rgba(245, 158, 11, 0.05)' : 'rgba(59, 130, 246, 0.05)';
    const borderColor = mode === 'paragraph' ? 'rgba(245, 158, 11, 0.3)' : 'rgba(59, 130, 246, 0.3)';
    const labelColor = mode === 'paragraph' ? '#f59e0b' : '#3b82f6';
    const badgeColor = mode === 'paragraph' ? '#f59e0b' : '#6366f1';
    const badgeText = mode === 'paragraph' ? '可选' : 'AI';
    
    const wrapper = document.createElement('div');
    wrapper.className = 'field-wrapper ai-field-wrapper';
    wrapper.dataset.aiField = 'true';
    wrapper.dataset.tag = pinyinTag;
    wrapper.dataset.mode = mode;
    wrapper.style.cssText = `
        position: relative;
        padding: 8px 12px;
        margin: 6px 0;
        ${isNested ? 'margin-left: 24px;' : ''}
        background: ${bgColor};
        border: 1px solid ${borderColor};
        border-radius: 6px;
    `;
    
    // 模式标签
    const modeBadge = document.createElement('span');
    modeBadge.className = 'ai-badge';
    modeBadge.textContent = badgeText;
    modeBadge.style.cssText = `
        position: absolute;
        top: -6px;
        right: 8px;
        background: ${badgeColor};
        color: white;
        font-size: 10px;
        padding: 1px 6px;
        border-radius: 3px;
    `;
    wrapper.appendChild(modeBadge);
    
    // 删除按钮
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'ai-field-delete';
    deleteBtn.innerHTML = '<i class="ms-Icon ms-Icon--Cancel"></i>';
    deleteBtn.style.cssText = `
        position: absolute;
        top: -6px;
        left: 8px;
        background: #ef4444;
        color: white;
        border: none;
        width: 18px;
        height: 18px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
    `;
    deleteBtn.title = '删除此字段';
    deleteBtn.onclick = () => removeAIField(pinyinTag, wrapper);
    wrapper.appendChild(deleteBtn);
    
    // Label（双击跳转到文档中的对应位置）
    const label = document.createElement('label');
    label.textContent = field.label;
    label.style.cssText = `
        display: block;
        font-size: 12px;
        color: ${labelColor};
        margin-bottom: 4px;
        font-weight: 500;
        cursor: pointer;
    `;
    label.title = '双击跳转到文档中的对应位置';
    label.ondblclick = () => scrollToContentControl(pinyinTag);
    wrapper.appendChild(label);
    
    if (mode === 'paragraph') {
        // paragraph 模式：渲染开关按钮
        const toggleContainer = document.createElement('div');
        toggleContainer.style.cssText = 'display: flex; align-items: center; gap: 8px;';
        
        const toggleBtn = document.createElement('button');
        toggleBtn.className = 'toggle-paragraph-btn';
        toggleBtn.dataset.tag = pinyinTag;
        toggleBtn.textContent = '显示段落';
        toggleBtn.style.cssText = `
            padding: 6px 12px;
            background: ${badgeColor};
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        `;
        
        toggleBtn.onclick = async () => {
            // 调用现有的 toggle 机制
            await toggleParagraph(pinyinTag, toggleBtn);
        };
        
        toggleContainer.appendChild(toggleBtn);
        wrapper.appendChild(toggleContainer);
        
    } else {
        // field 模式：渲染输入框
        let input;
        if (field.type === 'date') {
            input = document.createElement('input');
            input.type = 'date';
        } else if (field.type === 'number') {
            input = document.createElement('input');
            input.type = 'number';
        } else {
            input = document.createElement('input');
            input.type = 'text';
        }
        
        input.id = `ai_${pinyinTag}`;
        input.name = pinyinTag;
        input.dataset.tag = pinyinTag;
        input.placeholder = field.text || field.label;
        input.style.cssText = `
            width: 100%;
            padding: 6px 8px;
            border: 1px solid ${borderColor};
            border-radius: 4px;
            font-size: 13px;
            background: white;
        `;
        
        // 输入时同步到文档
        // #region agent log
        input.addEventListener('input', debounce(async () => {
            console.log(`[DEBUG-INPUT] AI字段输入触发: tag=${pinyinTag}, value="${input.value}"`);
            await applyAIFieldToDocument(pinyinTag, input.value || `[${field.label}]`);
        }, 500));
        // #endregion
        
        wrapper.appendChild(input);
    }
    
    return wrapper;
}

/**
 * 注册 Content Control 进入事件
 * 当用户在文档中点击 Content Control 时，侧边栏滚动到对应的表单字段
 */
async function registerContentControlEvents() {
    if (typeof Word === 'undefined') {
        console.warn("[Events] Word API 不可用");
        return;
    }
    
    try {
        await Word.run(async (context) => {
            // 注册文档级别的 ContentControl 进入事件
            context.document.onContentControlEntered.add(handleContentControlEntered);
            await context.sync();
            console.log("[Events] ✓ Content Control 进入事件已注册");
        });
    } catch (error) {
        console.warn("[Events] 注册事件失败（可能不支持此 API）:", error.message);
    }
}

/**
 * Content Control 进入事件处理函数
 * @param {Word.ContentControlEnteredEventArgs} event 事件参数
 */
async function handleContentControlEntered(event) {
    console.log("[Events] Content Control 进入事件触发");
    
    try {
        await Word.run(async (context) => {
            // 获取被点击的 Content Control
            const ccId = event.ids[0];
            if (!ccId) return;
            
            const cc = context.document.contentControls.getById(ccId);
            cc.load("tag, title");
            await context.sync();
            
            const tag = cc.tag;
            if (!tag) return;
            
            console.log(`[Events] 进入 Content Control: tag=${tag}, title=${cc.title}`);
            
            // 在侧边栏中找到对应的表单字段并滚动到它
            scrollToFormField(tag);
        });
    } catch (error) {
        console.warn("[Events] 处理进入事件失败:", error.message);
    }
}

/**
 * 滚动侧边栏到指定 tag 的表单字段
 * @param {string} tag 字段的 tag
 */
function scrollToFormField(tag) {
    // 尝试多种选择器找到对应的表单字段
    const selectors = [
        `input[data-tag="${tag}"]`,
        `input[name="${tag}"]`,
        `#ai_${tag}`,
        `button[data-tag="${tag}"]`,
        `.ai-field-wrapper[data-tag="${tag}"]`
    ];
    
    let element = null;
    for (const selector of selectors) {
        element = document.querySelector(selector);
        if (element) break;
    }
    
    if (element) {
        // 滚动到视图中并高亮
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
        
        // 添加临时高亮效果
        const wrapper = element.closest('.ai-field-wrapper') || element.closest('.field-wrapper') || element.parentElement;
        if (wrapper) {
            wrapper.style.transition = 'background-color 0.3s';
            wrapper.style.backgroundColor = '#fef3c7'; // 淡黄色高亮
            setTimeout(() => {
                wrapper.style.backgroundColor = '';
            }, 2000);
        }
        
        // 聚焦输入框
        if (element.tagName === 'INPUT') {
            element.focus();
        }
        
        console.log(`[Events] ✓ 已滚动到表单字段: ${tag}`);
    } else {
        console.log(`[Events] 未找到表单字段: ${tag}`);
    }
}

/**
 * 滚动文档到指定 tag 的 Content Control 位置
 * @param {string} tag Content Control 的 tag
 */
async function scrollToContentControl(tag) {
    console.log(`[Scroll] 跳转到 Content Control: ${tag}`);
    
    if (typeof Word === 'undefined') {
        console.warn("[Scroll] Word API 不可用");
        return;
    }
    
    try {
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls;
            contentControls.load("items");
            await context.sync();
            
            // 找到匹配的 Content Control
            for (const cc of contentControls.items) {
                cc.load("tag");
            }
            await context.sync();
            
            const targetCC = contentControls.items.find(cc => cc.tag === tag);
            if (targetCC) {
                // 选中该 Content Control，Word 会自动滚动到可视区域
                targetCC.select("Select");
                await context.sync();
                console.log(`[Scroll] ✓ 已跳转到: ${tag}`);
            } else {
                console.warn(`[Scroll] 未找到 Content Control: ${tag}`);
                showNotification(`未找到对应的埋点: ${tag}`, "warning", 2000);
            }
        });
    } catch (error) {
        console.error("[Scroll] 跳转失败:", error);
    }
}

/**
 * 切换 paragraph 段落的显示/隐藏
 */
async function toggleParagraph(tag, button) {
    try {
        // 复用现有的 toggle 逻辑
        // 这里简化实现，实际可以调用现有的 toggleRoundVisibility 等函数
        const isHidden = button.textContent === '显示段落';
        
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls;
            contentControls.load("items");
            await context.sync();
            
            for (const cc of contentControls.items) {
                cc.load("tag");
            }
            await context.sync();
            
            for (const cc of contentControls.items) {
                if (cc.tag === tag) {
                    if (isHidden) {
                        // 显示：恢复内容
                        cc.load("text");
                        await context.sync();
                        // 这里简化，实际可能需要从 Settings 恢复
                    } else {
                        // 隐藏：清空内容
                        await insertTextPreserveFormat(cc, '', context);
                    }
                }
            }
            await context.sync();
        });
        
        // 切换按钮状态
        button.textContent = isHidden ? '隐藏段落' : '显示段落';
        button.style.background = isHidden ? '#ef4444' : '#f59e0b';
        
    } catch (error) {
        console.error('[Toggle] 切换失败:', error);
        showNotification(`切换失败: ${error.message}`, 'error');
    }
}

/**
 * 将 AI 字段值同步到文档中的 Content Control
 */
async function applyAIFieldToDocument(tag, value) {
    // #region agent log
    console.log(`[DEBUG-SYNC] applyAIFieldToDocument 开始: tag="${tag}", value="${value}"`);
    // #endregion
    try {
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls;
            contentControls.load("items");
            await context.sync();
            
            // #region agent log
            console.log(`[DEBUG-SYNC] 文档中共有 ${contentControls.items.length} 个 Content Control`);
            // #endregion
            
            for (const cc of contentControls.items) {
                cc.load("tag");
            }
            await context.sync();
            
            // #region agent log
            const allTags = contentControls.items.map(cc => cc.tag);
            console.log(`[DEBUG-SYNC] 所有 CC tags: ${JSON.stringify(allTags.slice(0, 10))}...`);
            // #endregion
            
            let foundMatch = false;
            for (const cc of contentControls.items) {
                if (cc.tag === tag) {
                    foundMatch = true;
                    await insertTextPreserveFormat(cc, value, context);
                    console.log(`[AI Sync] 已更新 ${tag}: ${value}`);
                }
            }
            
            // #region agent log
            if (!foundMatch) {
                console.warn(`[DEBUG-SYNC] 未找到匹配的 Content Control: tag="${tag}"`);
            }
            // #endregion
            
            await context.sync();
        });
    } catch (error) {
        console.warn(`[AI Sync] 同步失败 (${tag}):`, error);
    }
}

/**
 * 删除 AI 识别的字段
 */
async function removeAIField(tag, wrapperElement) {
    // 从 UI 移除
    wrapperElement.remove();
    
    // 从存储中移除
    try {
        const fields = await loadAIFieldsFromDocument();
        const updatedFields = fields.filter(f => {
            const fTag = pinyin(f.label, { toneType: 'none' })
                .split(' ')
                .map(w => w.charAt(0).toUpperCase() + w.slice(1))
                .join('');
            return fTag !== tag;
        });
        await saveAIFieldsToDocument(updatedFields);
        console.log(`[AI] 已删除字段: ${tag}`);
    } catch (error) {
        console.warn(`[AI] 删除字段失败:`, error);
    }
}

/**
 * 读取当前文档的文本内容
 */
async function getDocumentText() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        return body.text;
    });
}

/**
 * AI 识别并埋点主流程
 */
async function aiRecognizeAndEmbed() {
    console.log("[AI] ========== 开始 AI 智能识别 ==========");
    
    let successCount = 0;
    let skippedCount = 0;
    
    try {
        // Step 1: 读取文档文本
        showNotification("正在读取文档...", "info");
        const docText = await getDocumentText();
        console.log(`[AI] 文档文本长度: ${docText.length}`);
        
        if (docText.length < 50) {
            showNotification("文档内容太少，无法进行 AI 分析", "warning");
            return;
        }
        
        // Step 2: 调用 AI 分析
        showNotification("AI 正在分析文档...(可能需要 10-30 秒)", "info", 30000);
        const aiResult = await callDoubaoAI(docText);
        
        if (aiResult.length === 0) {
            showNotification("AI 未识别到需要填写的变量", "info");
            return;
        }
        
        showNotification(`AI 识别到 ${aiResult.length} 个变量，正在埋点...`, "info");
        
        // Step 3: 在文档中创建 Content Control
        await Word.run(async (context) => {
            for (const variable of aiResult) {
                try {
                    // 搜索变量文本
                    const searchResults = context.document.body.search(variable.text, { 
                        matchCase: false,
                        matchWholeWord: false 
                    });
                    searchResults.load("items");
                    await context.sync();
                    
                    if (searchResults.items.length === 0) {
                        console.log(`[AI] 未找到文本: "${variable.text}"`);
                        skippedCount++;
                        continue;
                    }
                    
                    // 处理第一个匹配项
                    const range = searchResults.items[0];
                    range.load("text, parentContentControlOrNullObject");
                    await context.sync();
                    
                    // 检查是否已有埋点
                    const parentCC = range.parentContentControlOrNullObject;
                    parentCC.load("isNullObject, tag");
                    await context.sync();
                    
                    if (!parentCC.isNullObject) {
                        console.log(`[AI] 跳过已埋点: "${variable.text}"`);
                        skippedCount++;
                        continue;
                    }
                    
                    // 生成拼音 tag（使用统一函数确保与表单一致）
                    const pinyinTag = generatePinyinTag(variable.label);
                    
                    // 创建 Content Control
                    const cc = range.insertContentControl("RichText");
                    cc.tag = pinyinTag;
                    cc.title = variable.label;
                    cc.appearance = "Tags"; // Tags 模式显示 title 标签
                    cc.color = "#6366f1"; // 紫色，区分 AI 识别
                    
                    await context.sync();
                    successCount++;
                    console.log(`[AI] ✓ 埋点成功: ${variable.label} → ${pinyinTag}`);
                    
                } catch (err) {
                    console.warn(`[AI] 处理变量失败 (${variable.text}):`, err.message);
                    skippedCount++;
                }
            }
            
            await context.sync();
        });
        
        // Step 4: 存储字段配置到文档
        await saveAIFieldsToDocument(aiResult);
        
        // Step 5: 更新表单 UI
        renderAIFieldsInForm(aiResult);
        
        // 完成通知
        const message = `AI 识别完成！\n成功埋点: ${successCount} 个\n跳过: ${skippedCount} 个`;
        showNotification(message, successCount > 0 ? "success" : "info", 5000);
        console.log(`[AI] ========== 完成: 成功 ${successCount}, 跳过 ${skippedCount} ==========`);
        
    } catch (error) {
        console.error("[AI] 识别失败:", error);
        showNotification(`AI 识别失败: ${error.message || error}`, "error");
    }
}

/**
 * 确认对话框：AI 智能识别
 */
async function confirmAIRecognize() {
    const confirmed = await showConfirmDialog(
        "AI 智能识别变量",
        "将使用 AI 分析当前文档，自动识别并埋点所有变量。\n\n建议先执行【占位符】埋点，再执行此操作。\n\n此操作需要联网，可能需要 10-30 秒。\n\n确定继续吗？"
    );
    
    if (confirmed) {
        showNotification("正在启动 AI 分析...", "info");
        await aiRecognizeAndEmbed();
    }
}

// 暴露到全局
window.confirmAIRecognize = confirmAIRecognize;
window.aiRecognizeAndEmbed = aiRecognizeAndEmbed;

// ================================================================================
// 自动生成表单 - 一键埋点（【】+ AI）
// ================================================================================

// 记录本次自动生成创建的埋点 tag
let autoGeneratedTags = [];

/**
 * 获取文档中所有 Content Control 的 tag 列表
 */
async function getAllContentControlTags() {
    if (typeof Word === 'undefined') return [];
    
    // #region agent log
    console.log("[DEBUG-H4] getAllContentControlTags: 开始获取所有CC的tag");
    // #endregion
    
    return await Word.run(async (context) => {
        const ccs = context.document.contentControls;
        ccs.load("items");
        await context.sync();
        
        // #region agent log
        console.log(`[DEBUG-H4] 找到 ${ccs.items.length} 个 Content Controls`);
        // #endregion
        
        const tags = [];
        for (let i = 0; i < ccs.items.length; i++) {
            try {
                ccs.items[i].load("tag");
            } catch (loadErr) {
                // #region agent log
                console.log(`[DEBUG-H4] load tag 失败 index=${i}: ${loadErr.message}`);
                // #endregion
            }
        }
        
        try {
            await context.sync();
        } catch (syncErr) {
            // #region agent log
            console.error(`[DEBUG-H4] sync 失败: ${syncErr.message}`, syncErr);
            // #endregion
            throw syncErr;
        }
        
        for (let i = 0; i < ccs.items.length; i++) {
            try {
                if (ccs.items[i].tag) tags.push(ccs.items[i].tag);
            } catch (readErr) {
                // #region agent log
                console.log(`[DEBUG-H4] 读取 tag 失败 index=${i}: ${readErr.message}`);
                // #endregion
            }
        }
        
        // #region agent log
        console.log(`[DEBUG-H4] 成功获取 ${tags.length} 个 tags`);
        // #endregion
        
        return tags;
    });
}

/**
 * AI 识别核心逻辑（不含确认对话框）
 */
async function aiRecognizeCore() {
    console.log("[AI Core] ========== 开始 AI 智能识别 ==========");
    
    let successCount = 0;
    let skippedCount = 0;
    
    // Step 1: 读取文档文本
    const docText = await getDocumentText();
    console.log(`[AI Core] 文档文本长度: ${docText.length}`);
    
    if (docText.length < 50) {
        console.log("[AI Core] 文档内容太少，跳过 AI 分析");
        return { success: 0, skipped: 0 };
    }
    
    // Step 2: 调用 AI 分析（支持全文分段）
    const aiResult = await analyzeFullDocument(docText);
    
    if (aiResult.length === 0) {
        console.log("[AI Core] AI 未识别到变量");
        return { success: 0, skipped: 0 };
    }
    
    console.log(`[AI Core] AI 识别到 ${aiResult.length} 个变量（已去重）`);
    
    // Step 3: 为每个变量创建 Content Control（支持多处埋点）
    // 使用单独的 Word.run 来隔离每个变量的处理，避免一个失败影响全部
    for (const variable of aiResult) {
        try {
            // 生成统一的 tag
            const pinyinTag = generatePinyinTag(variable.label);
            
            // 根据 mode 确定颜色
            const color = variable.mode === 'paragraph' ? '#f59e0b' : '#3b82f6'; // 橙色/蓝色
            const title = variable.mode === 'paragraph' ? `[可选] ${variable.label}` : variable.label;
            
            // 收集所有需要搜索的文本（原文 + 同义词）
            const searchTexts = [variable.text];
            if (variable.alternativeTexts && variable.alternativeTexts.length > 0) {
                searchTexts.push(...variable.alternativeTexts);
            }
            
            console.log(`[AI Core] 处理变量: ${variable.label}, 搜索 ${searchTexts.length} 种表述`);
            
            // 为每种表述搜索并创建埋点
            let embedCount = 0;
            
            for (const searchText of searchTexts) {
                // 跳过过于通用的搜索词（可能导致误匹配）
                if (searchText.length < 2) {
                    console.log(`[AI Core] 跳过过短文本: "${searchText}"`);
                    continue;
                }
                
                // 跳过纯空格/下划线的文本（Word search 对纯空格不可靠）
                const trimmed = searchText.replace(/[\s_]/g, '');
                if (trimmed.length === 0) {
                    console.log(`[AI Core] 跳过纯空格/下划线: "${searchText}" (无法可靠搜索)`);
                    continue;
                }
                
                try {
                    await Word.run(async (context) => {
                        const searchResults = context.document.body.search(searchText, { 
                            matchCase: false,
                            matchWholeWord: false 
                        });
                        searchResults.load("items");
                        await context.sync();
                        
                        if (searchResults.items.length === 0) {
                            console.log(`[AI Core] 未找到文本: "${searchText}"`);
                            return;
                        }
                        
                        // 限制每种表述最多处理前5个匹配（避免过多重复）
                        const maxMatches = Math.min(searchResults.items.length, 5);
                        
                        // 处理匹配项
                        for (let i = 0; i < maxMatches; i++) {
                            const range = searchResults.items[i];
                            // #region agent log
                            console.log(`[DEBUG-H5] 处理匹配项 ${i}/${maxMatches}, range存在=${!!range}`);
                            // #endregion
                            try {
                                // #region agent log
                                console.log(`[DEBUG-H6] 准备加载 range 属性, i=${i}`);
                                // #endregion
                                range.load("text, parentContentControlOrNullObject");
                                await context.sync();
                                // #region agent log
                                console.log(`[DEBUG-H5] 匹配项 ${i} 加载成功, text="${range.text?.substring(0,30)}"`);
                                // #endregion
                                
                                // #region agent log
                                console.log(`[DEBUG-H6] 准备加载 parentCC 属性`);
                                // #endregion
                                
                                // 检查是否已有埋点
                                const parentCC = range.parentContentControlOrNullObject;
                                // #region agent log
                                console.log(`[DEBUG-H6] parentCC 对象存在: ${!!parentCC}`);
                                // #endregion
                                parentCC.load("isNullObject, tag, title, color");
                                // #region agent log
                                console.log(`[DEBUG-H6] 准备 sync parentCC 属性...`);
                                // #endregion
                                await context.sync();
                                // #region agent log
                                console.log(`[DEBUG-H6] parentCC sync 完成, isNullObject=${parentCC.isNullObject}`);
                                // #endregion
                                
                                if (!parentCC.isNullObject) {
                                    // 已有埋点，检查是否需要修正
                                    // 注意：颜色比对忽略大小写（#3B82F6 和 #3b82f6 是同一个颜色）
                                    const tagMatch = parentCC.tag === pinyinTag;
                                    const titleMatch = parentCC.title === title;
                                    const colorMatch = (parentCC.color || '').toLowerCase() === (color || '').toLowerCase();
                                    
                                    // #region agent log
                                    console.log(`[DEBUG-H1] 比较CC: text="${searchText.substring(0,20)}", i=${i}`);
                                    console.log(`[DEBUG-H1]   oldTag="${parentCC.tag}" (type=${typeof parentCC.tag})`);
                                    console.log(`[DEBUG-H1]   newTag="${pinyinTag}" (type=${typeof pinyinTag})`);
                                    console.log(`[DEBUG-H1]   tagMatch=${tagMatch}, titleMatch=${titleMatch}, colorMatch=${colorMatch}`);
                                    console.log(`[DEBUG-H1]   oldTag===newTag: ${parentCC.tag === pinyinTag}`);
                                    console.log(`[DEBUG-H1]   oldTag JSON: ${JSON.stringify(parentCC.tag)}`);
                                    console.log(`[DEBUG-H1]   newTag JSON: ${JSON.stringify(pinyinTag)}`);
                                    // #endregion
                                    
                                    const needsUpdate = !tagMatch || !titleMatch;
                                    // 如果只是颜色大小写不同，不需要更新
                                    
                                    // #region agent log
                                    console.log(`[DEBUG-H1]   needsUpdate=${needsUpdate} (!tagMatch=${!tagMatch} || !titleMatch=${!titleMatch})`);
                                    // #endregion
                                    
                                    if (needsUpdate) {
                                        console.log(`[AI Core] 检测到不一致，重新埋点: "${searchText}"`);
                                        console.log(`  - tag: ${parentCC.tag} → ${pinyinTag} (${tagMatch ? '✓' : '✗'})`);
                                        console.log(`  - title: ${parentCC.title} → ${title} (${titleMatch ? '✓' : '✗'})`);
                                        
                                        try {
                                            // #region agent log
                                            console.log(`[DEBUG-H2] 准备删除旧CC并重新创建...`);
                                            // #endregion
                                            
                                            // 删除旧的 Content Control（保留文本）
                                            parentCC.delete(false);
                                            await context.sync();
                                            
                                            // #region agent log
                                            console.log(`[DEBUG-H2] 旧CC已删除，准备重新搜索...`);
                                            // #endregion
                                            
                                            // 重新搜索该位置
                                            const reSearch = context.document.body.search(searchText, { 
                                                matchCase: false,
                                                matchWholeWord: false 
                                            });
                                            reSearch.load("items");
                                            await context.sync();
                                            
                                            // #region agent log
                                            console.log(`[DEBUG-H2] 重新搜索结果: ${reSearch.items.length} 个匹配`);
                                            // #endregion
                                            
                                            if (reSearch.items.length > 0) {
                                                // 找到对应位置，创建新的 Content Control
                                                const newIndex = Math.min(i, reSearch.items.length - 1);
                                                // #region agent log
                                                console.log(`[DEBUG-H2] 使用索引 ${newIndex} (原索引 ${i})`);
                                                // #endregion
                                                const newRange = reSearch.items[newIndex];
                                                const newCC = newRange.insertContentControl("RichText");
                                                newCC.tag = pinyinTag;
                                                newCC.title = title;
                                                newCC.appearance = "Tags"; // Tags 模式显示 title 标签
                                                newCC.color = color;
                                                await context.sync();
                                                
                                                embedCount++;
                                                console.log(`[AI Core] ✓ 已修正埋点`);
                                            }
                                        } catch (updateErr) {
                                            // #region agent log
                                            console.error(`[DEBUG-H2] 修正埋点异常:`, updateErr);
                                            // #endregion
                                            console.warn(`[AI Core] 修正埋点失败: "${searchText}" - ${updateErr.message}`);
                                        }
                                    } else {
                                        console.log(`[AI Core] 已验证正确，保留: "${searchText}"`);
                                        embedCount++; // 计入成功数量
                                    }
                                    continue;
                                }
                                
                                // 创建 Content Control（所有同义词共享同一个 tag）
                                // #region agent log - 假设 B/C: 检查 tag 和 range 状态
                                const tagInfo = {
                                    pinyinTag,
                                    hasChinese: /[\u4e00-\u9fa5]/.test(pinyinTag),
                                    hasSpecialChars: /[（）()【】\[\]""]/.test(pinyinTag),
                                    tagLength: pinyinTag.length,
                                    rangeText: range.text?.substring(0, 50)
                                };
                                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:aiRecognizeCore:beforeInsertCC',message:'准备创建CC',data:tagInfo,timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'B'})}).catch(()=>{});
                                // #endregion
                                console.log(`[DEBUG-H6] 准备创建新 CC, tag=${pinyinTag}`);
                                
                                const cc = range.insertContentControl("RichText");
                                cc.tag = pinyinTag;
                                cc.title = title;
                                cc.appearance = "Tags"; // Tags 模式显示 title 标签
                                cc.color = color;
                                
                                console.log(`[DEBUG-H6] CC 属性已设置, 准备 sync...`);
                                // #region agent log - 假设 E: 检查 CC 对象状态
                                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:aiRecognizeCore:beforeSync',message:'CC属性已设置准备sync',data:{tag:pinyinTag,title,color,ccExists:!!cc},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'E'})}).catch(()=>{});
                                // #endregion
                                await context.sync();
                                // #region agent log
                                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:aiRecognizeCore:afterSync',message:'CC创建成功',data:{tag:pinyinTag},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'B'})}).catch(()=>{});
                                // #endregion
                                console.log(`[DEBUG-H6] ✓ CC 创建成功`);
                                embedCount++;
                            } catch (rangeErr) {
                                // 单个匹配项失败，继续处理下一个
                                // #region agent log - 假设 B/C/E: 详细错误信息
                                const errorInfo = {
                                    searchText: searchText?.substring(0, 50),
                                    matchIndex: i,
                                    pinyinTag,
                                    errorMessage: rangeErr.message,
                                    errorName: rangeErr.name,
                                    debugInfo: rangeErr.debugInfo ? JSON.stringify(rangeErr.debugInfo) : null,
                                    hasChinese: /[\u4e00-\u9fa5]/.test(pinyinTag),
                                    hasSpecialChars: /[（）()【】\[\]""]/.test(pinyinTag)
                                };
                                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:aiRecognizeCore:rangeError',message:'CC创建失败',data:errorInfo,timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'BCE'})}).catch(()=>{});
                                // #endregion
                                console.error(`[DEBUG-H5] 匹配项 ${i} 处理异常:`, rangeErr);
                                console.log(`[DEBUG-H5] 错误栈: ${rangeErr.stack}`);
                                console.log(`[AI Core] 单个匹配处理失败: "${searchText}" - ${rangeErr.message}`);
                            }
                        }
                    });
                } catch (searchErr) {
                    console.log(`[AI Core] 搜索失败: "${searchText}" - ${searchErr.message}`);
                }
            }
            
            if (embedCount > 0) {
                successCount++;
                console.log(`[AI Core] ✓ 埋点成功: ${variable.label} → ${pinyinTag} (${embedCount} 处)`);
            } else {
                skippedCount++;
                console.log(`[AI Core] 跳过: ${variable.label} (未找到任何匹配或全部失败)`);
            }
            
        } catch (err) {
            console.warn(`[AI Core] 处理变量失败 (${variable.label}):`, err.message);
            skippedCount++;
        }
    }
    
    // Step 4: 存储字段配置到文档
    await saveAIFieldsToDocument(aiResult);
    
    // Step 5: 更新表单 UI
    renderAIFieldsInForm(aiResult);
    
    console.log(`[AI Core] ========== 完成: 成功 ${successCount}, 跳过 ${skippedCount} ==========`);
    return { success: successCount, skipped: skippedCount };
}

/**
 * 自动生成表单主函数
 * 一键执行：备份 → 【】埋点 → AI 识别
 */
async function autoGenerateForm() {
    const btn = document.getElementById('btn-auto-generate');
    if (!btn || btn.disabled) return;
    
    btn.disabled = true;
    const originalHTML = btn.innerHTML;
    
    try {
        // Step 1: 尝试备份（复用现有逻辑）
        btn.innerHTML = '<i class="ms-Icon ms-Icon--Sync"></i> 正在备份...';
        console.log("[AutoGenerate] Step 1: 尝试备份...");
        
        const backupResult = await tryBackupCurrentDocument();
        
        if (backupResult.success) {
            if (backupResult.skippedByUser) {
                showNotification("用户选择跳过备份，继续执行", "info", 2000);
            } else {
                showNotification(`已备份: ${backupResult.fileName}`, "success", 3000);
            }
        } else {
            // 备份失败 - 提示用户手动备份
            let msg = backupResult.skipped 
                ? `⚠️ ${backupResult.reason || "无法自动备份"}\n\n`
                : `⚠️ 备份失败：${backupResult.error}\n\n`;
            msg += "请先手动备份当前文档：\n";
            msg += "1. 在 OneDrive 中找到此文件\n";
            msg += "2. 右键选择「复制」创建副本\n\n";
            msg += "确认已完成备份后，点击「确定」继续。";
            
            const confirmed = await showConfirmDialog(msg, {
                title: "⚠️ 需要手动备份",
                confirmText: "已备份，继续",
                cancelText: "取消"
            });
            
            if (!confirmed) {
                showNotification("操作已取消", "warning");
                return;
            }
        }
        
        // Step 2: 记录现有埋点
        btn.innerHTML = '<i class="ms-Icon ms-Icon--Sync"></i> 正在扫描...';
        console.log("[AutoGenerate] Step 2: 记录现有埋点...");
        const beforeTags = await getAllContentControlTags();
        console.log(`[AutoGenerate] 现有埋点数量: ${beforeTags.length}`);
        
        // Step 3: 执行 AI 识别（AI 会同时处理【】占位符和其他变量）
        btn.innerHTML = '<i class="ms-Icon ms-Icon--Sync"></i> AI 分析中...';
        console.log("[AutoGenerate] Step 3: 执行 AI 识别（包含【】占位符）...");
        const aiResult = await aiRecognizeCore();
        console.log(`[AutoGenerate] AI 识别结果: 成功 ${aiResult.success}, 跳过 ${aiResult.skipped}`);
        
        // Step 4: 记录新增埋点
        const afterTags = await getAllContentControlTags();
        autoGeneratedTags = afterTags.filter(t => !beforeTags.includes(t));
        console.log(`[AutoGenerate] 新增埋点: ${autoGeneratedTags.length} 个`);
        
        // Step 5: 显示撤销按钮（只要有埋点就显示，支持撤销所有或仅本次）
        const undoBtn = document.getElementById('btn-undo-embed');
        if (undoBtn && afterTags.length > 0) {
            undoBtn.style.display = 'block';
            // 更新按钮文本显示可撤销的数量
            undoBtn.title = autoGeneratedTags.length > 0 
                ? `撤销本次新增的 ${autoGeneratedTags.length} 个埋点，或撤销所有 ${afterTags.length} 个埋点`
                : `撤销所有 ${afterTags.length} 个埋点`;
        }
        
        // 完成通知
        showNotification(`生成完成！新增 ${aiResult.success || 0} 个埋点`, "success", 5000);
        console.log("[AutoGenerate] ========== 完成 ==========");
        
    } catch (error) {
        console.error("[AutoGenerate] 错误:", error);
        showNotification(`生成失败: ${error.message || error}`, "error");
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalHTML;
    }
}

/**
 * 检查并显示撤销按钮
 * 在页面加载时调用，如果文档中已有埋点则显示撤销按钮
 */
async function checkAndShowUndoButton() {
    try {
        const allTags = await getAllContentControlTags();
        const undoBtn = document.getElementById('btn-undo-embed');
        
        if (undoBtn && allTags.length > 0) {
            undoBtn.style.display = 'block';
            undoBtn.title = `撤销所有 ${allTags.length} 个埋点`;
            console.log(`[Init] 检测到 ${allTags.length} 个埋点，显示撤销按钮`);
        }
    } catch (e) {
        console.warn("[Init] 检查撤销按钮失败:", e);
    }
}

/**
 * 撤销自动埋点
 * 支持撤销本次新增的埋点或所有埋点
 */
async function undoAutoEmbed() {
    const btn = document.getElementById('btn-undo-embed');
    if (!btn) return;
    
    // 获取当前所有埋点
    let allTags = [];
    try {
        allTags = await getAllContentControlTags();
    } catch (e) {
        console.error("[Undo] 获取埋点失败:", e);
        showNotification("获取埋点失败", "error");
        return;
    }
    
    if (allTags.length === 0) {
        showNotification("文档中没有埋点", "info");
        btn.style.display = 'none';
        return;
    }
    
    // 确定撤销范围
    let tagsToDelete = [];
    let undoType = 'all'; // 'new' 或 'all'
    
    if (autoGeneratedTags.length > 0 && autoGeneratedTags.length < allTags.length) {
        // 有本次新增的埋点，且不是全部 - 询问用户
        const msg = `选择撤销范围：\n\n` +
                    `• 仅撤销本次新增: ${autoGeneratedTags.length} 个埋点\n` +
                    `• 撤销所有埋点: ${allTags.length} 个埋点\n\n` +
                    `点击「确定」撤销本次新增，点击「取消」撤销所有埋点`;
        
        const undoOnlyNew = await showConfirmDialog(msg, {
            title: "选择撤销范围",
            confirmText: "仅本次新增",
            cancelText: "撤销所有"
        });
        
        if (undoOnlyNew) {
            tagsToDelete = autoGeneratedTags;
            undoType = 'new';
        } else {
            tagsToDelete = allTags;
            undoType = 'all';
        }
    } else {
        // 没有新增记录或全部都是新增的 - 直接撤销所有
        tagsToDelete = allTags;
        undoType = 'all';
    }
    
    if (tagsToDelete.length === 0) {
        showNotification("没有可撤销的埋点", "info");
        return;
    }
    
    btn.disabled = true;
    const originalHTML = btn.innerHTML;
    btn.innerHTML = '<i class="ms-Icon ms-Icon--Sync"></i> 正在撤销...';
    
    try {
        let deletedCount = 0;
        
        // #region agent log - 假设 C: 记录要删除的 tags
        fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:beforeDelete',message:'准备删除埋点',data:{tagsToDelete,tagsCount:tagsToDelete.length,undoType,autoGeneratedTagsCount:autoGeneratedTags.length},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'C'})}).catch(()=>{});
        // #endregion
        
        await Word.run(async (context) => {
            try {
                // #region agent log - 步骤1: 开始加载
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:step1',message:'开始Word.run',data:{},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                console.log('[Undo-Step] 1. 开始 Word.run');
                // #endregion
                
                const ccs = context.document.contentControls;
                ccs.load("items");
                await context.sync();
                
                // #region agent log - 步骤2: CC加载完成
                const totalCCCount = ccs.items.length;
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:step2',message:'CC加载完成',data:{totalCCCount},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                console.log(`[Undo-Step] 2. 加载了 ${totalCCCount} 个 CC`);
                // #endregion
                
                // 加载所有 tag
                for (const cc of ccs.items) {
                    cc.load("tag");
                }
                await context.sync();
                
                // #region agent log - 步骤3: tags加载完成
                const allCCTags = ccs.items.map(cc => cc.tag);
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:step3',message:'tags加载完成',data:{allCCTags:allCCTags.slice(0,10),totalTags:allCCTags.length,matchCount:allCCTags.filter(t => tagsToDelete.includes(t)).length},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                console.log(`[Undo-Step] 3. 找到 ${allCCTags.filter(t => tagsToDelete.includes(t)).length} 个匹配的 tag`);
                // #endregion
                
                // 删除指定的 Content Control
                const toDelete = [];
                for (const cc of ccs.items) {
                    if (tagsToDelete.includes(cc.tag)) {
                        toDelete.push(cc);
                    }
                }
                
                // #region agent log - 步骤4: 准备删除
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:step4',message:'准备删除',data:{totalToDelete:toDelete.length,deleteMethod:'batch',deleteParam:false},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                console.log(`[Undo-Step] 4. 准备删除 ${toDelete.length} 个 CC (使用 delete(false))`);
                // #endregion
                
                // 批量删除（不在循环中 sync）
                for (let i = 0; i < toDelete.length; i++) {
                    const cc = toDelete[i];
                    try {
                        cc.delete(false); // 根据代码库其他地方的注释，false = 保留内容
                        deletedCount++;
                        // #region agent log - 每个删除操作
                        if (i === 0 || i === toDelete.length - 1) {
                            fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:deleteOne',message:'删除单个CC',data:{index:i,tag:cc.tag,total:toDelete.length},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                        }
                        // #endregion
                    } catch (delErr) {
                        // #region agent log - 删除失败
                        fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:deleteError',message:'删除单个CC失败',data:{index:i,tag:cc.tag,error:delErr.message},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                        console.error(`[Undo-Step] 删除失败 [${i}]: ${cc.tag}`, delErr);
                        // #endregion
                    }
                }
                
                // #region agent log - 步骤5: 准备sync
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:step5',message:'准备sync',data:{deletedCount,totalCCCount},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                console.log(`[Undo-Step] 5. 准备 sync (已标记删除 ${deletedCount} 个)`);
                // #endregion
                
                await context.sync();
                
                // #region agent log - 步骤6: sync完成
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:step6',message:'sync完成',data:{deletedCount},timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'STEP'})}).catch(()=>{});
                console.log('[Undo-Step] 6. sync 完成');
                // #endregion
            } catch (innerError) {
                // #region agent log - Word.run内部错误
                const innerErrorDetail = {
                    message: innerError.message,
                    name: innerError.name,
                    debugInfo: innerError.debugInfo ? JSON.stringify(innerError.debugInfo) : null
                };
                fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:innerError',message:'Word.run内部错误',data:innerErrorDetail,timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'ERROR'})}).catch(()=>{});
                console.error('[Undo-Step] Word.run 内部错误:', innerError);
                if (innerError.debugInfo) {
                    console.error('[Undo-Step] debugInfo:', innerError.debugInfo);
                }
                // #endregion
                throw innerError;
            }
        });
        
        // 清除 AI 表单项（如果撤销所有）
        if (undoType === 'all') {
            document.querySelectorAll('.ai-field-wrapper').forEach(el => el.remove());
            document.querySelectorAll('.ai-fields-section').forEach(el => el.remove());
            
            // 清除存储的 AI 字段
            try {
                await saveAIFieldsToDocument([]);
            } catch (e) {
                console.warn("[Undo] 清除 AI 字段存储失败:", e);
            }
        }
        
        // 更新状态
        if (undoType === 'all' || deletedCount === autoGeneratedTags.length) {
            // 如果撤销了所有埋点或本次所有新增，隐藏按钮
            btn.style.display = 'none';
            autoGeneratedTags = [];
        } else {
            // 否则更新 autoGeneratedTags
            autoGeneratedTags = autoGeneratedTags.filter(t => !tagsToDelete.includes(t));
        }
        
        const typeText = undoType === 'new' ? '本次新增的' : '所有';
        showNotification(`已撤销${typeText} ${deletedCount} 个埋点`, "success");
        console.log(`[Undo] 撤销完成: 删除${typeText} ${deletedCount} 个埋点`);
        
        btn.disabled = false;
        btn.innerHTML = originalHTML;
        
    } catch (error) {
        console.error("[Undo] 撤销失败:", error);
        
        // #region agent log - 捕获详细错误信息
        const errorDetail = {
            message: error.message,
            name: error.name,
            debugInfo: error.debugInfo ? JSON.stringify(error.debugInfo) : null,
            stack: error.stack
        };
        fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:undoAutoEmbed:error',message:'撤销失败详情',data:errorDetail,timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'ERROR'})}).catch(()=>{});
        
        // 打印详细的 debugInfo
        if (error.debugInfo) {
            console.error("[Undo] debugInfo:", error.debugInfo);
            showNotification(`撤销失败: ${error.message}\n详情: ${JSON.stringify(error.debugInfo)}`, "error");
        } else {
            showNotification(`撤销失败: ${error.message}`, "error");
        }
        // #endregion
        
        btn.disabled = false;
        btn.innerHTML = originalHTML;
    }
}

// 暴露到全局
window.autoGenerateForm = autoGenerateForm;
window.undoAutoEmbed = undoAutoEmbed;

// ---------------- 插入 Content Control ----------------
async function insertControl(tag, title, isWrapper = false, specificRoundId = null) {
    console.log(`[InsertControl] 开始插入: tag=${tag}, title=${title}, isWrapper=${isWrapper}, roundId=${specificRoundId}`);
    
    return wordActionQueue.add(async () => {
    if (typeof Word === 'undefined') {
        console.warn("[InsertControl] Word API 不可用");
        return;
    }
    try {
        await Word.run(async (context) => {
            console.log(`[InsertControl] Word.run 开始执行...`);
            
            // 获取选区
            const selection = context.document.getSelection();
            selection.load("text, parentTableCellOrNullObject, parentContentControlOrNullObject");
            await context.sync();
            
            console.log(`[InsertControl] 当前选中文本: "${selection.text ? selection.text.substring(0, 50) : '(空)'}..."`);
            
            if (!selection.text || selection.text.trim() === '') {
                console.warn(`[InsertControl] ⚠️ 警告: 未选中任何文本！请先选中要埋点的段落。`);
                showNotification(`请先在 Word 中选中要埋点的"${title}"段落，然后再点击插入按钮。`, "warning");
                return;
            }
            
            // 【嵌套检测】允许嵌套，但给出提示
            const parentCC = selection.parentContentControlOrNullObject;
            parentCC.load("tag, title, isNullObject");
            await context.sync();
            
            if (!parentCC.isNullObject) {
                const parentName = parentCC.title || parentCC.tag;
                if (isWrapper) {
                    // "插入段落"嵌套在另一个 Content Control 内部：允许，但警告
                    console.warn(`[InsertControl] ⚠️ 嵌套"插入段落"！在 "${parentName}" 内部插入 "${title}"`);
                    showNotification(`ℹ️ 嵌套提示\n\n在"${parentName}"内部插入了"${title}"段落。\n\n注意：\n• 隐藏外层时，内层也会一起隐藏\n• 恢复外层后，内层保持隐藏前的状态`, "info", 6000);
                } else {
                    // 普通埋点在 Content Control 内部：正常允许
                    console.log(`[InsertControl] ℹ️ 在 "${parentName}" 内部插入普通埋点 "${tag}"`);
                }
            }
            
            let targetRange = selection;
            
            // 对于表格单元格内的选择，直接使用用户选择的内容，不扩展到整个单元格
            // 这样可以正常在表格单元格内创建 Content Control 埋点
            
            // 插入 Content Control (RichText 类型)
            console.log(`[InsertControl] 正在插入 Content Control...`);
            const contentControl = targetRange.insertContentControl("RichText");
            contentControl.tag = tag;
            contentControl.title = title;
            contentControl.appearance = "Tags"; // Tags 模式显示 title 标签
            contentControl.color = "blue";
            contentControl.cannotEdit = false;  // 确保可编辑
            contentControl.cannotDelete = false; // 确保可删除
                
            // 插入时显示占位符 [字段名]，方便识别
            // 段落模式 (isWrapper=true) 保留原内容
            if (!isWrapper) {
                contentControl.insertText(`[${title}]`, "Replace");
            }
            
            // 同步写入
            await context.sync();
            
            // 额外等待一下，让 Word Online 处理完毕
            await new Promise(r => setTimeout(r, 500));
            
            console.log(`✅ [InsertControl] 成功插入 Content Control: ${tag}`);
        });
    } catch (error) {
        console.error(`❌ [InsertControl] 插入失败 (${tag}):`, error.message || error);
        
        // 获取详细的调试信息
        if (error.debugInfo) {
            console.error(`❌ [InsertControl] debugInfo:`, JSON.stringify(error.debugInfo, null, 2));
        }
        if (error.code) {
            console.error(`❌ [InsertControl] errorCode: ${error.code}`);
        }
        if (error.traceMessages) {
            console.error(`❌ [InsertControl] traceMessages:`, error.traceMessages);
        }
        
        showNotification(`插入"${title}"失败: ${error.message || error}。请确保选中的是普通文本（不在已有埋点或表格内）`, "error");
    }
    });
}

// =====================================================================
// 股东字段显示/隐藏控制
// =====================================================================
// 对于表格中的可选股东（SH2/SH3/SH4）：
// - 复选框控制表单区域 + 文档中对应埋点的内容
// - 取消勾选时：将所有字段埋点内容设为占位符
// - 勾选时：如果内容是占位符则清空，等待用户填写
// =====================================================================

const SHAREHOLDER_HIDDEN_PLACEHOLDER = "[▶已隐藏]";

// 获取股东字段的标签和中文名称映射
function getShareholderFieldsInfo(tagPrefix) {
    // 从 contractConfig 中获取股东字段定义
    const shSection = contractConfig.find(s => s.type === "existing_shareholders");
    if (!shSection || !shSection.shareholderFields) return [];
    
    return shSection.shareholderFields.map(f => ({
        tag: tagPrefix + f.tag,
        label: f.label  // 中文名称，如"姓名/名称"、"认缴注册资本(万元)"
    }));
}

/**
 * 切换股东字段的可见性（设置/清除占位符）
 * @param {string} tagPrefix - 股东标签前缀（如 "SH2"）
 * @param {boolean} isVisible - true=显示（显示中文名称），false=隐藏（设置占位符）
 */
async function toggleShareholderFieldsVisibility(tagPrefix, isVisible) {
    const fieldsInfo = getShareholderFieldsInfo(tagPrefix);
    if (fieldsInfo.length === 0) {
        console.warn(`[ShareholderFields] 未找到 ${tagPrefix} 的字段定义`);
        return;
    }
    
    console.log(`[ShareholderFields] ${isVisible ? '显示' : '隐藏'} ${tagPrefix} 的 ${fieldsInfo.length} 个字段`);
    
    try {
        await Word.run(async (context) => {
            let processedCount = 0;
            
            for (const fieldInfo of fieldsInfo) {
                // 查找该标签的 Content Control
                const controls = context.document.contentControls.getByTag(fieldInfo.tag);
                controls.load("items");
                await context.sync();
                
                if (controls.items.length === 0) {
                    // 该字段尚未埋点，跳过
                    continue;
                }
                
                for (const ctrl of controls.items) {
                    ctrl.load("text");
                    await context.sync();
                    
                    if (isVisible) {
                        // 显示：如果内容是占位符，设置为中文字段名称
                        if (ctrl.text === SHAREHOLDER_HIDDEN_PLACEHOLDER) {
                            await insertTextPreserveFormat(ctrl, `[${fieldInfo.label}]`, context);
                            processedCount++;
                        }
                    } else {
                        // 隐藏：设置为占位符
                        if (ctrl.text !== SHAREHOLDER_HIDDEN_PLACEHOLDER) {
                            await insertTextPreserveFormat(ctrl, SHAREHOLDER_HIDDEN_PLACEHOLDER, context);
                            processedCount++;
                        }
                    }
                }
                
                await context.sync();
            }
            
            console.log(`[ShareholderFields] ✅ ${tagPrefix} 处理完成，${processedCount} 个字段已${isVisible ? '显示' : '隐藏'}`);
        });
    } catch (error) {
        console.error(`[ShareholderFields] ❌ 处理失败:`, error.message || error);
    }
}

// =====================================================================
// 切换轮次段落可见性 (彻底删除 + 恢复方案) —— 分步稳健版
// =====================================================================
// 核心原则：拆解操作步骤，降低单次事务负载，使用"空格"保活控件。
// 1. 获取 OOXML
// 2. 存储 Settings (独立 Sync)
// 3. 删除内容 (独立 Sync，用空格占位)

const BACKUP_PREFIX = "Bk_"; // 缩短前缀，避免 Key 长度限制问题
const CHUNK_SIZE = 100 * 1024; // 100KB per chunk - 极度保守以避免 Word Online 同步崩溃

/**
 * 精简 OOXML：移除不需要的 media（图片等）资源，大幅减小大小
 * @param {string} ooxml - 原始 OOXML 字符串
 * @returns {string} - 精简后的 OOXML
 */
/**
 * 【诊断函数】分析 OOXML 包结构
 * 列出所有 pkg:part，检查内部引用
 * @param {string} ooxml - OOXML 字符串
 * @param {string} label - 标签（用于日志）
 */
function diagnoseOoxml(ooxml, label = "OOXML") {
    console.log(`\n========== [${label}] 诊断开始 ==========`);
    console.log(`[${label}] 总大小: ${(ooxml.length / 1024).toFixed(1)} KB`);
    
    // 1. 提取所有 pkg:part 的名称
    const partNameRegex = /<pkg:part[^>]*pkg:name="([^"]+)"/g;
    const parts = [];
    let match;
    while ((match = partNameRegex.exec(ooxml)) !== null) {
        parts.push(match[1]);
    }
    
    console.log(`[${label}] 包含 ${parts.length} 个 pkg:part:`);
    parts.forEach((p, i) => {
        // 估算每个 part 的大小
        const partRegex = new RegExp(`<pkg:part[^>]*pkg:name="${p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>[\\s\\S]*?<\\/pkg:part>`, 'i');
        const partMatch = ooxml.match(partRegex);
        const size = partMatch ? partMatch[0].length : 0;
        console.log(`   ${i + 1}. ${p} (${(size / 1024).toFixed(1)} KB)`);
    });
    
    // 2. 检查关键文件是否存在
    const criticalParts = [
        '/word/document.xml',
        '/word/styles.xml',
        '/word/_rels/document.xml.rels',
        '/_rels/.rels'
    ];
    console.log(`[${label}] 关键文件检查:`);
    criticalParts.forEach(cp => {
        const exists = parts.some(p => p === cp);
        console.log(`   ${exists ? '✓' : '✗'} ${cp}`);
    });
    
    // 3. 检查 document.xml 中的 rId 引用
    const docPartMatch = ooxml.match(/<pkg:part[^>]*pkg:name="\/word\/document\.xml"[^>]*>([\s\S]*?)<\/pkg:part>/i);
    if (docPartMatch) {
        const docContent = docPartMatch[1];
        const rIdRefs = docContent.match(/r:id="([^"]+)"/g) || [];
        const rIdEmbed = docContent.match(/r:embed="([^"]+)"/g) || [];
        console.log(`[${label}] document.xml 中的外部引用:`);
        console.log(`   - r:id 引用: ${rIdRefs.length} 个`);
        console.log(`   - r:embed 引用: ${rIdEmbed.length} 个`);
        
        if (rIdRefs.length > 0) {
            const uniqueRefs = [...new Set(rIdRefs.slice(0, 10))];
            console.log(`   - 前 10 个 r:id: ${uniqueRefs.join(', ')}`);
        }
    } else {
        console.log(`[${label}] ⚠️ 未找到 /word/document.xml 部分!`);
    }
    
    // 4. 检查 /_rels/.rels 中的关系
    const relsMatch = ooxml.match(/<pkg:part[^>]*pkg:name="\/_rels\/\.rels"[^>]*>([\s\S]*?)<\/pkg:part>/i);
    if (relsMatch) {
        const relsContent = relsMatch[1];
        const relationships = relsContent.match(/<Relationship[^>]+>/g) || [];
        console.log(`[${label}] /_rels/.rels 中的关系 (${relationships.length} 个):`);
        relationships.forEach((rel, i) => {
            const target = rel.match(/Target="([^"]+)"/);
            const type = rel.match(/Type="[^"]*\/([^"\/]+)"/);
            console.log(`   ${i + 1}. ${type ? type[1] : '?'} → ${target ? target[1] : '?'}`);
        });
    }
    
    // 5. 检查 /word/_rels/document.xml.rels
    const docRelsMatch = ooxml.match(/<pkg:part[^>]*pkg:name="\/word\/_rels\/document\.xml\.rels"[^>]*>([\s\S]*?)<\/pkg:part>/i);
    if (docRelsMatch) {
        const docRelsContent = docRelsMatch[1];
        const docRelationships = docRelsContent.match(/<Relationship[^>]+>/g) || [];
        console.log(`[${label}] /word/_rels/document.xml.rels 中的关系 (${docRelationships.length} 个):`);
        docRelationships.slice(0, 10).forEach((rel, i) => {
            const id = rel.match(/Id="([^"]+)"/);
            const target = rel.match(/Target="([^"]+)"/);
            const type = rel.match(/Type="[^"]*\/([^"\/]+)"/);
            console.log(`   ${i + 1}. ${id ? id[1] : '?'}: ${type ? type[1] : '?'} → ${target ? target[1] : '?'}`);
        });
        if (docRelationships.length > 10) {
            console.log(`   ... 还有 ${docRelationships.length - 10} 个关系`);
        }
    } else {
        console.log(`[${label}] ⚠️ 未找到 /word/_rels/document.xml.rels 部分!`);
    }
    
    console.log(`========== [${label}] 诊断结束 ==========\n`);
}

/**
 * 清理 OOXML 中的 webextensions 引用和数据
 * 保留 document.xml 等核心内容，移除加载项相关的大数据
 * 【重要】同时移除文件和对应的引用，避免引用断裂
 * @param {string} ooxml - 完整的 OOXML 包
 * @returns {string} - 清理后的 OOXML
 */
function cleanOoxmlForRestore(ooxml) {
    if (!ooxml) return ooxml;
    
    let cleaned = ooxml;
    
    // ========== 第一步：移除 pkg:part 文件 ==========
    
    // 1. 移除所有 webextensions 的 pkg:part
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*webextensions[^"]*"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    
    // 2. 移除所有 taskpanes 的 pkg:part
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*taskpanes[^"]*"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    
    // 3. 移除 media（图片等）
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*\/media\/[^"]*"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    
    // 4. 移除 theme
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*\/theme\/[^"]*"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    
    // 5. 移除 footnotes 和 endnotes（用户确认不使用）
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*\/footnotes\.xml"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*\/endnotes\.xml"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    
    // 6. 【实验性】移除 settings.xml（占用 270+ KB，可能包含不兼容设置）
    cleaned = cleaned.replace(/<pkg:part[^>]*pkg:name="[^"]*\/settings\.xml"[^>]*>[\s\S]*?<\/pkg:part>/gi, '');
    
    // ========== 第二步：移除对已删除文件的引用（避免引用断裂）==========
    
    // 6. 移除 /_rels/.rels 中对 webextensions/taskpanes 的引用
    cleaned = cleaned.replace(/<Relationship[^>]*webextension[^>]*\/>/gi, '');
    cleaned = cleaned.replace(/<Relationship[^>]*taskpanes[^>]*\/>/gi, '');
    
    // 7. 【关键修复】移除 document.xml.rels 中对 theme 的引用
    cleaned = cleaned.replace(/<Relationship[^>]*Type="[^"]*\/theme"[^>]*\/>/gi, '');
    
    // 8. 移除 document.xml.rels 中对 media/image 的引用
    cleaned = cleaned.replace(/<Relationship[^>]*Target="media\/[^"]*"[^>]*\/>/gi, '');
    
    // 9. 移除 document.xml.rels 中对 footnotes/endnotes 的引用
    cleaned = cleaned.replace(/<Relationship[^>]*Type="[^"]*\/footnotes"[^>]*\/>/gi, '');
    cleaned = cleaned.replace(/<Relationship[^>]*Type="[^"]*\/endnotes"[^>]*\/>/gi, '');
    
    // 10. 【实验性】移除 document.xml.rels 中对 settings 的引用
    cleaned = cleaned.replace(/<Relationship[^>]*Type="[^"]*\/settings"[^>]*\/>/gi, '');
    
    // 11. 清理可能残留的空 Relationship（Target 为空或无效）
    cleaned = cleaned.replace(/<Relationship[^>]*Target=""[^>]*\/>/gi, '');
    
    // 12. 清理 ? → ? 的无效 Relationship（没有 Id 或 Target 的）
    // 移除只有换行符或空白的 Relationship
    cleaned = cleaned.replace(/<Relationship\s*\/>/gi, '');
    
    return cleaned;
}

/**
 * 精简 OOXML 用于保存
 * 不做任何修改，保存完整 OOXML（但会过滤掉一些不需要的大数据）
 * @param {string} ooxml - 完整的 OOXML 包
 * @returns {string} - 精简后的 OOXML
 */
function slimOoxml(ooxml) {
    if (!ooxml) return ooxml;
    
    const originalSize = ooxml.length;
    
    // 直接使用 cleanOoxmlForRestore 进行精简
    // 保存和恢复使用相同的清理逻辑
    const slimmed = cleanOoxmlForRestore(ooxml);
    
    const slimmedSize = slimmed.length;
    const savedPercent = ((1 - slimmedSize / originalSize) * 100).toFixed(1);
    
    console.log(`[SlimOoxml] ${(originalSize / 1024).toFixed(1)} KB → ${(slimmedSize / 1024).toFixed(1)} KB (节省 ${savedPercent}%)`);
    
    return slimmed;
}

/**
 * 【测试函数】用最小有效 OOXML 测试 insertOoxml 是否工作
 * 这是一个已知有效的最小 OOXML，只包含一段简单文字
 */
function getTestMinimalOoxml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
<pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
</pkg:xmlData>
</pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
<pkg:xmlData>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
<w:r>
<w:t>测试内容 - insertOoxml 有效性测试</w:t>
</w:r>
</w:p>
</w:body>
</w:document>
</pkg:xmlData>
</pkg:part>
</pkg:package>`;
}

/**
 * 构建用于恢复的 OOXML（清理无效引用）
 * @param {string} savedOoxml - 保存的 OOXML
 * @returns {string} - 清理后可用于恢复的 OOXML
 */
function buildMinimalOoxml(savedOoxml) {
    // 如果不是完整的 OOXML 包，尝试包装
    if (!savedOoxml.includes('<pkg:package')) {
        if (savedOoxml.includes('<w:document') || savedOoxml.includes('<w:body')) {
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
<pkg:xmlData>${savedOoxml}</pkg:xmlData>
</pkg:part>
</pkg:package>`;
        }
    }
    
    // 对已经是完整包的 OOXML，再次清理（确保无效引用被移除）
    return cleanOoxmlForRestore(savedOoxml);
}

/**
 * 分块存储大数据到 document.settings
 * @param {Word.RequestContext} context 
 * @param {Word.Settings} settings 
 * @param {string} key - 基础 key 名称
 * @param {string} data - 要存储的数据
 * @returns {Promise<number>} - 存储的块数
 */
async function saveToSettingsChunked(context, settings, key, data) {
    const chunks = [];
    for (let i = 0; i < data.length; i += CHUNK_SIZE) {
        chunks.push(data.slice(i, i + CHUNK_SIZE));
    }
    
    const chunkCount = chunks.length;
    console.log(`[Chunked] Saving ${key}: ${data.length} bytes -> ${chunkCount} chunks`);
    
    // 存储元数据
    settings.add(`${key}_Meta`, JSON.stringify({ count: chunkCount, length: data.length }));
    
    // 分批写入，每 5 块 sync 一次
    for (let i = 0; i < chunks.length; i++) {
        settings.add(`${key}_${i}`, chunks[i]);
        
        // 每 5 块同步一次，或者最后一块
        if ((i + 1) % 5 === 0 || i === chunks.length - 1) {
            await context.sync();
            console.log(`[Chunked] Synced chunks ${Math.max(0, i - 4)}-${i} for ${key}`);
        }
    }
    
    return chunkCount;
}

/**
 * 从 document.settings 读取分块数据
 * @param {Word.RequestContext} context 
 * @param {Word.Settings} settings 
 * @param {string} key - 基础 key 名称
 * @returns {Promise<string|null>} - 拼接后的完整数据，或 null 如果不存在
 */
async function readFromSettingsChunked(context, settings, key) {
    // 先读取元数据
    const metaSetting = settings.getItemOrNullObject(`${key}_Meta`);
    metaSetting.load("value,isNullObject");
    await context.sync();
    
    if (metaSetting.isNullObject || !metaSetting.value) {
        // 检查是否为旧版非分块数据
        const legacySetting = settings.getItemOrNullObject(key);
        legacySetting.load("value,isNullObject");
        await context.sync();
        
        if (!legacySetting.isNullObject && legacySetting.value) {
            console.log(`[Chunked] Legacy (non-chunked) data found for ${key}`);
            return legacySetting.value;
        }
        
        console.log(`[Chunked] No data found for ${key}`);
        return null;
    }
    
    const meta = JSON.parse(metaSetting.value);
    const chunkCount = meta.count;
    console.log(`[Chunked] Reading ${key}: ${chunkCount} chunks, expected ${meta.length} bytes`);
    
    // 加载所有块
    const chunkSettings = [];
    for (let i = 0; i < chunkCount; i++) {
        const chunkSetting = settings.getItemOrNullObject(`${key}_${i}`);
        chunkSetting.load("value,isNullObject");
        chunkSettings.push(chunkSetting);
    }
    await context.sync();
    
    // 拼接
    let result = "";
    for (let i = 0; i < chunkCount; i++) {
        if (chunkSettings[i].isNullObject || !chunkSettings[i].value) {
            console.error(`🔴 [Chunked] Missing chunk ${i} for ${key}!`);
            return null; // 数据不完整
        }
        result += chunkSettings[i].value;
    }
    
    if (result.length !== meta.length) {
        console.warn(`⚠️ [Chunked] Length mismatch for ${key}: got ${result.length}, expected ${meta.length}`);
    }
    
    console.log(`✅ [Chunked] Successfully read ${key}: ${result.length} bytes`);
    return result;
}

/**
 * 删除分块存储的数据
 * @param {Word.RequestContext} context 
 * @param {Word.Settings} settings 
 * @param {string} key - 基础 key 名称
 */
async function deleteFromSettingsChunked(context, settings, key) {
    // 读取元数据获取块数
    const metaSetting = settings.getItemOrNullObject(`${key}_Meta`);
    metaSetting.load("value,isNullObject");
    await context.sync();
    
    if (metaSetting.isNullObject) {
        // 尝试删除旧版非分块数据
        const legacySetting = settings.getItemOrNullObject(key);
        legacySetting.load("isNullObject");
        await context.sync();
        if (!legacySetting.isNullObject) {
            legacySetting.delete();
            await context.sync();
            console.log(`[Chunked] Deleted legacy setting for ${key}`);
        }
        return;
    }
    
    const meta = JSON.parse(metaSetting.value);
    const chunkCount = meta.count;
    
    // 删除所有块
    for (let i = 0; i < chunkCount; i++) {
        const chunkSetting = settings.getItemOrNullObject(`${key}_${i}`);
        chunkSetting.load("isNullObject");
    }
    await context.sync();
    
    for (let i = 0; i < chunkCount; i++) {
        const chunkSetting = settings.getItemOrNullObject(`${key}_${i}`);
        chunkSetting.delete();
    }
    
    // 删除元数据
    metaSetting.delete();
    await context.sync();
    
    console.log(`[Chunked] Deleted ${chunkCount} chunks for ${key}`);
}

/**
 * 自动控制法定代表人段落的显示/隐藏
 * 当投资人类型切换时自动调用
 * @param {string} paraTag - 法定代表人段落的 Tag (如 Inv_Lead_LegalRep_Para)
 * @param {boolean} shouldShow - 是否应该显示
 */
async function autoToggleLegalRepParagraph(paraTag, shouldShow) {
    try {
        let shouldProceed = false;
        
        await Word.run(async (context) => {
            const controls = context.document.contentControls;
            controls.load("items/tag,items/text");
            await context.sync();
            
            // 查找是否存在法定代表人段落的内容控件
            const target = controls.items.find(c => c.tag === paraTag);
            if (!target) {
                console.log(`[AutoToggle] ${paraTag} 未埋点，跳过`);
                return;
            }
            
            // 检查当前状态（占位符格式为 [▶xxx]）
            const currentlyHidden = /^\[▶.+\]$/.test((target.text || "").trim());
            
            // 判断是否需要操作
            if (shouldShow && !currentlyHidden) {
                console.log(`[AutoToggle] ${paraTag} 已显示，无需操作`);
                return;
            }
            if (!shouldShow && currentlyHidden) {
                console.log(`[AutoToggle] ${paraTag} 已隐藏，无需操作`);
                return;
            }
            
            shouldProceed = true;
        });
        
        // 只有当控件存在且状态需要改变时才调用 toggleRoundVisibility
        if (shouldProceed) {
            console.log(`[AutoToggle] 执行 ${shouldShow ? '恢复' : '隐藏'}: ${paraTag}`);
            await toggleRoundVisibility(paraTag, shouldShow);
        }
    } catch (error) {
        console.warn(`[AutoToggle] 失败:`, error.message || error);
    }
}

async function toggleRoundVisibility(tag, isVisible) {
    // 强制排队执行
    return wordActionQueue.add(async () => {
        if (typeof Word === 'undefined') {
            console.log(`[Toggle] Mock: Tag=${tag}, Visible=${isVisible}`);
            return;
        }
        
        let retryCount = 0;
        const maxRetries = 3;

        while (retryCount < maxRetries) {
            try {
                await Word.run(async (context) => {
                    const settings = context.document.settings;
                    const settingKey = `${BACKUP_PREFIX}${tag}`;
                    
                    // ========== 阶段 1：获取目标控件 ==========
                    const targets = context.document.contentControls.getByTag(tag);
                    targets.load("items,text,title");
                    await context.sync(); // 【同步 1：只读】
                    
                    if (targets.items.length === 0) {
                        console.warn(`[Toggle] Tag ${tag} not found.`);
                        return;
                    }
                    const ctrl = targets.items[0];
                    // 判断是否有实际内容（排除占位符 [▶xxx] 格式）
                    const isHiddenPlaceholder = /^\[▶.+\]$/.test((ctrl.text || "").trim());
                    const hasContent = ctrl.text && ctrl.text.trim().length > 0 && !isHiddenPlaceholder;

                    if (isVisible) {
                        // ========== 恢复逻辑 (使用最小 OOXML 包重建) ==========
                        
                        // 【新增】如果当前有实际内容（非隐藏占位符），跳过恢复
                        if (hasContent) {
                            console.log(`[Toggle] ${tag} 当前有内容，无需恢复`);
                            return;
                        }
                        
                        console.log(`[Toggle] Attempting to restore ${tag}...`);
                        const savedContent = await readFromSettingsChunked(context, settings, settingKey);
                        
                        if (savedContent) {
                            console.log(`[Toggle] Restoring content for ${tag} (saved: ${savedContent.length} bytes)...`);
                            
                            // 使用 buildMinimalOoxml 重建最小有效 OOXML 包
                            const finalOoxml = buildMinimalOoxml(savedContent);
                            
                            // 预先插入一个空格，激活控件编辑状态（防止 ItemNotFound）
                            ctrl.insertText(" ", "Replace");
                            await context.sync();
                            
                            // 恢复 OOXML
                            ctrl.insertOoxml(finalOoxml, "Replace");
                            await context.sync();
                            console.log(`✅ [Toggle] Restored ${tag} with OOXML successfully`);
                            
                            // 删除备份
                            await deleteFromSettingsChunked(context, settings, settingKey);
                        } else {
                            console.log(`[Toggle] No backup found for ${tag}, skipping restore.`);
                        }
                    } else {
                        // ========== 隐藏逻辑 (纯 OOXML，带大小拦截) ==========
                        if (!hasContent) {
                            console.log(`[Toggle] ${tag} already empty/hidden.`);
                            return;
                        }

                        console.log(`[Toggle] [Step 1] Getting OOXML for ${tag}...`);
                        const ooxmlResult = ctrl.getOoxml();
                        await context.sync(); // 【同步 2：获取数据】
                        
                        const originalOoxml = ooxmlResult.value || "";
                        const originalLength = originalOoxml.length;
                        console.log(`[Toggle] 原始 OOXML 大小 for ${tag}: ${(originalLength / 1024).toFixed(1)} KB`);
                        
                        // 【OOXML 精简】移除 media、theme 等大型资源
                        const slimmedOoxml = slimOoxml(originalOoxml);
                        const slimmedLength = slimmedOoxml.length;
                        
                        // 【嵌套控件保护】检查精简后的 OOXML 大小，发出警告
                        const slimmedSizeKB = (slimmedLength / 1024).toFixed(1);
                        if (slimmedLength > 500 * 1024) { // 超过 500KB 警告
                            console.warn(`⚠️ [Toggle] ${tag} OOXML 较大 (${slimmedSizeKB} KB)，可能是嵌套控件或复杂内容`);
                        }
                        
                        // 【方案 G 改进】精简后仍超过 3.5MB 才拦截
                        const MAX_OOXML_SIZE = 3.5 * 1024 * 1024; // 3.5MB
                        if (slimmedLength > MAX_OOXML_SIZE) {
                            const sizeMB = (slimmedLength / (1024 * 1024)).toFixed(2);
                            console.error(`❌ [Toggle] 精简后 OOXML 仍过大 (${sizeMB} MB) for ${tag}. Hide operation aborted.`);
                            throw new Error(`段落 "${tag}" 精简后仍过大 (${sizeMB} MB)，禁止隐藏操作。`);
                        }

                        console.log(`[Toggle] [Step 2] Saving slimmed OOXML (chunked) with key: ${settingKey}...`);
                        const chunkCount = await saveToSettingsChunked(context, settings, settingKey, slimmedOoxml);
                        console.log(`✅ [Toggle] Saved OOXML for ${tag} in ${chunkCount} chunks`);

                        console.log(`[Toggle] [Step 3] Clearing content...`);
                        // 【关键】使用可见占位符显示具体名称，让用户知道这里隐藏的是什么
                        // 【保留格式】
                        const displayName = ctrl.title || tag;
                        await insertTextPreserveFormat(ctrl, `[▶${displayName}]`, context);
                        // 【同步：删正文】
                    }
                    
                    console.log(`[Toggle] ${tag} operation completed.`);
                });
                return; // 成功退出
            } catch (error) {
                console.error(`[Toggle] Attempt ${retryCount + 1} failed:`, error);
                
                // 专门检测 forceSaveFailed
                if (error.code === "forceSaveFailed" || (error.message && error.message.includes("forceSaveFailed"))) {
                    console.warn("⚠️ Word Online Save Failed. Document might be in read-only state.");
                }

                retryCount++;
                if (retryCount < maxRetries) {
                    const delay = retryCount * 2500; // 稍微增加重试间隔 2.5s, 5s
                    console.warn(`[Toggle] Retrying in ${delay}ms...`);
                    await new Promise(r => setTimeout(r, delay));
                    continue;
                }
                throw error;
            }
        }
    });
}

/**
 * 从本地 Settings 中提取所有备份数据 (用于同步)
 * 支持分块存储：识别 _Meta 后缀，使用 readFromSettingsChunked 读取
 * 返回格式: { "Tag1": "OOXML1", "Tag2": "OOXML2" }
 */
async function getBackupsFromSettings() {
    if (typeof Word === 'undefined') return {};
    return await Word.run(async (context) => {
        const settings = context.document.settings;
        settings.load("items");
        await context.sync();
        
        const backups = {};
        const processedTags = new Set();
        
        // 先找出所有有 _Meta 后缀的 key（分块存储）
        for (const item of settings.items) {
            if (item.key.startsWith(BACKUP_PREFIX) && item.key.endsWith("_Meta")) {
                // 分块存储的 key 格式: Bk_Round_Seed_Meta
                const baseKey = item.key.replace("_Meta", ""); // Bk_Round_Seed
                const tag = baseKey.replace(BACKUP_PREFIX, ""); // Round_Seed
                
                if (!processedTags.has(tag)) {
                    processedTags.add(tag);
                    const fullData = await readFromSettingsChunked(context, settings, baseKey);
                    if (fullData) {
                        backups[tag] = fullData;
                    }
                }
            }
        }
        
        // 再处理旧版非分块数据（兼容性）
        for (const item of settings.items) {
            if (item.key.startsWith(BACKUP_PREFIX) && 
                !item.key.endsWith("_Meta") && 
                !item.key.match(/_\d+$/)) { // 排除 _0, _1 等块数据
                
                const tag = item.key.replace(BACKUP_PREFIX, "");
                if (!processedTags.has(tag)) {
                    backups[tag] = item.value;
                }
            }
        }
        
        return backups;
    });
}

// ---------------- 更新文档内容 (简单字段) ----------------
// label: 可选的中文标签，用于空值时显示占位符
async function updateContent(tag, value, label = null) {
    // 强制排队
    return wordActionQueue.add(async () => {
    if (typeof Word === 'undefined') {
        console.log(`[Mock] Update Word content: Tag=${tag}, Value=${value}`);
        return;
    }
    try {
        await Word.run(async (context) => {
            let contentControls = context.document.contentControls.getByTag(tag);
            contentControls.load("items");
            await context.sync();

            // 空值时使用中文标签，如 [姓名/名称]；否则使用 tag，如 [SH2_Name]
            const displayLabel = label || tag;
            const textToInsert = value === "" ? `[${displayLabel}]` : value;

            if (contentControls.items.length > 0) {
                for (const ctrl of contentControls.items) {
                    // 【保留格式】使用通用函数更新内容
                    await insertTextPreserveFormat(ctrl, textToInsert, context);
                }
            }
        });
    } catch (error) {
        console.warn(`[UpdateContent] 更新 ${tag} 时出错:`, error.message || error);
        // 忽略简单的更新错误，队列会自动处理下一个
    }
    });
}

// ---------------- 应用表单数据到当前文档 (离线同步) ----------------
/**
 * 将 LocalStorage 中保存的表单数据批量写入当前打开文档的所有匹配 Content Control
 * 用途：打开新文档模板后，一键将之前填写的数据同步到当前文档
 */
async function applyFormToCurrentDocument() {
    console.log("[ApplyForm] 开始将表单数据应用到当前文档...");
    
    // 读取 LocalStorage 中的表单数据
    const lsState = loadFormStateFromLocalStorage();
    if (!lsState || !lsState.formData || Object.keys(lsState.formData).length === 0) {
        showNotification("没有找到已保存的表单数据。请先填写表单。", "warning");
        console.warn("[ApplyForm] LocalStorage 中没有表单数据");
        return;
    }
    
    const formData = lsState.formData;
    const totalFields = Object.keys(formData).length;
    let updatedCount = 0;
    let skippedCount = 0;
    
    showNotification(`正在同步 ${totalFields} 个字段到当前文档...`, "info");
    
    try {
        await Word.run(async (context) => {
            // 获取文档中所有的 Content Control
            const allControls = context.document.contentControls;
            allControls.load("items/tag,items/title");
            await context.sync();
            
            console.log(`[ApplyForm] 文档中共有 ${allControls.items.length} 个 Content Control`);
            
            // 创建 tag -> control 映射，方便快速查找
            const tagMap = new Map();
            for (const ctrl of allControls.items) {
                if (ctrl.tag) {
                    tagMap.set(ctrl.tag, ctrl);
                }
            }
            
            // 遍历表单数据，更新匹配的 Content Control
            for (const [tag, value] of Object.entries(formData)) {
                if (value === undefined || value === null) {
                    skippedCount++;
                    continue;
                }
                
                const ctrl = tagMap.get(tag);
                if (ctrl) {
                    // 找到匹配的 Content Control，更新内容
                    const textToInsert = value === "" ? `[${ctrl.title || tag}]` : String(value);
                    await insertTextPreserveFormat(ctrl, textToInsert, context);
                    updatedCount++;
                    console.log(`[ApplyForm] ✓ 更新 ${tag}: "${textToInsert.substring(0, 30)}..."`);
                } else {
                    // 文档中没有这个 tag 的 Content Control
                    skippedCount++;
                }
            }
            
            await context.sync();
        });
        
        showNotification(`✅ 同步完成！已更新 ${updatedCount} 个字段，跳过 ${skippedCount} 个（文档中无对应埋点）`, "success", 5000);
        console.log(`[ApplyForm] 完成: 更新 ${updatedCount}, 跳过 ${skippedCount}`);
        
    } catch (error) {
        console.error("[ApplyForm] 同步失败:", error);
        showNotification(`同步失败: ${error.message || error}`, "error");
    }
}

// 暴露到全局
window.applyFormToCurrentDocument = applyFormToCurrentDocument;

// ---------------- AI 按钮占位 ----------------
async function handleAIFill() {
    const btn = document.getElementById("btn-ai-fill");
    if (btn) {
        btn.disabled = true;
        btn.textContent = "🧠 正在分析...";
        setTimeout(() => {
            btn.disabled = false;
            btn.textContent = "✨ 开始智能识别并填充";
        }, 1500);
    }
}
window.handleAIFill = handleAIFill;

// ---------------- 文档生成函数 (Docxtemplater) ----------------
async function generateDocx(blob, data) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (evt) {
            if (evt.target.readyState !== 2) return;
            const content = evt.target.result;
            try {
                const zip = new PizZip(content);
                const doc = new window.docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                    delimiters: { start: "【", end: "】" },
                    nullGetter: () => ""
                });
                doc.render(data);
                const out = zip.generate({
                    type: "blob",
                    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                });
                resolve(out);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsBinaryString(blob);
    });
}
window.generateDocx = generateDocx;

// ---------------- 按 Tag 替换内容控件 (用于批量同步) ----------------
// 说明：直接使用 XML 替换 Content Control 内容。
// 【重构】移除状态注入逻辑，仅做纯内容替换，避免 500 错误。
// 状态同步通过 LocalStorage 实现（双轨同步方案）。
async function processDocxContentControls(blob, formData, uiState = null) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const zip = new PizZip(evt.target.result);
                
                // 【简化】直接使用传入的 uiState (来自 LocalStorage)，不再从 ZIP 读取
                const activeState = uiState || { enabledRounds: enabledRounds, formData: formData };
                
                // 用于记录该文件内部的备份（仅在内存中处理，不持久化到 ZIP）
                let fileBackups = {};
                
                // 处理正文 XML 替换与内容捕获/恢复
                    const targets = zip.file(/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/) || [];
                if (targets.length === 0 && zip.file("word/document.xml")) {
                    targets.push(zip.file("word/document.xml"));
                }
                    
                    let totalChanges = false;
                    targets.forEach((f) => {
                        const xml = f.asText();
                    // 使用 activeState 里的勾选状态
                    const { xml: newXml, hasChanges, updatedBackups } = replaceContentControlsXmlIndependent(
                        xml, 
                        formData, 
                        fileBackups, 
                        activeState.enabledRounds
                    );
                        if (hasChanges) {
                            zip.file(f.name, newXml);
                        fileBackups = updatedBackups;
                            totalChanges = true;
                        }
                    });
                    
                // 【重要】不再调用 saveStateToZip，避免修改 ZIP 结构导致 500 错误
                // 状态同步完全依赖 LocalStorage
                
                if (!totalChanges) {
                    resolve(null);
                    return;
                }
                
                const finalBlob = zip.generate({ 
                    type: "blob", 
                    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
                });
                resolve(finalBlob);
    } catch (e) {
                console.warn("XML processing failed:", e);
                reject(e);
            }
        };
        reader.onerror = reject;
            reader.readAsBinaryString(blob);
        });
    }

/**
 * 【简化版】从 ZIP 中读取状态
 * 
 * 新策略：状态主要来自 LocalStorage，ZIP 中不再存储状态。
 * 此函数仅为向后兼容保留，尝试读取旧版 Custom XML 数据。
 */
function loadStateFromZip(zip) {
    const defaultState = { backups: {}, enabledRounds: {}, formData: {} };
    
    // 尝试读取旧版 Custom XML (向后兼容)
    const backupXmlPath = "customXml/item_addin_state.xml";
    const file = zip.file(backupXmlPath) || zip.file("customXml/addin_backups.xml");
    
    if (!file) {
        // 没有旧数据，直接返回默认值
        return defaultState;
    }
    
    try {
        const xml = file.asText();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xml, "application/xml");
        const dataEl = xmlDoc.getElementsByTagName("Data")[0] || xmlDoc.getElementsByTagName("ns0:Data")[0];
        if (dataEl) {
            const encodedData = dataEl.textContent;
            const jsonStr = decodeURIComponent(escape(atob(encodedData)));
            const parsed = JSON.parse(jsonStr);
            // 兼容更旧格式 (旧格式直接是 backups 对象)
            if (!parsed.backups && !parsed.enabledRounds && !parsed.formData) {
                return { ...defaultState, backups: parsed };
            }
            return { ...defaultState, ...parsed };
        }
    } catch (e) {
        console.warn("[LoadStateFromZip] Legacy state read failed:", e);
    }
    return defaultState;
}

/**
 * 【简化版】将状态存回 ZIP
 * 
 * 新策略：不再修改 [Content_Types].xml 和 _rels/.rels，避免 500 错误。
 * 状态同步依赖 LocalStorage（同浏览器），ZIP 仅负责内容替换。
 * 
 * 如果未来需要跨设备/跨浏览器同步，可选择将状态 JSON 存入文档属性
 * (docProps/core.xml 的 description 字段) 作为备选方案。
 */
function saveStateToZip(zip, state) {
    // 新方案：不往 ZIP 注入任何新零件。
    // 状态通过 LocalStorage 实现浏览器内同步。
    // 这里仅保留接口兼容性，实际不做任何操作。
    console.log("[SaveStateToZip] Skipped - using LocalStorage for state sync instead.");
    
    // 【可选的安全存储方案】将状态存入 docProps/core.xml 的 description 字段
    // 但这会覆盖用户可能设置的文档描述，暂不启用。
    // 如需跨设备同步，可在此处启用：
    /*
    try {
        const coreXmlPath = "docProps/core.xml";
        const coreFile = zip.file(coreXmlPath);
        if (coreFile) {
            const parser = new DOMParser();
            const serializer = new XMLSerializer();
            const coreXml = coreFile.asText();
            const coreDoc = parser.parseFromString(coreXml, "application/xml");
            
            // 将状态编码后存入 dc:description
            const jsonStr = JSON.stringify(state);
            const encodedData = "ADDIN_STATE:" + btoa(unescape(encodeURIComponent(jsonStr)));
            
            let descEl = coreDoc.getElementsByTagName("dc:description")[0];
            if (!descEl) {
                descEl = coreDoc.createElementNS("http://purl.org/dc/elements/1.1/", "dc:description");
                coreDoc.documentElement.appendChild(descEl);
            }
            descEl.textContent = encodedData;
            
            zip.file(coreXmlPath, serializer.serializeToString(coreDoc));
        }
    } catch (e) {
        console.warn("[SaveStateToZip] Optional core.xml update failed:", e);
    }
    */
}

/**
 * 核心逻辑：独立处理每个文件的 XML 替换与捕获
 * 增加 targetEnabledRounds 参数，确保使用快照中的勾选状态
 */
function replaceContentControlsXmlIndependent(xml, formData, backups, targetEnabledRounds = null) {
    let hasChanges = false;
    const currentBackups = { ...backups };
    // 如果没有传入目标状态（比如本地同步），则回退到当前内存中的状态
    const activeRounds = targetEnabledRounds || enabledRounds;
    
    try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, "application/xml");
        const allSdts = doc.getElementsByTagName("w:sdt");
        
        const roundTags = [
            // 历轮投资人
            { id: "seed", tag: "Round_Seed", stateObj: "enabledRounds" },
            { id: "angel", tag: "Round_Angel", stateObj: "enabledRounds" },
            { id: "preA", tag: "Round_PreA", stateObj: "enabledRounds" },
            { id: "seriesA", tag: "Round_SeriesA", stateObj: "enabledRounds" },
            { id: "seriesB", tag: "Round_SeriesB", stateObj: "enabledRounds" },
            // 本轮投资人
            { id: "lead", tag: "Inv_Lead", stateObj: "enabledCurrentInvestors" },
            { id: "follow1", tag: "Inv_Follow1", stateObj: "enabledCurrentInvestors" },
            { id: "follow2", tag: "Inv_Follow2", stateObj: "enabledCurrentInvestors" },
            { id: "follow3", tag: "Inv_Follow3", stateObj: "enabledCurrentInvestors" }
        ];
        
        for (let i = 0; i < allSdts.length; i++) {
            const sdt = allSdts[i];
            const pr = sdt.getElementsByTagName("w:sdtPr")[0];
            if (!pr) continue;
            
            const tagEl = pr.getElementsByTagName("w:tag")[0];
            const tagVal = tagEl ? (tagEl.getAttribute("w:val") || tagEl.getAttribute("val")) : "";
            
            const roundInfo = roundTags.find(r => r.tag === tagVal);
            if (roundInfo) {
                const isEnabled = activeRounds[roundInfo.id];
                const content = sdt.getElementsByTagName("w:sdtContent")[0];
                if (!content) continue;

                const hasContent = content.childNodes.length > 0;

                if (!isEnabled && hasContent) {
                    // --- 场景：需要隐藏，但目前有内容 -> 捕获并物理删除 ---
                    const serializer = new XMLSerializer();
                    let contentXml = "";
                    for (let j = 0; j < content.childNodes.length; j++) {
                        contentXml += serializer.serializeToString(content.childNodes[j]);
                    }
                    currentBackups[tagVal] = contentXml;
                    while (content.firstChild) content.removeChild(content.firstChild);
                    hasChanges = true;
                }
                else if (isEnabled && !hasContent) {
                    // --- 场景：需要显示，但目前是空的 -> 从备份恢复 ---
                    const backupXml = currentBackups[tagVal];
                    if (backupXml) {
                        const backupDoc = parser.parseFromString(`<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${backupXml}</root>`, "application/xml");
                        const nodes = backupDoc.documentElement.childNodes;
                        while (nodes.length > 0) {
                            const importedNode = doc.importNode(nodes[0], true);
                            content.appendChild(importedNode);
                        }
                        delete currentBackups[tagVal];
                        hasChanges = true;
                    }
                }
            }
        }
// ... 后面字段替换逻辑保持不变 ...
        
        // 替换普通字段
        for (let i = 0; i < allSdts.length; i++) {
            const sdt = allSdts[i];
            const pr = sdt.getElementsByTagName("w:sdtPr")[0];
            if (!pr) continue;
            
            const tagEl = pr.getElementsByTagName("w:tag")[0];
            const aliasEl = pr.getElementsByTagName("w:alias")[0];
            const tagVal = tagEl ? (tagEl.getAttribute("w:val") || tagEl.getAttribute("val")) : "";
            const aliasVal = aliasEl ? (aliasEl.getAttribute("w:val") || aliasEl.getAttribute("val")) : "";
            
            let key = "";
            if (tagVal && formData[tagVal] !== undefined) key = tagVal;
            else if (aliasVal && formData[aliasVal] !== undefined) key = aliasVal;
            else continue;

            const target = String(formData[key] ?? "");
            const content = sdt.getElementsByTagName("w:sdtContent")[0];
            if (content) {
                const texts = content.getElementsByTagName("w:t");
                let currentText = "";
                for (let j = 0; j < texts.length; j++) currentText += texts[j].textContent;
                
                if (currentText !== target) {
                    hasChanges = true;
                    if (texts.length > 0) {
                        texts[0].textContent = target;
                        for (let j = 1; j < texts.length; j++) texts[j].textContent = "";
                    } else {
                        const t = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:t");
                        t.textContent = target;
                        const r = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:r");
                        r.appendChild(t);
                        content.appendChild(r);
                    }
                }
            }
        }

        const serializer = new XMLSerializer();
        return { 
            xml: serializer.serializeToString(doc), 
            hasChanges, 
            updatedBackups: currentBackups 
        };
    } catch (e) {
        console.error("Independent sync failed:", e);
        return { xml, hasChanges: false, updatedBackups: backups };
    }
}

/**
 * 从 Custom XML 中“解冻”备份数据到本地 Settings
 * 用于别人打开云端同步后的文件时，自动获取恢复能力。
 */
async function thawBackupsFromCustomXml() {
    if (typeof Word === 'undefined') return;
    try {
        await Word.run(async (context) => {
            const customXmlParts = context.document.customXmlParts;
            customXmlParts.load("items");
            await context.sync();
            
            console.log(`Searching Custom XML Parts... total items: ${customXmlParts.items.length}`);
            
            let foundState = null;
            for (const part of customXmlParts.items) {
                const xmlResult = part.getXml();
                await context.sync();
                const xml = xmlResult.value;
                
                // 识别全量状态 XML (支持旧版 addin_backups 和新版 AddinState)
                if (xml && (xml.includes("AddinState") || xml.includes("AddinBackups") || xml.includes("item_addin_state.xml"))) {
                    console.log("Found matching Custom XML Part, parsing...");
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(xml, "application/xml");
                    const dataEl = xmlDoc.getElementsByTagName("Data")[0] || xmlDoc.getElementsByTagName("ns0:Data")[0];
                    if (dataEl) {
                        const encodedData = dataEl.textContent;
                        const jsonStr = decodeURIComponent(escape(atob(encodedData)));
                        const parsed = JSON.parse(jsonStr);
                        
                        // 兼容旧格式 (如果是旧格式，将其包装为标准 State 结构)
                        if (!parsed.backups && !parsed.enabledRounds && !parsed.formData) {
                            console.log("Converting legacy backup format to modern State...");
                            foundState = { backups: parsed, enabledRounds: {}, formData: {} };
                        } else {
                            foundState = parsed;
                        }
                        break;
                    }
                }
            }
            
            if (foundState) {
                console.log("Cloud state found! enabledRounds:", foundState.enabledRounds);
                
                // 1. 恢复全局勾选状态
                if (foundState.enabledRounds) {
                    enabledRounds = { ...enabledRounds, ...foundState.enabledRounds };
                }
                
                // 2. 恢复全局表单内容
                if (foundState.formData) {
                    lastLoadedFormData = { ...foundState.formData };
                    console.log(`Loaded form data for ${Object.keys(lastLoadedFormData).length} fields.`);
                }

                // 3. 恢复备份段落到本地 Settings (用于勾选恢复)
                const settings = context.document.settings;
                if (foundState.backups) {
                    for (const [tag, ooxml] of Object.entries(foundState.backups)) {
                        settings.add(`${BACKUP_PREFIX}${tag}`, ooxml);
                    }
                }
                await context.sync();
                
                // 4. 联动修复：如果 UI 是开启但文档是空的，自动复活
                await autoRestoreEnabledRounds(foundState.backups || {});
            } else {
                console.log("No cloud state found in Custom XML Parts.");
            }
        });
    } catch (err) {
        console.warn("Thaw state failed:", err);
    }
}

/**
 * 自动恢复那些在 UI 中已开启但在文档中可能已被云端物理删除的段落
 */
async function autoRestoreEnabledRounds(backups) {
    const roundTags = [
        // 历轮投资人
        { id: "seed", tag: "Round_Seed", stateObj: "enabledRounds" },
        { id: "angel", tag: "Round_Angel", stateObj: "enabledRounds" },
        { id: "preA", tag: "Round_PreA", stateObj: "enabledRounds" },
        { id: "seriesA", tag: "Round_SeriesA", stateObj: "enabledRounds" },
        { id: "seriesB", tag: "Round_SeriesB", stateObj: "enabledRounds" },
        // 本轮投资人
        { id: "lead", tag: "Inv_Lead", stateObj: "enabledCurrentInvestors" },
        { id: "follow1", tag: "Inv_Follow1", stateObj: "enabledCurrentInvestors" },
        { id: "follow2", tag: "Inv_Follow2", stateObj: "enabledCurrentInvestors" },
        { id: "follow3", tag: "Inv_Follow3", stateObj: "enabledCurrentInvestors" }
    ];

    await Word.run(async (context) => {
        const allControls = context.document.contentControls;
        allControls.load("items,tag");
        await context.sync();

        for (const ctrl of allControls.items) {
            const roundInfo = roundTags.find(r => r.tag === ctrl.tag);
            if (roundInfo) {
                // 如果该轮次已启用 (enabledRounds[id])
                // 注意：enabledRounds 可能还没被表单初始化，或者默认为 false
                // 这里我们要看 backups 里是否有它
                const backupXml = backups[ctrl.tag];
                if (backupXml) {
                    // 如果本地变量显示应该开启，则恢复
                    if (enabledRounds[roundInfo.id]) {
                        console.log(`Auto-restoring ${ctrl.tag} on document load...`);
                        // 如果 backupXml 是纯 XML 节点字符串，直接 insertOoxml 
                        // 注意：我们在 cloud capture 时保存的是节点 XML 字符串
                        ctrl.insertOoxml(`<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${backupXml}</w:document></pkg:xmlData></pkg:part></pkg:package>`, "Replace");
                    }
                }
            }
        }
        await context.sync();
    });
}

// =====================================================================
// 启动时自动对齐 (Auto-Alignment)
// 当打开 B 文件时，根据 LocalStorage 中的表单状态，自动刷新文档正文
// 【优化】合并所有轮次的显隐操作到单个 Word.run 事务，避免 500 错误
// =====================================================================
async function autoAlignDocumentOnStartup() {
    console.log("[AutoAlign] Starting document alignment check...");
    
    // 从 LocalStorage 获取最新状态
    const lsState = loadFormStateFromLocalStorage();
    if (!lsState || Object.keys(lsState.formData).length === 0) {
        console.log("[AutoAlign] No LocalStorage state found, skipping alignment.");
        return;
    }
    
    // 1. 应用占位符替换（确保正文内容与表单一致）
    console.log("[AutoAlign] Applying placeholder values from LocalStorage...");
    await applyPlaceholderToCurrentDoc(lsState.formData);
    
    // 2. 【批量化】同步所有轮次段落可见性 - 单个 Word.run 事务
    console.log("[AutoAlign] Syncing round visibility (batched)...");
    await batchAlignRoundVisibility(lsState.enabledRounds);
    
    console.log("[AutoAlign] Document alignment completed.");
}

/**
 * 【分块存储版】批量对齐轮次段落可见性
 * 
 * 使用 Chunked Settings 存储大 OOXML，避免 forceSaveFailed
 * 由于分块需要多次 sync，改用顺序处理每个操作
 */
async function batchAlignRoundVisibility(targetEnabledRounds, targetEnabledInvestors = null) {
    // 强制排队
    return wordActionQueue.add(async () => {
        if (typeof Word === 'undefined') {
            console.log("[Mock] Batch align round visibility:", targetEnabledRounds, targetEnabledInvestors);
            return;
        }
        
        const roundTags = [
            // 历轮投资人
            { id: "seed", tag: "Round_Seed", stateObj: "enabledRounds" },
            { id: "angel", tag: "Round_Angel", stateObj: "enabledRounds" },
            { id: "preA", tag: "Round_PreA", stateObj: "enabledRounds" },
            { id: "seriesA", tag: "Round_SeriesA", stateObj: "enabledRounds" },
            { id: "seriesB", tag: "Round_SeriesB", stateObj: "enabledRounds" },
            // 本轮投资人
            { id: "lead", tag: "Inv_Lead", stateObj: "enabledCurrentInvestors" },
            { id: "follow1", tag: "Inv_Follow1", stateObj: "enabledCurrentInvestors" },
            { id: "follow2", tag: "Inv_Follow2", stateObj: "enabledCurrentInvestors" },
            { id: "follow3", tag: "Inv_Follow3", stateObj: "enabledCurrentInvestors" }
        ];
        
        try {
            await Word.run(async (context) => {
                const settings = context.document.settings;
                const allControls = context.document.contentControls;
                allControls.load("items,tag,text,title");
                await context.sync(); // 【同步 1: 加载控件】

                // ========== 阶段 1：分析需要的操作 ==========
                const restoreOps = [];
                const hideOps = [];
                
                // 合并两个状态对象用于查询
                const mergedState = {
                    ...targetEnabledRounds,
                    ...(targetEnabledInvestors || enabledCurrentInvestors)
                };
                
                for (const roundInfo of roundTags) {
                    const shouldBeVisible = mergedState[roundInfo.id] || false;
                    const targets = allControls.items.filter(c => c.tag === roundInfo.tag);
                    if (targets.length === 0) continue;
                    
                    const ctrl = targets[0];
                    // 判断是否有实际内容（排除占位符 [▶xxx] 格式）
                    const isHiddenPlaceholder = /^\[▶.+\]$/.test((ctrl.text || "").trim());
                    const hasContent = ctrl.text && ctrl.text.trim().length > 0 && !isHiddenPlaceholder;
                    const settingKey = `${BACKUP_PREFIX}${roundInfo.tag}`;
                    
                    if (shouldBeVisible && !hasContent) {
                        restoreOps.push({ ctrl, settingKey, tag: roundInfo.tag });
                    } else if (!shouldBeVisible && hasContent) {
                        hideOps.push({ ctrl, settingKey, tag: roundInfo.tag, title: ctrl.title });
                    }
                }
                
                if (restoreOps.length === 0 && hideOps.length === 0) {
                    console.log("[BatchAlign] No operations needed.");
                    return;
                }
                
                console.log(`[BatchAlign] Processing: ${restoreOps.length} restore, ${hideOps.length} hide`);

                // ========== 阶段 2：执行恢复操作 (使用最小 OOXML 包重建) ==========
                for (const op of restoreOps) {
                    console.log(`[BatchAlign] Restoring ${op.tag}...`);
                    const savedContent = await readFromSettingsChunked(context, settings, op.settingKey);
                    
                    if (savedContent) {
                        // 使用 buildMinimalOoxml 重建最小有效 OOXML 包
                        const finalOoxml = buildMinimalOoxml(savedContent);
                        console.log(`[BatchAlign] Built minimal OOXML for ${op.tag}: ${finalOoxml.length} bytes`);
                        
                        op.ctrl.insertText(" ", "Replace");
                        await context.sync();
                        
                        // 恢复 OOXML
                        op.ctrl.insertOoxml(finalOoxml, "Replace");
                        await context.sync();
                        console.log(`✅ [BatchAlign] Restored ${op.tag} with OOXML`);
                        
                        await deleteFromSettingsChunked(context, settings, op.settingKey);
                    } else {
                        console.log(`[BatchAlign] No backup for ${op.tag}, skipping.`);
                    }
                    
                    op.ctrl.appearance = "Tags"; // Tags 模式显示 title 标签
                }

                // ========== 阶段 3：执行隐藏操作 (纯 OOXML，带大小拦截) ==========
                const MAX_OOXML_SIZE = 3.5 * 1024 * 1024; // 3.5MB
                
                for (const op of hideOps) {
                    console.log(`[BatchAlign] Hiding ${op.tag}...`);
                    
                    const ooxmlResult = op.ctrl.getOoxml();
                    await context.sync();
                    
                    const originalOoxml = ooxmlResult.value || "";
                    const originalLength = originalOoxml.length;
                    console.log(`[BatchAlign] 原始 OOXML for ${op.tag}: ${(originalLength / 1024).toFixed(1)} KB`);
                    
                    // 【OOXML 精简】移除 media、theme 等大型资源
                    const slimmedOoxml = slimOoxml(originalOoxml);
                    const slimmedLength = slimmedOoxml.length;
                    
                    // 【方案 G 改进】精简后仍超过 3.5MB 才拦截
                    if (slimmedLength > MAX_OOXML_SIZE) {
                        const sizeMB = (slimmedLength / (1024 * 1024)).toFixed(2);
                        console.error(`❌ [BatchAlign] 精简后 OOXML 仍过大 (${sizeMB} MB) for ${op.tag}. Skipping.`);
                        continue; // 跳过此段落，继续处理其他段落
                    }
                    
                    // 使用分块存储精简后的 OOXML
                    const chunkCount = await saveToSettingsChunked(context, settings, op.settingKey, slimmedOoxml);
                    console.log(`✅ [BatchAlign] Saved OOXML for ${op.tag} in ${chunkCount} chunks`);
                    
                    // 使用可见占位符显示具体名称，让用户知道这里隐藏的是什么
                    // 【保留格式】
                    const displayName = op.title || op.tag;
                    await insertTextPreserveFormat(op.ctrl, `[▶${displayName}]`, context);
                    
                    console.log(`✅ [BatchAlign] Hidden ${op.tag}`);
                }
                
                console.log("[BatchAlign] All operations completed.");
            });
        } catch (error) {
            console.error("[BatchAlign] Failed:", error);
        }
    });
}

// ---------------- Office 初始化 ----------------
if (typeof Office !== 'undefined') {
    Office.onReady(async (info) => {
        // 在 Word Online 环境下，即便 info.host 为空也尝试初始化表单
        console.log("Office.onReady triggered", info);
        
        // #region agent log - 假设 A: 检测 pinyin-pro 全局变量
        const pinyinCheck = {
            pinyinPro: typeof pinyinPro,
            pinyin: typeof pinyin,
            window_pinyinPro: typeof window?.pinyinPro,
            window_pinyin: typeof window?.pinyin,
            windowKeys: Object.keys(window).filter(k => k.toLowerCase().includes('pinyin'))
        };
        console.log('[DEBUG-A] pinyin-pro 全局变量检测:', pinyinCheck);
        fetch('http://127.0.0.1:7242/ingest/43fd6a23-dd95-478c-a700-bed9820a26db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({location:'taskpane.js:Office.onReady',message:'pinyin-pro全局变量检测',data:pinyinCheck,timestamp:Date.now(),sessionId:'debug-session',hypothesisId:'A'})}).catch(()=>{});
        // #endregion
        
        // 0. 【新增】先加载表单配置
        await loadFormConfig();
        
        // 1. 【双轨同步】优先从 LocalStorage 加载状态
        // LocalStorage 是"真理源"，确保同浏览器下的文件状态一致
        const lsState = loadFormStateFromLocalStorage();
        if (lsState) {
            console.log("[Init] Using LocalStorage as primary state source");
            if (lsState.formData) lastLoadedFormData = lsState.formData;
            if (lsState.enabledRounds) enabledRounds = { ...enabledRounds, ...lsState.enabledRounds };
        }
        
        // 2. 【兼容旧版】尝试从 Custom XML 中"解冻"状态（如果 LocalStorage 为空）
        /* 已移除 Custom XML 干扰逻辑以提高稳定性 */
        
        // 3. 数据准备好后再构建表单
        buildForm();
        
        // 4. 【新增】加载 AI 识别的字段并渲染到表单
        try {
            const aiFields = await loadAIFieldsFromDocument();
            if (aiFields && aiFields.length > 0) {
                console.log(`[Init] 从文档加载了 ${aiFields.length} 个 AI 字段`);
                renderAIFieldsInForm(aiFields);
            }
        } catch (e) {
            console.warn("[Init] 加载 AI 字段失败:", e);
        }
        
        // 5. 绑定紧急工具按钮
        bindEmergencyTools();
        
        // 6. 【新增】注册 Content Control 进入事件，实现双向跳转
        try {
            await registerContentControlEvents();
        } catch (e) {
            console.warn("[Init] Content Control 事件注册失败:", e);
        }
        
        // 7. 【新增】检查并显示撤销按钮
        try {
            await checkAndShowUndoButton();
        } catch (e) {
            console.warn("[Init] 检查撤销按钮失败:", e);
        }
    });
} else {
    // 允许在浏览器预览模式下加载表单
    document.addEventListener("DOMContentLoaded", async () => {
        await loadFormConfig();
        buildForm();
        // 浏览器预览模式下不加载 AI 字段（需要 Word API）
        bindEmergencyTools();
    });
}

// ---------------- 紧急工具：解锁所有 Content Control ----------------
function bindEmergencyTools() {
    const unlockBtn = document.getElementById("btn-unlock-all");
    const deleteInvLeadBtn = document.getElementById("btn-delete-inv-lead");
    const clearBtn = document.getElementById("btn-clear-backups");
    const statusDiv = document.getElementById("emergency-status");
    
    if (unlockBtn) {
        unlockBtn.addEventListener("click", async () => {
            if (statusDiv) statusDiv.textContent = "正在解锁...";
            try {
                await unlockAllContentControls();
                if (statusDiv) statusDiv.textContent = "✅ 解锁完成！现在可以编辑/删除段落了";
                showNotification("所有 Content Control 已解锁", "success");
            } catch (e) {
                console.error("解锁失败:", e);
                if (statusDiv) statusDiv.textContent = "❌ 解锁失败: " + e.message;
                showNotification("解锁失败: " + e.message, "error");
            }
        });
    }
    
    if (deleteInvLeadBtn) {
        deleteInvLeadBtn.addEventListener("click", async () => {
            if (statusDiv) statusDiv.textContent = "正在删除领投方...";
            try {
                await forceDeleteContentControl("Inv_Lead");
                if (statusDiv) statusDiv.textContent = "✅ 领投方已删除（内容保留）";
                showNotification("领投方 Content Control 已删除，内容保留", "success");
            } catch (e) {
                console.error("删除失败:", e);
                if (statusDiv) statusDiv.textContent = "❌ 删除失败: " + e.message;
                showNotification("删除失败: " + e.message, "error");
            }
        });
    }
    
    if (clearBtn) {
        clearBtn.addEventListener("click", () => {
            if (statusDiv) statusDiv.textContent = "正在清除备份...";
            clearAllBackups();
            clearDocumentSettings();
            if (statusDiv) statusDiv.textContent = "✅ 备份已清除";
            showNotification("所有备份已清除", "success");
        });
    }
}

// 解锁所有 Content Control
async function unlockAllContentControls() {
    if (typeof Word === 'undefined') {
        throw new Error("Word API 不可用");
    }
    
    await Word.run(async (context) => {
        const controls = context.document.contentControls;
        controls.load("items/tag,items/title,items/cannotEdit,items/cannotDelete");
        await context.sync();
        
        console.log(`[Unlock] 找到 ${controls.items.length} 个 Content Control`);
        
        let unlockedCount = 0;
        for (const cc of controls.items) {
            if (cc.cannotEdit || cc.cannotDelete) {
                console.log(`[Unlock] 解锁: ${cc.tag} (cannotEdit=${cc.cannotEdit}, cannotDelete=${cc.cannotDelete})`);
                cc.cannotEdit = false;
                cc.cannotDelete = false;
                unlockedCount++;
            }
        }
        
        await context.sync();
        console.log(`[Unlock] ✅ 已解锁 ${unlockedCount} 个 Content Control`);
    });
}

// 清除所有备份
function clearAllBackups() {
    const keysToRemove = [];
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && (key.startsWith("Bk_") || key.startsWith("contract_addin:Bk_"))) {
            keysToRemove.push(key);
        }
    }
    keysToRemove.forEach(key => {
        localStorage.removeItem(key);
        console.log(`[ClearBackups] 已删除: ${key}`);
    });
    console.log(`[ClearBackups] ✅ 共清除 ${keysToRemove.length} 个备份`);
}

// 清除 Document Settings 中的所有备份
async function clearDocumentSettings() {
    if (typeof Word === 'undefined') return;
    
    try {
        await Word.run(async (context) => {
            const settings = context.document.settings;
            settings.load("items/key");
            await context.sync();
            
            let deletedCount = 0;
            for (const setting of settings.items) {
                if (setting.key && setting.key.startsWith("Bk_")) {
                    setting.delete();
                    deletedCount++;
                    console.log(`[ClearDocSettings] 删除: ${setting.key}`);
                }
            }
            
            await context.sync();
            console.log(`[ClearDocSettings] ✅ 共清除 ${deletedCount} 个文档设置备份`);
        });
    } catch (e) {
        console.error("[ClearDocSettings] 清除失败:", e);
    }
}

// 强制删除指定 tag 的 Content Control（保留内容）
async function forceDeleteContentControl(tag) {
    if (typeof Word === 'undefined') {
        throw new Error("Word API 不可用");
    }
    
    await Word.run(async (context) => {
        const controls = context.document.contentControls.getByTag(tag);
        controls.load("items/tag,items/title,items/cannotEdit,items/cannotDelete");
        await context.sync();
        
        if (controls.items.length === 0) {
            throw new Error(`未找到 tag="${tag}" 的 Content Control`);
        }
        
        console.log(`[ForceDelete] 找到 ${controls.items.length} 个 "${tag}" Content Control`);
        
        for (const cc of controls.items) {
            // 先解锁
            cc.cannotEdit = false;
            cc.cannotDelete = false;
        }
        await context.sync();
        
        // 再删除（保留内容）
        for (const cc of controls.items) {
            cc.delete(false); // false = 保留内容
            console.log(`[ForceDelete] 已删除: ${cc.tag}`);
        }
        await context.sync();
        
        console.log(`[ForceDelete] ✅ 删除完成`);
    });
}

// =====================================================================
// Cloud Sync 模块 (MSAL + Graph API)
// =====================================================================

// 动态获取 redirectUri，支持 localhost 和 ngrok 两种环境
const currentRedirectUri = window.location.origin + "/taskpane.html";
console.log("[MSAL] Using redirectUri:", currentRedirectUri);

const msalConfig = {
    auth: {
        clientId: "c5f5a0d8-569b-4d5c-b790-041826a5497d", 
        authority: "https://login.microsoftonline.com/common",
        redirectUri: currentRedirectUri,
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
    }
};

let msalInstance;
let accessToken = null;
let currentAccount = null;
let cloudFiles = [];

function setSyncStatus(msg, level) {
    const el = document.getElementById("sync-status");
    if (el) {
        el.textContent = msg;
        el.style.color = level === "error" ? "red" : (level === "success" ? "green" : "#666");
    }
}

function setLoginStatus(msg, level) {
    const el = document.getElementById("auth-login-status");
    if (el) {
        el.textContent = msg;
        el.style.color = level === "error" ? "red" : "green";
    }
}

async function ensureAccessToken(interactive) {
    if (accessToken) return accessToken;
    if (!msalInstance) return null;
    const account = currentAccount || msalInstance.getAllAccounts()[0];
    if (!account) return null;
    try {
        const resp = await msalInstance.acquireTokenSilent({ scopes: ["Files.ReadWrite"], account });
        accessToken = resp.accessToken;
        return accessToken;
    } catch (e) {
        if (interactive) {
            try {
                const resp = await msalInstance.acquireTokenPopup({ scopes: ["Files.ReadWrite"] });
                accessToken = resp.accessToken;
                return accessToken;
            } catch (e2) { console.error(e2); return null; }
        }
        return null;
    }
}

async function initMSAL() {
    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) handleLoginSuccess(accounts[0]);
    } catch (e) { console.error(e); }
}

document.addEventListener("DOMContentLoaded", () => {
    initMSAL();
    const btnLogin = document.getElementById("btn-login");
    const btnList = document.getElementById("btn-list-files");
    const btnBatch = document.getElementById("btn-batch-sync");
    const toggleAutoSync = document.getElementById("toggle-auto-sync");
    
    if (btnLogin) btnLogin.addEventListener("click", signIn);
    if (btnList) btnList.addEventListener("click", listFiles);
    if (btnBatch) btnBatch.addEventListener("click", batchSyncFiles);
    
    if (toggleAutoSync) {
        toggleAutoSync.addEventListener("change", () => {
            autoSyncEnabled = toggleAutoSync.checked;
            try { localStorage.setItem(LS_AUTO_SYNC, String(autoSyncEnabled)); } catch(_) {}
        });
    }
});

async function signIn() {
    try {
        const resp = await msalInstance.loginPopup({ scopes: ["User.Read", "Files.ReadWrite"] });
        handleLoginSuccess(resp.account);
    } catch (e) { showNotification("登录失败: " + e.message, "error"); console.error("登录失败:", e); }
}

function handleLoginSuccess(account) {
    document.getElementById("auth-login-container").style.display = "none";
    document.getElementById("auth-connected-container").style.display = "block";
    document.getElementById("user-name").textContent = account.name;
    currentAccount = account;
    ensureAccessToken(true).then(token => {
        if (token) {
            setSyncStatus("已连接云端", "success");
            try { listFiles(); } catch(_) {}
        }
    });
}

async function listFiles() {
    const token = await ensureAccessToken(true);
    if (!token) return;
    
    const container = document.getElementById("cloud-file-list");
    container.innerHTML = "扫描中...";
    const pathInput = document.getElementById("drive-folder-path");
    const path = pathInput ? pathInput.value.trim() : "/";
    
    try {
        let url = "https://graph.microsoft.com/v1.0/me/drive/root/children";
        if (path !== "/") {
            url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path.replace(/^\/+|\/+$/g, ""))}:/children`;
        }
        const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
        if (!res.ok) throw new Error(res.statusText);
        const data = await res.json();
        const files = (data.value || []).filter(f => f.name.endsWith(".docx"));
        cloudFiles = files;
        
        container.innerHTML = "";
        if (files.length === 0) container.innerHTML = "无 docx 文件";
        files.forEach(f => {
            const div = document.createElement("div");
            div.innerHTML = `<label><input type="checkbox" class="file-checkbox" value="${f.id}" checked> ${f.name}</label>`;
            container.appendChild(div);
        });
        document.getElementById("btn-batch-sync").disabled = false;
    } catch (e) {
        container.innerHTML = "扫描失败: " + e.message;
    }
}

// ---------------- 批量处理队列助手 ----------------
async function runConcurrentPool(items, concurrency, taskFn, onProgress) {
    const results = [];
    let index = 0;
    let completed = 0;

    const workers = new Array(Math.min(concurrency, items.length)).fill(0).map(async () => {
        while (index < items.length) {
            const i = index++;
            const item = items[i];
            try {
                const res = await taskFn(item, i);
                results.push(res);
            } catch (err) {
                results.push({ success: false, error: err, item });
            }
            completed++;
            if (onProgress) onProgress(completed, items.length, item);
        }
    });

    await Promise.all(workers);
    return results;
}

async function batchSyncFiles(arg) {
    const options = (arg && typeof arg === "object") ? arg : {};
    const silent = !!options.silent;
    const reason = options.reason || "manual";
    
    const checks = document.querySelectorAll(".file-checkbox:checked");
    if (checks.length === 0) { if (!silent) showNotification("请选择文件", "warning"); return; }
    
    // 防并发
    if (autoSyncInProgress) { autoSyncPending = true; return; }
    autoSyncInProgress = true;

    const formData = collectFormData();
    // 【修改】构建包含 UI 状态的全量快照数据
    const uiState = {
        enabledRounds: enabledRounds,
        formData: formData
    };
    
    // 指纹检测
    const ids = Array.from(checks).map(c => c.value);
    const fingerprint = buildAutoSyncFingerprint(formData, ids);
    if (reason === "auto" && fingerprint === lastAutoSyncFingerprint) {
        autoSyncInProgress = false;
        return;
    }

    const token = await ensureAccessToken(true);
    if (!token) { autoSyncInProgress = false; return; }
    
    const statusDiv = document.getElementById("sync-status");
    
    const fileItems = Array.from(checks).map(cb => ({ id: cb.value, name: cb.nextSibling.textContent }));
    
    try {
        const results = await runConcurrentPool(fileItems, 4, async (item) => {
            try {
                const dl = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${item.id}/content`, { headers: { Authorization: `Bearer ${token}` } });
                const blob = await dl.blob();
                // 【修改】传入 uiState
                const newBlob = await processDocxContentControls(blob, formData, uiState);
                
                if (newBlob) {
                    await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${item.id}/content`, {
                        method: "PUT",
                        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                        body: newBlob
                    });
                    return { status: "success" };
                } else {
                    return { status: "skipped" };
                }
            } catch (e) {
                if (e.message.includes("423")) return { status: "locked" };
                return { status: "fail" };
            }
        }, (done, total) => {
            if (statusDiv) statusDiv.textContent = `正在同步 ${done}/${total}...`;
        });
        
        let s = 0, f = 0, l = 0, k = 0;
        results.forEach(r => {
            if (r.status === "success") s++;
            else if (r.status === "fail") f++;
            else if (r.status === "locked") l++;
            else k++;
        });
        
        lastAutoSyncFingerprint = fingerprint;
        if (statusDiv) {
            statusDiv.textContent = `完成: 成功${s}, 失败${f}, 锁定${l}, 跳过${k}`;
            statusDiv.style.color = (f + l) > 0 ? "orange" : "green";
        }
    } finally {
        autoSyncInProgress = false;
        if (autoSyncPending) {
            autoSyncPending = false;
            scheduleAutoSync();
        }
    }
}

// =====================================================================
// 合同完成功能 - 清理隐藏标记并生成干净版本
// =====================================================================

// 延迟函数
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/* ==================================================================
 * 未填写字段检查系统
 * ================================================================== */

// 当前未填写字段列表
let unfilledFields = [];
let currentUnfilledIndex = 0;

/**
 * 检查所有未填写的字段
 * 返回未填写字段列表 [{id, label, sectionLabel, element}]
 */
function checkUnfilledFields() {
    const unfilled = [];
    
    contractConfig.forEach(section => {
        if (!section.fields) return;
        
        const sectionLabel = section.header?.label || section.id;
        
        section.fields.forEach((field, idx) => {
            // 查找表单中的输入元素
            const inputId = `input-${field.tag}`;
            const input = document.getElementById(inputId);
            
            if (!input) return;
            
            let isEmpty = false;
            const value = input.value?.trim() || "";
            
            // 检查是否为空或占位符
            if (input.tagName === "SELECT") {
                isEmpty = !value || value === "";
            } else if (input.tagName === "INPUT" || input.tagName === "TEXTAREA") {
                // 空值或占位符格式 [xxx]
                isEmpty = !value || /^\[.+\]$/.test(value);
            }
            
            if (isEmpty) {
                unfilled.push({
                    id: field.id || `${section.id}-${idx}`,
                    tag: field.tag,
                    label: field.label,
                    sectionId: section.id,
                    sectionLabel: sectionLabel,
                    element: input
                });
            }
        });
    });
    
    console.log(`[UnfilledCheck] 发现 ${unfilled.length} 个未填写字段`);
    return unfilled;
}

/**
 * 显示未填写字段导航条
 */
function showUnfilledFieldsBar(fields) {
    unfilledFields = fields;
    currentUnfilledIndex = 0;
    
    // 创建或获取导航条
    let bar = document.getElementById("unfilled-fields-bar");
    if (!bar) {
        bar = document.createElement("div");
        bar.id = "unfilled-fields-bar";
        bar.className = "unfilled-fields-bar";
        document.body.appendChild(bar);
    }
    
    bar.innerHTML = `
        <div class="unfilled-bar-content">
            <div class="unfilled-bar-icon">⚠️</div>
            <div class="unfilled-bar-info">
                <span class="unfilled-bar-title">发现 <strong>${fields.length}</strong> 个未填写字段</span>
                <span class="unfilled-bar-current" id="unfilled-current-label"></span>
            </div>
            <div class="unfilled-bar-nav">
                <button class="unfilled-nav-btn" id="unfilled-prev" title="上一个">◄</button>
                <span class="unfilled-nav-counter" id="unfilled-counter">1/${fields.length}</span>
                <button class="unfilled-nav-btn" id="unfilled-next" title="下一个">►</button>
            </div>
            <button class="unfilled-bar-done" id="unfilled-done">检查完毕</button>
            <button class="unfilled-bar-close" id="unfilled-close">×</button>
        </div>
    `;
    
    bar.classList.add("show");
    document.body.classList.add("has-unfilled-bar");
    
    // 跳转到第一个未填写字段
    jumpToUnfilledField(0);
    
    // 绑定事件
    document.getElementById("unfilled-prev").onclick = () => navigateUnfilledField(-1);
    document.getElementById("unfilled-next").onclick = () => navigateUnfilledField(1);
    
    // "检查完毕"按钮返回 Promise
    return new Promise((resolve) => {
        document.getElementById("unfilled-done").onclick = () => {
            hideUnfilledFieldsBar();
            resolve(true);
        };
        // 关闭按钮也返回 true（允许继续）
        document.getElementById("unfilled-close").onclick = () => {
            hideUnfilledFieldsBar();
            resolve(false); // 取消操作
        };
    });
}

/**
 * 隐藏未填写字段导航条
 */
function hideUnfilledFieldsBar() {
    const bar = document.getElementById("unfilled-fields-bar");
    if (bar) {
        bar.classList.remove("show");
    }
    document.body.classList.remove("has-unfilled-bar");
    // 移除所有高亮
    document.querySelectorAll(".unfilled-highlight").forEach(el => {
        el.classList.remove("unfilled-highlight");
    });
    unfilledFields = [];
    currentUnfilledIndex = 0;
}

/**
 * 导航到上一个/下一个未填写字段
 */
function navigateUnfilledField(direction) {
    if (unfilledFields.length === 0) return;
    
    // 移除当前高亮
    const currentField = unfilledFields[currentUnfilledIndex];
    if (currentField?.element) {
        currentField.element.closest(".form-group")?.classList.remove("unfilled-highlight");
    }
    
    // 计算新索引
    currentUnfilledIndex += direction;
    if (currentUnfilledIndex < 0) currentUnfilledIndex = unfilledFields.length - 1;
    if (currentUnfilledIndex >= unfilledFields.length) currentUnfilledIndex = 0;
    
    // 跳转到新字段
    jumpToUnfilledField(currentUnfilledIndex);
}

/**
 * 跳转到指定未填写字段
 */
function jumpToUnfilledField(index) {
    if (index < 0 || index >= unfilledFields.length) return;
    
    const field = unfilledFields[index];
    const element = field.element;
    const formGroup = element?.closest(".form-group");
    
    if (!element || !formGroup) return;
    
    // 更新计数器
    document.getElementById("unfilled-counter").textContent = `${index + 1}/${unfilledFields.length}`;
    document.getElementById("unfilled-current-label").textContent = `${field.sectionLabel} → ${field.label}`;
    
    // 滚动到字段位置
    formGroup.scrollIntoView({ behavior: "smooth", block: "center" });
    
    // 添加高亮 + 抖动效果
    formGroup.classList.add("unfilled-highlight");
    
    // 聚焦到输入框
    setTimeout(() => {
        element.focus();
    }, 300);
}

const HIDDEN_PLACEHOLDER_TEXT = "[▶已隐藏]";

/**
 * 完成合同 - 主入口函数
 * 1. 备份当前文档（如果已登录 Graph）
 * 2. 删除所有 [▶已隐藏] 标记
 * 3. 清理 Settings 中的 OOXML 备份
 * 4. 可选：移除所有 Content Control（保留内容）
 */
async function finalizeContract(options = {}) {
    const removeContentControls = options.removeContentControls || false;
    const statusDiv = document.getElementById("finalize-status");
    
    const updateStatus = (msg, color = "#666") => {
        if (statusDiv) {
            statusDiv.textContent = msg;
            statusDiv.style.color = color;
        }
        console.log(`[Finalize] ${msg}`);
    };
    
    try {
        updateStatus("正在检查表单...");
        
        // Step 0: 检查未填写字段
        const unfilled = checkUnfilledFields();
        if (unfilled.length > 0) {
            updateStatus(`⚠️ 发现 ${unfilled.length} 个未填写字段，请检查`, "orange");
            
            // 显示导航条，等待用户确认
            const userConfirmed = await showUnfilledFieldsBar(unfilled);
            
            if (!userConfirmed) {
                updateStatus("❌ 用户取消操作", "red");
                showNotification("操作已取消。请填写表单后重试。", "warning");
                return;
            }
            
            updateStatus("✅ 用户确认已检查未填写字段");
            await delay(300);
        } else {
            updateStatus("✅ 表单检查通过");
            await delay(200);
        }
        
        updateStatus("正在准备备份...");
        
        // Step 1: 尝试备份（如果已登录）
        const backupResult = await tryBackupCurrentDocument();
        
        if (backupResult.success) {
            if (backupResult.skippedByUser) {
                // 用户主动选择跳过备份
                updateStatus("⚠️ 用户选择跳过备份", "orange");
                await delay(300);
            } else {
                updateStatus(`✅ 备份已创建: ${backupResult.fileName}`);
                await delay(500);
            }
        } else {
            // 备份失败或跳过 - 要求用户确认手动备份
            let manualBackupMsg = "";
            if (backupResult.skipped) {
                const reason = backupResult.reason || "无法自动备份";
                manualBackupMsg = `⚠️ ${reason}\n\n`;
            } else {
                manualBackupMsg = `⚠️ 自动备份失败：${backupResult.error}\n\n`;
            }
            manualBackupMsg += "请先手动备份当前文档：\n";
            manualBackupMsg += "1. 在 OneDrive 中找到此文件\n";
            manualBackupMsg += "2. 右键选择「复制」创建副本\n\n";
            manualBackupMsg += "确认已完成备份后，点击「确定」继续。\n";
            manualBackupMsg += "点击「取消」放弃操作。";
            
            updateStatus("⚠️ 需要手动备份，等待用户确认...", "orange");
            
            const userConfirmed = await showConfirmDialog(manualBackupMsg, {
                title: "⚠️ 需要手动备份",
                confirmText: "已完成备份，继续",
                cancelText: "取消操作"
            });
            if (!userConfirmed) {
                updateStatus("❌ 用户取消操作", "red");
                showNotification("操作已取消。请先备份文档再重试。", "warning");
                return; // 终止操作
            }
            
            updateStatus("✅ 用户确认已完成手动备份");
            await delay(300);
        }
        
        // Step 2: 备份完成后，弹出最终确认对话框
        let confirmMsg = "备份已完成！现在确认执行清理操作：\n\n";
        confirmMsg += "• 删除所有 [▶已隐藏] 标记\n";
        confirmMsg += "• 清理内部存储的备份数据\n";
        if (removeContentControls) {
            confirmMsg += "• 移除所有埋点（Content Control）\n";
        }
        confirmMsg += "\n⚠️ 此操作不可撤销！";
        
        const finalConfirm = await showConfirmDialog(confirmMsg, {
            title: "✅ 备份完成，确认清理",
            confirmText: "确认清理",
            cancelText: "取消"
        });
        
        if (!finalConfirm) {
            updateStatus("❌ 用户取消清理操作", "red");
            showNotification("已取消清理操作。备份文件已保留。", "info");
            return;
        }
        
        // Step 3: 删除所有隐藏标记
        updateStatus("正在删除隐藏标记...");
        const deleteResult = await deleteAllHiddenPlaceholders();
        const deletedCount = (deleteResult && deleteResult.count) ? deleteResult.count : 0;
        updateStatus(`✅ 已删除 ${deletedCount} 个隐藏标记`);
        await delay(300);
        
        // Step 4: 清理 Settings 中的备份数据
        updateStatus("正在清理内部存储...");
        clearAllBackups();
        clearDocumentSettings();
        updateStatus("✅ 内部存储已清理");
        await delay(300);
        
        // Step 5: 可选 - 移除 Content Control
        if (removeContentControls) {
            updateStatus("正在移除埋点...");
            const ccResult = await removeAllContentControls();
            const removedCount = (ccResult && ccResult.count) ? ccResult.count : 0;
            updateStatus(`✅ 已移除 ${removedCount} 个埋点（内容保留）`);
            await delay(300);
        }
        
        // 完成
        updateStatus("🎉 合同已完成！文档已准备好交付。", "green");
        showNotification("合同已完成！隐藏标记已清理，文档已准备好交付。", "success");
        
    } catch (error) {
        console.error("[Finalize] 错误:", error);
        updateStatus(`❌ 操作失败: ${error.message}`, "red");
        showNotification(`合同完成失败: ${error.message}`, "error");
    }
}

/**
 * 获取已扫描的文件列表
 */
function getScannedFiles() {
    const checkboxes = document.querySelectorAll(".file-checkbox");
    const files = [];
    checkboxes.forEach(cb => {
        const label = cb.parentElement;
        const name = label ? label.textContent.trim() : "";
        files.push({
            id: cb.value,
            name: name,
            checked: cb.checked
        });
    });
    return files;
}

/**
 * 从当前文档 URL 中提取文件名
 */
function extractCurrentDocumentName() {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.document) {
        return null;
    }
    
    const docUrl = Office.context.document.url;
    if (!docUrl) return null;
    
    console.log("[Backup] Document URL:", docUrl);
    
    // 尝试多种方式提取文件名
    // 方式1: URL 路径末尾
    const pathMatch = docUrl.match(/\/([^\/]+\.docx)$/i);
    if (pathMatch) {
        return decodeURIComponent(pathMatch[1]);
    }
    
    // 方式2: URL 参数中的文件名
    const nameMatch = docUrl.match(/[?&]file=([^&]+)/i);
    if (nameMatch) {
        return decodeURIComponent(nameMatch[1]);
    }
    
    // 方式3: SharePoint/OneDrive 格式
    const spMatch = docUrl.match(/Documents[\/\\](.+\.docx)/i);
    if (spMatch) {
        return decodeURIComponent(spMatch[1]);
    }
    
    return null;
}

/**
 * 显示文件选择对话框
 * 让用户确认或选择要备份的文件
 */
function showFileSelectDialog(files, suggestedFile) {
    return new Promise((resolve) => {
        // 移除已有对话框
        const existingDialog = document.getElementById("app-file-select-dialog");
        if (existingDialog) existingDialog.remove();
        
        // 创建遮罩层
        const overlay = document.createElement("div");
        overlay.id = "app-file-select-dialog";
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 10000;
            display: flex;
            align-items: center;
            justify-content: center;
        `;
        
        // 创建对话框
        const dialog = document.createElement("div");
        dialog.style.cssText = `
            background: white;
            border-radius: 12px;
            padding: 24px;
            max-width: 450px;
            width: 90%;
            max-height: 80vh;
            overflow-y: auto;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
        `;
        
        // 标题
        const titleEl = document.createElement("h3");
        titleEl.style.cssText = `margin: 0 0 12px 0; font-size: 16px; font-weight: 600; color: #333;`;
        titleEl.textContent = "📁 选择要备份的文件";
        
        // 说明
        const descEl = document.createElement("p");
        descEl.style.cssText = `font-size: 13px; color: #666; margin: 0 0 16px 0; line-height: 1.5;`;
        if (suggestedFile) {
            descEl.innerHTML = `系统检测到您可能正在编辑 <strong>${suggestedFile.name}</strong>，请确认或选择其他文件：`;
        } else {
            descEl.textContent = "请选择当前正在编辑的文件进行备份：";
        }
        
        // 文件列表容器
        const listContainer = document.createElement("div");
        listContainer.style.cssText = `
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            margin-bottom: 16px;
        `;
        
        // 生成文件列表
        let selectedFileId = suggestedFile ? suggestedFile.id : (files.length > 0 ? files[0].id : null);
        
        files.forEach((file, index) => {
            const item = document.createElement("label");
            item.style.cssText = `
                display: flex;
                align-items: center;
                padding: 12px;
                cursor: pointer;
                border-bottom: 1px solid #f0f0f0;
                transition: background 0.2s;
            `;
            item.onmouseover = () => item.style.background = "#f5f5f5";
            item.onmouseout = () => item.style.background = "";
            
            const radio = document.createElement("input");
            radio.type = "radio";
            radio.name = "backup-file";
            radio.value = file.id;
            radio.checked = file.id === selectedFileId;
            radio.style.cssText = `margin-right: 10px; width: 16px; height: 16px;`;
            radio.onchange = () => { selectedFileId = file.id; };
            
            const nameSpan = document.createElement("span");
            nameSpan.style.cssText = `font-size: 13px; color: #333;`;
            nameSpan.textContent = file.name;
            
            // 如果是推荐文件，添加标记
            if (suggestedFile && file.id === suggestedFile.id) {
                const badge = document.createElement("span");
                badge.style.cssText = `
                    margin-left: 8px;
                    font-size: 11px;
                    background: #e3f2fd;
                    color: #1976d2;
                    padding: 2px 6px;
                    border-radius: 4px;
                `;
                badge.textContent = "推荐";
                nameSpan.appendChild(badge);
            }
            
            item.appendChild(radio);
            item.appendChild(nameSpan);
            listContainer.appendChild(item);
        });
        
        // 按钮容器
        const btnContainer = document.createElement("div");
        btnContainer.style.cssText = `display: flex; gap: 12px; justify-content: flex-end;`;
        
        // 跳过按钮
        const skipBtn = document.createElement("button");
        skipBtn.style.cssText = `
            padding: 10px 16px;
            border: 1px solid #ddd;
            background: #f5f5f5;
            color: #666;
            border-radius: 6px;
            font-size: 13px;
            cursor: pointer;
        `;
        skipBtn.textContent = "跳过备份";
        skipBtn.onclick = () => {
            overlay.remove();
            resolve({ action: "skip" });
        };
        
        // 手动备份按钮
        const manualBtn = document.createElement("button");
        manualBtn.style.cssText = `
            padding: 10px 16px;
            border: 1px solid #ddd;
            background: #fff;
            color: #333;
            border-radius: 6px;
            font-size: 13px;
            cursor: pointer;
        `;
        manualBtn.textContent = "我自己备份";
        manualBtn.onclick = () => {
            overlay.remove();
            resolve({ action: "manual" });
        };
        
        // 确认备份按钮
        const confirmBtn = document.createElement("button");
        confirmBtn.style.cssText = `
            padding: 10px 16px;
            border: none;
            background: #107c10;
            color: white;
            border-radius: 6px;
            font-size: 13px;
            cursor: pointer;
            font-weight: 500;
        `;
        confirmBtn.textContent = "备份此文件";
        confirmBtn.onclick = () => {
            const selectedFile = files.find(f => f.id === selectedFileId);
            overlay.remove();
            resolve({ action: "backup", file: selectedFile });
        };
        
        btnContainer.appendChild(skipBtn);
        btnContainer.appendChild(manualBtn);
        btnContainer.appendChild(confirmBtn);
        
        dialog.appendChild(titleEl);
        dialog.appendChild(descEl);
        dialog.appendChild(listContainer);
        dialog.appendChild(btnContainer);
        overlay.appendChild(dialog);
        
        document.body.appendChild(overlay);
    });
}

/**
 * 尝试备份当前文档到 OneDrive
 * 综合方案：自动匹配 + 用户确认/选择
 */
async function tryBackupCurrentDocument() {
    // 检查是否已登录
    if (!msalInstance) {
        return { success: false, skipped: true, reason: "未登录" };
    }
    
    const accounts = msalInstance.getAllAccounts();
    if (!accounts || accounts.length === 0) {
        return { success: false, skipped: true, reason: "未登录" };
    }
    
    try {
        const token = await ensureAccessToken(true);
        if (!token) {
            return { success: false, skipped: true, reason: "无法获取访问令牌" };
        }
        
        // 获取已扫描的文件列表
        const scannedFiles = getScannedFiles();
        if (scannedFiles.length === 0) {
            return { success: false, skipped: true, reason: "请先扫描 OneDrive 文件夹" };
        }
        
        // 尝试从当前文档 URL 提取文件名
        const currentDocName = extractCurrentDocumentName();
        console.log("[Backup] 当前文档名:", currentDocName);
        
        // 尝试匹配已扫描的文件
        let suggestedFile = null;
        if (currentDocName) {
            // 精确匹配
            suggestedFile = scannedFiles.find(f => f.name === currentDocName);
            
            // 模糊匹配（去掉扩展名比较）
            if (!suggestedFile) {
                const baseName = currentDocName.replace(/\.docx$/i, "").toLowerCase();
                suggestedFile = scannedFiles.find(f => 
                    f.name.replace(/\.docx$/i, "").toLowerCase() === baseName
                );
            }
            
            // 包含匹配
            if (!suggestedFile) {
                const baseName = currentDocName.replace(/\.docx$/i, "").toLowerCase();
                suggestedFile = scannedFiles.find(f => 
                    f.name.toLowerCase().includes(baseName) || baseName.includes(f.name.replace(/\.docx$/i, "").toLowerCase())
                );
            }
        }
        
        console.log("[Backup] 推荐文件:", suggestedFile ? suggestedFile.name : "无");
        
        // 显示文件选择对话框
        const userChoice = await showFileSelectDialog(scannedFiles, suggestedFile);
        
        if (userChoice.action === "skip") {
            // 用户选择跳过 - 视为已确认不需要备份
            return { success: true, fileName: "(用户跳过备份)", skippedByUser: true };
        }
        
        if (userChoice.action === "manual") {
            // 用户选择手动备份
            return { success: false, skipped: true, reason: "用户选择手动备份" };
        }
        
        if (userChoice.action === "backup" && userChoice.file) {
            // 用户确认备份
            const file = userChoice.file;
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
            const backupFileName = `${file.name.replace(/\.docx$/i, '')}_备份_${timestamp}.docx`;
            
            // 【重要】先强制保存当前文档，确保备份是最新版本
            try {
                await Word.run(async (context) => {
                    context.document.save();
                    await context.sync();
                    console.log("[Backup] 文档已保存");
                });
                // 等待一下让保存同步到服务器
                await new Promise(resolve => setTimeout(resolve, 3000));
            } catch (saveError) {
                console.warn("[Backup] 保存文档失败:", saveError.message);
                // 继续尝试备份
            }
            
            // 【改进】使用"下载 + 上传"方式，避免 copy 的缓存问题
            try {
                // Step 1: 获取原文件的父文件夹信息
                console.log("[Backup] 获取文件信息...");
                const fileInfoResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}?select=parentReference`, {
                    headers: { "Authorization": `Bearer ${token}` }
                });
                
                if (!fileInfoResponse.ok) {
                    const errText = await fileInfoResponse.text();
                    return { success: false, error: `获取文件信息失败: ${errText}` };
                }
                
                const fileInfo = await fileInfoResponse.json();
                const parentId = fileInfo.parentReference?.id;
                
                if (!parentId) {
                    return { success: false, error: "无法获取父文件夹信息" };
                }
                
                console.log("[Backup] 父文件夹 ID:", parentId);
                
                // Step 2: 下载当前文件内容
                console.log("[Backup] 下载文件内容...");
                const downloadResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/content`, {
                    headers: { "Authorization": `Bearer ${token}` }
                });
                
                if (!downloadResponse.ok) {
                    const errText = await downloadResponse.text();
                    return { success: false, error: `下载文件失败: ${errText}` };
                }
                
                const fileBlob = await downloadResponse.blob();
                console.log("[Backup] 文件大小:", (fileBlob.size / 1024).toFixed(2), "KB");
                
                // Step 3: 上传到同目录作为新文件
                console.log("[Backup] 上传备份文件...");
                const uploadResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${parentId}:/${encodeURIComponent(backupFileName)}:/content`, {
                    method: "PUT",
                    headers: {
                        "Authorization": `Bearer ${token}`,
                        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    },
                    body: fileBlob
                });
                
                if (uploadResponse.ok || uploadResponse.status === 201) {
                    console.log("[Backup] 备份成功:", backupFileName);
                    return { success: true, fileName: backupFileName };
                } else {
                    const errText = await uploadResponse.text();
                    return { success: false, error: `上传备份失败: ${errText}` };
                }
                
            } catch (backupError) {
                console.error("[Backup] 备份过程出错:", backupError);
                return { success: false, error: backupError.message };
            }
        }
        
        return { success: false, skipped: true, reason: "未知操作" };
        
    } catch (error) {
        console.error("[Backup] 错误:", error);
        return { success: false, error: error.message };
    }
}

/**
 * 删除文档中所有的 [▶xxx] 格式占位符
 */
async function deleteAllHiddenPlaceholders() {
    return wordActionQueue.add(async () => {
        let deletedCount = 0;
        
        await Word.run(async (context) => {
            const body = context.document.body;
            
            // 搜索所有 [▶ 开头的文本（占位符格式为 [▶xxx]）
            const searchResults = body.search("[▶", {
                matchCase: true,
                matchWildcards: false
            });
            
            context.load(searchResults, "items,text");
            await context.sync();
            
            console.log(`[Finalize] 找到 ${searchResults.items.length} 个潜在隐藏标记`);
            
            // 从后向前处理，避免索引问题
            for (let i = searchResults.items.length - 1; i >= 0; i--) {
                const range = searchResults.items[i];
                
                // 获取包含该范围的段落
                const paragraph = range.paragraphs.getFirst();
                context.load(paragraph, "text");
                await context.sync();
                
                const paraText = paragraph.text.trim();
                
                // 检查段落是否只包含 [▶xxx] 格式的占位符
                if (/^\[▶.+\]$/.test(paraText)) {
                    // 整个段落只有占位符，删除整个段落
                    paragraph.delete();
                    deletedCount++;
                } else if (paraText.includes("[▶") && paraText.includes("]")) {
                    // 段落还有其他内容，尝试精确删除占位符
                    // 扩展 range 到包含完整的 [▶xxx]
                    const expandedRange = range.expandTo(range.getRange("End").expandTo(
                        paragraph.getRange("Content")
                    ));
                    context.load(expandedRange, "text");
                    await context.sync();
                    
                    // 使用正则匹配 [▶xxx] 并删除
                    const match = expandedRange.text.match(/\[▶[^\]]+\]/);
                    if (match) {
                        const placeholderRange = paragraph.search(match[0], { matchCase: true });
                        context.load(placeholderRange, "items");
                        await context.sync();
                        if (placeholderRange.items.length > 0) {
                            placeholderRange.items[0].delete();
                            deletedCount++;
                        }
                    }
                }
            }
            
            await context.sync();
            console.log(`[Finalize] 已删除 ${deletedCount} 个隐藏标记`);
        });
        
        return { count: deletedCount };
    });
}

/**
 * 移除所有 Content Control，保留其内容
 */
async function removeAllContentControls() {
    return wordActionQueue.add(async () => {
        let removedCount = 0;
        
        await Word.run(async (context) => {
            const contentControls = context.document.contentControls;
            context.load(contentControls, "items, tag, title");
            await context.sync();
            
            console.log(`[Finalize] 找到 ${contentControls.items.length} 个 Content Control`);
            
            // 遍历所有 Content Control
            for (const cc of contentControls.items) {
                try {
                    // 获取 CC 的文本内容
                    cc.load("text");
                    await context.sync();
                    
                    const text = cc.text;
                    
                    // 删除 CC 但保留内容
                    // Word API 的 delete() 方法默认保留内容
                    cc.delete(false); // false = 保留内容
                    removedCount++;
                } catch (e) {
                    console.warn(`[Finalize] 移除 CC 失败:`, e.message);
                }
            }
            
            await context.sync();
            console.log(`[Finalize] 已移除 ${removedCount} 个 Content Control`);
        });
        
        return { count: removedCount };
    });
}

/**
 * 清理文档 Settings（所有 backup_ooxml_ 开头的设置）
 */
function clearDocumentSettings() {
    if (typeof Word === 'undefined') return;
    
    Word.run(async (context) => {
        try {
            const settings = context.document.settings;
            context.load(settings, "items");
            await context.sync();
            
            let clearedCount = 0;
            for (const setting of settings.items) {
                if (setting.key && setting.key.startsWith("backup_ooxml_")) {
                    setting.delete();
                    clearedCount++;
                }
            }
            
            await context.sync();
            console.log(`[Finalize] 已清理 ${clearedCount} 个文档设置`);
        } catch (e) {
            console.warn("[Finalize] 清理文档设置失败:", e.message);
        }
    });
}

/**
 * 初始化合同完成区域的 UI 事件
 */
function initFinalizeUI() {
    const btn = document.getElementById("btn-finalize-contract");
    const checkbox = document.getElementById("finalize-remove-cc");
    
    if (!btn) return;
    
    btn.addEventListener("click", async () => {
        const removeCC = checkbox ? checkbox.checked : false;
        
        // 禁用按钮防止重复点击
        btn.disabled = true;
        btn.textContent = "处理中...";
        
        try {
            // 直接调用 finalizeContract，备份和确认逻辑都在里面
            await finalizeContract({ removeContentControls: removeCC });
        } finally {
            btn.disabled = false;
            btn.textContent = "🎯 完成合同并清理";
        }
    });
}

// 在 Office 初始化时调用
if (typeof Office !== 'undefined') {
    Office.onReady(() => {
        // 延迟初始化，确保 DOM 已加载
        setTimeout(initFinalizeUI, 500);
        // 初始化自定义字段管理器
        setTimeout(initCustomFieldsManager, 600);
    });
}

/* ==================================================================
 * 字段管理工具
 * ================================================================== */

// 待放置的自定义字段（创建后暂存，拖拽后移入 contractConfig）
let pendingFields = [];
const PENDING_FIELDS_KEY = "contract_addin:pendingFields";

// 当前拖拽的待放置字段
let draggingPendingField = null;

/**
 * 加载待放置字段
 */
function loadPendingFields() {
    try {
        const stored = localStorage.getItem(PENDING_FIELDS_KEY);
        if (stored) {
            pendingFields = JSON.parse(stored);
            console.log("[PendingFields] 已加载", pendingFields.length, "个待放置字段");
        }
    } catch (e) {
        console.warn("[PendingFields] 加载失败:", e.message);
        pendingFields = [];
    }
}

/**
 * 保存待放置字段
 */
function savePendingFields() {
    try {
        localStorage.setItem(PENDING_FIELDS_KEY, JSON.stringify(pendingFields));
        console.log("[PendingFields] 已保存", pendingFields.length, "个待放置字段");
    } catch (e) {
        console.warn("[PendingFields] 保存失败:", e.message);
    }
}

/**
 * 渲染自定义字段列表（底部面板横向布局）
 */
function renderCustomFieldsPanel() {
    const listContainer = document.getElementById("custom-field-list");
    if (!listContainer) return;
    
    const typeLabels = {
        text: "文本",
        number: "数字",
        date: "日期",
        select: "下拉",
        radio: "单选"
    };
    
    listContainer.innerHTML = '';
    
    // 添加字段按钮
    const addBtn = document.createElement("div");
    addBtn.className = "add-field-card";
    addBtn.id = "btn-add-field";
    addBtn.innerHTML = `<i class="ms-Icon ms-Icon--Add" aria-hidden="true"></i> 添加`;
    addBtn.onclick = showAddFieldModal;
    listContainer.appendChild(addBtn);
    
    // 如果没有待放置字段，不显示任何提示
    if (pendingFields.length === 0) {
        return;
    }
    
    // 渲染每个待放置字段卡片
    pendingFields.forEach(field => {
        const card = document.createElement("div");
        card.className = "custom-field-card";
        card.id = `pending-card-${field.id}`;
        card.draggable = true;
        card.dataset.fieldId = field.id;
        
        card.innerHTML = `
            <div class="field-label" title="${escapeHtml(field.label)}">${escapeHtml(field.label)}</div>
            <div class="field-meta">${typeLabels[field.type] || field.type}</div>
            <button class="field-delete-btn" data-action="delete" title="删除">×</button>
        `;
        
        // 删除按钮事件
        card.querySelector(".field-delete-btn").onclick = (e) => {
            e.stopPropagation();
            deletePendingField(field.id);
        };
        
        // 拖拽事件
        card.addEventListener("dragstart", handlePendingFieldDragStart);
        card.addEventListener("dragend", handlePendingFieldDragEnd);
        
        listContainer.appendChild(card);
    });
}

/**
 * 删除待放置字段
 */
function deletePendingField(fieldId) {
    const field = pendingFields.find(f => f.id === fieldId);
    if (!field) return;
    
    showConfirmDialog(`确定要删除字段 "${field.label}" 吗？`, {
        confirmText: "删除",
        cancelText: "取消",
        confirmStyle: "background:#ef4444;color:#fff;"
    }).then(confirmed => {
        if (confirmed) {
            pendingFields = pendingFields.filter(f => f.id !== fieldId);
            savePendingFields();
            renderCustomFieldsPanel();
            showNotification(`已删除字段: ${field.label}`, "success");
        }
    });
}

/**
 * 待放置字段拖拽开始
 */
function handlePendingFieldDragStart(e) {
    const card = e.target.closest(".custom-field-card");
    if (!card) return;
    
    draggingPendingField = pendingFields.find(f => f.id === card.dataset.fieldId);
    if (!draggingPendingField) return;
    
    card.classList.add("dragging");
    document.body.classList.add("dragging-field");
    
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", JSON.stringify({ type: "pending", fieldId: draggingPendingField.id }));
    
    // 显示放置区
    showDropZones();
    
    console.log("[DragDrop] 开始拖拽待放置字段:", draggingPendingField.label);
}

/**
 * 待放置字段拖拽结束
 */
function handlePendingFieldDragEnd(e) {
    const card = e.target.closest(".custom-field-card");
    if (card) card.classList.remove("dragging");
    
    document.body.classList.remove("dragging-field");
    draggingPendingField = null;
    
    hideDropZones();
    
    console.log("[DragDrop] 拖拽结束");
}

/**
 * HTML 转义
 */
function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

/**
 * 显示添加字段弹窗
 */
function showAddFieldModal() {
    const modal = document.getElementById("add-field-modal");
    if (modal) {
        modal.classList.add("show");
        // 清空表单
        const labelInput = document.getElementById("field-label");
        labelInput.value = "";
        document.getElementById("field-type").value = "text";
        document.getElementById("options-group").style.display = "none";
        document.getElementById("tag-preview").style.display = "none";
        document.getElementById("tag-preview-text").textContent = "";
        
        // 重置选项列表
        resetAddOptions();
        
        // 重置插入模式选择
        document.querySelectorAll("#add-field-modal .insert-mode-option").forEach(opt => {
            opt.classList.remove("selected");
            if (opt.dataset.mode === "insert") {
                opt.classList.add("selected");
                opt.querySelector("input").checked = true;
            }
        });
        
        // 设置弹窗标题
        document.getElementById("modal-title").textContent = "添加新字段";
        document.getElementById("modal-confirm").textContent = "创建字段";
        
        // 聚焦到名称输入框
        setTimeout(() => labelInput.focus(), 100);
    }
}

/**
 * 隐藏添加字段弹窗
 */
function hideAddFieldModal() {
    const modal = document.getElementById("add-field-modal");
    if (modal) {
        modal.classList.remove("show");
    }
}

/**
 * 隐藏编辑字段弹窗
 */
function hideEditFieldModal() {
    const modal = document.getElementById("edit-field-modal");
    if (modal) {
        modal.classList.remove("show");
    }
}

/**
 * 更新 Tag 预览（基于拼音转换）
 */
function updateTagPreview() {
    const label = document.getElementById("field-label").value.trim();
    const tagPreview = document.getElementById("tag-preview");
    const tagPreviewText = document.getElementById("tag-preview-text");
    
    if (label) {
        const tag = toPinyin(label);
        tagPreviewText.textContent = tag;
        tagPreview.style.display = "block";
    } else {
        tagPreview.style.display = "none";
    }
}

// ========== 选项管理（支持多行选项） ==========

// 临时选项存储（添加字段弹窗用）
let tempAddOptions = [];
// 临时选项存储（编辑字段弹窗用）
let tempEditOptions = [];

/**
 * 渲染选项列表
 * @param {string} containerId - 列表容器 ID
 * @param {Array} options - 选项数组
 * @param {string} mode - 'add' 或 'edit'
 */
function renderOptionsList(containerId, options, mode) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    if (options.length === 0) {
        container.innerHTML = '<div class="options-empty">暂无选项，点击下方按钮添加</div>';
        return;
    }
    
    container.innerHTML = options.map((opt, idx) => `
        <div class="option-item" data-index="${idx}">
            <span class="option-index">${idx + 1}</span>
            <div class="option-text" title="${escapeHtml(opt)}">${escapeHtml(opt)}</div>
            <button type="button" class="option-delete" onclick="removeOption(${idx}, '${mode}')">×</button>
        </div>
    `).join('');
}

/**
 * 显示添加选项的弹窗
 * @param {string} mode - 'add' 或 'edit'
 */
function showAddOptionModal(mode) {
    // 移除已存在的弹窗
    const existing = document.getElementById("option-input-modal");
    if (existing) existing.remove();
    
    const modal = document.createElement("div");
    modal.id = "option-input-modal";
    modal.className = "option-input-modal";
    modal.innerHTML = `
        <div class="option-input-box">
            <h4>添加选项</h4>
            <textarea id="new-option-text" placeholder="输入选项内容（支持多行）"></textarea>
            <div class="option-input-actions">
                <button type="button" class="btn-cancel" onclick="closeAddOptionModal()">取消</button>
                <button type="button" class="btn-confirm" onclick="confirmAddOption('${mode}')">确定</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);
    
    // 自动聚焦
    setTimeout(() => {
        document.getElementById("new-option-text")?.focus();
    }, 100);
}

/**
 * 关闭添加选项弹窗
 */
function closeAddOptionModal() {
    const modal = document.getElementById("option-input-modal");
    if (modal) modal.remove();
}

/**
 * 确认添加选项
 */
function confirmAddOption(mode) {
    const textarea = document.getElementById("new-option-text");
    const value = textarea?.value.trim();
    
    if (!value) {
        showNotification("请输入选项内容", "error");
        return;
    }
    
    if (mode === 'add') {
        tempAddOptions.push(value);
        renderOptionsList("field-options-list", tempAddOptions, 'add');
    } else if (mode === 'edit') {
        tempEditOptions.push(value);
        renderOptionsList("ufm-options-list", tempEditOptions, 'edit');
    }
    
    closeAddOptionModal();
}

/**
 * 移除选项
 */
function removeOption(index, mode) {
    if (mode === 'add') {
        tempAddOptions.splice(index, 1);
        renderOptionsList("field-options-list", tempAddOptions, 'add');
    } else if (mode === 'edit') {
        tempEditOptions.splice(index, 1);
        renderOptionsList("ufm-options-list", tempEditOptions, 'edit');
    }
}

/**
 * 重置添加字段弹窗的选项
 */
function resetAddOptions() {
    tempAddOptions = [];
    renderOptionsList("field-options-list", tempAddOptions, 'add');
}

/**
 * 设置编辑字段弹窗的选项
 */
function setEditOptions(options) {
    tempEditOptions = [...(options || [])];
    renderOptionsList("ufm-options-list", tempEditOptions, 'edit');
}

/**
 * 添加新字段（统一添加到 contractConfig）
 */
function addCustomFieldFromModal() {
    const label = document.getElementById("field-label").value.trim();
    const type = document.getElementById("field-type").value;
    const insertMode = document.querySelector('#add-field-modal input[name="insert-mode"]:checked')?.value || "insert";
    
    // 验证
    if (!label) {
        showNotification("请输入字段名称", "error");
        return;
    }
    
    // 自动生成 Tag（拼音）
    let tag = toPinyin(label);
    
    // 检查 tag 是否在 contractConfig 和 pendingFields 中重复
    let counter = 1;
    let originalTag = tag;
    const allTags = [];
    contractConfig.forEach(sec => {
        if (sec.fields) {
            sec.fields.forEach(f => allTags.push(f.tag));
        }
    });
    pendingFields.forEach(f => allTags.push(f.tag));
    while (allTags.includes(tag)) {
        tag = originalTag + counter;
        counter++;
    }
    
    // 从临时数组获取选项
    let options = [...tempAddOptions];
    
    // 选择类型需要至少一个选项
    if (type === "select" || type === "radio") {
        if (options.length === 0) {
            showNotification("请添加至少一个选项", "error");
            return;
        }
    }
    
    // 创建字段对象
    const newField = {
        id: "pending_" + Date.now(),
        label,
        tag,
        type,
        options: options.length > 0 ? options : undefined,
        hasParagraphToggle: insertMode === "paragraph" || insertMode === "both"
    };
    
    // 添加到待放置区
    pendingFields.push(newField);
    savePendingFields();
    
    // 重置选项
    resetAddOptions();
    
    // 重新渲染底部面板
    renderCustomFieldsPanel();
    
    hideAddFieldModal();
    showNotification(`已创建字段: ${label}，请拖拽到表单中放置`, "success");
}


/* ------------------------------------------------------------------
 * 字段编辑系统 (Field Editor)
 * ------------------------------------------------------------------ */

// 正在拖拽的表单字段信息
let draggingFormFieldInfo = null;

/**
 * 表单字段拖拽开始
 */
function handleFieldDragStart(e) {
    const formGroup = e.target.closest(".form-group");
    if (!formGroup) return;
    
    // 如果从 input/select/textarea/button 开始拖拽，阻止（让用户正常操作这些元素）
    const targetTag = e.target.tagName.toLowerCase();
    if (["input", "select", "textarea", "button"].includes(targetTag)) {
        e.preventDefault();
        return;
    }
    
    const sectionId = formGroup.dataset.sectionId;
    const fieldId = formGroup.dataset.fieldId;
    const fieldIndex = parseInt(formGroup.dataset.fieldIndex || 0);
    
    draggingFormFieldInfo = { sectionId, fieldId, fieldIndex, element: formGroup };
    formGroup.classList.add("dragging");
    document.body.classList.add("dragging-field");
    
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", JSON.stringify({ type: "form-field", sectionId, fieldId, fieldIndex }));
    
    // 显示放置区
    showDropZones();
    
    console.log("[FieldDrag] 开始拖拽:", { sectionId, fieldId, fieldIndex });
}

/**
 * 表单字段拖拽结束
 */
function handleFieldDragEnd(e) {
    const formGroup = e.target.closest(".form-group");
    if (formGroup) formGroup.classList.remove("dragging");
    
    document.body.classList.remove("dragging-field");
    draggingFormFieldInfo = null;
    
    hideDropZones();
    
    console.log("[FieldDrag] 拖拽结束");
}

/**
 * 编辑 Section 中的字段
 */
function editFieldInSection(sectionId, fieldIdOrIndex) {
    // 查找 section 和字段
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields) {
        console.error("[EditField] 未找到 section:", sectionId);
        return;
    }
    
    let field;
    let fieldIndex;
    if (typeof fieldIdOrIndex === "number") {
        fieldIndex = fieldIdOrIndex;
        field = section.fields[fieldIndex];
    } else {
        fieldIndex = section.fields.findIndex(f => f.id === fieldIdOrIndex);
        field = section.fields[fieldIndex];
    }
    
    if (!field) {
        console.error("[EditField] 未找到字段:", fieldIdOrIndex);
        return;
    }
    
    // 打开编辑弹窗
    showFieldEditModal(sectionId, fieldIndex, field);
}

/**
 * 显示字段编辑弹窗
 */
function showFieldEditModal(sectionId, fieldIndex, field) {
    // 如果弹窗不存在，创建它
    let modal = document.getElementById("universal-field-edit-modal");
    if (!modal) {
        modal = document.createElement("div");
        modal.id = "universal-field-edit-modal";
        modal.className = "modal-overlay";
        modal.innerHTML = `
            <div class="modal-content" style="max-width: 480px;">
                <div class="modal-header">
                    <h3 id="ufm-title">编辑字段</h3>
                    <button class="modal-close" onclick="hideFieldEditModal()">&times;</button>
                </div>
                <div class="modal-body">
                    <input type="hidden" id="ufm-section-id">
                    <input type="hidden" id="ufm-field-index">
                    
                    <div class="form-group-modal">
                        <label>字段名称</label>
                        <input type="text" id="ufm-label" class="modal-input" placeholder="如：签订时间">
                    </div>
                    
                    <div class="form-group-modal">
                        <label>Tag 标签 (只读)</label>
                        <input type="text" id="ufm-tag" class="modal-input" readonly style="background:#f1f5f9;color:#64748b;">
                    </div>
                    
                    <div class="form-group-modal">
                        <label>字段类型</label>
                        <select id="ufm-type" class="modal-input" onchange="onUfmTypeChange()">
                            <option value="text">文本</option>
                            <option value="number">数字</option>
                            <option value="date">日期</option>
                            <option value="select">下拉选择</option>
                            <option value="radio">单选按钮</option>
                        </select>
                    </div>
                    
                    <div class="form-group-modal" id="ufm-options-group" style="display:none;">
                        <label>选项列表</label>
                        <div class="options-list" id="ufm-options-list"></div>
                        <button type="button" class="add-option-btn" onclick="showAddOptionModal('edit')">
                            <i class="ms-Icon ms-Icon--Add" aria-hidden="true"></i> 添加选项
                        </button>
                    </div>
                    
                    <div class="form-group-modal">
                        <label>移动到 Section</label>
                        <select id="ufm-target-section" class="modal-input">
                        </select>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="modal-btn danger" onclick="deleteFieldInSection()">
                        <i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i> 删除字段
                    </button>
                    <div style="flex:1;"></div>
                    <button class="modal-btn secondary" onclick="hideFieldEditModal()">取消</button>
                    <button class="modal-btn primary" onclick="saveFieldEdit()">保存</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }
    
    // 填充表单
    document.getElementById("ufm-section-id").value = sectionId;
    document.getElementById("ufm-field-index").value = fieldIndex;
    document.getElementById("ufm-label").value = field.label || "";
    document.getElementById("ufm-tag").value = field.tag || "";
    document.getElementById("ufm-type").value = field.type || "text";
    
    // 选项 - 使用选项列表
    const optionsGroup = document.getElementById("ufm-options-group");
    if (field.type === "select" || field.type === "radio") {
        optionsGroup.style.display = "block";
        setEditOptions(field.options || []);
    } else {
        optionsGroup.style.display = "none";
        setEditOptions([]);
    }
    
    // 填充目标 Section 下拉框
    const targetSelect = document.getElementById("ufm-target-section");
    targetSelect.innerHTML = "";
    contractConfig.forEach(sec => {
        if (sec.fields) { // 只显示有 fields 的普通 section
            const opt = document.createElement("option");
            opt.value = sec.id;
            opt.textContent = sec.header.label;
            if (sec.id === sectionId) opt.selected = true;
            targetSelect.appendChild(opt);
        }
    });
    
    modal.classList.add("show");
}

/**
 * 类型改变时显示/隐藏选项区
 */
function onUfmTypeChange() {
    const type = document.getElementById("ufm-type").value;
    const optionsGroup = document.getElementById("ufm-options-group");
    const isOptionType = (type === "select" || type === "radio");
    optionsGroup.style.display = isOptionType ? "block" : "none";
    
    // 如果切换到非选项类型，清空选项列表
    if (!isOptionType) {
        setEditOptions([]);
    }
}

/**
 * 隐藏字段编辑弹窗
 */
function hideFieldEditModal() {
    const modal = document.getElementById("universal-field-edit-modal");
    if (modal) modal.classList.remove("show");
}

/**
 * 保存字段编辑
 */
function saveFieldEdit() {
    const sectionId = document.getElementById("ufm-section-id").value;
    const fieldIndex = parseInt(document.getElementById("ufm-field-index").value);
    const newLabel = document.getElementById("ufm-label").value.trim();
    const newType = document.getElementById("ufm-type").value;
    const targetSectionId = document.getElementById("ufm-target-section").value;
    
    if (!newLabel) {
        showNotification("请输入字段名称", "error");
        return;
    }
    
    // 查找并更新字段
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields || !section.fields[fieldIndex]) {
        showNotification("字段不存在", "error");
        return;
    }
    
    const field = section.fields[fieldIndex];
    
    // 更新字段属性
    field.label = newLabel;
    field.type = newType;
    
    // 更新选项 - 从临时数组读取
    if (newType === "select" || newType === "radio") {
        if (tempEditOptions.length === 0) {
            showNotification("请添加至少一个选项", "error");
            return;
        }
        field.options = [...tempEditOptions];
    } else {
        // 非选项类型，清除选项
        delete field.options;
    }
    
    // 如果需要移动到其他 Section
    if (targetSectionId !== sectionId) {
        // 从原 section 移除
        section.fields.splice(fieldIndex, 1);
        
        // 添加到目标 section
        const targetSection = contractConfig.find(s => s.id === targetSectionId);
        if (targetSection && targetSection.fields) {
            targetSection.fields.push(field);
        }
    }
    
    // 保存配置
    saveFormConfig();
    
    // 重新构建表单
    buildForm();
    
    hideFieldEditModal();
    showNotification(`字段 "${newLabel}" 已更新`, "success");
}

/**
 * 删除 Section 中的字段
 */
function deleteFieldInSection() {
    const sectionId = document.getElementById("ufm-section-id").value;
    const fieldIndex = parseInt(document.getElementById("ufm-field-index").value);
    
    const section = contractConfig.find(s => s.id === sectionId);
    if (!section || !section.fields || !section.fields[fieldIndex]) {
        showNotification("字段不存在", "error");
        hideFieldEditModal(); // 关闭弹窗
        return;
    }
    
    const field = section.fields[fieldIndex];
    const fieldLabel = field.label; // 保存名称用于提示
    
    // 先关闭编辑弹窗
    hideFieldEditModal();
    
    showConfirmDialog(`确定要删除字段 "${fieldLabel}" 吗？此操作不可撤销。`, {
        confirmText: "删除",
        cancelText: "取消",
        confirmStyle: "background:#ef4444;color:#fff;"
    }).then(confirmed => {
        if (confirmed) {
            // 删除字段
            section.fields.splice(fieldIndex, 1);
            
            // 保存配置
            saveFormConfig();
            
            // 重新构建表单
            buildForm();
            
            showNotification(`字段 "${fieldLabel}" 已删除`, "success");
        }
    });
}

/**
 * 在配置中移动字段位置
 */
function moveFieldInConfig(fromSectionId, toSectionId, fromIndex, toIndex) {
    const fromSection = contractConfig.find(s => s.id === fromSectionId);
    if (!fromSection || !fromSection.fields) return false;
    
    const field = fromSection.fields[fromIndex];
    if (!field) return false;
    
    // 从原位置移除
    fromSection.fields.splice(fromIndex, 1);
    
    if (fromSectionId === toSectionId) {
        // 同 Section 内移动
        // 如果目标位置在原位置之后，需要调整索引
        const adjustedIndex = toIndex > fromIndex ? toIndex - 1 : toIndex;
        fromSection.fields.splice(adjustedIndex, 0, field);
    } else {
        // 跨 Section 移动
        const toSection = contractConfig.find(s => s.id === toSectionId);
        if (!toSection || !toSection.fields) return false;
        toSection.fields.splice(toIndex, 0, field);
    }
    
    return true;
}

/* ------------------------------------------------------------------
 * 拖拽系统 (Drag & Drop)
 * ------------------------------------------------------------------ */


/**
 * 显示放置区
 */
function showDropZones() {
    // 移除旧的放置区，重新创建以确保位置正确
    document.querySelectorAll(".drop-zone").forEach(z => z.remove());
    
    // 在每个 section header 后添加放置区
    document.querySelectorAll(".section-header-container").forEach(header => {
        const sectionId = header.id.replace("section-nav-", "");
        
        const zone = document.createElement("div");
        zone.className = "drop-zone";
        zone.dataset.headerId = header.id;
        zone.dataset.targetSectionId = sectionId;
        zone.dataset.targetIndex = "0"; // 放在 section 最前面
        zone.addEventListener("dragover", handleDragOver);
        zone.addEventListener("dragleave", handleDragLeave);
        zone.addEventListener("drop", handleDrop);
        
        // 插入到 section-fields 之前（header 后面）
        const nextSibling = header.nextElementSibling;
        if (nextSibling) {
            header.parentNode.insertBefore(zone, nextSibling);
        } else {
            header.parentNode.appendChild(zone);
        }
        zone.style.display = "block";
    });
    
    // 在每个 form-group 后添加放置区（更精细的位置控制）
    document.querySelectorAll("#dynamic-form-container .form-group").forEach((formGroup) => {
        const sectionId = formGroup.dataset.sectionId;
        const fieldIndex = parseInt(formGroup.dataset.fieldIndex || 0);
        
        const zone = document.createElement("div");
        zone.className = "drop-zone";
        zone.dataset.afterGroup = formGroup.id;
        zone.dataset.afterElement = formGroup.id;
        zone.dataset.targetSectionId = sectionId;
        zone.dataset.targetIndex = String(fieldIndex + 1); // 放在该字段之后
        zone.addEventListener("dragover", handleDragOver);
        zone.addEventListener("dragleave", handleDragLeave);
        zone.addEventListener("drop", handleDrop);
        
        formGroup.parentNode.insertBefore(zone, formGroup.nextSibling);
        zone.style.display = "block";
    });
}

/**
 * 隐藏放置区
 */
function hideDropZones() {
    document.querySelectorAll(".drop-zone").forEach(zone => {
        zone.style.display = "none";
        zone.classList.remove("drag-over");
    });
}

/**
 * 拖拽经过放置区
 */
function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    e.target.classList.add("drag-over");
}

/**
 * 拖拽离开放置区
 */
function handleDragLeave(e) {
    e.target.classList.remove("drag-over");
}

/**
 * 放置
 */
function handleDrop(e) {
    e.preventDefault();
    e.target.classList.remove("drag-over");
    
    const zone = e.target.closest(".drop-zone");
    if (!zone) {
        console.log("[DragDrop] 未找到放置区");
        return;
    }
    
    const targetSectionId = zone.dataset.targetSectionId;
    const targetIndex = parseInt(zone.dataset.targetIndex || 0);
    
    // 1. 表单字段拖拽（在 contractConfig 中移动位置）
    if (draggingFormFieldInfo) {
        const { sectionId: fromSectionId, fieldIndex: fromIndex } = draggingFormFieldInfo;
        
        console.log("[DragDrop] 表单字段移动:", { fromSectionId, fromIndex, targetSectionId, targetIndex });
        
        if (fromSectionId === targetSectionId && fromIndex === targetIndex) {
            console.log("[DragDrop] 位置未变化，忽略");
        } else {
            const success = moveFieldInConfig(fromSectionId, targetSectionId, fromIndex, targetIndex);
            if (success) {
                // 【修复】保存滚动位置
                const scrollTop = document.querySelector(".main-content")?.scrollTop || 0;
                
                saveFormConfig();
                buildForm();
                
                // 【修复】延迟恢复滚动位置，确保 DOM 渲染完成
                requestAnimationFrame(() => {
                    const mc = document.querySelector(".main-content");
                    if (mc) mc.scrollTop = scrollTop;
                });
                
                showNotification("字段位置已更新", "success");
            } else {
                showNotification("移动失败", "error");
            }
        }
        
        draggingFormFieldInfo = null;
    }
    // 2. 待放置字段拖拽（从 pendingFields 移入 contractConfig）
    else if (draggingPendingField) {
        console.log("[DragDrop] 放置待放置字段:", draggingPendingField.label, "到", targetSectionId, "位置", targetIndex);
        
        // 找到目标 section
        const targetSection = contractConfig.find(s => s.id === targetSectionId);
        if (targetSection && targetSection.fields) {
            // 创建正式字段对象（修改 ID 前缀）
            const newField = {
                ...draggingPendingField,
                id: draggingPendingField.id.replace("pending_", "field_")
            };
            
            // 插入到目标位置
            targetSection.fields.splice(targetIndex, 0, newField);
            
            // 从 pendingFields 移除
            pendingFields = pendingFields.filter(f => f.id !== draggingPendingField.id);
            
            // 保存
            saveFormConfig();
            savePendingFields();
            
            // 【修复】保存滚动位置
            const scrollTop = document.querySelector(".main-content")?.scrollTop || 0;
            
            // 重新渲染
            buildForm();
            renderCustomFieldsPanel();
            
            // 【修复】延迟恢复滚动位置，确保 DOM 渲染完成
            requestAnimationFrame(() => {
                const mc = document.querySelector(".main-content");
                if (mc) mc.scrollTop = scrollTop;
            });
            
            showNotification(`已将 "${newField.label}" 放置到表单`, "success");
        } else {
            showNotification("无法放置到该位置", "error");
        }
        
        draggingPendingField = null;
    }
    else {
        console.log("[DragDrop] 没有正在拖拽的字段");
    }
    
    // 隐藏放置区
    hideDropZones();
    document.body.classList.remove("dragging-field");
}




/**
 * 初始化字段管理器（统一版本）
 */
function initCustomFieldsManager() {
    console.log("[FieldManager] 初始化字段管理器...");
    
    // 加载待放置字段
    loadPendingFields();
    
    // 渲染底部面板（显示待放置字段卡片）
    renderCustomFieldsPanel();
    
    // FAB 按钮点击 - 切换底部面板
    const fab = document.getElementById("custom-field-fab");
    const drawer = document.getElementById("custom-field-drawer");
    
    if (fab && drawer) {
        fab.addEventListener("click", () => {
            const isOpen = drawer.classList.contains("open");
            if (isOpen) {
                drawer.classList.remove("open");
                fab.classList.remove("active");
                hideDropZones();
            } else {
                drawer.classList.add("open");
                fab.classList.add("active");
            }
        });
    }
    
    // 关闭抽屉按钮
    const closeBtn = document.getElementById("drawer-close");
    if (closeBtn && drawer && fab) {
        closeBtn.addEventListener("click", () => {
            drawer.classList.remove("open");
            fab.classList.remove("active");
            hideDropZones();
        });
    }
    
    // 弹窗关闭按钮
    const modalClose = document.getElementById("modal-close");
    const modalCancel = document.getElementById("modal-cancel");
    if (modalClose) modalClose.addEventListener("click", hideAddFieldModal);
    if (modalCancel) modalCancel.addEventListener("click", hideAddFieldModal);
    
    // 弹窗确认按钮
    const modalConfirm = document.getElementById("modal-confirm");
    if (modalConfirm) {
        modalConfirm.addEventListener("click", addCustomFieldFromModal);
    }
    
    // 字段名称输入 - 实时更新 Tag 预览
    const fieldLabel = document.getElementById("field-label");
    if (fieldLabel) {
        fieldLabel.addEventListener("input", updateTagPreview);
    }
    
    // 字段类型切换显示选项输入
    const fieldType = document.getElementById("field-type");
    const optionsGroup = document.getElementById("options-group");
    if (fieldType && optionsGroup) {
        fieldType.addEventListener("change", () => {
            if (fieldType.value === "select" || fieldType.value === "radio") {
                optionsGroup.style.display = "block";
            } else {
                optionsGroup.style.display = "none";
                // 清空选项
                resetAddOptions();
            }
        });
    }
    
    // 添加选项按钮事件
    const addOptionBtn = document.getElementById("add-option-btn");
    if (addOptionBtn) {
        addOptionBtn.addEventListener("click", () => showAddOptionModal('add'));
    }
    
    // 插入模式选择
    document.querySelectorAll("#add-field-modal .insert-mode-option").forEach(option => {
        option.addEventListener("click", () => {
            document.querySelectorAll("#add-field-modal .insert-mode-option").forEach(o => o.classList.remove("selected"));
            option.classList.add("selected");
            option.querySelector("input").checked = true;
        });
    });
    
    // 导出完整配置按钮
    const exportBtn = document.getElementById("btn-export-config");
    if (exportBtn) {
        exportBtn.addEventListener("click", exportFullFormConfig);
    }
    
    // 导入完整配置按钮
    const importBtn = document.getElementById("btn-import-config");
    const importInput = document.getElementById("import-config-input");
    if (importBtn && importInput) {
        importBtn.addEventListener("click", () => importInput.click());
        importInput.addEventListener("change", (e) => {
            if (e.target.files.length > 0) {
                importFullFormConfig(e.target.files[0]);
                e.target.value = ""; // 重置以允许再次选择同一文件
            }
        });
    }
    
    // 重置配置按钮
    const resetBtn = document.getElementById("btn-reset-config");
    if (resetBtn) {
        resetBtn.addEventListener("click", () => {
            showConfirmDialog("确定要重置为默认配置吗？所有自定义修改将丢失。", {
                confirmText: "重置",
                cancelText: "取消"
            }).then(confirmed => {
                if (confirmed) resetFormConfig();
            });
        });
    }
    
    // ========== 编辑字段弹窗事件 ==========
    const editModalClose = document.getElementById("edit-modal-close");
    const editModalCancel = document.getElementById("edit-modal-cancel");
    const editModalConfirm = document.getElementById("edit-modal-confirm");
    const editModalDelete = document.getElementById("edit-modal-delete");
    
    if (editModalClose) editModalClose.addEventListener("click", hideEditFieldModal);
    if (editModalCancel) editModalCancel.addEventListener("click", hideEditFieldModal);
    if (editModalConfirm) editModalConfirm.addEventListener("click", saveEditedField);
    if (editModalDelete) editModalDelete.addEventListener("click", deleteFieldFromEditModal);
    
    // 编辑弹窗 - 字段类型切换
    const editFieldType = document.getElementById("edit-field-type");
    const editOptionsGroup = document.getElementById("edit-options-group");
    if (editFieldType && editOptionsGroup) {
        editFieldType.addEventListener("change", () => {
            if (editFieldType.value === "select" || editFieldType.value === "radio") {
                editOptionsGroup.style.display = "block";
            } else {
                editOptionsGroup.style.display = "none";
            }
        });
    }
    
    console.log("[CustomFields] 初始化完成");
}

