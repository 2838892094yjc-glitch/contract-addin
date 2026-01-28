const PEVC_DEFAULT_TEMPLATE = 
[
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
