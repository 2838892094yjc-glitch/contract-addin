// 合同变量配置清单
// 格式说明：
// id: HTML元素的唯一ID（不要重复）
// label: 界面上显示的提示文字
// tag: Word文档里内容控件的 Tag（必须精确匹配）
// placeholder: 输入框里的灰色提示字
// type: 输入框类型 (text: 普通文本, date: 日期, number: 数字)

const contractConfig = [
    {
        id: "companyName",
        label: "目标公司名称",
        tag: "CompanyName",
        placeholder: "例如：北京某某科技有限公司",
        type: "text"
    },
    {
        id: "partyA",
        label: "甲方 (投资人)",
        tag: "PartyA",
        placeholder: "例如：GGV Capital",
        type: "text"
    },
    {
        id: "partyB",
        label: "乙方 (现有股东)",
        tag: "PartyB",
        placeholder: "例如：张三",
        type: "text"
    },
    {
        id: "capital",
        label: "注册资本",
        tag: "RegCapital",
        placeholder: "例如：100万元",
        type: "text"
    },
    {
        id: "shareRatio",
        label: "持股比例",
        tag: "ShareRatio",
        placeholder: "例如：10%",
        type: "text"
    },
    {
        id: "boardSeats",
        label: "董事会席位",
        tag: "BoardSeats",
        placeholder: "例如：1",
        type: "number"
    },
    {
        id: "signDate",
        label: "签署日期",
        tag: "SignDate",
        placeholder: "",
        type: "date"
    }
];

