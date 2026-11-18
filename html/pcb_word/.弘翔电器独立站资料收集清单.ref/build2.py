# -*- coding: utf-8 -*-
"""生成「弘翔电器 PCBA 独立站 · 资料收集清单」（v2：只提示收集什么，不讲来源）"""
import math

try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.properties import PageSetupProperties

def xl_color(h: str) -> str:
    v = h.removeprefix("#").upper()
    if len(v) != 6:
        raise ValueError(h)
    return "FF" + v

XL_PRIMARY = xl_color("#1F4E79")
XL_HEADER = xl_color("#2E6BA8")
XL_SUB = xl_color("#5B7FA6")
XL_INPUT = xl_color("#F4F8FE")
XL_STRIPE = xl_color("#FAFBFD")
XL_BORDER = xl_color("#D6DEE8")
XL_GREEN_BG = xl_color("#C6EFCE")
XL_GREEN_FG = xl_color("#006100")
XL_GREY_FG = xl_color("#8C8C8C")
XL_RED_FG = xl_color("#C00000")

F = "微软雅黑"
thin = Side(style="thin", color=XL_BORDER)
BD = Border(left=thin, right=thin, top=thin, bottom=thin)

# ============================================================
# 主清单：批次 / 收集项 / 为什么需要 / 合格标准 / 常见坑 / 优先级
# ============================================================
ITEMS = [
    # ---------- A 先决项 ----------
    ("A 先决项", "正式域名",
     "全站多语言地址、企业邮箱后缀、Google 收录都挂在域名上。不定下来，邮箱、hreflang、分享图全都做不了；定了再改要重来一遍。",
     "已注册、能解析、与品牌拼音一致（如 hongxiang-pcba.com）",
     "别用免费二级域名，也别选冗长难拼的名字——客户要靠它记住你", "高"),
    ("A 先决项", "企业域名邮箱",
     "B2B 询盘最忌 gmail / qq 邮箱：客户会怀疑你是不是贸易公司，邮件也容易进垃圾箱。",
     "sales@公司域名 与 inquiry@公司域名 各一个，附件上限 ≥ 50MB",
     "邮箱后缀与域名不一致，或用个人邮箱，等于主动丢掉信任分", "高"),
    ("A 先决项", "对外统一英文名",
     "9 个外语版本、报价单、合同、报关资料都用它。中途改名，网站和所有印刷资料都要重做。",
     "一次定死，全公司统一；与营业执照英文名称或海关备案名称一致",
     "网站一个写法、PI 另一个写法，客户会直接产生疑问", "高"),
    ("A 先决项", "网站上线主体",
     "决定页脚版权行、隐私政策主体、收款方。主体不清，欧洲客户会直接质疑合规性。",
     "与营业执照、收款账户三者同名",
     "网站写 A 公司、收款是 B 公司，是最典型的诈骗信号", "高"),

    # ---------- B 公司事实 ----------
    ("B 公司事实", "营业执照",
     "成立年份、公司名称、统一社会信用代码都从它来。海外客户会用天眼查、海关数据、邓白氏核对你写的每一个字。",
     "最新版原件彩色扫描 PDF，四角完整、无遮挡",
     "注册地在写字楼、实际工厂在别处时，网站要写实际生产地址", "高"),
    ("B 公司事实", "实际生产地址（中 + 英）",
     "客户会拿去 Google Maps 定位、算物流、安排审厂。只写「广东东莞」会让人怀疑你有没有真工厂。",
     "精确到门牌号；英文按邮政标准翻译，含邮编",
     "地址写错一个字，样品寄不到，客户第一印象就没了", "高"),
    ("B 公司事实", "厂房性质与面积",
     "「自有厂房」是 B2B 客户眼里交期稳定的硬证据，比任何形容词都管用。",
     "自有 → 产权证明；租赁 → 合同剩余年限。面积精确到百平方米",
     "虚报面积在审厂现场一票否决，且会永久失去该客户", "高"),
    ("B 公司事实", "员工总数 / 工程师人数",
     "「280 人，其中 35 名工程师」比「我们技术实力强」有说服力一百倍，也是客户评估产能的锚点。",
     "当月在册人数；分别给出总数与工程、品质人员数",
     "与社保人数、平台填写数差异过大，会被判定为信息不实", "中"),
    ("B 公司事实", "成立年份与主营沿革",
     "年限是 PCBA 行业最直观的信任指标，客户默认「做得久 = 工艺稳」。",
     "以营业执照为准；主营变更过的话说明何时切入 PCBA",
     "把老板其他公司的年限算进来，一旦被追问会很尴尬", "高"),
    ("B 公司事实", "联系电话（固话 + 手机）",
     "页脚和联系页的核心转化入口。海外客户打不通，就直接去找下一家。",
     "带 +86 国际区号；工作时间有人接听",
     "只留固话、无人值守或号码已停用，是最常见的转化杀手", "高"),
    ("B 公司事实", "WhatsApp / WeChat 号码",
     "海外 B2B 询盘首选 WhatsApp，尤其印度、中东、南美客户。没有它等于放弃一半询盘。",
     "真实开通、头像用公司 LOGO、号码含 +86",
     "给了号码却没人看消息，比不给更伤", "高"),
    ("B 公司事实", "询盘接收邮箱",
     "表单提交后往哪发。错一个字母，询盘就石沉大海。",
     "至少两个（主收 + 主管抄送），确认能收 20MB 以上附件",
     "用长期不登录的邮箱，客户等 3 天没回复就走了", "高"),
    ("B 公司事实", "工作时间与时区",
     "客户据此判断多久能等到回复，也是「48 小时响应」这类承诺的前提。",
     "明确到时段与时区，例：Mon–Sat 9:00–18:00 (GMT+8)",
     "写 7×24 却做不到，第一次超时回复就失去信任", "中"),
    ("B 公司事实", "收款主体与付款方式",
     "决定报价单与 PI 的抬头，客户付款前一定会核对。",
     "收款账户名 = 营业执照名；说明接受的付款方式与币种",
     "收款方与网站公司名不一致，客户会直接判定为风险交易", "中"),

    # ---------- C 产能与工艺 ----------
    ("C 产能工艺", "SMT 产线数量与线体配置",
     "产线数是客户判断你能否承接他订单量的第一指标，也是「交期稳定」的依据。",
     "几条线、每条线的贴片机品牌与型号、是否含双轨",
     "把闲置或已退役的线算进去，客户数设备时会当场拆穿", "高"),
    ("C 产能工艺", "DIP 插件 / 波峰焊线数量",
     "很多板卡必须过波峰焊，这是判断是否真具备全流程能力的关键。",
     "条数、是否含选择性波峰焊",
     "只写 SMT 不写 DIP，客户会认为你只能做单面贴装", "高"),
    ("C 产能工艺", "日 / 月贴装点数与产能",
     "让客户自己算出「我的量在你这里排得上号吗」，减少无效询盘。",
     "以实际月度产出反推，注明是理论值还是实测值",
     "按设备极限值报，实际排产做不出来，交付时必然扯皮", "高"),
    ("C 产能工艺", "关键工艺参数",
     "工程师拿这些参数先做一轮自我筛选：不匹配的不会来询盘，匹配的则直接建立专业信任。",
     "层数、最小贴装封装、贴装精度（mm）、最大板幅（mm）、最小线宽线距",
     "写实验室极限值而非量产稳定值，打样能做、量产做不了", "高"),
    ("C 产能工艺", "设备清单（品牌 + 型号 + 数量）",
     "海外工程师认设备：贴片机、回流焊、AOI、X-Ray 的品牌型号，比「设备先进」有说服力得多。",
     "列表：类别、品牌、型号、数量、投产年份",
     "不列设备清单，客户会默认你没有，或不敢公开", "高"),
    ("C 产能工艺", "检测能力清单",
     "PCBA 客户最怕批量不良。检测设备清单直接回答「你怎么保证质量」。",
     "AOI / SPI / X-Ray / ICT / FCT / 老化各自数量 + 覆盖哪些工序 + 是否全检",
     "有 X-Ray 却没写，等于白买了这台设备带来的信任", "高"),
    ("C 产能工艺", "良率 / 一次交验通过率",
     "最有杀伤力的质量数字，但必须有统计口径，否则等于自曝其短。",
     "近 12 个月实测值 + 统计口径（如 FQC 一次通过率）",
     "拍脑袋写 99.9%，客户后续要报表时拿不出来", "高"),
    ("C 产能工艺", "交期承诺（分档）",
     "客户最常问的第一个问题。写清楚能显著减少来回沟通，也能筛掉不匹配的客户。",
     "按打样 / 小批量 / 量产三档给天数，注明前提是 BOM 齐套",
     "承诺 24 小时却受制于物料采购，第一次违约就难有第二次", "高"),
    ("C 产能工艺", "年交付规模",
     "规模数字是首页数据条的核心，用来快速建立「这家做得很大」的印象。",
     "近 12 个月实际数据，可用区间表述（如 500 万–800 万片）",
     "把历年累计说成年交付量，是最容易被识破的夸大", "高"),
    ("C 产能工艺", "出口市场与行业分布",
     "客户看到「你做过德国工业客户」，会觉得你懂他们的标准与沟通方式。",
     "前 5 个国家或地区 + 占比；主要行业与大致比例",
     "客户名单未经授权不要公开，写行业分布已足够", "中"),
    ("C 产能工艺", "来料加工 / 客供料政策",
     "很多客户已有自己的料，最关心你接不接受来料加工。这是高频问题，值得单独写清。",
     "说明 Turnkey、客供 PCB、客供料、部分客供四种模式分别接不接",
     "页面只写「一站式交钥匙」，会流失一批自带物料的优质客户", "中"),

    # ---------- D 证书与质量证据 ----------
    ("D 证书证据", "ISO 9001 证书",
     "体系认证的底线。没有它，多数欧美客户不会进入下一轮沟通。",
     "扫描 PDF + 发证机构 + 证书编号 + 有效期 + 认证范围",
     "认证范围只覆盖贸易不含生产，被客户发现会直接出局", "高"),
    ("D 证书证据", "IATF 16949（汽车）",
     "进入汽车电子供应链的门票。没有它就别在汽车电子板块出现，否则是硬伤。",
     "扫描件 + 编号 + 有效期 + 适用范围",
     "宣传里写「符合 IATF 要求」却无证书，客户一查就穿帮", "高"),
    ("D 证书证据", "ISO 13485（医疗）",
     "医疗客户的准入条件。有它就是差异化优势，没有就别写医疗行业。",
     "扫描件 + 编号 + 有效期 + 适用范围",
     "拿 ISO 9001 冒充医疗体系认证，属严重误导", "高"),
    ("D 证书证据", "UL / RoHS / REACH 符合性",
     "出口欧美的合规底线，客户合规部门会逐项索要。",
     "UL 档案号或符合性声明；RoHS / REACH 检测报告（注意有效期）",
     "检测报告过期仍在用，属提供虚假材料", "高"),
    ("D 证书证据", "ISO 14001 / 高新技术企业 / 其他",
     "锦上添花但真实有效：欧洲客户看重环保体系，高企在国内客户中加分。",
     "扫描件 + 有效期",
     "过期证书继续挂着，是网站最严重的信任风险之一", "中"),
    ("D 证书证据", "IPC-A-610 执行等级",
     "工艺水平的国际通用语言。写明 Class 2 还是 Class 3，工程师一眼就懂你的水准。",
     "明确可执行等级 + 检验依据",
     "写了 Class 3 却按 Class 2 检验，验厂时会暴露", "中"),
    ("D 证书证据", "质量流程文件",
     "审厂前客户必看。能拿出流程图，说明你的质量不是靠老师傅手感。",
     "IQC / IPQC / OQC / 追溯 / 8D 流程图 + 关键记录样张（可脱敏）",
     "流程文件里有其他客户的名称与料号，属于泄密", "中"),
    ("D 证书证据", "证书有效期台账",
     "一次性建立，避免证书过期仍挂在网站上——这是最容易被竞争对手举报的点。",
     "列表：证书名、编号、有效期、换证提醒时间",
     "到期前 3 个月未启动换证，网站会被迫临时下架证书内容", "中"),

    # ---------- E 视觉素材 ----------
    ("E 视觉素材", "公司 LOGO",
     "站点目前还没有 LOGO。没有 LOGO 的 B2B 网站会被当成未完成的临时页面。",
     "横版、竖版、反白（深色底用）三版；SVG / AI 矢量 + PNG 透明底",
     "只给 JPG 白底图，放到深色页头会出现难看的白块", "高"),
    ("E 视觉素材", "产线与工厂实拍 7 张",
     "海外客户无法来厂，照片就是他的验厂依据。用网图或渲染图，一次视频会议就会露馅。",
     "真实拍摄、横构图、长边 ≥ 2000px，按《拍摄脚本》命名",
     "用网上下载的工厂图，可能被索赔，且客户反向搜图即可识破", "高"),
    ("E 视觉素材", "典型 PCBA 产品图 3–5 张",
     "客户先看你做过什么板，才判断你有没有同类经验。",
     "自有产品、白底、单板平铺、45° 微俯拍",
     "把客户定制板拍照放上去，等于违反保密协议", "高"),
    ("E 视觉素材", "团队 / 工程师工作照",
     "让人相信屏幕背后是真实的工厂和工程师，而不是贸易中间商。",
     "真实工作场景 1–2 张，取得出镜者书面授权",
     "使用网图「工程师」或未经授权的员工照片，存在肖像权风险", "中"),
    ("E 视觉素材", "工厂 / 产线视频",
     "首页放视频可显著提升停留时长，对首次来访的海外客户说服力最强。",
     "横版 MP4，1080P，含产线运行、检测环节、厂区外观",
     "竖版或抖动的手机视频，反而显得不专业", "低"),
    ("E 视觉素材", "图片版权与肖像授权确认",
     "不是收集素材，而是确认素材能不能用。省了这步，上线后可能被迫全部撤图。",
     "自有拍摄默认可用；员工出镜需签字；第三方素材需购买凭证",
     "默认「网上找的图都能用」，是最常见的法律雷区", "高"),

    # ---------- F 说服力内容 ----------
    ("F 说服力内容", "客户案例 3 个",
     "海外客户真正关心的只有一件事：你做过和我类似的吗。三个案例胜过十页能力介绍。",
     "行业 + 客户痛点 + 做了什么 + 可量化结果；无授权时写成「某德系工业客户」",
     "只写「服务众多世界 500 强」而无任何细节，客户会当成空话", "高"),
    ("F 说服力内容", "核心产品 / 板卡目录",
     "让客户 30 秒内判断你对不对口，也顺便过滤掉不合适询盘。",
     "10–20 项：名称、应用、关键参数、层数、典型批量",
     "参数写错会被工程师当场质疑专业性", "中"),
    ("F 说服力内容", "差异化能力 3–5 条",
     "回答「为什么选你而不是隔壁那家」。这是整个网站最该认真想的部分。",
     "具体、可验证、别人不易抄的，例如 X-Ray 全检、48 小时打样、BOM 降本 8–15%",
     "写「品质优良、价格合理、服务周到」，等于什么都没说", "高"),
    ("F 说服力内容", "真实 FAQ 8–10 条",
     "客户反复问的就是他最在意的。用真实提问做 FAQ，转化率远高于自编的。",
     "近半年客户高频问题 + 你们的真实答复",
     "自编 FAQ 常答非所问，客户看完依然要发邮件确认", "中"),
    ("F 说服力内容", "合作流程与交付节点",
     "把「发资料 → 评估 → 打样 → 测试 → 量产」讲清楚，降低客户对未知的顾虑。",
     "每步的输入、输出、耗时",
     "流程写得很美好但实际执行不同，合作中会产生落差", "中"),
    ("F 说服力内容", "付款方式与贸易条款",
     "海外客户必问。提前写清，可以过滤掉账期谈不拢的客户。",
     "T/T 定金比例、是否接受信用证、FOB / EXW / DDP 各自条件",
     "含糊其辞会让客户下单前反复确认，拖长成交周期", "中"),
    ("F 说服力内容", "样品政策",
     "客户迈出第一步的成本。写清楚能明显提高打样询盘量。",
     "样品是否收费、能否抵扣批量货款、运费谁承担",
     "样品免费但运费规则不明，容易在小额订单上产生纠纷", "中"),
    ("F 说服力内容", "NDA 模板",
     "工程师最在意图纸安全。提供 NDA 下载，是把「我们重视保密」从口号变成证据。",
     "中英文版 PDF，含双方主体、保密范围、期限、违约责任",
     "只说「可签 NDA」却拿不出模板，客户会认为流程不规范", "中"),

    # ---------- G 上线准备 ----------
    ("G 上线准备", "隐私政策与服务条款",
     "页脚已有这两个链接但无内容。欧盟客户会真去看，Google 收录也看这个。",
     "可由我按 GDPR 起草一版，你确认公司主体、收集字段与用途即可",
     "直接复制同行的隐私政策，条款里出现别家公司名称", "高"),
    ("G 上线准备", "Cookie / GDPR 提示",
     "面向欧洲访客的合规要求，缺了可能被投诉甚至罚款。",
     "确认是否需要 Cookie 同意条；面向欧洲则默认加上",
     "按国内站心态做外贸而忽视 GDPR，会留下长期隐患", "中"),
    ("G 上线准备", "9 种外语的审校安排",
     "机器翻译在德语、俄语、阿拉伯语的工业术语上容易出错，而最先看到的恰恰是最专业的那批客户。",
     "安排海外销售 / 代理抽校关键页面，或由我按行业术语库人工审校",
     "完全不校就上线，专业买家看到 Chinglish 会直接关闭页面", "中"),
    ("G 上线准备", "统计代码与转化追踪",
     "上线后才能知道访客来自哪些国家、询盘从哪个页面来，否则推广就是盲投。",
     "GA4 衡量 ID，或告诉我用哪种统计工具",
     "上线后才想起加统计，会永久丢失前期流量数据", "中"),
    ("G 上线准备", "服务器与访问速度",
     "决定海外打开速度。放在国内需备案、海外访问慢；放境外刚好相反。",
     "明确目标市场：欧美为主建议香港或新加坡节点",
     "没想清楚就上线，海外打开要 5 秒以上，客户直接关掉", "中"),
    ("G 上线准备", "上线前事实核对",
     "网站上的每个数字、每张证书、每张图都要能拿出证据。上线前统一核一遍。",
     "逐条过一遍《上线前自查》，全部为「是」再上线",
     "上线后被客户或竞争对手指出造假，修复成本远高于事前核对", "高"),
]

SHOTS = [
    ("1", "SMT 车间全景", "首页首屏下方", "横向，站在产线一端向纵深拍，能看出产线长度与设备排列",
     "上午自然光 + 车间照明全开，避免单侧强光造成过曝", "地面无杂物、物料箱归位、员工不拍正脸", "hero-smt.jpg", "必拍"),
    ("2", "工厂外景 / 大门", "关于页首屏", "斜 45° 拍厂区建筑，尽量带上公司招牌",
     "白天晴天，顺光", "招牌完整可读；避免拍到路边杂物与无关车辆", "factory-exterior.jpg", "必拍"),
    ("3", "SMT 产线特写", "关于页大图", "对准贴片机工作段，最好拍到吸嘴正在贴装",
     "车间照明，快门别太慢以免拖影", "可露设备品牌铭牌；不要拍到操作屏上的客户料号", "line-smt.jpg", "必拍"),
    ("4", "AOI 检测画面", "关于页", "镜头对着检测屏幕，屏幕上的板子图像要清晰",
     "关掉附近强光源，避免屏幕反光", "屏幕内容须为自有板卡，不要出现客户型号", "aoi.jpg", "必拍"),
    ("5", "老化测试区", "关于页", "老化架通电状态，指示灯亮起",
     "室内，可补光", "线缆收纳整齐，不要露出插线板乱线", "burn-in.jpg", "必拍"),
    ("6", "仓储与物料", "关于页", "料架全景，能看出分区与标识",
     "室内均匀照明", "拍到条码标签与防潮柜最佳；不要拍到客户料号", "warehouse.jpg", "必拍"),
    ("7", "工程实验室 / 调试", "关于页", "工程师看原理图或调试板卡的真实瞬间",
     "自然光优先", "需取得出镜者书面授权；屏幕内容脱敏", "lab.jpg", "必拍"),
    ("8", "典型 PCBA 产品图", "服务页 / 首页", "白底或浅灰底，单板平铺，可 45° 微俯拍",
     "两侧柔光，避免反光与阴影", "必须为自有产品，不得出现客户丝印与品牌 LOGO", "product-01~05.jpg", "建议"),
    ("9", "焊点 / 器件微距", "首页能力区", "微距拍 BGA 或 01005 器件，体现工艺精度",
     "微距专用光，或用显微镜成像", "确保画面里没有可识别的客户 LOGO", "micro-01.jpg", "建议"),
    ("10", "团队工作照", "关于页", "真实工作状态优于摆拍合影",
     "自然光，避免顶光造成眼影", "出镜员工需签字授权", "team-01.jpg", "建议"),
    ("11", "证书 / 报告实拍", "认证区（可选）", "证书平铺拍摄，或直接用扫描件",
     "光线均匀，四角完整", "不要露出证书编号以外的敏感信息", "cert-*.jpg", "可选"),
    ("12", "社媒分享图", "分享预览", "首屏视觉 + 公司名 + 一句话卖点",
     "—", "无图时我可用代码生成，不必单独拍", "og-cover.jpg", "可选"),
]

CHECKS = [
    ("证书可出示", "页面上写的每一个认证，都能立刻拿出证书扫描件"),
    ("证书在有效期内", "所有展示的证书均未过期，且已记录换证提醒"),
    ("认证范围匹配", "认证范围覆盖 PCBA 组装，而不只是贸易或设计"),
    ("数字有出处", "面积、产线数、客户数、良率等每个数字都能说出来源"),
    ("产能表述保守", "承诺的产能与交期是量产稳定值，不是极限值"),
    ("无绝对化用语", "未出现「最好」「第一」「最大」等无依据表述"),
    ("联系方式可接通", "电话、WhatsApp、邮箱均已实测可正常联系"),
    ("邮箱为企业域名", "对外邮箱后缀与公司域名一致"),
    ("图片版权清晰", "所有图片为自拍或已获授权，无网络下载图"),
    ("肖像已授权", "出镜员工的照片已取得书面授权"),
    ("无客户信息泄露", "图片与案例中不出现客户名称、LOGO、料号"),
    ("案例已获准公开", "引用客户案例前已取得同意，或已做匿名处理"),
    ("多语言已审校", "关键页面外语版本经母语者或专业人员看过"),
    ("阿拉伯语显示正常", "阿语页面整体右对齐，无内容错位"),
    ("表单可收附件", "询盘表单实测可上传 Gerber / BOM，收件邮箱能收到"),
    ("移动端可用", "手机上页面不横向滚动，菜单与表单正常"),
    ("隐私政策已上线", "隐私政策与服务条款有真实内容，主体为公司本身"),
    ("主体信息一致", "网站公司名、收款方、营业执照三者一致"),
    ("加载速度可接受", "目标市场打开首页不超过 3 秒"),
    ("统计已接入", "已安装统计代码，能看到访客来源与询盘转化"),
]

BATCHES = ["A 先决项", "B 公司事实", "C 产能工艺", "D 证书证据",
           "E 视觉素材", "F 说服力内容", "G 上线准备"]
BATCH_NOTE = [
    "这 4 项不定，后面所有内容都可能返工，务必最先定",
    "客户会拿去核实的信息，必须真实",
    "海外工程师真正看的部分，需要生产与品质配合",
    "多数是现成的，重点是找出扫描件并核对有效期",
    "唯一需要安排拍摄的部分，建议提前约时间",
    "不是找资料，而是做判断：我们到底强在哪",
    "可与我并行推进，你确认主体，其余我来起草",
]

wb = Workbook()

def title_block(ws, ncols, title, subtitle, h=42, sh=34):
    last = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last}1")
    c = ws["A1"]; c.value = title
    c.font = Font(name=F, size=16, bold=True, color="FFFFFFFF")
    c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = h
    ws.merge_cells(f"A2:{last}2")
    c2 = ws["A2"]; c2.value = subtitle
    c2.font = Font(name=F, size=10, color=XL_SUB)
    c2.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[2].height = sh


def est_height(pairs, minimum=30):
    lines = 1
    for text, width in pairs:
        if not text:
            continue
        cap = max(8, int(width / 2.05))
        lines = max(lines, math.ceil(len(str(text)) / cap))
    return max(minimum, min(lines * 15.0 + 8, 132))


def header_row(ws, headers, widths, row=3):
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=F, size=10.5, bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor=XL_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BD
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30


def fill_rows(ws, rows, widths, start=4, wrap_cols=(), center_cols=(1,), input_col=None):
    for idx, row in enumerate(rows):
        r = start + idx
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = Font(name=F, size=10, color=XL_SUB) if c == 1 else Font(name=F, size=10)
            cell.border = BD
            halign = "center" if c in center_cols else "left"
            cell.alignment = Alignment(horizontal=halign, vertical="center",
                                       wrap_text=c in wrap_cols, indent=0 if halign == "center" else 1)
            if idx % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=XL_STRIPE)
        if input_col:
            ws.cell(row=r, column=input_col).fill = PatternFill("solid", fgColor=XL_INPUT)
        ws.row_dimensions[r].height = est_height([(row[c - 1], widths[c - 1]) for c in wrap_cols])


def dv_list(ws, rng, options):
    d = DataValidation(type="list", formula1='"%s"' % ",".join(options), allow_blank=True, showDropDown=False)
    ws.add_data_validation(d)
    d.add(rng)


def status_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"已收集"'],
        fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(name=F, size=10, color=XL_GREEN_FG)))
    for k in ("暂无", "不需要"):
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal", formula=[f'"{k}"'], font=Font(name=F, size=10, color=XL_GREY_FG)))


def page_setup(ws, last_row, ncols):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:3"
    ws.print_area = f"A1:{get_column_letter(ncols)}{last_row}"


# ==================== Sheet 1：先看这里 ====================
ws0 = wb.active
ws0.title = "先看这里"
ws0.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [8, 30, 46, 30, 24]):
    ws0.column_dimensions[col].width = w

ws0.merge_cells("A1:E1")
c = ws0["A1"]; c.value = "弘翔电器 PCBA 独立站 · 资料收集清单"
c.font = Font(name=F, size=20, bold=True, color="FFFFFFFF")
c.fill = PatternFill("solid", fgColor=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws0.row_dimensions[1].height = 56

ws0.merge_cells("A2:E2")
c = ws0["A2"]
c.value = ("下面每一条都写清了「要收集什么、为什么需要、什么算合格、容易踩什么坑」。"
           "按顺序过一遍，收集到的内容填在最右侧一列；状态改成「已收集」后会自动变绿。")
c.font = Font(name=F, size=10.5, color=XL_SUB)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
ws0.row_dimensions[2].height = 38

ws0.merge_cells("A4:E4")
c = ws0["A4"]; c.value = "建议的收集顺序（按批次推进，不要跳）"
c.font = Font(name=F, size=12, bold=True, color=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

r = 5
for i, (b, note) in enumerate(zip(BATCHES, BATCH_NOTE)):
    cnt = sum(1 for x in ITEMS if x[0] == b)
    cc = ws0.cell(row=r, column=1, value=b[0])
    cc.font = Font(name=F, size=12, bold=True, color="FFFFFFFF")
    cc.fill = PatternFill("solid", fgColor=XL_HEADER)
    cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.border = BD
    ws0.cell(row=r, column=2, value=f"{b[2:]}（{cnt} 项）").font = Font(name=F, size=10.5, bold=True)
    ws0.cell(row=r, column=3, value=note).font = Font(name=F, size=10)
    for col in (2, 3):
        ws0.cell(row=r, column=col).border = BD
        ws0.cell(row=r, column=col).alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws0.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    ws0.cell(row=r, column=4).border = BD
    ws0.row_dimensions[r].height = 30
    r += 1

r += 1
ws0.merge_cells(f"A{r}:E{r}")
c = ws0.cell(row=r, column=1, value="三条红线（比收集本身更重要）")
c.font = Font(name=F, size=12, bold=True, color=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
r += 1
RED = [
    ("不写没有证书支撑的认证", "写了 ISO 13485 却拿不出证书，是最严重的信任崩塌，也可能被举报"),
    ("不报做不到的产能与交期", "客户按你的承诺排产，第一次违约基本就没有第二次合作"),
    ("不用来源不明的图片", "网图可能侵权被索赔，客户反向搜图就能识破「工厂不是你的」"),
]
for i, (t, d) in enumerate(RED, start=1):
    cc = ws0.cell(row=r, column=1, value=i)
    cc.font = Font(name=F, size=11, bold=True, color=XL_RED_FG)
    cc.fill = PatternFill("solid", fgColor=xl_color("#FDE9E9"))
    cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.border = BD
    ws0.cell(row=r, column=2, value=t).font = Font(name=F, size=10, bold=True)
    ws0.cell(row=r, column=3, value=d).font = Font(name=F, size=10, color=XL_SUB)
    for col in (2, 3):
        ws0.cell(row=r, column=col).border = BD
        ws0.cell(row=r, column=col).alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws0.row_dimensions[r].height = 30
    r += 1

r += 1
ws0.merge_cells(f"A{r}:E{r}")
c = ws0.cell(row=r, column=1, value="进度总览（自动统计，不用手动改）")
c.font = Font(name=F, size=12, bold=True, color=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
r += 1
for i, h in enumerate(["", "指标", "数值", "说明", ""], start=1):
    cc = ws0.cell(row=r, column=i, value=h)
    cc.font = Font(name=F, size=10, bold=True, color="FFFFFFFF")
    cc.fill = PatternFill("solid", fgColor=XL_HEADER)
    cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.border = BD
ws0.row_dimensions[r].height = 26
r += 1
STAT_START = r
STAT_KEYS = ["需要收集的总项数", "已收集", "确认拿不到（暂无）", "收集完成率", "高优先级中未完成的"]
for i, name in enumerate(STAT_KEYS):
    rr = STAT_START + i
    ws0.cell(row=rr, column=1, value="").border = BD
    cc = ws0.cell(row=rr, column=2, value=name)
    cc.font = Font(name=F, size=10)
    cc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    cc.border = BD
    cv = ws0.cell(row=rr, column=3)
    cv.font = Font(name=F, size=11, bold=True, color=XL_PRIMARY)
    cv.alignment = Alignment(horizontal="center", vertical="center")
    cv.border = BD
    ws0.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=5)
    ws0.cell(row=rr, column=4).border = BD
    ws0.row_dimensions[rr].height = 26
STAT_END = STAT_START + len(STAT_KEYS) - 1
ws0.cell(row=STAT_START + 3, column=3).number_format = "0.0%"

r = STAT_END + 2
ws0.merge_cells(f"A{r}:E{r}")
c = ws0.cell(row=r, column=1)
c.value = ("拿不准要不要收集的，就先收集——多出来的我可以不放进页面，但缺了的会让网站显得空洞。"
           "填完连同证书、图片一起回传，我做全站 10 语言替换与一致性校验。")
c.font = Font(name=F, size=10, italic=True, color=XL_PRIMARY)
c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
ws0.row_dimensions[r].height = 34
LAST0 = r

# ==================== Sheet 2：收集清单 ====================
ws1 = wb.create_sheet("收集清单")
ws1.sheet_properties.tabColor = xl_color("#2E6BA8")
ws1.sheet_view.showGridLines = False
H1 = ["序号", "批次", "收集项", "为什么需要（不收集会怎样）", "合格标准 · 交付形式", "常见坑", "优先级", "状态", "收集到的内容 / 文件名"]
W1 = [5, 12, 24, 48, 38, 36, 8, 10, 28]
title_block(ws1, len(H1), "资料收集清单",
            "第 4 列是这份表的核心：先读懂「为什么需要」和「常见坑」，再决定怎么收集。"
            "收集到的内容填在最右列，状态改成「已收集」会自动变绿。")
header_row(ws1, H1, W1)
rows1 = [[i, b, item, why, std, pit, prio, "待收集", ""] for i, (b, item, why, std, pit, prio) in enumerate(ITEMS, start=1)]
fill_rows(ws1, rows1, W1, start=4, wrap_cols=(3, 4, 5, 6, 9), center_cols=(1, 2, 7, 8), input_col=9)
END1 = 4 + len(rows1) - 1
dv_list(ws1, f"H4:H{END1}", ["待收集", "已收集", "暂无"])
dv_list(ws1, f"G4:G{END1}", ["高", "中", "低"])
status_cf(ws1, f"H4:H{END1}")
ws1.conditional_formatting.add(f"G4:G{END1}", CellIsRule(
    operator="equal", formula=['"高"'], font=Font(name=F, size=10, bold=True, color=XL_RED_FG)))
ws1.auto_filter.ref = f"A3:I{END1}"
ws1.freeze_panes = "D4"
page_setup(ws1, END1, len(H1))

# ==================== Sheet 3：拍摄脚本 ====================
ws2 = wb.create_sheet("拍摄脚本")
ws2.sheet_properties.tabColor = xl_color("#2E6BA8")
ws2.sheet_view.showGridLines = False
H2 = ["编号", "拍什么", "用在哪个位置", "机位与构图", "光线要求", "必须避开的内容", "文件命名", "必要性"]
W2 = [6, 22, 20, 40, 26, 34, 20, 9]
title_block(ws2, len(H2), "工厂拍摄脚本（交给拍照的人照着做）",
            "手机横持、开最高分辨率即可，但请按「机位与构图」执行。1–7 是站点已预留的图片位，缺图会显示灰色占位块。")
header_row(ws2, H2, W2)
fill_rows(ws2, SHOTS, W2, start=4, wrap_cols=(2, 3, 4, 5, 6), center_cols=(1, 8))
END2 = 4 + len(SHOTS) - 1
ws2.conditional_formatting.add(f"H4:H{END2}", CellIsRule(
    operator="equal", formula=['"必拍"'],
    fill=PatternFill("solid", fgColor=xl_color("#FFF2CC")), font=Font(name=F, size=10, bold=True)))
ws2.auto_filter.ref = f"A3:H{END2}"
ws2.freeze_panes = "C4"
page_setup(ws2, END2, len(H2))

# ==================== Sheet 4：批次进度 ====================
ws3 = wb.create_sheet("批次进度")
ws3.sheet_properties.tabColor = xl_color("#5B7FA6")
ws3.sheet_view.showGridLines = False
H3 = ["批次", "事项数", "已收集", "完成率", "说明"]
W3 = [18, 12, 12, 12, 52]
title_block(ws3, len(H3), "各批次进度（自动统计）",
            "按批次看进度，方便知道卡在哪一步。数值来自《收集清单》，不用手动填。")
header_row(ws3, H3, W3)
rows3 = [[b, None, None, None, note] for b, note in zip(BATCHES, BATCH_NOTE)]
fill_rows(ws3, rows3, W3, start=4, wrap_cols=(5,), center_cols=(1, 2, 3, 4))
END3 = 4 + len(rows3) - 1
for i, b in enumerate(BATCHES):
    r = 4 + i
    ws3.cell(row=r, column=2, value=f'=COUNTIF(收集清单!B4:B{END1},A{r})')
    ws3.cell(row=r, column=3, value=f'=COUNTIFS(收集清单!B4:B{END1},A{r},收集清单!H4:H{END1},"已收集")')
    ws3.cell(row=r, column=4, value=f'=IF(OR(B{r}="",B{r}=0),"",C{r}/B{r})')
    ws3.cell(row=r, column=4).number_format = "0.0%"
ws3.conditional_formatting.add(f"D4:D{END3}", CellIsRule(
    operator="greaterThanOrEqual", formula=["1"],
    fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(name=F, size=10, color=XL_GREEN_FG)))
ws3.freeze_panes = "A4"
page_setup(ws3, END3, len(H3))

# ==================== Sheet 5：上线前自查 ====================
ws4 = wb.create_sheet("上线前自查")
ws4.sheet_properties.tabColor = xl_color("#5B7FA6")
ws4.sheet_view.showGridLines = False
H4 = ["序号", "检查项", "具体要确认什么", "结论", "备注"]
W4 = [6, 24, 56, 12, 26]
title_block(ws4, len(H4), "上线前事实核对（全部为「是」再上线）",
            "网站上的每个数字、每张证书、每张图都要能拿出证据。上线后被指出问题，修复成本远高于事前核对。")
header_row(ws4, H4, W4)
rows4 = [[i, name, desc, "待确认", ""] for i, (name, desc) in enumerate(CHECKS, start=1)]
fill_rows(ws4, rows4, W4, start=4, wrap_cols=(2, 3, 5), center_cols=(1, 4), input_col=5)
END4 = 4 + len(rows4) - 1
dv_list(ws4, f"D4:D{END4}", ["是", "否", "待确认", "不适用"])
ws4.conditional_formatting.add(f"D4:D{END4}", CellIsRule(
    operator="equal", formula=['"是"'],
    fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(name=F, size=10, color=XL_GREEN_FG)))
ws4.conditional_formatting.add(f"D4:D{END4}", CellIsRule(
    operator="equal", formula=['"否"'],
    fill=PatternFill("solid", fgColor=xl_color("#FFC7CE")), font=Font(name=F, size=10, bold=True, color=XL_RED_FG)))
ws4.auto_filter.ref = f"A3:E{END4}"
ws4.freeze_panes = "C4"

SUM4 = END4 + 2
ws4.cell(row=SUM4, column=2, value="「是」的项数").font = Font(name=F, size=10, bold=True)
ws4.cell(row=SUM4, column=3, value=f'=COUNTIF(D4:D{END4},"是")')
ws4.cell(row=SUM4 + 1, column=2, value="「否」的项数").font = Font(name=F, size=10, bold=True)
ws4.cell(row=SUM4 + 1, column=3, value=f'=COUNTIF(D4:D{END4},"否")')
ws4.cell(row=SUM4 + 2, column=2, value="通过率").font = Font(name=F, size=10, bold=True)
ws4.cell(row=SUM4 + 2, column=3, value=f'=IF((B{SUM4}+B{SUM4+1})=0,"",B{SUM4}/(B{SUM4}+B{SUM4+1}))')
ws4.cell(row=SUM4 + 2, column=3).number_format = "0.0%"
for rr in range(SUM4, SUM4 + 3):
    ws4.cell(row=rr, column=3).font = Font(name=F, size=11, bold=True, color=XL_PRIMARY)
    ws4.cell(row=rr, column=3).alignment = Alignment(horizontal="center", vertical="center")
    ws4.cell(row=rr, column=2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
LAST4 = SUM4 + 2
page_setup(ws4, LAST4, len(H4))

# ==================== 回填 Sheet1 统计公式 ====================
NOTES = ["清单里需要处理的条目总数", "状态已改为「已收集」的项", "确认拿不到的，标「暂无」后我从页面移除",
         "已收集 ÷ 总数，100% 即齐活", "高优先级里还没拿到的，这些决定能否上线"]
fmts = [
    f"=COUNTA(收集清单!C4:C{END1})",
    f'=COUNTIF(收集清单!H4:H{END1},"已收集")',
    f'=COUNTIF(收集清单!H4:H{END1},"暂无")',
    f'=IF(OR(C{STAT_START}="",C{STAT_START}=0),"",C{STAT_START+1}/C{STAT_START})',
    f'=COUNTIFS(收集清单!G4:G{END1},"高",收集清单!H4:H{END1},"待收集")',
]
for i, f_ in enumerate(fmts):
    ws0.cell(row=STAT_START + i, column=3, value=f_)
for i, note in enumerate(NOTES):
    ws0.cell(row=STAT_START + i, column=4, value=note)
    ws0.cell(row=STAT_START + i, column=4).font = Font(name=F, size=9.5, color=XL_SUB)
    ws0.cell(row=STAT_START + i, column=4).alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
ws0.conditional_formatting.add(f"C{STAT_START+3}:C{STAT_START+3}", CellIsRule(
    operator="greaterThanOrEqual", formula=["1"],
    fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(name=F, size=11, color=XL_GREEN_FG)))
ws0.conditional_formatting.add(f"C{STAT_START+4}:C{STAT_START+4}", CellIsRule(
    operator="greaterThan", formula=["0"], font=Font(name=F, size=11, bold=True, color=XL_RED_FG)))

# ---------- 文档属性 ----------
wb.properties.title = "弘翔电器 PCBA 独立站资料收集清单"
wb.properties.subject = "PCBA 外贸独立站内容与素材收集"
wb.properties.creator = "WorkBuddy"
wb.calculation.fullCalcOnLoad = True
wb.active = 0

OUT = r"C:/Users/Administrator/WorkBuddy/2026-10-21-17-15-59/hongxiang-pcba/弘翔电器独立站资料收集清单.xlsx"
wb.save(OUT)


def inject(path, values):
    import zipfile, re, os
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", item.filename)
            if m and int(m.group(1)) in values:
                xml = data.decode("utf-8")
                for cell, val in values[int(m.group(1))].items():
                    pat = re.compile(r'(<c r="%s"[^>]*>)(<f>.*?</f>)<v />' % cell)
                    txt = repr(round(val, 10)) if isinstance(val, float) else str(int(val))
                    xml, n = pat.subn(lambda mo, t=txt: f"{mo.group(1)}{mo.group(2)}<v>{t}</v>", xml)
                    if n != 1:
                        raise RuntimeError(f"inject failed: {cell} ({n})")
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


n_item = len(ITEMS)
n_high = sum(1 for x in ITEMS if x[5] == "高")
vals0 = {
    f"C{STAT_START}": n_item, f"C{STAT_START+1}": 0, f"C{STAT_START+2}": 0,
    f"C{STAT_START+3}": 0.0, f"C{STAT_START+4}": n_high,
}
vals3 = {}
for i, b in enumerate(BATCHES):
    r = 4 + i
    cnt = sum(1 for x in ITEMS if x[0] == b)
    vals3[f"B{r}"] = cnt
    vals3[f"C{r}"] = 0
    vals3[f"D{r}"] = 0.0
vals4 = {f"C{SUM4}": 0, f"C{SUM4+1}": 0, f"C{SUM4+2}": 0.0}

inject(OUT, {1: vals0, 4: vals3, 5: vals4})
print("ITEMS", n_item, "HIGH", n_high)
print("BATCH", {b: sum(1 for x in ITEMS if x[0] == b) for b in BATCHES})
print("SAVED", OUT)
