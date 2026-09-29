# -*- coding: utf-8 -*-
"""生成润微岗位汇总 Excel（与简女士「发布职位·10」清单一致）。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

jobs = [
    dict(no="A56129", name="纸样主管/高级技术（外单优先）", line="技术/生产", salary="20-30K", exp="5-10年", edu="大专",
         tags="内衣/泳衣/家居服、服装、熟悉面/辅料市场、擅长ET、FOB批办",
         duty="版型研究与开发；竞争品牌版型及前沿资讯调研；负责设计款纸样及出品；负责版型的修改与完善、版型确认的跟进；负责大货生产可行性的分析；FOB批办跟进及大货生产中技术问题的解决；负责跟进试穿试洗测评及大货风险性；完成上级交代任务",
         req="5年以上内衣纸样工作经验（需外单经验）；能独立完成纸样、版型的修改与完善，放码，排唛架，熟悉贴合生产工艺；对设计作品理解力较强；熟练操作ET、格柏系统、Excel等；沟通能力、执行力强，有较高团队精神",
         bonus="外单经验是硬门槛；内衣/泳衣/家居服垂直经验优先", note="技术线高阶岗"),
    dict(no="A08930", name="质量管理工程师", line="质量管理", salary="15-18K", exp="3-5年", edu="大专",
         tags="团队管理经验、服装/纺织经验、贴身衣物、需要出差、项目管理经验、SQE（供应商）、DQE（研发）、内衣",
         duty="质量体系建设与维护；QC/QA团队管理；供应商质量管理；客户分析与改善闭环",
         req="大专及以上，纺织工程/服装设计与工程/质量管理相关专业优先；3年以上内衣/贴身衣物行业管理经验；熟悉内衣面料特性（针织/蕾丝/莫代尔/冰丝）与工艺（裁剪/缝制/后整/包装）及常见缺陷；掌握AQL抽样检验标准（GB/T 2828.1）；熟悉国标（GB18401、GB31701）及内衣行业标准（FZ/T 73012、GB/T 8878）；会用Excel做质量统计与趋势分析，了解鱼骨图/5Why/PDCA；能与工厂有效博弈推动整改",
         bonus="内衣品牌方或电商外发QC主管经验优先；熟悉潮汕/佛山内衣工厂资源优先；会开车、有驾照优先", note="需出差"),
    dict(no="A175757", name="品质经理", line="质量管理", salary="18-25K", exp="不限", edu="大专",
         tags="内衣、服装/纺织经验、需要出差、项目管理经验、QC七大手法、PQE经验（制程）、有质量管理经验",
         duty="①负责文胸/内裤/家居服等内衣品类全链路品质管理，搭建并落地原辅料/半成品/成品检验标准与品控SOP；②管控供应链品质：供应商/工厂准入审核、驻厂巡检、关键工序质控、品质异常整改与8D闭环；③主导成品出厂检验（FQC/OQC），处理客诉与退换货品质判定，输出品质数据报表；④协同研发/生产/采购，参与打样评审与工艺确认；⑤管理品控团队",
         req="5年以上内衣/家居服行业品质管理经验，熟悉内衣生产工艺与面料特性，有文胸品类经验优先；精通服装质检标准、AQL抽样、纺织品安全标准（GB 18401等）；具备供应链管控/工厂审核/异常处理与改善推进能力，有电商品牌/品牌内衣公司经验优先；熟练使用办公软件，能输出检验标准/SOP/品质分析报告",
         bonus="文胸品类经验；电商品牌/品牌内衣公司经验；QC七大手法、PQE制程经验；全勤奖+绩效奖金+社保+带薪假期+节日福利", note="⭐品质线管理岗，比A08930高一档；公司提供福利"),
    dict(no="A164643", name="货架电商资深推广", line="货架电商运营", salary="15-18K", exp="不限", edu="大专",
         tags="店铺站内推广、品类结构优化、爆款运营、店铺站外推广、有天猫运营经验、服饰鞋帽",
         duty="负责货架电商服饰内衣店铺推广；建立有效推广计划、把控全店费率、完成销售额及利润目标；优化付费端流量效率带动人群资产流转；定期输出店铺数据分析报告；关注行业及竞店推广策略",
         req="全日制大专及以上，5年以上天猫推广经验；精通淘系付费工具（直通车/引力魔方等）；有年推广费2000万+/爆品主计划日预算2万+项目经验；精通付费带动免费流量底层逻辑；擅长优化点击率；精通数据银行/策略中心/DMP等；主要负责天猫，辅助京东/唯品会/拼多多",
         bonus="熟悉天猫京东品牌广告及IP流量玩法；精通站外引流工具（CID、UDS等）", note="要求硬指标高，适合资深推广"),
    dict(no="A09339", name="高级纸样师", line="技术/生产", salary="15-20K", exp="3-5年", edu="中专/中技",
         tags="内衣/泳衣/家居服、服装、擅长ET",
         duty="版型研究与开发；负责设计款纸样及出品；版型修改与完善、版型确认跟进；大货生产可行性分析；试穿试洗测评跟进",
         req="3-5年以上内衣纸样工作经验；能独立完成纸样/放码/排唛架，熟悉贴合生产工艺；熟练操作ET、格柏系统、Excel等；沟通能力、执行力强，团队精神",
         bonus="—", note="中级纸样岗，晋升下一站=A56129"),
    dict(no="A207002", name="商务媒介主管", line="内容/达人BD", salary="10-15K", exp="3-5年", edu="大专",
         tags="不需要出差、企业品牌方经验、懂内容、要求媒介商务BD经验、品牌/代理商客户资源",
         duty="搭建BD流程+谈判框架+达人评估模型，带2-5人团队完成GMV/ROI；主导年度付费达人合作谈判；搭建达人ROI数据模型，主导周/月复盘；搭建品牌内容审核标准；分层运营头部/腰部/尾部达人；跟踪TikTok/Instagram/YouTube平台趋势",
         req="3-5年电商/新媒体/营销经验；1年以上TikTok直接运营经验+团队管理经验；主导过付费达人合作谈判；独立搭建过ROI数据模型；熟悉TikTok平台规则/算法/玩法；有出海/跨境项目经验",
         bonus="有海外达人/MCN资源优先", note="TikTok跨境达人方向"),
    dict(no="A122331", name="拼多多运营", line="货架电商运营", salary="7-10K", exp="1-3年", edu="大专",
         tags="拼多多、女装+内衣类目、爆款运营、付费推广",
         duty="拼多多店铺运营推广；活动提报、资源位争取；页面/活动/装修设计需求对接；推广效果评估与数据分析；竞品数据采集分析",
         req="大专以上，2年以上经验；做过女装和内衣类目（操盘年销千万以上店铺优先）；有打爆品+付费推广经验；熟悉线上活动推广手段；市场销售策略、数据分析、市场洞察力",
         bonus="操盘过年销千万以上店铺优先", note="电商线入门岗"),
    dict(no="A26839", name="亚马逊运营经理", line="货架电商运营", salary="18-25K", exp="3-5年", edu="本科",
         tags="店铺站内推广、不接受居家办公、亚马逊、店铺0-1搭建、爆款运营、合规管理、有跨境电商运营经验",
         duty="制定亚马逊整体运营策略、优化流程、完善线上运营体系、确保销售目标；全面负责店铺管理与推广（打造单品爆款）；定期分析运营数据，监控销售与流量，提高库存周转率；关注行业动态和竞品，优化产品线和运营策略",
         req="统招本科以上，英语四级；0-1亚马逊开店经验+3-5年亚马逊店铺运营经验；思维灵活有逻辑性；优秀的沟通、组织协调能力",
         bonus="—", note="唯一要求本科+英语四级；薪资最高档"),
    dict(no="A257211", name="唯品运营专员", line="货架电商运营", salary="8-10K", exp="1-3年", edu="大专",
         tags="品类结构优化、不接受居家办公、唯品会、爆品运营、有国内电商运营经验、服饰鞋帽、活动策划执行",
         duty="品牌线上渠道日常商品运营管理；店铺日常运营推广（规划爆款方案）；商家群/平台小二资源对接，争取官方资源；品牌产品线规划、货盘调配；营销数据统计分析与运营方案输出",
         req="大专以上，2年以上服饰行业经验，有服饰类电商商品运营经验；擅长沟通、商品数据分析及逻辑思维；熟悉办公软件+数据分析；有审美和选品能力，能独立管理货品、组织货盘；对产品趋势敏锐",
         bonus="—", note="电商线入门岗，与拼多多岗同级"),
    dict(no="A69663", name="Codex技能开发工程师", line="AI/数字化", salary="15-25K", exp="3-5年", edu="本科",
         tags="爬虫经验、Codex Skill、自动化脚本落地经验、运维开发经验、MCP、Python、AI Agent",
         duty="①Codex Skill开发与落地40%：拆解部门工作→识别可自动化任务→设计开发可复用Skill→测试调试→输出标准模板、提示词库、最佳实践；②跨部门项目推动30%：主导AI提效项目、协调部门落地、向管理层汇报、建立AI应用排行与激励机制；③人效数据与效果管理20%：建立人效指标，量化每个Skill效果，输出周/月人效报告；④培训与赋能10%：组织Codex及AI工具培训、搭建Skill库/知识库/案例库、培养AI种子用户",
         req="统招本科以上；精通Python（爬虫、自动化脚本、数据处理、API对接）；精通Codex（需求拆解、Skill设计、代码生成、调试测试、落地交付）；熟悉Requests/BeautifulSoup/Selenium/Playwright/Scrapy/XPath/CSS Selector；能独立完成域名/服务器/前后端/数据库/部署维护全链路；懂业务流程改造与效率提升项目；跨部门推动能力，结果导向",
         bonus="直播电商/内容电商/内衣服饰经验；飞书/钉钉/企业微信自动化搭建经验；成熟Codex Skill/AI Agent/MCP/爬虫项目案例；企业内部AI工具推广培训经验", note="全公司AI自动化改造核心岗，技术栈硬"),
]

# 与简女士截图「发布职位·10」校验
assert len(jobs) == 10, f"岗位数应为10，实际{len(jobs)}"

wb = Workbook()
ws = wb.active
ws.title = "岗位汇总"

header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
body_font = Font(name="微软雅黑", size=10)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
line_fill_map = {
    "技术/生产": "FFF2CC", "质量管理": "DDEBF7", "货架电商运营": "E2EFDA",
    "内容/达人BD": "FCE4D6", "AI/数字化": "E4DFEC",
}

headers = ["序号", "岗位编号", "岗位名称", "业务线", "薪资", "经验要求", "学历", "标签", "岗位职责", "任职要求", "加分项", "备注"]
ws.append(headers)
for c in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=c)
    cell.fill = header_fill; cell.font = header_font; cell.alignment = center; cell.border = border

for i, j in enumerate(jobs, start=1):
    ws.append([i, j["no"], j["name"], j["line"], j["salary"], j["exp"], j["edu"],
               j["tags"], j["duty"], j["req"], j["bonus"], j["note"]])
    r = ws.max_row
    fill = PatternFill("solid", fgColor=line_fill_map.get(j["line"], "FFFFFF"))
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = body_font; cell.alignment = wrap; cell.border = border; cell.fill = fill
    for c in (1, 2, 4, 5, 6, 7):
        ws.cell(row=r, column=c).alignment = center

widths = [5, 10, 26, 12, 9, 14, 10, 30, 55, 60, 30, 20]
for c, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:L{ws.max_row}"

ws2 = wb.create_sheet("速读总结")
summary = [
    ["维度", "内容"],
    ["岗位总数", "10 个（与简女士 BOSS 直聘「发布职位·10」清单一致；全部为广州润微科技·润微内衣，办公地：广州海珠区中大，HR：简女士）"],
    ["技术/生产线", "纸样主管 A56129（20-30K，5-10年，需外单）、高级纸样师 A09339（15-20K，3-5年）——晋升路径 A09339→A56129"],
    ["质量线", "品质经理 A175757（18-25K，5年+内衣，管理岗）、质量管理工程师 A08930（15-18K，3-5年，需出差）"],
    ["货架电商运营线", "货架电商资深推广 A164643（15-18K，5年+天猫，年推广费2000万+）、亚马逊运营经理 A26839（18-25K，本科+英语四级，0-1开店）、拼多多运营 A122331（7-10K，1-3年）、唯品运营专员 A257211（8-10K，1-3年）"],
    ["内容/达人BD线", "商务媒介主管 A207002（10-15K，TikTok达人BD，1年+TikTok经验）"],
    ["AI/数字化线", "Codex技能开发工程师 A69663（15-25K，Python+Codex+爬虫+全链路）"],
    ["业务线分布", "技术/生产2、质量管理2、货架电商运营4、内容/达人BD1、AI/数字化1"],
    ["新业务信号", "招聘Codex技能开发工程师，公司正在进行内部AI化改造，是切入AI岗的好时机"],
    ["求职建议", "电商运营岗最多且门槛低（拼多多/唯品1-3年7-10K）适合作为跳板；技术线需内衣+外单经验；AI线适合有Python/Codex/爬虫背景者"],
]
for row in summary:
    ws2.append(row)
ws2.column_dimensions["A"].width = 16
ws2.column_dimensions["B"].width = 110
for r in range(1, ws2.max_row + 1):
    for c in range(1, 3):
        cell = ws2.cell(row=r, column=c)
        cell.font = body_font; cell.alignment = wrap; cell.border = border
    if r == 1:
        for c in range(1, 3):
            cell = ws2.cell(row=r, column=c)
            cell.fill = header_fill; cell.font = header_font; cell.alignment = center

out = "/workspace/润微岗位汇总_10个岗位.xlsx"
wb.save(out)
print("saved:", out, "| 岗位数:", len(jobs))
