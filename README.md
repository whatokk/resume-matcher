# 简历岗位匹配（HR 提效）

> WorkBuddy Skill · 屋里涛说

HR 招聘全流程助手：JD 拆解入库 → 简历匹配打分 → BOSS 直聘简历包导入。把招聘从「一份份看」变成「系统初筛 + 人看终面」。

## 技能清单

| 技能 | 说明 |
|---|---|
| **resume-matcher** · 招聘全流程 | ① JD 颗粒度拆解（薪资/学历/经验/行业词/技能词/职责词/硬性条件）入库；② 简历库解析与岗位匹配打分，输出 Excel 初筛；③ BOSS 直聘企业版简历 ZIP 自动解压、识别姓名意向、重命名入库。 |


## 安装

把 `skills/` 下的技能目录拷贝到 WorkBuddy 的技能目录：

```bash
cp -r skills/* ~/.workbuddy/skills/
```

Windows PowerShell：

```powershell
Copy-Item .\skills\* "$env:USERPROFILE\.workbuddy\skills\" -Recurse -Force
```

重启 WorkBuddy 后，技能列表即可看到。

## 使用要点

- 触发词：简历匹配、简历筛选、人岗匹配、匹配打分、岗位拆解、JD解析、BOSS直聘、简历导入、HR 提效。
- **⚠️ 隐私**：`skills/resume-matcher/data/` 存放真实简历库与匹配输出，含候选人姓名与联系方式，**已通过 .gitignore 排除，不上传**。克隆后该目录为空属正常，请放入自己的简历数据。

## 环境依赖

- Python 3.13
- openpyxl（Excel 输出）
- BOSS 直聘企业后台（导出简历 ZIP）

## 目录规范

```
resume-matcher/
└── skills/
    ├── resume-matcher/
```

每个技能遵循统一结构：`SKILL.md`（必需，含 name/description frontmatter）+ `scripts/`（可选）+ `references/`（可选）。

---

## License

MIT — 随意取用、修改、二次分发。
