<h1 align="center">Beautiful Chinese Resume</h1>

<p align="center">
  <strong>一键生成 HR 友好、目标岗位适配、排版精致的一页中文 DOCX 简历</strong>
</p>

<p align="center">
  <a href="https://github.com/hao0401/beautiful-chinese-resume">
    <img alt="Codex Skill" src="https://img.shields.io/badge/Codex-Skill-2F6F73?style=for-the-badge">
  </a>
  <img alt="DOCX Output" src="https://img.shields.io/badge/Output-DOCX-31516A?style=for-the-badge">
  <img alt="Chinese Resume" src="https://img.shields.io/badge/Resume-Chinese-3C7A89?style=for-the-badge">
</p>

<p align="center">
  <sub>不是模板堆砌，也不是夸张包装。它把原始经历变成一份干净、专业、可投递的中文简历。</sub>
</p>

---

## 一句话介绍

`beautiful-chinese-resume` 是一个 Codex skill，用于把用户的原始经历、目标公司、目标岗位和 JD，整理成一份 **中文一页 A4 简历**。

它关注完整交付：内容规范、岗位匹配、美观排版、Word 渲染检查。

> 只优化表达、结构和排版，不编造经历。

## 视觉与内容标准

| 维度 | 标准 |
| --- | --- |
| 第一眼 | 干净、清爽、有留白，不是密密麻麻表格 |
| 顶部信息 | 姓名、求职方向、联系方式清晰可扫读 |
| 内容结构 | 教育背景 / 项目经历 / 实习经历 / 技能证书 |
| 语言表达 | 简历化、专业化，避免口语和空泛大词 |
| 岗位适配 | 根据目标公司、JD 和岗位关键词重排重点 |
| 最终交付 | 生成 `.docx`，并用 Word 检查是否一页、是否错位 |

## 适合的场景

| 场景 | 它会重点优化 |
| --- | --- |
| 跨境电商运营实习 | Amazon、Listing、ASIN、BSR、竞品调研、Excel |
| 内容运营实习 | 文案、选题、内容排期、账号维护、数据复盘 |
| 数据/运营助理 | 表格维护、数据清洗、指标跟踪、周报、SOP |
| 校招/实习投递 | HR 10 秒扫读、经历取舍、一页紧凑排版 |

## 工作流

```mermaid
flowchart LR
  A["原始经历"] --> B["事实边界"]
  B --> C["目标公司 / JD 匹配"]
  C --> D["语言规范化"]
  D --> E["一页 A4 排版"]
  E --> F["Word 渲染检查"]
  F --> G["可投递 DOCX"]
```

## 安装

把仓库克隆到 Codex skills 目录：

```powershell
git clone https://github.com/hao0401/beautiful-chinese-resume.git "$env:USERPROFILE\.codex\skills\beautiful-chinese-resume"
```

然后重启 Codex。

## 使用方式

在 Codex 中直接说：

```text
使用 beautiful-chinese-resume，帮我生成一页中文 DOCX 简历。
目标公司：XXX
目标岗位：跨境电商运营实习生
岗位 JD：……
原始经历：……
```

或使用默认提示：

```text
Use $beautiful-chinese-resume to create a one-page Chinese DOCX resume from my raw experience, target company, and target role.
```

## 输出效果

生成的简历默认遵循这些规则：

- 一页 A4，不把内容塞成密集表格。
- 顶部姓名、求职方向、联系方式清楚。
- 分区标题有层级，使用浅色分割线。
- 正文控制在可读字号，不靠极小字体硬压缩。
- 关键词自然出现，不做机械堆砌。
- 内容太多时优先压缩弱相关经历，而不是牺牲版面。

## 文件结构

```text
beautiful-chinese-resume/
├─ SKILL.md
├─ agents/
│  └─ openai.yaml
├─ references/
│  ├─ content-rules.md
│  ├─ layout-rules.md
│  └─ targeting-keywords.md
└─ scripts/
   ├─ build_resume_docx.py
   └─ verify_resume_docx.py
```

## 脚本能力

从结构化 JSON 生成中文 DOCX：

```powershell
python .\scripts\build_resume_docx.py .\resume.json .\姓名-目标公司-岗位-中文简历.docx --style campus
```

检查 DOCX 是否满足投递要求：

```powershell
python .\scripts\verify_resume_docx.py .\姓名-目标公司-岗位-中文简历.docx --require-word --expect-one-page
```

验证项：

| 检查项 | 说明 |
| --- | --- |
| Word 打开 | 确认 `.docx` 不是只在代码层面有效 |
| 页数 | 默认要求一页 |
| 分区 | 检查教育、项目、实习、技能等核心分区 |
| 关键词 | 检查目标岗位关键词是否自然出现 |
| PDF | 可按需导出投递版 PDF |

## License

暂未选择开源许可证。如需公开复用或再分发，请先补充 License。
