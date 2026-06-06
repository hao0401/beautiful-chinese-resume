<h1 align="center">Beautiful Chinese Resume</h1>

<p align="center">
  <strong>一键生成清爽、专业、HR 友好的一页中文 DOCX 简历</strong>
</p>

<p align="center">
  <img alt="Codex Skill" src="https://img.shields.io/badge/Codex-Skill-2F6F73?style=for-the-badge">
  <img alt="A4 One Page" src="https://img.shields.io/badge/A4-One_Page-31516A?style=for-the-badge">
  <img alt="Word Verified" src="https://img.shields.io/badge/Word-Verified-3C7A89?style=for-the-badge">
  <img alt="Fact Safe" src="https://img.shields.io/badge/Fact-Safe-6B7C8F?style=for-the-badge">
</p>

<p align="center">
  <sub>中文表达规范 · 目标公司适配 · 岗位关键词自然嵌入 · Word 渲染检查</sub>
</p>

<p align="center">
  <img src="./assets/readme-preview.svg" alt="Beautiful Chinese Resume preview" width="940">
</p>

---

## 为什么需要它

很多简历不是经历不够，而是 **重点不突出、表达太口语、版面太拥挤**。

`beautiful-chinese-resume` 把原始经历、目标公司、目标岗位和 JD，整理成一份 **中文、一页、A4、可投递** 的 `.docx` 简历。它不是夸张包装，也不是模板堆砌，而是把事实写清楚、把重点放对、把版面做干净。

<p align="center">
  <img src="./assets/readme-before-after.svg" alt="Before and after resume wording" width="900">
</p>

## 核心亮点

<table>
  <tr>
    <td><strong>一页 A4</strong><br>自动压缩弱相关内容，避免把简历塞成密集表格。</td>
    <td><strong>HR 扫读</strong><br>姓名、求职方向、教育背景和强相关经历优先呈现。</td>
    <td><strong>目标适配</strong><br>根据目标公司、岗位 JD 和关键词重排内容重点。</td>
  </tr>
  <tr>
    <td><strong>语言规范</strong><br>把流水账改成“动作 + 工具 + 产出”的简历表达。</td>
    <td><strong>视觉克制</strong><br>留白、层级、浅色分割线和少量强调色。</td>
    <td><strong>Word 检查</strong><br>生成后检查能否打开、是否一页、分区是否完整。</td>
  </tr>
</table>

## 适合的岗位

| 方向 | 会自然强化的关键词 |
| --- | --- |
| 跨境电商运营实习 | Amazon、Listing、ASIN、BSR、竞品调研、Excel |
| 内容运营实习 | 文案、选题、内容排期、账号维护、数据复盘 |
| 数据/运营助理 | 表格维护、数据清洗、指标跟踪、周报、SOP |
| 校招/实习投递 | HR 10 秒扫读、经历取舍、一页紧凑排版 |

## 快速开始

```powershell
git clone https://github.com/hao0401/beautiful-chinese-resume.git "$env:USERPROFILE\.codex\skills\beautiful-chinese-resume"
```

重启 Codex 后，直接这样说：

```text
使用 beautiful-chinese-resume，帮我生成一页中文 DOCX 简历。

目标公司：XXX
目标岗位：跨境电商运营实习生
岗位 JD：……
原始经历：……
```

也可以显式调用：

```text
Use $beautiful-chinese-resume to create a one-page Chinese DOCX resume from my raw experience, target company, and target role.
```

## 它如何工作

```mermaid
flowchart LR
  A["原始经历"] --> B["事实边界"]
  B --> C["目标公司 / JD 匹配"]
  C --> D["中文表达规范化"]
  D --> E["一页 A4 排版"]
  E --> F["Word 渲染检查"]
  F --> G["可投递 DOCX"]
```

## 生成标准

| 标准 | 要求 |
| --- | --- |
| 页面 | 一页 A4，清爽留白，不做密集表格 |
| 顶部 | 姓名、求职方向、联系方式清晰 |
| 分区 | 教育背景 / 项目经历 / 实习经历 / 技能证书 |
| 语言 | 不写空话，不编造结果，不堆砌关键词 |
| 排版 | 字号层级、浅色分割线、少量强调色 |
| 验证 | Word 打开检查、页数检查、分区检查 |

<details>
<summary><strong>项目结构</strong></summary>

```text
beautiful-chinese-resume/
├─ SKILL.md
├─ agents/
│  └─ openai.yaml
├─ assets/
│  ├─ readme-before-after.svg
│  └─ readme-preview.svg
├─ references/
│  ├─ content-rules.md
│  ├─ layout-rules.md
│  └─ targeting-keywords.md
└─ scripts/
   ├─ build_resume_docx.py
   └─ verify_resume_docx.py
```

</details>

<details>
<summary><strong>脚本用法</strong></summary>

从结构化 JSON 生成中文 DOCX：

```powershell
python .\scripts\build_resume_docx.py .\resume.json .\姓名-目标公司-岗位-中文简历.docx --style campus
```

检查 DOCX 是否满足投递要求：

```powershell
python .\scripts\verify_resume_docx.py .\姓名-目标公司-岗位-中文简历.docx --require-word --expect-one-page
```

| 检查项 | 说明 |
| --- | --- |
| Word 打开 | 确认 `.docx` 不是只在代码层面有效 |
| 页数 | 默认要求一页 |
| 分区 | 检查教育、项目、实习、技能等核心分区 |
| 关键词 | 检查目标岗位关键词是否自然出现 |
| PDF | 可按需导出投递版 PDF |

</details>

## License

暂未选择开源许可证。如需公开复用或再分发，请先补充 License。
