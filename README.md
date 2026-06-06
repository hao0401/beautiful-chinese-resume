<h1 align="center">Beautiful Chinese Resume</h1>

<p align="center">
  <strong>把原始经历变成一份干净、专业、可投递的一页中文 DOCX 简历</strong>
</p>

<p align="center">
  <img alt="Codex Skill" src="https://img.shields.io/badge/Codex-Skill-2F6F73?style=for-the-badge">
  <img alt="One Page" src="https://img.shields.io/badge/A4-One_Page-31516A?style=for-the-badge">
  <img alt="Word Verified" src="https://img.shields.io/badge/Word-Verified-3C7A89?style=for-the-badge">
  <img alt="DOCX Output" src="https://img.shields.io/badge/Output-DOCX-6B7C8F?style=for-the-badge">
</p>

<p align="center">
  <sub>目标公司适配 · 中文表达规范 · HR 10 秒扫读 · Word 渲染检查</sub>
</p>

<p align="center">
  <img src="./assets/readme-preview.svg" alt="Beautiful Chinese Resume preview" width="920">
</p>

---

## 它解决什么

很多简历不是内容差，而是 **重点埋得太深、语言太像流水账、版面太像表格**。

`beautiful-chinese-resume` 是一个 Codex skill，用来把原始经历、目标公司、目标岗位和 JD，整理成一份 **中文、一页、A4、视觉清爽、岗位匹配** 的 `.docx` 简历。

> 核心原则：只优化表达、结构和排版，不编造经历。

## 核心能力

| 能力 | 结果 |
| --- | --- |
| 一键整理 | 从原始经历直接产出可投递 `.docx` |
| 语言规范 | 把口语、流水账改成简历表达 |
| 目标适配 | 根据目标公司、岗位 JD 和关键词调整重点 |
| 美观排版 | 留白、字号层级、浅色分割线、少量强调色 |
| Word 检查 | 确认能打开、是一页、分区完整、没有明显错位 |

## 默认简历风格

<table>
  <tr>
    <td><strong>清爽</strong><br>不用密集网格和花哨模板，第一眼干净。</td>
    <td><strong>聚焦</strong><br>姓名、求职方向、学校、强相关经历优先出现。</td>
    <td><strong>克制</strong><br>少量强调色和细分割线，避免过度装饰。</td>
  </tr>
  <tr>
    <td><strong>真实</strong><br>不虚构公司、数据、奖项、证书和经历。</td>
    <td><strong>专业</strong><br>用动作、工具、产出组织每条经历。</td>
    <td><strong>可投递</strong><br>生成 DOCX 后用 Word 做最终检查。</td>
  </tr>
</table>

## 适合的岗位

| 方向 | 会自然强化的关键词 |
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

```powershell
git clone https://github.com/hao0401/beautiful-chinese-resume.git "$env:USERPROFILE\.codex\skills\beautiful-chinese-resume"
```

重启 Codex 后即可使用。

## 使用方式

```text
使用 beautiful-chinese-resume，帮我生成一页中文 DOCX 简历。

目标公司：XXX
目标岗位：跨境电商运营实习生
岗位 JD：……
原始经历：……
```

也可以直接调用：

```text
Use $beautiful-chinese-resume to create a one-page Chinese DOCX resume from my raw experience, target company, and target role.
```

## 生成标准

生成的简历默认满足：

- 一页 A4，不把内容塞成密集表格。
- 顶部姓名、求职方向、联系方式清楚。
- 分区标题有层级，使用浅色分割线。
- 正文保持可读字号，不靠极小字体硬压缩。
- 关键词自然出现，不机械堆砌。
- 内容太多时优先压缩弱相关经历，而不是牺牲版面。

## 项目结构

```text
beautiful-chinese-resume/
├─ SKILL.md
├─ agents/
│  └─ openai.yaml
├─ assets/
│  └─ readme-preview.svg
├─ references/
│  ├─ content-rules.md
│  ├─ layout-rules.md
│  └─ targeting-keywords.md
└─ scripts/
   ├─ build_resume_docx.py
   └─ verify_resume_docx.py
```

## 脚本

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

## License

暂未选择开源许可证。如需公开复用或再分发，请先补充 License。
