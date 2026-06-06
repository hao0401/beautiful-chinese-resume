<p align="center">
  <strong>Beautiful Chinese Resume</strong>
</p>

<p align="center">
  一键生成面向目标公司和实习岗位的中文一页 A4 美观简历
</p>

<p align="center">
  <img alt="Codex Skill" src="https://img.shields.io/badge/Codex-Skill-2F6F73">
  <img alt="DOCX" src="https://img.shields.io/badge/Output-DOCX-31516A">
  <img alt="Language" src="https://img.shields.io/badge/Language-Chinese-3C7A89">
</p>

---

## 这是什么

`beautiful-chinese-resume` 是一个 Codex skill，用来把原始经历整理成一份 **中文、单页、A4、HR 友好、排版精致** 的 `.docx` 简历。

它不是普通的简历润色器。它关注的是完整交付：

- 语言规范：把口语化经历改成简历表达。
- 事实安全：不编造公司、数据、证书、奖项和经历。
- 岗位适配：根据目标公司、目标岗位和 JD 调整重点。
- 视觉排版：留白、层级、浅色分割线、少量强调色。
- Word 验证：生成后检查能否打开、是否一页、是否错位。

## 适合谁

适合正在投递实习、校招或初级岗位的人，尤其是：

- 跨境电商运营实习生
- 内容运营实习生
- 数据/运营助理
- 电商商品运营
- Amazon 相关运营岗位

对跨境电商岗位，会自然强化这些关键词：

`Amazon`、`Listing`、`ASIN`、`BSR`、`竞品调研`、`内容运营`、`关键词优化`、`Excel`、`数据透视表`、`运营复盘`

## 设计目标

一份好的中文实习简历，不应该像密密麻麻的表格，也不应该像模板网站的花哨海报。

这个 skill 的默认风格是：

- 第一眼干净。
- 10 秒内能看到重点。
- 内容紧凑但不拥挤。
- 版面有层级，但不过度装饰。
- 用词专业，但不夸大经历。

## 工作流

```mermaid
flowchart LR
  A["原始经历"] --> B["提取事实边界"]
  B --> C["匹配目标公司/JD"]
  C --> D["规范中文表达"]
  D --> E["生成一页 DOCX"]
  E --> F["Word 打开与页数检查"]
  F --> G["最终可投递简历"]
```

## 安装

把仓库克隆到 Codex skills 目录：

```powershell
git clone https://github.com/hao0401/beautiful-chinese-resume.git "$env:USERPROFILE\.codex\skills\beautiful-chinese-resume"
```

然后重启 Codex。

## 使用方式

在 Codex 中这样说：

```text
使用 beautiful-chinese-resume，帮我根据这份原始经历生成一页中文 DOCX 简历。
目标公司：XXX
目标岗位：跨境电商运营实习生
岗位 JD：……
原始经历：……
```

也可以直接用默认提示：

```text
Use $beautiful-chinese-resume to create a one-page Chinese DOCX resume from my raw experience, target company, and target role.
```

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

验证项包括：

- Word 是否能打开。
- 渲染后是否为一页。
- 核心分区是否存在。
- 关键词是否自然出现。
- 是否可以按需导出 PDF。

## 隐私与事实边界

这个仓库只包含 skill 规则、参考文档和生成/检查脚本。

请不要把以下内容提交到公开仓库：

- 真实个人简历。
- 真实手机号、邮箱、地址、身份证号。
- 生成出的 `.docx` 或 `.pdf` 投递文件。
- 包含目标公司内部信息的 JD 或面试材料。

skill 的核心原则是：

> 只优化表达、结构和排版，不编造经历。

## License

No license has been selected yet. Add one before public reuse or redistribution if needed.
