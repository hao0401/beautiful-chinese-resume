---
name: beautiful-chinese-resume
description: Create, rewrite, typeset, and verify polished one-page Chinese DOCX resumes for internships and entry-level roles. Use when the user wants a beautiful Chinese resume/CV, one-click resume generation, HR-friendly resume layout, target-company or job-description tailoring, Chinese resume language normalization, Word-rendered .docx output, PDF export, or resume optimization for ecommerce/operations roles using keywords such as Amazon, Listing, ASIN, BSR, competitor research, content operations, and Excel.
---

# Beautiful Chinese Resume

## Purpose

Turn raw user experience into a polished, one-page Chinese A4 resume for a target company and role. Produce a real `.docx`, normalize language, preserve facts, make the first 10 seconds easy for HR, and verify the rendered document in Word before calling it final.

## Default Mode

Use one-click mode unless the user explicitly asks for planning only:

1. Read the user's raw resume, notes, target company, target role, and JD if provided.
2. Extract a fact ledger: name, contact, education, projects, internships, skills, certificates, dates, tools, metrics, and constraints.
3. Ask only for missing name/contact/target role when they are required for a final resume. Otherwise proceed and mark non-critical gaps as `待补充`.
4. Tailor content to the target company/JD by selecting, ordering, and wording facts. Do not invent facts.
5. Generate a one-page A4 Chinese `.docx`.
6. Open/render-check in Word. If Word is unavailable or the platform is not Windows, say the resume was structurally checked but the true Word-rendered page count was not confirmed.
7. Iterate until the document is one page, aligned, readable, and not dense.

## Content Rules

Read `references/content-rules.md` before rewriting raw experiences.

Hard rules:

- Do not fabricate companies, titles, schools, dates, awards, certificates, metrics, tools, platforms, or results.
- Rewrite with professional resume language, not exaggerated claims.
- Prefer `动作 + 工具/方法 + 产出/影响` for bullet points.
- If no measurable result was provided, use accurate process/result words such as `整理`, `协助`, `跟进`, `维护`, `分析`, `复盘`, `输出`.
- Keep sections natural for Chinese internship resumes: `教育背景`, `项目经历`, `实习经历`, `技能证书`. Add `校园经历` or `获奖经历` only when useful.
- Keep the most relevant target-role evidence in the upper half of page one.

## Targeting Rules

Read `references/targeting-keywords.md` when the user provides a target company, JD, industry, or role.

- Use target company and JD for prioritization, keyword selection, file naming, and role wording.
- Do not normally print `目标公司` on the resume unless the user asks or the application context requires it. Use it to tailor the resume and file name.
- For ecommerce/operations internships, naturally surface relevant provided facts around Amazon, Listing, ASIN, BSR, competitor research, content operations, Excel, data organization, and review workflows.
- If the JD is absent, adapt lightly from the company/role without claiming company-specific internal knowledge.

## Layout Rules

Read `references/layout-rules.md` before producing DOCX.

Required visual standard:

- One A4 page, not a dense table resume.
- Clear top: name, target role, contact.
- Clean Chinese business style with whitespace, font hierarchy, thin dividers, and one restrained accent color.
- Body font around 9.5-10.5 pt; name around 18-22 pt.
- Use hidden tables only for alignment when needed, never as a dense grid.
- Compress by removing weak/irrelevant content before shrinking fonts.

## DOCX Generation

Prefer the bundled builder when the resume can be represented as structured JSON:

```powershell
python .\scripts\lint_resume_json.py .\resume.json
python .\scripts\build_resume_docx.py .\resume.json .\姓名-目标公司-岗位-中文简历.docx --style campus
```

Available bundled styles are `campus`, `business`, `fresh`, and `minimal`. Use `campus` by default for internship resumes. Use the builder's `STYLE_PRESETS` instead of hard-coding new layout constants when adding templates.

The builder accepts this shape:

```json
{
  "name": "姓名",
  "target_role": "跨境电商运营实习生",
  "target_company": "目标公司",
  "contact": ["电话", "邮箱", "城市"],
  "summary": ["一句核心优势，可省略"],
  "sections": [
    {
      "title": "教育背景",
      "items": [
        {
          "heading": "学校 | 专业 | 本科",
          "date": "2022.09 - 2026.06",
          "details": ["相关课程：电子商务、市场营销、数据分析"]
        }
      ]
    }
  ]
}
```

Use direct `python-docx` or document tools only when the user's source document requires preservation beyond the builder. Still follow the same layout and verification rules.

## Verification

Run the verifier after generating every final `.docx`:

```powershell
python .\scripts\verify_resume_docx.py .\姓名-目标公司-岗位-中文简历.docx --require-word --expect-one-page --required-section 教育背景 --required-section 项目经历 --required-section 实习经历 --required-section 技能证书
```

Verification must check:

- Word can open the document when Windows + Microsoft Word are available.
- The rendered page count is one when Word rendering is available.
- Sections exist and are not empty.
- Top identity/contact information appears.
- No visible overflow into a second page.
- Optional PDF export renders if requested.

If verification fails, revise the content or layout and rerun verification before final delivery. If Word rendering is unavailable, clearly report that only structural validation was completed.

## Final Response

Report only the important outcome:

- Final `.docx` path.
- PDF path if generated.
- Word verification result and page count.
- Any factual gaps left as `待补充`.

Do not paste the full resume text unless the user asks.
