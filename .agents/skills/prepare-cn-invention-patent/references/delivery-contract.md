# 交付与版本

## 文件名

单独交底书的文件名为 `{规范化案件名}_{YYYYMMDDHHmmss}.md` 和同名 `.docx`。时间戳用落盘地的本地时间。目录默认使用案件工作区的 `outputs/{案件标识}/`，用户指定时从其指定。每次实质修订另存，追加简短的合并或纠正摘要。同一秒发生重名时等到下一秒或选择新目录。

全套申请使用一个时间戳目录，目录内的分件文件名采用上游固定编号。下表每个文件均生成独立 MD 和 DOCX。

| 文件主名 | 范围 |
| --- | --- |
| `00_专利草案总览` | 可选 |
| `01_技术交底书` | 必备 |
| `02_权利要求书` | 必备 |
| `03_说明书` | 必备 |
| `04_说明书摘要` | 必备 |
| `05_说明书附图` | 适用时必备 |
| `06_摘要附图` | 适用时必备 |
| `07_初步查新笔记` | 必备，注明已检索及未覆盖范围 |
| `08_发明专利请求书信息确认表` | 必备 |
| `09_查新与差异化分析` | 按需 |
| `提交文件清单` | 必备，列出已生成文件和待确认项 |

路径辅助脚本只输出建议路径。全套的 `--output-root` 指向私有案件输出根目录，单份修订的 `--case-dir` 指向既有私有案件目录。

```powershell
python scripts/make_output_paths.py --case-name "一种样例方法" --package --also-docx
python scripts/make_output_paths.py --case-name "一种样例方法" --case-dir "<案件目录>" --also-docx
python scripts/append_revision_log.py --case-dir "<案件目录>" --kind correct --user "纠正采样条件" --summary "更新第三章及图2，检查引用一致性" --artifacts "<本轮Markdown>,<本轮Word>"
```

修订记录保存在私有案件目录的 `交底书修订对话记录.md`。只写变更摘要、文件、证据来源及对查新和权项的影响。响应合并或纠正任务时附对应的留档摘要，让发明人确认变更。

## A4 Word 导出

安装 `python-docx`（含 `lxml`）。从 Skill 目录运行以下命令，每个分件独立渲染。图片路径相对源稿目录，也可使用 `--base-dir` 指定允许读取图片的目录。先把 Mermaid 渲染为本地 PNG，在源码块后附 `<!-- ![图示 系统框图](附图/图1.png) -->`。独立位图也可直接用 Markdown 图片语法。复杂公式可用文档工具生成可编辑 Word 公式，或预先渲染为本地图像并用 `<!-- ![公式](附图/公式1.png) -->` 附在公式块后。逐一核对符号和可读性。公开文献 URL 写为正文文本，以便交接核验。

```powershell
python scripts/render_patent_docx.py --input "<源稿.md>" --output "<交付稿.docx>"
```

用户指定 DOCX 模板时，用文档工具按原模板填写。完成后渲染并检查全部页面、图表和权项编号。用户要求 PDF 时，从已校对的分件源生成，并按 [电子提交边界](filing-boundary.md) 标记审阅用途。旧八段式 DOCX 和旧合订 PDF 版式仅用于维护历史文件。

## 内部审计

将事实账本、权项支撑矩阵、审计和页面预览保存在私有位置。完整分件包在交付前运行以下审计。包内保留 SHA-256 清单 `02-文件校验.sha256` 和 `99-内部校验/provenance.json`。哈希清单覆盖每个实际交付的 Markdown、DOCX、附图及可选 PDF，清单自身除外。

```json
{
  "artifact_type": "application_docx",
  "workspace_root": "<授权的工作区绝对路径>",
  "invention_topic": "<发明名称>",
  "canonical_application_sources": {
    "abstract": "04_说明书摘要.md",
    "claims": "02_权利要求书.md",
    "description": "03_说明书.md",
    "drawings": [],
    "abstract_figure": null
  },
  "implementation_evidence": ["<工作区中的事实账本或原始证据路径>"],
  "layout_authority": "<本 Skill 的 document-contracts.md 绝对路径>",
  "official_rule_authorities": ["https://www.cnipa.gov.cn/art/2023/12/21/art_98_189197.html"],
  "build_pipeline": {
    "document_builders": ["<实际使用的构建脚本路径>"],
    "pdf_exporters": [],
    "orchestrators": []
  },
  "source_hashes": {"<规范源、附图、事实证据、版式和构建脚本路径>": "<实际 sha256>"},
  "confirmed_identity": {"applicant": "", "inventors": []}
}
```

相对路径从案件版本目录解析。使用附图时，`drawings` 列出说明书使用的位图，`abstract_figure` 指向其中一幅。添加审阅 PDF 时设置 `output_pdf` 和 `review_label`。每个来源只承担一个角色，公开专利归入 `raw_publication_references`，不充当本案事实。哈希键应与所记录的实际路径逐项对应。

```powershell
python scripts/audit_delivery_bundle.py --root "<案件版本目录>" --profile full
```

在 Git 项目内交付时，使用已忽略且未被跟踪的目录，并加 `--repo "<项目仓库>" --require-git-isolated` 核对。旧包继续支持原有 `review_pdf`、`application_sources` 和 `xml_preparation` 审计。分析包或章节包可分别用 `--profile analysis` / `--profile section`，仅在接受部分交付时加 `--allow-partial`。隐私、安全和哈希核验仍须通过。

## 版本与保密

按任务需要记录工作区基线、证据版本、公开披露的时间与访问范围、发明人具体贡献及权属缺口。用户陈述、公开证据和专业判断分别记录。源材料、既有稿、身份原件及上传/发送记录放在授权的受控位置。外部搜索只用可公开词汇。交付说明列出文件、检索边界、待确认信息和渲染结果。
