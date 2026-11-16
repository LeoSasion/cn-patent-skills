# 中国专利 Skills

这是一组面向 Codex 的中文专利工作流 Skill。首个正式版本把“准备一件新的中国发明专利申请”和“读懂一件已有专利”拆成两个边界清晰的 Skill。

## 包含的 Skills

### `prepare-cn-invention-patent`

显示名称：**中国发明专利申请准备**。

从创意、论文、实验、代码或已有交底书出发，筛选发明点、核验技术证据、检索现有技术、撰写和修订交底书及申请草稿。全套草稿按权利要求、说明书、摘要、适用的附图、查新笔记、请求书信息确认表和提交文件清单分件交付。

它不会替代专利代理师或律师，不会正式提交申请、承诺授权，或把未经官方工具和系统校验的文件称为可提交 XML。

### `explain-cn-patent`

显示名称：**中国专利通俗解读**。

读取用户提供的专利 PDF、全文或公开号，梳理权利要求树，用通俗语言解释技术方案，并把特征定位到权利要求、说明书和附图。它不提供侵权、无效或可专利性法律意见。

## 安装

把需要的目录复制到 Codex 可发现的 Skill 目录，例如：

```text
.agents/skills/prepare-cn-invention-patent/
.agents/skills/explain-cn-patent/
```

安装后可显式调用：

```text
$prepare-cn-invention-patent
$explain-cn-patent
```

旧测试名称 `cn-invention-patent-delivery` 已在 1.0.0 中由 `prepare-cn-invention-patent` 取代。正式发布不保留旧名称兼容壳，以免旧版继续触发过宽职责和默认策略。

## 默认交底书与分件申请草稿

默认采用 [六章交底书模板](.agents/skills/prepare-cn-invention-patent/assets/default-disclosure-template.md) 和 Markdown + A4 DOCX 双格式。交底书可按用户提供的模板制作。完整申请的文件名、渲染和审计方法见 [交付与版本](.agents/skills/prepare-cn-invention-patent/references/delivery-contract.md)。

```powershell
python -m pip install python-docx
python .agents/skills/prepare-cn-invention-patent/scripts/render_patent_docx.py `
  --input "<交底书源稿.md>" --output "<交底书交付稿.docx>"
```

旧八段式模板及填充脚本供历史稿维护。交付审计在遇到 PDF 时需要 `pypdf`，Microsoft Word 可作为 Windows 上的只读 PDF 导出后端。

## 验证

```powershell
python -m unittest discover `
  -s .agents/skills/prepare-cn-invention-patent/tests `
  -v
```

测试覆盖分件 Word 导出与审计、A4 页面、权项号、图片、文件哈希、隐私元数据和历史模板兼容。只有分析包或章节包可以显式使用 `--allow-partial`。

## 保密与法律边界

- 不要把真实申请材料、身份原件、证照、客户数据、令牌或未公开代码提交到本仓库。
- 本仓库只包含工作流、脚本、参考规则和空白模板，不包含任何具体案件材料或真实申请文件。
- 正式提交前，重新核对 CNIPA 当期规则、工具和系统校验结果。
- 申请人、发明人、权属、公开披露、优先权和最终文本应由用户与具备相应资质的专业人员确认。

## 许可证

本项目原创内容按 [Apache License 2.0](LICENSE) 发布。采用的上游模板和脚本随 Skill 附有各自的 MIT 许可声明，详见 [资产来源与上游版本](ASSET_PROVENANCE.md)。
