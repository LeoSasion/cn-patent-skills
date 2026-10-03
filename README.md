# 中国专利 Skills

这是一组面向 Codex 的中文专利工作流 Skill。首个正式版本把“准备一件新的中国发明专利申请”和“读懂一件已有专利”拆成两个边界清晰的 Skill。

## 包含的 Skills

### `prepare-cn-invention-patent`

显示名称：**中国发明专利申请准备**。

从当前代码、测试和技术材料出发，建立事实账本，检索现有技术，比较申请策略，设计权利要求，生成技术交底书与申请前校对材料，并审计版本、来源、隐私和 Git 隔离。

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

## 内置技术交底书

主 Skill 自带一份从零生成的 A4 空白 DOCX：

```text
.agents/skills/prepare-cn-invention-patent/assets/技术交底书_空白模板.docx
```

模板使用 14 个稳定内容控件标签，不依赖段落或表格索引。构建和填写命令：

```powershell
python .agents/skills/prepare-cn-invention-patent/scripts/build_disclosure_template.py
python .agents/skills/prepare-cn-invention-patent/scripts/check_template_hash.py
python .agents/skills/prepare-cn-invention-patent/scripts/fill_disclosure_template.py `
  --values values.json `
  --output 技术交底书.docx `
  --require-all
```

构建脚本需要 `python-docx` 和 `lxml`。交付审计在遇到 PDF 时需要 `pypdf`。Microsoft Word 仅作为 Windows 上的可选只读 PDF 导出后端。

## 验证

```powershell
python -m unittest discover `
  -s .agents/skills/prepare-cn-invention-patent/tests `
  -v
```

测试覆盖确定性模板构建、内容控件填充、隐私元数据、空包失败、显式部分交付、无附图说明边界和旧 XML 类型拒绝。正式完整包的审计默认失败关闭；只有分析包或章节包可以显式使用 `--allow-partial`。

## 保密与法律边界

- 不要把真实申请材料、身份原件、证照、客户数据、令牌或未公开代码提交到本仓库。
- 本仓库只包含工作流、脚本、参考规则和空白模板，不包含任何具体案件材料或真实申请文件。
- 正式提交前，重新核对 CNIPA 当期规则、工具和系统校验结果。
- 申请人、发明人、权属、公开披露、优先权和最终文本应由用户与具备相应资质的专业人员确认。

## 许可证

代码、Skill 文本和项目生成的空白模板按 [Apache License 2.0](LICENSE) 发布。资产来源见 [ASSET_PROVENANCE.md](ASSET_PROVENANCE.md)。
