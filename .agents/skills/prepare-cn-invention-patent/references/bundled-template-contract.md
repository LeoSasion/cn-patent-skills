# 内置发明专利交底书模板合同

> 此模板用于复现和维护历史交底书。新案件的默认版式见 [六章交底书模板](../assets/default-disclosure-template.md)。

本合同只适用于 `assets/技术交底书_空白模板.docx`。用户提供其他 DOCX 时，以用户文件为准，重新盘点字段、内容控件和版式，不套用本合同。

## 资产基线

- SHA-256：`d407d40f679acbe1cf02760699709c92dc085b38be55edf3f4f380e36ecbea91`
- 来源：由本项目的 `scripts/build_disclosure_template.py` 从空白文档确定性生成，不复用第三方模板内容。
- 页面：A4 纵向，210 × 297 mm；上、下边距 20 mm，左、右边距 22 mm。
- 结构：1 个节；基本信息区、术语区和八个技术问题区；响应区域无固定高度，可随内容跨页扩展。
- 性质：内部技术交底表，不是请求书、法律意见或 CNIPA 正式提交表格。
- 隐私：不得含个人作者、最近修改者、自定义属性、批注、修订、宏、嵌入对象或外部关系。

使用前运行：

```powershell
python scripts/check_template_hash.py
```

哈希、页面几何、内容控件标签或隐私检查不一致时停止使用，先重新生成或人工核验。

## 稳定内容控件

不要依赖段落、表格或空白行索引。按下面的 `tag` 写入：

| tag | 内容 | 类型 |
|---|---|---|
| `PATENT_TITLE` | 发明专利名称 | 单行纯文本 |
| `CONTACT_NAME` | 技术联系人 | 单行纯文本 |
| `CONTACT_PHONE` | 电话 | 单行纯文本 |
| `CONTACT_FAX` | 传真 | 单行纯文本 |
| `CONTACT_EMAIL` | 电子邮箱 | 单行纯文本 |
| `TERMS` | 术语、缩写和定义 | 多段文本 |
| `SECTION_1_PROBLEM` | 实际要解决的技术问题 | 多段文本 |
| `SECTION_2_BACKGROUND` | 背景与最接近方案 | 多段文本 |
| `SECTION_3_DEFECTS_PURPOSE` | 缺点与发明目的 | 多段文本 |
| `SECTION_4_SOLUTION` | 详细技术方案 | 多段文本 |
| `SECTION_5_KEY_POINTS` | 关键点与拟保护组合 | 多段文本 |
| `SECTION_6_ADVANTAGES` | 技术效果及因果关系 | 多段文本 |
| `SECTION_7_ALTERNATIVES` | 可替代实施方式 | 多段文本 |
| `SECTION_8_SUPPORTING_MATERIALS` | 版本、测试、附图和待补资料 | 多段文本 |

## 填充合同

1. 复制模板或指定不同的输出路径；禁止把输出路径指向模板本身。
2. 将字段保存为 UTF-8 JSON 对象，键必须是上表 `tag`，值必须是字符串、数字、布尔值或 `null`。
3. 运行：

```powershell
python scripts/fill_disclosure_template.py --template assets/技术交底书_空白模板.docx --values values.json --output 技术交底书.docx
```

4. 未提供的内容控件保留为空白填写区；明确传入 `null` 或空字符串时清空该控件。
5. 未确认的身份和联系方式必须为空，不得用“待确认”、猜测值或参考专利信息填充。
6. 生成后重新盘点内容控件、元数据和外部关系，并渲染全部页面。

可使用 `--require-all` 强制 JSON 包含全部标签；可使用 `--list-tags` 只列出模板标签。脚本遇到未知标签、重复标签或缺失模板标签时失败。
