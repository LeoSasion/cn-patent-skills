# 资产来源

## 上游整合

| 上游固定版本 | 纳入本 Skill 的内容 | 许可 |
| --- | --- | --- |
| [HuangXinzhe/cn-patent-drafting@9be95a9](https://github.com/HuangXinzhe/cn-patent-drafting/tree/9be95a93c096d682f1a72613725cda990c2ef594) | 分件申请包、证据与权项支持流程、`make_output_paths.py`、`append_revision_log.py` | [MIT 原始声明](.agents/skills/prepare-cn-invention-patent/licenses/HuangXinzhe-cn-patent-drafting-MIT.txt) |
| [yjmm10/patent-skills@29704d2](https://github.com/yjmm10/patent-skills/tree/29704d22ef2f5ef0f90f09c3611b7968954059ed) | 六章 Markdown 交底书模板、时间戳与双格式交付、规范核验、脱敏检查。由 `tools/md_to_docx.py` 改造的 `render_patent_docx.py` | [MIT 原始声明](.agents/skills/prepare-cn-invention-patent/licenses/yjmm10-patent-skills-MIT.txt) |
| [handsomestWei/patent-disclosure-skill](https://github.com/handsomestWei/patent-disclosure-skill) | `yjmm10/patent-skills` 的工具说明列明的 Word 转换器原始来源 | [MIT 原始声明](.agents/skills/prepare-cn-invention-patent/licenses/handsomestWei-patent-disclosure-skill-MIT.txt) |

Word 渲染器补设 A4 尺寸、保留源稿编号、限制本地图像路径并在缺图时报错，复用本项目的 OOXML 元数据清理。

## `技术交底书_空白模板.docx`

- 用途：维护和复现旧版八段式交底书。
- 路径：`.agents/skills/prepare-cn-invention-patent/assets/技术交底书_空白模板.docx`
- 版本：1.0.0
- SHA-256：`d407d40f679acbe1cf02760699709c92dc085b38be55edf3f4f380e36ecbea91`
- 生成器：`.agents/skills/prepare-cn-invention-patent/scripts/build_disclosure_template.py`
- 生成方式：从新的空白 DOCX 确定性构建；不复用第三方模板的文字、样式、图片或元数据。
- 授权：随本仓库按 Apache License 2.0 发布。

生成器固定 OOXML ZIP 时间戳，清除作者、最近修改者、公司、管理者、自定义属性、自定义 XML、外部关系和修订标识。相同版本的生成器和依赖应产生相同 SHA-256；发布前仍以 `scripts/check_template_hash.py` 为准。

测试和渲染产生的 PDF、PNG 与临时目录只是本地质量保证中间件，不属于发布资产。
