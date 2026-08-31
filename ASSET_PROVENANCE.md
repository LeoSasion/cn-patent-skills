# 资产来源

## `技术交底书_空白模板.docx`

- 路径：`.agents/skills/prepare-cn-invention-patent/assets/技术交底书_空白模板.docx`
- 版本：1.0.0
- SHA-256：`d407d40f679acbe1cf02760699709c92dc085b38be55edf3f4f380e36ecbea91`
- 生成器：`.agents/skills/prepare-cn-invention-patent/scripts/build_disclosure_template.py`
- 生成方式：从新的空白 DOCX 确定性构建；不复用第三方模板的文字、样式、图片或元数据。
- 授权：随本仓库按 Apache License 2.0 发布。

生成器固定 OOXML ZIP 时间戳，清除作者、最近修改者、公司、管理者、自定义属性、自定义 XML、外部关系和修订标识。相同版本的生成器和依赖应产生相同 SHA-256；发布前仍以 `scripts/check_template_hash.py` 为准。

测试和渲染产生的 PDF、PNG 与临时目录只是本地质量保证中间件，不属于发布资产。
