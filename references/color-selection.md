# 主色色板选择

内容模式与配色分别选择。用户已明确主色/品牌要求则不重复问；参考文件颜色是候选，不自动当作用户本次选择。保持参考结构，主色由用户决定；硬性品牌规范发生冲突时说明并澄清。

1. 用view_image展示随包assets/palette-board.png，配合可点选的request_user_input_async选项A—F；可同次问内容模式。不要只给颜色名称而不展示实际颜色。配置变更时用scripts/export_palette_board.ps1从assets/palettes.json重新导出；随包palette-board.json记录色板配置哈希，发布前核对。
2. 六套为候选而非限制，允许用户文字输入自定义HEX或明确沿用参考色。不能把默认预选、超时或推荐当实际提交。
3. 保存palette_id、primary_color、palette_confirmation_message实际回复。没有选择时继续内容解析与映射，等待后再生成配色相关素材。
4. 运行create_theme.ps1输出项目theme.json；根据角色填计划、图表与ImageGen提示词。脚本只创建配置，不会自动修改已有PPT/母版，也不会替换图片像素颜色。

## 角色

primary主色用于标题、主结构与必要标签；accent用于少量差异/关键数字；pale用于轻分区；body深中性色为默认正文；background默认白。不要把正文/照片都改成主色。不用整页高饱和填充追求“有配色”。

科学语义色（损伤/警示、对照系列）与主色分开记录，确保不能只依赖颜色识别。所有重要文字对背景有足够对比；白字放深色，正文不要放在复杂照片上。新主题用实际样稿验证，不宣称色值本身保证漂亮。

```powershell
& ./scripts/export_palette_board.ps1 -OutputDirectory ./planning/palette-preview
& ./scripts/create_theme.ps1 -PaletteId A -OutputPath ./planning/theme.json
& ./scripts/create_theme.ps1 -CustomPrimary '254B7A' -OutputPath ./planning/theme-custom.json
```

自定义色需要核对对比、强调色协调；脚本生成浅色候选，不代替设计判断。色板PNG为原生Office色块，无需ImageGen；它是选色界面，不是正式PPT页面素材。
