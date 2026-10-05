# 设计契约和执行接口

计划 JSON 的单位统一为 point（1 inch = 72 point）。元素数组顺序是从后到前的层级。示例 `examples/slide-plan.json` 给出完整可运行的文字页面；实际制作必须加入登记 ImageGen 图片。示例不包含教学、研究或用户材料。

必需字段：`canvas.width/height`、`title_style`、`page_style`、`slides`。正文页的标题强制由全局 `title_style` 生成，页面局部标题参数无效。其他页面可使用 `title_box`。所有页均生成右下角页码。标题坐标、框宽高、字体、字号不可随页改变；内容超长应先精简或统一调整全局规范。

每个 slide 包含 `kind`（cover/toc/body/ending）、`title`、`core_message`、`evidence`、`sources`、`elements`。元素仅支持 `text` 和 `image`；文字字段 `text,x,y,w,h,size,font,color,bold,align`；图片字段 `asset_id,x,y,w,h,fit`，fit 是 `contain` 或 `cover`，不得拉伸。禁止原生图形、线条、外部图片链接和整页探索图。

`asset-manifest.json` 必含 `assets` 数组。每项 `id,path,sha256,origin,prompt,purpose`，origin 必为 `imagegen`。路径相对清单文件且必须在清单目录内；建议清单位于项目根目录、素材在 assets 子目录。PNG/JPEG 文件，图片不可是符号链接到项目之外。哈希来自实际文件；来源标签需与生成记录人工核对，脚本无法证明图像由哪个模型创建。

原生背景仅允许 canvas 的单色填充。带纹理、渐变和装饰的背景必须是登记图像。原生文字只用于语义文字，不用符号、重复字符或表情绕过图形来源约束。

## 命令

```powershell
& ./scripts/check_environment.ps1
& ./scripts/register_asset.ps1 -ManifestPath ./asset-manifest.json -AssetPath ./assets/illustration.png -Id illustration-v1 -Prompt '实际生成提示词' -Purpose '正文插画' -GenerationRecord '实际工具调用或生成记录标识'
& ./scripts/assemble_powerpoint.ps1 -PlanPath ./planning/slide-plan.json -ManifestPath ./asset-manifest.json -OutputPath ./output/deck.pptx
& ./scripts/export_slides.ps1 -PresentationPath ./output/deck.pptx -OutputDirectory ./previews/final
& ./scripts/export_contact_sheet.ps1 -PreviewDirectory ./previews/final -OutputDirectory ./previews/contact-first-six
& ./scripts/validate_package.ps1
```

环境检查启动 PowerPoint 并创建/关闭一个空文稿，不退出应用。组装和导出拒绝覆盖已有输出，修改时使用新版本路径。脚本不自动安装依赖，不访问网络，不执行输入中的代码。

`assemble_powerpoint.ps1` 附带保存 `.audit.json`，记录页面数、形状、源素材及文字溢出。溢出只是几何初筛，必须再看渲染图片。ImageGen 图片中的数字、文本、透明边界需人工视觉核查。原生图表例外必须另行扩展本脚本并验证，当前不得传入不支持元素。

Office 已在用户会话运行时可能复用同一进程；脚本不调用 Quit、不修改其他已打开文稿。正式成果在存在未满足要求时不得标为完成。
