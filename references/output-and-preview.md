# 输出模式、四风格预览与标题栏（v2.6）

## 独立选择

初期依次确认content_mode=concise/detailed及output_mode=editable/image。已有本次明确选择不重复问。内容模式约束两种输出，均保护名称、数字、单位、条件、例子、互动和论证链。图片版不等于允许省略详细信息。

editable：ImageGen生成语义图标、独立配图；Office原生文字、色块、流程、准确图表。图像仍是栅格，不宣称全部可编辑。

image：每页完整标题/正文/图标/配图/页码由ImageGen一体生成，Office仅将一张整页图铺满一页。没有对象级文字/图形编辑，无内部点击揭示动画，互动题与答案可在同页分区或经确认拆页。无需另拆每个图标与配图。Logo用用户真实标识作输入，禁止生成伪造身份标识。

图片版以更强整体设计感为目标，不承诺最强或像素一致。先锁定逐页文案、版式、页序、画幅、章节编号、标题栏、标题区域与页码，再逐页生成；核对中文名称、数字单位、引线/流程，错页用ImageGen修正/重绘，不在Office覆盖原生文字绕过整页图规则。多信息详细页先优化构图，仍不可读则征求拆页，不静默删减。不复用场景配图，图标语言与Logo允许保持一致。

## 前六页，四种风格：直接缩略图

两种输出模式均采用同一选风格流程。先规划完整结构，提取前六页（含封面、目录与正文；不足六页用全部）的标题、核心观点、关键图文关系。A—D保持相同页序、内容核心、画幅和用户主色，分别在构图、图像语言、卡片组织及阅读节奏上形成差异。

每个方向直接调用ImageGen一次绘制一张2列×3行缩略图，共四张；每个格子为16:9页面示意，格外标页号和方向名称。此阶段禁止先逐页绘制、制作24张独立候选页面、Office组装预览PPTX或导出逐页图再拼接。export_contact_sheet.ps1仅用于正式稿概览，不用于风格选择。

缩略图采用标题、关键词和核心关系表达，避免塞入全部细节；不得改变原稿观点、捏造数据或事实。它只用于视觉方向选择，不是详细模式信息完整性或标题物理一致性的验收依据。可编辑版需说明正式稿使用原生对象实现选定方向，缩略图不构成像素一致承诺。

用户选择后从源稿与锁定规范重新制作全部正式页；不能裁切缩略图格子作为正式页，也不能视为前六页已经完成。保存preview_strategy=direct_contact_sheet、preview_styles、preview_page_ids、preview_content_mapping、style_confirmation_message及selected_style。未选择时可整理内容，不得将预选/超时当确认。用户明确要求例外时遵守并记录。

## 四种标题栏

展示实际T1—T4样本后点选；随包assets/titlebar-board.png为学院蓝示例，主色已选则使用scripts/export_titlebar_board.ps1的Primary/Pale重新导出匹配样本。与整体风格独立，不默认做4×4共16套。优先先确认标题栏，再制作四张缩略图；尚未选择时标明候选标题栏暂定。风格与标题栏均选定后才制作正式页。

- T1 简洁横线：编号＋标题＋细横线，正式/学术。
- T2 深色通栏：深主色带＋白字，结构强。
- T3 左侧标记：竖色条或编号徽章＋标题，灵活。
- T4 浅色圆角：浅色标题区＋深字，科普/教学。

正文标题区域、字号、编号与subtitle锁定；封面、目录、结束页独立构图。图片版按[物理规范](image-titlebar-contract.md)固定字号、坐标并逐页传入同一参考图，核验实际成图。可编辑版通过apply_titlebar_style.ps1将样式转成实际元素，不是只存一个标签。

## 桌面交付与脚本

默认正式PPTX保存到系统实际桌面路径，以[Environment]::GetFolderPath('Desktop')解析，兼容迁移/同步。用户指定路径优先。桌面无法解析或不可写时报告并请求可用位置，不偷偷改目录。素材、逐页图、预览和记录放桌面项目子文件夹；正式PPTX直接位于桌面。文件重名使用版本号，不覆盖。

assemble_powerpoint.ps1与assemble_image_powerpoint.ps1可省略OutputPath，此时生成桌面版本化文件名；预览/测试应显式指定工作目录。显式输出存在仍拒绝覆盖。

图片版plan结构：canvas(width,height)，presentation_spec(speaker,presentation_date)，slides每项kind/title/subtitle（正文）、chapter_id/chapter_title（正文）、core_message、asset_id、text_transcript、text_verified、sources、notes。manifest整页项asset_kind=slide，origin=imagegen并保存准确prompt/generation_record/sha256。text_verified代表已人工逐字检查，不是机器OCR证明。脚本核对可见转录、页序、来源/hash、整页图片比例；仍须实际图片验收。

最终交付PPTX、PNG、概览、素材清单与已选模式说明；图片版明确只能整体替换图片。图片版不存在原生文字检查，不能声称无文字溢出就证明图内文字正确。
