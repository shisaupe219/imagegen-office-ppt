# 设计契约与脚本

单位point，元素按数组顺序从后到前。canvas/title_style/page_style/slides必填。正文标题全局统一，其余可title_box。slide含kind(cover/toc/body/ending)、title、core_message、evidence、sources、elements；可加notes、reference_page、layout_id、title_color、page_color。

planning/theme.json由create_theme.ps1生成，包含colors角色和推荐标题/正文规范；组装脚本仍读取计划中的具体HEX，不会自动加载theme或改已有PPT。生成计划时按角色填值，并在state保存主题路径与用户选择。可增加first_focus、reading_order、visual_review记录设计决策；脚本不自动评判这些语义字段。

## 元素

- text：text,x,y,w,h,size,font,color,bold,align；可highlights数组，项text/color/bold，所有同短语局部强调。
- image：asset_id,x,y,w,h,fit=contain/cover，不拉伸。
- 图像内必要文字旁注属于同一image素材，登记时无需拆成text；核对其准确与可读性，并说明图内文字不可单独编辑。正文标题、页码仍为原生文字。不得生成底部灰色注释；来源/待核/制作说明存notes与独立清单，必要技术限定存正文或图内。
- shape：x,y,w,h,shape=rect/roundrect/ellipse/rightArrow/downArrow，fill（none透明）、line_color、line_width。文字另放text。
- shape可加fill_gradient={from:'E4F0FA',to:'3D8ED0',direction:'horizontal'}，direction仅horizontal/vertical；两色原生渐变优先于fill。其他渐变、阴影和SVG图标需另行验证，不假装支持。
- line：x,y,w,h,color,width,end_arrow布尔；w或h可0；斜线从左上到右下。
- chart：x,y,w,h,chart_type=bar/column/line，categories字符串数组、series数组(name/color/values)，min默认0、max自动、unit、show_values默认true、size默认14。仅非负数据，line至少两类，柱状图min必须0。图表由准确可编辑形状组成，不是Excel联动Chart。负值、误差线、双轴、对数、堆叠需独立实现并验证，不假装支持。

默认asset_policy=hybrid允许全部对象；strict-imagegen仅text/image并要求图片均ImageGen，须说明准确图表限制。精确结构图用shape/text按真实比例组成。

## 清单与命令

asset-manifest.json含assets，每项id/path/sha256/origin/purpose。ImageGen需prompt/generation_record；用户项origin=user-provided并有source。相对路径PNG/JPEG，不越界、不符号链接。登记不是自动模型溯源证明。

```powershell
& ./scripts/check_environment.ps1
& ./scripts/register_asset.ps1 -ManifestPath ./asset-manifest.json -AssetPath ./assets/scene.png -Id scene-v1 -Prompt '实际提示词' -Purpose '场景示意' -GenerationRecord '实际记录'
& ./scripts/register_asset.ps1 -ManifestPath ./asset-manifest.json -AssetPath ./assets/project.jpg -Id project-v1 -Origin user-provided -Source '用户上传照片' -Purpose '项目现场'
& ./scripts/assemble_powerpoint.ps1 -PlanPath ./planning/plan.json -ManifestPath ./asset-manifest.json -OutputPath ./output/deck-v1.pptx
& ./scripts/export_slides.ps1 -PresentationPath ./output/deck-v1.pptx -OutputDirectory ./previews/v1
& ./scripts/export_contact_sheet.ps1 -PreviewDirectory ./previews/v1 -OutputDirectory ./previews/overview-v1 -Limit 20 -Columns 4
& ./scripts/validate_package.ps1
```

examples/slide-plan.json为无个人数据技术测试，examples/layout-library.json为布局坐标。正式项目须补图；空manifest允许无图技术测试。脚本拒绝覆盖，仅关闭本次文稿，audit几何初筛后还须视觉检查。

## v2.4 文字与生产规格

text/title_style/page_style支持padding_left/right/top/bottom（point，默认0）、vertical_align=top/middle/bottom（默认top）、line_spacing（字号倍数，默认1.2）、paragraph_before/after（point，默认0）。align仍为left/center/right。内边距与行距参与实际PowerPoint排版；文本框需留足面积，无溢出仍须视觉检查。

新正式plan必须含presentation_spec={speaker:'用户确认的姓名',presentation_date:'用户确认的时间'}；目录图片确有用户例外时记录toc_image_exception。每个body记录chapter_id（如一）、chapter_title、subtitle、logic_relation、icon_plan、image_plan；title为带编号完整一级标题，subtitle另建可见text。manifest每项asset_kind=icon/illustration/logo，可记录parent_asset_id、page_usage。ImageGen图标为栅格，不宣称可编辑矢量。

validate_presentation_plan.ps1在组装前校验已启用规格的计划，检查封面显示、章节字段与编号、可见二级标题、图标来源和跨页配图重复。旧无presentation_spec示例仅作技术兼容测试，不作为新正式稿模板；不得为绕过规则在正式稿省略该字段。脚本不自动命名章节、生成图标或渲染subtitle，不能代替语义/视觉核验。
