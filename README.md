# imagegen-office-ppt v2

参考版式复用＋ImageGen主要图片＋Windows PowerPoint原生排版。

将仓库文件夹放入 `~/.codex/skills/imagegen-office-ppt/`。需要Windows、PowerShell与Microsoft PowerPoint；无需python-pptx。

用法：“使用 $imagegen-office-ppt，根据逐页稿与参考PPT，沿用参考版式，主要图片调用ImageGen，先制作三张代表性正文样稿。”

v2支持局部红字、原生色块/边框/箭头、准确技术几何和基础bar/column/line图表。图表为可编辑形状，不是Excel联动Chart。默认不做四风格，参考版式优先，增加视觉验收门槛。用户明确要求时仍可四风格或严格ImageGen模式。

主要图片生成，用户项目照片可登记；生成图不冒充真实证据。PDF只参考结构，不截图代替可编辑页面。公开仓库不含用户模板、原稿、项目数据、生成素材和正式PPT。质量仍需逐页评审，不承诺跨设备像素一致。
