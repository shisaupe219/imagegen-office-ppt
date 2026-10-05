# ImageGen 素材与 Office PPT

做PPT就是这么简单。

可复用 Codex Skill：逐页 MD 与参考 PPT → 四种风格预览 → 用户选择并继续 → ImageGen 素材库 → 本机 PowerPoint 精确组装 → 正式效果图与 PPTX。

## 使用条件

Windows、本机 Microsoft PowerPoint、可用的 ImageGen 工具、桌面会话和所需字体。无需 python-pptx，无需为内置 ImageGen 配置 API key。Skill 不提供 Office 或图像生成服务，不能单靠 GitHub Actions 完成整个交互流程。

## 安装

将此仓库完整下载到 Codex 的 skills 目录下 `imagegen-office-ppt` 文件夹，使 `SKILL.md` 位于该文件夹根目录。保留所有 scripts、references、examples 和 agents 文件；不要只复制 SKILL.md。重新加载技能后调用：

> 使用 $imagegen-office-ppt，根据我的逐页 MD 和参考 PPT，先提供四种风格预览，待我选择并继续后制作正式 PPT。

## 工作方式

先设计版式再生成独立素材，图片和图形只使用登记的 ImageGen 文件。标题、正文和页码为原生可编辑文字，正式效果图由交付 PPTX 导出。四种风格预览默认也由 Office 组装，确保不新增图形素材。用户明确允许自由探索后，可改为 ImageGen 整页探索图，但不承诺原始素材保持不变。

严格模式下图表、表格、流程图也是 ImageGen 图片，不能原生编辑数据。用户可明确允许原生证据对象，但当前组装脚本仅支持文字和图片，扩展实现前须告知限制。图像中的数据与科学关系必须核对。

## 验证

`scripts/validate_package.ps1` 检查文件与脚本语法；`scripts/check_environment.ps1` 检查 Office。examples 提供无隐私输入和技术测试页面，非视觉质量展示。实际 ImageGen 生成与真实材料制作需要在对应工具会话中测试。

不把用户原材料、生成成果、账号凭证或本机绝对路径提交到仓库。

