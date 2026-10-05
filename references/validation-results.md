# 发布前验证记录

日期：2026-10-05。验证环境：Windows、Microsoft PowerPoint 16.0。

- 文件完整性、PowerShell 脚本语法、示例 JSON：通过。
- PowerPoint COM 创建演示文稿：通过。
- 4 页技术示例保存为 PPTX，中文文字、标题及右下角页码导出为 1920×1080 PNG：通过，逐页视觉检查。
- 已修复 PowerPoint 默认文本框 AutoSize 继承问题；修复后无几何溢出提示。
- 使用一次真实内置 ImageGen 生成桥梁插画，登记 SHA256 后组装测试；contain 等比放入、cover 居中裁切：通过，检查导出页面。
- 四页导出图由 PowerPoint 拼成 2×3 概览：通过。
- 未登记素材与素材哈希不匹配均在创建正式输出前被拒绝：通过。
- 标准 Skill Creator Python 验证器因环境缺 PyYAML 无法执行；未安装依赖。另用随包 PowerShell 检查必要 frontmatter、引用路径及包完整性。

这是技术验证，不是完整四风格真实项目的端到端验收。需要用户逐页材料、参考文件及风格选择后才可验证正式设计质量。素材 provenance 字段与哈希不能独立证明模型来源，必须与实际工具记录核对。
