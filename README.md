# Pandajay Skills

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> 我的 Claude Code Skills 收藏夹 - 提升开发效率的实用工具集

## 简介

这是我个人收集和开发的各种 Claude Code Skills，用于提升日常开发工作效率。所有 skills 都已经过实战验证，可直接使用。

## Skills 列表

| Skill 名称 | 描述 | 备注 |
|-----------|------|------|
| `prompt-reverse-engineer` | 分析文本写作风格并生成通用的模仿风格提示词。适用于需要理解某段文字风格并创作类似内容的场景。 | 独立使用，无外部依赖 |
| `chat-to-notes` | 将当前 AI 对话整理成一篇叙事风格的复盘文章，像写博客一样自然地还原探索过程。输出为 Markdown 文件，适合复盘、个人阅读和对外分享。 | 独立使用，输出到 `docs/notes/` 目录 |
| `req-to-dev-doc` | 将简单需求描述转换为结构化开发文档，包含用户故事、功能描述、影响范围、风险项和验收标准。适用于内部系统的需求整理。 | 依赖 `references/module-structure.md` 和 `references/impact_rules.md` |
| `dev-doc-match-resource` | 基于开发文档分析匹配当前开发方案的资源配置。结合能力资源矩阵，匹配合适的开发人员、测试人员及其能力等级。 | 依赖 `references/employee-data.md` |
| `req-to-hours-estimate` | 基于用户需求生成售前工时评估报告，包含需求分析、功能范围、技术方案、资源配置、工时评估和报价建议。 | 调用 `req-to-dev-doc` 和 `dev-doc-match-resource` skill，依赖 `references/example-report.md` 和 `references/generate_word_report.py` |

## 如何使用

1. 将 `skills/` 目录下的内容复制到你的 Claude Code skills 目录
2. 重启 Claude Code 或重新加载配置
3. 直接调用对应的 skill 即可

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

本项目采用 [MIT License](LICENSE) 开源协议。

---

Made with ❤️ by [熊猫 Jay](https://github.com/jaylpp)
