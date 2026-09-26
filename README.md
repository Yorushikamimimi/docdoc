# docdoc

docdoc 是一个给 Codex 用的 Word 修改辅助工具。你已经有一份 DOCX，只想改其中一小部分时，它会先看清目标，再选合适的修改方式，在副本上动手，并检查有没有误伤其他内容。

例如：改标题字号、改一段文字颜色、给某个单元格加底色、调整已有图片，或填写现有 Word 模板。**目前不负责从零创建新的 Word。**

## 为什么做这个

修改 DOCX 的难点常常不是“改不出来”，而是“只改一个地方，保存后其他结构也跟着变了”。docdoc 的流程是：**确认目标 → 选择方式 → 修改副本 → 比较前后 → 合格后交付**。

## 它怎么选修改方式

一次任务选一条主要路线，不把几种编辑器接力用在同一个输出文件上。

| 路线 | 适合什么情况 |
| --- | --- |
| OOXML Surgical | 目标和原值都明确，能唯一定位。例如把某个标题从 15 pt 改为 16 pt，尽量只改对应位置。 |
| OpenAI Documents | 需要先理解文档结构，再用程序修改。例如找到第三章的风险表，改指定单元格。 |
| Microsoft Word 原生 | 确实需要 Word 自己处理字段、修订或其他专有行为。 |
| Computer Use + Word | 主要靠视觉判断。例如把 Logo 往右挪一点，避免挡住文字。 |

没有一种路线永远最好。目标找不准、原值不符，或修改后出现非目标变化时，应停下来说明问题，不把有风险的副本当成成功结果。

## 最重要的原则

- 保留原文件，默认只修改副本。
- 修改前确认目标及周围结构。
- 修改后检查 DOCX 内部差异；可能影响排版时，再用同一环境比较前后页面。
- 页面看起来一样，不等于内部结构也一样。

## MCP 提供什么

- `inspect_docx`：先看 DOCX 的组成，例如段落、表格、书签、图片和关系。
- `compare_docx`：比较修改前后，区分文件内容变化和规范化后的 XML 结构变化。
- `patch_docx_part_exact`：确认目标唯一后，在新副本的指定位置做严格的小范围替换。
- `validate_docx_edit`：按这次任务允许的变化范围检查结果，发现越界就报告失败。

MCP 负责检查和精确替换；选择编辑路线与判断能否交付由 Skill 指导 Codex 完成。

## 怎么安装

以下步骤适用于 macOS/Linux。需要 `git`、`uv` 和 Python 3.10 或更新版本；使用 Word 原生或界面路线时，还需要本机可用的 Microsoft Word。

1. 把 Skill 放到 Codex 的用户级目录：

   ```sh
   git clone https://github.com/Yorushikamimimi/docdoc.git ~/.codex/skills/word-enhanced
   ```

2. 在 `~/.codex/config.toml` 中**追加**这个 MCP 条目，不要替换已有配置：

   ```toml
   [mcp_servers.word_enhanced]
   type = "stdio"
   command = "sh"
   args = ["-c", "exec uv run --no-project --with mcp==1.27.0 python \"$HOME/.codex/skills/word-enhanced/mcp/server.py\""]
   ```

3. 新开 Codex 对话；如果新工具仍未出现，重启 Codex。`uv` 首次启动可能需要下载 MCP SDK。

## 怎么使用

正常描述需求即可，不必自己挑选路线。例如：

- “帮我修改这个已有 Word，把第三章标题字号改成 16 pt，其他地方别动。”
- “把这个 Word 里‘成绩’那个单元格改成黄色，别改其他表格。”
- “把封面的 Logo 往右调整一点，不要挡住下面的字。”

## 当前边界与验证

docdoc 重点处理**修改已有 DOCX**。它不保证所有 Word 对象都能无损修改，也不保证 Word 保存或复杂文档的自动处理一定没有副作用。无法唯一定位目标，或检查发现非目标变化时，应停止并告知用户。

已用普通文字格式、表格单元格、目录字段及书签、浮动图片位置等已有 DOCX 修改场景验证流程。实验文件和记录不随仓库公开。

## 许可证

MIT，见 [LICENSE](LICENSE)。

## 贡献者

[@Yorushikamimimi](https://github.com/Yorushikamimimi)
