# write-feasibility-study-sgcc
Codex skill for drafting, formatting, and validating SGCC science and technology project feasibility studies.

## 国网科技项目可研写作

面向国家电网科技项目，支持章节讨论、研究逻辑梳理、多人合稿、局部 Word 修订、引用核对、图表、进度和预算。以用户确认的母版和研究内容为准，不是官方模板或审批意见。

## 安装与依赖

将仓库克隆到自己的 Codex skills 目录，目录名保持 `write-feasibility-study-sgcc`。也可通过 skill-installer 安装本仓库。

Python 辅助脚本使用标准库，建议 Python 3.10 或更新版本。Word 交付还需要可用的 `documents:documents` 和 `word-mathtype-native-equations` 技能；原生 COM 操作需要 Windows、Microsoft Word、MathType 和 PowerShell 7。纯聊天文字不启动这些软件。绘图和数学模型审计按需使用对应技能。

## 使用示例

调用 `$write-feasibility-study-sgcc`，然后说明：

> 根据我提供的研究内容，检查研究任务、实施方法和成果指标是否对应。先在聊天中给出修改意见，不生成 Word。

> 以我指定的最新 DOCX 为母版，只修改第三章，另存新版，保留其余文字、公式、域和图表；对本轮修改做内容和包级检查。

脚本接口可用以下无文档操作检查：

```sh
python scripts/inspect_docx.py --help
python scripts/patch_docx_text.py --help
python scripts/qa_docx.py --help
```

## 安全与验证边界

- 不覆盖母版；输出提交时拒绝覆盖已有文件。
- 保留 XML 命名空间作用域及兼容性声明；只在允许的普通文本范围执行替换。
- 不内置个人路径、真实项目材料或私人参考资料清单；资料记录留在当前项目工作目录，不提交到仓库。
- 默认不生成 PDF 或渲染 DOCX 页面。只有用户当次明确要求检查最终版式时才执行。
- 包检查、内容提取和 Word/MathType COM 检查分别报告。脚本检查不能证明文献真实性、科学结论正确性或完整 Word 兼容性。

## 许可证

MIT，详见 [LICENSE](LICENSE)。Word、MathType 及其他第三方软件不随本仓库提供，仍需遵守各自许可。
