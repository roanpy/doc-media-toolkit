# AGENTS.md — 面向 AI 编程助手的仓库约定

本文件面向 Codex、Gemini、Claude Code 等 AI 编程助手，说明在本仓库中工作时的硬约束与验证路径。人类贡献者请优先阅读 [`CONTRIBUTING.md`](CONTRIBUTING.md)。

## 硬性约束

- **不触发 GitHub Actions**：本仓库的 Actions 额度受限，未经维护者明确要求，禁止 push 会触发 CI/Release 工作流的变更、禁止手动 `workflow_dispatch`。提交信息也不需要 `[skip ci]` 标记——约束在助手侧，而非污染提交历史。
- **安装包不入库、不发布**：DMG/EXE/构建中间产物不进入 Git；GitHub Release 保持仅源码（源码自带），除非维护者明确要求。
- **敏感信息不入库**：源文档、截图样例、本机路径、客户/项目名、Token、`.env`、模型缓存等一律不入库。提交前运行 `scripts/check_public_safety.py`。
- **行为变更必须带回归测试与同步文档**；新增运行时依赖需说明标准库或现有依赖为何不足。

## 验证路径（本地，无需 CI）

```bash
uv sync --locked --all-extras
QT_QPA_PLATFORM=offscreen PYTHONPATH=src:tests .venv/bin/python scripts/run_tests_isolated.py
.venv/bin/ruff check src tests scripts
.venv/bin/ruff format --check src tests scripts
.venv/bin/python scripts/check_public_safety.py
```

打包验证（仅在维护者要求时）：`.venv/bin/python scripts/build_standalone.py --help`，产物仅在 `dist/`，不入库。

## 提交与合并

- 分支命名 `agent/<描述>`，通过 PR 合入 `main`（main 受保护，直接 push 会被拒）。
- 提交信息用英文、单行主题，不附带 AI 署名（仓库策略：正文标注由 Codex 主导、其他 AI 协助，但不写进提交）。
- 合并后保持本地与 `origin/main` 同步，删除临时分支。

## 语言

Issue、PR、文档可使用英文或简体中文；README 保持双语结构，英文在前。
