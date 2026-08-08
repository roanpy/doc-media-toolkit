# Contributing / 参与贡献

Thank you for improving Doc Media Toolkit. Issues and pull requests may be written in English or Simplified Chinese.

感谢参与改进文档媒体工具箱。Issue 和 Pull Request 均可使用英文或简体中文。

## Before opening an issue / 提交 Issue 前

- Search existing issues and confirm the behavior on the latest `main` branch.
- Remove confidential documents, local paths, customer names, credentials, and personal data from logs and samples.
- For vulnerabilities, follow [`SECURITY.md`](SECURITY.md) instead of opening a public issue.

- 请先搜索已有 Issue，并在最新 `main` 分支复现。
- 日志和样例必须移除机密文档、本机路径、客户名称、凭据和个人信息。
- 安全漏洞请按 [`SECURITY.md`](SECURITY.md) 私下报告，不要提交公开 Issue。

## Development / 开发

```bash
bash setup_env.sh
QT_QPA_PLATFORM=offscreen PYTHONPATH=src:tests .venv/bin/python scripts/run_tests_isolated.py
.venv/bin/ruff check src tests scripts
.venv/bin/python scripts/check_public_safety.py
```

Keep changes focused and reuse the existing architecture. A behavior change needs a regression test and matching documentation. Do not add runtime dependencies without explaining why the standard library or an existing dependency is insufficient.

变更应聚焦并遵循现有架构。行为修改必须包含回归测试和同步文档；新增运行时依赖时，需说明标准库和现有依赖为何不足。

## Pull requests / Pull Request 要求

- Describe the user-visible problem and the chosen behavior.
- List tests actually run and any platform or external-runtime boundary not tested.
- Keep source documents and generated packages out of Git.
- Do not enable or trigger GitHub Actions unless a maintainer explicitly requests it.

- 说明用户可见的问题与最终行为。
- 列出实际运行的测试，以及未覆盖的平台或外部运行时边界。
- 不要提交源文档和生成的安装包。
- 未经维护者明确要求，不要启用或触发 GitHub Actions。
