# Open-source release checklist / 开源发布清单

This file separates repository preparation from owner-only publication actions.

本文件把代码仓库准备工作与必须由仓库所有者确认的公开操作分开。

## Completed in the repository / 仓库内已完成

- MIT source license and Python package metadata.
- English `README.md` as the default entry point, plus `README.zh-CN.md` with an accurate language-support statement.
- Contribution, security, conduct, issue, and pull-request guidance.
- Third-party dependency and bundled-runtime notices.
- Current-tree public-safety scan and local release checks.
- GitHub workflows remain manual-only; ordinary commits do not consume Actions minutes.
- Clean public source repository: <https://github.com/roanpy/doc-media-toolkit>.
- Private vulnerability reporting, secret scanning, and push protection enabled.

- MIT 源码许可证和 Python 包元数据。
- 中英文 README 入口，并准确说明界面语言覆盖范围。
- 贡献、安全、行为准则、Issue 和 Pull Request 规范。
- 第三方依赖及打包运行时许可说明。
- 当前工作树公开安全扫描和本地发布检查。
- GitHub workflow 保持手动触发，普通提交不会消耗 Actions 时长。
- 干净的公开源码仓库：<https://github.com/roanpy/doc-media-toolkit>。
- 已启用私密漏洞报告、Secret scanning 和 Push protection。

## Owner confirmation required / 需要仓库所有者确认

1. **Rights clearance**: owner confirmation received for the project's own source, generated icons/screenshots, and test fixtures. Bundled Noto Sans SC and Phosphor icons remain third-party assets and are distributed under their included OFL/MIT notices; no unknown or incompatible asset is currently listed.
2. **Release trust**: publish only target-platform builds. macOS public releases should use Developer ID signing and notarization; Windows builds should be produced and signed on Windows. Ad-hoc local DMGs are test artifacts, not trusted public releases.
3. **Language scope**: do not claim full bilingual UI until the video and image library workspaces are translated and tested in English.

1. **权利确认**：已收到所有者确认：项目自有源码、自行生成的图标/截图和测试样例均可公开许可。随包的 Noto Sans SC 和 Phosphor 图标仍是第三方资源，分别按随附的 OFL/MIT 说明分发；当前没有列出来源不明或许可不兼容的资产。
2. **发布可信度**：只发布目标平台原生构建。macOS 公开包应使用 Developer ID 签名和公证；Windows 应在 Windows 上构建并签名。本机 ad-hoc DMG 只能作为测试产物。
3. **语言范围**：视频库和图片库完成英文翻译及测试前，不得宣称界面完整双语。
