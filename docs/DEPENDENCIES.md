# Dependency rationale / 依赖说明

The default installation represents the complete desktop product. It intentionally installs every Python dependency required by any shipped workspace so a user does not discover a missing extra only after dropping a supported file. CLI-only or feature-sliced packages are not published today.

默认安装代表完整桌面产品，主动安装四个工作区真实需要的全部 Python 依赖，避免用户拖入已声明支持的文件后才发现缺少 extra。目前不单独发布“仅 CLI”或按功能拆分的安装包。

## Direct runtime dependencies / 直接运行时依赖

| Dependency | Required for | Decision |
| --- | --- | --- |
| Pillow | image decode/encode, previews, fingerprints, watermarks, all image compression paths | Required by most workspaces / 多数工作区必需 |
| PySide6 | desktop GUI, media playback, dialogs, settings | Required for the desktop app; not conceptually required by CLI-only use / 桌面应用必需，纯 CLI 理论上可不装 |
| python-pptx | PPTX structure and media operations | Required / 必需 |
| pypdf | PDF structure, metadata, and page operations | Required for PDF workflows / PDF 功能必需 |
| pypdfium2 | PDF rendering and validation | Required for PDF/document validation / PDF 与文档校验必需 |
| ReportLab | PDF/image-page generation and watermark output | Required for PDF generation / PDF 生成必需 |
| pikepdf | PDF embedded-image compression | Used only by the PDF compression branch, but included so the advertised PDF input works without a second install / 只服务 PDF 压缩，但为保证公开功能开箱可用而默认安装 |
| comtypes | Microsoft Office/WPS COM automation | Windows-only through an environment marker; never installed on macOS/Linux / 仅 Windows 条件安装 |

`lxml`, `packaging`, `charset-normalizer`, `typing-extensions`, `XlsxWriter`, the split PySide6 packages, Shiboken6, and PDFium component libraries are transitive dependencies selected by the direct packages. They must not be manually removed or pinned independently without testing the parent package.

`lxml`、`packaging`、`charset-normalizer`、`typing-extensions`、`XlsxWriter`、拆分的 PySide6 包、Shiboken6 及 PDFium 内部组件均由直接依赖带入，不应脱离上游父包单独删除或随意锁版本。

## Non-runtime dependencies / 非运行时依赖

- `PyInstaller`: `build` extra; used to create standalone apps. Its bootloader becomes part of the artifact, so its exception notice is bundled.
- `Ruff`: `dev` extra; lint and formatting only.
- FFmpeg/FFprobe, LibreOffice, Microsoft Office, WPS Office, Keynote, Pages, and Poppler are native/external runtimes rather than Python package dependencies. Availability depends on the selected workflow and platform.
- The manual Windows release workflow retrieves Poppler `26.02.0-0` from the upstream Windows packaging project and verifies its SHA-256 before running PDF tests; this pinned test input is not redistributed with the app.

- `PyInstaller` 属于 `build` extra，用于生成独立应用；bootloader 会进入产物，因此打包其例外条款。
- `Ruff` 属于 `dev` extra，只用于检查与格式化。
- FFmpeg/FFprobe、LibreOffice、Microsoft Office、WPS Office、Keynote、Pages、Poppler 是原生或外部运行时，不属于 Python 包依赖；是否需要取决于功能和平台。
- 手动 Windows 发布工作流从上游 Windows 打包项目下载 Poppler `26.02.0-0`，先校验 SHA-256 再运行 PDF 测试；该固定测试依赖不会随应用分发。

## Future split / 未来拆分条件

Do not split optional extras merely to reduce the dependency list. Add a separate `cli` or `pdf` extra only when the project publishes and tests a genuinely separate distribution whose supported feature set and error messages are clear. Until then, one complete dependency set is smaller operationally and safer for users.

不要只为缩短依赖列表就拆 extras。只有真正发布并持续测试独立 CLI/功能包，且支持范围和缺失提示明确时，才增加 `cli`、`pdf` 等拆分；在此之前，一套完整依赖对用户更稳定，工程维护成本也更低。
