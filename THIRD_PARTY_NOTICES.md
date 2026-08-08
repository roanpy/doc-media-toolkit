# Third-Party Notices / 第三方许可说明

Doc Media Toolkit source code is licensed under MIT. That license applies only to this project's own source and does not replace any third-party license. The installed package metadata and the license texts carried by each release artifact are authoritative.

文档媒体工具箱自有源码采用 MIT 许可证。MIT 只覆盖本项目自有代码，不会替代任何第三方许可证；安装环境中的包元数据及发布产物实际携带的许可文本具有优先效力。

## Python runtime and transitive dependencies / Python 运行时及传递依赖

| Component | Relationship | License |
| --- | --- | --- |
| Python runtime | packaged interpreter/stdlib | PSF License |
| Pillow | direct | MIT-CMU |
| pikepdf | direct | MPL-2.0 |
| PySide6, PySide6-Addons, PySide6-Essentials, Shiboken6 | direct/transitive Qt for Python runtime | LGPL-3.0-only OR GPL-2.0-only OR GPL-3.0-only |
| python-pptx | direct | MIT |
| pypdf | direct | BSD-3-Clause |
| pypdfium2, PDFium, and the PDFium binary dependencies recorded by its wheel | direct/transitive | BSD-3-Clause, Apache-2.0, CC-BY-4.0, and component-specific licenses included by pypdfium2 |
| ReportLab | direct | BSD-3-Clause |
| comtypes | Windows direct | MIT |
| charset-normalizer | transitive | MIT |
| lxml | transitive | BSD-3-Clause and notices included by lxml |
| packaging | transitive | Apache-2.0 OR BSD-2-Clause |
| typing-extensions | transitive | PSF-2.0 |
| XlsxWriter | transitive | BSD-2-Clause |

The standalone build collects each installed distribution's `METADATA` and `LICENSE`, `COPYING`, `NOTICE`, or `licenses/` files. It fails when a required runtime package has no usable license text, except for the Qt for Python packages whose LGPL/GPL texts are maintained in this repository and bundled separately.

独立安装包会从每个已安装分发包收集 `METADATA` 以及 `LICENSE`、`COPYING`、`NOTICE` 或 `licenses/` 文件。必需运行时依赖缺少可用许可文本时构建失败；Qt for Python 的 LGPL/GPL 正文由本仓库保存并单独打入安装包。

## Build and development tools / 构建与开发工具

| Component | License / distribution status |
| --- | --- |
| PyInstaller | GPL-2.0-or-later with the PyInstaller bootloader exception; the packaged bootloader notice is bundled |
| pyinstaller-hooks-contrib | Apache-2.0 / GPL-2.0; build-time hooks, not imported as application runtime code |
| altgraph, macholib, setuptools | MIT; build-time dependencies |
| Ruff | MIT; development-only lint/format tool |
| uv | development/release tool; not embedded in application artifacts |

## Bundled assets / 内置资源

- `assets/fonts/NotoSansSC[wght].ttf`: SIL Open Font License 1.1; complete text: `assets/fonts/OFL.txt`.
- Phosphor icons: MIT; complete text: `assets/icons/PHOSPHOR-LICENSE.txt`.
- Qt for Python LGPL/GPL texts: `licenses/LGPL-3.0-only.txt` and `licenses/GPL-3.0-only.txt`.

## Optional or packaged native runtimes / 可选或打包的原生运行时

- **FFmpeg/FFprobe**: licensing is determined by the exact build configuration. The packaging script refuses to bundle binaries without matching `LICENSE`, `COPYING`, or `NOTICE` files. A build containing `--enable-gpl` is a GPL build; distributors must also satisfy the corresponding-source obligations for that exact binary. Do not describe the whole installer as “MIT only.”
- **LibreOffice**: included only by an explicit offline onedir build. The complete runtime and its own license notices must remain together.
- **Microsoft Office, WPS Office, Keynote, Pages, and system Poppler**: discovered and invoked externally; they are not included in the standard package.
- **Platform-native libraries**: the exact set collected by PyInstaller can differ on macOS and Windows. A public binary release requires an artifact-specific inventory and license audit in addition to this source-level list.

## Upstream projects / 上游项目

- Qt for Python: <https://www.qt.io/qt-for-python>
- FFmpeg legal information: <https://ffmpeg.org/legal.html>
- pypdfium2: <https://github.com/pypdfium2-team/pypdfium2>
- PyInstaller license: <https://pyinstaller.org/en/stable/license.html>

See `docs/LICENSING.md` before publishing source or binary releases. This file is an engineering compliance inventory, not legal advice.
