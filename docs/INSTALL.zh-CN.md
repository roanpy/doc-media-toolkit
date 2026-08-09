# 安装与平台支持

[English](INSTALL.md)

## 当前分发状态

源码版本可以公开使用。macOS 和 Windows 预构建包仍属于候选产物；只有平台签名、
适用时的公证、产物级许可清单、恶意软件扫描、SBOM，以及内置 FFmpeg 对应源码交付
证据全部通过公开二进制门禁后，才允许作为正式安装包公开。未签名候选包不能等同于
可信安装包。

下一批候选会内置由固定 FFmpeg/x264 源码构建的 FFmpeg 8.1.2，并在同一 Release
提供匹配的对应源码包。旧 0.2.0 Homebrew/Gyan 候选已经撤回，不得重新公开。

## 已支持和已验证环境

| 平台 | 当前证据 | 公开包状态 |
| --- | --- | --- |
| macOS 13 及以上，Apple Silicon | 原生构建、包审计、DMG 校验和离屏启动冒烟通过 | 仅候选；ad-hoc 签名且未公证 |
| macOS 13 及以上，Intel | 原生构建、包审计、DMG 校验和离屏启动冒烟通过 | 仅候选；ad-hoc 签名且未公证 |
| 64 位 Windows | 在 Windows Server 2025 托管构建环境完成 x64 构建和离屏启动冒烟 | 仅候选；Windows 客户端最低版本及 Authenticode 签名尚未定版 |
| Linux | 非平台专用链路可尝试源码运行 | 暂无正式桌面包 |

## 从源码安装

推荐 Python 3.12。Git 以及下方按功能列出的外部运行时需要预先安装。

macOS 或 Linux：

```bash
git clone https://github.com/roanpy/doc-media-toolkit.git
cd doc-media-toolkit
bash setup_env.sh
.venv/bin/pptx-tools-gui
```

Windows PowerShell：

```powershell
git clone https://github.com/roanpy/doc-media-toolkit.git
cd doc-media-toolkit
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev,build]"
pptx-tools-gui
```

## 未来正式签名包的选择方式

- `macOS-arm64.dmg`：Apple Silicon Mac。
- `macOS-x64.dmg`：Intel Mac。
- `windows-x64-portable.zip`：64 位 Windows；先完整解压目录再运行。目录形式让
  Qt 动态库和许可文件保持可替换、可检查。
- `SHA256SUMS.txt`：用于核对传输完整性。哈希一致是必要条件，但不能替代代码签名
  和公证。

签名后的 macOS 包公开后，应从 DMG 将应用拖入“应用程序”。正式包应能正常通过
Gatekeeper，不要全局关闭 Gatekeeper。签名后的 Windows 包公开后，应先在文件属性
中核对发布者；发布者未知或不一致时，不要绕过 SmartScreen。

## 按功能使用的外部运行时

| 工作流 | 运行时 |
| --- | --- |
| 视频压缩、画质评估、视频水印和视频库转码 | 正式签名包内置固定源码构建的 FFmpeg 8.1.2；源码安装解析显式配置或系统二进制 |
| Windows 上 PPTX 转 PDF/图片 | PowerPoint 或 WPS 优先，LibreOffice 回退 |
| Windows 上 DOCX 转 PDF | Word 或 WPS 优先，LibreOffice 回退 |
| macOS 上 PPTX 转 PDF/图片 | LibreOffice 优先，Keynote 回退并需要 Automation 权限 |
| macOS 上 DOCX 转 PDF | LibreOffice 优先，Pages 回退并需要 Automation 权限 |
| PDF 内嵌图片分类 | 使用该压缩链路时需要 Poppler `pdfimages` |

Office、WPS、LibreOffice、Keynote、Pages 和系统 Poppler 都是外部程序，不属于标准
安装包内容。

## 升级与数据安全

- 普通压缩和水印默认另存新文件。
- 升级应用前先结束所有任务，并把每个视频库或图片库的完整目录复制到备份位置。
  仅备份 manifest 和 `.bak`、不备份对应媒体文件是不完整的。
- 0.2.0 的视频库和图片库都使用 schema version 1。未来若提高不兼容版本，必须先
  提供明确迁移流程。
- 在新版本中用代表性文档完成验证，并对每个库执行只读体检前，保留旧应用版本。

## 核验产物

使用操作系统提供的 SHA-256 工具，并与 `SHA256SUMS.txt` 对照；同时核对平台签名、
Release 标签和对应提交。发现不一致时不要运行文件，应按仓库安全策略私密报告。
