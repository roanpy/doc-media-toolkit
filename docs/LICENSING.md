# Licensing and distribution / 许可与分发

## 1. Project license / 项目许可证

Doc Media Toolkit's own source is released under the MIT License. MIT is appropriate for this project because it permits personal, commercial, closed-source, and open-source reuse while requiring preservation of the copyright and license notice.

文档媒体工具箱的自有源码采用 MIT License。它允许个人、商业、闭源和开源复用，同时要求保留版权与许可声明，适合希望降低采用门槛的桌面工具项目。

MIT does **not** relicense bundled libraries, fonts, icons, FFmpeg, PDFium, Qt, or office applications. Their licenses remain independent.

MIT **不会**重新许可依赖库、字体、图标、FFmpeg、PDFium、Qt 或办公软件；这些组件继续适用各自许可证。

## 2. Source release / 源码发布

A source snapshot may be published when all of the following are true:

- the repository contains `LICENSE`, `THIRD_PARTY_NOTICES.md`, and the original asset license files;
- the owner has confirmed rights to publish copied source, screenshots, icons, fonts, fixtures, and documentation;
- public-safety and credential-pattern checks pass;
- package metadata identifies MIT with an SPDX license expression;
- the public history contains only reviewed content.

源码快照可以在以下条件全部满足时公开：仓库包含项目及第三方许可文件；所有者确认有权发布复制源码和全部资源；公开安全及凭据扫描通过；包元数据使用 SPDX `MIT`；公开历史只包含经审核内容。

## 3. Standalone binary release / 独立安装包发布

MIT source readiness is not sufficient evidence for a compliant DMG or EXE. Each target-platform artifact must additionally:

1. include the project license, third-party inventory, Python license, Qt LGPL/GPL texts, asset licenses, and every installed Python distribution's metadata/license files;
2. preserve dynamic Qt libraries and avoid technical restrictions that prevent lawful replacement or relinking under the selected LGPL route;
3. record the exact FFmpeg version and build configuration;
4. include the exact FFmpeg distribution notices and, for GPL builds, provide the corresponding source or another legally sufficient source-delivery mechanism for that exact binary;
5. include the complete LibreOffice runtime and notices together when the offline build is used;
6. inventory native libraries actually collected by PyInstaller on that platform;
7. complete malware scanning, code signing, checksum generation, and platform-native smoke tests.

MIT 源码合规不等于 DMG/EXE 已满足分发要求。每个平台产物还必须实际携带项目、Python、Qt、资源和全部 Python 依赖的许可材料；保留 Qt 动态库及 LGPL 路径需要的替换/重新链接能力；记录 FFmpeg 精确版本和构建参数；GPL FFmpeg 必须为对应二进制提供对应源码或其他合法交付方式；离线 LibreOffice 必须连同完整许可材料一起分发；还要按真实产物清点 PyInstaller 收集的原生库，并完成恶意软件扫描、签名、哈希及目标平台冒烟测试。

The current local Homebrew FFmpeg reports `--enable-gpl` and `--enable-version3`; a package containing that binary has GPLv3-related distribution obligations even though the application source remains MIT. Until the corresponding-source and artifact inventory steps are automated and verified, local DMGs are test artifacts and must not be published as formal open-source releases.

当前本机 Homebrew FFmpeg 明确包含 `--enable-gpl` 和 `--enable-version3`。打入该二进制的安装包需要履行 GPLv3 相关分发义务，尽管应用自有源码仍是 MIT。在对应源码交付和真实产物依赖清单完成自动化并验证前，本机 DMG 只能作为测试产物，不得作为正式开源发行包发布。

## 4. External applications / 外部应用

Microsoft Office, WPS Office, Keynote, Pages, and a system Poppler installation are external programs invoked through supported operating-system mechanisms. They are not redistributed by the standard package and are not covered by the project MIT license.

Microsoft Office、WPS Office、Keynote、Pages 及系统 Poppler 通过操作系统支持的方式从外部调用，不随标准安装包分发，也不属于项目 MIT 许可范围。

## 5. Release decision / 发布结论

- Public source repository: allowed after the repository gates pass.
- Python source/wheel: allowed after verifying included metadata and license files.
- Unsigned or ad-hoc local DMG: testing only.
- Public macOS/Windows installer: blocked until target-specific signing and binary-license audit pass.

- 公开源码仓库：仓库门禁通过后可发布。
- Python 源码包/wheel：确认包内元数据与许可文件后可发布。
- 未签名或 ad-hoc 本机 DMG：仅限测试。
- macOS/Windows 公开安装包：目标平台签名和二进制许可审计通过前禁止发布。

This document records engineering controls and is not legal advice. For commercial distribution or unusual FFmpeg/Qt configurations, obtain qualified legal review.

本文记录工程控制，不构成法律意见。商业分发或使用特殊 FFmpeg/Qt 构建时，应由合格专业人士进行法律复核。
