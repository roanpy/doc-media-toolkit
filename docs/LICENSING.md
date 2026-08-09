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
5. keep LibreOffice external in the standard package; an offline bundle must include the complete runtime and notices and provide the matching source-code access required by its MPL and component licenses;
6. inventory native libraries actually collected by PyInstaller on that platform;
7. complete malware scanning, code signing, checksum generation, and platform-native smoke tests.

MIT 源码合规不等于 DMG/EXE 已满足分发要求。每个平台产物还必须实际携带项目、Python、Qt、资源和全部 Python 依赖的许可材料；保留 Qt 动态库及 LGPL 路径需要的替换/重新链接能力；记录 FFmpeg 精确版本和构建参数；GPL FFmpeg 必须为对应二进制提供对应源码或其他合法交付方式；标准包不内置 LibreOffice，离线 LibreOffice 包必须连同完整运行时、许可材料及其 MPL/组件许可证要求的匹配源码获取方式一起交付；还要按真实产物清点 PyInstaller 收集的原生库，并完成恶意软件扫描、签名、哈希及目标平台冒烟测试。

The withdrawn local Homebrew FFmpeg reports `--enable-gpl` and `--enable-version3`; a package containing that binary has GPLv3-related distribution obligations even though the application source remains MIT. Those old local DMGs are test artifacts and must not be published. Replacement candidates use the pinned source build described below.

已撤回的本机 Homebrew FFmpeg 明确包含 `--enable-gpl` 和 `--enable-version3`。打入该二进制的安装包需要履行 GPLv3 相关分发义务，尽管应用自有源码仍是 MIT。旧本机 DMG 只能作为测试产物，不得公开；替换候选使用下文的固定源码构建。

The withdrawn 0.2.0 candidates confirm that macOS arm64/x64 use Homebrew FFmpeg
8.1.2 GPL builds and Windows uses the static Gyan FFmpeg 8.1.2 essentials GPLv3
build. Each package contains license text, but none currently ships a verified
corresponding-source bundle for the exact binary. The preferred remediation is a
repeatable, source-pinned minimal GPL build whose FFmpeg and linked GPL component sources, build
scripts, patches, configuration, and hashes are released together. Removing bundled
FFmpeg and documenting an external runtime remains a compliant fallback, but reduces
the product's out-of-box core capability. Replacing the GPL build with an LGPL-only
build is not accepted without proving that the required H.264 quality and hardware
paths remain equivalent.

已撤回的 0.2.0 候选包确认：macOS arm64/x64 使用 Homebrew FFmpeg 8.1.2 GPL
构建，Windows 使用 Gyan FFmpeg 8.1.2 essentials 静态 GPLv3 构建。包内已有许可
正文，但尚未随准确二进制提供经过验证的对应源码包。推荐修复方式是建立固定源码、可重复执行的最小
GPL 构建，并同时发布 FFmpeg 及链接的 GPL 组件精确源码、构建脚本、补丁、配置和
哈希。移除内置 FFmpeg、改为外部依赖可以作为合规回退，但会削弱核心功能的开箱体验；
除非证明 H.264 画质与硬件链路等价，否则不以 LGPL-only 构建替换当前能力。

The replacement release path now uses `scripts/build_ffmpeg_runtime.sh`. It pins the
official FFmpeg 8.1.2 archive, x264 commit, and zlib 1.3.2 archive with SHA-256, builds only the required
GPL runtime plus VideoToolbox on macOS or Media Foundation on Windows, and emits a
platform-specific corresponding-source archive containing all three pristine source
archives, the build script, an empty patch record, the exact configure arguments,
toolchain information, licenses, and hashes. A public release must upload that source
archive on the same GitHub Release as each matching binary and bind it by hash in the
public-binary evidence file.

替换后的正式构建路径使用 `scripts/build_ffmpeg_runtime.sh`：以 SHA-256 固定官方
FFmpeg 8.1.2 源码包、x264 提交和 zlib 1.3.2 源码包，只构建项目所需的 GPL 运行时，并在 macOS 保留
VideoToolbox、在 Windows 保留 Media Foundation。脚本同时生成平台对应源码包，内含
三份原始源码、构建脚本、空补丁记录、准确 configure 参数、工具链信息、许可和哈希。
公开时必须把它与匹配二进制放在同一 GitHub Release，并在公开二进制证据文件中用
哈希绑定。

## 4. External applications / 外部应用

LibreOffice, Microsoft Office, WPS Office, Keynote, Pages, and a system Poppler installation are external programs invoked through supported operating-system mechanisms. They are not redistributed by the standard package and are not covered by the project MIT license. The explicit LibreOffice offline build is for private/internal use unless its complete artifact-level notices and matching source-access obligations have been independently verified.

LibreOffice、Microsoft Office、WPS Office、Keynote、Pages 及系统 Poppler 通过操作系统支持的方式从外部调用，不随标准安装包分发，也不属于项目 MIT 许可范围。显式 LibreOffice 离线包默认仅限私下/内部使用；只有其完整产物级许可材料与匹配源码获取义务另行核验通过后，才可公开。

## 5. Release decision / 发布结论

- Public source repository: allowed after the repository gates pass.
- Python source/wheel: allowed after verifying included metadata and license files.
- Unsigned or ad-hoc local DMG: testing only.
- Public macOS/Windows installer: blocked until target-specific signing and binary-license audit pass.
- Candidate workflow artifacts: allowed for private review only; the workflow must
  not have repository write permission or a release-publication step.

- 公开源码仓库：仓库门禁通过后可发布。
- Python 源码包/wheel：确认包内元数据与许可文件后可发布。
- 未签名或 ad-hoc 本机 DMG：仅限测试。
- macOS/Windows 公开安装包：目标平台签名和二进制许可审计通过前禁止发布。
- 候选 workflow 产物：仅可用于私下核验；workflow 不得拥有仓库写权限或 Release
  发布步骤。

This document records engineering controls and is not legal advice. For commercial distribution or unusual FFmpeg/Qt configurations, obtain qualified legal review.

本文记录工程控制，不构成法律意见。商业分发或使用特殊 FFmpeg/Qt 构建时，应由合格专业人士进行法律复核。
