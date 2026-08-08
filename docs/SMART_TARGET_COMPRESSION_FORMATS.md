# DOCX / PDF / XLSX 图片压缩后端

状态：已合并 `main`（8d13d72，经用户明确批准），三个后端随包分发；正式 GUI 压缩页与 CLI 已接入（`compact_input_path` 按后缀路由），与 PPTX/图片/视频共用队列和设置。

## 目的

把三个格式后端接入同一个正式压缩入口，统一复用 PPTX 核心的目标容量、显示面积、
内容分类、质量门禁、有限纠偏、报告和安全输出规则；格式专用结构和布局校验仍留在
各自适配器中。容量纠偏会先比较素材编码计划，计划未变化时停止，避免重复重写文档
或 PDF。库级 API（`compact_docx` / `compact_pdf` / `compact_xlsx`）经
`pptx_video_compactor.compact_input_path` 对正式 GUI 和 CLI 暴露：队列混入这三类
文件时必须显式填写目标容量（MB），且图片预设不能为"不压缩"；GUI 在开始前校验并
给出明确提示。进度按文件粒度推进，取消在文件边界生效。

DOCX、PDF、XLSX/XLSM 中已经因质量门禁恢复原件的图片会在后续容量轮次锁定为原件；
其他图片仍可继续优化，因此不会为了一个不可安全压缩的图片而阻止整份文档收敛。

## 包含能力

- `docx_image_compactor`：`.docx` / `.docm` 图片关系替换，非媒体和 VBA 字节校验，
  Word/WPS/LibreOffice/Pages 逐页布局门禁。
- `pdf_image_compactor`：数字、扫描、混合 PDF 的安全 Image XObject 替换，OCR 层、
  页面几何、文字、表单、附件和链接保护；禁用有损 JBIG2。
- `xlsx_image_compactor`：`.xlsx` / `.xlsm` 图片关系替换，工作簿成员/宏/公式/图表
  保持，Excel/WPS/LibreOffice Calc PDF 布局门禁。

三个后端均失败关闭：签名、加密、非法包、无法解码或无法完成视觉验证时保留源文件，
不自动生成未验证输出。格式细则分别见 `SMART_TARGET_COMPRESSION_DOCX.md`、
`SMART_TARGET_COMPRESSION_PDF.md`、`SMART_TARGET_COMPRESSION_XLSX.md`。

## 依赖边界

`pyproject.toml` 声明 `pikepdf>=10.0`，并把三个后端列入 wheel；不声明
`python-docx`，DOCX 使用 ZIP/XML 原位处理。`Pillow`、`pypdf`、`pypdfium2` 等基础
依赖沿用项目现有依赖。Office、WPS、LibreOffice、Pages 和 Poppler `pdfimages` /
`pdftocairo` 是按任务探测的外部运行时，缺失时报告并拒绝发布未验证输出，不随 Python
包静默下载或捆绑。

`main` 包包含三个格式模块和 `pikepdf` 依赖，随 macOS 安装包分发；正式 UI 的
压缩页已直接暴露这三类格式（与 PPTX/图片/视频同队列），打包清单通过 hidden
imports 显式声明三个后端和 `pikepdf`。

## 验证和产物

统一分支已合并 `main`；合并及后续变更必须通过 Ruff、全量单元/容器测试、wheel
元数据与模块清单检查。实验 Python
包放在 `dist/experimental/<branch>/<commit>/`；三个格式已接入正式 UI，无需再构建
隔离实验应用分发格式能力。旧的三个独立实验包属于过程产物，统一包验证后
删除，不再继续发布。
