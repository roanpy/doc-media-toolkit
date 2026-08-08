from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from pptx_output_watermark.pptx_video_support import scan_fidelity_warnings

# Minimal OOXML fragments for building test PPTX packages.

_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="text/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="mp4" ContentType="video/mp4"/>
<Default Extension="mp3" ContentType="audio/mpeg"/>
</Types>"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

_PRESENTATION = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="1" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="2" r:id="rId2"/></p:sldIdLst>
<p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"""

_PRESENTATION_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"""

_MASTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:pic>
<p:nvPicPr><p:cNvPr id="100" name="master video"/><p:cNvPicPr/><p:nvPr>
<a:videoFile r:link="rId1"/>
</p:nvPr></p:nvPicPr>
<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
</p:pic>
</p:spTree></p:cSld>
</p:sldMaster>"""

_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="media/master.mp4"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/poster.jpg"/>
</Relationships>"""


def _slide_with(
    *,
    audio: bool = False,
    linked_video: bool = False,
    grouped_video: bool = False,
    unsupported: bool = False,
) -> str:
    """Build slide1.xml with the requested media shapes."""
    shapes: list[str] = []
    if audio:
        shapes.append(
            "<p:pic>"
            "<p:nvPicPr><p:cNvPr id='10' name='audio1'/><p:cNvPicPr/><p:nvPr>"
            "<a:audioFile r:link='rId10'/>"
            "</p:nvPr></p:nvPicPr>"
            "<p:blipFill><a:blip r:embed='rId11'/></p:blipFill>"
            "<p:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='100' cy='100'/></a:xfrm></p:spPr>"
            "</p:pic>"
        )
    if linked_video:
        shapes.append(
            "<p:pic>"
            "<p:nvPicPr><p:cNvPr id='20' name='linked_video'/><p:cNvPicPr/><p:nvPr>"
            "<a:videoFile r:link='rId20'/>"
            "</p:nvPr></p:nvPicPr>"
            "<p:blipFill><a:blip r:embed='rId21'/></p:blipFill>"
            "<p:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='100' cy='100'/></a:xfrm></p:spPr>"
            "</p:pic>"
        )
    if grouped_video:
        shapes.append(
            "<p:grpSp>"
            "<p:nvGrpSpPr><p:cNvPr id='30' name='group1'/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            "<p:grpSpPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='200' cy='200'/>"
            "<a:chOff x='0' y='0'/><a:chExt cx='200' cy='200'/></a:xfrm></p:grpSpPr>"
            "<p:pic>"
            "<p:nvPicPr><p:cNvPr id='31' name='grouped_video'/><p:cNvPicPr/><p:nvPr>"
            "<a:videoFile r:link='rId30'/>"
            "</p:nvPr></p:nvPicPr>"
            "<p:blipFill><a:blip r:embed='rId31'/></p:blipFill>"
            "<p:spPr><a:xfrm><a:off x='10' y='10'/><a:ext cx='100' cy='100'/></a:xfrm></p:spPr>"
            "</p:pic>"
            "</p:grpSp>"
        )
    if unsupported:
        shapes.append(
            "<p:pic>"
            "<p:nvPicPr><p:cNvPr id='40' name='unsupported_video'/><p:cNvPicPr/><p:nvPr>"
            "<a:videoFile r:link='rId40'/>"
            "</p:nvPr></p:nvPicPr>"
            "<p:blipFill><a:blip r:embed='rId41'/></p:blipFill>"
            "<p:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='100' cy='100'/></a:xfrm></p:spPr>"
            "</p:pic>"
        )
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<p:sld xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'"
        " xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'"
        " xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>"
        "<p:cSld><p:spTree>" + "".join(shapes) + "</p:spTree></p:cSld></p:sld>"
    )


def _slide_rels(
    *,
    audio: bool = False,
    linked_video: bool = False,
    grouped_video: bool = False,
    unsupported: bool = False,
) -> str:
    """Build slide1.xml.rels with the needed relationships."""
    rels: list[str] = []
    if audio:
        rels.append(
            "<Relationship Id='rId10' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio' Target='media/audio1.mp3'/>"
            "<Relationship Id='rId11' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/audio_poster.jpg'/>"
        )
    if linked_video:
        rels.append(
            "<Relationship Id='rId20' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/video' Target='https://example.com/external.mp4' TargetMode='External'/>"
            "<Relationship Id='rId21' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/linked_poster.jpg'/>"
        )
    if grouped_video:
        rels.append(
            "<Relationship Id='rId30' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/video' Target='media/grouped.mp4'/>"
            "<Relationship Id='rId31' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/grouped_poster.jpg'/>"
        )
    if unsupported:
        rels.append(
            "<Relationship Id='rId40' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/video' Target='media/clip.f4v'/>"
            "<Relationship Id='rId41' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/unsupported_poster.jpg'/>"
        )
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        + "".join(rels)
        + "</Relationships>"
    )


def _build_pptx(
    path: Path,
    *,
    include_master: bool = False,
    audio: bool = False,
    linked_video: bool = False,
    grouped_video: bool = False,
    unsupported: bool = False,
) -> Path:
    with ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("ppt/presentation.xml", _PRESENTATION)
        zf.writestr("ppt/_rels/presentation.xml.rels", _PRESENTATION_RELS)
        if include_master:
            zf.writestr("ppt/slideMasters/slideMaster1.xml", _MASTER)
            zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _MASTER_RELS)
        zf.writestr(
            "ppt/slides/slide1.xml",
            _slide_with(
                audio=audio,
                linked_video=linked_video,
                grouped_video=grouped_video,
                unsupported=unsupported,
            ),
        )
        zf.writestr(
            "ppt/slides/_rels/slide1.xml.rels",
            _slide_rels(
                audio=audio,
                linked_video=linked_video,
                grouped_video=grouped_video,
                unsupported=unsupported,
            ),
        )
    return path


class ScanFidelityWarningsTest(unittest.TestCase):
    def test_no_warnings_for_clean_pptx(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(Path(d) / "clean.pptx")
            self.assertEqual(scan_fidelity_warnings(path), [])

    def test_detects_embedded_audio(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(Path(d) / "audio.pptx", audio=True)
            warnings = scan_fidelity_warnings(path)
            self.assertEqual(len(warnings), 1)
            self.assertIn("音频", warnings[0])
            self.assertIn("1", warnings[0])

    def test_detects_linked_video(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(Path(d) / "linked.pptx", linked_video=True)
            warnings = scan_fidelity_warnings(path)
            self.assertEqual(len(warnings), 1)
            self.assertIn("链接视频", warnings[0])

    def test_detects_grouped_video(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(Path(d) / "grouped.pptx", grouped_video=True)
            warnings = scan_fidelity_warnings(path)
            self.assertEqual(len(warnings), 1)
            self.assertIn("组合", warnings[0])

    def test_detects_unsupported_extension(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(Path(d) / "unsupported.pptx", unsupported=True)
            warnings = scan_fidelity_warnings(path)
            self.assertEqual(len(warnings), 1)
            self.assertIn("白名单", warnings[0])

    def test_detects_master_video(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(Path(d) / "master.pptx", include_master=True)
            warnings = scan_fidelity_warnings(path)
            self.assertEqual(len(warnings), 1)
            self.assertIn("母版", warnings[0])

    def test_aggregates_multiple_categories(self) -> None:
        with tempfile.TemporaryDirectory() as d:
            path = _build_pptx(
                Path(d) / "multi.pptx",
                audio=True,
                linked_video=True,
                grouped_video=True,
                include_master=True,
            )
            warnings = scan_fidelity_warnings(path)
            self.assertEqual(len(warnings), 4)
            joined = " ".join(warnings)
            self.assertIn("音频", joined)
            self.assertIn("链接视频", joined)
            self.assertIn("组合", joined)
            self.assertIn("母版", joined)


if __name__ == "__main__":
    unittest.main()
