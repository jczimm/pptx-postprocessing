#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.9"
# dependencies = ["fonttools>=4.50"]
# ///
"""
make_portable.py — post-process a .pptx to make the presentation portable.

Hybrid tool: Python `zipfile` does the OpenXML package surgery; `officecli` is used
to validate the result and is reachable as the system `officecli` binary.

Modes (compose freely; default with no mode flag = --audit, read-only):

  --audit                Report every portability risk; never writes.
  --fix-links            Repair dangling/external relationships (e.g. a video link whose Target is "NULL").
  --embed-fonts          Embed every used font whose fsType permits it; report the rest.
  --externalize-videos   Portability for presentation, but not for the .pptx: pull embedded videos
                         OUT to a sidecar folder and link to them, shrinking the .pptx for transport.

Writes go to dist/<input-name> by default (override with -o).

Run:   uv run make_portable.py INPUT.pptx --audit
       uv run make_portable.py INPUT.pptx --fix-links --embed-fonts
       uv run make_portable.py INPUT.pptx --externalize-videos
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".avi", ".wmv", ".mkv", ".webm"}
AUDIO_EXTS = {".mp3", ".m4a", ".wav", ".aac", ".wma"}
MEDIA_EXTS = VIDEO_EXTS | AUDIO_EXTS

REL_VIDEO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
REL_AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
REL_MEDIA = "http://schemas.microsoft.com/office/2007/relationships/media"
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"

# Fonts present on essentially every PowerPoint host; no need to embed.
SAFE_FONTS = {
    "arial", "calibri", "cambria", "times new roman", "courier new",
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "tahoma", "verdana", "georgia", "trebuchet ms",
}

# Tokens that are theme references, not real font names.
THEME_TOKENS = {"+mj-lt", "+mn-lt", "+mj-ea", "+mn-ea", "+mj-cs", "+mn-cs"}

FONT_DIRS = [
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"),
]


# ---------------------------------------------------------------------------
# Small XML / relationship helpers (regex-based, to preserve round-trip fidelity)
# ---------------------------------------------------------------------------

REL_RE = re.compile(r"<Relationship\b[^>]*?/>")
ATTR_RE = re.compile(r'([A-Za-z:]+)="([^"]*)"')


def parse_attrs(tag: str) -> Dict[str, str]:
    return dict(ATTR_RE.findall(tag))


@dataclass
class Rel:
    raw: str
    id: str
    type: str
    target: str
    mode: str  # "" or "External"

    @property
    def is_external(self) -> bool:
        return self.mode == "External"


def parse_rels(xml: str) -> List[Rel]:
    rels = []
    for m in REL_RE.finditer(xml):
        tag = m.group(0)
        a = parse_attrs(tag)
        rels.append(Rel(
            raw=tag,
            id=a.get("Id", ""),
            type=a.get("Type", ""),
            target=a.get("Target", ""),
            mode=a.get("TargetMode", ""),
        ))
    return rels


def basename(target: str) -> str:
    return target.replace("\\", "/").rstrip("/").split("/")[-1]


def ext_of(target: str) -> str:
    return os.path.splitext(basename(target))[1].lower()


def resolve_target(part_name: str, target: str) -> str:
    """Resolve a relationship Target (relative to the .rels owner) to a package part path."""
    owner_dir = os.path.dirname(os.path.dirname(part_name))  # drop /_rels/<file>.rels
    return os.path.normpath(os.path.join(owner_dir, target)).replace("\\", "/")


def rels_part_for(part_name: str) -> str:
    d, f = os.path.split(part_name)
    return f"{d}/_rels/{f}.rels" if d else f"_rels/{f}.rels"


# ---------------------------------------------------------------------------
# Package: read-only access + streaming rewrite
# ---------------------------------------------------------------------------

class Package:
    """Read access to a .pptx (a zip). Use rewrite() to emit a modified copy."""

    def __init__(self, path: str):
        self.path = path
        self._zf = zipfile.ZipFile(path, "r")
        self.names: List[str] = self._zf.namelist()
        self._info: Dict[str, zipfile.ZipInfo] = {i.filename: i for i in self._zf.infolist()}

    def close(self):
        self._zf.close()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.close()

    def has(self, name: str) -> bool:
        return name in self._info

    def text(self, name: str) -> str:
        return self._zf.read(name).decode("utf-8")

    def size(self, name: str) -> int:
        return self._info[name].file_size

    def rewrite(
        self,
        out_path: str,
        modified: Dict[str, bytes],
        dropped: Set[str],
        added: List[Tuple[str, bytes]],
    ) -> None:
        """Write a new package: copy unchanged members streaming, apply edits.

        modified: part name -> new bytes (utf-8 already encoded)
        dropped:  part names to omit entirely
        added:    list of (name, bytes) new parts (compressed with DEFLATE)
        """
        with zipfile.ZipFile(out_path, "w") as zout:
            for info in self._zf.infolist():
                name = info.filename
                if name in dropped:
                    continue
                if name in modified:
                    zi = zipfile.ZipInfo(name, date_time=info.date_time)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zi.external_attr = info.external_attr
                    zout.writestr(zi, modified[name])
                    continue
                # Stream-copy unchanged member without loading it into RAM.
                zi = zipfile.ZipInfo(name, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                with self._zf.open(info, "r") as src, zout.open(zi, "w") as dst:
                    shutil.copyfileobj(src, dst, length=1024 * 1024)
            for name, data in added:
                zi = zipfile.ZipInfo(name)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(zi, data)


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

@dataclass
class Report:
    sections: List[Tuple[str, List[str]]] = field(default_factory=list)
    data: Dict[str, object] = field(default_factory=dict)

    def section(self, title: str, lines: List[str]):
        self.sections.append((title, lines))

    def render(self) -> str:
        out = []
        for title, lines in self.sections:
            out.append(f"\n{title}")
            out.append("-" * len(title))
            out.extend(f"  {ln}" for ln in (lines or ["(none)"]))
        return "\n".join(out)


def human(n: int) -> str:
    f = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if f < 1024 or unit == "GB":
            return f"{f:.1f} {unit}" if unit != "B" else f"{int(f)} B"
        f /= 1024
    return f"{f:.1f} GB"


# ---------------------------------------------------------------------------
# Font discovery / typeface collection
# ---------------------------------------------------------------------------

def collect_used_typefaces(pkg: Package) -> Set[str]:
    """Every literal typeface used in slides/masters/layouts/notes/theme."""
    used: Set[str] = set()

    for name in pkg.names:
        if not (name.startswith("ppt/") and name.endswith(".xml")):
            continue
        if not re.search(r"ppt/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|theme)/", name):
            continue
        xml = pkg.text(name)
        if "/theme/" in name:
            # Themes carry a large script-fallback table (<a:font script=...>) that is
            # NOT real usage. Only the major/minor latin/ea/cs heads are actual defaults.
            for block in ("majorFont", "minorFont"):
                m = re.search(rf"<a:{block}>(.*?)</a:{block}>", xml, re.S)
                if not m:
                    continue
                for face in re.findall(r"<a:(?:latin|ea|cs)[^>]*typeface=\"([^\"]*)\"", m.group(1)):
                    if face and face not in THEME_TOKENS:
                        used.add(face)
            continue
        for tf in re.findall(r'typeface="([^"]*)"', xml):
            if tf and tf not in THEME_TOKENS:
                used.add(tf)
    return {u for u in used if u.strip()}


def _fstype_blocks_embedding(fstype: int) -> Optional[str]:
    """Return a reason string if fsType forbids embedding, else None."""
    # Bit 1 (0x0002) = Restricted License (no embedding at all).
    if fstype & 0x0002:
        return "fsType: restricted-license (embedding not permitted)"
    return None


@dataclass
class Face:
    path: str
    ttc_index: int
    family: str
    bold: bool
    italic: bool
    fstype: int
    is_truetype: bool  # glyf outlines (PowerPoint embeds TrueType only)


def _read_faces(path: str) -> List[Face]:
    from fontTools.ttLib import TTFont, TTCollection

    def face_from(font: TTFont, idx: int) -> Optional[Face]:
        try:
            name = font["name"]
            fam = name.getDebugName(16) or name.getDebugName(1) or name.getDebugName(4) or ""
            full = name.getDebugName(4) or fam
            os2 = font["OS/2"] if "OS/2" in font else None
            fstype = int(getattr(os2, "fsType", 0)) if os2 else 0
            sel = int(getattr(os2, "fsSelection", 0)) if os2 else 0
            head = font["head"] if "head" in font else None
            mac = int(getattr(head, "macStyle", 0)) if head else 0
            bold = bool(sel & 0x20) or bool(mac & 0x01)
            italic = bool(sel & 0x01) or bool(mac & 0x02)
            is_tt = "glyf" in font
            # Record both the typographic family and the full name so weight-specific
            # run typefaces (e.g. "Helvetica Neue Medium") can be matched.
            f = Face(path, idx, fam.strip(), bold, italic, fstype, is_tt)
            f._full = full.strip()  # type: ignore[attr-defined]
            f._sub = (name.getDebugName(17) or name.getDebugName(2) or "").strip()  # type: ignore
            return f
        except Exception:
            return None

    faces: List[Face] = []
    try:
        if path.lower().endswith(".ttc"):
            coll = TTCollection(path, lazy=True)
            for idx, font in enumerate(coll.fonts):
                f = face_from(font, idx)
                if f:
                    faces.append(f)
        else:
            font = TTFont(path, lazy=True, fontNumber=0)
            f = face_from(font, 0)
            if f:
                faces.append(f)
    except Exception:
        pass
    return faces


def build_font_index(wanted: Set[str]) -> Dict[str, List[Face]]:
    """Map each wanted typeface (lowercased) -> matching faces found on disk."""
    wanted_norm = {w.lower(): w for w in wanted}
    nospace = {w.lower().replace(" ", "") for w in wanted}

    candidates: List[str] = []
    for d in FONT_DIRS:
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if os.path.splitext(fn)[1].lower() not in (".ttf", ".ttc", ".otf"):
                continue
            stem = os.path.splitext(fn)[0].lower().replace(" ", "")
            # Prefilter: only inspect files whose name hints at a wanted family.
            if any(ns and (ns in stem or stem in ns) for ns in nospace):
                candidates.append(os.path.join(d, fn))

    index: Dict[str, List[Face]] = {w: [] for w in wanted}
    for path in candidates:
        for face in _read_faces(path):
            full = getattr(face, "_full", face.family)
            keys = {face.family.lower(), full.lower(),
                    f"{face.family} {getattr(face, '_sub', '')}".strip().lower()}
            for key in keys:
                if key in wanted_norm:
                    index[wanted_norm[key]].append(face)
    return {k: v for k, v in index.items() if v}


# ---------------------------------------------------------------------------
# AUDIT
# ---------------------------------------------------------------------------

def run_audit(pkg: Package, report: Report) -> Dict[str, object]:
    rels_parts = [n for n in pkg.names if n.endswith(".rels")]

    external, dangling, local_hyperlinks, linked_ole, linked_charts = [], [], [], [], []
    for rp in rels_parts:
        for r in parse_rels(pkg.text(rp)):
            owner = rp
            if r.is_external:
                external.append(f"{owner}: {r.id} {r.type.split('/')[-1]} -> {r.target!r}")
                if r.target in ("", "NULL") or not r.target.strip():
                    dangling.append(f"{owner}: {r.id} ({r.type.split('/')[-1]}) Target={r.target!r}")
                t = r.target.lower()
                if t.startswith("file:") or re.match(r"[a-z]:\\", t) or t.startswith("\\\\"):
                    if "hyperlink" in r.type:
                        local_hyperlinks.append(f"{owner}: {r.target}")
            if "oleObject" in r.type and r.is_external:
                linked_ole.append(f"{owner}: {r.target}")
            if r.type.endswith("/oleObject") and "external" in r.target.lower():
                linked_ole.append(f"{owner}: {r.target}")

    # Linked chart data (externalData r:id with TargetMode External lives in chart rels)
    for n in pkg.names:
        if "/charts/" in n and n.endswith(".xml") and "_rels" not in n:
            if "<c:externalData" in pkg.text(n):
                linked_charts.append(n)

    # Media: embedded vs linked
    embedded_media, linked_media = [], []
    for rp in rels_parts:
        for r in parse_rels(pkg.text(rp)):
            if r.type in (REL_VIDEO, REL_AUDIO, REL_MEDIA) or ext_of(r.target) in MEDIA_EXTS:
                if r.is_external:
                    linked_media.append(f"{rp}: {r.id} -> {r.target}")
                elif ext_of(r.target) in MEDIA_EXTS:
                    embedded_media.append(resolve_target(rp, r.target))

    # Fonts
    pres = pkg.text("ppt/presentation.xml") if pkg.has("ppt/presentation.xml") else ""
    embed_flag = re.search(r'embedTrueTypeFonts="([^"]*)"', pres)
    has_fontlst = "<p:embeddedFontLst" in pres
    used = collect_used_typefaces(pkg)
    embedded_faces = {tf.lower() for tf in re.findall(r'<p:font typeface="([^"]*)"', pres)}
    non_safe = sorted(u for u in used
                      if u.lower() not in SAFE_FONTS and u.lower() not in embedded_faces)

    # Sizes by media type
    sizes: Dict[str, int] = {}
    for n in pkg.names:
        if n.startswith("ppt/media/"):
            e = ext_of(n) or ".?"
            sizes[e] = sizes.get(e, 0) + pkg.size(n)
    total = os.path.getsize(pkg.path)
    slide_count = sum(1 for n in pkg.names if re.match(r"ppt/slides/slide\d+\.xml$", n))

    # officecli schema validation
    validation = officecli_validate(pkg.path)

    report.section("Package", [
        f"path: {pkg.path}",
        f"total size: {human(total)}",
        f"parts: {len(pkg.names)}",
        f"slides: {slide_count}",
    ])
    report.section("Dangling / broken external relationships (BLOCKS portability)", dangling)
    report.section("All external relationships", external)
    report.section("Linked media (not embedded)", linked_media)
    report.section("Embedded media", sorted(set(embedded_media)))
    report.section("Linked OLE objects", linked_ole)
    report.section("Linked chart data", linked_charts)
    report.section("Local/UNC hyperlinks", local_hyperlinks)
    report.section("Fonts", [
        f"embedTrueTypeFonts: {embed_flag.group(1) if embed_flag else 'NOT SET'}",
        f"embeddedFontLst present: {has_fontlst}",
        f"distinct typefaces used: {len(used)}",
        f"embedded faces: {len(embedded_faces)}",
        f"non-universal AND not embedded (would substitute elsewhere): "
        f"{', '.join(non_safe) if non_safe else '(none)'}",
    ])
    report.section("Media size breakdown", [f"{e}: {human(s)}" for e, s in
                                           sorted(sizes.items(), key=lambda kv: -kv[1])])
    report.section("officecli validate", [validation])

    return {
        "total_size": total,
        "dangling": dangling,
        "external": external,
        "linked_media": linked_media,
        "non_safe_fonts": non_safe,
        "embed_fonts_set": bool(embed_flag and embed_flag.group(1) in ("1", "true")),
        "validation": validation,
    }


def officecli_validate(path: str) -> str:
    if shutil.which("officecli") is None:
        return "officecli not found on PATH (skipped)"
    try:
        p = subprocess.run(["officecli", "validate", path, "--json"],
                           capture_output=True, text=True, timeout=600)
        out = (p.stdout or p.stderr).strip()
        return out[:500] if out else f"exit {p.returncode}"
    except Exception as e:  # pragma: no cover
        return f"error: {e}"


# ---------------------------------------------------------------------------
# FIX-LINKS
# ---------------------------------------------------------------------------

def fix_links(pkg: Package, modified: Dict[str, bytes], report: Report) -> int:
    """Repair dangling external relationships. Returns count fixed."""
    lines: List[str] = []
    fixed = 0
    for rp in [n for n in pkg.names if n.endswith(".rels")]:
        xml = modified.get(rp, pkg.text(rp).encode()).decode("utf-8") \
            if rp in modified else pkg.text(rp)
        rels = parse_rels(xml)
        changed = False
        # Index internal media relationships in this rels file by target ext.
        media_targets = [r for r in rels
                         if (r.type == REL_MEDIA or ext_of(r.target) in MEDIA_EXTS)
                         and not r.is_external and r.target not in ("", "NULL")]
        for r in rels:
            if not r.is_external:
                continue
            if r.target not in ("", "NULL") and r.target.strip():
                continue  # external but resolvable-looking; leave alone
            if r.type in (REL_VIDEO, REL_AUDIO) and media_targets:
                # Repoint the dangling link to the embedded media (matches healthy slides).
                tgt = media_targets[0].target
                new_tag = f'<Relationship Id="{r.id}" Type="{r.type}" Target="{tgt}"/>'
                xml = xml.replace(r.raw, new_tag)
                lines.append(f"{rp}: {r.id} repointed NULL -> {tgt}")
                changed = True
                fixed += 1
            else:
                # No embedded counterpart: drop the dangling rel + its reference.
                xml = xml.replace(r.raw, "")
                owner = rp.replace("/_rels/", "/").rsplit(".rels", 1)[0]
                if pkg.has(owner):
                    oxml = modified.get(owner, pkg.text(owner).encode()).decode("utf-8") \
                        if owner in modified else pkg.text(owner)
                    oxml = strip_rel_reference(oxml, r.id)
                    modified[owner] = oxml.encode("utf-8")
                lines.append(f"{rp}: {r.id} ({r.type.split('/')[-1]}) removed (no embedded copy)")
                changed = True
                fixed += 1
        if changed:
            modified[rp] = xml.encode("utf-8")
    report.section("fix-links", lines or ["no dangling relationships found"])
    return fixed


def strip_rel_reference(xml: str, rel_id: str) -> str:
    """Remove elements/attributes in a part XML that reference a removed rel id."""
    # videoFile/audioFile link elements
    xml = re.sub(rf'<a:(video|audio)File\b[^>]*r:link="{rel_id}"[^>]*/>', "", xml)
    # p14:media embed (self-closing or wrapping)
    xml = re.sub(rf'<p14:media\b[^>]*r:embed="{rel_id}"[^>]*/>', "", xml)
    xml = re.sub(rf'<p14:media\b[^>]*r:embed="{rel_id}".*?</p14:media>', "", xml, flags=re.S)
    return xml


# ---------------------------------------------------------------------------
# EXTERNALIZE VIDEOS
# ---------------------------------------------------------------------------

def externalize_videos(
    pkg: Package,
    modified: Dict[str, bytes],
    dropped: Set[str],
    media_dirname: str,
    report: Report,
) -> List[str]:
    """Move embedded videos to a sidecar folder; link to them. Returns extracted part names."""
    lines: List[str] = []
    extract: Set[str] = set()  # package part names to pull out + drop

    slide_rels = [n for n in pkg.names
                  if re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", n)]
    for rp in slide_rels:
        owner = rp.replace("/_rels/", "/").rsplit(".rels", 1)[0]  # ppt/slides/slideN.xml
        xml = modified.get(rp, pkg.text(rp).encode()).decode("utf-8") \
            if rp in modified else pkg.text(rp)
        oxml = modified.get(owner, pkg.text(owner).encode()).decode("utf-8") \
            if owner in modified else pkg.text(owner)
        rels = parse_rels(xml)
        changed = changed_o = False

        for r in rels:
            is_video_file = ext_of(r.target) in MEDIA_EXTS and r.target not in ("", "NULL")
            part = resolve_target(rp, r.target) if is_video_file else None

            if r.type in (REL_VIDEO, REL_AUDIO) and is_video_file:
                # Link relationship -> point at the external sidecar file.
                rel_path = f"{media_dirname}/{basename(r.target)}"
                new_tag = (f'<Relationship Id="{r.id}" Type="{r.type}" '
                           f'Target="{rel_path}" TargetMode="External"/>')
                xml = xml.replace(r.raw, new_tag)
                extract.add(part)
                changed = True
                lines.append(f"{owner}: {r.id} linked -> {rel_path}")
            elif (r.type == REL_MEDIA or is_video_file) and is_video_file:
                # Embed relationship -> remove it and its <p14:media> element.
                xml = xml.replace(r.raw, "")
                oxml = strip_rel_reference(oxml, r.id)
                extract.add(part)
                changed = changed_o = True
                lines.append(f"{owner}: {r.id} embed removed ({basename(r.target)})")

        # Repair a still-dangling videoFile link in this slide by pointing it at the
        # extracted media (handles the slide-2 NULL case in one pass).
        for r in rels:
            if r.type in (REL_VIDEO, REL_AUDIO) and r.is_external and r.target in ("", "NULL"):
                cand = next((p for p in extract), None)
                if cand:
                    rel_path = f"{media_dirname}/{basename(cand)}"
                    new_tag = (f'<Relationship Id="{r.id}" Type="{r.type}" '
                               f'Target="{rel_path}" TargetMode="External"/>')
                    xml = xml.replace(r.raw, new_tag)
                    changed = True
                    lines.append(f"{owner}: {r.id} NULL link -> {rel_path}")

        if changed:
            modified[rp] = xml.encode("utf-8")
        if changed_o:
            modified[owner] = oxml.encode("utf-8")

    for part in extract:
        dropped.add(part)

    report.section("externalize-videos", lines or ["no embedded videos found"])
    report.data["extracted_parts"] = sorted(extract)
    return sorted(extract)


# ---------------------------------------------------------------------------
# EMBED FONTS
# ---------------------------------------------------------------------------

def embed_fonts(pkg: Package, modified: Dict[str, bytes], added: List[Tuple[str, bytes]],
                report: Report) -> int:
    from fontTools.ttLib import TTFont, TTCollection

    used = collect_used_typefaces(pkg)
    candidates = sorted(u for u in used if u.lower() not in SAFE_FONTS)
    index = build_font_index(set(candidates))

    embedded, skipped = [], []
    font_parts: List[Tuple[str, bytes, str, str]] = []  # (partname, bytes, typeface, slot)
    counter = 1

    def slot_of(f: Face) -> str:
        return ("boldItalic" if f.bold and f.italic else
                "bold" if f.bold else "italic" if f.italic else "regular")

    def extract_bytes(f: Face) -> bytes:
        font = (TTCollection(f.path, lazy=False).fonts[f.ttc_index]
                if f.path.lower().endswith(".ttc") else TTFont(f.path, fontNumber=0))
        buf = io.BytesIO()
        font.save(buf)
        return buf.getvalue()

    for typeface in candidates:
        faces = index.get(typeface, [])
        if not faces:
            skipped.append(f"{typeface}: not found in {', '.join(FONT_DIRS)}")
            continue
        # Choose the best face per slot. Prefer a face whose subfamily names the slot
        # exactly (e.g. "Regular") so weight-split families pick the canonical face.
        chosen: Dict[str, Face] = {}
        for f in faces:
            chosen.setdefault(slot_of(f), f)
        for f in faces:  # let an exact "Regular" subfamily win the regular slot
            if slot_of(f) == "regular" and getattr(f, "_sub", "").lower() == "regular":
                chosen["regular"] = f
        any_slot = False
        for slot, face in chosen.items():
            if not face.is_truetype:
                skipped.append(f"{typeface} [{slot}]: OpenType/CFF — PowerPoint embeds TrueType only")
                continue
            reason = _fstype_blocks_embedding(face.fstype)
            if reason:
                skipped.append(f"{typeface} [{slot}]: {reason}")
                continue
            try:
                data = extract_bytes(face)
            except Exception as e:
                skipped.append(f"{typeface} [{slot}]: extract failed ({e})")
                continue
            part = f"ppt/fonts/font{counter}.fntdata"
            counter += 1
            font_parts.append((part, data, typeface, slot))
            added.append((part, data))
            embedded.append(f"{typeface} [{slot}] <- {os.path.basename(face.path)} "
                            f"(fsType={face.fstype}, {human(len(data))})")
            any_slot = True
        if not any_slot and typeface not in {ln.split(' [')[0] for ln in skipped}:
            skipped.append(f"{typeface}: no embeddable face")

    if font_parts:
        _wire_font_parts(pkg, modified, font_parts)

    report.section("embed-fonts (embedded)", embedded)
    report.section("embed-fonts (skipped — recommend remap to a safe font)", skipped)
    return len(font_parts)


def _wire_font_parts(pkg: Package, modified: Dict[str, bytes],
                     font_parts: List[Tuple[str, bytes, str, str]]) -> None:
    # 1) [Content_Types].xml: ensure fntdata default
    ct_name = "[Content_Types].xml"
    ct = modified.get(ct_name, pkg.text(ct_name).encode()).decode("utf-8") \
        if ct_name in modified else pkg.text(ct_name)
    if 'Extension="fntdata"' not in ct:
        ct = ct.replace("</Types>",
                        '<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>')
        modified[ct_name] = ct.encode("utf-8")

    # 2) presentation rels: add a font relationship per part
    rels_name = "ppt/_rels/presentation.xml.rels"
    rels = modified.get(rels_name, pkg.text(rels_name).encode()).decode("utf-8") \
        if rels_name in modified else pkg.text(rels_name)
    existing_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    next_id = (max(existing_ids) + 1) if existing_ids else 1

    rid_for: Dict[str, str] = {}
    new_rel_tags = []
    for part, _data, _tf, _slot in font_parts:
        rid = f"rId{next_id}"
        next_id += 1
        rid_for[part] = rid
        target = part[len("ppt/"):]  # relative to ppt/
        new_rel_tags.append(f'<Relationship Id="{rid}" Type="{REL_FONT}" Target="{target}"/>')
    rels = rels.replace("</Relationships>", "".join(new_rel_tags) + "</Relationships>")
    modified[rels_name] = rels.encode("utf-8")

    # 3) presentation.xml: embeddedFontLst + flags
    pres_name = "ppt/presentation.xml"
    pres = modified.get(pres_name, pkg.text(pres_name).encode()).decode("utf-8") \
        if pres_name in modified else pkg.text(pres_name)

    # Group parts by typeface, collecting slots.
    by_tf: Dict[str, Dict[str, str]] = {}
    for part, _data, tf, slot in font_parts:
        by_tf.setdefault(tf, {})[slot] = rid_for[part]

    entries = []
    for tf, slots in by_tf.items():
        slot_xml = "".join(f'<p:{s} r:id="{rid}"/>' for s, rid in slots.items())
        entries.append(f'<p:embeddedFont><p:font typeface="{tf}"/>{slot_xml}</p:embeddedFont>')
    fontlst = f"<p:embeddedFontLst>{''.join(entries)}</p:embeddedFontLst>"

    if "<p:embeddedFontLst" not in pres:
        # Insert after notesSz (correct CT_Presentation child order), else before custShowLst,
        # else just before the close tag as a fallback.
        if re.search(r"<p:notesSz\b[^>]*/>", pres):
            pres = re.sub(r"(<p:notesSz\b[^>]*/>)", r"\1" + fontlst, pres, count=1)
        else:
            pres = pres.replace("</p:presentation>", fontlst + "</p:presentation>")

    # Set flags on the root element.
    def set_root_attr(xml: str, attr: str, value: str) -> str:
        m = re.search(r"<p:presentation\b[^>]*>", xml)
        tag = m.group(0)
        if f'{attr}="' in tag:
            new = re.sub(rf'{attr}="[^"]*"', f'{attr}="{value}"', tag)
        else:
            new = tag[:-1] + f' {attr}="{value}"' + ">"
        return xml.replace(tag, new, 1)

    pres = set_root_attr(pres, "embedTrueTypeFonts", "1")
    pres = set_root_attr(pres, "saveSubsetFonts", "0")  # we embed full faces
    modified[pres_name] = pres.encode("utf-8")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(argv: List[str]) -> int:
    ap = argparse.ArgumentParser(description="Make a .pptx fully portable.")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--audit", action="store_true")
    ap.add_argument("--fix-links", action="store_true")
    ap.add_argument("--embed-fonts", action="store_true")
    ap.add_argument("--externalize-videos", action="store_true")
    ap.add_argument("--media-dir", help="Sidecar folder name for externalized videos")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.input):
        print(f"error: no such file: {args.input}", file=sys.stderr)
        return 2

    do_write = args.fix_links or args.embed_fonts or args.externalize_videos
    report = Report()

    with Package(args.input) as pkg:
        if not do_write or args.audit:
            data = run_audit(pkg, report)
            report.data.update(data)

        if not do_write:
            _emit(report, args)
            return 0

        out_path = os.path.abspath(
            args.output or os.path.join("dist", os.path.basename(args.input)))
        if out_path == os.path.abspath(args.input):
            print("error: refusing to overwrite the input in place", file=sys.stderr)
            return 2
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        media_dirname = args.media_dir or (
            os.path.splitext(os.path.basename(out_path))[0] + "_media")
        media_dir_path = os.path.join(os.path.dirname(out_path), media_dirname)

        modified: Dict[str, bytes] = {}
        dropped: Set[str] = set()
        added: List[Tuple[str, bytes]] = []

        if args.embed_fonts and args.externalize_videos:
            report.section("WARNING", [
                "--embed-fonts and externalizing videos are opposite intents: "
                "embedding makes the deck self-contained; externalizing makes it small "
                "but dependent on the sidecar media. Applying both."])

        if args.fix_links:
            fix_links(pkg, modified, report)
        if args.externalize_videos:
            externalize_videos(pkg, modified, dropped, media_dirname, report)
        if args.embed_fonts:
            embed_fonts(pkg, modified, added, report)

        pkg.rewrite(out_path, modified, dropped, added)

        # Externalize: write the extracted media binaries to the sidecar folder.
        if args.externalize_videos:
            _extract_media(args.input, report.data.get("extracted_parts", []),
                           media_dir_path, report)

        before = os.path.getsize(args.input)
        after = os.path.getsize(out_path)
        report.section("Output", [
            f"written: {out_path}",
            f"size: {human(before)} -> {human(after)}",
            f"officecli validate: {officecli_validate(out_path)}",
        ])

    _emit(report, args)
    return 0


def _extract_media(input_path: str, parts: List[str], dest_dir: str, report: Report):
    os.makedirs(dest_dir, exist_ok=True)
    written = []
    with zipfile.ZipFile(input_path, "r") as zf:
        for part in parts:
            out = os.path.join(dest_dir, basename(part))
            with zf.open(part, "r") as src, open(out, "wb") as dst:
                shutil.copyfileobj(src, dst, length=1024 * 1024)
            written.append(f"{basename(part)} ({human(os.path.getsize(out))})")
    report.section("Sidecar media folder", [f"dir: {dest_dir}"] + written +
                   ["KEEP this folder next to the .pptx — linked videos are not self-contained."])


def _emit(report: Report, args):
    if args.json:
        print(json.dumps(report.data, indent=2, default=str))
    else:
        print(report.render())


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
