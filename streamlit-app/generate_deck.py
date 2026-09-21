#!/usr/bin/env python3
"""Generate a Sunday service .pptx from the Sunday_Service_Template.pptx.

Keeps the recurring slides (opening, worship header, prayers, doxology,
benediction, announcements header, closing) and rebuilds the weekly-changing
sections (宣召, psalm, hymns, scripture reading, sermon, announcements) by
cloning scaffold slides from the template and filling their text.

The songs/sermon/announcements sections are each either imported in full from
an external deck (songs_import + songs_file, sermon_import + sermon_file,
announcements_import + announcements_file) or compiled from the config
(see build_hymns_section / build_sermon_section /
build_announcements_section).  build_section_deck("hymns"|"sermon"|
"announcements", cfg) produces a standalone 16:9 deck for one section, ready
to be shared and later merged via the *_import flow.

Usage:
    python3 -m pip install python-pptx
    python3 generate_deck.py --config weekly_config.json
    python3 generate_deck.py --config weekly_config.json --dry-run  # print plan only

Config keys are documented in weekly_config.example.json.
"""
import argparse
import copy
import json
import os
import re
import shutil
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}

# 0-based slide index in the template for each scaffold the generator reuses.
SCAFFOLD = {
    "opening": 1,            # slide 2  welcome screen
    "call": 2,               # slide 3  宣召
    "psalm": 3,              # slide 4  psalm / responsive reading
    "worship_header": 4,     # slide 5  詩歌敬拜
    "hymn_title": 8,         # slide 9  hymn title card
    "hymn_verse": 9,         # slide 10 hymn verse / refrain
    "prayer_header": 20,     # slide 21 禱告
    "lords_prayer": 21,      # slide 22 主禱文
    "scripture_header": 22,  # slide 23 讀經
    "scripture_verse": 23,   # slide 24 傳經 verse slide
    "sermon_header": 26,     # slide 27 信息
    "sermon_transition": 27, # slide 28 亮金色的過場
    "sermon_title": 28,      # slide 29 講道題目
    "sermon_body": 32,       # slide 33 講道內容 (title + content placeholders)
    "announce_img": 47,      # slide 48 相片式通告
    "announce_text": 58,     # slide 59 文字式通告 (title + body + 2 photos)
    "doxology": 44,          # slide 45 三一頌
    "benediction": 45,       # slide 46 祝福
    "announce_header": 46,   # slide 47 歡迎 家事分享
    "closing": 60,           # slide 61 結束
}

# Section slides that can be replaced by a full-bleed image. The config can
# override the path with a key of the same name as the SCAFFOLD entry, e.g.
#   "worship_header": "media/worship.jpg"
# Falls back to the template duplicate when the file is missing.
IMAGE_SLIDES = {
    "worship_header": os.path.join("media", "worship.jpg"),
    "prayer_header": os.path.join("media", "prayer.jpg"),
    "lords_prayer": os.path.join("media", "lordsprayer.jpg"),
    "communion": os.path.join("media", "communion.jpg"),
    "apostles_creed": os.path.join("media", "apostles_creed.jpg"),
    "scripture_header": os.path.join("media", "scripture_reading.jpg"),
    "sermon_header": os.path.join("media", "sermon.jpg"),
    "response_header": os.path.join("media", "response_hymn.jpg"),
    "doxology": os.path.join("media", "doxology.jpg"),
    "benediction": os.path.join("media", "benediction.jpg"),
    "announce_header": os.path.join("media", "announcements.jpg"),
    "closing": os.path.join("media", "closing.jpg"),
}

# Which shape within a duplicated slide gets filled (0-based shape order).
SHAPE = {
    "call_text": 1,
    "psalm_text": 1,
    "hymn_title_text": 0,
    "hymn_title_source": 1,
    "hymn_body": 0,
    "scripture_text": 1,
    "sermon_title_text": 0,
    "sermon_title_meta": 1,
    "sermon_body_title": 0,
    "sermon_body_content": 1,
    "announce_title": 0,
    "announce_body": 1,
    "announce_pic": 0,
}


def clear_slides(pres):
    sldIdLst = pres.slides._sldIdLst
    part = pres.part
    for sld in list(sldIdLst):
        rid = sld.get(qn("r:id"))
        sldIdLst.remove(sld)
        if rid:
            try:
                part.drop_rel(rid)
            except (KeyError, ValueError):
                pass


def _relate(new_part, src_part, rel):
    try:
        new_rel = new_part.rels.get_or_add(rel.reltype, rel.target_part)
    except Exception:
        return None
    if hasattr(new_rel, "rId"):
        return new_rel.rId
    return str(new_rel)


def _remap_rels(new_part, src_part, el):
    for node in el.iter():
        for attr, val in list(node.attrib.items()):
            if not attr.startswith("{%s}" % R_NS):
                continue
            rel = src_part.rels.get(val)
            if rel is None:
                continue
            new_rid = _relate(new_part, src_part, rel)
            if new_rid:
                node.set(attr, new_rid)


def _copy_slide_meta(dest_el, src_el, remap):
    """Copy per-slide background + clrMapOvr/transition so the duplicated
    slide keeps the exact same background as its template source."""
    src_cSld = src_el.find(qn("p:cSld"))
    dest_cSld = dest_el.find(qn("p:cSld"))
    if src_cSld is not None and dest_cSld is not None:
        bg = src_cSld.find(qn("p:bg"))
        old = dest_cSld.find(qn("p:bg"))
        if old is not None:
            dest_cSld.remove(old)
        if bg is not None:
            new_bg = copy.deepcopy(bg)
            dest_cSld.insert(0, new_bg)
            remap.append(new_bg)
    for tag in ("p:clrMapOvr", "p:transition", "p:timing"):
        el = src_el.find(qn(tag))
        old = dest_el.find(qn(tag))
        if old is not None:
            dest_el.remove(old)
        if el is not None:
            new_el = copy.deepcopy(el)
            dest_el.append(new_el)
            remap.append(new_el)


def duplicate_slide(pres, source):
    slide = pres.slides.add_slide(source.slide_layout)
    tree = slide.shapes._spTree
    for shp in list(slide.shapes):
        tree.remove(shp._element)
    elements = []
    for shp in source.shapes:
        el = copy.deepcopy(shp._element)
        elements.append(el)
        tree.append(el)
    _copy_slide_meta(slide._element, source._element, elements)
    for el in elements:
        _remap_rels(slide.part, source.part, el)
    return slide


_SKIP_IMPORT_REL_TYPES = (RT.NOTES_SLIDE, RT.NOTES_MASTER,
                           RT.PRES_PROPS, RT.VIEW_PROPS)


def _partname_tmpl(src_part):
    """printf-style template for a fresh partname in the same folder/profile,
    e.g. /ppt/slides/slide7.xml -> /ppt/slides/slide%d.xml."""
    text = str(src_part.partname)
    m = re.match(r"^(.*?)(\d+)([^/]*)$", text)
    if m and m.group(1):
        return "%s%%d%s" % (m.group(1), m.group(3))
    if "." in text:
        stem, ext = text.rsplit(".", 1)
        return "%s-%%d.%s" % (stem, ext)
    return text + "-%d"


def _remap_part_rids(el, rmap):
    """Rewrite every relationship attribute so the XML references the fresh
    rIds held by the cloned part.  Unmapped refs are dropped so no dangling
    relationship survives to corrupt the package."""
    for node in el.iter():
        for attr, val in list(node.attrib.items()):
            if not attr.startswith("{%s}" % R_NS):
                continue
            if val in rmap:
                node.set(attr, rmap[val])
            else:
                del node.attrib[attr]


def _alloc_partname(pkg, src_part, used):
    """First free partname in the same folder/profile as `src_part`, tracked in
    `used`.  next_partname() alone is not enough: clones are not reachable from
    the package until their slide is appended, so consecutive allocations would
    otherwise collide."""
    tmpl = _partname_tmpl(src_part)
    existing = {str(p.partname) for p in pkg.iter_parts()}
    n = 1
    while True:
        cand = str(tmpl % n)
        if cand not in used and cand not in existing:
            used.add(cand)
            return PackURI(cand)
        n += 1


class _PackageCloner:
    """Clones slide parts (plus layout/master/theme/media) between packages."""

    def __init__(self, pres):
        self.pres = pres
        self.package = pres.part.package
        self.memo = {}
        self.used = set()

    def clone(self, src_part):
        if src_part in self.memo:
            return self.memo[src_part]
        partname = _alloc_partname(self.package, src_part, self.used)
        new_part = PartFactory(partname, src_part.content_type,
                               self.package, src_part.blob)
        self.memo[src_part] = new_part
        rmap = {}
        for rel in src_part.rels.values():
            if rel.reltype in _SKIP_IMPORT_REL_TYPES:
                continue
            if rel.is_external:
                try:
                    new_rid = new_part.rels.get_or_add_ext_rel(
                        rel.reltype, rel.target_ref)
                except Exception:
                    continue
            else:
                target = self.clone(rel.target_part)
                try:
                    new_rid = new_part.rels.get_or_add(rel.reltype, target)
                except Exception:
                    continue
            rmap[rel.rId] = new_rid
        if isinstance(new_part, XmlPart):
            _remap_part_rids(new_part._element, rmap)
        if src_part.content_type == CT.PML_SLIDE_MASTER:
            self._register_master(new_part)
        return new_part

    def _register_master(self, master_part):
        """A slide master only becomes valid once the presentation references
        it: add the relationship and a matching `p:sldMasterId` entry.
        Without this PowerPoint sees orphan masters and offers to repair."""
        part = self.pres.part
        rId = part.rels.get_or_add(RT.SLIDE_MASTER, master_part)
        lst = part._element.get_or_add_sldMasterIdLst()
        entry = lst._add_sldMasterId()
        entry.set(qn("r:id"), rId)
        entry.set("id", str(self._next_ml_id(master_part)))
        # Renumber this master's (cloned) layout ids too: a source deck can
        # reuse ids that already exist in the destination, and a collision in
        # the shared id space also makes PowerPoint offer to repair.
        layout_ids = master_part._element.sldLayoutIdLst
        if layout_ids is not None:
            for e in layout_ids:
                if e.get("id"):
                    e.set("id", str(self._next_ml_id(master_part)))

    def _next_ml_id(self, master_part):
        """Next free id in PowerPoint's shared sldMasterId/sldLayoutId space.

        PowerPoint treats both id kinds as one space (values >= 2147483648)
        and repairs the file when a new master id collides with an existing
        layout id.  So the max must be taken across the presentation's master
        list, every master's layout list, and this master's own cloned layouts
        (the latter are not reachable via `pres.slide_masters` yet)."""
        min_id = 2147483648
        used = set()

        def collect(lst):
            if lst is None:
                return
            for e in lst:
                val = e.get("id")
                if val:
                    used.add(int(val))

        collect(self.pres.part._element.sldMasterIdLst)
        for m in self.pres.slide_masters:
            collect(m.part._element.sldLayoutIdLst)
        collect(master_part._element.sldLayoutIdLst)
        return max(used | {min_id - 1}) + 1


def import_slide(pres, source, cloner):
    """Append a copy of `source` (from another presentation) onto `pres`,
    preserving its own slide part + layout/master/theme + media so the merged
    slide looks identical to the source (source formatting)."""
    slide_part = cloner.clone(source.part)
    rId = pres.part.rels.get_or_add(RT.SLIDE, slide_part)
    pres.slides._sldIdLst.add_sldId(rId)
    return slide_part


def _effective_rpr(p):
    """rPr of paragraph p's first non-empty run, cloning its formatting.

    Empty paragraphs return None so the filler inherits from a neighbouring
    content paragraph instead of the paragraph's spacer defaults."""
    for r in p.runs:
        if r.text.strip():
            rPr = r._r.find(qn("a:rPr"))
            if rPr is not None:
                return copy.deepcopy(rPr)
    return None


def _shape_default_rpr(shape):
    """Fallback rPr (tag normalized to a:rPr) from paragraph/list-style
    defaults, used only when no paragraph in the shape has any text."""
    tf = shape.text_frame
    paras = tf.paragraphs
    for p in paras:
        pPr = p._p.find(qn("a:pPr"))
        if pPr is not None:
            defRPr = pPr.find(qn("a:defRPr"))
            if defRPr is not None and len(p.runs) == 0:
                el = copy.deepcopy(defRPr)
                el.tag = qn("a:rPr")
                return el
    lst = tf._txBody.find(qn("a:lstStyle"))
    if lst is not None:
        node = lst.find(qn("a:lvl1pPr"))
        if node is not None:
            defRPr = node.find(qn("a:defRPr"))
            if defRPr is not None:
                el = copy.deepcopy(defRPr)
                el.tag = qn("a:rPr")
                return el
    return None


def _proto_rprs(shape):
    """One candidate rPr per original paragraph, None where the paragraph is
    empty. The filler picks the nearest non-empty paragraph's style for those."""
    return [_effective_rpr(p) for p in shape.text_frame.paragraphs]


def _style_run(r, rc, proto):
    """Style run `r` per config `rc`, inheriting typography from `proto` rPr.

    No rPr element is added when nothing needs to be set, so PowerPoint falls
    back to the paragraph/master default (matching the template)."""
    rPr = copy.deepcopy(proto) if proto is not None else None
    if rPr is not None and rPr.tag == qn("a:defRPr"):
        rPr.tag = qn("a:rPr")
    need = rPr is not None
    if rc.get("bold"):
        need = True
    if rc.get("italic"):
        need = True
    if rc.get("size"):
        need = True
    if rc.get("font"):
        need = True
    if rc.get("color"):
        need = True
    if need and rPr is None:
        rPr = OxmlElement("a:rPr")
    if rPr is not None:
        r._r.insert(0, rPr)
    if rc.get("bold"):
        rPr.set("b", "1")
    if rc.get("italic"):
        rPr.set("i", "1")
    if rc.get("size"):
        rPr.set("sz", str(int(round(rc["size"] * 100))))
    if rc.get("font"):
        latin = rPr.find(qn("a:latin"))
        if latin is None:
            latin = OxmlElement("a:latin")
            rPr.append(latin)
        latin.set("typeface", rc["font"])
    if rc.get("color"):
        for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill",
                    "a:pattFill", "a:grpFill"):
            for node in rPr.findall(qn(tag)):
                rPr.remove(node)
        fill = OxmlElement("a:solidFill")
        color = OxmlElement("a:srgbClr")
        color.set("val", rc["color"])
        fill.append(color)
        rPr.insert(0, fill)


def _set_paragraph(p, item, proto):
    if isinstance(item, str):
        item = {"runs": [{"text": item}]}
    if item.get("align") in ALIGN:
        p.alignment = ALIGN[item["align"]]
    if item.get("space_before") is not None:
        p.space_before = Pt(item["space_before"])
    if item.get("space_after") is not None:
        p.space_after = Pt(item["space_after"])
    if item.get("line_spacing") is not None:
        p.line_spacing = item["line_spacing"]
    for r in list(p.runs):
        p._p.remove(r._r)
    runs = item.get("runs")
    if not runs:
        runs = [{"text": item.get("text", "")}]
    for rc in runs:
        if isinstance(rc, str):
            rc = {"text": rc}
        r = p.add_run()
        r.text = rc.get("text", "")
        _style_run(r, rc, proto)


def set_text(shape, items):
    """Fill `shape` with `items`, mapping each onto a template content slot.

    Content slots are the template's non-empty paragraphs (in order), followed
    by the empty paragraphs after the last content paragraph. Empty spacer
    paragraphs *between* content paragraphs are kept untouched, so each new
    paragraph inherits the exact font size of the template slot it fills."""
    tf = shape.text_frame
    paras = list(tf.paragraphs)
    protos = _proto_rprs(shape)
    fallback = _shape_default_rpr(shape)

    slots = [i for i, p in enumerate(paras) if p.text.strip()]
    if slots:
        last = slots[-1]
        slots.extend(j for j in range(last + 1, len(paras))
                     if not paras[j].text.strip())
    else:
        slots = list(range(len(paras)))

    def pick(i):
        if i < len(protos) and protos[i] is not None:
            return protos[i]
        for j in range(i + 1, len(protos)):
            if protos[j] is not None:
                return protos[j]
        for j in range(len(protos) - 1, -1, -1):
            if protos[j] is not None:
                return protos[j]
        return fallback

    txBody = tf._txBody
    used = []
    for i, item in enumerate(items):
        if i < len(slots):
            idx = slots[i]
            p = paras[idx]
        else:
            idx = len(paras) + (i - len(slots))
            p = tf.add_paragraph()
        used.append(idx)
        _set_paragraph(p, item, pick(idx))

    cutoff = max(used) if used else -1
    if cutoff < len(paras) - 1:
        for j in range(len(paras) - 1, cutoff, -1):
            txBody.remove(paras[j]._p)


def _shape(slide, index):
    return list(slide.shapes)[index]


def _normalize_paragraphs(shape, algn="just"):
    """Strip per-paragraph layout leftovers (indents, blank spacers, level
    overrides) so every line shares the same uniform vertical rhythm."""
    txBody = shape.text_frame._txBody
    for p in list(shape.text_frame.paragraphs):
        if not p.text.strip():
            p._p.getparent().remove(p._p)
            continue
        pPr = p._p.find(qn("a:pPr"))
        if pPr is None:
            pPr = OxmlElement("a:pPr")
            p._p.insert(0, pPr)
        for attr in ("marL", "marR", "indent", "lvl", "defTabSz", "rtl",
                     "eaLnBrk", "hangingPunct", "latinLnBrk"):
            if attr in pPr.attrib:
                del pPr.attrib[attr]
        for ch in list(pPr):
            if ch.tag in (qn("a:lnSpc"), qn("a:spcBef"), qn("a:spcAft")):
                pPr.remove(ch)
        pPr.set("algn", algn)
        pPr.set("fontAlgn", "base")


def _stretch_to_background(shape, slide_width):
    """Make a shape span the full slide width, matching the full-bleed
    background image behind it."""
    shape.left = 0
    shape.width = slide_width


def _margin_textbox(shape, slide_width, margin_in):
    """Give a shape symmetric left/right insets on the slide."""
    m = int(margin_in * 914400)
    shape.left = m
    shape.width = slide_width - 2 * m


def _line_width_in(line, pt):
    """Approx rendered width in inches: CJK chars ~1em, ASCII ~0.5em."""
    ems = sum(1.0 if ord(c) > 0x2E80 else 0.5 for c in line)
    return ems * pt / 72.0


def fit_size(lines, max_pt, box_w, min_pt=36, inset=0.2):
    """Largest pt in [min_pt, max_pt] so no line wraps inside the box."""
    usable = max(box_w - inset, 0.5)
    for size in range(max_pt, min_pt - 1, -1):
        if all(_line_width_in(l, size) <= usable for l in lines):
            return size
    return min_pt


def _block_fit(lines, box_w, box_h, max_pt, min_pt=36, inset=0.2):
    """Largest pt in [min_pt, max_pt] so all lines fit both width (no wrap)
    and height (auto line spacing ≈1.25em).  For 詩歌 lyrics so short stanzas
    render as big as possible (cap 54pt)."""
    usable = max(box_w - inset, 0.5)
    for size in range(max_pt, min_pt - 1, -1):
        block_h = len(lines) * 1.25 * size / 72.0
        if block_h <= box_h and all(_line_width_in(l, size) <= usable
                                    for l in lines):
            return size
    return min_pt


def _set_solid_bg(slide, hexval):
    """Give the slide a solid-color <p:bg> (overrides layout/master art).

    Mirrors the template's proven form exactly: <p:bg><p:bgPr>
    <a:solidFill><a:srgbClr/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    placed as the FIRST child of <p:cSld> (schema order: bg before spTree).
    """
    cSld = slide._element.find(qn("p:cSld"))
    old = cSld.find(qn("p:bg"))
    if old is not None:
        cSld.remove(old)
    bg = OxmlElement("p:bg")
    bgPr = OxmlElement("p:bgPr")
    solid = OxmlElement("a:solidFill")
    rgb = OxmlElement("a:srgbClr")
    rgb.set("val", hexval)
    solid.append(rgb)
    bgPr.append(solid)
    bgPr.append(OxmlElement("a:effectLst"))
    bg.append(bgPr)
    cSld.insert(0, bg)


def _force_text_color(shape, hexval):
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            rPr = r._r.find(qn("a:rPr"))
            if rPr is None:
                rPr = OxmlElement("a:rPr")
                r._r.insert(0, rPr)
            for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill",
                        "a:pattFill", "a:grpFill"):
                for node in rPr.findall(qn(tag)):
                    rPr.remove(node)
            fill = OxmlElement("a:solidFill")
            color = OxmlElement("a:srgbClr")
            color.set("val", hexval)
            fill.append(color)
            rPr.insert(0, fill)


def style_blue(slide, text_index, color="1F3A6E"):
    """Strip the decorative art, paint the slide solid blue, force white text."""
    shapes = list(slide.shapes)
    keep = shapes[text_index]
    for i, sh in enumerate(shapes):
        if i != text_index:
            sh._element.getparent().remove(sh._element)
    _set_solid_bg(slide, color)
    _force_text_color(keep, "FFFFFF")


def remove_shapes(slide, indices):
    shapes = list(slide.shapes)
    for ix in sorted(indices, reverse=True):
        shapes[ix]._element.getparent().remove(shapes[ix]._element)


def replace_picture(slide, index, path):
    pic = list(slide.shapes)[index]
    try:
        image_part, rId = slide.part.get_or_add_image_part(path)
    except Exception as exc:
        raise SystemExit("failed to load image %s: %s" % (path, exc))
    blipFill = pic._element.find(qn("p:blipFill"))
    if blipFill is None:
        raise SystemExit("target shape is not a picture: %s" % path)
    blip = blipFill.find(qn("a:blip"))
    if blip is None:
        blip = OxmlElement("a:blip")
        blipFill.insert(0, blip)
    blip.set(qn("r:embed"), rId)


def _pics(slide):
    return [i for i, sh in enumerate(list(slide.shapes))
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]


def _stretch_pic(slide, index, slide_width, slide_height):
    pic = list(slide.shapes)[index]
    pic.left = 0
    pic.top = 0
    pic.width = slide_width
    pic.height = slide_height


def _set_font(shape, typeface):
    """Set latin + East-Asian typeface on every run of the shape."""
    for para in shape.text_frame.paragraphs:
        for r in para.runs:
            rPr = r._r.find(qn("a:rPr"))
            if rPr is None:
                rPr = OxmlElement("a:rPr")
                r._r.insert(0, rPr)
            for tag in ("a:latin", "a:ea", "a:cs"):
                node = rPr.find(qn(tag))
                if node is None:
                    node = OxmlElement(tag)
                    rPr.append(node)
                node.set("typeface", typeface)


def chunk(lines, max_lines):
    return [lines[i:i + max_lines] for i in range(0, len(lines), max_lines)]


# 經文（宣召／讀經）自動分頁 -------------------------------------------------
# Chinese line-break (禁則) constraints: a rendered line may END with an opening
# mark ／ no, it must not END with one; and it must not START with a closing
# mark (，．；：！？…）.  Ending a line with a comma is fine and idiomatic.
LINE_OPEN = "「（《〔『“〈"
LINE_CLOSE = "，。、；：！？」』）】》〉”’!?,.;:…—·"


def _reading_box_lines(pt, box_h_in=5.7):
    """How many lines fit in the reading textbox at `pt` (single spacing)."""
    line_h = pt * 1.2 / 72.0
    return max(1, int(box_h_in / line_h))


def _reading_break(text, start, end, n):
    """Nudge a character-count break so a line never ends on an opening mark,
    never starts on a closing mark, and never leaves a single character (e.g.
    `並` in `…馬念，並掃羅。`) dangling at the end of the line.  Prefer taking
    a few extra characters, else a few fewer; falls back to the original cut
    when no clean point is nearby."""
    for delta in range(0, 10):
        for e in (end + delta, end - delta):
            if start + 1 < e < n and text[e - 1] not in LINE_OPEN \
                    and text[e] not in LINE_CLOSE:
                return e
    return end


def _split_verses(text):
    """Split a passage into verse units at verse-number markers.  Handles both
    fetched forms — `1（…） 2 …` (number flush against the text) and earlier
    stored copies `1 （…） 2 …` (number + space + text).  Returns the units, or
    [] when the text has no verse markers (then the caller falls back to plain
    paragraph pagination)."""
    marker = re.compile(
        r"(?:^|\s)(\d{1,3}(?::\d{1,3})?)"
        r"(?=\s*[\u3400-\u9fff\u3000-\u303f"
        r"\uff01-\uff5e\u2018\u2019\u201c\u201d\u300a\u300b\u300c\u300d])")
    cuts = [m.start(1) for m in marker.finditer(text)]
    if not cuts:
        return []
    spans = list(zip([0] + cuts, cuts + [len(text)]))
    return [text[a:b].strip() for a, b in spans
            if text[a:b].strip()]


def _reading_pages(text, pt, box_w_in=11.67, inset=0.2, whole_verses=False):
    """Wrap continuous 經文 into pages of explicit lines.

    Capacity comes from the font size: a CJK glyph is about one em wide, so
    ``per_line = usable_width / (pt / 72)`` and ``max_lines`` fits the box
    height.  Paragraphs (blank-line separated, e.g. one per fetched book, each
    headed by its 書卷出處) are wrapped independently and never joined across
    a paragraph boundary.  Every break is nudged so a rendered line neither
    ends on an opening mark nor starts with a closing mark, and no single
    character is left dangling at the end of a line.

    With ``whole_verses`` set (讀經), page breaks fall only between verse
    markers when the passage has them — a single verse never spans two
    slides."""
    paras = [p for p in str(text).split("\n") if p.strip()]
    if not paras:
        return []
    usable = max(box_w_in - inset, 0.5)
    per_line = max(1, int(usable / (pt / 72.0)))
    max_lines = _reading_box_lines(pt)

    def wrap(joined):
        lines, i, n = [], 0, len(joined)
        while i < n:
            end = min(i + per_line, n)
            if end < n:
                end = _reading_break(joined, i, end, n)
            lines.append(joined[i:end])
            i = end
        return lines

    if whole_verses:
        # Group each book's heading together with its verse units, and keep
        # pages breaking only at verse boundaries (a verse never spans two
        # slides).  The heading is placed together with the first ensuing verse.
        blocks, cur = [], {"heading": None, "units": []}
        found = False
        for para in paras:
            if para[:1].isdigit():
                parts = _split_verses(para)
                if parts:
                    found = True
                    cur["units"].extend(parts)
                    continue
            if cur["units"] or cur["heading"]:
                blocks.append(cur)
            cur = {"heading": " ".join(para.split()), "units": []}
        if cur["units"] or cur["heading"]:
            blocks.append(cur)
        if found:
            pages, stack = [], []
            for blk in blocks:
                vs = [wrap(" ".join(u.split())) for u in blk["units"]]
                if blk["heading"]:
                    if stack:                 # each book starts a fresh slide
                        pages.append(stack)
                        stack = []
                    stack.append(blk["heading"])
                for vlines in vs:
                    if not vlines:
                        continue
                    if stack and len(stack) + len(vlines) > max_lines:
                        pages.append(stack)
                        stack = []
                    stack.extend(vlines)
                    if len(stack) >= max_lines:
                        pages.append(stack)
                        stack = []
            if stack:
                pages.append(stack)
            return pages

    pages, lines = [], []
    for para in paras:
        joined = " ".join(para.split())
        if not joined:
            continue
        for ln in wrap(joined):
            if len(lines) >= max_lines:
                pages.append(lines)
                lines = []
            lines.append(ln)
    if lines:
        pages.append(lines)
    return pages


def _reading_slides(pres, ref, text, font_size, ref_size, typeface,
                    whole_verses=False):
    """Render continuous 經文 onto as many fresh slides as it needs.

    宣召經文／讀經經文 use a darker navy background (READING_BG) than the
    default slide blue so the white text stands out more. 出處 is bold and
    appears only on the FIRST slide of the passage; later slides repurpose
    that vertical space for more 內容."""
    pages = _reading_pages(text, font_size, whole_verses=whole_verses)
    for i, page in enumerate(pages):
        slide = _new_slide(pres, color=READING_BG)
        _fresh_reading_slide(slide, ref, page, font_size, ref_size, typeface,
                             ref_only=(i == 0))


def as_paragraphs(lines, size=None):
    if isinstance(lines, str):
        lines = [lines]
    runs = []
    for line in lines:
        rc = {"text": line}
        if size:
            rc["size"] = size
        runs.append(rc)
    return [{"runs": [rc]} for rc in runs]


def _plain_runs(shape, size=None):
    """Strip leftover typography from content runs copied off template slots
    (underline, bold, italic, negative char-spacing, baseline shift, …) and
    equalize their size."""
    FORBIDDEN = ("u", "b", "i", "spc", "baseline", "cap", "strike", "kern",
                 "lit", "lig", "compact", "forceClr", "poorKerning")
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            rPr = r._r.find(qn("a:rPr"))
            if rPr is None:
                if size is None:
                    continue
                rPr = OxmlElement("a:rPr")
                r._r.insert(0, rPr)
            for attr in FORBIDDEN:
                if attr in rPr.attrib:
                    del rPr.attrib[attr]
            for tag in ("a:uLnTx", "a:uLn", "a:uFill", "a:highlight",
                        "a:effectLst", "a:effectDag"):
                for node in rPr.findall(qn(tag)):
                    rPr.remove(node)
            if size is not None:
                rPr.set("sz", str(int(round(size * 100))))
        end = p._p.find(qn("a:endParaRPr"))
        if end is not None:
            for attr in FORBIDDEN:
                if attr in end.attrib:
                    del end.attrib[attr]
            if size is not None:
                end.set("sz", str(int(round(size * 100))))


def _split_ref_box(slide, ref_text, ref_size, content_shape,
                   slide_width, slide_height):
    """Lay REF（出處）in its own header textbox and pin the content box
    to the area below it, so the two never share geometry."""
    mar = int(0.83 * 914400)
    top_content = int(1.5 * 914400)
    content_shape.left = mar
    content_shape.top = top_content
    content_shape.width = slide_width - 2 * mar
    content_shape.height = max(slide_height - top_content - int(0.3 * 914400),
                               int(2 * 914400))
    box = slide.shapes.add_textbox(mar, int(0.3 * 914400),
                                   slide_width - 2 * mar, int(1.0 * 914400))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = ref_text
    r.font.size = Pt(ref_size)
    r.font.color.rgb = RGBColor.from_string("FFFFFF")
    return box


def _blank_layout(pres):
    for layout in pres.slide_layouts:
        if layout.name and layout.name.strip().lower() == "blank":
            return layout
    for layout in pres.slide_layouts:
        if not list(layout.placeholders):
            return layout
    return pres.slide_layouts[0]


def _fresh_textbox(slide, left_in, top_in, width_in, height_in,
                   lines, size, typeface, algn, shadow=False, bold=False):
    """Add a brand-new textbox from scratch (no template inheritance)."""
    tb = slide.shapes.add_textbox(int(left_in * 914400),
                                  int(top_in * 914400),
                                  int(width_in * 914400),
                                  int(height_in * 914400))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = algn
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
        if bold:
            r.font.bold = True
        if shadow:
            _run_shadow(r)
    _set_font(tb, typeface)
    return tb


def _run_shadow(run, color="000000", alpha_pct=90, blur_in=0.80, dist_in=0.50):
    """Text shadow on a run — the same <a:effectLst><a:outerShdw> XML that
    PowerPoint itself writes for text shadows, so Microsoft PowerPoint,
    ONLYOFFICE and LibreOffice render it identically."""
    rPr = run._r.get_or_add_rPr()
    for node in rPr.findall(qn("a:effectLst")):
        rPr.remove(node)
    eff = OxmlElement("a:effectLst")
    sh = OxmlElement("a:outerShdw")
    sh.set("blurRad", str(int(blur_in * 63500)))
    sh.set("dist", str(int(dist_in * 63500)))
    sh.set("dir", "5400000")
    sh.set("algn", "tl")
    sh.set("rotWithShape", "0")
    clr = OxmlElement("a:srgbClr")
    clr.set("val", color)
    al = OxmlElement("a:alpha")
    al.set("val", str(int(alpha_pct * 1000)))
    clr.append(al)
    sh.append(clr)
    eff.append(sh)
    before = None
    for tag in ("a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick",
                "a:hlinkMouseOver", "a:rtl", "a:extLst"):
        before = rPr.find(qn(tag))
        if before is not None:
            break
    if before is not None:
        before.addprevious(eff)
    else:
        rPr.append(eff)


def _fresh_reading_slide(slide, ref, lines, content_size, ref_size,
                         typeface, ref_only=True, bold_ref=True):
    top = 1.5
    if ref_only and ref:
        _fresh_textbox(slide, 0.83, 0.30, 11.67, 0.95, [ref],
                       ref_size, typeface, algn=PP_ALIGN.CENTER,
                       bold=bold_ref)
        top = 1.5
    else:
        top = 0.85
    _fresh_verse_box(slide, 0.83, top, 11.67, 6.35, lines,
                     content_size, typeface)


def _verse_sup_runs(line):
    """Split a reading line into runs so EVERY verse number (e.g. `1 `, `1在`,
    `10:2 `) becomes its own superscript run, wherever it appears in the
    line — leading, after sentence punctuation, or glued mid-line because
    verses were joined without a separator; the surrounding text stays
    normal. A number is treated as a verse marker only when it follows the
    line start or sentence punctuation, then a space, a CJK character, a
    full-width punctuation/bracket, or a letter — so plain numbers inside
    sentences (dates, counts) are left alone."""
    pat = re.compile(r"(?:^|[。；，」』）】])(\d{1,3}(?::\d{1,3})?)(?=[ 　]|[一-龥A-Za-z\uFF01-\uFF65])")
    out, pos = [], 0
    for m in pat.finditer(line):
        if m.start() > pos:
            out.append((line[pos:m.start(1)], False))
        out.append((m.group(1), True))
        pos = m.end(1)
    if pos < len(line):
        out.append((line[pos:], False))
    return out or [(line, False)]


def _fresh_verse_box(slide, left_in, top_in, width_in, height_in,
                     lines, size, typeface):
    """Reading-slide content textbox: leading verse numbers in superscript
    (same size as body, baseline-raised 30% like PowerPoint writes)."""
    tb = slide.shapes.add_textbox(int(left_in * 914400),
                                  int(top_in * 914400),
                                  int(width_in * 914400),
                                  int(height_in * 914400))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        for text, sup in _verse_sup_runs(line):
            r = p.add_run()
            r.text = text
            r.font.color.rgb = RGBColor.from_string("FFFFFF")
            if sup:
                r.font.size = Pt(size)
                rPr = r._r.get_or_add_rPr()
                rPr.set("baseline", "30000")
                if "vertAlign" in rPr.attrib:
                    del rPr.attrib["vertAlign"]
            else:
                r.font.size = Pt(size)
    _set_font(tb, typeface)
    return tb


def _image_cover(slide, path, width, height):
    pics = _pics(slide)
    if pics:
        idx = pics[0]
        replace_picture(slide, idx, path)
        _stretch_pic(slide, idx, width, height)
        remove_shapes(slide, [i for i in range(len(list(slide.shapes))) if i != idx])
        return
    slide.shapes.add_picture(path, 0, 0, width, height)
    remove_shapes(slide, list(range(len(list(slide.shapes)) - 1)))


READING_BG = "0B2A5B"          # 宣召經文／讀經經文 slides: darker navy bg


def _new_slide(pres, color="1F3A6E", bg=None):
    """Brand-new slide from the Blank layout (nothing duplicated from a
    template slide). Solid background color, or a full-bleed image."""
    slide = pres.slides.add_slide(_blank_layout(pres))
    for shp in list(slide.shapes):
        shp._element.getparent().remove(shp._element)
    if bg and os.path.isfile(bg):
        slide.shapes.add_picture(bg, 0, 0, pres.slide_width, pres.slide_height)
        return slide
    _set_solid_bg(slide, color)
    return slide


def _image_only(pres, path):
    """Full-bleed image slide (plain blue fallback when the file is missing)."""
    if path and os.path.isfile(path):
        return _new_slide(pres, bg=path)
    return _new_slide(pres)


def _section_geometry(cfg, pres):
    return (cfg.get("hymn_margin_in", 0.83),
            pres.slide_width / 914400 - 2 * cfg.get("hymn_margin_in", 0.83),
            cfg.get("call_font", "DFKai-SB"))


def build_hymns_section(pres, cfg, hymns):
    """詩歌敬拜：每首詩歌 = 標題卡 + 各段歌詞（＋副歌），背景 song_bg*.jpg。

    Prints nothing; appends slides straight onto `pres`.  Also used to compile
    a standalone songs deck (build_section_deck with section="hymns")."""
    margin, box_w, typeface = _section_geometry(cfg, pres)
    hymn_font_min = cfg.get("hymn_font_min", 36)
    hymn_font_max = cfg.get("hymn_font_max", 54)
    for i, hymn in enumerate(hymns):
        bg = hymn.get("bg") or ""
        if bg:
            if not os.path.isfile(bg):
                cand = os.path.join(os.getcwd(), bg)
                if os.path.isfile(cand):
                    bg = cand
            if not os.path.isfile(bg):
                bg = ""
        if not bg:
            bg = os.path.join(os.getcwd(), "media", "song_bg%d.jpg" % (i + 1))
        if not os.path.isfile(bg):
            bg = None

        # title card
        slide = _new_slide(pres, bg=bg)
        title = hymn["title"]
        tsize = fit_size([title], 60, box_w, min_pt=34)
        shadow = cfg.get("hymn_text_shadow", True)
        _fresh_textbox(slide, margin, 2.0, box_w, 1.4, [title],
                       tsize, typeface, PP_ALIGN.CENTER, shadow=shadow)
        if hymn.get("subtitle"):
            _fresh_textbox(slide, margin, 3.5, box_w, 0.9, [hymn["subtitle"]],
                           30, typeface, PP_ALIGN.CENTER, shadow=shadow)
        if hymn.get("source"):
            _fresh_textbox(slide, margin, 4.6, box_w, 0.7, [hymn["source"]],
                           36, typeface, PP_ALIGN.CENTER, shadow=shadow)

        refrain = hymn.get("refrain")
        repeat_refrain = bool(refrain) and hymn.get("refrain_after_every_verse",
                                                    True)

        def add_verse(lines, bg=bg):
            slide = _new_slide(pres, bg=bg)
            size = _block_fit(lines, box_w, 5.5, hymn_font_max,
                              min_pt=hymn_font_min)
            _fresh_textbox(slide, margin, 1.2, box_w, 5.5, lines,
                           size, typeface, PP_ALIGN.CENTER,
                           shadow=cfg.get("hymn_text_shadow", True))

        for verse_lines in hymn["verses"]:
            add_verse(verse_lines)
            if repeat_refrain:
                add_verse(refrain)
        if refrain and not repeat_refrain:
            add_verse(refrain)


def build_response_section(pres, cfg):
    """詩歌回應：單首回應詩歌（標題卡 + 各段歌詞 + 副歌），緊接信息之後。
    無上載流程，一律由內容現場編譯；背景沿用詩歌背景圖。"""
    response = cfg.get("response") or {}
    if not (response.get("title") or response.get("verses")):
        return 0
    build_hymns_section(pres, cfg, [{
        "title": response.get("title", ""),
        "subtitle": response.get("subtitle", ""),
        "source": response.get("source", ""),
        "refrain": response.get("refrain"),
        "refrain_after_every_verse": response.get(
            "refrain_after_every_verse", True),
        "verses": response.get("verses", []),
    }])
    return _response_slide_count(cfg)


def build_sermon_section(pres, cfg, sermon):
    """講道信息：亮金色過場 + 題目頁（題目/副題/經文/講員/日期）+ 內容頁。"""
    margin, box_w, typeface = _section_geometry(cfg, pres)
    if sermon.get("title"):
        _new_slide(pres, "C9A227")                   # 亮金色過場

        slide = _new_slide(pres)
        items = []
        items.append((sermon["title"], 40))
        if sermon.get("title_2"):
            items.append((sermon["title_2"], 32))
        if sermon.get("scripture"):
            items.append((sermon["scripture"],
                          sermon.get("ref_size", 28)))
        # centered title block
        top = 1.5
        for text, size in items:
            _fresh_textbox(slide, margin, top, box_w, 1.0, [text],
                           size, typeface, PP_ALIGN.CENTER)
            top += 1.25
        meta = "  ".join(x for x in (sermon.get("speaker", ""),
                                     sermon.get("date", "")) if x)
        if meta:
            _fresh_textbox(slide, margin, 5.6, box_w, 0.7, [meta],
                           20, typeface, PP_ALIGN.CENTER)

        for content in sermon.get("slides", []):
            body_slide = _new_slide(pres)
            if content.get("title"):
                _fresh_textbox(body_slide, margin, 0.35, box_w, 0.8,
                               [content["title"]], 30, typeface,
                               PP_ALIGN.CENTER)
            body = content.get("body") or []
            if body:
                nsize = fit_size(body, 40, box_w, min_pt=24)
                _fresh_textbox(body_slide, margin, 1.3, box_w, 5.9, body,
                               nsize, typeface, PP_ALIGN.LEFT)


def build_announcements_section(pres, cfg, announcements):
    """家事分享：圖片或文字（標題 + 內容）。"""
    margin, box_w, typeface = _section_geometry(cfg, pres)
    for a in announcements:
        if a.get("image"):
            _image_only(pres, a["image"])
            continue
        slide = _new_slide(pres)
        if a.get("title"):
            _fresh_textbox(slide, margin, 0.5, box_w, 1.0, [a["title"]],
                           36, typeface, PP_ALIGN.CENTER)
        body = a.get("body") or []
        if body:
            nsize = fit_size(body, 36, box_w, min_pt=24)
            _fresh_textbox(slide, margin, 1.6, box_w, 5.6, body,
                           nsize, typeface, PP_ALIGN.LEFT)


def _sermon_slide_count(cfg):
    """Slides build_sermon_section adds for `cfg` (no import)."""
    n = 0
    if cfg.get("sermon", {}).get("title"):
        n += 2                         # 過場 + 題目頁
        n += len(cfg.get("sermon", {}).get("slides", []))
    return n


def _response_slide_count(cfg):
    """Slides 詩歌回應 content contributes (title card + verses/refrains; no
    header), 0 when the section is empty."""
    response = cfg.get("response") or {}
    if not (response.get("title") or response.get("verses")):
        return 0
    verses = response.get("verses") or []
    n = 1                             # title card
    if response.get("refrain"):
        n += len(verses) if response.get(
            "refrain_after_every_verse", True) and verses else 1
    n += len(verses)
    return n


def build_section_deck(section, cfg, template=None):
    """Compile one section as its own standalone 16:9 deck:
    section = "hymns" | "sermon" | "announcements" | "scripture" | "response".  The result
    is exactly what build() would insert between the fixed header slides, so
    the section deck can be handed to someone else and merged by uploading it.
    ``scripture`` is a separate workflow: verse slides only (讀經), no service
    headers, for people who just want the 經文 on its own."""
    template = template or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "blank_16x9.pptx")
    pres = Presentation(template)
    clear_slides(pres)
    if section == "hymns":
        build_hymns_section(pres, cfg, cfg.get("hymns", []))
    elif section == "sermon":
        build_sermon_section(pres, cfg, cfg.get("sermon", {}))
    elif section == "announcements":
        build_announcements_section(pres, cfg, cfg.get("announcements", []))
    elif section == "scripture":
        sc = cfg.get("scripture", {}) or {}
        if sc.get("verses"):
            typeface = cfg.get("call_font", "DFKai-SB")
            _reading_slides(pres, sc.get("ref"), " ".join(sc["verses"]),
                            max(44, int(sc.get("font_size")
                                        or cfg.get("scripture_font_size", 44))),
                            sc.get("ref_size", 48), typeface,
                            whole_verses=True)
    elif section == "response":
        build_response_section(pres, cfg)
    else:
        raise ValueError("unknown section: %r (use hymns|sermon|"
                         "announcements|scripture)" % section)
    return pres


def build(pres, cfg):
    clear_slides(pres)
    songs_path = None      # external songs pptx (songs_import), if used
    sermons_path = None    # external sermon pptx (sermon_import), if used
    announcements_path = None  # external announcements pptx, if used
    offering_path = None   # external offering pptx (offering_import), if used
    response_path = None   # external response pptx (response_import), if used

    def image_path(name):
        val = cfg.get(name)
        if not isinstance(val, str) or not val.strip():
            val = IMAGE_SLIDES.get(name)
        if name == "doxology" and not val:
            val = cfg.get("doxolog") or IMAGE_SLIDES.get("doxology")
        if not val:
            return None
        p = val
        if not os.path.isabs(p):
            p = os.path.join(os.getcwd(), p)
        return p if os.path.isfile(p) else None

    def add_image(name):
        return _image_only(pres, image_path(name))

    psalm = cfg.get("psalm", {})
    scripture = cfg.get("scripture", {})
    sermon = cfg.get("sermon", {})
    announcements = cfg.get("announcements", [])
    hymns = cfg.get("hymns", [])
    typeface = cfg.get("call_font", "DFKai-SB")

    # opening（welcome image）
    welcome = cfg.get("welcome_image") or os.path.join(
        os.getcwd(), "media", "welcome.jpg")
    _image_only(pres, welcome)

    # call（宣召 image）
    call = cfg.get("call_image") or os.path.join(
        os.getcwd(), "media", "call_to_worship.jpg")
    _image_only(pres, call)

    # psalm（宣召經文，自動分頁；字型下限 44pt）
    if psalm.get("verses"):
        psalm_font = max(44, int(psalm.get("font_size")
                                 or cfg.get("psalm_font_size", 44)))
        psalm_ref = psalm.get("ref_size", 48)
        _reading_slides(pres, psalm.get("ref"), " ".join(psalm["verses"]),
                        psalm_font, psalm_ref, typeface, whole_verses=True)

    # 詩歌敬拜 header
    add_image("worship_header")

    # hymns: normally generated here; with songs_import the complete songs
    # pptx (songs_slides_<date>.pptx) is imported as-is instead.
    if cfg.get("songs_import"):
        songs_path = locate_songs_file(cfg)
        if songs_path:
            try:
                if import_deck(pres, songs_path, "songs") < 1:
                    print("warning: songs pptx has no slides (%s), "
                          "generating hymns from config" % songs_path)
                    songs_path = None
            except Exception as exc:
                print("warning: could not import %s (%s), generating hymns from config"
                      % (songs_path, exc))
                songs_path = None
        if not songs_path:
            print("warning: songs pptx not found (%s), generating hymns from config"
                  % cfg.get("songs_file", "songs_slides_*.pptx"))

    if not songs_path:
        # hymns（title card + verses + refrain，全部 fresh）
        build_hymns_section(pres, cfg, hymns)

    # 禱告、主禱文
    add_image("prayer_header")
    add_image("lords_prayer")

    # 獻詩（緊接禱告之後；上載的 pptx 原封不動合併，可選）
    if cfg.get("offering_import"):
        offering_path = locate_offering_file(cfg)
        if offering_path:
            try:
                if import_deck(pres, offering_path, "offering") < 1:
                    print("warning: offering pptx has no slides (%s), skipping"
                          % offering_path)
                    offering_path = None
            except Exception as exc:
                print("warning: could not import %s (%s)" % (offering_path, exc))
                offering_path = None
        if not offering_path:
            print("warning: offering pptx not found (%s), skipping"
                  % cfg.get("offering_file", "offering_*.pptx"))

    # 讀經 header + scripture verses（fresh reading slides）
    add_image("scripture_header")
    if scripture.get("verses"):
        scripture_font = max(44, int(scripture.get("font_size")
                                      or cfg.get("scripture_font_size", 44)))
        scripture_ref = scripture.get("ref_size", 48)
        _reading_slides(pres, scripture.get("ref"),
                        " ".join(scripture["verses"]),
                        scripture_font, scripture_ref, typeface,
                        whole_verses=True)

    # 信息 header（always follows 經文）＋（合併講道 pptx）或產生過場/題目/內容
    add_image("sermon_header")
    if cfg.get("sermon_import"):
        sermons_path = locate_sermon_file(cfg)
        if sermons_path:
            try:
                if import_deck(pres, sermons_path, "sermon") < 1:
                    print("warning: sermon pptx has no slides (%s), "
                          "generating sermon from config" % sermons_path)
                    sermons_path = None
            except Exception as exc:
                print("warning: could not import %s (%s), generating sermon from config"
                      % (sermons_path, exc))
                sermons_path = None
        if not sermons_path:
            print("warning: sermon pptx not found (%s), generating from config"
                  % cfg.get("sermon_file", "sermon_*.pptx"))
    if not sermons_path:
        build_sermon_section(pres, cfg, sermon)

    # 詩歌回應（緊接信息之後、聖餐之前；上載的 pptx 以檔案為準，否則由內容編譯）
    response_has_content = bool(cfg.get("response", {}).get("title") or
                                cfg.get("response", {}).get("verses"))
    if cfg.get("response_import"):
        response_path = locate_response_file(cfg)
        if response_path:
            try:
                if response_has_content:
                    add_image("response_header")
                if import_deck(pres, response_path, "response") < 1:
                    print("warning: response pptx has no slides (%s), "
                          "generating from config" % response_path)
                    response_path = None
            except Exception as exc:
                print("warning: could not import %s (%s)" % (response_path, exc))
                response_path = None
        if not response_path:
            print("warning: response pptx not found (%s), generating from config"
                  % cfg.get("response_file", "response_*.pptx"))
    elif response_has_content:
        add_image("response_header")
    if not response_path and response_has_content:
        build_response_section(pres, cfg)

    # 聖餐、使徒信經（first Sunday only）
    if cfg.get("communion"):
        add_image("communion")
        add_image("apostles_creed")

    # 三一頌
    add_image("doxology")

    # 祝福（緊接三一頌）
    add_image("benediction")

    # 歡迎/家事分享（announcements.jpg）
    add_image("announce_header")

    # announcements: imported external deck (announcements_<date>.pptx) or
    # freshly generated from the announcements list.
    if cfg.get("announcements_import"):
        announcements_path = locate_announcements_file(cfg)
        if announcements_path:
            try:
                if import_deck(pres, announcements_path, "announcements") < 1:
                    print("warning: announcements pptx has no slides (%s), "
                          "generating from config" % announcements_path)
                    announcements_path = None
            except Exception as exc:
                print("warning: could not import %s (%s), generating announcements from config"
                      % (announcements_path, exc))
                announcements_path = None
        if not announcements_path:
            print("warning: announcements pptx not found (%s), generating from config"
                  % cfg.get("announcements_file", "announcements_*.pptx"))

    if not announcements_path:
        build_announcements_section(pres, cfg, announcements)

    # 結束（最後一頁，接在家事分享之後）
    add_image("closing")


def plan(cfg):
    count = 1                          # opening
    count += 1                         # call
    psalm = cfg.get("psalm", {})
    if psalm.get("verses"):
        count += len(_reading_pages(
            " ".join(psalm["verses"]),
            max(44, int(psalm.get("font_size")
                         or cfg.get("psalm_font_size", 44))),
            whole_verses=True))
    count += 1                         # worship header
    count += _hymn_slide_count(cfg)    # hymns（或匯入 songs_slides pptx）
    count += 2                         # prayer header + lordsprayer
    if cfg.get("offering_import"):
        path = locate_offering_file(cfg)
        if path:
            try:
                count += len(Presentation(path).slides._sldIdLst)
            except Exception:
                pass
    count += 1                         # scripture header
    scripture = cfg.get("scripture", {})
    if scripture.get("verses"):
        count += len(_reading_pages(
            " ".join(scripture["verses"]),
            max(44, int(scripture.get("font_size")
                         or cfg.get("scripture_font_size", 44))),
            whole_verses=True))
    count += 1                         # sermon header（always）
    if cfg.get("sermon_import"):
        path = locate_sermon_file(cfg)
        if path:
            try:
                count += len(Presentation(path).slides._sldIdLst)
            except Exception:
                count += _sermon_slide_count(cfg)
        else:
            count += _sermon_slide_count(cfg)
    else:
        count += _sermon_slide_count(cfg)
    response_has_content = bool(cfg.get("response", {}).get("title") or
                                cfg.get("response", {}).get("verses"))
    if response_has_content:
        count += 1                         # response header
    if cfg.get("response_import"):
        path = locate_response_file(cfg)
        if path:
            try:
                count += len(Presentation(path).slides._sldIdLst)
            except Exception:
                count += _response_slide_count(cfg)
        else:
            count += _response_slide_count(cfg)
    else:
        count += _response_slide_count(cfg)
    if cfg.get("communion"):
        count += 2                         # communion + apostles creed
    count += 1                         # announce header（announcements）
    count += _announcements_slide_count(cfg)
    count += 2                         # doxology + benediction
    count += 1                         # closing
    return count


def _announcements_slide_count(cfg):
    """Slide count for announcements: imported deck when enabled (and found),
    else the number of announcement entries in the config."""
    if cfg.get("announcements_import"):
        path = locate_announcements_file(cfg)
        if path:
            try:
                return len(Presentation(path).slides._sldIdLst)
            except Exception:
                pass
    return len(cfg.get("announcements", []))


def load_config(path):
    with open(path, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg.setdefault("template", os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "blank_16x9.pptx"))
    cfg["template"] = os.path.expanduser(cfg["template"])
    if "output" not in cfg:
        date = cfg.get("date", "new")
        cfg["output"] = "Sunday_Service_%s.pptx" % date
    return cfg


DOWNLOADS = os.path.expanduser("~/Downloads")


def _locate_pptx(cfg, file_key, default_name, fallback_prefix):
    """Locate an external deck pptx.

    Preference: absolute path from config -> exact filename in the project
    folder or ~/Downloads -> newest <fallback_prefix>*.pptx in either location.
    """
    name = cfg.get(file_key) or default_name
    if isinstance(name, str) and name.strip():
        name = name.strip()
    if os.path.isabs(name) and os.path.isfile(name):
        return name
    here = os.path.dirname(os.path.abspath(__file__))
    for d in (here, DOWNLOADS):
        p = os.path.normpath(os.path.join(d, name))
        if os.path.isfile(p):
            return p
    cands = []
    for d in (here, DOWNLOADS):
        if not os.path.isdir(d):
            continue
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for fn in names:
            if fn.lower().endswith(".pptx") and fn.lower().startswith(fallback_prefix):
                cands.append(os.path.normpath(os.path.join(d, fn)))
    if cands:
        return max(cands, key=os.path.getmtime)
    return None


def locate_songs_file(cfg):
    date = cfg.get("date", "")
    return _locate_pptx(cfg, "songs_file",
                        "songs_slides_%s.pptx" % date if date else "songs_slides.pptx",
                        "songs")


def locate_sermon_file(cfg):
    date = cfg.get("date", "")
    return _locate_pptx(cfg, "sermon_file",
                        "sermon_%s.pptx" % date if date else "sermon.pptx",
                        "sermon")


def locate_announcements_file(cfg):
    date = cfg.get("date", "")
    return _locate_pptx(cfg, "announcements_file",
                        "announcements_%s.pptx" % date if date else "announcements.pptx",
                        "announcements")


def locate_offering_file(cfg):
    date = cfg.get("date", "")
    return _locate_pptx(cfg, "offering_file",
                        "offering_%s.pptx" % date if date else "offering.pptx",
                        "offering")


def locate_response_file(cfg):
    date = cfg.get("date", "")
    return _locate_pptx(cfg, "response_file",
                        "response_%s.pptx" % date if date else "response.pptx",
                        "response")


def import_deck(pres, path, label):
    """Import every slide of `path` into `pres`. Returns the slide count."""
    deck = Presentation(path)
    cloner = _PackageCloner(pres)
    for source_slide in deck.slides:
        import_slide(pres, source_slide, cloner)
    count = len(deck.slides._sldIdLst)
    print("imported %s slides (%d) from %s" % (label, count, path))
    return count


def _hymn_slide_count(cfg):
    """Number of slides the hymns section contributes: the imported songs
    deck's slide count when songs_import is on (and found), else the count
    this generator would produce from cfg["hymns"]."""
    if cfg.get("songs_import"):
        songs_path = locate_songs_file(cfg)
        if songs_path:
            try:
                return len(Presentation(songs_path).slides._sldIdLst)
            except Exception:
                pass
    n = 0
    for hymn in cfg.get("hymns", []):
        n += 1
        verses = hymn["verses"]
        repeats = 0
        if hymn.get("refrain") and hymn.get("refrain_after_every_verse", True):
            repeats = len(verses)
        elif hymn.get("refrain"):
            repeats = 1
        n += len(verses) + repeats
    return n


def pick_weekly_config(here):
    """Return the weekly config to build from.

    With no --config, take the newest weekly*.json among the project folder
    and ~/Downloads (the app's ＋ 製作 PPTX downloads land there). If a
    download is newer, copy it into the project folder as weekly_config.json
    so the app export is used automatically — no manual copy needed.
    """
    local = os.path.join(here, "weekly_config.json")
    candidates = []
    seen = set()
    for d in (here, DOWNLOADS):
        if not os.path.isdir(d):
            continue
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for fn in names:
            lfn = fn.lower()
            if not lfn.endswith(".json") or "example" in lfn:
                continue
            if not lfn.startswith("weekly"):
                continue
            full = os.path.normpath(os.path.join(d, fn))
            if full in seen or not os.path.isfile(full):
                continue
            seen.add(full)
            candidates.append(full)
    if not candidates:
        return None
    best = max(candidates, key=os.path.getmtime)
    if best != local and os.path.exists(best):
        try:
            shutil.copyfile(best, local)
            print("picked %s (%s) -> %s"
                  % (os.path.basename(best),
                     os.path.dirname(best),
                     os.path.basename(local)))
        except OSError as e:
            print("warning: could not copy %s: %s" % (best, e))
            local = best
    return local


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--config", help="weekly config json (default: newest "
                    "weekly*.json in the folder or ~/Downloads)")
    ap.add_argument("--template")
    ap.add_argument("--output")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--copy-only", action="store_true",
                    help="copy the template unchanged instead of rebuilding")
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    config_path = args.config or pick_weekly_config(here)
    if not config_path or not os.path.exists(config_path):
        sys.exit("no weekly config found (use --config or export from the app)")

    cfg = load_config(config_path)
    if args.template:
        cfg["template"] = args.template
    if args.output:
        cfg["output"] = args.output

    if args.dry_run:
        print("template : %s" % cfg["template"])
        print("output   : %s" % cfg["output"])
        print("slides   : %d" % plan(cfg))
        return

    if not os.path.exists(cfg["template"]):
        sys.exit("template not found: %s" % cfg["template"])

    if args.copy_only:
        shutil.copyfile(cfg["template"], cfg["output"])
        print("copied template to %s (unchanged)" % cfg["output"])
        return

    pres = Presentation(cfg["template"])
    build(pres, cfg)
    pres.save(cfg["output"])
    print("wrote %s (%d slides)" % (cfg["output"], len(pres.slides._sldIdLst)))


if __name__ == "__main__":
    main()