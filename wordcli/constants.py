"""XML namespace URIs, tag constants, and heading regex."""

import re
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

# Fully qualified tag names
P_TAG = f"{{{W_NS}}}p"
R_TAG = f"{{{W_NS}}}r"
T_TAG = f"{{{W_NS}}}t"
DEL_TAG = f"{{{W_NS}}}del"
INS_TAG = f"{{{W_NS}}}ins"
DELTEXT_TAG = f"{{{W_NS}}}delText"
AUTHOR_ATTR = f"{{{W_NS}}}author"
DATE_ATTR = f"{{{W_NS}}}date"
ID_ATTR = f"{{{W_NS}}}id"
FOOTNOTE_TAG = f"{{{W_NS}}}footnote"
FOOTNOTE_REF_TAG = f"{{{W_NS}}}footnoteReference"
COMMENT_TAG = f"{{{W_NS}}}comment"
COMMENT_RANGE_START_TAG = f"{{{W_NS}}}commentRangeStart"
COMMENT_RANGE_END_TAG = f"{{{W_NS}}}commentRangeEnd"
COMMENT_REFERENCE_TAG = f"{{{W_NS}}}commentReference"
RSTYLE_TAG = f"{{{W_NS}}}rStyle"
INITIALS_ATTR = f"{{{W_NS}}}initials"
ANNOTATION_REF_TAG = f"{{{W_NS}}}annotationRef"
RPR_TAG = f"{{{W_NS}}}rPr"
PPR_TAG = f"{{{W_NS}}}pPr"
PSTYLE_TAG = f"{{{W_NS}}}pStyle"
VAL_ATTR = f"{{{W_NS}}}val"
TBL_TAG = f"{{{W_NS}}}tbl"
TR_TAG = f"{{{W_NS}}}tr"
TC_TAG = f"{{{W_NS}}}tc"
TCPR_TAG = f"{{{W_NS}}}tcPr"
GRIDSPAN_TAG = f"{{{W_NS}}}gridSpan"
VMERGE_TAG = f"{{{W_NS}}}vMerge"
BODY_TAG = f"{{{W_NS}}}body"
SECTPR_TAG = f"{{{W_NS}}}sectPr"
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"
BOOKMARK_START_TAG = f"{{{W_NS}}}bookmarkStart"
BOOKMARK_END_TAG = f"{{{W_NS}}}bookmarkEnd"
FLDCHAR_TAG = f"{{{W_NS}}}fldChar"
FLDCHARTYPE_ATTR = f"{{{W_NS}}}fldCharType"
INSTRTEXT_TAG = f"{{{W_NS}}}instrText"
NAME_ATTR = f"{{{W_NS}}}name"
HYPERLINK_TAG = f"{{{W_NS}}}hyperlink"
STYLE_TAG = f"{{{W_NS}}}style"
STYLE_ID_ATTR = f"{{{W_NS}}}styleId"
STYLE_TYPE_ATTR = f"{{{W_NS}}}type"
STYLE_NAME_TAG = f"{{{W_NS}}}name"
PPR_CHANGE_TAG = f"{{{W_NS}}}pPrChange"
RPR_CHANGE_TAG = f"{{{W_NS}}}rPrChange"
B_TAG = f"{{{W_NS}}}b"
I_TAG = f"{{{W_NS}}}i"
U_TAG = f"{{{W_NS}}}u"
STRIKE_TAG = f"{{{W_NS}}}strike"

DRAWING_TAG = f"{{{W_NS}}}drawing"

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RID_ATTR = f"{{{R_NS}}}id"

# Drawing/image namespaces
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

WP_INLINE_TAG = f"{{{WP_NS}}}inline"
WP_ANCHOR_TAG = f"{{{WP_NS}}}anchor"
WP_EXTENT_TAG = f"{{{WP_NS}}}extent"
WP_DOCPR_TAG = f"{{{WP_NS}}}docPr"
A_BLIP_TAG = f"{{{A_NS}}}blip"
R_EMBED_ATTR = f"{{{R_NS}}}embed"

# Register namespaces so ET.tostring() uses proper prefixes instead of ns0:
_ns_registered = set()


def _register_namespaces(raw_xml_bytes):
    """Extract and register all namespace prefixes from raw XML."""
    text = raw_xml_bytes.decode("utf-8", errors="ignore")
    for m in re.finditer(r'xmlns:(\w+)="([^"]+)"', text):
        prefix, uri = m.group(1), m.group(2)
        if (prefix, uri) not in _ns_registered:
            ET.register_namespace(prefix, uri)
            _ns_registered.add((prefix, uri))


# Heading style patterns (English and German)
HEADING_RE = re.compile(
    r"^(?:Heading|berschrift|Überschrift)(\d+)$", re.IGNORECASE
)
