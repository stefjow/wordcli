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
