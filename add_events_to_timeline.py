"""
Adds Q1 events to the existing Tiendanube timeline PPT.
Reads the original PPT, preserves style/background, and places events
chronologically along the ENE-MAR timeline with smart overlap avoidance.
"""

from datetime import datetime

from lxml import etree
from pptx import Presentation
from pptx.util import Pt, Emu

NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

FONT = "Plus Jakarta Sans"

TIMELINE_X_START = 730250
TIMELINE_X_END = 730250 + 7831800
TIMELINE_Y = 2593025

Q1_START = datetime(2025, 1, 1)
Q1_END = datetime(2025, 3, 31)
Q1_DAYS = (Q1_END - Q1_START).days

EVENTS = [
    ("2025-01-13", "Reel Belulucius"),
    ("2025-01-19", "Reel Juli Savioli"),
    ("2025-01-29", "E. Nube enero"),
    ("2025-02-02", "Contenido Grammys"),
    ("2025-02-05", "Inicio Escuelita de\nVerano Emprendedora"),
    ("2025-02-21", "Creafest"),
    ("2025-02-26", "NubeCommerce"),
    ("2025-02-27", "E. Nube febrero"),
    ("2025-03-10", "Caso de éxito\nJuan Valdéz"),
    ("2025-03-16", "Contenido Oscars"),
    ("2025-03-27", "Caso de éxito\nLucciano's"),
    ("2025-03-28", "E. Nube marzo\nespecial 1"),
    ("2025-03-31", "E. Nube marzo\nespecial 2"),
]

LABEL_WIDTH = 820000
MIN_SAME_SIDE_GAP = LABEL_WIDTH + 30000
SLIDE_WIDTH = 9144000
LABEL_MIN_X = LABEL_WIDTH // 2 + 50000
LABEL_MAX_X = SLIDE_WIDTH - LABEL_WIDTH // 2 - 50000


def date_to_x(date_str: str) -> int:
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    frac = (dt - Q1_START).days / Q1_DAYS
    return int(TIMELINE_X_START + frac * (TIMELINE_X_END - TIMELINE_X_START))


def resolve_overlaps(events_with_positions):
    """Nudge label centers on the same side so they don't overlap."""
    above = [(i, x) for i, (x, is_above) in enumerate(events_with_positions) if is_above]
    below = [(i, x) for i, (x, is_above) in enumerate(events_with_positions) if not is_above]

    result = dict(events_with_positions)

    for group in [above, below]:
        if len(group) < 2:
            continue

        centers = [x for _, x in group]
        indices = [i for i, _ in group]

        for _ in range(100):
            changed = False
            for k in range(len(centers) - 1):
                gap = centers[k + 1] - centers[k]
                if gap < MIN_SAME_SIDE_GAP:
                    shift = (MIN_SAME_SIDE_GAP - gap) // 2 + 1
                    centers[k] -= shift
                    centers[k + 1] += shift
                    changed = True

            for k in range(len(centers)):
                clamped = max(LABEL_MIN_X, min(LABEL_MAX_X, centers[k]))
                if clamped != centers[k]:
                    centers[k] = clamped
                    changed = True

            if not changed:
                break

        for k, idx in enumerate(indices):
            is_above = events_with_positions[idx][1]
            result[idx] = (centers[k], is_above)

    return result


def make_text_box_xml(shape_id, name, x, y, w, h, text, font_size_hundredths, bold=False, align="ctr"):
    bold_attr = "1" if bold else "0"
    lines = text.split("\n")

    paragraphs_xml = ""
    for line_text in lines:
        paragraphs_xml += f"""
            <a:p>
              <a:pPr indent="0" lvl="0" marL="0" marR="0" rtl="0" algn="{align}">
                <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
                <a:spcBef><a:spcPts val="0"/></a:spcBef>
                <a:spcAft><a:spcPts val="0"/></a:spcAft>
                <a:buNone/>
              </a:pPr>
              <a:r>
                <a:rPr b="{bold_attr}" lang="es-419" sz="{font_size_hundredths}" dirty="0">
                  <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
                  <a:latin typeface="{FONT}"/>
                  <a:ea typeface="{FONT}"/>
                  <a:cs typeface="{FONT}"/>
                  <a:sym typeface="{FONT}"/>
                </a:rPr>
                <a:t>{line_text}</a:t>
              </a:r>
            </a:p>"""

    xml_str = f"""
    <p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="{name}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{x}" y="{y}"/>
          <a:ext cx="{w}" cy="{h}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr anchorCtr="0" anchor="t" bIns="0" lIns="0" spcFirstLastPara="1" rIns="0" wrap="square" tIns="0">
          <a:spAutoFit/>
        </a:bodyPr>
        <a:lstStyle/>
        {paragraphs_xml}
      </p:txBody>
    </p:sp>
    """
    return etree.fromstring(xml_str)


def make_oval_xml(shape_id, name, cx, cy, radius):
    x = cx - radius
    y = cy - radius
    d = radius * 2
    xml_str = f"""
    <p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="{name}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{x}" y="{y}"/>
          <a:ext cx="{d}" cy="{d}"/>
        </a:xfrm>
        <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
    </p:sp>
    """
    return etree.fromstring(xml_str)


def make_connector_xml(shape_id, name, x1, y1, x2, y2):
    w = abs(x2 - x1) if x2 != x1 else 0
    h = abs(y2 - y1) if y2 != y1 else 0
    ox = min(x1, x2)
    oy = min(y1, y2)
    if w == 0:
        w = 1
    xml_str = f"""
    <p:cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvCxnSpPr>
        <p:cNvPr id="{shape_id}" name="{name}"/>
        <p:cNvCxnSpPr/>
        <p:nvPr/>
      </p:nvCxnSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{ox}" y="{oy}"/>
          <a:ext cx="{w}" cy="{h}"/>
        </a:xfrm>
        <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln cap="flat" cmpd="sng" w="12700">
          <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:round/>
          <a:headEnd type="none"/>
          <a:tailEnd type="none"/>
        </a:ln>
      </p:spPr>
    </p:cxnSp>
    """
    return etree.fromstring(xml_str)


def add_feb_label(sp_tree, shape_id_start):
    feb_x = date_to_x("2025-02-01")
    feb_label = make_text_box_xml(
        shape_id_start, "FEB_label",
        feb_x, 2244790, 6568200, 184800,
        "FEB", 1200, bold=True, align="l"
    )
    sp_tree.append(feb_label)
    return shape_id_start + 1


def add_events(sp_tree, shape_id_start):
    sid = shape_id_start
    dot_radius = Emu(Pt(4))

    date_label_height = 155000
    event_label_height = 500000
    vertical_stem = 200000

    raw_positions = []
    for i, (date_str, _title) in enumerate(EVENTS):
        cx = date_to_x(date_str)
        is_above = i % 2 == 0
        raw_positions.append((cx, is_above))

    nudged = resolve_overlaps(raw_positions)

    for i, (date_str, title) in enumerate(EVENTS):
        dot_cx = date_to_x(date_str)
        label_cx, is_above = nudged[i]

        dot = make_oval_xml(sid, f"dot_{i}", dot_cx, TIMELINE_Y, dot_radius)
        sp_tree.append(dot)
        sid += 1

        if is_above:
            stem_top = TIMELINE_Y - vertical_stem
            stem_bot = TIMELINE_Y - dot_radius
        else:
            stem_top = TIMELINE_Y + dot_radius
            stem_bot = TIMELINE_Y + vertical_stem

        stem = make_connector_xml(sid, f"stem_{i}", dot_cx, stem_top, dot_cx, stem_bot)
        sp_tree.append(stem)
        sid += 1

        dt = datetime.strptime(date_str, "%Y-%m-%d")
        month_names = {1: "ENE", 2: "FEB", 3: "MAR"}
        display_date = f"{dt.day}/{month_names[dt.month]}"

        text_x = label_cx - LABEL_WIDTH // 2

        if is_above:
            date_y = stem_top - date_label_height - 10000
            event_y = date_y - event_label_height
        else:
            date_y = stem_bot + 15000
            event_y = date_y + date_label_height + 10000

        date_box = make_text_box_xml(
            sid, f"date_{i}", text_x, date_y,
            LABEL_WIDTH, date_label_height,
            display_date, 800, bold=True, align="ctr"
        )
        sp_tree.append(date_box)
        sid += 1

        event_box = make_text_box_xml(
            sid, f"event_{i}", text_x, event_y,
            LABEL_WIDTH, event_label_height,
            title, 700, bold=False, align="ctr"
        )
        sp_tree.append(event_box)
        sid += 1

    return sid


def main():
    prs = Presentation("original_timeline.pptx")
    slide = prs.slides[0]

    sp_tree = slide._element.find(".//p:spTree", NSMAP)

    next_id = 100

    next_id = add_feb_label(sp_tree, next_id)
    next_id = add_events(sp_tree, next_id)

    output_path = "timeline_q1_con_eventos.pptx"
    prs.save(output_path)
    print(f"Saved to {output_path} with {len(EVENTS)} events added.")


if __name__ == "__main__":
    main()
