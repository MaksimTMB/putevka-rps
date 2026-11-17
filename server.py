import os, sys
if sys.stdout is None: sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w', encoding='utf-8')
import os, io, shutil, re, csv, base64, json, datetime, zipfile, urllib.parse, urllib.request, hashlib
import subprocess, threading, time
from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "v8.1.2"

app = FastAPI(title="ПУТЁВКА v8.1.2 — Форма № 3 (ОКУД 0345001)")

import sys
if getattr(sys, "frozen", False):
    BUNDLE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = BUNDLE_DIR

IS_FROZEN = getattr(sys, "frozen", False)

DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
DOCS_DIR = os.path.join(DATA_DIR, "documents")
os.makedirs(DOCS_DIR, exist_ok=True)
DB_FILE = os.path.join(DATA_DIR, "putevka_db.json")
INIT_FILE = os.path.join(DATA_DIR, "app_data.json")
if not os.path.exists(INIT_FILE) and os.path.exists(os.path.join(BUNDLE_DIR, "data", "app_data.json")):
    shutil.copy2(os.path.join(BUNDLE_DIR, "data", "app_data.json"), INIT_FILE)
if not os.path.exists(DB_FILE) and os.path.exists(INIT_FILE):
    shutil.copy2(INIT_FILE, DB_FILE)


def load_valid_state():
    for path in (DB_FILE, INIT_FILE):
        if os.path.exists(path) and os.path.getsize(path) > 50:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and "profiles" in data:
                    return data
            except Exception:
                pass
    return {}


def save_valid_state(payload: dict):
    os.makedirs(os.path.dirname(DB_FILE), exist_ok=True)
    existing = load_valid_state()
    if "xls_schema" not in payload and "xls_schema" in existing:
        payload["xls_schema"] = existing["xls_schema"]
    if "dwg_previews" not in payload and "dwg_previews" in existing:
        payload["dwg_previews"] = existing["dwg_previews"]
    if "adminSetupCompleted" not in payload and "adminSetupCompleted" in existing:
        payload["adminSetupCompleted"] = existing["adminSetupCompleted"]
    if "headerTemplates" not in payload and "headerTemplates" in existing:
        payload["headerTemplates"] = existing["headerTemplates"]
    payload["appVersion"] = APP_VERSION
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def xml_esc(s):
    return (
        str(s if s is not None else "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def get_front_overrides(prof: dict, day_obj: dict) -> dict:
    d_str, fg, ref = day_obj.get("date_ru", ""), str(day_obj.get("fuel_grade") or prof.get("defaultFuelGrade") or "92"), float(day_obj.get("refueled") or 0)
    legs = day_obj.get("legs") or []
    first_leg = legs[0] if legs else None
    last_leg = legs[-1] if legs else None
    dep_time = (first_leg.get("dep") if isinstance(first_leg, dict) and first_leg.get("dep") else (prof.get("depTimeDefault") or "9-00"))
    arr_time = (last_leg.get("arr") if isinstance(last_leg, dict) and last_leg.get("arr") else (prof.get("arrTimeDefault") or "18-00"))
    return {
        "1,18": prof.get("waybillSeries") or "", "1,22": day_obj.get("waybill_num") or prof.get("waybillNumberPrefix") or "", "3,7": d_str, "3,13": d_str, "4,22": prof.get("okud") or "0345001", "5,4": prof.get("orgFull") or "", "5,22": prof.get("okpo") or "",
        "8,5": prof.get("carModel") or "", "9,8": prof.get("carPlate") or "", "9,22": prof.get("garageNumber") or "", "10,2": prof.get("driverFull") or "", "10,22": prof.get("tabNumber") or "", "12,5": prof.get("license") or "", "12,20": prof.get("licenseClass") or "В",
        "13,6": prof.get("snils") or "", "14,6": prof.get("transportInfo") or "перевозка для собственных нужд", "17,0": prof.get("medOrgPre") or "", "28,0": prof.get("orgShort") or "", "28,22": prof.get("mechanic") or "", "34,23": prof.get("dispatcher") or "",
        "37,22": prof.get("driverShort") or "", "40,20": "АИ", "40,24": fg, "41,10": d_str, "42,10": dep_time, "44,9": prof.get("dispatcher") or "", "47,10": d_str, "48,10": arr_time, "50,9": prof.get("dispatcher") or "",
        "47,24": str(day_obj.get("refueled") or "") if ref > 0 else "", "49,18": str(day_obj.get("receipt_num") or "1") if ref > 0 else "", "50,24": str(day_obj.get("start_tank") or ""), "52,24": str(day_obj.get("end_tank") or ""), "54,24": str(day_obj.get("burned") or ""), "55,24": str(day_obj.get("burned") or ""),
        "62,22": prof.get("dispatcher") or "", "67,9": prof.get("driverShort") or "", "68,22": prof.get("driverShort") or ""}

def get_back_overrides(prof: dict, day_obj: dict) -> dict:
    wb_num = day_obj.get("waybill_num") or prof.get("waybillNumberPrefix") or ""
    ov = {
        "1,0": f"Оборотная сторона путевого листа легкового автомобиля № {wb_num} от {day_obj.get('date_ru', '')}",
        "24,7": day_obj.get("km", 0),
        "25,7": day_obj.get("km", 0),
        "29,5": prof.get("accountantRole") or "гл.бухгалтер",
        "29,17": prof.get("accountant") or "",
    }
    for rr in range(6, 19):
        ov[f"{rr},0"] = ""
        ov[f"{rr},5"] = ""
        ov[f"{rr},10"] = ""
        ov[f"{rr},15"] = ""
        ov[f"{rr},19"] = ""
        ov[f"{rr},23"] = ""
    for idx, leg in enumerate(day_obj.get("legs", [])):
        rr = 6 + idx
        if rr <= 18:
            ov[f"{rr},0"] = idx + 1
            ov[f"{rr},5"] = leg.get("from", "")
            ov[f"{rr},10"] = leg.get("to", "")
            ov[f"{rr},15"] = leg.get("dep", "")
            ov[f"{rr},19"] = leg.get("arr", "")
            ov[f"{rr},23"] = Number_or_str(leg.get("km", ""))
    return ov


def Number_or_str(v):
    try:
        return int(v)
    except Exception:
        return v


def build_word_table_xml(s_obj, styles, overrides, target_w_dxa=11220, target_h_dxa=15450, font_scale=1.05):
    ncols = 29
    nrows = 70 if s_obj["nrows"] >= 70 else 31
    raw_col_w = [float(s_obj["col_w"].get(str(c), 16.0)) for c in range(ncols)]
    sum_w = sum(raw_col_w)
    col_dxa = [max(40, int(round(w / sum_w * target_w_dxa))) for w in raw_col_w]
    col_dxa[-1] += target_w_dxa - sum(col_dxa)

    default_rh = 11.2 if nrows == 70 else 15.0
    raw_row_h = [float(s_obj["row_h"].get(str(r), default_rh)) for r in range(nrows)]
    sum_h = sum(raw_row_h)
    row_dxa = [max(40, int(round(h / sum_h * target_h_dxa))) for h in raw_row_h]

    row_lookup = {r: cells for r, cells in s_obj["rows"]}
    occupied = {}

    for r in range(nrows):
        cells = row_lookup.get(r, [])
        for c, orig_val, sid, ma, md in cells:
            if c >= ncols:
                continue
            if c + ma >= ncols:
                ma = ncols - 1 - c
            if r + md >= nrows:
                md = nrows - 1 - r
            key = f"{r},{c}"
            val = overrides[key] if key in overrides else orig_val
            occupied[(r, c)] = ("origin", c, val, sid, ma, md)
            for dr in range(0, md + 1):
                for dc in range(0, ma + 1):
                    if dr == 0 and dc == 0:
                        continue
                    elif dr > 0 and dc == 0:
                        occupied[(r + dr, c)] = ("vmerge_cont", r, c, sid, ma)
                    else:
                        occupied[(r + dr, c + dc)] = ("hmerge_cont",)

    b_map = {
        0: ("nil", 0),
        1: ("single", 4),
        2: ("single", 12),
        3: ("dashed", 4),
        4: ("dotted", 4),
        5: ("single", 18),
        6: ("double", 6),
        7: ("single", 4),
    }
    h_map = {0: "left", 1: "left", 2: "center", 3: "right", 4: "left", 5: "both", 6: "center"}
    v_map = {0: "top", 1: "center", 2: "bottom", 3: "center"}

    def border_xml(tb, rb, bb, lb):
        if not (tb or rb or bb or lb):
            return ""
        out = "<w:tcBorders>"
        for tag, code in (("top", tb), ("left", lb), ("bottom", bb), ("right", rb)):
            if code and code in b_map and b_map[code][0] != "nil":
                btype, bsz = b_map[code]
                out += f'<w:{tag} w:val="{btype}" w:sz="{bsz}" w:space="0" w:color="000000"/>'
        out += "</w:tcBorders>"
        return out

    xml = [
        "<w:tbl><w:tblPr>",
        f'<w:tblW w:w="{target_w_dxa}" w:type="dxa"/>',
        '<w:jc w:val="center"/><w:tblLayout w:type="fixed"/>',
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="15" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar>',
        "</w:tblPr><w:tblGrid>",
    ]
    for w in col_dxa:
        xml.append(f'<w:gridCol w:w="{w}"/>')
    xml.append("</w:tblGrid>")

    for r in range(nrows):
        rh = row_dxa[r]
        xml.append(f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="{rh}" w:hRule="exact"/></w:trPr>')
        c = 0
        while c < ncols:
            info = occupied.get((r, c))
            if info is None:
                w = col_dxa[c]
                xml.append(
                    f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="auto"/></w:pPr></w:p></w:tc>'
                )
                c += 1
            elif info[0] == "hmerge_cont":
                c += 1
            elif info[0] == "vmerge_cont":
                _, orig_r, orig_c, sid, ma = info
                span = ma + 1
                w = sum(col_dxa[c : c + span])
                st = styles[sid]
                tb, rb, bb, lb = st[0], st[1], st[2], st[3]
                xml.append("<w:tc><w:tcPr>")
                xml.append(f'<w:tcW w:w="{w}" w:type="dxa"/>')
                if span > 1:
                    xml.append(f'<w:gridSpan w:val="{span}"/>')
                xml.append("<w:vMerge/>")
                xml.append(border_xml(tb, rb, bb, lb))
                xml.append(
                    '</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="auto"/></w:pPr></w:p></w:tc>'
                )
                c += span
            elif info[0] == "origin":
                _, _, val, sid, ma, md = info
                span = ma + 1
                w = sum(col_dxa[c : c + span])
                st = styles[sid]
                tb, rb, bb, lb, fName, fPt, bold, italic, under, ha, va, wrap = st
                xml.append("<w:tc><w:tcPr>")
                xml.append(f'<w:tcW w:w="{w}" w:type="dxa"/>')
                if span > 1:
                    xml.append(f'<w:gridSpan w:val="{span}"/>')
                if md > 0:
                    xml.append('<w:vMerge w:val="restart"/>')
                xml.append(border_xml(tb, rb, bb, lb))
                xml.append(f'<w:vAlign w:val="{v_map.get(va, "center")}"/>')
                xml.append("</w:tcPr>")
                jc = h_map.get(ha, "left")
                hp = max(12, int(round((fPt or 8) * 2 * font_scale)))
                rpr = f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>'
                if bold:
                    rpr += "<w:b/>"
                if italic:
                    rpr += "<w:i/>"
                if under:
                    rpr += '<w:u w:val="single"/>'
                text_str = xml_esc(val).strip() if isinstance(val, str) else xml_esc(val)
                if text_str != "":
                    xml.append(
                        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="auto"/><w:jc w:val="{jc}"/></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text_str}</w:t></w:r></w:p>'
                    )
                else:
                    xml.append(
                        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="auto"/><w:jc w:val="{jc}"/><w:rPr>{rpr}</w:rPr></w:pPr></w:p>'
                    )
                xml.append("</w:tc>")
                c += span
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)


def generate_docx_bytes(prof: dict, days_list: list, schema: dict) -> bytes:
    sect_break = (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        '<w:rPr><w:sz w:val="2"/></w:rPr>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="280" w:right="340" w:bottom="280" w:left="340" w:header="0" w:footer="0" w:gutter="0"/>'
        "</w:sectPr></w:pPr></w:p>"
    )
    final_sect = (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="280" w:right="340" w:bottom="280" w:left="340" w:header="0" w:footer="0" w:gutter="0"/>'
        "</w:sectPr>"
    )
    body_parts = []
    total_days = len(days_list)
    for d_idx, day_obj in enumerate(days_list):
        front_tbl = build_word_table_xml(
            schema["s1"], schema["styles"], get_front_overrides(prof, day_obj),
            target_w_dxa=11220, target_h_dxa=15450, font_scale=1.05
        )
        back_tbl = build_word_table_xml(
            schema["s2"], schema["styles"], get_back_overrides(prof, day_obj),
            target_w_dxa=11220, target_h_dxa=15450, font_scale=1.10
        )
        body_parts.append(front_tbl)
        body_parts.append(sect_break)
        body_parts.append(back_tbl)
        if d_idx < total_days - 1:
            body_parts.append(sect_break)
        else:
            body_parts.append(final_sect)

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>" + "".join(body_parts) + "</w:body></w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", doc_xml)
    return buf.getvalue()


def populate_xlsx_sheet(ws, s_obj, styles, overrides, is_back=False):
    ncols = 29
    nrows = s_obj["nrows"]
    tot_w_pt = sum(float(s_obj["col_w"].get(str(c), 16.0)) for c in range(ncols))
    for c in range(ncols):
        w_pt = float(s_obj["col_w"].get(str(c), 16.0))
        col_letter = get_column_letter(c + 1)
        ws.column_dimensions[col_letter].width = round((w_pt / tot_w_pt) * 108.0, 2)

    tot_h_pt = sum(float(s_obj["row_h"].get(str(r), 12.0)) for r in range(nrows))
    target_h = 790.0
    for r in range(nrows):
        h_pt = float(s_obj["row_h"].get(str(r), 12.0))
        ws.row_dimensions[r + 1].height = round((h_pt / tot_h_pt) * target_h, 2)

    b_map = {
        0: None,
        1: Side(style="thin", color="000000"),
        2: Side(style="medium", color="000000"),
        3: Side(style="dashed", color="000000"),
        4: Side(style="dotted", color="000000"),
        5: Side(style="thick", color="000000"),
        6: Side(style="double", color="000000"),
        7: Side(style="thin", color="000000"),
    }
    h_map = {0: "left", 1: "left", 2: "center", 3: "right", 4: "fill", 5: "justify", 6: "centerContinuous"}
    v_map = {0: "top", 1: "center", 2: "bottom", 3: "justify"}

    font_cache = {}
    align_cache = {}
    border_cache = {}

    for r, cells in s_obj["rows"]:
        if r >= nrows:
            continue
        for c, orig_val, sid, ma, md in cells:
            if c >= ncols:
                continue
            if c + ma >= ncols:
                ma = ncols - 1 - c
            if r + md >= nrows:
                md = nrows - 1 - r
            key = f"{r},{c}"
            val = overrides[key] if key in overrides else orig_val
            cell = ws.cell(row=r + 1, column=c + 1)
            if is_back and r == 24 and c == 7:
                cell.value = "=SUM(X7:Z19)"
            elif is_back and r == 25 and c == 7:
                cell.value = "=SUM(X7:Z19)"
            elif val != "" and val is not None:
                cell.value = val

            st = styles[sid]
            tb, rb, bb, lb, fName, fPt, bold, italic, under, ha, va, wrap = st
            f_key = (fName, fPt, bold, italic, under)
            if f_key not in font_cache:
                font_cache[f_key] = Font(
                    name="Arial",
                    size=max(7, round((fPt or 8) * 1.05, 1)),
                    bold=bool(bold),
                    italic=bool(italic),
                    underline="single" if under else None,
                )
            a_key = (ha, va, wrap)
            if a_key not in align_cache:
                align_cache[a_key] = Alignment(
                    horizontal=h_map.get(ha, "left"),
                    vertical=v_map.get(va, "center"),
                    wrap_text=bool(wrap),
                )
            b_key = (tb, rb, bb, lb)
            if b_key not in border_cache:
                border_cache[b_key] = Border(
                    top=b_map.get(tb),
                    right=b_map.get(rb),
                    bottom=b_map.get(bb),
                    left=b_map.get(lb),
                )

            cell.font = font_cache[f_key]
            cell.alignment = align_cache[a_key]
            if ma > 0 or md > 0:
                ws.merge_cells(
                    start_row=r + 1, start_column=c + 1, end_row=r + 1 + md, end_column=c + 1 + ma
                )
                for rr in range(r + 1, r + 2 + md):
                    for cc in range(c + 1, c + 2 + ma):
                        mc = ws.cell(row=rr, column=cc)
                        mc.border = border_cache[b_key]
                        mc.font = font_cache[f_key]
                        mc.alignment = align_cache[a_key]
            else:
                cell.border = border_cache[b_key]

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True


def generate_xlsx_bytes(prof: dict, days_list: list, schema: dict) -> bytes:
    wb = openpyxl.Workbook()
    default_ws = wb.active
    for d_idx, day_obj in enumerate(days_list):
        d_short = day_obj.get("date_ru", f"День_{d_idx+1}")[:5]
        ws_front = default_ws if d_idx == 0 else wb.create_sheet()
        ws_front.title = f"{d_short} Шапка"[:31]
        populate_xlsx_sheet(ws_front, schema["s1"], schema["styles"], get_front_overrides(prof, day_obj), is_back=False)

        ws_back = wb.create_sheet(title=f"{d_short} Маршрут"[:31])
        populate_xlsx_sheet(ws_back, schema["s2"], schema["styles"], get_back_overrides(prof, day_obj), is_back=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@app.get("/")
def index():
    # v7.3.5: в exe-версии страница ВСЕГДА из комплекта самого exe — версия страницы
    # совпадает с версией сервера «по построению», и «наполовину обновлённое» состояние
    # (новые файлы на диске при старом exe) больше не может показать жёлтый баннер.
    if IS_FROZEN:
        candidates = [
            os.path.join(BUNDLE_DIR, "static", "index.html"),
            os.path.join(BASE_DIR, "static", "index.html"),
        ]
    else:
        candidates = [
            os.path.join(BASE_DIR, "static", "index.html"),
            os.path.join(BUNDLE_DIR, "static", "index.html"),
        ]
    page_path = next((p for p in candidates if os.path.exists(p)), candidates[0])
    return FileResponse(
        page_path,
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


@app.get("/api/version")
def get_version():
    st = load_valid_state()
    return {
        "version": APP_VERSION,
        "release_date": "2026-11-17",
        "has_saved_db": os.path.exists(DB_FILE),
        "adminSetupCompleted": bool(st.get("adminSetupCompleted", False)),
    }


@app.get("/api/state")
def get_state():
    return load_valid_state()


@app.post("/api/state")
def save_state(payload: dict = Body(...)):
    save_valid_state(payload)
    return {
        "status": "ok",
        "saved_at": datetime.datetime.now().isoformat(),
        "version": APP_VERSION,
    }


def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


@app.post("/api/auth/setup")
def setup_admin(payload: dict = Body(...)):
    login = str(payload.get("login") or "").strip()
    password = str(payload.get("password") or "").strip()
    full_name = str(payload.get("fullName") or "Главный Администратор").strip()
    if not login or not password:
        raise HTTPException(status_code=400, detail="Логин и пароль обязательны")

    st = load_valid_state()
    users = st.get("users", [])
    admin_user = None
    for u in users:
        if u.get("role") == "admin":
            u["login"] = login
            u["password"] = _hash_password(password)
            u["fullName"] = full_name
            u["shortName"] = full_name
            u["active"] = True
            admin_user = u
            break
    if not admin_user:
        admin_user = {
            "id": "usr_admin",
            "login": login,
            "password": _hash_password(password),
            "role": "admin",
            "fullName": full_name,
            "shortName": full_name,
            "assignedProfileId": "ALL",
            "canEditSettings": True,
            "canEditFuel": True,
            "canExportXls": True,
            "active": True,
        }
        users.insert(0, admin_user)
    st["users"] = users
    st["adminSetupCompleted"] = True
    save_valid_state(st)
    return {"status": "ok", "user": admin_user}


@app.post("/api/auth/login")
def login_user(payload: dict = Body(...)):
    login = str(payload.get("login") or "").strip()
    password = str(payload.get("password") or "").strip()
    st = load_valid_state()
    users = st.get("users", [])
    for u in users:
        stored = u.get("password") or ""
        if u.get("login") == login and (stored == password or stored == _hash_password(password)):
            if u.get("active") is False:
                return {"status": "error", "message": "Учетная запись отключена администратором"}
            return {"status": "ok", "user": u}
    return {"status": "error", "message": "Неверный логин или пароль"}




RU_MONTH_MAP = {
    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "май": 5, "мая": 5, "июн": 6,
    "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12
}

AZS_KEYWORDS = [
    ("301", "Астраханская, 173"),
    ("470", "Астраханская, 39"),
    ("341", "Стрельцы"),
    ("381", "Базарная"),
    ("304", "Моршанское"),
    ("468", "Студенецкая"),
    ("383", "Никифоровская"),
    ("302", "Мичуринская"),
    ("305", "Киквидзе"),
    ("312", "Советская"),
    ("318", "Рассказовское"),
    ("325", "Посконкина"),
    ("326", "Строитель"),
]


def match_azs_from_text(text_str: str, azs_locs: list) -> str:
    s = (text_str or "").strip()
    if not s:
        return ""
    s_low = s.lower()
    # 1. Direct match by AZS number (e.g. №301, АЗС 470, 341)
    for num_code, street_kw in AZS_KEYWORDS:
        if re.search(r"(?:азс|ти|№|#|\b)" + num_code + r"\b", s_low) or street_kw.lower() in s_low:
            for loc in azs_locs:
                addr = loc.get("addr", "")
                if num_code in addr or street_kw.lower() in addr.lower():
                    return addr
    # 2. Substring match against known AZS locations
    for loc in azs_locs:
        addr = loc.get("addr", "")
        short_a = loc.get("short_addr", "")
        if addr and addr.lower() in s_low:
            return addr
        if short_a and len(short_a) > 6 and short_a.lower() in s_low:
            return addr
    return ""


def parse_time_cell_lenient(val):
    """Мягкое чтение времени из ячейки «Время»: стандартные и нестандартные форматы.

    Понимает: 8:42, 08:42:00, 8-42, 8.42, 8,42 (строка!), 8ч42м, 8 ч 42 мин,
    8 42, 0842 (ЧЧММ), а также Excel-время дробью (0.3625 = 8:42) и типы
    datetime.time / datetime.datetime.
    """
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, datetime.datetime):
        return f"{val.hour:02d}:{val.minute:02d}"
    if isinstance(val, datetime.time):
        return f"{val.hour:02d}:{val.minute:02d}"
    if isinstance(val, (int, float)):
        v = float(val)
        if 0 <= v < 1:  # Excel хранит время как дробную часть суток
            total_min = round(v * 24 * 60)
            if 0 <= total_min < 24 * 60:
                return f"{total_min // 60:02d}:{total_min % 60:02d}"
        return None
    s = str(val).strip().lower().replace("мск", "").strip()
    if not s:
        return None
    # 8:42:00 / 8:42 / 8-42 / 8.42 / 8,42
    m = re.match(r"^([01]?\d|2[0-3])[.,:\-;]([0-5]\d)(?::[0-5]\d)?$", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    # 8ч42м / 8 ч 42 мин / 8ч 42мин
    m = re.match(r"^([01]?\d|2[0-3])\s*ч\.?\s*([0-5]\d)\s*м", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    # 8 42 (время через пробел)
    m = re.match(r"^([01]?\d|2[0-3])\s+([0-5]\d)$", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    # 0842 (ЧЧММ)
    m = re.match(r"^([01]\d|2[0-3])([0-5]\d)$", s)
    if m:
        return f"{m.group(1)}:{m.group(2)}"
    return None


def parse_cell_datetime(val, default_year=2026, default_month=9):
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return (val.year, val.month, val.day, f"{val.hour:02d}:{val.minute:02d}")
    if isinstance(val, datetime.date):
        return (val.year, val.month, val.day, None)
    if isinstance(val, (int, float)) and 38000 <= float(val) <= 55000:
        base = datetime.datetime(1899, 12, 30)
        dt = base + datetime.timedelta(days=float(val))
        t_str = f"{dt.hour:02d}:{dt.minute:02d}" if (float(val) % 1) > 0.0005 else None
        return (dt.year, dt.month, dt.day, t_str)

    s = str(val).strip()
    if not s or len(s) < 4:
        return None
    # Skip date ranges like "01.09.2026 - 30.09.2026" or "с 01.09 по 30.09"
    if re.search(r"\d{2}\.\d{2}\.\d{2,4}\s*[-–—]\s*\d{2}\.\d{2}\.\d{2,4}", s):
        return None
    if "период" in s.lower() and (" по " in s.lower() or "-" in s):
        return None

    time_match = re.search(r"\b([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?\b", s)
    t_found = f"{int(time_match.group(1)):02d}:{time_match.group(2)}" if time_match else None

    # DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY
    m1 = re.search(r"\b(0?[1-9]|[12]\d|3[01])[./\-](0?[1-9]|1[0-2])[./\-](20[2-3]\d|[2-3]\d)\b", s)
    if m1:
        dd, mm, yy = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
        if yy < 100:
            yy += 2000
        return (yy, mm, dd, t_found)

    # YYYY-MM-DD
    m2 = re.search(r"\b(20[2-3]\d)[./\-](0?[1-9]|1[0-2])[./\-](0?[1-9]|[12]\d|3[01])\b", s)
    if m2:
        return (int(m2.group(1)), int(m2.group(2)), int(m2.group(3)), t_found)

    # Russian text month: "3 сентября 2025" or "03 сен"
    s_low = s.lower()
    for m_prefix, m_num in RU_MONTH_MAP.items():
        m3 = re.search(r"\b(0?[1-9]|[12]\d|3[01])\s+" + m_prefix + r"[а-я]*\s*(20[2-3]\d)?\b", s_low)
        if m3:
            dd = int(m3.group(1))
            yy = int(m3.group(2)) if m3.group(2) else default_year
            return (yy, m_num, dd, t_found)

    return None


def parse_float_cell(val):
    if val is None or isinstance(val, (bool, datetime.date, datetime.datetime, datetime.time)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace(" ", "").replace(",", ".")
    s = re.sub(r"(?:л\.?|руб\.?|р\.?)$", "", s, flags=re.IGNORECASE)
    if re.match(r"^[+-]?\d+(?:\.\d+)?$", s):
        try:
            return float(s)
        except Exception:
            return None
    return None


def extract_receipts_from_rows(rows: list, sheet_name: str, azs_locs: list, default_year=2026, default_month=9) -> list:
    ctx_year = default_year
    ctx_month = default_month

    # Detect context year/month from sheet_name or top rows
    header_scan_text = sheet_name.lower() + " "
    for r in rows[:12]:
        header_scan_text += " ".join(str(c or "") for c in r).lower() + " "
    ym_match = re.search(r"\b(20[2-3]\d)\b", header_scan_text)
    if ym_match:
        ctx_year = int(ym_match.group(1))
    for m_pref, m_num in RU_MONTH_MAP.items():
        if m_pref in header_scan_text:
            ctx_month = m_num
            break

    col_map = {}
    header_row_idx = -1

    for r_idx, row in enumerate(rows[:40]):
        texts = [str(c or "").strip().lower() for c in row]
        non_empty = [t for t in texts if t]
        if len(non_empty) < 2:
            continue
        cand = {}
        for c_idx, t in enumerate(texts):
            if not t:
                continue
            if any(w in t for w in ("дата и время", "дата транз", "дата чека", "дата пл", "рабочая дата")) or t in ("дата", "date", "период"):
                cand.setdefault("date", c_idx)
            elif any(w in t for w in ("время", "time", "час")) and "дата" not in t:
                cand.setdefault("time", c_idx)
            elif any(w in t for w in ("количество", "кол-во", "объем", "объём", "литр", "залито", "заправлено", "отпущено")) and not any(b in t for b in ("сумма", "стоим", "цена", "руб", "остаток")):
                cand.setdefault("liters", c_idx)
            elif any(w in t for w in ("товар", "услуга", "марка", "топливо", "гсм", "вид топл", "наименование", "продукт", "нефтепродукт")):
                cand.setdefault("fuel", c_idx)
            elif any(w in t for w in ("азс", "адрес", "терминал", "точка обслуж", "место", "станция", "локация", "номер ти")):
                cand.setdefault("azs", c_idx)
            elif any(w in t for w in ("операция", "вид опер", "тип опер", "направление", "дебет")):
                cand.setdefault("op", c_idx)
            elif any(w in t for w in ("цена", "тариф")):
                cand.setdefault("price", c_idx)
            elif any(w in t for w in ("сумма", "стоимость")):
                cand.setdefault("sum", c_idx)
            elif any(w in t for w in ("водитель", "держатель", "фио", "тс", "госномер", "машина")):
                cand.setdefault("driver", c_idx)
        if ("date" in cand and ("liters" in cand or "fuel" in cand or "azs" in cand)) or len(cand) >= 3:
            col_map = cand
            header_row_idx = r_idx
            break

    receipts = []
    start_r = header_row_idx + 1 if header_row_idx >= 0 else 0

    for r_idx in range(start_r, len(rows)):
        row = rows[r_idx]
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        row_str = " ".join(str(c or "").strip() for c in row)
        row_low = row_str.lower()

        if any(skip_w in row_low for skip_w in ("итого", "всего", "оборот за", "остаток на ", "отчет сформирован", "период с ")):
            continue
        if any(non_f in row_low for non_f in ("пополнение", "зачисление", "оплата счета", "платежное поручение", "мойка", "омыватель", "абонентск")):
            continue

        # 1. Extract Date & Time
        dt_res = None
        date_col_used = None
        if "date" in col_map and col_map["date"] < len(row):
            dt_res = parse_cell_datetime(row[col_map["date"]], ctx_year, ctx_month)
            if dt_res:
                date_col_used = col_map["date"]
        if not dt_res:
            for c_idx, cell in enumerate(row):
                dt_res = parse_cell_datetime(cell, ctx_year, ctx_month)
                if dt_res:
                    date_col_used = c_idx
                    break
        if not dt_res:
            continue

        yy, mm, dd, t_str = dt_res

        # Время чека: отдельная колонка «Время» — главный источник (нестандартные
        # форматы поддерживает parse_time_cell_lenient). Время, «зашитое» в ячейку
        # даты, используем только если колонки времени нет/она пустая.
        t_str = None
        if "time" in col_map and col_map["time"] < len(row):
            t_str = parse_time_cell_lenient(row[col_map["time"]])
        if not t_str:
            t_str = dt_res[3]
        if not t_str:
            for c_idx, cell in enumerate(row):
                if c_idx == date_col_used or cell is None:
                    continue
                if isinstance(cell, datetime.time):
                    t_str = f"{cell.hour:02d}:{cell.minute:02d}"
                    break
                s_c = str(cell).strip()
                if re.match(r"^([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?$", s_c):
                    p_t = s_c.split(":")
                    t_str = f"{int(p_t[0]):02d}:{p_t[1]}"
                    break
        if not t_str:
            t_str = "10:30"

        # 2. Check return / credit operation
        is_return = any(w in row_low for w in ("возврат", "кредит", "сторно", "отмена"))

        # 3. Extract Liters
        liters = None
        if "liters" in col_map and col_map["liters"] < len(row):
            lv = parse_float_cell(row[col_map["liters"]])
            if lv is not None and 0.5 <= abs(lv) <= 200:
                if lv < 0:
                    is_return = True
                liters = round(abs(lv), 2)

        if liters is None:
            # Try mathematical identity: Liters * Price == Sum
            num_cells = []
            for c_idx, cell in enumerate(row):
                if c_idx in (date_col_used, col_map.get("time"), col_map.get("azs")):
                    continue
                fv = parse_float_cell(cell)
                if fv is not None and fv != 0:
                    num_cells.append((c_idx, fv))
            for idx_a, val_a in num_cells:
                if 0.5 <= abs(val_a) <= 200:
                    for idx_b, val_b in num_cells:
                        if idx_b == idx_a or not (32.0 <= val_b <= 99.0):
                            continue
                        for idx_c, val_c in num_cells:
                            if idx_c in (idx_a, idx_b):
                                continue
                            if abs(abs(val_a) * val_b - abs(val_c)) <= max(1.5, abs(val_c) * 0.015):
                                if val_a < 0:
                                    is_return = True
                                liters = round(abs(val_a), 2)
                                break
                        if liters is not None:
                            break
                if liters is not None:
                    break

        if liters is None:
            # Fallback: pick candidate volume in [1.0 .. 180.0]
            excluded_cols = {date_col_used, col_map.get("time"), col_map.get("price"), col_map.get("sum"), col_map.get("fuel"), col_map.get("azs")}
            for c_idx, cell in enumerate(row):
                if c_idx in excluded_cols:
                    continue
                fv = parse_float_cell(cell)
                if fv is not None and 1.0 <= abs(fv) <= 180.0:
                    # Skip integer 92 or 95 if it looks like fuel grade
                    if abs(fv) in (92.0, 95.0, 98.0, 100.0) and ("аи" in row_low or len(row) >= 4):
                        continue
                    if fv < 0:
                        is_return = True
                    liters = round(abs(fv), 2)
                    break

        if liters is None or liters <= 0:
            continue

        # 4. Extract Fuel Grade
        fuel_grade = "92"
        fuel_text = str(row[col_map["fuel"]] or "") if ("fuel" in col_map and col_map["fuel"] < len(row)) else row_str
        if re.search(r"(?:аи[-\s]*95|\b95[-\s]*к5|\b95\b)", fuel_text.lower()):
            fuel_grade = "95"
        elif re.search(r"(?:аи[-\s]*92|\b92[-\s]*к5|\b92\b)", fuel_text.lower()):
            fuel_grade = "92"
        elif re.search(r"(?:аи[-\s]*95|\b95[-\s]*к5)", row_low):
            fuel_grade = "95"

        # 5. Extract AZS Address
        azs_addr = ""
        if "azs" in col_map and col_map["azs"] < len(row):
            raw_azs = str(row[col_map["azs"]] or "").strip()
            azs_addr = match_azs_from_text(raw_azs, azs_locs)
            if not azs_addr and len(raw_azs) > 3 and not re.match(r"^\d+$", raw_azs):
                azs_addr = raw_azs
        if not azs_addr:
            azs_addr = match_azs_from_text(row_str, azs_locs)
        if not azs_addr and azs_locs:
            azs_addr = azs_locs[0].get("addr", "г. Тамбов, ул. Астраханская, 39В (АЗС Роснефть №470)")

        driver_hint = ""
        if "driver" in col_map and col_map["driver"] < len(row):
            driver_hint = str(row[col_map["driver"]] or "").strip()

        date_iso = f"{yy:04d}-{mm:02d}-{dd:02d}"
        date_ru = f"{dd:02d}.{mm:02d}.{yy:04d}"

        receipts.append({
            "year": yy,
            "month": mm,
            "day": dd,
            "dateIso": date_iso,
            "dateRu": date_ru,
            "time": t_str,
            "liters": liters,
            "fuelGrade": fuel_grade,
            "azsAddr": azs_addr,
            "isReturn": is_return,
            "driverHint": driver_hint,
        })

    return receipts


@app.post("/api/fuel/parse-upload")
def parse_fuel_upload(payload: dict = Body(...)):
    filename = str(payload.get("filename") or "upload.xlsx").strip()
    b64 = str(payload.get("contentBase64") or "").strip()
    default_year = int(payload.get("defaultYear") or 2026)
    default_month = int(payload.get("defaultMonth") or 9)

    if not b64:
        raise HTTPException(status_code=400, detail="Пустой файл")

    raw_bytes = base64.b64decode(b64)
    st = load_valid_state()
    azs_locs = [l for l in st.get("locations", []) if l.get("zone") == "AZS"]

    all_receipts = []
    fname_low = filename.lower()

    # 1. Try openpyxl for .xlsx
    if fname_low.endswith(".xlsx") or raw_bytes[:2] == b"PK":
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                all_receipts.extend(extract_receipts_from_rows(rows, sname, azs_locs, default_year, default_month))
        except Exception:
            pass

    # 2. Try CSV / TSV / TXT / HTML-XLS if no receipts yet
    if not all_receipts:
        text_content = None
        for enc in ("utf-8-sig", "utf-8", "cp1251", "latin1"):
            try:
                text_content = raw_bytes.decode(enc)
                break
            except Exception:
                continue
        if text_content:
            # Check if JSON
            if text_content.lstrip().startswith(("{", "[")):
                try:
                    jdata = json.loads(text_content)
                    if isinstance(jdata, list):
                        for item in jdata:
                            if isinstance(item, dict) and item.get("liters"):
                                all_receipts.append(item)
                except Exception:
                    pass
            # Check if HTML table inside .xls
            if not all_receipts and "<tr" in text_content.lower():
                tr_blocks = re.findall(r"<tr[^>]*>(.*?)</tr>", text_content, flags=re.IGNORECASE | re.DOTALL)
                rows = []
                for tr in tr_blocks:
                    cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, flags=re.IGNORECASE | re.DOTALL)
                    clean_cells = [re.sub(r"<[^>]+>", " ", c).replace("&nbsp;", " ").strip() for c in cells]
                    if clean_cells:
                        rows.append(clean_cells)
                all_receipts.extend(extract_receipts_from_rows(rows, filename, azs_locs, default_year, default_month))
            # Otherwise parse as CSV / TSV / delimited lines
            if not all_receipts:
                lines = [ln for ln in text_content.splitlines() if ln.strip()]
                delim = ";" if sum(ln.count(";") for ln in lines[:10]) >= sum(ln.count(",") for ln in lines[:10]) else ","
                if sum(ln.count("\t") for ln in lines[:10]) > sum(ln.count(delim) for ln in lines[:10]):
                    delim = "\t"
                reader = csv.reader(lines, delimiter=delim)
                rows = list(reader)
                all_receipts.extend(extract_receipts_from_rows(rows, filename, azs_locs, default_year, default_month))

    # v8.1.0: чей чек — по колонке «водитель / держатель карты / госномер» (если она есть)
    profiles = st.get("profiles", []) or []
    for r in all_receipts:
        hint = str(r.get("driverHint") or "") if isinstance(r, dict) else ""
        if isinstance(r, dict):
            r["profileId"] = doc_match_profile(hint, profiles)[0] if hint else ""

    return {
        "status": "ok",
        "filename": filename,
        "count": len(all_receipts),
        "receipts": all_receipts,
    }


@app.post("/api/export/docx")
def export_docx(payload: dict = Body(...)):
    st = load_valid_state()
    schema = st.get("xls_schema")
    prof = payload.get("profile")
    days_list = payload.get("days", []); [d.update({"receipt_num": str(d.get("receipt_num") or i)}) for i, d in enumerate([x for x in days_list if float(x.get("refueled") or 0) > 0], 1)]
    file_name = payload.get("fileName") or "putevoy_list.docx"
    if not prof or not days_list or not schema:
        raise HTTPException(status_code=400, detail="Недостаточно данных для генерации документа Word")
    docx_bytes = generate_docx_bytes(prof, days_list, schema)
    encoded_name = urllib.parse.quote(file_name)
    return StreamingResponse(
        io.BytesIO(docx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )


@app.post("/api/export/xlsx")
def export_xlsx(payload: dict = Body(...)):
    st = load_valid_state()
    schema = st.get("xls_schema")
    prof = payload.get("profile")
    days_list = payload.get("days", []); [d.update({"receipt_num": str(d.get("receipt_num") or i)}) for i, d in enumerate([x for x in days_list if float(x.get("refueled") or 0) > 0], 1)]
    file_name = payload.get("fileName") or "putevoy_list.xlsx"
    if not prof or not days_list or not schema:
        raise HTTPException(status_code=400, detail="Недостаточно данных для генерации документа Excel")
    xlsx_bytes = generate_xlsx_bytes(prof, days_list, schema)
    encoded_name = urllib.parse.quote(file_name)
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"},
    )


app.mount("/static", StaticFiles(directory=os.path.join(BUNDLE_DIR if not os.path.exists(os.path.join(BASE_DIR, "static")) else BASE_DIR, "static")), name="static")

UPDATE_REPO = "MaksimTMB/putevka-rps"
UPDATE_FILES = ["version.txt", "server.py", "desktop_app.py", "requirements.txt", "static/index.html", "data/app_data.json"]

# Кэш последнего релиза: экономим запросы к API GitHub (лимит 60/час без токена)
_LAST_RELEASE_CACHE = {"tag": None, "ok": False, "ts": 0.0}
_RELEASE_CACHE_TTL = 600.0  # секунд


def _version_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:3])


def _fetch_latest_release():
    """Последний тег релиза с GitHub. Два источника + кэш 10 минут.

    Возвращает тег (например 'v7.3.5'). Если ни один источник не ответил —
    возвращаем текущую APP_VERSION, но запоминаем в кэше, что проверка
    НЕ удалась (флаг ok=False), чтобы /api/update/check мог честно сказать
    «не удалось проверить», а не притворяться «актуальная версия».
    """
    now = time.time()
    if _LAST_RELEASE_CACHE["tag"] is not None and (now - _LAST_RELEASE_CACHE["ts"]) < _RELEASE_CACHE_TTL:
        return _LAST_RELEASE_CACHE["tag"]

    tag = None
    ok = False
    # Способ 1: официальное API (может быть недоступно или лимит 60 запросов/час)
    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest",
            headers={"User-Agent": "PutevkaUpdater/1.0", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=8) as r:
            t = str(json.load(r).get("tag_name") or "")
            if re.match(r"^v?\d+\.\d+\.\d+$", t):
                tag, ok = t, True
    except Exception:
        pass
    # Способ 2: безлимитный переход releases/latest — конечный адрес содержит тег
    if not ok:
        try:
            req = urllib.request.Request(
                f"https://github.com/{UPDATE_REPO}/releases/latest",
                headers={"User-Agent": "PutevkaUpdater/1.0"},
            )
            with urllib.request.urlopen(req, timeout=10) as r:
                m = re.search(r"/tag/(v\d+\.\d+\.\d+)$", str(r.url))
                if m:
                    tag, ok = m.group(1), True
        except Exception:
            pass
    # Способ 3 (v7.3.4): версия из raw version.txt — другой домен GitHub,
    # работает даже когда недоступны API и сам сайт; тот же источник, откуда
    # качаются файлы обновления (/api/update/apply).
    if not ok:
        try:
            req = urllib.request.Request(
                f"https://raw.githubusercontent.com/{UPDATE_REPO}/main/version.txt",
                headers={"User-Agent": "PutevkaUpdater/1.0"},
            )
            with urllib.request.urlopen(req, timeout=8) as r:
                t = r.read().decode("utf-8", "ignore").strip()
                m = re.match(r"^v?(\d+\.\d+\.\d+)$", t)
                if m:
                    tag, ok = "v" + m.group(1), True
        except Exception:
            pass

    if ok:
        _LAST_RELEASE_CACHE.update({"tag": tag, "ok": True, "ts": now})
        return tag
    # Проверка не удалась: честно запоминаем это (кэш на 2 минуты, чтобы не долбить)
    _LAST_RELEASE_CACHE.update({"tag": APP_VERSION, "ok": False, "ts": now - _RELEASE_CACHE_TTL + 120.0})
    return APP_VERSION


def _latest_release_known():
    """True, если последнюю проверку релиза удалось выполнить успешно."""
    _fetch_latest_release()
    return _LAST_RELEASE_CACHE["ok"]


@app.get('/api/update/check')
def api_upd_check():
    latest = _fetch_latest_release()
    known = _latest_release_known()
    update_available = _version_tuple(latest) > _version_tuple(APP_VERSION)
    return {
        'status': 'ok',
        'currentVersion': APP_VERSION,
        'latestVersion': latest,
        'latestKnown': known,
        'updateAvailable': update_available and known,
    }


def _schedule_server_restart():
    def _restart():
        # Пауза 4 секунды: ответ /api/update/apply и сообщение об обновлении
        # должны успеть дойти до браузера ДО того, как сервер перезапустится.
        time.sleep(4.0)
        try:
            os.execv(sys.executable, [sys.executable] + sys.argv)
        except Exception:
            pass
    threading.Thread(target=_restart, daemon=True).start()


def _download_new_exe(timeout=180):
    """v7.3.5: скачивает новый Putevka.exe рядом с текущим (Putevka.exe.new).

    До 3 попыток. Возвращает путь к скачанному файлу или None.
    Вынос в отдельную функцию нужен, чтобы при обновлении exe-версии качать
    новый exe ЗАРАНЕЕ: если не скачался — обновление вообще не начинаем.
    """
    exe_path = sys.executable
    new_path = exe_path + ".new"
    url = f"https://github.com/{UPDATE_REPO}/releases/latest/download/Putevka.exe"
    for _ in range(3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "PutevkaUpdater/1.0"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                content = r.read()
            if not content or len(content) < 1_000_000:
                raise RuntimeError("скачанный exe подозрительно мал")
            with open(new_path, "wb") as f:
                f.write(content)
            return new_path
        except Exception:
            continue
    return None


def _replace_windows_exe(new_path=None):
    try:
        exe_path = sys.executable
        if not new_path or not os.path.exists(new_path):
            new_path = _download_new_exe()
        if not new_path:
            return False
        log_path = os.path.join(os.path.dirname(exe_path), "update.log")
        bat = exe_path + ".update.bat"
        with open(bat, "w", encoding="utf-8") as f:
            f.write("@echo off\r\n")
            f.write("chcp 65001 >nul\r\n")
            # журнал замены файла для диагностики проблем обновления
            f.write(f'echo [%date% %time%] start update > "{log_path}" 2>&1\r\n')
            # пауза, чтобы ответ сервера и сообщение дошли до окна приложения
            f.write("timeout /t 3 /nobreak >nul\r\n")
            f.write("taskkill /F /IM Putevka.exe >nul 2>&1\r\n")
            f.write(f'echo [%date% %time%] old exe stopped >> "{log_path}" 2>&1\r\n')
            f.write("timeout /t 2 /nobreak >nul\r\n")
            f.write(f'move /Y "{new_path}" "{exe_path}" >nul 2>&1\r\n')
            f.write(f'if exist "{new_path}" timeout /t 3 /nobreak >nul\r\n')
            f.write(f'move /Y "{new_path}" "{exe_path}" >nul 2>&1\r\n')
            f.write(f'if exist "{new_path}" timeout /t 3 /nobreak >nul\r\n')
            f.write(f'move /Y "{new_path}" "{exe_path}" >nul 2>&1\r\n')
            f.write(f'if exist "{new_path}" (echo [%date% %time%] WARNING: new exe NOT installed, file busy >> "{log_path}" 2>&1) else (echo [%date% %time%] new exe installed >> "{log_path}" 2>&1)\r\n')
            # запускаем программу в любом случае: если замена не удалась — запустится прежняя версия
            f.write(f'start "" "{exe_path}"\r\n')
            f.write('del "%~f0" >nul 2>&1\r\n')
        subprocess.Popen(["cmd", "/c", bat], creationflags=0x00000008, close_fds=True)
        return True
    except Exception:
        return False


@app.post('/api/update/apply')
def api_upd_apply():
    latest = _fetch_latest_release()
    if os.path.exists(DB_FILE):
        try:
            shutil.copy2(DB_FILE, DB_FILE + ".backup.json")
        except Exception:
            pass

    # v7.3.5: в exe-версии СНАЧАЛА скачиваем новый exe. Если он не скачался —
    # обновление не начинаем вовсе. Раньше файлы на диске обновлялись, а exe мог
    # не замениться/не перезапуститься → «наполовину обновлённое» состояние:
    # новая страница на диске при старом сервере в памяти (жёлтый баннер, F5 не помогает).
    pre_exe = None
    if IS_FROZEN:
        pre_exe = _download_new_exe()
        if not pre_exe:
            return {
                'status': 'error',
                'updatedTo': APP_VERSION,
                'updated': [],
                'failed': ['Putevka.exe'],
                'message': ('Не удалось скачать новый файл программы (Putevka.exe). '
                            'Обновление НЕ применено — программа продолжает работать прежней версией без изменений. '
                            'Проверьте интернет и попробуйте ещё раз через минуту, либо запустите update.ps1 '
                            '(правый клик → «Выполнить с помощью PowerShell»).'),
            }

    # 1. Сначала скачиваем ВСЕ файлы во временные (до 3 попыток на файл).
    #    Если хоть один не скачался — ничего не меняем (никаких «полуобновлений»).
    downloaded = {}
    failed = []
    for rel in UPDATE_FILES:
        url = f"https://raw.githubusercontent.com/{UPDATE_REPO}/main/{rel}"
        ok_file = None
        for attempt in range(3):
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "PutevkaUpdater/1.0"})
                with urllib.request.urlopen(req, timeout=60) as r:
                    data = r.read()
                if not data:
                    raise RuntimeError("пустой файл")
                tmp_path = os.path.join(BASE_DIR, rel + ".updtmp")
                parent = os.path.dirname(tmp_path)
                if parent:
                    os.makedirs(parent, exist_ok=True)
                with open(tmp_path, "wb") as f:
                    f.write(data)
                ok_file = tmp_path
                break
            except Exception:
                continue
        if ok_file:
            downloaded[rel] = ok_file
        else:
            failed.append(rel)

    if failed or not downloaded:
        return {
            'status': 'error',
            'updatedTo': APP_VERSION,
            'updated': [],
            'failed': failed,
            'message': ('Не удалось скачать файлы: ' + ', '.join(failed) +
                        '. Обновление НЕ применено (программа работает прежней версией). '
                        'Проверьте интернет и попробуйте ещё раз через минуту.'),
        }

    # 2. Все файлы на месте — заменяем. index.html последним: если что-то пойдёт не так,
    #    страница останется согласована со старым кодом сервера до полной замены.
    updated = []
    ordered = [rel for rel in UPDATE_FILES if rel != 'static/index.html'] + ['static/index.html']
    for rel in ordered:
        if rel not in downloaded:
            continue
        dest = os.path.join(BASE_DIR, rel)
        try:
            os.replace(downloaded[rel], dest)
            updated.append(rel)
        except Exception:
            failed.append(rel)

    # Реальная установленная версия — из скачанного version.txt, а не из API релизов
    real_version = latest
    try:
        with open(os.path.join(BASE_DIR, 'version.txt'), 'r', encoding='utf-8') as f:
            real_version = 'v' + f.read().strip().lstrip('v')
    except Exception:
        pass

    if updated:
        message = f'Обновлено файлов: {len(updated)}. Версия: {real_version}.'
        if failed:
            message += ' Не заменены: ' + ', '.join(failed) + '.'
        if IS_FROZEN:
            if _replace_windows_exe(pre_exe):
                message += ' Скачан новый exe. Приложение сейчас перезапустится с новой версией.'
            else:
                message += ' Перезапустите приложение, чтобы изменения вступили в силу. Рабочая база сохранена.'
        else:
            _schedule_server_restart()
            message += ' Сервер перезапускается с новой версией. Рабочая база сохранена.'
        return {
            'status': 'ok',
            'updatedTo': real_version,
            'updated': updated,
            'failed': failed,
            'message': message,
        }
    return {
        'status': 'error',
        'updatedTo': APP_VERSION,
        'updated': [],
        'failed': failed,
        'message': 'Не удалось заменить файлы программы. Обновление не применено.',
    }


# =====================================================================================
# v8.1.0 — БРАУЗЕР ДОКУМЕНТОВ: РАСПОЗНАВАНИЕ И РАСКЛАДКА ПО ПАПКАМ
# Любой загруженный файл читается (xlsx, docx, csv/txt/json/html; pdf и старый xls —
# текстом, который присылает браузер), после чего определяется:
#   • раздел (category): ведомости чеков / путевые листы / документы водителя /
#     документы автомобиля / прочее, и подпапка (kind) — вид документа;
#   • владелец (profileId): по госномеру, номеру ВУ, СНИЛС, ФИО — из профилей базы;
#   • период (year/month) для ведомостей и путевых листов, срок действия (validUntil).
# Если уверенности нет — needsReview=True: файл ложится в папку «Требуют проверки».
# Всё в одном файле server.py: обновление из программы качает только его (UPDATE_FILES).
# =====================================================================================

DOC_CATEGORY_LABELS = {
    "receipts": "Ведомости чеков ГСМ",
    "waybills": "Путевые листы",
    "driver": "Документы водителя",
    "car": "Документы автомобиля",
    "other": "Прочее",
}

# (раздел, подпапка, шаблоны). Порядок важен: ПТС раньше «паспорта», ОСАГО раньше прочего.
DOC_KIND_RULES = [
    ("car", "Страховка (ОСАГО, КАСКО)", [r"осаго", r"каско", r"страхов\w*\s+полис", r"полис\w*\s+(?:обязательного|страхования)", r"\bosago\b", r"\bkasko\b", r"\bpolis\b"]),
    ("car", "ПТС", [r"паспорт\w*\s+транспортн\w*\s+средств", r"\bэ?птс\b", r"\bpts\b"]),
    ("car", "СТС", [r"свидетельств\w*\s+о\s+регистрации\s+(?:тс|транспортн)", r"\bстс\b", r"\bsts\b"]),
    ("car", "Техосмотр (диагностическая карта)", [r"диагностическ\w*\s+карт", r"техосмотр", r"технического\s+осмотра", r"\btehosmotr\b", r"\bdiagnostic"]),
    ("car", "Ремонт и обслуживание", [r"заказ[\s-]*наряд", r"\bремонт", r"техническ\w*\s+обслуживан", r"сервисн\w*\s+книжк", r"замена\s+масла", r"\bremont\b"]),
    ("car", "Штрафы", [r"постановлени\w*\s+(?:по\s+делу|об\s+административн)", r"\bштраф", r"\bshtraf"]),
    ("car", "Договоры на автомобиль", [r"договор\w*\s+(?:аренды|купли)[\s\w-]{0,30}(?:транспорт|автомоб|\bтс\b)", r"акт\w*\s+при[её]ма[\s-]+передачи\s+(?:транспорт|автомоб)"]),
    ("driver", "Водительское удостоверение", [r"водительск\w*\s+удостоверен", r"\bв/?у\b", r"\bправа\b", r"\bвод\.?\s*удост", r"\bprava\b", r"\bvu\b", r"voditel\w*\s*udost"]),
    ("driver", "Медицинские документы", [r"медицинск\w*\s+(?:заключен|справк|осмотр|освидетельств)", r"медсправк", r"медосмотр", r"003[\s-]*в/?у", r"\bmed(?:spravka)?\b"]),
    ("driver", "СНИЛС", [r"снилс", r"страхов\w*\s+номер\w*\s+индивидуальн", r"\bsnils\b"]),
    ("driver", "Паспорт", [r"паспорт(?!\w*\s+транспортн)", r"\bpassport\b", r"\bpasport\b"]),
    ("driver", "Тахограф и карта водителя", [r"тахограф", r"карт\w*\s+водителя"]),
    ("driver", "Кадровые документы", [r"трудов\w*\s+(?:договор|книжк)", r"приказ\w*\s+о\s+при[её]ме", r"доверенност", r"\bdogovor\b"]),
]

WAYBILL_STRONG = [r"путев\w*\s+лист\w*\s+легков", r"0345001", r"оборотная\s+сторона\s+путев"]
WAYBILL_WEAK = [r"путев\w*\s+лист", r"\bputev\w*", r"\bпл[\s_-]+\d{2}[._]\d{2}"]
RECEIPTS_WORDS = [r"ведомост", r"транзакц", r"топливн\w*\s+карт", r"роснефть", r"отч[её]т\w*\s+по\s+карт", r"чек\w*\s+(?:азс|гсм)", r"\bgsm\b", r"\bvedomost"]

RU_MONTH_FULL = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма[йя]": 5, "июн": 6,
    "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
}
# латинские двойники букв номера + частая транслитерация (V→В, R→Р, N→Н, S→С, U→У)
_LAT2CYR = str.maketrans("ABEKMHOPCTYXVRNSU", "АВЕКМНОРСТУХВРНСУ")


def _doc_norm(s: str) -> str:
    s = str(s or "").lower().replace("ё", "е")
    s = re.sub(r"[_+]+", " ", s)
    return re.sub(r"\s+", " ", s)


def _plate_norm(s: str) -> str:
    return re.sub(r"[^0-9А-Я]", "", str(s or "").upper().replace("Ё", "Е").translate(_LAT2CYR))


def doc_extract_text(filename: str, raw: bytes, limit: int = 60000) -> str:
    """Текст файла силами сервера: xlsx (все листы), docx, csv/txt/json/html/xml. Без новых библиотек."""
    low = (filename or "").lower()
    out = []
    total = 0
    try:
        if raw[:2] == b"PK" and (low.endswith((".xlsx", ".xlsm")) or b"xl/workbook" in raw[:4000] or b"[Content_Types]" in raw[:200]):
            try:
                wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
                for ws in wb.worksheets:
                    out.append(ws.title)
                    for row in ws.iter_rows(values_only=True):
                        cells = [str(c) for c in row if c is not None and str(c).strip()]
                        if cells:
                            line = " | ".join(cells)
                            out.append(line)
                            total += len(line)
                            if total > limit:
                                break
                    if total > limit:
                        break
                wb.close()
                return "\n".join(out)[:limit]
            except Exception:
                pass
        if raw[:2] == b"PK":
            try:
                zf = zipfile.ZipFile(io.BytesIO(raw))
                names = zf.namelist()
                if "word/document.xml" in names:
                    xml = zf.read("word/document.xml").decode("utf-8", "ignore")
                    xml = re.sub(r"</w:p>", "\n", xml)
                    xml = re.sub(r"<w:tab/>|</w:tc>", " ", xml)
                    text = re.sub(r"<[^>]+>", "", xml)
                    return _html_unescape(text)[:limit]
            except Exception:
                pass
            return ""
        if low.endswith((".pdf", ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".heic", ".webp", ".doc", ".xls", ".zip", ".rar", ".7z")) or raw[:4] in (b"%PDF", b"\xd0\xcf\x11\xe0"):
            return ""
        for enc in ("utf-8-sig", "cp1251"):
            try:
                text = raw[: limit * 4].decode(enc)
                break
            except Exception:
                text = None
        if text is None:
            return ""
        if "<" in text and re.search(r"<(?:tr|td|p|div|table)\b", text, re.I):
            text = re.sub(r"</(?:tr|p|div|li|h\d)>", "\n", text, flags=re.I)
            text = _html_unescape(re.sub(r"<[^>]+>", " ", text))
        return text[:limit]
    except Exception:
        return ""


def _html_unescape(s: str) -> str:
    import html as _h
    return _h.unescape(s)


def _doc_rows_from_raw(filename: str, raw: bytes):
    """Строки таблиц файла (для поиска чеков): xlsx через openpyxl, csv/html — как в parse-upload."""
    low = (filename or "").lower()
    sheets = []
    if raw[:2] == b"PK" and not low.endswith((".docx", ".doc")):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
            for ws in wb.worksheets:
                sheets.append((ws.title, [list(r) for r in ws.iter_rows(values_only=True)]))
            return sheets
        except Exception:
            return []
    if low.endswith((".csv", ".txt", ".tsv", ".html", ".htm")) or (low.endswith(".xls") and raw[:1] == b"<"):
        text = None
        for enc in ("utf-8-sig", "cp1251"):
            try:
                text = raw.decode(enc)
                break
            except Exception:
                continue
        if not text:
            return []
        if "<tr" in text.lower():
            rows = []
            for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", text, flags=re.I | re.S):
                cells = [re.sub(r"<[^>]+>", " ", c).replace("&nbsp;", " ").strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, flags=re.I | re.S)]
                if cells:
                    rows.append(cells)
            return [(filename, rows)]
        lines = [ln for ln in text.splitlines() if ln.strip()]
        delim = ";" if sum(ln.count(";") for ln in lines[:10]) >= sum(ln.count(",") for ln in lines[:10]) else ","
        if sum(ln.count("\t") for ln in lines[:10]) > sum(ln.count(delim) for ln in lines[:10]):
            delim = "\t"
        return [(filename, list(csv.reader(lines, delimiter=delim)))]
    return []


def _rows_from_text_table(text: str):
    """Таблица из текста, который прислал браузер для старого .xls (строки «a | b | c»)."""
    rows = []
    for ln in (text or "").splitlines():
        if " | " in ln:
            rows.append([c.strip() for c in ln.split(" | ")])
    return rows


def doc_match_profile(text: str, profiles: list, extra_hints: list = None):
    """Чей документ. Возвращает (profileId|'' , способ, {profileId: очки}). Неоднозначность → ''."""
    norm = _doc_norm(text)
    plate_text = _plate_norm(text)
    digits_text = re.sub(r"\D", "", text or "")
    scores, why = {}, {}

    def add(pid, pts, reason):
        if pts > scores.get(pid, 0):
            scores[pid], why[pid] = pts, reason

    for p in profiles or []:
        pid = p.get("id")
        if not pid:
            continue
        plate = _plate_norm(p.get("carPlate"))
        if len(plate) >= 6:
            core = re.match(r"^([А-Я]\d{3}[А-Я]{2})", plate)
            if plate in plate_text or (core and core.group(1) in plate_text):
                add(pid, 10, "госномер")
        for fld, label in (("license", "номер ВУ"), ("snils", "СНИЛС")):
            d = re.sub(r"\D", "", str(p.get(fld) or ""))
            if len(d) >= 8 and d in digits_text:
                add(pid, 10, label)
        full = _doc_norm(p.get("driverFull"))
        parts = [x for x in re.split(r"[\s.]+", full) if x]
        if len(parts) >= 2 and full in norm:
            add(pid, 8, "ФИО")
        if parts:
            sur = parts[0]
            if len(parts) >= 3:
                i1, i2 = parts[1][0], parts[2][0]
                if re.search(rf"\b{re.escape(sur)}\s*{re.escape(i1)}\.?\s*{re.escape(i2)}\b", norm) or re.search(rf"\b{re.escape(i1)}\.?\s*{re.escape(i2)}\.?\s*{re.escape(sur)}\b", norm):
                    add(pid, 7, "фамилия и инициалы")
            if len(sur) >= 4 and re.search(rf"\b{re.escape(sur)}\w{{0,2}}\b", norm):
                add(pid, 4, "фамилия")
        short = _doc_norm(p.get("driverShort"))
        if len(short) >= 5 and short in norm:
            add(pid, 7, "фамилия и инициалы")
    if not scores:
        return "", "", {}
    best = max(scores.values())
    top = [pid for pid, sc in scores.items() if sc == best]
    if len(top) != 1 or best < 4:
        return "", "несколько совпадений" if len(top) > 1 else "", scores
    return top[0], why[top[0]], scores


def _doc_find_period(text: str):
    """(год, месяц) из текста/имени: «октябрь 2026», «10.2026», «2026-10», «2026_10»."""
    t = _doc_norm(text)
    for stem, num in RU_MONTH_FULL.items():
        m = re.search(rf"\b{stem}\w*\s+(20[2-3]\d)", t)
        if m:
            return int(m.group(1)), num
    m = re.search(r"\b(0?[1-9]|1[0-2])[._/-](20[2-3]\d)\b", t)
    if m:
        return int(m.group(2)), int(m.group(1))
    m = re.search(r"\b(20[2-3]\d)[._/-](0?[1-9]|1[0-2])\b(?![._/-]\d)", t)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def _doc_dates(text: str):
    out = []
    for m in re.finditer(r"\b(0?[1-9]|[12]\d|3[01])[./](0?[1-9]|1[0-2])[./](20[2-3]\d)\b", text or ""):
        try:
            out.append((m.start(), datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))))
        except ValueError:
            pass
    return out


def _doc_valid_until(text: str):
    """Срок действия: дата после «до/по/действителен/окончания» — самая поздняя из таких."""
    t = text or ""
    best = None
    for m in re.finditer(r"(?:действит\w*|срок\w*\s+действ\w*|окончани\w*|\bдо\b|\bпо\b)[^0-9\n]{0,40}(0?[1-9]|[12]\d|3[01])[./](0?[1-9]|1[0-2])[./](20[2-4]\d)", t, re.I):
        try:
            d = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        if best is None or d > best:
            best = d
    return best.isoformat() if best else ""


def classify_document(filename: str, raw: bytes, extracted_text: str = "", state: dict = None,
                      fallback_profile_id: str = "", fallback_year: int = 0, fallback_month: int = 0) -> dict:
    st = state if state is not None else load_valid_state()
    profiles = st.get("profiles", []) or []
    azs_locs = [l for l in st.get("locations", []) if l.get("zone") == "AZS"]
    server_text = doc_extract_text(filename, raw)
    text = (server_text + "\n" + (extracted_text or ""))[:120000]
    name_text = _doc_norm(os.path.splitext(filename or "")[0])
    body = _doc_norm(text)
    both = name_text + " \n " + body
    reasons = []

    def hits(patterns, where):
        return sum(1 for p in patterns if re.search(p, where))

    # 1. Чеки внутри таблиц
    receipts = []
    sheets = _doc_rows_from_raw(filename, raw)
    if not sheets and extracted_text and (filename or "").lower().endswith((".xls", ".ods")):
        sheets = [(filename, _rows_from_text_table(extracted_text))]
    fy, fm = (fallback_year or 2026), (fallback_month or 9)
    fp = _doc_find_period(name_text + " " + body[:3000])
    if fp:
        fy, fm = fp
    for sname, rows in sheets:
        try:
            receipts.extend(extract_receipts_from_rows(rows, sname, azs_locs, fy, fm))
        except Exception:
            pass

    wb_strong = hits(WAYBILL_STRONG, both)
    wb_weak = hits(WAYBILL_WEAK, both)
    rc_words = hits(RECEIPTS_WORDS, both)

    category, kind, needs_review = "other", "", False
    if wb_strong:
        category = "waybills"
        reasons.append("форма путевого листа (ОКУД 0345001)")
    elif receipts and (len(receipts) >= 2 or rc_words):
        category = "receipts"
        reasons.append(f"найдено чеков АЗС: {len(receipts)}")
    elif wb_weak:
        category = "waybills"
        reasons.append("в названии или тексте «путевой лист»")
    else:
        best = None
        for cat, k, pats in DOC_KIND_RULES:
            sc = 3 * hits(pats, name_text) + hits(pats, body)
            if sc and (best is None or sc > best[0]):
                best = (sc, cat, k)
        if best:
            _, category, kind = best
            reasons.append(f"распознан вид: {kind}")
        elif rc_words and sheets:
            category = "receipts"
            reasons.append("похоже на ведомость, но чеки не найдены")
            needs_review = True
        else:
            needs_review = True
            reasons.append("вид документа не распознан")

    # 2. Владелец: по чекам (колонка «водитель/карта/госномер»), иначе по тексту и имени файла
    receipt_profiles = {}
    if category == "receipts" and receipts:
        for r in receipts:
            hint = str(r.get("driverHint") or "")
            pid = doc_match_profile(hint, profiles)[0] if hint else ""
            r["profileId"] = pid
            receipt_profiles[pid] = receipt_profiles.get(pid, 0) + 1
    profile_id, how, scores = doc_match_profile(filename + "\n" + text, profiles)
    named = {k: v for k, v in receipt_profiles.items() if k}
    if len(named) >= 2:
        profile_id, how = "", "чеки нескольких водителей"
        reasons.append("в ведомости чеки нескольких водителей: " + ", ".join(
            f"{(next((p.get('driverShort') or p.get('driverFull') for p in profiles if p.get('id') == k), '?'))} — {v}" for k, v in named.items()))
    elif len(named) == 1:
        profile_id, how = next(iter(named)), "водитель в чеках"
    if profile_id:
        reasons.append(f"владелец найден: {how}")
    elif how == "чеки нескольких водителей":
        pass
    elif len(profiles) == 1:
        profile_id = profiles[0].get("id", "")
        reasons.append("единственный профиль в базе")
    elif fallback_profile_id and any(p.get("id") == fallback_profile_id for p in profiles):
        profile_id = fallback_profile_id
        needs_review = True
        reasons.append("владелец не найден в документе — положен к выбранному водителю, проверьте")
    else:
        needs_review = True
        reasons.append("владелец не найден")

    # 3. Период и срок действия
    year = month = 0
    if category == "receipts" and receipts:
        cnt = {}
        for r in receipts:
            key = (int(r.get("year") or 0), int(r.get("month") or 0))
            cnt[key] = cnt.get(key, 0) + 1
        (year, month), _ = max(cnt.items(), key=lambda kv: kv[1])
        if len(cnt) > 1:
            reasons.append("чеки за несколько месяцев: " + ", ".join(f"{m:02d}.{y}" for (y, m) in sorted(cnt)))
    elif category == "waybills":
        m = re.search(r"срок\s+действия\W{0,5}с\W{0,3}(0?[1-9]|[12]\d|3[01])[./](0?[1-9]|1[0-2])[./](20[2-3]\d)", body)
        if m:
            year, month = int(m.group(3)), int(m.group(2))
        else:
            per = _doc_find_period(name_text) or _doc_find_period(body[:4000])
            ds = _doc_dates(name_text + " " + text[:20000])
            if per:
                year, month = per
            elif ds:
                year, month = ds[0][1].year, ds[0][1].month
    if category in ("receipts", "waybills") and not year:
        if fallback_year and fallback_month:
            year, month = fallback_year, fallback_month
            needs_review = True
            reasons.append("месяц не найден в документе — взят выбранный период, проверьте")
    valid_until = _doc_valid_until(text) if category in ("driver", "car") else ""

    if not category or (category == "other" and not text.strip() and not needs_review):
        needs_review = True
    confidence = "high" if not needs_review and (profile_id or how == "чеки нескольких водителей") else ("medium" if not needs_review else "low")
    return {
        "category": category,
        "categoryLabel": DOC_CATEGORY_LABELS.get(category, "Прочее"),
        "kind": kind,
        "profileId": profile_id,
        "profileMatch": how,
        "year": year,
        "month": month,
        "validUntil": valid_until,
        "needsReview": bool(needs_review),
        "confidence": confidence,
        "reasons": reasons,
        "receiptsCount": len(receipts),
        "receiptProfiles": receipt_profiles,
        "receipts": receipts if category == "receipts" else [],
        "textChars": len(text.strip()),
    }


def _doc_meta_from_type(doc_type: str, period_key: str) -> dict:
    """Раскладка без распознавания (экспорт ПЛ, ведомость из «Чеков ГСМ»): по типу и периоду."""
    cat = {"receipt": "receipts", "waybill": "waybills"}.get(doc_type, "")
    m = re.search(r"(20\d{2})_(\d{1,2})", period_key or "")
    return {
        "category": cat or "other",
        "categoryLabel": DOC_CATEGORY_LABELS.get(cat or "other"),
        "kind": "",
        "year": int(m.group(1)) if m else 0,
        "month": int(m.group(2)) if m else 0,
        "needsReview": not cat,
        "confidence": "high" if cat else "low",
        "reasons": ["создан программой" if doc_type == "waybill" else ("загружен как ведомость чеков" if cat else "тип не указан")],
    }


@app.post("/api/documents/upload")
def documents_upload(payload: dict = Body(...)):
    """Сохранить файл в data/documents/ и разложить его.
    autoClassify=True — распознать содержимое (раздел, водитель, период, срок действия);
    иначе — раскладка по переданному типу (экспорт ПЛ, ведомость из «Чеков ГСМ»)."""
    filename = str(payload.get("fileName") or "document").strip()
    b64 = str(payload.get("contentBase64") or "").strip()
    doc_type = str(payload.get("type") or "document")
    profile_id = str(payload.get("profileId") or "")
    period_key = str(payload.get("periodKey") or "")
    auto = bool(payload.get("autoClassify"))
    if not b64:
        raise HTTPException(status_code=400, detail="Пустой файл")
    try:
        raw = base64.b64decode(b64)
    except Exception:
        raise HTTPException(status_code=400, detail="Некорректное содержимое файла")
    if len(raw) > 50 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="Файл слишком большой (максимум 50 МБ)")

    safe_name = re.sub(r'[^\w.\-]+', '_', filename) or "document"
    safe_name = safe_name[:80]
    ext = os.path.splitext(safe_name)[1].lower() or ".bin"
    doc_id = "doc_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_" + str(int(time.time() * 1000) % 10000)
    stored_name = doc_id + ext
    os.makedirs(DOCS_DIR, exist_ok=True)
    path = os.path.join(DOCS_DIR, stored_name)
    with open(path, "wb") as f:
        f.write(raw)

    meta = {
        "id": doc_id,
        "fileName": filename,
        "type": doc_type,
        "profileId": profile_id,
        "periodKey": period_key,
        "storedName": stored_name,
        "size": len(raw),
        "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "sha1": hashlib.sha1(raw).hexdigest(),
    }
    receipts = []
    if auto:
        cls = classify_document(filename, raw, str(payload.get("extractedText") or ""), None,
                                profile_id, int(payload.get("fallbackYear") or 0), int(payload.get("fallbackMonth") or 0))
        receipts = cls.pop("receipts", [])
        meta.update(cls)
        meta["type"] = {"receipts": "receipt", "waybills": "waybill"}.get(cls["category"], "document")
    else:
        meta.update(_doc_meta_from_type(doc_type, period_key))
    return {"status": "ok", "document": meta, "receipts": receipts}


@app.post("/api/documents/reclassify")
def documents_reclassify(payload: dict = Body(...)):
    """Повторно распознать уже сохранённый файл (старые документы без раздела, кнопка «Разобрать заново»)."""
    safe = os.path.basename(str(payload.get("storedName") or ""))
    path = os.path.join(DOCS_DIR, safe)
    if not safe or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    with open(path, "rb") as f:
        raw = f.read()
    cls = classify_document(str(payload.get("fileName") or safe), raw, str(payload.get("extractedText") or ""), None,
                            str(payload.get("fallbackProfileId") or ""), int(payload.get("fallbackYear") or 0), int(payload.get("fallbackMonth") or 0))
    receipts = cls.pop("receipts", [])
    cls["sha1"] = hashlib.sha1(raw).hexdigest()
    return {"status": "ok", "classification": cls, "receipts": receipts}


@app.post("/api/documents/zip")
def documents_zip(payload: dict = Body(...)):
    """v8.1.1: выбранные документы одним архивом; внутри — те же папки, что в программе
    (Водитель/Раздел/Подпапка/имя файла)."""
    items = payload.get("items") or []
    if not isinstance(items, list) or not items:
        raise HTTPException(status_code=400, detail="Не выбраны документы")

    def clean(part: str) -> str:
        part = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", str(part or "")).strip(" .")
        return part[:80] or "_"

    buf = io.BytesIO()
    used = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for it in items[:2000]:
            safe = os.path.basename(str(it.get("storedName") or ""))
            path = os.path.join(DOCS_DIR, safe)
            if not safe or not os.path.exists(path):
                continue
            folders = [clean(x) for x in (it.get("folders") or []) if str(x).strip()][:4]
            name = clean(it.get("fileName") or safe)
            arc = "/".join(folders + [name])
            base, ext = os.path.splitext(arc)
            n = 2
            while arc in used:
                arc = f"{base} ({n}){ext}"
                n += 1
            used.add(arc)
            zf.write(path, arc)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d")
    return StreamingResponse(io.BytesIO(buf.getvalue()), media_type="application/zip",
                             headers={"Content-Disposition": "attachment; filename*=UTF-8''" + urllib.parse.quote(f"Документы_{stamp}.zip")})


@app.get("/api/documents/text/{stored_name}")
def documents_text(stored_name: str):
    """Текст документа для предпросмотра (Word, таблицы, текстовые файлы)."""
    safe = os.path.basename(stored_name)
    path = os.path.join(DOCS_DIR, safe)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    with open(path, "rb") as f:
        raw = f.read()
    return {"status": "ok", "text": doc_extract_text(safe, raw, 200000)}


@app.get("/api/documents/download/{stored_name}")
def documents_download(stored_name: str, inline: int = 0, name: str = ""):
    safe = os.path.basename(stored_name)
    path = os.path.join(DOCS_DIR, safe)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    import mimetypes
    media = mimetypes.guess_type(safe)[0] or "application/octet-stream"
    # name — исходное имя файла (кириллица), чтобы скачивался «ОСАГО.pdf», а не doc_2026…pdf
    return FileResponse(path, filename=os.path.basename(name) or safe, media_type=media,
                        content_disposition_type="inline" if inline else "attachment")


@app.post("/api/documents/delete")
def documents_delete(payload: dict = Body(...)):
    stored_name = str(payload.get("storedName") or "")
    safe = os.path.basename(stored_name)
    path = os.path.join(DOCS_DIR, safe)
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass
    return {"status": "ok"}


# v8.1.2: запуск — СТРОГО в самом конце файла. Раньше этот блок стоял посередине, и при
# запуске «python server.py» (серверная версия: systemd, Docker) всё, что ниже него,
# не успевало подключиться: проверка обновлений и весь раздел «Документы» отвечали 404.
# Адрес и порт — из окружения (PUTEVKA_HOST / PUTEVKA_PORT или PORT), по умолчанию как раньше.
if __name__ == "__main__":
    import uvicorn
    _host = os.environ.get("PUTEVKA_HOST") or "127.0.0.1"
    try:
        _port = int(os.environ.get("PUTEVKA_PORT") or os.environ.get("PORT") or 8080)
    except ValueError:
        _port = 8080
    uvicorn.run(app, host=_host, port=_port)
