import os, sys
if sys.stdout is None: sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w', encoding='utf-8')
import os, io, shutil, re, csv, base64, json, datetime, zipfile, urllib.parse, urllib.request, hashlib
import subprocess, threading, time
from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, StreamingResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "v8.5.6"

app = FastAPI(title="ПУТЁВКА v8.5.6 — Форма № 3 (ОКУД 0345001)")

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


# =====================================================================================
# v8.1.2 — ВХОД НА СЕРВЕРЕ
# Раньше сервер отдавал всю базу (вместе с хэшами паролей) любому, кто откроет /api/state,
# принимал её перезапись без входа, а пароль сверялся в браузере. Теперь:
#   • без входа доступны только страница, /api/version, /api/update/check и /api/auth/*;
#   • вход — на сервере, сеанс в cookie (HttpOnly, SameSite=Strict), сеансы переживают
#     перезапуск программы (data/sessions.json);
#   • хэши паролей никогда не уходят в браузер (вместо них hasPassword);
#   • водитель («просмотр и печать») получает только свой профиль и ничего не записывает;
#   • повторная «первичная настройка» администратора невозможна, если он уже настроен.
# =====================================================================================
import secrets
from fastapi import Request
from fastapi.responses import JSONResponse

SESSION_COOKIE = "putevka_sid"
SESSION_TTL = 7 * 24 * 3600
SESSION_IDLE = 24 * 3600  # v8.3.0: сутки без действий — сеанс закрывается
SESSIONS_FILE = os.path.join(DATA_DIR, "sessions.json")
_SESS_LOCK = threading.Lock()
_SESSIONS = {}  # v8.2.0: ключ — SHA-256 от токена (сам токен на диске не хранится)
_PUBLIC_API = {"/api/version", "/api/update/check", "/api/auth/login", "/api/auth/setup",
               "/api/auth/logout", "/api/auth/me", "/api/auth/forgot", "/api/auth/reset"}
# что можно водителю (метод, путь или префикс пути, заканчивающийся на «/»)
_DRIVER_ALLOWED = [("GET", "/api/state"), ("POST", "/api/export/docx"), ("POST", "/api/export/xlsx"),
                   ("GET", "/api/documents/download/"), ("GET", "/api/documents/text/"), ("GET", "/api/deadlines"),
                   ("POST", "/api/auth/change-password")]
# v8.2.0: программа на одном компьютере (exe) слушает только 127.0.0.1 — тогда принимаем запросы
# только с адресом 127.0.0.1/localhost (защита от «DNS rebinding»: чужой сайт не достучится до программы)
LOCAL_ONLY = os.environ.get("PUTEVKA_LOCAL_ONLY") == "1"


def _tok_key(tok: str) -> str:
    return hashlib.sha256(("putevka-session:" + str(tok)).encode("utf-8")).hexdigest()


def _sessions_load():
    try:
        with open(SESSIONS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        now = time.time()
        # сеансы v8.1.2 хранились открытым токеном — такие не берём (один раз войти заново)
        _SESSIONS.update({k: v for k, v in data.items()
                          if isinstance(v, dict) and v.get("exp", 0) > now and re.fullmatch(r"[0-9a-f]{64}", k)})
    except Exception:
        pass


def _sessions_save():
    try:
        tmp = SESSIONS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_SESSIONS, f)
        os.replace(tmp, SESSIONS_FILE)
    except Exception:
        pass


_sessions_load()

_USERS_CACHE = {"mtime": -1.0, "users": [], "setup": False}


def _users_now():
    try:
        mt = os.path.getmtime(DB_FILE)
    except Exception:
        mt = 0.0
    if mt != _USERS_CACHE["mtime"]:
        st = load_valid_state()
        _USERS_CACHE.update({"mtime": mt, "users": st.get("users", []) or [],
                             "setup": bool(st.get("adminSetupCompleted"))})
    return _USERS_CACHE["users"]


def _public_user(u: dict) -> dict:
    out = {k: v for k, v in (u or {}).items() if k not in ("password", "newPassword")}
    out["hasPassword"] = bool((u or {}).get("password"))
    return out


def _is_admin(u) -> bool:
    return bool(u) and u.get("role") == "admin"


def _session_user(request: Request):
    tok = request.cookies.get(SESSION_COOKIE)
    if not tok:
        return None
    now = time.time()
    with _SESS_LOCK:
        s = _SESSIONS.get(_tok_key(tok))
        if s and now - s.get("last", now) > SESSION_IDLE:
            _SESSIONS.pop(_tok_key(tok), None)
            _sessions_save()
            s = None
        elif s and now - s.get("last", 0) > 300:
            s["last"] = now  # отметка активности (на диск — не чаще раза в 5 минут)
            _sessions_save()
    if not s or s.get("exp", 0) < now:
        return None
    u = next((x for x in _users_now() if x.get("id") == s.get("uid")), None)
    if not u or u.get("active") is False:
        return None
    return u


def _start_session(response, uid: str, secure: bool = False):
    tok = secrets.token_urlsafe(32)
    now = time.time()
    with _SESS_LOCK:
        for k in [k for k, v in _SESSIONS.items() if v.get("exp", 0) < now]:
            _SESSIONS.pop(k, None)
        _SESSIONS[_tok_key(tok)] = {"uid": uid, "exp": now + SESSION_TTL, "last": now}
        _sessions_save()
    response.set_cookie(SESSION_COOKIE, tok, httponly=True, samesite="strict", path="/", secure=secure)


def _drop_user_sessions(uid: str):
    with _SESS_LOCK:
        for k in [k for k, v in _SESSIONS.items() if v.get("uid") == uid]:
            _SESSIONS.pop(k, None)
        _sessions_save()


def _driver_may(method: str, path: str) -> bool:
    for m, p in _DRIVER_ALLOWED:
        if m == method and (path == p or (p.endswith("/") and path.startswith(p))):
            return True
    return False


# --- v8.2.0: пароли с солью (PBKDF2-SHA256) вместо голого SHA-256 -------------------------
_PBKDF2_ITER = 200_000


def _hash_password(password: str) -> str:
    """Старый формат (SHA-256 без соли) — только для проверки паролей из прошлых версий."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _make_password(password: str) -> str:
    salt = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), _PBKDF2_ITER).hex()
    return f"pbkdf2_sha256${_PBKDF2_ITER}${salt}${h}"


def _check_password(stored: str, password: str):
    """(подходит ли, нужно ли перевести на новый формат)."""
    stored = stored or ""
    if not stored or not password:
        return False, False
    if stored.startswith("pbkdf2_sha256$"):
        try:
            _, it, salt, h = stored.split("$")
            calc = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), int(it)).hex()
            return secrets.compare_digest(calc, h), False
        except Exception:
            return False, False
    if re.fullmatch(r"[0-9a-f]{64}", stored):
        return secrets.compare_digest(stored, _hash_password(password)), True
    # совсем старые базы: пароль открытым текстом (сам хэш паролем НЕ работает — он не открытый текст)
    return secrets.compare_digest(stored, password), True


# --- v8.2.0: ограничение подбора пароля: 10 ошибок за 10 минут с одного адреса → пауза ------
_LOGIN_FAILS = {}
_LOGIN_LOCK = threading.Lock()


def _login_blocked(ip: str) -> bool:
    now = time.time()
    with _LOGIN_LOCK:
        lst = [t for t in _LOGIN_FAILS.get(ip, []) if now - t < 600]
        _LOGIN_FAILS[ip] = lst
        return len(lst) >= 10


def _login_failed(ip: str):
    with _LOGIN_LOCK:
        _LOGIN_FAILS.setdefault(ip, []).append(time.time())


# --- v8.2.0: защитные заголовки -------------------------------------------------------------
# Страница может обращаться по сети ТОЛЬКО к самой программе (connect-src 'self') — даже если
# чужой код попадёт в окно (через подменённую библиотеку или загруженный файл), данные он
# никуда не отправит. Библиотеки грузятся только с перечисленных CDN и сверяются по хэшу (SRI).
CSP_PAGE = ("default-src 'self'; "
            "script-src 'self' 'unsafe-inline' 'wasm-unsafe-eval' https://cdn.tailwindcss.com https://cdn.jsdelivr.net "
            "https://cdnjs.cloudflare.com https://cdn.sheetjs.com; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' data: https://fonts.gstatic.com; "
            "img-src 'self' data: blob:; connect-src 'self' https://cdn.jsdelivr.net; worker-src 'self' blob:; "
            "frame-src 'self' blob:; object-src 'none'; base-uri 'self'; form-action 'self'; frame-ancestors 'none'")
CSP_DOC = "sandbox; default-src 'none'; img-src 'self' data:; style-src 'unsafe-inline'; frame-ancestors 'self'"


def _host_ok(request: Request) -> bool:
    if not LOCAL_ONLY:
        return True
    host = (request.headers.get("host") or "").split(":")[0].strip("[]").lower()
    return host in ("127.0.0.1", "localhost", "::1", "")


LAST_CLIENT_SEEN = 0.0  # v8.2.1: когда окно программы последний раз обращалось к серверу


@app.middleware("http")
async def _auth_guard(request: Request, call_next):
    global LAST_CLIENT_SEEN
    if not request.headers.get("user-agent", "").startswith("Python-urllib"):  # свои проверки программы не в счёт
        LAST_CLIENT_SEEN = time.time()
    if not _host_ok(request):
        return JSONResponse({"status": "error", "detail": "Недопустимый адрес"}, status_code=400)
    path = request.url.path
    if path.startswith("/api/") and path not in _PUBLIC_API:
        u = _session_user(request)
        if not u:
            return JSONResponse({"status": "error", "detail": "Требуется вход в программу",
                                 "message": "Требуется вход в программу"}, status_code=401)
        if not _is_admin(u) and not _driver_may(request.method, path):
            msg = "Недостаточно прав: водителю доступны только просмотр и печать своих путевых листов"
            return JSONResponse({"status": "error", "detail": msg, "message": msg}, status_code=403)
        request.state.user = u
    resp = await call_next(request)
    h = resp.headers
    h.setdefault("X-Content-Type-Options", "nosniff")
    h.setdefault("Referrer-Policy", "no-referrer")
    h.setdefault("X-Frame-Options", "SAMEORIGIN")
    h.setdefault("Permissions-Policy", "camera=(self), microphone=(), geolocation=()")
    if path.startswith("/api/"):
        h.setdefault("Cache-Control", "no-store")
    if (h.get("content-type") or "").startswith("text/html") and not path.startswith("/api/documents/"):
        h.setdefault("Content-Security-Policy", CSP_PAGE)
    return resp


def _strip_state_for(u: dict, st: dict) -> dict:
    if _is_admin(u):
        st["users"] = [_public_user(x) for x in st.get("users", []) or []]
        return st
    st["users"] = [_public_user(u)]
    pid = u.get("assignedProfileId")
    if pid and pid != "ALL":
        st["profiles"] = [p for p in st.get("profiles", []) or [] if p.get("id") == pid]
        st["savedDocuments"] = [d for d in st.get("savedDocuments", []) or [] if d.get("profileId") == pid]
        st["fines"] = [f for f in st.get("fines", []) or [] if f.get("profileId") == pid]
    return st


@app.get("/api/state")
def get_state(request: Request):
    return _strip_state_for(request.state.user, load_valid_state())


@app.post("/api/state")
def save_state(request: Request, payload: dict = Body(...)):
    if not isinstance(payload, dict) or not isinstance(payload.get("profiles"), list):
        raise HTTPException(status_code=400, detail="Некорректные данные: нет списка профилей — база не перезаписана")
    existing = load_valid_state()
    # пароли: браузер их не видит; новый пароль приходит один раз в newPassword и хэшируется здесь
    old_users = {x.get("id"): x for x in existing.get("users", []) or []}
    changed_pw = []
    for x in payload.get("users") or []:
        if not isinstance(x, dict):
            continue
        x.pop("hasPassword", None)
        newpw = str(x.pop("newPassword", "") or "").strip()
        if newpw:
            prob = password_problem(newpw, x.get("login"))
            if prob:
                raise HTTPException(status_code=400, detail=f"Пароль пользователя «{x.get('login') or x.get('fullName') or ''}»: {prob}. База не перезаписана.")
            x["password"] = _make_password(newpw)
            changed_pw.append(x.get("id"))
        elif not x.get("password"):
            old = old_users.get(x.get("id"))
            if old and old.get("password"):
                x["password"] = old["password"]
            else:
                x.pop("password", None)
    if existing.get("adminSetupCompleted"):
        payload["adminSetupCompleted"] = True
    save_valid_state(payload)
    me = getattr(request.state, "user", {}) or {}
    for uid in changed_pw:
        _set_weak(uid, False)
        _sec_log("password-set-by-admin", request, uid)
        if uid != me.get("id"):
            _drop_user_sessions(uid)  # сменили пароль — старые входы этого пользователя закрываются
    return {
        "status": "ok",
        "saved_at": datetime.datetime.now().isoformat(),
        "version": APP_VERSION,
    }


@app.post("/api/auth/setup")
def setup_admin(request: Request, payload: dict = Body(...)):
    login = str(payload.get("login") or "").strip()
    password = str(payload.get("password") or "").strip()
    full_name = str(payload.get("fullName") or "Главный Администратор").strip()
    if not login or not password:
        raise HTTPException(status_code=400, detail="Логин и пароль обязательны")
    prob = password_problem(password, login)
    if prob:
        raise HTTPException(status_code=400, detail=prob)

    st = load_valid_state()
    users = st.get("users", []) or []
    if st.get("adminSetupCompleted") and any(u.get("role") == "admin" and u.get("password") for u in users):
        raise HTTPException(status_code=403, detail="Администратор уже настроен — войдите под своим логином")
    admin_user = None
    for u in users:
        if u.get("role") == "admin":
            u["login"] = login
            u["password"] = _make_password(password)
            u["fullName"] = full_name
            u["shortName"] = full_name
            u["active"] = True
            admin_user = u
            break
    if not admin_user:
        admin_user = {
            "id": "usr_admin",
            "login": login,
            "password": _make_password(password),
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
    _set_weak(admin_user["id"], False)
    code = _new_recovery_code(admin_user["id"])
    _sec_log("admin-setup", request, login)
    resp = JSONResponse({"status": "ok", "user": _public_user(admin_user), "recoveryCode": code})
    _start_session(resp, admin_user["id"], secure=request.url.scheme == "https")
    return resp


@app.post("/api/auth/login")
def login_user(request: Request, payload: dict = Body(...)):
    ip = request.client.host if request.client else "?"
    if _login_blocked(ip):
        return JSONResponse({"status": "error", "message": "Слишком много неудачных попыток входа. Подождите 10 минут."}, status_code=429)
    login = str(payload.get("login") or "").strip()
    password = str(payload.get("password") or "").strip()
    left = _acct_locked_left(login)
    if left:
        _sec_log("login-while-locked", request, login)
        return JSONResponse({"status": "error", "message": _lock_message(left)}, status_code=429)
    st = load_valid_state()
    users = st.get("users", []) or []
    for u in users:
        if not login or u.get("login") != login:
            continue
        ok, upgrade = _check_password(u.get("password") or "", password)
        if not ok:
            continue
        if u.get("active") is False:
            return {"status": "error", "message": "Учетная запись отключена администратором"}
        if upgrade:
            # пароль из прошлых версий (SHA-256 без соли или открытый текст) — переводим на новый формат
            u["password"] = _make_password(password)
            save_valid_state(st)
        weak = bool(password_problem(password, u.get("login")))
        _set_weak(u.get("id"), weak)
        _acct_ok(login)
        _sec_log("login-ok", request, login)
        resp = JSONResponse({"status": "ok", "user": _public_user(u), "mustChangePassword": weak})
        _start_session(resp, u.get("id"), secure=request.url.scheme == "https")
        return resp
    _login_failed(ip)
    _acct_failed(login, request)
    _sec_log("login-bad", request, login)
    time.sleep(0.4)  # притормаживаем подбор пароля
    return {"status": "error", "message": "Неверный логин или пароль"}


@app.post("/api/auth/logout")
def logout_user(request: Request):
    tok = request.cookies.get(SESSION_COOKIE)
    if tok:
        with _SESS_LOCK:
            if _SESSIONS.pop(_tok_key(tok), None) is not None:
                _sessions_save()
    resp = JSONResponse({"status": "ok"})
    resp.delete_cookie(SESSION_COOKIE, path="/")
    return resp


@app.get("/api/auth/me")
def auth_me(request: Request):
    u = _session_user(request)
    if not u:
        return JSONResponse({"status": "error", "message": "Требуется вход"}, status_code=401)
    sec = _sec_load()
    return {"status": "ok", "user": _public_user(u), "mustChangePassword": bool((sec.get("weak") or {}).get(u.get("id"))),
            "hasRecoveryCode": bool((sec.get("recovery") or {}).get(u.get("id"))), "mailReady": _smtp_ready()}


def _doc_allowed(request: Request, stored_name: str) -> bool:
    """Водитель видит только документы своего профиля."""
    u = getattr(request.state, "user", None)
    if _is_admin(u):
        return True
    pid = (u or {}).get("assignedProfileId")
    if not pid or pid == "ALL":
        return True
    st = load_valid_state()
    return any(d.get("storedName") == stored_name and d.get("profileId") == pid for d in st.get("savedDocuments", []) or [])



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


# v8.5.0: модели распознавания водительского удостоверения (PP-OCR, PaddleOCR, Apache-2.0).
# В exe они внутри файла программы, у установки «из файлов» — в static/ocr. Отдаём оттуда, где есть:
# после обновления static/ берётся из папки установки, а модели остаются внутри нового exe.
OCR_MODEL_FILES = ("ppocr_det_v3.onnx", "ppocr_eslav_rec_v5.onnx", "ppocr_eslav_dict.txt")


@app.get("/ocr-models/{name}")
def get_ocr_model(name: str):
    if name not in OCR_MODEL_FILES:
        raise HTTPException(status_code=404, detail="Нет такого файла")
    for root in (BASE_DIR, BUNDLE_DIR):
        p = os.path.join(root, "static", "ocr", name)
        if os.path.isfile(p):
            media = "text/plain; charset=utf-8" if name.endswith(".txt") else "application/octet-stream"
            return FileResponse(p, media_type=media, headers={"Cache-Control": "public, max-age=31536000, immutable"})
    raise HTTPException(status_code=404, detail="Модель распознавания не найдена — ВУ читается прежним способом")


# v8.5.2: библиотеки страницы — в программе. cdn.tailwindcss.com, cdn.sheetjs.com и cdnjs.cloudflare.com из России
# открываются не всегда (Cloudflare): без оформления страница показывала все экраны разом и вход не работал.
# Отдаём из папки установки или из пакета exe; файла нет (сервер обновлён старой версией и ещё не докачал) — с прежнего CDN.
VENDOR_FILES = {
    "tailwind-3.4.17.js": "https://cdn.tailwindcss.com/3.4.17",
    "chart-4.4.1.umd.min.js": "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js",
    "xlsx-0.20.3.full.min.js": "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
    "pdf-3.11.174.min.js": "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
    "pdf.worker-3.11.174.min.js": "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
}
ASSET_FILES = ["static/ocr/" + n for n in OCR_MODEL_FILES] + ["static/vendor/" + n for n in VENDOR_FILES] + ["static/vendor/NOTICE.txt"]


@app.get("/vendor/{name}")
def get_vendor(name: str):
    if name not in VENDOR_FILES:
        raise HTTPException(status_code=404, detail="Нет такого файла")
    for root in (BASE_DIR, BUNDLE_DIR):
        p = os.path.join(root, "static", "vendor", name)
        if os.path.isfile(p):
            return FileResponse(p, media_type="application/javascript; charset=utf-8", headers={"Cache-Control": "public, max-age=31536000, immutable"})
    return RedirectResponse(VENDOR_FILES[name], status_code=302)


def _fetch_missing_assets():
    """v8.5.2: установка «из файлов» (сервер): докачать модели ВУ и библиотеки страницы, если их нет. В exe они внутри."""
    got = []
    if IS_FROZEN:
        return got
    for rel in ASSET_FILES:
        dest = os.path.join(BASE_DIR, *rel.split("/"))
        if os.path.isfile(dest):
            continue
        try:
            req = urllib.request.Request(f"{UPDATE_RAW_BASE}/{rel}", headers={"User-Agent": "PutevkaUpdater/1.0"})
            with urllib.request.urlopen(req, timeout=180) as r:
                data = r.read()
            if not data:
                continue
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            with open(dest + ".updtmp", "wb") as f:
                f.write(data)
            os.replace(dest + ".updtmp", dest)
            got.append(rel)
        except Exception as e:
            _err_log("assets", e, rel)
    return got


@app.on_event("startup")
def _assets_on_start():
    threading.Thread(target=_fetch_missing_assets, daemon=True).start()


app.mount("/static", StaticFiles(directory=os.path.join(BUNDLE_DIR if not os.path.exists(os.path.join(BASE_DIR, "static")) else BASE_DIR, "static")), name="static")

UPDATE_REPO = "MaksimTMB/putevka-rps"
# v8.2.1: источники обновления можно подменить для проверки на тестовой машине (только переменными окружения)
UPDATE_EXE_URL = os.environ.get("PUTEVKA_UPDATE_EXE_URL") or f"https://github.com/{UPDATE_REPO}/releases/latest/download/Putevka.exe"
UPDATE_RAW_BASE = (os.environ.get("PUTEVKA_UPDATE_RAW_BASE") or f"https://raw.githubusercontent.com/{UPDATE_REPO}/main").rstrip("/")
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
    url = UPDATE_EXE_URL
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


def _clean_child_env():
    """Окружение для запуска НОВОГО экземпляра программы: без переменных распаковщика PyInstaller."""
    env = {k: v for k, v in os.environ.items()
           if not (k.startswith("_PYI_") or k.startswith("_MEIPASS") or k == "PUTEVKA_RELAUNCHED")}
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return env


def _replace_windows_exe(new_path=None):
    """v8.2.1: новый exe сам встаёт на место старого: запускаем скачанный файл с
    «--finish-update <куда> <pid> <pid загрузчика>». Он ждёт, закрывает прежнюю версию ПО НОМЕРУ
    процесса (раньше — по имени Putevka.exe: «Putevka (1).exe» не закрывался), копирует себя на место
    (с повторами, если файл занят) и запускает программу. Без батника: русские пути не ломаются,
    а логика замены — всегда из НОВОЙ версии."""
    try:
        exe_path = sys.executable
        if not new_path or not os.path.exists(new_path):
            new_path = _download_new_exe()
        if not new_path:
            return False
        subprocess.Popen([new_path, "--finish-update", exe_path, str(os.getpid()), str(os.getppid())],
                         creationflags=0x00000008 | 0x00000200, close_fds=True,
                         cwd=os.path.dirname(exe_path), env=_clean_child_env())
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
        url = f"{UPDATE_RAW_BASE}/{rel}"
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

    # v8.5.0/8.5.2: модели ВУ и библиотеки страницы — один раз, если их нет (только установка «из файлов»; в exe они внутри).
    # Не скачались — не беда: страница возьмёт библиотеки с прежних CDN, ВУ прочитается прежним способом.
    if updated and not IS_FROZEN:
        updated.extend(_fetch_missing_assets())

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

    # v8.2.0: постановления и выписки штрафов — раздел «Документы автомобиля / Штрафы» + разобранные штрафы
    fines = parse_fines_text(text) if re.search(r"постановлен|штраф|уин|коап|12\.\d", both) else []
    fines = [f for f in fines if f.get("uin") or (f.get("amount") and f.get("article"))]

    wb_strong = hits(WAYBILL_STRONG, both)
    wb_weak = hits(WAYBILL_WEAK, both)
    rc_words = hits(RECEIPTS_WORDS, both)

    category, kind, needs_review = "other", "", False
    if wb_strong:
        category = "waybills"
        reasons.append("форма путевого листа (ОКУД 0345001)")
    elif fines:
        category, kind = "car", "Штрафы"
        reasons.append(f"найдено штрафов: {len(fines)}")
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
    if fines:
        _fines_with_profiles(fines, profiles)
        owners = {pid for f in fines for pid in f.get("carProfileIds") or []}
        if len(owners) == 1:
            profile_id, how = next(iter(owners)), "госномер/СТС в постановлении"
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
        "fines": fines,
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
    path = os.path.join(DOCS_DIR, stored_name)
    try:
        os.makedirs(DOCS_DIR, exist_ok=True)
        with open(path, "wb") as f:
            f.write(raw)
    except Exception as e:
        # v8.5.1: раньше — 500 без объяснения, в окне «сервер не ответил»; теперь причина видна и записана в журнал
        _err_log("documents/upload", e, f"{DOCS_DIR} ({len(raw)} байт)")
        raise HTTPException(status_code=500, detail=f"не удалось записать файл в папку документов программы: {type(e).__name__}: {e}")

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
    receipts, fines = [], []
    if payload.get("category") in DOC_CATEGORY_LABELS and not auto:
        cat = payload.get("category")
        meta.update({"category": cat, "categoryLabel": DOC_CATEGORY_LABELS[cat], "kind": str(payload.get("kind") or "")[:80],
                     "year": 0, "month": 0, "validUntil": str(payload.get("validUntil") or "")[:10], "needsReview": False,
                     "confidence": "high", "reasons": [str(payload.get("reason") or "загружен в карточку водителя")[:120]]})
        return {"status": "ok", "document": meta, "receipts": [], "fines": []}
    if auto:
        cls = classify_document(filename, raw, str(payload.get("extractedText") or ""), None,
                                profile_id, int(payload.get("fallbackYear") or 0), int(payload.get("fallbackMonth") or 0))
        receipts = cls.pop("receipts", [])
        fines = cls.pop("fines", [])
        meta.update(cls)
        meta["type"] = {"receipts": "receipt", "waybills": "waybill"}.get(cls["category"], "document")
    else:
        meta.update(_doc_meta_from_type(doc_type, period_key))
    return {"status": "ok", "document": meta, "receipts": receipts, "fines": fines}


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
    fines = cls.pop("fines", [])
    cls["sha1"] = hashlib.sha1(raw).hexdigest()
    return {"status": "ok", "classification": cls, "receipts": receipts, "fines": fines}


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
def documents_text(stored_name: str, request: Request):
    """Текст документа для предпросмотра (Word, таблицы, текстовые файлы)."""
    safe = os.path.basename(stored_name)
    if not _doc_allowed(request, safe):
        raise HTTPException(status_code=403, detail="Нет доступа к этому документу")
    path = os.path.join(DOCS_DIR, safe)
    if not safe or not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    with open(path, "rb") as f:
        raw = f.read()
    return {"status": "ok", "text": doc_extract_text(safe, raw, 200000)}


@app.get("/api/documents/download/{stored_name}")
def documents_download(stored_name: str, request: Request, inline: int = 0, name: str = ""):
    safe = os.path.basename(stored_name)
    if not safe or not _doc_allowed(request, safe):
        raise HTTPException(status_code=403, detail="Нет доступа к этому документу")
    path = os.path.join(DOCS_DIR, safe)
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    import mimetypes
    media = mimetypes.guess_type(safe)[0] or "application/octet-stream"
    ext = os.path.splitext(safe)[1].lower()
    # v8.2.0: загруженный файл не должен исполняться в окне программы. Показываем внутри окна только
    # PDF, картинки (не SVG) и простой текст; HTML/SVG/XML и прочее — только скачиванием и в «песочнице».
    inline_ok = bool(inline) and (media == "application/pdf" or (media.startswith("image/") and ext != ".svg") or media == "text/plain")
    if media == "text/plain":
        media = "text/plain; charset=utf-8"
    # name — исходное имя файла (кириллица), чтобы скачивался «ОСАГО.pdf», а не doc_2026…pdf
    resp = FileResponse(path, filename=os.path.basename(name) or safe, media_type=media,
                        content_disposition_type="inline" if inline_ok else "attachment")
    if not (media == "application/pdf" and inline_ok):
        resp.headers["Content-Security-Policy"] = CSP_DOC
    return resp


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


# =====================================================================================
# v8.2.0 — ШТРАФЫ: разбор постановлений и текста проверки (ГИБДД, Госуслуги, банк)
# Автоматической бесплатной проверки по госномеру/СТС нет (сайт ГИБДД — с капчей, официального API нет),
# поэтому штрафы приходят из документов (PDF/фото постановления) и из текста страницы проверки,
# который пользователь копирует сам. Разбор — здесь, привязка к машине/водителю — в программе.
# =====================================================================================
_PLATE_LETTERS = "АВЕКМНОРСТУХABEKMHOPCTYX"
_PLATE_RE = re.compile(r"(?<![0-9A-Za-zА-Яа-я])([" + _PLATE_LETTERS + r"])\s?(\d{3})\s?([" + _PLATE_LETTERS + r"]{2})\s?(\d{2,3})(?![0-9])", re.I)
_UIN_RE = re.compile(r"(?<!\d)(\d(?:[\s ]?\d){19}(?:(?:[\s ]?\d){5})?)(?!\d)")
_DATE_RE = r"(0?[1-9]|[12]\d|3[01])[./](0?[1-9]|1[0-2])[./](20[2-3]\d)"


def _fine_date(m):
    return f"{int(m.group(3)):04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"


def _money(s):
    s = re.sub(r"[\s ]", "", s or "").replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def parse_fines_text(text: str) -> list:
    t = (text or "").replace("\r", "")
    if not t.strip():
        return []
    # УИН — 20 или 25 цифр (у ГИБДД начинается с 188…); номера длиннее/короче пропускаем
    uins = []
    for m in _UIN_RE.finditer(t):
        d = re.sub(r"\D", "", m.group(1))
        if len(d) in (20, 25) and d not in [u for u, _, _ in uins]:
            ctx = t[max(0, m.start() - 60):m.start()].lower()
            if d.startswith("188") or "уин" in ctx or "постановлен" in ctx or "идентификатор" in ctx:
                uins.append((d, m.start(), m.end()))
    blocks = []
    if uins:
        for i, (u, a, b) in enumerate(uins):
            prev_end = uins[i - 1][2] if i else 0
            end = uins[i + 1][1] if i + 1 < len(uins) else min(len(t), b + 1500)
            # данные штрафа — сначала ПОСЛЕ его УИН (до следующего), иначе в коротком куске ПЕРЕД ним
            blocks.append((u, t[b:end], t[max(prev_end, a - 350):a]))
    elif re.search(r"постановлен|штраф", t, re.I) and re.search(r"12\.\d", t):
        blocks.append(("", t, ""))
    out = []
    for uin, after, before in blocks:
        f = {"uin": uin}
        # запись вида «номер → данные» (ГИБДД, постановление): кусок перед номером — чужой штраф
        if re.search(r"\d\s*(?:руб|₽)", after, re.I):
            before = ""

        def find(pattern, flags=re.I):
            return re.search(pattern, after, flags) or (re.search(pattern, before, flags) if before else None)

        m = find(r"(?:дат[аы]\s+(?:и\s+время\s+)?(?:совершения\s+)?(?:правонарушения|нарушения)|нарушени[ея]\s+от|время\s+нарушения)\D{0,25}" + _DATE_RE + r"(?:\D{0,12}?([01]?\d|2[0-3])[:.]([0-5]\d))?")
        if not m:
            m = find(_DATE_RE + r"\s*(?:г\.?)?\s*(?:в\s*)?([01]?\d|2[0-3])[:.]([0-5]\d)", 0)
        if m:
            f["date"] = _fine_date(m)
            if m.lastindex and m.lastindex >= 5 and m.group(4):
                f["time"] = f"{int(m.group(4)):02d}:{m.group(5)}"
        m = re.search(r"^\D{0,6}от\s+" + _DATE_RE, after, re.I) or find(r"(?:постановлени[ея][^\n]{0,40}?от|дата\s+постановления|вынесено)\D{0,15}" + _DATE_RE)
        if m:
            f["decisionDate"] = _fine_date(m)
        m = find(r"(?:сумма\s+(?:штрафа|к\s+оплате|начисления)?|размер\s+штрафа|штраф\w*\s+в\s+размере|к\s+оплате|начислено)\D{0,20}?(\d[\d\s ]{0,9}(?:[.,]\d{1,2})?)\s*(?:руб|₽|р\.)")
        if not m:
            m = find(r"(\d[\d\s ]{1,9}(?:[.,]\d{2})?)\s*(?:руб|₽)")
        if m:
            f["amount"] = _money(m.group(1))
        m = find(r"(?:скидк\w*|со\s+скидкой|льготн\w*)[^\n]{0,80}?до\s+" + _DATE_RE)
        if m:
            f["discountUntil"] = _fine_date(m)
        m = find(r"(?:со\s+скидкой|с\s+учетом\s+скидки|сумма\s+со\s+скидкой)\D{0,20}(\d[\d\s ]{0,9}(?:[.,]\d{1,2})?)\s*(?:руб|₽)")
        if not m:
            m = find(r"скидк[^\n]{0,80}?до\s+" + _DATE_RE + r"\D{0,12}?(\d[\d\s ]{0,9}(?:[.,]\d{1,2})?)\s*(?:руб|₽)")
            if m:
                f["discountAmount"] = _money(m.group(4))
                m = None
        if m:
            f["discountAmount"] = _money(m.group(1))
        src = after if re.search(r"12\.\d", after) else before
        m = re.search(r"(?:ст(?:атья|\.)?\s*)?(12\.\d{1,2}(?:\.\d)?)\s*(?:,?\s*ч(?:асть|\.)?\s*(\d+(?:\.\d)?))?(?:\s*коап(?:\s*рф)?)?", src, re.I)
        if m:
            f["article"] = m.group(1) + (f" ч.{m.group(2)}" if m.group(2) else "")
            tail = src[m.end():m.end() + 200].split("\n")[0].strip(" -–—:.,")
            if len(tail) >= 6 and not re.match(r"^\d", tail):
                f["description"] = tail[:160]
        m = find(r"место\s+(?:совершения\s+)?(?:правонарушения|нарушения)\s*[:\-]?\s*([^\n]{6,200})")
        if m:
            f["place"] = m.group(1).strip()[:200]
        m = _PLATE_RE.search(after) or (_PLATE_RE.search(before) if before else None)
        if m:
            f["plate"] = (m.group(1) + m.group(2) + m.group(3) + m.group(4)).upper()
        m = find(r"(?:стс|свидетельств\w*\s+о\s+регистрации(?:\s+тс)?)\D{0,20}(\d{2}\s?[0-9А-ЯA-Z]{2}\s?\d{6})")
        if m:
            f["sts"] = re.sub(r"\s", "", m.group(1)).upper()
        m = re.search(r"(ЦАФАП[^\n]{0,120}|ОБ\s+ДПС[^\n]{0,120}|ГИБДД[^\n]{0,80})", after) or (re.search(r"(ЦАФАП[^\n]{0,120}|ГИБДД[^\n]{0,80})", before) if before else None)
        if m:
            f["issuer"] = m.group(1).strip()[:160]
        near = after[:700].lower()
        f["status"] = "paid" if re.search(r"(?<!не\s)оплачен|погашен|исполнен", near) and not re.search(r"не\s+оплачен|к\s+оплате", near) else "unpaid"
        if f.get("amount") or f.get("uin"):
            out.append(f)
    return out


def _fines_with_profiles(fines: list, profiles: list) -> list:
    """Подсказка, чья машина: по госномеру или СТС из постановления."""
    for f in fines:
        pids = []
        for p in profiles or []:
            plate = _plate_norm(p.get("carPlate"))
            sts = re.sub(r"\s", "", str(p.get("carSts") or "")).upper()
            if (f.get("plate") and plate and _plate_norm(f["plate"]) == plate) or (f.get("sts") and sts and f["sts"] == sts):
                pids.append(p.get("id"))
        f["carProfileIds"] = pids
    return fines


@app.post("/api/fines/parse")
def fines_parse(payload: dict = Body(...)):
    text = str(payload.get("text") or "")[:300000]
    st = load_valid_state()
    return {"status": "ok", "fines": _fines_with_profiles(parse_fines_text(text), st.get("profiles", []))}


# =====================================================================================
# v8.2.0 — СРОКИ И НАПОМИНАНИЯ: ОСАГО, ВУ, техосмотр, медсправка, документы со сроком, штрафы
# В программе — список «Сроки» (/api/deadlines). По желанию — сообщения в Telegram через бота
# владельца: токен хранится только на сервере (data/notify.json), в окно не отдаётся.
# В сообщения идут только ФИО кратко, госномер, вид документа и дата — без СНИЛС/номеров ВУ.
# =====================================================================================
NOTIFY_FILE = os.path.join(DATA_DIR, "notify.json")
_NOTIFY_LOCK = threading.Lock()
NOTIFY_DEFAULT_DAYS = [30, 14, 7, 3, 1, 0]
PROFILE_DEADLINE_FIELDS = [
    ("licenseValidUntil", "Водительское удостоверение", "Водительское удостоверение"),
    ("osagoUntil", "Полис ОСАГО", "Страховка (ОСАГО, КАСКО)"),
    ("tehosmotrUntil", "Техосмотр (диагностическая карта)", "Техосмотр (диагностическая карта)"),
    ("medUntil", "Медицинская справка", "Медицинские документы"),
]


def _notify_load() -> dict:
    try:
        with open(NOTIFY_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _notify_save(d: dict):
    tmp = NOTIFY_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    try:
        os.chmod(tmp, 0o600)
    except Exception:
        pass
    os.replace(tmp, NOTIFY_FILE)


def _iso_date(s):
    try:
        return datetime.date.fromisoformat(str(s or "")[:10])
    except Exception:
        return None


def collect_deadlines(st: dict, today=None, profile_filter: str = "") -> list:
    today = today or datetime.date.today()
    items = []
    profiles = st.get("profiles") or []
    byid = {p.get("id"): p for p in profiles}

    def who(p):
        return (p or {}).get("driverShort") or (p or {}).get("driverFull") or "без имени"

    def add(key, label, p, until, source, extra=None):
        it = {"key": key, "label": label, "profileId": (p or {}).get("id", ""), "driver": who(p) if p else "",
              "car": (p or {}).get("carPlate", ""), "until": until.isoformat(), "daysLeft": (until - today).days, "source": source}
        if extra:
            it.update(extra)
        items.append(it)

    for p in profiles:
        if profile_filter and p.get("id") != profile_filter:
            continue
        for field, label, _ in PROFILE_DEADLINE_FIELDS:
            d = _iso_date(p.get(field))
            if d:
                add(f"{p.get('id')}:{field}:{d}", label, p, d, "карточка")
    for doc in st.get("savedDocuments") or []:
        d = _iso_date(doc.get("validUntil"))
        if not d:
            continue
        p = byid.get(doc.get("profileId"))
        if profile_filter and doc.get("profileId") != profile_filter:
            continue
        # срок из карточки главнее: документ того же вида не дублируем
        if p and any(kind == doc.get("kind") and _iso_date(p.get(field)) for field, _, kind in PROFILE_DEADLINE_FIELDS):
            continue
        add(f"doc:{doc.get('id')}:{d}", doc.get("kind") or doc.get("fileName") or "Документ", p, d, "документ",
            {"docId": doc.get("id")})
    for f in st.get("fines") or []:
        if f.get("status") in ("paid", "cancelled"):
            continue
        p = byid.get(f.get("profileId"))
        if profile_filter and f.get("profileId") != profile_filter:
            continue
        amount = f.get("amount")
        what = f"штраф {('%g' % amount) if amount else ''} ₽".replace("  ", " ") + (f", ст. {f['article']}" if f.get("article") else "")
        d = _iso_date(f.get("discountUntil"))
        if d and d >= today:
            add(f"fine:{f.get('id')}:disc:{d}", f"Скидка на {what}", p, d, "штраф", {"fineId": f.get("id"), "car": f.get("plate") or (p or {}).get("carPlate", "")})
        dd = _iso_date(f.get("payUntil"))
        if not dd:
            dec = _iso_date(f.get("decisionDate"))
            dd = dec + datetime.timedelta(days=70) if dec else None  # ≈ 10 дней на вступление в силу + 60 на оплату
        if dd:
            add(f"fine:{f.get('id')}:pay:{dd}", f"Оплатить {what}", p, dd, "штраф", {"fineId": f.get("id"), "car": f.get("plate") or (p or {}).get("carPlate", "")})
    items.sort(key=lambda x: x["daysLeft"])
    return items


@app.get("/api/deadlines")
def api_deadlines(request: Request):
    u = request.state.user
    pid = "" if _is_admin(u) or (u.get("assignedProfileId") in (None, "", "ALL")) else u.get("assignedProfileId")
    return {"status": "ok", "today": datetime.date.today().isoformat(), "items": collect_deadlines(load_valid_state(), profile_filter=pid)}


def _tg_send(token: str, chat: str, text: str):
    try:
        data = json.dumps({"chat_id": chat, "text": text[:4000], "disable_web_page_preview": True}).encode("utf-8")
        req = urllib.request.Request(f"https://api.telegram.org/bot{token}/sendMessage", data=data,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=20) as r:
            ok = json.load(r).get("ok") is True
        return ok, "" if ok else "Telegram ответил ошибкой"
    except urllib.error.HTTPError as e:
        try:
            desc = json.load(e).get("description", "")
        except Exception:
            desc = ""
        return False, f"Telegram: {e.code} {desc}".strip()
    except Exception as e:
        return False, f"нет связи с Telegram ({type(e).__name__})"


def _fmt_ru(iso):
    d = _iso_date(iso)
    return d.strftime("%d.%m.%Y") if d else str(iso or "")


def _deadline_line(it):
    dl = it["daysLeft"]
    when = "сегодня" if dl == 0 else (f"через {dl} дн." if dl > 0 else f"ПРОСРОЧЕНО на {-dl} дн.")
    who = ", ".join(x for x in (it.get("driver"), it.get("car")) if x)
    return f"• {it['label']}{' — ' + who if who else ''}: до {_fmt_ru(it['until'])} ({when})"


def _notify_check() -> dict:
    """v8.3.0: каналы — Telegram и/или почта. Каждый канал помнит, что уже отправил (почта — с приставкой «mail:»)."""
    with _NOTIFY_LOCK:
        cfg = _notify_load()
        tg = cfg.get("telegram") or {}
        tg_on = bool(tg.get("enabled") and tg.get("token") and tg.get("chats"))
        mc = _smtp_load()
        mail_on = bool(mc.get("notifyEnabled") and mc.get("notifyTo") and _smtp_ready(mc))
        if not tg_on and not mail_on:
            return {"sent": 0, "reason": "уведомления выключены или не настроены"}
        st = load_valid_state()
        today = datetime.date.today()
        days = sorted({int(x) for x in (cfg.get("days") or NOTIFY_DEFAULT_DAYS) if str(x).lstrip("-").isdigit()}, reverse=True)
        sent = cfg.setdefault("sent", {})
        items = []  # (вид, ключ, строка)
        for it in collect_deadlines(st, today):
            dl = it["daysLeft"]
            if dl < 0:
                mark = "expired"
            else:
                ths = [d for d in days if dl <= d]
                if not ths:
                    continue
                mark = str(min(ths))
            items.append(("deadline", it["key"] + "|" + mark, _deadline_line(it)))
        byid = {p.get("id"): p for p in st.get("profiles") or []}
        for f in st.get("fines") or []:
            if f.get("status") in ("paid", "cancelled") or not f.get("id"):
                continue
            p = byid.get(f.get("profileId"))
            parts = [f.get("plate") or (p or {}).get("carPlate", ""), f"{('%g' % f['amount']) if f.get('amount') else '?'} ₽",
                     " ".join(x for x in (_fmt_ru(f.get("date")) if f.get("date") else "", f.get("time", "")) if x),
                     f"ст. {f['article']}" if f.get("article") else "",
                     f"водитель: {(p or {}).get('driverShort') or (p or {}).get('driverFull')}" if p else "водитель не определён",
                     f"скидка до {_fmt_ru(f['discountUntil'])}" if f.get("discountUntil") else ""]
            items.append(("fine", f"fine-new:{f.get('id')}", "• " + "; ".join(x for x in parts if x)))

        def build(sel):
            fines = [x[2] for x in sel if x[0] == "fine"]
            dls = [x[2] for x in sel if x[0] == "deadline"]
            text = ""
            if fines:
                text += "🚨 Новые штрафы (ПУТЁВКА):\n" + "\n".join(fines) + "\n\n"
            if dls:
                text += "⏰ Сроки документов (ПУТЁВКА):\n" + "\n".join(dls)
            return text.strip()

        total, errs = 0, []
        for ch in (["tg"] if tg_on else []) + (["mail"] if mail_on else []):
            pref = "" if ch == "tg" else "mail:"
            pend = [x for x in items if pref + x[1] not in sent]
            if not pend:
                continue
            text = build(pend)
            ok_all = True
            if ch == "tg":
                for chat in tg.get("chats") or []:
                    ok, err = _tg_send(tg["token"], str(chat), text)
                    if not ok:
                        ok_all = False
                        errs.append("Telegram: " + err)
            else:
                ok, err = _mail_send(mc.get("notifyTo") or [], "ПУТЁВКА: сроки документов и штрафы", text, mc)
                if not ok:
                    ok_all = False
                    errs.append("почта: " + err)
            if ok_all:
                for x in pend:
                    sent[pref + x[1]] = today.isoformat()
                total += len(pend)
        cfg["lastRun"] = datetime.datetime.now().isoformat(timespec="seconds")
        cfg["lastError"] = "; ".join(sorted(set(errs)))
        cutoff = (today - datetime.timedelta(days=400)).isoformat()
        cfg["sent"] = {k: v for k, v in sent.items() if v >= cutoff}
        _notify_save(cfg)
        return {"sent": total, "error": cfg["lastError"]}


def _notify_loop():
    time.sleep(30)
    while True:
        try:
            _notify_check()
        except Exception:
            pass
        time.sleep(1800)


@app.on_event("startup")
def _start_notify_thread():
    threading.Thread(target=_notify_loop, daemon=True).start()


@app.get("/api/notify/settings")
def notify_settings_get():
    cfg = _notify_load()
    tg = cfg.get("telegram") or {}
    tok = tg.get("token") or ""
    return {"status": "ok", "enabled": bool(tg.get("enabled")), "hasToken": bool(tok),
            "tokenHint": ("…" + tok[-4:]) if tok else "", "chats": tg.get("chats") or [],
            "days": cfg.get("days") or NOTIFY_DEFAULT_DAYS, "lastRun": cfg.get("lastRun", ""), "lastError": cfg.get("lastError", "")}


@app.post("/api/notify/settings")
def notify_settings_set(payload: dict = Body(...)):
    with _NOTIFY_LOCK:
        cfg = _notify_load()
        tg = cfg.setdefault("telegram", {})
        tok = str(payload.get("token") or "").strip()
        if payload.get("clearToken"):
            tg["token"] = ""
        elif tok:
            if not re.fullmatch(r"\d{5,}:[A-Za-z0-9_-]{30,}", tok):
                raise HTTPException(status_code=400, detail="Это не похоже на токен бота (вида 123456789:AA…). Его выдаёт @BotFather.")
            tg["token"] = tok
        if "chats" in payload:
            chats = [str(c).strip() for c in (payload.get("chats") or []) if str(c).strip()]
            for c in chats:
                if not re.fullmatch(r"-?\d{3,20}|@[A-Za-z0-9_]{4,64}", c):
                    raise HTTPException(status_code=400, detail=f"Неверный адрес чата: {c} (нужен числовой id или @канал)")
            tg["chats"] = chats[:10]
        if "enabled" in payload:
            tg["enabled"] = bool(payload.get("enabled"))
        if "days" in payload:
            d = sorted({int(x) for x in payload.get("days") or [] if str(x).isdigit() and 0 <= int(x) <= 365}, reverse=True)
            cfg["days"] = d or NOTIFY_DEFAULT_DAYS
        _notify_save(cfg)
    return notify_settings_get()


@app.post("/api/notify/test")
def notify_test():
    cfg = _notify_load()
    tg = cfg.get("telegram") or {}
    if not tg.get("token") or not tg.get("chats"):
        return {"status": "error", "message": "Сначала укажите токен бота и хотя бы один чат"}
    res = [{"chat": c, "ok": ok, "error": err} for c in tg["chats"]
           for ok, err in [_tg_send(tg["token"], str(c), "✅ ПУТЁВКА: проверка уведомлений. Сюда будут приходить напоминания о сроках ОСАГО, ВУ, документов и о новых штрафах.")]]
    return {"status": "ok" if all(r["ok"] for r in res) else "error", "results": res}


@app.post("/api/notify/run")
def notify_run():
    return {"status": "ok", **_notify_check()}


# =====================================================================================
# v8.3.0 — НАДЁЖНЫЙ ВХОД, ВОССТАНОВЛЕНИЕ ПАРОЛЯ, ПОЧТА, РЕЗЕРВНАЯ КОПИЯ БАЗЫ
#  • пароль не короче 8 символов, не совпадает с логином, не из распространённых;
#  • подбор: 10 ошибок с адреса за 10 мин (было) + 5 ошибок на логин → вход этим логином закрыт на 15 мин;
#  • сеанс закрывается после 24 ч бездействия (и через 7 дней в любом случае);
#  • код восстановления администратора (выдаётся при первичной настройке, меняется в «Настройках»);
#  • сброс пароля кодом из письма — почта своего ящика (SMTP: Яндекс, Mail.ru, Gmail и др.);
#  • та же почта — для напоминаний о сроках и штрафах (вместе с Telegram или вместо него);
#  • выгрузка всей базы одним файлом, загрузка на другом компьютере, чистая база; ежедневная копия.
# Секреты (хэши кодов, пароль почты) — в отдельных файлах data/, браузер их не получает.
# =====================================================================================
import smtplib, ssl, tempfile
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from starlette.background import BackgroundTask

SECURITY_FILE = os.path.join(DATA_DIR, "security.json")
SECURITY_LOG = os.path.join(DATA_DIR, "security.log")
SMTP_FILE = os.path.join(DATA_DIR, "smtp.json")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
_SEC_LOCK = threading.Lock()
_BACKUP_LOCK = threading.Lock()
PW_MIN = 8
ACCT_MAX_FAILS, ACCT_WINDOW, ACCT_LOCK_SEC = 5, 900, 900
_ACCT_FAILS = {}
_ACCT_LOCK_UNTIL = {}
_COMMON_PW = {
    "password", "password1", "passw0rd", "qwerty123", "qwertyui", "qwerty12", "1q2w3e4r", "1qaz2wsx", "zaq12wsx",
    "iloveyou", "admin123", "administrator", "welcome1", "letmein1", "11111111", "12345678", "123456789",
    "1234567890", "87654321", "00000000", "abc12345", "abcd1234", "12341234", "11223344", "qazwsxedc",
    "пароль123", "йцукенгш", "putevka1", "путевка1", "путёвка1", "admin1234", "user1234", "driver12",
}
_SEQ_ROWS = ["01234567890", "abcdefghijklmnopqrstuvwxyz", "qwertyuiopasdfghjklzxcvbnm",
             "абвгдеёжзийклмнопрстуфхцчшщъыьэюя", "йцукенгшщзхъфывапролджэячсмитьбю"]
_RC_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"


def password_problem(pw, login: str = "") -> str:
    """Почему пароль не годится («» — годится). Одни правила для всех способов задать пароль."""
    pw = str(pw or "")
    if len(pw) < PW_MIN:
        return f"Пароль должен быть не короче {PW_MIN} символов"
    low = pw.lower()
    if login and low == str(login).strip().lower():
        return "Пароль не должен совпадать с логином"
    if len(set(low)) <= 2:
        return "Пароль слишком простой: почти одни и те же символы"
    if low in _COMMON_PW or any(low in row or low in row[::-1] for row in _SEQ_ROWS):
        return "Пароль слишком простой: распространённый или символы подряд с клавиатуры"
    return ""


def _sec_load() -> dict:
    try:
        with open(SECURITY_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _sec_save(d: dict):
    tmp = SECURITY_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    try:
        os.chmod(tmp, 0o600)
    except Exception:
        pass
    os.replace(tmp, SECURITY_FILE)


ERRORS_LOG = os.path.join(DATA_DIR, "errors.log")


def _err_log(where: str, exc: Exception, extra: str = ""):
    """v8.5.1: журнал ошибок data/errors.log (сбои записи файлов и т. п.) — чтобы причину можно было найти."""
    try:
        with open(ERRORS_LOG, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {where}: {type(exc).__name__}: {exc} {extra}\n")
    except Exception:
        pass


def _sec_log(event: str, request: Request = None, login: str = ""):
    """Журнал безопасности data/security.log: входы, блокировки, смены паролей, загрузка базы."""
    try:
        ip = (request.client.host if request is not None and request.client else "-")
        line = f"{datetime.datetime.now().isoformat(timespec='seconds')}\t{ip}\t{event}\t{str(login)[:64]}\n"
        if os.path.exists(SECURITY_LOG) and os.path.getsize(SECURITY_LOG) > 1_000_000:
            os.replace(SECURITY_LOG, SECURITY_LOG + ".1")
        with open(SECURITY_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _acct_key(login: str) -> str:
    return str(login or "").strip().lower()[:64]


def _acct_locked_left(login: str) -> int:
    left = _ACCT_LOCK_UNTIL.get(_acct_key(login), 0) - time.time()
    return int(left) if left > 0 else 0


def _acct_failed(login: str, request: Request = None):
    k, now = _acct_key(login), time.time()
    if not k:
        return
    with _LOGIN_LOCK:
        lst = [t for t in _ACCT_FAILS.get(k, []) if now - t < ACCT_WINDOW] + [now]
        _ACCT_FAILS[k] = lst
        if len(lst) >= ACCT_MAX_FAILS:
            _ACCT_LOCK_UNTIL[k] = now + ACCT_LOCK_SEC
            _ACCT_FAILS[k] = []
            _sec_log("login-locked", request, login)


def _acct_ok(login: str):
    k = _acct_key(login)
    with _LOGIN_LOCK:
        _ACCT_FAILS.pop(k, None)
        _ACCT_LOCK_UNTIL.pop(k, None)


def _lock_message(left: int) -> str:
    return f"Вход этим логином временно закрыт из-за неверных паролей. Попробуйте через {max(1, (left + 59) // 60)} мин."


def _norm_code(code: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", str(code or "")).upper()


def _new_recovery_code(uid: str) -> str:
    raw = "".join(secrets.choice(_RC_ALPHABET) for _ in range(16))
    code = "-".join(raw[i:i + 4] for i in range(0, 16, 4))
    with _SEC_LOCK:
        d = _sec_load()
        d.setdefault("recovery", {})[uid] = _make_password(raw)
        _sec_save(d)
    return code


def _set_weak(uid: str, weak: bool):
    with _SEC_LOCK:
        d = _sec_load()
        w = d.setdefault("weak", {})
        if weak:
            w[uid] = True
        else:
            w.pop(uid, None)
        _sec_save(d)


def _set_user_password(uid: str, new_pw: str) -> dict:
    st = load_valid_state()
    u = next((x for x in st.get("users", []) or [] if x.get("id") == uid), None)
    if not u:
        raise HTTPException(status_code=404, detail="Пользователь не найден")
    u["password"] = _make_password(new_pw)
    save_valid_state(st)
    _set_weak(uid, False)
    _drop_user_sessions(uid)
    return u


def _find_user_by_login(login: str):
    lk = _acct_key(login)
    return next((x for x in (load_valid_state().get("users", []) or []) if _acct_key(x.get("login")) == lk and lk), None)


# ---------- почта своего ящика (SMTP) ----------
SMTP_PRESETS = {"yandex": ("smtp.yandex.ru", 465, "ssl"), "mailru": ("smtp.mail.ru", 465, "ssl"),
                "gmail": ("smtp.gmail.com", 587, "starttls")}


def _smtp_load() -> dict:
    try:
        with open(SMTP_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _smtp_save(d: dict):
    tmp = SMTP_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    try:
        os.chmod(tmp, 0o600)
    except Exception:
        pass
    os.replace(tmp, SMTP_FILE)


def _smtp_ready(cfg: dict = None) -> bool:
    cfg = cfg if cfg is not None else _smtp_load()
    return bool(cfg.get("host") and (cfg.get("fromAddr") or cfg.get("user")))


_EMAIL_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


def _mail_send(to, subject: str, text: str, cfg: dict = None):
    """(успех, ошибка). Письмо простым текстом через почтовый сервер своего ящика."""
    cfg = cfg if cfg is not None else _smtp_load()
    to = [t for t in ([to] if isinstance(to, str) else list(to or [])) if _EMAIL_RE.match(str(t))]
    if not _smtp_ready(cfg):
        return False, "почта для отправки не настроена (Настройки приложения → Почта)"
    if not to:
        return False, "не указан адрес получателя"
    sender = cfg.get("fromAddr") or cfg.get("user")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr((cfg.get("fromName") or "ПУТЁВКА", sender))
    msg["To"] = ", ".join(to)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(sender.split("@")[-1] if "@" in sender else "putevka.local"))
    msg.set_content(text)
    sec = cfg.get("security") or "ssl"
    port = int(cfg.get("port") or (465 if sec == "ssl" else 587))
    try:
        ctx = ssl.create_default_context()
        if sec == "ssl":
            srv = smtplib.SMTP_SSL(cfg["host"], port, timeout=25, context=ctx)
        else:
            srv = smtplib.SMTP(cfg["host"], port, timeout=25)
            srv.ehlo()
            if sec == "starttls":
                srv.starttls(context=ctx)
                srv.ehlo()
        try:
            if cfg.get("user"):
                srv.login(cfg["user"], cfg.get("password") or "")
            srv.send_message(msg)
        finally:
            try:
                srv.quit()
            except Exception:
                pass
        return True, ""
    except smtplib.SMTPAuthenticationError:
        return False, "почтовый сервер не принял логин или пароль (для Яндекса, Mail.ru и Gmail нужен «пароль приложения», не обычный пароль)"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"[:300]


@app.get("/api/mail/settings")
def mail_settings_get():
    c = _smtp_load()
    pw = c.get("password") or ""
    return {"status": "ok", "host": c.get("host", ""), "port": c.get("port") or "", "security": c.get("security") or "ssl",
            "user": c.get("user", ""), "fromName": c.get("fromName", ""), "fromAddr": c.get("fromAddr", ""),
            "hasPassword": bool(pw), "notifyEnabled": bool(c.get("notifyEnabled")), "notifyTo": c.get("notifyTo") or [],
            "ready": _smtp_ready(c), "lastError": c.get("lastError", "")}


@app.post("/api/mail/settings")
def mail_settings_set(payload: dict = Body(...)):
    with _SEC_LOCK:
        c = _smtp_load()
        preset = str(payload.get("preset") or "")
        if preset in SMTP_PRESETS:
            c["host"], c["port"], c["security"] = SMTP_PRESETS[preset]
        if "host" in payload:
            host = str(payload.get("host") or "").strip()
            if host and not re.fullmatch(r"[A-Za-z0-9.\-]{1,253}", host):
                raise HTTPException(status_code=400, detail="Адрес почтового сервера — вида smtp.yandex.ru")
            c["host"] = host
        if "port" in payload and str(payload.get("port") or "").strip():
            try:
                port = int(payload.get("port"))
                assert 1 <= port <= 65535
            except Exception:
                raise HTTPException(status_code=400, detail="Порт — число, обычно 465 или 587")
            c["port"] = port
        if "security" in payload:
            sec = str(payload.get("security") or "ssl")
            if sec not in ("ssl", "starttls", "none"):
                raise HTTPException(status_code=400, detail="Шифрование: ssl, starttls или none")
            c["security"] = sec
        for k in ("user", "fromName", "fromAddr"):
            if k in payload:
                c[k] = str(payload.get(k) or "").strip()[:200]
        if c.get("fromAddr") and not _EMAIL_RE.match(c["fromAddr"]):
            raise HTTPException(status_code=400, detail="Адрес отправителя — вида name@yandex.ru")
        if payload.get("clearPassword"):
            c["password"] = ""
        elif str(payload.get("password") or ""):
            c["password"] = str(payload.get("password"))[:200]
        if "notifyTo" in payload:
            lst = [str(x).strip() for x in (payload.get("notifyTo") or []) if str(x).strip()]
            bad = [x for x in lst if not _EMAIL_RE.match(x)]
            if bad:
                raise HTTPException(status_code=400, detail="Неверный адрес почты: " + ", ".join(bad[:3]))
            c["notifyTo"] = lst[:10]
        if "notifyEnabled" in payload:
            c["notifyEnabled"] = bool(payload.get("notifyEnabled"))
        _smtp_save(c)
    return mail_settings_get()


@app.post("/api/mail/test")
def mail_test(request: Request, payload: dict = Body(default={})):
    to = str((payload or {}).get("to") or "").strip()
    if not to:
        c = _smtp_load()
        to = (c.get("notifyTo") or [""])[0] or c.get("fromAddr") or c.get("user") or ""
    ok, err = _mail_send(to, "ПУТЁВКА: проверка почты",
                         "Это проверочное письмо программы «Путёвка».\n\nСюда будут приходить коды для сброса пароля и, если "
                         "включено, напоминания о сроках документов и о новых штрафах.\n")
    with _SEC_LOCK:
        c = _smtp_load()
        c["lastError"] = err
        _smtp_save(c)
    return {"status": "ok" if ok else "error", "to": to, "message": err or f"Письмо отправлено на {to}"}


# ---------- вход: смена пароля, код восстановления, сброс ----------
@app.post("/api/auth/change-password")
def auth_change_password(request: Request, payload: dict = Body(...)):
    u = request.state.user
    cur = str(payload.get("current") or "")
    new = str(payload.get("new") or "")
    ok, _ = _check_password(u.get("password") or "", cur)
    if not ok:
        _acct_failed(u.get("login"), request)
        _sec_log("password-change-bad-current", request, u.get("login"))
        raise HTTPException(status_code=403, detail="Текущий пароль указан неверно")
    prob = password_problem(new, u.get("login"))
    if prob:
        raise HTTPException(status_code=400, detail=prob)
    if new == cur:
        raise HTTPException(status_code=400, detail="Новый пароль совпадает с текущим")
    _set_user_password(u["id"], new)
    _sec_log("password-changed", request, u.get("login"))
    resp = JSONResponse({"status": "ok", "message": "Пароль изменён"})
    _start_session(resp, u["id"], secure=request.url.scheme == "https")
    return resp


@app.post("/api/auth/recovery-code")
def auth_recovery_code(request: Request, payload: dict = Body(...)):
    u = request.state.user
    ok, _ = _check_password(u.get("password") or "", str(payload.get("password") or ""))
    if not ok:
        _acct_failed(u.get("login"), request)
        raise HTTPException(status_code=403, detail="Пароль указан неверно")
    code = _new_recovery_code(u["id"])
    _sec_log("recovery-code-new", request, u.get("login"))
    return {"status": "ok", "recoveryCode": code}


_GENERIC_FORGOT = ("Если у этого логина указана почта и в программе настроена отправка писем, код уже отправлен — он "
                   "действует 15 минут. Письма нет — введите код восстановления администратора или попросите "
                   "администратора задать новый пароль (раздел «Пользователи и доступ»).")


@app.post("/api/auth/forgot")
def auth_forgot(request: Request, payload: dict = Body(...)):
    ip = request.client.host if request.client else "?"
    if _login_blocked(ip):
        return JSONResponse({"status": "error", "message": "Слишком много попыток. Подождите 10 минут."}, status_code=429)
    login = str(payload.get("login") or "").strip()
    _login_failed(ip)  # каждый запрос кода считается попыткой: рассылку кодов не превратить в спам
    u = _find_user_by_login(login)
    cfg = _smtp_load()
    sent = False
    if u and u.get("active") is not False and _EMAIL_RE.match(str(u.get("email") or "")) and _smtp_ready(cfg):
        code = f"{secrets.randbelow(1_000_000):06d}"
        with _SEC_LOCK:
            d = _sec_load()
            d.setdefault("resets", {})[u["id"]] = {"hash": _make_password(code), "exp": time.time() + 900, "tries": 0}
            _sec_save(d)
        sent, err = _mail_send(u["email"], "ПУТЁВКА: код для сброса пароля",
                               f"Код для сброса пароля в программе «Путёвка»: {code}\n\nЛогин: {u.get('login')}\n"
                               "Код действует 15 минут. Если вы не просили сбросить пароль — просто удалите это письмо: "
                               "без кода пароль не изменится.\n", cfg)
        _sec_log("reset-code-sent" if sent else "reset-code-mail-error", request, login)
    else:
        _sec_log("reset-code-not-sent", request, login)
    return {"status": "ok", "message": _GENERIC_FORGOT, "mailReady": _smtp_ready(cfg)}


@app.post("/api/auth/reset")
def auth_reset(request: Request, payload: dict = Body(...)):
    ip = request.client.host if request.client else "?"
    login = str(payload.get("login") or "").strip()
    if _login_blocked(ip):
        return JSONResponse({"status": "error", "message": "Слишком много неудачных попыток. Подождите 10 минут."}, status_code=429)
    left = _acct_locked_left(login)
    if left:
        return JSONResponse({"status": "error", "message": _lock_message(left)}, status_code=429)
    code = _norm_code(payload.get("code"))
    new = str(payload.get("newPassword") or "")
    u = _find_user_by_login(login)
    used = ""
    if u and code:
        with _SEC_LOCK:
            d = _sec_load()
            r = (d.get("resets") or {}).get(u["id"])
            if r and r.get("exp", 0) > time.time() and r.get("tries", 0) < 5:
                if _check_password(r.get("hash") or "", code)[0]:
                    used = "email"
                else:
                    r["tries"] = r.get("tries", 0) + 1
                    _sec_save(d)
            if not used and _check_password((d.get("recovery") or {}).get(u["id"]) or "", code)[0]:
                used = "recovery"
    if not used:
        _login_failed(ip)
        _acct_failed(login, request)
        _sec_log("reset-bad-code", request, login)
        time.sleep(0.4)
        return JSONResponse({"status": "error", "message": "Код неверный или устарел"}, status_code=400)
    prob = password_problem(new, u.get("login"))
    if prob:
        return JSONResponse({"status": "error", "message": prob}, status_code=400)
    if u.get("active") is False:
        return JSONResponse({"status": "error", "message": "Учётная запись отключена администратором"}, status_code=403)
    _set_user_password(u["id"], new)
    with _SEC_LOCK:
        d = _sec_load()
        (d.get("resets") or {}).pop(u["id"], None)
        if used == "recovery":
            (d.get("recovery") or {}).pop(u["id"], None)
        _sec_save(d)
    out = {"status": "ok", "message": "Пароль изменён — войдите с новым паролем"}
    if used == "recovery":
        out["recoveryCode"] = _new_recovery_code(u["id"])  # старый код одноразовый — сразу выдаём новый
    _acct_ok(login)
    _sec_log("password-reset-" + used, request, login)
    return out


# ---------- резервная копия базы ----------
_BK_NAME = re.compile(r"^putevka_backup_[0-9A-Za-z_\-]+\.zip$")


def _backup_make(reason: str, dest_dir: str = None, with_documents: bool = True) -> str:
    dest_dir = dest_dir or BACKUP_DIR
    os.makedirs(dest_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    path = os.path.join(dest_dir, f"putevka_backup_{ts}_{reason}.zip")
    tmp = path + ".part"
    st = load_valid_state()
    docs = []
    if with_documents and os.path.isdir(DOCS_DIR):
        for root, _dirs, files in os.walk(DOCS_DIR):
            for fn in files:
                full = os.path.join(root, fn)
                docs.append((full, "documents/" + os.path.relpath(full, DOCS_DIR).replace(os.sep, "/")))
    manifest = {"app": "putevka", "appVersion": APP_VERSION, "created": datetime.datetime.now().isoformat(timespec="seconds"),
                "reason": reason, "profiles": len(st.get("profiles") or []), "users": len(st.get("users") or []),
                "documents": len(docs), "withDocuments": with_documents}
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as zf:
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=1))
        if os.path.exists(DB_FILE):
            zf.write(DB_FILE, "putevka_db.json")
        if os.path.exists(SECURITY_FILE):
            zf.write(SECURITY_FILE, "security.json")
        for full, arc in docs:
            zf.write(full, arc)
    os.replace(tmp, path)
    if dest_dir == BACKUP_DIR:
        _backup_prune()
    return path


def _backup_prune():
    try:
        files = sorted((f for f in os.listdir(BACKUP_DIR) if _BK_NAME.match(f)), reverse=True)
        auto = [f for f in files if f.endswith("_auto.zip")]
        other = [f for f in files if not f.endswith("_auto.zip")]
        for f in auto[14:] + other[5:]:
            os.remove(os.path.join(BACKUP_DIR, f))
    except Exception:
        pass


def _confirm_admin_password(request: Request, pw: str):
    u = request.state.user
    left = _acct_locked_left(u.get("login"))
    if left:
        raise HTTPException(status_code=429, detail=_lock_message(left))
    ok, _ = _check_password(u.get("password") or "", pw or "")
    if not ok:
        _acct_failed(u.get("login"), request)
        raise HTTPException(status_code=403, detail="Пароль указан неверно")


@app.get("/api/backup/export")
def backup_export(request: Request):
    tmpdir = tempfile.mkdtemp(prefix="putevka_export_")
    with _BACKUP_LOCK:
        path = _backup_make("export", dest_dir=tmpdir)
    _sec_log("backup-export", request, request.state.user.get("login"))
    name = "Путёвка_база_" + datetime.datetime.now().strftime("%Y-%m-%d_%H%M") + ".zip"
    return FileResponse(path, media_type="application/zip", filename=name,
                        background=BackgroundTask(shutil.rmtree, tmpdir, True))


def _safe_member(name: str) -> bool:
    if not name or name.startswith(("/", "\\")) or ":" in name or "\\" in name:
        return False
    parts = name.split("/")
    return ".." not in parts and all(p not in ("", ".") for p in parts[:-1])


@app.post("/api/backup/import")
async def backup_import(request: Request):
    try:
        pw = base64.b64decode(request.headers.get("x-confirm-password-b64") or "").decode("utf-8")
    except Exception:
        pw = ""
    _confirm_admin_password(request, pw)
    fd, tmp = tempfile.mkstemp(prefix="putevka_import_", suffix=".zip")
    size = 0
    try:
        with os.fdopen(fd, "wb") as f:
            async for chunk in request.stream():
                size += len(chunk)
                if size > 4 * 1024 ** 3:
                    raise HTTPException(status_code=413, detail="Файл больше 4 ГБ")
                f.write(chunk)
        try:
            zf = zipfile.ZipFile(tmp)
        except Exception:
            raise HTTPException(status_code=400, detail="Это не файл резервной копии (.zip программы «Путёвка»)")
        with zf:
            infos = zf.infolist()
            names = {i.filename for i in infos}
            if "putevka_db.json" not in names:
                raise HTTPException(status_code=400, detail="В файле нет базы (putevka_db.json) — это не резервная копия «Путёвки»")
            if len(infos) > 100_000 or sum(i.file_size for i in infos) > 8 * 1024 ** 3:
                raise HTTPException(status_code=400, detail="Файл слишком большой или повреждён")
            for i in infos:
                if i.is_dir():
                    continue
                if not _safe_member(i.filename) or not (i.filename in ("manifest.json", "putevka_db.json", "security.json")
                                                         or i.filename.startswith("documents/")):
                    raise HTTPException(status_code=400, detail=f"Недопустимый файл в архиве: {i.filename[:80]}")
            try:
                st = json.loads(zf.read("putevka_db.json").decode("utf-8"))
            except Exception:
                raise HTTPException(status_code=400, detail="База в файле повреждена (не читается)")
            if not isinstance(st, dict) or not isinstance(st.get("profiles"), list) or not isinstance(st.get("users", []), list):
                raise HTTPException(status_code=400, detail="База в файле не похожа на базу «Путёвки»")
            if not any(x.get("role") == "admin" and x.get("password") and x.get("active") is not False
                       for x in st.get("users") or [] if isinstance(x, dict)):
                raise HTTPException(status_code=400, detail="В загружаемой базе нет администратора с паролем — войти в неё будет нельзя")
            with _BACKUP_LOCK:
                before = _backup_make("before-import")
                new_docs = DOCS_DIR + ".import"
                shutil.rmtree(new_docs, ignore_errors=True)
                os.makedirs(new_docs, exist_ok=True)
                for i in infos:
                    if i.is_dir() or not i.filename.startswith("documents/"):
                        continue
                    dest = os.path.join(new_docs, *i.filename[len("documents/"):].split("/"))
                    os.makedirs(os.path.dirname(dest), exist_ok=True)
                    with zf.open(i) as src, open(dest, "wb") as out:
                        shutil.copyfileobj(src, out)
                with open(DB_FILE + ".import", "w", encoding="utf-8") as f:
                    json.dump(st, f, ensure_ascii=False, indent=2)
                os.replace(DB_FILE + ".import", DB_FILE)
                if "security.json" in names:
                    with open(SECURITY_FILE + ".import", "wb") as f:
                        f.write(zf.read("security.json"))
                    os.replace(SECURITY_FILE + ".import", SECURITY_FILE)
                elif os.path.exists(SECURITY_FILE):
                    os.remove(SECURITY_FILE)
                old_docs = DOCS_DIR + ".old"
                shutil.rmtree(old_docs, ignore_errors=True)
                if os.path.isdir(DOCS_DIR):
                    os.replace(DOCS_DIR, old_docs)
                os.replace(new_docs, DOCS_DIR)
                shutil.rmtree(old_docs, ignore_errors=True)
                with _SESS_LOCK:
                    _SESSIONS.clear()
                    _sessions_save()
                _USERS_CACHE["mtime"] = -1.0
        _sec_log("backup-import", request, request.state.user.get("login"))
        return {"status": "ok", "backupBefore": os.path.basename(before), "profiles": len(st.get("profiles") or []),
                "users": len(st.get("users") or []),
                "message": "База загружена. Войдите логином и паролем из загруженной базы."}
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


@app.post("/api/backup/reset")
def backup_reset(request: Request, payload: dict = Body(...)):
    _confirm_admin_password(request, str(payload.get("password") or ""))
    if str(payload.get("confirm") or "").strip().upper() not in ("ОЧИСТИТЬ", "ОЧИСТИТЬ БАЗУ"):
        raise HTTPException(status_code=400, detail="Для подтверждения введите слово ОЧИСТИТЬ")
    with _BACKUP_LOCK:
        before = _backup_make("before-reset")
        tpl = INIT_FILE if os.path.exists(INIT_FILE) else os.path.join(BUNDLE_DIR, "data", "app_data.json")
        with open(tpl, "r", encoding="utf-8") as f:
            st = json.load(f)
        st["users"] = []
        st["profiles"] = []
        st["savedDocuments"] = []
        st["savedFolders"] = []
        st["fines"] = []
        st["adminSetupCompleted"] = False
        with open(DB_FILE + ".reset", "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
        os.replace(DB_FILE + ".reset", DB_FILE)
        shutil.rmtree(DOCS_DIR, ignore_errors=True)
        os.makedirs(DOCS_DIR, exist_ok=True)
        if os.path.exists(SECURITY_FILE):
            os.remove(SECURITY_FILE)
        with _SESS_LOCK:
            _SESSIONS.clear()
            _sessions_save()
        _USERS_CACHE["mtime"] = -1.0
    _sec_log("backup-reset", request, request.state.user.get("login"))
    return {"status": "ok", "backupBefore": os.path.basename(before),
            "message": "База очищена. Создайте администратора заново. Прежняя база сохранена в резервных копиях."}


@app.get("/api/backup/list")
def backup_list():
    out = []
    try:
        for f in sorted((x for x in os.listdir(BACKUP_DIR) if _BK_NAME.match(x)), reverse=True):
            p = os.path.join(BACKUP_DIR, f)
            m = re.match(r"^putevka_backup_(\d{4}-\d{2}-\d{2})_(\d{2})(\d{2})(\d{2})_(.+)\.zip$", f)
            out.append({"name": f, "size": os.path.getsize(p), "reason": m.group(5) if m else "",
                        "created": (f"{m.group(1)} {m.group(2)}:{m.group(3)}" if m else "")})
    except FileNotFoundError:
        pass
    return {"status": "ok", "items": out}


@app.get("/api/backup/file/{name}")
def backup_file(name: str):
    if not _BK_NAME.match(name or "") or not os.path.exists(os.path.join(BACKUP_DIR, name)):
        raise HTTPException(status_code=404, detail="Нет такой копии")
    return FileResponse(os.path.join(BACKUP_DIR, name), media_type="application/zip", filename=name)


def _backup_daily_loop():
    """Раз в сутки — копия базы без документов (на случай поломки), хранятся 14 последних."""
    time.sleep(120)
    while True:
        try:
            today = datetime.date.today().isoformat()
            have = any(f.startswith("putevka_backup_" + today) and f.endswith("_auto.zip")
                       for f in (os.listdir(BACKUP_DIR) if os.path.isdir(BACKUP_DIR) else []))
            st = load_valid_state()
            if not have and (st.get("users") or st.get("profiles")):
                with _BACKUP_LOCK:
                    _backup_make("auto", with_documents=False)
        except Exception:
            pass
        time.sleep(3600)


@app.on_event("startup")
def _start_backup_thread():
    threading.Thread(target=_backup_daily_loop, daemon=True).start()



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
    if _host in ("127.0.0.1", "localhost", "::1"):
        LOCAL_ONLY = True
    # v8.2.0: HTTPS — если заданы сертификат и ключ (для серверной версии, доступной из сети)
    _cert, _key = os.environ.get("PUTEVKA_SSL_CERTFILE"), os.environ.get("PUTEVKA_SSL_KEYFILE")
    if _cert and _key:
        uvicorn.run(app, host=_host, port=_port, ssl_certfile=_cert, ssl_keyfile=_key)
    else:
        uvicorn.run(app, host=_host, port=_port)
