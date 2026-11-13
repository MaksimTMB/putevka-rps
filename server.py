import os, sys
if sys.stdout is None: sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w', encoding='utf-8')
import os, io, shutil, re, csv, base64, json, datetime, zipfile, urllib.parse, urllib.request
from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "v7.0.5"

app = FastAPI(title="ПУТЁВКА v7.0.5 — Форма № 3 (ОКУД 0345001)")

import sys
if getattr(sys, "frozen", False):
    BUNDLE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = BUNDLE_DIR

DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
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
    return {
        "1,18": prof.get("waybillSeries") or "", "1,22": day_obj.get("waybill_num") or prof.get("waybillNumberPrefix") or "", "3,7": d_str, "3,13": d_str, "4,22": prof.get("okud") or "0345001", "5,4": prof.get("orgFull") or "", "5,22": prof.get("okpo") or "",
        "8,5": prof.get("carModel") or "", "9,8": prof.get("carPlate") or "", "9,22": prof.get("garageNumber") or "", "10,2": prof.get("driverFull") or "", "10,22": prof.get("tabNumber") or "", "12,5": prof.get("license") or "", "12,20": prof.get("licenseClass") or "В",
        "13,6": prof.get("snils") or "", "14,6": prof.get("transportInfo") or "перевозка для собственных нужд", "17,0": prof.get("medOrgPre") or "", "28,0": prof.get("orgShort") or "", "28,22": prof.get("mechanic") or "", "34,23": prof.get("dispatcher") or "",
        "37,22": prof.get("driverShort") or "", "40,20": "АИ", "40,24": fg, "41,10": d_str, "42,10": prof.get("depTimeDefault") or "9-00", "44,9": prof.get("dispatcher") or "", "47,10": d_str, "48,10": prof.get("arrTimeDefault") or "18-00", "50,9": prof.get("dispatcher") or "",
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
    return FileResponse(
        os.path.join(BUNDLE_DIR if not os.path.exists(os.path.join(BASE_DIR, "static", "index.html")) else BASE_DIR, "static", "index.html"),
        headers={"Cache-Control": "no-store, no-cache, must-revalidate, max-age=0"},
    )


@app.get("/api/version")
def get_version():
    st = load_valid_state()
    return {
        "version": APP_VERSION,
        "release_date": "2026-11-09",
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
            u["password"] = password
            u["fullName"] = full_name
            u["shortName"] = full_name
            u["active"] = True
            admin_user = u
            break
    if not admin_user:
        admin_user = {
            "id": "usr_admin",
            "login": login,
            "password": password,
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
        if u.get("login") == login and u.get("password") == password:
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

        # Check separate time column or other cells for time if t_str not found yet
        if not t_str:
            if "time" in col_map and col_map["time"] < len(row):
                t_val = row[col_map["time"]]
                if isinstance(t_val, (datetime.time, datetime.datetime)):
                    t_str = f"{t_val.hour:02d}:{t_val.minute:02d}"
                elif t_val is not None:
                    tm = re.search(r"\b([01]?\d|2[0-3])[:\-]([0-5]\d)\b", str(t_val))
                    if tm:
                        t_str = f"{int(tm.group(1)):02d}:{tm.group(2)}"
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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)

UPDATE_REPO = "MaksimTMB/putevka-rps"
UPDATE_FILES = ["version.txt", "server.py", "desktop_app.py", "requirements.txt", "static/index.html", "data/app_data.json"]


def _version_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:3])


def _fetch_latest_release():
    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest",
            headers={"User-Agent": "PutevkaUpdater/1.0", "Accept": "application/vnd.github+json"},
        )
        with urllib.request.urlopen(req, timeout=8) as r:
            return str(json.load(r).get("tag_name") or APP_VERSION)
    except Exception:
        return APP_VERSION


@app.get('/api/update/check')
def api_upd_check():
    latest = _fetch_latest_release()
    update_available = _version_tuple(latest) > _version_tuple(APP_VERSION)
    return {'status': 'ok', 'currentVersion': APP_VERSION, 'latestVersion': latest, 'updateAvailable': update_available}


@app.post('/api/update/apply')
def api_upd_apply():
    latest = _fetch_latest_release()
    if os.path.exists(DB_FILE):
        try:
            shutil.copy2(DB_FILE, DB_FILE + ".backup.json")
        except Exception:
            pass

    updated, failed = [], []
    for rel in UPDATE_FILES:
        url = f"https://raw.githubusercontent.com/{UPDATE_REPO}/main/{rel}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "PutevkaUpdater/1.0"})
            with urllib.request.urlopen(req, timeout=25) as r:
                data = r.read()
            dest = os.path.join(BASE_DIR, rel)
            parent = os.path.dirname(dest)
            if parent:
                os.makedirs(parent, exist_ok=True)
            with open(dest, "wb") as f:
                f.write(data)
            updated.append(rel)
        except Exception:
            failed.append(rel)

    if updated:
        return {
            'status': 'ok',
            'updatedTo': latest,
            'updated': updated,
            'failed': failed,
            'message': f'Обновлено файлов: {len(updated)}. Перезапустите приложение (или сервер), чтобы изменения вступили в силу. Рабочая база сохранена.',
        }
    return {
        'status': 'error',
        'updatedTo': APP_VERSION,
        'updated': [],
        'failed': failed,
        'message': 'Не удалось загрузить обновление. Проверьте доступ к интернету.',
    }
