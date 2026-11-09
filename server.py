import os, io, json, datetime, zipfile, urllib.parse
from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_VERSION = "v6.2.0"

app = FastAPI(title="ПУТЁВКА v6.0.0 — Форма № 3 (ОКУД 0345001)")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "data", "putevka_db.json")
INIT_FILE = os.path.join(BASE_DIR, "data", "app_data.json")


def load_valid_state():
    for path in (DB_FILE, INIT_FILE):
        if os.path.exists(path) and os.path.getsize(path) > 50:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict) and data.get("profiles"):
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
    d_str = day_obj.get("date_ru", "")
    fg = str(day_obj.get("fuel_grade") or prof.get("defaultFuelGrade") or "92")
    return {
        "1,18": prof.get("waybillSeries") or "",
        "1,22": day_obj.get("waybill_num") or prof.get("waybillNumberPrefix") or "",
        "3,7": d_str,
        "3,13": d_str,
        "4,22": prof.get("okud") or "0345001",
        "5,4": prof.get("orgFull") or "",
        "5,22": prof.get("okpo") or "",
        "8,5": prof.get("carModel") or "",
        "9,8": prof.get("carPlate") or "",
        "9,22": prof.get("garageNumber") or "",
        "10,2": prof.get("driverFull") or "",
        "10,22": prof.get("tabNumber") or "",
        "12,5": prof.get("license") or "",
        "12,20": prof.get("licenseClass") or "В",
        "13,6": prof.get("snils") or "",
        "14,6": prof.get("transportInfo") or "перевозка для собственных нужд",
        "17,0": prof.get("medOrgPre") or "",
        "28,0": prof.get("orgShort") or "",
        "28,22": prof.get("mechanic") or "",
        "34,23": prof.get("dispatcher") or "",
        "37,22": prof.get("driverShort") or "",
        "40,20": "АИ",
        "40,24": fg,
        "41,10": d_str,
        "42,10": prof.get("depTimeDefault") or "9-00",
        "44,9": prof.get("dispatcher") or "",
        "47,10": d_str,
        "48,10": prof.get("arrTimeDefault") or "18-00",
        "50,9": prof.get("dispatcher") or "",
        "62,22": prof.get("dispatcher") or "",
        "67,9": prof.get("driverShort") or "",
        "68,22": prof.get("driverShort") or "",
    }


def get_back_overrides(prof: dict, day_obj: dict) -> dict:
    wb_num = day_obj.get("waybill_num") or prof.get("waybillNumberPrefix") or ""
    ov = {
        "1,0": f"Оборотная сторона путевого листа легкового автомобиля № {wb_num} от {day_obj.get('date_ru', '')}",
        "24,7": day_obj.get("km", 0),
        "25,7": day_obj.get("city_km", 0),
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
                cell.value = "=X7"
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
        os.path.join(BASE_DIR, "static", "index.html"),
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


@app.post("/api/export/docx")
def export_docx(payload: dict = Body(...)):
    st = load_valid_state()
    schema = st.get("xls_schema")
    prof = payload.get("profile")
    days_list = payload.get("days", [])
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
    days_list = payload.get("days", [])
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


app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
