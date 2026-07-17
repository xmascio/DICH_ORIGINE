from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
from xml.sax.saxutils import escape as xml_escape
import csv
import getpass
import io
import json
import os
import re
import subprocess
import sys
import textwrap
import time
import zipfile

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas


BASE_DIR = Path(__file__).resolve().parent
ASSET_DIR = BASE_DIR / "assets"
LOGO_PATH = ASSET_DIR / "logo-nordlaser.png"
SIGNATURE_PATH = ASSET_DIR / "firma-nordlaser.png"
FOOTER_TEMPLATE_PATH = ASSET_DIR / "footer-template.png"
SHARED_LOG_PATH = Path(r"\\nas01\QUALITA\7-UTENTI\MASCELLONI\UTILITY\AUTOMAZIONI\DASHBOARD_DICH_ORIGINE\registro_attestazioni.csv")
SHARED_LOG_LOCK = SHARED_LOG_PATH.with_suffix(".lock")
LOG_FIELDS = [
    "data_ora",
    "utente",
    "lingua",
    "codice_cliente",
    "cliente",
    "validita_da",
    "validita_a",
    "articoli",
    "file_pdf",
]

COUNTRY_GROUPS = {
    "GP001": "AL;BA;CH;EG;FO;GE;IS;LI;MD;ME;MK;NO;PS;RS;UA",
    "GP002": "BW;CI;CL;DZ;GH;IL;JO;KM;LB;LS;MA;MG;MU;MX;MZ;NA;SC;SZ;TN;TR;XC;XL;ZA;ZW",
    "GP003": "CA;GB;JP;KR;VN",
    "GP004": "CO;EC;PE;SG",
    "GP005": "AG;BB;BS;BZ;CM;CR;DM;DO;FJ;GD;GT;GY;HN;JM;KN;LC;NC;NI;PA;PF;PG;PM;SB;SR;SV;TT;VC;WS",
    "MERCOSUR": "AR;BR;PY;UY",
}


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states) + 1
        self._saved_page_states.append(dict(self.__dict__))
        for state in self._saved_page_states:
            self.__dict__.update(state)
            draw_footer(self, total_pages)
            super().showPage()
        super().save()


def safe_filename(value):
    cleaned = re.sub(r'[<>:"/\\|?*]+', "-", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or "dichiarazione-origine"


def acquire_log_lock(timeout=10):
    deadline = time.time() + timeout
    SHARED_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    while True:
        try:
            return os.open(str(SHARED_LOG_LOCK), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                if time.time() - SHARED_LOG_LOCK.stat().st_mtime > 60:
                    SHARED_LOG_LOCK.unlink()
                    continue
            except FileNotFoundError:
                continue
            if time.time() >= deadline:
                raise RuntimeError("Registro stampe temporaneamente occupato. Riprova tra qualche secondo.")
            time.sleep(0.1)


def release_log_lock(handle):
    os.close(handle)
    try:
        SHARED_LOG_LOCK.unlink()
    except FileNotFoundError:
        pass


def append_print_log(data, output_path, language, materials):
    handle = acquire_log_lock()
    try:
        write_header = not SHARED_LOG_PATH.exists() or SHARED_LOG_PATH.stat().st_size == 0
        with SHARED_LOG_PATH.open("a", newline="", encoding="utf-8-sig") as stream:
            writer = csv.DictWriter(stream, fieldnames=LOG_FIELDS, delimiter=";")
            if write_header:
                writer.writeheader()
            client = data.get("cliente", {})
            writer.writerow({
                "data_ora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "utente": os.environ.get("USERNAME") or getpass.getuser(),
                "lingua": language.upper(),
                "codice_cliente": client.get("code", ""),
                "cliente": client.get("name", ""),
                "validita_da": parse_date(data.get("validFrom", "")),
                "validita_a": parse_date(data.get("validTo", "")),
                "articoli": len(materials),
                "file_pdf": str(output_path),
            })
    finally:
        release_log_lock(handle)


def read_print_log():
    if not SHARED_LOG_PATH.exists():
        return []
    handle = acquire_log_lock()
    try:
        with SHARED_LOG_PATH.open("r", newline="", encoding="utf-8-sig") as stream:
            return list(csv.DictReader(stream, delimiter=";"))
    finally:
        release_log_lock(handle)


def xlsx_column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def build_xlsx(rows):
    headers = [
        "Data e ora", "Utente", "Lingua", "Codice cliente", "Cliente",
        "Validita da", "Validita a", "Articoli", "File PDF",
    ]
    keys = LOG_FIELDS
    sheet_rows = [headers] + [[str(row.get(key, "")) for key in keys] for row in rows]
    xml_rows = []
    for row_index, values in enumerate(sheet_rows, 1):
        cells = []
        for column_index, value in enumerate(values, 1):
            ref = f"{xlsx_column_name(column_index)}{row_index}"
            style = ' s="1"' if row_index == 1 else ""
            cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t>{xml_escape(value)}</t></is></c>')
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="20" customWidth="1"/>'
        '<col min="3" max="4" width="16" customWidth="1"/><col min="5" max="5" width="38" customWidth="1"/>'
        '<col min="6" max="8" width="15" customWidth="1"/><col min="9" max="9" width="70" customWidth="1"/></cols>'
        f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'
    )
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        archive.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Attestazioni" sheetId="1" r:id="rId1"/></sheets></workbook>')
        archive.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        archive.writestr("xl/styles.xml", '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf xfId="0"/><xf xfId="0" fontId="1" applyFont="1"/></cellXfs></styleSheet>')
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
    return output.getvalue()


def select_folder():
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d=New-Object System.Windows.Forms.FolderBrowserDialog;"
        "$d.Description='Seleziona la cartella di destinazione dei PDF';"
        "$d.ShowNewFolderButton=$true;"
        "if($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        "Write-Output $d.SelectedPath}"
    )
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-STA", "-Command", script],
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        creationflags=flags,
        timeout=300,
    )
    if result.returncode:
        raise RuntimeError("Impossibile aprire la selezione cartella.")
    return result.stdout.strip()


def parse_date(value):
    parts = value.split("-")
    if len(parts) == 3:
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return value


def wrap(text, width):
    return textwrap.wrap(str(text or ""), width=width, break_long_words=False) or [""]


def draw_wrapped(c, text, x, y, width_chars, leading=4.5 * mm, font="Helvetica", size=10):
    c.setFont(font, size)
    for line in wrap(text, width_chars):
        c.drawString(x, y, line)
        y -= leading
    return y


def draw_header(c):
    width, height = A4
    margin = 18 * mm
    y = height - 9 * mm

    if LOGO_PATH.exists():
        c.drawImage(str(LOGO_PATH), margin, y - 19.8 * mm, width=63 * mm, height=17.1 * mm, preserveAspectRatio=True, mask="auto")

    c.setFont("Helvetica", 7.5)
    c.setFillColor(colors.HexColor("#526074"))
    c.drawRightString(width - margin, y - 16 * mm, "Nord Laser s.r.l.")
    c.drawRightString(width - margin, y - 21 * mm, "Via L. Da Vinci s.n.c. - 33010 Reana del Rojale, UD - Italy")
    c.setFillColor(colors.black)

    line_y = height - 33 * mm
    c.setStrokeColor(colors.HexColor("#9aa6b2"))
    c.setLineWidth(0.6)
    c.line(margin, line_y, width - margin, line_y)
    c.setStrokeColor(colors.black)
    return line_y - 9 * mm


def draw_footer(c, total_pages):
    width, _height = A4
    margin = 18 * mm
    file_name = getattr(c, "_pdf_file_name", "")
    c.saveState()
    if FOOTER_TEMPLATE_PATH.exists():
        c.drawImage(str(FOOTER_TEMPLATE_PATH), margin, 8 * mm, width=width - (2 * margin), height=15.8 * mm, preserveAspectRatio=True, mask="auto")
    c.setFont("Helvetica", 7)
    c.setFillColor(colors.HexColor("#1d2433"))
    c.drawRightString(width - margin, 4 * mm, f"Pag {c.getPageNumber()} di {total_pages}")
    if file_name:
        c.setFont("Helvetica", 6)
        c.setFillColor(colors.HexColor("#526074"))
        c.drawString(margin, 4 * mm, file_name)
    c.restoreState()


def ensure_page(c, y, needed=42 * mm):
    if y > needed:
        return y
    c.showPage()
    return draw_header(c)


def language_texts(language, cliente, start_date, end_date):
    if language == "en":
        return {
            "file_label": "Origin declaration",
            "title": "Long-term supplier's declaration for products having preferential originating status",
            "subtitle": "under preferential trade arrangements pursuant to Commission Implementing Regulation (EU) No. 2015/2447 as amended.",
            "paragraphs": [
                "By signing this declaration:",
                "the undersigned Marco Bovolini, acting as President of Nord Laser s.r.l.",
                "located at Via L. Da Vinci s.n.c. Reana del Rojale (UD)",
                f"with reference to the products regularly supplied to {cliente}",
                "declares and guarantees pursuant to Commission Implementing Regulation (EU) No. 2447/2015 that:",
                "All products listed below are of ITALY origin and comply with the rules of origin governing preferential trade with the countries listed in the table:",
            ],
            "headers": ["ITEM CODE", "DESCRIPTION", "COUNTRY GROUP"],
            "customs_header": "CUSTOMS CODE",
            "groups_label": "* Country groups",
            "closing": [
                "It also declares that: cumulation not applied",
                f"The supplier undertakes to immediately notify {cliente} of any loss of validity of this declaration.",
                "The supplier undertakes to submit to the Customs Authorities all documentation concerning the country of origin and/or preferential origin of the products sold to customers, upon simple request.",
                f"This declaration is valid for all shipments of said products from {start_date} to {end_date}",
            ],
            "place_date": "Reana del Rojale",
            "signature": "Signature",
        }

    return {
        "file_label": "Dichiarazione origine",
        "title": "Dichiarazione a lungo termine del fornitore per prodotti aventi carattere originario",
        "subtitle": "nell'ambito di un regime preferenziale ai sensi del Regolamento di Esecuzione (UE) n. 2015/2447 e ss.mm.ii.",
        "paragraphs": [
            "Con la sottoscrizione della presente dichiarazione:",
            "il sottoscritto Marco Bovolini, nel ruolo di Presidente della societa Nord Laser s.r.l.",
            "sita in Via L. Da Vinci s.n.c. Reana del Rojale (UD)",
            f"con riferimento ai prodotti, regolarmente forniti a {cliente}",
            "dichiara e garantisce, ai sensi del Regolamento di Esecuzione della Commissione (UE) n.2447/2015, che:",
            "Tutti i prodotti di seguito elencati hanno origine ITALIA e rispettano le norme di origine che disciplinano gli scambi preferenziali con i Paesi elencati in tabella:",
        ],
        "headers": ["CODICE ARTICOLO", "DESCRIZIONE", "GRUPPO PAESE"],
        "customs_header": "CODICE DOGANALE",
        "groups_label": "* Gruppi paese",
        "closing": [
            "Dichiara inoltre che: cumulo non applicato",
            f"Il fornitore si impegna a notificare immediatamente a {cliente} della perdita di validita della presente dichiarazione.",
            "Il fornitore si impegna a presentare alle Autorita Doganali tutta la documentazione sul paese di origine e/o sulla origine preferenziale dei prodotti venduti ai clienti, dietro semplice richiesta.",
            f"La presente dichiarazione e valida per tutti gli invii di detti prodotti dal {start_date} al {end_date}",
        ],
        "place_date": "Reana del Rojale",
        "signature": "Firma",
    }


def generate_pdf(data):
    output_dir = Path(data.get("outputDir", "")).expanduser()
    if not output_dir:
        raise ValueError("Inserisci un percorso di destinazione.")
    output_dir.mkdir(parents=True, exist_ok=True)

    cliente = data.get("cliente", {}).get("name", "").strip()
    if not cliente:
        raise ValueError("Seleziona o inserisci un cliente.")

    if not data.get("validFrom") or not data.get("validTo"):
        raise ValueError("Inserisci data inizio validita e data fine validita.")
    start_date = parse_date(data.get("validFrom", ""))
    end_date = parse_date(data.get("validTo", ""))
    language = "en" if data.get("language") == "en" else "it"
    texts = language_texts(language, cliente, start_date, end_date)

    materials = [m for m in data.get("materials", []) if m.get("code") or m.get("description")]
    if not materials:
        raise ValueError("Inserisci almeno una riga materiale.")
    show_customs_code = any(str(m.get("customsCode", "")).strip() for m in materials)

    print_day = datetime.now()
    cliente_file = cliente.strip(" .")
    file_name = safe_filename(f"{print_day.strftime('%Y%m%d_%H%M%S')} - {texts['file_label']} {cliente_file}.pdf")
    output_path = output_dir / file_name

    c = NumberedCanvas(str(output_path), pagesize=A4)
    c._pdf_file_name = file_name
    width, _height = A4
    margin = 18 * mm
    y = draw_header(c)

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, y, texts["title"])
    y -= 5 * mm
    c.setFont("Helvetica", 9)
    c.drawCentredString(width / 2, y, texts["subtitle"])
    y -= 12 * mm

    for paragraph in texts["paragraphs"]:
        y = ensure_page(c, y)
        y = draw_wrapped(c, paragraph, margin, y, 112, leading=5 * mm, size=10)
        y -= 1 * mm

    y -= 2 * mm
    if show_customs_code:
        col_w = [34 * mm, 72 * mm, 34 * mm, 28 * mm]
        headers = [texts["headers"][0], texts["headers"][1], texts["customs_header"], texts["headers"][2]]
    else:
        col_w = [38 * mm, 92 * mm, 38 * mm]
        headers = texts["headers"]
    table_x = (width - sum(col_w)) / 2
    row_h = 8 * mm
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(colors.HexColor("#eef2f7"))
    c.rect(table_x, y - row_h, sum(col_w), row_h, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#3c4654"))
    x = table_x
    for i, header in enumerate(headers):
        c.rect(x, y - row_h, col_w[i], row_h, fill=0, stroke=1)
        c.drawString(x + 2 * mm, y - 5 * mm, header)
        x += col_w[i]
    y -= row_h

    c.setFillColor(colors.black)
    for item in materials:
        lines_code = wrap(item.get("code", ""), 18)
        lines_desc = wrap(item.get("description", ""), 40 if show_customs_code else 48)
        lines_customs = wrap(item.get("customsCode", ""), 16)
        lines_groups = ["*"]
        cells = [lines_code, lines_desc, lines_customs, lines_groups] if show_customs_code else [lines_code, lines_desc, lines_groups]
        lines = max(*(len(cell) for cell in cells), 1)
        h = max(row_h, (lines * 4.4 + 4) * mm)
        y = ensure_page(c, y, h + 42 * mm)
        x = table_x
        for w in col_w:
            c.rect(x, y - h, w, h, fill=0, stroke=1)
            x += w
        c.setFont("Helvetica", 8.5)
        x = table_x
        for i, cell_lines in enumerate(cells):
            ty = y - 5 * mm
            for line in cell_lines:
                c.drawString(x + 2 * mm, ty, line)
                ty -= 4.4 * mm
            x += col_w[i]
        y -= h

    y -= 8 * mm
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(margin, y, texts["groups_label"])
    y -= 5 * mm
    c.setFont("Helvetica", 8)
    for group, countries in COUNTRY_GROUPS.items():
        y = ensure_page(c, y)
        y = draw_wrapped(c, f"{group}: {countries}", margin, y, 118, leading=4.2 * mm, size=8)

    y -= 4 * mm
    for paragraph in texts["closing"]:
        y = ensure_page(c, y)
        y = draw_wrapped(c, paragraph, margin, y, 104, leading=5 * mm, size=10)
        y -= 1 * mm

    y = ensure_page(c, y, 65 * mm)
    y -= 11 * mm
    c.setFont("Helvetica", 10)
    print_date = print_day.strftime("%d/%m/%Y")
    c.drawString(margin, y, f"{texts['place_date']}, {print_date}")
    y -= 9 * mm
    c.drawString(margin, y, texts["signature"])
    if SIGNATURE_PATH.exists():
        c.drawImage(str(SIGNATURE_PATH), margin + 18 * mm, y - 20 * mm, width=55 * mm, height=25 * mm, preserveAspectRatio=True, mask="auto")

    c.save()
    append_print_log(data, output_path, language, materials)
    return output_path


class Handler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_POST(self):
        if self.path == "/api/select-folder":
            try:
                response = {"ok": True, "path": select_folder()}
                self.send_response(200)
            except Exception as exc:
                response = {"ok": False, "error": str(exc)}
                self.send_response(500)
            return self.send_json(response)
        if self.path != "/api/generate":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length).decode("utf-8"))
            output_path = generate_pdf(data)
            response = {"ok": True, "path": str(output_path)}
            self.send_response(200)
        except Exception as exc:
            response = {"ok": False, "error": str(exc)}
            self.send_response(400)
        self.send_json(response)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/logs":
            try:
                rows = list(reversed(read_print_log()))
                self.send_response(200)
                return self.send_json({"ok": True, "rows": rows})
            except Exception as exc:
                self.send_response(500)
                return self.send_json({"ok": False, "error": str(exc)})
        if parsed.path == "/api/export-log":
            try:
                rows = read_print_log()
                export_format = parse_qs(parsed.query).get("format", ["xlsx"])[0]
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                if export_format == "txt":
                    output = io.StringIO()
                    writer = csv.writer(output, delimiter="\t", lineterminator="\r\n")
                    writer.writerow(["Data e ora", "Utente", "Lingua", "Codice cliente", "Cliente", "Validita da", "Validita a", "Articoli", "File PDF"])
                    writer.writerows([[row.get(key, "") for key in LOG_FIELDS] for row in rows])
                    body = output.getvalue().encode("utf-8-sig")
                    content_type = "text/plain; charset=utf-8"
                    filename = f"registro_attestazioni_{stamp}.txt"
                else:
                    body = build_xlsx(rows)
                    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    filename = f"registro_attestazioni_{stamp}.xlsx"
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            except Exception as exc:
                self.send_response(500)
                return self.send_json({"ok": False, "error": str(exc)})
        super().do_GET()

    def send_json(self, response):
        body = json.dumps(response, ensure_ascii=False).encode("utf-8")
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def translate_path(self, path):
        path = unquote(path.split("?", 1)[0].split("#", 1)[0])
        if path == "/":
            path = "/index.html"
        return str(BASE_DIR / path.lstrip("/"))


def main():
    os.chdir(BASE_DIR)
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    actual_port = server.server_address[1]
    port_file = Path(os.environ.get("LOCALAPPDATA", BASE_DIR)) / "NordLaser" / "DichiarazioniOrigine" / "dashboard.port"
    port_file.parent.mkdir(parents=True, exist_ok=True)
    port_file.write_text(str(actual_port), encoding="ascii")
    try:
        server.serve_forever()
    finally:
        try:
            port_file.unlink()
        except FileNotFoundError:
            pass


if __name__ == "__main__":
    main()
