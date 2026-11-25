import csv
import functools
import io
import json
import mimetypes
import sys
import uuid
from pathlib import Path
from typing import Literal

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from PIL import Image as PILImage
from pptx import Presentation
from pydantic import BaseModel, Field

SCOPES = ["https://www.googleapis.com/auth/drive"]  # full read/write on the user's Drive
HERE = Path(__file__).parent
TOKEN = HERE / "token.json"
mcp = MCPServer("gdrive")

DOC = "application/vnd.google-apps.document"
SHEET = "application/vnd.google-apps.spreadsheet"
SLIDES = "application/vnd.google-apps.presentation"
DRAWING = "application/vnd.google-apps.drawing"
EXPORT_AS = {DOC: "text/markdown", SLIDES: "text/plain", DRAWING: "image/svg+xml"}
KIND_TO_MIME = {"doc": DOC, "sheet": SHEET, "slides": SLIDES}


# ---------- auth ----------

def login():
    """Interactive OAuth login (opens a browser). Run once: python gdrive_server.py auth"""
    creds = InstalledAppFlow.from_client_secrets_file(
        HERE / "credentials.json", SCOPES).run_local_server(port=0)
    TOKEN.write_text(creds.to_json())
    print(f"Saved {TOKEN} with scopes {SCOPES}")


def credentials() -> Credentials:
    # The MCP server runs over stdio, so it must never open a browser itself.
    relogin = f"run: {sys.executable} {Path(__file__)} auth"
    if not TOKEN.exists():
        raise ToolError(f"Not logged in to Google Drive; {relogin}")
    if not set(SCOPES) <= set(json.loads(TOKEN.read_text()).get("scopes", [])):
        raise ToolError(f"Saved Google token lacks the required scopes {SCOPES}; {relogin}")
    creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
    if not creds.valid:
        if not creds.refresh_token:
            raise ToolError(f"Google token expired and cannot be refreshed; {relogin}")
        creds.refresh(Request())
        TOKEN.write_text(creds.to_json())
    return creds


@functools.cache
def service(name: str, version: str):
    # Built once per server process; google-auth refreshes the access token on its own.
    return build(name, version, credentials=credentials())


def drive():
    return service("drive", "v3")


def tool(fn):
    """Register an MCP tool whose Google API errors reach the client as readable messages."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except HttpError as e:
            raise ToolError(f"Google API error {e.status_code}: {e.reason}") from e
    return mcp.tool()(wrapper)


def mime_of(file_id: str) -> dict:
    return drive().files().get(fileId=file_id, fields="name,mimeType").execute()


def quote_tab(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"  # A1-notation quoting for tab names


# ---------- read ----------

@tool
def search_files(query: str, max_results: int = 10) -> list[dict]:
    """Search Google Drive by file name. Returns id, name, type and last-modified date."""
    safe = query.replace("\\", "\\\\").replace("'", "\\'")
    q = f"name contains '{safe}' and trashed = false"
    res = drive().files().list(q=q, pageSize=max_results,
                               fields="files(id,name,mimeType,modifiedTime)").execute()
    return res.get("files", [])


@tool
def read_file(file_id: str) -> str:
    """Read a Drive file as text. Docs come back as Markdown, Sheets as one CSV block
    per tab (all tabs), Slides as plain text, Drawings as SVG. Plain-text uploads are
    returned as-is; other binary uploads (PDF, .docx, images...) are not supported."""
    meta = mime_of(file_id)
    mime = meta["mimeType"]
    if mime == SHEET:
        return read_sheet(file_id)
    if mime in EXPORT_AS:
        return drive().files().export(fileId=file_id, mimeType=EXPORT_AS[mime]).execute().decode()
    if mime.startswith("application/vnd.google-apps."):
        raise ToolError(f"'{meta['name']}' is a {mime}, which Google cannot export as text")
    if not (mime.startswith("text/") or mime in ("application/json", "application/xml")):
        raise ToolError(f"'{meta['name']}' is a binary {mime} file; it cannot be read as text")
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, drive().files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue().decode(errors="replace")


def read_sheet(file_id: str) -> str:
    sheets = service("sheets", "v4").spreadsheets()
    titles = [s["properties"]["title"] for s in sheets.get(
        spreadsheetId=file_id, fields="sheets.properties.title").execute()["sheets"]]
    ranges = sheets.values().batchGet(
        spreadsheetId=file_id, ranges=[quote_tab(t) for t in titles]).execute()["valueRanges"]
    out = []
    for title, vr in zip(titles, ranges):
        buf = io.StringIO()
        csv.writer(buf).writerows(vr.get("values", []))
        out.append(f"## Tab: {title}\n{buf.getvalue()}")
    return "\n".join(out)


# ---------- create ----------

@tool
def create_file(name: str, kind: Literal["doc", "sheet", "slides"], content: str = "",
                folder_id: str | None = None, template_id: str | None = None) -> dict:
    """Create a native Google file. `content` is Markdown for a doc and CSV for a sheet
    (first tab); for slides it must be empty (add slides afterwards with add_slide).
    With template_id, the new file is a copy of that file instead (e.g. a themed deck),
    and content must be empty. Goes in My Drive root unless folder_id is given.
    Returns id, name and link."""
    if (kind == "slides" or template_id) and content:
        raise ToolError("content is only for new docs and sheets, not slides or template copies")
    body = {"name": name}
    if folder_id:
        body["parents"] = [folder_id]
    if template_id:
        if mime_of(template_id)["mimeType"] != KIND_TO_MIME[kind]:
            raise ToolError(f"template_id is not a {kind}")
        return drive().files().copy(fileId=template_id, body=body,
                                    fields="id,name,webViewLink").execute()
    body["mimeType"] = KIND_TO_MIME[kind]
    media = None
    if content:
        src = "text/markdown" if kind == "doc" else "text/csv"
        media = MediaIoBaseUpload(io.BytesIO(content.encode()), mimetype=src)
    return drive().files().create(body=body, media_body=media,
                                  fields="id,name,webViewLink").execute()


# ---------- edit: sheets ----------

@tool
def update_cells(file_id: str, range: str, values: list[list[str]]) -> dict:
    """Overwrite cells of a Google Sheet. `range` is A1 notation, e.g. "Sheet1!I5" or
    "'My tab'!A2:C3"; `values` is a list of rows. Input is parsed as if typed by a user
    (numbers, dates, formulas starting with = are interpreted)."""
    res = service("sheets", "v4").spreadsheets().values().update(
        spreadsheetId=file_id, range=range, valueInputOption="USER_ENTERED",
        body={"values": values}).execute()
    return {"updatedRange": res.get("updatedRange"), "updatedCells": res.get("updatedCells", 0)}


@tool
def append_rows(file_id: str, tab: str, rows: list[list[str]]) -> dict:
    """Append rows after the last non-empty row of a Google Sheet tab."""
    res = service("sheets", "v4").spreadsheets().values().append(
        spreadsheetId=file_id, range=quote_tab(tab), valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS", body={"values": rows}).execute()
    upd = res.get("updates", {})
    return {"updatedRange": upd.get("updatedRange"), "updatedRows": upd.get("updatedRows", 0)}


# ---------- edit: docs ----------

@tool
def insert_text(file_id: str, text: str, at: Literal["end", "start"] = "end") -> dict:
    """Insert plain text into a Google Doc, at the end or the start of its body."""
    loc = ({"endOfSegmentLocation": {}} if at == "end" else {"location": {"index": 1}})
    service("docs", "v1").documents().batchUpdate(
        documentId=file_id, body={"requests": [{"insertText": {"text": text, **loc}}]}).execute()
    return {"inserted_chars": len(text), "at": at}


# ---------- edit: docs + slides ----------

@tool
def replace_text(file_id: str, find: str, replace: str, match_case: bool = True) -> dict:
    """Replace every occurrence of `find` in a Google Doc or Google Slides presentation."""
    req = {"requests": [{"replaceAllText": {
        "containsText": {"text": find, "matchCase": match_case}, "replaceText": replace}}]}
    mime = mime_of(file_id)["mimeType"]
    if mime == DOC:
        res = service("docs", "v1").documents().batchUpdate(documentId=file_id, body=req).execute()
    elif mime == SLIDES:
        res = service("slides", "v1").presentations().batchUpdate(
            presentationId=file_id, body=req).execute()
    else:
        raise ToolError(f"replace_text works on Docs and Slides, not {mime} (for Sheets use update_cells)")
    reply = res.get("replies", [{}])[0].get("replaceAllText", {})
    return {"occurrencesChanged": reply.get("occurrencesChanged", 0)}


# ---------- edit: slides ----------
# Slides text indices count UTF-16 code units, not Python characters.

Layout = Literal["TITLE", "SECTION_HEADER", "SECTION_TITLE_AND_DESCRIPTION", "TITLE_AND_BODY",
                 "TITLE_AND_TWO_COLUMNS", "TITLE_ONLY", "ONE_COLUMN_TEXT", "MAIN_POINT",
                 "BIG_NUMBER", "CAPTION_ONLY", "BLANK"]
EMU_PER_PT = 12700
PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
IMAGE_MIMES = ("image/png", "image/jpeg", "image/gif")
# Image boxes as fractions of the slide (left, top, width, height). The image is
# scaled to fit inside its box, keeping its aspect ratio, and centred in it.
ImageRegion = Literal["full", "left", "right", "slide"]
IMAGE_REGIONS = {
    "full": (0.05, 0.22, 0.90, 0.72),   # under a normal title
    "left": (0.05, 0.22, 0.43, 0.72),   # left column
    "right": (0.52, 0.22, 0.43, 0.72),  # right column
    "slide": (0.02, 0.17, 0.96, 0.81),  # as big as possible under a one-line title
}


def slides_api():
    return service("slides", "v1").presentations()


def slides_batch(file_id: str, reqs: list[dict]) -> dict:
    return slides_api().batchUpdate(presentationId=file_id, body={"requests": reqs}).execute()


def u16(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2


def rgb(hex_color: str) -> dict:
    h = hex_color.lstrip("#")
    if len(h) != 6 or any(c not in "0123456789abcdefABCDEF" for c in h):
        raise ToolError(f"Colour must be hex like '#1A2B3C', got {hex_color!r}")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}


def shape_text(el: dict) -> str:
    # textRun and autoText elements tile the shape's text, so their concatenation
    # lines up with the API's indices.
    return "".join((te.get("textRun") or te.get("autoText") or {}).get("content", "")
                   for te in el.get("shape", {}).get("text", {}).get("textElements", []))


def keep_bold(el: dict, start: int = 0, end: int | None = None) -> list[dict]:
    """Requests that re-apply bold to the bold runs of a shape (clipped to [start, end)).
    Needed after a fontFamily change, which the Slides API makes reset bold to off."""
    reqs = []
    for te in el.get("shape", {}).get("text", {}).get("textElements", []):
        if not te.get("textRun", {}).get("style", {}).get("bold"):
            continue
        s, e = max(te.get("startIndex", 0), start), te["endIndex"] if end is None else min(te["endIndex"], end)
        if s < e:
            reqs.append({"updateTextStyle": {"objectId": el["objectId"], "style": {"bold": True},
                                             "fields": "bold", "textRange": {
                                                 "type": "FIXED_RANGE", "startIndex": s, "endIndex": e}}})
    return reqs


def text_requests(object_id: str, text: str, bullets: bool = False) -> list[dict]:
    """Insert `text` into a shape; **x** spans become bold, and with bullets=True each
    line becomes a bullet whose nesting level is its number of leading tabs."""
    parts = text.split("**")
    if len(parts) % 2 == 0:  # unbalanced ** -> keep literally
        parts = [text]
    plain, pos, bold = "", 0, []
    for i, part in enumerate(parts):
        if i % 2 and part:
            bold.append((pos, pos + u16(part)))
        plain += part
        pos += u16(part)
    reqs = [{"insertText": {"objectId": object_id, "text": plain}}]
    reqs += [{"updateTextStyle": {"objectId": object_id, "style": {"bold": True}, "fields": "bold",
                                  "textRange": {"type": "FIXED_RANGE", "startIndex": s, "endIndex": e}}}
             for s, e in bold]
    if bullets:  # applied last: it strips the leading tabs, which shifts indices
        reqs.append({"createParagraphBullets": {"objectId": object_id, "textRange": {"type": "ALL"},
                                                "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE"}})
    return reqs


def notes_requests(page: dict, notes: str) -> list[dict]:
    notes_page = page["slideProperties"]["notesPage"]
    nid = notes_page["notesProperties"]["speakerNotesObjectId"]
    existing = next((el for el in notes_page.get("pageElements", []) if el["objectId"] == nid), {})
    reqs = [{"deleteText": {"objectId": nid, "textRange": {"type": "ALL"}}}] if shape_text(existing) else []
    return reqs + (text_requests(nid, notes) if notes else [])


def get_page(file_id: str, page_id: str) -> dict:
    return slides_api().pages().get(presentationId=file_id, pageObjectId=page_id).execute()


@tool
def list_slides(file_id: str) -> dict:
    """List a presentation's slides in order with their ids, layout, text boxes (id,
    placeholder type, text), images and speaker notes. Use the ids with the other
    slide tools. Slide size is in points."""
    p = slides_api().get(presentationId=file_id, fields=(
        "pageSize,layouts(objectId,layoutProperties(name)),"
        "slides(objectId,pageElements,slideProperties(layoutObjectId,notesPage))")).execute()
    layouts = {l["objectId"]: l["layoutProperties"].get("name") for l in p.get("layouts", [])}
    out = []
    for i, s in enumerate(p.get("slides", [])):
        texts, images = [], []
        for el in s.get("pageElements", []):
            if "image" in el:
                images.append(el["objectId"])
            elif "shape" in el:
                texts.append({"id": el["objectId"],
                              "placeholder": el["shape"].get("placeholder", {}).get("type"),
                              "text": shape_text(el).strip()})
        np_ = s["slideProperties"]["notesPage"]
        nid = np_["notesProperties"]["speakerNotesObjectId"]
        notes = next((shape_text(el) for el in np_.get("pageElements", []) if el["objectId"] == nid), "")
        out.append({"index": i, "slideId": s["objectId"],
                    "layout": layouts.get(s["slideProperties"]["layoutObjectId"]),
                    "texts": texts, "images": images, "notes": notes.strip()})
    size = p["pageSize"]
    return {"widthPt": size["width"]["magnitude"] / EMU_PER_PT,
            "heightPt": size["height"]["magnitude"] / EMU_PER_PT, "slides": out}


def fill_requests(page: dict, layout: str, title: str, body: str, right: str,
                  notes: str, bullets: bool) -> list[dict]:
    """Requests that write title/body/right/notes into a slide's placeholders.
    `body` falls back to the subtitle box on layouts without a body (TITLE...)."""
    ph = {}
    for el in page.get("pageElements", []):
        p = el.get("shape", {}).get("placeholder")
        if p:
            ph.setdefault(p["type"], []).append(
                (p.get("index", 0), el["transform"].get("translateX", 0), el["objectId"]))
    title_ids = ph.get("TITLE") or ph.get("CENTERED_TITLE") or []
    body_ids = [oid for *_, oid in sorted(ph.get("BODY", []))]
    sub_ids = [oid for *_, oid in sorted(ph.get("SUBTITLE", []))]
    targets = [("title", title, title_ids[0][2] if title_ids else None, False),
               ("body", body, (body_ids or sub_ids or [None])[0], bullets and bool(body_ids)),
               ("right", right, body_ids[1] if len(body_ids) > 1 else None, bullets)]
    reqs = []
    for name, txt, oid, bul in targets:
        if not txt:
            continue
        if oid is None:
            raise ToolError(f"Layout {layout} has no box for `{name}` (it has: {sorted(ph)})")
        reqs += text_requests(oid, txt, bul)
    if notes:
        reqs += notes_requests(page, notes)
    return reqs


@tool
def add_slide(file_id: str, layout: Layout = "TITLE_AND_BODY", title: str = "", body: str = "",
              right: str = "", notes: str = "", index: int | None = None,
              bullets: bool = True) -> dict:
    """Add a slide to a Google Slides presentation (at the end, or at 0-based `index`).
    `body` fills the main text box (the subtitle on TITLE / SECTION_HEADER layouts);
    `right` fills the second column of TITLE_AND_TWO_COLUMNS. In any text, **x** is
    bold. With bullets=True, body/right lines become bullets and leading tabs nest
    them. `notes` sets the speaker notes. Layouts come from the deck's theme.
    To leave room for a figure the user will insert by hand, use TITLE_ONLY."""
    slide = f"s_{uuid.uuid4().hex}"
    create = {"objectId": slide, "slideLayoutReference": {"predefinedLayout": layout}}
    if index is not None:
        create["insertionIndex"] = index
    slides_batch(file_id, [{"createSlide": create}])
    try:
        reqs = fill_requests(get_page(file_id, slide), layout, title, body, right, notes, bullets)
        if reqs:
            slides_batch(file_id, reqs)
    except Exception:
        slides_batch(file_id, [{"deleteObject": {"objectId": slide}}])  # no half-filled slides
        raise
    return {"slideId": slide}


class SlideSpec(BaseModel):
    layout: Layout = "TITLE_AND_BODY"
    title: str = ""
    body: str = ""
    right: str = ""
    notes: str = ""
    bullets: bool = True
    image: str | None = Field(None, description="Local file path or Drive file id of a PNG/JPEG/GIF")
    image_region: ImageRegion = "full"


def image_bytes(source: str) -> bytes:
    """Read an image privately: from disk, or from Drive through the user's login."""
    path = Path(source).expanduser()
    if path.is_file():
        mime = mimetypes.guess_type(path.name)[0]
        if mime not in IMAGE_MIMES:
            raise ToolError(f"{path.name} is {mime}; use PNG, JPEG or GIF")
        return path.read_bytes()
    meta = mime_of(source)  # not a local file: treat as a Drive file id
    if meta["mimeType"] not in IMAGE_MIMES:
        raise ToolError(f"'{meta['name']}' is {meta['mimeType']}; use PNG, JPEG or GIF")
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, drive().files().get_media(fileId=source))
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()


@tool
def create_presentation(name: str, slides: list[SlideSpec], template_id: str | None = None,
                        folder_id: str | None = None) -> dict:
    """Create a Google Slides deck in one go, including figures, without ever making
    any file public. Each slide takes the same fields as add_slide plus an optional
    `image` (local path or Drive file id; PNG/JPEG/GIF) placed under the title in
    `image_region` (full width, or the left/right column, keeping its aspect ratio).
    The look comes from `template_id` (a themed Google Slides deck, whose own slides
    are not copied) or Google's default theme. Figures can only be embedded here, at
    creation: to add one to an existing deck, the user inserts it in the editor."""
    # 1. Template: a Google Slides theme, exported privately to .pptx.
    temp = None
    if template_id:
        if mime_of(template_id)["mimeType"] != SLIDES:
            raise ToolError("template_id must be a Google Slides deck")
        src = template_id
    else:
        temp = src = drive().files().create(body={"name": "tmp-blank-template", "mimeType": SLIDES},
                                            fields="id").execute()["id"]
    try:
        prs = Presentation(io.BytesIO(drive().files().export(fileId=src, mimeType=PPTX).execute()))
    finally:
        if temp:
            drive().files().delete(fileId=temp).execute()

    # 2. Slide structure and figures, built locally.
    for sid in list(prs.slides._sldIdLst):  # keep the theme, drop the template's slides
        prs.part.drop_rel(sid.rId)
        prs.slides._sldIdLst.remove(sid)
    layouts = {l.name: l for l in prs.slide_layouts}
    W, H = prs.slide_width, prs.slide_height
    for i, spec in enumerate(slides):
        if spec.layout not in layouts:
            raise ToolError(f"Slide {i}: the theme has no {spec.layout} layout (it has {sorted(layouts)})")
        slide = prs.slides.add_slide(layouts[spec.layout])
        if spec.image:
            data = image_bytes(spec.image)
            px_w, px_h = PILImage.open(io.BytesIO(data)).size
            x, y, w, h = (f * d for f, d in zip(IMAGE_REGIONS[spec.image_region], (W, H, W, H)))
            scale = min(w / px_w, h / px_h)  # fit inside the box, centred
            slide.shapes.add_picture(io.BytesIO(data), int(x + (w - px_w * scale) / 2),
                                     int(y + (h - px_h * scale) / 2),
                                     int(px_w * scale), int(px_h * scale))
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    # 3. Private upload, converted to Google Slides.
    body = {"name": name, "mimeType": SLIDES}
    if folder_id:
        body["parents"] = [folder_id]
    deck = drive().files().create(body=body, media_body=MediaIoBaseUpload(buf, mimetype=PPTX),
                                  fields="id,name,webViewLink").execute()

    # 4. Text through the Slides API, the same way add_slide writes it.
    try:
        pages = slides_api().get(presentationId=deck["id"],
                                 fields="slides(objectId,pageElements,slideProperties)").execute()["slides"]
        reqs = []
        for spec, page in zip(slides, pages):
            reqs += fill_requests(page, spec.layout, spec.title, spec.body, spec.right,
                                  spec.notes, spec.bullets)
        if reqs:
            slides_batch(deck["id"], reqs)
    except Exception:
        drive().files().delete(fileId=deck["id"]).execute()  # no half-built decks
        raise
    return {**deck, "slideIds": [p["objectId"] for p in pages]}


@tool
def set_notes(file_id: str, slide_id: str, notes: str) -> dict:
    """Replace a slide's speaker notes (empty string clears them). **x** is bold."""
    reqs = notes_requests(get_page(file_id, slide_id), notes)
    if reqs:
        slides_batch(file_id, reqs)
    return {"slideId": slide_id, "notes_chars": len(notes)}


@tool
def add_image(file_id: str, slide_id: str, source: str,
              region: ImageRegion = "full",
              left: float | None = None, top: float | None = None,
              width: float | None = None, height: float | None = None) -> dict:
    """Place a PNG, JPEG or GIF on a slide. `source` must be an already-public http(s)
    URL (private local/Drive images are disabled). The image keeps its aspect ratio
    inside its box: `region` picks the area under the title (full width, or the
    left/right column), or give an explicit box in points (left, top, width, height)."""
    p = slides_api().get(presentationId=file_id, fields="pageSize").execute()["pageSize"]
    W, H = p["width"]["magnitude"] / EMU_PER_PT, p["height"]["magnitude"] / EMU_PER_PT
    default = [f * d for f, d in zip(IMAGE_REGIONS[region], (W, H, W, H))]
    box = [d if v is None else v for d, v in zip(default, (left, top, width, height))]

    # Private images (local files, Drive files) are refused: the Slides API only
    # fetches URLs reachable without login, and we never make the user's files public.
    if not source.startswith(("http://", "https://")):
        raise ToolError("Only already-public http(s) image URLs are supported for now; "
                        "private local/Drive images are disabled until a private upload route exists")
    img = f"i_{uuid.uuid4().hex}"
    slides_batch(file_id, [{"createImage": {"objectId": img, "url": source, "elementProperties": {
        "pageObjectId": slide_id,
        "size": {"width": {"magnitude": box[2], "unit": "PT"},
                 "height": {"magnitude": box[3], "unit": "PT"}},
        "transform": {"scaleX": 1, "scaleY": 1, "translateX": box[0], "translateY": box[1],
                      "unit": "PT"}}}}])
    return {"imageId": img}


@tool
def format_text(file_id: str, find: str, bold: bool | None = None, italic: bool | None = None,
                underline: bool | None = None, color: str | None = None,
                font: str | None = None, size_pt: float | None = None,
                slide_id: str | None = None, match_case: bool = True) -> dict:
    """Style every occurrence of `find` in a presentation's text boxes (or only on
    `slide_id`): bold/italic/underline on or off, hex colour, font family, size."""
    style, fields = {}, []
    for key, val in (("bold", bold), ("italic", italic), ("underline", underline)):
        if val is not None:
            style[key] = val
            fields.append(key)
    if color:
        style["foregroundColor"] = {"opaqueColor": {"rgbColor": rgb(color)}}
        fields.append("foregroundColor")
    if font:
        style["fontFamily"] = font
        fields.append("fontFamily")
    if size_pt:
        style["fontSize"] = {"magnitude": size_pt, "unit": "PT"}
        fields.append("fontSize")
    if not fields:
        raise ToolError("Give at least one style to apply")
    slides = slides_api().get(presentationId=file_id,
                              fields="slides(objectId,pageElements(objectId,shape))").execute()
    needle = find if match_case else find.lower()
    reqs, found = [], 0
    for s in slides.get("slides", []):
        if slide_id and s["objectId"] != slide_id:
            continue
        for el in s.get("pageElements", []):
            text = shape_text(el)
            hay = text if match_case else text.lower()
            i = hay.find(needle)
            while needle and i != -1:
                start = u16(text[:i])
                end = start + u16(text[i:i + len(find)])
                reqs.append({"updateTextStyle": {
                    "objectId": el["objectId"], "style": style, "fields": ",".join(fields),
                    "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end}}})
                found += 1
                if font and bold is None:
                    reqs += keep_bold(el, start, end)
                i = hay.find(needle, i + len(find))
    if reqs:
        slides_batch(file_id, reqs)
    return {"occurrencesStyled": found}


@tool
def style_deck(file_id: str, font: str | None = None, title_font: str | None = None,
               text_color: str | None = None, title_color: str | None = None,
               background: str | None = None, text_size: float | None = None,
               nested_text_size: float | None = None) -> dict:
    """Apply a look to every slide that exists now: font family for all text (or
    title_font for titles only), hex colours for body text, titles and the slide
    background. text_size (pt) sets the size of body text boxes (bullets and
    columns; titles and subtitles keep theirs), with nested bullet lines at
    nested_text_size (default: same as text_size). Run it again after adding
    slides, then use format_text for one-off
    accents (this overrides their colour and font). For a full Google theme, create the
    deck from a themed template instead (create_file with template_id)."""
    slides = slides_api().get(presentationId=file_id,
                              fields="slides(objectId,pageElements(objectId,shape))").execute()
    reqs = []
    for s in slides.get("slides", []):
        if background:
            reqs.append({"updatePageProperties": {
                "objectId": s["objectId"], "fields": "pageBackgroundFill.solidFill.color",
                "pageProperties": {"pageBackgroundFill": {"solidFill": {"color": {"rgbColor": rgb(background)}}}}}})
        for el in s.get("pageElements", []):
            if not shape_text(el):
                continue
            is_title = el["shape"].get("placeholder", {}).get("type") in ("TITLE", "CENTERED_TITLE")
            f = (title_font or font) if is_title else font
            c = (title_color or text_color) if is_title else text_color
            style, fields = {}, []
            if f:
                style["fontFamily"] = f
                fields.append("fontFamily")
            if c:
                style["foregroundColor"] = {"opaqueColor": {"rgbColor": rgb(c)}}
                fields.append("foregroundColor")
            if fields:
                reqs.append({"updateTextStyle": {"objectId": el["objectId"], "style": style,
                                                 "fields": ",".join(fields), "textRange": {"type": "ALL"}}})
                if f:
                    reqs += keep_bold(el)
            if text_size and el["shape"].get("placeholder", {}).get("type") == "BODY":
                reqs += size_requests(el, text_size, nested_text_size or text_size)
    if reqs:
        slides_batch(file_id, reqs)
    return {"slides": len(slides.get("slides", [])), "changes": len(reqs)}


def size_requests(el: dict, top_size: float, nested_size: float) -> list[dict]:
    """Font-size requests per paragraph: top-level lines at top_size, nested
    bullet lines (nesting level >= 1) at nested_size."""
    reqs = []
    for te in el["shape"].get("text", {}).get("textElements", []):
        if "paragraphMarker" not in te:
            continue
        level = te["paragraphMarker"].get("bullet", {}).get("nestingLevel", 0)
        start, end = te.get("startIndex", 0), te["endIndex"]
        if start < end:
            reqs.append({"updateTextStyle": {
                "objectId": el["objectId"], "fields": "fontSize",
                "style": {"fontSize": {"magnitude": nested_size if level else top_size, "unit": "PT"}},
                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end}}})
    return reqs


@tool
def move_slides(file_id: str, slide_ids: list[str], index: int) -> dict:
    """Move one or more slides (kept in the given order) to 0-based position `index`,
    counted in the deck as it is before the move."""
    slides_batch(file_id, [{"updateSlidesPosition": {"slideObjectIds": slide_ids, "insertionIndex": index}}])
    return {"moved": slide_ids, "to": index}


@tool
def delete_slide_or_element(file_id: str, object_id: str) -> dict:
    """Delete a slide, or one element on a slide (image, text box), by the id that
    list_slides shows. Undo is only possible from the deck's version history."""
    slides_batch(file_id, [{"deleteObject": {"objectId": object_id}}])
    return {"deleted": object_id}


if __name__ == "__main__":
    if sys.argv[1:] == ["auth"]:
        login()
    else:
        mcp.run()  # stdio transport by default
