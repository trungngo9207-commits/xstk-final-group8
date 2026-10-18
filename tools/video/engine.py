"""Tiny screencast engine: spreadsheet mock-up + slides + narration -> MP4."""
import os, re, math, wave, hashlib, subprocess
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from fontTools.ttLib import TTFont

HERE = os.path.dirname(os.path.abspath(__file__))
W, H, FPS = 1920, 1080, 30
SR = 22050

# ---------------------------------------------------------------- palette
BG = (243, 242, 238)
BAND = (32, 32, 30)
WHITE = (255, 255, 255)
CARD = (252, 252, 251)
INK = (11, 11, 11)
INK2 = (82, 81, 78)
MUTED = (137, 135, 129)
GRID = (225, 224, 217)
AXIS = (195, 194, 183)
CHROME = (233, 232, 228)
HEAD = (240, 239, 235)
HEAD_ON = (214, 226, 242)
BLUE = (42, 120, 214)
ORANGE = (235, 104, 52)
AQUA = (27, 175, 122)
VIOLET = (98, 80, 214)
GREEN = (0, 131, 0)
MAGENTA = (213, 81, 129)
YELLOW = (237, 161, 0)
RED = (227, 73, 72)
SEL = (28, 92, 171)
REF_COLORS = [BLUE, ORANGE, VIOLET, GREEN, MAGENTA, AQUA]

# ---------------------------------------------------------------- fonts
FONT_FILES = {
    "ui": "/usr/share/fonts/truetype/ubuntu/Ubuntu-R.ttf",
    "ui-b": "/usr/share/fonts/truetype/ubuntu/Ubuntu-B.ttf",
    "ui-i": "/usr/share/fonts/truetype/ubuntu/Ubuntu-RI.ttf",
    "cell": "/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
    "cell-b": "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf",
    "mono": "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    "mono-b": "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
    "fb": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "fb-b": "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
}
_cmap, _fonts = {}, {}
MACRON = "̄"


def cmap(key):
    if key not in _cmap:
        _cmap[key] = set(TTFont(FONT_FILES[key]).getBestCmap().keys())
    return _cmap[key]


def font(key, size):
    k = (key, size)
    if k not in _fonts:
        _fonts[k] = ImageFont.truetype(FONT_FILES[key], size)
    return _fonts[k]


def _runs(text, key):
    fb = "fb-b" if key.endswith("-b") else "fb"
    out = []
    for ch in text:
        k = key if (ord(ch) in cmap(key) and ch != MACRON) else fb
        if ch == MACRON:
            k = "macron"
        if out and out[-1][0] == k and k != "macron":
            out[-1][1] += ch
        else:
            out.append([k, ch])
    return out


def text_w(text, key, size):
    return sum(font(k, size).getlength(t) for k, t in _runs(text, key) if k != "macron")


def draw_text(d, x, y, text, key, size, fill, anchor="ls"):
    """Draw with per-glyph font fallback. y is the baseline unless anchor='lm'/'lt'."""
    f = font(key, size)
    asc, desc = f.getmetrics()
    if anchor == "lm":
        y = y + (asc - desc) / 2
    elif anchor == "lt":
        y = y + asc
    elif anchor == "mm":
        x = x - text_w(text, key, size) / 2
        y = y + (asc - desc) / 2
    elif anchor == "ms":
        x = x - text_w(text, key, size) / 2
    elif anchor == "rs":
        x = x - text_w(text, key, size)
    elif anchor == "rm":
        x = x - text_w(text, key, size)
        y = y + (asc - desc) / 2
    last_w = 0
    for k, t in _runs(text, key):
        if k == "macron":  # draw a bar over the previous glyph
            xb = x - last_w
            d.line([(xb + last_w * 0.12, y - size * 0.62), (x - last_w * 0.08, y - size * 0.62)],
                   fill=fill, width=max(1, size // 14))
            continue
        ff = font(k, size)
        d.text((x, y), t, font=ff, fill=fill, anchor="ls")
        last_w = ff.getlength(t[-1]) if t else 0
        x += ff.getlength(t)
    return x


def wrap(text, key, size, maxw):
    lines = []
    for para in text.split("\n"):
        words, cur = para.split(" "), ""
        for w_ in words:
            trial = (cur + " " + w_).strip()
            if text_w(trial, key, size) <= maxw or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = w_
        lines.append(cur)
    return lines


def rrect(d, box, r, fill=None, outline=None, width=1):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=width)


# ---------------------------------------------------------------- refs & formatting
def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_ref(ref):
    m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", ref)
    return col2n(m.group(1)), int(m.group(2))


def parse_range(rng):
    a, _, b = rng.partition(":")
    c1, r1 = parse_ref(a)
    c2, r2 = parse_ref(b) if b else (c1, r1)
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def cells_in(rng):
    c1, r1, c2, r2 = parse_range(rng)
    return [f"{n2col(c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


REF_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:[A-Za-z]+!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)")


def ref_segments(formula):
    """Split a formula into (text, colour_or_None, ref_or_None) segments, colouring refs like a spreadsheet."""
    segs, pos, colours = [], 0, {}
    for m in REF_RE.finditer(formula):
        if m.start() > pos:
            segs.append((formula[pos:m.start()], None, None))
        tok = m.group(1)
        key = tok.replace("$", "")
        if key not in colours:
            colours[key] = REF_COLORS[len(colours) % len(REF_COLORS)]
        segs.append((tok, colours[key], tok))
        pos = m.end()
    if pos < len(formula):
        segs.append((formula[pos:], None, None))
    return segs


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    v = float(v)
    if v.is_integer() and abs(v) < 1e12:
        return str(int(v))
    a = abs(v)
    if a < 1e-4:
        m, e = f"{v:.2E}".split("E")
        return f"{m}E{int(e)}"
    if a < 1e-3:
        return f"{v:.6f}".rstrip("0")
    return f"{v:.4f}".rstrip("0").rstrip(".")


# ---------------------------------------------------------------- sheet model
class Cell:
    def __init__(self, value="", formula=None, bold=False, num=None, fill=None, color=None):
        self.value, self.formula, self.bold, self.fill, self.color = value, formula, bold, fill, color
        self.num = num if num is not None else bool(re.fullmatch(r"-?[\d.]+(E-?\d+)?", value or ""))


class Sheet:
    RH = 30

    def __init__(self, name, widths, default_w=100, hidden=()):
        self.name, self.widths, self.default_w, self.hidden = name, widths, default_w, set(hidden)
        self.cells = {}
        self.sx = 0.0  # horizontal scroll in px
        self.top = 1.0  # first visible row (float)

    def cw(self, c):
        letter = n2col(c)
        if letter in self.hidden:
            return 0
        return self.widths.get(letter, self.default_w)

    def col_x(self, c):  # absolute x of column start
        return sum(self.cw(i) for i in range(1, c))

    def col_at_px(self, px):
        x, c = 0, 1
        while True:
            w = self.cw(c)
            if x + w > px:
                return c
            x += w
            c += 1


class Workbook:
    def __init__(self, sheets, active, filename):
        self.sheets = {s.name: s for s in sheets}
        self.order = [s.name for s in sheets]
        self.active, self.filename = active, filename
        self.cur = None       # active cell ref
        self.sel = None       # selection range string
        self.callouts = []    # dicts(range, color, label)
        self.badges = []      # dicts(text, x, y)
        self.cursor = None    # (x, y) mouse pointer
        self.cursor_last = None
        self.note = None
        self.version = 0
        self.transient()

    def transient(self):
        self.editing = None   # dict(ref, text)
        self.keycaps = None
        self.ripple = None    # (x, y, phase 0..1)
        self.caret_on = True

    @property
    def sh(self):
        return self.sheets[self.active]


# window geometry (sheet mode)
SX0, SY0, SWW, SWH = 24, 76, 1340, 838
TITLE_H, FBAR_H, CHEAD_H, RHEAD_W, TABS_H = 34, 46, 30, 52, 36
GX0 = SX0 + RHEAD_W
GY0 = SY0 + TITLE_H + FBAR_H + CHEAD_H
GX1 = SX0 + SWW
GY1 = SY0 + SWH - TABS_H
PANEL = (1388, 76, 1896, 914)


def cell_rect(wb, ref, sheet=None):
    sh = sheet or wb.sh
    c, r = parse_ref(ref)
    x = GX0 + sh.col_x(c) - sh.sx
    y = GY0 + (r - sh.top) * Sheet.RH
    return x, y, sh.cw(c), Sheet.RH


def range_rect(wb, rng):
    c1, r1, c2, r2 = parse_range(rng)
    x1, y1, _, _ = cell_rect(wb, f"{n2col(c1)}{r1}")
    x2, y2, w2, h2 = cell_rect(wb, f"{n2col(c2)}{r2}")
    return x1, y1, x2 + w2, y2 + h2


def tab_rects(wb):
    out, tx = {}, SX0 + 70
    for name in wb.order:
        tw = text_w(name, "ui-b", 17) + 34
        out[name] = (tx, GY1, tx + tw, SY0 + SWH)
        tx += tw + 4
    return out


def render_sheet(img, wb):
    d = ImageDraw.Draw(img, "RGBA")
    sh = wb.sh
    # window shadow + frame
    rrect(d, (SX0 + 3, SY0 + 5, SX0 + SWW + 3, SY0 + SWH + 5), 12, fill=(0, 0, 0, 28))
    rrect(d, (SX0, SY0, SX0 + SWW, SY0 + SWH), 12, fill=WHITE, outline=AXIS)
    # title bar
    rrect(d, (SX0, SY0, SX0 + SWW, SY0 + TITLE_H + 10), 12, fill=CHROME)
    d.rectangle((SX0 + 1, SY0 + TITLE_H, SX0 + SWW - 1, SY0 + TITLE_H + 10), fill=WHITE)
    for i, c in enumerate([(205, 204, 198)] * 3):
        d.ellipse((SX0 + 16 + i * 22, SY0 + 11, SX0 + 28 + i * 22, SY0 + 23), fill=c)
    draw_text(d, SX0 + SWW / 2, SY0 + TITLE_H / 2 + 1, f"{wb.filename}  —  sheet: {sh.name}", "ui", 17, INK2, anchor="mm")
    # formula bar
    fy = SY0 + TITLE_H
    d.line((SX0, fy + FBAR_H - 1, SX0 + SWW, fy + FBAR_H - 1), fill=GRID, width=1)
    rrect(d, (SX0 + 10, fy + 8, SX0 + 130, fy + FBAR_H - 8), 4, fill=WHITE, outline=AXIS)
    name = wb.sel if (wb.sel and wb.editing is None) else (wb.editing["ref"] if wb.editing else (wb.cur or ""))
    draw_text(d, SX0 + 20, fy + FBAR_H / 2, name, "cell", 19, INK, anchor="lm")
    draw_text(d, SX0 + 146, fy + FBAR_H / 2, "fx", "ui-i", 20, MUTED, anchor="lm")
    d.line((SX0 + 176, fy + 9, SX0 + 176, fy + FBAR_H - 9), fill=GRID, width=1)
    fx_x = SX0 + 190
    if wb.editing is not None:
        ftxt, colored = wb.editing["text"], True
    else:
        c = sh.cells.get(wb.cur) if wb.cur else None
        ftxt = (c.formula if c and c.formula else (c.value if c else "")) or ""
        colored = False
    x = fx_x
    maxx = SX0 + SWW - 14
    for seg, colr, _ in (ref_segments(ftxt) if ftxt.startswith(("=", "{")) else [(ftxt, None, None)]):
        col = colr if (colored and colr) else INK
        for ch in seg:
            cwid = font("mono", 19).getlength(ch)
            if x + cwid > maxx:
                break
            draw_text(d, x, fy + FBAR_H / 2, ch, "mono", 19, col, anchor="lm")
            x += cwid
    if wb.editing is not None and wb.caret_on:
        d.line((x + 1, fy + 12, x + 1, fy + FBAR_H - 12), fill=INK, width=2)

    # visible columns
    gw = GX1 - GX0
    cols, c = [], sh.col_at_px(sh.sx)
    while True:
        x0 = GX0 + sh.col_x(c) - sh.sx
        if x0 >= GX1:
            break
        if sh.cw(c) > 0:
            cols.append((c, x0, sh.cw(c)))
        c += 1
        if c > 200:
            break
    first_row = max(1, int(math.floor(sh.top)))
    rows = []
    r = first_row
    while True:
        y0 = GY0 + (r - sh.top) * Sheet.RH
        if y0 >= GY1:
            break
        rows.append((r, y0))
        r += 1

    # selection extents for header shading
    sel_rng = wb.sel or wb.cur
    sc1 = sr1 = sc2 = sr2 = None
    if sel_rng:
        sc1, sr1, sc2, sr2 = parse_range(sel_rng)

    # clip region for grid
    grid = Image.new("RGB", (GX1 - GX0, GY1 - GY0), (255, 255, 255))
    g = ImageDraw.Draw(grid, "RGBA")
    ox, oy = GX0, GY0

    # fills
    for (r, y0) in rows:
        for (c, x0, w) in cols:
            cell = sh.cells.get(f"{n2col(c)}{r}")
            if cell and cell.fill:
                g.rectangle((x0 - ox, y0 - oy, x0 - ox + w, y0 - oy + Sheet.RH), fill=cell.fill)
    for co in wb.callouts:
        if co.get("sheet", sh.name) != sh.name:
            continue
        x1, y1, x2, y2 = range_rect(wb, co["range"])
        g.rectangle((x1 - ox, y1 - oy, x2 - ox, y2 - oy), fill=co["color"] + (34,))
    if wb.sel:
        x1, y1, x2, y2 = range_rect(wb, wb.sel)
        g.rectangle((x1 - ox, y1 - oy, x2 - ox, y2 - oy), fill=SEL + (26,))
    edit_refs = []
    if wb.editing is not None:
        for seg, colr, ref in ref_segments(wb.editing["text"]):
            if ref and "!" not in ref:
                edit_refs.append((ref.replace("$", ""), colr))
        for ref, colr in edit_refs:
            x1, y1, x2, y2 = range_rect(wb, ref)
            g.rectangle((x1 - ox, y1 - oy, x2 - ox, y2 - oy), fill=colr + (30,))
    # gridlines
    for (c, x0, w) in cols:
        g.line((x0 - ox + w, 0, x0 - ox + w, GY1 - GY0), fill=GRID)
    for (r, y0) in rows:
        g.line((0, y0 - oy + Sheet.RH, GX1 - GX0, y0 - oy + Sheet.RH), fill=GRID)
    # hidden-column markers
    for (c, x0, w) in cols:
        if c > 1 and sh.cw(c - 1) == 0:
            g.line((x0 - ox - 1, 0, x0 - ox - 1, GY1 - GY0), fill=AXIS, width=3)

    # text
    occupied = lambda c, r: (f"{n2col(c)}{r}" in sh.cells and sh.cells[f"{n2col(c)}{r}"].value != "") or \
        (wb.editing is not None and wb.editing["ref"] == f"{n2col(c)}{r}")
    for (r, y0) in rows:
        base = y0 - oy + Sheet.RH / 2
        for idx, (c, x0, w) in enumerate(cols):
            ref = f"{n2col(c)}{r}"
            if wb.editing is not None and wb.editing["ref"] == ref:
                continue
            cell = sh.cells.get(ref)
            if not cell or cell.value == "":
                continue
            key = "cell-b" if cell.bold else "cell"
            colr = cell.color or INK
            if cell.num:
                draw_text(g, x0 - ox + w - 7, base, cell.value, key, 19, colr, anchor="rm")
            else:
                avail = w - 10
                j = idx + 1
                while j < len(cols) and not occupied(cols[j][0], r):
                    avail += cols[j][2]
                    j += 1
                txt = cell.value
                if j >= len(cols):
                    avail += 4000  # runs past the visible edge; the grid image clips it
                while txt and text_w(txt, key, 19) > avail:
                    txt = txt[:-1]
                tw = text_w(txt, key, 19)
                if tw > w - 10:  # overflowing text hides the gridlines it crosses
                    g.rectangle((x0 - ox + w - 1, y0 - oy + 1, x0 - ox + 6 + tw + 2, y0 - oy + Sheet.RH - 1),
                                fill=cell.fill or (255, 255, 255, 255))
                draw_text(g, x0 - ox + 6, base, txt, key, 19, colr, anchor="lm")
    # in-cell editing
    if wb.editing is not None:
        x0, y0, w, hh = cell_rect(wb, wb.editing["ref"])
        txt = wb.editing["text"]
        tw = max(w, text_w(txt, "cell", 19) + 16)
        g.rectangle((x0 - ox, y0 - oy, x0 - ox + tw, y0 - oy + hh), fill=(255, 255, 255, 255))
        xx = x0 - ox + 6
        for seg, colr, _ in (ref_segments(txt) if txt.startswith("=") else [(txt, None, None)]):
            xx = draw_text(g, xx, y0 - oy + hh / 2, seg, "cell", 19, colr or INK, anchor="lm")
        if wb.caret_on:
            g.line((xx + 1, y0 - oy + 6, xx + 1, y0 - oy + hh - 6), fill=INK, width=2)
        g.rectangle((x0 - ox, y0 - oy, x0 - ox + tw, y0 - oy + hh), outline=SEL, width=3)
    # ref borders while editing
    for ref, colr in edit_refs:
        x1, y1, x2, y2 = range_rect(wb, ref)
        g.rectangle((x1 - ox, y1 - oy, x2 - ox, y2 - oy), outline=colr, width=2)
    # callout borders + labels
    for co in wb.callouts:
        if co.get("sheet", sh.name) != sh.name:
            continue
        x1, y1, x2, y2 = range_rect(wb, co["range"])
        g.rounded_rectangle((x1 - ox - 2, y1 - oy - 2, x2 - ox + 2, y2 - oy + 2), radius=4, outline=co["color"], width=3)
        if co.get("label"):
            lw = text_w(co["label"], "ui-b", 17) + 18
            side = co.get("side", "top")
            if side == "right":
                lx, ly = x2 - ox + 10, (y1 + y2) / 2 - oy - 14
            elif side == "left":
                lx, ly = x1 - ox - lw - 10, (y1 + y2) / 2 - oy - 14
            elif side == "bottom":
                lx, ly = x1 - ox, y2 - oy + 6
            else:
                lx, ly = x1 - ox, y1 - oy - 32
            g.rounded_rectangle((lx, ly, lx + lw, ly + 28), radius=6, fill=co["color"])
            draw_text(g, lx + 9, ly + 14, co["label"], "ui-b", 17, WHITE, anchor="lm")
    # selection & active cell
    if wb.sel and wb.editing is None:
        x1, y1, x2, y2 = range_rect(wb, wb.sel)
        g.rectangle((x1 - ox, y1 - oy, x2 - ox, y2 - oy), outline=SEL, width=2)
    if wb.cur and wb.editing is None:
        x0, y0, w, hh = cell_rect(wb, wb.cur)
        g.rectangle((x0 - ox - 1, y0 - oy - 1, x0 - ox + w + 1, y0 - oy + hh + 1), outline=SEL, width=3)
        if not wb.sel:
            g.rectangle((x0 - ox + w - 3, y0 - oy + hh - 3, x0 - ox + w + 4, y0 - oy + hh + 4), fill=SEL, outline=WHITE)
    if wb.sel and wb.editing is None:
        x1, y1, x2, y2 = range_rect(wb, wb.sel)
        g.rectangle((x2 - ox - 3, y2 - oy - 3, x2 - ox + 4, y2 - oy + 4), fill=SEL, outline=WHITE)
    img.paste(grid, (GX0, GY0))

    # headers: row numbers first, then the column strip on top of them
    d.rectangle((SX0 + 1, GY0, GX0, GY1), fill=HEAD)
    for (r, y0) in rows:
        on = sr1 is not None and sr1 <= r <= sr2
        y1c, y2c = max(y0, GY0), min(y0 + Sheet.RH, GY1)
        if y2c <= y1c:
            continue
        if on:
            d.rectangle((SX0 + 1, y1c, GX0, y2c), fill=HEAD_ON)
        if y0 + Sheet.RH / 2 + 8 < GY1:
            draw_text(d, SX0 + RHEAD_W / 2, y0 + Sheet.RH / 2, str(r), "cell-b" if on else "cell", 17, INK2, anchor="mm")
        if y0 + Sheet.RH < GY1:
            d.line((SX0 + 1, y0 + Sheet.RH, GX0, y0 + Sheet.RH), fill=AXIS)
    d.rectangle((SX0 + 1, GY0 - CHEAD_H, GX1 - 1, GY0), fill=HEAD)
    for (c, x0, w) in cols:
        x1c, x2c = max(x0, GX0), min(x0 + w, GX1)
        if x2c <= x1c:
            continue
        on = sc1 is not None and sc1 <= c <= sc2
        if on:
            d.rectangle((x1c, GY0 - CHEAD_H, x2c, GY0), fill=HEAD_ON)
        if x0 + w < GX1:
            d.line((x0 + w, GY0 - CHEAD_H, x0 + w, GY0), fill=AXIS)
        if c > 1 and sh.cw(c - 1) == 0:
            d.line((x0 - 1, GY0 - CHEAD_H, x0 - 1, GY0), fill=INK2, width=3)
        if x0 >= GX0 - w / 2 and x0 + w / 2 + 10 < GX1:
            draw_text(d, x0 + w / 2, GY0 - CHEAD_H / 2, n2col(c), "cell-b" if on else "cell", 18, INK2, anchor="mm")
    d.rectangle((SX0 + 1, GY0 - CHEAD_H, GX0 - 1, GY0 - 1), fill=HEAD)
    d.line((GX0, GY0 - CHEAD_H, GX0, GY1), fill=AXIS)
    d.line((SX0, GY0, GX1, GY0), fill=AXIS)
    # tabs
    d.rectangle((SX0 + 1, GY1, SX0 + SWW - 1, SY0 + SWH - 10), fill=CHROME)
    rrect(d, (SX0 + 1, SY0 + SWH - 20, SX0 + SWW - 1, SY0 + SWH - 1), 11, fill=CHROME)
    for name, (tx, _, tx2, _) in tab_rects(wb).items():
        tw = tx2 - tx
        on = name == wb.active
        if on:
            d.rectangle((tx, GY1, tx + tw, SY0 + SWH - 4), fill=WHITE)
            d.line((tx + 6, SY0 + SWH - 6, tx + tw - 6, SY0 + SWH - 6), fill=BLUE, width=3)
        draw_text(d, tx + tw / 2, GY1 + TABS_H / 2, name, "ui-b" if on else "ui", 17, INK if on else INK2, anchor="mm")
    d.line((SX0, GY1, SX0 + SWW, GY1), fill=AXIS)

    # badges
    for b in wb.badges:
        bw = text_w(b["text"], "ui-b", 19) + 30
        x, y = b["x"], b["y"]
        rrect(d, (x, y, x + bw, y + 40), 20, fill=b.get("color", INK))
        draw_text(d, x + 15, y + 20, b["text"], "ui-b", 19, WHITE, anchor="lm")
    # keycaps
    if wb.keycaps:
        keys = wb.keycaps
        widths = [max(64, text_w(k, "ui-b", 24) + 36) for k in keys]
        total = sum(widths) + 44 * (len(keys) - 1)
        x = SX0 + SWW / 2 - total / 2
        y = GY1 - 92
        rrect(d, (x - 22, y - 18, x + total + 22, y + 70), 16, fill=(32, 32, 30, 225))
        for i, (k, kw) in enumerate(zip(keys, widths)):
            rrect(d, (x, y + 4, x + kw, y + 54), 9, fill=(170, 169, 162))
            rrect(d, (x, y, x + kw, y + 50), 9, fill=WHITE)
            draw_text(d, x + kw / 2, y + 25, k, "ui-b", 24, INK, anchor="mm")
            x += kw
            if i < len(keys) - 1:
                draw_text(d, x + 22, y + 25, "+", "ui-b", 28, WHITE, anchor="mm")
                x += 44
    # ripple + cursor
    if wb.ripple:
        x, y, p = wb.ripple
        rad = 6 + 22 * p
        d.ellipse((x - rad, y - rad, x + rad, y + rad), outline=BLUE + (int(220 * (1 - p)),), width=3)
    if wb.cursor:
        draw_cursor(d, *wb.cursor)


def draw_cursor(d, x, y):
    pts = [(0, 0), (0, 25), (6, 19.5), (10.5, 29.5), (14.5, 27.8), (10.2, 18), (18, 18)]
    pts = [(x + px, y + py) for px, py in pts]
    d.polygon([(px + 1.5, py + 2) for px, py in pts], fill=(0, 0, 0, 60))
    d.polygon(pts, fill=WHITE, outline=INK)
    d.line(pts + [pts[0]], fill=INK, width=2)


# ---------------------------------------------------------------- note panel
def render_note(img, note):
    if not note:
        return
    d = ImageDraw.Draw(img, "RGBA")
    x0, y0, x1, y1 = PANEL
    rrect(d, (x0 + 3, y0 + 5, x1 + 3, y1 + 5), 14, fill=(0, 0, 0, 22))
    rrect(d, (x0, y0, x1, y1), 14, fill=CARD, outline=AXIS)
    x, y, maxw = x0 + 28, y0 + 26, (x1 - x0) - 56
    for kind, val in note:
        if kind == "eyebrow":
            draw_text(d, x, y + 16, val.upper(), "ui-b", 17, BLUE, anchor="ls")
            y += 34
        elif kind == "title":
            for ln in wrap(val, "ui-b", 30, maxw):
                draw_text(d, x, y + 30, ln, "ui-b", 30, INK)
                y += 38
            y += 10
        elif kind == "text":
            for ln in wrap(val, "ui", 22, maxw):
                draw_text(d, x, y + 22, ln, "ui", 22, INK2)
                y += 31
            y += 12
        elif kind == "strong":
            for ln in wrap(val, "ui-b", 23, maxw):
                draw_text(d, x, y + 23, ln, "ui-b", 23, INK)
                y += 32
            y += 12
        elif kind == "code":
            lines = []
            for part in val.split("\n"):
                cur = ""
                for ch in part:
                    if text_w(cur + ch, "mono", 18) > maxw - 24:
                        cut = cur.rfind(",") + 1
                        if cut <= 0:
                            cut = len(cur)
                        lines.append(cur[:cut])
                        cur = cur[cut:]
                    cur += ch
                lines.append(cur)
            hgt = 18 + 27 * len(lines)
            rrect(d, (x, y, x + maxw, y + hgt), 8, fill=(244, 243, 239), outline=GRID)
            yy = y + 9
            for ln in lines:
                xx = x + 12
                for seg, colr, _ in (ref_segments(ln) if "!" in ln or ln.startswith("=") else [(ln, None, None)]):
                    xx = draw_text(d, xx, yy + 20, seg, "mono", 18, colr or INK)
                yy += 27
            y += hgt + 16
        elif kind == "kv":
            for k, v in val:
                draw_text(d, x, y + 22, k, "ui", 21, INK2)
                draw_text(d, x + maxw, y + 22, v, "ui-b", 21, INK, anchor="rs")
                d.line((x, y + 33, x + maxw, y + 33), fill=GRID)
                y += 38
            y += 12
        elif kind == "img":
            im = load_img(val)
            sc = maxw / im.width
            im2 = im.resize((int(im.width * sc), int(im.height * sc)), Image.LANCZOS)
            img.paste(im2, (int(x), int(y)))
            y += im2.height + 14
        elif kind == "space":
            y += val


_imgs = {}


def load_img(path):
    if path not in _imgs:
        _imgs[path] = Image.open(os.path.join(HERE, path)).convert("RGB")
    return _imgs[path]


# ---------------------------------------------------------------- chrome
def render_chrome(img, chapter, chap_idx, chap_total, progress, subtitle):
    d = ImageDraw.Draw(img, "RGBA")
    d.rectangle((0, 0, W, 60), fill=BAND)
    draw_text(d, 32, 31, chapter, "ui-b", 24, WHITE, anchor="lm")
    draw_text(d, W - 32, 31, f"Group 8 · Part B1 & B2 guide   ·   chapter {chap_idx} of {chap_total}", "ui", 19,
              (195, 194, 183), anchor="rm")
    d.rectangle((0, 57, int(W * progress), 60), fill=BLUE)
    d.rectangle((0, 930, W, H), fill=BAND)
    if subtitle:
        size = 33
        lines = wrap(subtitle, "ui", size, 1800)
        if len(lines) > 2:
            size = 29
            lines = wrap(subtitle, "ui", size, 1820)
        lh = size * 1.32
        y = 1005 - (len(lines) - 1) * lh / 2
        for ln in lines:
            draw_text(d, W / 2, y, ln, "ui", size, WHITE, anchor="mm")
            y += lh


# ---------------------------------------------------------------- TTS
_voice = None


def tts(text, length_scale=1.07):
    global _voice
    cache = os.path.join(HERE, "tts_cache")
    os.makedirs(cache, exist_ok=True)
    key = hashlib.md5(f"{length_scale}|{text}".encode()).hexdigest()
    path = os.path.join(cache, key + ".wav")
    if not os.path.exists(path):
        if _voice is None:
            from piper import PiperVoice
            _voice = PiperVoice.load(os.path.join(HERE, "voices", "en_US-lessac-medium.onnx"))
        from piper import SynthesisConfig
        with wave.open(path + ".tmp", "wb") as wf:
            _voice.synthesize_wav(text, wf, syn_config=SynthesisConfig(length_scale=length_scale))
        os.replace(path + ".tmp", path)
    with wave.open(path, "rb") as wf:
        assert wf.getframerate() == SR
        data = np.frombuffer(wf.readframes(wf.getnframes()), dtype=np.int16)
    return data


SPEECH_SUBS = [
    (r"\$(\d[\d,]*(?:\.\d+)?)", r"\1 dollars"),
    (r"H₀", "H zero"), (r"H₁", "H one"), (r"β₀", "beta zero"), (r"β₁", "beta one"), (r"β̂₀", "beta zero hat"),
    (r"x̄", "x bar"), (r"\bH0\b", "H zero"), (r"μ", "mu"), (r"α", "alpha"), (r"≠", "is not equal to"), (r"×", "times"),
    (r"T\.INV\.2T", "T dot INV dot 2 T"), (r"T\.DIST\.2T", "T dot DIST dot 2 T"), (r"STDEV\.S", "STDEV dot S"),
    (r"VAR\.S", "VAR dot S"), (r"CONFIDENCE\.T", "CONFIDENCE dot T"), (r"T\.TEST", "T dot TEST"),
    (r"E−(\d+)", r"E minus \1"), (r"E-(\d+)", r"E minus \1"),
    (r"([A-Z]+\d+):([A-Z]+\d+)", r"\1 to \2"), (r"Ctrl\+Shift\+Enter", "Control, Shift, Enter"),
    (r"−", " minus "), (r"s²", "s squared"), (r"R²", "R squared"), (r"n₁", "n one"), (r"n₂", "n two"),
    (r"\bvs\.?\b", "versus"), (r"≈", "about"), (r"→", ","), (r"…", "..."),
]


def to_speech(s):
    for a, b in SPEECH_SUBS:
        s = re.sub(a, b, s)
    return re.sub(r"\s+", " ", s).strip()


# ---------------------------------------------------------------- actions
def ease(p):
    p = min(1.0, max(0.0, p))
    return p * p * (3 - 2 * p)


class Action:
    def duration(self, wb, ctx):
        return 0.0

    def start(self, wb):
        pass

    def frame(self, wb, t):  # returns a hashable key describing the visual state
        return None

    def commit(self, wb):
        pass


class Wait(Action):
    def __init__(self, s):
        self.s = s

    def duration(self, wb, ctx):
        return self.s


class Sync(Action):
    """Wait until sentence k of this beat starts."""
    def __init__(self, k, offset=0.0):
        self.k, self.offset = k, offset

    def duration(self, wb, ctx):
        return max(0.0, ctx["sent_starts"][self.k] + self.offset - ctx["now"])


class MoveTo(Action):
    def __init__(self, target, dur=0.45, click=True):
        self.target, self.dur, self.click = target, dur, click

    def resolve(self, wb):
        if isinstance(self.target, tuple):
            return self.target
        if self.target.startswith("tab:"):
            x1, y1, x2, y2 = tab_rects(wb)[self.target[4:]]
            return ((x1 + x2) / 2, (y1 + y2) / 2)
        x, y, w, h = cell_rect(wb, self.target)
        return (x + min(w * 0.5, 60), y + h * 0.55)

    def duration(self, wb, ctx):
        return self.dur + (0.18 if self.click else 0)

    def start(self, wb):
        self.p0 = wb.cursor or wb.cursor_last or (GX0 + 400, GY0 + 300)
        self.p1 = self.resolve(wb)

    def frame(self, wb, t):
        p = ease(t / self.dur)
        x = self.p0[0] + (self.p1[0] - self.p0[0]) * p
        y = self.p0[1] + (self.p1[1] - self.p0[1]) * p
        wb.cursor = (x, y)
        if self.click and t > self.dur:
            wb.ripple = (self.p1[0], self.p1[1], (t - self.dur) / 0.18)
        return ("mv", round(x), round(y), wb.ripple and round(wb.ripple[2], 2))

    def commit(self, wb):
        wb.cursor = self.p1


class Seq(Action):
    """Run sub-actions back to back as one action."""
    def __init__(self, *acts):
        self.acts = list(acts)

    def duration(self, wb, ctx):
        return sum(a.duration(wb, ctx) for a in self.acts)

    def start(self, wb):
        self.i, self.t0 = 0, 0.0
        self._durs = [a.duration(wb, {"sent_starts": [], "now": 0}) for a in self.acts]
        self.acts[0].start(wb)

    def frame(self, wb, t):
        while self.i < len(self.acts) - 1 and t >= self.t0 + self._durs[self.i]:
            self.acts[self.i].commit(wb)
            self.t0 += self._durs[self.i]
            self.i += 1
            self.acts[self.i].start(wb)
        return (self.i, self.acts[self.i].frame(wb, t - self.t0))

    def commit(self, wb):
        for a in self.acts[self.i:]:
            if a is not self.acts[self.i]:
                a.start(wb)
            a.commit(wb)


class Click(Action):
    """Move the mouse to a cell and select it."""
    def __init__(self, ref, sel=None):
        self.ref, self.selr = ref, sel
        self.mv = MoveTo(ref)

    def duration(self, wb, ctx):
        return self.mv.duration(wb, ctx) + 0.1

    def start(self, wb):
        self.mv.start(wb)

    def frame(self, wb, t):
        k = self.mv.frame(wb, min(t, 0.63))
        if t > 0.45:
            wb.cur, wb.sel = self.ref, self.selr
        return ("click", k, t > 0.45)

    def commit(self, wb):
        self.mv.commit(wb)
        wb.cur, wb.sel = self.ref, self.selr
        wb.version += 1


class TypeCell(Action):
    """Click a cell, type its content, press a key, show the result."""
    def __init__(self, ref, text, value, keys=("Enter",), cps=None, num=None, bold=False, move=True, hold=0.25):
        self.ref, self.text, self.value, self.keys = ref, text, value, keys
        self.cps = cps or (15 if len(text) <= 14 else 20 if len(text) <= 32 else 30)
        self.num, self.bold, self.move, self.hold = num, bold, move, hold
        self.mv = MoveTo(ref)

    def tparts(self):
        mv = 0.63 if self.move else 0.0
        typ = min(4.5, len(self.text) / self.cps)
        return mv, typ

    def duration(self, wb, ctx):
        mv, typ = self.tparts()
        return mv + typ + self.hold + (0.75 if self.keys else 0.1)

    def start(self, wb):
        if self.move:
            self.mv.start(wb)

    def frame(self, wb, t):
        mv, typ = self.tparts()
        if t < mv:
            k = self.mv.frame(wb, t)
            if t > 0.45:
                wb.cur, wb.sel = self.ref, None
            return ("tmv", k)
        if self.move:
            wb.cursor_last = self.mv.p1
        wb.cursor = None  # the pointer hides while typing
        wb.cur, wb.sel = self.ref, None
        tt = t - mv
        if tt < typ + self.hold:
            n = min(len(self.text), int(len(self.text) * tt / typ) + 1) if typ > 0 else len(self.text)
            wb.editing = {"ref": self.ref, "text": self.text[:n]}
            wb.caret_on = (int(tt * 2.2) % 2 == 0) or n < len(self.text)
            return ("typ", n, wb.caret_on)
        self._put(wb)
        if self.keys:
            wb.keycaps = list(self.keys)
        return ("done", bool(self.keys))

    def _put(self, wb):
        is_formula = self.text.startswith("=")
        wb.sh.cells[self.ref] = Cell(self.value, self.text if is_formula else None, bold=self.bold, num=self.num)

    def commit(self, wb):
        if self.move:
            self.mv.commit(wb)
            wb.cursor_last = self.mv.p1
        wb.cursor = None
        self._put(wb)
        wb.cur, wb.sel = self.ref, None
        wb.version += 1


class ArrayFormula(Action):
    """Select a block, type one formula, confirm with Ctrl+Shift+Enter, block fills."""
    def __init__(self, rng, text, values, cps=30):
        self.rng, self.text, self.values, self.cps = rng, text, values, cps
        c1, r1, c2, r2 = parse_range(rng)
        self.first, self.last = f"{n2col(c1)}{r1}", f"{n2col(c2)}{r2}"
        self.mv = MoveTo(self.first, click=False)

    def parts(self):
        return 0.5, 0.7, min(4.5, len(self.text) / self.cps), 0.3, 1.1

    def duration(self, wb, ctx):
        return sum(self.parts())

    def start(self, wb):
        self.mv.start(wb)
        x, y, w, h = cell_rect(wb, self.last)
        self.p_end = (x + w * 0.6, y + h * 0.55)

    def frame(self, wb, t):
        a, b, c, dd, e = self.parts()
        if t < a:
            return ("amv", self.mv.frame(wb, t))
        if t < a + b:
            p = ease((t - a) / b)
            x0, y0 = self.mv.p1
            wb.cursor = (x0 + (self.p_end[0] - x0) * p, y0 + (self.p_end[1] - y0) * p)
            c1, r1, c2, r2 = parse_range(self.rng)
            cc = c1 + round((c2 - c1) * p)
            rr = r1 + round((r2 - r1) * p)
            wb.cur, wb.sel = self.first, f"{self.first}:{n2col(cc)}{rr}"
            return ("drag", round(wb.cursor[0]), round(wb.cursor[1]), wb.sel)
        wb.cursor, wb.cursor_last = None, self.p_end
        wb.cur, wb.sel = self.first, self.rng
        tt = t - a - b
        if tt < c + dd:
            n = min(len(self.text), int(len(self.text) * tt / c) + 1)
            wb.editing = {"ref": self.first, "text": self.text[:n]}
            wb.caret_on = (int(tt * 2.2) % 2 == 0) or n < len(self.text)
            return ("atyp", n, wb.caret_on)
        self._put(wb)
        wb.keycaps = ["Ctrl", "Shift", "Enter"]
        return ("adone",)

    def _put(self, wb):
        for ref, v in zip(cells_in(self.rng), self.values):
            wb.sh.cells[ref] = Cell(v, "{" + self.text + "}")

    def commit(self, wb):
        self._put(wb)
        wb.cursor, wb.cursor_last = None, self.p_end
        wb.cur, wb.sel = self.first, self.rng
        wb.version += 1


class Pop(Action):
    """Cells appear one after another (labels, repeated formulas)."""
    def __init__(self, items, gap=0.12, select=False):
        self.items, self.gap, self.select = items, gap, select  # items: (ref, Cell)

    def duration(self, wb, ctx):
        return self.gap * len(self.items) + 0.1

    def frame(self, wb, t):
        n = min(len(self.items), int(t / self.gap) + 1)
        for ref, cell in self.items[:n]:
            wb.sh.cells[ref] = cell
        if self.select:
            wb.cur, wb.sel = self.items[n - 1][0], None
        return ("pop", n)

    def commit(self, wb):
        for ref, cell in self.items:
            wb.sh.cells[ref] = cell
        if self.select:
            wb.cur, wb.sel = self.items[-1][0], None
        wb.version += 1


class Scroll(Action):
    def __init__(self, top=None, sx=None, dur=0.9):
        self.top, self.sx, self.dur = top, sx, dur

    def duration(self, wb, ctx):
        return self.dur

    def start(self, wb):
        sh = wb.sh
        self.t0, self.x0 = sh.top, sh.sx
        self.t1 = self.top if self.top is not None else sh.top
        if isinstance(self.sx, str):
            self.x1 = sh.col_x(col2n(self.sx))
        else:
            self.x1 = self.sx if self.sx is not None else sh.sx
        self.cur0 = wb.cursor

    def frame(self, wb, t):
        p = ease(t / self.dur)
        wb.sh.top = self.t0 + (self.t1 - self.t0) * p
        wb.sh.sx = self.x0 + (self.x1 - self.x0) * p
        return ("scr", round(wb.sh.top, 3), round(wb.sh.sx))

    def commit(self, wb):
        wb.sh.top, wb.sh.sx = self.t1, self.x1
        wb.version += 1


class GotoSheet(Action):
    def __init__(self, name):
        self.name = name
        self.mv = MoveTo("tab:" + name)

    def duration(self, wb, ctx):
        return 0.75

    def start(self, wb):
        self.mv.start(wb)

    def frame(self, wb, t):
        k = self.mv.frame(wb, min(t, 0.63))
        if t > 0.45:
            wb.active = self.name
        return ("goto", k, wb.active)

    def commit(self, wb):
        self.mv.commit(wb)
        wb.active = self.name
        wb.cur, wb.sel = None, None
        wb.version += 1


class FillDown(Action):
    """Double-click the fill handle of a 1-row selection; values cascade down."""
    def __init__(self, rng, cells_by_row, last_row):
        self.rng, self.cells_by_row, self.last_row = rng, cells_by_row, last_row
        c1, r1, c2, r2 = parse_range(rng)
        self.c1, self.c2, self.r = c1, c2, r1

    def duration(self, wb, ctx):
        return 0.5 + 0.35 + 1.0

    def start(self, wb):
        x1, y1, x2, y2 = range_rect(wb, self.rng)
        self.mv = MoveTo((x2 + 1, y2 + 1), dur=0.5, click=False)
        self.mv.start(wb)
        self.vis_rows = [r for r in range(self.r + 1, self.last_row + 1)
                         if GY0 + (r - wb.sh.top) * Sheet.RH < GY1]

    def frame(self, wb, t):
        wb.sel = self.rng
        if t < 0.5:
            return ("fmv", self.mv.frame(wb, t))
        wb.cursor = self.mv.p1
        if t < 0.85:
            ph = ((t - 0.5) % 0.17) / 0.17
            wb.ripple = (self.mv.p1[0], self.mv.p1[1], ph)
            return ("dbl", round(ph, 2))
        p = (t - 0.85) / 0.8
        n = min(len(self.vis_rows), int(len(self.vis_rows) * p) + 1)
        for r in self.vis_rows[:n]:
            for ref, cell in self.cells_by_row(r):
                wb.sh.cells[ref] = cell
        last = self.vis_rows[n - 1]
        wb.sel = f"{n2col(self.c1)}{self.r}:{n2col(self.c2)}{last}"
        return ("fill", n)

    def commit(self, wb):
        for r in range(self.r + 1, self.last_row + 1):
            for ref, cell in self.cells_by_row(r):
                wb.sh.cells[ref] = cell
        wb.cursor = self.mv.p1
        wb.sel = f"{n2col(self.c1)}{self.r}:{n2col(self.c2)}{self.last_row}"
        wb.version += 1


class Do(Action):
    """Instant state change: fn(wb)."""
    def __init__(self, fn):
        self.fn = fn

    def frame(self, wb, t):
        return None

    def commit(self, wb):
        self.fn(wb)
        wb.version += 1


def callout(rng, color, label=None, side="top", sheet=None):
    def f(wb):
        co = {"range": rng, "color": color, "label": label, "side": side}
        if sheet:
            co["sheet"] = sheet
        wb.callouts.append(co)
    return Do(f)


def clear_callouts():
    return Do(lambda wb: wb.callouts.clear())


def badge(text, x, y, color=INK):
    return Do(lambda wb: wb.badges.append({"text": text, "x": x, "y": y, "color": color}))


def clear_badges():
    return Do(lambda wb: wb.badges.clear())


def set_note(note):
    def f(wb):
        wb.note = note
    return Do(f)


def hide_cursor():
    return Do(lambda wb: setattr(wb, "cursor", None))


# ---------------------------------------------------------------- timeline
class Beat:
    def __init__(self, kind, say, actions=(), slide=None, note=None, chapter=None):
        self.kind, self.say, self.actions, self.slide, self.note, self.chapter = kind, say, list(actions), slide, note, chapter


class Video:
    def __init__(self, wb):
        self.wb = wb
        self.beats = []
        self.chapters = []
        self._chapter = ""

    def chapter(self, name):
        self._chapter = name
        self.chapters.append(name)

    def slide(self, fn, say, note=None):
        self.beats.append(Beat("slide", say, slide=fn, chapter=self._chapter))

    def sheet(self, say, acts=(), note=None):
        self.beats.append(Beat("sheet", say, acts, note=note, chapter=self._chapter))

    # -- compile timing
    def compile(self, lead=0.35, gap=0.32, tail=0.5):
        t = 0.0
        self.audio_parts = []
        sim = self.wb  # actions need geometry at start; we time with static durations
        for b in self.beats:
            b.start = t
            b.sents = []
            st = t + lead
            for s in b.say:
                disp, sp = (s if isinstance(s, tuple) else (s, to_speech(s)))
                audio = tts(sp)
                dur = len(audio) / SR
                b.sents.append((st, st + dur, disp))
                self.audio_parts.append((st, audio))
                st += dur + gap
            speech_end = st - gap if b.say else t
            b.rel_sent_starts = [s0 - t for s0, _, _ in b.sents]
            b.speech_end = speech_end
            t = max(speech_end, t) + tail  # provisional; finalised during render (action durations)
            b.end = t
        return self

    def render(self, out_path, preview_times=None, preview_dir=None):
        return self._stream(out_path, preview_times, preview_dir)

    def plan(self):
        """Dry-run every action on a copy of the workbook to fix exact start/end times."""
        if getattr(self, "_planned", False):
            return
        import copy
        wb = self.wb
        sim = copy.deepcopy(wb)
        t = 0.0
        for b in self.beats:
            shift = t - b.start
            b.start = t
            b.sents = [(s0 + shift, s1 + shift, disp) for s0, s1, disp in b.sents]
            b.speech_end += shift
            now = t + 0.2
            b.sched = []
            if b.note is not None:
                sim.note = b.note
            for a in b.actions:
                ctx = {"sent_starts": [s0 for s0, _, _ in b.sents], "now": now}
                dur = a.duration(sim, ctx)
                a.start(sim)
                a.commit(sim)
                b.sched.append((now, now + dur, a))
                now += dur
            b.end = max(b.speech_end, now) + 0.5
            t = b.end
        self.total = t
        self._planned = True

    def _stream(self, out_path, preview_times, preview_dir):
        import imageio_ffmpeg
        wb = self.wb
        # First pass: dry-run actions to get exact durations (needs geometry, so we simulate state)
        self.plan()
        # audio
        audio = np.zeros(int(math.ceil(self.total * SR)) + SR, dtype=np.int32)
        k = 0  # audio_parts are in sentence order
        for b in self.beats:
            for (s0, s1, disp) in b.sents:
                _, a = self.audio_parts[k]
                i = int(round(s0 * SR))
                audio[i:i + len(a)] += a
                k += 1
        audio = np.clip(audio, -32768, 32767).astype(np.int16)
        wav_path = out_path + ".wav"
        with wave.open(wav_path, "wb") as wf:
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(SR)
            wf.writeframes(audio[: int(self.total * SR)].tobytes())
        # chapters
        chap_starts = []
        seen = set()
        for b in self.beats:
            if b.chapter not in seen:
                seen.add(b.chapter)
                chap_starts.append((b.start, b.chapter))
        self.chap_starts = chap_starts
        meta_path = out_path + ".meta.txt"
        with open(meta_path, "w") as f:
            f.write(";FFMETADATA1\ntitle=Final Project Group 8 - B1 & B2 step by step\n")
            for i, (s, name) in enumerate(chap_starts):
                e = chap_starts[i + 1][0] if i + 1 < len(chap_starts) else self.total
                f.write(f"[CHAPTER]\nTIMEBASE=1/1000\nSTART={int(s * 1000)}\nEND={int(e * 1000)}\ntitle={name}\n")
        chap_names = [c for _, c in chap_starts]

        nframes = int(math.ceil(self.total * FPS))
        proc = None
        if preview_times is None:
            ff = imageio_ffmpeg.get_ffmpeg_exe()
            cmd = [ff, "-y", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}",
                   "-r", str(FPS), "-i", "-", "-i", wav_path, "-i", meta_path, "-map", "0:v", "-map", "1:a",
                   "-map_metadata", "2", "-map_chapters", "2", "-c:v", "libx264", "-preset", "fast", "-crf", "18",
                   "-tune", "stillimage", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "160k", "-ar", "44100",
                   "-movflags", "+faststart", "-shortest", out_path]
            proc = subprocess.Popen(cmd, stdin=subprocess.PIPE)
            frames = range(nframes)
        else:
            frames = sorted({int(tt * FPS) for tt in preview_times})
            os.makedirs(preview_dir, exist_ok=True)

        bi, ai = 0, 0
        cur_beat = None
        last_key, last_bytes = None, None
        fade_from, fade_until = None, -1
        prev_kind, prev_img = None, None
        renders = 0
        wanted = set(frames) if preview_times is not None else None
        for f in range(nframes if preview_times is None else max(frames) + 1):
            t = f / FPS
            while bi < len(self.beats) - 1 and t >= self.beats[bi].end:
                bi += 1
            b = self.beats[bi]
            if b is not cur_beat:
                # finish previous beat's actions
                if cur_beat is not None:
                    for (s0, s1, a) in cur_beat.sched[ai:]:
                        if not getattr(a, "_started", False):
                            a.start(wb)
                        a.commit(wb)
                    if cur_beat.kind != b.kind and prev_img is not None:
                        fade_from, fade_until = prev_img, b.start + 0.35
                cur_beat, ai = b, 0
                if b.note is not None:
                    wb.note = b.note
                for _, _, a in b.sched:
                    a._started = False
            # commit finished actions
            while ai < len(b.sched) and t >= b.sched[ai][1]:
                a = b.sched[ai][2]
                if not a._started:
                    a.start(wb)
                    a._started = True
                a.commit(wb)
                ai += 1
            wb.transient()
            akey = None
            if ai < len(b.sched) and t >= b.sched[ai][0]:
                s0, s1, a = b.sched[ai]
                if not a._started:
                    a.start(wb)
                    a._started = True
                akey = (ai, a.frame(wb, t - s0))
            sub = ""
            for s0, s1, disp in b.sents:
                if s0 - 0.05 <= t <= s1 + 0.6:
                    sub = disp
            ci = chap_names.index(b.chapter) + 1
            prog = round(W * t / self.total)
            fade_p = None
            if fade_from is not None and t < fade_until:
                fade_p = round((fade_until - t) / 0.35, 2)
            key = (bi, ai, akey, wb.version, sub, prog, id(wb.note), fade_p, b.kind)
            if wanted is not None and f not in wanted:
                continue
            if key != last_key:
                img = Image.new("RGB", (W, H), BG)
                if b.kind == "slide":
                    img.paste(b.slide(), (0, 60))
                else:
                    render_sheet(img, wb)
                    render_note(img, wb.note)
                render_chrome(img, b.chapter, ci, len(chap_names), t / self.total, sub)
                prev_img = img
                if fade_p is not None and fade_p > 0:
                    img = Image.blend(img, fade_from, min(1, fade_p))
                last_bytes = img.tobytes()
                last_key = key
                renders += 1
                if preview_times is not None:
                    img.save(os.path.join(preview_dir, f"t{t:07.2f}.png"))
            if proc:
                proc.stdin.write(last_bytes)
        if proc:
            proc.stdin.close()
            proc.wait()
        return {"total": self.total, "frames": nframes, "renders": renders, "chapters": chap_starts}
