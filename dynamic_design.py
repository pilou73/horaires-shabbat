"""
Rendu « dynamique » de l'affiche hebdomadaire (portrait, 1080 px de large, optimisé WhatsApp).

- Texte hébreu RTL rendu de façon déterministe (python-bidi + layout BASIC de Pillow), donc
  identique sous Windows, Linux et GitHub Actions — avec ou sans libraqm.
- En-tête : logo du template + titre de la parasha + dates + 3 illustrations de la parasha
  (assets/parsha/<slug>/), sinon tuiles décoratives (Magen David).
- Sections « זמני תפילות שבת » / « זמני תפילות אמצע השבוע » / informations du mois.
"""
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps

try:
    from bidi import get_display as _bidi_display
except ImportError:  # python-bidi < 0.5
    try:
        from bidi.algorithm import get_display as _bidi_display
    except ImportError:
        _bidi_display = None

from ai_images import get_parsha_pictures

# ---------------------------------------------------------------------------
# Textes fixes de l'affiche (modifiables ici sans toucher au code de rendu)
# ---------------------------------------------------------------------------
SYNAGOGUE_NAME = "בית הכנסת מרום נווה"
ADDRESS = "רחוב ד\"ר יוחנן בדר 5 רמת גן"
WELCOME = "בואו בשמחה!"
SHABBAT_TITLE = "זמני תפילות שבת"
WEEKDAY_TITLE = "זמני תפילות אמצע השבוע"

# (clé de temps, libellé, note) — même ordre et mêmes notes que le template
SHABBAT_ROWS = [
    ("shir_hashirim", "שיר השירים", ""),
    ("candle_lighting", "כניסת שבת", ""),
    ("mincha_kabbalat", "מנחה וקבלת שבת", ""),
    ("shacharit", "שחרית (קורבנות)", ""),
    ("mincha_gdola", "מנחה גדולה", "לאחר מכן שיעור הדף היומי"),
    ("tehilim", "תהילים לילדים", ""),
    ("shiur_nashim", "שיעור לנשים", ""),
    ("parashat_hashavua", "שיעור פרשת השבוע", "אוליביה (אברהם) כהן"),
    ("shiur_rav", "שיעור", "הרב שוקרון"),
    ("mincha_2", "מנחה נוספת", "סעודה שלישית לאחר מנחה"),
    ("arvit_motsach", "ערבית מוצ״ש", ""),
    ("shabbat_end", "מוצאי שבת קודש", ""),
]
WEEKDAY_ROWS = [
    ("shacharit_hol", "שחרית (הודו)", "ראשון - חמישי / שישי", "06:30 / 06:00"),
    ("mincha_hol", "מנחה", "ראשון - חמישי", None),
    ("arvit_hol", "ערבית", "ראשון - חמישי", None),
]

# Palette (reprise du template : bleu-lavande + turquoise)
NAVY = (36, 48, 112)
TEAL = (52, 160, 182)
TEAL_LIGHT = (222, 242, 246)
LAVENDER = (128, 130, 204)
LAVENDER_LIGHT = (238, 238, 251)
GREEN = (40, 108, 58)
GREY = (95, 100, 120)
WHITE = (255, 255, 255)
PURPLE = (110, 60, 150)
RED = (190, 40, 40)

W = 1080
MARGIN = 48


# ---------------------------------------------------------------------------
# Texte RTL
# ---------------------------------------------------------------------------
_HEB = re.compile(r"[\u0590-\u05FF]")


# HEBREW_RENDER=bidi (défaut) : ordre visuel calculé ici + layout BASIC -> identique partout.
# HEBREW_RENDER=raqm : on laisse libraqm gérer le RTL (seulement si installée sur la machine).
import os
from PIL import features as _features
_USE_RAQM = os.environ.get("HEBREW_RENDER", "bidi").lower() == "raqm" and _features.check("raqm")


def rtl(text: str) -> str:
    """Convertit un texte logique (hébreu + chiffres) en ordre visuel pour un rendu LTR."""
    if _USE_RAQM:
        return text or ""
    if not text or not _HEB.search(text):
        return text or ""
    if _bidi_display:
        # python-bidi inverse l'ordre mais ne « miroite » pas les parenthèses : on le fait ici
        return _bidi_display(text, base_dir="R").translate(str.maketrans("()[]", ")(]["))
    # Repli minimal sans python-bidi : inverse l'ordre des mots et les mots hébreux
    words = text.split(" ")
    out = [w[::-1] if _HEB.search(w) else w for w in reversed(words)]
    return " ".join(out).translate(str.maketrans("()[]", ")(]["))


class Fonts:
    def __init__(self, resources: Path, fallback_regular: Path, fallback_bold: Path):
        fdir = resources / "fonts"
        self._paths = {
            "regular": fdir / "Heebo-Regular.ttf",
            "medium": fdir / "Heebo-Medium.ttf",
            "bold": fdir / "Heebo-Bold.ttf",
            "black": fdir / "Heebo-ExtraBold.ttf",
            "title": fdir / "SuezOne-Regular.ttf",
            "brand": fdir / "SecularOne-Regular.ttf",
        }
        self._fallback = {"regular": fallback_regular, "medium": fallback_regular}
        self._fallback_bold = fallback_bold
        self._cache = {}

    def get(self, kind: str, size: int):
        key = (kind, size)
        if key not in self._cache:
            path = self._paths.get(kind)
            if not path or not path.exists():
                path = self._fallback.get(kind, self._fallback_bold)
            self._cache[key] = ImageFont.truetype(
                str(path), size,
                layout_engine=ImageFont.Layout.RAQM if _USE_RAQM else ImageFont.Layout.BASIC)
        return self._cache[key]


def text_w(draw, text, font):
    return draw.textlength(rtl(text), font=font)


def draw_text(draw, xy, text, font, fill, anchor="ra"):
    """anchor 'ra' = aligné à droite (naturel en hébreu), 'ma' centré, 'la' à gauche."""
    draw.text(xy, rtl(text), font=font, fill=fill, anchor=anchor)


def fit_font(draw, fonts, kind, text, max_w, size, min_size=18):
    while size > min_size and text_w(draw, text, fonts.get(kind, size)) > max_w:
        size -= 2
    return fonts.get(kind, size)


# ---------------------------------------------------------------------------
# Éléments graphiques
# ---------------------------------------------------------------------------
def vertical_gradient(size, top, bottom):
    w, h = size
    grad = Image.new("RGB", (1, h))
    for y in range(h):
        t = y / max(h - 1, 1)
        grad.putpixel((0, y), tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return grad.resize((w, h))


def star_of_david(draw, cx, cy, r, outline, width=3, fill=None):
    import math
    for rot in (90, 270):
        pts = [(cx + r * math.cos(math.radians(rot + k * 120)), cy - r * math.sin(math.radians(rot + k * 120)))
               for k in range(3)]
        draw.polygon(pts, outline=outline, fill=fill, width=width)


def rounded_mask(size, radius):
    mask = Image.new("L", size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius, fill=255)
    return mask


def paste_with_shadow(canvas, img, xy, radius=22, shadow=10):
    x, y = xy
    w, h = img.size
    sh = Image.new("RGBA", (w + shadow * 4, h + shadow * 4), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle((shadow * 2, shadow * 2, shadow * 2 + w, shadow * 2 + h),
                                         radius, fill=(20, 30, 80, 90))
    sh = sh.filter(ImageFilter.GaussianBlur(shadow))
    canvas.alpha_composite(sh, (x - shadow * 2 + 4, y - shadow * 2 + 8))
    canvas.paste(img, (x, y), rounded_mask((w, h), radius))


def decorative_tile(size, seed=0):
    """Tuile de repli quand aucune illustration n'existe pour la parasha."""
    w, h = size
    tile = vertical_gradient(size, (LAVENDER_LIGHT, TEAL_LIGHT, (232, 236, 252))[seed % 3], WHITE).convert("RGBA")
    pattern = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(pattern)
    step = 56
    for row, y in enumerate(range(-step, h + step, step)):
        for x in range(-step, w + step, step):
            star_of_david(d, x + (step // 2 if row % 2 else 0), y, 14, outline=(*LAVENDER, 70), width=2)
    tile.alpha_composite(pattern)
    d = ImageDraw.Draw(tile)
    star_of_david(d, w // 2, h // 2, min(w, h) // 4, outline=TEAL, width=7)
    return tile.convert("RGB")


def prepare_picture(img, size):
    """Recadre (cover) et, pour les gravures en niveaux de gris, applique un duotone de la palette."""
    img = ImageOps.fit(img, size, Image.LANCZOS, centering=(0.5, 0.4))
    small = img.resize((32, 32))
    is_grey = all(max(p) - min(p) < 12 for p in small.getdata())
    if is_grey:
        img = ImageOps.colorize(ImageOps.autocontrast(img.convert("L"), cutoff=1), black=(22, 30, 78),
                                white=(252, 248, 236), mid=(118, 140, 186))
    return img


# ---------------------------------------------------------------------------
# Rendu principal
# ---------------------------------------------------------------------------
def render_schedule(*, resources: Path, template_path: Path, fonts: Fonts, parasha: str,
                    times_text: dict, extra_lines: list, date_line: str, hebrew_date_line: str,
                    is_mevarchim: bool = False, row_notes: dict = None):
    """
    row_notes : {clé: note} — remplace/ajoute la note d'une ligne de SHABBAT_ROWS (ex. hiver)
    times_text : {clé: 'HH:MM'} pour toutes les clés de SHABBAT_ROWS / WEEKDAY_ROWS
    extra_lines : [(texte, couleur_rgb, icône_PIL_ou_None), ...] (molad, lune, tekoufa, jeûnes)
    Retourne une PIL.Image RGB.
    """
    probe = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    row_h, week_row_h, info_h = 54, 56, 50
    header_h = 250
    pics_top = header_h + 10
    pic_w, pic_h = 300, 250
    caption_h = 56
    shabbat_top = pics_top + pic_h + caption_h + 40
    shabbat_h = 80 + row_h * len(SHABBAT_ROWS) + 20
    week_top = shabbat_top + shabbat_h + 26
    week_h = 80 + week_row_h * len(WEEKDAY_ROWS) + 16
    info_top = week_top + week_h + 22
    info_block = (len(extra_lines) * info_h + 40) if (extra_lines or is_mevarchim) else 0
    if is_mevarchim:
        info_block += 70
    footer_h = 150
    H = info_top + info_block + footer_h + 20

    canvas = vertical_gradient((W, H), (247, 249, 255), (240, 248, 250)).convert("RGBA")
    d = ImageDraw.Draw(canvas)

    # Cadre double (rappel du template)
    d.rounded_rectangle((18, 18, W - 18, H - 18), 34, outline=TEAL, width=6)
    d.rounded_rectangle((32, 32, W - 32, H - 32), 26, outline=LAVENDER, width=3)

    # --- En-tête : bandeau + logo ---------------------------------------------------
    band = vertical_gradient((W - 76, header_h - 40), (232, 233, 251), (220, 241, 246)).convert("RGBA")
    canvas.paste(band, (38, 38), rounded_mask(band.size, 22))
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    for x, y in ((70, 66), (70, 206), (W - 420, 66), (W - 420, 206)):  # Magen David discrets
        star_of_david(od, x, y, 18, outline=(*LAVENDER, 90), width=2)
    canvas.alpha_composite(overlay)

    logo_box = (W - 38 - 330, 48, W - 50, header_h - 12)
    logo = None
    try:
        with Image.open(template_path) as tpl:
            logo = tpl.convert("RGB").crop((34, 22, 330, 208))
            # efface les filets du cadre du template qui traversent le haut du logo
            ImageDraw.Draw(logo).rectangle((158 - 34, 68 - 22, logo.width, 99 - 22), fill=WHITE)
    except Exception as e:
        print(f"⚠️ Logo du template indisponible ({e}) — texte utilisé à la place")
    if logo is not None:
        lw = logo_box[2] - logo_box[0]
        lh = int(logo.height * lw / logo.width)
        logo = logo.resize((lw, lh), Image.LANCZOS)
        card = Image.new("RGBA", (lw + 20, lh + 16), (255, 255, 255, 235))
        cy = (header_h + 38 - card.height) // 2 + 6
        canvas.paste(card, (logo_box[0] - 10, cy), rounded_mask(card.size, 18))
        canvas.paste(logo, (logo_box[0], cy + 8))
    else:
        f = fonts.get("brand", 54)
        draw_text(d, (logo_box[2], 80), "בית הכנסת", f, LAVENDER)
        draw_text(d, (logo_box[2], 140), "מרום נווה", fonts.get("brand", 70), LAVENDER)

    # Titre parasha (zone à gauche du logo)
    title_cx = (MARGIN + logo_box[0] - 20) // 2
    title_w = logo_box[0] - 20 - MARGIN - 20
    draw_text(d, (title_cx, 60), "שבת קודש · " + date_line, fonts.get("medium", 28), GREY, anchor="ma")
    ptitle = "פרשת " + parasha.replace("=", "-") if parasha else "שבת שלום"
    tf = fit_font(d, fonts, "title", ptitle, title_w, 84, 44)
    draw_text(d, (title_cx, 98), ptitle, tf, NAVY, anchor="ma")
    draw_text(d, (title_cx, 200), hebrew_date_line, fonts.get("bold", 30), TEAL, anchor="ma")

    # --- 3 illustrations ------------------------------------------------------------
    pictures = get_parsha_pictures(parasha, 3)
    gap = (W - 2 * MARGIN - 3 * pic_w) // 2
    cap_font = fonts.get("bold", 26)
    for i in range(3):
        x = W - MARGIN - (i + 1) * pic_w - i * gap  # ordre RTL : 1re image à droite
        if i < len(pictures):
            img, caption, _ = pictures[i]
            tile = prepare_picture(img, (pic_w, pic_h))
        else:
            tile = decorative_tile((pic_w, pic_h), i)
            caption = "שבת שלום" if (i == 1 and not pictures) else ""
        frame = Image.new("RGB", (pic_w + 12, pic_h + 12), WHITE)
        frame.paste(tile, (6, 6), rounded_mask((pic_w, pic_h), 18))
        paste_with_shadow(canvas, frame, (x - 6, pics_top - 6), radius=24)
        if caption:
            cf = fit_font(d, fonts, "bold", caption, pic_w - 30, 26, 18)
            cw = text_w(d, caption, cf) + 36
            cx = x + pic_w // 2
            cyy = pics_top + pic_h + 14
            d.rounded_rectangle((cx - cw / 2, cyy, cx + cw / 2, cyy + 40), 20, fill=NAVY)
            draw_text(d, (cx, cyy + 20), caption, cf, WHITE, anchor="mm")
    del cap_font

    # --- Section Shabbat --------------------------------------------------------------
    def section(top, height, title, color, light):
        d.rounded_rectangle((MARGIN, top, W - MARGIN, top + height), 26, fill=WHITE, outline=light, width=3)
        tfnt = fonts.get("black", 38)
        tw = text_w(d, title, tfnt) + 130
        d.rounded_rectangle((W / 2 - tw / 2, top - 4, W / 2 + tw / 2, top + 58), 30, fill=color)
        star_of_david(d, W / 2 + tw / 2 - 34, top + 27, 13, outline=WHITE, width=2)
        star_of_david(d, W / 2 - tw / 2 + 34, top + 27, 13, outline=WHITE, width=2)
        draw_text(d, (W / 2, top + 27), title, tfnt, WHITE, anchor="mm")

    label_x = W - MARGIN - 40
    time_cx = MARGIN + 125
    note_cx = (label_x - 300 + time_cx + 90) // 2 - 20

    def time_pill(cy, text, color, wide=False):
        f = fonts.get("bold", 34 if not wide else 26)
        tw = d.textlength(text, font=f) + 36
        tw = max(tw, 130)
        d.rounded_rectangle((time_cx - tw / 2, cy - 22, time_cx + tw / 2, cy + 22), 22, fill=color)
        d.text((time_cx, cy), text, font=f, fill=WHITE, anchor="mm")

    section(shabbat_top, shabbat_h, SHABBAT_TITLE, NAVY, LAVENDER_LIGHT)
    y = shabbat_top + 80
    lf = fonts.get("medium", 33)
    nf = fonts.get("regular", 25)
    for idx, (key, label, note) in enumerate(SHABBAT_ROWS):
        note = (row_notes or {}).get(key, note)
        cy = y + row_h // 2
        if idx % 2 == 0:
            d.rounded_rectangle((MARGIN + 16, y + 3, W - MARGIN - 16, y + row_h - 3), 16, fill=LAVENDER_LIGHT)
        draw_text(d, (label_x, cy), label, lf, NAVY, anchor="rm")
        if note:
            nfit = fit_font(d, fonts, "regular", note, 330, 25, 18)
            draw_text(d, (note_cx, cy), note, nfit, GREY, anchor="mm")
        t = times_text.get(key, "")
        if t:
            time_pill(cy, t, TEAL if key not in ("candle_lighting", "shabbat_end") else LAVENDER)
        y += row_h

    # --- Section semaine --------------------------------------------------------------
    section(week_top, week_h, WEEKDAY_TITLE, GREEN, TEAL_LIGHT)
    y = week_top + 80
    for idx, (key, label, days, fixed) in enumerate(WEEKDAY_ROWS):
        cy = y + week_row_h // 2
        if idx % 2 == 0:
            d.rounded_rectangle((MARGIN + 16, y + 3, W - MARGIN - 16, y + week_row_h - 3), 16,
                                fill=(232, 245, 236))
        draw_text(d, (label_x, cy), label, lf, GREEN, anchor="rm")
        draw_text(d, (note_cx, cy), days, nf, GREY, anchor="mm")
        t = fixed or times_text.get(key, "")
        if t:
            time_pill(cy, t, GREEN, wide=bool(fixed))
        y += week_row_h

    # --- Informations (molad, lune, tekoufa, jeûnes) -----------------------------------
    if info_block:
        top = info_top
        d.rounded_rectangle((MARGIN, top, W - MARGIN, top + info_block - 10), 26, fill=(250, 250, 255),
                            outline=TEAL_LIGHT, width=3)
        y = top + 15
        if is_mevarchim:
            bf = fonts.get("title", 48)
            draw_text(d, (W / 2, y + 28), "שבת מברכין", bf, PURPLE, anchor="mm")
            y += 70
        for text, color, icon in extra_lines:
            cy = y + info_h // 2
            right = W - MARGIN - 36
            if icon is not None:
                ic = icon.convert("RGBA").resize((40, 40), Image.LANCZOS)
                canvas.alpha_composite(ic, (right - 40, cy - 20))
                right -= 54
            f = fit_font(d, fonts, "medium", text, right - MARGIN - 40, 28, 18)
            draw_text(d, (right, cy), text, f, color, anchor="rm")
            y += info_h

    # --- Pied de page -------------------------------------------------------------------
    fy = H - footer_h - 20
    d.line((MARGIN + 30, fy + 10, W - MARGIN - 30, fy + 10), fill=TEAL, width=3)
    star_of_david(d, W // 2, fy + 10, 16, outline=LAVENDER, width=3, fill=WHITE)
    draw_text(d, (W - MARGIN - 30, fy + 72), WELCOME, fonts.get("title", 54), LAVENDER, anchor="rm")
    draw_text(d, (MARGIN + 30, fy + 60), SYNAGOGUE_NAME, fonts.get("bold", 30), NAVY, anchor="lm")
    draw_text(d, (MARGIN + 30, fy + 100), ADDRESS, fonts.get("regular", 28), GREY, anchor="lm")

    return canvas.convert("RGB")
