"""
Gestion des illustrations par parasha : assets/parsha/<slug>/{1,2,3}.jpg + meta.json

meta.json :
{
  "parsha": "וירא",
  "images": [
    {"file": "1.jpg", "caption": "הכנסת אורחים", "origin": "external" | "ai-generated" | "own",
     "source": "https://i.ytimg.com/vi/hd8rDyheqbI/maxresdefault.jpg...", "author": "...", "license": "..."}
  ]
}
"""
import json
import re
from pathlib import Path

ASSETS_DIR = Path(__file__).resolve().parent / "assets" / "parsha"

# Ordre canonique des 54 parashiot -> slug ASCII (noms de dossiers stables, sans hébreu)
PARSHIOT = [
    ("בראשית", "bereshit"), ("נח", "noach"), ("לך לך", "lech-lecha"), ("וירא", "vayera"),
    ("חיי שרה", "chayei-sara"), ("תולדות", "toldot"), ("ויצא", "vayetzei"), ("וישלח", "vayishlach"),
    ("וישב", "vayeshev"), ("מקץ", "miketz"), ("ויגש", "vayigash"), ("ויחי", "vayechi"),
    ("שמות", "shemot"), ("וארא", "vaera"), ("בא", "bo"), ("בשלח", "beshalach"), ("יתרו", "yitro"),
    ("משפטים", "mishpatim"), ("תרומה", "terumah"), ("תצוה", "tetzaveh"), ("כי תשא", "ki-tisa"),
    ("ויקהל", "vayakhel"), ("פקודי", "pekudei"), ("ויקרא", "vayikra"), ("צו", "tzav"),
    ("שמיני", "shmini"), ("תזריע", "tazria"), ("מצרע", "metzora"), ("אחרי מות", "achrei-mot"),
    ("קדשים", "kedoshim"), ("אמור", "emor"), ("בהר", "behar"), ("בחקתי", "bechukotai"),
    ("במדבר", "bamidbar"), ("נשא", "nasso"), ("בהעלתך", "behaalotecha"), ("שלח לך", "shlach"),
    ("קורח", "korach"), ("חקת", "chukat"), ("בלק", "balak"), ("פינחס", "pinchas"), ("מטות", "matot"),
    ("מסעי", "masei"), ("דברים", "devarim"), ("ואתחנן", "vaetchanan"), ("עקב", "eikev"),
    ("ראה", "reeh"), ("שופטים", "shoftim"), ("כי תצא", "ki-teitzei"), ("כי תבוא", "ki-tavo"),
    ("נצבים", "nitzavim"), ("וילך", "vayeilech"), ("האזינו", "haazinu"), ("וזאת הברכה", "vezot-haberakhah"),
]
_SLUGS = {name: slug for name, slug in PARSHIOT}
_ALIASES = {"קרח": "קורח", "שלח": "שלח לך", "מצורע": "מצרע", "בחוקותי": "בחקתי", "פנחס": "פינחס",
            "בהעלותך": "בהעלתך", "תצווה": "תצוה", "פקודה": "פקודי"}


def normalize_parsha(name: str) -> str:
    """'לך־לך' / 'פרשת לך-לך' -> 'לך לך' (maqaf, tirets et 'פרשת' neutralisés)."""
    name = (name or "").strip()
    name = re.sub(r"^פרשת\s+", "", name)
    name = name.replace("\u05be", " ").replace("־", " ")
    name = re.sub(r"\s+", " ", name).strip()
    return _ALIASES.get(name, name)


def parsha_slug(name: str) -> str:
    """Slug ASCII d'une parasha (ou de parashiot jumelées 'ויקהל-פקודי' -> 'vayakhel-pekudei')."""
    norm = normalize_parsha(name)
    if norm in _SLUGS:
        return _SLUGS[norm]
    parts = [normalize_parsha(p) for p in re.split(r"[-=]", norm) if p.strip()]
    slugs = [_SLUGS.get(p) for p in parts]
    if parts and all(slugs):
        return "-".join(slugs)
    return "other-" + "".join(f"{ord(c):x}" for c in norm)[:40]


def _candidate_folders(name: str):
    """Dossier exact, puis (pour les parashiot jumelées) chaque parasha séparément."""
    norm = normalize_parsha(name)
    seen = [parsha_slug(norm)]
    for part in re.split(r"[-=]", norm):
        part = normalize_parsha(part)
        if part in _SLUGS and _SLUGS[part] not in seen:
            seen.append(_SLUGS[part])
    for slug in seen:
        yield ASSETS_DIR / slug


def load_parsha_images(name: str, max_images: int = 3):
    """Retourne [(PIL.Image RGB, légende, meta), ...] — liste vide si aucun visuel disponible."""
    from PIL import Image

    results = []
    for folder in _candidate_folders(name):
        meta_path = folder / "meta.json"
        if not folder.is_dir():
            continue
        entries = []
        if meta_path.exists():
            try:
                entries = json.loads(meta_path.read_text(encoding="utf-8")).get("images", [])
            except Exception as e:
                print(f"⚠️ meta.json illisible ({meta_path}): {e}")
        if not entries:  # pas de meta : on prend les images présentes, sans légende
            entries = [{"file": p.name, "caption": ""} for p in sorted(folder.iterdir())
                       if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp")]
        for entry in entries:
            path = folder / entry.get("file", "")
            try:
                with Image.open(path) as im:
                    results.append((im.convert("RGB"), entry.get("caption", ""), entry))
            except Exception as e:
                print(f"⚠️ Image de parasha ignorée ({path}): {e}")
            if len(results) >= max_images:
                return results
    return results
