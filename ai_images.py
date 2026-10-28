"""
Illustrations de parasha générées par IA, à la volée (rien à stocker à l'avance dans le dépôt).

Ordre de recherche des 3 images (variable d'environnement PARSHA_IMAGE_SOURCE) :
  "ai"    (défaut) : IA -> images locales assets/parsha/<slug>/ -> tuiles décoratives
  "local"          : images locales -> tuiles décoratives (aucun appel réseau)
  "none"           : tuiles décoratives uniquement

Fournisseur IA (variable AI_IMAGE_PROVIDER) :
  "pollinations" (défaut) : service gratuit https://pollinations.ai, sans clé API
  "openai"                : nécessite le secret OPENAI_API_KEY (payant, ~0.04 $ par image)

Les images générées sont mises en cache dans output/ai_cache/<slug>/ (dossier ignoré par git),
pour ne pas régénérer à chaque exécution de la même semaine.
TOUTES ces images sont produites par IA : elles ne proviennent d'aucun site ni d'aucune œuvre existante.
"""
import base64
import hashlib
import io
import json
import os
import time
import urllib.parse
from pathlib import Path

import requests
from PIL import Image

from parsha_assets import load_parsha_images, normalize_parsha, parsha_slug

CACHE_DIR = Path(__file__).resolve().parent / "output" / "ai_cache"

# Style commun. Les modèles d'image gèrent mal les négations (« no X » fait souvent apparaître X) :
# on décrit donc ce qu'on VEUT voir, et on ne parle de personnages que dans les scènes qui en ont.
STYLE = ("soft watercolor illustration, ancient Land of Israel, gentle golden light, "
         "clean simple composition, without any text or lettering")

# parasha -> 3 x (légende hébraïque, description de la scène en anglais pour l'IA)
THEMES = {
    "בראשית": [("בריאת העולם", "the creation of the world, sun moon and stars over sea and land"),
               ("גן עדן", "lush garden of Eden with a tree of fruit, no people"),
               ("השבת", "two shabbat candles glowing on a table, peaceful evening")],
    "נח": [("תיבת נח", "Noah's wooden ark floating on a great flood"),
           ("החיות בתיבה", "pairs of animals walking into a wooden ark"),
           ("הקשת בענן", "rainbow over mountains after the flood, a white dove with olive branch")],
    "לך לך": [("לך לך מארצך", "Abraham and Sarah journeying with camels and sheep through the desert"),
              ("כוכבי השמים", "old man looking up at a sky full of countless stars in the desert"),
              ("ברית בין הבתרים", "desert night with a torch of fire and a smoking furnace")],
    "וירא": [("הכנסת אורחים", "an elderly bearded man in a long robe and head covering standing at the open entrance "
                              "of a large woven goat-hair tent in the desert, welcoming three male travelers "
                              "in long robes, a big tree giving shade, bread and water on a low table"),
             ("הפיכת סדום", "a burning ancient city, huge orange flames and black smoke, fire falling from a dark red sky"),
             ("עקדת יצחק", "close-up of a male sheep ram with spiral horns stuck in a thicket, "
                           "stacked firewood on stone altar behind, hills at dawn")],
    "חיי שרה": [("מערת המכפלה", "an ancient stone cave entrance in Hebron hills"),
                ("רבקה ליד הבאר", "young woman giving water to camels from a well"),
                ("יצחק ורבקה", "a caravan of camels arriving at sunset in a field")],
    "תולדות": [("יעקב ועשו", "twin brothers, one a hunter with a bow, one a shepherd near tents"),
               ("נזיד עדשים", "a pot of red lentil stew with bread in a tent"),
               ("ברכת יצחק", "an elderly blind father blessing his son inside a tent")],
    "ויצא": [("סולם יעקב", "a ladder reaching to the sky with angels, a man sleeping on stones"),
             ("רחל הרועה", "shepherdess with a flock of sheep near a well"),
             ("צאן יעקב", "flock of speckled and spotted sheep and goats in pasture")],
    "וישלח": [("מנחה לעשו", "herds of goats, camels and cattle sent across the desert"),
              ("מעבר יבוק", "a river crossing at dawn, a man wrestling with an angel"),
              ("פגישת האחים", "two brothers embracing in the desert, camps behind them")],
    "וישב": [("כתונת הפסים", "a colorful striped coat"),
             ("חלומות יוסף", "sheaves of wheat bowing to one sheaf, sun moon and eleven stars"),
             ("יוסף בבור", "an empty pit in the desert, a caravan of camels passing")],
    "מקץ": [("חלום פרעה", "seven fat cows and seven thin cows by the Nile"),
            ("שבע שבלים", "seven full ears of grain and seven thin ears"),
            ("אוצרות התבואה", "ancient Egyptian granaries full of wheat")],
    "ויגש": [("יוסף מתוודע", "a viceroy of Egypt embracing his brothers in tears"),
             ("העגלות", "wagons travelling from Egypt to Canaan"),
             ("ארץ גושן", "green pastures of Goshen with flocks and tents")],
    "ויחי": [("ברכת יעקב", "an elderly patriarch blessing his two grandsons"),
             ("שנים עשר השבטים", "twelve symbols of the tribes of Israel on banners"),
             ("עלייה לקבורה", "a procession of chariots traveling to Canaan")],
    "שמות": [("משה ביאור", "a baby in a woven basket among reeds of the Nile"),
             ("הסנה הבוער", "a desert bush burning with fire but not consumed"),
             ("עבודת פרך", "Egyptian pyramids and bricks of clay"),],
    "וארא": [("מטה לתנין", "a shepherd's staff turning into a serpent in pharaoh's palace"),
             ("מכת דם", "the Nile river turning red"),
             ("מכת צפרדע", "frogs everywhere in an Egyptian palace")],
    "בא": [("מכת ארבה", "a swarm of locusts over Egyptian fields"),
           ("קרבן פסח", "lamb, matzah and bitter herbs on a table at night"),
           ("יציאת מצרים", "a great multitude leaving Egypt at dawn with flocks")],
    "בשלח": [("קריעת ים סוף", "the sea split into two walls of water with a dry path"),
             ("שירת הים", "women dancing with tambourines by the sea"),
             ("המן", "white manna on the desert ground at morning")],
    "יתרו": [("הר סיני", "Mount Sinai with clouds, smoke and lightning"),
             ("עשרת הדברות", "two stone tablets with rounded tops, blank, no text"),
             ("עצת יתרו", "an old Midianite priest advising a leader in a tent")],
    "משפטים": [("משפטים", "ancient scales of justice"),
               ("שלוש רגלים", "pilgrims walking to Jerusalem with baskets of first fruits"),
               ("נעשה ונשמע", "the people of Israel gathered at the foot of a mountain")],
    "תרומה": [("ארון הברית", "the golden Ark of the Covenant with two cherubim"),
              ("המנורה", "the golden seven-branched menorah"),
              ("המשכן", "the Tabernacle in the desert with curtains and courtyard")],
    "תצוה": [("בגדי כהונה", "the High Priest garments with the breastplate of twelve stones"),
             ("שמן זית", "pure olive oil and an olive branch"),
             ("מזבח הקטורת", "a golden incense altar with fragrant smoke")],
    "כי תשא": [("מחצית השקל", "ancient silver half-shekel coins"),
               ("לוחות הברית", "two stone tablets on a mountain, blank, no text"),
               ("הכיור", "a copper laver basin with water")],
    "ויקהל": [("השבת", "shabbat candles and challah bread"),
              ("נדבת המשכן", "people bringing gold, fabrics and wood as gifts"),
              ("בצלאל האומן", "an artisan crafting golden vessels")],
    "פקודי": [("הקמת המשכן", "the Tabernacle erected in the desert camp"),
              ("ענן השכינה", "a cloud resting over the Tabernacle"),
              ("כלי המשכן", "golden menorah, table of showbread and altar")],
    "ויקרא": [("המזבח", "a stone and copper altar in a courtyard"),
              ("קרבן מנחה", "fine flour, oil and frankincense in vessels"),
              ("המלח", "a bowl of salt beside bread")],
    "צו": [("אש תמיד", "an eternal fire burning on an altar"),
           ("הכהנים", "priests in white linen garments"),
           ("מילואים", "the Tabernacle courtyard at dawn")],
    "שמיני": [("אש מן השמים", "fire descending onto an altar"),
              ("סימני כשרות", "a fish with fins and scales, and a cow"),
              ("השמחה במשכן", "people rejoicing before the Tabernacle")],
    "תזריע": [("ברית מילה", "a baby wrapped in a white tallit, candles"),
              ("שמירת הלשון", "a quiet path, a person with finger on lips"),
              ("הכהן", "a priest in white garments")],
    "מצרע": [("טהרה", "two birds, cedar wood and hyssop"),
             ("הבית", "an ancient stone house in Israel"),
             ("המעיין", "a spring of living water")],
    "אחרי מות": [("יום הכיפורים", "the High Priest in white linen garments"),
                 ("השעיר לעזאזל", "a goat on a desert cliff"),
                 ("הקטורת", "incense smoke rising in a sanctuary")],
    "קדשים": [("ואהבת לרעך כמוך", "two people helping each other, warm light"),
              ("פאת השדה", "a corner of a wheat field left for the poor"),
              ("והדרת פני זקן", "a young man rising to honor an elder")],
    "אמור": [("המועדים", "shofar, matzah, sukkah and lulav together"),
             ("ספירת העומר", "sheaves of barley in a field"),
             ("המנורה", "the golden menorah with olive oil")],
    "בהר": [("שמיטה", "a field resting, vines and wildflowers"),
            ("היובל", "a shofar being blown over the land"),
            ("הר סיני", "Mount Sinai at sunrise")],
    "בחקתי": [("ונתתי גשמיכם", "rain falling on green fields"),
              ("פרי העץ", "trees full of fruit in an orchard"),
              ("שלום בארץ", "a peaceful village in Israel")],
    "במדבר": [("דגלי השבטים", "colorful tribal banners around a desert camp"),
              ("מחנה ישראל", "tents arranged in a square around the Tabernacle"),
              ("המדבר", "the Sinai desert at sunrise")],
    "נשא": [("ברכת כהנים", "hands of a priest raised in blessing under a tallit"),
            ("קרבנות הנשיאים", "twelve wagons and oxen brought as gifts"),
            ("המשכן", "the Tabernacle in the desert")],
    "בהעלתך": [("הדלקת המנורה", "a priest lighting the golden menorah"),
               ("החצוצרות", "two silver trumpets"),
               ("השליו", "quails flying over the desert camp")],
    "שלח לך": [("המרגלים", "scouts carrying a huge cluster of grapes on a pole"),
               ("ארץ זבת חלב ודבש", "fertile land of Israel with honey, milk, grapes and figs"),
               ("ציצית", "tzitzit fringes with a blue thread")],
    "קורח": [("מחלוקת קורח", "a crowd gathered around tents in the desert"),
             ("מטה אהרן", "a wooden staff blossoming with almond flowers"),
             ("המחתות", "copper fire pans")],
    "חקת": [("פרה אדומה", "a red heifer in a field"),
            ("מי מריבה", "water gushing from a rock in the desert"),
            ("נחש הנחושת", "a copper serpent on a pole")],
    "בלק": [("האתון", "a donkey stopping on a narrow vineyard path"),
            ("מה טובו אהליך", "tents of Israel spread in a valley seen from a mountain"),
            ("במות בעל", "a mountain top with seven altars at sunrise")],
    "פינחס": [("בנות צלפחד", "five sisters standing before elders"),
              ("יהושע", "an elder laying hands on a young leader"),
              ("קרבנות המועדים", "festival offerings of wheat and wine")],
    "מטות": [("נדרים", "a man raising his hand in an oath"),
             ("שבטי ראובן וגד", "flocks of sheep on green land east of the Jordan"),
             ("ערי מקלט", "a walled ancient city on a hill")],
    "מסעי": [("מסעי בני ישראל", "a long caravan traveling through the desert"),
             ("גבולות הארץ", "map-like landscape of the land of Israel, no text"),
             ("נהר הירדן", "the Jordan river among trees")],
    "דברים": [("נאום משה", "an old leader addressing a great crowd"),
              ("ערבות מואב", "plains of Moab with the Jordan in the distance"),
              ("ספר התורה", "a Torah scroll"),],
    "ואתחנן": [("שמע ישראל", "a mezuzah on a doorpost"),
               ("תפילין", "tefillin and a tallit"),
               ("הר נבו", "view of the land of Israel from Mount Nebo")],
    "עקב": [("שבעת המינים", "wheat, barley, grapes, figs, pomegranates, olives and dates"),
            ("ברכת המזון", "a set table with bread"),
            ("ארץ טובה", "streams and valleys of the land of Israel")],
    "ראה": [("הר גריזים והר עיבל", "two mountains facing each other, one green and one bare"),
            ("צדקה", "a hand giving a coin to the poor"),
            ("שלושה רגלים", "pilgrims going up to Jerusalem")],
    "שופטים": [("צדק צדק תרדוף", "ancient scales of justice"),
               ("השופטים", "elders sitting at the city gate"),
               ("עץ השדה", "a fruit tree standing in a field")],
    "כי תצא": [("שילוח הקן", "a mother bird flying away from a nest with eggs"),
               ("מעקה לגג", "an ancient house with a railing on the roof"),
               ("זכור את עמלק", "the desert of Sinai")],
    "כי תבוא": [("ביכורים", "a basket of first fruits"),
                ("אבנים גדולות", "large stones plastered on a mountain"),
                ("ארץ זבת חלב ודבש", "honey, milk and fruits of Israel")],
    "נצבים": [("אתם נצבים", "the whole people standing together before a leader"),
              ("ובחרת בחיים", "a tree of life full of fruit"),
              ("התשובה", "sunrise over Jerusalem")],
    "וילך": [("חזק ואמץ", "an elder encouraging a young leader"),
             ("הקהל", "a large gathering listening to the Torah"),
             ("ספר התורה", "a Torah scroll in an ark")],
    "האזינו": [("שירת האזינו", "heaven and earth, sky and land"),
               ("כנשר יעיר קנו", "an eagle hovering over its nest"),
               ("הר נבו", "Mount Nebo at sunset")],
    "וזאת הברכה": [("ברכת השבטים", "twelve tribal banners"),
                   ("שמחת תורה", "people dancing with Torah scrolls"),
                   ("הר נבו", "view from a mountain over the promised land")],
}


def _themes_for(parasha: str):
    """3 thèmes pour une parasha (ou parashiot jumelées : on mélange les deux)."""
    norm = normalize_parsha(parasha)
    if norm in THEMES:
        return THEMES[norm]
    parts = [normalize_parsha(p) for p in norm.replace("=", "-").split("-") if p.strip()]
    found = [THEMES[p] for p in parts if p in THEMES]
    if len(found) >= 2:
        return [found[0][0], found[1][0], found[0][1]]
    if found:
        return found[0]
    return [("שבת שלום", "shabbat candles and challah on a white tablecloth"),
            ("ספר התורה", "an open Torah scroll"),
            ("ירושלים", "the old city of Jerusalem at golden hour")]


_last_call = [0.0]


def _gen_pollinations(prompt: str, seed: int) -> Image.Image:
    """Le palier gratuit (anonyme) limite à ~1 requête / 15 s : on espace les appels et on réessaie."""
    url = ("https://image.pollinations.ai/prompt/" + urllib.parse.quote(prompt)
           + f"?width=640&height=512&nologo=true&seed={seed}")
    for attempt in range(3):
        wait = 16 - (time.time() - _last_call[0])
        if wait > 0:
            time.sleep(wait)
        r = requests.get(url, timeout=120)
        _last_call[0] = time.time()
        if r.status_code not in (402, 429):
            break
        print(f"   … limite de débit Pollinations ({r.status_code}), nouvel essai {attempt + 2}/3")
        time.sleep(20)
    r.raise_for_status()
    if not r.headers.get("content-type", "").startswith("image/"):
        raise RuntimeError(f"réponse inattendue : {r.headers.get('content-type')}")
    return Image.open(io.BytesIO(r.content)).convert("RGB")


def _gen_openai(prompt: str, seed: int) -> Image.Image:
    key = os.environ.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("OPENAI_API_KEY absent")
    r = requests.post("https://lh7-rt.googleusercontent.com/docsz/AD_4nXdaOMUQlf2nGVeu5yCty8MX8iLC76poqhqF2kzzxesX_ahHfkDS6HKBuRA607X4SF16Bj-djBwb6XRl7BSET6lpqEuUOvtpVZbRhSvzbQK7FEeBlmEcTtwFFIzwEfqrKJYDPnCL?key=k24e6zQVyhH8gZBIv2TGLQ",
                      headers={"Authorization": f"Bearer {key}"},
                      json={"model": os.environ.get("OPENAI_IMAGE_MODEL", "gpt-image-1"),
                            "prompt": prompt, "size": "1024x1024", "n": 1},
                      timeout=180)
    r.raise_for_status()
    data = r.json()["data"][0]
    raw = base64.b64decode(data["b64_json"]) if data.get("b64_json") else requests.get(data["url"], timeout=60).content
    return Image.open(io.BytesIO(raw)).convert("RGB")


def generate_ai_images(parasha: str, max_images: int = 3):
    """Retourne [(image, légende, meta), ...] générées par IA (avec cache local)."""
    provider = os.environ.get("AI_IMAGE_PROVIDER", "pollinations").lower()
    gen = _gen_openai if provider == "openai" else _gen_pollinations
    folder = CACHE_DIR / parsha_slug(parasha)
    folder.mkdir(parents=True, exist_ok=True)
    results = []
    for i, (caption, scene) in enumerate(_themes_for(parasha)[:max_images], start=1):
        prompt = f"{scene}, {STYLE}"
        key = hashlib.sha1(f"{provider}|{prompt}".encode("utf-8")).hexdigest()[:8]
        path = folder / f"{i}_{key}.jpg"  # nouveau nom si le thème est modifié -> régénération
        meta = {"file": path.name, "caption": caption, "origin": "ai-generated", "provider": provider,
                "prompt": prompt}
        try:
            if path.exists():
                img = Image.open(path).convert("RGB")
            else:
                print(f"🎨 Génération IA ({provider}) : {caption}")
                img = gen(meta["prompt"], seed=1000 + i)
                img.save(path, quality=90)
            results.append((img, caption, meta))
        except Exception as e:
            print(f"⚠️ Génération IA impossible pour « {caption} » : {e}")
    (folder / "meta.json").write_text(json.dumps({"parsha": parasha, "images": [m for _, _, m in results]},
                                                 ensure_ascii=False, indent=2), encoding="utf-8")
    return results


def get_parsha_pictures(parasha: str, max_images: int = 3):
    """Point d'entrée unique utilisé par le rendu. Ne lève jamais d'exception."""
    if not parasha:
        return []
    source = os.environ.get("PARSHA_IMAGE_SOURCE", "ai").lower()
    pictures = []
    try:
        if source == "ai":
            pictures = generate_ai_images(parasha, max_images)
        if source in ("ai", "local") and len(pictures) < max_images:
            pictures += load_parsha_images(parasha, max_images - len(pictures))
    except Exception as e:
        print(f"⚠️ Illustrations indisponibles : {e}")
    return pictures[:max_images]
