import re
import unicodedata


def clean_text(value):
    if value is None:
        return ""

    value = str(value).strip()
    value = re.sub(r"\s+", " ", value)

    return value


def normalize_key(value):
    value = clean_text(value).lower()

    value = unicodedata.normalize("NFKD", value)
    value = "".join(c for c in value if not unicodedata.combining(c))

    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()

    return value


def title_name(value):
    value = clean_text(value)

    if not value:
        return ""

    small_words = {"de", "del", "la", "las", "los", "y", "en", "con"}

    words = []
    for word in value.lower().split():
        if word in small_words:
            words.append(word)
        else:
            words.append(word.capitalize())

    return " ".join(words)


def normalize_species(value):
    value = normalize_key(value)

    if not value:
        return ""

    if "can" in value and "fel" in value:
        return "Canino y felino"

    if "perro" in value or "canino" in value or value == "can":
        return "Canino"

    if "gato" in value or "felino" in value or value == "fel":
        return "Felino"

    return title_name(value)


def normalize_route(value):
    value = normalize_key(value)

    route_map = {
        "vo": "VO",
        "oral": "VO",
        "via oral": "VO",
        "sc": "SC",
        "subcutanea": "SC",
        "subcutaneo": "SC",
        "im": "IM",
        "intramuscular": "IM",
        "ev": "EV",
        "iv": "EV",
        "endovenosa": "EV",
        "intravenosa": "EV",
        "topica": "Tópica",
        "otica": "Ótica",
        "oftalmica": "Oftálmica",
    }

    return route_map.get(value, title_name(value))


def normalize_category(value):
    return title_name(value)


def normalize_laboratory(value):
    return title_name(value)


def split_active_ingredients(value):
    value = clean_text(value)

    if not value:
        return []

    parts = re.split(r"\s*\+\s*|,|/|;", value)
    parts = [title_name(p) for p in parts if clean_text(p)]

    return parts
