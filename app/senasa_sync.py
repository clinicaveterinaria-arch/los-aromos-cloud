"""
senasa_sync.py
Sincronizador SENASA -> Vademécum Aromos Cloud.

Uso esperado desde main.py:

    from .senasa_sync import update_from_senasa

    summary = update_from_senasa(db)

No usa Selenium. Consume la API JSON pública que usa el Vademécum oficial de SENASA.
"""

from __future__ import annotations

import json
import re
import time
import unicodedata
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

from sqlalchemy import text
from sqlalchemy.orm import Session


API_BASE = "https://aps2.senasa.gov.ar/adt_api/api"
PRODUCTS_BASE = f"{API_BASE}/productosFarmacos"
INDEX_ENDPOINT = f"{PRODUCTS_BASE}/search/publicSearchProductoFarmacoDTO"
DETAIL_ENDPOINT = f"{PRODUCTS_BASE}/search/publicSearchProducto"
DETAIL_PROJECTION = "productoFarmacoDetallePublicoProjection"
INDEX_PROJECTION = "productoFarmacoLiteProjection"

DEFAULT_PAGE_SIZE = 500
DEFAULT_TIMEOUT = 40
DEFAULT_SLEEP_SECONDS = 0.03


@dataclass
class SenasaRecord:
    active_ingredient: str
    brand_name: str = ""
    laboratory: str = ""
    presentation: str = ""
    concentration: str = ""
    species: str = ""
    category: str = ""
    route: str = ""
    frequency: str = ""
    indications: str = ""
    dog_dose: str = ""
    cat_dose: str = ""
    observations: str = ""
    senasa_id: str = ""
    certificate: str = ""


@dataclass
class SenasaSummary:
    total_index: int = 0
    details_ok: int = 0
    details_error: int = 0
    skipped: int = 0
    new_active: int = 0
    updated_active: int = 0
    new_brands: int = 0
    updated_brands: int = 0
    errors: List[str] = field(default_factory=list)

    def as_dict(self) -> Dict[str, Any]:
        return {
            "total_index": self.total_index,
            "details_ok": self.details_ok,
            "details_error": self.details_error,
            "skipped": self.skipped,
            "new_active": self.new_active,
            "updated_active": self.updated_active,
            "new_brands": self.new_brands,
            "updated_brands": self.updated_brands,
            "errors": self.errors,
        }


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple, set)):
        return ", ".join(clean_text(v) for v in value if clean_text(v))
    text_value = str(value).replace("\r", "\n")
    text_value = re.sub(r"[ \t]+", " ", text_value)
    text_value = re.sub(r"\n{3,}", "\n\n", text_value)
    return text_value.strip()


def normalize_key(value: Any) -> str:
    value = clean_text(value).lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def unique_join(values: Iterable[Any], sep: str = ", ") -> str:
    seen = set()
    result = []
    for value in values:
        text_value = clean_text(value)
        if not text_value:
            continue
        key = normalize_key(text_value)
        if key in seen:
            continue
        seen.add(key)
        result.append(text_value)
    return sep.join(result)


def http_json(url: str, params: Optional[Dict[str, Any]] = None, timeout: int = DEFAULT_TIMEOUT) -> Any:
    if params:
        query = urllib.parse.urlencode(params, doseq=True)
        separator = "&" if "?" in url else "?"
        url = f"{url}{separator}{query}"

    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0 AromosCloud/1.0",
        },
        method="GET",
    )

    with urllib.request.urlopen(request, timeout=timeout) as response:
        raw = response.read().decode("utf-8", errors="replace")

    return json.loads(raw)


def extract_items(payload: Any) -> List[Dict[str, Any]]:
    """Soporta HAL, listas, content/page y respuestas mixtas."""
    if payload is None:
        return []
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if not isinstance(payload, dict):
        return []

    if isinstance(payload.get("content"), list):
        return [item for item in payload["content"] if isinstance(item, dict)]

    embedded = payload.get("_embedded")
    if isinstance(embedded, dict):
        items: List[Dict[str, Any]] = []
        for value in embedded.values():
            if isinstance(value, list):
                items.extend([item for item in value if isinstance(item, dict)])
        return items

    for key in ["data", "results", "items", "productos", "productosFarmacos"]:
        if isinstance(payload.get(key), list):
            return [item for item in payload[key] if isinstance(item, dict)]

    if "id" in payload or "numeroInscripcion" in payload:
        return [payload]

    return []


def payload_has_next(payload: Any, current_page: int, items_count: int) -> bool:
    if not isinstance(payload, dict):
        return False

    page_info = payload.get("page")
    if isinstance(page_info, dict):
        total_pages = page_info.get("totalPages")
        number = page_info.get("number", current_page)
        if isinstance(total_pages, int):
            return number + 1 < total_pages

    links = payload.get("_links")
    if isinstance(links, dict) and links.get("next"):
        return True

    return items_count >= DEFAULT_PAGE_SIZE


def get_product_href(row: Dict[str, Any]) -> str:
    links = row.get("_links") or {}
    if isinstance(links, dict):
        for key in ["productoFarmaco", "producto", "self"]:
            link = links.get(key)
            if isinstance(link, dict) and link.get("href"):
                return clean_href(link["href"])

    for key in ["producto", "productoFarmaco"]:
        value = row.get(key)
        if isinstance(value, dict):
            nested_links = value.get("_links") or {}
            if isinstance(nested_links, dict):
                self_link = nested_links.get("self")
                if isinstance(self_link, dict) and self_link.get("href"):
                    return clean_href(self_link["href"])
            if value.get("id"):
                return f"{PRODUCTS_BASE}/{value['id']}"

    product_id = row.get("id") or row.get("idProducto")
    if product_id:
        return f"{PRODUCTS_BASE}/{product_id}"

    return ""


def clean_href(href: str) -> str:
    href = clean_text(href)
    href = href.replace("{?projection}", "")
    return href


def fetch_index(max_pages: int = 0, page_size: int = DEFAULT_PAGE_SIZE) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    page = 0

    while True:
        params = {
            "projection": INDEX_PROJECTION,
            "page": page,
            "size": page_size,
        }
        payload = http_json(INDEX_ENDPOINT, params=params)
        items = extract_items(payload)
        rows.extend(items)

        if max_pages and page + 1 >= max_pages:
            break
        if not payload_has_next(payload, page, len(items)):
            break
        if not items:
            break

        page += 1

    return rows


def fetch_product_detail(product_href: str) -> Dict[str, Any]:
    params = {
        "producto": product_href,
        "projection": DETAIL_PROJECTION,
    }
    return http_json(DETAIL_ENDPOINT, params=params)


def component_name(component_row: Dict[str, Any]) -> str:
    component = component_row.get("componente")
    if isinstance(component, dict):
        return clean_text(component.get("nombre"))
    return clean_text(component_row.get("nombre"))


def is_active_component(component_row: Dict[str, Any]) -> bool:
    component = component_row.get("componente")
    if not isinstance(component, dict):
        return True

    tipo = component.get("tipoComponente")
    if isinstance(tipo, dict):
        tipo_text = clean_text(tipo.get("nombre") or tipo.get("descripcion")).upper()
        if "PRINCIPIO ACTIVO" in tipo_text:
            return True
        if "EXCIPIENTE" in tipo_text:
            return False

    return True


def component_concentration(component_row: Dict[str, Any]) -> str:
    amount = clean_text(
        component_row.get("cantidadCompleja")
        or component_row.get("cantidad")
        or component_row.get("cantidadAgregada")
    )
    unit = ""
    unidad = component_row.get("unidadMedida")
    if isinstance(unidad, dict):
        unit = clean_text(unidad.get("siglaEstandarizada") or unidad.get("descripcion"))

    formulation = ""
    form = component_row.get("formulacion")
    if isinstance(form, dict):
        formulation = clean_text(form.get("nombre"))

    parts = []
    if amount:
        parts.append(amount)
    if unit:
        parts.append(unit)
    text_value = " ".join(parts)
    if formulation and text_value:
        text_value = f"{text_value} ({formulation})"
    return text_value


def parse_active_ingredients(detail: Dict[str, Any]) -> Tuple[str, str]:
    components = detail.get("componentesPorProducto") or []
    names = []
    concentrations = []

    for row in components:
        if not isinstance(row, dict):
            continue
        if not is_active_component(row):
            continue
        name = component_name(row)
        if not name:
            continue
        names.append(name)
        concentration = component_concentration(row)
        if concentration:
            concentrations.append(f"{name}: {concentration}")

    return unique_join(names, " + "), unique_join(concentrations, " | ")
def parse_brand(detail: Dict[str, Any]) -> str:
    signatures = detail.get("productosFirmas") or []
    names = []
    for row in signatures:
        if isinstance(row, dict):
            names.append(row.get("nombreComercial"))
    return unique_join(names) or clean_text(detail.get("nombreProducto"))


def parse_laboratory(detail: Dict[str, Any]) -> str:
    signatures = detail.get("productosFirmas") or []
    labs = []
    for row in signatures:
        if not isinstance(row, dict):
            continue
        firma = row.get("firma")
        if isinstance(firma, dict):
            labs.append(firma.get("nombre") or firma.get("nombreFantasia"))
    return unique_join(labs)


def parse_species_and_dose(detail: Dict[str, Any]) -> Tuple[str, str, str]:
    rows = detail.get("productoEspecieCategoria") or []
    species = []
    dog_doses = []
    cat_doses = []
    all_doses = []

    for row in rows:
        if not isinstance(row, dict):
            continue
        especie = row.get("especie")
        species_text = ""
        if isinstance(especie, dict):
            species_text = clean_text(especie.get("descripcion"))
        species.append(species_text)

        dose = clean_text(row.get("dosificacion"))
        if dose:
            label = f"{species_text}: {dose}" if species_text else dose
            all_doses.append(label)
            key = normalize_key(species_text)
            if "canino" in key or "perro" in key:
                dog_doses.append(dose)
            if "felino" in key or "gato" in key:
                cat_doses.append(dose)

    if not dog_doses:
        dog_doses = all_doses
    if not cat_doses:
        cat_doses = all_doses

    return unique_join(species), unique_join(dog_doses, "\n"), unique_join(cat_doses, "\n")


def parse_routes(detail: Dict[str, Any]) -> str:
    routes = []
    for row in detail.get("productosFarmacoViaPorProducto") or []:
        if not isinstance(row, dict):
            continue
        via = row.get("via") or row.get("viaAdministracion") or row.get("productoFarmacoVia")
        if isinstance(via, dict):
            routes.append(via.get("descripcion") or via.get("nombre"))
        else:
            routes.append(row.get("descripcion") or row.get("nombre"))

    indications = clean_text(detail.get("indicacionesYVias"))
    if not routes and indications:
        matches = re.findall(r"V[ií]a\s+([^\.\n]+)", indications, flags=re.IGNORECASE)
        routes.extend(matches)

    return unique_join(routes)


def parse_presentation(detail: Dict[str, Any]) -> str:
    values = []

    tipo_presentacion = detail.get("tipoPresentacion")
    if isinstance(tipo_presentacion, dict):
        values.append(tipo_presentacion.get("descripcion") or tipo_presentacion.get("nombre"))

    for envase in detail.get("envases") or []:
        if not isinstance(envase, dict):
            continue
        parts = []
        for key in ["descripcion", "nombre", "cantidad", "contenido", "capacidad"]:
            if envase.get(key):
                parts.append(clean_text(envase.get(key)))
        unidad = envase.get("unidadMedida")
        if isinstance(unidad, dict):
            parts.append(clean_text(unidad.get("siglaEstandarizada") or unidad.get("descripcion")))
        if parts:
            values.append(" ".join(parts))

    return unique_join(values)


def parse_category(detail: Dict[str, Any]) -> str:
    values = []
    tipo = detail.get("tipoProducto")
    if isinstance(tipo, dict):
        values.append(tipo.get("descripcion"))
    nomenclador = detail.get("nomenclador")
    if isinstance(nomenclador, dict):
        values.append(nomenclador.get("descripcion"))
    return unique_join(values)


def parse_observations(detail: Dict[str, Any]) -> str:
    lines = []
    certificate = clean_text(detail.get("numeroInscripcion"))
    senasa_id = clean_text(detail.get("id"))
    estado = detail.get("estadoProducto")
    estado_text = ""
    if isinstance(estado, dict):
        estado_text = clean_text(estado.get("descripcion"))

    if certificate:
        lines.append(f"Certificado SENASA: {certificate}")
    if senasa_id:
        lines.append(f"ID SENASA: {senasa_id}")
    if estado_text:
        lines.append(f"Estado SENASA: {estado_text}")

    for label, key in [
        ("Observaciones SENASA", "observaciones"),
        ("Restricción preordeñe", "restriccionPreordenie"),
        ("Restricción huevos", "restriccionHuevos"),
        ("Restricción miel", "restriccionMiel"),
        ("Fecha inscripción", "fechaInscripcion"),
        ("Fecha validez", "fechaValidez"),
    ]:
        value = clean_text(detail.get(key))
        if value:
            lines.append(f"{label}: {value}")

    return "\n".join(lines)


def detail_to_record(detail: Dict[str, Any]) -> Optional[SenasaRecord]:
    active, concentration = parse_active_ingredients(detail)
    brand = parse_brand(detail)

    if not active or not brand:
        return None

    species, dog_dose, cat_dose = parse_species_and_dose(detail)

    return SenasaRecord(
        active_ingredient=active,
        brand_name=brand,
        laboratory=parse_laboratory(detail),
        presentation=parse_presentation(detail),
        concentration=concentration,
        species=species,
        category=parse_category(detail),
        route=parse_routes(detail),
        frequency="",
        indications=clean_text(detail.get("indicacionesYVias")),
        dog_dose=dog_dose,
        cat_dose=cat_dose,
        observations=parse_observations(detail),
        senasa_id=clean_text(detail.get("id")),
        certificate=clean_text(detail.get("numeroInscripcion")),
    )


def get_existing_active(db: Session, active_name: str) -> Optional[Dict[str, Any]]:
    return db.execute(
        text("""
            SELECT *
            FROM vademecum_active_ingredients
            WHERE active = TRUE
            AND lower(name) = lower(:name)
            LIMIT 1
        """),
        {"name": active_name},
    ).mappings().first()


def create_active(db: Session, record: SenasaRecord) -> int:
    row = db.execute(
        text("""
            INSERT INTO vademecum_active_ingredients (
                name, category, species, dog_dose, cat_dose, route,
                frequency, indications, contraindications, interactions,
                warnings, observations, active
            )
            VALUES (
                :name, :category, :species, :dog_dose, :cat_dose, :route,
                :frequency, :indications, '', '', '', :observations, TRUE
            )
            RETURNING id
        """),
        {
            "name": record.active_ingredient,
            "category": record.category,
            "species": record.species,
            "dog_dose": record.dog_dose,
            "cat_dose": record.cat_dose,
            "route": record.route,
            "frequency": record.frequency,
            "indications": record.indications,
            "observations": record.observations,
        },
    ).mappings().first()
    return int(row["id"])


def update_active_preserving_clinical(db: Session, active_id: int, existing: Dict[str, Any], record: SenasaRecord) -> bool:
    updates = {}

    for column, value in [
        ("category", record.category),
        ("species", record.species),
        ("route", record.route),
        ("frequency", record.frequency),
        ("indications", record.indications),
    ]:
        if value and not clean_text(existing.get(column)):
            updates[column] = value

    for column, value in [
        ("dog_dose", record.dog_dose),
        ("cat_dose", record.cat_dose),
        ("observations", record.observations),
    ]:
        if value and not clean_text(existing.get(column)):
            updates[column] = value

    if not updates:
        return False

    set_sql = ", ".join(f"{column} = :{column}" for column in updates)
    params = dict(updates)
    params["id"] = active_id

    db.execute(
        text(f"""
            UPDATE vademecum_active_ingredients
            SET {set_sql}
            WHERE id = :id
        """),
        params,
    )
    return True
def find_brand(db: Session, active_id: int, brand_name: str, laboratory: str) -> Optional[Dict[str, Any]]:
    return db.execute(
        text("""
            SELECT *
            FROM vademecum_brands
            WHERE active = TRUE
            AND active_ingredient_id = :active_id
            AND lower(brand_name) = lower(:brand_name)
            AND (
                :laboratory = ''
                OR laboratory = ''
                OR lower(laboratory) = lower(:laboratory)
            )
            LIMIT 1
        """),
        {
            "active_id": active_id,
            "brand_name": brand_name,
            "laboratory": laboratory,
        },
    ).mappings().first()


def upsert_brand(db: Session, active_id: int, record: SenasaRecord) -> str:
    existing = find_brand(db, active_id, record.brand_name, record.laboratory)

    if not existing:
        db.execute(
            text("""
                INSERT INTO vademecum_brands (
                    active_ingredient_id, brand_name, laboratory,
                    presentation, concentration, active
                )
                VALUES (
                    :active_id, :brand_name, :laboratory,
                    :presentation, :concentration, TRUE
                )
            """),
            {
                "active_id": active_id,
                "brand_name": record.brand_name,
                "laboratory": record.laboratory,
                "presentation": record.presentation,
                "concentration": record.concentration,
            },
        )
        return "new"

    changed = False
    updates = {}
    for column, value in [
        ("laboratory", record.laboratory),
        ("presentation", record.presentation),
        ("concentration", record.concentration),
    ]:
        if value and clean_text(existing.get(column)) != value:
            updates[column] = value
            changed = True

    if changed:
        params = dict(updates)
        params["id"] = existing["id"]
        set_sql = ", ".join(f"{column} = :{column}" for column in updates)
        db.execute(
            text(f"""
                UPDATE vademecum_brands
                SET {set_sql}
                WHERE id = :id
            """),
            params,
        )
        return "updated"

    return "same"


def save_record(db: Session, record: SenasaRecord, summary: SenasaSummary) -> None:
    existing = get_existing_active(db, record.active_ingredient)

    if existing:
        active_id = int(existing["id"])
        if update_active_preserving_clinical(db, active_id, existing, record):
            summary.updated_active += 1
    else:
        active_id = create_active(db, record)
        summary.new_active += 1

    brand_result = upsert_brand(db, active_id, record)
    if brand_result == "new":
        summary.new_brands += 1
    elif brand_result == "updated":
        summary.updated_brands += 1


def update_from_senasa(
    db: Session,
    limit: int = 0,
    max_pages: int = 0,
    page_size: int = DEFAULT_PAGE_SIZE,
    sleep_seconds: float = DEFAULT_SLEEP_SECONDS,
) -> Dict[str, Any]:
    """
    Sincroniza Vademécum desde SENASA.

    limit=0 importa todo.
    Para prueba rápida: limit=10.
    """
    summary = SenasaSummary()

    index_rows = fetch_index(max_pages=max_pages, page_size=page_size)
    summary.total_index = len(index_rows)

    if limit and limit > 0:
        index_rows = index_rows[:limit]

    for position, row in enumerate(index_rows, start=1):
        href = get_product_href(row)
        if not href:
            summary.skipped += 1
            if len(summary.errors) < 30:
                summary.errors.append(f"Fila {position}: no se pudo obtener href/id de producto.")
            continue

        try:
            detail = fetch_product_detail(href)
            record = detail_to_record(detail)

            if not record:
                summary.skipped += 1
                if len(summary.errors) < 30:
                    summary.errors.append(f"Producto {href}: sin principio activo o marca comercial.")
                continue

            save_record(db, record, summary)
            summary.details_ok += 1

            if sleep_seconds:
                time.sleep(sleep_seconds)

        except Exception as exc:
            db.rollback()
            summary.details_error += 1
            if len(summary.errors) < 30:
                summary.errors.append(f"Producto {href}: {exc}")
            continue

    db.commit()
    return summary.as_dict()
