from sqlalchemy import text


def import_vademecum(db, records):
    summary = {
        "new_active": 0,
        "updated_active": 0,
        "new_brands": 0,
        "updated_brands": 0,
        "skipped": 0,
        "errors": []
    }

    for record in records:
        try:
            active_name = (record.get("active_name") or "").strip()
            brand_name = (record.get("brand_name") or "").strip()

            if not active_name:
                summary["skipped"] += 1
                continue

            active = db.execute(
                text("""
                    SELECT id
                    FROM vademecum_active_ingredients
                    WHERE LOWER(name)=LOWER(:name)
                    LIMIT 1
                """),
                {"name": active_name}
            ).mappings().first()

            if active:
                active_id = active["id"]

                db.execute(
                    text("""
                        UPDATE vademecum_active_ingredients
                        SET
                            category = COALESCE(NULLIF(:category, ''), category),
                            species = COALESCE(NULLIF(:species, ''), species),
                            route = COALESCE(NULLIF(:route, ''), route),
                            indications = COALESCE(NULLIF(:indications, ''), indications),
                            active = TRUE
                        WHERE id=:id
                    """),
                    {
                        "id": active_id,
                        "category": record.get("category", ""),
                        "species": record.get("species", ""),
                        "route": record.get("route", ""),
                        "indications": record.get("indications", "")
                    }
                )

                summary["updated_active"] += 1

            else:
                active_id = db.execute(
                    text("""
                        INSERT INTO vademecum_active_ingredients
                        (
                            name, category, species, route, indications, active
                        )
                        VALUES
                        (
                            :name, :category, :species, :route, :indications, TRUE
                        )
                        RETURNING id
                    """),
                    {
                        "name": active_name,
                        "category": record.get("category", ""),
                        "species": record.get("species", ""),
                        "route": record.get("route", ""),
                        "indications": record.get("indications", "")
                    }
                ).scalar()

                summary["new_active"] += 1

            if not brand_name:
                summary["skipped"] += 1
                continue

            brand = db.execute(
                text("""
                    SELECT id
                    FROM vademecum_brands
                    WHERE active_ingredient_id=:active_id
                    AND LOWER(brand_name)=LOWER(:brand)
                    LIMIT 1
                """),
                {
                    "active_id": active_id,
                    "brand": brand_name
                }
            ).mappings().first()

            if brand:
                db.execute(
                    text("""
                        UPDATE vademecum_brands
                        SET
                            laboratory = COALESCE(NULLIF(:laboratory, ''), laboratory),
                            presentation = COALESCE(NULLIF(:presentation, ''), presentation),
                            concentration = COALESCE(NULLIF(:concentration, ''), concentration),
                            active = TRUE
                        WHERE id=:id
                    """),
                    {
                        "id": brand["id"],
                        "laboratory": record.get("laboratory", ""),
                        "presentation": record.get("presentation", ""),
                        "concentration": record.get("concentration", "")
                    }
                )

                summary["updated_brands"] += 1

            else:
                db.execute(
                    text("""
                        INSERT INTO vademecum_brands
                        (
                            active_ingredient_id, brand_name, laboratory,
                            presentation, concentration, active
                        )
                        VALUES
                        (
                            :active_id, :brand, :laboratory,
                            :presentation, :concentration, TRUE
                        )
                    """),
                    {
                        "active_id": active_id,
                        "brand": brand_name,
                        "laboratory": record.get("laboratory", ""),
                        "presentation": record.get("presentation", ""),
                        "concentration": record.get("concentration", "")
                    }
                )

                summary["new_brands"] += 1

        except Exception as e:
            summary["errors"].append(str(e))

    db.commit()
    return summary
