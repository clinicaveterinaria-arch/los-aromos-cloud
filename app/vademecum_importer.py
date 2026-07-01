from sqlalchemy import text


def import_vademecum(db, records):
    summary = {
        "new_active": 0,
        "updated_active": 0,
        "new_brands": 0,
        "updated_brands": 0,
        "errors": []
    }

    for record in records:

        try:

            active = db.execute(
                text("""
                    SELECT id
                    FROM vademecum_active_ingredients
                    WHERE LOWER(name)=LOWER(:name)
                    LIMIT 1
                """),
                {
                    "name": record["active_name"]
                }
            ).mappings().first()

            if active:

                active_id = active["id"]

                db.execute(
                    text("""
                        UPDATE vademecum_active_ingredients
                        SET
                            category=:category,
                            species=:species,
                            route=:route,
                            indications=:indications
                        WHERE id=:id
                    """),
                    {
                        "id": active_id,
                        "category": record["category"],
                        "species": record["species"],
                        "route": record["route"],
                        "indications": record["indications"]
                    }
                )

                summary["updated_active"] += 1

            else:

                active_id = db.execute(
                    text("""
                        INSERT INTO vademecum_active_ingredients
                        (
                            name,
                            category,
                            species,
                            route,
                            indications,
                            active
                        )
                        VALUES
                        (
                            :name,
                            :category,
                            :species,
                            :route,
                            :indications,
                            TRUE
                        )
                        RETURNING id
                    """),
                    {
                        "name": record["active_name"],
                        "category": record["category"],
                        "species": record["species"],
                        "route": record["route"],
                        "indications": record["indications"]
                    }
                ).scalar()

                summary["new_active"] += 1

            brand = db.execute(
                text("""
                    SELECT id
                    FROM vademecum_brands
                    WHERE
                        active_ingredient_id=:active_id
                        AND LOWER(brand_name)=LOWER(:brand)
                    LIMIT 1
                """),
                {
                    "active_id": active_id,
                    "brand": record["brand_name"]
                }
            ).mappings().first()

            if brand:

                db.execute(
                    text("""
                        UPDATE vademecum_brands
                        SET
                            laboratory=:laboratory,
                            presentation=:presentation,
                            concentration=:concentration
                        WHERE id=:id
                    """),
                    {
                        "id": brand["id"],
                        "laboratory": record["laboratory"],
                        "presentation": record["presentation"],
                        "concentration": record["concentration"]
                    }
                )

                summary["updated_brands"] += 1

            else:

                db.execute(
                    text("""
                        INSERT INTO vademecum_brands
                        (
                            active_ingredient_id,
                            brand_name,
                            laboratory,
                            presentation,
                            concentration,
                            active
                        )
                        VALUES
                        (
                            :active_id,
                            :brand,
                            :laboratory,
                            :presentation,
                            :concentration,
                            TRUE
                        )
                    """),
                    {
                        "active_id": active_id,
                        "brand": record["brand_name"],
                        "laboratory": record["laboratory"],
                        "presentation": record["presentation"],
                        "concentration": record["concentration"]
                    }
                )

                summary["new_brands"] += 1

        except Exception as e:

            summary["errors"].append(str(e))

    db.commit()

    return summary
