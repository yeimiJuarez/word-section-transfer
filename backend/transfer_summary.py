def count_list_groups(list_specs):
    count = 0
    previous_key = None

    for spec in list_specs:
        if spec is None:
            previous_key = None
            continue

        key = spec["source_num_id"]

        if key != previous_key:
            count += 1

        previous_key = key

    return count


def summarize_plan(plan):
    summary = {
        "copied_paragraphs": 0,
        "copied_tables": 0,
        "copied_images": 0,
        "created_lists": 0,
    }

    previous_list_key = None

    for item in plan:
        if "table" in item:
            table_plan = item["table"]

            summary["copied_tables"] += 1

            for row in table_plan["rows"]:
                for cell in row:
                    summary["copied_paragraphs"] += len(cell)

                    summary["copied_images"] += sum(
                        1
                        for paragraph_content in cell
                        for content in paragraph_content
                        if "image" in content
                    )

            for list_specs in table_plan.get(
                "cell_lists",
                {},
            ).values():
                summary["created_lists"] += count_list_groups(
                    list_specs
                )

            previous_list_key = None
            continue

        summary["copied_paragraphs"] += 1

        summary["copied_images"] += sum(
            1
            for content in item["runs"]
            if "image" in content
        )

        list_spec = item.get("list")

        if list_spec is None:
            previous_list_key = None
            continue

        list_key = list_spec["source_num_id"]

        if list_key != previous_list_key:
            summary["created_lists"] += 1

        previous_list_key = list_key

    return summary


def summarize_batch(prepared):
    replacements = []

    totals = {
        "copied_paragraphs": 0,
        "copied_tables": 0,
        "copied_images": 0,
        "created_lists": 0,
    }

    for item in prepared:
        counts = summarize_plan(item["plan"])

        replacements.append(
            {
                "source_title": item["source"]["title"],
                "target_title": item["target"]["title"],
                **counts,
            }
        )

        for name in totals:
            totals[name] += counts[name]

    return {
        "replacement_count": len(prepared),
        **totals,
        "replacements": replacements,
    }