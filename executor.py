from excel_engine import ExcelEngine

def execute_response(response, excel_path):
    engine = ExcelEngine(excel_path)

    # ================= ANALYSIS =================
    if response["type"] == "analysis":
        p = response["action"]["params"]

        # -------- Defensive extraction --------
        filter_column = p.get("filter_column")
        group_by = p.get("group_by")
        value = p.get("value")
        operator = p.get("operator", "contains")

        # -------- Smart operator correction --------
        # Free-text + multi-word value should not use equals
        if operator == "equals" and isinstance(value, str) and " " in value:
            operator = "contains"

        # -------- Apply primary filter --------
        if filter_column and value:
            engine.filter(
                column=filter_column,
                operator=operator,
                value=value
            )

        # -------- SMART FALLBACK (KEY FIX) --------
        # If phrase matching returns 0 rows, retry with token matching
        if (
            value
            and isinstance(value, str)
            and " " in value
            and engine.df.empty
        ):
            # Reload original data
            engine = ExcelEngine(excel_path)

            tokens = [t.strip().lower() for t in value.split() if t.strip()]
            series = engine.df[filter_column].astype(str).str.lower()

            mask = False
            for token in tokens:
                mask = mask | series.str.contains(token, na=False)

            engine.df = engine.df[mask]

        # -------- Grouping --------
        grouped = engine.group_count(group_by)

        total_rows = int(grouped["count"].sum())
        total_groups = len(grouped)

        return {
            "type": "analysis_result",
            "dataframe": grouped,   # FULL result (for export)
            "meta": {
                "filter": f"{filter_column} contains '{value}'"
                if value else "No filter applied",
                "grouped_by": group_by,
                "total_rows": total_rows,
                "total_groups": total_groups
            },
            "preview": grouped.sort_values(
                "count", ascending=False
            ).head(5),
            "message": "Analysis complete. No file was created."
        }

    # ================= PLAN (RESERVED) =================
    elif response["type"] == "plan":
        # No write actions here by design
        return engine.df

    else:
        raise ValueError("Unknown response type")
