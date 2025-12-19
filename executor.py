from excel_engine import ExcelEngine

def execute_response(response, excel_path):
    engine = ExcelEngine(excel_path)

    # ================= ANALYSIS =================
    if response["type"] == "analysis":
        p = response["action"]["params"]

        engine.filter(
            column=p["filter_column"],
            operator=p["operator"],
            value=p["value"]
        )

        grouped = engine.group_count(p["group_by"])

        total_rows = int(grouped["count"].sum())
        total_groups = len(grouped)

        return {
            "type": "analysis_result",
            "dataframe": grouped,   # FULL result (for export)
            "meta": {
                "filter": f"{p['filter_column']} {p['operator']} '{p['value']}'",
                "grouped_by": p["group_by"],
                "total_rows": total_rows,
                "total_groups": total_groups
            },
            "preview": grouped.sort_values("count", ascending=False).head(5),
            "message": "Analysis complete. No file was created."
        }

    # ================= PLAN (NOT USED FOR EXPORT) =================
    elif response["type"] == "plan":
        # Reserved for future write actions
        return engine.df

    else:
        raise ValueError("Unknown response type")
