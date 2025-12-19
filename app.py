from llm_client import ask_llm
from executor import execute_response
from excel_engine import ExcelEngine
import os
import pandas as pd

# ---------------- CONFIG ----------------
EXCEL_FILE = "Plumbing_RR_sample.csv"
OUTPUT_FILE = "exported_result.xlsx"
# ----------------------------------------

# Load file once to get schema
engine = ExcelEngine(EXCEL_FILE)
columns = list(engine.df.columns)

# Session memory
last_result_df = None
last_meta = None

print("\n📊 Excel AI Assistant (type 'exit' to quit)\n")

while True:
    user_input = input("You: ").strip()

    if user_input.lower() == "exit":
        break

    # ---------- SHOW (READ-ONLY) ----------
    if user_input.lower().startswith("show"):
        if last_result_df is None:
            print("❌ Nothing to show yet. Run a query first.\n")
        else:
            print("\nBot (preview of last result):\n")
            print(last_result_df.head(10).to_string(index=False))
            print()
        continue

    # ---------- EXPORT (WRITE-ONLY) ----------
    if user_input.lower().startswith("export"):
        if last_result_df is None or last_meta is None:
            print("❌ Nothing to export yet. Run a query first.\n")
            continue

        os.makedirs("output", exist_ok=True)
        path = os.path.join("output", OUTPUT_FILE)

        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                # -------- Summary Sheet --------
                summary_df = pd.DataFrame({
                    "Metric": [
                        "Filter Applied",
                        "Grouped By",
                        "Total Records",
                        "Total Unique Groups"
                    ],
                    "Value": [
                        last_meta["filter"],
                        last_meta["grouped_by"],
                        last_meta["total_rows"],
                        last_meta["total_groups"]
                    ]
                })

                summary_df.to_excel(
                    writer,
                    sheet_name="Summary",
                    index=False
                )

                # -------- Data Sheet --------
                last_result_df.to_excel(
                    writer,
                    sheet_name="Data",
                    index=False
                )

            print(f"\nBot: Exported result to {path}\n")

        except PermissionError:
            print(
                "\n❌ Cannot write file.\n"
                f"The file '{OUTPUT_FILE}' is currently open in Excel.\n"
                "👉 Please close the file and type 'export' again.\n"
            )

        except Exception as e:
            print(f"\n❌ Export failed: {e}\n")

        continue

    # ---------- ANALYSIS (READ-ONLY) ----------
    try:
        response = ask_llm(user_input, columns)
        result = execute_response(response, EXCEL_FILE)

        if isinstance(result, dict) and result.get("type") == "analysis_result":
            last_result_df = result["dataframe"]
            last_meta = result["meta"]

            print("\n📊 Analysis Summary")
            print(f"- Filter applied     : {last_meta['filter']}")
            print(f"- Grouped by         : {last_meta['grouped_by']}")
            print(f"- Total records      : {last_meta['total_rows']}")
            print(f"- Total unique groups: {last_meta['total_groups']}\n")

            print("Top 5 results:")
            print(result["preview"].to_string(index=False))

            print(
                "\nNote:",
                result["message"]
            )
            print("👉 Use `show` to preview or `export` to save.\n")

        else:
            print("\nBot:", result, "\n")

    except Exception as e:
        # Keep chatbot alive no matter what
        print("❌ Error:", e, "\n")
