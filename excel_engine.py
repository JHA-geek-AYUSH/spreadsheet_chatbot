import pandas as pd
import os
import difflib

class ExcelEngine:
    def __init__(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".csv":
            self.df = pd.read_csv(file_path)
        elif ext in [".xlsx", ".xls"]:
            self.df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file type")

    # ---------- COLUMN RESOLUTION ----------
    def resolve_column(self, col):
        if col in self.df.columns:
            return col

        matches = difflib.get_close_matches(col, self.df.columns, n=1, cutoff=0.6)
        if not matches:
            raise KeyError(f"Column '{col}' not found. Available: {list(self.df.columns)}")

        return matches[0]

    # ---------- FILTER ----------
    def filter(self, column, operator, value):
        column = self.resolve_column(column)
        series = self.df[column].astype(str)

        if operator == "contains":
            self.df = self.df[series.str.contains(value, case=False, na=False)]
        elif operator == "equals":
            self.df = self.df[series.str.lower() == str(value).lower()]
        else:
            raise ValueError(f"Unsupported operator: {operator}")

    # ---------- GROUP ----------
    def group_count(self, column):
        column = self.resolve_column(column)
        return self.df.groupby(column).size().reset_index(name="count")
