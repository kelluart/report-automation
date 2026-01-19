from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = BASE_DIR / "data" / "input.csv"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

EXCEL_PATH = OUTPUT_DIR / "report.xlsx"


# -------------------------
# Load
# -------------------------
def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path)


# -------------------------
# Clean (basic + predictable)
# -------------------------
def basic_clean(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    info = {}
    df = df.copy()

    # normalize headers
    df.columns = [c.strip().lower() for c in df.columns]

    # drop duplicates
    before = len(df)
    df = df.drop_duplicates()
    info["duplicates_removed"] = int(before - len(df))

    # best-effort conversions
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce")

    return df, info


# -------------------------
# Organize Clean_Data (sorting)
# -------------------------
def organize_clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    sort_cols = []
    if "category" in df.columns:
        # ensure consistent ordering even with missing categories
        df["category"] = df["category"].astype("string")
        sort_cols.append("category")

    if "date" in df.columns:
        sort_cols.append("date")

    if "amount" in df.columns:
        sort_cols.append("amount")

    if sort_cols:
        df = df.sort_values(by=sort_cols, ascending=True, na_position="last")

    return df


# -------------------------
# Build Summary Tables
# -------------------------
def build_metrics_table(df: pd.DataFrame, info: dict) -> pd.DataFrame:
    rows = []
    rows.append(("rows_total", int(len(df))))
    rows.append(("columns_total", int(len(df.columns))))
    rows.append(("duplicates_removed", int(info.get("duplicates_removed", 0))))

    # nulls overview
    nulls = df.isna().sum()
    rows.append(("null_columns_count", int((nulls > 0).sum())))
    rows.append(("null_cells_total", int(nulls.sum())))

    # amount metrics (if exists)
    if "amount" in df.columns:
        s = df["amount"]
        rows.append(("amount_non_null_count", int(s.notna().sum())))
        rows.append(("amount_sum", float(s.sum(skipna=True))))
        rows.append(("amount_mean", float(s.mean(skipna=True))))
        rows.append(("amount_min", float(s.min(skipna=True))))
        rows.append(("amount_max", float(s.max(skipna=True))))

    return pd.DataFrame(rows, columns=["metric", "value"])


def build_category_totals(df: pd.DataFrame) -> pd.DataFrame | None:
    if "category" not in df.columns:
        return None

    out = df.copy()
    out["category"] = out["category"].fillna("NULL").astype("string")

    # count always exists; sum only if amount exists
    if "amount" in out.columns:
        cat = (
            out.groupby("category", dropna=False)
            .agg(
                transactions=("category", "count"),
                total_amount=("amount", "sum"),
            )
            .reset_index()
            .sort_values(by=["total_amount", "transactions"], ascending=False)
        )
    else:
        cat = (
            out.groupby("category", dropna=False)
            .size()
            .reset_index(name="transactions")
            .sort_values(by=["transactions"], ascending=False)
        )
        cat["total_amount"] = ""

    return cat


def build_month_totals(df: pd.DataFrame) -> pd.DataFrame | None:
    if "date" not in df.columns:
        return None

    out = df.dropna(subset=["date"]).copy()
    if len(out) == 0:
        return None

    out["month"] = out["date"].dt.to_period("M").astype(str)

    if "amount" in out.columns:
        m = (
            out.groupby("month")["amount"]
            .sum()
            .reset_index(name="total_amount")
            .sort_values("month")
        )
    else:
        m = (
            out.groupby("month")
            .size()
            .reset_index(name="transactions")
            .sort_values("month")
        )
        m["total_amount"] = ""

    return m


# -------------------------
# Export (2 sheets only)
# -------------------------
def export_excel(clean_df: pd.DataFrame, metrics_df: pd.DataFrame,
                 category_df: pd.DataFrame | None, month_df: pd.DataFrame | None,
                 path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Sheet 1: organized clean data
        clean_df.to_excel(writer, sheet_name="Clean_Data", index=False)

        # Sheet 2: Summary (stack tables with spacing)
        start = 0
        metrics_df.to_excel(writer, sheet_name="Summary", index=False, startrow=start)
        start += len(metrics_df) + 2

        if category_df is not None:
            pd.DataFrame([["Category Totals", ""]], columns=["metric", "value"]).to_excel(
                writer, sheet_name="Summary", index=False, startrow=start, header=False
            )
            start += 1
            category_df.to_excel(writer, sheet_name="Summary", index=False, startrow=start)
            start += len(category_df) + 2

        if month_df is not None:
            pd.DataFrame([["Month Totals", ""]], columns=["metric", "value"]).to_excel(
                writer, sheet_name="Summary", index=False, startrow=start, header=False
            )
            start += 1
            month_df.to_excel(writer, sheet_name="Summary", index=False, startrow=start)

        # Formatting (simple and readable)
        wb = writer.book

        def autofit(ws, max_width=45):
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    val = "" if cell.value is None else str(cell.value)
                    if len(val) > max_len:
                        max_len = len(val)
                ws.column_dimensions[col_letter].width = min(max_len + 2, max_width)

        # Clean_Data formatting
        ws = wb["Clean_Data"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autofit(ws)

        # Date format if present
        headers = [cell.value for cell in ws[1]]
        if "date" in headers:
            date_col_idx = headers.index("date") + 1
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=date_col_idx)
                if cell.value is not None:
                    cell.number_format = "yyyy-mm-dd"

        # Summary formatting
        ws2 = wb["Summary"]
        ws2.freeze_panes = "A2"
        autofit(ws2, max_width=55)


# -------------------------
# Main
# -------------------------
def main():
    print("SCRIPT STARTED")

    df_raw = load_csv(DATA_PATH)
    df_clean, info = basic_clean(df_raw)
    df_clean = organize_clean_data(df_clean)

    metrics_df = build_metrics_table(df_clean, info)
    category_df = build_category_totals(df_clean)
    month_df = build_month_totals(df_clean)

    export_excel(df_clean, metrics_df, category_df, month_df, EXCEL_PATH)

    print("DONE")
    print(f"Excel: {EXCEL_PATH}")


if __name__ == "__main__":
    main()
