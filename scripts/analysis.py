from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent
COMBINED = BASE / "cleaned_data" / "Quick_Bite_Combined_Data.xlsx"
OUT = BASE / "analysis" / "Quick_Bite_Analysis.xlsx"
LATE_AFTER_MINUTES = 45
DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
def delivery_stats(g):
    """Delivery-quality numbers for a group of delivered orders.
    Late % is measured only on orders that have a delivery time recorded."""
    timed = g["delivery_minutes"].dropna()
    late = (timed > LATE_AFTER_MINUTES).sum()
    return pd.Series({
        "avg_delivery_minutes": round(timed.mean(), 1),
        "late_deliveries": int(late),
        "late_pct": round(100 * late / len(timed), 1) if len(timed) else None,
        "avg_customer_rating": round(g["rating"].mean(), 2),
    })
def area_performance(d):
    d = d.dropna(subset=["area_name"])  
    rows = []
    for (area, city), g in d.groupby(["area_name", "city"]):
        rows.append({"area_name": area, "city": city,
                     "delivered_orders": len(g),
                     "total_revenue": round(g["order_amount"].sum(), 2),
                     "avg_order_value": round(g["order_amount"].mean(), 2),
                     **delivery_stats(g).to_dict()})
    out = pd.DataFrame(rows).sort_values("total_revenue", ascending=False).reset_index(drop=True)
    out.insert(0, "revenue_rank", out.index + 1)
    out.insert(out.columns.get_loc("total_revenue") + 1, "revenue_share_pct",
               (100 * out["total_revenue"] / out["total_revenue"].sum()).round(1))
    overall_late = 100 * (d["delivery_minutes"] > LATE_AFTER_MINUTES).sum() / d["delivery_minutes"].notna().sum()
    top_half = out["revenue_rank"] <= len(out) / 2
    out["service_flag"] = "OK"
    out.loc[out["late_pct"] > overall_late, "service_flag"] = "Late rate above average"
    out.loc[top_half & (out["late_pct"] > overall_late), "service_flag"] = "Top-revenue area, poor delivery"
    return out
def top_cuisine_by_day(d):
    rows = []
    for day in DAY_ORDER:
        g = d[d["order_day"] == day]
        counts = g.groupby("cuisine").agg(orders=("order_id", "count"),
                                          revenue=("order_amount", "sum"))
        counts = counts.sort_values(["orders", "revenue"], ascending=False)
        top = counts.iloc[0]
        tied = (counts["orders"] == top["orders"]).sum() > 1
        rows.append({"order_day": day,
                     "top_cuisine": counts.index[0],
                     "orders": int(top["orders"]),
                     "total_orders_that_day": len(g),
                     "share_of_day_pct": round(100 * top["orders"] / len(g), 1),
                     "runner_up": counts.index[1],
                     "runner_up_orders": int(counts.iloc[1]["orders"]),
                     "tie_for_first": "Yes (broken by revenue)" if tied else "No"})
    return pd.DataFrame(rows)

def vehicle_performance(d):
    d = d.dropna(subset=["vehicle_type"])
    rows = []
    for veh, g in d.groupby("vehicle_type"):
        timed = g["delivery_minutes"].dropna()
        rows.append({"vehicle_type": veh,
                     "delivered_orders": len(g),
                     "orders_with_time": len(timed),
                     "avg_delivery_minutes": round(timed.mean(), 1),
                     "median_delivery_minutes": round(timed.median(), 1),
                     "late_deliveries": int((timed > LATE_AFTER_MINUTES).sum()),
                     "late_pct": round(100 * (timed > LATE_AFTER_MINUTES).mean(), 1),
                     "on_time_pct": round(100 * (timed <= LATE_AFTER_MINUTES).mean(), 1),
                     "avg_customer_rating": round(g["rating"].mean(), 2)})
    out = pd.DataFrame(rows).sort_values(["late_pct", "avg_delivery_minutes"]).reset_index(drop=True)
    out.insert(0, "rank_best_to_worst", out.index + 1)  
    return out

def save(sheets):
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                w = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(w + 2, 40)

def main():
    df = pd.read_excel(COMBINED)
    d = df[df["order_status"] == "Delivered"].copy()
    print(f"delivered orders used: {len(d)} of {len(df)}")
    sheets = {"Area_Performance": area_performance(d),
              "Top_Cuisine_By_Day": top_cuisine_by_day(d),
              "Vehicle_Performance": vehicle_performance(d)}
    save(sheets)
    for k, v in sheets.items():
        print(f"\n== {k} ==\n{v.to_string(index=False)}")

if __name__ == "__main__":
    main()
