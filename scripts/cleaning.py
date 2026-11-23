from pathlib import Path

import pandas as pd
BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "data" / "QuickBite_Assessment.xlsx"
CLEAN_OUT = BASE / "cleaned_data" / "Quick_Bite_Data_Cleaning.xlsx"
COMBINED_OUT = BASE / "cleaned_data" / "Quick_Bite_Combined_Data.xlsx"
LATE_AFTER_MINUTES = 45 

log = []  
def note(table, msg):
    log.append(f"[{table}] {msg}")
def read_all():
    return pd.read_excel(RAW, sheet_name=None, dtype=str)
def strip_all(df):
    """Trim whitespace in every text column; empty strings become real blanks."""
    for c in df.columns:
        df[c] = df[c].str.strip().replace("", pd.NA)
    return df

def clean_id(s):
    """' r121' / 'c0045 ' -> 'R121' / 'C0045'"""
    return s.str.strip().str.upper()

def clean_name(s):
    return s.str.strip().str.replace(r"\s+", " ", regex=True).str.title()

def clean_phone(s):
    """Keep the last 10 digits (drops +91, 91, leading 0, dashes, spaces)."""
    return s.str.replace(r"\D", "", regex=True).str[-10:]

def map_labels(s, mapping, table, col):
    """Standardise a category column using a lower-case lookup."""
    key = s.str.strip().str.lower()
    out = key.map(mapping)
    unknown = key[out.isna() & key.notna()].unique()
    if len(unknown):
        raise ValueError(f"{table}.{col}: unmapped values {unknown}")
    return out

def parse_date(s, formats):
    """Try each known format in turn; anything still unparsed is reported."""
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for fmt in formats:
        todo = out.isna() & s.notna()
        out[todo] = pd.to_datetime(s[todo], format=fmt, errors="coerce")
    bad = s[out.isna() & s.notna()]
    if len(bad):
        raise ValueError(f"Unparsed dates: {bad.unique()[:5]}")
    return out

def drop_dupes(df, table, subset=None):
    before = len(df)
    df = df.drop_duplicates(subset=subset).reset_index(drop=True)
    note(table, f"removed {before - len(df)} duplicate rows")
    return df

def drop_blank(df, table, cols):
    before = len(df)
    df = df.dropna(subset=cols).reset_index(drop=True)
    note(table, f"removed {before - len(df)} rows with blank mandatory {cols}")
    return df

CITY = {"bengaluru": "Bengaluru", "blr": "Bengaluru",
        "hyderabad": "Hyderabad", "hyd": "Hyderabad",
        "chennai": "Chennai", "madras": "Chennai",
        "mumbai": "Mumbai"}

CUISINE = {"chinese": "Chinese",
           "biryani": "Biryani", "biriyani": "Biryani",
           "burgers": "Burgers",
           "desserts": "Desserts", "dessert": "Desserts",
           "south indian": "South Indian", "south-indian": "South Indian",
           "north indian": "North Indian", "n. indian": "North Indian",
           "north-indian": "North Indian",
           "pizza": "Pizza", "italian - pizza": "Pizza"}

VEHICLE = {"bicycle": "Bicycle", "cycle": "Bicycle",
           "bike": "Bike", "motor bike": "Bike", "motorbike": "Bike",
           "scooter": "Scooter", "scooty": "Scooter", "activa": "Scooter",
           "ev scooter": "EV Scooter", "ev": "EV Scooter",
           "e-scooter": "EV Scooter"}

PAYMENT = {"upi": "UPI", "gpay": "UPI", "phonepe": "UPI",
           "card": "Card", "debit card": "Card", "credit card": "Card",
           "cash": "Cash", "cod": "Cash", "cash on delivery": "Cash",
           "wallet": "Wallet", "paytm wallet": "Wallet", "amazon pay": "Wallet"}

STATUS = {"delivered": "Delivered", "completed": "Delivered",
          "cancelled": "Cancelled", "canceled": "Cancelled"}

def clean_areas(df):
    t = "areas"
    df = strip_all(df)
    df["area_id"] = clean_id(df["area_id"])
    df["area_name"] = df["area_name"].str.replace(r"\s+", " ", regex=True)
    df["city"] = map_labels(df["city"], CITY, t, "city")
    df["zone"] = df["zone"].str.title()
    df = drop_blank(df, t, ["area_id", "area_name", "city"])
    return drop_dupes(df, t)

def clean_restaurants(df):
    t = "restaurants"
    df = strip_all(df)
    df["restaurant_id"] = clean_id(df["restaurant_id"])
    df["area_id"] = clean_id(df["area_id"])
    df["restaurant_name"] = clean_name(df["restaurant_name"])
    df["cuisine"] = map_labels(df["cuisine"], CUISINE, t, "cuisine")
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce") 
    df["onboard_date"] = parse_date(df["onboard_date"], ["%Y-%m-%d", "%d-%m-%Y"]).dt.date
    df = drop_blank(df, t, ["restaurant_id", "restaurant_name", "cuisine"])
    return drop_dupes(df, t, subset=["restaurant_id"])

def clean_partners(df):
    t = "delivery_partners"
    df = strip_all(df)
    df["partner_id"] = clean_id(df["partner_id"])
    df["partner_name"] = clean_name(df["partner_name"])
    df["phone"] = clean_phone(df["phone"])
    df["vehicle_type"] = map_labels(df["vehicle_type"], VEHICLE, t, "vehicle_type")
    df["join_date"] = parse_date(df["join_date"], ["%Y-%m-%d", "%d/%m/%Y"]).dt.date
    df = drop_blank(df, t, ["partner_id", "vehicle_type"])
    return drop_dupes(df, t, subset=["partner_id"])

def clean_customers(df):
    t = "customers"
    df = strip_all(df)
    df["customer_id"] = clean_id(df["customer_id"])
    df["customer_name"] = clean_name(df["customer_name"])
    df["phone"] = clean_phone(df["phone"])
    df["email"] = df["email"].str.lower()
    df["area_id"] = clean_id(df["area_id"])
    df["signup_date"] = parse_date(df["signup_date"],
                                   ["%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y"]).dt.date
    df = drop_blank(df, t, ["customer_id", "customer_name"])
    note(t, f"kept {df['email'].isna().sum()} blank emails and "
            f"{df['area_id'].isna().sum()} blank area_ids (not mandatory)")
    return drop_dupes(df, t, subset=["customer_id"])
def clean_orders(df):
    t = "orders"
    df = strip_all(df)
    for c in ["order_id", "customer_id", "restaurant_id", "partner_id"]:
        df[c] = clean_id(df[c])
    df["order_status"] = map_labels(df["order_status"], STATUS, t, "order_status")
    df["payment_mode"] = map_labels(df["payment_mode"], PAYMENT, t, "payment_mode")
    df["coupon_code"] = df["coupon_code"].str.upper()
    df["order_amount"] = pd.to_numeric(
        df["order_amount"].str.replace(r"[^0-9.]", "", regex=True), errors="coerce")
    df["delivery_minutes"] = pd.to_numeric(
        df["delivery_minutes"].str.extract(r"(\d+(?:\.\d+)?)")[0], errors="coerce")
    df["order_datetime"] = parse_date(
        df["order_datetime"],
        ["%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M", "%d %b %Y %I:%M %p"])

    df = drop_dupes(df, t, subset=["order_id"])
    df = drop_blank(df, t, ["order_id", "customer_id", "restaurant_id",
                            "order_amount", "order_datetime"])
    return df

def clean_feedback(df):
    t = "feedback"
    df = strip_all(df)
    df["feedback_id"] = clean_id(df["feedback_id"])
    df["order_id"] = clean_id(df["order_id"])
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce")
    df["comment"] = df["comment"].str.strip()
    df["feedback_date"] = parse_date(df["feedback_date"], ["%Y-%m-%d", "%d/%m/%Y"]).dt.date
    df = drop_dupes(df, t)
    df = drop_blank(df, t, ["feedback_id", "order_id", "rating"])
    return drop_dupes(df, t, subset=["feedback_id"])

def combine(areas, restaurants, partners, customers, orders, feedback):
    n = len(orders)
    orders = orders[orders.customer_id.isin(customers.customer_id)
                    & orders.restaurant_id.isin(restaurants.restaurant_id)
                    & (orders.partner_id.isna() | orders.partner_id.isin(partners.partner_id))]
    note("combined", f"dropped {n - len(orders)} orders pointing to unknown customer/restaurant/partner")

    df = (orders
          .merge(customers[["customer_id", "customer_name", "area_id"]], on="customer_id", how="left")
          .merge(areas[["area_id", "area_name", "city"]], on="area_id", how="left")
          .merge(restaurants[["restaurant_id", "restaurant_name", "cuisine"]], on="restaurant_id", how="left")
          .merge(partners[["partner_id", "partner_name", "vehicle_type"]], on="partner_id", how="left")
          .merge(feedback[["order_id", "rating", "comment"]], on="order_id", how="left"))

    dt = df["order_datetime"]
    df["order_date"] = dt.dt.date
    df["order_time"] = dt.dt.strftime("%H:%M")
    df["order_day"] = dt.dt.day_name()
    df["is_late"] = (df["delivery_minutes"] > LATE_AFTER_MINUTES).astype("boolean")
    df.loc[df["delivery_minutes"].isna(), "is_late"] = pd.NA

    cols = ["order_id", "order_date", "order_time", "order_day", "customer_name",
            "restaurant_name", "cuisine", "area_name", "city", "partner_name",
            "vehicle_type", "order_amount", "payment_mode", "coupon_code",
            "order_status", "delivery_minutes", "is_late", "rating", "comment"]
    df = df[cols].sort_values("order_id").reset_index(drop=True)
    assert df["order_id"].is_unique, "combined data must have one row per order"
    return df

def save_xlsx(sheets, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, d in sheets.items():
            d.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:  
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 45)

def main():
    raw = read_all()
    for k, v in raw.items():
        note(k, f"raw rows: {len(v)}")

    areas = clean_areas(raw["areas"])
    restaurants = clean_restaurants(raw["restaurants"])
    partners = clean_partners(raw["delivery_partners"])
    customers = clean_customers(raw["customers"])
    orders = clean_orders(raw["orders"])
    feedback = clean_feedback(raw["feedback"])

    cleaned = {"areas": areas, "restaurants": restaurants,
               "delivery_partners": partners, "customers": customers,
               "orders": orders, "feedback": feedback}
    save_xlsx(cleaned, CLEAN_OUT)
    combined = combine(areas, restaurants, partners, customers, orders, feedback)
    save_xlsx({"Combined_Data": combined}, COMBINED_OUT)

    for k, v in cleaned.items():
        note(k, f"clean rows: {len(v)}")
    note("combined", f"rows: {len(combined)}")
    print("\n".join(log))

if __name__ == "__main__":
    main()
