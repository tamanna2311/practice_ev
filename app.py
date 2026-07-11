import html
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Voltwise — Find your EV",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_FILE = Path(__file__).with_name("ev_data.xlsx")
SHEET_NAME = "EV_Data"

NUMERIC_COLUMNS = [
    "seating_capacity", "price_lakh", "claimed_range_km", "battery_kwh",
    "safety_rating", "dc_fast_charging_time_min", "ac_charging_time_hours",
    "launch_year", "user_rating", "rating_count", "average_monthly_sales",
]


st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root { --ink:#17221b; --muted:#647069; --green:#1f7a4d; --lime:#dff36b; --paper:#f7f7f2; --line:#dfe4dc; }
.stApp { background: var(--paper); color: var(--ink); }
html, body, [class*="css"] { font-family: "DM Sans", sans-serif; }
h1,h2,h3 { font-family:"Manrope",sans-serif !important; letter-spacing:-.035em !important; }
#MainMenu, footer, header { visibility:hidden; }
.stMainBlockContainer { max-width:1180px; padding:1.25rem 2rem 4rem; }
[data-testid="stSidebar"] { display:none; }
.vw-nav { display:flex; justify-content:space-between; align-items:center; padding:.4rem 0 1.5rem; }
.vw-brand { font:800 1.15rem "Manrope"; letter-spacing:-.04em; display:flex; align-items:center; gap:.55rem; }
.vw-mark { width:29px;height:29px;border-radius:9px;background:var(--ink);color:var(--lime);display:grid;place-items:center;font-size:.9rem; }
.vw-navmeta { color:var(--muted); font-size:.82rem; }
.vw-hero { background:var(--ink); color:white; border-radius:26px; padding:3rem 3.2rem; position:relative; overflow:hidden; min-height:310px; }
.vw-hero:after { content:""; position:absolute; width:370px;height:370px;right:-80px;top:-145px;border:1px solid rgba(223,243,107,.4);border-radius:50%;box-shadow:0 0 0 55px rgba(223,243,107,.05),0 0 0 110px rgba(223,243,107,.025); }
.vw-kicker { color:var(--lime); font-weight:700; font-size:.76rem; letter-spacing:.13em; text-transform:uppercase; }
.vw-hero h1 { color:white; font-size:clamp(2.3rem,5vw,4.3rem); line-height:1.01; margin:.75rem 0 1rem; max-width:700px; }
.vw-hero p { color:#bdc8c0; max-width:610px; font-size:1.03rem; line-height:1.65; margin:0; }
.vw-trust { display:flex;gap:1.5rem;margin-top:2rem;color:#d8ded9;font-size:.78rem; }
.vw-trust span:before { content:"✓"; color:var(--lime); margin-right:.45rem;font-weight:800; }
.vw-section { margin:3.2rem 0 1.4rem; }
.vw-eyebrow { color:var(--green);font-weight:700;font-size:.76rem;letter-spacing:.12em;text-transform:uppercase; }
.vw-section h2 { font-size:2rem;margin:.35rem 0 .45rem; }
.vw-section p { color:var(--muted);margin:0; }
[data-testid="stForm"] { background:white;border:1px solid var(--line);border-radius:22px;padding:1rem 1.3rem 1.35rem;box-shadow:0 12px 34px rgba(29,45,34,.055); }
[data-testid="stForm"] h3 { font-size:1rem !important; letter-spacing:-.015em !important; margin:.35rem 0 .15rem; }
[data-testid="stForm"] p { color:var(--muted);font-size:.82rem; }
[data-testid="stSelectbox"] label, [data-testid="stNumberInput"] label, [data-testid="stSlider"] label { font-weight:650;color:var(--ink); }
[data-baseweb="select"] > div, [data-testid="stNumberInput"] input { border-color:var(--line);background:#fbfcf9;border-radius:10px; }
.stButton > button, [data-testid="stFormSubmitButton"] button { border:0;border-radius:11px;background:var(--green);color:white;font-weight:700;min-height:3.05rem;transition:.2s ease; }
.stButton > button:hover, [data-testid="stFormSubmitButton"] button:hover { background:#17643e;color:white;transform:translateY(-1px);box-shadow:0 8px 20px rgba(31,122,77,.18); }
.vw-chip { display:inline-flex;align-items:center;background:#edf4ee;color:#285c3e;padding:.42rem .65rem;border-radius:99px;font-size:.72rem;font-weight:700;margin:.15rem .2rem .15rem 0; }
.vw-card { background:white;border:1px solid var(--line);border-radius:20px;padding:1.35rem 1.45rem;margin:.8rem 0;box-shadow:0 9px 28px rgba(29,45,34,.045); }
.vw-cardtop { display:flex;justify-content:space-between;gap:1rem;align-items:start; }
.vw-rank { color:var(--green);font-weight:800;font-size:.72rem;letter-spacing:.1em;text-transform:uppercase; }
.vw-card h3 { margin:.25rem 0 .1rem;font-size:1.35rem; }
.vw-variant { color:var(--muted);font-size:.83rem; }
.vw-score { background:var(--ink);color:white;border-radius:14px;padding:.65rem .8rem;text-align:center;min-width:73px; }
.vw-score strong { display:block;color:var(--lime);font:800 1.15rem "Manrope"; }
.vw-score small { color:#b9c4bc;font-size:.65rem;text-transform:uppercase;letter-spacing:.06em; }
.vw-specs { display:grid;grid-template-columns:repeat(4,1fr);gap:.55rem;margin:1.2rem 0; }
.vw-spec { background:#f5f7f3;border-radius:11px;padding:.72rem; }
.vw-spec small { display:block;color:var(--muted);font-size:.67rem;margin-bottom:.22rem; }
.vw-spec strong { font-size:.9rem; }
.vw-reason { border-top:1px solid var(--line);padding-top:.9rem;color:#4e5b53;font-size:.82rem; }
.vw-empty { border:1px dashed #c7d0c7;border-radius:18px;padding:2.1rem;text-align:center;color:var(--muted);background:rgba(255,255,255,.45); }
.vw-note { background:#eef2eb;border-radius:14px;padding:1rem 1.15rem;color:#526057;font-size:.79rem;margin-top:1.2rem; }
@media(max-width:700px){.stMainBlockContainer{padding:.8rem 1rem 3rem}.vw-navmeta{display:none}.vw-hero{padding:2rem 1.35rem;min-height:0}.vw-hero h1{font-size:2.35rem}.vw-trust{flex-direction:column;gap:.45rem}.vw-specs{grid-template-columns:repeat(2,1fr)}.vw-card{padding:1.1rem}.vw-score{min-width:65px}.vw-section{margin-top:2.4rem}}
</style>
""", unsafe_allow_html=True)


def minmax_score(series, missing_value=0.0):
    values = pd.to_numeric(series, errors="coerce")
    result = pd.Series(missing_value, index=series.index, dtype=float)
    valid = values.notna()
    if not valid.any():
        return result
    low, high = values[valid].min(), values[valid].max()
    result.loc[valid] = 1.0 if high == low else (values.loc[valid] - low) / (high - low)
    return result.clip(0, 1)


def reverse_minmax_score(series, missing_value=np.nan):
    values = pd.to_numeric(series, errors="coerce")
    result = pd.Series(missing_value, index=series.index, dtype=float)
    valid = values.notna()
    if not valid.any():
        return result
    low, high = values[valid].min(), values[valid].max()
    result.loc[valid] = 1.0 if high == low else (high - values.loc[valid]) / (high - low)
    return result.clip(0, 1)


def is_any(value):
    return value is None or str(value).strip().lower() in {"", "any", "all", "none", "no preference"}


@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    for column in NUMERIC_COLUMNS:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")
    for column in ["brand", "model", "variant", "body_type", "fast_charging_supported", "sales_period", "image_url"]:
        if column in df.columns:
            cleaned = df[column].astype("string").str.strip()
            missing_tokens = cleaned.str.lower().isin({
                "", "na", "n/a", "none", "null", "nan", "-", "--",
                "not publicly available",
            })
            df[column] = cleaned.mask(missing_tokens, pd.NA)
    df["brand"] = df["brand"].fillna("Unknown Brand")
    df["model"] = df["model"].fillna("Unknown Model")
    df["variant"] = df["variant"].fillna("Base / entry configuration")
    df["body_type"] = df["body_type"].fillna("Unknown")
    charging = df["fast_charging_supported"].fillna("").astype("string").str.lower()
    charging_map = {
        "yes": "Yes", "y": "Yes", "true": "Yes", "1": "Yes", "supported": "Yes",
        "no": "No", "n": "No", "false": "No", "0": "No", "not supported": "No",
    }
    df["fast_charging_supported"] = charging.map(charging_map).fillna("Unknown").astype(str)
    df["rating_score"] = np.where(df["user_rating"].notna() & df["rating_count"].fillna(0).gt(0), (df["user_rating"] / 5).clip(0, 1), 0.0)
    max_count = df["rating_count"].fillna(0).clip(lower=0).max()
    df["rating_count_score"] = np.log1p(df["rating_count"].fillna(0).clip(lower=0)) / np.log1p(max_count) if max_count > 0 else 0.0
    max_sales = df["average_monthly_sales"].fillna(0).clip(lower=0).max()
    df["sales_score"] = np.log1p(df["average_monthly_sales"].fillna(0).clip(lower=0)) / np.log1p(max_sales) if max_sales > 0 else 0.0
    df["recency_score"] = minmax_score(df["launch_year"], 0)
    dc_times = df["dc_fast_charging_time_min"].where(df["fast_charging_supported"] == "Yes")
    df["dc_speed_score"] = reverse_minmax_score(dc_times)
    df["popularity_score"] = (df["rating_score"] + df["rating_count_score"] + df["sales_score"] + df["recency_score"]) / 4
    return df


def recommend_evs(data, preferred_brand="Any", max_budget=None, min_range=None,
                  preferred_body_type="Any", passengers=None, min_safety=None,
                  fast_charging="Any", top_n=5, popularity_weight=.25):
    filtered = data.copy()
    if max_budget is not None:
        filtered = filtered[filtered["price_lakh"].notna() & filtered["price_lakh"].le(max_budget)]
    if min_range is not None:
        filtered = filtered[filtered["claimed_range_km"].notna() & filtered["claimed_range_km"].ge(min_range)]
    if passengers is not None:
        filtered = filtered[filtered["seating_capacity"].notna() & filtered["seating_capacity"].ge(passengers)]
    if min_safety is not None:
        filtered = filtered[filtered["safety_rating"].notna() & filtered["safety_rating"].ge(min_safety)]
    if fast_charging in {"Yes", "No"}:
        filtered = filtered[filtered["fast_charging_supported"] == fast_charging]
    if filtered.empty:
        return filtered

    scores = []
    if not is_any(preferred_brand):
        filtered["brand_match_score"] = filtered["brand"].str.lower().eq(str(preferred_brand).lower()).astype(float)
        scores.append("brand_match_score")
    if max_budget is not None:
        filtered["budget_match_score"] = ((max_budget - filtered["price_lakh"]) / max_budget).clip(0, 1)
        scores.append("budget_match_score")
    if min_range is not None:
        filtered["range_match_score"] = minmax_score(filtered["claimed_range_km"], 0)
        scores.append("range_match_score")
    if not is_any(preferred_body_type):
        selected = str(preferred_body_type).lower()
        filtered["body_type_match_score"] = filtered["body_type"].fillna("").str.lower().apply(lambda body: 1.0 if selected in body or body in selected else 0.0)
        scores.append("body_type_match_score")
    if min_safety is not None:
        filtered["safety_match_score"] = (filtered["safety_rating"] / 5).clip(0, 1)
        scores.append("safety_match_score")
    if fast_charging == "Yes":
        filtered["charging_match_score"] = reverse_minmax_score(filtered["dc_fast_charging_time_min"]).fillna(.5)
        scores.append("charging_match_score")
    if scores:
        filtered["preference_score"] = filtered[scores].mean(axis=1)
        filtered["final_score"] = (1 - popularity_weight) * filtered["preference_score"] + popularity_weight * filtered["popularity_score"]
    else:
        filtered["preference_score"] = np.nan
        filtered["final_score"] = filtered["popularity_score"]
    return (filtered.sort_values(["final_score", "popularity_score", "rating_count"], ascending=False, na_position="last")
            .drop_duplicates(["brand", "model"]).head(top_n).reset_index(drop=True))


def safe(value):
    return html.escape(str(value))


def match_reason(row, brand, body, min_range, passengers):
    reasons = []
    if not is_any(brand) and str(row["brand"]).lower() == str(brand).lower():
        reasons.append("your preferred brand")
    if not is_any(body) and str(body).lower() in str(row["body_type"]).lower():
        reasons.append("the body style you chose")
    if min_range and pd.notna(row["claimed_range_km"]):
        reasons.append(f"{int(row['claimed_range_km'] - min_range)} km above your minimum range")
    if passengers and row["seating_capacity"] > passengers:
        reasons.append("extra passenger room")
    if row["fast_charging_supported"] == "Yes":
        reasons.append("DC fast charging")
    return ", ".join(reasons[:3]).capitalize() or "Strong overall popularity, recency, and owner-rating signals"


st.markdown("""
<div class="vw-nav"><div class="vw-brand"><span class="vw-mark">⚡</span> Voltwise</div><div class="vw-navmeta">Independent EV discovery · India</div></div>
<section class="vw-hero"><div class="vw-kicker">The simpler way to shortlist an EV</div><h1>Find the EV that fits your life.</h1><p>Tell us what matters. We’ll rank the Indian market around your budget, range and everyday needs—with a clear reason behind every match.</p><div class="vw-trust"><span>62 EV variants reviewed</span><span>No sponsored rankings</span><span>About 60 seconds</span></div></section>
<div class="vw-section"><div class="vw-eyebrow">Your priorities</div><h2>Build your shortlist</h2><p>Requirements narrow the market. Preferences improve the ranking without hiding good alternatives.</p></div>
""", unsafe_allow_html=True)

try:
    ev_data = load_data()
except Exception as error:
    st.error(f"We couldn't load the EV catalogue. Please check the workbook: {error}")
    st.stop()

with st.form("matcher"):
    left, middle, right = st.columns(3, gap="large")
    with left:
        st.markdown("### 01 · Budget & range")
        st.caption("Your non-negotiables")
        max_budget = st.slider("Maximum budget (₹ lakh)", 5, 100, 25, 1)
        min_range = st.slider("Minimum claimed range (km)", 100, 700, 300, 25)
    with middle:
        st.markdown("### 02 · Everyday fit")
        st.caption("Space for the way you travel")
        passengers = st.selectbox("Minimum seats", [2, 4, 5, 6, 7], index=2)
        body = st.selectbox("Preferred body style", ["Any", *sorted(ev_data["body_type"].dropna().unique())])
    with right:
        st.markdown("### 03 · Final touches")
        st.caption("Helpful ranking signals")
        brand = st.selectbox("Preferred brand", ["Any", *sorted(ev_data["brand"].dropna().unique())])
        safety = st.selectbox("Minimum safety rating", ["Any", 4, 5])
        charging = st.selectbox("DC fast charging", ["Any", "Yes", "No"])
    submitted = st.form_submit_button("Find my best matches →", use_container_width=True)

if submitted:
    min_safety = None if safety == "Any" else float(safety)
    results = recommend_evs(ev_data, brand, max_budget, min_range, body, passengers, min_safety, charging, 5)
    st.markdown('<div class="vw-section"><div class="vw-eyebrow">Your shortlist</div><h2>Best matches for you</h2><p>Ranked by fit, then strengthened by owner ratings, popularity and recency.</p></div>', unsafe_allow_html=True)
    if results.empty:
        st.markdown('<div class="vw-empty"><h3>No exact match—yet.</h3><p>Try increasing your budget, lowering the minimum range, or choosing “Any” for safety and charging.</p></div>', unsafe_allow_html=True)
    else:
        for i, row in results.iterrows():
            safety_text = f"{row['safety_rating']:.0f}/5" if pd.notna(row["safety_rating"]) else "Not rated"
            rating_text = f"{row['user_rating']:.1f}/5" if pd.notna(row["user_rating"]) else "New"
            score = int(round(row["final_score"] * 100))
            reason = match_reason(row, brand, body, min_range, passengers)
            st.markdown(f'''<article class="vw-card"><div class="vw-cardtop"><div><div class="vw-rank">Match {i + 1}</div><h3>{safe(row['brand'])} {safe(row['model'])}</h3><div class="vw-variant">{safe(row['variant'])} · {safe(row['body_type'])}</div></div><div class="vw-score"><strong>{score}%</strong><small>match</small></div></div><div class="vw-specs"><div class="vw-spec"><small>Indicative price</small><strong>₹{row['price_lakh']:.2f} lakh</strong></div><div class="vw-spec"><small>Claimed range</small><strong>{row['claimed_range_km']:.0f} km</strong></div><div class="vw-spec"><small>Battery</small><strong>{row['battery_kwh']:.1f} kWh</strong></div><div class="vw-spec"><small>Seats · safety</small><strong>{int(row['seating_capacity'])} · {safety_text}</strong></div></div><div class="vw-reason"><span class="vw-chip">Why it fits</span> {safe(reason)} · Owner rating {rating_text}.</div></article>''', unsafe_allow_html=True)
        st.markdown('<div class="vw-note"><strong>Good to know:</strong> Prices are indicative ex-showroom figures. Claimed range can differ from real-world range based on speed, weather, load and driving style. Verify the latest variant, price and charging specifications with the manufacturer before buying.</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="vw-empty"><strong>Your matches will appear here.</strong><br>Start with the three essentials: budget, usable range and seats.</div>', unsafe_allow_html=True)
