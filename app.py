
import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Indian EV Recommendation System",
    page_icon="🚗",
    layout="wide"
)

DATA_FILE = "ev_data.xlsx"
SHEET_NAME = "EV_Data"


NUMERIC_COLUMNS = [
    "seating_capacity",
    "price_lakh",
    "claimed_range_km",
    "battery_kwh",
    "safety_rating",
    "dc_fast_charging_time_min",
    "ac_charging_time_hours",
    "launch_year",
    "user_rating",
    "rating_count",
    "average_monthly_sales"
]


def minmax_score(series, missing_value=0.0):
    values = pd.to_numeric(series, errors="coerce")

    result = pd.Series(
        missing_value,
        index=series.index,
        dtype=float
    )

    valid = values.notna()

    if not valid.any():
        return result

    minimum = values[valid].min()
    maximum = values[valid].max()

    if maximum == minimum:
        result.loc[valid] = 1.0
    else:
        result.loc[valid] = (
            values.loc[valid] - minimum
        ) / (
            maximum - minimum
        )

    return result.clip(0, 1)


def reverse_minmax_score(series, missing_value=np.nan):
    values = pd.to_numeric(series, errors="coerce")

    result = pd.Series(
        missing_value,
        index=series.index,
        dtype=float
    )

    valid = values.notna()

    if not valid.any():
        return result

    minimum = values[valid].min()
    maximum = values[valid].max()

    if maximum == minimum:
        result.loc[valid] = 1.0
    else:
        result.loc[valid] = (
            maximum - values.loc[valid]
        ) / (
            maximum - minimum
        )

    return result.clip(0, 1)


def is_any(value):
    if value is None:
        return True

    return str(value).strip().lower() in {
        "",
        "any",
        "all",
        "none",
        "no preference"
    }


@st.cache_data
def load_data():
    df = pd.read_excel(
        DATA_FILE,
        sheet_name=SHEET_NAME
    )

    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    missing_values = [
        "",
        " ",
        "NA",
        "N/A",
        "None",
        "null",
        "nan",
        "-",
        "--",
        "Not publicly available"
    ]

    df = df.replace(missing_values, np.nan)

    for column in NUMERIC_COLUMNS:
        if column in df.columns:
            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            )

    text_columns = [
        "brand",
        "model",
        "variant",
        "body_type",
        "fast_charging_supported",
        "sales_period",
        "image_url"
    ]

    for column in text_columns:
        if column in df.columns:
            df[column] = (
                df[column]
                .astype("string")
                .str.strip()
            )

    df["brand"] = df["brand"].fillna("Unknown Brand")
    df["model"] = df["model"].fillna("Unknown Model")
    df["variant"] = df["variant"].fillna(
        "Base / entry configuration"
    )
    df["body_type"] = df["body_type"].fillna("Unknown")

    charging_text = (
        df["fast_charging_supported"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df["fast_charging_supported"] = np.select(
        [
            charging_text.isin([
                "yes",
                "y",
                "true",
                "1",
                "supported"
            ]),
            charging_text.isin([
                "no",
                "n",
                "false",
                "0",
                "not supported"
            ])
        ],
        [
            "Yes",
            "No"
        ],
        default="Unknown"
    )

    df["rating_score"] = np.where(
        df["user_rating"].notna()
        & df["rating_count"].fillna(0).gt(0),
        (df["user_rating"] / 5).clip(0, 1),
        0.0
    )

    maximum_rating_count = (
        df["rating_count"]
        .fillna(0)
        .clip(lower=0)
        .max()
    )

    if maximum_rating_count > 0:
        df["rating_count_score"] = (
            np.log1p(
                df["rating_count"]
                .fillna(0)
                .clip(lower=0)
            )
            / np.log1p(maximum_rating_count)
        )
    else:
        df["rating_count_score"] = 0.0

    maximum_sales = (
        df["average_monthly_sales"]
        .fillna(0)
        .clip(lower=0)
        .max()
    )

    if maximum_sales > 0:
        df["sales_score"] = (
            np.log1p(
                df["average_monthly_sales"]
                .fillna(0)
                .clip(lower=0)
            )
            / np.log1p(maximum_sales)
        )
    else:
        df["sales_score"] = 0.0

    df["recency_score"] = minmax_score(
        df["launch_year"],
        missing_value=0
    )

    dc_times = df[
        "dc_fast_charging_time_min"
    ].where(
        df["fast_charging_supported"] == "Yes"
    )

    df["dc_speed_score"] = reverse_minmax_score(
        dc_times
    )

    df["popularity_score"] = (
        df["rating_score"]
        + df["rating_count_score"]
        + df["sales_score"]
        + df["recency_score"]
    ) / 4

    return df


def recommend_evs(
    data,
    preferred_brand="Any",
    min_budget=None,
    max_budget=None,
    min_range=None,
    preferred_body_type="Any",
    passengers=None,
    min_safety=None,
    fast_charging="Any",
    top_n=5,
    popularity_weight=0.25
):
    filtered = data.copy()

    if min_budget is not None:
        filtered = filtered[
            filtered["price_lakh"].notna()
            & filtered["price_lakh"].ge(min_budget)
        ]

    if max_budget is not None:
        filtered = filtered[
            filtered["price_lakh"].notna()
            & filtered["price_lakh"].le(max_budget)
        ]

    if min_range is not None:
        filtered = filtered[
            filtered["claimed_range_km"].notna()
            & filtered["claimed_range_km"].ge(min_range)
        ]

    if passengers is not None:
        filtered = filtered[
            filtered["seating_capacity"].notna()
            & filtered["seating_capacity"].ge(passengers)
        ]

    if min_safety is not None:
        filtered = filtered[
            filtered["safety_rating"].notna()
            & filtered["safety_rating"].ge(min_safety)
        ]

    if fast_charging == "Yes":
        filtered = filtered[
            filtered["fast_charging_supported"] == "Yes"
        ]

    if filtered.empty:
        return filtered

    selected_scores = []

    if not is_any(preferred_brand):
        filtered["brand_match_score"] = (
            filtered["brand"]
            .str.lower()
            .eq(preferred_brand.lower())
            .astype(float)
        )

        selected_scores.append("brand_match_score")

    if min_budget is not None or max_budget is not None:
        if (
            min_budget is not None
            and max_budget is not None
            and max_budget > min_budget
        ):
            filtered["budget_match_score"] = (
                (max_budget - filtered["price_lakh"])
                / (max_budget - min_budget)
            ).clip(0, 1)

        elif max_budget is not None and max_budget > 0:
            filtered["budget_match_score"] = (
                (max_budget - filtered["price_lakh"])
                / max_budget
            ).clip(0, 1)

        else:
            filtered["budget_match_score"] = 1.0

        selected_scores.append("budget_match_score")

    if min_range is not None:
        filtered["range_match_score"] = minmax_score(
            filtered["claimed_range_km"],
            missing_value=0
        )

        selected_scores.append("range_match_score")

    if not is_any(preferred_body_type):
        selected_body = preferred_body_type.lower()

        filtered["body_type_match_score"] = (
            filtered["body_type"]
            .fillna("")
            .str.lower()
            .apply(
                lambda body: 1.0
                if selected_body in body
                or body in selected_body
                else 0.0
            )
        )

        selected_scores.append("body_type_match_score")

    if min_safety is not None:
        filtered["safety_match_score"] = (
            filtered["safety_rating"] / 5
        ).clip(0, 1)

        selected_scores.append("safety_match_score")

    if fast_charging == "Yes":
        charging_score = reverse_minmax_score(
            filtered["dc_fast_charging_time_min"],
            missing_value=np.nan
        )

        filtered["charging_match_score"] = (
            charging_score.fillna(0.5)
        )

        selected_scores.append("charging_match_score")

    if selected_scores:
        filtered["preference_score"] = (
            filtered[selected_scores]
            .mean(axis=1)
        )

        filtered["final_score"] = (
            (1 - popularity_weight)
            * filtered["preference_score"]
            + popularity_weight
            * filtered["popularity_score"]
        )
    else:
        filtered["preference_score"] = np.nan
        filtered["final_score"] = (
            filtered["popularity_score"]
        )

    filtered = filtered.sort_values(
        by=[
            "final_score",
            "popularity_score",
            "rating_count"
        ],
        ascending=[
            False,
            False,
            False
        ],
        na_position="last"
    )

    filtered = filtered.drop_duplicates(
        subset=["brand", "model"],
        keep="first"
    )

    return filtered.head(top_n).reset_index(drop=True)


st.title("Indian EV Recommendation System")

try:
    ev_data = load_data()
except Exception as error:
    st.error(f"Dataset could not be loaded: {error}")
    st.stop()


with st.sidebar:
    st.header("Select your preferences")

    brand_options = [
        "Any",
        *sorted(ev_data["brand"].dropna().unique())
    ]

    preferred_brand = st.selectbox(
        "Preferred brand",
        brand_options
    )

    use_budget = st.checkbox("Select budget range")

    min_budget = None
    max_budget = None

    if use_budget:
        minimum_price = float(
            ev_data["price_lakh"].dropna().min()
        )

        maximum_price = float(
            ev_data["price_lakh"].dropna().max()
        )

        min_budget, max_budget = st.slider(
            "Budget in ₹ lakh",
            min_value=minimum_price,
            max_value=maximum_price,
            value=(
                minimum_price,
                min(30.0, maximum_price)
            ),
            step=0.5
        )

    use_range = st.checkbox(
        "Select minimum required range"
    )

    min_range = None

    if use_range:
        min_range = st.number_input(
            "Minimum range in km",
            min_value=0,
            max_value=1000,
            value=300,
            step=25
        )

    body_options = [
        "Any",
        *sorted(
            ev_data["body_type"]
            .dropna()
            .unique()
        )
    ]

    preferred_body_type = st.selectbox(
        "Preferred body type",
        body_options
    )

    passenger_selection = st.selectbox(
        "Number of passengers",
        ["Any", 2, 3, 4, 5, 6, 7]
    )

    passengers = (
        None
        if passenger_selection == "Any"
        else int(passenger_selection)
    )

    safety_selection = st.selectbox(
        "Minimum safety rating",
        ["Any", 3, 4, 5]
    )

    min_safety = (
        None
        if safety_selection == "Any"
        else float(safety_selection)
    )

    fast_charging = st.selectbox(
        "Fast charging required",
        ["Any", "Yes", "No"]
    )

    top_n = st.slider(
        "Number of recommendations",
        min_value=3,
        max_value=10,
        value=5
    )

    recommend_button = st.button(
        "Recommend EVs",
        use_container_width=True
    )


if recommend_button:
    results = recommend_evs(
        data=ev_data,
        preferred_brand=preferred_brand,
        min_budget=min_budget,
        max_budget=max_budget,
        min_range=min_range,
        preferred_body_type=preferred_body_type,
        passengers=passengers,
        min_safety=min_safety,
        fast_charging=fast_charging,
        top_n=top_n,
        popularity_weight=0.25
    )

    if results.empty:
        st.warning(
            "No EV matches all selected requirements."
        )
    else:
        st.subheader("Recommended EVs")

        for index, row in results.iterrows():
            with st.container(border=True):
                st.subheader(
                    f"{index + 1}. "
                    f"{row['brand']} {row['model']}"
                )

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Price",
                    f"₹{row['price_lakh']:.2f} lakh"
                )

                col2.metric(
                    "Claimed range",
                    f"{row['claimed_range_km']:.0f} km"
                )

                col3.metric(
                    "Recommendation score",
                    f"{row['final_score']:.3f}"
                )

                st.write(
                    f"**Variant:** {row['variant']}"
                )

                st.write(
                    f"**Body type:** {row['body_type']}"
                )

                st.write(
                    f"**Seats:** "
                    f"{int(row['seating_capacity'])}"
                )

                if pd.notna(row["safety_rating"]):
                    st.write(
                        f"**Safety rating:** "
                        f"{row['safety_rating']:.0f}/5"
                    )
                else:
                    st.write(
                        "**Safety rating:** Not available"
                    )

                st.write(
                    f"**Fast charging:** "
                    f"{row['fast_charging_supported']}"
                )

                if pd.notna(row["user_rating"]):
                    st.write(
                        f"**User rating:** "
                        f"{row['user_rating']:.1f}/5"
                    )

                st.caption(
                    f"Popularity score: "
                    f"{row['popularity_score']:.3f}"
                )

else:
    st.info(
        "Select preferences from the sidebar and "
        "click Recommend EVs. Leave everything as Any "
        "for popularity-based recommendations."
    )
