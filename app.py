# app.py
from flask import Flask, render_template, abort, redirect, url_for, make_response
import plotly.express as px
import plotly
import json
import re
import base64
from collections import defaultdict
from collections import OrderedDict
from collections import Counter
from jinja2 import Environment, FileSystemLoader
from pathlib import Path
import pandas as pd
from datetime import datetime

from services import google_sheets as gs

BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__, template_folder=str(BASE_DIR), static_folder=str(BASE_DIR / "static"))

# ===== HELPER FUNCTIONS (PLACE AT TOP) =====

def safe_get_first(dlist, key_candidates):
    """Return first matching key value from a dict list's first record."""
    if not dlist:
        return None
    row = dlist[0]
    for k in key_candidates:
        if k in row and row[k] not in (None, ""):
            return row[k]
    for k in row:
        if k.lower() in [c.lower() for c in key_candidates]:
            return row[k]
    return None

def clean_number(val):
    """Convert '15,750' or '$\\mathbf{5,250}$' to float 15750.0"""
    if pd.isna(val) or val == "":
        return None
    s = str(val)
    s = re.sub(r'[^0-9.]', '', s)
    return float(s) if s else None

def clean_latex_math(text):
    """Clean LaTeX math notation for display: $17\%$ → 17%, \rightarrow → →, etc."""
    if not isinstance(text, str):
        return ""
    text = re.sub(r'\\mathbf\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\rightarrow', '→', text)
    text = re.sub(r'\$(.*?)\$', r'\1', text)
    return text.strip()
def parse_number(value):
    if value in (None, ''):
        return None
    s = str(value).replace(',', '').replace(' ', '')
    s = re.sub(r"[^\d\.\-]", '', s)
    try:
        return float(s)
    except Exception:
        return None

def fmt_money(value, decimals=2):
    num = parse_number(value)
    if num is None:
        return ''
    return f"{num:,.{decimals}f}"


def build_dashboard_context():
    ecom_target = build_ecom_target_context()
    ecom_comp = build_ecom_comparison_context()
    strategy = get_strategy_plan_context()
    roadmap = get_roadmap_context()

    primary_rows = []
    columns = ecom_target.get('target_columns') or []
    first_col = columns[0] if columns else ''
    value_cols = columns[1:] if len(columns) > 1 else []
    for row in ecom_target.get('target_rows', [])[:5]:
        label = row.get(first_col, '') if first_col else ''
        value = ''
        for col in value_cols:
            val = row.get(col)
            if val not in (None, '', 'nan'):
                numeric = parse_number(val)
                value = fmt_money(val) if numeric is not None else val
                break
        primary_rows.append({'label': label, 'value': value})

    secondary_rows = [
        {
            'label': r.get('Months', ''),
            'value': f"{r.get('2024_fmt', '')} → {r.get('2025_fmt', '')}"
        }
        for r in ecom_comp.get('comp_rows', [])[:5]
    ]

    cards = [
        {
            'title': '2026 Target Plan',
            'rows': primary_rows,
            'link': url_for('ecom'),
            'download': url_for('download_ecom_target')
        },
        {
            'title': '2024 vs 2025 Performance',
            'rows': secondary_rows,
            'link': url_for('ecom_comparison'),
            'download': url_for('download_ecom_comparison')
        }
    ]

    return {
        'cards': cards,
        'strategy': strategy,
        'roadmap': roadmap,
        'title': 'Strategy Dashboard'
    }



def render_dashboard():
    context = build_dashboard_context()
    return render_template("index.html", **context)


@app.route("/")
def index():
    return render_dashboard()


@app.route("/home")
def home_page():
    return render_dashboard()


@app.route("/dashboard")
def dashboard():
    return render_dashboard()


def build_ecom_target_context():
    try:
        df = gs.sheet_to_df("2026 Ecom Target")
    except Exception as e:
        print("Error loading sheet '2026 Ecom Target':", e)
        df = pd.DataFrame()

    cols = []
    formatted_rows = []
    insight_text = ""

    if not df.empty:
        df.columns = (
            df.columns.astype(str)
            .str.replace(" ", " ", regex=False)
            .str.replace("	", " ", regex=False)
            .str.replace("  ", " ", regex=False)
            .str.strip()
        )
        first_col = df.columns[0]
        mask = df[first_col].astype(str).str.strip().str.lower().str.contains(r"insight", na=False)
        if mask.any():
            insight_row = df[mask].iloc[0]
            parts = []
            for c in df.columns[1:]:
                clean_val = str(insight_row.get(c, "")).strip()
                if clean_val and clean_val.lower() != 'nan':
                    parts.append(clean_val)
            insight_text = " ".join(parts).strip()
        else:
            long_text = ""
            for i in range(len(df) - 1, -1, -1):
                row = df.iloc[i].astype(str).fillna("")
                values = [str(x).strip() for x in row.tolist() if x and str(x).strip().lower() != 'nan']
                joined = " ".join(values)
                if len(joined) > 40 and not re.search(r"\d{3,}", joined):
                    long_text = joined
                    break
            insight_text = long_text.strip()

        cols = list(df.columns)
        for _, r in df.iterrows():
            row_out = {}
            for c in cols:
                val = r.get(c, "")
                if any(k in c.lower() for k in ["amount", "target", "sales", "moonshot", "fulfillment"]):
                    num = parse_number(val)
                    row_out[c] = fmt_money(num) if num is not None else ""
                else:
                    row_out[c] = val if (val is not None and str(val).strip() != "nan") else ""
            formatted_rows.append(row_out)

    return {
        "active_tab": "target",
        "target_columns": cols,
        "target_rows": formatted_rows,
        "target_insight": insight_text,
        "comp_rows": [],
        "comp_summary": {},
        "title": "E-commerce Performance",
        "description": "2026 target plan across the funnel."
    }


def build_ecom_comparison_context():
    try:
        df = gs.sheet_to_df("ecom 2024 vs 2025")
    except Exception as e:
        print("Error loading sheet 'ecom 2024 vs 2025':", e)
        df = pd.DataFrame()

    df.columns = df.columns.astype(str).str.strip() if not df.empty else []
    if not df.empty:
        keep = [c for c in ["Months", "2024", "2025"] if c in df.columns]
        df = df[keep]

    def fmt_int(val):
        try:
            return f"{float(val):,.0f}"
        except Exception:
            return "0"

    def fmt_pct(val):
        try:
            return f"{float(val):.1f}%"
        except Exception:
            return "—"

    records = []
    total_2024 = total_2025 = 0
    max_row = min_row = None

    if not df.empty:
        for _, row in df.iterrows():
            month = str(row.get("Months", "")).strip()
            val_2024 = parse_number(row.get("2024")) or 0
            val_2025 = parse_number(row.get("2025")) or 0
            delta = val_2025 - val_2024
            delta_pct = ((delta / val_2024) * 100) if val_2024 else None
            entry = {
                "Months": month,
                "2024": val_2024,
                "2025": val_2025,
                "2024_fmt": fmt_int(val_2024),
                "2025_fmt": fmt_int(val_2025),
                "delta": delta,
                "delta_fmt": fmt_int(delta),
                "delta_pct": delta_pct,
                "delta_pct_fmt": fmt_pct(delta_pct) if delta_pct is not None else "—"
            }
            records.append(entry)
            total_2024 += val_2024
            total_2025 += val_2025

        if records:
            max_row = max(records, key=lambda r: r["2025"])
            eligible_min = [r for r in records if r["Months"].lower() not in ("dec", "december")]
            min_source = eligible_min if eligible_min else records
            min_row = min(min_source, key=lambda r: r["2025"])

    count = len(records) if records else 1
    comp_summary = {
        "total_2024": fmt_int(total_2024),
        "total_2025": fmt_int(total_2025),
        "avg_2024": fmt_int(total_2024 / count),
        "avg_2025": fmt_int(total_2025 / count),
        "max_month": max_row["Months"] if max_row else "-",
        "max_value": fmt_int(max_row["2025"]) if max_row else "0",
        "min_month": min_row["Months"] if min_row else "-",
        "min_value": fmt_int(min_row["2025"]) if min_row else "0",
    }

    return {
        "active_tab": "comparison",
        "target_columns": [],
        "target_rows": [],
        "target_insight": "",
        "comp_rows": records,
        "comp_summary": comp_summary,
        "title": "E-commerce Performance",
        "description": "2024 vs 2025 performance snapshot."
    }


@app.route("/ecom")
def ecom():
    return render_template("ecom.html", **build_ecom_target_context())


@app.route("/ecom_comp")
def ecom_comparison():
    return render_template("ecom.html", **build_ecom_comparison_context())


@app.route("/ecom/download")
def download_ecom_target():
    return make_download_response("ecom.html", "ecom_target", **build_ecom_target_context())


@app.route("/ecom_comp/download")
def download_ecom_comparison():
    return make_download_response("ecom.html", "ecom_comparison", **build_ecom_comparison_context())


def get_strategy_plan_context():
    try:
        df = gs.sheet_to_df("2026 Strategy plan")
    except Exception as e:
        print("Error loading 2026 Strategy plan:", e)
        df = pd.DataFrame()

    context = {
        "goal_text": "2026 Strategy Plan",
        "pillars": {},
        "title": "2026 Strategy Plan"
    }

    if df.empty:
        return context

    df.columns = df.columns.astype(str).str.strip()
    required = ["Goal", "Strategy Pillar", "Phase", "Quarter", "Action", "Photo_URL 1", "Photo_URL 2", "Photo_URL 3"]
    for col in required:
        if col not in df.columns:
            df[col] = ""

    df = df.fillna("")

    current_goal = ""
    pillars = OrderedDict()

    for _, row in df.iterrows():
        goal_raw = str(row.get("Goal", "")).strip()
        if goal_raw:
            current_goal = goal_raw
        pillar = str(row.get("Strategy Pillar", "")).strip() or "General"
        entry = {
            "goal": current_goal,
            "phase": row.get("Phase", ""),
            "quarter": row.get("Quarter", ""),
            "action": row.get("Action", ""),
            "photos": [
                url for url in [row.get("Photo_URL 1", ""), row.get("Photo_URL 2", ""), row.get("Photo_URL 3", "")]
                if url and str(url).strip()
            ]
        }
        pillars.setdefault(pillar, []).append(entry)

    context["goal_text"] = current_goal or "2026 Strategy Plan"
    context["pillars"] = pillars
    return context


@app.route("/strategy_plan")
def strategy_plan():
    return render_template("strategy_plan.html", **get_strategy_plan_context())


@app.route("/strategy_plan/download")
def download_strategy_plan():
    return make_download_response("strategy_plan.html", "strategy_plan", **get_strategy_plan_context())


def get_roadmap_context():
    try:
        df = gs.get_roadmap()
    except Exception as e:
        print(f"❌ Error loading roadmap: {e}")
        df = pd.DataFrame()

    if df.empty:
        return {"quarters": {}, "quarter_order": []}

    df.columns = [c.strip() for c in df.columns]
    records = df.to_dict(orient="records")

    def get_value(row, *keys):
        for key in keys:
            key_lower = key.strip().lower()
            for actual_key, value in row.items():
                if actual_key and actual_key.strip().lower() == key_lower:
                    return "" if value in (None, "") else str(value)
        return ""

    quarters = {}
    for row in records:
        q = get_value(row, "Quarter") or "Unassigned"
        quarters.setdefault(q, []).append({
            "Activity_ID": get_value(row, "Activity_ID", "Activity ID"),
            "Topic": get_value(row, "Key Topic", "Key_Topic", "Key Activity"),
            "Owner": get_value(row, "Owner"),
        })

    quarter_order = sorted(quarters.keys())
    return {"quarters": quarters, "quarter_order": quarter_order}


@app.route("/roadmap")
def roadmap_page():
    return render_template("roadmap.html", **get_roadmap_context())


@app.route("/roadmap/download")
def download_roadmap():
    return make_download_response("roadmap.html", "roadmap", **get_roadmap_context())


@app.route("/swot")
def swot_page():
    try:
        df = gs.sheet_to_df("swot")
    except Exception:
        df = pd.DataFrame()

    rows = []
    try:
        rows = df.to_dict(orient="records") if hasattr(df, "to_dict") else []
    except Exception:
        rows = []

    sections = OrderedDict()
    key_insights = []

    for row in rows:
        category = (row.get("Category") or "").strip()
        point_id = row.get("Point_ID") or ""
        key_item = row.get("Key_Item") or row.get("Key Item") or ""
        insight_2025 = row.get("2025") or row.get("2025 Insight") or ""
        strategy_2026 = row.get("2026") or row.get("2026 Strategy") or ""

        if not category:
            continue

        if category.lower() == "key insight":
            key_insights.append({
                "title": key_item,
                "content": insight_2025 or strategy_2026 or point_id
            })
            continue

        sections.setdefault(category, []).append({
            "id": point_id,
            "title": key_item,
            "details_2025": insight_2025,
            "details_2026": strategy_2026
        })

    return render_template(
        "swot.html",
        sections=sections,
        key_insights=key_insights
    )


@app.route("/cost_per_x")
def cost_per_x():
    try:
        df = gs.sheet_to_df("Cost per X")
    except Exception as e:
        print("GS error (Cost per X):", e)
        df = pd.DataFrame()

    rows = []
    if not df.empty:
        df.columns = (
            df.columns.astype(str)
            .str.replace("\u00A0", " ", regex=False)
            .str.replace("\t", " ", regex=False)
            .str.replace("  ", " ", regex=False)
            .str.strip()
        )

        # Map probable headers to canonical names
        rename_map = {}
        for c in df.columns:
            lc = c.lower()
            if "cost per x" in lc:
                rename_map[c] = "Cost per X"
            elif lc.startswith("facts"):
                rename_map[c] = "Facts"
            elif lc.startswith("why"):
                rename_map[c] = "Why?"
            elif "what to improve" in lc or "improve" in lc:
                rename_map[c] = "What to Improve More?"

        if rename_map:
            df = df.rename(columns=rename_map)

        # Ensure required columns exist
        for col in ["Cost per X", "Facts", "Why?", "What to Improve More?"]:
            if col not in df.columns:
                df[col] = ""

        # Normalize newlines for display
        for col in ["Facts", "Why?", "What to Improve More?"]:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.replace("\r\n", "\n")
                .str.replace("\r", "\n")
            )

        rows = (
            df[["Cost per X", "Facts", "Why?", "What to Improve More?"]]
            .fillna("")
            .to_dict(orient="records")
        )

    return render_template("cost_per_x.html", rows=rows, title="Cost per X")


@app.route("/okr")
def okr_page():
    try:
        df = gs.sheet_to_df("okr")  # ← Sheet name = "okr"
    except Exception as e:
        print(f"❌ Error loading OKR: {e}")
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    # Separate 2025 and 2026
    okr_2025 = []
    okr_2026 = []

    for row in rows:
        year = str(row.get("Years", "")).strip()
        if "2025" in year:
            okr_2025.append(row)
        elif "2026" in year:
            okr_2026.append(row)

    # Group by Functional Team
    teams_2025 = defaultdict(list)
    teams_2026 = defaultdict(list)

    for row in okr_2025:
        team = row.get("Functional POVs", "Other").strip()
        teams_2025[team].append(row)

    for row in okr_2026:
        team = row.get("Functional POVs", "Other").strip()
        teams_2026[team].append(row)

    # Merge all teams (union of keys)
    all_teams = sorted(set(teams_2025.keys()) | set(teams_2026.keys()))

    # Build comparison structure
    comparison = []

    def compute_avg(rows, field="Average"):
        values = []
        for r in rows:
            v = r.get(field)
            if v in (None, ""):
                continue
            s = str(v).replace("%", "").strip()
            try:
                # Treat sheet values like 0.7 as 70%
                values.append(float(s) * 100.0)
            except Exception:
                continue
        if not values:
            return None
        return round(sum(values) / len(values), 1)

    for team in all_teams:
        items_2025 = teams_2025.get(team, [])
        items_2026 = teams_2026.get(team, [])

        # Group by Objective
        obj_map = {}
        for item in items_2025:
            obj = item.get("Objective", "No Objective")
            if obj not in obj_map:
                obj_map[obj] = {"2025": [], "2026": []}
            obj_map[obj]["2025"].append(item)

        for item in items_2026:
            obj = item.get("Objective", "No Objective")
            if obj not in obj_map:
                obj_map[obj] = {"2025": [], "2026": []}
            obj_map[obj]["2026"].append(item)

        comparison.append({
            "team": team,
            "objectives": [
                {
                    "objective": obj,
                    "items_2025": data["2025"],
                    "items_2026": data["2026"]
                }
                for obj, data in obj_map.items()
            ],
            "avg_2025": compute_avg(items_2025),
            "avg_2026": compute_avg(items_2026),
        })

    return render_template(
        "okr.html",
        comparison=comparison,
        title="OKR Dashboard: 2025 vs 2026"
    )

def clean_number(val):
    """Convert '3.7 B' or '180 L' to float"""
    if not isinstance(val, str):
        return val
    val = val.strip()
    if not val:
        return val
    
    # Handle "3.7 B" → 3700000000
    match = re.search(r'([\d.]+)\s*([BM]?)', val)
    if match:
        num = float(match.group(1))
        unit = match.group(2)
        if unit == 'B':
            return num * 1_000_000_000
        elif unit == 'M':
            return num * 1_000_000
        elif unit == 'L':
            return num * 100_000
        else:
            return num
    return val


def load_fna_performance_from_excel():
    excel_path = Path(__file__).resolve().parent / "strategic_insight.xlsx"
    try:
        df = pd.read_excel(excel_path, sheet_name="fna_performance")
    except Exception as e:
        print(f"❌ Error reading FNA performance sheet: {e}")
        return [], Counter()

    if df.empty:
        return [], Counter()

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\u00A0", " ", regex=False)
    )

    records = df.fillna("").to_dict(orient="records")

    category_counter = Counter()
    for row in records:
        category = str(row.get("KPI Category", "General")).strip() or "General"
        category_counter[category] += 1
        for key, value in row.items():
            if isinstance(value, str):
                cleaned = re.sub(r"\\mathbf\{([^}]*)\}", r"\1", value)
                cleaned = re.sub(r"\\rightarrow", "→", cleaned)
                cleaned = re.sub(r"\$(.*?)\$", r"\1", cleaned)
                row[key] = cleaned.strip()

    return records, category_counter


@app.route("/fna_performance")
def fna_performance_page():
    fna_rows, category_counter = load_fna_performance_from_excel()

    palette = [
        "#1976d2", "#ef6c00", "#2e7d32",
        "#6a1b9a", "#00838f", "#c62828"
    ]
    category_meta = {}
    for idx, (category, count) in enumerate(category_counter.items()):
        category_meta[category] = {
            "color": palette[idx % len(palette)],
            "count": count
        }

    return render_template(
        "fna_performance.html",
        fna=fna_rows,
        category_meta=category_meta,
        title="FNA Performance"
    )


def split_operations(rows):
    insights = []
    stages = defaultdict(list)
    status_counter = Counter()
    for row in rows:
        stage = str(row.get("Funnel Stage", "")).strip()
        status = str(row.get("Status", "")).strip()
        if stage.lower() == "insight":
            insights.append(row)
            continue
        stages[stage or "Unassigned"].append(row)
        if status:
            status_counter[status] += 1
    return insights, stages, status_counter


@app.route("/operation_health")
def operation_health_page():
    try:
        df = gs.get_operation_health()
    except Exception as e:
        print(f"❌ Error loading operation health: {e}")
        df = pd.DataFrame()

    if df.empty:
        return render_template(
            "operation_health.html",
            grouped_ops={},
            insights=[],
            status_counter={},
            title="Operations Health"
        )

    df.columns = df.columns.astype(str).str.strip()
    records = df.fillna("").to_dict(orient="records")
    insights, grouped_ops, status_counter = split_operations(records)

    return render_template(
        "operation_health.html",
        grouped_ops=grouped_ops,
        insights=insights,
        status_counter=status_counter,
        title="Operations Health"
    )


@app.route("/bob")
def bob_page():
    bob_df = read_local_excel_sheet("BOB")
    review_df = read_local_excel_sheet("BOB_review")

    def normalize_columns(df):
        if df.empty:
            return df
        df.columns = df.columns.astype(str).str.strip()
        return df

    bob_df = normalize_columns(bob_df)
    review_df = normalize_columns(review_df)

    # Ensure numeric table columns exist even if sheet uses variants like 'Grand_Total '
    bob_column_map = {
        "months": "Months",
        "month": "Months",
        "bob order": "BOB Order",
        "bob": "BOB Order",
        "boborder": "BOB Order",
        "self order": "Self Order",
        "self": "Self Order",
        "grand total": "Grand Total",
        "total": "Grand Total",
        "cs%": "CS%",
        "cs %": "CS%",
        "cs percentage": "CS%"
    }

    if not bob_df.empty:
        rename_map = {}
        for col in bob_df.columns:
            key = col.strip().lower()
            if key in bob_column_map:
                rename_map[col] = bob_column_map[key]
        if rename_map:
            bob_df = bob_df.rename(columns=rename_map)

    bob_rows = []
    chart_data = {"months": [], "bob": [], "self": [], "cs": []}
    totals = {"bob": 0, "self": 0, "grand": 0, "cs": []}
    best_month = {"label": "-", "value": 0}

    if not bob_df.empty:
        for _, row in bob_df.iterrows():
            month = str(row.get("Months", "")).strip()
            bob_val = parse_number(row.get("BOB Order")) or 0
            self_val = parse_number(row.get("Self Order")) or 0
            grand_val = parse_number(row.get("Grand Total")) or 0
            cs_val = parse_number(row.get("CS%"))

            bob_rows.append({
                "Months": month,
                "BOB Order": format_number(bob_val),
                "Self Order": format_number(self_val),
                "Grand Total": format_number(grand_val),
                "CS%": format_percent(cs_val)
            })

            chart_data["months"].append(month)
            chart_data["bob"].append(bob_val)
            chart_data["self"].append(self_val)
            chart_data["cs"].append((cs_val * 100) if (cs_val is not None and abs(cs_val) <= 1) else (cs_val or 0))

            totals["bob"] += bob_val
            totals["self"] += self_val
            totals["grand"] += grand_val
            if cs_val is not None:
                totals["cs"].append(cs_val if abs(cs_val) <= 1 else cs_val / 100)

            if grand_val > best_month["value"]:
                best_month = {"label": month, "value": grand_val}

    review_sections = [
        {"key": "worked", "label": "What Worked?"},
        {"key": "scale", "label": "What needs to scale?"},
        {"key": "not_work", "label": "What did not work?"},
        {"key": "lesson", "label": "What is the lesson learned?"},
        {"key": "next_goal", "label": "What is the next goal for BOB?"}
    ]

    reviews = []
    if not review_df.empty:
        column_map = {
            "what worked?": "worked",
            "what needs to scale?": "scale",
            "what did not work?": "not_work",
            "what is the lesson learned?": "lesson",
            "what is the next goal for bob?": "next_goal"
        }
        keys_defaults = {sec["key"]: [] for sec in review_sections}
        for _, row in review_df.iterrows():
            entry = dict(keys_defaults)
            for col in review_df.columns:
                normalized = str(col).strip().lower()
                if normalized in column_map:
                    entry[column_map[normalized]] = parse_review_text(row.get(col, ""))
            if any(entry.values()):
                reviews.append(entry)

    summary = {
        "total_bob": format_number(totals["bob"]),
        "total_self": format_number(totals["self"]),
        "total_grand": format_number(totals["grand"]),
        "avg_cs": format_percent(sum(totals["cs"]) / len(totals["cs"]) if totals["cs"] else None),
        "best_month": best_month["label"],
        "best_month_value": format_number(best_month["value"])
    }

    return render_template(
        "bob.html",
        rows=bob_rows,
        reviews=reviews,
        review_sections=review_sections,
        summary=summary,
        chart_data=chart_data,
        title="BOB Performance",
        description="Monthly BOB volume split with qualitative learnings"
    )


if __name__ == "__main__":  
    app.run(debug=True, host="0.0.0.0", port=5000)
    