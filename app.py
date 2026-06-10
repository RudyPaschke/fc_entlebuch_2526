from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

APP_TITLE = "FC Entlebuch – Team- & Spielerstatistik 2025/26"
DATA_FILE = Path(__file__).parent / "data" / "FC_Entlebuch_Daten_26_27.xlsx"
TEAM_NAME = "FC Entlebuch"

# FC-Entlebuch-Branding: grün als Hauptfarbe, rot/schwarz/weiss als Akzentfarben.
FCE_GREEN = "#00843D"
FCE_DARK_GREEN = "#005A2B"
FCE_LIGHT_GREEN = "#EAF7EF"
FCE_RED = "#E30613"
FCE_BLACK = "#111111"
FCE_GREY = "#F5F7F6"
FCE_COLOR_MAP = {
    "FC Entlebuch": FCE_GREEN,
    TEAM_NAME: FCE_GREEN,
    "Gegner": FCE_BLACK,
    "Tore": FCE_GREEN,
    "Gegentore": FCE_RED,
    "Tore_FCE": FCE_GREEN,
    "Tore_Gegner": FCE_RED,
}

st.set_page_config(
    page_title="FC Entlebuch Statistik",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CUSTOM_CSS = f"""
<style>
:root {{
    --fce-green: {FCE_GREEN};
    --fce-dark-green: {FCE_DARK_GREEN};
    --fce-light-green: {FCE_LIGHT_GREEN};
    --fce-red: {FCE_RED};
    --fce-black: {FCE_BLACK};
    --fce-grey: {FCE_GREY};
}}

.block-container {{padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px;}}

/* Gesamter Hintergrund */
[data-testid="stAppViewContainer"] {{
    background: linear-gradient(180deg, #ffffff 0%, var(--fce-grey) 55%, #ffffff 100%);
}}

/* Sidebar */
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, var(--fce-dark-green) 0%, var(--fce-green) 100%);
}}
[data-testid="stSidebar"] * {{
    color: white !important;
}}
[data-testid="stSidebar"] [data-baseweb="select"] * {{
    color: #111111 !important;
}}
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stSelectbox label {{
    font-weight: 700;
}}

/* Hero */
.fce-hero {{
    border-radius: 22px;
    padding: 28px 30px;
    margin-bottom: 22px;
    background: linear-gradient(135deg, var(--fce-dark-green) 0%, var(--fce-green) 62%, #10a862 100%);
    color: white;
    box-shadow: 0 14px 35px rgba(0, 90, 43, 0.22);
    border-left: 8px solid var(--fce-red);
}}
.fce-kicker {{
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 0.13em;
    font-weight: 800;
    opacity: 0.92;
    margin-bottom: 5px;
}}
.fce-title {{
    font-size: clamp(2rem, 4vw, 3.3rem);
    font-weight: 900;
    line-height: 1.02;
    margin: 0;
}}
.fce-subtitle {{
    font-size: 1.05rem;
    opacity: 0.95;
    margin-top: 10px;
}}
.fce-pill {{
    display: inline-block;
    margin-top: 16px;
    background: rgba(255,255,255,0.16);
    border: 1px solid rgba(255,255,255,0.28);
    border-radius: 999px;
    padding: 7px 13px;
    font-weight: 700;
}}

/* Metrics */
[data-testid="stMetric"] {{
    background: white;
    border: 1px solid #e5e7eb;
    border-top: 5px solid var(--fce-green);
    padding: 16px 18px;
    border-radius: 16px;
    box-shadow: 0 8px 22px rgba(17, 17, 17, 0.06);
}}
[data-testid="stMetricValue"] {{
    font-size: 1.75rem;
    font-weight: 900;
    color: var(--fce-dark-green);
}}
[data-testid="stMetricLabel"] {{
    font-weight: 800;
    color: #374151;
}}

/* Tabs */
button[data-baseweb="tab"] {{
    font-weight: 800;
}}
button[data-baseweb="tab"][aria-selected="true"] {{
    color: var(--fce-green) !important;
    border-bottom-color: var(--fce-green) !important;
}}

/* Buttons */
.stDownloadButton button, .stButton button {{
    background: var(--fce-green);
    color: white;
    border: 0;
    border-radius: 12px;
    font-weight: 800;
}}
.stDownloadButton button:hover, .stButton button:hover {{
    background: var(--fce-dark-green);
    color: white;
}}

h1, h2, h3 {{
    color: var(--fce-black);
}}
.small-note {{color: #6b7280; font-size: 0.9rem;}}
.fce-section-note {{
    background: white;
    border-left: 5px solid var(--fce-red);
    padding: 11px 14px;
    border-radius: 10px;
    color: #374151;
    margin-bottom: 14px;
}}

.fce-card {{
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 16px 18px;
    box-shadow: 0 8px 22px rgba(17, 17, 17, 0.06);
    margin-bottom: 12px;
}}
.fce-card-title {{
    font-size: 0.85rem;
    color: #6b7280;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: .06em;
}}
.fce-card-value {{
    font-size: 1.7rem;
    font-weight: 900;
    color: var(--fce-dark-green);
    line-height: 1.2;
    margin-top: 3px;
}}
.fce-card-sub {{
    font-size: 0.86rem;
    color: #6b7280;
    margin-top: 3px;
}}


/* Mobile-Optimierung */
@media (max-width: 768px) {{
    .block-container {{padding: 0.75rem 0.7rem 1.4rem 0.7rem;}}
    .fce-hero {{border-radius: 16px; padding: 20px 18px; margin-bottom: 14px;}}
    .fce-title {{font-size: 2.0rem; line-height: 1.05;}}
    .fce-subtitle {{font-size: 0.95rem;}}
    .fce-pill {{font-size: 0.82rem; padding: 6px 10px;}}
    [data-testid="stMetric"] {{padding: 12px 12px; border-radius: 12px;}}
    [data-testid="stMetricValue"] {{font-size: 1.35rem;}}
    button[data-baseweb="tab"] {{font-size: 0.82rem; padding-left: 0.4rem; padding-right: 0.4rem;}}
    div[data-testid="stHorizontalBlock"] {{gap: 0.5rem;}}
}}
.fce-metric-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px;
    margin: 8px 0 16px 0;
}}
.fce-mini-table-note {{
    color: #6b7280;
    font-size: 0.88rem;
    margin-top: -6px;
    margin-bottom: 10px;
}}

</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def _to_int_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)


def _result_sort_key(value):
    try:
        return int(value)
    except Exception:
        return str(value)


@st.cache_data(show_spinner=False)
def load_data(path: str | Path):
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Daten-Datei nicht gefunden: {path}")

    spieler = pd.read_excel(path, sheet_name="Spieler")
    spiele = pd.read_excel(path, sheet_name="Spiele")
    stats = pd.read_excel(path, sheet_name="SpielerSpielStatistik")
    tore = pd.read_excel(path, sheet_name="Tore")

    # Basis-Bereinigung
    for df in [spieler, spiele, stats, tore]:
        df.columns = [str(c).strip() for c in df.columns]

    spieler = spieler.dropna(how="all")
    spiele = spiele.dropna(how="all")
    stats = stats.dropna(how="all")
    tore = tore.dropna(how="all")

    # Datentypen
    for col in ["Spieler_ID", "Nr"]:
        if col in spieler.columns:
            spieler[col] = _to_int_series(spieler[col])

    if "Spieler" not in spieler.columns:
        spieler["Spieler"] = spieler.get("Vorname", "").fillna("").astype(str) + " " + spieler.get("Name", "").fillna("").astype(str)
    spieler["Spieler"] = spieler["Spieler"].fillna("").astype(str).str.strip()

    for col in ["Spiel_ID", "Tore_FCE", "Tore_Gegner"]:
        if col in spiele.columns:
            spiele[col] = _to_int_series(spiele[col])
    if "Datum" in spiele.columns:
        spiele["Datum"] = pd.to_datetime(spiele["Datum"], errors="coerce")
        # Datenkorrektur: Meisterschaft Runde 9 gegen Ruswil fand am 18.10.2025 statt.
        try:
            mask_ruswil = (spiele["Wettbewerb"].astype(str).str.strip().eq("Meisterschaft")) & (spiele["Runde"].astype(str).str.strip().eq("9")) & (spiele["Gegner"].astype(str).str.contains("Ruswil", case=False, na=False))
            spiele.loc[mask_ruswil, "Datum"] = pd.Timestamp("2025-10-18")
        except Exception:
            pass
    for col in ["Phase", "Wettbewerb", "Heim/Auswärts", "Gegner", "Ergebnis", "Resultat", "Runde"]:
        if col in spiele.columns:
            spiele[col] = spiele[col].fillna("").astype(str)

    for col in ["Spiel_ID", "Spieler_ID", "Startelf", "Eingewechselt", "Minuten", "Tore", "Assists", "Scorer", "Gelb", "Gelb-Rot", "Rot"]:
        if col in stats.columns:
            stats[col] = _to_int_series(stats[col])
    if "Status" in stats.columns:
        stats["Status"] = stats["Status"].fillna("").astype(str).str.strip()
        # Einheitliche Status-Bezeichnungen für die App:
        # S = Startelf, E = Eingewechselt, 0 Minuten = Kein Einsatz.
        stats["Status"] = stats["Status"].replace({
            "S": "Startelf",
            "s": "Startelf",
            "E": "Eingewechselt",
            "e": "Eingewechselt",
            "Ohne Einsatz": "Kein Einsatz",
            "ohne Einsatz": "Kein Einsatz",
        })
        stats.loc[stats["Minuten"].fillna(0) == 0, "Status"] = "Kein Einsatz"
        stats.loc[(stats["Minuten"].fillna(0) > 0) & (stats["Status"].str.strip() == "") & (stats.get("Startelf", 0).fillna(0) > 0), "Status"] = "Startelf"
        stats.loc[(stats["Minuten"].fillna(0) > 0) & (stats["Status"].str.strip() == "") & (stats.get("Eingewechselt", 0).fillna(0) > 0), "Status"] = "Eingewechselt"

    for col in ["Tor_ID", "Spiel_ID", "Minute", "Spieler_ID"]:
        if col in tore.columns:
            tore[col] = pd.to_numeric(tore[col], errors="coerce")
    tore["Spiel_ID"] = tore["Spiel_ID"].fillna(0).astype(int)
    tore["Minute"] = tore["Minute"].fillna(0).astype(int)
    tore["Spieler_ID"] = tore["Spieler_ID"].fillna(0).astype(int)
    for col in ["Team", "Spielstand_Nach_Tor"]:
        if col in tore.columns:
            tore[col] = tore[col].fillna("").astype(str)

    # Angereicherte Tabellen
    stats_full = (
        stats.merge(spieler[["Spieler_ID", "Nr", "Name", "Vorname", "Spieler"]], on="Spieler_ID", how="left")
        .merge(
            spiele[[
                "Spiel_ID", "Saison", "Runde", "Phase", "Wettbewerb", "Datum", "Heim/Auswärts",
                "Heimteam", "Auswärtsteam", "Gegner", "Resultat", "Resultat Halbzeit", "Ergebnis", "Punkte",
                "Tore_FCE", "Tore_Gegner"
            ]],
            on="Spiel_ID",
            how="left",
        )
    )
    stats_full["Einsatz"] = (stats_full["Minuten"] > 0).astype(int)

    tore_full = tore.merge(spiele[["Spiel_ID", "Datum", "Wettbewerb", "Phase", "Heim/Auswärts", "Gegner"]], on="Spiel_ID", how="left")
    tore_full = tore_full.merge(spieler[["Spieler_ID", "Nr", "Spieler"]], on="Spieler_ID", how="left")
    tore_full["Torschütze"] = tore_full["Spieler"].fillna("")
    tore_full.loc[tore_full["Spieler_ID"] == 0, "Torschütze"] = "Gegner"

    return spieler, spiele, stats, tore, stats_full, tore_full


def make_game_label(row: pd.Series) -> str:
    date = row["Datum"].strftime("%d.%m.%Y") if pd.notna(row.get("Datum")) else "ohne Datum"
    runde = row.get("Runde", "")
    wettbewerb = row.get("Wettbewerb", "")
    gegner = row.get("Gegner", "")
    resultat = row.get("Resultat", "")
    return f"{date} | {wettbewerb} {runde} | {gegner} | {resultat}"


def torfolge_for_game(tore_df: pd.DataFrame, spiel_id: int) -> str:
    g = tore_df[tore_df["Spiel_ID"] == spiel_id].sort_values(["Minute", "Tor_ID"] if "Tor_ID" in tore_df.columns else ["Minute"])
    if g.empty:
        return ""
    parts = []
    for _, row in g.iterrows():
        minute = int(row.get("Minute", 0)) if pd.notna(row.get("Minute")) else 0
        stand = str(row.get("Spielstand_Nach_Tor", "")).strip()
        team = str(row.get("Team", "")).strip()
        scorer = str(row.get("Torschütze", "")).strip()
        if team == TEAM_NAME and scorer and scorer != "nan":
            parts.append(f"{minute}. {stand} {scorer}")
        else:
            parts.append(f"{minute}. {stand}")
    return " · ".join(parts)


def aggregate_players(filtered_stats: pd.DataFrame) -> pd.DataFrame:
    if filtered_stats.empty:
        return pd.DataFrame(columns=[
            "Nr", "Spieler", "Einsätze", "Startelf", "Einwechslungen", "Minuten", "Tore", "Assists",
            "Scorer", "Gelb", "Gelb-Rot", "Rot", "Ø Min./Einsatz", "Tore/90", "Assists/90", "Scorer/90", "Startelfquote"
        ])

    grouped = filtered_stats.groupby(["Spieler_ID", "Nr", "Spieler"], dropna=False).agg(
        Einsätze=("Einsatz", "sum"),
        Startelf=("Startelf", "sum"),
        Einwechslungen=("Eingewechselt", "sum"),
        Minuten=("Minuten", "sum"),
        Tore=("Tore", "sum"),
        Assists=("Assists", "sum"),
        Scorer=("Scorer", "sum"),
        Gelb=("Gelb", "sum"),
        **{"Gelb-Rot": ("Gelb-Rot", "sum")},
        Rot=("Rot", "sum"),
    ).reset_index()

    grouped["Ø Min./Einsatz"] = np.where(grouped["Einsätze"] > 0, grouped["Minuten"] / grouped["Einsätze"], 0)
    grouped["Tore/90"] = np.where(grouped["Minuten"] > 0, grouped["Tore"] / grouped["Minuten"] * 90, 0)
    grouped["Assists/90"] = np.where(grouped["Minuten"] > 0, grouped["Assists"] / grouped["Minuten"] * 90, 0)
    grouped["Scorer/90"] = np.where(grouped["Minuten"] > 0, grouped["Scorer"] / grouped["Minuten"] * 90, 0)
    grouped["Startelfquote"] = np.where(grouped["Einsätze"] > 0, grouped["Startelf"] / grouped["Einsätze"], 0)

    num_cols = ["Ø Min./Einsatz", "Tore/90", "Assists/90", "Scorer/90", "Startelfquote"]
    grouped[num_cols] = grouped[num_cols].round(2)
    grouped = grouped.sort_values(["Nr", "Spieler"])
    return grouped[[
        "Nr", "Spieler", "Einsätze", "Startelf", "Einwechslungen", "Minuten", "Tore", "Assists", "Scorer",
        "Gelb", "Gelb-Rot", "Rot", "Ø Min./Einsatz", "Tore/90", "Assists/90", "Scorer/90", "Startelfquote"
    ]]


def make_player_phase_comparison(player_stats: pd.DataFrame) -> pd.DataFrame:
    """Vergleich für einen Spieler: Total Meisterschaft, Vorrunde, Rückrunde und Veränderung."""
    cols = ["Einsatzminuten", "Startelf-Einsätze", "Einwechslungen", "Tore", "Assists"]
    ms = player_stats[player_stats["Wettbewerb"].astype(str).str.strip().eq("Meisterschaft")].copy()

    def vals(df: pd.DataFrame) -> dict:
        if df.empty:
            return {c: 0 for c in cols}
        return {
            "Einsatzminuten": int(df["Minuten"].sum()),
            "Startelf-Einsätze": int(df["Startelf"].sum()),
            "Einwechslungen": int(df["Eingewechselt"].sum()),
            "Tore": int(df["Tore"].sum()),
            "Assists": int(df["Assists"].sum()),
        }

    total = vals(ms)
    vor = vals(ms[ms["Phase"].astype(str).str.strip().eq("Vorrunde")])
    rueck = vals(ms[ms["Phase"].astype(str).str.strip().eq("Rückrunde")])

    rows = []
    for c in cols:
        diff = rueck[c] - vor[c]
        rows.append({
            "Kennzahl": c,
            "Total Meisterschaft": total[c],
            "Vorrunde": vor[c],
            "Rückrunde": rueck[c],
            "Veränderung": diff,
        })
    return pd.DataFrame(rows)


def format_change(value: int | float) -> str:
    try:
        value = int(value)
    except Exception:
        return str(value)
    if value > 0:
        return f"+{value}"
    return str(value)


def aggregate_games(filtered_games: pd.DataFrame) -> dict:
    if filtered_games.empty:
        return {
            "spiele": 0, "siege": 0, "unentschieden": 0, "niederlagen": 0, "punkte": 0,
            "tore": 0, "gegentore": 0, "tordiff": 0, "schnitt_tore": 0.0, "schnitt_gegentore": 0.0,
        }
    spiele_count = len(filtered_games)
    tore = int(filtered_games["Tore_FCE"].sum())
    gegentore = int(filtered_games["Tore_Gegner"].sum())
    meisterschaft = filtered_games[filtered_games["Wettbewerb"] == "Meisterschaft"]
    return {
        "spiele": spiele_count,
        "siege": int((filtered_games["Ergebnis"] == "Sieg").sum()),
        "unentschieden": int((filtered_games["Ergebnis"] == "Unentschieden").sum()),
        "niederlagen": int((filtered_games["Ergebnis"] == "Niederlage").sum()),
        "punkte": int(pd.to_numeric(meisterschaft.get("Punkte", 0), errors="coerce").fillna(0).sum()),
        "tore": tore,
        "gegentore": gegentore,
        "tordiff": tore - gegentore,
        "schnitt_tore": round(tore / spiele_count, 2),
        "schnitt_gegentore": round(gegentore / spiele_count, 2),
    }



def _game_minutes_state(tore_game: pd.DataFrame, heim_aus: str) -> dict:
    """Berechnet Minuten in Führung/Rückstand/Unentschieden für ein Spiel über 90 Minuten."""
    fce_goals = 0
    opp_goals = 0
    last_minute = 0
    lead = 0
    draw = 0
    trail = 0

    g = tore_game.copy()
    if g.empty:
        return {"Minuten Führung": 0, "Minuten Unentschieden": 90, "Minuten Rückstand": 0}

    sort_cols = [c for c in ["Minute", "Tor_ID"] if c in g.columns]
    g = g.sort_values(sort_cols if sort_cols else ["Minute"])

    def add_duration(duration: int):
        nonlocal lead, draw, trail, fce_goals, opp_goals
        duration = max(0, int(duration))
        if fce_goals > opp_goals:
            lead += duration
        elif fce_goals < opp_goals:
            trail += duration
        else:
            draw += duration

    for _, row in g.iterrows():
        minute = int(row.get("Minute", 0) or 0)
        minute = max(0, min(90, minute))
        add_duration(minute - last_minute)

        team = str(row.get("Team", "")).strip()
        if team == TEAM_NAME or team == "FC Entlebuch":
            fce_goals += 1
        elif team:
            opp_goals += 1
        last_minute = minute

    add_duration(90 - last_minute)
    return {"Minuten Führung": lead, "Minuten Unentschieden": draw, "Minuten Rückstand": trail}


def aggregate_team_overview(games_df: pd.DataFrame, stats_df: pd.DataFrame, goals_df: pd.DataFrame) -> dict:
    """Erweiterte Team-Kennzahlen für eine Spielauswahl."""
    if games_df.empty:
        return {
            "Spiele": 0, "Bilanz": "0-0-0", "Siege": 0, "Unentschieden": 0, "Niederlagen": 0,
            "Tore": 0, "Gegentore": 0, "Tordifferenz": 0, "Ø Tore": 0.0, "Ø Gegentore": 0.0,
            "Minuten Führung": 0, "Minuten Rückstand": 0, "Minuten Unentschieden": 0,
            "Spiele ohne Gegentor": 0, "Spiele ohne eigenes Tor": 0, "Eingesetzte Spieler": 0,
        }

    game_ids = set(games_df["Spiel_ID"].astype(int).tolist())
    g_stats = stats_df[stats_df["Spiel_ID"].isin(game_ids)] if not stats_df.empty else pd.DataFrame()
    g_goals = goals_df[goals_df["Spiel_ID"].isin(game_ids)] if not goals_df.empty else pd.DataFrame()

    minutes = {"Minuten Führung": 0, "Minuten Unentschieden": 0, "Minuten Rückstand": 0}
    for _, game in games_df.iterrows():
        game_goals = g_goals[g_goals["Spiel_ID"] == int(game["Spiel_ID"])]
        m = _game_minutes_state(game_goals, str(game.get("Heim/Auswärts", "")))
        for k, v in m.items():
            minutes[k] += int(v)

    spiele_count = int(len(games_df))
    tore = int(pd.to_numeric(games_df["Tore_FCE"], errors="coerce").fillna(0).sum())
    gegentore = int(pd.to_numeric(games_df["Tore_Gegner"], errors="coerce").fillna(0).sum())
    siege = int((games_df["Ergebnis"] == "Sieg").sum())
    remis = int((games_df["Ergebnis"] == "Unentschieden").sum())
    niederlagen = int((games_df["Ergebnis"] == "Niederlage").sum())

    return {
        "Spiele": spiele_count,
        "Bilanz": f"{siege}-{remis}-{niederlagen}",
        "Siege": siege,
        "Unentschieden": remis,
        "Niederlagen": niederlagen,
        "Tore": tore,
        "Gegentore": gegentore,
        "Tordifferenz": tore - gegentore,
        "Ø Tore": round(tore / spiele_count, 2) if spiele_count else 0.0,
        "Ø Gegentore": round(gegentore / spiele_count, 2) if spiele_count else 0.0,
        "Minuten Führung": minutes["Minuten Führung"],
        "Minuten Rückstand": minutes["Minuten Rückstand"],
        "Minuten Unentschieden": minutes["Minuten Unentschieden"],
        "Spiele ohne Gegentor": int((pd.to_numeric(games_df["Tore_Gegner"], errors="coerce").fillna(0) == 0).sum()),
        "Spiele ohne eigenes Tor": int((pd.to_numeric(games_df["Tore_FCE"], errors="coerce").fillna(0) == 0).sum()),
        "Eingesetzte Spieler": int(g_stats.loc[pd.to_numeric(g_stats.get("Minuten", 0), errors="coerce").fillna(0) > 0, "Spieler_ID"].nunique()) if not g_stats.empty else 0,
    }


def overview_table(games_subset: pd.DataFrame, stats_df: pd.DataFrame, goals_df: pd.DataFrame, label: str) -> pd.DataFrame:
    d = aggregate_team_overview(games_subset, stats_df, goals_df)
    row = {"Bereich": label, **d}
    return pd.DataFrame([row])


def render_metric_card(title: str, value, sub: str = ""):
    st.markdown(
        f"""
        <div class="fce-card">
            <div class="fce-card-title">{title}</div>
            <div class="fce-card-value">{value}</div>
            <div class="fce-card-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_overview_section(title: str, data: dict):
    st.markdown(f"### {title}")
    metrics = [
        ("Spiele", data["Spiele"], "Anzahl Spiele"),
        ("Bilanz", data["Bilanz"], "Siege · Remis · Niederlagen"),
        ("Tore", f"{data['Tore']}:{data['Gegentore']}", f"Tordifferenz {data['Tordifferenz']:+d}"),
        ("Ø Tore", data["Ø Tore"], "pro Spiel"),
        ("Ø Gegentore", data["Ø Gegentore"], "pro Spiel"),
        ("Min. Führung", data["Minuten Führung"], "von 90 Minuten je Spiel"),
        ("Min. Rückstand", data["Minuten Rückstand"], "von 90 Minuten je Spiel"),
        ("Min. Unentschieden", data["Minuten Unentschieden"], "inkl. 0:0-Phasen"),
        ("Ohne Gegentor", data["Spiele ohne Gegentor"], "Clean Sheets"),
        ("Ohne eigenes Tor", data["Spiele ohne eigenes Tor"], "Spiele ohne FCE-Tor"),
        ("Eingesetzte Spieler", data["Eingesetzte Spieler"], "mind. 1 Minute"),
    ]
    cards = []
    for title_, value, sub in metrics:
        cards.append(
            f'<div class="fce-card">'
            f'<div class="fce-card-title">{title_}</div>'
            f'<div class="fce-card-value">{value}</div>'
            f'<div class="fce-card-sub">{sub}</div>'
            f'</div>'
        )
    st.markdown('<div class="fce-metric-grid">' + ''.join(cards) + '</div>', unsafe_allow_html=True)


def make_overview_comparison_table(ms_games: pd.DataFrame, stats_df: pd.DataFrame, goals_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    total = aggregate_team_overview(ms_games, stats_df, goals_df)
    vor = aggregate_team_overview(ms_games[ms_games["Phase"] == "Vorrunde"], stats_df, goals_df)
    rueck = aggregate_team_overview(ms_games[ms_games["Phase"] == "Rückrunde"], stats_df, goals_df)

    rows = []
    for label, d in [("Gesamte Meisterschaft", total), ("Vorrunde", vor), ("Rückrunde", rueck)]:
        rows.append({
            "Bereich": label,
            "Spiele": d["Spiele"],
            "Bilanz": d["Bilanz"],
            "Tore": d["Tore"],
            "Gegentore": d["Gegentore"],
            "Tordiff": d["Tordifferenz"],
            "Ø Tore": d["Ø Tore"],
            "Ø Gegentore": d["Ø Gegentore"],
            "Min. Führung": d["Minuten Führung"],
            "Min. Rückstand": d["Minuten Rückstand"],
            "Min. Unentschieden": d["Minuten Unentschieden"],
            "Ohne Gegentor": d["Spiele ohne Gegentor"],
            "Ohne eigenes Tor": d["Spiele ohne eigenes Tor"],
            "Eingesetzte Spieler": d["Eingesetzte Spieler"],
        })
    table = pd.DataFrame(rows)

    diff_metrics = [
        ("Spiele", "Spiele"), ("Tore", "Tore"), ("Gegentore", "Gegentore"), ("Tordiff", "Tordifferenz"),
        ("Ø Tore", "Ø Tore"), ("Ø Gegentore", "Ø Gegentore"), ("Min. Führung", "Minuten Führung"),
        ("Min. Rückstand", "Minuten Rückstand"), ("Min. Unentschieden", "Minuten Unentschieden"),
        ("Ohne Gegentor", "Spiele ohne Gegentor"), ("Ohne eigenes Tor", "Spiele ohne eigenes Tor"),
        ("Eingesetzte Spieler", "Eingesetzte Spieler"),
    ]
    diffs = []
    for display, key in diff_metrics:
        diffs.append({"Kennzahl": display, "Vorrunde": vor[key], "Rückrunde": rueck[key], "Differenz Rückrunde - Vorrunde": round(rueck[key] - vor[key], 2)})
    diff_table = pd.DataFrame(diffs)
    return table, diff_table


def make_home_away_comparison_table(games_df: pd.DataFrame, stats_df: pd.DataFrame, goals_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Erstellt Heim-/Auswärtsvergleich mit denselben Kennzahlen wie die Übersicht."""
    heim = aggregate_team_overview(games_df[games_df["Heim/Auswärts"].astype(str).str.strip().eq("Heim")], stats_df, goals_df)
    aus = aggregate_team_overview(games_df[games_df["Heim/Auswärts"].astype(str).str.strip().eq("Auswärts")], stats_df, goals_df)

    def row(label: str, d: dict) -> dict:
        return {
            "Bereich": label,
            "Spiele": d["Spiele"],
            "Bilanz": d["Bilanz"],
            "Tore": d["Tore"],
            "Gegentore": d["Gegentore"],
            "Tordiff": d["Tordifferenz"],
            "Ø Tore": d["Ø Tore"],
            "Ø Gegentore": d["Ø Gegentore"],
            "Min. Führung": d["Minuten Führung"],
            "Min. Rückstand": d["Minuten Rückstand"],
            "Min. Unentschieden": d["Minuten Unentschieden"],
            "Ohne Gegentor": d["Spiele ohne Gegentor"],
            "Ohne eigenes Tor": d["Spiele ohne eigenes Tor"],
            "Eingesetzte Spieler": d["Eingesetzte Spieler"],
        }

    table = pd.DataFrame([row("Heim", heim), row("Auswärts", aus)])
    diff_metrics = [
        ("Spiele", "Spiele"), ("Tore", "Tore"), ("Gegentore", "Gegentore"),
        ("Tordiff", "Tordifferenz"), ("Ø Tore", "Ø Tore"), ("Ø Gegentore", "Ø Gegentore"),
        ("Min. Führung", "Minuten Führung"), ("Min. Rückstand", "Minuten Rückstand"),
        ("Min. Unentschieden", "Minuten Unentschieden"), ("Ohne Gegentor", "Spiele ohne Gegentor"),
        ("Ohne eigenes Tor", "Spiele ohne eigenes Tor"), ("Eingesetzte Spieler", "Eingesetzte Spieler"),
    ]
    diff_rows = []
    for display, key in diff_metrics:
        diff_rows.append({
            "Kennzahl": display,
            "Heim": heim[key],
            "Auswärts": aus[key],
            "Differenz Auswärts - Heim": round(aus[key] - heim[key], 2),
        })
    diff_table = pd.DataFrame(diff_rows)
    return table, diff_table


def render_home_away_section(title: str, games_df: pd.DataFrame, stats_df: pd.DataFrame, goals_df: pd.DataFrame):
    st.markdown(f"### {title}")
    heim = aggregate_team_overview(games_df[games_df["Heim/Auswärts"].astype(str).str.strip().eq("Heim")], stats_df, goals_df)
    aus = aggregate_team_overview(games_df[games_df["Heim/Auswärts"].astype(str).str.strip().eq("Auswärts")], stats_df, goals_df)

    col_home, col_away = st.columns(2)
    with col_home:
        render_overview_section("Heimspiele", heim)
    with col_away:
        render_overview_section("Auswärtsspiele", aus)

    table, diff_table = make_home_away_comparison_table(games_df, stats_df, goals_df)
    st.markdown("**Heim/Auswärts-Tabelle**")
    st.dataframe(table, use_container_width=True, hide_index=True)
    st.markdown("**Differenz Auswärts - Heim**")
    st.caption("Positive Werte bedeuten: auswärts höher als zuhause. Bei Gegentoren/Rückstand ist ein tieferer Wert natürlich besser.")
    st.dataframe(diff_table, use_container_width=True, hide_index=True)

def show_metric_row(summary: dict):
    cols = st.columns(6)
    cols[0].metric("Spiele", summary["spiele"])
    cols[1].metric("Bilanz", f"{summary['siege']}-{summary['unentschieden']}-{summary['niederlagen']}", help="Siege-Unentschieden-Niederlagen")
    cols[2].metric("Punkte MS", summary["punkte"], help="Nur Meisterschaftsspiele")
    cols[3].metric("Tore", f"{summary['tore']}:{summary['gegentore']}", delta=summary["tordiff"])
    cols[4].metric("Ø Tore", summary["schnitt_tore"])
    cols[5].metric("Ø Gegentore", summary["schnitt_gegentore"])


def make_goal_bins(tore_df: pd.DataFrame) -> pd.DataFrame:
    if tore_df.empty:
        return pd.DataFrame(columns=["Zeitfenster", "Team", "Tore"])
    bins = [0, 15, 30, 45, 60, 75, 90, 200]
    labels = ["0–15", "16–30", "31–45", "46–60", "61–75", "76–90", "90+"]
    d = tore_df.copy()
    d["Zeitfenster"] = pd.cut(d["Minute"], bins=bins, labels=labels, include_lowest=True, right=True)
    d["Team"] = d["Team"].replace({TEAM_NAME: "FC Entlebuch"}).fillna("Unbekannt")
    agg = d.groupby(["Zeitfenster", "Team"], observed=False).size().reset_index(name="Tore")
    return agg[agg["Tore"] > 0]


def style_figure(fig, height: int | None = None):
    """Einheitliches FC-Entlebuch-Layout für Plotly-Grafiken."""
    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color=FCE_BLACK),
        title=dict(font=dict(size=20, color=FCE_BLACK)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=20, r=20, t=65, b=35),
    )
    fig.update_xaxes(showgrid=False, linecolor="#e5e7eb")
    fig.update_yaxes(gridcolor="#edf2f0", zerolinecolor="#e5e7eb")
    if height:
        fig.update_layout(height=height)
    return fig


try:
    spieler, spiele, stats, tore, stats_full, tore_full = load_data(DATA_FILE)
except Exception as exc:
    st.error(f"Die Daten konnten nicht geladen werden: {exc}")
    st.stop()

st.markdown(
    f"""
    <div class="fce-hero">
        <div class="fce-kicker">FC Entlebuch · Saisonstatistik</div>
        <div class="fce-title">Spielerstatistik 2025/26</div>
        <div class="fce-subtitle">Interaktive Saisonübersicht mit Teamkennzahlen, Spielerstatistiken, Spielübersicht und Tor-Zeitpunkten.</div>
        <div class="fce-pill">⚽ FCE-Dashboard · Auf dem Handy Menü oben links öffnen</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Sidebar-Filter
st.sidebar.markdown("## ⚽ Filter")

wettbewerbe = sorted([x for x in spiele["Wettbewerb"].dropna().unique() if str(x).strip()])
phasen = sorted([x for x in spiele["Phase"].dropna().unique() if str(x).strip()])
heim_aus = sorted([x for x in spiele["Heim/Auswärts"].dropna().unique() if str(x).strip()])
gegner = sorted([x for x in spiele["Gegner"].dropna().unique() if str(x).strip()])

selected_wettbewerbe = st.sidebar.multiselect("Wettbewerb", wettbewerbe, default=wettbewerbe)
selected_phasen = st.sidebar.multiselect("Phase", phasen, default=phasen)
selected_heim_aus = st.sidebar.multiselect("Heim/Auswärts", heim_aus, default=heim_aus)
selected_gegner = st.sidebar.multiselect("Gegner", gegner, default=gegner)

filtered_games = spiele[
    spiele["Wettbewerb"].isin(selected_wettbewerbe)
    & spiele["Phase"].isin(selected_phasen)
    & spiele["Heim/Auswärts"].isin(selected_heim_aus)
    & spiele["Gegner"].isin(selected_gegner)
].copy()
filtered_game_ids = set(filtered_games["Spiel_ID"].tolist())
filtered_stats = stats_full[stats_full["Spiel_ID"].isin(filtered_game_ids)].copy()
filtered_tore = tore_full[tore_full["Spiel_ID"].isin(filtered_game_ids)].copy()

summary = aggregate_games(filtered_games)
show_metric_row(summary)

st.divider()

tab_overview, tab_players, tab_games, tab_goals, tab_profile = st.tabs([
    "🏠 Übersicht", "👥 Spieler", "📅 Spiele", "⚽ Tore", "🔎 Profil"
])

with tab_overview:
    st.subheader("Übersicht")
    st.markdown(
        """
        <div class="fce-section-note">
        Die Übersicht ist bewusst als Startseite gebaut: zuerst die gesamte Meisterschaft, dann der direkte Vergleich Vorrunde/Rückrunde und danach der Cup separat. Die Filter in der Sidebar gelten für die übrigen Tabs; diese Startseite zeigt immer die fixen Saisonblöcke.
        </div>
        """,
        unsafe_allow_html=True,
    )

    ms_games = spiele[spiele["Wettbewerb"] == "Meisterschaft"].copy()
    cup_games = spiele[spiele["Wettbewerb"] == "Cup"].copy()

    overview_tab_ms, overview_tab_halves, overview_tab_homeaway, overview_tab_cup, overview_tab_games = st.tabs([
        "Meisterschaft", "Vorrunde/Rückrunde", "Heim/Auswärts", "Cup", "Spiele"
    ])

    with overview_tab_ms:
        ms_total = aggregate_team_overview(ms_games, stats_full, tore_full)
        render_overview_section("Meisterschaft · Gesamt", ms_total)

    with overview_tab_halves:
        st.markdown("### Vorrunde vs. Rückrunde")
        overview_df, diff_df = make_overview_comparison_table(ms_games, stats_full, tore_full)
        st.markdown('<div class="fce-mini-table-note">Die Differenz zeigt jeweils: Rückrunde minus Vorrunde.</div>', unsafe_allow_html=True)
        st.dataframe(overview_df, use_container_width=True, hide_index=True)
        st.write("**Differenz Rückrunde - Vorrunde**")
        st.dataframe(diff_df, use_container_width=True, hide_index=True)

    with overview_tab_homeaway:
        st.markdown("### Heim- und Auswärtsvergleich")
        st.markdown(
            '<div class="fce-section-note">Dieser Bereich zeigt die gleichen Kennzahlen wie die Startübersicht, getrennt nach Heim- und Auswärtsspielen.</div>',
            unsafe_allow_html=True,
        )
        ha_tab_all, ha_tab_ms, ha_tab_cup = st.tabs(["Alle Pflichtspiele", "Meisterschaft", "Cup"])
        with ha_tab_all:
            render_home_away_section("Alle Pflichtspiele", spiele, stats_full, tore_full)
        with ha_tab_ms:
            render_home_away_section("Meisterschaft", ms_games, stats_full, tore_full)
        with ha_tab_cup:
            render_home_away_section("Cup", cup_games, stats_full, tore_full)

    with overview_tab_cup:
        cup_total = aggregate_team_overview(cup_games, stats_full, tore_full)
        render_overview_section("Cup", cup_total)

    with overview_tab_games:
        st.markdown("### Meisterschaft: Spiele im Überblick")
        ms_view = ms_games.sort_values(["Datum", "Spiel_ID"]).copy()
        if not ms_view.empty:
            ms_view["Datum"] = ms_view["Datum"].dt.strftime("%d.%m.%Y")
            ms_view["Spiel"] = ms_view["Heimteam"].astype(str) + " - " + ms_view["Auswärtsteam"].astype(str)
            ms_view["Torfolge"] = ms_view["Spiel_ID"].apply(lambda sid: torfolge_for_game(tore_full, sid))
            st.dataframe(
                ms_view[["Datum", "Runde", "Phase", "Heim/Auswärts", "Spiel", "Resultat", "Resultat Halbzeit", "Ergebnis", "Punkte", "Torfolge"]],
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("### Cup: Spiele im Überblick")
        cup_view = cup_games.sort_values(["Datum", "Spiel_ID"]).copy()
        if not cup_view.empty:
            cup_view["Datum"] = cup_view["Datum"].dt.strftime("%d.%m.%Y")
            cup_view["Spiel"] = cup_view["Heimteam"].astype(str) + " - " + cup_view["Auswärtsteam"].astype(str)
            cup_view["Torfolge"] = cup_view["Spiel_ID"].apply(lambda sid: torfolge_for_game(tore_full, sid))
            st.dataframe(
                cup_view[["Datum", "Runde", "Heim/Auswärts", "Spiel", "Resultat", "Resultat Halbzeit", "Ergebnis", "Torfolge"]],
                use_container_width=True,
                hide_index=True,
            )


with tab_players:
    st.subheader("Spielerstatistik")
    player_agg = aggregate_players(filtered_stats)
    st.dataframe(
        player_agg,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Startelfquote": st.column_config.ProgressColumn("Startelfquote", format="%.0f%%", min_value=0, max_value=1),
            "Tore/90": st.column_config.NumberColumn("Tore/90", format="%.2f"),
            "Assists/90": st.column_config.NumberColumn("Assists/90", format="%.2f"),
            "Scorer/90": st.column_config.NumberColumn("Scorer/90", format="%.2f"),
        },
    )

    st.download_button(
        "Spielerstatistik als CSV herunterladen",
        player_agg.to_csv(index=False).encode("utf-8-sig"),
        file_name="fc_entlebuch_spielerstatistik.csv",
        mime="text/csv",
    )

with tab_games:
    st.subheader("Spielübersicht")
    games_view = filtered_games.sort_values("Datum").copy()
    games_view["Datum"] = games_view["Datum"].dt.strftime("%d.%m.%Y")
    games_view["Spiel"] = games_view["Heimteam"].astype(str) + " - " + games_view["Auswärtsteam"].astype(str)
    games_view["Torfolge"] = games_view["Spiel_ID"].apply(lambda sid: torfolge_for_game(tore_full, sid))
    cols = ["Datum", "Wettbewerb", "Runde", "Phase", "Heim/Auswärts", "Spiel", "Resultat", "Resultat Halbzeit", "Ergebnis", "Punkte", "Torfolge"]
    st.dataframe(games_view[cols], use_container_width=True, hide_index=True)

    st.subheader("Spieldetails")
    if not filtered_games.empty:
        label_df = filtered_games.sort_values("Datum").copy()
        label_df["Label"] = label_df.apply(make_game_label, axis=1)
        selected_label = st.selectbox("Spiel auswählen", label_df["Label"].tolist())
        selected_game_id = int(label_df.loc[label_df["Label"] == selected_label, "Spiel_ID"].iloc[0])

        selected_game = spiele[spiele["Spiel_ID"] == selected_game_id].iloc[0]
        st.write(f"**{selected_game['Heimteam']} - {selected_game['Auswärtsteam']} {selected_game['Resultat']}**")
        st.caption(f"{selected_game['Wettbewerb']} · Runde {selected_game['Runde']} · {selected_game['Datum'].strftime('%d.%m.%Y') if pd.notna(selected_game['Datum']) else ''}")

        detail_stats = stats_full[stats_full["Spiel_ID"] == selected_game_id].sort_values(["Startelf", "Minuten", "Nr"], ascending=[False, False, True])
        eingesetzte = detail_stats[detail_stats["Minuten"] > 0].copy()
        nicht_eingesetzt = detail_stats[detail_stats["Minuten"] == 0].sort_values(["Nr", "Spieler"]).copy()

        st.write("**Eingesetzte Spieler**")
        st.dataframe(eingesetzte[["Nr", "Spieler", "Status", "Minuten", "Tore", "Assists", "Scorer", "Gelb", "Gelb-Rot", "Rot"]], use_container_width=True, hide_index=True)

        if not nicht_eingesetzt.empty:
            with st.expander(f"Nicht eingesetzt ({len(nicht_eingesetzt)})"):
                st.dataframe(nicht_eingesetzt[["Nr", "Spieler", "Status", "Minuten"]], use_container_width=True, hide_index=True)

        goal_detail = tore_full[tore_full["Spiel_ID"] == selected_game_id].sort_values("Minute")
        if not goal_detail.empty:
            st.write("**Torfolge**")
            st.dataframe(goal_detail[["Minute", "Team", "Torschütze", "Spielstand_Nach_Tor"]], use_container_width=True, hide_index=True)

with tab_goals:
    st.subheader("Tore & Zeitpunkte")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        goal_bins = make_goal_bins(filtered_tore)
        if not goal_bins.empty:
            fig = px.bar(goal_bins, x="Zeitfenster", y="Tore", color="Team", barmode="group", title="Tore nach Spielabschnitt", color_discrete_map={"FC Entlebuch": FCE_GREEN, TEAM_NAME: FCE_GREEN, "Gegner": FCE_RED})
            fig.update_layout(legend_title_text="")
            fig = style_figure(fig, height=420)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Keine Tore für die gewählten Filter vorhanden.")

    with c2:
        if not filtered_tore.empty:
            team_goals = filtered_tore.groupby("Team").size().reset_index(name="Tore").sort_values("Tore", ascending=False)
            st.dataframe(team_goals, use_container_width=True, hide_index=True)
            fce_goals = filtered_tore[filtered_tore["Team"] == TEAM_NAME]
            scorers = fce_goals.groupby(["Nr", "Torschütze"]).size().reset_index(name="Tore").sort_values("Tore", ascending=False)
            st.write("**Torschützen nach Tor-Tabelle**")
            st.dataframe(scorers, use_container_width=True, hide_index=True)

    st.write("**Alle Tore**")
    goals_view = filtered_tore.sort_values(["Datum", "Minute"]).copy()
    if not goals_view.empty:
        goals_view["Datum"] = goals_view["Datum"].dt.strftime("%d.%m.%Y")
        st.dataframe(goals_view[["Datum", "Wettbewerb", "Phase", "Gegner", "Minute", "Team", "Torschütze", "Spielstand_Nach_Tor"]], use_container_width=True, hide_index=True)

with tab_profile:
    st.subheader("Spielerprofil")
    player_options = spieler.sort_values(["Nr", "Spieler"])[["Spieler_ID", "Nr", "Spieler"]].copy()
    player_options["Label"] = player_options["Nr"].astype(str) + " – " + player_options["Spieler"]
    selected_player_label = st.selectbox("Spieler auswählen", player_options["Label"].tolist())
    selected_player_id = int(player_options.loc[player_options["Label"] == selected_player_label, "Spieler_ID"].iloc[0])

    p_stats = filtered_stats[filtered_stats["Spieler_ID"] == selected_player_id].copy()
    p_agg = aggregate_players(p_stats)
    if p_agg.empty:
        st.info("Für diesen Spieler gibt es im aktuellen Filter keine Spiele.")
    else:
        row = p_agg.iloc[0]
        pc = st.columns(6)
        pc[0].metric("Einsätze", int(row["Einsätze"]))
        pc[1].metric("Startelf", int(row["Startelf"]))
        pc[2].metric("Minuten", int(row["Minuten"]))
        pc[3].metric("Tore", int(row["Tore"]))
        pc[4].metric("Assists", int(row["Assists"]))
        pc[5].metric("Scorer", int(row["Scorer"]))

        if int(row["Einsätze"]) == 0:
            st.info("Dieser Spieler hat im aktuellen Filter keine Einsatzminuten. Die Spiele werden trotzdem unten aufgeführt.")

        st.markdown("### Vorrunde vs. Rückrunde")
        st.caption("Vergleich für die Meisterschaft. Cupspiele sind hier bewusst nicht enthalten.")
        comparison_df = make_player_phase_comparison(p_stats)
        display_comparison = comparison_df.copy()
        display_comparison["Veränderung"] = display_comparison["Veränderung"].apply(format_change)
        st.dataframe(display_comparison, use_container_width=True, hide_index=True)

        mobile_rows = []
        for _, comp_row in comparison_df.iterrows():
            mobile_rows.append(
                f'<div class="fce-card">'
                f'<div class="fce-card-title">{comp_row["Kennzahl"]}</div>'
                f'<div class="fce-card-value">{comp_row["Total Meisterschaft"]}</div>'
                f'<div class="fce-card-sub">VR {comp_row["Vorrunde"]} · RR {comp_row["Rückrunde"]} · Veränderung {format_change(comp_row["Veränderung"])}</div>'
                f'</div>'
            )
        st.markdown('<div class="fce-metric-grid">' + ''.join(mobile_rows) + '</div>', unsafe_allow_html=True)

        p_games = p_stats.sort_values("Datum").copy()
        p_games["Datum"] = p_games["Datum"].dt.strftime("%d.%m.%Y")
        p_games["Spiel"] = p_games["Heimteam"].astype(str) + " - " + p_games["Auswärtsteam"].astype(str)
        st.dataframe(p_games[["Datum", "Wettbewerb", "Runde", "Phase", "Spiel", "Resultat", "Status", "Minuten", "Tore", "Assists", "Scorer", "Gelb", "Gelb-Rot", "Rot"]], use_container_width=True, hide_index=True)

        trend = p_stats.sort_values("Datum").copy()
        trend = trend[pd.notna(trend["Datum"])]
        if not trend.empty:
            trend["Spiel"] = trend["Datum"].dt.strftime("%d.%m.") + " " + trend["Gegner"].astype(str)
            fig = px.bar(trend, x="Spiel", y="Minuten", title=f"Einsatzminuten: {row['Spieler']}", color_discrete_sequence=[FCE_GREEN])
            fig.update_layout(xaxis_tickangle=-35)
            fig = style_figure(fig, height=330)
            st.plotly_chart(fig, use_container_width=True)

st.sidebar.divider()
st.sidebar.caption(f"Datenquelle: {DATA_FILE.name}")
st.sidebar.caption("Version: 1.3.4 · Spielerprofil-Vergleich")
