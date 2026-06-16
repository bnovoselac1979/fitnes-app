# streamlit_app_fitnes_gsheets.py
# Fitnes klub aplikacija preko Google Sheets baze
# V6: testiranje svih vežbi + plan + admin panel + A4 izveštaj vežbača

from __future__ import annotations

import re
from datetime import datetime, date
from typing import Dict, List, Optional, Tuple, Any

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ------------------------------------------------------------
# 1) OSNOVNA PODESAVANJA
# ------------------------------------------------------------

APP_TITLE = "Fitnes klub — 5/3/1 + RIR"

SHEET_VEZBACI = "VEZBACI"
SHEET_BAZA_VEZBI = "BAZA_VEZBI"
SHEET_PODESAVANJA = "PODESAVANJA"
SHEET_TEST_INPUT = "UNOS_TESTOVA"
SHEET_DNEVNIK = "DNEVNIK_UNOS"
SHEET_PLAN = "A4_MESEC_KOMPAKT"

HEADER_ROWS = {
    SHEET_VEZBACI: 2,
    SHEET_BAZA_VEZBI: 2,
    SHEET_TEST_INPUT: 2,
    SHEET_DNEVNIK: 3,
    SHEET_PLAN: 4,
}

REQUIRED_VEZBACI_COLS = ["ID", "Ime i Prezime", "Aktivan", "Email", "Sifra", "Poruka_Blokada"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ROUND_KG_STEP_DEFAULT = 2.5

DEFAULT_GOAL_SETTINGS = {
    "Snaga": {"TM %": 0.90, "RIR Cilj": 3, "Serije+": 0, "Kg Faktor": 1.00, "RIR Asist.": 3, "Zona Asist.": "5-10"},
    "Hipertrofija": {"TM %": 0.875, "RIR Cilj": 2, "Serije+": 1, "Kg Faktor": 0.97, "RIR Asist.": 2, "Zona Asist.": "10-20"},
    "Snaga+Hipertrofija": {"TM %": 0.90, "RIR Cilj": 2, "Serije+": 0, "Kg Faktor": 1.00, "RIR Asist.": 2, "Zona Asist.": "10-20"},
    "Rekompozicija": {"TM %": 0.85, "RIR Cilj": 3, "Serije+": 0, "Kg Faktor": 0.97, "RIR Asist.": 3, "Zona Asist.": "10-20"},
    "Sportista": {"TM %": 0.85, "RIR Cilj": 4, "Serije+": -1, "Kg Faktor": 0.95, "RIR Asist.": 4, "Zona Asist.": "5-10"},
    "Početnik": {"TM %": 0.85, "RIR Cilj": 4, "Serije+": -1, "Kg Faktor": 0.95, "RIR Asist.": 4, "Zona Asist.": "10-20"},
}

DEFAULT_531_WEEKS = {
    1: [(0.65, "5"), (0.75, "5"), (0.85, "5+")],
    2: [(0.70, "3"), (0.80, "3"), (0.90, "3+")],
    3: [(0.75, "5"), (0.85, "3"), (0.95, "1+")],
    4: [(0.40, "5"), (0.50, "5"), (0.60, "5")],
}

ASSIST_WEEK_SETS = {1: "2", 2: "2", 3: "3", 4: "1-2"}
ASSIST_WEEK_FACTOR = {1: 1.00, 2: 1.00, 3: 0.98, 4: 0.85}

# ------------------------------------------------------------
# 2) POMOCNE FUNKCIJE
# ------------------------------------------------------------

def normalize_text(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip()


def normalize_email(x) -> str:
    return normalize_text(x).lower()


def normalize_sifra(x) -> str:
    s = normalize_text(x)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def to_float(x, default: float = 0.0) -> float:
    s = normalize_text(x).replace(",", ".")
    if not s:
        return default
    try:
        return float(s)
    except Exception:
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group(0)) if m else default


def to_int(x, default: int = 0) -> int:
    try:
        return int(float(str(x).replace(",", ".")))
    except Exception:
        return default


def estimate_1rm(kg: float, reps: int) -> float:
    kg = float(kg or 0)
    reps = int(reps or 1)
    if reps <= 1:
        return round(kg, 1)
    return round(kg * (1 + reps / 30), 1)


def round_to_step(value: float, step: float = ROUND_KG_STEP_DEFAULT) -> float:
    if not value:
        return 0.0
    return round(round(float(value) / step) * step, 2)


def format_kg(x: Any) -> str:
    if x is None or x == "":
        return ""
    val = to_float(x, 0)
    if val == int(val):
        return str(int(val))
    return str(val).replace(".", ",")


def parse_range(value: Any, fallback: Tuple[int, int] = (10, 15)) -> Tuple[int, int]:
    text = normalize_text(value)
    nums = [int(n) for n in re.findall(r"\d+", text)]
    if len(nums) >= 2:
        return nums[0], nums[1]
    if len(nums) == 1:
        return nums[0], nums[0]
    return fallback


def to_sheet_value(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return value

# ------------------------------------------------------------
# 3) GOOGLE SHEETS KONEKCIJA
# ------------------------------------------------------------

def _service_account_info() -> Dict[str, str]:
    if "gcp_service_account" in st.secrets:
        return dict(st.secrets["gcp_service_account"])

    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        conn_secret = dict(st.secrets["connections"]["gsheets"])
        conn_secret.pop("spreadsheet", None)
        conn_secret.pop("worksheet", None)
        return conn_secret

    st.error("Nisu pronađeni Google service account podaci u Streamlit secrets.")
    st.stop()


def _spreadsheet_url() -> str:
    if "gsheets" in st.secrets and "spreadsheet_url" in st.secrets["gsheets"]:
        return st.secrets["gsheets"]["spreadsheet_url"]

    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        spreadsheet = st.secrets["connections"]["gsheets"].get("spreadsheet", "")
        if spreadsheet:
            return spreadsheet

    st.error("Nije pronađen spreadsheet URL u secrets.")
    st.stop()


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    info = _service_account_info()
    if "private_key" in info:
        info["private_key"] = info["private_key"].replace("\\n", "\n")
    credentials = Credentials.from_service_account_info(info, scopes=SCOPES)
    client = gspread.authorize(credentials)
    return client.open_by_url(_spreadsheet_url())


def get_ws(sheet_name: str):
    sh = get_spreadsheet()
    try:
        return sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        st.error(f"U Google Sheet fajlu ne postoji tab/list: {sheet_name}")
        st.stop()


@st.cache_data(ttl=10, show_spinner=False)
def read_sheet(sheet_name: str) -> pd.DataFrame:
    ws = get_ws(sheet_name)
    values = ws.get_all_values()
    header_row = HEADER_ROWS.get(sheet_name, 1)

    if len(values) < header_row:
        return pd.DataFrame()

    headers = [str(h).strip() for h in values[header_row - 1]]
    data_rows = values[header_row:]
    data_rows = [r for r in data_rows if any(str(c).strip() for c in r)]

    width = len(headers)
    normalized_rows = []
    for row in data_rows:
        if len(row) < width:
            row = row + [""] * (width - len(row))
        elif len(row) > width:
            row = row[:width]
        normalized_rows.append(row)

    df = pd.DataFrame(normalized_rows, columns=headers)
    df = df.loc[:, [c for c in df.columns if c != ""]]
    return df


@st.cache_data(ttl=10, show_spinner=False)
def load_core_data() -> Dict[str, pd.DataFrame]:
    return {
        "vezbaci": read_sheet(SHEET_VEZBACI),
        "vezbe": read_sheet(SHEET_BAZA_VEZBI),
        "testovi_unos": read_sheet(SHEET_TEST_INPUT),
        "dnevnik": read_sheet(SHEET_DNEVNIK),
        "plan": read_sheet(SHEET_PLAN),
    }


def clear_caches() -> None:
    st.cache_data.clear()


def get_headers(sheet_name: str) -> List[str]:
    ws = get_ws(sheet_name)
    header_row = HEADER_ROWS.get(sheet_name, 1)
    return [str(h).strip() for h in ws.row_values(header_row)]


def append_rows_to_sheet(sheet_name: str, rows: List[Dict[str, object]]) -> None:
    if not rows:
        return
    ws = get_ws(sheet_name)
    headers = get_headers(sheet_name)
    if not headers:
        st.error(f"List {sheet_name} nema definisana zaglavlja.")
        st.stop()
    prepared_rows = [[to_sheet_value(row.get(h, "")) for h in headers] for row in rows]
    ws.append_rows(prepared_rows, value_input_option="USER_ENTERED")
    clear_caches()

# ------------------------------------------------------------
# 4) PODESAVANJA TRENINGA
# ------------------------------------------------------------

@st.cache_data(ttl=30, show_spinner=False)
def load_training_settings() -> Dict[str, Any]:
    settings = {
        "round_step": ROUND_KG_STEP_DEFAULT,
        "goals": DEFAULT_GOAL_SETTINGS.copy(),
        "weeks_531": DEFAULT_531_WEEKS.copy(),
    }
    try:
        ws = get_ws(SHEET_PODESAVANJA)
        values = ws.get_all_values()
        # B2 je korak za zaokruzivanje kg u nasem Excelu
        if len(values) >= 2 and len(values[1]) >= 2:
            settings["round_step"] = to_float(values[1][1], ROUND_KG_STEP_DEFAULT) or ROUND_KG_STEP_DEFAULT

        # Ciljevi: red 4 header, redovi 5-10 podaci
        if len(values) >= 5:
            headers = [normalize_text(x) for x in values[3]]
            goal_rows = values[4:10]
            parsed_goals = {}
            for r in goal_rows:
                if not r or not normalize_text(r[0]):
                    continue
                row = {headers[i]: r[i] if i < len(r) else "" for i in range(len(headers))}
                goal_name = normalize_text(row.get("Cilj"))
                if goal_name:
                    parsed_goals[goal_name] = {
                        "TM %": to_float(row.get("TM %"), DEFAULT_GOAL_SETTINGS.get(goal_name, {}).get("TM %", 0.9)),
                        "RIR Cilj": to_int(row.get("RIR Cilj"), DEFAULT_GOAL_SETTINGS.get(goal_name, {}).get("RIR Cilj", 3)),
                        "Serije+": to_int(row.get("Serije+"), DEFAULT_GOAL_SETTINGS.get(goal_name, {}).get("Serije+", 0)),
                        "Kg Faktor": to_float(row.get("Kg Faktor"), DEFAULT_GOAL_SETTINGS.get(goal_name, {}).get("Kg Faktor", 1.0)),
                        "RIR Asist.": to_int(row.get("RIR Asist."), DEFAULT_GOAL_SETTINGS.get(goal_name, {}).get("RIR Asist.", 3)),
                        "Zona Asist.": normalize_text(row.get("Zona Asist.")) or DEFAULT_GOAL_SETTINGS.get(goal_name, {}).get("Zona Asist.", "10-15"),
                    }
            if parsed_goals:
                settings["goals"] = {**DEFAULT_GOAL_SETTINGS, **parsed_goals}
    except Exception:
        pass
    return settings


def get_goal_settings(goal: str) -> Dict[str, Any]:
    settings = load_training_settings()
    goals = settings["goals"]
    goal = normalize_text(goal)
    return goals.get(goal) or goals.get("Snaga+Hipertrofija") or DEFAULT_GOAL_SETTINGS["Snaga+Hipertrofija"]

# ------------------------------------------------------------
# 5) LOGOVANJE I PAYWALL
# ------------------------------------------------------------

def get_user_by_login(email: str, sifra: str, vezbaci: pd.DataFrame) -> Optional[pd.Series]:
    for col in REQUIRED_VEZBACI_COLS:
        if col not in vezbaci.columns:
            st.error(f"U listu VEZBACI nedostaje kolona: {col}")
            st.stop()

    email_norm = normalize_email(email)
    sifra_norm = normalize_sifra(sifra)
    df = vezbaci.copy()
    df["__email"] = df["Email"].apply(normalize_email)
    df["__sifra"] = df["Sifra"].apply(normalize_sifra)
    match = df[(df["__email"] == email_norm) & (df["__sifra"] == sifra_norm)]
    if match.empty:
        return None
    return match.iloc[0]


def show_login_screen() -> None:
    st.title(APP_TITLE)
    st.subheader("Prijava vežbača")
    st.info("Unesite email i šifru iz registra vežbača.")
    data = load_core_data()
    vezbaci = data["vezbaci"]
    email = st.text_input("Email", placeholder="npr. ana@email.com", key="login_email")
    sifra = st.text_input("Šifra", type="password", placeholder="4 broja", key="login_sifra")
    if st.button("Uđi u aplikaciju", type="primary", use_container_width=True):
        user = get_user_by_login(email, sifra, vezbaci)
        if user is None:
            st.error("Pogrešan email ili šifra.")
            return
        st.session_state["logged_in"] = True
        st.session_state["user"] = user.to_dict()
        st.rerun()


def enforce_paywall(user: Dict[str, object]) -> None:
    aktivan = normalize_text(user.get("Aktivan", ""))
    if aktivan.lower() in ["ne", "no", "0", "false", "nije aktivan"]:
        poruka = normalize_text(user.get("Poruka_Blokada", "")) or "Članarina je istekla. Kontaktirajte trenera za produženje."
        st.markdown(
            f"""
            <div style="height:70vh; display:flex; align-items:center; justify-content:center; text-align:center;">
                <div style="max-width:760px; padding:40px; border:2px solid #cc0000; border-radius:20px;">
                    <h1>Pristup je zaključan</h1>
                    <p style="font-size:28px; font-weight:600;">{poruka}</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.stop()

# ------------------------------------------------------------
# 6) VEZBE, TESTOVI I PLAN
# ------------------------------------------------------------

def exercise_column_name(vezbe: pd.DataFrame) -> str:
    for possible in ["Vežba", "Naziv vežbe", "Vezba", "Naziv"]:
        if possible in vezbe.columns:
            return possible
    st.error("U listu BAZA_VEZBI ne nalazim kolonu sa nazivom vežbe. Očekujem 'Vežba' ili 'Naziv vežbe'.")
    st.stop()


def get_exercise_meta(vezbe: pd.DataFrame, exercise: str) -> Dict[str, Any]:
    if vezbe.empty:
        return {}
    col = exercise_column_name(vezbe)
    df = vezbe.copy()
    df["__ex"] = df[col].apply(lambda x: normalize_text(x).lower())
    m = df[df["__ex"] == normalize_text(exercise).lower()]
    if m.empty:
        return {}
    return m.iloc[0].to_dict()


def user_has_tests(email: str, testovi_unos: pd.DataFrame) -> bool:
    if testovi_unos.empty or "Email" not in testovi_unos.columns:
        return False
    df = testovi_unos.copy()
    df["__email"] = df["Email"].apply(normalize_email)
    return (df["__email"] == normalize_email(email)).any()


def latest_test_for(email: str, exercise: str, testovi_unos: pd.DataFrame) -> Optional[Dict[str, Any]]:
    required = ["Email", "Naziv vežbe", "Procena 1RM"]
    if testovi_unos.empty or any(c not in testovi_unos.columns for c in required):
        return None
    df = testovi_unos.copy()
    df["__email"] = df["Email"].apply(normalize_email)
    df["__ex"] = df["Naziv vežbe"].apply(lambda x: normalize_text(x).lower())
    df = df[(df["__email"] == normalize_email(email)) & (df["__ex"] == normalize_text(exercise).lower())]
    if df.empty:
        return None
    if "Timestamp" in df.columns:
        sort_col = "Timestamp"
    elif "Datum" in df.columns:
        sort_col = "Datum"
    else:
        sort_col = None
    if sort_col:
        df["__date"] = pd.to_datetime(df[sort_col], errors="coerce")
        df = df.sort_values("__date")
    row = df.iloc[-1].to_dict()
    one_rm = to_float(row.get("Procena 1RM"), 0)
    if one_rm <= 0:
        kg = to_float(row.get("Kilaža (kg)"), 0)
        reps = to_int(row.get("Broj ponavljanja"), 1)
        one_rm = estimate_1rm(kg, reps)
    row["__1rm"] = one_rm
    return row if one_rm > 0 else None


def parse_assistants(text: str) -> List[str]:
    text = normalize_text(text)
    if not text:
        return []
    parts = re.split(r"\s*/\s*|\s*;\s*|\s*,\s*", text)
    return [p.strip() for p in parts if p.strip()]


def all_plan_exercises(plan: pd.DataFrame) -> List[str]:
    out: List[str] = []
    if plan.empty:
        return out
    for _, r in plan.iterrows():
        main = normalize_text(r.get("Glavna vežba"))
        if main:
            out.append(main)
        for a in parse_assistants(r.get("Asistencije — skraćeno", "")):
            out.append(a)
    # zadrzi redosled, ukloni duplikate
    seen = set()
    unique = []
    for x in out:
        key = x.lower()
        if key not in seen:
            unique.append(x)
            seen.add(key)
    return unique


def missing_tests_for_user(email: str, exercises: List[str], testovi: pd.DataFrame) -> List[str]:
    missing = []
    for ex in exercises:
        if latest_test_for(email, ex, testovi) is None:
            missing.append(ex)
    return missing


def show_initial_tests_screen(user: Dict[str, object]) -> None:
    st.header("Testiranje vežbi")
    st.write("Za svaku vežbu unosi se kilaža i broj ponavljanja. Aplikacija računa procenjeni 1RM i od toga kasnije pravi plan.")

    data = load_core_data()
    vezbe = data["vezbe"]
    plan = data["plan"]
    testovi = data["testovi_unos"]
    email = normalize_email(user.get("Email"))

    exercise_col = exercise_column_name(vezbe)
    exercise_options = sorted([v for v in vezbe[exercise_col].dropna().astype(str).unique().tolist() if v.strip()])

    plan_exercises = all_plan_exercises(plan)
    missing = missing_tests_for_user(email, plan_exercises, testovi)

    if missing:
        st.warning(f"Nedostaje test za {len(missing)} vežbi iz plana. Prvo unesi ove vežbe, pa će aplikacija moći da računa predlog.")
        default_selection = missing
    else:
        st.success("Za sve vežbe iz plana postoji bar jedan test. Možeš dodati novi test ako želiš da ažuriraš plan.")
        default_selection = plan_exercises[:6] if plan_exercises else exercise_options[:6]

    selected = st.multiselect(
        "Izaberi vežbe za testiranje / ažuriranje testa",
        options=exercise_options,
        default=[x for x in default_selection if x in exercise_options],
    )

    if not selected:
        st.warning("Izaberi bar jednu vežbu.")
        return

    rows = []
    with st.form("testovi_form"):
        test_date = st.date_input("Datum testa", value=date.today())
        st.caption("Kod asistentskih vežbi test ne mora biti pravi maksimum. Unesi sigurnu radnu kilažu i reps do procene, uz dobru tehniku.")
        for ex in selected:
            meta = get_exercise_meta(vezbe, ex)
            max_reps = to_int(meta.get("Max reps test", 15), 15) or 15
            rep_zone = normalize_text(meta.get("Rep zona", ""))
            st.markdown(f"**{ex}**  ·  zona {rep_zone or '-'}  ·  max test reps {max_reps}")
            c1, c2, c3 = st.columns(3)
            kg = c1.number_input(f"Kilaža kg — {ex}", min_value=0.0, step=2.5, key=f"kg_{ex}")
            reps = c2.number_input(f"Broj ponavljanja — {ex}", min_value=1, max_value=max(30, max_reps), value=min(max_reps, 10), step=1, key=f"reps_{ex}")
            one_rm = estimate_1rm(kg, reps) if kg > 0 else 0
            c3.metric("Procena 1RM", f"{format_kg(one_rm)} kg" if one_rm else "—")
            rows.append((ex, kg, reps, one_rm))
        submitted = st.form_submit_button("Sačuvaj testove")

    if submitted:
        valid_rows = []
        for ex, kg, reps, one_rm in rows:
            if kg <= 0:
                continue
            valid_rows.append({
                "Datum": test_date,
                "Email": email,
                "Naziv vežbe": ex,
                "Kilaža (kg)": kg,
                "Broj ponavljanja": int(reps),
                "Procena 1RM": one_rm,
                "ID (auto)": user.get("ID", ""),
                "Ime (auto)": user.get("Ime i Prezime", ""),
                "Napomena": "Unos preko Streamlit aplikacije",
            })
        if not valid_rows:
            st.warning("Nema validnih unosa. Unesite kilažu za bar jednu vežbu.")
            return
        append_rows_to_sheet(SHEET_TEST_INPUT, valid_rows)
        st.success("Testovi su sačuvani u Google Sheets list UNOS_TESTOVA.")
        st.rerun()

# ------------------------------------------------------------
# 7) RACUNANJE PLANA
# ------------------------------------------------------------

def get_current_week_number() -> int:
    return ((date.today().isocalendar().week - 1) % 4) + 1


def get_last_log(email: str, exercise: str, dnevnik: pd.DataFrame) -> Optional[Dict[str, Any]]:
    if dnevnik.empty or "Email" not in dnevnik.columns or "Vežba" not in dnevnik.columns:
        return None
    df = dnevnik.copy()
    df["__email"] = df["Email"].apply(normalize_email)
    df["__ex"] = df["Vežba"].apply(lambda x: normalize_text(x).lower())
    df = df[(df["__email"] == normalize_email(email)) & (df["__ex"] == normalize_text(exercise).lower())]
    if df.empty:
        return None
    sort_col = "Timestamp" if "Timestamp" in df.columns else "Datum" if "Datum" in df.columns else None
    if sort_col:
        df["__date"] = pd.to_datetime(df[sort_col], errors="coerce")
        df = df.sort_values("__date")
    return df.iloc[-1].to_dict()


def apply_progression_correction(base_kg: float, exercise_meta: Dict[str, Any], target_sets: int, target_reps: int, target_rir: int, last_log: Optional[Dict[str, Any]], round_step: float) -> Tuple[float, str]:
    if not base_kg or base_kg <= 0:
        return base_kg, "nema testa"
    if not last_log:
        return base_kg, "prvi plan posle testa"

    step = to_float(exercise_meta.get("Napredak (kg)"), round_step) or round_step
    done_kg = to_float(last_log.get("Urađeno kg"), 0) or to_float(last_log.get("Plan kg"), 0)
    done_sets = to_int(last_log.get("Urađeno ser."), 0) or to_int(last_log.get("Plan ser."), 0)
    done_reps = to_int(last_log.get("Urađeno reps"), 0) or to_int(last_log.get("Plan reps"), 0)
    rir = to_int(last_log.get("RIR stvarni"), target_rir)

    completed = True
    if target_sets and done_sets and done_sets < target_sets:
        completed = False
    if target_reps and done_reps and done_reps < target_reps:
        completed = False
    if done_kg and done_kg < base_kg - 0.01:
        completed = False

    if completed and rir >= target_rir + 2:
        return round_to_step(base_kg + step, round_step), f"bilo lako (RIR {rir}) → +{format_kg(step)} kg"
    if (not completed) or rir <= max(0, target_rir - 2):
        return max(0, round_to_step(base_kg - step, round_step)), f"bilo teško/neodrađeno (RIR {rir}) → -{format_kg(step)} kg ili zadržati"
    return base_kg, f"dobro pogođeno (RIR {rir}) → zadrži"


def calc_main_plan(email: str, exercise: str, week: int, user: Dict[str, Any], vezbe: pd.DataFrame, testovi: pd.DataFrame, dnevnik: pd.DataFrame) -> Dict[str, Any]:
    settings = load_training_settings()
    round_step = settings["round_step"]
    goal = normalize_text(user.get("Cilj"))
    goal_set = get_goal_settings(goal)
    test = latest_test_for(email, exercise, testovi)
    meta = get_exercise_meta(vezbe, exercise)
    target_rir = to_int(goal_set.get("RIR Cilj"), 3)

    if not test:
        return {"Vežba": exercise, "Tip": "Glavna", "Plan tekst": "Čeka test", "Plan kg": "", "Plan ser.": "", "Plan reps": "", "Plan RIR": target_rir, "Korekcija": "unesi test"}

    one_rm = to_float(test.get("__1rm"), 0)
    tm = round_to_step(one_rm * to_float(goal_set.get("TM %"), 0.9), round_step)
    scheme = DEFAULT_531_WEEKS.get(int(week), DEFAULT_531_WEEKS[1])
    set_texts = []
    planned_volume = 0.0
    top_kg = 0.0
    top_reps_numeric = 0
    for pct, reps in scheme:
        kg = round_to_step(tm * pct, round_step)
        set_texts.append(f"{format_kg(kg)}×{reps}")
        rep_num = to_int(reps, 1)
        planned_volume += kg * rep_num
        top_kg = kg
        top_reps_numeric = rep_num

    last_log = get_last_log(email, exercise, dnevnik)
    corrected_top_kg, correction_note = apply_progression_correction(top_kg, meta, 3, top_reps_numeric, target_rir, last_log, round_step)
    if corrected_top_kg != top_kg:
        set_texts[-1] = f"{format_kg(corrected_top_kg)}×{scheme[-1][1]}"

    plan_text = f"1RM {format_kg(one_rm)} kg | TM {format_kg(tm)} kg | " + " / ".join(set_texts)
    return {
        "Vežba": exercise,
        "Tip": "Glavna",
        "Plan tekst": plan_text,
        "Plan kg": corrected_top_kg,
        "Plan ser.": 3,
        "Plan reps": "/".join([str(r) for _, r in scheme]),
        "Plan RIR": target_rir,
        "Korekcija": correction_note,
        "Plan volumen": round(planned_volume, 1),
    }


def calc_assist_plan(email: str, exercise: str, week: int, user: Dict[str, Any], vezbe: pd.DataFrame, testovi: pd.DataFrame, dnevnik: pd.DataFrame) -> Dict[str, Any]:
    settings = load_training_settings()
    round_step = settings["round_step"]
    goal = normalize_text(user.get("Cilj"))
    goal_set = get_goal_settings(goal)
    test = latest_test_for(email, exercise, testovi)
    meta = get_exercise_meta(vezbe, exercise)

    rep_zone = normalize_text(meta.get("Rep zona", "")) or normalize_text(goal_set.get("Zona Asist.")) or "10-15"
    low_rep, high_rep = parse_range(rep_zone, fallback=parse_range(goal_set.get("Zona Asist."), (10, 15)))
    max_reps_test = to_int(meta.get("Max reps test"), high_rep) or high_rep
    target_rir = to_int(goal_set.get("RIR Asist."), 3)
    sets_text = ASSIST_WEEK_SETS.get(int(week), "2")
    base_sets = 2 if sets_text == "1-2" else to_int(sets_text, 2)
    base_sets = max(1, base_sets + to_int(goal_set.get("Serije+"), 0))
    week_factor = ASSIST_WEEK_FACTOR.get(int(week), 1.0)
    kg_factor = to_float(goal_set.get("Kg Faktor"), 1.0)

    if not test:
        return {"Vežba": exercise, "Tip": "Asistencija", "Plan tekst": f"Čeka test | cilj zona {rep_zone}, RIR {target_rir}", "Plan kg": "", "Plan ser.": base_sets, "Plan reps": rep_zone, "Plan RIR": target_rir, "Korekcija": "unesi test"}

    one_rm = to_float(test.get("__1rm"), 0)
    # Epley unazad: za ciljni broj ponavljanja + RIR od procenjenog 1RM dobijamo radnu kilazu.
    target_reps_for_calc = high_rep
    base_kg = one_rm / (1 + ((target_reps_for_calc + target_rir) / 30))
    base_kg = round_to_step(base_kg * week_factor * kg_factor, round_step)

    last_log = get_last_log(email, exercise, dnevnik)
    corrected_kg, correction_note = apply_progression_correction(base_kg, meta, base_sets, high_rep, target_rir, last_log, round_step)

    plan_text = f"{base_sets}×{rep_zone} @ {format_kg(corrected_kg)} kg | cilj RIR {target_rir} | 1RM {format_kg(one_rm)} kg"
    return {
        "Vežba": exercise,
        "Tip": "Asistencija",
        "Plan tekst": plan_text,
        "Plan kg": corrected_kg,
        "Plan ser.": base_sets,
        "Plan reps": rep_zone,
        "Plan RIR": target_rir,
        "Korekcija": correction_note,
        "Plan volumen": round(corrected_kg * base_sets * high_rep, 1),
    }


def build_training_plan_for_user(user: Dict[str, object], week: int, trening: int) -> pd.DataFrame:
    data = load_core_data()
    plan = data["plan"].copy()
    vezbe = data["vezbe"]
    testovi = data["testovi_unos"]
    dnevnik = data["dnevnik"]
    email = normalize_email(user.get("Email"))

    required = ["Ned.", "Trening", "Glavna vežba", "Asistencije — skraćeno"]
    missing = [c for c in required if c not in plan.columns]
    if missing:
        st.error(f"U listu {SHEET_PLAN} nedostaju kolone: {missing}")
        st.stop()

    plan["Ned."] = pd.to_numeric(plan["Ned."], errors="coerce")
    plan["Trening"] = pd.to_numeric(plan["Trening"], errors="coerce")
    plan = plan.dropna(subset=["Ned.", "Trening"])
    plan["Ned."] = plan["Ned."].astype(int)
    plan["Trening"] = plan["Trening"].astype(int)
    plan = plan[(plan["Ned."] == int(week)) & (plan["Trening"] == int(trening))]

    rows = []
    for _, r in plan.iterrows():
        main = normalize_text(r.get("Glavna vežba"))
        if main:
            rows.append(calc_main_plan(email, main, week, user, vezbe, testovi, dnevnik))
        for a in parse_assistants(r.get("Asistencije — skraćeno", "")):
            rows.append(calc_assist_plan(email, a, week, user, vezbe, testovi, dnevnik))
    return pd.DataFrame(rows)

# ------------------------------------------------------------
# 8) DANAŠNJI TRENING
# ------------------------------------------------------------

def show_today_training_screen(user: Dict[str, object]) -> None:
    st.header("Današnji trening")
    st.write("Aplikacija računa predlog iz testova. Ti upisuješ samo odstupanja: ako je nešto urađeno drugačije, upiši stvarno stanje.")

    goal = normalize_text(user.get("Cilj")) or "Snaga+Hipertrofija"
    st.info(f"Cilj programa: **{goal}**. Glavne vežbe idu po 5/3/1, asistencije po testu, rep zoni i ciljanom RIR-u.")

    c1, c2 = st.columns(2)
    week = c1.selectbox("Nedelja ciklusa", options=[1, 2, 3, 4], index=get_current_week_number() - 1)
    trening = c2.selectbox("Trening", options=[1, 2, 3, 4], index=0)

    plan_df = build_training_plan_for_user(user, week, trening)
    if plan_df.empty:
        st.warning("Nema pronađenog plana za izabranu nedelju i trening.")
        return

    st.subheader("Predlog plana")
    display_cols = ["Vežba", "Tip", "Plan tekst", "Plan RIR", "Korekcija"]
    st.dataframe(plan_df[display_cols], use_container_width=True, hide_index=True)

    if (plan_df["Korekcija"].astype(str).str.contains("unesi test", case=False, na=False)).any():
        st.warning("Neke vežbe nemaju test. Pređi na ekran Testovi i unesi test da bi aplikacija računala kilažu.")

    with st.form("dnevnik_form"):
        st.subheader("Unos odstupanja")
        st.caption("Ako je urađeno tačno po planu, ne moraš ništa da upisuješ. Ako želiš da zabeležiš trening i bez odstupanja, upiši bar RIR ili napomenu.")
        entries = []
        for idx, row in plan_df.iterrows():
            ex = row["Vežba"]
            st.markdown(f"**{ex}** — {row['Tip']}  ·  plan: {row['Plan tekst']}")
            c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 2])
            plan_kg_value = to_float(row.get("Plan kg"), 0)
            default_rir = to_int(row.get("Plan RIR"), 3)
            done_kg = c1.number_input("Urađeno kg", min_value=0.0, step=2.5, key=f"done_kg_{idx}")
            done_sets = c2.number_input("Serije", min_value=0, step=1, key=f"sets_{idx}")
            done_reps = c3.number_input("Reps", min_value=0, step=1, key=f"reps_{idx}")
            rir = c4.number_input("RIR", min_value=0, max_value=10, value=default_rir, step=1, key=f"rir_{idx}")
            note = c5.text_input("Napomena", key=f"note_{idx}")
            entries.append({
                "exercise": ex,
                "plan_kg": plan_kg_value if plan_kg_value > 0 else "",
                "plan_sets": row.get("Plan ser.", ""),
                "plan_reps": row.get("Plan reps", ""),
                "done_kg": done_kg if done_kg > 0 else None,
                "done_sets": int(done_sets) if done_sets > 0 else None,
                "done_reps": int(done_reps) if done_reps > 0 else None,
                "rir": int(rir),
                "note": note,
            })
        submitted = st.form_submit_button("Sačuvaj trening")

    if submitted:
        now = datetime.now()
        rows_to_write = []
        for e in entries:
            # Upisujemo red ako postoji odstupanje, RIR ili napomena.
            has_change = any([e["done_kg"], e["done_sets"], e["done_reps"], e["note"]])
            # Pošto RIR ima default, nećemo samo zbog njega upisivati sve vežbe automatski.
            if not has_change:
                continue

            kg_for_volume = e["done_kg"] or to_float(e["plan_kg"], 0)
            # Za Plan ser/reps mogu biti tekstovi; za volumen uzimamo prvi broj.
            sets_for_volume = e["done_sets"] or to_int(e["plan_sets"], 0)
            reps_for_volume = e["done_reps"] or to_int(e["plan_reps"], 0)
            volume = kg_for_volume * sets_for_volume * reps_for_volume if kg_for_volume and sets_for_volume and reps_for_volume else ""

            rows_to_write.append({
                "Email": normalize_email(user.get("Email")),
                "Timestamp": now,
                "Datum": now.date(),
                "ID": user.get("ID", ""),
                "Ime": user.get("Ime i Prezime", ""),
                "Nedelja": week,
                "Trening": trening,
                "Vežba": e["exercise"],
                "Plan kg": e["plan_kg"] or "",
                "Plan ser.": e["plan_sets"] or "",
                "Plan reps": e["plan_reps"] or "",
                "Urađeno kg": e["done_kg"] or "",
                "Urađeno ser.": e["done_sets"] or "",
                "Urađeno reps": e["done_reps"] or "",
                "RIR stvarni": e["rir"],
                "Status": "Odstupanje",
                "Napomena": e["note"],
                "Volumen (kg)": volume,
            })
        if not rows_to_write:
            st.info("Nema odstupanja za upis. To znači da je plan urađen kako je zadat.")
            return
        append_rows_to_sheet(SHEET_DNEVNIK, rows_to_write)
        st.success("Trening je sačuvan u Google Sheets list DNEVNIK_UNOS.")
        st.rerun()




# ------------------------------------------------------------
# 9A) ADMIN IZVESTAJ - A4 STAMPA
# ------------------------------------------------------------

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Vraca ime kolone uz toleranciju na mala/velika slova i razmake."""
    if df is None or df.empty:
        return None
    norm_map = {normalize_text(c).lower(): c for c in df.columns}
    for cand in candidates:
        key = normalize_text(cand).lower()
        if key in norm_map:
            return norm_map[key]
    # labavije: bez specijalnih znakova
    def simple(x: str) -> str:
        return re.sub(r"[^a-zA-Z0-9čćžšđČĆŽŠĐ]+", "", normalize_text(x).lower())
    simple_map = {simple(c): c for c in df.columns}
    for cand in candidates:
        key = simple(cand)
        if key in simple_map:
            return simple_map[key]
    return None


def add_datetime_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    date_col = find_col(df, ["Timestamp", "Vreme unosa", "Datum", "Date"])
    if date_col:
        df["__dt"] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
    else:
        df["__dt"] = pd.NaT
    return df


def add_volume_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    vol_col = find_col(df, ["Volumen (kg)", "Volumen", "Volume"])
    if vol_col:
        df["__volume"] = pd.to_numeric(df[vol_col], errors="coerce").fillna(0.0)
    else:
        kg_col = find_col(df, ["Urađeno kg", "Uradjeno kg", "Done kg", "Plan kg"])
        set_col = find_col(df, ["Urađeno ser.", "Uradjeno ser.", "Serije", "Plan ser."])
        rep_col = find_col(df, ["Urađeno reps", "Uradjeno reps", "Reps", "Plan reps"])
        kg = pd.to_numeric(df[kg_col], errors="coerce").fillna(0.0) if kg_col else 0.0
        sets = pd.to_numeric(df[set_col], errors="coerce").fillna(0.0) if set_col else 0.0
        reps = pd.to_numeric(df[rep_col], errors="coerce").fillna(0.0) if rep_col else 0.0
        df["__volume"] = kg * sets * reps
    return df


def one_rm_series_from_tests(df: pd.DataFrame) -> pd.Series:
    one_col = find_col(df, ["Procena 1RM", "Ostvareni 1RM", "1RM", "e1RM"])
    if one_col:
        one = pd.to_numeric(df[one_col], errors="coerce")
    else:
        one = pd.Series([None] * len(df), index=df.index, dtype="float")
    kg_col = find_col(df, ["Kilaža (kg)", "Kilaza (kg)", "Kilaža", "Kilaza", "kg"])
    reps_col = find_col(df, ["Broj ponavljanja", "Ponavljanja", "Reps", "reps"])
    if kg_col and reps_col:
        kg = pd.to_numeric(df[kg_col], errors="coerce")
        reps = pd.to_numeric(df[reps_col], errors="coerce")
        calc = kg * (1 + reps / 30)
        one = one.fillna(calc)
    return one.fillna(0.0)


def prepare_user_report(email: str, start_date: date, end_date: date, vezbaci: pd.DataFrame, testovi: pd.DataFrame, dnevnik: pd.DataFrame) -> Dict[str, Any]:
    email = normalize_email(email)
    user_row = {}
    if not vezbaci.empty and "Email" in vezbaci.columns:
        v = vezbaci[vezbaci["Email"].apply(normalize_email) == email]
        if not v.empty:
            user_row = v.iloc[0].to_dict()

    # Dnevnik za period
    logs = dnevnik.copy()
    if not logs.empty and "Email" in logs.columns:
        logs["__email"] = logs["Email"].apply(normalize_email)
        logs = logs[logs["__email"] == email]
    logs = add_datetime_column(logs)
    logs = add_volume_column(logs)
    if "__dt" in logs.columns:
        mask = (logs["__dt"].dt.date >= start_date) & (logs["__dt"].dt.date <= end_date)
        logs_period = logs[mask].copy()
    else:
        logs_period = logs.copy()

    ex_col = find_col(logs_period, ["Vežba", "Vezba", "Naziv vežbe", "Exercise"])
    rir_col = find_col(logs_period, ["RIR stvarni", "RIR", "rir"])
    kg_col = find_col(logs_period, ["Urađeno kg", "Uradjeno kg", "Plan kg"])
    sets_col = find_col(logs_period, ["Urađeno ser.", "Uradjeno ser.", "Plan ser."])
    reps_col = find_col(logs_period, ["Urađeno reps", "Uradjeno reps", "Plan reps"])
    trening_col = find_col(logs_period, ["Trening", "Training"])

    total_volume = float(logs_period["__volume"].sum()) if "__volume" in logs_period.columns else 0.0
    exercise_rows = len(logs_period)
    avg_rir = None
    if rir_col:
        rir_vals = pd.to_numeric(logs_period[rir_col], errors="coerce").dropna()
        if not rir_vals.empty:
            avg_rir = float(rir_vals.mean())

    if not logs_period.empty and "__dt" in logs_period.columns and logs_period["__dt"].notna().any():
        if trening_col:
            keys = logs_period["__dt"].dt.strftime("%Y-%m-%d") + "_T" + logs_period[trening_col].astype(str)
            training_count = int(keys.nunique())
        else:
            training_count = int(logs_period["__dt"].dt.date.nunique())
    else:
        training_count = 0
    avg_vol_per_training = total_volume / training_count if training_count else 0.0

    session_volume = pd.DataFrame()
    if not logs_period.empty and "__dt" in logs_period.columns and logs_period["__dt"].notna().any():
        logs_period["__day"] = logs_period["__dt"].dt.date.astype(str)
        if trening_col:
            logs_period["__session"] = logs_period["__day"] + " / T" + logs_period[trening_col].astype(str)
        else:
            logs_period["__session"] = logs_period["__day"]
        session_volume = logs_period.groupby("__session", as_index=False)["__volume"].sum().sort_values("__session")
    max_session_volume = float(session_volume["__volume"].max()) if not session_volume.empty else 0.0

    best_lifts = pd.DataFrame()
    if ex_col and kg_col and not logs_period.empty:
        tmp = logs_period.copy()
        tmp["__kg"] = pd.to_numeric(tmp[kg_col], errors="coerce").fillna(0.0)
        tmp["__sets"] = pd.to_numeric(tmp[sets_col], errors="coerce").fillna(0.0) if sets_col else 0.0
        tmp["__reps"] = pd.to_numeric(tmp[reps_col], errors="coerce").fillna(0.0) if reps_col else 0.0
        best_lifts = tmp.groupby(ex_col).agg(
            **{"Najveća kg": ("__kg", "max"), "Ukupan volumen": ("__volume", "sum"), "Broj unosa": (ex_col, "count")}
        ).reset_index().rename(columns={ex_col: "Vežba"}).sort_values("Ukupan volumen", ascending=False).head(10)

    # Testovi i napredak
    tests = testovi.copy()
    if not tests.empty and "Email" in tests.columns:
        tests["__email"] = tests["Email"].apply(normalize_email)
        tests = tests[tests["__email"] == email]
    tests = add_datetime_column(tests)
    tests["__1rm"] = one_rm_series_from_tests(tests) if not tests.empty else []
    test_ex_col = find_col(tests, ["Naziv vežbe", "Naziv vezbe", "Vežba", "Vezba", "Exercise"])
    progress = pd.DataFrame()
    if test_ex_col and not tests.empty:
        if "__dt" in tests.columns:
            tests = tests.sort_values("__dt")
        rows = []
        for ex, g in tests.groupby(test_ex_col):
            g = g[g["__1rm"] > 0]
            if g.empty:
                continue
            first = g.iloc[0]
            last = g.iloc[-1]
            first_rm = float(first["__1rm"])
            last_rm = float(last["__1rm"])
            diff = last_rm - first_rm
            pct = (diff / first_rm * 100) if first_rm else 0.0
            rows.append({
                "Vežba": normalize_text(ex),
                "Početni 1RM": round(first_rm, 1),
                "Poslednji 1RM": round(last_rm, 1),
                "Napredak kg": round(diff, 1),
                "Napredak %": round(pct, 1),
                "Broj testova": len(g),
            })
        if rows:
            progress = pd.DataFrame(rows).sort_values("Napredak kg", ascending=False)

    weekly = pd.DataFrame()
    if not logs_period.empty and "__dt" in logs_period.columns and logs_period["__dt"].notna().any():
        tmp = logs_period.copy()
        tmp["Nedelja"] = tmp["__dt"].dt.to_period("W").astype(str)
        weekly = tmp.groupby("Nedelja", as_index=False)["__volume"].sum().rename(columns={"__volume": "Volumen"})

    return {
        "user": user_row,
        "logs_period": logs_period,
        "tests": tests,
        "progress": progress,
        "best_lifts": best_lifts,
        "weekly": weekly,
        "training_count": training_count,
        "exercise_rows": exercise_rows,
        "total_volume": total_volume,
        "avg_vol_per_training": avg_vol_per_training,
        "max_session_volume": max_session_volume,
        "avg_rir": avg_rir,
        "period": (start_date, end_date),
    }


def fmt_num(x: Any, decimals: int = 0) -> str:
    try:
        val = float(x)
    except Exception:
        return "-"
    if decimals == 0:
        return f"{val:,.0f}".replace(",", ".")
    return f"{val:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def table_to_html(df: pd.DataFrame, max_rows: int = 8) -> str:
    if df is None or df.empty:
        return "<p class='muted'>Nema podataka za prikaz.</p>"
    show = df.head(max_rows).copy()
    return show.to_html(index=False, classes="report-table", border=0, escape=False)


def build_a4_report_html(report: Dict[str, Any], trainer_note: str = "") -> str:
    user = report.get("user", {}) or {}
    name = normalize_text(user.get("Ime i Prezime")) or normalize_text(user.get("Ime")) or "Vežbač"
    email = normalize_email(user.get("Email")) or ""
    start_date, end_date = report["period"]
    avg_rir = report.get("avg_rir")
    avg_rir_text = fmt_num(avg_rir, 1) if avg_rir is not None else "-"
    progress_html = table_to_html(report.get("progress", pd.DataFrame()), max_rows=8)
    best_html = table_to_html(report.get("best_lifts", pd.DataFrame()), max_rows=8)
    weekly = report.get("weekly", pd.DataFrame())
    weekly_html = table_to_html(weekly, max_rows=6)
    note_html = normalize_text(trainer_note).replace("\n", "<br>") or "&nbsp;<br>&nbsp;<br>&nbsp;"
    generated = datetime.now().strftime("%d.%m.%Y. %H:%M")

    return f"""
<style>
.report-a4 {{
  background: white;
  color: #111;
  width: 210mm;
  min-height: 297mm;
  padding: 12mm 12mm 10mm 12mm;
  margin: 0 auto;
  border: 1px solid #ddd;
  box-shadow: 0 0 8px rgba(0,0,0,.08);
  font-family: Arial, sans-serif;
  font-size: 11px;
  line-height: 1.25;
}}
.report-title {{font-size: 20px; font-weight: 700; margin: 0 0 4px 0;}}
.report-subtitle {{font-size: 11px; color:#555; margin-bottom: 10px;}}
.report-section-title {{font-size: 13px; font-weight: 700; margin: 10px 0 5px 0; border-bottom: 1px solid #222; padding-bottom: 2px;}}
.metric-grid {{display:grid; grid-template-columns: repeat(4, 1fr); gap: 6px; margin: 8px 0;}}
.metric-card {{border:1px solid #ccc; padding:6px; border-radius:6px;}}
.metric-label {{font-size:9px; color:#555; text-transform:uppercase;}}
.metric-value {{font-size:15px; font-weight:700; margin-top:2px;}}
.report-table {{width:100%; border-collapse: collapse; font-size: 10px;}}
.report-table th {{background:#f2f2f2; border:1px solid #bbb; padding:4px; text-align:left;}}
.report-table td {{border:1px solid #ccc; padding:4px;}}
.two-col {{display:grid; grid-template-columns: 1fr 1fr; gap: 8px;}}
.muted {{color:#777; font-style:italic;}}
.note-box {{border:1px solid #ccc; min-height:45px; padding:6px;}}
.footer {{margin-top:8px; font-size:9px; color:#777; display:flex; justify-content:space-between;}}
@media print {{
  body * {{ visibility: hidden; }}
  .report-a4, .report-a4 * {{ visibility: visible; }}
  .report-a4 {{ position: absolute; left: 0; top: 0; width: 190mm; min-height: 277mm; margin:0; border:none; box-shadow:none; padding:10mm; }}
  @page {{ size: A4; margin: 10mm; }}
}}
</style>
<div class="report-a4">
  <div class="report-title">MESEČNI IZVEŠTAJ VEŽBAČA</div>
  <div class="report-subtitle">Vežbač: <b>{name}</b> &nbsp; | &nbsp; Email: {email} &nbsp; | &nbsp; Period: {start_date.strftime('%d.%m.%Y.')} – {end_date.strftime('%d.%m.%Y.')}</div>

  <div class="metric-grid">
    <div class="metric-card"><div class="metric-label">Broj treninga</div><div class="metric-value">{report.get('training_count',0)}</div></div>
    <div class="metric-card"><div class="metric-label">Ukupan volumen</div><div class="metric-value">{fmt_num(report.get('total_volume',0))} kg</div></div>
    <div class="metric-card"><div class="metric-label">Prosek / trening</div><div class="metric-value">{fmt_num(report.get('avg_vol_per_training',0))} kg</div></div>
    <div class="metric-card"><div class="metric-label">Prosečan RIR</div><div class="metric-value">{avg_rir_text}</div></div>
  </div>

  <div class="two-col">
    <div>
      <div class="report-section-title">Napredak u testovima</div>
      {progress_html}
    </div>
    <div>
      <div class="report-section-title">Najbolji rezultati u periodu</div>
      {best_html}
    </div>
  </div>

  <div class="report-section-title">Volumen po nedeljama</div>
  {weekly_html}

  <div class="report-section-title">Komentar trenera</div>
  <div class="note-box">{note_html}</div>

  <div class="footer"><span>Generisano: {generated}</span><span>Fitnes klub — 5/3/1 + RIR</span></div>
</div>
"""


def show_admin_report_screen(vezbaci: pd.DataFrame, testovi: pd.DataFrame, dnevnik: pd.DataFrame) -> None:
    st.subheader("A4 izveštaj vežbača")
    st.caption("Izaberi vežbača i period. Izveštaj je formatiran da stane na jedan A4 list. Za štampu koristi Ctrl+P ili download HTML fajla.")

    if vezbaci.empty or "Email" not in vezbaci.columns:
        st.warning("Nema podataka u listu VEZBACI.")
        return
    emails = sorted([normalize_email(e) for e in vezbaci["Email"].dropna().tolist() if normalize_email(e)])
    if not emails:
        st.warning("Nema email adresa u listu VEZBACI.")
        return

    c1, c2, c3 = st.columns([2, 1, 1])
    selected_email = c1.selectbox("Vežbač", emails, key="report_email")
    today = date.today()
    default_start = today.replace(day=1)
    start_date = c2.date_input("Od", value=default_start, key="report_start")
    end_date = c3.date_input("Do", value=today, key="report_end")
    trainer_note = st.text_area("Komentar trenera za izveštaj", height=90, placeholder="Upiši kratak komentar koji će se pojaviti na A4 izveštaju.")

    if start_date > end_date:
        st.error("Datum 'Od' ne može biti posle datuma 'Do'.")
        return

    report = prepare_user_report(selected_email, start_date, end_date, vezbaci, testovi, dnevnik)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Treninzi", report["training_count"])
    m2.metric("Ukupan volumen", f"{fmt_num(report['total_volume'])} kg")
    m3.metric("Prosek / trening", f"{fmt_num(report['avg_vol_per_training'])} kg")
    m4.metric("Prosečan RIR", fmt_num(report["avg_rir"], 1) if report["avg_rir"] is not None else "-")

    if report["progress"].empty and report["best_lifts"].empty and report["training_count"] == 0:
        st.info("Za izabrani period nema dovoljno podataka za izveštaj. Proveri da li postoje unosi u UNOS_TESTOVA i DNEVNIK_UNOS.")

    st.markdown("### Pregled za štampu")
    html = build_a4_report_html(report, trainer_note)
    st.components.v1.html(html, height=1120, scrolling=True)

    st.download_button(
        "Preuzmi A4 izveštaj kao HTML",
        data=html.encode("utf-8"),
        file_name=f"izvestaj_{selected_email}_{start_date}_{end_date}.html".replace("@", "_"),
        mime="text/html",
        use_container_width=True,
    )

    with st.expander("Detaljne tabele"):
        st.write("Napredak u testovima")
        st.dataframe(report["progress"], use_container_width=True, hide_index=True)
        st.write("Najbolji rezultati / volumen po vežbi")
        st.dataframe(report["best_lifts"], use_container_width=True, hide_index=True)
        st.write("Volumen po nedeljama")
        st.dataframe(report["weekly"], use_container_width=True, hide_index=True)


# ------------------------------------------------------------
# 9) ADMIN DEO
# ------------------------------------------------------------

def is_admin_user(user: Dict[str, object]) -> bool:
    """Admin pristup možeš dati na dva načina:
    1) u Google Sheets listu VEZBACI dodaj kolonu Uloga=Admin ili Admin=Da
    2) u Streamlit Secrets dodaj:
       [admin]
       emails = "tvoj@email.com, drugi@email.com"
    """
    email = normalize_email(user.get("Email"))

    # 1) preko kolona u VEZBACI
    uloga = normalize_text(user.get("Uloga", "")).lower()
    admin_flag = normalize_text(user.get("Admin", "")).lower()
    if uloga in ["admin", "administrator"] or admin_flag in ["da", "yes", "1", "true"]:
        return True

    # 2) preko Secrets
    try:
        admin_cfg = st.secrets.get("admin", {})
        emails_raw = admin_cfg.get("emails", "") if hasattr(admin_cfg, "get") else ""
        if isinstance(emails_raw, str):
            admin_emails = [normalize_email(x) for x in emails_raw.split(",") if normalize_email(x)]
        else:
            admin_emails = [normalize_email(x) for x in list(emails_raw)]
        return email in admin_emails
    except Exception:
        return False


def _sheet_headers_and_indexes(sheet_name: str) -> Tuple[List[str], Dict[str, int], List[List[str]]]:
    ws = get_ws(sheet_name)
    values = ws.get_all_values()
    header_row = HEADER_ROWS.get(sheet_name, 1)
    if len(values) < header_row:
        return [], {}, values
    headers = [normalize_text(h) for h in values[header_row - 1]]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    return headers, idx, values


def find_row_numbers_by_email(sheet_name: str, email: str) -> List[int]:
    headers, idx, values = _sheet_headers_and_indexes(sheet_name)
    if "Email" not in idx:
        return []
    email_col_zero = idx["Email"] - 1
    header_row = HEADER_ROWS.get(sheet_name, 1)
    target = normalize_email(email)
    rows = []
    for sheet_row_number, row in enumerate(values[header_row:], start=header_row + 1):
        cell = row[email_col_zero] if email_col_zero < len(row) else ""
        if normalize_email(cell) == target:
            rows.append(sheet_row_number)
    return rows


def delete_rows_by_email(sheet_name: str, email: str) -> int:
    ws = get_ws(sheet_name)
    rows = find_row_numbers_by_email(sheet_name, email)
    # Brisanje ide od dna da se brojevi redova ne pomere.
    for r in sorted(rows, reverse=True):
        ws.delete_rows(r)
    clear_caches()
    return len(rows)


def update_vezbac_aktivan(email: str, active_value: str) -> bool:
    ws = get_ws(SHEET_VEZBACI)
    headers, idx, values = _sheet_headers_and_indexes(SHEET_VEZBACI)
    if "Email" not in idx or "Aktivan" not in idx:
        return False
    rows = find_row_numbers_by_email(SHEET_VEZBACI, email)
    if not rows:
        return False
    for r in rows:
        ws.update_cell(r, idx["Aktivan"], active_value)
    clear_caches()
    return True


def mask_sifra(x: Any) -> str:
    s = normalize_sifra(x)
    if not s:
        return ""
    return "•" * max(4, len(s))


def show_admin_screen(user: Dict[str, object]) -> None:
    if not is_admin_user(user):
        st.error("Ovaj ekran je dostupan samo administratoru.")
        return

    st.header("Admin panel")
    st.caption("Ovde možeš da pregledaš vežbače, blokiraš članarinu i brišeš testove/dnevnik. Za trajnu evidenciju je bolje deaktivirati vežbača nego brisati redove.")

    data = load_core_data()
    vezbaci = data["vezbaci"].copy()
    testovi = data["testovi_unos"].copy()
    dnevnik = data["dnevnik"].copy()

    if vezbaci.empty or "Email" not in vezbaci.columns:
        st.warning("List VEZBACI je prazan ili nema kolonu Email.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Vežbači", len(vezbaci))
    c2.metric("Unosi testova", len(testovi))
    c3.metric("Unosi treninga", len(dnevnik))

    tab_pregled, tab_izvestaj, tab_brisanje = st.tabs(["Pregled i članarina", "A4 izveštaj", "Brisanje podataka"])

    with tab_pregled:
        st.subheader("Pregled vežbača")
        show_cols = [c for c in ["ID", "Ime i Prezime", "Email", "Telefon", "Aktivan", "Cilj", "Program_ID", "Uloga", "Sifra"] if c in vezbaci.columns]
        preview = vezbaci[show_cols].copy() if show_cols else vezbaci.copy()
        if "Sifra" in preview.columns:
            preview["Sifra"] = preview["Sifra"].apply(mask_sifra)
        st.dataframe(preview, use_container_width=True, hide_index=True)

        emails = sorted([normalize_email(e) for e in vezbaci["Email"].dropna().tolist() if normalize_email(e)])
        if not emails:
            st.warning("Nema email adresa u listu VEZBACI.")
            return

        st.subheader("Članarina / aktivacija")
        selected_email = st.selectbox("Izaberi vežbača", emails, key="admin_status_email")
        vrow = vezbaci[vezbaci["Email"].apply(normalize_email) == selected_email]
        current_name = vrow.iloc[0].get("Ime i Prezime", "") if not vrow.empty else ""
        current_active = vrow.iloc[0].get("Aktivan", "") if not vrow.empty else ""
        tests_count = len(testovi[testovi["Email"].apply(normalize_email) == selected_email]) if "Email" in testovi.columns else 0
        logs_count = len(dnevnik[dnevnik["Email"].apply(normalize_email) == selected_email]) if "Email" in dnevnik.columns else 0
        st.info(f"Izabran: **{current_name}** · {selected_email} · Aktivan: **{current_active}** · Testovi: **{tests_count}** · Unosi treninga: **{logs_count}**")

        a1, a2 = st.columns(2)
        with a1:
            if st.button("Blokiraj članarinu — Aktivan = Ne", use_container_width=True):
                if update_vezbac_aktivan(selected_email, "Ne"):
                    st.success("Vežbač je blokiran.")
                    st.rerun()
                else:
                    st.error("Nisam uspeo da promenim status.")
        with a2:
            if st.button("Aktiviraj članarinu — Aktivan = Da", use_container_width=True):
                if update_vezbac_aktivan(selected_email, "Da"):
                    st.success("Vežbač je aktiviran.")
                    st.rerun()
                else:
                    st.error("Nisam uspeo da promenim status.")

    with tab_izvestaj:
        show_admin_report_screen(vezbaci, testovi, dnevnik)

    with tab_brisanje:
        emails = sorted([normalize_email(e) for e in vezbaci["Email"].dropna().tolist() if normalize_email(e)])
        if not emails:
            st.warning("Nema email adresa u listu VEZBACI.")
            return
        st.subheader("Brisanje podataka")
        st.warning("Brisanje je trajno: redovi se uklanjaju iz Google Sheets-a. Za vežbača koji je prestao da trenira obično je bolje staviti Aktivan = Ne.")
        selected_email = st.selectbox("Izaberi vežbača za brisanje", emails, key="admin_delete_email")
        confirm = st.checkbox(f"Potvrđujem da želim da brišem podatke za: {selected_email}")
        confirm_text = st.text_input("Za potvrdu upiši OBRISI", value="")
        can_delete = confirm and normalize_text(confirm_text).upper() == "OBRISI"

        d1, d2, d3 = st.columns(3)
        with d1:
            if st.button("Obriši samo testove", disabled=not can_delete, use_container_width=True):
                n = delete_rows_by_email(SHEET_TEST_INPUT, selected_email)
                st.success(f"Obrisano redova iz UNOS_TESTOVA: {n}")
                st.rerun()
        with d2:
            if st.button("Obriši samo dnevnik", disabled=not can_delete, use_container_width=True):
                n = delete_rows_by_email(SHEET_DNEVNIK, selected_email)
                st.success(f"Obrisano redova iz DNEVNIK_UNOS: {n}")
                st.rerun()
        with d3:
            if st.button("Obriši vežbača + sve podatke", disabled=not can_delete, use_container_width=True):
                n_tests = delete_rows_by_email(SHEET_TEST_INPUT, selected_email)
                n_logs = delete_rows_by_email(SHEET_DNEVNIK, selected_email)
                n_user = delete_rows_by_email(SHEET_VEZBACI, selected_email)
                st.success(f"Obrisano: VEZBACI {n_user}, TESTOVI {n_tests}, DNEVNIK {n_logs}")
                st.rerun()


# ------------------------------------------------------------
# 10) GLAVNI TOK APLIKACIJE
# ------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🏋️", layout="wide")
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if not st.session_state["logged_in"]:
        show_login_screen()
        return

    user = st.session_state["user"]
    enforce_paywall(user)

    with st.sidebar:
        st.title("Meni")
        st.write(f"**{user.get('Ime i Prezime', '')}**")
        st.write(user.get("Email", ""))
        if st.button("Odjavi se"):
            st.session_state.clear()
            st.rerun()
        pages = ["Današnji trening", "Testovi", "Moji podaci"]
        if is_admin_user(user):
            pages.append("Admin")
        page = st.radio("Izaberi ekran", pages)
        if st.button("Osveži podatke"):
            clear_caches()
            st.rerun()

    data = load_core_data()
    email = normalize_email(user.get("Email"))
    plan_exercises = all_plan_exercises(data["plan"])
    missing = missing_tests_for_user(email, plan_exercises, data["testovi_unos"])

    if missing and page == "Današnji trening":
        st.warning("Za potpun plan nedostaju testovi za neke vežbe. Možeš ipak otvoriti trening, ali će za te vežbe pisati 'Čeka test'.")

    if page == "Današnji trening":
        show_today_training_screen(user)
    elif page == "Testovi":
        show_initial_tests_screen(user)
    elif page == "Admin":
        show_admin_screen(user)
    else:
        st.header("Moji podaci")
        safe_user = {k: v for k, v in user.items() if k not in ["Sifra", "Poruka_Blokada", "__email", "__sifra"]}
        st.json(safe_user)
        st.caption("Šifra se ne prikazuje na ovom ekranu.")


if __name__ == "__main__":
    main()
