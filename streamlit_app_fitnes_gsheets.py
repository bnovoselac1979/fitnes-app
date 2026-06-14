# streamlit_app_fitnes_gsheets.py
# Fitnes klub aplikacija preko Google Sheets baze
# Radi na Streamlit Cloud-u uz gspread + Streamlit secrets.

from __future__ import annotations

import re
from datetime import datetime, date
from typing import Dict, List, Optional, Tuple

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ------------------------------------------------------------
# 1) OSNOVNA PODEŠAVANJA
# ------------------------------------------------------------

APP_TITLE = "Fitnes klub — 5/3/1 + RIR"

SHEET_VEZBACI = "VEZBACI"
SHEET_BAZA_VEZBI = "BAZA_VEZBI"
SHEET_TEST_INPUT = "UNOS_TESTOVA"   # app upisuje testove ovde
SHEET_TEST_CALC = "TESTOVI"         # može ostati kao obračunski/pregledni list u Google Sheets
SHEET_DNEVNIK = "DNEVNIK_UNOS"
SHEET_PLAN = "A4_MESEC_KOMPAKT"

# Brojevi redova gde su zaglavlja u Google Sheet-u.
# Ako kasnije pojednostaviš Google Sheet pa svuda zaglavlja budu u prvom redu,
# ovde samo promeni vrednosti na 1.
HEADER_ROWS = {
    SHEET_VEZBACI: 2,
    SHEET_BAZA_VEZBI: 2,
    SHEET_TEST_INPUT: 2,
    SHEET_TEST_CALC: 3,
    SHEET_DNEVNIK: 3,
    SHEET_PLAN: 4,
}

REQUIRED_VEZBACI_COLS = [
    "ID", "Ime i Prezime", "Aktivan", "Email", "Sifra", "Poruka_Blokada"
]

REQUIRED_TEST_COLS = [
    "Datum", "Email", "Naziv vežbe", "Kilaža (kg)", "Broj ponavljanja",
    "Procena 1RM", "ID (auto)", "Ime (auto)", "Napomena"
]

REQUIRED_DNEVNIK_COLS = [
    "Email", "Timestamp", "Datum", "ID", "Ime", "Nedelja", "Trening", "Vežba",
    "Plan kg", "Plan ser.", "Plan reps", "Urađeno kg", "Urađeno ser.",
    "Urađeno reps", "RIR stvarni", "Status", "Napomena", "Volumen (kg)"
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ------------------------------------------------------------
# 2) POMOĆNE FUNKCIJE
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
    # Google Sheets nekad broj pročita kao 3069.0 ili 3069
    s = normalize_text(x)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def estimate_1rm(kg: float, reps: int) -> float:
    """Epley formula: kg * (1 + reps/30)."""
    kg = float(kg or 0)
    reps = int(reps or 1)
    if reps <= 1:
        return round(kg, 1)
    return round(kg * (1 + reps / 30), 1)


def extract_first_number(value) -> Optional[float]:
    """Iz teksta tipa '80×5 / 90×3' uzima prvi broj kao plan kg."""
    if value is None:
        return None
    text = str(value).replace(",", ".")
    match = re.search(r"(\d+(?:\.\d+)?)", text)
    return float(match.group(1)) if match else None


def to_sheet_value(value):
    """Pretvara Python vrednosti u format koji Google Sheets lepo prima."""
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
    """
    Podržava dva načina za secrets:
    1) preporučeno u ovom projektu:
       [gcp_service_account]
       type = "service_account"
       ...
       [gsheets]
       spreadsheet_url = "..."

    2) Streamlit/GSheetsConnection stil:
       [connections.gsheets]
       spreadsheet = "..."
       type = "service_account"
       ...
    """
    if "gcp_service_account" in st.secrets:
        return dict(st.secrets["gcp_service_account"])

    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        conn_secret = dict(st.secrets["connections"]["gsheets"])
        # uklanjamo polja koja nisu deo JSON ključa
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

    # Privatni ključ u TOML-u često mora da ima \n; ovde ga vraćamo u pravi oblik.
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
    """Čita list iz Google Sheets-a, uzimajući u obzir red zaglavlja."""
    ws = get_ws(sheet_name)
    values = ws.get_all_values()
    header_row = HEADER_ROWS.get(sheet_name, 1)

    if len(values) < header_row:
        return pd.DataFrame()

    headers = [str(h).strip() for h in values[header_row - 1]]
    data_rows = values[header_row:]

    # ukloni potpuno prazne redove
    data_rows = [r for r in data_rows if any(str(c).strip() for c in r)]

    if not headers:
        return pd.DataFrame()

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
    headers = ws.row_values(header_row)
    return [str(h).strip() for h in headers]


def append_rows_to_sheet(sheet_name: str, rows: List[Dict[str, object]]) -> None:
    """Dodaje redove u Google Sheets po imenima kolona iz header reda."""
    if not rows:
        return

    ws = get_ws(sheet_name)
    headers = get_headers(sheet_name)
    if not headers:
        st.error(f"List {sheet_name} nema definisana zaglavlja.")
        st.stop()

    prepared_rows = []
    for row in rows:
        prepared_rows.append([to_sheet_value(row.get(h, "")) for h in headers])

    ws.append_rows(prepared_rows, value_input_option="USER_ENTERED")
    clear_caches()

# ------------------------------------------------------------
# 4) LOGOVANJE I PAYWALL
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

    with st.form("login_form"):
        email = st.text_input("Email", placeholder="npr. ana@email.com")
        sifra = st.text_input("Šifra", type="password", placeholder="4 broja")
        submitted = st.form_submit_button("Uđi u aplikaciju")

    if submitted:
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
# 5) TESTOVI
# ------------------------------------------------------------

def user_has_tests(email: str, testovi_unos: pd.DataFrame) -> bool:
    if testovi_unos.empty or "Email" not in testovi_unos.columns:
        return False
    df = testovi_unos.copy()
    df["__email"] = df["Email"].apply(normalize_email)
    return (df["__email"] == normalize_email(email)).any()


def _exercise_column_name(vezbe: pd.DataFrame) -> str:
    for possible in ["Vežba", "Naziv vežbe", "Vezba", "Naziv"]:
        if possible in vezbe.columns:
            return possible
    st.error("U listu BAZA_VEZBI ne nalazim kolonu sa nazivom vežbe. Očekujem 'Vežba' ili 'Naziv vežbe'.")
    st.stop()


def show_initial_tests_screen(user: Dict[str, object]) -> None:
    st.header("Početni testovi")
    st.write("Unesite kilažu i broj ponavljanja. Aplikacija automatski računa procenjeni 1RM.")

    data = load_core_data()
    vezbe = data["vezbe"]
    exercise_col = _exercise_column_name(vezbe)

    exercise_options = sorted([
        v for v in vezbe[exercise_col].dropna().astype(str).unique().tolist()
        if v.strip()
    ])

    default_exercises = [v for v in ["Čučanj", "Bench press", "Iskorak", "Rameni potisak"] if v in exercise_options]

    selected = st.multiselect(
        "Izaberi vežbe za testiranje",
        options=exercise_options,
        default=default_exercises if default_exercises else exercise_options[:4],
    )

    if not selected:
        st.warning("Izaberi bar jednu vežbu.")
        return

    rows = []
    with st.form("testovi_form"):
        test_date = st.date_input("Datum testa", value=date.today())
        for ex in selected:
            st.markdown(f"**{ex}**")
            c1, c2, c3 = st.columns(3)
            kg = c1.number_input(f"Kilaža kg — {ex}", min_value=0.0, step=2.5, key=f"kg_{ex}")
            reps = c2.number_input(f"Broj ponavljanja — {ex}", min_value=1, max_value=30, value=1, step=1, key=f"reps_{ex}")
            one_rm = estimate_1rm(kg, reps) if kg > 0 else 0
            c3.metric("Procena 1RM", f"{one_rm} kg" if one_rm else "—")
            rows.append((ex, kg, reps, one_rm))

        submitted = st.form_submit_button("Sačuvaj testove")

    if submitted:
        email = normalize_email(user.get("Email"))
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
# 6) DANAŠNJI TRENING
# ------------------------------------------------------------

def get_current_week_number() -> int:
    return ((date.today().isocalendar().week - 1) % 4) + 1


def build_training_plan_for_user(user: Dict[str, object], week: int, trening: int) -> pd.DataFrame:
    data = load_core_data()
    plan = data["plan"].copy()

    required = ["Ned.", "Trening", "Glavna vežba", "Plan glavne vežbe", "Asistencije — skraćeno"]
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
        glavna = normalize_text(r.get("Glavna vežba"))
        if glavna:
            rows.append({
                "Vežba": glavna,
                "Tip": "Glavna",
                "Plan tekst": normalize_text(r.get("Plan glavne vežbe")),
                "Plan kg": extract_first_number(r.get("Plan glavne vežbe")),
                "Plan ser.": None,
                "Plan reps": None,
            })

        asist = normalize_text(r.get("Asistencije — skraćeno"))
        if asist:
            for a in [x.strip() for x in asist.split("/") if x.strip()]:
                rows.append({
                    "Vežba": a,
                    "Tip": "Asistencija",
                    "Plan tekst": "Po planu trenera",
                    "Plan kg": None,
                    "Plan ser.": None,
                    "Plan reps": None,
                })

    return pd.DataFrame(rows)


def show_today_training_screen(user: Dict[str, object]) -> None:
    st.header("Današnji trening")
    st.write("Upisuješ samo odstupanja. Ako je urađeno tačno po planu, polja mogu ostati prazna.")

    c1, c2 = st.columns(2)
    week = c1.selectbox("Nedelja ciklusa", options=[1, 2, 3, 4], index=get_current_week_number() - 1)
    trening = c2.selectbox("Trening", options=[1, 2, 3, 4], index=0)

    plan_df = build_training_plan_for_user(user, week, trening)
    if plan_df.empty:
        st.warning("Nema pronađenog plana za izabranu nedelju i trening.")
        return

    st.dataframe(plan_df[["Vežba", "Tip", "Plan tekst"]], use_container_width=True, hide_index=True)

    with st.form("dnevnik_form"):
        st.subheader("Unos odstupanja")
        entries = []
        for idx, row in plan_df.iterrows():
            ex = row["Vežba"]
            st.markdown(f"**{ex}** — {row['Tip']}")
            c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 2])
            done_kg = c1.number_input("Urađeno kg", min_value=0.0, step=2.5, key=f"done_kg_{idx}")
            done_sets = c2.number_input("Serije", min_value=0, step=1, key=f"sets_{idx}")
            done_reps = c3.number_input("Reps", min_value=0, step=1, key=f"reps_{idx}")
            rir = c4.number_input("RIR", min_value=0, max_value=10, step=1, key=f"rir_{idx}")
            note = c5.text_input("Napomena", key=f"note_{idx}")
            entries.append({
                "exercise": ex,
                "plan_kg": row.get("Plan kg"),
                "plan_sets": row.get("Plan ser."),
                "plan_reps": row.get("Plan reps"),
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
            has_change = any([e["done_kg"], e["done_sets"], e["done_reps"], e["note"]])
            if not has_change:
                continue

            kg_for_volume = e["done_kg"] or e["plan_kg"] or 0
            sets_for_volume = e["done_sets"] or e["plan_sets"] or 0
            reps_for_volume = e["done_reps"] or e["plan_reps"] or 0
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
# 7) GLAVNI TOK APLIKACIJE
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

        page = st.radio("Izaberi ekran", ["Današnji trening", "Testovi", "Moji podaci"])

        if st.button("Osveži podatke"):
            clear_caches()
            st.rerun()

    data = load_core_data()
    email = normalize_email(user.get("Email"))

    if not user_has_tests(email, data["testovi_unos"]) and page != "Moji podaci":
        st.warning("Prvo treba uneti početne testove da bi plan mogao pravilno da se računa.")
        show_initial_tests_screen(user)
        return

    if page == "Današnji trening":
        show_today_training_screen(user)
    elif page == "Testovi":
        show_initial_tests_screen(user)
    else:
        st.header("Moji podaci")
        safe_user = {k: v for k, v in user.items() if k not in ["Sifra", "Poruka_Blokada", "__email", "__sifra"]}
        st.json(safe_user)
        st.caption("Šifra se ne prikazuje na ovom ekranu.")


if __name__ == "__main__":
    main()
