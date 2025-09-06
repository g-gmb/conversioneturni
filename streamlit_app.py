import io
import re
import runpy
import uuid
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Convertitore Turni", page_icon="🗓️", layout="centered")
st.title("Convertitore Turni")
st.caption(
    "Carica il tuo file Excel, inserisci il cognome e scarica il risultato in formato CSV o ICS per importarlo sul calendario."
)

# --- UI --------------------------------------------------------------------
uploaded_xlsx = st.file_uploader("Carica il file Excel", type=["xlsx"], accept_multiple_files=False)
surname_input = st.text_input("Cognome", help="Scrivi il tuo cognome completo. Non c'è bisogno di differenziare maiuscole/minuscole.")
run_btn = st.button("Converti")

# Helper per sanitizzare i nomi file
def _sanitize(s: str) -> str:
    s = (s or "").strip()
    s = s.replace(" ", "_")
    return re.sub(r"[^A-Za-z0-9_-]", "", s)

# --- Core runner ------------------------------------------------------------
def run_conversion_script(script_path: Path, excel_path: Path, surname: str):
    init_globals = {
        "__name__": "__main__",
        "pd": pd,
        "surname": surname,
        "excel_file_path": str(excel_path),
    }

    exec_globals = runpy.run_path(str(script_path), init_globals=init_globals)
    exec_globals["surname"] = surname

    df_final = exec_globals.get("df_final")
    return df_final, exec_globals

# --- Generatore ICS (manuale) -----------------------------------------------
def csv_text_to_ics(csv_text: str) -> tuple[str, int]:
    df = pd.read_csv(io.StringIO(csv_text))

    subj_col = "Subject"
    sd = "Start Date"
    stime = "Start Time"
    ed = "End Date"
    etime = "End Time"

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Turni2ICS//EN"
    ]

    count = 0
    for _, row in df.iterrows():
        try:
            subject = str(row[subj_col]).strip()

            start_str = f"{row[sd]} {row[stime]}"
            end_str = f"{row[ed]} {row[etime]}"

            # Parsing date/time (giorno primo, orari 12h)
            start_dt = pd.to_datetime(start_str, dayfirst=True, errors="coerce")
            end_dt = pd.to_datetime(end_str, dayfirst=True, errors="coerce")

            if pd.isna(start_dt) or pd.isna(end_dt):
                continue

            # ICS richiede formato: YYYYMMDDTHHMMSS
            dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
            dtstart = start_dt.strftime("%Y%m%dT%H%M%S")
            dtend = end_dt.strftime("%Y%m%dT%H%M%S")
            uid = str(uuid.uuid4())

            lines.extend([
                "BEGIN:VEVENT",
                f"DTSTAMP:{dtstamp}",
                f"DTSTART:{dtstart}",
                f"DTEND:{dtend}",
                f"SUMMARY:{subject}",
                f"UID:{uid}",
                "END:VEVENT"
            ])
            count += 1
        except Exception:
            continue

    lines.append("END:VCALENDAR")
    return "\n".join(lines), count

# --- Azione principale ------------------------------------------------------
if run_btn:
    if not uploaded_xlsx:
        st.error("Per favore carica un file .xlsx prima di procedere.")
        st.stop()

    if not surname_input.strip():
        st.error("Per favore inserisci il tuo cognome.")
        st.stop()

    tmp_dir = Path(st.session_state.get("_tmp_dir", ".tmp_uploads"))
    tmp_dir.mkdir(exist_ok=True)
    xlsx_name = Path(uploaded_xlsx.name).stem
    excel_path = tmp_dir / f"{_sanitize(xlsx_name)}.xlsx"
    excel_bytes = uploaded_xlsx.read()
    excel_path.write_bytes(excel_bytes)

    surname = surname_input.strip().lower()

    app_dir = Path(__file__).parent
    script_path = app_dir / "conversione_turni.py"

    try:
        if not script_path.exists():
            st.warning("`conversione_turni.py` non è stato trovato accanto a questa app. Verrà utilizzato direttamente il contenuto del file Excel caricato come `df_final`.")
            df_final = pd.read_excel(excel_path)
        else:
            with st.spinner("Esecuzione di conversione_turni.py..."):
                try:
                    df_final, _ = run_conversion_script(script_path, excel_path, surname)
                except NameError:
                    st.error("⚠️ Cognome non trovato nel file Excel.")
                    st.stop()

        if not isinstance(df_final, pd.DataFrame):
            try:
                df_final = pd.DataFrame(df_final)
            except Exception:
                st.error("`df_final` non può essere convertito in DataFrame. Assicurati che lo script produca un pandas DataFrame chiamato `df_final`.")
                st.stop()

        st.success("Conversione completata!")
        st.dataframe(df_final, use_container_width=True)

        out_name_csv = f"{_sanitize(xlsx_name)}_{_sanitize(surname)}.csv"
        out_name_ics = f"{_sanitize(xlsx_name)}_{_sanitize(surname)}.ics"

        csv_buf = io.StringIO()
        df_final.to_csv(csv_buf, index=False)
        csv_text = csv_buf.getvalue()
        st.download_button(
            label=f"Scarica {out_name_csv}",
            data=csv_text.encode("utf-8"),
            file_name=out_name_csv,
            mime="text/csv",
        )

        try:
            ics_str, n_events = csv_text_to_ics(csv_text)
            st.download_button(
                label=f"Scarica {out_name_ics}",
                data=ics_str.encode("utf-8"),
                file_name=out_name_ics,
                mime="text/calendar",
            )
            st.caption(f"Eventi nel file ICS: **{n_events}**")
            if n_events == 0:
                st.warning("Nessun evento rilevato per l'ICS. Controlla che le intestazioni del CSV rispettino lo schema (Subject, Start Date/Time, End Date/Time, ecc.).")
        except Exception as e:
            st.warning(f"Impossibile generare ICS dal CSV: {e}")

        save_path_csv = app_dir / out_name_csv
        try:
            save_path_csv.write_text(csv_text, encoding="utf-8")
            st.caption(f"Copia CSV salvata in: `{save_path_csv}`")
        except Exception:
            pass

    except Exception as e:
        st.exception(e)
        st.stop()

# --- Pannello aiuto ---------------------------------------------------------
with st.expander("Come funziona / Note"):
    st.markdown(
        """
        ### Come usare l'app

        1. **Carica il file Excel** con i tuoi turni utilizzando il pulsante in alto.
        2. **Inserisci il tuo cognome** nella casella di testo.
        3. Premi **Esegui conversione**.
        4. Dopo l'elaborazione puoi:
            - **Scaricare il file CSV** → questo formato è accettato direttamente da **Google Calendar**.
                - Apri Google Calendar sul web
                - Vai su *Impostazioni → Importa*
                - Seleziona il CSV scaricato e importa nel calendario desiderato
            - **Scaricare il file ICS** → questo formato è universale e può essere importato in molti altri calendari, ad esempio:
                - **Apple Calendar (macOS/iOS)**
                - **Outlook (Windows)**
                - **Thunderbird Lightning**
                - Altri gestori di calendari compatibili con ICS

        In questo modo puoi avere i tuoi turni sincronizzati nel calendario che preferisci.
        """
    )

# --- Copyright --------------------------------------------------------------
st.markdown("<div style='text-align: right; font-size: small; color: gray;'>© Gioele Gambato</div>", unsafe_allow_html=True)
