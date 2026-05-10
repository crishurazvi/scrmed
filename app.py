from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

from src.anonymizer import anonymize_text
from src.letter_utils import LetterData, build_docx, build_pdf, build_plain_letter, treatment_rows_from_dataframe
from src.options_store import add_unique_item, labels, load_options, save_options, strip_count

APP_DIR = Path(__file__).parent
OPTIONS_PATH = APP_DIR / "data" / "options.json"

st.set_page_config(page_title="Generator scrisoare medicală", page_icon="🫀", layout="wide")


def init_state() -> None:
    defaults = {
        "treatment_df": pd.DataFrame(columns=["Medicament", "Dimineața", "Prânz", "Seara", "Observații"]),
        "manual_diagnoses": "",
        "anamneza": "",
        "alte_investigatii": "De efectuat în ambulator peste 6 luni: hemoleucograma, ionograma, creatinina, acid uric, colesterol total, LDL, trigliceride, ASAT, ALAT, glicemie, HbA1c",
        "recomandari": "Control peste 6 luni, cu programare.",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_form() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith("form_") or key in {"treatment_df", "manual_diagnoses", "anamneza", "alte_investigatii", "recomandari"}:
            del st.session_state[key]
    init_state()


def append_text_state(key: str, text: str) -> None:
    current = st.session_state.get(key, "")
    if current and not current.endswith((" ", "\n")):
        current += "\n"
    st.session_state[key] = current + text


def add_treatment_row(med: str) -> None:
    med = med.strip()
    if not med:
        return
    df = st.session_state.treatment_df.copy()
    df.loc[len(df)] = [med, "", "", "", ""]
    st.session_state.treatment_df = df


init_state()
options = load_options(OPTIONS_PATH)

diagnosis_options = labels(options.get("diagnoses", []), include_counts=True)
medication_options = labels(options.get("medications", []), include_counts=True)
phrase_items = options.get("standard_phrases", [])
investigation_items = options.get("investigations", [])

st.title("Generator rapid de scrisoare medicală")
st.caption("Sugestiile sunt extrase din exemplele încărcate și servesc doar pentru redactare. Diagnosticul și tratamentul trebuie confirmate de medic.")

with st.sidebar:
    st.header("Date administrative")
    furnizor = st.text_input("Furnizor", value="Institutul de Boli Cardiovasculare TIMIȘOARA", key="form_furnizor")
    medic = st.text_input("Medic", value="Dr. OLARIU IOAN", key="form_medic")
    contract = st.text_input("Contract/convenție", value="II/AS/6/2023", key="form_contract")
    casa = st.text_input("CAS", value="CAS Timiș", key="form_casa")
    st.divider()
    st.header("Pacient")
    pacient_nume = st.text_input("Nume pacient *", key="form_pacient_nume")
    data_nastere = st.text_input("Data nașterii", placeholder="zz.ll.aaaa", key="form_data_nastere")
    cnp = st.text_input("CNP / cod unic", type="password", help="Nu salva date reale pe cloud public. Pentru testare folosește date anonimizate.", key="form_cnp")
    data_consult = st.text_input("Data consultului *", value=date.today().strftime("%d.%m.%Y"), key="form_data_consult")
    nr_registru = st.text_input("Nr. registru consultații", key="form_nr_registru")
    oncologic = st.radio("Pacient oncologic", ["NU", "DA"], horizontal=True, key="form_oncologic")
    st.button("Resetează formularul", on_click=reset_form, use_container_width=True)

main_tab, lists_tab, privacy_tab = st.tabs(["Scrisoare", "Liste editabile", "Anonimizare"])

with main_tab:
    c1, c2 = st.columns([1.05, 1])

    with c1:
        st.subheader("1. Motiv, diagnostice, anamneză")
        motiv = st.text_input("Motivele prezentării", value="Evaluare cardiologică", key="form_motiv")

        selected_diag_labels = st.multiselect(
            "Diagnostice frecvente din exemple",
            options=diagnosis_options,
            help="Bifează doar diagnostice confirmate clinic. Aplicația nu pune diagnostice automat.",
            key="form_diagnoses",
        )
        selected_diagnoses = [strip_count(x) for x in selected_diag_labels]
        st.text_area("Diagnostice manuale / completări", key="manual_diagnoses", height=120)

        with st.expander("Butoane rapide pentru fraze standard", expanded=False):
            cols = st.columns(2)
            for i, item in enumerate(phrase_items[:12]):
                text = item.get("text") or item.get("label") or ""
                label = item.get("label") or text[:50]
                if cols[i % 2].button(label, key=f"phrase_{i}"):
                    target = "alte_investigatii" if "EKG" in text.upper() or "ECO" in text.upper() or "AMBULATOR" in text.upper() else "recomandari"
                    append_text_state(target, text)

        st.text_area("Anamneză", key="anamneza", height=100)
        factori_risc = st.text_area("Factori de risc", key="form_factori_risc", height=70)

        st.subheader("2. Examen clinic și investigații")
        col_ta, col_fc = st.columns(2)
        ta = col_ta.text_input("TA", placeholder="140/80", key="form_ta")
        fc = col_fc.text_input("FC", placeholder="70", key="form_fc")
        examen_local = st.text_area("Examen local", key="form_examen_local", height=70)
        laborator_normal = st.text_area("Laborator cu valori normale", key="form_laborator_normal", height=70)
        laborator_patologic = st.text_area("Laborator cu valori patologice", key="form_laborator_patologic", height=70)

        ekg = st.text_input("EKG", value="EKG anexat", key="form_ekg")
        eco = st.text_input("ECO", value="ECO", key="form_eco")
        rx = st.text_area("Rx / consulturi / investigații specifice", key="form_rx", height=70)
        st.text_area("Altele / analize de efectuat", key="alte_investigatii", height=90)

        with st.expander("Sugestii de investigații extrase", expanded=False):
            for i, item in enumerate(investigation_items[:12]):
                text = item.get("text") or item.get("label") or ""
                if st.button(text[:90], key=f"inv_{i}"):
                    append_text_state("alte_investigatii", text)

    with c2:
        st.subheader("3. Tratament recomandat")
        med_selected = st.selectbox("Caută / alege medicament din exemple", [""] + medication_options, key="form_med_select")
        med_manual = st.text_input("Sau introdu medicament nou", key="form_med_manual")
        add_col1, add_col2 = st.columns(2)
        if add_col1.button("Adaugă medicament", use_container_width=True):
            med_to_add = med_manual.strip() or strip_count(med_selected)
            add_treatment_row(med_to_add)
        if add_col2.button("Golește tabel tratament", use_container_width=True):
            st.session_state.treatment_df = pd.DataFrame(columns=["Medicament", "Dimineața", "Prânz", "Seara", "Observații"])

        st.session_state.treatment_df = st.data_editor(
            st.session_state.treatment_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Medicament": st.column_config.TextColumn(required=False),
                "Dimineața": st.column_config.TextColumn(width="small"),
                "Prânz": st.column_config.TextColumn(width="small"),
                "Seara": st.column_config.TextColumn(width="small"),
                "Observații": st.column_config.TextColumn(width="medium"),
            },
            key="form_treatment_editor",
        )

        tratament_efectuat = st.text_area("Tratament efectuat", key="form_tratament_efectuat", height=70)
        alte_info = st.text_area("Alte informații", value="Scrisoare medicală valabilă 6 luni", key="form_alte_info", height=70)
        st.text_area("Recomandări / control", key="recomandari", height=90)

        st.subheader("4. Administrativ final")
        prescriptie_status = st.selectbox(
            "Prescripție medicală",
            ["Nu s-a eliberat prescripție medicală", "Nu s-a eliberat prescripție medicală deoarece nu a fost necesar", "S-a eliberat prescripție medicală"],
            key="form_prescriptie_status",
        )
        concediu_status = st.selectbox(
            "Concediu medical",
            ["Nu s-a eliberat concediu medical la externare", "Nu s-a eliberat concediu medical la externare deoarece nu a fost necesar", "S-a eliberat concediu medical"],
            key="form_concediu_status",
        )
        ingrijiri_status = st.selectbox(
            "Îngrijiri la domiciliu",
            ["Nu s-a eliberat recomandare pentru îngrijiri medicale la domiciliu/paliative la domiciliu, deoarece nu a fost necesar", "S-a eliberat recomandare pentru îngrijiri medicale la domiciliu/paliative la domiciliu"],
            key="form_ingrijiri_status",
        )
        dispozitive_status = st.selectbox(
            "Dispozitive medicale",
            ["Nu s-a eliberat prescripție medicală pentru dispozitive medicale în ambulatoriu deoarece nu a fost necesar", "S-a eliberat prescripție medicală pentru dispozitive medicale în ambulatoriu"],
            key="form_dispozitive_status",
        )
        cale_transmitere = st.selectbox("Calea de transmitere", ["prin asigurat", "prin poștă", "prin poștă electronică"], key="form_cale_transmitere")

    data = LetterData(
        furnizor=furnizor,
        medic=medic,
        contract=contract,
        casa=casa,
        pacient_nume=pacient_nume,
        data_nastere=data_nastere,
        cnp=cnp,
        data_consult=data_consult,
        nr_registru=nr_registru,
        motiv=motiv,
        pacient_oncologic=oncologic,
        diagnostice=selected_diagnoses,
        diagnostice_manual=st.session_state.manual_diagnoses,
        anamneza=st.session_state.anamneza,
        factori_risc=factori_risc,
        ta=ta,
        fc=fc,
        examen_local=examen_local,
        laborator_normal=laborator_normal,
        laborator_patologic=laborator_patologic,
        ekg=ekg,
        eco=eco,
        rx=rx,
        alte_investigatii=st.session_state.alte_investigatii,
        tratament_efectuat=tratament_efectuat,
        alte_informatii=alte_info,
        tratament=treatment_rows_from_dataframe(st.session_state.treatment_df),
        recomandari=st.session_state.recomandari,
        prescriptie_status=prescriptie_status,
        concediu_status=concediu_status,
        ingrijiri_status=ingrijiri_status,
        dispozitive_status=dispozitive_status,
        cale_transmitere=cale_transmitere,
    )

    st.divider()
    st.subheader("Previzualizare live")

    required_errors = []
    if not pacient_nume.strip():
        required_errors.append("Nume pacient")
    if not data_consult.strip():
        required_errors.append("Data consultului")
    if not (selected_diagnoses or st.session_state.manual_diagnoses.strip()):
        required_errors.append("Diagnostic")
    if required_errors:
        st.warning("Câmpuri obligatorii necompletate: " + ", ".join(required_errors))

    preview = build_plain_letter(data)
    st.markdown(preview)

    export_col1, export_col2, export_col3 = st.columns(3)
    safe_name = (pacient_nume.strip().replace(" ", "_") or "scrisoare_medicala")
    export_col1.download_button(
        "Descarcă DOCX",
        data=build_docx(data),
        file_name=f"{safe_name}_scrisoare_medicala.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    export_col2.download_button(
        "Descarcă PDF simplu",
        data=build_pdf(data),
        file_name=f"{safe_name}_scrisoare_medicala.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
    export_col3.download_button(
        "Descarcă TXT",
        data=preview.encode("utf-8"),
        file_name=f"{safe_name}_scrisoare_medicala.txt",
        mime="text/plain",
        use_container_width=True,
    )

with lists_tab:
    st.subheader("Liste editabile")
    st.caption("Modificările salvate local se scriu în data/options.json. Pe Streamlit Community Cloud, stocarea poate fi resetată la redeploy sau restart.")

    list_type = st.selectbox(
        "Listă",
        ["diagnoses", "medications", "standard_phrases", "investigations"],
        format_func={
            "diagnoses": "Diagnostice",
            "medications": "Medicamente",
            "standard_phrases": "Fraze standard",
            "investigations": "Investigații / recomandări de analize",
        }.get,
    )
    current_items = options.get(list_type, [])
    df_items = pd.DataFrame(current_items)
    edited_df = st.data_editor(df_items, num_rows="dynamic", use_container_width=True, key=f"edit_{list_type}")

    new_item = st.text_input("Adaugă rapid element nou")
    col_a, col_b, col_c = st.columns(3)
    if col_a.button("Adaugă în listă", use_container_width=True):
        key_name = "text" if list_type in {"standard_phrases", "investigations"} else "label"
        if add_unique_item(options, list_type, new_item, label_key=key_name):
            save_options(OPTIONS_PATH, options)
            st.success("Element adăugat.")
            st.rerun()
        else:
            st.info("Element gol sau deja existent.")

    if col_b.button("Salvează tabelul editat", use_container_width=True):
        cleaned = edited_df.fillna("").to_dict(orient="records")
        options[list_type] = cleaned
        save_options(OPTIONS_PATH, options)
        st.success("Lista a fost salvată local.")
        st.rerun()

    col_c.download_button(
        "Exportă JSON opțiuni",
        data=json.dumps(options, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="options_scrisoare_medicala.json",
        mime="application/json",
        use_container_width=True,
    )

with privacy_tab:
    st.subheader("Anonimizare text copiat din documente")
    st.caption("Este o anonimizare tehnică minimă. Verificarea manuală rămâne obligatorie înainte de a urca date pe cloud public.")
    raw_text = st.text_area("Lipește textul de anonimizat", height=220)
    mask_dates = st.checkbox("Anonimizează și datele calendaristice")
    if st.button("Anonimizează"):
        st.session_state["anon_text"] = anonymize_text(raw_text, mask_dates=mask_dates)
    st.text_area("Rezultat anonimizat", value=st.session_state.get("anon_text", ""), height=220)
