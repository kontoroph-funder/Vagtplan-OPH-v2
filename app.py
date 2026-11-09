from hashlib import sha256
from datetime import datetime
from pathlib import Path

import streamlit as st

from workhours_tool.reconciliation import reconcile_workbook
from workhours_tool.scraper import apply_adjustments, export_to_excel, parse_shift_rows


st.set_page_config(page_title="Vagtplan til Excel", page_icon="📅", layout="wide")

page_style = """
<style>
body {
    background: #060a14;
    color: #f2f5fa;
}
section.main {
    padding-top: 0 !important;
}
.stApp {
    background: linear-gradient(135deg, #060a14 0%, #0a1120 45%, #101a2e 100%);
}
.stButton>button {
    border-radius: 10px;
}
.css-1kyxreq {
    background: #0d1526 !important;
}
.css-1d391kg {
    background: transparent !important;
}
.status-pill {
    background: #0f172a;
    border: 1px solid #1c2b45;
    border-radius: 12px;
    color: #d8e3ff;
    padding: 10px 16px;
    display: inline-block;
    font-size: 0.95rem;
}
.card {
    border: 1px solid #1c2b45;
    background: rgba(13, 21, 38, 0.88);
    border-radius: 18px;
    padding: 24px;
    margin-bottom: 22px;
}
.card h2 {
    margin: 0;
    font-size: 1.1rem;
}
.card .req {
    font-size: 0.9rem;
    color: #9aa5c2;
}
.card-head {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 16px;
    margin-bottom: 18px;
}
.card-head h2 {
    margin: 0;
    font-size: 1.18rem;
}
.card-head .req {
    margin-top: 6px;
}
.breadcrumbs {
    font-size: 0.92rem;
    color: #9aa5c2;
    margin-bottom: 18px;
}
.hero {
    display: flex;
    gap: 20px;
    align-items: flex-start;
    margin-bottom: 24px;
    padding: 22px;
    border: 1px solid rgba(28, 43, 69, 0.35);
    border-radius: 20px;
    background: rgba(15, 23, 42, 0.7);
}
.hero-left {
    flex: 1;
}
.hero-title {
    margin: 0 0 12px;
    font-size: 2.3rem;
    line-height: 1.05;
}
.hero-subtitle {
    margin: 0;
    color: #9aa5c2;
    max-width: 740px;
    font-size: 1rem;
    line-height: 1.7;
}
.hero-actions {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-top: 18px;
}
.field-block {
    margin-bottom: 20px;
}
.field-label {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    margin-bottom: 8px;
    font-weight: 700;
    color: #e2e8f0;
}
.field-title {
    font-size: 0.95rem;
}
.help-icon {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
    border-radius: 50%;
    border: 1px solid #2563eb;
    color: #a5b4fc;
    background: rgba(37, 99, 235, 0.1);
    font-size: 0.9rem;
    cursor: help;
}
.help-icon:hover {
    background: rgba(37, 99, 235, 0.18);
}
.btn-row {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-bottom: 22px;
}
.brand-logo {
    width: 52px;
    height: 52px;
    border-radius: 14px;
    background: #fff;
    display: grid;
    place-items: center;
    color: #0f172a;
    font-weight: 800;
    font-size: 1.15rem;
}
.brand-name {
    font-size: 1rem;
    font-weight: 700;
}
.brand-sub {
    font-size: 0.92rem;
    color: #9aa5c2;
}
.step {
    border-radius: 14px;
    border: 1px solid transparent;
    padding: 16px;
    margin-bottom: 12px;
}
.step.active {
    background: #101a2e;
    border-color: #1c2b45;
}
.step-num {
    width: 28px;
    height: 28px;
    border-radius: 50%;
    display: grid;
    place-items: center;
    font-size: 0.85rem;
    font-weight: 700;
    margin-right: 12px;
}
.step.active .step-num {
    background: #3b82f6;
    color: #fff;
}
.step-title {
    font-size: 0.95rem;
    font-weight: 600;
    margin: 0;
}
.step-desc {
    margin: 4px 0 0;
    color: #9aa5c2;
    font-size: 0.9rem;
}
.hint-card {
    background: #0f172a;
    border: 1px solid #1c2b45;
    border-radius: 16px;
    padding: 18px;
}
.hint-title {
    font-weight: 700;
    margin-bottom: 8px;
}
.hint-text {
    color: #9aa5c2;
    font-size: 0.95rem;
    line-height: 1.6;
}
.overview-card {
    background: linear-gradient(160deg, rgba(52,211,153,0.08), rgba(13,21,38,0.4));
    border: 1px solid rgba(52,211,153,0.18);
}
.summary-row {
    display: flex;
    justify-content: space-between;
    gap: 16px;
    font-size: 0.95rem;
    margin-bottom: 8px;
}
.summary-row .k {
    color: #9aa5c2;
}
.summary-row .v {
    color: #f8fafc;
    font-weight: 600;
}
.summary-row .v.muted {
    color: #9aa5c2;
    font-weight: 500;
}
.checklist {
    margin-top: 22px;
    display: flex;
    flex-direction: column;
    gap: 12px;
}
.check-item {
    display: flex;
    gap: 12px;
    color: #9aa5c2;
    font-size: 0.95rem;
    line-height: 1.5;
}
.check-mark {
    width: 22px;
    height: 22px;
    border-radius: 50%;
    background: #34d399;
    color: #08130f;
    display: grid;
    place-items: center;
    flex-shrink: 0;
    font-size: 0.9rem;
    font-weight: 800;
}
.btn-row {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-bottom: 24px;
}
.btn-ghost button,
.btn-primary button {
    width: 100%;
    border-radius: 10px;
}
.help-icon {
    margin-left: 6px;
    color: #9aa5c2;
}
</style>
"""
st.markdown(page_style, unsafe_allow_html=True)


def build_preview_rows(rows: list[dict], apply_night_reduction_to_bagvagt: bool) -> list[dict]:
    preview_rows: list[dict] = []
    for row in rows:
        adjusted = apply_adjustments(row, apply_night_reduction_to_bagvagt=apply_night_reduction_to_bagvagt)
        preview_rows.append(
            {
                "Dato": adjusted["date"],
                "Dag": adjusted["day"],
                "Starttid": adjusted["start_time"],
                "Sluttid": adjusted["end_time"],
                "Vagtform": adjusted["work_type"],
                "Person": adjusted["person"],
                "Summering efter reduktioner": adjusted["adjusted_duration_hours"],
                "Bemærkning": adjusted["comments"],
                "Justeringsnote": "; ".join(adjusted["adjustment_notes"]),
                "Summering før reduktioner": adjusted["original_duration_hours"],
                "Afstemning": adjusted.get("night_minus_status", ""),
            }
        )
    return preview_rows


def clear_download():
    st.session_state.pop("download_result", None)


def build_download(mode, source, html_value, upload_bytes, output_name,
                   institution_overrides, apply_night_reduction_to_bagvagt, pay_period_mode):
    """Build bytes from the actual exported file or the uploaded workbook."""
    name = output_name if output_name.lower().endswith(".xlsx") else f"{output_name}.xlsx"
    if source == "Excel" and mode != "Normal vagtkalender":
        if not upload_bytes:
            raise ValueError("Upload venligst en Excel-fil først.")
        workbook_bytes = upload_bytes
    else:
        if not html_value.strip():
            raise ValueError("Indsæt venligst HTML først.")
        parsed_rows = parse_shift_rows(html_value, institution_overrides=institution_overrides)
        actual_path = export_to_excel(
            parsed_rows, template_path="data/template.xlsx", output_path=Path("output") / name,
            apply_night_reduction_to_bagvagt=apply_night_reduction_to_bagvagt,
            pay_period_mode=pay_period_mode,
        )
        workbook_bytes = Path(actual_path).read_bytes()
    if mode != "Normal vagtkalender":
        workbook_bytes = reconcile_workbook(workbook_bytes, mode)
    return {"bytes": workbook_bytes, "name": name}


default_output_name = f"arbejdstimer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
if "output_name" not in st.session_state:
    st.session_state["output_name"] = default_output_name
if "html_input" not in st.session_state:
    st.session_state["html_input"] = ""

with st.container():
    col_left, col_right = st.columns([0.9, 2.1], gap="large")

    with col_left:
        st.markdown(
            """
            <div class="card">
                <div class="card-head">
                    <div>
                        <h2>Opsætning</h2>
                        <div class="req">Konfigurer navngivning, periode og natte-regler.</div>
                    </div>
                    <div class="status-pill">Klar til eksport</div>
                </div>
                <div class="hint-text">Brug denne side til at indlæse gemt kalender-HTML og eksportere til en ny Excel-fil uden gamle data.</div>
            </div>
            <div class="card">
                <div class="card-head">
                    <div>
                        <h2>Trin</h2>
                        <div class="req">Følg trinene for en hurtig arbejdsgang.</div>
                    </div>
                </div>
                <div class="step active">
                    <div class="step-num">1</div>
                    <div>
                        <div class="step-title">Indstillinger</div>
                        <div class="step-desc">Navn, periode og regler</div>
                    </div>
                </div>
                <div class="step">
                    <div class="step-num">2</div>
                    <div>
                        <div class="step-title">Indsæt HTML</div>
                        <div class="step-desc">Indsæt gemt kalender-HTML</div>
                    </div>
                </div>
                <div class="step">
                    <div class="step-num">3</div>
                    <div>
                        <div class="step-title">Eksportér</div>
                        <div class="step-desc">Download din Excel-fil</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_right:
        st.markdown('<div class="breadcrumbs">Dashboard <span style="color:#1c2b45;">›</span> Webscraper <span style="color:#1c2b45;">›</span> <span style="color:#9aa5c2;">Excel eksport</span></div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="hero">
                <div class="hero-left">
                    <h1 class="hero-title">Vagtplan til Excel</h1>
                    <p class="hero-subtitle">Konvertér gemt HTML fra kalenderen til en struktureret Excel-arbejdsmappe. Arbejd hurtigt med indstillinger, validering og eksport i ét samlet workflow.</p>
                </div>
                <div class="hero-actions">
                    <button class="btn btn-primary">Parse og eksportér</button>
                    <button class="btn btn-ghost">Ryd felter</button>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        button_col_1, button_col_2 = st.columns([1, 1], gap="small")
        with button_col_1:
            if st.button("Ryd felter", key="clear_fields"):
                clear_download()
                st.session_state["html_input"] = ""
                st.session_state["output_name"] = default_output_name
        with button_col_2:
            export_main = st.button("Parse og eksportér", key="top_export")

        content_col, overview_col = st.columns([2, 1], gap="large")
        with content_col:
            st.markdown('<div class="card"><div class="card-head"><h2>1. Indstillinger</h2><div class="req">Påkrævede felter er markeret</div></div></div>', unsafe_allow_html=True)
            st.markdown(
                """
                <div class="field-block">
                    <div class="field-label"><span class="field-title">Filnavn til Excel</span><span class="help-icon" title="Excel eksportfilnavn. .xlsx tilføjes automatisk, hvis ikke du skriver det selv.">?</span></div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            output_name = st.text_input(
                "Filnavn til Excel",
                value=st.session_state["output_name"],
                key="output_name",
                help="Excel eksportfilnavn. .xlsx tilføjes automatisk, hvis ikke du skriver det selv.",
            )
            reconciliation_mode = st.selectbox(
                "Vagtkalender", ["2:1", "1:1", "Normal vagtkalender"], index=2,
                key="reconciliation_mode",
            )
            input_source = "HTML"
            if reconciliation_mode != "Normal vagtkalender":
                input_source = st.radio("Datagrundlag", ["HTML", "Excel"], key="input_source")
            rows = []
            html_value = ""
            upload_bytes = None
            institution_overrides = {}
            apply_night_reduction_to_bagvagt = True
            pay_period_mode = "16-15"
            if input_source == "HTML":
                st.markdown(
                    """
                    <div class="field-block">
                        <div class="field-label"><span class="field-title">Bagvagt minus timer nat</span><span class="help-icon" title="Vælg Alpha-Afdeling for at undlade nat-2-timers reduktion for Bagvagt og skrive B i stedet for BN ved nattearbejde.">?</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                bagvagt_night_mode = st.selectbox(
                    "Bagvagt minus timer nat",
                    ["Andre Afdelinger", "Alpha-Afdeling"],
                    index=0,
                    help="Vælg Alpha-Afdeling for at undlade nat-2-timers reduktion for Bagvagt og skrive B i stedet for BN ved nattearbejde.",
                )
                apply_night_reduction_to_bagvagt = bagvagt_night_mode == "Andre Afdelinger"
                st.markdown(
                    """
                    <div class="field-block">
                        <div class="field-label"><span class="field-title">Vælg periode</span><span class="help-icon" title="Vælg om perioden skal være 16-15 eller hele kalender-måneden.">?</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                pay_period_mode = st.selectbox(
                    "Vælg periode",
                    ["16-15", "01 - Sidste i måneden"],
                    index=0,
                    help="Vælg om perioden skal være 16-15 eller hele kalender-måneden.",
                )
                st.caption("Hver eksport skriver en frisk fil, så gamle data bliver ikke medtaget.")

                st.markdown('<div class="card"><div class="card-head"><h2>Institution override mapping</h2></div></div>', unsafe_allow_html=True)
                st.markdown(
                    """
                    <div class="field-block">
                        <div class="field-label"><span class="field-title">Brugerdefinerede institutioner</span><span class="help-icon" title="Skriv institutionens rå navn og den ønskede normaliserede version på hver sin linje.">?</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.write("Brug `raw name => normalized name` på separate linjer.")
                institution_override_text = st.text_area(
                    "Brugerdefinerede institutioner",
                    value="Alpha.dk - Lyngbygade => Alpha Lyngbygade\nOpholdssteder - Solo => Solo",
                    height=140,
                    help="Skriv institutionens rå navn og den ønskede normaliserede version på hver sin linje.",
                )

                institution_overrides = {}
                for mapping_line in institution_override_text.splitlines():
                    mapping_line = mapping_line.strip()
                    if not mapping_line or mapping_line.startswith("#"):
                        continue
                    if "=>" in mapping_line:
                        raw, normalized = mapping_line.split("=>", 1)
                    elif ":" in mapping_line:
                        raw, normalized = mapping_line.split(":", 1)
                    elif "=" in mapping_line:
                        raw, normalized = mapping_line.split("=", 1)
                    else:
                        continue
                    institution_overrides[raw.strip()] = normalized.strip()

                st.markdown('<div class="card"><div class="card-head"><h2>2. Indsæt HTML fra kalenderen</h2><div class="req">Gem siden som HTML og indsæt indholdet her</div></div></div>', unsafe_allow_html=True)
                st.markdown(
                    """
                    <div class="field-block">
                        <div class="field-label"><span class="field-title">HTML indhold</span><span class="help-icon" title="Indsæt den gemte kalender-HTML, som importeres til Excel-eksporten.">?</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                html_value = st.text_area(
                    "Indsæt HTML her",
                    value=st.session_state["html_input"],
                    key="html_input",
                    height=380,
                    placeholder="Indsæt den gemte HTML fra kalender-siden her...",
                    help="Indsæt den gemte kalender-HTML, som importeres til Excel-eksporten.",
                )

                if html_value:
                    try:
                        rows = parse_shift_rows(html_value, institution_overrides=institution_overrides)
                        preview_rows = build_preview_rows(rows, apply_night_reduction_to_bagvagt)
                        st.markdown("<h2 style='margin-top:30px;'>Forhåndsvisning af eksport</h2>", unsafe_allow_html=True)
                        st.dataframe(preview_rows, use_container_width=True)
                    except Exception as exc:
                        st.error(f"Kunne ikke parse HTML: {exc}")
                else:
                    rows = []

            else:
                uploaded = st.file_uploader("Upload personfaner", type=["xlsx"], accept_multiple_files=False)
                upload_bytes = uploaded.getvalue() if uploaded is not None else None
                st.caption("Alle institutioner indgår i én opgørelse. Timerne læses uden yderligere reduktioner eller periodefiltrering.")

            input_signature = (
                reconciliation_mode, input_source, html_value,
                sha256(upload_bytes).hexdigest() if upload_bytes is not None else None,
                output_name, tuple(sorted(institution_overrides.items())),
                apply_night_reduction_to_bagvagt, pay_period_mode,
            )
            if st.session_state.get("download_input_signature") != input_signature:
                clear_download()
                st.session_state["download_input_signature"] = input_signature

            export_bottom = st.button("Eksportér", key="bottom_export")
            if export_main or export_bottom:
                clear_download()
                try:
                    st.session_state["download_result"] = build_download(
                        reconciliation_mode, input_source, html_value, upload_bytes, output_name,
                        institution_overrides, apply_night_reduction_to_bagvagt, pay_period_mode,
                    )
                    st.success("Excel-arbejdsmappe blev oprettet.")
                except Exception as exc:
                    st.error(str(exc))
            if "download_result" in st.session_state:
                result = st.session_state["download_result"]
                st.download_button(
                    label="Download Excel", data=result["bytes"], file_name=result["name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        with overview_col:
            if input_source == "Excel":
                st.subheader("Eksport-overblik")
                st.write(f"Model: {reconciliation_mode}")
                st.write("Excel-fil klar til afstemning" if upload_bytes else "Afventer Excel-fil")
                st.caption("Perioden bestemmes af første og sidste vagtdato i personfanerne.")
            else:
                st.markdown('<div class="card overview-card"><div class="card-head"><h2>Eksport-overblik</h2><p style="color:#9aa5c2;">Kontrollér dine valg før filen oprettes. Når HTML er indsat, vises antal vagter, timer og institutioner her.</p></div></div>', unsafe_allow_html=True)
                st.markdown(
                    """
                    <div class='checklist'>
                      <div class='check-item'><div class='check-mark'>✓</div><div>Filnavn er udfyldt og klar til ny Excel-arbejdsmappe.</div></div>
                      <div class='check-item'><div class='check-mark'>✓</div><div>Periode sat til <b>16-15</b>.</div></div>
                      <div class='check-item'><div class='check-mark'>✓</div><div>Nattrafdrag er aktivt for <b>Andre Afdelinger</b>.</div></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"""
                    <div class="card" style="padding-top:18px;">
                        <div class="summary-row"><div class="k">Forventet fil</div><div class="v">{output_name}</div></div>
                        <div class="summary-row"><div class="k">Vagter fundet</div><div class="v muted">{len(rows) if rows else 'Afventer HTML'}</div></div>
                        <div class="summary-row"><div class="k">Institutioner</div><div class="v muted">{len({r.get('institution', '') for r in rows if r.get('institution')}) if rows else 'Afventer HTML'}</div></div>
                        <div class="summary-row"><div class="k">Status</div><div class="v">{'Ingen fejl' if rows else 'Afventer HTML'}</div></div>
                    </div>
                    <div class="tip">Tip: Gem altid kalenderen som en ren HTML-fil, og brug preview før eksport.</div>
                    """,
                    unsafe_allow_html=True,
                )
