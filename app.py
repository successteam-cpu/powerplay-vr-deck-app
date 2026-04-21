"""
Powerplay VR Deck Generator — Streamlit Web App
================================================
Run locally:  streamlit run app.py
Deploy:       push to GitHub → Streamlit Community Cloud (free)

This is a thin UI layer over vr_deck_generator.py — all deck logic lives there.
"""
import os
import io
import json
import shutil
import tempfile
import zipfile
import datetime as dt
from pathlib import Path

import streamlit as st
import pandas as pd

import vr_deck_generator as vg


# =============================================================================
# PAGE CONFIG + GLOBAL STYLING
# =============================================================================
st.set_page_config(
    page_title="Powerplay VR Deck Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* Tighten default padding so the app feels less cavernous */
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1100px; }
/* Bigger primary button */
.stButton > button[kind="primary"] {
    background-color: #1E2761;
    color: white;
    border: none;
    padding: 0.6rem 1.2rem;
    font-size: 1.05rem;
    font-weight: 600;
}
.stButton > button[kind="primary"]:hover {
    background-color: #151C48;
    color: white;
}
/* Amber accent on download buttons */
.stDownloadButton > button {
    background-color: #F59E0B;
    color: #151C48;
    border: none;
    font-weight: 600;
}
.stDownloadButton > button:hover {
    background-color: #D97706;
    color: white;
}
/* Tame the huge default headings */
h1 { font-size: 2rem !important; margin-bottom: 0.3rem !important; }
h2 { font-size: 1.4rem !important; margin-top: 1.5rem !important; }
h3 { font-size: 1.15rem !important; }
/* Subtle info boxes */
.info-box {
    background: #F8FAFC;
    border-left: 3px solid #F59E0B;
    padding: 0.8rem 1rem;
    border-radius: 4px;
    margin: 0.5rem 0;
    font-size: 0.92rem;
    color: #475569;
}
</style>
""", unsafe_allow_html=True)


# =============================================================================
# HEADER
# =============================================================================
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown("# 📊 Powerplay VR Deck Generator")
    st.caption("Upload a usage CSV → fill in the story → download a renewal-ready deck.")
with col2:
    st.markdown(
        f"<div style='text-align:right; color:#64748B; padding-top:1.5rem; font-size:0.85rem;'>"
        f"Powerplay CS · v1.0<br/>{dt.date.today().strftime('%B %Y')}"
        f"</div>",
        unsafe_allow_html=True,
    )

st.divider()


# =============================================================================
# SESSION STATE
# =============================================================================
if 'csv_df' not in st.session_state:
    st.session_state.csv_df = None
if 'csv_name' not in st.session_state:
    st.session_state.csv_name = None
if 'orgs' not in st.session_state:
    st.session_state.orgs = []
if 'generated' not in st.session_state:
    st.session_state.generated = None


# =============================================================================
# STEP 1 — UPLOAD CSV
# =============================================================================
st.markdown("## Step 1 · Upload your usage CSV")
st.markdown(
    '<div class="info-box">Use the <strong>AA Event Depth Funnel — Week Wise</strong> '
    'export. The file needs these columns: <code>org_name, project_name, user_name, '
    'role, module, event_count, activity_date</code>. Works for one org or many.</div>',
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("Drop your CSV here", type=['csv'], label_visibility='collapsed')

if uploaded is not None:
    # Cache parse so we don't re-read on every rerun
    if st.session_state.csv_name != uploaded.name:
        try:
            df = pd.read_csv(uploaded, low_memory=False)
            required = {'org_name', 'project_name', 'user_name', 'role',
                        'module', 'event_count', 'activity_date'}
            missing = required - set(df.columns)
            if missing:
                st.error(f"CSV is missing required columns: {sorted(missing)}")
                st.stop()
            st.session_state.csv_df = df
            st.session_state.csv_name = uploaded.name
            st.session_state.orgs = sorted(df['org_name'].dropna().unique().tolist())
            st.session_state.generated = None  # reset any old output
        except Exception as e:
            st.error(f"Could not read CSV: {e}")
            st.stop()

    df = st.session_state.csv_df
    orgs = st.session_state.orgs
    c1, c2, c3 = st.columns(3)
    c1.metric("Rows in CSV", f"{len(df):,}")
    c2.metric("Orgs found", len(orgs))
    c3.metric("Date range", f"{df['activity_date'].min()[:10]} → {df['activity_date'].max()[:10]}"
              if len(df) else "—")


# =============================================================================
# STEP 2 — PICK THE ORG
# =============================================================================
if st.session_state.csv_df is not None:
    st.markdown("## Step 2 · Which org is this deck for?")

    orgs = st.session_state.orgs
    ALL_OPT = "[All orgs — generate one deck per org]"
    picker_options = orgs + ([ALL_OPT] if len(orgs) > 1 else [])
    default_idx = 0

    picked = st.selectbox(
        "Org",
        options=picker_options,
        index=default_idx,
        label_visibility='collapsed',
    )

    if picked == ALL_OPT:
        selected_orgs = orgs
        st.info(f"Batch mode: will generate **{len(orgs)} decks** (one per org). "
                f"The customization form below will apply to the first org; "
                f"use the power-user JSON editor for per-org detail.")
    else:
        selected_orgs = [picked]

    # =============================================================================
    # STEP 3 — CUSTOMIZATION FORM
    # =============================================================================
    st.markdown("## Step 3 · Tell us the story (optional but recommended)")
    st.markdown(
        '<div class="info-box">Every field is optional. Leave things blank if you don\'t '
        'know them — the deck will just skip the related slide. The more you fill in, '
        'the more personalized the deck feels to the client.</div>',
        unsafe_allow_html=True,
    )

    form_target = selected_orgs[0]  # form drives the first (or only) org
    st.caption(f"Customizing for: **{form_target}**")

    # Basic
    with st.expander("👤 Who the deck is for", expanded=True):
        owner_name = st.text_input(
            "Owner / decision-maker name",
            placeholder="e.g. Mitranjan Ganguly",
            help="Appears on the cover slide and on the 'Our Ask' slide.",
        )
        c1, c2 = st.columns(2)
        with c1:
            owner_title = st.text_input("Owner title", value="Owner",
                                        placeholder="e.g. Owner / MD / Director")
        with c2:
            company_display = st.text_input(
                "Company name (for the cover)",
                placeholder=form_target.title() if form_target.isupper() else form_target,
                help="Leave blank to auto-format the org name from the CSV.",
            )
        c1, c2 = st.columns(2)
        with c1:
            business_desc = st.text_input(
                "Business description",
                placeholder="e.g. Water & wastewater infra — STP / WTP / UGR",
            )
        with c2:
            region = st.text_input("Region", placeholder="e.g. West Bengal")

    # Goal slide
    with st.expander("🎯 Onboarding goal (enables the 'Goal Delivered' slide)", expanded=True):
        onboarding_goal = st.text_area(
            "What they asked for at onboarding (full quote)",
            placeholder="e.g. We want to track labour attendance across our sites.",
            height=80,
            help="Shown as a quote on slide 4. Keep it under ~80 characters.",
        )
        onboarding_goal_short = st.text_input(
            "Short phrase version",
            placeholder="e.g. labour attendance tracking",
            help="Used in running copy on slides 2 and 14 ('...solved labour attendance tracking').",
        )

    # Pain points
    with st.expander("⚠️ Pain points (enables 'What you've flagged' slide)"):
        pain_points_raw = st.text_area(
            "One pain point per line",
            placeholder="Material leakage across sites hard to trace\n"
                        "5 dormant sites with no visibility\n"
                        "Supervisors not updating task status consistently",
            height=100,
            help="Each line becomes a row on slide 10. Keep each under 50 chars.",
        )

    # Year 1 comparison
    with st.expander("📈 Year 1 comparison data (enables Y1 vs Y2 slide)"):
        st.caption("If you have Year 1 numbers handy, enter them here. All three are optional.")
        c1, c2, c3 = st.columns(3)
        with c1:
            y1_events = st.number_input("Y1 total activities", min_value=0, value=0, step=100)
        with c2:
            y1_users = st.number_input("Y1 trained users", min_value=0, value=0, step=1)
        with c3:
            y1_sites = st.number_input("Y1 active sites", min_value=0, value=0, step=1)

    # Renewal ask
    with st.expander("💰 Renewal ask", expanded=True):
        ask_type = st.radio(
            "What are we pitching?",
            options=['Same ARR renewal', 'Upsell', 'Multi-year lock-in', 'Custom'],
            horizontal=True,
        )
        upsell_amount = None
        renewal_custom = None
        if ask_type == 'Upsell':
            upsell_amount = st.number_input(
                "Upsell target (annual ₹)", min_value=0, value=0, step=10000
            )
        elif ask_type == 'Custom':
            renewal_custom = st.text_input(
                "Custom ask line",
                placeholder="e.g. Renew + add Labour Payroll module at ₹200,000",
            )

    # KAM / Success
    with st.expander("👥 Powerplay team"):
        c1, c2 = st.columns(2)
        with c1:
            kam_name = st.text_input("KAM name", placeholder="e.g. Altaf")
            kam_email = st.text_input("KAM email", placeholder="e.g. altaf@getpowerplay.in")
        with c2:
            trainer_name = st.text_input("Success partner name", placeholder="e.g. Wasim")
            trainer_email = st.text_input("Success partner email",
                                          placeholder="e.g. wasim@getpowerplay.in")

    # Build customizations dict
    pain_points_list = [p.strip() for p in (pain_points_raw or "").splitlines() if p.strip()]
    y1 = {}
    if y1_events: y1['events'] = int(y1_events)
    if y1_users:  y1['users']  = int(y1_users)
    if y1_sites:  y1['sites']  = int(y1_sites)

    ask_map = {
        'Same ARR renewal': 'same_arr',
        'Upsell':           'upsell',
        'Multi-year lock-in': 'multi_year',
        'Custom':           'custom',
    }

    form_customs = {
        'owner_name':            owner_name or None,
        'owner_title':           owner_title or None,
        'company_display_name':  company_display or None,
        'business_description':  business_desc or None,
        'region':                region or None,
        'onboarding_goal':       onboarding_goal or None,
        'onboarding_goal_short': onboarding_goal_short or None,
        'pain_points':           pain_points_list or None,
        'year1_comparison':      y1 or None,
        'renewal_ask':           ask_map[ask_type],
        'upsell_amount':         upsell_amount,
        'renewal_custom_text':   renewal_custom,
        'kam_name':              kam_name or None,
        'kam_email':             kam_email or None,
        'trainer_name':          trainer_name or None,
        'trainer_email':         trainer_email or None,
    }
    # Drop None values so defaults in the engine take over
    form_customs = {k: v for k, v in form_customs.items() if v is not None}

    customizations = {form_target: form_customs}

    # Advanced JSON override (for batch mode or power users)
    if len(selected_orgs) > 1:
        with st.expander("🛠️ Power-user: per-org customizations (JSON)"):
            st.caption("For batch mode, paste a JSON object of "
                       "`{\"ORG NAME\": {...fields...}}` to override any org. "
                       "Leave blank to use the form above for the first org only.")
            json_override = st.text_area(
                "Customizations JSON",
                height=200,
                placeholder='{\n  "ACME CONSTRUCTION LTD": {\n'
                            '    "owner_name": "Alice",\n'
                            '    "onboarding_goal": "We need faster billing."\n'
                            '  }\n}',
            )
            if json_override.strip():
                try:
                    extra = json.loads(json_override)
                    customizations.update(extra)
                    st.success(f"Loaded overrides for {len(extra)} org(s).")
                except Exception as e:
                    st.error(f"Invalid JSON: {e}")

    # =============================================================================
    # STEP 4 — GENERATE
    # =============================================================================
    st.markdown("## Step 4 · Generate the deck")

    if st.button("🚀 Generate deck", type="primary", use_container_width=True):
        with st.spinner(f"Building deck{'s' if len(selected_orgs) > 1 else ''}..."):
            tmpdir = tempfile.mkdtemp(prefix="vrdeck_")
            # Save uploaded CSV to disk so generator can re-read it
            csv_tmp = os.path.join(tmpdir, "usage.csv")
            st.session_state.csv_df.to_csv(csv_tmp, index=False)

            try:
                results = vg.generate_vr_deck(
                    csv_path       = csv_tmp,
                    org_filter     = selected_orgs if len(selected_orgs) > 1 else selected_orgs[0],
                    output_dir     = tmpdir,
                    customizations = customizations,
                    verbose        = False,
                )
                st.session_state.generated = {
                    'tmpdir': tmpdir,
                    'results': results,
                }
                st.success(f"Generated {len(results)} deck(s).")
            except Exception as e:
                st.error(f"Generation failed: {e}")
                shutil.rmtree(tmpdir, ignore_errors=True)
                st.stop()


# =============================================================================
# STEP 5 — DOWNLOAD
# =============================================================================
if st.session_state.generated:
    st.markdown("## Step 5 · Download")
    results = st.session_state.generated['results']
    tmpdir = st.session_state.generated['tmpdir']

    if len(results) == 1:
        # Single deck — direct download
        r = results[0]
        st.success(f"Deck ready for **{r['org']}** ({r['slides']} slides)")
        with open(r['path'], 'rb') as f:
            st.download_button(
                label=f"⬇️ Download {os.path.basename(r['path'])}",
                data=f.read(),
                file_name=os.path.basename(r['path']),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
    else:
        # Multiple decks — zip
        st.success(f"{len(results)} decks ready")
        zip_path = os.path.join(tmpdir, f"VR_decks_{dt.date.today().strftime('%Y%m%d')}.zip")
        if not os.path.exists(zip_path):
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
                for r in results:
                    z.write(r['path'], arcname=os.path.basename(r['path']))

        with open(zip_path, 'rb') as f:
            st.download_button(
                label=f"⬇️ Download all {len(results)} decks (ZIP)",
                data=f.read(),
                file_name=os.path.basename(zip_path),
                mime="application/zip",
                use_container_width=True,
            )

        with st.expander(f"Or download individually ({len(results)} files)"):
            for r in results:
                with open(r['path'], 'rb') as f:
                    st.download_button(
                        label=f"{r['org']} ({r['slides']} slides)",
                        data=f.read(),
                        file_name=os.path.basename(r['path']),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"dl_{r['org']}",
                    )


# =============================================================================
# FOOTER
# =============================================================================
st.divider()
st.markdown(
    "<div style='text-align:center; color:#94A3B8; font-size:0.85rem;'>"
    "Built by Powerplay CS · Questions? Post in #cs-team-help"
    "</div>",
    unsafe_allow_html=True,
)
