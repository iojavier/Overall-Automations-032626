from pathlib import Path
import runpy

import streamlit as st


APP_FILES = {
    "ExternalReport": "ExternalReport.py",
    "TexxenExternalV2": "TexxenExternalV2.py",
    "IntensityReport": "IntensityReport.py",
    "AgentsprodTexxen": "AgentsProdTexxen.py",
    "AgentsprodVolare": "AgentsProdVolare.py",
    "HourlyPerformance": "HourlyPerformance.py",
}

APP_META = {
    "ExternalReport": {
        "label": "External Report",
        "description": "Volare external debt collection summary and export workflow.",
    },
    "TexxenExternalV2": {
        "label": "Texxen External",
        "description": "Texxen summary, balances, and raw breakdown reporting.",
    },
    "IntensityReport": {
        "label": "Intensity Report",
        "description": "Intensity reporting workflow from the same launcher.",
    },
    "AgentsprodTexxen": {
        "label": "Agentsprod Texxen",
        "description": "Texxen agent production monitoring and reporting workflow.",
    },
    "AgentsprodVolare": {
        "label": "Agentsprod Volare",
        "description": "Volare agent production monitoring and reporting workflow.",
    },
    "HourlyPerformance": {
        "label": "Hourly Performance",
        "description": "Hourly performance monitoring and reporting workflow.",
    },
}


def _noop_set_page_config(*args, **kwargs):
    return None


def launch_selected_app(app_name: str) -> None:
    target = Path(__file__).with_name(APP_FILES[app_name])
    if not target.exists():
        st.error(f"Missing automation file: {target.name}")
        return

    original_set_page_config = st.set_page_config
    st.set_page_config = _noop_set_page_config
    try:
        runpy.run_path(str(target), run_name="__main__")
    finally:
        st.set_page_config = original_set_page_config


st.set_page_config(
    layout="wide",
    page_title="Automation Launcher",
    page_icon="A",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Manrope', sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(180, 139, 74, 0.18), transparent 28%),
            radial-gradient(circle at top right, rgba(76, 90, 120, 0.22), transparent 24%),
            linear-gradient(180deg, #0b1118 0%, #0f1722 45%, #121b27 100%);
        color: #edf2f7;
    }

    .block-container {
        max-width: 1060px;
        padding-top: 2.4rem;
        padding-bottom: 2rem;
    }

    .launch-hero {
        background:
            linear-gradient(135deg, rgba(18, 28, 40, 0.96) 0%, rgba(10, 16, 24, 0.94) 100%);
        border: 1px solid rgba(214, 183, 120, 0.16);
        border-radius: 26px;
        padding: 28px 30px 24px;
        box-shadow: 0 24px 80px rgba(0, 0, 0, 0.35);
        position: relative;
        overflow: hidden;
    }

    .launch-hero::after {
        content: "";
        position: absolute;
        inset: 0;
        background:
            linear-gradient(120deg, transparent 0%, rgba(214, 183, 120, 0.05) 48%, transparent 100%);
        pointer-events: none;
    }

    .launch-kicker {
        display: inline-block;
        padding: 6px 12px;
        border-radius: 999px;
        background: rgba(214, 183, 120, 0.12);
        border: 1px solid rgba(214, 183, 120, 0.16);
        color: #d6b778;
        font-size: 0.78rem;
        font-weight: 800;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }

    .launch-title {
        margin: 16px 0 8px;
        color: #f5f7fb;
        font-size: 2.6rem;
        line-height: 1.02;
        font-weight: 800;
    }

    .launch-copy {
        margin: 0;
        max-width: 700px;
        color: #aeb8c7;
        font-size: 1rem;
        line-height: 1.7;
    }

    .launch-card {
        margin-top: 20px;
        background: rgba(14, 22, 32, 0.88);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 20px;
        padding: 22px;
        box-shadow: 0 16px 48px rgba(0, 0, 0, 0.24);
    }

    .launch-label {
        color: #8ea0ba;
        font-size: 0.82rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 10px;
    }

    .launch-app {
        color: #f5f7fb;
        font-size: 1.35rem;
        font-weight: 800;
        margin-bottom: 6px;
    }

    .launch-app-copy {
        color: #aeb8c7;
        font-size: 0.96rem;
        line-height: 1.6;
        margin: 0;
    }

    div[data-testid="stSelectbox"] > label {
        font-weight: 700;
        color: #d7e0ec;
    }

    div[data-testid="stSelectbox"] > div[data-baseweb="select"] > div {
        min-height: 54px;
        border-radius: 14px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(14, 22, 32, 0.9);
        color: #edf2f7;
        box-shadow: none;
    }

    div[data-testid="stSelectbox"] svg {
        fill: #d6b778;
    }

    div[data-testid="stButton"] button {
        min-height: 54px;
        border-radius: 14px;
        border: none;
        background: linear-gradient(135deg, #d0a85f 0%, #9f7837 100%);
        color: white;
        font-weight: 800;
        box-shadow: 0 12px 28px rgba(159, 120, 55, 0.28);
    }

    div[data-testid="stButton"] button:hover {
        background: linear-gradient(135deg, #ddb56a 0%, #a68145 100%);
    }

    div[data-testid="stMarkdownContainer"] p {
        color: inherit;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "selected_automation" not in st.session_state:
    st.session_state.selected_automation = "ExternalReport"

st.markdown(
    """
    <div class="launch-hero">
        <div class="launch-kicker">Control Center</div>
        <div class="launch-title">Launch the right automation.</div>
        <p class="launch-copy">
            A single entrypoint for your reporting tools with a cleaner operating surface.
            Choose the workflow you need and run it without changing the original report files.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

selected_option = st.selectbox(
    "Select automation",
    list(APP_FILES.keys()),
    index=list(APP_FILES.keys()).index(st.session_state.selected_automation),
    format_func=lambda key: APP_META[key]["label"],
)

st.markdown(
    f"""
    <div class="launch-card">
        <div class="launch-label">Selected Automation</div>
        <div class="launch-app">{APP_META[selected_option]["label"]}</div>
        <p class="launch-app-copy">{APP_META[selected_option]["description"]}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.button("Run Automation", use_container_width=True):
    st.session_state.selected_automation = selected_option

launch_selected_app(st.session_state.selected_automation)
