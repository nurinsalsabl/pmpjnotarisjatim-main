"""
ui_styles.py
Modul tampilan (UI/UX) untuk Kuisioner PMPJ Notaris - Kanwil Kemenkumham Jawa Timur.
Berisi styling saja (CSS, header, sidebar stepper, progress bar).
Tidak ada logika perhitungan/integrasi di file ini.
"""

import streamlit as st

PRIMARY = "#0B3D6E"      # navy - warna institusi
PRIMARY_DARK = "#062845"
ACCENT = "#C9A227"       # emas - aksen
SUCCESS = "#1E7B45"
BG_SOFT = "#F4F6F9"
CARD_BG = "#FFFFFF"
TEXT_MUTED = "#5B6B7C"


def load_css():
    st.markdown(
        f"""
        <style>
        /* ===== Global ===== */
        .stApp {{
            background: {BG_SOFT};
        }}
        html, body, [class*="css"] {{
            font-family: "Segoe UI", "Source Sans Pro", sans-serif;
        }}
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}

        /* ===== Header banner ===== */
        .app-header {{
            background: linear-gradient(135deg, {PRIMARY} 0%, {PRIMARY_DARK} 100%);
            padding: 1.6rem 2rem;
            border-radius: 14px;
            margin-bottom: 1.4rem;
            box-shadow: 0 6px 18px rgba(6, 40, 69, 0.25);
        }}
        .app-header h1 {{
            color: #FFFFFF;
            font-size: 1.55rem;
            font-weight: 700;
            margin: 0;
            line-height: 1.3;
        }}
        .app-header p {{
            color: #DCE6F0;
            margin: 0.35rem 0 0 0;
            font-size: 0.92rem;
        }}
        .app-header .badge {{
            display: inline-block;
            background: {ACCENT};
            color: {PRIMARY_DARK};
            font-weight: 700;
            font-size: 0.72rem;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            padding: 0.18rem 0.6rem;
            border-radius: 999px;
            margin-bottom: 0.5rem;
        }}

        /* ===== Step card (form container) ===== */
        .step-card {{
            background: {CARD_BG};
            border: 1px solid #E4E9F0;
            border-radius: 14px;
            padding: 1.6rem 1.8rem;
            margin-bottom: 1.2rem;
            box-shadow: 0 2px 10px rgba(15, 34, 58, 0.05);
        }}
        .step-card h2 {{
            color: {PRIMARY_DARK};
            font-size: 1.2rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
            border-left: 5px solid {ACCENT};
            padding-left: 0.7rem;
        }}
        .step-subtext {{
            color: {TEXT_MUTED};
            font-size: 0.9rem;
            margin: 0.3rem 0 1.1rem 0.75rem;
        }}

        /* ===== Progress bar (top) ===== */
        .progress-wrap {{
            margin-bottom: 1.3rem;
        }}
        .progress-track {{
            width: 100%;
            height: 10px;
            background: #E4E9F0;
            border-radius: 999px;
            overflow: hidden;
        }}
        .progress-fill {{
            height: 100%;
            background: linear-gradient(90deg, {ACCENT} 0%, {PRIMARY} 100%);
            border-radius: 999px;
            transition: width 0.35s ease;
        }}
        .progress-label {{
            display: flex;
            justify-content: space-between;
            font-size: 0.8rem;
            color: {TEXT_MUTED};
            margin-top: 0.4rem;
        }}

        /* ===== Sidebar stepper ===== */
        section[data-testid="stSidebar"] {{
            background: {PRIMARY_DARK};
        }}
        section[data-testid="stSidebar"] * {{
            color: #EAF0F7 !important;
        }}
        .sidebar-title {{
            font-weight: 700;
            font-size: 0.95rem;
            margin-bottom: 0.9rem;
            padding-bottom: 0.6rem;
            border-bottom: 1px solid rgba(255,255,255,0.15);
        }}
        .step-item {{
            display: flex;
            align-items: center;
            gap: 0.55rem;
            padding: 0.42rem 0.5rem;
            border-radius: 8px;
            font-size: 0.84rem;
            margin-bottom: 0.15rem;
        }}
        .step-item.active {{
            background: rgba(201, 162, 39, 0.18);
            font-weight: 700;
            color: {ACCENT} !important;
        }}
        .step-item.done .step-dot {{
            background: {SUCCESS};
            border-color: {SUCCESS};
        }}
        .step-dot {{
            flex: 0 0 20px;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            border: 2px solid rgba(255,255,255,0.4);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.68rem;
            font-weight: 700;
        }}

        /* ===== Buttons ===== */
        .stButton > button {{
            border-radius: 10px;
            font-weight: 600;
            padding: 0.55rem 1rem;
        }}
        .stButton > button[kind="primary"] {{
            background: {PRIMARY};
            border: none;
        }}
        .stButton > button[kind="primary"]:hover {{
            background: {PRIMARY_DARK};
        }}

        /* ===== Metric cards on summary step ===== */
        div[data-testid="stMetric"] {{
            background: {CARD_BG};
            border: 1px solid #E4E9F0;
            border-radius: 12px;
            padding: 0.8rem 1rem;
        }}

        /* ===== Report-style summary (Step 9) ===== */
        .report-block {{
            font-family: Arial, Helvetica, sans-serif;
            margin-bottom: 1.4rem;
        }}
        .report-title {{
            font-family: Arial, Helvetica, sans-serif;
            font-size: 1rem;
            font-weight: 700;
            color: {PRIMARY_DARK};
            margin-bottom: 0.5rem;
            padding-bottom: 0.3rem;
            border-bottom: 2px solid {ACCENT};
        }}
        .report-table {{
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, Helvetica, sans-serif;
            font-size: 0.9rem;
        }}
        .report-table tr {{
            border-bottom: 1px solid #E4E9F0;
        }}
        .report-table tr:last-child {{
            border-bottom: none;
        }}
        .report-table td {{
            padding: 0.45rem 0.6rem;
            vertical-align: top;
        }}
        .report-table td.label {{
            width: 40%;
            color: {TEXT_MUTED};
            font-weight: 600;
        }}
        .report-table td.value {{
            color: {PRIMARY_DARK};
            font-weight: 500;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_report_table(rows):
    """Render list of (label, value) tuples as an Arial-styled report table.

    rows: list of (label:str, value:str)
    """
    html = '<table class="report-table">'
    for label, value in rows:
        html += (
            f'<tr><td class="label">{label}</td>'
            f'<td class="value">{value if value not in (None, "") else "-"}</td></tr>'
        )
    html += "</table>"
    return html


def render_header():
    st.markdown(
        """
        <div class="app-header">
            <span class="badge">Kanwil Kemenkumham Jawa Timur</span>
            <h1>📊 Kuisioner PMPJ Notaris</h1>
            <p>Penilaian Risiko Prinsip Mengenali Pengguna Jasa (PMPJ) bagi Kantor Notaris</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar_stepper(step_titles, current_step):
    with st.sidebar:
        st.markdown('<div class="sidebar-title">🧭 Tahapan Pengisian</div>', unsafe_allow_html=True)
        html = ""
        for i, title in enumerate(step_titles, start=1):
            if i < current_step:
                cls = "step-item done"
                dot = "✓"
            elif i == current_step:
                cls = "step-item active"
                dot = str(i)
            else:
                cls = "step-item"
                dot = str(i)
            html += (
                f'<div class="{cls}">'
                f'<div class="step-dot">{dot}</div>'
                f'<div>{title}</div>'
                f'</div>'
            )
        st.markdown(html, unsafe_allow_html=True)


def render_progress_bar(step_titles, current_step):
    total = len(step_titles)
    pct = int((current_step - 1) / (total - 1) * 100) if total > 1 else 100
    st.markdown(
        f"""
        <div class="progress-wrap">
            <div class="progress-track">
                <div class="progress-fill" style="width:{pct}%;"></div>
            </div>
            <div class="progress-label">
                <span>Langkah {current_step} dari {total}</span>
                <span>{step_titles[current_step - 1]}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )