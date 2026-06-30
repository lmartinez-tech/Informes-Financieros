from __future__ import annotations

from html import escape
from textwrap import dedent

import streamlit as st


def inject_app_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --ink: #172033;
            --muted: #667085;
            --line: #E4E7EC;
            --surface: #FFFFFF;
            --canvas: #F6F7F9;
            --brand: #0B6B57;
            --brand-dark: #075446;
            --accent: #C68A2D;
            --soft: #EAF4F1;
        }
        html, body, [class*="css"] {
            font-family: "Segoe UI", Inter, Arial, sans-serif;
            color: var(--ink);
        }
        .stApp { background: var(--canvas); }
        .main .block-container {
            max-width: 1440px;
            padding: 1.4rem 2.2rem 3rem;
        }
        [data-testid="stSidebar"] {
            background: #101828;
            border-right: 1px solid #1D2939;
        }
        [data-testid="stSidebar"] * { color: #F2F4F7; }
        [data-testid="stSidebar"] [role="radiogroup"] label {
            padding: .62rem .72rem;
            border-radius: 6px;
            margin-bottom: .18rem;
        }
        [data-testid="stSidebar"] [role="radiogroup"] label:hover {
            background: #1D2939;
        }
        [data-testid="stSidebar"] hr { border-color: #344054; }
        .js-brand {
            display: flex;
            align-items: center;
            gap: .78rem;
            padding: .4rem 0 1.25rem;
        }
        .js-brand-mark {
            width: 42px;
            height: 42px;
            display: grid;
            place-items: center;
            background: #D6A44B;
            color: #101828;
            font-weight: 800;
            font-size: .9rem;
            border-radius: 6px;
        }
        .js-brand-name { font-size: 1rem; font-weight: 700; color: #FFFFFF; }
        .js-brand-role { font-size: .74rem; color: #98A2B3; margin-top: .1rem; }
        .js-topline {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--line);
            padding: 0 0 .75rem;
            margin-bottom: 1.35rem;
        }
        .js-topline-label {
            color: var(--muted);
            font-size: .78rem;
            font-weight: 650;
            text-transform: uppercase;
        }
        .js-status {
            display: inline-flex;
            align-items: center;
            gap: .45rem;
            color: #344054;
            font-size: .78rem;
            font-weight: 600;
        }
        .js-status:before {
            content: "";
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #12B76A;
        }
        .js-page-header {
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            gap: 1.5rem;
            align-items: end;
            margin-bottom: 1.45rem;
        }
        .js-eyebrow {
            color: var(--brand);
            font-weight: 750;
            font-size: .76rem;
            text-transform: uppercase;
            margin-bottom: .45rem;
        }
        .js-page-title {
            font-size: 2rem;
            line-height: 1.12;
            font-weight: 760;
            color: var(--ink);
            margin: 0;
        }
        .js-page-subtitle {
            color: var(--muted);
            font-size: .96rem;
            line-height: 1.55;
            max-width: 880px;
            margin-top: .55rem;
        }
        .js-header-badge {
            border: 1px solid #B7D8CF;
            background: var(--soft);
            color: var(--brand-dark);
            border-radius: 6px;
            padding: .52rem .72rem;
            font-size: .78rem;
            font-weight: 700;
            white-space: nowrap;
        }
        .js-feature-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
            background: var(--surface);
        }
        .js-feature-card {
            min-height: 168px;
            padding: 1.2rem;
            border-right: 1px solid var(--line);
        }
        .js-feature-card:nth-child(3n) { border-right: 0; }
        .js-feature-card:nth-child(n+4) { border-top: 1px solid var(--line); }
        .js-feature-index {
            color: var(--accent);
            font-size: .72rem;
            font-weight: 800;
            margin-bottom: 1.35rem;
        }
        .js-feature-title { font-weight: 730; color: var(--ink); margin-bottom: .45rem; }
        .js-feature-copy { color: var(--muted); font-size: .88rem; line-height: 1.5; }
        .js-panel {
            border: 1px solid var(--line);
            background: var(--surface);
            border-radius: 8px;
            padding: 1.1rem 1.2rem;
            margin-bottom: 1rem;
        }
        .js-panel-title { font-weight: 720; color: var(--ink); margin-bottom: .28rem; }
        .js-panel-copy { color: var(--muted); font-size: .86rem; line-height: 1.45; }
        .js-steps {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            border-top: 1px solid var(--line);
            border-bottom: 1px solid var(--line);
            margin: .2rem 0 1.3rem;
        }
        .js-step { padding: .8rem .4rem; color: var(--muted); font-size: .8rem; }
        .js-step strong { color: var(--brand); margin-right: .35rem; }
        div[data-testid="stFileUploader"] {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: .45rem .8rem;
        }
        div[data-testid="stMetric"] {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: .85rem 1rem;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            overflow: hidden;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: .25rem;
            border-bottom: 1px solid var(--line);
        }
        .stTabs [data-baseweb="tab"] {
            height: 42px;
            border-radius: 5px 5px 0 0;
            color: var(--muted);
            font-weight: 620;
        }
        .stTabs [aria-selected="true"] { color: var(--brand) !important; }
        .stButton > button, .stDownloadButton > button {
            border-radius: 6px;
            min-height: 42px;
            font-weight: 680;
            border: 1px solid #CBD5E1;
        }
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
            background: var(--brand);
            border-color: var(--brand);
        }
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover {
            background: var(--brand-dark);
            border-color: var(--brand-dark);
        }
        h1, h2, h3 { letter-spacing: 0; color: var(--ink); }
        @media (max-width: 900px) {
            .main .block-container { padding: 1rem; }
            .js-page-header { grid-template-columns: 1fr; }
            .js-header-badge { width: fit-content; }
            .js-feature-grid, .js-steps { grid-template-columns: 1fr; }
            .js-feature-card { border-right: 0; border-top: 1px solid var(--line); }
            .js-feature-card:first-child { border-top: 0; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def sidebar_brand() -> None:
    st.sidebar.markdown(
        dedent(
            """
            <div class="js-brand">
                <div class="js-brand-mark">JS</div>
                <div>
                    <div class="js-brand-name">Julio Salazar</div>
                    <div class="js-brand-role">Gestión financiera y tributaria</div>
                </div>
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )


def top_bar(section: str) -> None:
    st.markdown(
        dedent(
            f"""
            <div class="js-topline">
                <div class="js-topline-label">{escape(section)}</div>
                <div class="js-status">Plataforma operativa</div>
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str, eyebrow: str = "Centro de análisis", badge: str = "") -> None:
    badge_html = f'<div class="js-header-badge">{escape(badge)}</div>' if badge else ""
    st.markdown(
        dedent(
            f"""
            <div class="js-page-header">
                <div>
                    <div class="js-eyebrow">{escape(eyebrow)}</div>
                    <h1 class="js-page-title">{escape(title)}</h1>
                    <div class="js-page-subtitle">{escape(subtitle)}</div>
                </div>
                {badge_html}
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )


def feature_grid(items: list[tuple[str, str]]) -> None:
    cards = []
    for index, (title, copy) in enumerate(items, start=1):
        cards.append(
            dedent(
                f"""
                <div class="js-feature-card">
                    <div class="js-feature-index">{index:02d}</div>
                    <div class="js-feature-title">{escape(title)}</div>
                    <div class="js-feature-copy">{escape(copy)}</div>
                </div>
                """
            ).strip()
        )
    st.markdown(f'<div class="js-feature-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def process_steps(items: list[str]) -> None:
    steps = "".join(
        f'<div class="js-step"><strong>{index:02d}</strong> {escape(item)}</div>'
        for index, item in enumerate(items, start=1)
    )
    st.markdown(f'<div class="js-steps">{steps}</div>', unsafe_allow_html=True)


def info_panel(title: str, copy: str) -> None:
    st.markdown(
        dedent(
            f"""
            <div class="js-panel">
                <div class="js-panel-title">{escape(title)}</div>
                <div class="js-panel-copy">{escape(copy)}</div>
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )
