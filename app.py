#!/usr/bin/env python3
"""
CB Reporting Hub — Streamlit App
Combines Weekly WoW Report, Monthly Campaign Summary, and Keyword Analysis
into a single app with a report selector.

Usage: streamlit run app.py
"""

import streamlit as st
import wow_report_app
import monthly_app
import keyword_analysis_app


def main():
    st.set_page_config(page_title="CB Reporting Hub", layout="wide")

    st.title("CB Reporting Hub")

    report_type = st.selectbox(
        "Select a report to run:",
        ["Weekly WoW Report", "Monthly Campaign Summary", "Keyword Analysis"]
    )

    st.divider()

    if report_type == "Weekly WoW Report":
        wow_report_app.main()
    elif report_type == "Monthly Campaign Summary":
        monthly_app.main()
    elif report_type == "Keyword Analysis":
        keyword_analysis_app.run()


if __name__ == "__main__":
    main()
