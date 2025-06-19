import pandas as pd
import plotly.express as px
import streamlit as st
import datetime
import os
import io
import requests
from streamlit_lottie import st_lottie
from plotly.io import to_image
import plotly.io as pio
pio.kaleido.scope.default_format = "png"

# Load animation from Lottie URL
def load_lottieurl(url):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

# Use session_state to preserve language selection
if "language" not in st.session_state:
    st.session_state.language = "English"

# Layout: Language selector + Title
col1, col2 = st.columns([3, 1])
with col2:
    selected_lang = st.radio("🌐", ["English", "Български"], index=["English", "Български"].index(st.session_state.language), label_visibility="collapsed")
    if selected_lang != st.session_state.language:
        st.session_state.language = selected_lang
        st.rerun()

# Translations
titles = {
    "English": "📊 Business Trips Statistics",
    "Български": "📊 Статистика за командировки"
}

labels = {
    "English": {
        "file_error": "File 'exported BTs.csv' not found.",
        "missing_columns": "Missing columns:",
        "stat_select": "Select a statistic to visualize:",
        "option1": "Business Trips per employee",
        "option2": "Business Trips by month (last year)",
        "option3": "Domestic vs Abroad statistics",
        "chart1": "Business Trips per employee",
        "chart2": "Business Trips by month (last year)",
        "chart3": "Domestic vs Abroad",
        "export": "Export chart data to CSV"
    },
    "Български": {
        "file_error": "Файлът 'exported BTs.csv' не е намерен.",
        "missing_columns": "Липсващи колони:",
        "stat_select": "Изберете каква статистика да визуализирате:",
        "option1": "Командировки по служители",
        "option2": "Командировки по месеци (последна година)",
        "option3": "Статистика по дестинация",
        "chart1": "Командировки по служители",
        "chart2": "Командировки по месеци (последна година)",
        "chart3": "Дестинации - в/извън страната",
        "export": "Експортирай диаграмата в CSV"
    }
}
L = labels[st.session_state.language]

with col1:
    st.markdown(f"""<h1 style='font-size: 34px;'>{titles[st.session_state.language]}</h1>
        </div>
    """, unsafe_allow_html=True)

# Intro animation
lottie_chart = load_lottieurl("https://lottie.host/1f6f5ae8-d54d-4a04-b54a-3d3f8e60c5e3/xkqsPAFq3H.json")
if lottie_chart:
    st_lottie(lottie_chart, height=200, key="intro")

# Choose separator based on language
csv_separator = ";" if st.session_state.language == "Български" else ","

# Load the actual CSV file
file_path = "exported BTs.csv"

if not os.path.exists(file_path):
    st.error(L["file_error"])
else:
    raw_df = pd.read_csv(file_path, encoding='utf-8', sep=';')
    raw_df.columns = raw_df.columns.str.strip()

    required_columns = ["Requestor Employee", "Type of Destination", "Start Date"]
    missing_columns = [col for col in required_columns if col not in raw_df.columns]

    if missing_columns:
        st.error(f"{L['missing_columns']} {missing_columns}")
    else:
        try:
            raw_df["Type of Destination"] = raw_df["Type of Destination"].str.extract(r'\d+\.\s*(.*)')
            raw_df["Start Date"] = pd.to_datetime(raw_df["Start Date"].str.replace(" г.", ""), format="%d.%m.%Y")

            df = raw_df[["Requestor Employee", "Type of Destination", "Start Date"]].copy()

            st.markdown("---")
            st.subheader("🔽 " + L["stat_select"])
            option = st.selectbox("", (L["option1"], L["option2"], L["option3"]))

            export_df = None
            chart_title = ""
            fig = None
            chart_key = ""

            if option == L["option1"]:
                counts = df["Requestor Employee"].value_counts().reset_index()
                counts.columns = ["Employee", "Business Trips Count"]
                export_df = counts
                chart_title = L["chart1"]
                fig = px.pie(counts, names="Employee", values="Business Trips Count", title=chart_title,
                             color_discrete_sequence=px.colors.qualitative.Set3,
                             hover_data=["Business Trips Count"])
                chart_key = "employee_chart"

            elif option == L["option2"]:
                df_last_year = df[df["Start Date"] >= (datetime.datetime.now() - pd.DateOffset(years=1))].copy()
                df_last_year["Month"] = df_last_year["Start Date"].dt.strftime("%B")
                counts = df_last_year["Month"].value_counts().reindex(
                    ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
                ).fillna(0)
                counts = counts[counts > 0]
                export_df = counts.reset_index()
                export_df.columns = ["Month", "Business Trips Count"]
                chart_title = L["chart2"]
                fig = px.pie(names=counts.index, values=counts.values, title=chart_title,
                             color_discrete_sequence=px.colors.qualitative.Set2,
                             hover_name=counts.index)
                chart_key = "month_chart"

            elif option == L["option3"]:
                counts = df["Type of Destination"].value_counts().reset_index()
                counts.columns = ["Type", "Business Trips Count"]
                export_df = counts
                chart_title = L["chart3"]
                fig = px.pie(counts, names="Type", values="Business Trips Count", title=chart_title,
                             color_discrete_sequence=px.colors.qualitative.Pastel,
                             hover_data=["Business Trips Count"])
                chart_key = "type_chart"

            if fig is not None:
                st.markdown("---")
                st.subheader("📈 " + chart_title)
                st.plotly_chart(fig, key=chart_key, use_container_width=True)

            if export_df is not None:
                csv = export_df.to_csv(index=False, sep=csv_separator, lineterminator='\n').encode('utf-8')
                with st.expander("📤 Export"):
                    st.download_button(L["export"], data=csv, file_name="exported_data.csv", mime="text/csv")

        except Exception as e:
            st.error(f"Data processing error: {e}")

# Footer
st.markdown("---")
st.markdown("© 2025 • Author: Seyhan Muskova • [GitHub Repo](https://github.com/yourrepo)", unsafe_allow_html=True)
