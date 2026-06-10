import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import re

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Аналитическая система КЦ",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# DESIGN
# =========================================================

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #F6FCFD 0%, #FFFFFF 45%, #E4F5F8 100%);
    color: #1F2D3D;
}
h1 {
    color: #0097A9 !important;
    font-size: 42px !important;
    font-weight: 900 !important;
}
h2, h3 {
    color: #005B8F !important;
    font-weight: 800 !important;
}
.hero {
    background: linear-gradient(135deg, #FFFFFF 0%, #E1F5F8 55%, #B9E5EB 100%);
    padding: 34px;
    border-radius: 30px;
    margin-bottom: 30px;
    box-shadow: 0 18px 45px rgba(0, 151, 169, 0.16);
    border: 1px solid #D5F0F4;
}
.hero-subtitle {
    color: #005B8F;
    font-size: 22px;
    font-weight: 700;
}
.hero-text {
    color: #34495E;
    font-size: 17px;
    line-height: 1.6;
}
.section-card {
    background: white;
    padding: 22px;
    border-radius: 24px;
    box-shadow: 0 10px 30px rgba(0, 91, 143, 0.07);
    border: 1px solid #DDF3F6;
    margin-bottom: 18px;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #E3F6F8 0%, #F7FCFD 100%);
}
[data-testid="stMetric"] {
    background: white;
    padding: 20px;
    border-radius: 22px;
    box-shadow: 0 10px 28px rgba(0, 91, 143, 0.08);
    border-left: 7px solid #00A6B4;
}
[data-testid="stMetricValue"] {
    color: #0097A9;
    font-weight: 900;
}
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(90deg, #0097A9, #00B8C8);
    color: white;
    border-radius: 14px;
    padding: 10px 24px;
    border: none;
    font-weight: 800;
}
.risk-high {
    background: #FFECEC;
    padding: 14px;
    border-radius: 16px;
    border-left: 6px solid #D64545;
    margin-bottom: 10px;
}
.risk-medium {
    background: #FFF7E6;
    padding: 14px;
    border-radius: 16px;
    border-left: 6px solid #F2A900;
    margin-bottom: 10px;
}
.risk-low {
    background: #EAF9F1;
    padding: 14px;
    border-radius: 16px;
    border-left: 6px solid #2EAD70;
    margin-bottom: 10px;
}
.small-note {
    color: #5B6B7A;
    font-size: 14px;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="hero">
    <h1>Интеллектуальная аналитическая система контроля качества контакт-центра</h1>
    <div class="hero-subtitle">
        Система поддержки управленческих решений для анализа претензионной деятельности
    </div>
    <p class="hero-text">
        Платформа выполняет анализ претензий пациентов, выявляет проблемные зоны,
        оценивает качество работы операторов, рассчитывает риск-индексы и формирует
        управленческие рекомендации для повышения эффективности обслуживания.
    </p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# DATA SOURCE
# =========================================================

st.markdown("### Источник данных")

data_source = st.radio(
    "Выберите источник данных",
    ["Excel-файл", "Google Sheets"],
    horizontal=True
)

df_original = None

if data_source == "Excel-файл":
    uploaded_file = st.file_uploader(
        "Загрузите Excel-файл с претензиями",
        type=["xlsx", "xls"]
    )
    if uploaded_file is not None:
        df_original = pd.read_excel(uploaded_file)

if data_source == "Google Sheets":
    google_url = st.text_input("Вставьте ссылку на Google Sheets")
    st.caption("Для Google Sheets должен быть открыт доступ по ссылке или таблица должна быть опубликована в интернете.")
    if google_url:
        try:
            if "/edit" in google_url:
                sheet_id = google_url.split("/d/")[1].split("/")[0]
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            else:
                csv_url = google_url
            df_original = pd.read_csv(csv_url)
        except Exception as e:
            st.error("Не удалось загрузить Google Sheets. Проверьте доступ по ссылке.")
            st.write(e)

if df_original is None:
    st.info("Загрузите Excel-файл или вставьте ссылку на Google Sheets.")
    st.stop()

df_original = df_original.dropna(how="all")
df_original = df_original.dropna(axis=1, how="all")

# =========================================================
# COLUMN SEARCH
# =========================================================

def find_column(df, possible_names):
    for col in df.columns:
        col_clean = str(col).lower().strip()
        for name in possible_names:
            if name.lower() in col_clean:
                return col
    return None

col_year = find_column(df_original, ["год"])
col_month = find_column(df_original, ["месяц"])
col_week = find_column(df_original, ["неделя"])
col_sv = find_column(df_original, ["св"])
col_city = find_column(df_original, ["город"])
col_region = find_column(df_original, ["агломерация"])
col_operator = find_column(df_original, ["фио оператора", "оператор"])
col_channel = find_column(df_original, ["канал поступления"])
col_consult_topic = find_column(df_original, ["тематика консультации"])
col_claim_topic = find_column(df_original, ["тематика жалобы"])
col_claim_text = find_column(df_original, ["жалоба пациента", "жалоба на", "жалоба"])
col_operator_error = find_column(df_original, ["в чем ошибка оператора", "ошибка оператора"])
col_error_reason = find_column(df_original, ["причина ошибки"])
col_validity = find_column(df_original, ["обоснованность"])
col_confirmed = find_column(df_original, ["за кем подтверждена"])
col_date = find_column(df_original, ["дата поступления жалобы", "дата составления жалобы", "дата"])

# =========================================================
# CLEANING
# =========================================================

df = df_original.copy()

def normalize_spaces(text):
    return " ".join(str(text).replace("\n", " ").replace("\r", " ").split())

def clean_text_value(value):
    if pd.isna(value):
        return "Не указано"

    text = normalize_spaces(value)

    if text.lower() in ["nan", "none", "", "nat"]:
        return "Не указано"

    lower = text.lower().replace("ё", "е")

    dictionary = {
        "нет": "Нет",
        "не": "Нет",
        "да": "Да",
        "обоснована": "Обоснована",
        "обоснованная": "Обоснована",
        "обосновано": "Обоснована",
        "не обоснована": "Не обоснована",
        "необоснована": "Не обоснована",
        "не обоснованная": "Не обоснована",
        "необоснованная": "Не обоснована",
        "оператор": "Оператор",
        "кц": "Контакт-центр",
        "контакт центр": "Контакт-центр",
        "контакт-центр": "Контакт-центр",
    }

    if lower in dictionary:
        return dictionary[lower]

    if lower.startswith("св"):
        number = "".join([ch for ch in lower if ch.isdigit()])
        return f"СВ{number}" if number else text.upper()

    return text[0].upper() + text[1:] if len(text) > 1 else text.upper()

for col in df.columns:
    if df[col].dtype == "object":
        df[col] = df[col].apply(clean_text_value)

# Analytical error columns:
if col_error_reason:
    df["_Причина_ошибки_аналитическая"] = df[col_error_reason].apply(
        lambda x: "Ошибки нет" if clean_text_value(x) == "Нет" else clean_text_value(x)
    )
else:
    df["_Причина_ошибки_аналитическая"] = "Не указано"

if col_operator_error:
    df["_Ошибка_оператора_аналитическая"] = df[col_operator_error].apply(
        lambda x: "Ошибки нет" if clean_text_value(x) == "Нет" else clean_text_value(x)
    )
else:
    df["_Ошибка_оператора_аналитическая"] = "Не указано"

# =========================================================
# DATE FEATURES
# =========================================================

if col_date and col_date in df.columns:
    df["_Дата"] = pd.to_datetime(df[col_date], errors="coerce", dayfirst=True)
    df["_День"] = df["_Дата"].dt.strftime("%d.%m").fillna("Не указано")
    df["_Неделя"] = df["_Дата"].dt.isocalendar().week.astype(str)
else:
    df["_Дата"] = pd.NaT
    df["_День"] = "Не указано"
    df["_Неделя"] = "Не указано"

if col_week and col_week in df.columns:
    df["_Неделя"] = df[col_week].astype(str).apply(clean_text_value)

if col_month and col_month in df.columns:
    df["_Месяц"] = df[col_month].astype(str).apply(clean_text_value)
else:
    df["_Месяц"] = df["_Дата"].dt.strftime("%m").fillna("Не указано")

# =========================================================
# CASCADING FILTERS
# =========================================================

st.sidebar.markdown("## Фильтры анализа")

if st.sidebar.button("Сбросить фильтры"):
    for key in list(st.session_state.keys()):
        if key.startswith("f_") or key == "operator_search":
            del st.session_state[key]
    st.rerun()

df_filtered = df.copy()

def cascading_filter(label, column, key):
    global df_filtered

    if column is None or column not in df.columns:
        return

    options = sorted([
        x for x in df_filtered[column].dropna().unique()
        if str(x) != "Не указано"
    ])

    if key in st.session_state:
        st.session_state[key] = [
            x for x in st.session_state[key]
            if x in options
        ]

    selected = st.sidebar.multiselect(label, options, key=key)

    if selected:
        df_filtered = df_filtered[df_filtered[column].isin(selected)]

cascading_filter("СВ", col_sv, "f_sv")
cascading_filter("Месяц", "_Месяц", "f_month")
cascading_filter("Неделя", "_Неделя", "f_week")
cascading_filter("День", "_День", "f_day")
cascading_filter("Город", col_city, "f_city")
cascading_filter("Агломерация", col_region, "f_region")

operator_search = st.sidebar.text_input("Поиск оператора", key="operator_search")

if col_operator and operator_search:
    df_filtered = df_filtered[
        df_filtered[col_operator]
        .astype(str)
        .str.contains(operator_search, case=False, na=False)
    ]

cascading_filter("Оператор", col_operator, "f_operator")
cascading_filter("Тематика жалобы", col_claim_topic, "f_claim_topic")
cascading_filter("Обоснованность", col_validity, "f_validity")

df = df_filtered.copy()

if len(df) == 0:
    st.warning("По выбранным фильтрам данных нет. Измените фильтры или нажмите «Сбросить фильтры».")
    st.stop()

# =========================================================
# KPI
# =========================================================

st.markdown("## Dashboard руководителя")

total_claims = len(df)

if col_validity and total_claims > 0:
    valid_count = (df[col_validity] == "Обоснована").sum()
    invalid_count = (df[col_validity] == "Не обоснована").sum()
    valid_share = valid_count / total_claims * 100
    invalid_share = invalid_count / total_claims * 100
else:
    valid_count = 0
    invalid_count = 0
    valid_share = 0
    invalid_share = 0

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

k1.metric("Всего претензий", total_claims)
k2.metric("Операторов", df[col_operator].nunique() if col_operator else "—")
k3.metric("Городов", df[col_city].nunique() if col_city else "—")
k4.metric("СВ", df[col_sv].nunique() if col_sv else "—")
k5.metric("Тем жалоб", df[col_claim_topic].nunique() if col_claim_topic else "—")
k6.metric("Обоснованных", f"{valid_share:.1f}%")
k7.metric("Необоснованных", f"{invalid_share:.1f}%")

# =========================================================
# OPERATOR RISK AND QUALITY INDEX
# =========================================================

operator_stats = None
risk_index = 0
quality_index = 100

if col_operator:
    df["_Обоснована"] = df[col_validity].apply(lambda x: 1 if x == "Обоснована" else 0) if col_validity else 0
    df["_Есть_ошибка"] = df["_Ошибка_оператора_аналитическая"].apply(lambda x: 0 if x == "Ошибки нет" else 1)

    operator_stats = (
        df.groupby(col_operator)
        .agg(
            Количество_претензий=(col_operator, "count"),
            Обоснованные_претензии=("_Обоснована", "sum"),
            Подтвержденные_ошибки=("_Есть_ошибка", "sum")
        )
        .sort_values("Количество_претензий", ascending=False)
    )

    operator_stats["Доля_обоснованных_%"] = (
        operator_stats["Обоснованные_претензии"] /
        operator_stats["Количество_претензий"] * 100
    ).round(1)

    max_claims = max(operator_stats["Количество_претензий"].max(), 1)

    operator_stats["Индекс_риска"] = (
        operator_stats["Количество_претензий"] / max_claims * 35 +
        operator_stats["Доля_обоснованных_%"] * 0.4 +
        operator_stats["Подтвержденные_ошибки"] / max_claims * 25
    ).clip(upper=100).round(1)

    operator_stats["Индекс_качества"] = (
        100 - operator_stats["Индекс_риска"]
    ).clip(lower=0).round(1)

    def risk_level(row):
        if row["Индекс_риска"] >= 70:
            return "Высокий риск"
        elif row["Индекс_риска"] >= 40:
            return "Средний риск"
        return "Низкий риск"

    operator_stats["Риск"] = operator_stats.apply(risk_level, axis=1)

    risk_index = float(operator_stats["Индекс_риска"].mean().round(1))
    quality_index = float(operator_stats["Индекс_качества"].mean().round(1))

q1, q2, q3 = st.columns(3)
q1.metric("Индекс качества операторов", f"{quality_index:.1f}/100")
q2.metric("Индекс риска", f"{risk_index:.1f}/100")
q3.metric("Уровень управленческого риска", "Высокий" if risk_index >= 70 else "Средний" if risk_index >= 40 else "Низкий")

with st.expander("Методика расчета индексов"):
    st.markdown("""
### Индекс риска оператора

```text
Индекс риска =
(количество претензий оператора / максимум претензий среди операторов × 35)
+ (доля обоснованных претензий × 0.4)
+ (подтвержденные ошибки / максимум претензий среди операторов × 25)
```

### Индекс качества оператора

```text
Индекс качества = 100 - Индекс риска
```

### Уровень управленческого риска

```text
0–39   — низкий риск
40–69  — средний риск
70–100 — высокий риск
```

Методика используется для приоритизации операторов, требующих дополнительного контроля,
аудита звонков или адресного обучения.
""")

# =========================================================
# EXECUTIVE DASHBOARD
# =========================================================

st.markdown("## Executive Dashboard")

d1, d2, d3 = st.columns(3)

critical_text = (
    "Высокая доля обоснованных претензий требует управленческого вмешательства."
    if valid_share >= 50
    else "Критическая доля обоснованных претензий не выявлена."
)

risk_text = (
    "Риск повышен из-за концентрации претензий и подтвержденных ошибок по отдельным операторам."
    if risk_index >= 40
    else "Риск находится на контролируемом уровне."
)

action_text = (
    "Рекомендуется адресное обучение операторов, аудит звонков и корректировка скриптов коммуникации."
)

with d1:
    st.markdown(f"""
    <div class="section-card">
    <h3>Критическая зона</h3>
    <p>{critical_text}</p>
    </div>
    """, unsafe_allow_html=True)

with d2:
    st.markdown(f"""
    <div class="section-card">
    <h3>Индекс риска</h3>
    <p>{risk_text}</p>
    </div>
    """, unsafe_allow_html=True)

with d3:
    st.markdown(f"""
    <div class="section-card">
    <h3>Рекомендуемое действие</h3>
    <p>{action_text}</p>
    </div>
    """, unsafe_allow_html=True)

with st.expander("Исходные данные после фильтрации"):
    st.dataframe(df, width="stretch")

with st.expander("Сопоставление найденных столбцов"):
    st.write({
        "Год": col_year,
        "Месяц": col_month,
        "Неделя": col_week,
        "СВ": col_sv,
        "Город": col_city,
        "Агломерация": col_region,
        "Оператор": col_operator,
        "Тематика жалобы": col_claim_topic,
        "Ошибка оператора": col_operator_error,
        "Причина ошибки": col_error_reason,
        "Обоснованность": col_validity,
        "Дата": col_date
    })

# =========================================================
# VISUAL FUNCTIONS
# =========================================================

def shorten_text(text, max_len=35):
    text = str(text)
    return text if len(text) <= max_len else text[:max_len] + "..."

def make_unique_columns(columns):
    seen = {}
    result = []
    for col in columns:
        col = str(col)
        if col not in seen:
            seen[col] = 0
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")
    return result

def show_top(title, column, top_n=15, exclude_no=False):
    if column is None or column not in df.columns:
        return None

    st.markdown(f"### {title}")

    series = df[column].fillna("Не указано").apply(clean_text_value)

    if exclude_no:
        series = series[series != "Ошибки нет"]
        series = series[series != "Нет"]

    values = series.value_counts().head(top_n)

    if len(values) == 0:
        st.info("По данному блоку отсутствуют подтвержденные ошибки.")
        return values

    st.dataframe(values.rename("Количество"), width="stretch")

    fig, ax = plt.subplots(figsize=(11, 6))
    values.sort_values().plot(kind="barh", ax=ax, color="#00A6B4")
    ax.set_title(title, color="#005B8F", fontsize=14, fontweight="bold")
    ax.set_xlabel("Количество")
    ax.set_ylabel("")
    plt.tight_layout()
    st.pyplot(fig)

    return values

def show_heatmap(title, row_col, col_col, top_rows=15, top_cols=15, exclude_no=False):
    if row_col is None or col_col is None:
        return None

    st.markdown(f"### {title}")

    work_df = df.copy()

    if exclude_no:
        work_df = work_df[
            (work_df[col_col] != "Ошибки нет") &
            (work_df[col_col] != "Нет")
        ]

    if len(work_df) == 0:
        st.info("Нет данных для построения матрицы.")
        return None

    top_row_values = work_df[row_col].value_counts().head(top_rows).index
    top_col_values = work_df[col_col].value_counts().head(top_cols).index

    filtered_df = work_df[
        work_df[row_col].isin(top_row_values) &
        work_df[col_col].isin(top_col_values)
    ]

    matrix = pd.crosstab(filtered_df[row_col], filtered_df[col_col])

    matrix.index = [shorten_text(x, 28) for x in matrix.index]
    matrix.columns = [shorten_text(x, 32) for x in matrix.columns]
    matrix.columns = make_unique_columns(matrix.columns)

    st.dataframe(matrix, width="stretch")

    if matrix.shape[0] > 0 and matrix.shape[1] > 0:
        fig, ax = plt.subplots(figsize=(12, 7))
        heatmap = ax.imshow(matrix.values, aspect="auto", cmap="PuBuGn")

        ax.set_xticks(range(len(matrix.columns)))
        ax.set_yticks(range(len(matrix.index)))
        ax.set_xticklabels(matrix.columns, rotation=35, ha="right", fontsize=8)
        ax.set_yticklabels(matrix.index, fontsize=9)
        ax.set_title(title, fontsize=16, color="#005B8F", fontweight="bold", pad=18)

        fig.colorbar(heatmap)
        plt.tight_layout()
        st.pyplot(fig)

    return matrix

# =========================================================
# TABS
# =========================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Основная аналитика",
    "👥 Операторы и риски",
    "🔥 Матрицы и Heatmap",
    "📝 Текстовый анализ",
    "🧠 Генеративная аналитика"
])

with tab1:
    top_claim_topics = show_top("Топ тематик жалоб", col_claim_topic)
    top_consult_topics = show_top("Топ тематик консультаций", col_consult_topic)
    top_errors = show_top("Топ подтвержденных ошибок операторов", "_Ошибка_оператора_аналитическая", exclude_no=True)
    top_error_reasons = show_top("Топ причин подтвержденных ошибок", "_Причина_ошибки_аналитическая", exclude_no=True)
    top_validity = show_top("Обоснованность жалоб", col_validity)
    top_sv = show_top("Анализ по СВ", col_sv)
    top_week = show_top("Анализ по неделям", "_Неделя")
    top_month = show_top("Анализ по месяцам", "_Месяц")
    top_day = show_top("Анализ по дням", "_День")

with tab2:
    st.markdown("### Рейтинг операторов по качеству и риску")

    if operator_stats is not None:
        st.dataframe(operator_stats, width="stretch")

        high_risk = operator_stats[operator_stats["Риск"] == "Высокий риск"]
        medium_risk = operator_stats[operator_stats["Риск"] == "Средний риск"]

        st.markdown("### Операторы высокого риска")

        if len(high_risk) > 0:
            for name, row in high_risk.head(10).iterrows():
                st.markdown(f"""
                <div class="risk-high">
                <b>{name}</b><br>
                Претензий: {row['Количество_претензий']}<br>
                Подтвержденные ошибки: {row['Подтвержденные_ошибки']}<br>
                Доля обоснованных: {row['Доля_обоснованных_%']}%<br>
                Индекс риска: {row['Индекс_риска']}<br>
                Индекс качества: {row['Индекс_качества']}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.success("Операторов высокого риска не выявлено.")

        st.markdown("### Операторы среднего риска")

        if len(medium_risk) > 0:
            for name, row in medium_risk.head(10).iterrows():
                st.markdown(f"""
                <div class="risk-medium">
                <b>{name}</b><br>
                Претензий: {row['Количество_претензий']}<br>
                Подтвержденные ошибки: {row['Подтвержденные_ошибки']}<br>
                Индекс риска: {row['Индекс_риска']}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.success("Операторов среднего риска не выявлено.")

with tab3:
    show_heatmap("Heatmap: оператор × подтвержденная ошибка", col_operator, "_Ошибка_оператора_аналитическая", exclude_no=True)
    show_heatmap("Heatmap: СВ × подтвержденная ошибка", col_sv, "_Ошибка_оператора_аналитическая", top_rows=10, exclude_no=True)
    show_heatmap("Heatmap: тематика жалобы × причина ошибки", col_claim_topic, "_Причина_ошибки_аналитическая", exclude_no=True)

with tab4:
    st.markdown("### Текстовый анализ жалоб")

    negative_words = [
        "долго", "ждал", "ждала", "ожидание", "не дозвонился", "не дозвонилась",
        "грубо", "хамство", "ошибка", "не помогли", "отказ", "неверно",
        "некорректно", "претензия", "возврат", "обман", "не сообщили",
        "не предупредили", "плохо", "не записали", "не пришло",
        "не отправили", "не перевели", "не объяснили"
    ]

    if col_claim_text:
        def count_problem_words(text):
            text = str(text).lower()
            return sum(1 for word in negative_words if word in text)

        df["Количество проблемных слов"] = df[col_claim_text].apply(count_problem_words)

        def priority(row):
            score = row["Количество проблемных слов"]
            if col_validity and row.get(col_validity) == "Обоснована":
                score += 2
            if row["_Ошибка_оператора_аналитическая"] != "Ошибки нет":
                score += 1

            if score >= 4:
                return "Высокий"
            elif score >= 2:
                return "Средний"
            return "Низкий"

        df["Приоритет жалобы"] = df.apply(priority, axis=1)

        p1, p2, p3 = st.columns(3)
        p1.metric("Высокий приоритет", len(df[df["Приоритет жалобы"] == "Высокий"]))
        p2.metric("Средний приоритет", len(df[df["Приоритет жалобы"] == "Средний"]))
        p3.metric("Низкий приоритет", len(df[df["Приоритет жалобы"] == "Низкий"]))

        show_top("Распределение приоритетов", "Приоритет жалобы")

        cols_to_show = [
            c for c in [
                col_sv, col_operator, col_city, col_claim_topic,
                "_Ошибка_оператора_аналитическая", col_claim_text,
                "Количество проблемных слов", "Приоритет жалобы"
            ]
            if c is not None and c in df.columns
        ]

        st.dataframe(
            df[cols_to_show].sort_values("Количество проблемных слов", ascending=False).head(50),
            width="stretch"
        )

# =========================================================
# RECOMMENDATIONS
# =========================================================

recommendations = []

if "top_claim_topics" in locals() and top_claim_topics is not None and len(top_claim_topics) > 0:
    recommendations.append(
        f"Основная тематика жалоб — «{top_claim_topics.index[0]}». Требуется детальный разбор процесса по этой тематике."
    )

if "top_errors" in locals() and top_errors is not None and len(top_errors) > 0:
    recommendations.append(
        f"Наиболее частая подтвержденная ошибка — «{top_errors.index[0]}». Рекомендуется обновить чек-листы и провести адресное обучение."
    )

if "top_error_reasons" in locals() and top_error_reasons is not None and len(top_error_reasons) > 0:
    recommendations.append(
        f"Ключевая причина подтвержденных ошибок — «{top_error_reasons.index[0]}». Следует уточнить регламенты и сценарии коммуникации."
    )

if valid_share > 50:
    recommendations.append(
        f"Доля обоснованных жалоб составляет {valid_share:.1f}%. Необходимы системные корректирующие мероприятия."
    )

if invalid_share > 40:
    recommendations.append(
        f"Доля необоснованных жалоб составляет {invalid_share:.1f}%. Рекомендуется улучшить информирование пациентов о правилах и ограничениях сервиса."
    )

if risk_index >= 70:
    recommendations.append("Индекс риска высокий. Требуется срочный аудит качества работы операторов.")
elif risk_index >= 40:
    recommendations.append("Индекс риска средний. Рекомендуется адресное обучение операторов группы риска.")
else:
    recommendations.append("Индекс риска низкий. Рекомендуется поддерживать текущий контроль качества и мониторинг динамики.")

# =========================================================
# GENERATIVE ANALYTICS
# =========================================================

with tab5:
    st.markdown("### Генеративная аналитика")

    question = st.text_area("Введите управленческий вопрос", height=120)

    if st.button("Сформировать экспертный вывод"):
        if not question.strip():
            st.warning("Введите вопрос.")
        else:
            top_operator_text = ""

            if operator_stats is not None and len(operator_stats) > 0:
                top_operator = operator_stats.index[0]
                top_operator_row = operator_stats.iloc[0]
                top_operator_text = (
                    f"Наибольшая концентрация претензий связана с оператором «{top_operator}»: "
                    f"{top_operator_row['Количество_претензий']} претензий, "
                    f"{top_operator_row['Подтвержденные_ошибки']} подтвержденных ошибок, "
                    f"индекс риска {top_operator_row['Индекс_риска']}."
                )

            st.success("Экспертный вывод:")

            st.write(f"""
**1. Краткий управленческий вывод**

По выбранной выборке проанализировано **{total_claims} претензий**.  
Доля обоснованных претензий составляет **{valid_share:.1f}%**, доля необоснованных — **{invalid_share:.1f}%**.  
Индекс качества операторов — **{quality_index:.1f}/100**, индекс риска — **{risk_index:.1f}/100**.

**2. Основные проблемные зоны**

{recommendations[0] if len(recommendations) > 0 else "Проблемные зоны требуют дополнительного анализа."}

{top_operator_text}

**3. Интерпретация для руководителя**

Если претензии обоснованы и сопровождаются подтвержденными ошибками операторов, проблема связана не только с восприятием пациента, но и с нарушением процесса обслуживания.  
Если преобладают необоснованные претензии, основной управленческий акцент должен быть направлен на информирование пациентов, корректность ожиданий и прозрачность регламентов.

**4. Рекомендации**

{chr(10).join("- " + r for r in recommendations)}

**5. План корректирующих мероприятий**

1. Провести выборочный аудит звонков по основным тематикам жалоб.  
2. Сформировать список операторов группы риска.  
3. Провести адресное обучение по частым подтвержденным ошибкам.  
4. Обновить чек-листы качества и скрипты коммуникации.  
5. Через 2–4 недели повторно измерить индекс риска и индекс качества операторов.
""")

# =========================================================
# EXPORT
# =========================================================

st.markdown("## Экспорт результатов")

report_text = f"""
ОТЧЕТ ПО АНАЛИЗУ ПРЕТЕНЗИОННОЙ ДЕЯТЕЛЬНОСТИ КОНТАКТ-ЦЕНТРА

Дата формирования: {datetime.now().strftime("%d.%m.%Y %H:%M")}

Всего претензий: {total_claims}
Доля обоснованных жалоб: {valid_share:.1f}%
Доля необоснованных жалоб: {invalid_share:.1f}%
Индекс качества операторов: {quality_index:.1f}/100
Индекс риска: {risk_index:.1f}/100

Управленческие рекомендации:
{chr(10).join("- " + r for r in recommendations) if recommendations else "Нет данных"}
"""

st.download_button(
    "Скачать управленческий отчет TXT",
    data=report_text,
    file_name="contact_center_quality_report.txt",
    mime="text/plain"
)

html_report = f"""
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    font-family: Arial, sans-serif;
    color: #1F2D3D;
    padding: 30px;
}}
h1, h2 {{
    color: #005B8F;
}}
.card {{
    background: #E4F5F8;
    padding: 16px;
    border-radius: 12px;
    margin-bottom: 12px;
}}
</style>
</head>
<body>
<h1>Управленческий отчет по анализу претензионной деятельности контакт-центра</h1>

<div class="card">
<p><b>Дата формирования:</b> {datetime.now().strftime("%d.%m.%Y %H:%M")}</p>
<p><b>Всего претензий:</b> {total_claims}</p>
<p><b>Доля обоснованных жалоб:</b> {valid_share:.1f}%</p>
<p><b>Доля необоснованных жалоб:</b> {invalid_share:.1f}%</p>
<p><b>Индекс качества операторов:</b> {quality_index:.1f}/100</p>
<p><b>Индекс риска:</b> {risk_index:.1f}/100</p>
</div>

<h2>Управленческие рекомендации</h2>
<ul>
{''.join(f'<li>{r}</li>' for r in recommendations)}
</ul>

<h2>Методика расчета</h2>
<p><b>Индекс риска</b> учитывает количество претензий, долю обоснованных жалоб и подтвержденные ошибки.</p>
<p><b>Индекс качества оператора</b> = 100 - индекс риска.</p>

</body>
</html>
"""

st.download_button(
    "Скачать HTML-отчет",
    data=html_report,
    file_name="contact_center_quality_report.html",
    mime="text/html"
)

excel_buffer = BytesIO()
df.to_excel(excel_buffer, index=False, engine="openpyxl")
excel_buffer.seek(0)

st.download_button(
    "Скачать очищенные данные Excel",
    data=excel_buffer,
    file_name="cleaned_contact_center_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)