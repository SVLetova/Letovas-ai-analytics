import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

st.set_page_config(page_title="LetovaS DSS", page_icon="📊", layout="wide")

# ===================== DESIGN =====================

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #F6FCFD 0%, #FFFFFF 45%, #E4F5F8 100%);
    color: #1F2D3D;
}
h1 {
    color: #0097A9 !important;
    font-size: 52px !important;
    font-weight: 900 !important;
}
h2, h3 {
    color: #005B8F !important;
    font-weight: 800 !important;
}
.hero {
    background: linear-gradient(135deg, #FFFFFF 0%, #E1F5F8 55%, #B9E5EB 100%);
    padding: 38px;
    border-radius: 30px;
    margin-bottom: 30px;
    box-shadow: 0 18px 45px rgba(0, 151, 169, 0.16);
    border: 1px solid #D5F0F4;
}
.hero-subtitle {
    color: #005B8F;
    font-size: 24px;
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
.stButton > button {
    background: linear-gradient(90deg, #0097A9, #00B8C8);
    color: white;
    border-radius: 14px;
    padding: 10px 24px;
    border: none;
    font-weight: 800;
}
.stDownloadButton > button {
    background: linear-gradient(90deg, #005B8F, #0097A9);
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
</style>
""", unsafe_allow_html=True)

# ===================== HEADER =====================

st.markdown("""
<div class="hero">
    <h1>LetovaS DSS</h1>
    <div class="hero-subtitle">
        Интеллектуальная система поддержки управленческих решений
    </div>
    <p class="hero-text">
        Аналитическая платформа для мониторинга претензионной деятельности,
        оценки качества обслуживания, выявления ошибок операторов,
        прогнозирования нагрузки, риск-анализа и формирования управленческих рекомендаций
        для контакт-центра медицинской сети.
    </p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Загрузите Excel-файл с претензиями", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Загрузите Excel-файл для запуска аналитического модуля.")
    st.stop()

# ===================== LOAD DATA =====================

df_original = pd.read_excel(uploaded_file)
df_original = df_original.dropna(how="all")
df_original = df_original.dropna(axis=1, how="all")

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
col_plan = find_column(df_original, ["план корректирующих", "корректирующих мероприятий"])
col_confirmed = find_column(df_original, ["за кем подтверждена"])
col_date = find_column(df_original, ["дата поступления жалобы", "дата составления жалобы", "дата"])

df = df_original.copy()

def clean_text_value(value):
    if pd.isna(value):
        return "Не указано"

    text = str(value).strip()
    text = " ".join(text.split())

    if text.lower() in ["nan", "none", "", "nat"]:
        return "Не указано"

    lower = text.lower()

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
        "клиент": "Клиент"
    }

    if lower in dictionary:
        return dictionary[lower]

    return text[0].upper() + text[1:] if len(text) > 1 else text.upper()

for col in df.columns:
    if df[col].dtype == "object":
        df[col] = df[col].apply(clean_text_value)

# ===================== DATE FEATURES =====================

if col_date and col_date in df.columns:
    df["_Дата"] = pd.to_datetime(df[col_date], errors="coerce", dayfirst=True)
    df["_День"] = df["_Дата"].dt.day.fillna("Не указано").astype(str)
    df["_Месяц_номер"] = df["_Дата"].dt.month
else:
    df["_Дата"] = pd.NaT
    df["_День"] = "Не указано"
    df["_Месяц_номер"] = None

if col_month and col_month in df.columns:
    df["_Месяц"] = df[col_month].astype(str).apply(clean_text_value)
else:
    df["_Месяц"] = df["_Месяц_номер"].astype(str)

# ===================== FILTERS =====================

st.sidebar.markdown("## Фильтры анализа")

if st.sidebar.button("Сбросить фильтры"):
    st.session_state.clear()
    st.rerun()

df_filtered = df.copy()

def add_filter(label, column, key):
    global df_filtered

    if column is None or column not in df.columns:
        return

    options = sorted([
        x for x in df[column].dropna().unique()
        if str(x) != "Не указано"
    ])

    selected = st.sidebar.multiselect(label, options, key=key)

    if selected:
        df_filtered = df_filtered[df_filtered[column].isin(selected)]

add_filter("СВ", col_sv, "f_sv")
add_filter("Месяц", "_Месяц", "f_month")
add_filter("День", "_День", "f_day")
add_filter("Город", col_city, "f_city")
add_filter("Агломерация", col_region, "f_region")
add_filter("Оператор", col_operator, "f_operator")
add_filter("Тематика жалобы", col_claim_topic, "f_claim_topic")
add_filter("Обоснованность", col_validity, "f_validity")

df = df_filtered.copy()

# ===================== KPI =====================

st.markdown("## Dashboard руководителя")

total_claims = len(df)

if col_validity and total_claims > 0:
    valid_share = (df[col_validity] == "Обоснована").sum() / total_claims * 100
else:
    valid_share = 0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Всего претензий", total_claims)
k2.metric("Операторов", df[col_operator].nunique() if col_operator else "—")
k3.metric("Городов", df[col_city].nunique() if col_city else "—")
k4.metric("СВ", df[col_sv].nunique() if col_sv else "—")
k5.metric("Тем жалоб", df[col_claim_topic].nunique() if col_claim_topic else "—")
k6.metric("Обоснованных", f"{valid_share:.1f}%")

# ===================== QUALITY AND RISK INDEXES =====================

operator_stats = None
risk_index = 0
quality_index = 100

if col_operator:
    df["_Обоснована"] = 0
    if col_validity:
        df["_Обоснована"] = df[col_validity].apply(lambda x: 1 if x == "Обоснована" else 0)

    operator_stats = (
        df.groupby(col_operator)
        .agg(
            Количество_претензий=(col_operator, "count"),
            Обоснованные_претензии=("_Обоснована", "sum")
        )
        .sort_values("Количество_претензий", ascending=False)
    )

    operator_stats["Доля_обоснованных_%"] = (
        operator_stats["Обоснованные_претензии"] /
        operator_stats["Количество_претензий"] * 100
    ).round(1)

    max_claims = operator_stats["Количество_претензий"].max() if len(operator_stats) > 0 else 1

    operator_stats["Индекс_риска"] = (
        operator_stats["Количество_претензий"] / max_claims * 50 +
        operator_stats["Доля_обоснованных_%"] * 0.5
    ).round(1)

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

    risk_index = operator_stats["Индекс_риска"].mean().round(1)
    quality_index = operator_stats["Индекс_качества"].mean().round(1)

q1, q2, q3 = st.columns(3)
q1.metric("Индекс качества операторов", f"{quality_index:.1f}/100")
q2.metric("Индекс риска", f"{risk_index:.1f}/100")
q3.metric("Уровень управленческого риска", "Высокий" if risk_index >= 70 else "Средний" if risk_index >= 40 else "Низкий")

# ===================== EXECUTIVE DASHBOARD =====================

st.markdown("## Executive Dashboard")

d1, d2, d3 = st.columns(3)

critical_text = "Доля обоснованных жалоб требует контроля." if valid_share >= 50 else "Критическая доля обоснованных жалоб не выявлена."
risk_text = "Риск повышен из-за концентрации претензий по отдельным операторам." if risk_index >= 40 else "Риск находится на контролируемом уровне."
action_text = "Провести адресное обучение операторов и актуализировать регламенты."

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

# ===================== VISUAL FUNCTIONS =====================

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

def show_top(title, column, top_n=15):
    if column is None or column not in df.columns:
        return None

    st.markdown(f"### {title}")

    values = (
        df[column]
        .fillna("Не указано")
        .apply(clean_text_value)
        .value_counts()
        .head(top_n)
    )

    st.dataframe(values.rename("Количество"), width="stretch")

    fig, ax = plt.subplots(figsize=(11, 6))
    values.sort_values().plot(kind="barh", ax=ax, color="#00A6B4")
    ax.set_title(title, color="#005B8F", fontsize=14, fontweight="bold")
    ax.set_xlabel("Количество")
    ax.set_ylabel("")
    plt.tight_layout()
    st.pyplot(fig)

    return values

def show_heatmap(title, row_col, col_col, top_rows=15, top_cols=15):
    if row_col is None or col_col is None:
        return None

    st.markdown(f"### {title}")

    top_row_values = df[row_col].value_counts().head(top_rows).index
    top_col_values = df[col_col].value_counts().head(top_cols).index

    filtered_df = df[
        df[row_col].isin(top_row_values) &
        df[col_col].isin(top_col_values)
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

# ===================== TABS =====================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Основная аналитика",
    "👥 Операторы и риски",
    "🔥 Матрицы и Heatmap",
    "📝 Текстовый анализ",
    "📈 Прогноз нагрузки",
    "🤖 Генеративная аналитика"
])

with tab1:
    top_claim_topics = show_top("Топ тематик жалоб", col_claim_topic)
    top_consult_topics = show_top("Топ тематик консультаций", col_consult_topic)
    top_errors = show_top("Топ ошибок операторов", col_operator_error)
    top_error_reasons = show_top("Топ причин ошибок", col_error_reason)
    top_validity = show_top("Обоснованность жалоб", col_validity)
    top_sv = show_top("Анализ по СВ", col_sv)
    top_month = show_top("Динамика по месяцам", "_Месяц")
    top_day = show_top("Распределение по дням", "_День")

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
                Индекс риска: {row['Индекс_риска']}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.success("Операторов среднего риска не выявлено.")

with tab3:
    matrix_operator_error = show_heatmap("Heatmap: оператор × ошибка", col_operator, col_operator_error)
    matrix_sv_error = show_heatmap("Heatmap: СВ × ошибка", col_sv, col_operator_error, top_rows=10)
    matrix_topic_reason = show_heatmap("Heatmap: тематика жалобы × причина ошибки", col_claim_topic, col_error_reason)

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
                col_operator_error, col_claim_text,
                "Количество проблемных слов", "Приоритет жалобы"
            ]
            if c is not None and c in df.columns
        ]

        st.dataframe(
            df[cols_to_show].sort_values("Количество проблемных слов", ascending=False).head(50),
            width="stretch"
        )

with tab5:
    st.markdown("### Прогноз нагрузки по претензиям")

    forecast_text = "Недостаточно данных для прогноза."

    if "_Дата" in df.columns and df["_Дата"].notna().sum() > 0:
        daily = df.dropna(subset=["_Дата"]).groupby(df["_Дата"].dt.date).size()
        daily.index = pd.to_datetime(daily.index)

        st.dataframe(daily.rename("Количество претензий"), width="stretch")

        if len(daily) >= 2:
            avg_daily = daily.tail(7).mean()
            forecast_next = round(avg_daily, 1)

            forecast_text = f"Прогноз количества претензий на следующий день: {forecast_next}"

            st.metric("Прогноз на следующий день", forecast_next)

            fig, ax = plt.subplots(figsize=(11, 5))
            daily.plot(ax=ax, marker="o", color="#0097A9")
            ax.axhline(avg_daily, linestyle="--", color="#005B8F", label="Среднее последних дней")
            ax.set_title("Динамика поступления претензий", color="#005B8F", fontweight="bold")
            ax.set_xlabel("Дата")
            ax.set_ylabel("Количество")
            ax.legend()
            plt.tight_layout()
            st.pyplot(fig)
        else:
            st.warning("Для прогноза нужна минимум 2 даты.")
    else:
        st.warning("Не удалось определить дату для построения прогноза.")

# ===================== RECOMMENDATIONS =====================

recommendations = []

if "top_claim_topics" in locals() and top_claim_topics is not None and len(top_claim_topics) > 0:
    recommendations.append(
        f"Основная тематика жалоб — «{top_claim_topics.index[0]}». Требуется разбор процесса и проверка скриптов."
    )

if "top_errors" in locals() and top_errors is not None and len(top_errors) > 0:
    recommendations.append(
        f"Наиболее частая ошибка оператора — «{top_errors.index[0]}». Необходимо обновить чек-листы и провести обучение."
    )

if "top_error_reasons" in locals() and top_error_reasons is not None and len(top_error_reasons) > 0:
    recommendations.append(
        f"Основная причина ошибки — «{top_error_reasons.index[0]}». Следует уточнить регламенты и инструкции."
    )

if valid_share > 50:
    recommendations.append(
        f"Доля обоснованных жалоб составляет {valid_share:.1f}%. Нужны системные корректирующие мероприятия."
    )

if risk_index >= 70:
    recommendations.append(
        "Индекс риска высокий. Необходимо срочно провести аудит качества работы операторов."
    )
elif risk_index >= 40:
    recommendations.append(
        "Индекс риска средний. Рекомендуется адресное обучение операторов из группы риска."
    )

with tab6:
    st.markdown("### Автоматические управленческие рекомендации")

    for rec in recommendations:
        st.warning(rec)

    st.markdown("### Генеративная аналитика")

    st.info(
        "В облачной версии генеративный модуль работает в безопасном режиме без Ollama. "
        "Для полноценного ИИ в облаке можно подключить OpenAI API. "
        "Локально Ollama может использоваться на вашем Mac."
    )

    question = st.text_area("Введите управленческий вопрос", height=120)

    if st.button("Сформировать экспертный вывод"):
        if not question.strip():
            st.warning("Введите вопрос.")
        else:
            st.success("Экспертный вывод:")
            st.write(
                f"""
                **Краткий вывод.**  
                Система выявила {total_claims} претензий, долю обоснованных жалоб {valid_share:.1f}%,
                индекс качества операторов {quality_index:.1f}/100 и индекс риска {risk_index:.1f}/100.

                **Основные причины.**  
                {recommendations[0] if len(recommendations) > 0 else "Основные причины требуют дополнительного анализа."}

                **Риски.**  
                Основные риски связаны с концентрацией претензий по отдельным операторам,
                повторяющимися ошибками и обоснованными жалобами.

                **Управленческие рекомендации.**  
                {chr(10).join("- " + r for r in recommendations) if recommendations else "Рекомендуется провести дополнительный аудит данных."}

                **План корректирующих мероприятий.**  
                1. Провести выборочный аудит звонков.  
                2. Обновить чек-листы контроля качества.  
                3. Назначить адресное обучение операторов группы риска.  
                4. Отслеживать динамику жалоб после корректирующих мероприятий.  
                """
            )

# ===================== REPORTS =====================

st.markdown("## Экспорт результатов")

report_text = f"""
ОТЧЕТ LETOVAS DSS

Дата формирования: {datetime.now().strftime("%d.%m.%Y %H:%M")}

Всего претензий: {total_claims}
Доля обоснованных жалоб: {valid_share:.1f}%
Индекс качества операторов: {quality_index:.1f}/100
Индекс риска: {risk_index:.1f}/100

Прогноз:
{forecast_text if 'forecast_text' in locals() else "Нет данных"}

Управленческие рекомендации:
{chr(10).join("- " + r for r in recommendations) if recommendations else "Нет данных"}
"""

st.download_button(
    "Скачать управленческий TXT-отчет",
    data=report_text,
    file_name="LetovaS_DSS_report.txt",
    mime="text/plain"
)

if REPORTLAB_AVAILABLE:
    pdf_buffer = BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    for line in report_text.split("\n"):
        if line.strip():
            story.append(Paragraph(line, styles["Normal"]))
            story.append(Spacer(1, 8))

    doc.build(story)
    pdf_buffer.seek(0)

    st.download_button(
        "Скачать PDF Executive Report",
        data=pdf_buffer,
        file_name="LetovaS_DSS_Executive_Report.pdf",
        mime="application/pdf"
    )
else:
    st.warning("PDF-отчет недоступен. Добавьте reportlab в requirements.txt.")

excel_buffer = BytesIO()
df.to_excel(excel_buffer, index=False, engine="openpyxl")
excel_buffer.seek(0)

st.download_button(
    "Скачать очищенные данные",
    data=excel_buffer,
    file_name="LetovaS_DSS_cleaned_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)