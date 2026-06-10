import streamlit as st
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from io import BytesIO
from datetime import datetime
import textwrap
import re
# Кириллица в графиках и PDF (DejaVu Sans входит в matplotlib и поддерживает русский)
matplotlib.rcParams["font.family"] = "DejaVu Sans"
matplotlib.rcParams["axes.unicode_minus"] = False
matplotlib.rcParams["figure.max_open_warning"] = 0
# Палитра проекта
TEAL = "#0097A9"
DEEP_BLUE = "#005B8F"
ACCENT = "#00A6B4"
# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
page_title="Аналитическая система КЦ",
page_icon=" ",
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
}
[data-testid="stSidebar"] {
background: linear-gradient(180deg, #E3F6F8 0%, #F7FCFD 100%);
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
if source_type == "Excel":
    uploaded_file = st.file_uploader(
        "Загрузите Excel-файл",
        type=["xlsx", "xls"]
    )
if uploaded_file is not None:
df_original = pd.read_excel(uploaded_file)
if data_source == "Google Sheets":
google_url = st.text_input("Вставьте ссылку на Google Sheets")
st.caption(
"Откройте доступ к таблице по ссылке (режим «Просмотр» для всех, у кого есть ссылка)"
"или опубликуйте её в интернете. Поддерживаются ссылки вида /edit, ?gid=, а также "
"опубликованные CSV-ссылки."
)
def build_gsheet_csv_url(url):
"""Преобразует любую ссылку Google Sheets в ссылку экспорта в CSV."""
url = url.strip()
# Уже готовая CSV-ссылка (опубликованная таблица)
if "format=csv" in url or "output=csv" in url or url.endswith(".csv"):
return url
# Извлекаем идентификатор таблицы
id_match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
if not id_match:
id_match = re.search(r"[?&]id=([a-zA-Z0-9-_]+)", url)
if not id_match:
return url # отдаём как есть, pandas попробует сам
sheet_id = id_match.group(1)
# Извлекаем идентификатор конкретного листа (gid), если он есть
gid_match = re.search(r"[#?&]gid=(\d+)", url)
gid = gid_match.group(1) if gid_match else "0"
return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid
if google_url:
csv_url = build_gsheet_csv_url(google_url)
loaded = False
with st.spinner("Загружаю данные из Google Sheets..."):
for enc in ("utf-8", "cp1251", None):
try:
df_original = pd.read_csv(csv_url, encoding=enc) if enc else pd.read_csv(
loaded = True
break
except UnicodeDecodeError:
continue
except Exception as e:
last_error = e
break
if loaded:
st.success("Таблица успешно загружена из Google Sheets.")
else:
st.error(
"Не удалось загрузить таблицу. Проверьте, что:\n"
"• ссылка ведёт на Google Sheets;\n"
"• открыт доступ «всем, у кого есть ссылка»;\n"
"• выбран нужный лист (gid в ссылке)."
)
if df_original is None:
st.info("Загрузите Excel-файл или вставьте ссылку на Google Sheets.")
st.stop()
df_original = df_original.dropna(how="all")
df_original = df_original.dropna(axis=1, how="all")
with st.expander("Предпросмотр загруженных данных", expanded=False):
st.write(f"Загружено строк: **{len(df_original)}**, столбцов: **{df_original.shape[1]}**"
st.dataframe(df_original.head(20), width="stretch", height=320)
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
col_date = find_column(df_original, ["дата поступления жалобы", "дата составления жалобы", "д
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
def normalize_sv(value):
"""Единый формат СВ: устраняет латинские двойники (CB → СВ), регистр и лишние символы."""
text = normalize_spaces(value)
if text.lower() in ["nan", "none", "", "nat", "не указано"]:
return "Не указано"
# Латинские буквы, визуально совпадающие с С и В, приводим к кириллице
low = text.lower().replace("c", "с").replace("b", "в")
m = re.match(r"^\s*с\s*в\D*0*(\d+)", low)
if m:
return f"СВ{m.group(1)}"
return text
def normalize_validity(value):
"""Единый формат обоснованности: убирает дубли по регистру и слитному/раздельному написан
text = normalize_spaces(value)
low = text.lower().replace("ё", "е")
if low in ["nan", "none", "", "nat", "не указано"]:
return "Не указано"
if re.search(r"\bне\s*обоснов", low) or low.startswith("необоснов"):
return "Не обоснована"
if "обоснов" in low:
return "Обоснована"
return text
for col in df.columns:
if df[col].dtype == "object":
df[col] = df[col].apply(clean_text_value)
# Целевая нормализация ключевых категориальных полей (устранение дублей в фильтрах и KPI)
if col_sv and col_sv in df.columns:
df[col_sv] = df[col_sv].apply(normalize_sv)
if col_validity and col_validity in df.columns:
df[col_validity] = df[col_validity].apply(normalize_validity)
# Analytical error columns:
if col_error_reason:
df["_Причина
ошибки
_
_
аналитическая"] = df[col_error_reason].apply(
lambda x: "Ошибки нет" if clean_text_value(x) == "Нет" else clean_text_value(x)
)
else:
df["_Причина
ошибки
аналитическая"] = "Не указано"
_
_
if col_operator_error:
df["_Ошибка
оператора
_
_
аналитическая"] = df[col_operator_error].apply(
lambda x: "Ошибки нет" if clean_text_value(x) == "Нет" else clean_text_value(x)
)
else:
df["_Ошибка
оператора
аналитическая"] = "Не указано"
_
_
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
st.sidebar.markdown("## Параметры отчёта")
with st.sidebar.expander("Реквизиты для PDF", expanded=False):
report_author = st.text_input("Автор (ФИО)", value="")
report_org = st.text_input("Организация / вуз", value="ВГУ")
report_supervisor = st.text_input("Научный руководитель", value="")
st.sidebar.markdown("## Фильтры анализа")
if st.sidebar.button("Сбросить фильтры"):
for key in list(st.session_state.keys()):
if key.startswith("f_") or key == "operator_search":
del st.session_state[key]
st.rerun()
df_filtered = df.copy()
def natural_key(value):
"""Ключ натуральной сортировки: СВ2 раньше СВ10, числа сравниваются как числа."""
parts = re.split(r"(\d+)", str(value))
return [int(p) if p.isdigit() else p.lower() for p in parts]
def cascading_filter(label, column, key):
global df_filtered
if column is None or column not in df.columns:
return
options = sorted(
key=natural_key
[x for x in df_filtered[column].dropna().unique() if str(x) != "Не указано"],
)
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
st.stop()
st.warning("По выбранным фильтрам данных нет. Измените фильтры или нажмите «Сбросить филь
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
df["_Обоснована"] = df[col_validity].apply(lambda x: 1 if x == "Обоснована" else 0) if co
df["_Есть
_
ошибка"] = df["_Ошибка
оператора
аналитическая"].apply(lambda x: 0 if x == "Оши
_
_
operator_stats = (
df.groupby(col_operator)
.agg(
Количество
_
претензий=(col_operator, "count"),
Обоснованные
_
претензии=("_Обоснована", "sum"),
Подтвержденные
_
ошибки=("_Есть
_
ошибка", "sum")
)
.sort_values("Количество
претензий", ascending=False)
_
)
operator_stats["Доля
_
обоснованных_%"] = (
operator_stats["Обоснованные
претензии"] /
_
operator_stats["Количество
претензий"] * 100
_
).round(1)
max_claims = max(operator_stats["Количество
_
претензий"].max(), 1)
operator_stats["Индекс
риска"] = (
_
operator_stats["Количество
_
претензий"] / max_claims * 35 +
operator_stats["Доля
_
обоснованных_%"] * 0.4 +
operator_stats["Подтвержденные
_
ошибки"] / max_claims * 25
).clip(upper=100).round(1)
operator_stats["Индекс
качества"] = (
_
100 - operator_stats["Индекс
риска"]_
).clip(lower=0).round(1)
def risk_level(row):
if row["Индекс
риска"] >= 70:
_
return "Высокий риск"
elif row["Индекс
риска"] >= 40:
_
return "Средний риск"
return "Низкий риск"
operator_stats["Риск"] = operator_stats.apply(risk_level, axis=1)
risk_index = float(operator_stats["Индекс
_
quality_index = float(operator_stats["Индекс
_
риска"].mean().round(1))
качества"].mean().round(1))
q1, q2, q3 = st.columns(3)
q1.metric("Индекс качества операторов", f"{quality_index:.1f}/100")
q2.metric("Индекс риска", f"{risk_index:.1f}/100")
q3.metric("Уровень управленческого риска", "Высокий" if risk_index >= 70 else "Средний" if ri
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
0–39 — низкий риск
40–69 — средний риск
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
"Риск повышен из-за концентрации претензий и подтвержденных ошибок по отдельным оператора
if risk_index >= 40
else "Риск находится на контролируемом уровне."
)
action_text = (
"Рекомендуется адресное обучение операторов, аудит звонков и корректировка скриптов комму
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
# Сюда собираются все построенные диаграммы для последующей выгрузки в PDF
report_figures = []
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
report_figures.append((title, fig))
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
report_figures.append((title, fig))
return matrix
# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
" Основная аналитика",
" Операторы и риски",
" Матрицы и Heatmap",
" Текстовый анализ",
" Генеративная аналитика"
])
with tab1:
top_claim_topics = show_top("Топ тематик жалоб", col_claim_topic)
top_consult_topics = show_top("Топ тематик консультаций", col_consult_topic)
top_errors = show_top("Топ подтвержденных ошибок операторов", "_Ошибка
оператора
_
_
top_error_reasons = show_top("Топ причин подтвержденных ошибок", "_Причина
ошибки
_
_
top_validity = show_top("Обоснованность жалоб", col_validity)
top_sv = show_top("Анализ по СВ", col_sv)
top_week = show_top("Анализ по неделям", "_Неделя")
top_month = show_top("Анализ по месяцам", "_Месяц")
top_day = show_top("Анализ по дням", "_День")
аналитич
аналити
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
Претензий: {row['Количество
претензий']}<br>
_
Подтвержденные ошибки: {row['Подтвержденные
ошибки']}<br>
Доля обоснованных: {row['Доля
Индекс риска: {row['Индекс
_
_
обоснованных_%']}%<br>
риска']}<br>
_
Индекс качества: {row['Индекс
качества']}
_
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
Претензий: {row['Количество
претензий']}<br>
_
Подтвержденные ошибки: {row['Подтвержденные
ошибки']}<br>
_
Индекс риска: {row['Индекс
</div>
риска']}
_
""", unsafe_allow_html=True)
else:
st.success("Операторов среднего риска не выявлено.")
with tab3:
show_heatmap("Heatmap: оператор × подтвержденная ошибка", col_operator, "_Ошибка
_
оператор
show_heatmap("Heatmap: СВ × подтвержденная ошибка", col_sv, "_Ошибка
оператора
_
_
аналитичес
show_heatmap("Heatmap: тематика жалобы × причина ошибки", col_claim_topic, "_Причина
_
ошиб
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
if row["_Ошибка
оператора
аналитическая"] != "Ошибки нет":
_
_
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
"_Ошибка
оператора
_
_
аналитическая", col_claim_text,
"Количество проблемных слов", "Приоритет жалобы"
]
if c is not None and c in df.columns
]
st.dataframe(
df[cols_to_show].sort_values("Количество проблемных слов", ascending=False).head(
width="stretch"
)
# =========================================================
# RECOMMENDATIONS
# =========================================================
recommendations = []
if "top_claim_topics" in locals() and top_claim_topics is not None and len(top_claim_topics)
recommendations.append(
f"Основная тематика жалоб — «{top_claim_topics.index[0]}». Требуется детальный разбор
)
if "top_errors" in locals() and top_errors is not None and len(top_errors) > 0:
recommendations.append(
f"Наиболее частая подтвержденная ошибка — «{top_errors.index[0]}». Рекомендуется обно
)
if "top_error_reasons" in locals() and top_error_reasons is not None and len(top_error_reason
recommendations.append(
f"Ключевая причина подтвержденных ошибок — «{top_error_reasons.index[0]}». Следует ут
)
if valid_share > 50:
recommendations.append(
f"Доля обоснованных жалоб составляет {valid_share:.1f}%. Необходимы системные коррект
)
if invalid_share > 40:
recommendations.append(
f"Доля необоснованных жалоб составляет {invalid_share:.1f}%. Рекомендуется улучшить и
)
if risk_index >= 70:
recommendations.append("Индекс риска высокий. Требуется срочный аудит качества работы опе
elif risk_index >= 40:
recommendations.append("Индекс риска средний. Рекомендуется адресное обучение операторов
else:
recommendations.append("Индекс риска низкий. Рекомендуется поддерживать текущий контроль
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
f"Наибольшая концентрация претензий связана с оператором «{top_operator}»
f"{top_operator_row['Количество
претензий']} претензий, "
_
f"{top_operator_row['Подтвержденные
ошибки']} подтвержденных ошибок, "
_
f"индекс риска {top_operator_row['Индекс
риска']}."
_
)
st.success("Экспертный вывод:")
st.write(f"""
**1. Краткий управленческий вывод**
По выбранной выборке проанализировано **{total_claims} претензий**.
Доля обоснованных претензий составляет **{valid_share:.1f}%**, доля необоснованных — **{inval
Индекс качества операторов — **{quality_index:.1f}/100**, индекс риска — **{risk_index:.1f}/1
**2. Основные проблемные зоны**
{recommendations[0] if len(recommendations) > 0 else "Проблемные зоны требуют дополнительного
{top_operator_text}
**3. Интерпретация для руководителя**
Если претензии обоснованы и сопровождаются подтвержденными ошибками операторов, проблема связ
Если преобладают необоснованные претензии, основной управленческий акцент должен быть направл
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
# ---------------------------------------------------------
# PDF-ОТЧЕТ (показатели + рекомендации + операторы + все диаграммы)
# ---------------------------------------------------------
def build_pdf_report():
buf = BytesIO()
with PdfPages(buf) as pdf:
# ----- Титульная / сводная страница (A4) -----
fig = plt.figure(figsize=(8.27, 11.69))
fig.patch.set_facecolor("white")
ax = fig.add_axes([0, 0, 1, 1])
ax.axis("off")
ax.add_patch(plt.Rectangle((0.06, 0.905), 0.88, 0.06, color=TEAL,
transform=ax.transAxes))
ax.text(0.5, 0.935,
"Аналитический отчёт по претензионной\nдеятельности контакт-центра",
ha="center", va="center", color="white", fontsize=13,
fontweight="bold", transform=ax.transAxes)
meta_lines = [f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}"]
if report_author:
meta_lines.append(f"Автор: {report_author}")
if report_org:
meta_lines.append(f"Организация: {report_org}")
if report_supervisor:
meta_lines.append(f"Научный руководитель: {report_supervisor}")
ax.text(0.06, 0.875, "\n".join(meta_lines), ha="left", va="top",
fontsize=10, color="#1F2D3D", transform=ax.transAxes)
ax.text(0.06, 0.79, "Ключевые показатели", fontsize=13, fontweight="bold",
color=DEEP_BLUE, transform=ax.transAxes)
kpis = [
("Всего претензий", f"{total_claims}"),
("Доля обоснованных", f"{valid_share:.1f}%"),
("Доля необоснованных", f"{invalid_share:.1f}%"),
("Индекс качества", f"{quality_index:.1f}/100"),
("Индекс риска", f"{risk_index:.1f}/100"),
("Уровень риска",
"Высокий" if risk_index >= 70 else "Средний" if risk_index >= 40 else "Низкий"),
]
x0, y0, w, h, gx, gy = 0.06, 0.69, 0.275, 0.07, 0.0275, 0.022
for i, (label, val) in enumerate(kpis):
r, c = divmod(i, 3)
cx = x0 + c * (w + gx)
cy = y0 - r * (h + gy)
ax.add_patch(plt.Rectangle((cx, cy), w, h, facecolor="#F2FAFB",
edgecolor=ACCENT, lw=1.2, transform=ax.transAxes))
ax.text(cx + 0.012, cy + h - 0.016, label, fontsize=8, color="#5B6B7A",
transform=ax.transAxes, va="top")
ax.text(cx + 0.012, cy + 0.012, val, fontsize=14, color=TEAL,
fontweight="bold", transform=ax.transAxes, va="bottom")
rec_top = 0.55
ax.text(0.06, rec_top, "Управленческие рекомендации", fontsize=13,
fontweight="bold", color=DEEP_BLUE, transform=ax.transAxes)
yy = rec_top - 0.035
for rec in recommendations:
wrapped = textwrap.wrap(rec, width=90)
if not wrapped:
continue
block = "• " + "\n ".join(wrapped)
ax.text(0.06, yy, block, fontsize=9.5, color="#1F2D3D", va="top",
transform=ax.transAxes)
yy -= 0.022 * len(wrapped) + 0.013
if yy < 0.06:
break
ax.text(0.5, 0.03,
"Сформировано интеллектуальной аналитической системой контроля качества конта
ha="center", fontsize=7.5, color="#9AA7B2", transform=ax.transAxes)
pdf.savefig(fig)
plt.close(fig)
# ----- Таблица операторов группы риска -----
if operator_stats is not None and len(operator_stats) > 0:
fig = plt.figure(figsize=(8.27, 11.69))
fig.patch.set_facecolor("white")
head = fig.add_axes([0, 0, 1, 1])
head.axis("off")
head.text(0.06, 0.95, "Операторы: качество и риск (топ-15)",
fontsize=13, fontweight="bold", color=DEEP_BLUE)
tbl = operator_stats.head(15)[[
"Количество
претензий", "Подтвержденные
ошибки",
"Доля
_
_
_
обоснованных_%", "Индекс
риска", "Индекс
_
_
качества", "Риск"
]].copy()
names = [shorten_text(i, 24) for i in tbl.index]
cell_text = [
[names[j]] + [str(v) for v in tbl.iloc[j].tolist()]
for j in range(len(tbl))
]
table_ax = fig.add_axes([0.04, 0.08, 0.92, 0.83])
table_ax.axis("off")
t = table_ax.table(
cellText=cell_text,
colLabels=["Оператор", "Претензий", "Ошибки", "Обосн.,%",
"Риск", "Качество", "Уровень"],
colWidths=[0.30, 0.11, 0.10, 0.11, 0.10, 0.11, 0.17],
loc="upper center", cellLoc="center"
)
t.auto_set_font_size(False)
t.set_fontsize(7.5)
t.scale(1, 1.6)
for (r, c), cell in t.get_celld().items():
if c == 0 and r > 0:
cell.set_text_props(ha="left")
if r == 0:
cell.set_text_props(fontweight="bold", color="white")
cell.set_facecolor(DEEP_BLUE)
pdf.savefig(fig)
plt.close(fig)
# ----- Все построенные диаграммы -----
for _title, f in report_figures:
try:
pdf.savefig(f, bbox_inches="tight")
except Exception:
pass
buf.seek(0)
return buf.getvalue()
st.markdown("### PDF-отчёт с диаграммами")
st.caption(
"В PDF включаются ключевые показатели, управленческие рекомендации, таблица операторов "
"группы риска и все построенные диаграммы с учётом выбранных фильтров."
)
if st.button("Сформировать PDF-отчёт"):
with st.spinner("Формирую PDF-отчёт..."):
st.session_state["pdf_bytes"] = build_pdf_report()
if st.session_state.get("pdf_bytes"):
st.download_button(
"Скачать PDF-отчёт",
data=st.session_state["pdf_bytes"],
file_name=f"contact_center_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
mime="application/pdf"
)
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
<p><b>Индекс риска</b> учитывает количество претензий, долю обоснованных жалоб и подтвержденн
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