import streamlit as st
import pandas as pd
import numpy as np
import io

# -------------------------------------------------
# НАСТРОЙКА СТРАНИЦЫ
# -------------------------------------------------

st.set_page_config(layout="wide", page_title="Аналитика")
st.title("📊 Аналитика и сравнение")

# -------------------------------------------------
# ФУНКЦИЯ ЗАГРУЗКИ
# -------------------------------------------------

@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

# -------------------------------------------------
# ЗАГРУЗКА ФАЙЛОВ
# -------------------------------------------------

st.sidebar.header("📂 Загрузка файлов")

f_curr = st.sidebar.file_uploader("Текущий период (Новый)", type=["xlsx"], key="main_curr")
f_old = st.sidebar.file_uploader("Прошлый период (Старый)", type=["xlsx"], key="main_old")

if not f_curr:
    st.info("👈 Загрузите основной файл.")
    st.stop()

df_c = load_excel(f_curr)
df_o = load_excel(f_old) if f_old else None

# -------------------------------------------------
# ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ ДЛЯ ОСНОВНЫХ СТОЛБЦОВ
# -------------------------------------------------

if "master_checkbox" not in st.session_state:
    st.session_state.master_checkbox = True

for col in df_c.columns:
    if f"col_{col}" not in st.session_state:
        st.session_state[f"col_{col}"] = True

def toggle_all():
    for col in df_c.columns:
        st.session_state[f"col_{col}"] = st.session_state.master_checkbox

# -------------------------------------------------
# 1. СТОЛБЦЫ
# -------------------------------------------------

st.sidebar.header("⚙️ 1. Столбцы")

st.sidebar.checkbox("☑️ Выбрать все", key="master_checkbox", on_change=toggle_all)

selected_columns = []

with st.sidebar.container(height=300):
    for col in df_c.columns:
        if st.checkbox(col, key=f"col_{col}"):
            selected_columns.append(col)

# -------------------------------------------------
# 2. ФИЛЬТРАЦИЯ
# -------------------------------------------------

st.sidebar.header("🔍 2. Фильтрация")

filtered_c = df_c[selected_columns].copy() if selected_columns else df_c.copy()
filtered_o = (
    df_o[selected_columns].copy()
    if (df_o is not None and selected_columns)
    else (df_o.copy() if df_o is not None else None)
)

if selected_columns:
    for col in selected_columns:
        v_c = filtered_c[col].dropna().unique().tolist()
        v_o = filtered_o[col].dropna().unique().tolist() if filtered_o is not None else []

        all_vals = sorted(list(set(map(str, v_c)) | set(map(str, v_o))))

        selected_vals = st.sidebar.multiselect(
            f"📁 {col}",
            all_vals,
            key=f"filter_{col}"
        )

        if selected_vals:
            filtered_c = filtered_c[filtered_c[col].astype(str).isin(selected_vals)]
            if filtered_o is not None:
                filtered_o = filtered_o[filtered_o[col].astype(str).isin(selected_vals)]

# -------------------------------------------------
# ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ ДЛЯ СТОЛБЦОВ ПО ТЕМАМ
# -------------------------------------------------

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 3. Столбцы по темам")

if "master_checkbox_theme" not in st.session_state:
    st.session_state.master_checkbox_theme = True

for col in df_c.columns:
    if f"theme_col_{col}" not in st.session_state:
        st.session_state[f"theme_col_{col}"] = True

def toggle_all_theme():
    for col in df_c.columns:
        st.session_state[f"theme_col_{col}"] = st.session_state.master_checkbox_theme

st.sidebar.checkbox(
    "☑️ Выбрать все (темы)",
    key="master_checkbox_theme",
    on_change=toggle_all_theme
)

selected_columns_theme = []

with st.sidebar.container(height=300):
    for col in df_c.columns:
        if st.checkbox(col, key=f"theme_col_{col}"):
            selected_columns_theme.append(col)

# -------------------------------------------------
# 4. ФИЛЬТРАЦИЯ ПО ТЕМАМ
# -------------------------------------------------

st.sidebar.header("🔍 4. Фильтрация по темам")

filtered_c_theme = df_c[selected_columns_theme].copy() if selected_columns_theme else df_c.copy()
filtered_o_theme = (
    df_o[selected_columns_theme].copy()
    if (df_o is not None and selected_columns_theme)
    else (df_o.copy() if df_o is not None else None)
)

if selected_columns_theme:
    for col in selected_columns_theme:
        v_c = filtered_c_theme[col].dropna().unique().tolist()
        v_o = filtered_o_theme[col].dropna().unique().tolist() if filtered_o_theme is not None else []

        all_vals = sorted(list(set(map(str, v_c)) | set(map(str, v_o))))

        selected_vals = st.sidebar.multiselect(
            f"📁 {col} (темы)",
            all_vals,
            key=f"filter_theme_{col}"
        )

        if selected_vals:
            filtered_c_theme = filtered_c_theme[filtered_c_theme[col].astype(str).isin(selected_vals)]
            if filtered_o_theme is not None:
                filtered_o_theme = filtered_o_theme[filtered_o_theme[col].astype(str).isin(selected_vals)]

# -------------------------------------------------
# ОБЩИЕ ФУНКЦИИ
# -------------------------------------------------

def calc_dyn(old, curr):
    if old == 0:
        return 100.0 if curr > 0 else 0.0
    return ((curr - old) / old) * 100

def color_dyn(val):
    if val > 0:
        return "color: green"
    elif val < 0:
        return "color: red"
    return "color: gray"

def style_dyn(series):
    return [color_dyn(v) for v in series]

# -------------------------------------------------
# ПУСТЫЕ ОБЪЕКТЫ ДЛЯ ОТЧЕТА
# -------------------------------------------------

p_df = pd.DataFrame()
p_df_theme = pd.DataFrame()
res = pd.DataFrame()
res_theme = pd.DataFrame()

# -------------------------------------------------
# ВКЛАДКИ
# -------------------------------------------------

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Аналитика",
    "📊 Аналитика по темам",
    "🔄 Сравнение",
    "🔄 Сравнение по темам",
    "📈 Графики"
])

# -------------------------------------------------
# АНАЛИТИКА
# -------------------------------------------------

with tab1:
    st.write(f"**Строк в текущем периоде: {len(filtered_c)}**")

    st.dataframe(filtered_c, use_container_width=True, height=400)

    st.markdown("---")

    if selected_columns:
        st.header("📊 Сводная таблица")
        c1, c2 = st.columns(2)

        with c1:
            p_row = st.selectbox("Что в СТРОКИ?", options=selected_columns, key="p1_main")

        with c2:
            p_col = st.selectbox("Что в КОЛОНКИ?", options=["Нет"] + selected_columns, key="p2_main")

        if p_col != "Нет":
            p_df = pd.crosstab(
                filtered_c[p_row],
                filtered_c[p_col],
                margins=True,
                margins_name="ИТОГО"
            )
        else:
            p_df = filtered_c[p_row].value_counts().reset_index()
            p_df.columns = [p_row, "Количество"]
            p_df.loc[len(p_df)] = ["ИТОГО", p_df["Количество"].sum()]
            p_df = p_df.set_index(p_row)

        st.dataframe(p_df, use_container_width=True)

        st.markdown("---")
        cd1, cd2 = st.columns(2)

        with cd1:
            buf1 = io.BytesIO()
            with pd.ExcelWriter(buf1, engine="openpyxl") as w:
                filtered_c.to_excel(w, index=False)

            st.download_button(
                "📄 Скачать подробные данные",
                buf1.getvalue(),
                "Details.xlsx"
            )

        with cd2:
            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine="openpyxl") as w:
                p_df.to_excel(w)

            st.download_button(
                "📊 Скачать сводную таблицу",
                buf2.getvalue(),
                "Summary.xlsx"
            )

# -------------------------------------------------
# АНАЛИТИКА ПО ТЕМАМ
# -------------------------------------------------

with tab2:
    st.write(f"**Строк в текущем периоде: {len(filtered_c_theme)}**")

    st.dataframe(filtered_c_theme, use_container_width=True, height=400)

    st.markdown("---")

    if selected_columns_theme:
        st.header("📊 Сводная таблица")
        c1, c2 = st.columns(2)

        with c1:
            p_row_t = st.selectbox("Что в СТРОКИ?", options=selected_columns_theme, key="p1_theme")

        with c2:
            p_col_t = st.selectbox("Что в КОЛОНКИ?", options=["Нет"] + selected_columns_theme, key="p2_theme")

        if p_col_t != "Нет":
            p_df_theme = pd.crosstab(
                filtered_c_theme[p_row_t],
                filtered_c_theme[p_col_t],
                margins=True,
                margins_name="ИТОГО"
            )
        else:
            p_df_theme = filtered_c_theme[p_row_t].value_counts().reset_index()
            p_df_theme.columns = [p_row_t, "Количество"]
            p_df_theme.loc[len(p_df_theme)] = ["ИТОГО", p_df_theme["Количество"].sum()]
            p_df_theme = p_df_theme.set_index(p_row_t)

        st.dataframe(p_df_theme, use_container_width=True)

        st.markdown("---")
        cd1t, cd2t = st.columns(2)

        with cd1t:
            buf1t = io.BytesIO()
            with pd.ExcelWriter(buf1t, engine="openpyxl") as w:
                filtered_c_theme.to_excel(w, index=False)

            st.download_button(
                "📄 Скачать подробные данные",
                buf1t.getvalue(),
                "Details_themes.xlsx"
            )

        with cd2t:
            buf2t = io.BytesIO()
            with pd.ExcelWriter(buf2t, engine="openpyxl") as w:
                p_df_theme.to_excel(w)

            st.download_button(
                "📊 Скачать сводную таблицу",
                buf2t.getvalue(),
                "Summary_themes.xlsx"
            )

# -------------------------------------------------
# СРАВНЕНИЕ
# -------------------------------------------------

with tab3:
    if f_old:
        common = [c for c in selected_columns if c in df_o.columns]

        if common:
            st.header("🔄 Сравнение динамики")

            comp_col = st.selectbox(
                "Параметр сравнения",
                options=common,
                key="compare_main"
            )

            n_val = filtered_c[comp_col].value_counts()
            o_val = filtered_o[comp_col].value_counts()

            res = pd.DataFrame({
                "Прошлый (шт)": o_val,
                "Текущий (шт)": n_val
            }).fillna(0)

            res.loc["ИТОГО"] = res.sum()

            res["Динамика %"] = res.apply(
                lambda r: calc_dyn(r["Прошлый (шт)"], r["Текущий (шт)"]),
                axis=1
            )

            st.dataframe(
                res.style
                    .format({
                        "Прошлый (шт)": "{:.0f}",
                        "Текущий (шт)": "{:.0f}",
                        "Динамика %": "{:+.1f}%"
                    })
                    .apply(style_dyn, subset=["Динамика %"]),
                use_container_width=True
            )

            buf_comp = io.BytesIO()
            with pd.ExcelWriter(buf_comp, engine="openpyxl") as w:
                res.to_excel(w)

            st.download_button(
                "📥 Скачать отчет о сравнении",
                buf_comp.getvalue(),
                "Comparison.xlsx"
            )
        else:
            st.info("Нет общих столбцов для сравнения.")
    else:
        st.warning("👈 Загрузите второй файл ('Прошлый период') слева для сравнения динамики.")

# -------------------------------------------------
# СРАВНЕНИЕ ПО ТЕМАМ
# -------------------------------------------------

with tab4:
    if f_old:
        common_theme = [c for c in selected_columns_theme if c in df_o.columns]

        if common_theme:
            st.header("🔄 Сравнение динамики")

            comp_col_t = st.selectbox(
                "Параметр сравнения",
                options=common_theme,
                key="compare_theme"
            )

            n_val_t = filtered_c_theme[comp_col_t].value_counts()
            o_val_t = filtered_o_theme[comp_col_t].value_counts()

            res_theme = pd.DataFrame({
                "Прошлый (шт)": o_val_t,
                "Текущий (шт)": n_val_t
            }).fillna(0)

            res_theme.loc["ИТОГО"] = res_theme.sum()

            res_theme["Динамика %"] = res_theme.apply(
                lambda r: calc_dyn(r["Прошлый (шт)"], r["Текущий (шт)"]),
                axis=1
            )

            st.dataframe(
                res_theme.style
                    .format({
                        "Прошлый (шт)": "{:.0f}",
                        "Текущий (шт)": "{:.0f}",
                        "Динамика %": "{:+.1f}%"
                    })
                    .apply(style_dyn, subset=["Динамика %"]),
                use_container_width=True
            )

            buf_comp_theme = io.BytesIO()
            with pd.ExcelWriter(buf_comp_theme, engine="openpyxl") as w:
                res_theme.to_excel(w)

            st.download_button(
                "📥 Скачать отчет о сравнении",
                buf_comp_theme.getvalue(),
                "Comparison_themes.xlsx"
            )
        else:
            st.info("Нет общих столбцов для сравнения по темам.")
    else:
        st.warning("👈 Загрузите второй файл ('Прошлый период') слева для сравнения динамики.")

# -------------------------------------------------
# ГРАФИКИ + ОДИН EXCEL-ОТЧЕТ
# -------------------------------------------------

with tab5:
    st.header("📈 Графики")

    if selected_columns:
        chart_col = st.selectbox("Показатель для графика", options=selected_columns, key="chart_main")
        chart_data = filtered_c[chart_col].value_counts()
        st.bar_chart(chart_data)
    else:
        st.info("Выберите хотя бы один столбец для графика.")

    st.markdown("---")
    st.header("📥 Один Excel-отчет")

    buffer_all = io.BytesIO()

    with pd.ExcelWriter(buffer_all, engine="openpyxl") as writer:
        # 4 листа — по одному на каждый этаж
        p_df.to_excel(writer, sheet_name="Аналитика")
        p_df_theme.to_excel(writer, sheet_name="Аналитика_темы")
        res.to_excel(writer, sheet_name="Сравнение")
        res_theme.to_excel(writer, sheet_name="Сравнение_темы")

    st.download_button(
        "📥 Один Excel-отчет",
        buffer_all.getvalue(),
        "analytics_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
