import io
import math
import os
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pulp
import streamlit as st

# ==========================================
# 1. ГЛОБАЛЬНІ НАЛАШТУВАННЯ ТА БАЗА ДАНИХ
# ==========================================
st.set_page_config(page_title="Генератор шкільного розкладу", layout="wide")

DAYS = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', "П'ятниця"]
SLOTS = [1, 2, 3, 4, 5, 6, 7, 8]

LESSONS_FILE = "saved_lessons.csv"
RULES_FILE = "saved_rules.csv"

def load_data():
    if os.path.exists(LESSONS_FILE):
        st.session_state['lessons_db'] = pd.read_csv(LESSONS_FILE)
    else:
        st.session_state['lessons_db'] = pd.DataFrame(columns=["Клас", "Предмет", "Вчитель", "Кількість годин", "Тиждень"])
        
    if os.path.exists(RULES_FILE):
        st.session_state['rules_db'] = pd.read_csv(RULES_FILE)
    else:
        st.session_state['rules_db'] = pd.DataFrame(columns=["Тип заборони", "Об'єкт (Назва)", "День тижня", "Номер уроку"])

def save_data():
    st.session_state['lessons_db'].to_csv(LESSONS_FILE, index=False)
    st.session_state['rules_db'].to_csv(RULES_FILE, index=False)

if 'lessons_db' not in st.session_state or 'rules_db' not in st.session_state:
    load_data()

# ==========================================
# 2. РОБОТА З EXCEL (ІМПОРТ ТА ШАБЛОН)
# ==========================================
def create_excel_template():
    """Створює готову матрицю шаблону input_data.xlsx у пам'яті"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_lessons_sample = pd.DataFrame([
            {"Клас": "10-А", "Предмет": "Математика", "Вчитель": "Іванов І.І.", "Кількість годин": 4, "Тиждень": "Кожен тиждень"},
            {"Клас": "10-А", "Предмет": "Фізика", "Вчитель": "Петров П.П.", "Кількість годин": 2, "Тиждень": "Кожен тиждень"},
            {"Клас": "11-Б", "Предмет": "Хімія", "Вчитель": "Сидорова С.С.", "Кількість годин": 1, "Тиждень": "Чисельник (Тиждень 1)"},
        ])
        df_lessons_sample.to_excel(writer, sheet_name="Lessons", index=False)

        df_teacher_rules_sample = pd.DataFrame([
            {"Вчитель": "Іванов І.І.", "День тижня": "П'ятниця", "Номер уроку": "8"},
            {"Вчитель": "Петров П.П.", "День тижня": "Понеділок", "Номер уроку": "Весь день"}
        ])
        df_teacher_rules_sample.to_excel(writer, sheet_name="Teacher_Rules", index=False)

        df_school_rules_sample = pd.DataFrame([
            {"Тип заборони": "Вся школа", "Об'єкт (Назва)": "Усі", "День тижня": "П'ятниця", "Номер уроку": "8"},
            {"Тип заборони": "Клас", "Об'єкт (Назва)": "10-А", "День тижня": "Понеділок", "Номер уроку": "1"}
        ])
        df_school_rules_sample.to_excel(writer, sheet_name="School_Rules", index=False)

    output.seek(0)
    return output


def process_uploaded_excel(uploaded_file):
    """Парсить завантажений Excel-файл і конвертує його у внутрішній формат бази"""
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        if "Lessons" in xls.sheet_names:
            df_lessons = pd.read_excel(xls, sheet_name="Lessons")
            if "Тиждень" not in df_lessons.columns:
                df_lessons["Тиждень"] = "Кожен тиждень"
            df_lessons["Тиждень"] = df_lessons["Тиждень"].fillna("Кожен тиждень")
            st.session_state['lessons_db'] = df_lessons[["Клас", "Предмет", "Вчитель", "Кількість годин", "Тиждень"]].dropna(subset=["Клас", "Предмет", "Вчитель"])
        
        new_rules = []

        if "Teacher_Rules" in xls.sheet_names:
            df_t_rules = pd.read_excel(xls, sheet_name="Teacher_Rules").dropna(subset=["Вчитель"])
            for _, row in df_t_rules.iterrows():
                new_rules.append({
                    "Тип заборони": "Вчитель",
                    "Об'єкт (Назва)": str(row["Вчитель"]).strip(),
                    "День тижня": str(row["День тижня"]).strip(),
                    "Номер уроку": str(row["Номер уроку"]).strip()
                })

        if "School_Rules" in xls.sheet_names:
            df_s_rules = pd.read_excel(xls, sheet_name="School_Rules")
            for _, row in df_s_rules.iterrows():
                r_type = row.get("Тип заборони") or row.get("Тип правила", "Вся школа")
                r_obj = row.get("Об'єкт (Назва)") or row.get("Предмет", "Усі")
                
                if pd.notna(row.get("День тижня")) and pd.notna(row.get("Номер уроку")):
                    new_rules.append({
                        "Тип заборони": str(r_type).strip(),
                        "Об'єкт (Назва)": str(r_obj).strip(),
                        "День тижня": str(row["День тижня"]).strip(),
                        "Номер уроку": str(row["Номер уроку"]).strip()
                    })

        if new_rules:
            st.session_state['rules_db'] = pd.DataFrame(new_rules)

        save_data()
        st.success("✅ Дані з Excel успішно імпортовано та збережено!")
        st.rerun()

    except Exception as e:
        st.error(f"❌ Помилка читання Excel-файлу: {e}")

# ==========================================
# 3. РОЗШИРЕНЕ МАТЕМАТИЧНЕ ЯДРО (PuLP)
# ==========================================
def solve_schedule(df_lessons, df_rules, opts):
    lessons = []
    for idx, row in df_lessons.iterrows():
        c_id = str(row["Клас"]).strip()
        subj = str(row["Предмет"]).strip()
        t = str(row["Вчитель"]).strip()
        try:
            count = int(row["Кількість годин"])
        except (ValueError, TypeError):
            count = 1
        w_type = str(row.get("Тиждень", "Кожен тиждень"))
        
        for i in range(count):
            lessons.append((c_id, subj, t, f"{subj}_{idx}_{i}", w_type))
            
    classes_list = sorted(df_lessons["Клас"].dropna().astype(str).unique().tolist())
    teachers_list = sorted(df_lessons["Вчитель"].dropna().astype(str).unique().tolist())

    prob = pulp.LpProblem("School_Schedule_Optimization_Advanced", pulp.LpMinimize)
    schedule_vars = {}

    for c_id, subj, t, l_id, wt in lessons:
        for d in DAYS:
            for s in SLOTS:
                var_name = f"g_{c_id}_{l_id}_{d}_{s}".replace(" ", "_").replace("'", "_")
                schedule_vars[(c_id, subj, t, l_id, wt, d, s)] = pulp.LpVariable(var_name, cat='Binary')

    # БАЗОВІ ОБМЕЖЕННЯ
    # 1. Кожен урок має відбутися 1 раз на тиждень
    for c_id, subj, t, l_id, wt in lessons:
        prob += pulp.lpSum(schedule_vars[(c_id, subj, t, l_id, wt, d, s)] for d in DAYS for s in SLOTS) == 1

    # 2. Не більше одного уроку для класу одночасно
    for c_id in classes_list:
        for d in DAYS:
            for s in SLOTS:
                w1 = [schedule_vars[(c, subj, t, l_id, wt, d, s)] for c, subj, t, l_id, wt in lessons if c == c_id and wt in ["Кожен тиждень", "Чисельник (Тиждень 1)"]]
                prob += pulp.lpSum(w1) <= 1
                w2 = [schedule_vars[(c, subj, t, l_id, wt, d, s)] for c, subj, t, l_id, wt in lessons if c == c_id and wt in ["Кожен тиждень", "Знаменник (Тиждень 2)"]]
                prob += pulp.lpSum(w2) <= 1

    # 3. Вчитель не може бути в двох класах одночасно
    for t_name in teachers_list:
        for d in DAYS:
            for s in SLOTS:
                w1_t = [schedule_vars[(c, subj, t, l_id, wt, d, s)] for c, subj, t, l_id, wt in lessons if t == t_name and wt in ["Кожен тиждень", "Чисельник (Тиждень 1)"]]
                prob += pulp.lpSum(w1_t) <= 1
                w2_t = [schedule_vars[(c, subj, t, l_id, wt, d, s)] for c, subj, t, l_id, wt in lessons if t == t_name and wt in ["Кожен тиждень", "Знаменник (Тиждень 2)"]]
                prob += pulp.lpSum(w2_t) <= 1

    # 4. Заборони з системи правил
    for _, row in df_rules.dropna(subset=["Тип заборони", "День тижня", "Номер уроку"]).iterrows():
        r_type = str(row["Тип заборони"]).strip()
        r_obj = str(row["Об'єкт (Назва)"]).strip()
        r_day = str(row["День тижня"]).strip()
        r_slot_val = str(row["Номер уроку"]).strip()
        
        slots_to_ban = SLOTS if r_slot_val.lower() == "весь день" else [int(r_slot_val)]
        
        for s in slots_to_ban:
            if r_type == "Вчитель":
                vars_to_ban = [schedule_vars[(c, subj, t, l_id, wt, r_day, s)] for c, subj, t, l_id, wt in lessons if t == r_obj]
            elif r_type == "Клас":
                vars_to_ban = [schedule_vars[(c, subj, t, l_id, wt, r_day, s)] for c, subj, t, l_id, wt in lessons if c == r_obj]
            elif r_type == "Вся школа":
                vars_to_ban = [schedule_vars[(c, subj, t, l_id, wt, r_day, s)] for c, subj, t, l_id, wt in lessons]
            else:
                vars_to_ban = []
                
            if vars_to_ban:
                prob += pulp.lpSum(vars_to_ban) == 0

    # 5. Без "вікон" у класах
    for c_id in classes_list:
        for d in DAYS:
            for i in range(1, len(SLOTS)):
                s_curr, s_prev = SLOTS[i], SLOTS[i-1]
                w1_curr = pulp.lpSum([schedule_vars[(c, subj, t, l_id, wt, d, s_curr)] for c, subj, t, l_id, wt in lessons if c == c_id and (wt == "Кожен тиждень" or wt == "Чисельник (Тиждень 1)")])
                w1_prev = pulp.lpSum([schedule_vars[(c, subj, t, l_id, wt, d, s_prev)] for c, subj, t, l_id, wt in lessons if c == c_id and (wt == "Кожен тиждень" or wt == "Чисельник (Тиждень 1)")])
                prob += w1_curr <= w1_prev
                
                w2_curr = pulp.lpSum([schedule_vars[(c, subj, t, l_id, wt, d, s_curr)] for c, subj, t, l_id, wt in lessons if c == c_id and (wt == "Кожен тиждень" or wt == "Знаменник (Тиждень 2)")])
                w2_prev = pulp.lpSum([schedule_vars[(c, subj, t, l_id, wt, d, s_prev)] for c, subj, t, l_id, wt in lessons if c == c_id and (wt == "Кожен тиждень" or wt == "Знаменник (Тиждень 2)")])
                prob += w2_curr <= w2_prev

    # 6. Спортзал (макс 1 фізкультура одночасно)
    for d in DAYS:
        for s in SLOTS:
            pe_vars = [schedule_vars[(c, subj, t, l_id, wt, d, s)] for c, subj, t, l_id, wt in lessons if subj.strip().lower() == "фізкультура"]
            if pe_vars: prob += pulp.lpSum(pe_vars) <= 1

    # ==========================================
    # НОВІ ДОДАТКОВІ ОБМЕЖЕННЯ (САНІТАРНІ ТА ВЧИТЕЛЬСЬКІ)
    # ==========================================

    # А. МАКСИМАЛЬНА КІЛЬКІСТЬ УРОКІВ НА ДЕНЬ ДЛЯ КЛАСУ
    if opts.get("max_daily_lessons"):
        max_daily = opts["max_daily_lessons"]
        for c_id in classes_list:
            for d in DAYS:
                prob += pulp.lpSum(schedule_vars[(c, subj, t, l_id, wt, d, s)]
                                   for c, subj, t, l_id, wt in lessons if c == c_id for s in SLOTS) <= max_daily

    # Б. ЗАБОРОНА СКЛАДНИХ ПРЕДМЕТІВ НА ОСТАННІХ УРОКАХ (7-8 уроки)
    if opts.get("ban_hard_on_late_slots"):
        hard_keywords = ["математика", "алгебра", "геометрія", "фізика", "хімія", "іноземна", "англійська", "інформатика"]
        for c, subj, t, l_id, wt in lessons:
            if any(kw in subj.lower() for kw in hard_keywords):
                for d in DAYS:
                    for s in [7, 8]:
                        prob += schedule_vars[(c, subj, t, l_id, wt, d, s)] == 0

    # В. НЕ БІЛЬШЕ 1 УРОКУ ОДНОГО ПРЕДМЕТА НА ДЕНЬ (якщо годин <= 5)
    if opts.get("max_one_subj_per_day"):
        for c_id in classes_list:
            class_subjs = set(subj for c, subj, t, l_id, wt in lessons if c == c_id)
            for subj_name in class_subjs:
                matching = [l for l in lessons if l[0] == c_id and l[1] == subj_name]
                if len(matching) <= 5:
                    for d in DAYS:
                        prob += pulp.lpSum(schedule_vars[(c, subj, t, l_id, wt, d, s)]
                                           for c, subj, t, l_id, wt in matching for s in SLOTS) <= 1

    # Г. ДОПОМІЖНІ ЗМІННІ ДЛЯ ВЧИТЕЛІВ
    teacher_active = {}
    for t_name in teachers_list:
        for d in DAYS:
            for s in SLOTS:
                var_t = pulp.LpVariable(f"t_act_{t_name}_{d}_{s}".replace(" ", "_").replace("'", "_"), cat='Binary')
                teacher_active[(t_name, d, s)] = var_t
                t_lessons = [schedule_vars[(c, subj, t, l_id, wt, d, s)] for c, subj, t, l_id, wt in lessons if t == t_name]
                if t_lessons:
                    prob += var_t == pulp.lpSum(t_lessons)
                else:
                    prob += var_t == 0

    # Д. МАКСИМУМ УРОКІВ ПОСПІЛЬ ДЛЯ ВЧИТЕЛЯ БЕЗ ПЕРЕРВИ
    if opts.get("max_consecutive_teacher"):
        max_c = opts["max_consecutive_teacher"]
        for t_name in teachers_list:
            for d in DAYS:
                for s in range(1, 9 - max_c):
                    prob += pulp.lpSum(teacher_active[(t_name, d, k)] for k in range(s, s + max_c + 1)) <= max_c

    # Е. МЕТОДИЧНИЙ ДЕНЬ ВЧИТЕЛЯ (1 повністю вільний день при навантаженні <= 20 год)
    if opts.get("teacher_method_day"):
        for t_name in teachers_list:
            t_total_hours = sum(1 for c, subj, t, l_id, wt in lessons if t == t_name)
            if t_total_hours <= 20:
                t_day_active = {}
                for d in DAYS:
                    d_var = pulp.LpVariable(f"t_day_act_{t_name}_{d}".replace(" ", "_").replace("'", "_"), cat='Binary')
                    t_day_active[d] = d_var
                    for s in SLOTS:
                        prob += teacher_active[(t_name, d, s)] <= d_var
                prob += pulp.lpSum(t_day_active[d] for d in DAYS) <= 4

    # Ж. РОЗПОДІЛ ПРЕДМЕТІВ З 2 ГОД/ТИЖДЕНЬ (не ставити в суміжні дні поспіль)
    if opts.get("spread_two_hour_subjs"):
        for c_id in classes_list:
            class_subjs = set(subj for c, subj, t, l_id, wt in lessons if c == c_id)
            for subj_name in class_subjs:
                matching = [l for l in lessons if l[0] == c_id and l[1] == subj_name]
                if len(matching) == 2:
                    c_day_var = {}
                    for d in DAYS:
                        cd_var = pulp.LpVariable(f"cd_act_{c_id}_{subj_name}_{d}".replace(" ", "_").replace("'", "_"), cat='Binary')
                        c_day_var[d] = cd_var
                        prob += pulp.lpSum(schedule_vars[(c, subj, t, l_id, wt, d, s)]
                                           for c, subj, t, l_id, wt in matching for s in SLOTS) <= cd_var
                    for i in range(len(DAYS) - 1):
                        d_curr, d_next = DAYS[i], DAYS[i+1]
                        prob += c_day_var[d_curr] + c_day_var[d_next] <= 1

    prob.solve(pulp.PULP_CBC_CMD(msg=False))
    return pulp.LpStatus[prob.status], schedule_vars, lessons, classes_list, teachers_list

# ==========================================
# 4. ГРАФІЧНИЙ ІНТЕРФЕЙС STREAMLIT
# ==========================================
st.title("🏫 Розумний шкільний розклад (Макс. 8 уроків)")
st.write("Автоматична оптимізація із урахуванням санітарних норм, зручності вчителів та матричним коригуванням.")

tab1, tab2, tab3 = st.tabs(["📋 Навантаження школи", "🚫 Керування заборонами", "🚀 Генерація та коригування"])

# --- ВКЛАДКА 1: НАВАНТАЖЕННЯ ---
with tab1:
    with st.expander("📁 Імпорт даних з Excel та завантаження шаблону", expanded=False):
        col_ex1, col_ex2 = st.columns([1, 1])
        with col_ex1:
            st.markdown("**1. Завантажте готовий шаблон Excel:**")
            template_bytes = create_excel_template()
            st.download_button(
                label="📥 Скачати шаблон (input_data.xlsx)",
                data=template_bytes,
                file_name="input_data_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with col_ex2:
            st.markdown("**2. Завантажте заповнений файл Excel:**")
            uploaded_file = st.file_uploader("Виберіть .xlsx файл", type=["xlsx"], label_visibility="collapsed")
            if uploaded_file is not None:
                if st.button("🔄 Зчитати та перезаписати базу", type="primary", use_container_width=True):
                    process_uploaded_excel(uploaded_file)
                    
    st.markdown("---")
    st.subheader("1. Швидке додавання предмета")
    with st.form("quick_add_form", clear_on_submit=True):
        f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns(5)
        c_id = f_col1.text_input("Клас (напр: 10-A):")
        subj = f_col2.text_input("Предмет (напр: Хімія):")
        teacher = f_col3.text_input("Вчитель (напр: Васильєва):")
        hours = f_col4.number_input("Годин на тиждень:", min_value=1, max_value=10, value=1)
        week_select = f_col5.selectbox("Періодичність:", ["Кожен тиждень", "Чисельник (Тиждень 1)", "Знаменник (Тиждень 2)"])
        
        if st.form_submit_button("➕ Додати у базу"):
            if c_id and subj and teacher:
                new_row = pd.DataFrame([{"Клас": c_id, "Предмет": subj, "Вчитель": teacher, "Кількість годин": hours, "Тиждень": week_select}])
                st.session_state['lessons_db'] = pd.concat([st.session_state['lessons_db'], new_row], ignore_index=True)
                save_data()
                st.rerun()

    st.subheader("2. Інтерактивна таблиця предметів")
    edited_lessons = st.data_editor(
        st.session_state['lessons_db'], 
        use_container_width=True, 
        num_rows="dynamic",
        column_config={
            "Тиждень": st.column_config.SelectboxColumn(options=["Кожен тиждень", "Чисельник (Тиждень 1)", "Знаменник (Тиждень 2)"])
        }
    )
    if not edited_lessons.equals(st.session_state['lessons_db']):
        st.session_state['lessons_db'] = edited_lessons
        save_data()
        st.toast("💾 Навантаження оновлено на диску!")

# --- ВКЛАДКА 2: КЕРУВАННЯ ЗАБОРОНАМИ ---
with tab2:
    st.subheader("1. Створити нове обмеження часу")
    active_classes = sorted(st.session_state['lessons_db']["Клас"].dropna().astype(str).unique().tolist())
    active_teachers = sorted(st.session_state['lessons_db']["Вчитель"].dropna().astype(str).unique().tolist())
    
    if not active_classes and not active_teachers:
        st.info("💡 Спочатку додайте предмети та вчителів на вкладці 'Навантаження школи', щоб налаштувати для них заборони.")
    else:
        with st.form("quick_rule_form", clear_on_submit=True):
            r_col1, r_col2, r_col3, r_col4 = st.columns(4)
            r_type = r_col1.selectbox("Кому забороняємо?", ["Вчитель", "Клас", "Вся школа"])
            
            if r_type == "Вчитель":
                r_obj = r_col2.selectbox("Виберіть вчителя:", active_teachers if active_teachers else ["Немає вчителів"])
            elif r_type == "Клас":
                r_obj = r_col2.selectbox("Виберіть клас:", active_classes if active_classes else ["Немає класів"])
            else:
                r_obj = r_col2.selectbox("Об'єкт:", ["Усі"])
                
            r_day = r_col3.selectbox("День тижня:", DAYS)
            r_slot = r_col4.selectbox("Який урок заблокувати?", ["Весь день", "1", "2", "3", "4", "5", "6", "7", "8"])
            
            if st.form_submit_button("➕ Зафіксувати заборону"):
                new_rule = pd.DataFrame([{"Тип заборони": r_type, "Об'єкт (Назва)": r_obj, "День тижня": r_day, "Номер уроку": r_slot}])
                st.session_state['rules_db'] = pd.concat([st.session_state['rules_db'], new_rule], ignore_index=True)
                save_data()
                st.rerun()

    st.subheader("2. Інтерактивна таблиця всіх заборон")
    edited_rules = st.data_editor(
        st.session_state['rules_db'], 
        use_container_width=True, 
        num_rows="dynamic",
        column_config={
            "Тип заборони": st.column_config.SelectboxColumn(options=["Вчитель", "Клас", "Вся школа"]),
            "День тижня": st.column_config.SelectboxColumn(options=DAYS),
            "Номер уроку": st.column_config.SelectboxColumn(options=["Весь день", "1", "2", "3", "4", "5", "6", "7", "8"])
        }
    )
    if not edited_rules.equals(st.session_state['rules_db']):
        st.session_state['rules_db'] = edited_rules
        save_data()
        st.toast("💾 Заборони автоматично синхронізовано на диску!")

# --- ВКЛАДКА 3: РОЗРАХУНОК ТА РУЧНЕ КОРИГУВАННЯ ПО ДНЯХ ---
with tab3:
    st.subheader("Генерація та ручне коригування розкладу")
    df_lessons_clean = st.session_state['lessons_db'].dropna(subset=["Клас", "Предмет", "Вчитель"])
    
    if df_lessons_clean.empty:
        st.info("Будь ласка, заповніть навантаження на першій вкладці.")
    else:
        with st.expander("⚙️ Додаткові санітарні та педагогічні правила", expanded=True):
            st.markdown("**Налаштуйте розумні обмеження для покращення якісних показників розкладу:**")
            opt_col1, opt_col2 = st.columns(2)
            
            with opt_col1:
                st.markdown("##### 🏫 Для учнів та класів")
                max_daily = st.number_input("Максимум уроків на день для класу:", min_value=4, max_value=8, value=7)
                ban_hard_late = st.checkbox("Заборонити складні предмети на 7-8 уроках (Математика, Фізика, Хімія)", value=True)
                max_one_subj = st.checkbox("Не більше 1 уроку одного предмета на день (якщо <= 5 год/тиждень)", value=True)
                spread_two_hour = st.checkbox("Розносити предмети з 2 год/тиждень (не в суміжні дні поспіль)", value=True)

            with opt_col2:
                st.markdown("##### 👨‍🏫 Для вчителів")
                max_consec = st.number_input("Максимум уроків поспіль для вчителя без перерви:", min_value=2, max_value=6, value=4)
                teacher_m_day = st.checkbox("Забезпечити 1 методичний (вільний) день при навантаженні <= 20 год", value=True)

        opts_dict = {
            "max_daily_lessons": max_daily,
            "ban_hard_on_late_slots": ban_hard_late,
            "max_one_subj_per_day": max_one_subj,
            "max_consecutive_teacher": max_consec,
            "teacher_method_day": teacher_m_day,
            "spread_two_hour_subjs": spread_two_hour
        }

        col_gen1, col_gen2 = st.columns([2, 1])
        with col_gen1:
            run_calc = st.button("🚀 ЗАПУСТИТИ АВТОМАТИЧНИЙ РОЗРАХУНОК", type="primary", use_container_width=True)
        with col_gen2:
            if 'generated_schedule' in st.session_state and not st.session_state['generated_schedule'].empty:
                if st.button("🔄 Скинути ручні правки", use_container_width=True):
                    del st.session_state['generated_schedule']
                    st.rerun()

        if run_calc:
            with st.spinner("Математична модель оптимізує розклад за усім комплексом правил..."):
                status, schedule_vars, lessons_data, classes_list, teachers_list = solve_schedule(df_lessons_clean, st.session_state['rules_db'], opts_dict)
                
                if status == "Optimal":
                    st.balloons()
                    st.success("🎉 Новий розклад успішно побудовано із урахуванням усіх санітарних норм!")
                    
                    weekly_rows = []
                    for day in DAYS:
                        for slot in SLOTS:
                            for class_id in classes_list:
                                w1_item, w2_item = "—", "—"
                                for c, subj, t, l_id, wt in lessons_data:
                                    if str(c) == str(class_id) and schedule_vars[(c, subj, t, l_id, wt, day, slot)].varValue is not None:
                                        if schedule_vars[(c, subj, t, l_id, wt, day, slot)].varValue > 0.9:
                                            if wt == "Кожен тиждень":
                                                w1_item = f"{subj} ({t})"
                                                w2_item = f"{subj} ({t})"
                                            elif wt == "Чисельник (Тиждень 1)":
                                                w1_item = f"🔼 {subj} ({t})"
                                            elif wt == "Знаменник (Тиждень 2)":
                                                w2_item = f"🔽 {subj} ({t})"
                                display_text = w1_item if w1_item == w2_item else f"{w1_item} / {w2_item}"
                                weekly_rows.append({
                                    "Клас": class_id, 
                                    "День тижня": day, 
                                    "Номер уроку": slot, 
                                    "Розклад (Чисельник / Знаменник)": display_text
                                })
                    
                    st.session_state['generated_schedule'] = pd.DataFrame(weekly_rows)
                else:
                    st.error(f"❌ Алгоритм зайшов у тупик (Статус: {status}). Спробуйте послабити деякі прапорці в налаштуваннях вище.")

        # ЯКЩО РОЗКЛАД ЗГЕНЕРОВАНО — ВІДОБРАЖАЄМО ТА ДОЗВОЛЯЄМО РЕДАГУВАТИ ПО ДНЯХ
        if 'generated_schedule' in st.session_state and not st.session_state['generated_schedule'].empty:
            df_sched = st.session_state['generated_schedule']
            
            st.markdown("---")
            st.subheader("✏️ Інструмент швидкого переміщення (Swap)")
            
            with st.expander("🔄 Обміняти два уроки місцями для класу", expanded=False):
                swap_col1, swap_col2, swap_col3 = st.columns(3)
                sel_class = swap_col1.selectbox("Виберіть клас:", sorted(df_sched["Клас"].unique()))
                
                with swap_col2:
                    st.caption("Урок №1 (Звідки):")
                    day1 = st.selectbox("День 1:", DAYS, key="d1")
                    slot1 = st.selectbox("Урок 1:", SLOTS, key="s1")
                
                with swap_col3:
                    st.caption("Урок №2 (Куди):")
                    day2 = st.selectbox("День 2:", DAYS, key="d2")
                    slot2 = st.selectbox("Урок 2:", SLOTS, key="s2")
                
                if st.button("🔄 Обміняти уроки", use_container_width=True):
                    idx1 = df_sched[(df_sched["Клас"] == sel_class) & (df_sched["День тижня"] == day1) & (df_sched["Номер уроку"] == slot1)].index
                    idx2 = df_sched[(df_sched["Клас"] == sel_class) & (df_sched["День тижня"] == day2) & (df_sched["Номер уроку"] == slot2)].index
                    
                    if not idx1.empty and not idx2.empty:
                        val1 = df_sched.loc[idx1[0], "Розклад (Чисельник / Знаменник)"]
                        val2 = df_sched.loc[idx2[0], "Розклад (Чисельник / Знаменник)"]
                        
                        df_sched.loc[idx1[0], "Розклад (Чисельник / Знаменник)"] = val2
                        df_sched.loc[idx2[0], "Розклад (Чисельник / Знаменник)"] = val1
                        
                        st.session_state['generated_schedule'] = df_sched
                        st.toast("✅ Уроки успішно поміняно місцями!")
                        st.rerun()

            # ДЕТЕКТОР КОНФЛІКТІВ ВЧИТЕЛІВ В РЕАЛЬНОМУ ЧАСІ
            teacher_conflicts = []
            temp_records = []
            for _, row in df_sched.iterrows():
                val = row["Розклад (Чисельник / Знаменник)"]
                if val != "—":
                    parts = val.split(" / ")
                    for p in parts:
                        if "(" in p and ")" in p:
                            t_name = p.split("(")[1].split(")")[0].strip()
                            temp_records.append({
                                "Вчитель": t_name,
                                "День": row["День тижня"],
                                "Слот": row["Номер уроку"],
                                "Клас": row["Клас"]
                            })
            
            if temp_records:
                check_df = pd.DataFrame(temp_records)
                duplicates = check_df[check_df.duplicated(subset=["Вчитель", "День", "Слот"], keep=False)]
                if not duplicates.empty:
                    for (t_name, d, s), group in duplicates.groupby(["Вчитель", "День", "Слот"]):
                        classes_str = ", ".join(group["Клас"].unique())
                        teacher_conflicts.append(f"⚠️ **{t_name}** має накладку в **{d}**, **{s}-й урок** у класах: {classes_str}")

            if teacher_conflicts:
                st.warning("Виявлено конфлікти після ручного редагування:")
                for conf in teacher_conflicts:
                    st.write(conf)

            # ВІДОБРАЖЕННЯ РОЗКЛАДУ ПО ДНЯХ Х КЛАСАХ (МАТРИЦЯ)
            st.subheader("📅 Розклад за днями тижня (Вся школа):")
            
            day_tabs = st.tabs([f"🗓️ {day}" for day in DAYS])
            
            for idx, day_name in enumerate(DAYS):
                with day_tabs[idx]:
                    df_day = df_sched[df_sched["День тижня"] == day_name]
                    
                    pivot_day = df_day.pivot(
                        index="Номер уроку", 
                        columns="Клас", 
                        values="Розклад (Чисельник / Знаменник)"
                    )
                    
                    st.caption("💡 Натисніть двічі на будь-яку комірку, щоб змінити предмет чи вчителя на цей день:")
                    
                    edited_pivot = st.data_editor(
                        pivot_day,
                        use_container_width=True,
                        key=f"editor_day_{day_name}"
                    )
                    
                    for class_col in edited_pivot.columns:
                        for slot_row in edited_pivot.index:
                            val = edited_pivot.loc[slot_row, class_col]
                            mask = (df_sched["День тижня"] == day_name) & (df_sched["Номер уроку"] == slot_row) & (df_sched["Клас"] == class_col)
                            df_sched.loc[mask, "Розклад (Чисельник / Знаменник)"] = val
            
            st.session_state['generated_schedule'] = df_sched

            # ГЕНЕРАЦІЯ ОФОРМЛЕНОГО EXCEL З УРАХУВАННЯМ ПРАВОК
            wb = Workbook()
            wb.remove(wb.active)
            font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            font_body = Font(name="Arial", size=10)
            fill_header = PatternFill(start_color="365F91", end_color="365F91", fill_type="solid")
            thin_border = Border(
                left=Side(style='thin', color='B0B0B0'), right=Side(style='thin', color='B0B0B0'),
                top=Side(style='thin', color='B0B0B0'), bottom=Side(style='thin', color='B0B0B0')
            )

            for class_id in sorted(df_sched["Клас"].unique()):
                ws = wb.create_sheet(title=f"Клас {class_id}")
                ws.sheet_view.showGridLines = True
                headers = ["День тижня", "Номер уроку", "Розклад (Чисельник / Знаменник)"]
                ws.append(headers)
                
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.font = font_header; cell.fill = fill_header; cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                class_data = df_sched[df_sched["Клас"] == class_id]
                for _, row_data in class_data.iterrows():
                    ws.append([row_data["День тижня"], f"Урок №{row_data['Номер уроку']}", row_data["Розклад (Чисельник / Знаменник)"]])
                    curr_row = ws.max_row
                    for col_num in range(1, 4):
                        c = ws.cell(row=curr_row, column=col_num)
                        c.font = font_body; c.border = thin_border
                        c.alignment = Alignment(horizontal="center" if col_num <= 2 else "left", vertical="center")

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
                ws.row_dimensions.height = 25

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)

            st.markdown("---")
            st.download_button(
                label="📥 СКАЧАТИ ВІДКОРИГОВАНИЙ РОЗКЛАД В EXCEL (.xlsx)", 
                data=excel_buffer, 
                file_name="school_schedule_edited.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True
            )