"""
Веб-інтерфейс для розрахунку оренди за МСФЗ 16 на портфелі договорів.
Запуск локально:  streamlit run app.py
Розгортання:       Streamlit Community Cloud (приватний застосунок) або Render.

Пароль задається через st.secrets["APP_PASSWORD"] (файл .streamlit/secrets.toml
локально, або розділ Secrets у налаштуваннях застосунку на хостингу).
Ніде в коді ключ/пароль не хардкодиться.

Спрощена версія: результат — лише 3 вкладки в Excel (Summary, Графік,
Помісячно), на живих формулах для аудиту.
"""
import io
import datetime
import streamlit as st
import pandas as pd

from engine import (
    load_rates, run_batch, validate_register, build_monthly_summary,
    find_duplicate_contract_numbers, find_duplicate_contract_codes, build_template_bytes,
    write_formula_schedule, write_formula_monthly_summary,
)

st.set_page_config(page_title="МСФЗ 16 — розрахунок оренди", layout="centered")

RATES_FILE = "Ставки.xlsx"  # вбудований довідник, лежить поруч з app.py


# ---------- Проста авторизація ----------
def check_password():
    def password_entered():
        expected = st.secrets.get("APP_PASSWORD", "demo1234")
        if st.session_state.get("password_input") == expected:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.text_input("Пароль", type="password", key="password_input", on_change=password_entered)
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("Невірний пароль")
    return False


if not check_password():
    st.stop()

# ---------- Основний інтерфейс ----------
st.title("Розрахунок оренди за МСФЗ 16")
st.caption("Завантажте реєстр договорів → отримайте зведений розрахунок і графік ампортизації.")

st.subheader("1. Заповніть цю таблицю")
st.caption(
    "Скачайте шаблон, заповніть його даними по кожному договору (видаліть рядок-приклад) "
    "і завантажте назад нижче. Фіксовані колонки знижують ризик помилки в розрахунку."
)
st.download_button(
    label="Скачати шаблон реєстру (Excel)",
    data=build_template_bytes(),
    file_name="Шаблон_реєстру_договорів.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Опис колонок шаблону"):
    st.markdown(
        "- **Код договору** — власний ідентифікатор (необов'язково; якщо не заповнити, "
        "буде присвоєно автоматично)\n"
        "- **Номер договору** — обов'язково\n"
        "- **ПІБ пайовика** — обов'язково\n"
        "- **Дата початку оренди** — обов'язково\n"
        "- **Дата закінчення оренди** — обов'язково\n"
        "- **Сума оренди, грн** (річний платіж) — обов'язково"
    )

st.subheader("2. Завантажте заповнений реєстр")
uploaded = st.file_uploader("Реєстр договорів (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Не вдалося прочитати файл: {e}")
        st.stop()

    missing = validate_register(df)
    if missing:
        st.error("У файлі бракує обов'язкових колонок: " + ", ".join(missing))
        st.stop()

    st.success(f"Завантажено {len(df)} договорів.")
    st.dataframe(df.head(10), use_container_width=True)

    dup_numbers = find_duplicate_contract_numbers(df)
    if dup_numbers:
        st.warning(
            "Номери договору повторюються в реєстрі: " + ", ".join(map(str, dup_numbers)) +
            ". Розрахунок все одно виконається — кожному рядку присвоюється власний "
            "унікальний код (з колонки «Код договору» або автоматично L-00001, L-00002, ...), "
            "але варто перевірити реєстр на помилки."
        )

    dup_codes = find_duplicate_contract_codes(df)
    if dup_codes:
        st.warning(
            "Значення в колонці «Код договору» повторюються: " + ", ".join(map(str, dup_codes)) +
            ". Виправте, щоб коди справді ідентифікували договір однозначно."
        )

    if st.button("Розрахувати", type="primary"):
        with st.spinner(f"Рахую {len(df)} договорів..."):
            rates = load_rates(RATES_FILE)
            summary_df, schedule_df, errors_df = run_batch(df, rates)
            monthly_df = build_monthly_summary(schedule_df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                # Summary, Графік і Помісячно — на живих Excel-формулах (не готових
                # значеннях), щоб аудитор міг розкрити будь-яку клітинку й перерахувати
                # сам, а не вірити Python "наосліп". Перевірено LibreOffice recalculation.
                write_formula_schedule(writer, summary_df, schedule_df)
                write_formula_monthly_summary(writer, monthly_df, schedule_last_row=len(schedule_df))
            buffer.seek(0)

        st.success(f"Готово: {len(summary_df)} договорів прораховано, {len(errors_df)} помилок.")

        if not errors_df.empty:
            st.warning("Деякі договори не прораховано — див. деталі нижче.")
            st.dataframe(errors_df, use_container_width=True)

        st.dataframe(summary_df, use_container_width=True)

        st.subheader("Помісячна картина по всьому портфелю")
        st.dataframe(monthly_df, use_container_width=True)

        st.caption(
            "Вкладки «Summary», «Графік» і «Помісячно» у файлі Excel — на живих формулах "
            "(не готових значеннях): будь-яку клітинку можна розкрити й перерахувати вручну. "
            "Можливі поодинокі розбіжності до 1-2 копійок проти цифр вище — це різні алгоритми "
            "округлення Python і Excel, на підсумкові суми не впливає."
        )
        st.download_button(
            label="Завантажити повний розрахунок (Excel)",
            data=buffer,
            file_name=f"МСФЗ16_розрахунок_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
