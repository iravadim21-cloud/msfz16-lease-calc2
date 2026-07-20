"""
Веб-інтерфейс для розрахунку оренди за МСФЗ 16 на портфелі договорів.
Запуск локально:  streamlit run app.py
Розгортання:       Streamlit Community Cloud (приватний застосунок) або Render.

Пароль задається через st.secrets["APP_PASSWORD"] (файл .streamlit/secrets.toml
локально, або розділ Secrets у налаштуваннях застосунку на хостингу).
Ніде в коді ключ/пароль не хардкодиться.
"""
import io
import streamlit as st
import pandas as pd

from engine import load_rates, run_batch, validate_register, build_monthly_summary

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

with st.expander("Формат реєстру (обов'язкові колонки)"):
    st.markdown(
        "- **Номер договору**\n"
        "- **ПІБ пайовика**\n"
        "- **Дата початку оренди**\n"
        "- **Дата закінчення оренди**\n"
        "- **Сума оренди, грн** (річний платіж)"
    )

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

    if st.button("Розрахувати", type="primary"):
        with st.spinner(f"Рахую {len(df)} договорів..."):
            rates = load_rates(RATES_FILE)
            summary_df, schedule_df, errors_df = run_batch(df, rates)
            monthly_df = build_monthly_summary(schedule_df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                schedule_df.to_excel(writer, sheet_name="Графік", index=False)
                monthly_df.to_excel(writer, sheet_name="Помісячно", index=False)
                if not errors_df.empty:
                    errors_df.to_excel(writer, sheet_name="Помилки", index=False)
            buffer.seek(0)

        st.success(f"Готово: {len(summary_df)} договорів прораховано, {len(errors_df)} помилок.")

        # Кнопка завантаження йде одразу після успіху — незалежно від того,
        # чи відмалюються прев'ю-таблиці нижче без збоїв.
        st.download_button(
            label="Скачати результат (Excel)",
            data=buffer,
            file_name="Розрахунок_МСФЗ16.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if not errors_df.empty:
            st.warning("Деякі договори не прораховано — див. деталі нижче та аркуш «Помилки» у файлі.")
            try:
                st.dataframe(errors_df, use_container_width=True)
            except Exception as e:
                st.info(f"(Прев'ю помилок не відмалювалось: {e}. Дивіться аркуш «Помилки» у файлі.)")

        st.subheader("Зведення по договорах (Summary)")
        try:
            st.dataframe(summary_df, use_container_width=True)
        except Exception as e:
            st.info(f"(Прев'ю Summary не відмалювалось: {e}. Дивіться аркуш «Summary» у файлі.)")

        st.subheader("Помісячна картина по всьому портфелю")
        try:
            # свіжий, "чистий" DataFrame — обходить рідкісні збої серіалізації
            # прев'ю-таблиці після groupby на деяких версіях pyarrow
            monthly_preview = pd.DataFrame(monthly_df.to_dict("records"))
            st.dataframe(monthly_preview, use_container_width=True)
        except Exception as e:
            st.info(f"(Прев'ю «Помісячно» не відмалювалось: {e}. Дивіться аркуш «Помісячно» у скачаному файлі — там дані повні.)")
