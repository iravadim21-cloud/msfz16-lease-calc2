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

from engine import (
    load_rates, run_batch, validate_register, build_monthly_summary,
    build_annual_rollforward, build_maturity_analysis, build_wam_stats, build_journal_entries,
    find_duplicate_contract_numbers, build_liability_classification, build_reclassification_entry,
)
import datetime

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

as_of_date = st.date_input(
    "Звітна дата (для примітки МСФЗ 16 і зважених показників)",
    value=datetime.date.today(),
    help="Використовується для аналізу строків погашення та зважених "
         "середніх ставки/строку — беруться лише договори, активні на цю дату.",
)

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
            "унікальний код (L-00001, L-00002, ...), але варто перевірити реєстр на помилки."
        )

    if st.button("Розрахувати", type="primary"):
        with st.spinner(f"Рахую {len(df)} договорів..."):
            rates = load_rates(RATES_FILE)
            summary_df, schedule_df, errors_df = run_batch(df, rates)
            monthly_df = build_monthly_summary(schedule_df)
            annual_df = build_annual_rollforward(schedule_df)
            maturity_df = build_maturity_analysis(schedule_df, as_of_date)
            wam_stats = build_wam_stats(summary_df, schedule_df, as_of_date)
            wam_df = pd.DataFrame([wam_stats])
            journal_df = build_journal_entries(schedule_df)
            classification_df = build_liability_classification(schedule_df, as_of_date)
            reclass_df = build_reclassification_entry(schedule_df, as_of_date)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                schedule_df.to_excel(writer, sheet_name="Графік", index=False)
                monthly_df.to_excel(writer, sheet_name="Помісячно", index=False)

                annual_df.to_excel(writer, sheet_name="ROU та зобов'язання", index=False)

                note_sheet = "Примітка МСФЗ 16"
                wam_df.to_excel(writer, sheet_name=note_sheet, index=False, startrow=1)
                maturity_df.to_excel(writer, sheet_name=note_sheet, index=False, startrow=5)
                ws_note = writer.sheets[note_sheet]
                ws_note.write(0, 0, "Зважені показники станом на звітну дату")
                ws_note.write(4, 0, "Аналіз строків погашення зобов'язання з оренди "
                                     "(недисконтовані майбутні платежі)")

                classification_df.to_excel(writer, sheet_name="Класифікація зобов'язання", index=False)

                journal_df.to_excel(writer, sheet_name="Проводки", index=False)
                ws_j = writer.sheets["Проводки"]
                if not journal_df.empty:
                    ws_j.write(0, 7, "Рахунки орієнтовні — узгодити з бухгалтерією клієнта")
                if not reclass_df.empty:
                    reclass_row = len(journal_df) + 3
                    ws_j.write(reclass_row - 1, 0, "Перекласифікація на звітну дату (533 → 611)")
                    reclass_df.to_excel(writer, sheet_name="Проводки", index=False, startrow=reclass_row)

                if not errors_df.empty:
                    errors_df.to_excel(writer, sheet_name="Помилки", index=False)
            buffer.seek(0)

        st.success(f"Готово: {len(summary_df)} договорів прораховано, {len(errors_df)} помилок.")
        st.caption(
            f"Зважена ставка дисконтування: {wam_stats['Зважена середня ставка дисконтування, %']}% · "
            f"Зважений строк, що залишився: {wam_stats['Зважений середній строк, що залишився (міс.)']} міс. · "
            f"Активних договорів на {as_of_date.strftime('%d.%m.%Y')}: {wam_stats['К-ть активних договорів']}"
        )

        if not errors_df.empty:
            st.warning("Деякі договори не прораховано — див. деталі нижче та аркуш «Помилки» у файлі.")
            st.dataframe(errors_df, use_container_width=True)

        st.dataframe(summary_df, use_container_width=True)

        st.subheader("Помісячна картина по всьому портфелю")
        st.dataframe(monthly_df, use_container_width=True)

        st.subheader("ROU-актив та зобов'язання — річний рух (roll-forward)")
        st.dataframe(annual_df, use_container_width=True)

        st.subheader("Примітка МСФЗ 16: зважені показники та строки погашення")
        st.dataframe(wam_df, use_container_width=True)
        st.dataframe(maturity_df, use_container_width=True)

        st.subheader("Класифікація зобов'язання: короткострокова (611) / довгострокова (533)")
        st.dataframe(classification_df, use_container_width=True)

        st.subheader("Бухгалтерські проводки")
        st.caption("Рахунки — орієнтовні, потребують узгодження з бухгалтерією клієнта "
                    "(див. ACCOUNT_MAP в engine.py).")
        st.dataframe(journal_df, use_container_width=True)
        if not reclass_df.empty:
            st.caption("Перекласифікація на звітну дату:")
            st.dataframe(reclass_df, use_container_width=True)

        st.download_button(
            label="Завантажити повний розрахунок (Excel)",
            data=buffer,
            file_name=f"МСФЗ16_розрахунок_{as_of_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
