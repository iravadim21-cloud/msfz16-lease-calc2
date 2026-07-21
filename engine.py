"""
IFRS16 (МСФЗ 16) batch lease amortization engine — перевикористовуваний модуль.
Імпортується і з CLI-тесту, і з веб-інтерфейсу (app.py).

Припущення методології (підлягає підтвердженню клієнтом):
  - Річний платіж сплачується в річницю дати початку договору (постнумерандо).
  - Неповний перший/останній період (менше 12 міс.) — пропорційна частка річної суми.
"""
import io
import datetime

import openpyxl
import pandas as pd

MONTHS_UA = {
    'січень': 1, 'лютий': 2, 'березень': 3, 'квітень': 4, 'травень': 5, 'червень': 6,
    'липень': 7, 'серпень': 8, 'вересень': 9, 'жовтень': 10, 'листопад': 11, 'грудень': 12
}

REQUIRED_COLUMNS = ['Номер договору', 'ПІБ пайовика', 'Дата початку оренди', 'Дата закінчення оренди', 'Сума оренди, грн']
CODE_COLUMN = 'Код договору'  # опційна колонка: якщо заповнена клієнтом — саме вона є ідентифікатором договору
TEMPLATE_COLUMNS = [CODE_COLUMN] + REQUIRED_COLUMNS


def load_rates(path_or_buffer):
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    ws = wb.active
    rates = {}
    current_year = None
    for row in ws.iter_rows(min_row=6, values_only=True):
        col0 = row[0]
        if col0 is None:
            continue
        if isinstance(col0, (int, float)):
            current_year = int(col0)
            continue
        month_key = str(col0).strip().lower()
        if month_key in MONTHS_UA and current_year is not None:
            m = MONTHS_UA[month_key]
            short_, mid_, long_ = row[5], row[6], row[7]
            if short_ is not None and mid_ is not None and long_ is not None:
                rates[(current_year, m)] = {'short': float(short_), 'mid': float(mid_), 'long': float(long_)}
    return rates


def term_category(start_date, end_date):
    diff_days = (end_date - start_date).days
    diff_years = diff_days / 365.25
    if diff_years < 1:
        return 'short'
    elif diff_years <= 5:
        return 'mid'
    else:
        return 'long'


def total_months(start_date, end_date):
    return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) + 1


def build_payment_schedule(annual_payment, months):
    payments = {}
    m = 12
    while m <= months:
        payments[m] = round(annual_payment, 2)
        m += 12
    last_full_payment_month = m - 12
    stub = months - last_full_payment_month
    if stub > 0:
        prorated = round(annual_payment * stub / 12, 2)
        payments[months] = payments.get(months, 0) + prorated
    return payments


def calc_contract(row, rates, unique_code=None):
    contract_id = row['Номер договору']
    if unique_code is None:
        unique_code = str(contract_id)
    lessor = row['ПІБ пайовика']
    start_date = row['Дата початку оренди']
    end_date = row['Дата закінчення оренди']
    annual_payment = float(row['Сума оренди, грн'])

    if pd.isna(start_date) or pd.isna(end_date):
        raise ValueError("Порожня дата початку або закінчення")
    if hasattr(start_date, 'to_pydatetime'):
        start_date = start_date.to_pydatetime()
    if hasattr(end_date, 'to_pydatetime'):
        end_date = end_date.to_pydatetime()
    if end_date <= start_date:
        raise ValueError("Дата закінчення не пізніше дати початку")
    if annual_payment <= 0:
        raise ValueError("Сума оренди має бути додатною")

    category = term_category(start_date, end_date)
    key = (start_date.year, start_date.month)
    if key not in rates:
        raise ValueError(f"Немає ставки в довіднику на {start_date.month:02d}.{start_date.year}")
    rate_percent = rates[key][category]
    rate_decimal = rate_percent / 100
    monthly_rate = rate_decimal / 12

    months = total_months(start_date, end_date)
    pay_schedule = build_payment_schedule(annual_payment, months)

    pv = sum(amt / ((1 + monthly_rate) ** t) for t, amt in pay_schedule.items())
    initial_liability = round(pv, 2)
    initial_rou = initial_liability
    monthly_depreciation = round(initial_rou / months, 2)

    schedule_rows = []
    liab_balance = initial_liability
    rou_balance = initial_rou
    cal_year, cal_month = start_date.year, start_date.month

    for m in range(1, months + 1):
        interest = round(liab_balance * monthly_rate, 2)
        payment = pay_schedule.get(m, 0.0)
        if m in pay_schedule:
            principal = round(payment - interest, 2)
            liab_end = round(liab_balance + interest - payment, 2)
        else:
            principal = 0.0
            liab_end = round(liab_balance + interest, 2)

        dep = monthly_depreciation
        rou_end = round(rou_balance - dep, 2)

        if m == months:
            liab_end = 0.0
            if m in pay_schedule:
                principal = round(liab_balance + interest, 2)
            rou_end = 0.0
            dep = round(rou_balance, 2)

        schedule_rows.append({
            'unique_code': unique_code, 'contract_id': contract_id, 'lessor': lessor,
            'month_number': m, 'calendar_year': cal_year, 'calendar_month': cal_month,
            'liability_balance_start': liab_balance, 'interest_charged': interest,
            'payment_amount': payment, 'principal_repayment': principal,
            'liability_balance_end': liab_end,
            'rou_asset_start': rou_balance, 'depreciation': dep, 'rou_asset_end': rou_end,
        })

        liab_balance = liab_end
        rou_balance = rou_end
        cal_month += 1
        if cal_month > 12:
            cal_month = 1
            cal_year += 1

    summary_row = {
        'unique_code': unique_code, 'contract_id': contract_id, 'lessor': lessor,
        'start_date': start_date.strftime('%d.%m.%Y'), 'end_date': end_date.strftime('%d.%m.%Y'),
        'term_category': {'short': 'До 1 року', 'mid': 'Від 1 до 5 років', 'long': 'Більше 5 років'}[category],
        'term_months': months, 'annual_payment': annual_payment,
        'discount_rate_percent': rate_percent,
        'initial_liability': initial_liability, 'initial_rou_asset': initial_rou,
        'total_interest': round(sum(r['interest_charged'] for r in schedule_rows), 2),
        'total_depreciation': round(sum(r['depreciation'] for r in schedule_rows), 2),
        'n_payments': len(pay_schedule),
    }
    return summary_row, schedule_rows


def validate_register(df):
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return missing


def find_duplicate_contract_numbers(df):
    """Повертає список номерів договору ('Номер договору'), які зустрічаються
    в реєстрі більше одного разу. Це не блокує розрахунок — кожному рядку
    все одно присвоюється унікальний unique_code (з колонки 'Код договору',
    якщо вона заповнена, або автоматично L-00001, L-00002, ...), але
    дублікати варто показати користувачу як попередження."""
    return _find_duplicates(df, 'Номер договору')


def find_duplicate_contract_codes(df):
    """Те саме, але для колонки 'Код договору' (якщо клієнт її заповнив)."""
    return _find_duplicates(df, CODE_COLUMN)


def _find_duplicates(df, column):
    if column not in df.columns:
        return []
    non_empty = df[column].dropna()
    non_empty = non_empty[non_empty.astype(str).str.strip() != '']
    counts = non_empty.value_counts()
    return counts[counts > 1].index.tolist()


def run_batch(df, rates):
    summaries, all_schedule, errors = [], [], []
    has_code_col = CODE_COLUMN in df.columns
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        fallback_code = f"L-{i:05d}"
        user_code = None
        if has_code_col:
            val = row.get(CODE_COLUMN)
            if pd.notna(val) and str(val).strip() != '':
                user_code = str(val).strip()
        unique_code = user_code or fallback_code
        try:
            s, sch = calc_contract(row, rates, unique_code=unique_code)
            summaries.append(s)
            all_schedule.extend(sch)
        except Exception as e:
            errors.append({'unique_code': unique_code, 'contract_id': row.get('Номер договору'), 'error': str(e)})
    return pd.DataFrame(summaries), pd.DataFrame(all_schedule), pd.DataFrame(errors)


MONTH_NAMES_UA = {
    1: 'Січень', 2: 'Лютий', 3: 'Березень', 4: 'Квітень', 5: 'Травень', 6: 'Червень',
    7: 'Липень', 8: 'Серпень', 9: 'Вересень', 10: 'Жовтень', 11: 'Листопад', 12: 'Грудень'
}


def build_monthly_summary(schedule_df):
    """Згортає графік по всіх договорах у зведення по кожному календарному
    місяцю — та картина, яку бачить бухгалтер для проводок за період,
    а не по кожному контрагенту окремо."""
    if schedule_df.empty:
        return pd.DataFrame(columns=[
            'Період', 'Рік', 'Місяць', 'К-ть договорів у періоді',
            "Зобов'язання на початок періоду", 'Нарахований відсоток',
            'Сплачено (грошовий потік)', 'Погашення тіла зобов\'язання',
            "Зобов'язання на кінець періоду",
            'ROU-актив на початок періоду', 'Амортизація ROU-активу', 'ROU-актив на кінець періоду',
        ])

    grp = schedule_df.groupby(['calendar_year', 'calendar_month']).agg(
        n_contracts=('contract_id', 'nunique'),
        liability_balance_start=('liability_balance_start', 'sum'),
        interest_charged=('interest_charged', 'sum'),
        payment_amount=('payment_amount', 'sum'),
        principal_repayment=('principal_repayment', 'sum'),
        liability_balance_end=('liability_balance_end', 'sum'),
        rou_asset_start=('rou_asset_start', 'sum'),
        depreciation=('depreciation', 'sum'),
        rou_asset_end=('rou_asset_end', 'sum'),
    ).reset_index()

    grp = grp.sort_values(['calendar_year', 'calendar_month']).reset_index(drop=True)

    num_cols = ['liability_balance_start', 'interest_charged', 'payment_amount',
                'principal_repayment', 'liability_balance_end',
                'rou_asset_start', 'depreciation', 'rou_asset_end']
    for c in num_cols:
        grp[c] = grp[c].round(2)

    grp.insert(0, 'Період', grp.apply(
        lambda r: f"{MONTH_NAMES_UA[int(r['calendar_month'])]} {int(r['calendar_year'])}", axis=1))

    grp = grp.rename(columns={
        'calendar_year': 'Рік', 'calendar_month': 'Місяць', 'n_contracts': 'К-ть договорів у періоді',
        'liability_balance_start': "Зобов'язання на початок періоду",
        'interest_charged': 'Нарахований відсоток',
        'payment_amount': 'Сплачено (грошовий потік)',
        'principal_repayment': "Погашення тіла зобов'язання",
        'liability_balance_end': "Зобов'язання на кінець періоду",
        'rou_asset_start': 'ROU-актив на початок періоду',
        'depreciation': 'Амортизація ROU-активу',
        'rou_asset_end': 'ROU-актив на кінець періоду',
    })
    return grp


# ---------------------------------------------------------------------------
# Roll-forward ROU-активу та зобов'язання (річний розріз для примітки)
# ---------------------------------------------------------------------------

def _rollforward_for_period(sub):
    """Рух зобов'язання та ROU-активу для підмножини рядків графіка (sub),
    яка вже містить колонку 'ym' = calendar_year*12 + calendar_month.

    Опорна тотожність (звіряється тестом нижче):
        opening + additions + interest - payments == closing        (зобов'язання)
        rou_opening + rou_additions - depreciation == rou_closing    (ROU-актив)
    """
    if sub.empty:
        return None
    first, last = sub['ym'].min(), sub['ym'].max()
    is_new = sub['month_number'] == 1          # перший місяць договору = визнання
    at_first = sub['ym'] == first
    at_last = sub['ym'] == last
    return {
        'liability_opening': round(sub.loc[at_first & ~is_new, 'liability_balance_start'].sum(), 2),
        'liability_additions': round(sub.loc[is_new, 'liability_balance_start'].sum(), 2),
        'interest_charged': round(sub['interest_charged'].sum(), 2),
        'payments': round(sub['payment_amount'].sum(), 2),
        'liability_closing': round(sub.loc[at_last, 'liability_balance_end'].sum(), 2),
        'rou_opening': round(sub.loc[at_first & ~is_new, 'rou_asset_start'].sum(), 2),
        'rou_additions': round(sub.loc[is_new, 'rou_asset_start'].sum(), 2),
        'depreciation': round(sub['depreciation'].sum(), 2),
        'rou_closing': round(sub.loc[at_last, 'rou_asset_end'].sum(), 2),
    }


def build_annual_rollforward(schedule_df):
    """Річний рух ROU-активу та зобов'язання з оренди по всьому портфелю —
    та сама логіка, що й у 'Помісячно', але згорнута по календарних роках,
    у форматі, придатному для примітки до фінзвітності."""
    cols = ['Рік',
            "Зобов'язання на початок року", "Визнано нових договорів (зобов'язання)",
            'Нарахований відсоток за рік', 'Сплачено за рік', "Зобов'язання на кінець року",
            'ROU-актив на початок року', 'Визнано нових договорів (ROU)',
            'Амортизація ROU за рік', 'ROU-актив на кінець року']
    if schedule_df.empty:
        return pd.DataFrame(columns=cols)

    df = schedule_df.copy()
    df['ym'] = df['calendar_year'] * 12 + df['calendar_month']

    rows = []
    for year in sorted(df['calendar_year'].unique()):
        r = _rollforward_for_period(df[df['calendar_year'] == year])
        if r is None:
            continue
        rows.append({
            'Рік': int(year),
            "Зобов'язання на початок року": r['liability_opening'],
            "Визнано нових договорів (зобов'язання)": r['liability_additions'],
            'Нарахований відсоток за рік': r['interest_charged'],
            'Сплачено за рік': r['payments'],
            "Зобов'язання на кінець року": r['liability_closing'],
            'ROU-актив на початок року': r['rou_opening'],
            'Визнано нових договорів (ROU)': r['rou_additions'],
            'Амортизація ROU за рік': r['depreciation'],
            'ROU-актив на кінець року': r['rou_closing'],
        })
    return pd.DataFrame(rows, columns=cols)


# ---------------------------------------------------------------------------
# Примітка МСФЗ 16: аналіз строків погашення + зважені показники
# ---------------------------------------------------------------------------

def _as_of_ym(as_of_date):
    """Звітна дата трактується як КІНЕЦЬ місяця, у якому вона знаходиться
    (відповідає квартальним датам 31.03 / 30.06 / 30.09 / 31.12 — це вже
    закритий місяць). Повертає ym = рік*12 + місяць цього місяця."""
    ts = pd.Timestamp(as_of_date)
    return ts.year * 12 + ts.month


MATURITY_COLS = ['До 1 року', 'Від 1 до 2 років', 'Від 2 до 3 років', 'Від 3 до 4 років',
                  'Від 4 до 5 років', 'Понад 5 років', 'Разом недисконтовані платежі']


def build_maturity_analysis(schedule_df, as_of_date):
    """Недисконтовані майбутні орендні платежі за річними періодами,
    станом на as_of_date (звітну дату, трактується як кінець місяця —
    напр. 31.03/30.06/30.09/31.12) — стандартна таблиця аналізу строків
    погашення зобов'язання з оренди за МСФЗ 16.93(б).

    Платіж місяця самої звітної дати вже врахований у закритому періоді
    (він включений у закриваючий залишок на цю дату) і в "майбутні" не
    потрапляє — рахуються лише платежі СТРОГО ПІСЛЯ звітного місяця."""
    if schedule_df.empty:
        return pd.DataFrame([[0.0] * 7], columns=MATURITY_COLS)

    df = schedule_df.copy()
    df['ym'] = df['calendar_year'] * 12 + df['calendar_month']
    as_of_ym = _as_of_ym(as_of_date)

    future = df[(df['ym'] > as_of_ym) & (df['payment_amount'] > 0)]
    if future.empty:
        return pd.DataFrame([[0.0] * 7], columns=MATURITY_COLS)

    months_ahead = future['ym'] - as_of_ym  # 1, 2, 3, ... (місяців після звітної дати)
    bucket = ((months_ahead - 1) // 12).clip(upper=5)
    buckets = future.groupby(bucket)['payment_amount'].sum()

    values = [round(float(buckets.get(i, 0.0)), 2) for i in range(6)]
    values.append(round(sum(values), 2))
    return pd.DataFrame([values], columns=MATURITY_COLS)


def build_wam_stats(summary_df, schedule_df, as_of_date):
    """Зважена середня ставка дисконтування і зважений середній строк, що
    залишився (у місяцях), по договорах, активних станом на as_of_date
    (трактується як кінець місяця). "Активний" = договір мав хоча б один
    місяць у графіку в цьому звітному місяці. Строк, що залишився,
    рахується від місяців СТРОГО ПІСЛЯ звітного (той місяць уже закрито).
    Вага — початкове зобов'язання договору (initial_liability)."""
    empty = {
        'Звітна дата': pd.Timestamp(as_of_date).strftime('%d.%m.%Y'),
        'К-ть активних договорів': 0,
        'Зважена середня ставка дисконтування, %': 0.0,
        'Зважений середній строк, що залишився (міс.)': 0.0,
    }
    if schedule_df.empty or summary_df.empty:
        return empty

    df = schedule_df.copy()
    df['ym'] = df['calendar_year'] * 12 + df['calendar_month']
    as_of_ym = _as_of_ym(as_of_date)

    active_ids = df.loc[df['ym'] == as_of_ym, 'contract_id'].unique()
    if len(active_ids) == 0:
        return empty

    remaining_months = (
        df[df['contract_id'].isin(active_ids) & (df['ym'] > as_of_ym)]
        .groupby('contract_id')['ym'].count()
    )

    s = summary_df.set_index('contract_id')
    weights = s.loc[active_ids, 'initial_liability']
    rates = s.loc[active_ids, 'discount_rate_percent']
    total_w = weights.sum()
    if total_w == 0:
        return empty

    wa_rate = (rates * weights).sum() / total_w
    wa_term = (remaining_months.reindex(active_ids).fillna(0) * weights).sum() / total_w

    return {
        'Звітна дата': pd.Timestamp(as_of_date).strftime('%d.%m.%Y'),
        'К-ть активних договорів': int(len(active_ids)),
        'Зважена середня ставка дисконтування, %': round(float(wa_rate), 2),
        'Зважений середній строк, що залишився (міс.)': round(float(wa_term), 1),
    }


# ---------------------------------------------------------------------------
# Бухгалтерські проводки (журнал)
# ---------------------------------------------------------------------------

# УВАГА: рахунки орієнтовні (типовий План рахунків України) і НЕ ПІДТВЕРДЖЕНІ
# клієнтом. Перед використанням проводок для імпорту в облікову систему —
# узгодити коди рахунків з головним бухгалтером і за потреби передати
# власний account_map у build_journal_entries().
DEFAULT_ACCOUNT_MAP = {
    'rou_asset': '109',                    # Право користування активом (ROU)
    'rou_amortization': '131',             # Знос (накопичена амортизація) ROU-активу
    'lease_liability': '533',              # Контрольний рахунок зобов'язання з оренди
                                            # (спрощення: нарахування % і платежі протягом
                                            # року йдуть тут; на звітну дату частина, що
                                            # підлягає погашенню протягом 12 міс., переноситься
                                            # на 611 проводкою перекласифікації нижче)
    'lease_liability_noncurrent': '533',   # Довгострокові зобов'язання з оренди
    'lease_liability_current': '611',      # Поточна заборгованість за довгостроковими зобов'язаннями
    'interest_expense': '952',             # Фінансові витрати (проценти за зобов'язанням)
    'amortization_expense': '943',         # Витрати на амортизацію ROU-активу
    'cash': '311',                         # Поточний рахунок
}

JOURNAL_COLS = ['Період', 'Рік', 'Місяць', 'Зміст операції', 'Дебет', 'Кредит', 'Сума, грн']


def build_journal_entries(schedule_df, account_map=None):
    """Портфельні бухгалтерські проводки за кожен календарний місяць:
    нарахування відсотка, амортизація ROU, сплата орендного платежу.
    Деталізацію по кожному договору окремо дивись у вкладці «Графік»."""
    accounts = {**DEFAULT_ACCOUNT_MAP, **(account_map or {})}
    if schedule_df.empty:
        return pd.DataFrame(columns=JOURNAL_COLS)

    grp = schedule_df.groupby(['calendar_year', 'calendar_month']).agg(
        interest_charged=('interest_charged', 'sum'),
        depreciation=('depreciation', 'sum'),
        payment_amount=('payment_amount', 'sum'),
    ).reset_index().sort_values(['calendar_year', 'calendar_month'])

    rows = []
    for _, r in grp.iterrows():
        period = f"{MONTH_NAMES_UA[int(r['calendar_month'])]} {int(r['calendar_year'])}"
        base = dict(Період=period, Рік=int(r['calendar_year']), Місяць=int(r['calendar_month']))

        if r['interest_charged'] > 0:
            rows.append({**base,
                         'Зміст операції': "Нарахування відсотка за зобов'язанням з оренди",
                         'Дебет': accounts['interest_expense'], 'Кредит': accounts['lease_liability'],
                         'Сума, грн': round(float(r['interest_charged']), 2)})
        if r['depreciation'] > 0:
            rows.append({**base,
                         'Зміст операції': 'Амортизація активу з права користування (ROU)',
                         'Дебет': accounts['amortization_expense'], 'Кредит': accounts['rou_amortization'],
                         'Сума, грн': round(float(r['depreciation']), 2)})
        if r['payment_amount'] > 0:
            rows.append({**base,
                         'Зміст операції': 'Сплата орендного платежу',
                         'Дебет': accounts['lease_liability'], 'Кредит': accounts['cash'],
                         'Сума, грн': round(float(r['payment_amount']), 2)})

    return pd.DataFrame(rows, columns=JOURNAL_COLS)


# ---------------------------------------------------------------------------
# Класифікація зобов'язання: довгострокова (533) / короткострокова (611) частина
# ---------------------------------------------------------------------------

LIABILITY_CLASS_COLS = ['unique_code', 'contract_id', 'lessor',
                         "Зобов'язання станом на звітну дату",
                         'Прогнозний залишок через 12 міс.',
                         'Короткострокова частина (до 12 міс.)', 'Рахунок (короткострокова)',
                         'Довгострокова частина (понад 12 міс.)', 'Рахунок (довгострокова)']


def build_liability_classification(schedule_df, as_of_date, account_map=None):
    """Розподіл зобов'язання з оренди на довгострокову і короткострокову
    частини станом на as_of_date, по кожному договору.

    as_of_date трактується як КІНЕЦЬ місяця (напр. 31.03/30.06/30.09/31.12) —
    тому за базу береться ЗАКРИВАЮЧИЙ залишок (liability_balance_end) місяця
    звітної дати, а не залишок на його початок: до звітної дати нарахування
    процента і платіж за цей місяць уже відбулись.

    Короткострокова частина = сума погашення тіла зобов'язання, запланована
    на найближчі 12 місяців ПІСЛЯ звітної дати (різниця між закриваючим
    залишком на звітну дату і закриваючим залишком через 12 місяців).
    Якщо договір закінчується раніше, ніж через 12 місяців — уся сума, що
    залишилась, вважається короткостроковою.
    """
    accounts = {**DEFAULT_ACCOUNT_MAP, **(account_map or {})}
    if schedule_df.empty:
        return pd.DataFrame(columns=LIABILITY_CLASS_COLS)

    df = schedule_df.copy()
    df['ym'] = df['calendar_year'] * 12 + df['calendar_month']
    as_of_ym = _as_of_ym(as_of_date)
    future_ym = as_of_ym + 12

    now_rows = df[df['ym'] == as_of_ym]
    if now_rows.empty:
        return pd.DataFrame(columns=LIABILITY_CLASS_COLS)
    now_rows = now_rows.set_index('contract_id')
    future_rows = df[df['ym'] == future_ym].set_index('contract_id')
    meta = df.drop_duplicates('contract_id').set_index('contract_id')[['unique_code', 'lessor']]

    rows = []
    for cid, r in now_rows.iterrows():
        balance_now = float(r['liability_balance_end'])
        balance_future = float(future_rows['liability_balance_end'].get(cid, 0.0))
        short = round(balance_now - balance_future, 2)
        long_ = round(balance_future, 2)
        rows.append({
            'unique_code': meta.loc[cid, 'unique_code'],
            'contract_id': cid,
            'lessor': meta.loc[cid, 'lessor'],
            "Зобов'язання станом на звітну дату": round(balance_now, 2),
            'Прогнозний залишок через 12 міс.': round(balance_future, 2),
            'Короткострокова частина (до 12 міс.)': short,
            'Рахунок (короткострокова)': accounts['lease_liability_current'],
            'Довгострокова частина (понад 12 міс.)': long_,
            'Рахунок (довгострокова)': accounts['lease_liability_noncurrent'],
        })

    result = pd.DataFrame(rows, columns=LIABILITY_CLASS_COLS).sort_values('unique_code').reset_index(drop=True)

    totals = pd.DataFrame([{
        'unique_code': '', 'contract_id': '', 'lessor': 'РАЗОМ ПО ПОРТФЕЛЮ',
        "Зобов'язання станом на звітну дату": round(result["Зобов'язання станом на звітну дату"].sum(), 2),
        'Прогнозний залишок через 12 міс.': round(result['Прогнозний залишок через 12 міс.'].sum(), 2),
        'Короткострокова частина (до 12 міс.)': round(result['Короткострокова частина (до 12 міс.)'].sum(), 2),
        'Рахунок (короткострокова)': accounts['lease_liability_current'],
        'Довгострокова частина (понад 12 міс.)': round(result['Довгострокова частина (понад 12 міс.)'].sum(), 2),
        'Рахунок (довгострокова)': accounts['lease_liability_noncurrent'],
    }], columns=LIABILITY_CLASS_COLS)

    return pd.concat([result, totals], ignore_index=True)


def build_reclassification_entry(schedule_df, as_of_date, account_map=None):
    """Проводка перекласифікації станом на звітну дату: частина зобов'язання,
    яка підлягає погашенню протягом 12 місяців, переноситься з довгострокового
    рахунка (533) на короткостроковий (611). Робиться раз на звітну дату
    (за замовчуванням — портфельно, одним рядком)."""
    accounts = {**DEFAULT_ACCOUNT_MAP, **(account_map or {})}
    cols = ['Дата', 'Зміст операції', 'Дебет', 'Кредит', 'Сума, грн']

    classification_df = build_liability_classification(schedule_df, as_of_date, account_map)
    if classification_df.empty:
        return pd.DataFrame(columns=cols)

    total_short = classification_df.loc[
        classification_df['lessor'] == 'РАЗОМ ПО ПОРТФЕЛЮ', 'Короткострокова частина (до 12 міс.)'
    ].sum()
    if total_short <= 0:
        return pd.DataFrame(columns=cols)

    return pd.DataFrame([{
        'Дата': pd.Timestamp(as_of_date).strftime('%d.%m.%Y'),
        'Зміст операції': "Перекласифікація частини зобов'язання з оренди на поточну "
                           "(підлягає погашенню протягом 12 міс. від звітної дати)",
        'Дебет': accounts['lease_liability_noncurrent'],
        'Кредит': accounts['lease_liability_current'],
        'Сума, грн': round(float(total_short), 2),
    }], columns=cols)


# ---------------------------------------------------------------------------
# Шаблон реєстру для заповнення клієнтом
# ---------------------------------------------------------------------------

TEMPLATE_EXAMPLE_ROW = ['001', 'Д-01/25', 'Іваненко Петро Миколайович',
                         datetime.date(2025, 1, 1), datetime.date(2027, 12, 31), 26000]


def build_template_bytes():
    """Генерує порожній шаблон реєстру договорів (.xlsx) із чітко визначеними,
    зафіксованими колонками (включно з 'Код договору' — власним ідентифікатором
    клієнта) — для заповнення. Фіксований формат знижує ризик помилок
    розрахунку через довільні назви/порядок/пропуски колонок у файлі клієнта."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        pd.DataFrame(columns=TEMPLATE_COLUMNS).to_excel(
            writer, sheet_name='Реєстр', index=False, startrow=2)
        workbook = writer.book
        ws = writer.sheets['Реєстр']

        title_fmt = workbook.add_format({'bold': True, 'font_size': 12})
        note_fmt = workbook.add_format({'italic': True, 'font_color': '#C00000'})
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        example_text_fmt = workbook.add_format({'italic': True, 'font_color': '#999999'})
        example_date_fmt = workbook.add_format({'italic': True, 'font_color': '#999999', 'num_format': 'dd.mm.yyyy'})
        example_money_fmt = workbook.add_format({'italic': True, 'font_color': '#999999', 'num_format': '#,##0.00'})
        date_fmt = workbook.add_format({'num_format': 'dd.mm.yyyy'})
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})

        last_col = len(TEMPLATE_COLUMNS) - 1
        ws.merge_range(0, 0, 0, last_col,
                        'Реєстр договорів оренди — заповніть таблицю нижче, рядок за рядком', title_fmt)
        ws.merge_range(1, 0, 1, last_col,
                        "Рядок 4 — приклад заповнення. Видаліть його перед завантаженням у розрахунок. "
                        "«Код договору» можна не заповнювати — тоді код буде присвоєно автоматично.", note_fmt)

        for i, col in enumerate(TEMPLATE_COLUMNS):
            ws.write(2, i, col, header_fmt)

        example = TEMPLATE_EXAMPLE_ROW
        ws.write(3, 0, example[0], example_text_fmt)
        ws.write(3, 1, example[1], example_text_fmt)
        ws.write(3, 2, example[2], example_text_fmt)
        ws.write_datetime(3, 3, example[3], example_date_fmt)
        ws.write_datetime(3, 4, example[4], example_date_fmt)
        ws.write(3, 5, example[5], example_money_fmt)

        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 16)
        ws.set_column(2, 2, 32)
        ws.set_column(3, 4, 20, date_fmt)
        ws.set_column(5, 5, 18, money_fmt)

        ws.freeze_panes(3, 0)

    buffer.seek(0)
    return buffer.getvalue()
