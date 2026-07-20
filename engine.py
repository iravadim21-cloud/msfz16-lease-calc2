"""
IFRS16 (МСФЗ 16) batch lease amortization engine — перевикористовуваний модуль.
Імпортується і з CLI-тесту, і з веб-інтерфейсу (app.py).

Припущення методології (підлягає підтвердженню клієнтом):
  - Річний платіж сплачується в річницю дати початку договору (постнумерандо).
  - Неповний перший/останній період (менше 12 міс.) — пропорційна частка річної суми.
"""
import openpyxl
import pandas as pd

MONTHS_UA = {
    'січень': 1, 'лютий': 2, 'березень': 3, 'квітень': 4, 'травень': 5, 'червень': 6,
    'липень': 7, 'серпень': 8, 'вересень': 9, 'жовтень': 10, 'листопад': 11, 'грудень': 12
}

REQUIRED_COLUMNS = ['Номер договору', 'ПІБ пайовика', 'Дата початку оренди', 'Дата закінчення оренди', 'Сума оренди, грн']


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


def calc_contract(row, rates):
    contract_id = row['Номер договору']
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
            'contract_id': contract_id, 'lessor': lessor,
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
        'contract_id': contract_id, 'lessor': lessor,
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


def run_batch(df, rates):
    summaries, all_schedule, errors = [], [], []
    for _, row in df.iterrows():
        try:
            s, sch = calc_contract(row, rates)
            summaries.append(s)
            all_schedule.extend(sch)
        except Exception as e:
            errors.append({'contract_id': row.get('Номер договору'), 'error': str(e)})
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
