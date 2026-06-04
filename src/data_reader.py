import pandas as pd


_SHEET_MAP = {
    'RU': {
        'sheet': 'RU',
        'col_map': {
            'Account':           'account',
            'RU Plan QTR':       'plan_qtr',
            'RU Plan':      'wk_plan',
            'RU Act':       'wk_act',
            'RU Gap':            'gap',
            'Prev WK GAP':       'prev_gap',
            'WOW':               'wow',
            'RU Plan BPM QTR':   'bpm_plan_qtr',
            'RU Plan BPM':  'bpm_wk_plan',
            'RU Act BPM':   'bpm_wk_act',
            'RU BPM Gap':        'bpm_gap',
            'Prev week gap':     'bpm_prev_gap',
            'WOW BPM':           'bpm_wow',
            'Delta Reason':      'delta_reason',
            'Recovery Plan':     'recovery_plan',
        },
    },
    'RD': {
        'sheet': 'RD',
        'col_map': {
            'Account':            'account',
            'RD Plan QTR':        'plan_qtr',
            'RD Plan QTD':   'wk_plan',
            'RD Act QTD':    'wk_act',
            'RD Gap':             'gap',
            'PREV WK GAP':        'prev_gap',
            'WOW':                'wow',
            'RD Plan BPM QTR':    'bpm_plan_qtr',
            'RD Plan BPM':   'bpm_wk_plan',
            'RD Act BPM':    'bpm_wk_act',
            'RD BPM Gap':         'bpm_gap',
            'PREV WEEK GAP':      'bpm_prev_gap',
            'WOW BPM':            'bpm_wow',
            'Delta Reason':       'delta_reason',
            'Recovery Plan':      'recovery_plan',
        },
    },
    'Netadd': {
        'sheet': 'Netadd',
        'col_map': {
            'Account':             'account',
            'Net Plan Q1':         'plan_qtr',
            'Net Plan':       'wk_plan',
            'Net Act':        'wk_act',
            'Net Gap':             'gap',
            'Prev WK GAP':         'prev_gap',
            'WOW':                 'wow',
            'Net Plan BPM':        'bpm_plan_qtr',
            'Net Plan BPM':   'bpm_wk_plan',
            'Net Act BPM':    'bpm_wk_act',
            'Net BPM Gap':         'bpm_gap',
            'PREV WEEK ACT BPM':   'bpm_prev_gap',
            'WOW BPM':             'bpm_wow',
            'Delta Reason':        'delta_reason',
            'Recovery Plan':       'recovery_plan',
        },
    },
}


def _fmt(val, decimals=1):
    """Format a numeric value, returning '—' for missing data."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return '—'
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return f'{f:.{decimals}f}'
    except (TypeError, ValueError):
        text = str(val).strip()
        return text if text else '—'


def _raw(val):
    """Return float or None."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _text(val):
    """Clean text value."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    return str(val).strip()


def load_metric(filepath: str, metric: str, accounts_filter) -> list[dict]:
    """Load one metric sheet and return a list of account dicts."""
    cfg = _SHEET_MAP[metric]
    df = pd.read_excel(filepath, sheet_name=cfg['sheet'], header=1)
    df = df.rename(columns=cfg['col_map'])

    # drop rows without a valid account name
    df = df[df['account'].notna()]
    df = df[df['account'].astype(str).str.strip() != '']
    df = df[~df['account'].astype(str).str.lower().str.contains('grand total|total')]

    if accounts_filter != 'all':
        upper = [a.upper() for a in accounts_filter]
        df = df[df['account'].str.upper().isin(upper)]

    records = []
    for _, row in df.iterrows():
        rec = {
            'account':       _text(row.get('account')),
            'metric':        metric,
            'plan_qtr':      _raw(row.get('plan_qtr')),
            'wk_plan':       _raw(row.get('wk_plan')),
            'wk_act':        _raw(row.get('wk_act')),
            'gap':           _raw(row.get('gap')),
            'prev_gap':      _raw(row.get('prev_gap')),
            'wow':           _raw(row.get('wow')),
            'bpm_plan_qtr':  _raw(row.get('bpm_plan_qtr')),
            'bpm_wk_plan':   _raw(row.get('bpm_wk_plan')),
            'bpm_wk_act':    _raw(row.get('bpm_wk_act')),
            'bpm_gap':       _raw(row.get('bpm_gap')),
            'bpm_prev_gap':  _raw(row.get('bpm_prev_gap')),
            'bpm_wow':       _raw(row.get('bpm_wow')),
            'delta_reason':  _text(row.get('delta_reason')),
            'recovery_plan': _text(row.get('recovery_plan')),
            # formatted display strings
            'fmt_plan_qtr':     _fmt(row.get('plan_qtr')),
            'fmt_wk_plan':      _fmt(row.get('wk_plan')),
            'fmt_wk_act':       _fmt(row.get('wk_act')),
            'fmt_gap':          _fmt(row.get('gap')),
            'fmt_wow':          _fmt(row.get('wow')),
            'fmt_bpm_plan_qtr': _fmt(row.get('bpm_plan_qtr')),
            'fmt_bpm_wk_plan':  _fmt(row.get('bpm_wk_plan')),
            'fmt_bpm_wk_act':   _fmt(row.get('bpm_wk_act')),
            'fmt_bpm_gap':      _fmt(row.get('bpm_gap')),
            'fmt_bpm_wow':      _fmt(row.get('bpm_wow')),
        }
        records.append(rec)
    return records


def get_all_data(filepath: str, metrics: list, accounts_filter) -> dict:
    result = {}
    for m in metrics:
        result[m] = load_metric(filepath, m, accounts_filter)
    return result
