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
            'Net Plan QTR':        'plan_qtr',
            'Net Plan':            'wk_plan',
            'Net Act':             'wk_act',
            'Net Gap':             'gap',
            'Prev WK GAP':         'prev_gap',
            'WOW':                 'wow',
            'Net Plan BPM QTR':    'bpm_plan_qtr',
            'Net Plan BPM':        'bpm_wk_plan',
            'Net Act BPM':         'bpm_wk_act',
            'Net BPM Gap':         'bpm_gap',
            'PREV WEEK ACT BPM':   'bpm_prev_gap',
            'WOW BPM':             'bpm_wow',
            'Delta Reason':        'delta_reason',
            'Recovery Plan':       'recovery_plan',
        },
    },
}


def _text(val):
    """Clean text value."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    return str(val).strip()


def _normalize_column_label(value: str) -> str:
    if value is None:
        return ''
    text = str(value).strip().lower()
    text = ''.join(ch if ch.isalnum() else ' ' for ch in text)
    return ' '.join(text.split())


def _match_column_name(raw_label: str, known_labels: dict) -> str | None:
    normalized = _normalize_column_label(raw_label)
    if normalized in known_labels:
        return known_labels[normalized]

    tokens = set(normalized.split())
    if 'account' in tokens:
        return 'account'
    if 'recovery' in tokens and 'plan' in tokens:
        return 'recovery_plan'
    if 'delta' in tokens and 'reason' in tokens:
        return 'delta_reason'
    if 'bpm' in tokens:
        if 'gap' in tokens:
            return 'bpm_gap'
        if 'wow' in tokens:
            return 'bpm_wow'
        if 'prev' in tokens and 'gap' in tokens:
            return 'bpm_prev_gap'
        if 'plan' in tokens and any(x in tokens for x in ('qtr', 'q1', 'qtd', 'quarter')):
            return 'bpm_plan_qtr'
        if 'plan' in tokens:
            return 'bpm_wk_plan'
        if 'act' in tokens:
            return 'bpm_wk_act'
    if 'gap' in tokens:
        return 'gap'
    if 'wow' in tokens:
        return 'wow'
    if 'prev' in tokens and 'gap' in tokens:
        return 'prev_gap'
    if 'plan' in tokens and any(x in tokens for x in ('qtr', 'q1', 'qtd', 'quarter')):
        return 'plan_qtr'
    if 'plan' in tokens:
        return 'wk_plan'
    if 'act' in tokens or 'actual' in tokens:
        return 'wk_act'
    return None


def _clean_number(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, str):
        text = val.strip().replace(',', '').replace(' ', '')
        if text in ('', '—', '-', 'na', 'n/a'):
            return None
        if text.startswith('(') and text.endswith(')'):
            text = '-' + text[1:-1]
        try:
            return float(text)
        except ValueError:
            return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _fmt(val, decimals=1):
    """Format a numeric value, returning '—' for missing data."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return '—'
    try:
        f = _clean_number(val)
        if f is None:
            return '—'
        if f == int(f):
            return str(int(f))
        return f'{f:.{decimals}f}'
    except (TypeError, ValueError):
        text = str(val).strip()
        return text if text else '—'


def _raw(val):
    """Return float or None."""
    return _clean_number(val)


def load_metric(filepath: str, metric: str, accounts_filter) -> list[dict]:
    """Load one metric sheet and return a list of account dicts."""
    cfg = _SHEET_MAP[metric]
    xls = pd.ExcelFile(filepath)
    sheet_name = cfg['sheet']
    if sheet_name not in xls.sheet_names:
        lower_names = {n.lower(): n for n in xls.sheet_names}
        sheet_name = lower_names.get(sheet_name.lower(), sheet_name)

    df = None
    for header in (0, 1):
        candidate = pd.read_excel(xls, sheet_name=sheet_name, header=header)
        normalized_map = {
            _normalize_column_label(raw): dest
            for raw, dest in cfg['col_map'].items()
        }
        rename_map = {}
        for col in candidate.columns:
            mapped = _match_column_name(col, normalized_map)
            if mapped:
                rename_map[col] = mapped
        candidate = candidate.rename(columns=rename_map)
        if 'account' in candidate.columns:
            df = candidate
            break
    if df is None or 'account' not in df.columns:
        raise ValueError(f"Unable to locate the account column for metric '{metric}' in sheet '{sheet_name}'")

    # drop rows without a valid account name
    df = df[df['account'].notna()]
    df = df[df['account'].astype(str).str.strip() != '']
    df = df[~df['account'].astype(str).str.lower().str.contains('grand total|total')]

    if accounts_filter != 'all':
        upper = [a.upper() for a in accounts_filter]
        df = df[df['account'].astype(str).str.upper().isin(upper)]

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
