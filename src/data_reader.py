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
    if 'plan' in tokens and 'qtd' in tokens:
        return 'wk_plan'
    if 'plan' in tokens and any(x in tokens for x in ('qtr', 'q1', 'quarter')):
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
    """Format a numeric value as a whole number, returning '—' for missing data."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return '—'
    try:
        f = _clean_number(val)
        if f is None:
            return '—'
        return str(int(round(f)))
    except (TypeError, ValueError):
        text = str(val).strip()
        return text if text else '—'


def _raw(val):
    """Return float or None."""
    return _clean_number(val)


def _positive_if_rd(metric: str, val):
    if metric == 'RD' and val is not None:
        try:
            return abs(float(val))
        except (TypeError, ValueError):
            return val
    return val


def _invert_if_rd(metric: str, val):
    if metric == 'RD' and val is not None:
        try:
            return -float(val)
        except (TypeError, ValueError):
            return val
    return val


def _calc_gap(plan, act):
    if plan is None or act is None:
        return None
    try:
        return float(act) - float(plan)
    except (TypeError, ValueError):
        return None


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
        plan_qtr = _positive_if_rd(metric, _raw(row.get('plan_qtr')))
        wk_plan = _positive_if_rd(metric, _raw(row.get('wk_plan')))
        wk_act = _positive_if_rd(metric, _raw(row.get('wk_act')))
        gap = _calc_gap(wk_plan, wk_act) if metric == 'RD' else _raw(row.get('gap'))
        prev_gap = _positive_if_rd(metric, _raw(row.get('prev_gap')))
        wow = _invert_if_rd(metric, _raw(row.get('wow')))
        bpm_plan_qtr = _positive_if_rd(metric, _raw(row.get('bpm_plan_qtr')))
        bpm_wk_plan = _positive_if_rd(metric, _raw(row.get('bpm_wk_plan')))
        bpm_wk_act = _positive_if_rd(metric, _raw(row.get('bpm_wk_act')))
        bpm_gap = _calc_gap(bpm_wk_plan, bpm_wk_act) if metric == 'RD' else _raw(row.get('bpm_gap'))
        bpm_prev_gap = _positive_if_rd(metric, _raw(row.get('bpm_prev_gap')))
        bpm_wow = _invert_if_rd(metric, _raw(row.get('bpm_wow')))

        rec = {
            'account':       _text(row.get('account')),
            'metric':        metric,
            'plan_qtr':      plan_qtr,
            'wk_plan':       wk_plan,
            'wk_act':        wk_act,
            'gap':           gap,
            'prev_gap':      prev_gap,
            'wow':           wow,
            'bpm_plan_qtr':  bpm_plan_qtr,
            'bpm_wk_plan':   bpm_wk_plan,
            'bpm_wk_act':    bpm_wk_act,
            'bpm_gap':       bpm_gap,
            'bpm_prev_gap':  bpm_prev_gap,
            'bpm_wow':       bpm_wow,
            'delta_reason':  _text(row.get('delta_reason')),
            'recovery_plan': _text(row.get('recovery_plan')),
            # formatted display strings
            'fmt_plan_qtr':     _fmt(plan_qtr),
            'fmt_wk_plan':      _fmt(wk_plan),
            'fmt_wk_act':       _fmt(wk_act),
            'fmt_gap':          _fmt(gap),
            'fmt_wow':          _fmt(wow),
            'fmt_bpm_plan_qtr': _fmt(bpm_plan_qtr),
            'fmt_bpm_wk_plan':  _fmt(bpm_wk_plan),
            'fmt_bpm_wk_act':   _fmt(bpm_wk_act),
            'fmt_bpm_gap':      _fmt(bpm_gap),
            'fmt_bpm_wow':      _fmt(bpm_wow),
        }
        records.append(rec)
    return records


def _derive_netadd_records(ru_records: list[dict], rd_records: list[dict]) -> list[dict]:
    ru_map = {r['account'].strip().upper(): r for r in ru_records}
    rd_map = {r['account'].strip().upper(): r for r in rd_records}
    account_keys = sorted(set(ru_map) | set(rd_map))
    records = []

    def _safe(v):
        return 0.0 if v is None else float(v)

    for key in account_keys:
        ru = ru_map.get(key, {})
        rd = rd_map.get(key, {})
        account = ru.get('account') or rd.get('account') or key

        plan_qtr = _safe(ru.get('plan_qtr')) - _safe(rd.get('plan_qtr'))
        wk_plan = _safe(ru.get('wk_plan')) - _safe(rd.get('wk_plan'))
        wk_act = _safe(ru.get('wk_act')) - _safe(rd.get('wk_act'))
        gap = wk_act - wk_plan
        prev_gap = _safe(ru.get('prev_gap')) - _safe(rd.get('prev_gap'))
        wow = _safe(ru.get('wow')) - _safe(rd.get('wow'))
        bpm_plan_qtr = _safe(ru.get('bpm_plan_qtr')) - _safe(rd.get('bpm_plan_qtr'))
        bpm_wk_plan = _safe(ru.get('bpm_wk_plan')) - _safe(rd.get('bpm_wk_plan'))
        bpm_wk_act = _safe(ru.get('bpm_wk_act')) - _safe(rd.get('bpm_wk_act'))
        bpm_gap = bpm_wk_act - bpm_wk_plan
        bpm_prev_gap = _safe(ru.get('bpm_prev_gap')) - _safe(rd.get('bpm_prev_gap'))
        bpm_wow = _safe(ru.get('bpm_wow')) - _safe(rd.get('bpm_wow'))

        rec = {
            'account': account,
            'metric': 'Netadd',
            'plan_qtr': plan_qtr,
            'wk_plan': wk_plan,
            'wk_act': wk_act,
            'gap': gap,
            'prev_gap': prev_gap,
            'wow': wow,
            'bpm_plan_qtr': bpm_plan_qtr,
            'bpm_wk_plan': bpm_wk_plan,
            'bpm_wk_act': bpm_wk_act,
            'bpm_gap': bpm_gap,
            'bpm_prev_gap': bpm_prev_gap,
            'bpm_wow': bpm_wow,
            'delta_reason': ru.get('delta_reason') or rd.get('delta_reason') or '',
            'recovery_plan': '',
            'fmt_plan_qtr': _fmt(plan_qtr),
            'fmt_wk_plan': _fmt(wk_plan),
            'fmt_wk_act': _fmt(wk_act),
            'fmt_gap': _fmt(gap),
            'fmt_wow': _fmt(wow),
            'fmt_bpm_plan_qtr': _fmt(bpm_plan_qtr),
            'fmt_bpm_wk_plan': _fmt(bpm_wk_plan),
            'fmt_bpm_wk_act': _fmt(bpm_wk_act),
            'fmt_bpm_gap': _fmt(bpm_gap),
            'fmt_bpm_wow': _fmt(bpm_wow),
        }
        records.append(rec)
    return records


def get_all_data(filepath: str, metrics: list, accounts_filter) -> dict:
    result = {}
    for m in metrics:
        if m == 'Netadd':
            continue
        result[m] = load_metric(filepath, m, accounts_filter)

    if 'Netadd' in metrics:
        if 'RU' in result and 'RD' in result:
            result['Netadd'] = _derive_netadd_records(result['RU'], result['RD'])
        else:
            result['Netadd'] = load_metric(filepath, 'Netadd', accounts_filter)

    return result
