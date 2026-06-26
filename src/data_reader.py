import pandas as pd
import re


_SHEET_MAP = {
    'RU': {
        'sheet': 'RU',
        'col_map': {
            'Account':           'account',
            'RU Plan QTR':       'plan_qtr',
            'RU Plan':           'wk_plan',
            'RU Act':            'wk_act',
            'RU Gap':            'gap',
            'Prev WK GAP':       'prev_gap',
            'WOW':               'wow',
            'RU Plan BPM QTR':   'bpm_plan_qtr',
            'RU Plan BPM':       'bpm_wk_plan',
            'RU Act BPM':        'bpm_wk_act',
            'RU BPM Gap':        'bpm_gap',
            'Prev week gap':     'bpm_prev_gap',
            'WOW BPM':           'bpm_wow',
            'Delta Reason':      'delta_reason',
            'Recovery Plan':     'recovery_plan',
            'ADH':               'adh',
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
            'Column1':            'adh',
            'Column 1':           'adh',
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
            'Column1':             'adh',
            'Column 1':            'adh',
        },
    },
}

KNOWN_ADH_NAMES = [
    "Kumar", "Prachi", "Pradeep", "Himadri", "Ajit Katankot",
    "Manu", "RK", "Manish", "VSV Ramesh", "LK", "Solomon",
]

def _normalize_text(value) -> str:
    if value is None:
        return ''
    if isinstance(value, (pd.Series, pd.Index)):
        if len(value) == 0:
            return ''
        value = value.iloc[0]
    text = str(value).strip().lower()
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return ' '.join(text.split())


def _normalize_column_label(value: str) -> str:
    return _normalize_text(value)

KNOWN_ADH_MAP = {
    _normalize_text(name): name
    for name in KNOWN_ADH_NAMES
}


def _text(val):
    """Clean text value."""
    if isinstance(val, (pd.Series, pd.Index)):
        if len(val) == 0:
            return ''
        val = val.iloc[0]
    if pd.isna(val):
        return ''
    text = str(val).strip()
    if not text or text.lower() in ('nan', 'none', 'na', 'n/a', '--'):
        return ''
    return text


def _normalize_adh_value(value):
    if isinstance(value, (pd.Series, pd.Index)):
        if len(value) == 0:
            return ''
        value = value.iloc[0]
    if pd.isna(value):
        return ''

    text = _text(value)
    if not text:
        return ''

    normalized = _normalize_column_label(text)
    # Remove common ADH prefixes or labels from the raw value.
    normalized = re.sub(
        r'^(adh|account delivery head|delivery head|account head|adh[:\-]?|adh\s+)',
        '', normalized,
        flags=re.IGNORECASE
    ).strip()

    if normalized in KNOWN_ADH_MAP:
        return KNOWN_ADH_MAP[normalized]

    for known_norm, canonical in KNOWN_ADH_MAP.items():
        if known_norm in normalized or normalized in known_norm:
            return canonical

    return text


def _find_sheet_name(xls, metric: str) -> str:
    requested = _normalize_text(_SHEET_MAP[metric]['sheet'])
    sheets = list(xls.sheet_names)
    lower_map = {name.lower(): name for name in sheets}
    if requested in lower_map:
        return lower_map[requested]

    metric_tokens = {
        'RU': ['ru', 'resource', 'utilization', 'utilisation'],
        'RD': ['rd', 'resource', 'deallocation', 'de-alloc', 'dealloc'],
        'Netadd': ['net', 'netadd', 'net add', 'net addition', 'netaddition'],
    }.get(metric, [requested])

    for name in sheets:
        norm = _normalize_text(name)
        if any(token in norm for token in metric_tokens):
            return name

    for name in sheets:
        if requested in _normalize_text(name):
            return name

    return xls.sheet_names[0]


def _find_header_row(xls, sheet_name: str, normalized_labels: dict) -> int | None:
    preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=20)
    best = (-1, None)
    for row_index, row in preview.iterrows():
        matches = 0
        for cell in row:
            if _match_column_name(cell, normalized_labels):
                matches += 1
        if matches > best[0]:
            best = (matches, row_index)
    return best[1] if best[0] >= 3 else None


def _match_column_name(raw_label: str, known_labels: dict) -> str | None:
    normalized = _normalize_column_label(raw_label)
    if normalized in known_labels:
        return known_labels[normalized]

    tokens = set(normalized.split())

    def has_qtr_token(tokens):
        return any(tok.startswith(('qtr', 'q1', 'q2', 'q3', 'q4', 'qtd', 'quarter', 'quarterly')) for tok in tokens)

    if 'account' in tokens or 'client' in tokens or 'name' in tokens and 'account' in tokens:
        return 'account'
    if 'recovery' in tokens and 'plan' in tokens:
        return 'recovery_plan'
    if 'delta' in tokens and 'reason' in tokens:
        return 'delta_reason'
    if ('adh' in tokens
            or ('delivery' in tokens and 'head' in tokens)
            or ('account' in tokens and 'head' in tokens)
            or 'accountdeliveryhead' in normalized
            or 'deliveryhead' in normalized
            or 'column1' in tokens
            or ('column' in tokens and ('1' in tokens or 'one' in tokens))
            or ('col' in tokens and '1' in tokens)):
        return 'adh'
    if 'bpm' in tokens:
        if 'gap' in tokens:
            return 'bpm_gap'
        if 'wow' in tokens:
            return 'bpm_wow'
        if 'prev' in tokens and 'gap' in tokens:
            return 'bpm_prev_gap'
        if 'plan' in tokens and has_qtr_token(tokens):
            return 'bpm_plan_qtr'
        if 'plan' in tokens:
            return 'bpm_wk_plan'
        if 'act' in tokens or 'actual' in tokens:
            return 'bpm_wk_act'
    if 'wow' in tokens or 'woow' in tokens:
        return 'wow'
    if 'prev' in tokens and 'gap' in tokens:
        return 'prev_gap'
    if 'gap' in tokens:
        return 'gap'
    if 'plan' in tokens and has_qtr_token(tokens):
        return 'plan_qtr'
    if 'plan' in tokens:
        return 'wk_plan'
    if 'act' in tokens or 'actual' in tokens:
        return 'wk_act'
    return None



def _clean_number(val):
    if isinstance(val, (pd.Series, pd.Index)):
        if len(val) == 0:
            return None
        val = val.iloc[0]
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
    # Preserve RD values as provided in the Excel source.
    return val


def _invert_if_rd(metric: str, val):
    # Preserve RD WOW values as provided in the Excel source.
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
    sheet_name = _find_sheet_name(xls, metric)
    normalized_map = {
        _normalize_column_label(raw): dest
        for raw, dest in cfg['col_map'].items()
    }

    df = None
    header_row = _find_header_row(xls, sheet_name, normalized_map)
    if header_row is not None:
        candidate = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
        rename_map = {}
        for col in candidate.columns:
            mapped = _match_column_name(col, normalized_map)
            if mapped:
                rename_map[col] = mapped
        candidate = candidate.rename(columns=rename_map)
        if candidate.columns.duplicated().any():
            candidate = candidate.loc[:, ~candidate.columns.duplicated()]
        if 'account' in candidate.columns:
            df = candidate

    if df is None:
        for header in (0, 1, 2):
            candidate = pd.read_excel(xls, sheet_name=sheet_name, header=header)
            rename_map = {}
            for col in candidate.columns:
                mapped = _match_column_name(col, normalized_map)
                if mapped:
                    rename_map[col] = mapped
            candidate = candidate.rename(columns=rename_map)
            if candidate.columns.duplicated().any():
                candidate = candidate.loc[:, ~candidate.columns.duplicated()]
            if 'account' in candidate.columns:
                df = candidate
                break
    if df is None or 'account' not in df.columns:
        raise ValueError(f"Unable to locate the account column for metric '{metric}' in sheet '{sheet_name}'")

    # normalize account values to text and drop invalid rows
    df['account'] = df['account'].apply(lambda v: _text(v))
    df = df[df['account'] != '']
    df = df[~df['account'].str.lower().str.contains('grand total|total', na=False)]

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
            'adh':           _normalize_adh_value(row.get('adh')),
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

        plan_qtr = _safe(ru.get('plan_qtr')) + _safe(rd.get('plan_qtr'))
        wk_plan = _safe(ru.get('wk_plan')) + _safe(rd.get('wk_plan'))
        wk_act = _safe(ru.get('wk_act')) + _safe(rd.get('wk_act'))
        gap = wk_act - wk_plan
        prev_gap = _safe(ru.get('prev_gap')) + _safe(rd.get('prev_gap'))
        wow = _safe(ru.get('wow')) + _safe(rd.get('wow'))
        bpm_plan_qtr = _safe(ru.get('bpm_plan_qtr')) + _safe(rd.get('bpm_plan_qtr'))
        bpm_wk_plan = _safe(ru.get('bpm_wk_plan')) + _safe(rd.get('bpm_wk_plan'))
        bpm_wk_act = _safe(ru.get('bpm_wk_act')) + _safe(rd.get('bpm_wk_act'))
        bpm_gap = bpm_wk_act - bpm_wk_plan
        bpm_prev_gap = _safe(ru.get('bpm_prev_gap')) + _safe(rd.get('bpm_prev_gap'))
        bpm_wow = _safe(ru.get('bpm_wow')) + _safe(rd.get('bpm_wow'))

        rec = {
            'account': account,
            'metric': 'Netadd',
            'adh': ru.get('adh') or rd.get('adh') or '',
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


def _merge_actual_netadd(derived_records: list[dict], actual_records: list[dict]) -> list[dict]:
    actual_map = {r['account'].strip().upper(): r for r in actual_records}
    merged = []
    seen_accounts = set()

    for derived in derived_records:
        account_key = derived['account'].strip().upper()
        actual = actual_map.get(account_key)
        if actual is not None:
            merged.append({**derived, **actual})
        else:
            merged.append(derived)
        seen_accounts.add(account_key)

    for actual in actual_records:
        account_key = actual['account'].strip().upper()
        if account_key not in seen_accounts:
            merged.append(actual)

    return merged


def get_all_data(filepath: str, metrics: list, accounts_filter) -> dict:
    result = {}
    for m in metrics:
        if m == 'Netadd':
            continue
        result[m] = load_metric(filepath, m, accounts_filter)

    if 'Netadd' in metrics:
        if 'RU' in result and 'RD' in result:
            derived_netadd = _derive_netadd_records(result['RU'], result['RD'])
            try:
                actual_netadd = load_metric(filepath, 'Netadd', accounts_filter)
            except Exception:
                result['Netadd'] = derived_netadd
            else:
                result['Netadd'] = _merge_actual_netadd(derived_netadd, actual_netadd)
        else:
            result['Netadd'] = load_metric(filepath, 'Netadd', accounts_filter)

    return result


def group_records_by_adh(records: list[dict]) -> dict[str, list[dict]]:
    adh_groups = {}
    for rec in records:
        adh = (rec.get('adh') or '').strip()
        if not adh:
            continue
        adh_groups.setdefault(adh, []).append(rec)
    return adh_groups
