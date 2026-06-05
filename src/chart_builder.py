import os
import tempfile
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# Wipro brand palette
C_NAVY  = '#1F3864'
C_BLUE  = '#2E75B6'
C_LBLUE = '#BDD7EE'
C_RED   = '#C00000'
C_GREEN = '#70AD47'
C_GRAY  = '#808080'


def _bar_color(val):
    if val is None:
        return C_GRAY
    return C_BLUE if val >= 0 else C_RED


def _chart_value(metric: str, value):
    if value is None:
        return 0
    try:
        v = float(value)
    except (TypeError, ValueError):
        return 0
    return abs(v) if metric == 'RD' else v


def build_bar_chart(rec: dict, tmp_dir: str) -> str:
    """
    Grouped bar chart: main metric vs BPM metric.
    Returns path to saved PNG.
    """
    metric = rec['metric']

    groups = ['Plan QTR', 'WK Plan', 'WK Act']
    main_vals = [_chart_value(metric, rec.get('plan_qtr')), _chart_value(metric, rec.get('wk_plan')), _chart_value(metric, rec.get('wk_act'))]
    bpm_vals  = [_chart_value(metric, rec.get('bpm_plan_qtr')), _chart_value(metric, rec.get('bpm_wk_plan')), _chart_value(metric, rec.get('bpm_wk_act'))]

    x = np.arange(len(groups))
    width = 0.35

    fig, ax = plt.subplots(figsize=(6.2, 4.0))
    fig.patch.set_facecolor('white')
    fig.subplots_adjust(top=0.88, bottom=0.12, left=0.1, right=0.97)

    bars1 = ax.bar(x - width/2, main_vals, width, label=metric,
                   color=[_bar_color(v) for v in main_vals], alpha=0.9, zorder=3)
    bars2 = ax.bar(x + width/2, bpm_vals,  width, label=f'{metric} BPM',
                   color=[_bar_color(v) for v in bpm_vals], alpha=0.55, zorder=3)

    # value labels on bars
    for bar in list(bars1) + list(bars2):
        h = bar.get_height()
        va = 'bottom' if h >= 0 else 'top'
        offset = 0.5 if h >= 0 else -0.5
        ax.text(bar.get_x() + bar.get_width()/2, h + offset,
                f'{h:.1f}' if h != int(h) else str(int(h)),
                ha='center', va=va, fontsize=7.5, fontweight='bold',
                color='#222222')

    ax.set_xticks(x)
    ax.set_xticklabels(groups, fontsize=9)
    ax.set_ylabel('Value', fontsize=8, color='#444444')
    ax.set_title(f'{metric} — Plan vs Actual vs BPM', fontsize=10,
                 fontweight='bold', color=C_NAVY, pad=8)
    ax.axhline(0, color='#444444', linewidth=0.8, zorder=2)
    ax.grid(axis='y', linestyle='--', alpha=0.4, zorder=1)
    ax.set_facecolor('#F8F9FA')
    ax.tick_params(colors='#444444')
    for spine in ax.spines.values():
        spine.set_edgecolor('#CCCCCC')

    patch1 = mpatches.Patch(color=C_BLUE,  label=metric, alpha=0.9)
    patch2 = mpatches.Patch(color=C_BLUE,  label=f'{metric} BPM', alpha=0.55)
    ax.legend(handles=[patch1, patch2], fontsize=7.5, loc='best',
              framealpha=0.8, edgecolor='#CCCCCC')

    path = os.path.join(tmp_dir, f'{rec["account"]}_{metric}_bar.png')
    fig.savefig(path, dpi=140, bbox_inches=None)
    plt.close(fig)
    return path


def build_gap_trend_chart(rec: dict, tmp_dir: str) -> str:
    """
    Mini horizontal bar showing current vs previous gap for both metric and BPM.
    """
    metric = rec['metric']
    labels = [f'{metric} Gap', f'{metric} BPM Gap']
    current = [rec.get('gap') or 0, rec.get('bpm_gap') or 0]
    prev    = [rec.get('prev_gap') or 0, rec.get('bpm_prev_gap') or 0]

    fig, ax = plt.subplots(figsize=(5.8, 2.8))
    fig.patch.set_facecolor('white')
    fig.subplots_adjust(top=0.88, bottom=0.14, left=0.14, right=0.97)

    y = np.arange(len(labels))
    ax.barh(y + 0.2, current, 0.35, label='Current WK',
            color=[_bar_color(v) for v in current], alpha=0.9, zorder=3)
    ax.barh(y - 0.2, prev,    0.35, label='Prev WK',
            color=C_LBLUE, alpha=0.8, zorder=3)

    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8.5)
    ax.set_xlabel('Gap Value', fontsize=8)
    ax.axvline(0, color='#444444', linewidth=0.8)
    ax.set_title('Gap Trend: Current vs Prev Week', fontsize=9,
                 fontweight='bold', color=C_NAVY)
    ax.grid(axis='x', linestyle='--', alpha=0.4, zorder=1)
    ax.set_facecolor('#F8F9FA')
    ax.legend(fontsize=7.5, framealpha=0.8)
    for spine in ax.spines.values():
        spine.set_edgecolor('#CCCCCC')

    path = os.path.join(tmp_dir, f'{rec["account"]}_{metric}_trend.png')
    fig.savefig(path, dpi=140, bbox_inches=None)
    plt.close(fig)
    return path


def build_group_comparison_chart(records: list, metric: str, tmp_dir: str) -> str:
    """
    Grouped bar chart comparing Gap and BPM Gap across tagged accounts.
    Used in the group/comparison slide.
    """
    accounts = [r['account'] for r in records]
    gaps     = [_chart_value(metric, r.get('gap'))     for r in records]
    bpm_gaps = [_chart_value(metric, r.get('bpm_gap')) for r in records]

    x = np.arange(len(accounts))
    width = 0.35

    fig, ax = plt.subplots(figsize=(6.0, max(3.2, len(accounts) * 0.55)))
    fig.patch.set_facecolor('white')
    fig.subplots_adjust(top=0.88, bottom=0.18, left=0.08, right=0.97)

    bars1 = ax.bar(x - width/2, gaps,     width, label=f'{metric} Gap',
                   color=[_bar_color(v) for v in gaps], alpha=0.9, zorder=3)
    bars2 = ax.bar(x + width/2, bpm_gaps, width, label=f'{metric} BPM Gap',
                   color=[_bar_color(v) for v in bpm_gaps], alpha=0.5, zorder=3)

    for bar in list(bars1) + list(bars2):
        h = bar.get_height()
        va = 'bottom' if h >= 0 else 'top'
        offset = 0.3 if h >= 0 else -0.3
        ax.text(bar.get_x() + bar.get_width() / 2, h + offset,
                f'{h:.1f}' if h != int(h) else str(int(h)),
                ha='center', va=va, fontsize=7, fontweight='bold', color='#222222')

    ax.set_xticks(x)
    ax.set_xticklabels(accounts, fontsize=8, rotation=15, ha='right')
    ax.axhline(0, color='#444444', linewidth=0.8, zorder=2)
    ax.set_title(f'{metric} — Gap Comparison', fontsize=10,
                 fontweight='bold', color=C_NAVY, pad=8)
    ax.grid(axis='y', linestyle='--', alpha=0.4, zorder=1)
    ax.set_facecolor('#F8F9FA')
    ax.legend(fontsize=7.5, framealpha=0.8, loc='best')
    for spine in ax.spines.values():
        spine.set_edgecolor('#CCCCCC')

    label = '_'.join(r['account'][:6] for r in records[:3])
    path = os.path.join(tmp_dir, f'group_{metric}_{label}.png')
    fig.savefig(path, dpi=140, bbox_inches=None)
    plt.close(fig)
    return path


def build_overview_chart(records: list, metric: str, tmp_dir: str) -> str:
    """
    Horizontal bar chart of WK Act for all accounts — used in per_metric layout.
    """
    accounts = [r['account'] for r in records]
    acts     = [_chart_value(metric, r.get('wk_act')) for r in records]
    plans    = [_chart_value(metric, r.get('wk_plan')) for r in records]

    fig, ax = plt.subplots(figsize=(9, max(4.0, len(accounts) * 0.35)))
    fig.subplots_adjust(top=0.92, bottom=0.08, left=0.22, right=0.97)
    fig.patch.set_facecolor('white')

    y = np.arange(len(accounts))
    ax.barh(y + 0.2, acts,  0.35, label='WK Act',
            color=[_bar_color(v) for v in acts], alpha=0.9, zorder=3)
    ax.barh(y - 0.2, plans, 0.35, label='WK Plan',
            color=C_LBLUE, alpha=0.85, zorder=3)

    ax.set_yticks(y)
    ax.set_yticklabels(accounts, fontsize=7.5)
    ax.axvline(0, color='#444444', linewidth=0.8)
    ax.set_xlabel('Value', fontsize=8)
    ax.set_title(f'{metric} — WK Plan vs Act (All Accounts)', fontsize=10,
                 fontweight='bold', color=C_NAVY, pad=8)
    ax.grid(axis='x', linestyle='--', alpha=0.4, zorder=1)
    ax.set_facecolor('#F8F9FA')
    ax.legend(fontsize=8, framealpha=0.8, loc='lower right')
    for spine in ax.spines.values():
        spine.set_edgecolor('#CCCCCC')

    path = os.path.join(tmp_dir, f'overview_{metric}.png')
    fig.savefig(path, dpi=130, bbox_inches=None)
    plt.close(fig)
    return path
