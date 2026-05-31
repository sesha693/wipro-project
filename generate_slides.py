#!/usr/bin/env python3
"""
Wipro Weekly Slide Generator
Usage:  python3.13 generate_slides.py [--config config.yaml]
"""

import os
import sys
import tempfile
import shutil
import argparse
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).parent))
from src.data_reader  import get_all_data
from src.chart_builder import build_bar_chart, build_gap_trend_chart, build_overview_chart
from src.slide_builder import (
    create_presentation, add_title_slide, add_account_metric_slide,
    add_metric_overview_slide, add_summary_slide,
)


def load_config(path: str) -> dict:
    with open(path) as f:
        cfg = yaml.safe_load(f)
    return cfg


def main():
    parser = argparse.ArgumentParser(description='Generate Wipro weekly slides from xlsx')
    parser.add_argument('--config', default='config.yaml',
                        help='Path to config.yaml (default: config.yaml)')
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.exists():
        print(f'ERROR: config file not found: {config_path}')
        sys.exit(1)

    cfg = load_config(str(config_path))
    base_dir = config_path.parent

    input_file   = base_dir / cfg['input_file']
    output_file  = base_dir / cfg['output_file']
    week_label   = cfg.get('week_label',   'WK??')
    quarter_label = cfg.get('quarter_label', 'Q?')
    accounts     = cfg.get('accounts', 'all')
    metrics      = cfg.get('metrics',  ['RU', 'RD', 'Netadd'])
    layout       = cfg.get('layout',   'per_account')
    chart_type   = cfg.get('chart_type', 'all')

    if not input_file.exists():
        print(f'ERROR: input file not found: {input_file}')
        sys.exit(1)

    print(f'Reading data from {input_file.name} ...')
    all_data = get_all_data(str(input_file), metrics, accounts)

    total_accounts = {m: len(all_data[m]) for m in metrics}
    for m, n in total_accounts.items():
        print(f'  {m}: {n} accounts loaded')

    tmp_dir = tempfile.mkdtemp(prefix='wipro_slides_')
    try:
        prs = create_presentation()
        add_title_slide(prs, week_label, quarter_label)

        # summary slide (always added)
        add_summary_slide(prs, all_data, week_label, quarter_label)

        if layout == 'per_metric':
            for metric in metrics:
                records = all_data[metric]
                if not records:
                    continue
                chart_path = None
                if chart_type in ('bar', 'all'):
                    chart_path = build_overview_chart(records, metric, tmp_dir)
                add_metric_overview_slide(prs, records, metric,
                                          week_label, quarter_label, chart_path)
                print(f'  Added overview slide: {metric} ({len(records)} accounts)')

        elif layout in ('per_account', 'per_account_combined'):
            for metric in metrics:
                for rec in all_data[metric]:
                    bar_path   = None
                    trend_path = None

                    if chart_type in ('bar', 'all'):
                        bar_path = build_bar_chart(rec, tmp_dir)

                    add_account_metric_slide(
                        prs, rec, week_label, quarter_label,
                        chart_type=chart_type,
                        bar_path=bar_path,
                        trend_path=trend_path,
                    )

                total = len(all_data[metric])
                print(f'  Added {total} slides for {metric}')

        prs.save(str(output_file))
        print(f'\nDone! Saved to: {output_file}')
        print(f'Total slides: {len(prs.slides)}')

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == '__main__':
    main()
