import io
import os
import re
import sys
import tempfile
import shutil
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from src.data_reader   import get_all_data, load_metric
from src.chart_builder import build_bar_chart, build_overview_chart, build_group_comparison_chart
from src.slide_builder import (
    create_presentation, add_title_slide, add_account_metric_slide,
    add_metric_overview_slide, add_summary_slide, add_grouped_accounts_slide,
)

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Wipro Slide Generator",
    page_icon="📊",
    layout="wide",
)


def infer_week_quarter_from_filename(filename: str) -> tuple[str | None, str | None]:
    if not filename:
        return None, None
    name = filename.lower()
    week = None
    quarter = None
    wk_match = re.search(r'\b(?:week|wk)[\s_-]*0*([0-9]{1,2})\b', name)
    if wk_match:
        week = f"WK{int(wk_match.group(1)):02d}"
    q_match = re.search(r"\b(q[1-4])(?:[\'’]?(\d{2,4}))?\b", name)
    if q_match:
        quarter = q_match.group(1).upper()
        if q_match.group(2):
            quarter = f"{quarter}'{q_match.group(2)}"
    return week, quarter

# ── custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 1.2rem 2rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; color: white; }
    .main-header p  { margin: 0.2rem 0 0; opacity: 0.8; font-size: 0.9rem; color: #BDD7EE; }

    .section-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        border-left: 4px solid #2E75B6;
        margin-bottom: 1rem;
    }
    .metric-badge {
        display: inline-block;
        background: #1F3864;
        color: white;
        padding: 2px 10px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: bold;
        margin-right: 4px;
    }
    .stat-box {
        text-align: center;
        padding: 0.6rem;
        border-radius: 6px;
        background: white;
        border: 1px solid #dee2e6;
    }
    .stat-box .num  { font-size: 1.6rem; font-weight: 700; color: #1F3864; }
    .stat-box .lbl  { font-size: 0.72rem; color: #666; }
    .gap-red   { color: #C00000; font-weight: bold; }
    .gap-green { color: #375623; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ── header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📊 Wipro Weekly Slide Generator</h1>
  <p>Upload your xlsx, configure options, and download a polished PowerPoint in seconds.</p>
</div>
""", unsafe_allow_html=True)

# ── file uploader — full width (avoids column rendering issues) ───────────────
st.markdown("### 📁 Upload Data File")
uploaded = st.file_uploader(
    "Select your weekly Netadd_BPM_ADH xlsx file",
    type=["xlsx"],
    help="Upload the weekly xlsx exported from the BPM tracker",
)
if uploaded:
    st.success(f"✅  **{uploaded.name}**  —  {uploaded.size / 1024:.0f} KB loaded")

week_label_hint, quarter_label_hint = infer_week_quarter_from_filename(uploaded.name if uploaded else "")

# ── sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="background:#1F3864;padding:10px 16px;border-radius:8px;margin-bottom:8px;">
      <span style="color:white;font-size:1.2rem;font-weight:700;letter-spacing:1px;">WIPRO</span>
      <span style="color:#BDD7EE;font-size:0.75rem;margin-left:6px;">Slide Generator</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### ⚙️ Generation Settings")

    week_label_default = "WK08"
    quarter_label_default = "Q1'27"
    if uploaded:
        if week_label_hint:
            week_label_default = week_label_hint
        if quarter_label_hint:
            quarter_label_default = quarter_label_hint

    if 'week_label' not in st.session_state:
        st.session_state.week_label = week_label_default
    if 'quarter_label' not in st.session_state:
        st.session_state.quarter_label = quarter_label_default

    week_label = st.text_input("Week Label", value=st.session_state.week_label, key="week_label")
    quarter_label = st.text_input("Quarter Label", value=st.session_state.quarter_label, key="quarter_label")

    if uploaded and week_label_hint:
        st.info(f"Detected week label from filename: {week_label_hint}")

    st.markdown("**Metrics to include**")
    inc_ru     = st.checkbox("RU — Resource Utilisation", value=True)
    inc_rd     = st.checkbox("RD — Resource Deallocation", value=True)
    inc_net    = st.checkbox("Netadd — Net Addition", value=True)

    selected_metrics = [m for m, on in [("RU", inc_ru), ("RD", inc_rd), ("Netadd", inc_net)] if on]

    layout = st.selectbox(
        "Slide Layout",
        options=["per_account", "per_metric", "per_account_combined"],
        format_func=lambda x: {
            "per_account":          "Per Account (one slide each)",
            "per_metric":           "Per Metric (all accounts, ranked table)",
            "per_account_combined": "Per Account — Combined (all 3 metrics)",
        }[x],
    )

    chart_type = st.selectbox(
        "Chart Type",
        options=["all", "bar", "kpi"],
        format_func=lambda x: {
            "all": "Bar + KPI Cards (Recommended)",
            "bar": "Bar Chart Only",
            "kpi": "KPI Cards Only",
        }[x],
    )

    if 'output_name' not in st.session_state:
        st.session_state.output_name = f"{week_label}_Slides.pptx"
    output_name = st.text_input("Output filename", value=st.session_state.output_name, key="output_name")

st.markdown("---")

# ── load data when file is uploaded ───────────────────────────────────────────
data_loaded = False
all_data    = {}
all_accounts_by_metric = {}

if uploaded and selected_metrics:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded.getvalue())
        tmp_xlsx = tmp.name

    try:
        for m in selected_metrics:
            all_data[m] = load_metric(tmp_xlsx, m, 'all')
            all_accounts_by_metric[m] = sorted({r['account'] for r in all_data[m]})
        data_loaded = True
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        st.stop()


# ── account groups ────────────────────────────────────────────────────────────
selected_accounts = "all"
groups            = []          # list of {"name": str, "accounts": [...]}
also_individual   = True

if data_loaded:
    all_union = sorted(set(a for accs in all_accounts_by_metric.values() for a in accs))

    # initialise session state
    if 'groups' not in st.session_state:
        st.session_state.groups = []

    col_filter, col_preview = st.columns([1, 2], gap="large")

    with col_filter:
        # ── individual account filter ──────────────────────────────────────────
        st.markdown("### 🏢 Individual Slides")
        filter_mode = st.radio("", ["All accounts", "Select specific"],
                               horizontal=True, key="filter_mode")
        if filter_mode == "Select specific":
            selected_accounts = st.multiselect(
                "Accounts for individual slides",
                options=all_union,
                default=all_union[:5],
            )
            if not selected_accounts:
                selected_accounts = "all"
        else:
            selected_accounts = "all"

        if st.session_state.groups:
            also_individual = st.checkbox(
                "Also generate individual slides for grouped accounts",
                value=True,
            )

        st.markdown("---")

        # ── named group builder ────────────────────────────────────────────────
        st.markdown("### 🏷️ Account Groups")
        st.caption("Each group generates **3 slides** (RU · RD · Netadd) with the group name as the title.")

        # render existing groups
        to_delete = None
        for i, grp in enumerate(st.session_state.groups):
            with st.expander(
                f"{'📌 ' + grp['name'] if grp['name'] else f'Group {i+1} (unnamed)'}",
                expanded=True
            ):
                c1, c2 = st.columns([5, 1])
                with c1:
                    new_name = st.text_input(
                        "Group tag / name",
                        value=grp['name'],
                        key=f"gname_{i}",
                        placeholder="e.g. Tier 1, Healthcare, APAC...",
                    )
                    st.session_state.groups[i]['name'] = new_name
                with c2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️", key=f"gdel_{i}", help="Remove this group"):
                        to_delete = i

                picked = st.multiselect(
                    "Accounts in this group",
                    options=all_union,
                    default=[a for a in grp['accounts'] if a in all_union],
                    key=f"gaccs_{i}",
                )
                st.session_state.groups[i]['accounts'] = picked

                if picked:
                    st.caption(f"{len(picked)} accounts · generates {len(selected_metrics)} slides")

        if to_delete is not None:
            st.session_state.groups.pop(to_delete)
            st.rerun()

        if st.button("➕  Add Group", use_container_width=True):
            st.session_state.groups.append({"name": "", "accounts": []})
            st.rerun()

    groups = [g for g in st.session_state.groups if g['accounts']]


# ── data preview ───────────────────────────────────────────────────────────────
    with col_preview:
        st.markdown("### 📋 Data Preview")

        tab_labels = [f"{m} ({len(all_data[m])})" for m in selected_metrics]
        tabs = st.tabs(tab_labels)

        for tab, metric in zip(tabs, selected_metrics):
            with tab:
                records = all_data[metric]
                if selected_accounts != "all":
                    upper = [a.upper() for a in selected_accounts]
                    records = [r for r in records if r['account'].upper() in upper]

                if not records:
                    st.info("No accounts match the filter.")
                    continue

                # summary stats
                s1, s2, s3, s4 = st.columns(4)
                gaps    = [r['gap'] for r in records if r['gap'] is not None]
                bpm_gaps = [r['bpm_gap'] for r in records if r['bpm_gap'] is not None]
                neg_gaps = sum(1 for g in gaps if g < 0)

                s1.markdown(f'<div class="stat-box"><div class="num">{len(records)}</div><div class="lbl">Accounts</div></div>', unsafe_allow_html=True)
                s2.markdown(f'<div class="stat-box"><div class="num">{sum(gaps):.0f}</div><div class="lbl">Total Gap</div></div>', unsafe_allow_html=True)
                s3.markdown(f'<div class="stat-box"><div class="num gap-red">{neg_gaps}</div><div class="lbl">Behind Plan</div></div>', unsafe_allow_html=True)
                s4.markdown(f'<div class="stat-box"><div class="num">{sum(bpm_gaps):.0f}</div><div class="lbl">Total BPM Gap</div></div>', unsafe_allow_html=True)

                st.markdown("")

                # table
                rows = []
                for r in records:
                    gap_v    = r['gap']
                    bpm_gap_v = r['bpm_gap']
                    rows.append({
                        "Account":      r['account'],
                        "Plan QTR":     r['fmt_plan_qtr'],
                        "WK Plan":      r['fmt_wk_plan'],
                        "WK Act":       r['fmt_wk_act'],
                        "Gap":          r['fmt_gap'],
                        "WoW":          r['fmt_wow'],
                        "BPM Plan QTR": r['fmt_bpm_plan_qtr'],
                        "BPM WK Act":   r['fmt_bpm_wk_act'],
                        "BPM Gap":      r['fmt_bpm_gap'],
                        "BPM WoW":      r['fmt_bpm_wow'],
                        "Delta Reason": (r['delta_reason'] or '')[:60],
                    })

                df_preview = pd.DataFrame(rows)

                def color_gap(val):
                    try:
                        v = float(val)
                        if v < 0:
                            return 'color: #C00000; font-weight: bold'
                        if v > 0:
                            return 'color: #375623; font-weight: bold'
                    except (ValueError, TypeError):
                        pass
                    return ''

                styled = df_preview.style.map(
                    color_gap, subset=["Gap", "BPM Gap", "WoW", "BPM WoW"]
                )
                st.dataframe(styled, use_container_width=True, height=320)


# ── generate button ────────────────────────────────────────────────────────────
st.markdown("---")

if not data_loaded:
    st.info("⬆️ Upload an xlsx file to get started.")
elif not selected_metrics:
    st.warning("Select at least one metric in the sidebar.")
else:
    group_slides  = len(groups) * len(selected_metrics)
    indiv_count   = sum(
        len(all_data[m]) if layout in ("per_account", "per_account_combined") else 1
        for m in selected_metrics
    )
    n_slides_est  = indiv_count + group_slides + 2  # title + summary

    col_btn, col_info = st.columns([1, 3])
    with col_info:
        st.markdown(f"**Estimated slides:** ~{n_slides_est}  "
                    f"| **Metrics:** {', '.join(selected_metrics)}  "
                    f"| **Layout:** {layout}  "
                    f"| **Chart:** {chart_type}")

    with col_btn:
        generate = st.button("▶  Generate Slides", type="primary", use_container_width=True)

    if generate:
        tmp_dir  = tempfile.mkdtemp(prefix="wipro_gui_")
        out_path = os.path.join(tmp_dir, output_name)

        progress = st.progress(0, text="Starting…")
        status   = st.empty()

        try:
            # filter by selected accounts
            filtered = {}
            for m in selected_metrics:
                recs = all_data[m]
                if selected_accounts != "all":
                    upper = [a.upper() for a in selected_accounts]
                    recs  = [r for r in recs if r['account'].upper() in upper]
                filtered[m] = recs

            prs = create_presentation()
            add_title_slide(prs, week_label, quarter_label)
            add_summary_slide(prs, filtered, week_label, quarter_label)

            chart_tmp   = tempfile.mkdtemp(prefix="wipro_charts_")
            all_grouped_uppers = {
                a.upper()
                for g in groups
                for a in g['accounts']
            }
            group_steps = len(groups) * len(selected_metrics)
            indiv_steps = sum(
                len(filtered[m]) if layout in ("per_account", "per_account_combined") else 1
                for m in selected_metrics
            )
            total_steps = max(indiv_steps + group_steps, 1)
            step = 0

            # ── named group slides (RU → RD → Netadd per group) ──────────────
            for grp in groups:
                grp_name   = grp['name'] or 'Group'
                grp_upper  = [a.upper() for a in grp['accounts']]

                for metric in selected_metrics:
                    grp_recs = [r for r in filtered[metric]
                                if r['account'].upper() in grp_upper]
                    if not grp_recs:
                        step += 1
                        continue
                    status.markdown(
                        f"Building **{grp_name}** — {metric} "
                        f"({len(grp_recs)} accounts)…"
                    )
                    grp_chart = build_group_comparison_chart(grp_recs, metric, chart_tmp)
                    add_grouped_accounts_slide(
                        prs, grp_recs, metric,
                        week_label, quarter_label,
                        chart_path=grp_chart,
                        group_name=grp_name,
                    )
                    step += 1
                    progress.progress(min(step / total_steps, 1.0),
                                      text=f"{grp_name} — {metric}")

            # ── individual slides ─────────────────────────────────────────────
            for metric in selected_metrics:
                records = filtered[metric]

                # skip grouped accounts if "also individual" is off
                if groups and not also_individual:
                    indiv_recs = [r for r in records
                                  if r['account'].upper() not in all_grouped_uppers]
                else:
                    indiv_recs = records

                if layout == "per_metric":
                    status.markdown(f"Building overview slide for **{metric}**…")
                    chart_path = None
                    if chart_type in ("bar", "all"):
                        chart_path = build_overview_chart(indiv_recs, metric, chart_tmp)
                    add_metric_overview_slide(prs, indiv_recs, metric,
                                              week_label, quarter_label, chart_path)
                    step += 1
                    progress.progress(min(step / total_steps, 1.0),
                                      text=f"Overview — {metric}")
                else:
                    for rec in indiv_recs:
                        status.markdown(
                            f"Building **{metric}** — {rec['account']} "
                            f"({step + 1}/{total_steps})"
                        )
                        bar_path = None
                        if chart_type in ("bar", "all"):
                            bar_path = build_bar_chart(rec, chart_tmp)
                        add_account_metric_slide(
                            prs, rec, week_label, quarter_label,
                            chart_type=chart_type,
                            bar_path=bar_path,
                        )
                        step += 1
                        progress.progress(min(step / total_steps, 1.0),
                                          text=f"{metric}: {rec['account']}")

            prs.save(out_path)
            shutil.rmtree(chart_tmp, ignore_errors=True)

            progress.progress(1.0, text="Done!")
            status.success(f"✅ Generated **{len(prs.slides)} slides** successfully!")

            with open(out_path, "rb") as f:
                pptx_bytes = f.read()

            st.download_button(
                label=f"⬇️  Download {output_name}",
                data=pptx_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Generation failed: {e}")
            raise
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            if 'tmp_xlsx' in dir():
                try:
                    os.unlink(tmp_xlsx)
                except Exception:
                    pass
