"""
ABC Аналіз по магазинах — Streamlit додаток
============================================
"""

import os
import tempfile
import json
import base64
import urllib.request
from datetime import date, datetime

import streamlit as st
import pandas as pd
import numpy as np
import anthropic

from shop_engine import (
    load_all_months, build_promo_map,
    build_global_abc, build_group_abc,
    build_shop_abc, build_shop_summary,
    build_monthly_by_shop, build_monthly_by_group,
)

st.set_page_config(
    page_title="ABC Магазини",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
h2{font-size:1.05rem!important;font-weight:600!important;color:#4f7cff!important;}
[data-testid="metric-container"]{background:#13161e;border:1px solid #252a38;border-radius:10px;padding:14px 18px!important;}
[data-testid="stMetricValue"]{font-size:1.3rem!important;font-weight:700!important;}
.stDownloadButton>button{border-radius:8px!important;border:1px solid #252a38!important;background:#13161e!important;}
.stDownloadButton>button:hover{border-color:#4f7cff!important;}
</style>
""", unsafe_allow_html=True)


# ── HELPERS ───────────────────────────────────────────────────────────────────
def save_upload(f):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(f.read())
    tmp.flush()
    return tmp.name

def get_secret(key, default=""):
    try:
        return st.secrets[key]
    except Exception:
        return default

def fmt_usd(v): return f"${float(v):,.0f}"
def fmt_pct(v): return f"{float(v):.1f}%"


# ── GITHUB HISTORY ────────────────────────────────────────────────────────────
HISTORY_FILE = "history_shops.json"

def gh_request(method, url, token, data=None):
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    req.add_header("Content-Type", "application/json")
    req.method = method
    body = json.dumps(data).encode() if data else None
    try:
        with urllib.request.urlopen(req, body, timeout=15) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"GitHub {e.code}: {e.read().decode()[:200]}")

def load_history_gh(token, repo):
    url = f"https://api.github.com/repos/{repo}/contents/{HISTORY_FILE}"
    try:
        r = gh_request("GET", url, token)
        return json.loads(base64.b64decode(r["content"]).decode()), r["sha"]
    except RuntimeError as e:
        if "404" in str(e):
            return [], None
        raise

def save_history_gh(token, repo, data, sha):
    url = f"https://api.github.com/repos/{repo}/contents/{HISTORY_FILE}"
    content = base64.b64encode(
        json.dumps(data, ensure_ascii=False, indent=2).encode()
    ).decode()
    payload = {"message": f"shop snapshot {date.today()}", "content": content}
    if sha:
        payload["sha"] = sha
    gh_request("PUT", url, token, payload)

def build_shop_snapshot(report_date, global_abc, group_abc, shop_summary, monthly_total):
    """Build a compact weekly snapshot for history."""
    smr = global_abc.groupby('ABC')['ВП_clean'].agg(['count','sum'])
    total_vp = global_abc['ВП_clean'].sum()

    top_groups = group_abc.head(10)[['Група','ВП','ABC_гр']].to_dict('records')
    top_shops  = shop_summary.head(10)[['Магазин','ВП (USD)','A SKU']].to_dict('records')
    monthly    = monthly_total.to_dict()

    return {
        "date":       str(report_date),
        "created_at": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "total_vp":   round(float(total_vp), 2),
        "total_sku":  len(global_abc),
        "n_shops":    len(shop_summary),
        "abc_counts": {k: int(v) for k, v in smr['count'].items()},
        "abc_vp":     {k: round(float(v), 2) for k, v in smr['sum'].items()},
        "top_groups": [{k: (round(float(v),2) if isinstance(v, float) else v)
                        for k,v in g.items()} for g in top_groups],
        "top_shops":  [{k: (round(float(v),2) if isinstance(v, float) else v)
                        for k,v in s.items()} for s in top_shops],
        "monthly":    {k: round(float(v), 2) for k, v in monthly.items()},
    }


# ── AI ANALYSIS ───────────────────────────────────────────────────────────────
def run_ai_global(api_key, snap, history, language, depth):
    """AI analysis for global shop report."""
    lang = "Відповідай виключно українською мовою." if language == "Українська" else "Reply in English."
    depth_instr = "350-450 слів." if "Короткий" in depth else "800-1000 слів з чіткими заголовками."

    past = [h for h in history if h.get("date") != snap.get("date")]
    has_hist = bool(past)

    # Delta vs prev week
    delta_text = ""
    if has_hist:
        prev = past[0]
        prev_vp = prev.get("total_vp", 0)
        cur_vp  = snap.get("total_vp", 0)
        if prev_vp > 0:
            d = (cur_vp - prev_vp) / prev_vp * 100
            delta_text = f"\nЗміна ВП vs {prev['date']}: {'+' if d>=0 else ''}{d:.1f}%"
            # Group deltas
            prev_grps = {g['Група']: g['ВП'] for g in prev.get('top_groups', [])}
            cur_grps  = {g['Група']: g['ВП'] for g in snap.get('top_groups', [])}
            gd_lines = []
            for grp, cvp in cur_grps.items():
                if grp in prev_grps and prev_grps[grp] > 0:
                    gd = (cvp - prev_grps[grp]) / prev_grps[grp] * 100
                    gd_lines.append(f"  {grp}: {'+' if gd>=0 else ''}{gd:.1f}%")
            if gd_lines:
                delta_text += "\nДинаміка груп:\n" + "\n".join(gd_lines[:8])

    top_groups_text = "\n".join([
        f"  {g['Група']}: ${g['ВП']:,.0f} (ABC: {g.get('ABC_гр','')})"
        for g in snap.get('top_groups', [])[:8]
    ])
    top_shops_text = "\n".join([
        f"  {s['Магазин']}: ${s['ВП (USD)']:,.0f} (A-SKU: {s['A SKU']})"
        for s in snap.get('top_shops', [])[:8]
    ])
    monthly_text = "\n".join([f"  {m}: ${v:,.0f}" for m, v in snap.get('monthly', {}).items()])

    ac = snap.get('abc_counts', {}); av = snap.get('abc_vp', {}); tot = snap.get('total_vp', 1)
    trend_instr = (
        "ОБОВ'ЯЗКОВО порівняй з попередніми тижнями: тренди ВП, зміни топ-груп, динаміка магазинів, прогноз."
        if has_hist else "Перший звіт — аналізуй поточний стан."
    )

    prompt = f"""Ти — досвідчений бізнес-аналітик роздрібної мережі нутриціональних добавок.
{lang} {depth_instr}
{trend_instr}

=== ЗВІТ ({snap.get('date')}) ===
Загальний ВП (без акцій): ${tot:,.2f}
SKU: {snap.get('total_sku')} | Магазинів: {snap.get('n_shops')}

ABC-розподіл:
- A: {ac.get('A',0)} SKU -> ${av.get('A',0):,.0f} ({av.get('A',0)/tot*100:.1f}%)
- B: {ac.get('B',0)} SKU -> ${av.get('B',0):,.0f} ({av.get('B',0)/tot*100:.1f}%)
- C: {ac.get('C',0)} SKU -> ${av.get('C',0):,.0f} ({av.get('C',0)/tot*100:.1f}%)

ТОП-8 НОМЕНКЛАТУРНИХ ГРУП:
{top_groups_text}

ТОП-8 МАГАЗИНІВ:
{top_shops_text}

МІСЯЧНА ДИНАМІКА (загальна мережа):
{monthly_text}
{delta_text}

СТРУКТУРА ВІДПОВІДІ:
1. Загальна оцінка мережі
2. {'Тренди vs попередній тиждень' if has_hist else 'ABC-структура портфелю'}
3. АНАЛІЗ НОМЕНКЛАТУРНИХ ГРУП — найважливіший розділ:
   - Топ-5 груп з динамікою і часткою
   - Групи що зростають / падають
   - {'Зміни в топ-10 груп' if has_hist else 'Розподіл по групах'}
4. Рейтинг магазинів: лідери і аутсайдери
5. {'Прогноз на наступний тиждень' if has_hist else 'Сезонна динаміка'}
6. Мінімум 4 конкретні рекомендації з цифрами"""

    client = anthropic.Anthropic(api_key=api_key)
    r = client.messages.create(
        model="claude-sonnet-4-20250514", max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    return r.content[0].text


def run_ai_shop(api_key, shop, shop_abc, group_abc_shop, language, depth):
    """AI analysis for a single shop."""
    lang = "Відповідай виключно українською мовою." if language == "Українська" else "Reply in English."
    depth_instr = "350-450 слів." if "Короткий" in depth else "800-1000 слів."

    smr = shop_abc.groupby('ABC')['ВП_clean'].agg(['count','sum'])
    total = shop_abc['ВП_clean'].sum()
    top15 = "\n".join([
        f"  {int(r['№'])}. {r['Артикул']} — ${r['ВП_clean']:,.0f} ({r['Частка%']:.1f}%): {str(r['Назва'])[:50]}"
        for _, r in shop_abc.head(15).iterrows()
    ])
    c_items = shop_abc[shop_abc['ABC'] == 'C']
    c_total = c_items['ВП_clean'].sum()
    c_count = len(c_items)

    grp_text = "\n".join([
        f"  {r['Група']}: ${r['ВП']:,.0f} (A: {r['A_sku']} SKU)"
        for _, r in group_abc_shop.head(10).iterrows()
    ])

    prompt = f"""Ти — досвідчений ритейл-аналітик. Проаналізуй товарну матрицю магазину та дай рекомендації щодо її оптимізації.
{lang} {depth_instr}

=== МАГАЗИН: {shop} ===
Загальний ВП: ${total:,.2f}
SKU всього: {len(shop_abc)}

ABC-розподіл:
- A: {int(smr.loc['A','count']) if 'A' in smr.index else 0} SKU -> ${smr.loc['A','sum'] if 'A' in smr.index else 0:,.0f} ({(smr.loc['A','sum'] if 'A' in smr.index else 0)/total*100:.1f}%)
- B: {int(smr.loc['B','count']) if 'B' in smr.index else 0} SKU -> ${smr.loc['B','sum'] if 'B' in smr.index else 0:,.0f} ({(smr.loc['B','sum'] if 'B' in smr.index else 0)/total*100:.1f}%)
- C: {c_count} SKU -> ${c_total:,.0f} ({c_total/total*100:.1f}%) — потенційний баласт

ТОП-15 SKU:
{top15}

ГРУПИ ТОВАРІВ (по ВП магазину):
{grp_text}

СТРУКТУРА ВІДПОВІДІ:
1. Оцінка товарної матриці магазину
2. Сильні позиції (категорія A) — що робить магазин добре
3. АНАЛІЗ НОМЕНКЛАТУРНИХ ГРУП:
   - Які групи представлені добре / слабко
   - Які групи варто розширити
   - Які групи варто скоротити
4. Категорія C — конкретні рекомендації:
   - Які товари вивести (критерії)
   - Які замінити
5. Категорія B — як перевести кращих у A
6. Мінімум 5 конкретних рекомендацій щодо оптимізації матриці з цифрами"""

    client = anthropic.Anthropic(api_key=api_key)
    r = client.messages.create(
        model="claude-sonnet-4-20250514", max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    return r.content[0].text


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ABC Налаштування")
    abc_a = st.slider("Поріг A (%)", 70, 85, 80)
    abc_b = st.slider("Поріг B (%)", 88, 98, 95)
    top_n = st.number_input("Топ SKU у таблиці", 5, 100, 20)

    st.divider()
    st.markdown("### AI-аналітика")
    _def_key = get_secret("ANTHROPIC_API_KEY")
    anthropic_key = st.text_input("Anthropic API Key", value=_def_key, type="password",
                                   placeholder="sk-ant-...")
    if anthropic_key: st.session_state["_ant_key"] = anthropic_key
    elif _def_key:    st.session_state["_ant_key"] = _def_key
    ai_lang  = st.selectbox("Мова", ["Українська", "English"])
    st.session_state["_ai_lang"] = ai_lang
    ai_depth = st.selectbox("Деталізація", ["Короткий (~400 слів)", "Повний (~900 слів)"], index=1)
    st.session_state["_ai_depth"] = ai_depth

    st.divider()
    st.markdown("### Збереження історії")
    _def_tok  = get_secret("GITHUB_TOKEN")
    _def_repo = get_secret("GITHUB_REPO")
    gh_token = st.text_input("GitHub Token", value=_def_tok, type="password", placeholder="ghp_...")
    gh_repo  = st.text_input("Репозиторій", value=_def_repo, placeholder="user/repo")
    report_date = st.date_input("Дата звіту", value=date.today())
    hist_enabled = bool(gh_token and gh_repo)
    if hist_enabled: st.success("✅ Збереження активне")
    else: st.caption("Заповніть Token і Репозиторій для накопичення даних.")

    st.divider()
    st.caption("Акційні місяці виключаються автоматично за аркушем 'акції'.")


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("## 🏪 ABC Аналіз по магазинах · Валовий прибуток")
st.divider()


# ── UPLOAD ────────────────────────────────────────────────────────────────────
st.markdown("## 01 · Завантажити файл")
f_upload = st.file_uploader(
    "Файл з місячними аркушами (кожен аркуш = один місяць) + аркуш 'акції'",
    type=["xlsx"], key="fu_shops"
)
if f_upload:
    st.success(f"Файл завантажено: {f_upload.name}")


# ── RUN ───────────────────────────────────────────────────────────────────────
st.divider()
st.markdown("## 02 · Запуск")
col_btn, col_inf = st.columns([1, 3])
with col_btn:
    run_btn = st.button("▶ Аналізувати", type="primary",
                        disabled=(f_upload is None), use_container_width=True)
with col_inf:
    if not f_upload:
        st.info("Завантажте файл для активації.")
    else:
        st.success("Готово до запуску.")


# ── ANALYSIS ─────────────────────────────────────────────────────────────────
if run_btn and f_upload:
    progress = st.progress(0, "Ініціалізація...")
    status   = st.empty()
    try:
        status.info("Збереження файлу...")
        progress.progress(5)
        fpath = save_upload(f_upload)

        status.info("Завантаження даних (може тривати 30-60 сек для великих файлів)...")
        progress.progress(15)
        df_long, shops, month_labels, month_dates, promo_df = load_all_months(fpath)
        promo_map = build_promo_map(promo_df, month_dates)

        status.info("Розрахунок глобального ABC...")
        progress.progress(40)
        global_abc = build_global_abc(df_long, promo_map, abc_a, abc_b)
        group_abc  = build_group_abc(global_abc, abc_a, abc_b)

        status.info("Зведення по магазинах...")
        progress.progress(60)
        shop_sum = build_shop_summary(df_long, promo_map, shops, abc_a, abc_b)

        status.info("Місячна динаміка...")
        progress.progress(75)
        monthly_shops  = build_monthly_by_shop(df_long, promo_map, month_labels, shops)
        monthly_groups = build_monthly_by_group(df_long, promo_map, month_labels)
        monthly_total  = monthly_shops.sum(axis=1)

        progress.progress(85)

        # Load history
        history, hist_sha = [], None
        if hist_enabled:
            status.info("Завантаження історії...")
            try:
                history, hist_sha = load_history_gh(gh_token, gh_repo)
            except Exception as e:
                st.warning(f"Історія: {e}")

        snap = build_shop_snapshot(report_date, global_abc, group_abc, shop_sum, monthly_total)

        # Store ALL in session_state
        st.session_state.update({
            "sh_ready":        True,
            "sh_global_abc":   global_abc,
            "sh_group_abc":    group_abc,
            "sh_shop_sum":     shop_sum,
            "sh_monthly_shops":  monthly_shops,
            "sh_monthly_groups": monthly_groups,
            "sh_monthly_total":  monthly_total,
            "sh_df_long":      df_long,
            "sh_shops":        shops,
            "sh_month_labels": month_labels,
            "sh_promo_map":    promo_map,
            "sh_snap":         snap,
            "sh_history":      history,
            "sh_hist_sha":     hist_sha,
            "sh_report_date":  str(report_date),
            "sh_gh_token":     gh_token,
            "sh_gh_repo":      gh_repo,
            "sh_hist_enabled": hist_enabled,
        })
        st.session_state.pop("sh_ai_global", None)
        st.session_state.pop("sh_ai_shop",   None)

        progress.progress(100, "Готово!")
        status.success("Аналіз завершено!")
        try: os.unlink(fpath)
        except: pass

    except Exception as e:
        import traceback
        progress.progress(100)
        status.error(f"Помилка: {e}")
        st.code(traceback.format_exc())


# ── RESULTS ───────────────────────────────────────────────────────────────────
if st.session_state.get("sh_ready"):
    global_abc    = st.session_state["sh_global_abc"]
    group_abc     = st.session_state["sh_group_abc"]
    shop_sum      = st.session_state["sh_shop_sum"]
    monthly_shops = st.session_state["sh_monthly_shops"]
    monthly_groups= st.session_state["sh_monthly_groups"]
    monthly_total = st.session_state["sh_monthly_total"]
    df_long       = st.session_state["sh_df_long"]
    shops         = st.session_state["sh_shops"]
    month_labels  = st.session_state["sh_month_labels"]
    promo_map     = st.session_state["sh_promo_map"]
    snap          = st.session_state["sh_snap"]
    history       = st.session_state["sh_history"]
    hist_sha      = st.session_state["sh_hist_sha"]
    today_str     = st.session_state["sh_report_date"]
    hist_enabled  = st.session_state["sh_hist_enabled"]
    gh_token      = st.session_state["sh_gh_token"]
    gh_repo       = st.session_state["sh_gh_repo"]

    total_vp = global_abc['ВП_clean'].sum()
    smr = global_abc.groupby('ABC')['ВП_clean'].agg(['count','sum'])
    a_vp = smr.loc['A','sum'] if 'A' in smr.index else 0
    b_vp = smr.loc['B','sum'] if 'B' in smr.index else 0
    c_vp = smr.loc['C','sum'] if 'C' in smr.index else 0

    # Delta
    past = [h for h in history if h.get("date") != today_str]
    prev = past[0] if past else None
    def delta(cur, prev_snap, key):
        if not prev_snap: return None
        pv = prev_snap.get(key, 0) or prev_snap.get("abc_vp", {}).get(key[-1], 0)
        if not pv: return None
        d = (cur - pv) / pv * 100
        return f"{'+' if d>=0 else ''}{d:.1f}% vs {prev_snap['date']}"

    # ── SECTION 03: METRICS ───────────────────────────────────────────────────
    st.divider()
    st.markdown("## 03 · Результати")
    m0, m1, m2, m3, m4 = st.columns(5)
    with m0: st.metric("Загальний ВП", fmt_usd(total_vp), help=f"{len(global_abc)} SKU, {len(shops)} магазинів")
    with m1: st.metric("Категорія A", fmt_usd(a_vp),
                        delta=delta(a_vp, prev, "A"),
                        help=f"{int(smr.loc['A','count'] if 'A' in smr.index else 0)} SKU · {a_vp/total_vp*100:.0f}%")
    with m2: st.metric("Категорія B", fmt_usd(b_vp),
                        delta=delta(b_vp, prev, "B"),
                        help=f"{int(smr.loc['B','count'] if 'B' in smr.index else 0)} SKU · {b_vp/total_vp*100:.0f}%")
    with m3: st.metric("Категорія C", fmt_usd(c_vp),
                        delta=delta(c_vp, prev, "C"),
                        help=f"{int(smr.loc['C','count'] if 'C' in smr.index else 0)} SKU · {c_vp/total_vp*100:.0f}%")
    with m4: st.metric("Магазинів", len(shops))
    if prev:
        st.caption(f"Дельти відносно звіту за {prev['date']}")

    st.divider()

    # ── SECTION 04: TABS ─────────────────────────────────────────────────────
    st.markdown("## 04 · Аналітика")
    tab_gl, tab_gr, tab_sh, tab_dyn_sh, tab_dyn_gr = st.tabs([
        "🏆 Топ SKU (загалом)",
        "📁 Групи товарів",
        "🏪 Магазини",
        "📈 Динаміка — магазини",
        "📈 Динаміка — групи",
    ])

    with tab_gl:
        st.caption(f"ABC по всій мережі ({len(global_abc)} SKU, без акційних місяців)")
        show = global_abc.head(int(top_n)).copy()
        show['ВП_clean'] = show['ВП_clean'].map(lambda x: f"${x:,.2f}")
        show['Частка%']  = show['Частка%'].map(lambda x: f"{x:.2f}%")
        show['Кумул%']   = show['Кумул%'].map(lambda x: f"{x:.1f}%")
        rename = {'ВП_clean':'ВП (USD)','Частка%':'Частка','Кумул%':'Кумул.',
                  'Акц_міс':'Виключені міс.','N_promo':'Акц. міс.'}
        st.dataframe(show.rename(columns=rename)[
            [c for c in ['№','Артикул','Група','Назва','ВП (USD)','Частка','Кумул.','ABC','Акц. міс.','Виключені міс.']
             if c in show.rename(columns=rename).columns]
        ], use_container_width=True, hide_index=True)

    with tab_gr:
        st.caption(f"{len(group_abc)} номенклатурних груп")
        show_g = group_abc.copy()
        show_g['ВП'] = show_g['ВП'].map(lambda x: f"${x:,.0f}")
        show_g['Частка%'] = show_g['Частка%'].map(lambda x: f"{x:.1f}%")
        show_g['Кумул%']  = show_g['Кумул%'].map(lambda x: f"{x:.1f}%")
        st.dataframe(show_g.rename(columns={'ABC_гр':'ABC груп.','A_sku':'A SKU','B_sku':'B SKU','C_sku':'C SKU','ВП':'ВП (USD)'}),
                     use_container_width=True, hide_index=True)
        # Bar chart top 20 groups
        chart_g = group_abc.head(20).copy()
        chart_g = chart_g.set_index('Група')[['ВП']]
        st.bar_chart(chart_g, height=300)

    with tab_sh:
        st.caption(f"{len(shop_sum)} магазинів, відсортовано по ВП")
        show_s = shop_sum.copy()
        show_s['ВП (USD)'] = show_s['ВП (USD)'].map(lambda x: f"${x:,.0f}")
        show_s['A ВП'] = show_s['A ВП'].map(lambda x: f"${x:,.0f}")
        show_s['B ВП'] = show_s['B ВП'].map(lambda x: f"${x:,.0f}")
        show_s['C ВП'] = show_s['C ВП'].map(lambda x: f"${x:,.0f}")
        show_s['Частка%'] = show_s['Частка%'].map(lambda x: f"{x:.1f}%")
        st.dataframe(show_s, use_container_width=True, hide_index=True)
        # Bar chart shops
        chart_s = shop_sum.head(20).set_index('Магазин')[['ВП (USD)']].copy()
        chart_s['ВП (USD)'] = shop_sum.head(20)['ВП (USD)'].values
        st.bar_chart(chart_s, height=300)

    with tab_dyn_sh:
        st.caption("Місячна динаміка ВП по магазинах (без акційних місяців)")
        top_shops_list = shop_sum['Магазин'].head(15).tolist()
        chart_ms = monthly_shops[top_shops_list]
        st.line_chart(chart_ms, height=350)
        # Table
        fmt_ms = monthly_shops.copy().apply(lambda col: col.map(lambda x: f"${x:,.0f}"))
        fmt_ms['ЗАГАЛОМ'] = monthly_total.map(lambda x: f"${x:,.0f}")
        st.dataframe(fmt_ms, use_container_width=True)

    with tab_dyn_gr:
        st.caption("Місячна динаміка ВП по номенклатурних групах")
        top_grps = group_abc['Група'].head(12).tolist()
        chart_mg = monthly_groups[top_grps]
        st.line_chart(chart_mg, height=350)
        fmt_mg = monthly_groups[top_grps].copy().apply(lambda col: col.map(lambda x: f"${x:,.0f}"))
        st.dataframe(fmt_mg, use_container_width=True)

    # ── SECTION 05: SAVE SNAPSHOT ────────────────────────────────────────────
    st.divider()
    st.markdown("## 05 · Збереження знімку")
    if not hist_enabled:
        st.info("Заповніть GitHub Token і Репозиторій у бічній панелі.")
    else:
        is_dup = any(h.get("date") == today_str for h in history)
        if is_dup:
            st.warning(f"Знімок за {today_str} вже існує — збереження перезапише.")
        cs, ci = st.columns([1, 3])
        with cs:
            if st.button("💾 Зберегти знімок", key="btn_save_sh", use_container_width=True):
                with st.spinner("Збереження..."):
                    try:
                        updated = [h for h in history if h.get("date") != today_str]
                        updated.append(snap)
                        updated = sorted(updated, key=lambda x: x.get("date",""), reverse=True)[:52]
                        save_history_gh(gh_token, gh_repo, updated, hist_sha)
                        st.session_state["sh_history"] = updated
                        st.session_state["sh_hist_sha"] = None
                        st.success(f"Збережено! В архіві: {len(updated)} тижнів.")
                    except Exception as e:
                        st.error(f"Помилка: {e}")
        with ci:
            st.caption(f"В архіві: {len(history)} тижн(ів).")

    # ── SECTION 06: HISTORY CHART ────────────────────────────────────────────
    if len(history) > 1:
        st.divider()
        st.markdown("## 06 · Архів — тренд ВП")
        sorted_h = sorted(history, key=lambda x: x.get("date",""))
        trend = pd.DataFrame([{
            "Дата": h["date"],
            "Загальний ВП": h.get("total_vp", 0),
            "A ВП": h.get("abc_vp", {}).get("A", 0),
        } for h in sorted_h]).set_index("Дата")
        st.line_chart(trend, height=250)
        hist_tbl = pd.DataFrame([{
            "Дата": h["date"], "ВП": f"${h.get('total_vp',0):,.0f}",
            "SKU": h.get("total_sku",0), "Магазинів": h.get("n_shops",0),
            "A SKU": h.get("abc_counts",{}).get("A",0),
        } for h in sorted_h[::-1]])
        st.dataframe(hist_tbl, use_container_width=True, hide_index=True)

    # ── SECTION 07: AI GLOBAL ────────────────────────────────────────────────
    st.divider()
    st.markdown("## 07 · AI-аналіз мережі")
    cur_key = st.session_state.get("_ant_key", "")
    if not cur_key:
        st.info("Введіть Anthropic API Key у бічній панелі.")
    else:
        n_hist = len([h for h in history if h.get("date") != today_str])
        if n_hist > 0:
            st.success(f"Є {n_hist} тижн(ів) — AI проаналізує тренди.")
        else:
            st.info("Перший звіт — AI зробить аналіз поточного стану.")

        if st.button("🤖 AI-аналіз мережі", type="primary", key="btn_ai_gl"):
            with st.spinner("AI аналізує мережу..."):
                try:
                    result = run_ai_global(
                        cur_key, snap, history,
                        st.session_state.get("_ai_lang","Українська"),
                        st.session_state.get("_ai_depth","Повний (~900 слів)")
                    )
                    st.session_state["sh_ai_global"] = result
                except anthropic.AuthenticationError:
                    st.error("Невірний API ключ.")
                except Exception as e:
                    import traceback
                    st.error(f"Помилка: {type(e).__name__}: {e}")
                    st.code(traceback.format_exc())

    if st.session_state.get("sh_ai_global"):
        st.markdown("---")
        st.markdown(st.session_state["sh_ai_global"])
        st.download_button("📄 Завантажити (.txt)",
                           data=st.session_state["sh_ai_global"],
                           file_name=f"AI_мережа_{today_str}.txt",
                           mime="text/plain", key="dl_ai_gl")

    # ── SECTION 08: SINGLE SHOP ANALYSIS ────────────────────────────────────
    st.divider()
    st.markdown("## 08 · ABC аналіз окремого магазину")
    shops_sorted = shop_sum['Магазин'].tolist()
    selected_shop = st.selectbox("Оберіть магазин", shops_sorted, key="sel_shop")

    if selected_shop:
        shop_abc_sel = build_shop_abc(df_long, selected_shop, promo_map, abc_a, abc_b)
        if not shop_abc_sel.empty:
            shop_total = shop_abc_sel['ВП_clean'].sum()
            sm2 = shop_abc_sel.groupby('ABC')['ВП_clean'].agg(['count','sum'])
            c_shop, c_info = st.columns([2, 1])
            with c_shop:
                st.markdown(f"**{selected_shop}** · ${shop_total:,.0f} · {len(shop_abc_sel)} SKU")
            with c_info:
                for abc, color in [('A','🟢'),('B','🟡'),('C','🔴')]:
                    if abc in sm2.index:
                        st.caption(f"{color} {abc}: {int(sm2.loc[abc,'count'])} SKU · ${sm2.loc[abc,'sum']:,.0f}")

            # Shop ABC table
            with st.expander(f"ABC таблиця — {selected_shop}", expanded=True):
                show_sh = shop_abc_sel.head(int(top_n)).copy()
                show_sh['ВП_clean'] = show_sh['ВП_clean'].map(lambda x: f"${x:,.2f}")
                show_sh['Частка%']  = show_sh['Частка%'].map(lambda x: f"{x:.2f}%")
                show_sh['Кумул%']   = show_sh['Кумул%'].map(lambda x: f"{x:.1f}%")
                st.dataframe(show_sh.rename(columns={'ВП_clean':'ВП (USD)','Частка%':'Частка','Кумул%':'Кумул.'})[
                    [c for c in ['№','Артикул','Група','Назва','ВП (USD)','Частка','Кумул.','ABC']
                     if c in show_sh.rename(columns={'ВП_clean':'ВП (USD)','Частка%':'Частка','Кумул%':'Кумул.'}).columns]
                ], use_container_width=True, hide_index=True)

            # Group breakdown for this shop
            grp_shop = shop_abc_sel.groupby('Група').agg(
                ВП=('ВП_clean','sum'), SKU=('Артикул','count'),
                A_sku=('ABC', lambda x:(x=='A').sum()),
            ).reset_index().sort_values('ВП', ascending=False)

            # AI for shop
            if cur_key:
                if st.button(f"🤖 AI-аналіз: {selected_shop[:30]}", key="btn_ai_sh"):
                    with st.spinner(f"AI аналізує {selected_shop}..."):
                        try:
                            result = run_ai_shop(
                                cur_key, selected_shop, shop_abc_sel, grp_shop,
                                st.session_state.get("_ai_lang","Українська"),
                                st.session_state.get("_ai_depth","Повний (~900 слів)")
                            )
                            st.session_state["sh_ai_shop"] = result
                            st.session_state["sh_ai_shop_name"] = selected_shop
                        except anthropic.AuthenticationError:
                            st.error("Невірний API ключ.")
                        except Exception as e:
                            import traceback
                            st.error(f"Помилка: {type(e).__name__}: {e}")
                            st.code(traceback.format_exc())
            else:
                st.info("Введіть API ключ для AI-аналізу магазину.")

            if (st.session_state.get("sh_ai_shop") and
                    st.session_state.get("sh_ai_shop_name") == selected_shop):
                st.markdown("---")
                st.markdown(st.session_state["sh_ai_shop"])
                st.download_button(
                    f"📄 Завантажити аналіз ({selected_shop[:20]}...)",
                    data=st.session_state["sh_ai_shop"],
                    file_name=f"AI_{selected_shop[:20]}_{today_str}.txt",
                    mime="text/plain", key="dl_ai_sh"
                )

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("ABC Аналіз Магазини · Акційні місяці виключаються за аркушем 'акції' · Дані — history_shops.json у GitHub")
