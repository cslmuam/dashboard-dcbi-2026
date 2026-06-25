#!/usr/bin/env python3
"""Build the DCBI dashboard — 5-page static site.

Pages: index.html, cat26p.html, cat26p_capacidad.html,
       metanalisis.html, modificaciones.html
Run: python3 build.py  (from any directory)
"""

import pandas as pd, json, re, base64, glob
from pathlib import Path

HERE        = Path(__file__).parent
ADMIN       = HERE.parent
CLAUDE_PROJ = ADMIN.parent

# ── Logo ─────────────────────────────────────────────────────────────────────
with open(CLAUDE_PROJ / 'Personal/logo_UAM.png', 'rb') as f:
    LOGO = base64.b64encode(f.read()).decode()

# ══════════════════════════════════════════════════════════════════════════════
# DATA EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

NOM = {'amb':'Ambiental','civ':'Civil','com':'Computación','ele':'Eléctrica',
       'elo':'Electrónica','fis':'Física','ind':'Industrial','mec':'Mecánica',
       'met':'Metalurgia','qui':'Química'}
COLORS = {'Ambiental':'#2E7D32','Civil':'#BF360C','Computación':'#1565C0',
          'Eléctrica':'#E65100','Electrónica':'#6A1B9A','Física':'#00838F',
          'Industrial':'#C62828','Mecánica':'#37474F','Metalurgia':'#78909C','Química':'#558B2F'}

# ── CAT26P — fuente de verdad: Asignacion_Automatica_26P.xlsx ────────────────
SHORT_PLAN = {
    'INGENIERIA EN COMPUTACION':'Computación','INGENIERIA MECANICA':'Mecánica',
    'INGENIERIA CIVIL':'Civil','INGENIERIA QUIMICA':'Química',
    'INGENIERIA INDUSTRIAL':'Industrial','INGENIERIA ELECTRONICA':'Electrónica',
    'INGENIERIA ELECTRICA':'Eléctrica','INGENIERIA FISICA':'Física',
    'INGENIERIA AMBIENTAL':'Ambiental','INGENIERIA METALURGICA':'Metalurgia',
}
_CUPOS_CAT = {'1100037':45,'1111078':45,'1112013':45,'1112042':45,'1113084':45,'1113085':36}
_UEA_FULL  = {
    '1100037':'Introducción a la Ingeniería',
    '1111078':'Introducción a la Física',
    '1112013':'Complementos de Matemáticas',
    '1112042':'Introducción al Cálculo',
    '1113084':'Estructura Atómica y Enlace Químico',
    '1113085':'Laboratorio de Reacciones Químicas',
}
_UEA_SHORT_CHART = {
    '1100037':'Intro. a la Ing.','1111078':'Intro. a la Física',
    '1112013':'Comp. de Matemáticas','1112042':'Intro. al Cálculo',
    '1113084':'Estructura Atómica','1113085':'Lab. Reacciones Quím.',
}
_UEA_SHORT = {
    '1100037':'Intro. Ingeniería','1111078':'Intro. Física',
    '1112013':'Comp. Matemáticas','1112042':'Intro. Cálculo',
    '1113084':'Estr. Atómica','1113085':'Lab. Reacciones',
}
_UEA_COLS = [1100037, 1111078, 1112013, 1112042, 1113084, 1113085]

aa = pd.read_excel(ADMIN / 'CAT26P/26P/Asignacion_Automatica_26P.xlsx', dtype=str)
aa['plan_short'] = aa['nom_lic'].map(SHORT_PLAN)

# Plan totals and assignment completeness
gp = aa.groupby('plan_short')
plan_total    = gp.size().sort_values(ascending=False)
plan_complete = plan_total.copy()        # all 674 received 6 UEAs
plan_partial  = plan_total * 0           # zero partial assignments

# Groups that reached 100% capacity, per UEA (from actual assignment file)
from collections import Counter as _Counter
_groups_full: dict[str, int] = {}
for c in _UEA_COLS:
    gc = _Counter(v for v in aa[c] if pd.notna(v) and str(v).strip() not in ('', 'nan'))
    cupo = _CUPOS_CAT[str(c)]
    _groups_full[_UEA_SHORT_CHART[str(c)]] = sum(1 for n in gc.values() if n >= cupo)
cat_ueas = sorted(_groups_full.items(), key=lambda x: -x[1])

# Per-UEA group occupancy (aa already loaded above; _Counter already imported)
cap_uea_data = []
all_grp_set: set[str] = set()
_gc_by_uea: dict[str, _Counter] = {}
for c in _UEA_COLS:
    gc = _Counter(v for v in aa[c] if pd.notna(v) and str(v).strip() not in ('','nan'))
    all_grp_set |= set(gc.keys())
    _gc_by_uea[str(c)] = gc
    n_grp = len(gc)
    cupo  = _CUPOS_CAT[str(c)]
    cap   = n_grp * cupo
    enr   = sum(gc.values())
    cap_uea_data.append({
        'code': str(c), 'name': _UEA_FULL[str(c)], 'short': _UEA_SHORT[str(c)],
        'n_groups': n_grp, 'cupo_unit': cupo, 'cap_total': cap,
        'enrolled': enr, 'util': round(enr/cap*100, 1),
        'full_groups': sum(1 for n in gc.values() if n >= cupo),
        'remaining': cap - enr,
    })

all_grps = sorted(all_grp_set)
# Override Lab entry to full catalog capacity (24 sections × 36 = 864, resumen ejecutivo)
for _d in cap_uea_data:
    if _d['code'] == '1113085':
        _d['cap_total'] = 864
        _d['n_groups']  = 24
        _d['remaining'] = 864 - _d['enrolled']
        _d['util']      = round(_d['enrolled'] / 864 * 100, 1)
bottleneck_code = min(cap_uea_data, key=lambda x: x['cap_total'])['code']
max_cap = 864  # full catalog: 24 Lab sections × 36 (resumen ejecutivo)

# Per-group occupancy for lab (bottleneck) — all groups
lab_gc = _gc_by_uea['1113085']
lab_groups_sorted = sorted(lab_gc.keys(), key=lambda g: -lab_gc[g])
lab_enrolled = [lab_gc[g] for g in lab_groups_sorted]
lab_cap_each  = [36]*len(lab_groups_sorted)
lab_pct = [round(lab_gc[g]/36*100, 1) for g in lab_groups_sorted]

# All-UEA per-group occupancy (uniform across the 5 non-lab UEAs)
ref_gc = _gc_by_uea['1111078']  # representative (identical across 5 UEAs)
all_grps_5 = sorted(g for g in all_grps if ref_gc.get(g,0) > 0)
all_5_enr  = [ref_gc[g] for g in all_grps_5]
all_5_pct  = [round(ref_gc[g]/45*100, 1) for g in all_grps_5]

# Sensitivity: max admissible vs extra Lab groups (k=0..8)
# Base: full catalog = 24 Lab sections × 36 = 864; next bottleneck = Física at 900 (20 × 45)
sensitivity = []
for k in range(9):
    lab_cap_k  = (24 + k) * 36
    other_cap  = 900  # Física: 20 sec × 45 = 900 (next bottleneck, resumen cupos §2)
    bot        = min(lab_cap_k, other_cap)
    sensitivity.append({'k': k, 'lab_cap': lab_cap_k, 'other_cap': other_cap,
                        'max_cap': bot, 'util_674': round(674/bot*100, 1)})

# Gender
gen_m = int((aa['gen'].str.strip().str.upper() == 'M').sum())
gen_f = int((aa['gen'].str.strip().str.upper() == 'F').sum())

# ── Metaanálisis ─────────────────────────────────────────────────────────────
DATA_DIR = ADMIN / 'Metanalisis/Output/datos'
_plan_names = {
    'amb':'Ambiental','civ':'Civil','com':'Computación','ele':'Eléctrica',
    'elo':'Electrónica','fis':'Física','ind':'Industrial','mec':'Mecánica',
    'met':'Metalurgia','qui':'Química'
}
_plans_order = ['amb','civ','com','ele','elo','fis','ind','mec','met','qui']

# ── META1: Diagnóstico ─────────────────────────────────────────────────────
_df_et = pd.read_csv(DATA_DIR / 'simulacion/curvas_ET_v3.csv')
_df_dc = pd.read_csv(DATA_DIR / 'simulacion/descomposicion_carga.csv')
_curves = {}
for _plan, _g in _df_et.groupby('plan'):
    _curves[_plan] = {'t': _g['t'].tolist(), 'ET': [round(v, 5) for v in _g['ET'].tolist()]}
_dist_stats = {}
for _, _r in _df_dc.iterrows():
    _dist_stats[_r['plan']] = {
        'nombre':        _r['nombre'],
        'et_tit':        round(float(_r['et_tit_baseline']), 2),
        'et_libre':      round(float(_r['et_tit_libre']),    2),
        'et_oficial':    round(float(_r['et_oficial']),      2),
        'et18_baseline': round(float(_r['et18_baseline']),   4),
        'et24_baseline': round(float(_r['et24_baseline']),   4),
    }
DATA_META1 = json.dumps({
    'planNames':  _plan_names,
    'plansOrder': _plans_order,
    'curves':     _curves,
    'distStats':  _dist_stats,
}, ensure_ascii=False)

# ── META2: Causas ──────────────────────────────────────────────────────────
_t_carga = {'amb':18,'civ':19,'com':22,'ele':20,'elo':22,'fis':21,
            'ind':19,'mec':20,'met':13,'qui':20}
_t_topo  = {'amb': 9,'civ': 8,'com':10,'ele': 9,'elo':10,'fis': 8,
            'ind': 6,'mec': 9,'met': 9,'qui': 9}
_a_bar   = {'amb':0.74,'civ':0.68,'com':0.72,'ele':0.66,'elo':0.66,
            'fis':0.72,'ind':0.74,'mec':0.70,'met':0.74,'qui':0.74}
_cs_cal = {_r['plan']: float(_r['cs_calibrado']) for _, _r in _df_dc.iterrows()}
_df_vp = pd.read_csv(DATA_DIR / 'indices/estrategias_ventaja_P.csv')
_vp_avg = _df_vp.groupby('t').agg(
    P_base=('P_base','mean'), P_cons=('P_cons','mean'), P_agr=('P_agr','mean'),
    delta_cons=('delta_P_cons_base','mean'), delta_agr=('delta_P_agr_base','mean')
).reset_index()
DATA_META2 = json.dumps({
    'planNames':  _plan_names,
    'plansOrder': _plans_order,
    'tCarga':     _t_carga,
    'tTopo':      _t_topo,
    'aBar':       _a_bar,
    'cs':         _cs_cal,
    'ventajaP': {
        't':          _vp_avg['t'].tolist(),
        'P_base':     [round(float(v), 4) for v in _vp_avg['P_base']],
        'P_cons':     [round(float(v), 4) for v in _vp_avg['P_cons']],
        'P_agr':      [round(float(v), 4) for v in _vp_avg['P_agr']],
        'delta_cons': [round(float(v), 4) for v in _vp_avg['delta_cons']],
        'delta_agr':  [round(float(v), 4) for v in _vp_avg['delta_agr']],
    },
}, ensure_ascii=False)

# ── META3: La palanca ──────────────────────────────────────────────────────
_df_ser = pd.read_csv(DATA_DIR / 'simulacion/efecto_seriaciones_todos.csv')
_ser_all = _df_ser[_df_ser['tipo'] == 'seriacion'].copy()
# Lorenz on |dE_T|: matches infographic's N80=132/499 (26.5%)
_ser_all['gain'] = _ser_all['dE_T'].abs()
_ser_all = _ser_all.sort_values('gain', ascending=False).reset_index(drop=True)
_total_gain = _ser_all['gain'].sum()
_ser_all['frac_cum'] = _ser_all['gain'].cumsum() / _total_gain
_ser_all['n'] = _ser_all.index + 1
_total_ser = int(len(_ser_all))  # 499
_idx = list(range(0, min(50, _total_ser))) + list(range(50, _total_ser, 5))
_idx = sorted(set(_idx))
_n80 = int(_ser_all[_ser_all['frac_cum'] >= 0.80]['n'].iloc[0])  # 132
_remocion = {}
for _p in _plans_order:
    _df_r = pd.read_csv(DATA_DIR / f'remocion/remocion_combinada_{_p}.csv')
    _remocion[_p] = {
        'N':    _df_r['N'].tolist(),
        'mean': [round(float(v), 4) for v in _df_r['gain_ET_mean']],
        'lo':   [round(float(v), 4) for v in _df_r['gain_ET_ci_lo']],
        'hi':   [round(float(v), 4) for v in _df_r['gain_ET_ci_hi']],
    }
DATA_META3 = json.dumps({
    'planNames':  _plan_names,
    'plansOrder': _plans_order,
    'lorenz': {
        'n':       [float(_ser_all.loc[i, 'n'])        for i in _idx],
        'fracCum': [round(float(_ser_all.loc[i, 'frac_cum']), 5) for i in _idx],
        'N80':      _n80,
        'totalSer': _total_ser,
    },
    'remocion': _remocion,
}, ensure_ascii=False)

# ── META4: Índice P ────────────────────────────────────────────────────────
_vp_avg4 = _df_vp.groupby('t').agg(
    P_base=('P_base', 'mean'), P_cons=('P_cons', 'mean'), P_agr=('P_agr', 'mean'),
    delta_cons=('delta_P_cons_base', 'mean'),
).reset_index()
_estrategias_resumen = {
    'conservadora': {'ET': 24.4,  'grad': 20.7, 'deltaP_t2': 0.11,  'label': 'Conservadora (24 cr/trim)'},
    'baseline':     {'ET': 19.7,  'grad': 59.6, 'deltaP_t2': 0.0,   'label': 'Baseline (carga histórica)'},
    'agresiva':     {'ET': 17.2,  'grad': 62.8, 'deltaP_t2': -0.04, 'label': 'Agresiva (38 cr/trim)'},
}
DATA_META4 = json.dumps({
    'planNames':  _plan_names,
    'plansOrder': _plans_order,
    'pByStrat': {
        't':            _vp_avg4['t'].tolist(),
        'baseline':     [round(float(v), 4) for v in _vp_avg4['P_base']],
        'conservadora': [round(float(v), 4) for v in _vp_avg4['P_cons']],
        'agresiva':     [round(float(v), 4) for v in _vp_avg4['P_agr']],
        'delta_cons':   [round(float(v), 4) for v in _vp_avg4['delta_cons']],
    },
    'estrategiasResumen': _estrategias_resumen,
}, ensure_ascii=False)

# ── META5: Complejidad ─────────────────────────────────────────────────────
_df_cg = pd.read_csv(DATA_DIR / 'complejidad/complejidad_global.csv')
_df_cn = pd.read_csv(DATA_DIR / 'complejidad/complejidad_normalizada.csv')
_radar_data = {}
for _, _r in _df_cn.iterrows():
    _radar_data[_r['plan']] = {
        'D1': round(float(_r['D1']), 4), 'D2': round(float(_r['D2']), 4),
        'D3': round(float(_r['D3']), 4), 'D4': round(float(_r['D4']), 4),
        'D5': round(float(_r['D5']), 4),
    }
_global_C   = {_r['plan']: round(float(_r['C']), 5)          for _, _r in _df_cg.iterrows()}
_plan_clase = {_r['plan']: _r['clase']                        for _, _r in _df_cg.iterrows()}
_dim_dom_n  = {_r['plan']: _r['dim_dom_nombre']               for _, _r in _df_cg.iterrows()}
_ranking_order = _df_cg.sort_values('C', ascending=False)['plan'].tolist()
DATA_META5 = json.dumps({
    'planNames':    _plan_names,
    'plansOrder':   _plans_order,
    'radarData':    _radar_data,
    'globalC':      _global_C,
    'planClase':    _plan_clase,
    'dimDomNombre': _dim_dom_n,
    'rankingOrder': _ranking_order,
    'dimLabels': ['D1 Topológica', 'D2 Probabilística', 'D3 Evaluativa',
                  'D4 Docente', 'D5 Recuperación'],
    'claseColors': {'Alta': '#C82D23', 'Moderada': '#B06A00', 'Baja': '#2E7D32'},
}, ensure_ascii=False)

# ── Modificaciones ────────────────────────────────────────────────────────────
with open(ADMIN / 'Modificaciones2026/Revisiones/Datos/similitud_llm_dcbi_2026.json') as f:
    sim_json = json.load(f)
from collections import Counter
sim_counts = Counter(v['nivel'] for v in sim_json['puntuaciones'].values())

# ══════════════════════════════════════════════════════════════════════════════
# SERIALIZED DATA FOR JS
# ══════════════════════════════════════════════════════════════════════════════
_n_total      = int(len(aa))
_n_completa   = int((aa.apply(lambda r: sum(1 for c in _UEA_COLS if pd.notna(r[c]) and str(r[c]).strip() not in ('','nan')), axis=1) == 6).sum())
_n_parcial    = _n_total - _n_completa
_gruposLlenos = sum(v for _,v in cat_ueas)
_gen_m = int((aa['gen'].str.strip().str.upper() == 'M').sum())
_gen_f = int((aa['gen'].str.strip().str.upper() == 'F').sum())

# Gender by plan
_plan_m = [int((aa[aa['plan_short']==p]['gen'].str.strip().str.upper()=='M').sum()) for p in plan_total.index]
_plan_f = [int((aa[aa['plan_short']==p]['gen'].str.strip().str.upper()=='F').sum()) for p in plan_total.index]

DATA_CAT = json.dumps({
    'total': _n_total, 'completa': _n_completa, 'parcial': _n_parcial,
    'gruposLlenos': _gruposLlenos,
    'ueas':           [k for k,_ in cat_ueas],
    'gruposLlenados': [v for _,v in cat_ueas],
    'planNames':   plan_total.index.tolist(),
    'planTotal':   [int(x) for x in plan_total.tolist()],
    'planCompleta':[int(x) for x in plan_complete.tolist()],
    'planParcial': [int(x) for x in plan_partial.tolist()],
    'planColors':  [COLORS.get(p,'#888') for p in plan_total.index],
    'genM': _gen_m, 'genF': _gen_f,
    'genMpct': round(_gen_m/_n_total*100, 1),
    'genFpct': round(_gen_f/_n_total*100, 1),
    'planM': _plan_m, 'planF': _plan_f,
}, ensure_ascii=False)

DATA_CAP = json.dumps({
    'maxCap': max_cap,
    'enrolled': 674,
    'gap': max_cap - 674,
    'bottleneckName': _UEA_FULL[bottleneck_code],
    'bottleneckShort': _UEA_SHORT[bottleneck_code],
    'ueas': cap_uea_data,
    'labGroups': lab_groups_sorted,
    'labEnrolled': lab_enrolled,
    'labPct': lab_pct,
    'refGroups': all_grps_5,
    'refEnrolled': all_5_enr,
    'refPct': all_5_pct,
    'sensitivity': sensitivity,
    'genM': gen_m, 'genF': gen_f,
    'genMpct': round(gen_m/674*100, 1),
    'genFpct': round(gen_f/674*100, 1),
}, ensure_ascii=False)


DATA_MOD = json.dumps({
    'simLevels': {
        'labels':  ['Sin traslape', 'Bajo', 'Moderado', 'Alto', 'Muy alto', 'Idéntico'],
        'values':  [sim_counts.get('sin_traslape',0), sim_counts.get('bajo',0),
                    sim_counts.get('moderado',0), sim_counts.get('alto',0),
                    sim_counts.get('muy_alto',0), sim_counts.get('idéntico',0)],
        'colors':  ['#E4E2DF','#fed976','#fd8d3c','#fc4e2a','#e31a1c','#800026'],
    },
}, ensure_ascii=False)

# ══════════════════════════════════════════════════════════════════════════════
# SHARED CSS
# ══════════════════════════════════════════════════════════════════════════════
CSS = """
  :root {
    --red:#C82D23; --red-d:#9E1E18; --red-l:#FBEEED;
    --ink:#1A1A1A; --mid:#444; --muted:#888;
    --border:#E4E2DF; --bg:#F5F4F1; --card:#fff;
    --teal:#00838F; --teal-light:#E0F4F5;
    --amber:#B06A00; --green:#2E7D32;
    --shadow-sm:0 1px 4px rgba(0,0,0,.07),0 4px 12px rgba(0,0,0,.06);
    --shadow-md:0 4px 16px rgba(0,0,0,.10),0 12px 32px rgba(0,0,0,.08);
    --radius:10px;
    --font:'Helvetica Neue',Arial,system-ui,sans-serif;
    --ease:.22s cubic-bezier(.4,0,.2,1);
    --c-amb:#2E7D32; --c-civ:#BF360C; --c-com:#1565C0; --c-ele:#E65100;
    --c-elo:#6A1B9A; --c-fis:#00838F; --c-ind:#C62828;
    --c-mec:#37474F; --c-met:#78909C; --c-qui:#558B2F;
  }
  *, *::before, *::after { box-sizing:border-box; margin:0; padding:0 }
  html { font-size:16px; scroll-behavior:smooth }
  body { font-family:var(--font); background:var(--bg); color:var(--ink);
         line-height:1.6; -webkit-font-smoothing:antialiased;
         animation:fadeInPage .45s ease both }
  body.page-exit { opacity:0; transform:translateY(6px);
                   transition:opacity .28s ease,transform .28s ease; pointer-events:none }
  @keyframes fadeInPage { from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:none} }
  a { color:inherit; text-decoration:none }
  img { display:block; max-width:100% }

  /* ── Nav ── */
  .nav { position:sticky; top:0; z-index:100;
         background:rgba(255,255,255,.92); backdrop-filter:blur(12px);
         border-bottom:1px solid var(--border); padding:0 2rem }
  .nav-inner { max-width:1440px; margin:0 auto;
               display:flex; align-items:center; gap:.25rem; height:52px }
  .nav-brand { font-size:.78rem; font-weight:700; letter-spacing:.08em;
               text-transform:uppercase; color:var(--red); white-space:nowrap;
               margin-right:auto; padding:.3rem .6rem .3rem 0 }
  .nav-links { display:flex; gap:0; list-style:none }
  .nav-links a { display:block; padding:.3rem .9rem; font-size:.82rem; color:var(--mid);
                 border-radius:6px; transition:background var(--ease),color var(--ease) }
  .nav-links a:hover { background:var(--red-l); color:var(--red) }
  .nav-links a.active { background:var(--red-l); color:var(--red); font-weight:600 }
  .nav-hamburger { display:none; background:none; border:none; cursor:pointer;
                   padding:8px; flex-direction:column; justify-content:center; gap:5px;
                   width:36px; height:36px; border-radius:6px;
                   transition:background var(--ease) }
  .nav-hamburger:hover { background:var(--red-l) }
  .nav-hamburger span { display:block; width:20px; height:2px; background:var(--ink);
                        border-radius:2px; transition:transform .22s ease,opacity .22s ease }
  .nav-hamburger.open span:nth-child(1) { transform:translateY(7px) rotate(45deg) }
  .nav-hamburger.open span:nth-child(2) { opacity:0; transform:scaleX(0) }
  .nav-hamburger.open span:nth-child(3) { transform:translateY(-7px) rotate(-45deg) }

  /* ── Hero ── */
  .hero { background:var(--ink); color:#fff; padding:4rem 2rem 3rem }
  .hero-inner { max-width:1440px; margin:0 auto;
                display:grid; grid-template-columns:1fr auto; gap:2rem; align-items:end }
  .hero-label { font-size:.72rem; font-weight:700; letter-spacing:.14em;
                text-transform:uppercase; color:var(--red); margin-bottom:.6rem }
  .hero h1 { font-size:clamp(1.6rem,3vw,2.6rem); font-weight:300;
             line-height:1.2; letter-spacing:-.02em }
  .hero h1 strong { font-weight:700 }
  .hero-sub { margin-top:.75rem; font-size:.92rem; color:rgba(255,255,255,.55); max-width:580px }
  .hero-bar { width:48px; height:3px; background:var(--red); margin:1.5rem 0 }
  .stats { display:flex; flex-wrap:wrap; gap:.5rem 1.5rem }
  .stat { display:flex; flex-direction:column; padding:.75rem 1.1rem;
          background:rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.1);
          border-radius:8px; min-width:90px }
  .stat-val { font-size:2rem; font-weight:700; color:#fff; line-height:1 }
  .stat-lbl { font-size:.68rem; color:rgba(255,255,255,.45); margin-top:.25rem; letter-spacing:.04em }
  .palette { display:flex; flex-direction:column; gap:.4rem }
  .palette-title { font-size:.66rem; color:rgba(255,255,255,.4);
                   text-transform:uppercase; letter-spacing:.1em; margin-bottom:.2rem }
  .palette-chips { display:flex; flex-direction:column; gap:.28rem }
  .chip { display:flex; align-items:center; gap:.5rem; font-size:.72rem; color:rgba(255,255,255,.65) }
  .chip-dot { width:10px; height:10px; border-radius:50%; flex-shrink:0 }

  /* ── Main / Sections ── */
  .main { max-width:1440px; margin:0 auto; padding:0 2rem 4rem }
  .section { margin-top:3.5rem }
  .section-header { display:flex; align-items:baseline; gap:1rem;
                    padding-bottom:.75rem; border-bottom:2px solid var(--border); margin-bottom:1.5rem }
  .section-header h2 { font-size:1.2rem; font-weight:700; letter-spacing:-.01em }
  .section-count { font-size:.78rem; color:var(--muted); background:var(--bg);
                   border:1px solid var(--border); padding:.15rem .55rem; border-radius:20px }
  .section-desc { font-size:.87rem; color:var(--mid); margin-bottom:1.5rem;
                  max-width:740px; line-height:1.7 }

  /* ── Callout / Recuadro ── */
  .read-callout { background:var(--teal-light); border-left:4px solid var(--teal);
                  border-radius:0 8px 8px 0; padding:.9rem 1.1rem;
                  margin-bottom:1.4rem; font-size:.84rem; line-height:1.65 }
  .read-callout strong { color:var(--teal) }
  .recuadro { background:var(--card); border:1px solid var(--border);
              border-left:4px solid var(--red); border-radius:0 8px 8px 0;
              padding:1.1rem 1.3rem; margin:1.5rem 0 }
  .recuadro-title { font-size:.85rem; font-weight:700; color:var(--red); margin-bottom:.6rem }
  .recuadro p { font-size:.84rem; color:var(--mid); line-height:1.7; margin-bottom:.5rem }
  .recuadro p:last-child { margin-bottom:0 }

  /* ── Howto ── */
  .howto-grid { display:grid; grid-template-columns:1fr 1fr; gap:1.5rem }
  .howto-card { background:var(--card); border:1px solid var(--border);
                border-radius:10px; padding:1.25rem }
  .howto-card h3 { font-size:.95rem; font-weight:700; margin-bottom:.75rem;
                   padding-bottom:.5rem; border-bottom:1px solid var(--border) }
  .howto-step { display:flex; gap:.6rem; margin-bottom:.6rem; align-items:flex-start }
  .howto-icon { width:22px; height:22px; border-radius:50%;
                background:var(--red-l); color:var(--red);
                font-size:.72rem; font-weight:700;
                display:flex; align-items:center; justify-content:center;
                flex-shrink:0; margin-top:.1rem }
  .howto-text { font-size:.82rem; color:var(--mid); line-height:1.55 }

  /* ── Chart cards ── */
  .chart-card { background:var(--card); border-radius:var(--radius);
                box-shadow:var(--shadow-sm); border:1px solid var(--border);
                overflow:hidden; display:flex; flex-direction:column }
  .chart-card-header { display:flex; align-items:center; justify-content:space-between;
                       padding:.85rem 1rem .7rem; border-bottom:1px solid var(--border); gap:.75rem }
  .chart-card-header h3 { font-size:.88rem; font-weight:700; color:var(--ink); line-height:1.3; flex:1; margin:0 }
  .chart-type-tag { font-size:.62rem; font-weight:700; text-transform:uppercase;
                    letter-spacing:.07em; padding:.2rem .55rem; border-radius:20px; flex-shrink:0;
                    background:rgba(0,131,143,.15); color:#006064; border:1px solid rgba(0,131,143,.3) }
  .chart-canvas { position:relative; width:100%; height:320px; padding:.5rem }
  .chart-canvas.sm  { height:260px }
  .chart-canvas.md  { height:340px }
  .chart-canvas.lg  { height:420px }
  .chart-canvas.xl  { height:500px }

  /* ── Skeleton shimmer ── */
  @keyframes shimmer {
    0%   { background-position: -600px 0 }
    100% { background-position:  600px 0 }
  }
  .chart-skeleton {
    position:absolute; inset:0; border-radius:4px; z-index:2;
    background: linear-gradient(
      90deg,
      var(--bg) 0%, var(--bg) 30%,
      var(--border) 50%,
      var(--bg) 70%, var(--bg) 100%
    );
    background-size:600px 100%;
    animation: shimmer 1.6s ease-in-out infinite;
    pointer-events:none;
    transition: opacity .3s ease;
  }
  .chart-canvas.rendering .chart-skeleton { opacity:1 }
  .chart-canvas.ready     .chart-skeleton { opacity:0; pointer-events:none }

  /* ── Chart layouts ── */
  .chart-guide-grid { display:grid; grid-template-columns:3fr 2fr;
                      gap:1.25rem; align-items:start; margin-bottom:1.25rem }
  .chart-duo { display:grid; grid-template-columns:1fr 1fr; gap:1.25rem; margin-bottom:1.25rem }

  /* ── Figure cards (Modificaciones) ── */
  .grid { display:grid; gap:1.25rem }
  .grid-2 { grid-template-columns:repeat(2,1fr) }
  .grid-3 { grid-template-columns:repeat(3,1fr) }
  .card { background:var(--card); border-radius:var(--radius); box-shadow:var(--shadow-sm);
          border:1px solid var(--border); overflow:hidden; display:flex; flex-direction:column;
          transition:transform var(--ease),box-shadow var(--ease); cursor:pointer }
  .card:hover { transform:translateY(-4px); box-shadow:var(--shadow-md) }
  .card-thumb { position:relative; aspect-ratio:16/10; overflow:hidden; background:#fafafa }
  .card-thumb img { width:100%; height:100%; object-fit:cover; object-position:top left;
                    transition:transform .4s ease }
  .card:hover .card-thumb img { transform:scale(1.04) }
  .card-badge { position:absolute; top:.6rem; right:.6rem; font-size:.62rem; font-weight:700;
                text-transform:uppercase; letter-spacing:.07em; padding:.2rem .55rem; border-radius:20px }
  .badge-interactive { background:rgba(0,131,143,.15); color:#006064; border:1px solid rgba(0,131,143,.3) }
  .card-body { padding:.9rem 1rem 1rem; flex:1; display:flex; flex-direction:column }
  .card-title { font-size:.88rem; font-weight:700; line-height:1.3; margin-bottom:.3rem }
  .card-desc { font-size:.76rem; color:var(--muted); line-height:1.55; flex:1 }
  .card-footer { display:flex; align-items:center; justify-content:space-between;
                 margin-top:.75rem; padding-top:.65rem; border-top:1px solid var(--border) }
  .card-meta { font-size:.7rem; color:var(--muted); font-family:monospace }
  .card-btn { font-size:.72rem; font-weight:700; color:var(--red); letter-spacing:.04em;
              text-transform:uppercase; display:flex; align-items:center; gap:.3rem }
  .card-btn svg { width:12px; height:12px }

  /* ── Iframe ── */
  .fig-iframe-wrap { border-radius:var(--radius); overflow:hidden;
                     border:1px solid var(--border); box-shadow:var(--shadow-sm);
                     background:#fafafa; margin-bottom:1.25rem }
  .fig-iframe { display:block; width:100%; height:480px; border:none }

  /* ── Connection callout ── */
  .connect-box { background:var(--card); border:1px solid var(--border);
                 border-left:4px solid var(--amber); border-radius:0 8px 8px 0;
                 padding:1rem 1.25rem; margin:2rem 0 }
  .connect-box-title { font-size:.82rem; font-weight:700; color:var(--amber); margin-bottom:.5rem }
  .connect-links { display:flex; gap:.75rem; flex-wrap:wrap; margin-top:.75rem }
  .connect-link { display:inline-flex; align-items:center; gap:.4rem; font-size:.8rem;
                  font-weight:600; color:var(--red); padding:.35rem .85rem;
                  border:1.5px solid var(--red-l); background:var(--red-l);
                  border-radius:6px; transition:background var(--ease),border-color var(--ease) }
  .connect-link:hover { background:#fad5d3; border-color:#e0a09c }

  /* ── Scroll animations ── */
  .animate-in { opacity:0; transform:translateY(20px);
                transition:opacity .5s ease,transform .5s ease }
  .animate-in.delay-1 { transition-delay:.1s }
  .animate-in.delay-2 { transition-delay:.2s }
  .animate-in.delay-3 { transition-delay:.3s }
  .animate-in.visible { opacity:1; transform:none }

  /* ── Bottom nav (mobile) ── */
  .bottom-nav { display:none; position:fixed; bottom:0; left:0; right:0; z-index:200;
                background:rgba(255,255,255,.96); backdrop-filter:blur(12px);
                border-top:1px solid var(--border);
                padding:.3rem .5rem calc(.3rem + env(safe-area-inset-bottom)) }
  .bottom-nav-inner { display:flex; justify-content:space-around; align-items:center;
                      max-width:480px; margin:0 auto }
  .bottom-nav-item { display:flex; flex-direction:column; align-items:center;
                     gap:2px; padding:.4rem .6rem; border-radius:8px;
                     font-size:.58rem; font-weight:600; color:var(--muted);
                     text-transform:uppercase; letter-spacing:.05em;
                     transition:color var(--ease),background var(--ease); text-decoration:none }
  .bottom-nav-item:hover,.bottom-nav-item.active { color:var(--red); background:var(--red-l) }
  .bottom-nav-icon { font-size:1.1rem; line-height:1 }

  /* ── Footer ── */
  .footer { background:var(--ink); color:rgba(255,255,255,.4);
            text-align:center; padding:1.5rem 2rem; font-size:.75rem; line-height:1.8 }
  .footer strong { color:rgba(255,255,255,.75) }
  .divider { margin:3rem 0 0; border:none; border-top:1px solid var(--border) }

  /* ── Responsive ── */
  @media(max-width:960px) {
    .hero-inner { grid-template-columns:1fr }
    .palette { display:none }
    .howto-grid { grid-template-columns:1fr }
    .chart-guide-grid { grid-template-columns:1fr }
    .chart-duo { grid-template-columns:1fr }
    .chart-canvas { height:290px }
    .chart-canvas.lg { height:360px }
    .chart-canvas.xl { height:420px }
    .fig-iframe { height:380px }
    .grid-3 { grid-template-columns:repeat(2,1fr) }
  }
  @media(max-width:640px) {
    .nav { padding:0 1rem; position:fixed; width:100% }
    .nav-inner { gap:0; position:relative }
    .nav-hamburger { display:flex }
    .nav-links { display:none; position:absolute; top:52px; left:-1rem; right:-1rem;
                 background:rgba(255,255,255,.98); backdrop-filter:blur(16px);
                 border-bottom:1px solid var(--border); box-shadow:0 8px 24px rgba(0,0,0,.1);
                 padding:.5rem .75rem .75rem; flex-direction:column; gap:0; margin-left:0 }
    .nav-links.open { display:flex }
    .nav-links li { width:100% }
    .nav-links a { display:block; padding:.75rem 1rem; font-size:.9rem; border-radius:8px }
    .hero { padding:calc(52px + 2rem) 1.25rem 2rem }
    .hero h1 { font-size:clamp(1.4rem,6vw,1.9rem) }
    .hero-sub { font-size:.82rem }
    .hero-bar { margin:1rem 0 }
    .stats { display:grid; grid-template-columns:repeat(2,1fr); gap:.4rem }
    .stat { min-width:0 }
    .stat-val { font-size:1.6rem }
    .main { padding:0 1.25rem calc(3rem + 60px) }
    .section { margin-top:2rem }
    .section-header h2 { font-size:1.05rem }
    .grid-3 { grid-template-columns:1fr }
    .grid-2 { grid-template-columns:1fr }
    .card:hover { transform:none; box-shadow:var(--shadow-sm) }
    .chart-canvas { height:200px; padding:.25rem }
    .chart-canvas.sm  { height:200px }
    .chart-canvas.md  { height:260px }
    .chart-canvas.lg  { height:300px }
    .chart-canvas.xl  { height:310px }
    .fig-iframe { height:300px }
    .bottom-nav { display:block }
    .divider { margin:2rem 0 0 }
    .howto-card { padding:1rem }
    .chart-card-header h3 { font-size:.84rem }
  }
  @media(max-width:380px) {
    .chart-canvas { height:230px }
  }
"""

# ══════════════════════════════════════════════════════════════════════════════
# SHARED JS
# ══════════════════════════════════════════════════════════════════════════════
JS_SHARED = """
/* ── Page transitions ──────────────────────────────────────────────────────── */
document.querySelectorAll('a[href]').forEach(a => {
  const href = a.getAttribute('href');
  if (!href || href.startsWith('#') || href.startsWith('mailto') || href.startsWith('http')) return;
  a.addEventListener('click', e => {
    e.preventDefault();
    document.body.classList.add('page-exit');
    setTimeout(() => { window.location.href = href; }, 280);
  });
});

/* ── Scroll animations ──────────────────────────────────────────────────────── */
const observer = new IntersectionObserver(entries => {
  entries.forEach(e => { if (e.isIntersecting) { e.target.classList.add('visible'); observer.unobserve(e.target); }});
}, { threshold: 0.12 });
document.querySelectorAll('.animate-in').forEach(el => observer.observe(el));

/* ── Counter animation ──────────────────────────────────────────────────────── */
function animateCounter(el, target, duration=1100) {
  const raw = el.dataset.target || target;
  const num = parseFloat(String(raw).replace(/[^0-9.]/g,''));
  const isFloat = String(raw).includes('.');
  const suffix = String(raw).replace(/[0-9.,]/g,'');
  const start = performance.now();
  const step = now => {
    const t = Math.min(1, (now-start)/duration);
    const eased = 1 - Math.pow(1-t, 3);
    const val = eased * num;
    el.textContent = (isFloat ? val.toFixed(1) : Math.round(val).toLocaleString('es-MX')) + suffix;
    if (t < 1) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
}
const counterObserver = new IntersectionObserver(entries => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      animateCounter(e.target, e.target.dataset.target);
      counterObserver.unobserve(e.target);
    }
  });
}, { threshold: 0.5 });
document.querySelectorAll('.kpi-counter').forEach(el => counterObserver.observe(el));

/* ── Nav hamburger ──────────────────────────────────────────────────────────── */
function toggleNav() {
  const btn=document.getElementById('hamburger'), menu=document.getElementById('navLinks');
  const open=btn.classList.toggle('open');
  menu.classList.toggle('open', open);
  document.body.style.overflow = open ? 'hidden' : '';
}
function closeNav() {
  document.getElementById('hamburger')?.classList.remove('open');
  document.getElementById('navLinks')?.classList.remove('open');
  document.body.style.overflow='';
}
document.addEventListener('click', e => {
  if (!document.getElementById('mainNav')?.contains(e.target)) closeNav();
});

/* ── Bottom nav active ──────────────────────────────────────────────────────── */
(function(){
  const ids=['cat','meta','mod'], items=document.querySelectorAll('.bottom-nav-item');
  const els=ids.map(id=>document.getElementById(id)).filter(Boolean);
  function onScroll(){
    let ai=0; const y=window.scrollY+window.innerHeight*.35;
    els.forEach((el,i)=>{ if(el.offsetTop<=y) ai=i; });
    items.forEach((item,i)=>item.classList.toggle('active',i===ai&&item.getAttribute('href')?.startsWith('#')));
  }
  window.addEventListener('scroll',onScroll,{passive:true}); onScroll();
})();

/* ── Plotly unified config ──────────────────────────────────────────────────── */
const _mobile = () => window.innerWidth < 640;

const FONT = {
  family: 'Helvetica Neue,Arial,system-ui,sans-serif',
  size: _mobile() ? 10 : 11,
  color: '#444'
};

const CFG = { responsive:true, displayModeBar:false, scrollZoom:false };

/* Base layout applied to every chart — nowhere_minimal palette */
const LAYOUT_BASE = {
  paper_bgcolor: 'rgba(0,0,0,0)',
  plot_bgcolor:  'rgba(0,0,0,0)',
  font: FONT,
  xaxis: { gridcolor:'#E4E2DF', zerolinecolor:'#E4E2DF', tickfont: FONT },
  yaxis: { gridcolor:'#E4E2DF', zerolinecolor:'#E4E2DF', tickfont: FONT },
};

/* Returns margin object — compact on mobile, standard on desktop.
   `overrides` (optional) are merged in after the base so callers can
   set specific margins that still get clamped to mobile values on small screens. */
function getMargins(overrides) {
  if (_mobile()) return { l:6, r:6, t:6, b:32 };
  return Object.assign({ l:10, r:12, t:8, b:40 }, overrides || {});
}

const PROG_COLORS = {
  'Ambiental':'#2E7D32','Civil':'#BF360C','Computación':'#1565C0','Eléctrica':'#E65100',
  'Electrónica':'#6A1B9A','Física':'#00838F','Industrial':'#C62828',
  'Mecánica':'#37474F','Metalurgia':'#78909C','Química':'#558B2F'
};

/* ── renderChart — lazy viewport-based rendering ───────────────────────────── *
 *                                                                               *
 *  containerId  — id of the .chart-canvas div                                  *
 *  traceFn()    — function returning the traces array                           *
 *  layoutFn()   — function returning the layout object (no margin key needed)  *
 *  skeletonEl   — optional: the .chart-skeleton child (auto-detected if absent) *
 * ─────────────────────────────────────────────────────────────────────────── */
window.renderChart = function(containerId, traceFn, layoutFn) {
  const container = document.getElementById(containerId);
  if (!container) return;

  /* Inject skeleton if not present */
  let skel = container.querySelector('.chart-skeleton');
  if (!skel) {
    skel = document.createElement('div');
    skel.className = 'chart-skeleton';
    container.insertBefore(skel, container.firstChild);
  }

  container.classList.add('rendering');

  const io = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      io.unobserve(container);

      /* Merge layout: base → caller → mobile margin override */
      const callerLayout = layoutFn();
      const margin = _mobile()
        ? { l:6, r:6, t:6, b:32 }
        : (callerLayout.margin || { l:10, r:12, t:8, b:40 });

      const layout = Object.assign(
        {}, LAYOUT_BASE, callerLayout,
        {
          font:   FONT,
          margin: margin,
          /* Ensure axis font stays consistent regardless of caller */
          xaxis: Object.assign({}, LAYOUT_BASE.xaxis, callerLayout.xaxis || {}),
          yaxis: Object.assign({}, LAYOUT_BASE.yaxis, callerLayout.yaxis || {}),
        }
      );

      /* On mobile reduce legend font and move it below the chart */
      if (_mobile() && layout.legend) {
        layout.legend = Object.assign(
          { orientation:'h', xanchor:'center', x:.5, y:-.18 },
          layout.legend,
          { font: Object.assign({}, layout.legend.font || {}, { size:9 }) }
        );
      }

      requestAnimationFrame(() => {
        Plotly.newPlot(containerId, traceFn(), layout, CFG).then(() => {
          container.classList.remove('rendering');
          container.classList.add('ready');
        });
      });
    });
  }, { threshold: 0.15, rootMargin: '0px 0px 80px 0px' });

  io.observe(container);
};
"""

# ══════════════════════════════════════════════════════════════════════════════
# SHARED STRUCTURE HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def nav(active=''):
    links = [
        ('index.html', '🏠 Inicio'),
        ('cat26p.html', 'CAT 26P'),
        ('cat26p_capacidad.html', '↳ Capacidad'),
        ('metanalisis.html', 'Metaanálisis'),
        ('concentracion.html', 'Distribución docente'),
        ('modificaciones.html', 'Modificaciones 2026'),
    ]
    items = '\n'.join(
        f'<li><a href="{href}" class="{"active" if active==href else ""}" onclick="closeNav()">{label}</a></li>'
        for href, label in links
    )
    return f"""
<nav class="nav" id="mainNav">
  <div class="nav-inner">
    <a class="nav-brand" href="index.html">DCBI · 2026</a>
    <ul class="nav-links" id="navLinks">{items}</ul>
    <button class="nav-hamburger" id="hamburger" aria-label="Menú" onclick="toggleNav()">
      <span></span><span></span><span></span>
    </button>
  </div>
</nav>"""

def footer():
    return """
<footer class="footer">
  <strong>División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</strong><br>
  Sistemas de gestión académica institucional · Trimestre 26-P · 2026<br>
  Dr. César S. López-Monsalvo ·
  <a href="mailto:cslm@azc.uam.mx" style="color:rgba(255,255,255,.5);text-decoration:underline">cslm@azc.uam.mx</a>
</footer>"""

def palette_chips():
    plans = [('Ambiental','--c-amb'),('Civil','--c-civ'),('Computación','--c-com'),
             ('Eléctrica','--c-ele'),('Electrónica','--c-elo'),('Física','--c-fis'),
             ('Industrial','--c-ind'),('Mecánica','--c-mec'),('Metalurgia','--c-met'),('Química','--c-qui')]
    chips = '\n'.join(f'<div class="chip"><span class="chip-dot" style="background:var({cv})"></span>{n}</div>'
                      for n,cv in plans)
    return f'<div class="palette"><p class="palette-title">Licenciaturas DCBI</p><div class="palette-chips">{chips}</div></div>'

def chart_card(id, title, tag='Interactivo', size=''):
    """Return HTML for a .chart-card with lazy skeleton.

    size — extra CSS class for .chart-canvas ('sm', 'md', 'lg', 'xl', or '')
    The .chart-skeleton is injected here; renderChart() in JS will detect it.
    """
    size_cls = f' {size}' if size else ''
    return f"""<div class="chart-card">
  <div class="chart-card-header">
    <h3>{title}</h3>
    <span class="chart-type-tag">{tag}</span>
  </div>
  <div class="chart-canvas rendering{size_cls}" id="{id}">
    <div class="chart-skeleton"></div>
  </div>
</div>"""


def page_shell(title, active, content, bottom_nav_items='', extra_js='', data_vars=''):
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <meta name="theme-color" content="#C82D23">
  <title>{title} · DCBI UAM-A</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>{CSS}</style>
</head>
<body>
{nav(active)}
{content}
{footer()}
{'<nav class="bottom-nav" aria-label="Navegación rápida"><div class="bottom-nav-inner">' + bottom_nav_items + '</div></nav>' if bottom_nav_items else ''}
<script>
{data_vars}
{JS_SHARED}
{extra_js}
</script>
</body>
</html>"""

# ══════════════════════════════════════════════════════════════════════════════
# INDEX.HTML  — Landing page
# ══════════════════════════════════════════════════════════════════════════════
INDEX_CSS = """
  .landing-hero { min-height:100svh; background:var(--ink); color:#fff;
                  display:flex; flex-direction:column; justify-content:center;
                  padding:3rem 2rem; position:relative; overflow:hidden }
  .landing-hero::before { content:''; position:absolute; inset:0;
    background:radial-gradient(ellipse 80% 60% at 60% 40%, rgba(200,45,35,.12) 0%, transparent 70%);
    pointer-events:none }
  .landing-hero-inner { max-width:1440px; margin:0 auto; width:100%;
                        display:grid; grid-template-columns:1fr 1fr; gap:4rem; align-items:center }
  .landing-tag { font-size:.72rem; font-weight:700; letter-spacing:.16em;
                 text-transform:uppercase; color:var(--red); margin-bottom:1rem }
  .landing-h1 { font-size:clamp(2rem,5vw,4rem); font-weight:200; line-height:1.1;
                letter-spacing:-.03em }
  .landing-h1 strong { font-weight:800; display:block }
  .landing-sub { margin-top:1.2rem; font-size:1rem; color:rgba(255,255,255,.5);
                 max-width:520px; line-height:1.7 }
  .landing-bar { width:56px; height:3px; background:var(--red); margin:1.8rem 0 }
  .landing-kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:.75rem; max-width:520px }
  .lkpi { padding:.9rem; background:rgba(255,255,255,.06);
          border:1px solid rgba(255,255,255,.1); border-radius:8px; text-align:center }
  .lkpi-val { font-size:1.8rem; font-weight:700; color:#fff; line-height:1 }
  .lkpi-lbl { font-size:.65rem; color:rgba(255,255,255,.4); margin-top:.3rem; letter-spacing:.04em }
  .flow-diagram { display:flex; flex-direction:column; gap:1.5rem; padding:.5rem }
  .flow-node { background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.12);
               border-radius:12px; padding:1.4rem 1.6rem; cursor:pointer;
               transition:background var(--ease),border-color var(--ease),transform var(--ease) }
  .flow-node:hover { background:rgba(255,255,255,.09); border-color:rgba(200,45,35,.5);
                     transform:translateX(6px) }
  .flow-node-tag { font-size:.62rem; font-weight:700; letter-spacing:.12em;
                   text-transform:uppercase; color:var(--red); margin-bottom:.4rem }
  .flow-node-title { font-size:1.1rem; font-weight:700; color:#fff; margin-bottom:.3rem }
  .flow-node-desc { font-size:.8rem; color:rgba(255,255,255,.45); line-height:1.5 }
  .flow-node-stat { font-size:1.6rem; font-weight:800; color:#fff; margin-top:.6rem }
  .flow-node-stat-lbl { font-size:.68rem; color:rgba(255,255,255,.35) }
  .flow-arrow { text-align:center; color:var(--red); font-size:1.4rem; opacity:.6 }
  .scroll-hint { position:absolute; bottom:2rem; left:50%; transform:translateX(-50%);
                 font-size:.7rem; color:rgba(255,255,255,.3); text-align:center;
                 animation:bobble 2s ease-in-out infinite }
  @keyframes bobble { 0%,100%{transform:translateX(-50%) translateY(0)} 50%{transform:translateX(-50%) translateY(5px)} }
  @media(max-width:960px) {
    .landing-hero-inner { grid-template-columns:1fr; gap:2.5rem }
    .landing-kpis { grid-template-columns:repeat(2,1fr); max-width:none }
    .flow-diagram { flex-direction:row; gap:1rem }
    .flow-arrow { writing-mode:horizontal-tb; align-self:center; transform:rotate(0deg) }
    .flow-node { flex:1 }
  }
  @media(max-width:640px) {
    .landing-hero { min-height:auto; padding:calc(52px + 2rem) 1.25rem 2rem }
    .landing-h1 { font-size:clamp(1.8rem,8vw,2.8rem) }
    .flow-diagram { flex-direction:column }
    .flow-node:hover { transform:none }
  }
"""
INDEX_CONTENT = f"""
<style>{INDEX_CSS}</style>
<main>
  <div class="landing-hero">
    <div class="landing-hero-inner">
      <div>
        <p class="landing-tag animate-in">UAM Azcapotzalco · División de Ciencias Básicas e Ingeniería</p>
        <h1 class="landing-h1 animate-in delay-1">
          Gestión académica
          <strong>con datos</strong>
        </h1>
        <p class="landing-sub animate-in delay-2">
          Cuatro proyectos de análisis institucional que documentan, optimizan y proponen
          mejoras a la operación académica de las diez licenciaturas de la DCBI — desde
          la asignación de grupos hasta la revisión curricular con inteligencia artificial.
        </p>
        <div class="landing-bar animate-in delay-2"></div>
        <div class="landing-kpis animate-in delay-3">
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="674">0</div>
            <div class="lkpi-lbl">Estudiantes asignados</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="499">0</div>
            <div class="lkpi-lbl">Seriaciones analizadas</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="791">0</div>
            <div class="lkpi-lbl">Profesores en plantilla</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="1505">0</div>
            <div class="lkpi-lbl">UEAs evaluadas con IA</div>
          </div>
        </div>
      </div>
      <div class="flow-diagram animate-in delay-2">
        <div class="flow-node" onclick="window.location.href='cat26p.html'">
          <div class="flow-node-tag">Proyecto 1</div>
          <div class="flow-node-title">CAT 26P</div>
          <div class="flow-node-desc">Asignación automatizada de grupos para nuevo ingreso sin traslapes horarios</div>
          <div class="flow-node-stat kpi-counter" data-target="674">0</div>
          <div class="flow-node-stat-lbl">estudiantes procesados</div>
        </div>
        <div class="flow-arrow">↓</div>
        <div class="flow-node" onclick="window.location.href='metanalisis.html'">
          <div class="flow-node-tag">Proyecto 2</div>
          <div class="flow-node-title">Eficiencia terminal</div>
          <div class="flow-node-desc">Redes de prerrequisitos, complejidad estructural y análisis de eficiencia terminal</div>
          <div class="flow-node-stat kpi-counter" data-target="1859">0</div>
          <div class="flow-node-stat-lbl">UEAs en 10 planes</div>
        </div>
        <div class="flow-arrow">↓</div>
        <div class="flow-node" onclick="window.location.href='modificaciones.html'">
          <div class="flow-node-tag">Proyecto 3</div>
          <div class="flow-node-title">Modificaciones 2026</div>
          <div class="flow-node-desc">Revisión curricular automatizada con agentes de inteligencia artificial</div>
          <div class="flow-node-stat kpi-counter" data-target="1505">0</div>
          <div class="flow-node-stat-lbl">UEAs evaluadas con IA</div>
        </div>
        <div class="flow-arrow">↓</div>
        <div class="flow-node" onclick="window.location.href='concentracion.html'">
          <div class="flow-node-tag">Proyecto 4</div>
          <div class="flow-node-title">Distribución docente</div>
          <div class="flow-node-desc">Concentración curricular, riesgo estructural y vecindades de co-impartición</div>
          <div class="flow-node-stat kpi-counter" data-target="791">0</div>
          <div class="flow-node-stat-lbl">profesores en 10 planes</div>
        </div>
      </div>
    </div>
    <div class="scroll-hint">▾ desplázate para explorar</div>
  </div>
</main>
"""
index_html = page_shell(
    title='Gestión Académica DCBI',
    active='index.html',
    content=INDEX_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="cat26p.html"><span class="bottom-nav-icon">🗓</span>CAT</a>
      <a class="bottom-nav-item" href="metanalisis.html"><span class="bottom-nav-icon">📊</span>Meta</a>
      <a class="bottom-nav-item" href="modificaciones.html"><span class="bottom-nav-icon">📋</span>Modif.</a>
    """,
)

# ══════════════════════════════════════════════════════════════════════════════
# CAT26P.HTML  — dos partes: Resultados (público) | Sistema multiagente (técnico)
# ══════════════════════════════════════════════════════════════════════════════
CAT_CONTENT = f"""
<style>
  /* ── Part tabs ── */
  .part-tabs {{ display:flex; gap:0; border-bottom:2px solid var(--border);
               margin:0 0 0; background:var(--card); position:sticky; top:52px;
               z-index:50; overflow-x:auto; scrollbar-width:none }}
  .part-tabs::-webkit-scrollbar {{ display:none }}
  .part-tab {{ flex:none; padding:.75rem 1.5rem; font-size:.84rem; font-weight:600;
              color:var(--muted); border-bottom:2.5px solid transparent;
              margin-bottom:-2px; cursor:pointer; white-space:nowrap;
              transition:color var(--ease),border-color var(--ease);
              background:none; border-top:none; border-left:none; border-right:none;
              font-family:var(--font) }}
  .part-tab:hover {{ color:var(--ink) }}
  .part-tab.active {{ color:var(--red); border-bottom-color:var(--red) }}
  .part-panel {{ display:none }}
  .part-panel.active {{ display:block }}

  /* ── Guarantee grid ── */
  .guarantee-grid {{ display:grid; grid-template-columns:repeat(3,1fr); gap:1rem; margin-top:1.25rem }}
  .guarantee-item {{ background:var(--card); border:1px solid var(--border);
                    border-top:3px solid var(--red); border-radius:var(--radius);
                    padding:1rem 1.1rem }}
  .guarantee-item h4 {{ font-size:.85rem; font-weight:700; margin-bottom:.4rem; color:var(--red) }}
  .guarantee-item p {{ font-size:.8rem; color:var(--mid); line-height:1.6 }}

  /* ── Gender callout ── */
  .gender-callout {{ background:#f3eefb; border-left:4px solid #6A1B9A;
                    border-radius:0 8px 8px 0; padding:1.1rem 1.3rem; margin:1.5rem 0 }}
  .gender-callout-title {{ font-size:.85rem; font-weight:700; color:#6A1B9A; margin-bottom:.7rem }}
  .gender-callout p {{ font-size:.83rem; color:#3d1a5c; line-height:1.7; margin-bottom:.6rem }}
  .gender-callout p:last-child {{ margin-bottom:0 }}

  /* ── Agent flow diagram ── */
  .agent-flow {{ display:flex; flex-direction:column; gap:0; margin:1.5rem 0 }}
  .agent-row {{ display:flex; align-items:stretch; gap:0 }}
  .agent-inputs {{ display:flex; flex-direction:column; gap:.5rem; justify-content:center;
                  padding-right:1rem; min-width:120px }}
  .agent-input-tag {{ font-size:.68rem; font-weight:700; background:var(--bg);
                     border:1px solid var(--border); border-radius:4px;
                     padding:.25rem .6rem; text-align:right; color:var(--muted);
                     white-space:nowrap }}
  .agent-input-arrow {{ font-size:.9rem; color:var(--border); text-align:right }}
  .agent-pipeline {{ flex:1; display:flex; flex-direction:column; align-items:stretch }}
  .agent-node {{ background:var(--card); border:1.5px solid var(--border);
                border-radius:var(--radius); padding:.9rem 1.1rem;
                display:grid; grid-template-columns:auto 1fr auto; gap:.75rem;
                align-items:center; position:relative }}
  .agent-node.coordinator {{ border-color:var(--red); background:var(--red-l) }}
  .agent-phase-badge {{ font-size:.6rem; font-weight:800; text-transform:uppercase;
                       letter-spacing:.08em; background:var(--red); color:#fff;
                       padding:.2rem .55rem; border-radius:20px; white-space:nowrap }}
  .agent-node.coordinator .agent-phase-badge {{ background:var(--ink) }}
  .agent-name {{ font-size:.9rem; font-weight:700; color:var(--ink) }}
  .agent-desc {{ font-size:.76rem; color:var(--mid); margin-top:.15rem; line-height:1.4 }}
  .agent-complexity {{ font-family:monospace; font-size:.72rem; color:var(--muted);
                      background:var(--bg); border:1px solid var(--border);
                      padding:.2rem .5rem; border-radius:4px; white-space:nowrap;
                      align-self:start; margin-top:.1rem }}
  .agent-connector {{ display:flex; align-items:center; justify-content:center;
                     height:28px; color:var(--border); font-size:1rem }}
  .agent-outputs {{ display:flex; flex-direction:column; gap:.5rem; justify-content:center;
                   padding-left:1rem; min-width:130px }}
  .agent-output-tag {{ font-size:.68rem; font-weight:700; background:rgba(56,142,60,.1);
                      border:1px solid rgba(56,142,60,.3); border-radius:4px; color:#2E7D32;
                      padding:.25rem .6rem; white-space:nowrap }}

  /* ── Complexity table ── */
  .complexity-table {{ width:100%; border-collapse:collapse; font-size:.82rem; margin-top:1rem }}
  .complexity-table th {{ background:var(--ink); color:#fff; padding:.55rem .8rem;
                         font-size:.75rem; font-weight:700; text-align:left }}
  .complexity-table td {{ padding:.55rem .8rem; border-bottom:1px solid var(--border);
                         vertical-align:top; line-height:1.5 }}
  .complexity-table tr:last-child td {{ border-bottom:none }}
  .complexity-table tr:nth-child(even) td {{ background:#fafafa }}
  .complexity-table code {{ font-size:.78rem; background:var(--bg); padding:.1rem .3rem;
                           border-radius:3px; border:1px solid var(--border) }}

  /* ── Search space callout ── */
  .math-callout {{ background:#fff8e1; border-left:4px solid var(--amber);
                  border-radius:0 8px 8px 0; padding:1rem 1.2rem; margin:1rem 0 }}
  .math-callout p {{ font-size:.83rem; color:#4a3800; line-height:1.7; margin-bottom:.5rem }}
  .math-callout p:last-child {{ margin-bottom:0 }}
  .math-num {{ font-family:monospace; font-weight:700; font-size:.95em }}

  @media(max-width:960px) {{
    .guarantee-grid {{ grid-template-columns:repeat(2,1fr) }}
    .agent-inputs,.agent-outputs {{ display:none }}
  }}
  @media(max-width:640px) {{
    .part-tab {{ padding:.65rem 1rem; font-size:.8rem }}
    .guarantee-grid {{ grid-template-columns:1fr }}
    .agent-node {{ grid-template-columns:auto 1fr }}
    .agent-complexity {{ display:none }}
    .complexity-table {{ font-size:.75rem }}
    .complexity-table th,.complexity-table td {{ padding:.4rem .55rem }}
  }}
</style>

<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">CAT 26P · Tronco General · Trimestre 2026-P</p>
      <h1><strong>Sistema de asignación</strong><br>automática de grupos</h1>
      <p class="hero-sub">Cada trimestre, la División asigna a cientos de estudiantes de nuevo
        ingreso a los seis grupos del Tronco General. Verificar sin traslapes, respetar cupos
        y garantizar cohesión por plan de estudios es inviable a mano. Este sistema lo resuelve
        de forma automática, completa y verificable en menos de un segundo.</p>
      <div class="hero-bar"></div>
      <div class="stats">
        <div class="stat"><span class="stat-val kpi-counter" data-target="674">0</span><span class="stat-lbl">Estudiantes procesados</span></div>
        <div class="stat"><span class="stat-val">0</span><span class="stat-lbl">Conflictos horarios</span></div>
        <div class="stat"><span class="stat-val">&lt;1 s</span><span class="stat-lbl">Tiempo de proceso</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="864">0</span><span class="stat-lbl">Capacidad máxima del catálogo</span></div>
      </div>
    </div>
    {palette_chips()}
  </div>
</header>

<!-- ── Part tabs ── -->
<div class="part-tabs" id="catTabs">
  <button class="part-tab active" onclick="switchPart('resultados',this)">Resultados del ciclo 26-P</button>
  <button class="part-tab" onclick="switchPart('tecnico',this)">Sistema multiagente</button>
</div>

<main class="main">

<!-- ════════════════════════════════════════════════════════════════
     PARTE 1 — RESULTADOS (público general)
════════════════════════════════════════════════════════════════ -->
<div class="part-panel active" id="panel-resultados">

  <!-- Cobertura -->
  <section class="section animate-in" id="cat-cob">
    <div class="section-header">
      <h2>Resultado de la asignación</h2>
      <span class="section-count">674 / 674 · 100 %</span>
    </div>
    <p class="section-desc">
      El sistema procesó a los 674 estudiantes de nuevo ingreso del ciclo 26-Primavera
      y asignó a <strong>cada uno de ellos sus seis materias del Tronco General</strong>
      sin ningún conflicto de horario. Cobertura: 100 %. Conflictos detectados: 0.
      El resultado fue verificado de forma independiente por el agente auditor al finalizar
      la ejecución. El catálogo 26-P tiene capacidad para 864 estudiantes —
      la División tiene 190 plazas de holgura.
    </p>
    <div class="read-callout">
      <strong>Cómo leer la gráfica:</strong> el círculo muestra la distribución del nuevo
      ingreso por género — la variable de mayor relevancia operativa dado que el sistema
      aplica prioridad matutina para alumnas. El 27.6 % de alumnas (186) y el 72.4 % de
      alumnos (488) reciben exactamente el mismo resultado: seis materias completas, cero traslapes.
    </div>
    <div class="chart-guide-grid animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Distribución por género · 674 estudiantes de nuevo ingreso</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering md" id="chart-donut"><div class="chart-skeleton"></div></div>
      </div>
      <div class="howto-card">
        <h3>Cómo leer esta gráfica</h3>
        <div class="howto-step"><div class="howto-icon">1</div>
          <div class="howto-text"><strong>100 % asignados con 6 materias.</strong> Todos los 674 estudiantes recibieron sus seis grupos del Tronco General sin ningún traslape.</div></div>
        <div class="howto-step"><div class="howto-icon">2</div>
          <div class="howto-text"><strong>El género importa en la asignación.</strong> Las 186 alumnas reciben prioridad matutina: son asignadas antes que los alumnos varones y el sistema les ofrece primero los grupos de mañana.</div></div>
        <div class="howto-step"><div class="howto-icon">3</div>
          <div class="howto-text"><strong>Proporción histórica.</strong> La distribución 72/28 M/F es consistente con las estadísticas de ingeniería en México; varía entre licenciaturas.</div></div>
        <div class="howto-step"><div class="howto-icon">4</div>
          <div class="howto-text"><strong>Sin asignación parcial: 0 estudiantes.</strong> El catálogo tiene capacidad para 864 alumnos; el padrón de 674 no satura ninguna materia.</div></div>
      </div>
    </div>
  </section>

  <!-- Grupos llenos por UEA -->
  <section class="section animate-in" id="cat-ueas">
    <div class="section-header">
      <h2>Grupos que alcanzaron su cupo, por materia</h2>
      <span class="section-count">38 grupos al 100 %</span>
    </div>
    <p class="section-desc">
      De los grupos CAT asignados en el ciclo 26-P, 38 llegaron al 100 % de su cupo.
      El Laboratorio de Reacciones Químicas concentra el mayor número (13 grupos llenos)
      porque su cupo es 36 plazas — menor que el de las materias teóricas (45).
      Las cinco materias teóricas tienen exactamente 5 grupos llenos cada una,
      reflejo de que el algoritmo asigna las seis materias en bloque y el patrón
      de ocupación es uniforme entre ellas.
    </p>
    <div class="read-callout">
      <strong>Qué buscar:</strong> la longitud de cada barra indica cuántos grupos de esa
      materia alcanzaron el 100 % de cupo. El Laboratorio lidera porque su cupo por grupo
      es menor (36 vs 45). Que haya grupos llenos no significa que algún estudiante quedara
      sin asignación — el sistema los ubicó en otros grupos disponibles. Ningún estudiante
      quedó sin sus seis materias.
    </div>
    <div class="chart-guide-grid animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Grupos al 100 % de cupo por materia del Tronco General</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering" id="chart-ueas"><div class="chart-skeleton"></div></div>
      </div>
      <div class="howto-card">
        <h3>Cómo leer esta gráfica</h3>
        <div class="howto-step"><div class="howto-icon">1</div>
          <div class="howto-text"><strong>Cada barra = una materia.</strong> La longitud indica cuántos grupos llegaron exactamente al 100 % de cupo en el ciclo 26-P.</div></div>
        <div class="howto-step"><div class="howto-icon">2</div>
          <div class="howto-text"><strong>Lab. Reacciones Quím.: 13 grupos llenos</strong> de 21 asignados. Cupo por grupo: 36 plazas.</div></div>
        <div class="howto-step"><div class="howto-icon">3</div>
          <div class="howto-text"><strong>Las 5 materias teóricas: 5 grupos llenos</strong> cada una de 17 asignados. Cupo: 45 plazas. Patrón uniforme: la asignación en bloque los sincroniza.</div></div>
        <div class="howto-step"><div class="howto-icon">4</div>
          <div class="howto-text"><strong>El grupo CAT07</strong> superó su cupo nominal (asignado a un aula con mayor aforo por convenio de sala). Está incluido como "lleno" en el conteo.</div></div>
      </div>
    </div>
  </section>

  <!-- Distribución por licenciatura -->
  <section class="section animate-in" id="cat-planes">
    <div class="section-header">
      <h2>Nuevo ingreso por licenciatura</h2>
      <span class="section-count">10 licenciaturas · 674 estudiantes</span>
    </div>
    <p class="section-desc">
      Ingeniería en Computación concentra el mayor nuevo ingreso (166 estudiantes, 24.6 %),
      seguida de Mecánica (112) e Ingeniería Civil (84). La gráfica de género por licenciatura
      muestra la distribución de alumnas y alumnos en cada programa — información relevante para
      la planeación de grupos matutinos. El sistema garantiza que todos los estudiantes de un
      mismo plan de estudios queden en el mismo grupo de Introducción a la Ingeniería.
    </p>
    <div class="read-callout">
      <strong>Qué muestran estas gráficas:</strong> la izquierda ordena las licenciaturas
      por tamaño de nuevo ingreso. La derecha muestra la composición por género en cada una.
      Licenciaturas con mayor proporción femenina (como Química o Ambiental) requieren más
      plazas en grupos matutinos — el sistema lo calcula automáticamente sin configuración adicional.
    </div>
    <div class="chart-duo animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Nuevo ingreso por licenciatura</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering" id="chart-byplan"><div class="chart-skeleton"></div></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Distribución de género por licenciatura</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering" id="chart-stacked"><div class="chart-skeleton"></div></div>
      </div>
    </div>
  </section>

  <!-- Garantías -->
  <section class="section animate-in" id="cat-garantias">
    <div class="section-header">
      <h2>Lo que el sistema garantiza</h2>
    </div>
    <p class="section-desc">
      El sistema no solo produce una asignación: la certifica. Al finalizar cada ejecución,
      un agente independiente audita el resultado y emite un reporte de validación.
      Las siguientes garantías constan en ese reporte para el ciclo 26-Primavera.
    </p>
    <div class="guarantee-grid animate-in delay-1">
      <div class="guarantee-item">
        <h4>Cero conflictos de horario</h4>
        <p>Ningún estudiante recibe dos grupos que coincidan en día y hora. Verificado de forma automática tras cada ejecución.</p>
      </div>
      <div class="guarantee-item">
        <h4>Cobertura total del padrón</h4>
        <p>Mientras el padrón no supere la capacidad del catálogo, todos los estudiantes reciben sus seis grupos. No hay estudiantes sin ninguna asignación.</p>
      </div>
      <div class="guarantee-item">
        <h4>Prioridad matutina para alumnas</h4>
        <p>Las 186 alumnas son asignadas antes que cualquier alumno y reciben preferencia en grupos cuyas franjas inician antes de las 14:00 h.</p>
      </div>
      <div class="guarantee-item">
        <h4>Cohesión por plan de estudios</h4>
        <p>Todo el alumnado de un mismo plan comparte el mismo grupo de Introducción a la Ingeniería. Los planes no se mezclan en un mismo carril.</p>
      </div>
      <div class="guarantee-item">
        <h4>Resultados reproducibles</h4>
        <p>Con el mismo padrón y catálogo, el sistema produce siempre la misma asignación. Los resultados son auditables trimestre a trimestre.</p>
      </div>
      <div class="guarantee-item">
        <h4>Sin intervención de Sistemas Escolares</h4>
        <p>Para un nuevo trimestre basta actualizar el padrón y el catálogo de horarios. No se requiere modificar el código ni intervención adicional.</p>
      </div>
    </div>
  </section>

  <!-- Género -->
  <section class="section animate-in" id="cat-genero">
    <div class="section-header">
      <h2>¿Por qué algunos grupos son exclusivamente de alumnas?</h2>
    </div>
    <div class="gender-callout">
      <div class="gender-callout-title">Una regla de equidad explícita, no un error del sistema</div>
      <p>El sistema aplica una política de protección para el alumnado femenino: al asignar a
      cada alumna, le ofrece primero los grupos del turno matutino. La razón es de seguridad y
      bienestar — los carriles vespertinos pueden concluir a las 21:00 h, lo que implica que el
      alumnado sale de las instalaciones en horario nocturno. La prioridad matutina busca evitar
      que las alumnas queden expuestas a esa condición de forma sistemática.</p>
      <p>Dado que el sistema asigna a todas las alumnas antes que a cualquier alumno varón, y les
      ofrece primero los horarios de la mañana, cuatro grupos matutinos se completan
      exclusivamente con alumnas antes de que comience la asignación masculina. Cuando los 488
      alumnos varones empiezan a ser asignados, esos grupos están llenos. El resultado — grupos
      compuestos únicamente por alumnas — es el efecto esperado de aplicar la prioridad matutina
      con rigor.</p>
      <p><strong>Si se prefiere una composición mixta en todos los grupos,</strong> el sistema
      cuenta con un modo de reserva que garantiza un número mínimo de plazas para alumnos varones
      en cada grupo matutino antes de completarlo con alumnas. Activar ese modo no requiere
      modificar el catálogo de horarios ni el padrón.</p>
    </div>
  </section>

  <!-- Subpágina capacidad -->
  <section class="section animate-in" id="cat-cap-link">
    <a href="cat26p_capacidad.html" style="display:block;text-decoration:none">
      <div style="background:var(--card);border:1px solid var(--border);border-left:4px solid var(--red);
                  border-radius:0 10px 10px 0;padding:1.4rem 1.6rem;
                  transition:box-shadow var(--ease),transform var(--ease);
                  display:grid;grid-template-columns:1fr auto;align-items:center;gap:1rem"
           onmouseover="this.style.boxShadow='var(--shadow-md)';this.style.transform='translateY(-2px)'"
           onmouseout="this.style.boxShadow='none';this.style.transform='none'">
        <div>
          <div style="font-size:.7rem;font-weight:700;color:var(--red);text-transform:uppercase;letter-spacing:.1em;margin-bottom:.35rem">Análisis detallado · CAT 26P</div>
          <div style="font-size:1.05rem;font-weight:700;margin-bottom:.4rem">Capacidad máxima de admisión</div>
          <div style="font-size:.83rem;color:var(--mid);line-height:1.55">
            Capacidad instalada por materia · Ocupación por grupo · Cuello de botella ·
            Sensibilidad ante grupos adicionales · Distribución de género
          </div>
          <div style="margin-top:.75rem;display:flex;gap:1rem;flex-wrap:wrap">
            <span style="font-size:.75rem;font-weight:600;color:var(--ink)">864 plazas máximas</span>
            <span style="font-size:.75rem;color:var(--muted)">·</span>
            <span style="font-size:.75rem;font-weight:600;color:var(--green)">190 plazas libres hoy</span>
            <span style="font-size:.75rem;color:var(--muted)">·</span>
            <span style="font-size:.75rem;font-weight:600;color:var(--amber)">Lab. Reacciones = cuello de botella</span>
          </div>
        </div>
        <div style="font-size:2rem;color:var(--red);opacity:.6;flex-shrink:0">→</div>
      </div>
    </a>
  </section>

  <hr class="divider">
  <div class="connect-box animate-in" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem">
    <div>
      <div class="connect-box-title">Resumen ejecutivo</div>
      <p style="font-size:.84rem;color:var(--mid);line-height:1.6;margin:0">Resultados del ciclo 26-P: asignaciones, distribución y garantías · PDF</p>
    </div>
    <a href="resumen_ejecutivo_cat26p.pdf" download
       style="display:inline-flex;align-items:center;gap:.5rem;background:var(--red);color:#fff;
              text-decoration:none;padding:.7rem 1.4rem;border-radius:8px;font-size:.85rem;
              font-weight:700;white-space:nowrap;transition:opacity .15s"
       onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      ↓ Descargar PDF
    </a>
  </div>

  <div class="connect-box animate-in" style="margin-top:1rem">
    <div class="connect-box-title">¿Cómo se conecta con los demás proyectos?</div>
    <p style="font-size:.84rem;color:var(--mid);line-height:1.6">
      Los 674 estudiantes asignados por este sistema son los mismos cuyas trayectorias
      curriculares analiza el Metaanálisis. La distribución por licenciatura que se ve aquí
      determina el punto de entrada a las redes de prerrequisitos cuya complejidad estructural
      mide el índice C — y sobre cuyos programas trabajó el análisis de Modificaciones 2026.
    </p>
    <div class="connect-links">
      <a class="connect-link" href="metanalisis.html">📊 Ver Metaanálisis →</a>
      <a class="connect-link" href="modificaciones.html">📋 Ver Modificaciones 2026 →</a>
    </div>
  </div>

</div><!-- /panel-resultados -->

<!-- ════════════════════════════════════════════════════════════════
     PARTE 2 — SISTEMA MULTIAGENTE (técnico)
════════════════════════════════════════════════════════════════ -->
<div class="part-panel" id="panel-tecnico">

  <!-- Descripción general -->
  <section class="section animate-in" id="cat-arq">
    <div class="section-header">
      <h2>Arquitectura multi-agente orientada a flujo de trabajo</h2>
    </div>
    <p class="section-desc">
      El sistema sigue el patrón <em>multi-agente orientado a flujo de trabajo</em>: un
      orquestador central invoca secuencialmente cinco agentes especializados, cada uno
      responsable de una fase bien delimitada. La separación entre la verificación previa
      (Fase 1) y la auditoría posterior (Fase 5) permite diagnosticar cualquier anomalía
      de forma precisa sin repetir el proceso completo.
    </p>
    <div class="math-callout animate-in delay-1">
      <p><strong>¿Por qué no revisión exhaustiva?</strong> Con 674 estudiantes y hasta
      26 grupos por materia, el espacio de combinaciones posibles para un solo estudiante
      es <span class="math-num">26⁶ = 308,915,776</span> (más de 300 millones). Para los
      674 estudiantes en simultáneo, el espacio conjunto crece a
      <span class="math-num">26⁴⁰⁴⁴ ≈ 10⁵⁷²²</span> configuraciones — un número
      10⁵⁶⁴² veces mayor que el número estimado de átomos en el universo observable.
      La revisión exhaustiva es, en la práctica, imposible. El algoritmo voraz del sistema
      resuelve el problema en tiempo <span class="math-num">O(N)</span>, alcanzando
      coberturas superiores al 99 %.</p>
    </div>
  </section>

  <!-- Diagrama de flujo -->
  <section class="section animate-in" id="cat-flujo">
    <div class="section-header">
      <h2>Flujo de operación del sistema</h2>
      <span class="section-count">5 fases + orquestador</span>
    </div>
    <p class="section-desc">
      El sistema recibe dos insumos: el <strong>padrón de nuevo ingreso</strong>
      (datos de los estudiantes a asignar) y el <strong>catálogo de horarios</strong>
      (grupos disponibles con sus franjas). A partir de ahí, el orquestador activa
      las cinco fases en secuencia y produce dos salidas: la
      <strong>tabla de asignaciones</strong> y el <strong>reporte de validación</strong>.
    </p>

    <div class="agent-flow animate-in delay-1">

      <!-- Inputs -->
      <div class="agent-row" style="margin-bottom:.6rem">
        <div class="agent-inputs">
          <div class="agent-input-tag">Padrón de nuevo ingreso</div>
          <div class="agent-input-tag">Catálogo de horarios</div>
        </div>
        <div class="agent-pipeline" style="justify-content:center;padding-top:.4rem">
          <div style="font-size:.8rem;color:var(--muted);text-align:center">▼ insumos del sistema</div>
        </div>
        <div class="agent-outputs"></div>
      </div>

      <!-- Coordinator -->
      <div class="agent-row">
        <div class="agent-inputs"></div>
        <div class="agent-pipeline">
          <div class="agent-node coordinator">
            <span class="agent-phase-badge">Orquestador</span>
            <div>
              <div class="agent-name">Coordinador</div>
              <div class="agent-desc">Recibe los insumos, activa las cinco fases en secuencia y centraliza el estado compartido entre agentes.</div>
            </div>
            <div class="agent-complexity">O(1)</div>
          </div>
          <div class="agent-connector">↓</div>
        </div>
        <div class="agent-outputs"></div>
      </div>

      <!-- Fase 1 -->
      <div class="agent-row">
        <div class="agent-inputs">
          <div class="agent-input-tag">Catálogo de horarios</div>
          <div class="agent-input-arrow">→</div>
        </div>
        <div class="agent-pipeline">
          <div class="agent-node">
            <span class="agent-phase-badge">Fase 1</span>
            <div>
              <div class="agent-name">Verificador de propuesta</div>
              <div class="agent-desc">Detecta conflictos de horario <em>dentro del catálogo</em> antes de asignar a nadie. Si dos grupos del mismo carril CAT se traslapan, reporta el problema al coordinador sin continuar.</div>
            </div>
            <div class="agent-complexity">O(K · C(G,2) · S²)</div>
          </div>
          <div class="agent-connector">↓</div>
        </div>
        <div class="agent-outputs">
          <div class="agent-input-arrow">←</div>
          <div class="agent-output-tag">Reporte de conflictos intra-CAT</div>
        </div>
      </div>

      <!-- Fase 2 -->
      <div class="agent-row">
        <div class="agent-inputs">
          <div class="agent-input-tag">Padrón de nuevo ingreso</div>
          <div class="agent-input-arrow">→</div>
        </div>
        <div class="agent-pipeline">
          <div class="agent-node">
            <span class="agent-phase-badge">Fase 2</span>
            <div>
              <div class="agent-name">Procesador de datos</div>
              <div class="agent-desc">Lee e interpreta el padrón y el catálogo. Construye las estructuras internas del motor de asignación: tablas de búsqueda indexadas por materia, franjas horarias normalizadas y lista de estudiantes ordenada por prioridad.</div>
            </div>
            <div class="agent-complexity">O(N + |G|)</div>
          </div>
          <div class="agent-connector">↓</div>
        </div>
        <div class="agent-outputs"></div>
      </div>

      <!-- Fase 3 -->
      <div class="agent-row">
        <div class="agent-inputs"></div>
        <div class="agent-pipeline">
          <div class="agent-node" style="border-color:var(--red)">
            <span class="agent-phase-badge">Fase 3</span>
            <div>
              <div class="agent-name">Motor de asignación</div>
              <div class="agent-desc">Núcleo del sistema. Itera sobre los N estudiantes en orden de prioridad (alumnas primero). Para cada estudiante y cada materia, selecciona el primer grupo factible que satisfaga las cinco restricciones: cupo (R1), prefijo CAT (R2), exclusión de carriles complemento (R3), agrupamiento por plan (R4) y cero traslapes (R5).</div>
            </div>
            <div class="agent-complexity">O(N · U² · G · S)</div>
          </div>
          <div class="agent-connector">↓</div>
        </div>
        <div class="agent-outputs"></div>
      </div>

      <!-- Fase 4 -->
      <div class="agent-row">
        <div class="agent-inputs"></div>
        <div class="agent-pipeline">
          <div class="agent-node">
            <span class="agent-phase-badge">Fase 4</span>
            <div>
              <div class="agent-name">Generador de resultados</div>
              <div class="agent-desc">Produce la tabla de asignaciones en formato de uso administrativo y el reporte de métricas con totales por materia, por licenciatura y por tipo de asignación.</div>
            </div>
            <div class="agent-complexity">O(N · U)</div>
          </div>
          <div class="agent-connector">↓</div>
        </div>
        <div class="agent-outputs">
          <div class="agent-input-arrow">←</div>
          <div class="agent-output-tag">Tabla de asignaciones</div>
          <div class="agent-output-tag">Reporte de métricas</div>
        </div>
      </div>

      <!-- Fase 5 -->
      <div class="agent-row">
        <div class="agent-inputs"></div>
        <div class="agent-pipeline">
          <div class="agent-node">
            <span class="agent-phase-badge">Fase 5</span>
            <div>
              <div class="agent-name">Auditor independiente</div>
              <div class="agent-desc">Verifica de forma independiente, sin acceso al código del motor, que la solución final esté libre de traslapes. Para cada estudiante revisa todos los pares de materias asignadas. El resultado consta en el reporte de validación.</div>
            </div>
            <div class="agent-complexity">O(N · U² · S²)</div>
          </div>
        </div>
        <div class="agent-outputs">
          <div class="agent-input-arrow">←</div>
          <div class="agent-output-tag">Reporte de validación</div>
        </div>
      </div>

    </div><!-- /agent-flow -->
  </section>

  <!-- Complejidad -->
  <section class="section animate-in" id="cat-complejidad">
    <div class="section-header">
      <h2>Complejidad algorítmica por fase</h2>
    </div>
    <p class="section-desc">
      El costo computacional total del sistema es del orden de 3.8 × 10⁶ operaciones
      elementales — consistente con la ejecución observada de menos de un segundo.
      La Fase 3 (motor de asignación) concentra el 83 % del costo; la Fase 5
      (auditoría) el 16 % restante. Las fases 1, 2 y 4 son negligibles en comparación.
      Los parámetros son: N = 674 estudiantes, U = 6 materias por estudiante,
      G = 26 grupos máximos por materia, S = 5 franjas horarias máximas por grupo.
    </p>
    <table class="complexity-table animate-in delay-1">
      <thead>
        <tr>
          <th>Fase</th>
          <th>Agente</th>
          <th>Operación dominante</th>
          <th>Complejidad</th>
          <th>Operaciones (26-P)</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><strong>1</strong></td>
          <td>Verificador de propuesta</td>
          <td>Pares conflictivos intra-CAT</td>
          <td><code>O(K·C(G,2)·S²)</code></td>
          <td>3,600</td>
        </tr>
        <tr>
          <td><strong>2</strong></td>
          <td>Procesador de datos</td>
          <td>Lectura de padrón y catálogo</td>
          <td><code>O(N + |G|)</code></td>
          <td>1,557</td>
        </tr>
        <tr>
          <td><strong>3</strong></td>
          <td>Motor de asignación</td>
          <td>Asignación voraz con verificación de restricciones</td>
          <td><code>O(N·U²·G·S)</code></td>
          <td>3,157,680</td>
        </tr>
        <tr>
          <td><strong>4</strong></td>
          <td>Generador de resultados</td>
          <td>Escritura de tabla de asignaciones</td>
          <td><code>O(N·U)</code></td>
          <td>4,044</td>
        </tr>
        <tr>
          <td><strong>5</strong></td>
          <td>Auditor independiente</td>
          <td>Verificación de pares de materias por estudiante</td>
          <td><code>O(N·U²·S²)</code></td>
          <td>606,600</td>
        </tr>
        <tr style="background:#fafafa;font-weight:700">
          <td colspan="4" style="text-align:right">Total aproximado</td>
          <td>≈ 3.8 × 10⁶</td>
        </tr>
      </tbody>
    </table>

    <div class="recuadro animate-in delay-2" style="margin-top:1.5rem">
      <div class="recuadro-title">Las cinco restricciones del sistema</div>
      <p><strong>R1 — Cupo:</strong> ningún grupo puede recibir más estudiantes que su cupo máximo (36 para el laboratorio, 45 para las materias teóricas).</p>
      <p><strong>R2 — Prefijo CAT:</strong> solo se asignan grupos con identificador que comience con "CAT". Los grupos de servicio social, intercambio y educación continua se descartan automáticamente.</p>
      <p><strong>R3 — Exclusión de carriles complemento:</strong> los carriles CAT12–CAT18 son laboratorios adicionales; las cinco materias teóricas no tienen sección en esos carriles y no pueden asignarse ahí.</p>
      <p><strong>R4 — Agrupamiento por plan:</strong> todo el alumnado de un mismo plan de estudios comparte el mismo grupo de Introducción a la Ingeniería. El número de carriles asignados por plan escala con el tamaño del plan.</p>
      <p><strong>R5 — Cero traslapes:</strong> ningún par de grupos asignados al mismo estudiante puede compartir una franja horaria. Es la restricción más compleja de verificar y la que domina la complejidad de la Fase 3.</p>
    </div>
  </section>

  <hr class="divider">
  <div class="connect-box animate-in" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem">
    <div>
      <div class="connect-box-title">Reporte técnico completo</div>
      <p style="font-size:.84rem;color:var(--mid);line-height:1.6;margin:0">Arquitectura del sistema multiagente, algoritmos y validación formal · 27 pp. · PDF</p>
    </div>
    <a href="reporte_tecnico_sistema_multiagente.pdf" download
       style="display:inline-flex;align-items:center;gap:.5rem;background:var(--red);color:#fff;
              text-decoration:none;padding:.7rem 1.4rem;border-radius:8px;font-size:.85rem;
              font-weight:700;white-space:nowrap;transition:opacity .15s"
       onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      ↓ Descargar PDF
    </a>
  </div>

  <div class="connect-box animate-in" style="margin-top:1rem">
    <div class="connect-box-title">Ver los resultados del ciclo 26-P</div>
    <div class="connect-links">
      <a class="connect-link" href="#" onclick="switchPart('resultados',document.querySelector('.part-tab'));return false">← Resultados del ciclo 26-P</a>
      <a class="connect-link" href="cat26p_capacidad.html">📊 Análisis de capacidad →</a>
    </div>
  </div>

</div><!-- /panel-tecnico -->

</main>
"""

CAT_JS = """
(function(){
  const D = window.__CAT__;

  /* Tab switcher */
  window.switchPart = function(panelId, clickedTab) {
    document.querySelectorAll('.part-panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.part-tab').forEach(t => t.classList.remove('active'));
    document.getElementById('panel-' + panelId)?.classList.add('active');
    if (clickedTab) clickedTab.classList.add('active');
    document.querySelectorAll('#panel-'+panelId+' .animate-in:not(.visible)').forEach(el => {
      const io = new IntersectionObserver(entries => {
        entries.forEach(e => { if(e.isIntersecting){ e.target.classList.add('visible'); io.unobserve(e.target); }});
      },{threshold:.1});
      io.observe(el);
    });
  };

  renderChart('chart-donut',
    () => [{
      type:'pie', hole:.52,
      labels:['Alumnas','Alumnos'],
      values:[D.genF, D.genM],
      marker:{colors:['#C82D23','#1565C0'], line:{color:'#fff',width:2}},
      textinfo:'label+percent',
      hovertemplate:'<b>%{label}</b><br>%{value} estudiantes (%{percent})<extra></extra>',
      textfont:{...FONT,size:11}, direction:'clockwise',
      pull:[0.04, 0],
    }],
    () => ({
      margin:{l:8,r:8,t:8,b:8}, showlegend:false,
      annotations:[{text:'<b>'+D.total.toLocaleString('es-MX')+'</b><br><span style="font-size:10px;color:#888">total</span>',
        showarrow:false, font:{size:16,color:'#1A1A1A',family:'Helvetica Neue,Arial'}}]
    })
  );

  renderChart('chart-ueas',
    () => [{
      type:'bar', orientation:'h',
      x:D.gruposLlenados, y:D.ueas,
      marker:{color:'#C82D23', opacity:.87, line:{color:'#9E1E18',width:.5}},
      hovertemplate:'<b>%{y}</b><br>%{x} grupos al 100 % de cupo<extra></extra>',
      text:D.gruposLlenados, textposition:'outside', textfont:{...FONT,size:12}, cliponaxis:false,
    }],
    () => ({
      margin:{l:10,r:36,t:8,b:40},
      xaxis:{title:'Grupos al 100 % de cupo',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',range:[0,22]},
      yaxis:{tickfont:{...FONT,size:11},automargin:true},
    })
  );

  renderChart('chart-byplan',
    () => [{
      type:'bar', orientation:'h',
      x:D.planTotal, y:D.planNames,
      marker:{color:D.planColors, opacity:.88, line:{color:'rgba(0,0,0,.1)',width:.5}},
      hovertemplate:'<b>%{y}</b><br>%{x} estudiantes<extra></extra>',
      text:D.planTotal, textposition:'outside', textfont:{...FONT,size:11}, cliponaxis:false,
    }],
    () => ({
      margin:{l:10,r:40,t:8,b:40},
      xaxis:{title:'Estudiantes',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',range:[0,195]},
      yaxis:{tickfont:{...FONT,size:11},automargin:true},
    })
  );

  renderChart('chart-stacked',
    () => [
      {type:'bar',orientation:'h',name:'Alumnas',
       x:D.planF, y:D.planNames,
       marker:{color:'#C82D23',opacity:.87},
       hovertemplate:'<b>%{y}</b><br>Alumnas: %{x}<extra></extra>'},
      {type:'bar',orientation:'h',name:'Alumnos',
       x:D.planM, y:D.planNames,
       marker:{color:'#1565C0',opacity:.8},
       hovertemplate:'<b>%{y}</b><br>Alumnos: %{x}<extra></extra>'},
    ],
    () => ({
      barmode:'stack',
      margin:{l:10,r:10,t:8,b:40},
      xaxis:{title:'Estudiantes',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF'},
      yaxis:{tickfont:{...FONT,size:11},automargin:true},
      legend:{font:{size:10},orientation:'h',xanchor:'center',x:.5,y:1.06},
    })
  );
})();
"""

cat_html = page_shell(
    title='CAT 26P — Asignación de grupos',
    active='cat26p.html',
    content=CAT_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="#cat-cob"><span class="bottom-nav-icon">🍩</span>Cobertura</a>
      <a class="bottom-nav-item" href="#cat-ueas"><span class="bottom-nav-icon">📊</span>Grupos</a>
      <a class="bottom-nav-item" href="#cat-planes"><span class="bottom-nav-icon">🎓</span>Planes</a>
    """,
    data_vars=f'window.__CAT__ = {DATA_CAT};',
    extra_js=CAT_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# METANALISIS.HTML  — índice (landing)
# ══════════════════════════════════════════════════════════════════════════════
META_CONTENT = f"""
<style>{INDEX_CSS}
  .meta-nodes {{ display:grid; grid-template-columns:1fr 1fr; gap:.7rem }}
  .meta-node {{ background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.12);
               border-radius:12px; padding:1rem 1.1rem; cursor:pointer;
               transition:background var(--ease),border-color var(--ease);
               text-decoration:none; display:block }}
  .meta-node:hover {{ background:rgba(255,255,255,.09); border-color:rgba(200,45,35,.5) }}
  .meta-node-tag {{ font-size:.6rem; font-weight:700; letter-spacing:.12em;
                   text-transform:uppercase; color:var(--red); margin-bottom:.3rem }}
  .meta-node-title {{ font-size:.95rem; font-weight:700; color:#fff; margin-bottom:.25rem }}
  .meta-node-desc {{ font-size:.75rem; color:rgba(255,255,255,.4); line-height:1.45 }}
  .meta-node-sub {{ margin-top:.55rem; padding-top:.5rem;
                   border-top:1px solid rgba(255,255,255,.1);
                   font-size:.72rem; color:rgba(200,45,35,.8); font-weight:600 }}
  .meta-node-wide {{ grid-column:1/-1 }}
  @media(max-width:960px) {{ .meta-nodes {{ grid-template-columns:1fr }} }}
</style>
<main>
  <div class="landing-hero">
    <div class="landing-hero-inner">
      <div>
        <p class="landing-tag animate-in">UAM Azcapotzalco · División de Ciencias Básicas e Ingeniería</p>
        <h1 class="landing-h1 animate-in delay-1">
          Análisis de
          <strong>eficiencia terminal</strong>
        </h1>
        <p class="landing-sub animate-in delay-2">
          Cinco infografías interactivas que diagnostican el rezago en los diez programas
          del Plan 2020: causas estructurales, palancas de mejora, índices de prelación
          y complejidad sistémica — con datos de 2018 a 2025.
        </p>
        <div class="landing-bar animate-in delay-2"></div>
        <div class="landing-kpis animate-in delay-3">
          <div class="lkpi">
            <div class="lkpi-val">10</div>
            <div class="lkpi-lbl">Licenciaturas</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="6562">0</div>
            <div class="lkpi-lbl">Estudiantes activos</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="499">0</div>
            <div class="lkpi-lbl">Seriaciones</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="1859">0</div>
            <div class="lkpi-lbl">UEAs analizadas</div>
          </div>
        </div>
      </div>
      <div class="meta-nodes animate-in delay-2">
        <a class="meta-node" href="meta_1.html">
          <div class="meta-node-tag">Infografía 1</div>
          <div class="meta-node-title">El diagnóstico</div>
          <div class="meta-node-desc">Los 10 programas tardan el doble del plazo nominal en titular su cohorte.</div>
        </a>
        <a class="meta-node" href="meta_2.html">
          <div class="meta-node-tag">Infografía 2</div>
          <div class="meta-node-title">Las causas</div>
          <div class="meta-node-desc">Dos mecanismos estructurales imponen cotas mínimas de permanencia.</div>
        </a>
        <a class="meta-node" href="meta_3.html">
          <div class="meta-node-tag">Infografía 3</div>
          <div class="meta-node-title">La palanca</div>
          <div class="meta-node-desc">El 26.5 % de seriaciones concentra el 80 % del beneficio alcanzable.</div>
        </a>
        <a class="meta-node" href="meta_4.html">
          <div class="meta-node-tag">Infografía 4</div>
          <div class="meta-node-title">El índice P</div>
          <div class="meta-node-desc">Sesgo de diseño que retrasa el egreso hasta 4.7 trimestres.</div>
          <div class="meta-node-sub">↳ Respuesta divisional: índices P y h →</div>
        </a>
        <a class="meta-node meta-node-wide" href="meta_5.html">
          <div class="meta-node-tag">Infografía 5</div>
          <div class="meta-node-title">Complejidad multidimensional</div>
          <div class="meta-node-desc">El índice C mapea la disfunción en cinco dimensiones y orienta la intervención.</div>
        </a>
      </div>
    </div>
    <div class="scroll-hint">▾ desplázate para explorar</div>
  </div>
  <div style="padding:3rem 2rem;max-width:1440px;margin:0 auto" id="content">
    <div class="connect-box animate-in">
      <div class="connect-box-title">¿Cómo se conecta con los demás proyectos?</div>
      <p style="font-size:.84rem;color:var(--mid);line-height:1.6">El análisis de eficiencia terminal documenta el problema; el proyecto de Modificaciones 2026 es, en parte, la respuesta. Evalúa con inteligencia artificial las propuestas de reforma curricular que aún no se han implementado — las mismas seriaciones y UEAs que aquí se identifican como palancas de mejora son las que el análisis de modificaciones examina como candidatas a reestructuración.</p>
      <div class="connect-links">
        <a class="connect-link" href="modificaciones.html">Ver Modificaciones 2026 →</a>
        <a class="connect-link" href="cat26p.html">Ver CAT 26P →</a>
      </div>
    </div>
    <div class="connect-box animate-in" style="margin-top:1rem;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem">
      <div>
        <div class="connect-box-title">Reporte técnico completo</div>
        <p style="font-size:.84rem;color:var(--mid);line-height:1.6;margin:0">Documento integral con metodología, modelos y resultados completos · 94 pp. · PDF</p>
      </div>
      <a href="eficiencia_terminal.pdf" download
         style="display:inline-flex;align-items:center;gap:.5rem;background:var(--red);color:#fff;
                text-decoration:none;padding:.7rem 1.4rem;border-radius:8px;font-size:.85rem;
                font-weight:700;white-space:nowrap;transition:opacity .15s"
         onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
        ↓ Descargar PDF
      </a>
    </div>
  </div>
</main>
"""

META_JS = ""

meta_html = page_shell(
    title='Análisis de Eficiencia Terminal — DCBI',
    active='metanalisis.html',
    content=META_CONTENT,
    data_vars='',
    extra_js=META_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# META_1.HTML  — Infografía 1: El diagnóstico
# ══════════════════════════════════════════════════════════════════════════════
META1_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Metaanálisis · Infografía 1 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>El diagnóstico</strong></h1>
      <p class="hero-sub">Los diez programas de ingeniería tardan el doble del plazo establecido en titular a su cohorte activa.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">La División de Ciencias Básicas e Ingeniería (CBI) de la UAM-Azcapotzalco cuenta con <strong>6,562 estudiantes activos</strong> en sus diez programas de ingeniería del Plan 2020. De ese total, <strong>2,384 se encuentran en riesgo de baja definitiva</strong> —el 36.3%— por acumular cinco evaluaciones no acreditadas en la misma unidad de enseñanza-aprendizaje (UEA).</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">El tiempo medio al egreso es de <strong>20.6 trimestres</strong>, frente al plazo nominal de 12 establecido en los planes de estudio: un exceso de 8.6 trimestres por egresado. Apenas el <strong>1.52%</strong> de los estudiantes titula dentro del plazo curricular (intervalo de confianza del 95%: [0.000%, 0.018%]).</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">La causa no reside en el perfil del alumnado sino en el diseño curricular. La brecha de <strong>9.7 puntos porcentuales</strong> entre el plan con mayor tasa de egreso (Ingeniería Industrial, 67.2%) y el de menor (Ingeniería Eléctrica, 57.5%) —sobre la misma base de estudiantes— lo confirma.</p>
  </section>

  <section class="section animate-in" id="chart-section-et">
    <div class="section-header">
      <h2>Curvas de eficiencia terminal acumulada</h2>
      <span class="section-count">10 planes · simulación Monte Carlo 2018–2025</span>
    </div>
    <p class="section-desc">Las curvas muestran qué fracción de la cohorte ha egresado al cabo de t trimestres. Ningún plan supera el 70% de eficiencia terminal acumulada al trimestre 30 —más de dos veces el plazo nominal de 12. La zona comprendida entre t=12 (plazo del plan) y t=17 (mediana mínima de permanencia) ilustra la brecha estructural.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> el punto donde cada curva despega (ninguno antes de t=12), la pendiente de subida (lenta en todos los planes) y la altura a la que se estabiliza. La diferencia entre Ambiental (~68%) y Eléctrica (~58%) al trimestre 30 —sobre la misma cohorte de ingreso— señala que el diseño curricular, no el perfil del estudiante, explica la variación.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>ET(t) acumulada por plan · plazo nominal t = 12</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering xl" id="chart-meta1-et">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <section class="section animate-in" id="chart-section-dist">
    <div class="section-header">
      <h2>Tiempo esperado al egreso por plan</h2>
      <span class="section-count">E[T|tit] modelo vs ET oficial UAM-A</span>
    </div>
    <p class="section-desc">El tiempo esperado condicional al egreso, E[T|tit], oscila entre 19 y 22 trimestres según el plan —siempre por encima del plazo nominal. El valor oficial registrado por la UAM-A (barras grises) subestima sistemáticamente el tiempo real porque no incluye a los estudiantes que abandonan sin egresar.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> la distancia entre el diamante rojo (modelo) y la barra gris (oficial) en cada plan; esa brecha cuantifica el sesgo de supervivencia en la estadística institucional. La línea discontinua roja marca t=12 (plazo del plan) y la naranja señala el promedio divisional de 20.6 trimestres.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>E[T|tit] modelo (diamante) vs ET oficial UAM-A (barra)</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta1-dist">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Conclusión</div>
    <p>El rezago no es aleatorio ni marginal: es <strong>universal y estructural</strong>. Los diez programas operan con una brecha de 8.6 trimestres entre el plazo nominal y el tiempo real de egreso. La estadística oficial subestima esta brecha al excluir a los estudiantes que no completan. El diseño del plan —no el perfil del estudiante— es la variable que distingue a los programas entre sí.</p>
  </div>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <div class="connect-links">
      <a class="connect-link" href="metanalisis.html">← Volver al índice</a>
      <a class="connect-link" href="meta_2.html">Infografía 2: Las causas →</a>
    </div>
  </div>

</main>
"""

META1_JS = """
(function(){
  renderChart('chart-meta1-et',
    function() {
      var d = window.__META1__;
      var colorMap = {
        ind: '#2E7D32', met: '#F2AC38', qui: '#5898E8', fis: '#5898E8',
        amb: '#8890A8', civ: '#8890A8', com: '#B878E8', mec: '#8890A8',
        ele: '#888888', elo: '#C82D23'
      };
      var dashMap = {
        ind: 'solid', met: 'dot', qui: 'dash', fis: 'dashdot',
        amb: 'solid', civ: 'dash', com: 'solid', mec: 'dot',
        ele: 'dash',  elo: 'solid'
      };
      var widthMap = { ind: 2.5, elo: 2.5 };
      return d.plansOrder.map(function(plan) {
        var c = d.curves[plan];
        return {
          type: 'scatter', mode: 'lines',
          name: d.planNames[plan],
          x: c.t, y: c.ET,
          line: {
            color: colorMap[plan] || '#8890A8',
            dash:  dashMap[plan]  || 'solid',
            width: widthMap[plan] || 1.5
          },
          hovertemplate: '<b>' + d.planNames[plan] + '</b><br>t = %{x} trim<br>ET = %{y:.1%}<extra></extra>'
        };
      });
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:40, r:10, t:12, b:40},
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Trimestre t', font: FONT},
          range: [11, 31], dtick: 2, gridcolor: '#E4E2DF'
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          title: {text: 'Fracción acumulada egresada', font: FONT},
          tickformat: '.0%', range: [0, 0.75], gridcolor: '#E4E2DF'
        }),
        legend: {
          orientation: 'h', x: 0, y: -0.22,
          font: Object.assign({}, FONT, {size: 9}), itemwidth: 30
        },
        shapes: [
          {type:'line', x0:12, x1:12, y0:0, y1:0.75,
           line:{color:'#C82D23', width:1.2, dash:'dot'}},
          {type:'line', x0:17, x1:17, y0:0, y1:0.75,
           line:{color:'#B06A00', width:1, dash:'dot'}}
        ],
        annotations: [
          {x:12.1, y:0.71, text:'Plan (12 trim)', showarrow:false,
           font:{size:8, color:'#C82D23'}, xanchor:'left'},
          {x:17.1, y:0.71, text:'Mediana min.', showarrow:false,
           font:{size:8, color:'#B06A00'}, xanchor:'left'}
        ]
      });
    }
  );

  renderChart('chart-meta1-dist',
    function() {
      var d = window.__META1__;
      var sorted = d.plansOrder.slice().sort(function(a,b){
        return d.distStats[a].et_tit - d.distStats[b].et_tit;
      });
      var labels  = sorted.map(function(p){ return d.planNames[p]; });
      var et_tit  = sorted.map(function(p){ return d.distStats[p].et_tit; });
      var et_off  = sorted.map(function(p){ return d.distStats[p].et_oficial; });
      return [
        {
          type: 'bar', orientation: 'h',
          name: 'ET oficial UAM-A',
          y: labels, x: et_off,
          marker: {color: '#8890A8', opacity: 0.45},
          hovertemplate: '<b>%{y}</b><br>ET oficial: %{x:.1f} trim<extra></extra>'
        },
        {
          type: 'scatter', mode: 'markers',
          name: 'E[T|tit] modelo',
          y: labels, x: et_tit,
          marker: {color: '#C82D23', size: 8, symbol: 'diamond'},
          hovertemplate: '<b>%{y}</b><br>E[T|tit]: %{x:.1f} trim<extra></extra>'
        }
      ];
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:72, r:20, t:12, b:40},
        barmode: 'overlay',
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Trimestres', font: FONT},
          range: [10, 26], gridcolor: '#E4E2DF'
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          gridcolor: 'rgba(0,0,0,0)'
        }),
        shapes: [
          {type:'line', x0:12, x1:12, y0:-0.5, y1:9.5,
           line:{color:'#C82D23', width:1.2, dash:'dot'}},
          {type:'line', x0:20.6, x1:20.6, y0:-0.5, y1:9.5,
           line:{color:'#B06A00', width:1, dash:'dash'}}
        ],
        annotations: [
          {x:12.1, y:9.4, text:'Plazo plan', showarrow:false,
           font:{size:8, color:'#C82D23'}, xanchor:'left'},
          {x:20.7, y:9.4, text:'Prom. Division', showarrow:false,
           font:{size:8, color:'#B06A00'}, xanchor:'left'}
        ],
        legend: {
          orientation: 'h', x: 0, y: -0.18, font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );
})();
"""

meta1_html = page_shell(
    title='Infografía 1 — El diagnóstico',
    active='metanalisis.html',
    content=META1_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="metanalisis.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="meta_2.html"><span class="bottom-nav-icon">&#8594;</span>Causas</a>
    """,
    data_vars='window.__META1__ = ' + DATA_META1 + ';',
    extra_js=META1_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# META_2.HTML  — Infografía 2: Las causas
# ══════════════════════════════════════════════════════════════════════════════
META2_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Metaanálisis · Infografía 2 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>Las causas</strong></h1>
      <p class="hero-sub">El rezago no es transitorio: dos mecanismos estructurales imponen cotas mínimas de permanencia superiores al plazo del plan en los diez programas.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">El plan de estudios diseña la trayectoria asumiendo una inscripción de <strong>45 créditos por trimestre</strong> con aprobación perfecta. Los datos históricos de la División (2016-I a 2025-O) muestran que la carga real oscila entre <strong>27 y 33 créditos por trimestre</strong>, con una tasa neta de aprobación de 0.66 a 0.74. Solo por ese desajuste de carga, acumular los créditos del plan requiere entre <strong>18 y 21 trimestres</strong>, sin contar todavía el efecto de las seriaciones.</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">El segundo mecanismo es topológico: las cadenas de seriación imponen una cota mínima independiente de la carga. Si la cadena más larga tiene d eslabones, se necesitan como mínimo d+1 trimestres incluso aprobando todo y con carga máxima. Electrónica y Computación tienen la mayor profundidad (d=9): <strong>10 trimestres mínimos</strong> solo por estructura del grafo.</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">La cota de carga domina en los <strong>10 planes</strong>: los estudiantes inscriben entre un 36% y un 40% menos de los créditos previstos en el diseño, y ese déficit no se recupera. Más del <strong>98%</strong> de los estudiantes opera por debajo del ritmo del plan desde el tercer trimestre y permanece así hasta el egreso.</p>
  </section>

  <section class="section animate-in" id="chart-section-cotas">
    <div class="section-header">
      <h2>Desajuste de carga crediticia</h2>
      <span class="section-count">27–33 cr/trim reales · 45 cr/trim asumidos · Déficit del 36–40%</span>
    </div>
    <p class="section-desc">La figura muestra las cotas mínimas de permanencia calculadas para los dos mecanismos en cada plan. La cota de carga (naranja) excede la cota topológica (roja) en los diez programas: es el mecanismo dominante. La línea discontinua marca t=12 (plazo del plan); todas las cotas de carga la superan.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> en qué planes la barra naranja (cota de carga) supera a la roja (cota topológica) y en cuánto exceden ambas la línea de t=12. Verifique que la cota de carga domina en los diez programas sin excepción.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Cotas mínimas de permanencia: carga vs topológica</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta2-cotas">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <section class="section animate-in" id="chart-section-ra">
    <div class="section-header">
      <h2>Déficit de ritmo de avance: universal e irreversible</h2>
      <span class="section-count">98% de estudiantes por debajo del ritmo · Desde el trimestre 3</span>
    </div>
    <p class="section-desc">El gráfico muestra la diferencia entre estrategias (ΔP) en función del trimestre. La zona sombreada delimita la ventana T∈[1, 2.5] donde la estrategia conservadora produce el índice P más alto —la llamada «trampa de acreditación» (ACT). A partir del trimestre 3, la ventaja se revierte de forma permanente.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> la zona sombreada donde la curva conservadora produce ΔP positivo y el cruce a partir del cual queda permanentemente negativo.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>ΔP entre estrategias en función del trimestre t</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta2-ra">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Conclusión</div>
    <p>El déficit de ritmo de avance es <strong>universal e irreversible</strong>: el 98% de los estudiantes opera por debajo del ritmo del plan desde el trimestre 3. La cota de carga —inscripción entre un 36% y un 40% inferior al diseño— domina en los diez programas. Ambos mecanismos actúan en paralelo y requieren atención simultánea; corregir solo uno no elimina el rezago.</p>
  </div>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <div class="connect-links">
      <a class="connect-link" href="meta_1.html">← El diagnóstico</a>
      <a class="connect-link" href="metanalisis.html">Índice</a>
      <a class="connect-link" href="meta_3.html">La palanca →</a>
    </div>
  </div>

</main>
"""

META2_JS = """
(function(){

  renderChart('chart-meta2-cotas',
    function() {
      var d = window.__META2__;
      var sorted = d.plansOrder.slice().sort(function(a,b){
        return d.tCarga[b] - d.tCarga[a];
      });
      var labels = sorted.map(function(p){ return d.planNames[p]; });
      var carga  = sorted.map(function(p){ return d.tCarga[p]; });
      var topo   = sorted.map(function(p){ return d.tTopo[p]; });
      return [
        {
          type: 'bar', orientation: 'h', name: 'T<sub>min</sub> carga',
          y: labels, x: carga,
          marker: {color: '#B06A00'},
          hovertemplate: '<b>%{y}</b><br>Cota carga: %{x} trim<extra></extra>'
        },
        {
          type: 'bar', orientation: 'h', name: 'T<sub>min</sub> topologica',
          y: labels, x: topo,
          marker: {color: '#C82D23', opacity: 0.75},
          hovertemplate: '<b>%{y}</b><br>Cota topologica: %{x} trim<extra></extra>'
        }
      ];
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:76, r:20, t:12, b:40},
        barmode: 'group',
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Trimestres minimos', font: FONT},
          range: [0, 25], gridcolor: '#E4E2DF', dtick: 4
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {gridcolor: 'rgba(0,0,0,0)'}),
        shapes: [
          {type:'line', x0:12, x1:12, y0:-0.5, y1:9.5,
           line:{color:'#C82D23', width:1.2, dash:'dot'}}
        ],
        annotations: [
          {x:12.3, y:9.3, text:'Plazo plan (12)', showarrow:false,
           font:{size:8, color:'#C82D23'}, xanchor:'left'}
        ],
        legend: {
          orientation: 'h', x: 0, y: -0.18, font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );

  renderChart('chart-meta2-ra',
    function() {
      var d = window.__META2__;
      var vp = d.ventajaP;
      var zero = vp.t.map(function(){ return 0; });
      return [
        {
          type: 'scatter', mode: 'lines',
          name: 'delta_P conservadora vs baseline',
          x: vp.t, y: vp.delta_cons,
          line: {color: '#C82D23', width: 2, dash: 'solid'},
          fill: 'tozeroy', fillcolor: 'rgba(200,45,35,0.12)',
          hovertemplate: 't = %{x}<br>delta_P (cons-base) = %{y:.3f}<extra></extra>'
        },
        {
          type: 'scatter', mode: 'lines',
          name: 'delta_P agresiva vs baseline',
          x: vp.t, y: vp.delta_agr,
          line: {color: '#2E7D32', width: 1.5, dash: 'dash'},
          hovertemplate: 't = %{x}<br>delta_P (agr-base) = %{y:.3f}<extra></extra>'
        },
        {
          type: 'scatter', mode: 'lines', showlegend: false,
          x: vp.t, y: zero,
          line: {color: '#646464', width: 1}
        }
      ];
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:46, r:10, t:12, b:40},
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Trimestre t', font: FONT},
          dtick: 3, gridcolor: '#E4E2DF'
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          title: {text: 'delta_P (estrategia - baseline)', font: FONT},
          gridcolor: '#E4E2DF', zeroline: true,
          zerolinecolor: '#646464', zerolinewidth: 1
        }),
        shapes: [
          {type:'rect', x0:1, x1:2.5, y0:-0.18, y1:0.14,
           fillcolor:'rgba(176,106,0,0.07)', line:{width:0}},
          {type:'line', x0:2.5, x1:2.5, y0:-0.18, y1:0.14,
           line:{color:'#B06A00', width:1, dash:'dot'}}
        ],
        annotations: [
          {x:1.75, y:0.12, text:'Trampa ACT', showarrow:false,
           font:{size:8, color:'#B06A00'}}
        ],
        legend: {
          orientation: 'h', x: 0, y: -0.20, font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );
})();
"""

meta2_html = page_shell(
    title='Infografía 2 — Las causas',
    active='metanalisis.html',
    content=META2_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="meta_1.html"><span class="bottom-nav-icon">&#8592;</span>Diagnóstico</a>
      <a class="bottom-nav-item" href="metanalisis.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="meta_3.html"><span class="bottom-nav-icon">&#8594;</span>Palanca</a>
    """,
    data_vars='window.__META2__ = ' + DATA_META2 + ';',
    extra_js=META2_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# META_3.HTML  — Infografía 3: La palanca
# ══════════════════════════════════════════════════════════════════════════════
META3_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Metaanálisis · Infografía 3 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>La palanca</strong></h1>
      <p class="hero-sub">El 26.5% de las seriaciones concentra el 80% del beneficio alcanzable: la reforma puede ser quirúrgica, no curricular.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">De las 499 seriaciones activas en los diez planes de ingeniería, <strong>132 concentran el 80% del beneficio total</strong> alcanzable si se eliminaran todas. Identificar y revisar esas 132 relaciones de prerrequisito —el 26.5% del total— permitiría obtener la mayor parte de la ganancia en eficiencia terminal con una intervención acotada.</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">La curva de Lorenz del impacto muestra que la distribución del beneficio entre seriaciones es altamente concentrada: las primeras 20 seriaciones (4% del total) acumulan cerca del 35% del beneficio total. Esta concentración es estructural —no depende del plan específico— y se reproduce en los diez programas.</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">Las curvas de dosis-respuesta confirman el mismo patrón a nivel de plan: la ganancia marginal decrece rápidamente. En Electrónica, las primeras 5 seriaciones eliminadas producen más de la mitad de la ganancia máxima simulada. En Metalurgia, el plan con menor número de seriaciones, el efecto se agota antes del umbral de N=15.</p>
  </section>

  <section class="section animate-in" id="chart-section-lorenz">
    <div class="section-header">
      <h2>Curva de Lorenz del impacto de seriaciones</h2>
      <span class="section-count">499 seriaciones · N80 = 132 (26.5%)</span>
    </div>
    <p class="section-desc">La curva de Lorenz muestra qué fracción del beneficio total acumulado se obtiene al eliminar las N seriaciones más impactantes (ordenadas de mayor a menor ganancia en eficiencia terminal). La diagonal representaría una distribución uniforme del impacto: cuanto más abombada la curva, más concentrado el beneficio.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> el punto N80 —marcado con línea discontinua— donde la curva alcanza 0.80 en el eje vertical. Ese punto indica cuántas seriaciones bastan para obtener el 80% del beneficio. La distancia entre la curva y la diagonal mide el grado de concentración.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Lorenz del impacto: fracción del beneficio vs número de seriaciones eliminadas</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta3-lorenz">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <section class="section animate-in" id="chart-section-remocion">
    <div class="section-header">
      <h2>Curvas dosis-respuesta por plan</h2>
      <span class="section-count">Ganancia en ET al eliminar N seriaciones · intervalo de confianza del 95%</span>
    </div>
    <p class="section-desc">Cada curva muestra la ganancia esperada en eficiencia terminal (en trimestres) al eliminar las N seriaciones más críticas del plan correspondiente. Las bandas sombreadas son los intervalos de confianza del 95% de la simulación Monte Carlo. La pendiente inicial mide la palanca del primer prerrequisito crítico.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> qué plan tiene la curva más empinada al inicio (mayor palanca de la primera seriación crítica), dónde se aplanan las curvas (rendimiento marginal decreciente) y qué planes muestran mayor incertidumbre (bandas más anchas).</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Ganancia en ET al eliminar las N seriaciones más críticas · por plan</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering lg" id="chart-meta3-remocion">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Conclusión</div>
    <p>La reforma curricular no requiere revisión total: <strong>132 de 499 seriaciones</strong> bastan para obtener el 80% del beneficio simulado. La intervención puede ser quirúrgica y graduada. Las curvas dosis-respuesta por plan permiten priorizar qué prerrequisitos revisar primero en cada licenciatura según su relación beneficio-costo de implementación.</p>
  </div>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <div class="connect-links">
      <a class="connect-link" href="meta_2.html">← Las causas</a>
      <a class="connect-link" href="metanalisis.html">Índice</a>
      <a class="connect-link" href="meta_4.html">El índice P →</a>
    </div>
  </div>

</main>
"""

META3_JS = """
(function(){

  renderChart('chart-meta3-lorenz',
    function() {
      var d = window.__META3__;
      var lor = d.lorenz;
      var diag = [0, lor.totalSer];
      return [
        {
          type: 'scatter', mode: 'lines',
          name: 'Curva de Lorenz del impacto',
          x: lor.n, y: lor.fracCum,
          line: {color: '#C82D23', width: 2},
          fill: 'tozeroy', fillcolor: 'rgba(200,45,35,0.08)',
          hovertemplate: 'N = %{x}<br>Fraccion acumulada = %{y:.1%}<extra></extra>'
        },
        {
          type: 'scatter', mode: 'lines',
          name: 'Distribucion uniforme (diagonal)',
          x: [0, lor.totalSer], y: [0, 1],
          line: {color: '#8890A8', width: 1, dash: 'dot'},
          showlegend: true
        }
      ];
    },
    function() {
      var d = window.__META3__;
      var lor = d.lorenz;
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:46, r:20, t:12, b:40},
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Numero de seriaciones eliminadas (N)', font: FONT},
          range: [0, lor.totalSer], gridcolor: '#E4E2DF'
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          title: {text: 'Fraccion del beneficio acumulado', font: FONT},
          tickformat: '.0%', range: [0, 1.05], gridcolor: '#E4E2DF'
        }),
        shapes: [
          {type:'line', x0:lor.N80, x1:lor.N80, y0:0, y1:0.80,
           line:{color:'#B06A00', width:1.2, dash:'dot'}},
          {type:'line', x0:0, x1:lor.N80, y0:0.80, y1:0.80,
           line:{color:'#B06A00', width:1.2, dash:'dot'}}
        ],
        annotations: [
          {x:lor.N80, y:0.83,
           text:'N80 = ' + lor.N80 + ' (' + (lor.N80/lor.totalSer*100).toFixed(1) + '%)',
           showarrow:false, font:{size:9, color:'#B06A00'}, xanchor:'center'}
        ],
        legend: {
          orientation: 'h', x: 0, y: -0.18, font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );

  renderChart('chart-meta3-remocion',
    function() {
      var d = window.__META3__;
      var planColors = {
        amb:'#2E7D32', civ:'#BF360C', com:'#1565C0', ele:'#E65100',
        elo:'#C82D23', fis:'#00838F', ind:'#C62828', mec:'#37474F',
        met:'#78909C', qui:'#558B2F'
      };
      var bandColors = {
        elo:'rgba(200,45,35,0.12)', qui:'rgba(176,106,0,0.12)',
        met:'rgba(21,101,192,0.12)', amb:'rgba(88,152,232,0.12)',
        ele:'rgba(158,30,24,0.12)',  mec:'rgba(46,125,50,0.12)',
        fis:'rgba(184,120,232,0.12)',ind:'rgba(76,203,114,0.12)',
        civ:'rgba(136,144,168,0.12)',com:'rgba(100,100,100,0.12)'
      };
      var traces = [];
      d.plansOrder.forEach(function(plan) {
        var rd = d.remocion[plan];
        var col = planColors[plan] || '#888';
        var bcol = bandColors[plan] || 'rgba(136,144,168,0.10)';
        traces.push({
          type: 'scatter', mode: 'lines',
          name: d.planNames[plan],
          x: rd.N, y: rd.mean,
          line: {color: col, width: 2},
          hovertemplate: '<b>' + d.planNames[plan] + '</b><br>N=%{x}<br>Ganancia: +%{y:.4f} trim<extra></extra>'
        });
        traces.push({
          type: 'scatter', mode: 'lines',
          showlegend: false,
          x: rd.N.concat(rd.N.slice().reverse()),
          y: rd.hi.concat(rd.lo.slice().reverse()),
          fill: 'toself', fillcolor: bcol,
          line: {width: 0},
          hoverinfo: 'skip'
        });
      });
      return traces;
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:52, r:12, t:12, b:40},
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Seriaciones eliminadas (N)', font: FONT},
          gridcolor: '#E4E2DF', zeroline: false
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          title: {text: 'Ganancia en ET (trimestres)', font: FONT},
          gridcolor: '#E4E2DF'
        }),
        legend: {
          orientation: 'h', x: 0, y: -0.18,
          font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );
})();
"""

meta3_html = page_shell(
    title='Infografía 3 — La palanca',
    active='metanalisis.html',
    content=META3_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="meta_2.html"><span class="bottom-nav-icon">&#8592;</span>Causas</a>
      <a class="bottom-nav-item" href="metanalisis.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="meta_4.html"><span class="bottom-nav-icon">&#8594;</span>Índice P</a>
    """,
    data_vars='window.__META3__ = ' + DATA_META3 + ';',
    extra_js=META3_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# META_4.HTML  — Infografía 4: El índice P
# ══════════════════════════════════════════════════════════════════════════════
META4_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Metaanálisis · Infografía 4 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>El índice P</strong></h1>
      <p class="hero-sub">Inscribir menos créditos eleva artificialmente el índice P: una trampa de diseño que retrasa el egreso hasta 4.7 trimestres y reduce la tasa de titulación tres veces.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">El índice de avance P mide la fracción de créditos del plan acumulados por el estudiante. Su construcción crea una trampa: inscribir menos créditos en un trimestre —aunque se aprueben todos— reduce la carga actual sin mejorar el numerador acumulado, lo que puede elevar artificialmente P a corto plazo. Un estudiante con estrategia conservadora (24 créditos por trimestre) alcanza un P más alto que uno con carga estándar durante los dos primeros trimestres.</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">A partir del trimestre 3, la ventaja se revierte: la estrategia conservadora produce un P permanentemente inferior porque el estudiante acumula créditos más lentamente. El costo es <strong>4.7 trimestres adicionales</strong> al tiempo de egreso en promedio, y una reducción de la tasa de titulación a los 18 trimestres de 19.7% (baseline) a 3.1% (conservadora).</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">La estrategia agresiva (38 créditos por trimestre) acorta el egreso pero exige una tasa de aprobación sostenida que los datos históricos no respaldan para la mayoría de los estudiantes. El diseño del índice P no distingue entre trayectorias con distinto ritmo: <strong>el indicador premia la inscripción, no el aprendizaje</strong>.</p>
  </section>

  <section class="section animate-in" id="chart-section-p">
    <div class="section-header">
      <h2>Evolución del índice P por estrategia de inscripción</h2>
      <span class="section-count">Promedio de los 10 planes · trimestres 1–30</span>
    </div>
    <p class="section-desc">Las tres curvas muestran cómo evoluciona el índice P promedio según la estrategia de carga: baseline (histórico, ~30 cr/trim), conservadora (24 cr/trim) y agresiva (38 cr/trim). La estrategia conservadora arranca con mayor P pero se cruza con la baseline en el trimestre 2-3 y queda permanentemente rezagada.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> el cruce entre la curva conservadora y la baseline alrededor del trimestre 3, y la divergencia posterior. La distancia vertical entre curvas al trimestre 12 mide el costo de la trampa de inscripción reducida.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Índice P(t) promedio según estrategia de inscripción</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta4-p">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <section class="section animate-in" id="chart-section-comp-estrategias">
    <div class="section-header">
      <h2>Comparativa de resultados por estrategia</h2>
      <span class="section-count">ET(18 trim) · ET(24 trim) · ΔP en t=2</span>
    </div>
    <p class="section-desc">La tabla comparativa muestra los resultados clave de cada estrategia: eficiencia terminal a los 18 y 24 trimestres, y la diferencia de P en el trimestre 2 respecto al baseline. La estrategia conservadora tiene el mayor ΔP inicial pero el peor resultado de egreso a largo plazo.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> la inversión del ranking entre el trimestre 2 (donde conservadora lidera) y los trimestres 18–24 (donde lidera agresiva). Esa inversión es la definición operacional de la trampa.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Resultados comparativos de las tres estrategias de inscripción</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta4-comp">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Conclusión</div>
    <p>El índice P contiene un sesgo de diseño que incentiva la inscripción reducida a corto plazo y penaliza al estudiante a largo plazo. Corregir este sesgo requiere reformular el indicador para que refleje el avance real —créditos aprobados acumulados sobre créditos del plan— sin que la carga corriente lo distorsione. Mientras el sesgo persista, los sistemas de alerta temprana basados en P pueden identificar erróneamente como «avanzados» a estudiantes con trayectorias de egreso más lentas.</p>
  </div>

  <hr class="divider">

  <div class="recuadro animate-in">
    <div class="recuadro-title">Respuesta divisional: ¿puede el criterio de prelación mejorar?</div>
    <p>El índice de desempeño <em>P</em> vigente castiga estructuralmente a los estudiantes bloqueados por seriaciones&#8202;—&#8202;precisamente cuando más necesitan acceso preferente a los grupos del cuello de botella. La subpágina siguiente analiza el diseño del índice alternativo&#8202;&#8202;<em>h</em>, compara sus incentivos con los de <em>P</em> y muestra por qué ningún índice puede eliminar la penalización topológica mientras las seriaciones permanezcan en el plan.</p>
    <div class="connect-links" style="margin-top:.75rem">
      <a class="connect-link" href="meta_6.html">Índices P y h — análisis comparativo →</a>
    </div>
  </div>

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <div class="connect-links">
      <a class="connect-link" href="meta_3.html">← La palanca</a>
      <a class="connect-link" href="metanalisis.html">Índice</a>
      <a class="connect-link" href="meta_5.html">Complejidad →</a>
    </div>
  </div>

</main>
"""

META4_JS = """
(function(){

  renderChart('chart-meta4-p',
    function() {
      var d = window.__META4__;
      var ps = d.pByStrat;
      return [
        {
          type: 'scatter', mode: 'lines',
          name: 'Baseline (~30 cr/trim)',
          x: ps.t, y: ps.baseline,
          line: {color: '#1A1A1A', width: 2},
          hovertemplate: 't = %{x}<br>P baseline = %{y:.3f}<extra></extra>'
        },
        {
          type: 'scatter', mode: 'lines',
          name: 'Conservadora (24 cr/trim)',
          x: ps.t, y: ps.conservadora,
          line: {color: '#B06A00', width: 2, dash: 'dash'},
          hovertemplate: 't = %{x}<br>P conservadora = %{y:.3f}<extra></extra>'
        },
        {
          type: 'scatter', mode: 'lines',
          name: 'Agresiva (38 cr/trim)',
          x: ps.t, y: ps.agresiva,
          line: {color: '#2E7D32', width: 2, dash: 'dot'},
          hovertemplate: 't = %{x}<br>P agresiva = %{y:.3f}<extra></extra>'
        }
      ];
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:46, r:10, t:12, b:40},
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Trimestre t', font: FONT},
          dtick: 3, gridcolor: '#E4E2DF'
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          title: {text: 'Indice P promedio', font: FONT},
          gridcolor: '#E4E2DF', range: [0, 1.05]
        }),
        shapes: [
          {type:'rect', x0:1, x1:2.5, y0:0, y1:1.05,
           fillcolor:'rgba(176,106,0,0.06)', line:{width:0}},
          {type:'line', x0:2.5, x1:2.5, y0:0, y1:1.05,
           line:{color:'#B06A00', width:1, dash:'dot'}}
        ],
        annotations: [
          {x:1.75, y:1.02, text:'Trampa P', showarrow:false,
           font:{size:8, color:'#B06A00'}}
        ],
        legend: {
          orientation: 'h', x: 0, y: -0.18, font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );

  renderChart('chart-meta4-comp',
    function() {
      var d = window.__META4__;
      var er = d.estrategiasResumen;
      var labels = ['Conservadora (24 cr/trim)', 'Baseline (carga historica)', 'Agresiva (38 cr/trim)'];
      var et_vals = [er.conservadora.ET, er.baseline.ET, er.agresiva.ET];
      var grad_vals = [er.conservadora.grad, er.baseline.grad, er.agresiva.grad];
      var colors = ['#B06A00', '#1A1A1A', '#2E7D32'];
      return [
        {
          type: 'bar', orientation: 'h',
          name: 'E[T|egreso] (trimestres)',
          y: labels, x: et_vals,
          marker: {color: colors, opacity: 0.85},
          xaxis: 'x',
          hovertemplate: '<b>%{y}</b><br>E[T]: %{x:.1f} trim<extra></extra>'
        },
        {
          type: 'bar', orientation: 'h',
          name: 'Tasa de graduacion ET(18 trim) %',
          y: labels, x: grad_vals,
          marker: {color: colors, opacity: 0.45},
          xaxis: 'x2',
          hovertemplate: '<b>%{y}</b><br>Tasa grad: %{x:.1f}%<extra></extra>'
        }
      ];
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:140, r:20, t:20, b:40},
        barmode: 'group',
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'E[T|egreso] trimestres', font: FONT},
          domain: [0, 0.48], gridcolor: '#E4E2DF'
        }),
        xaxis2: {
          title: {text: 'Tasa graduacion ET(18 trim) %', font: FONT},
          domain: [0.52, 1], gridcolor: '#E4E2DF',
          tickfont: FONT, zeroline: false,
          anchor: 'free', overlaying: 'x', side: 'bottom', position: 0
        },
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          gridcolor: 'rgba(0,0,0,0)'
        }),
        legend: {
          orientation: 'h', x: 0, y: -0.20, font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );
})();
"""

meta4_html = page_shell(
    title='Infografía 4 — El índice P',
    active='metanalisis.html',
    content=META4_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="meta_3.html"><span class="bottom-nav-icon">&#8592;</span>Palanca</a>
      <a class="bottom-nav-item" href="metanalisis.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="meta_5.html"><span class="bottom-nav-icon">&#8594;</span>Complejidad</a>
    """,
    data_vars='window.__META4__ = ' + DATA_META4 + ';',
    extra_js=META4_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# META_5.HTML  — Infografía 5: Complejidad multidimensional
# ══════════════════════════════════════════════════════════════════════════════
META5_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Metaanálisis · Infografía 5 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>Complejidad multidimensional</strong></h1>
      <p class="hero-sub">El índice C mapea la naturaleza de la disfunción en cinco dimensiones y orienta la intervención hacia el frente prioritario de cada plan.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">El índice de complejidad C no mide la dificultad intelectual de los planes de estudio: mide la <strong>fragilidad estructural del camino hacia el egreso</strong>. Combina cinco dimensiones independientes: D1 topológica (profundidad y densidad de seriaciones), D2 probabilística (variabilidad de trayectorias), D3 evaluativa (sesgo del índice P y trampas de acreditación), D4 docente (varianza de desempeño entre grupos) y D5 de recuperación (riesgo latente de bloqueo).</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">Tres planes alcanzan la clase Alta (C &gt; 0.50): Computación, Química y Mecánica. Cuatro planes se ubican en Moderada y tres en Baja. La dimensión dominante varía por plan: Computación está dominada por D1 (topológica), mientras Química lo está por D3 (evaluativa). Esta heterogeneidad implica que la intervención debe ser específica por plan, no genérica.</p>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:.9rem">El índice C correlaciona negativamente con la eficiencia terminal: los planes de complejidad Alta tienen en promedio 5.8 puntos porcentuales menos de egresados a los 24 trimestres que los de complejidad Baja. Esta brecha es causalmente atribuible al diseño curricular, no al perfil de ingreso, dado que los planes comparten la misma base de admisión.</p>
  </section>

  <section class="section animate-in" id="chart-section-radar">
    <div class="section-header">
      <h2>Perfil de complejidad por dimensión</h2>
      <span class="section-count">5 dimensiones · polígonos radar normalizados</span>
    </div>
    <p class="section-desc">Cada polígono muestra el perfil de un plan en las cinco dimensiones del índice C (normalizadas a [0,1]). La forma del polígono revela cuál dimensión domina la complejidad: un perfil estrellado en D1 indica problemas topológicos (seriaciones profundas); en D3, problemas evaluativos (sesgo del índice P).</p>
    <div class="read-callout"><strong>Qué buscar:</strong> qué vértice del radar sobresale más en cada plan (dimensión dominante), y si los planes de la misma clase de complejidad comparten un patrón de perfil o difieren en su estructura interna.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Radar de dimensiones del índice C por plan</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering lg" id="chart-meta5-radar">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <section class="section animate-in" id="chart-section-ranking">
    <div class="section-header">
      <h2>Ranking del índice C global</h2>
      <span class="section-count">10 planes · clases Alta / Moderada / Baja</span>
    </div>
    <p class="section-desc">Las barras muestran el valor del índice C global para cada plan, ordenado de mayor a menor complejidad. El color indica la clase: rojo para Alta, naranja para Moderada, verde para Baja. La línea discontinua marca el umbral C=0.50 que separa Alta de Moderada.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> cuántos planes superan el umbral C=0.50 (clase Alta), la magnitud de la brecha entre el plan más complejo y el menos complejo, y la dimensión dominante indicada en la etiqueta de cada barra.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Índice C global por plan · clasificacion y dimension dominante</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-meta5-ranking">
        <div class="chart-skeleton"></div>
      </div>
    </div>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Conclusión</div>
    <p>El índice C revela que la complejidad estructural de los planes de ingeniería <strong>no es homogénea ni unidimensional</strong>. Cada plan tiene un frente prioritario distinto: reducir seriaciones en Computación, reformar el indicador P en Química, reducir la varianza docente en Mecánica. Un programa de mejora basado únicamente en el índice global perdería esa especificidad. El valor del índice C está en su descomposición, no en el escalar.</p>
  </div>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <div class="connect-links">
      <a class="connect-link" href="meta_4.html">← El índice P</a>
      <a class="connect-link" href="metanalisis.html">Índice</a>
    </div>
  </div>

</main>
"""

META5_JS = """
(function(){

  renderChart('chart-meta5-radar',
    function() {
      var d = window.__META5__;
      var planColors = {
        amb:'#2E7D32', civ:'#BF360C', com:'#1565C0', ele:'#E65100',
        elo:'#C82D23', fis:'#00838F', ind:'#C62828', mec:'#37474F',
        met:'#78909C', qui:'#558B2F'
      };
      return d.plansOrder.map(function(plan) {
        var rd = d.radarData[plan];
        var col = planColors[plan] || '#888';
        var vals = [rd.D1, rd.D2, rd.D3, rd.D4, rd.D5, rd.D1];
        var theta = d.dimLabels.concat([d.dimLabels[0]]);
        return {
          type: 'scatterpolar', mode: 'lines+markers',
          name: d.planNames[plan],
          r: vals, theta: theta,
          line: {color: col, width: 1.5},
          marker: {color: col, size: 5},
          fill: 'toself', fillcolor: col.replace(')', ',0.07)').replace('rgb', 'rgba'),
          hovertemplate: '<b>' + d.planNames[plan] + '</b><br>%{theta}: %{r:.3f}<extra></extra>'
        };
      });
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:40, r:40, t:30, b:40},
        polar: {
          radialaxis: {
            visible: true, range: [0, 1],
            tickfont: Object.assign({}, FONT, {size: 8}),
            gridcolor: '#E4E2DF'
          },
          angularaxis: {
            tickfont: Object.assign({}, FONT, {size: 9}),
            gridcolor: '#E4E2DF'
          }
        },
        legend: {
          orientation: 'h', x: 0, y: -0.12,
          font: Object.assign({}, FONT, {size: 9})
        }
      });
    }
  );

  renderChart('chart-meta5-ranking',
    function() {
      var d = window.__META5__;
      var claseColors = d.claseColors;
      var ranked = d.rankingOrder.slice().reverse();
      var labels = ranked.map(function(p){ return d.planNames[p]; });
      var vals   = ranked.map(function(p){ return d.globalC[p]; });
      var cols   = ranked.map(function(p){ return claseColors[d.planClase[p]] || '#888'; });
      var texts  = ranked.map(function(p){ return d.globalC[p].toFixed(3); });
      return [
        {
          type: 'bar', orientation: 'h',
          y: labels, x: vals,
          marker: {color: cols, opacity: 0.88},
          text: texts, textposition: 'outside',
          textfont: Object.assign({}, FONT, {size: 10}),
          cliponaxis: false,
          hovertemplate: '<b>%{y}</b><br>C = %{x:.5f}<extra></extra>'
        }
      ];
    },
    function() {
      return Object.assign({}, LAYOUT_BASE, {
        margin: {l:72, r:52, t:12, b:40},
        xaxis: Object.assign({}, LAYOUT_BASE.xaxis, {
          title: {text: 'Indice C global', font: FONT},
          range: [0, 0.85], gridcolor: '#E4E2DF'
        }),
        yaxis: Object.assign({}, LAYOUT_BASE.yaxis, {
          gridcolor: 'rgba(0,0,0,0)'
        }),
        shapes: [
          {type:'line', x0:0.50, x1:0.50, y0:-0.5, y1:9.5,
           line:{color:'#C82D23', width:1.2, dash:'dot'}}
        ],
        annotations: [
          {x:0.51, y:9.4, text:'Alta (C > 0.50)', showarrow:false,
           font:{size:8, color:'#C82D23'}, xanchor:'left'}
        ]
      });
    }
  );
})();
"""

meta5_html = page_shell(
    title='Infografía 5 — Complejidad multidimensional',
    active='metanalisis.html',
    content=META5_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="meta_4.html"><span class="bottom-nav-icon">&#8592;</span>Índice P</a>
      <a class="bottom-nav-item" href="metanalisis.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
    """,
    data_vars='window.__META5__ = ' + DATA_META5 + ';',
    extra_js=META5_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: meta_6.html — Infografía 6: Índices de prelación P y h
# ══════════════════════════════════════════════════════════════════════════════
META6_CONTENT = """
<header class="hero" id="indice-hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Metaanálisis · Infografía 4 · Respuesta divisional · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1>¿<strong>Qué mide el índice P</strong> y qué debería medir?</h1>
      <p class="hero-sub">El <em>índice P</em> vigente recompensa inscribir menos créditos y castiga a quienes más necesitan acceso preferencial. El <em>índice h</em> alternativo elimina ambos efectos mediante un principio de máxima entropía. Ninguno resuelve el bloqueo topológico.</p>
      <div class="hero-bar"></div>
      <div class="stats">
        <div class="stat">
          <span class="stat-val">100%</span>
          <span class="stat-lbl">de estudiantes con RA&nbsp;&lt;&nbsp;1 estructural (todos los planes, todos los trimestres)</span>
        </div>
        <div class="stat">
          <span class="stat-val">+0.111</span>
          <span class="stat-lbl">ventaja ΔP de la estrategia conservadora en T&#8202;=&#8202;2, antes del cruce</span>
        </div>
        <div class="stat">
          <span class="stat-val">&#8722;4.8</span>
          <span class="stat-lbl">trimestres adicionales al egreso por la trampa ACT a largo plazo</span>
        </div>
        <div class="stat">
          <span class="stat-val">Δh&nbsp;&lt;&nbsp;0</span>
          <span class="stat-lbl">el índice h nunca premia el subregistro: inscribir menos siempre reduce h</span>
        </div>
      </div>
    </div>
    <div class="palette">
      <p class="palette-title">Índices y estrategias</p>
      <div class="palette-chips">
        <div class="chip">
          <span class="chip-dot" style="background:#C82D23"></span>
          <span><em>P</em> vigente (Acuerdo 551.4.4.1&#8209;2015)</span>
        </div>
        <div class="chip">
          <span class="chip-dot" style="background:#00838F"></span>
          <span><em>h</em> alternativo (máx. entropía)</span>
        </div>
        <div class="chip">
          <span class="chip-dot" style="background:#2E7D32"></span>
          <span>Estrategia agresiva (38 cr/trim)</span>
        </div>
        <div class="chip">
          <span class="chip-dot" style="background:#B06A00"></span>
          <span>Estrategia conservadora (24 cr/trim)</span>
        </div>
        <div class="chip">
          <span class="chip-dot" style="background:#1565C0"></span>
          <span>h&#8209;óptima (38 cr/trim, orden topológico)</span>
        </div>
      </div>
    </div>
  </div>
</header>

<main class="main">

<!-- ── Sección 1: RA crónico ── -->
<section class="section animate-in" id="indice-p-deficit">
  <div class="section-header">
    <h2>RA crónico&#8202;—&#8202;el calibrador imposible del índice P</h2>
    <span class="section-count">RA &lt; 1 en el 100% de los estudiantes · todos los trimestres · todos los planes</span>
  </div>
  <p class="section-desc">El índice <em>P</em> incorpora un componente llamado <strong>RA</strong> (ritmo de avance), definido como la razón entre los créditos aprobados por el estudiante en el trimestre en curso y un valor de referencia fijo de <strong>38 créditos por trimestre</strong>. El componente RA se trunca en 1 cuando el estudiante alcanza o supera ese umbral. El problema estructural es que ningún estudiante lo alcanza: la carga real promedio en los diez planes de la División de Ciencias Básicas e Ingeniería oscila alrededor de 30 créditos por trimestre, no 38. La brecha no proviene de bajo rendimiento académico sino de las <strong>cadenas de seriación</strong> que bloquean la inscripción&#8202;—&#8202;un estudiante que ha reprobado un nodo topológicamente crítico pierde acceso a todas las unidades de enseñanza y aprendizaje (UEA) que dependen de él, y la oferta de UEA sin prerrequisito pendiente es estructuralmente inferior a 38 créditos en la mayoría de los trimestres.</p>
  <p class="section-desc">La consecuencia es que RA opera, en la práctica, como un componente fijo cercano a 0.79 para la población general, y cerca de cero para los estudiantes bloqueados por seriaciones. El índice P hereda este déficit en su término de mayor peso: la componente RA contribuye con un coeficiente que oscila entre 0.40 y 0.50 según el plan, de modo que el techo real de P para un estudiante no bloqueado es sensiblemente inferior a 1 antes incluso de considerar los demás factores.</p>

  <div class="read-callout">
    <strong>Cómo leer la gráfica:</strong> cada barra muestra la carga efectiva c<sub>s</sub> calibrada para ese plan (créditos promedio aprobados por trimestre en la trayectoria histórica). La línea punteada marca el umbral de referencia de 38 cr/trim que usa el denominador de RA. Todas las barras <strong>rojas</strong> quedan por debajo de esa línea&#8202;—&#8202;lo que convierte el déficit de RA en un resultado estructural, no en una característica de las cohortes. La barra <strong>verde</strong> (Metalurgia, c<sub>s</sub>=41) es la única excepción; su superávit se debe a que su plan de estudios tiene menor densidad de seriaciones bloqueantes. Pase el cursor sobre cada barra para ver el valor exacto de RA = c<sub>s</sub>/38.
  </div>

  <div class="chart-card animate-in delay-1">
    <div class="chart-card-header">
      <h3>Carga efectiva c<sub>s</sub> vs. calibrador de RA (38 cr/trim) por plan</h3>
      <span class="chart-type-tag">Interactivo</span>
    </div>
    <div class="chart-canvas rendering md" id="chart-ra-deficit"><div class="chart-skeleton"></div></div>
  </div>

  <div class="howto-card animate-in delay-1">
    <h3>Cómo interpretar el déficit de RA</h3>
    <div class="howto-step">
      <div class="howto-icon">1</div>
      <div class="howto-text"><strong>Localice las franjas de color uniforme.</strong> Si una columna de trimestre aparece completamente saturada en un plan, el umbral de 38 cr/trim es inalcanzable estructuralmente&#8202;—&#8202;no como excepción sino como norma. Eso no indica que los estudiantes reprobarán más, sino que el calibrador del índice está mal fijado para ese contexto.</div>
    </div>
    <div class="howto-step">
      <div class="howto-icon">2</div>
      <div class="howto-text"><strong>Compare entre planes.</strong> Metalurgia es el único plan donde una fracción de estudiantes alcanza RA&#8202;=&#8202;1 en algunos trimestres. Esto no refleja mayor esfuerzo sino menor densidad de seriaciones: la oferta desbloqueada supera los 38 créditos en esos tramos. La heterogeneidad entre planes revela que el calibrador beneficia estructuralmente a planes con topologías más simples.</div>
    </div>
    <div class="howto-step">
      <div class="howto-icon">3</div>
      <div class="howto-text"><strong>Lea el eje temporal.</strong> El déficit se agudiza en los trimestres intermedios (5&#8209;10), que corresponden al núcleo de las cadenas de seriación. En los trimestres finales, los estudiantes que permanecen en el sistema son los que superaron los cuellos de botella críticos, por lo que el déficit se atenúa&#8202;—&#8202;efecto de supervivencia, no de mejora sistémica.</div>
    </div>
  </div>
</section>

<!-- ── Sección 2: La trampa ACT ── -->
<section class="section animate-in" id="indice-trampa-act">
  <div class="section-header">
    <h2>La trampa ACT&#8202;—&#8202;cuando el índice premia inscribir menos</h2>
    <span class="section-count">ΔP(T&#8202;=&#8202;2)&#8202;=&#8202;+0.111 → cruce en T&#8202;≈&#8202;2.5 → egreso +4.8 trim</span>
  </div>
  <p class="section-desc">La componente <strong>ACT</strong> del índice <em>P</em> mide la razón entre créditos aprobados e inscritos en el trimestre anterior: <em>ACT&#8202;=&#8202;C<sub>A,T&#8722;1</sub>/C<sub>I,T&#8722;1</sub></em>. Su peso es <strong>0.40</strong>, el más alto de todos los componentes del índice. Un estudiante que inscribe 24 créditos y aprueba los 24 obtiene ACT&#8202;=&#8202;1.0; uno que inscribe 38 créditos y aprueba 30 obtiene ACT&#8202;=&#8202;0.79. El índice percibe al primero como más eficiente, aunque el segundo haya aprobado seis créditos más. Este diseño crea un incentivo activo al <strong>subregistro</strong>: inscribir pocos créditos fáciles es la estrategia racional para maximizar P en el corto plazo.</p>
  <p class="section-desc">La trampa se desarrolla en tres fases. En el trimestre 2, la estrategia conservadora genera una ventaja de <strong>ΔP&#8202;=&#8202;+0.111</strong> sobre la línea base. Esta ventaja desaparece alrededor del trimestre 2.5, cuando el menor ritmo de acumulación de créditos comienza a pesar más que el ACT alto. A partir de ese punto la curva conservadora cae por debajo de la línea base y no la recupera: en el trimestre 8, el diferencial se ha invertido a <strong>ΔP&#8202;=&#8202;&#8722;0.070</strong>. A largo plazo, la estrategia conservadora produce <strong>4.8 trimestres adicionales</strong> al tiempo esperado de egreso y reduce la tasa de graduación en <strong>29 puntos porcentuales</strong>.</p>

  <div class="read-callout">
    <strong>Cómo leer la gráfica de incentivos (ΔP y Δh):</strong> las dos curvas muestran, trimestre a trimestre, la diferencia entre la estrategia conservadora y la línea base bajo cada índice. Para el índice <em>P</em> (curva roja), el valor es positivo en T&#8202;=&#8202;1&#8209;2 y cruza el cero en T&#8202;≈&#8202;2.5: esa discontinuidad de signo es la trampa. Para el índice <em>h</em> (curva teal), la diferencia es negativa desde T&#8202;=&#8202;1 y se mantiene en esa dirección en todo el horizonte temporal: conservadora nunca es preferible bajo h. El área sombreada entre las dos curvas cuantifica la magnitud del sesgo inducido por la diferencia de diseño.
  </div>

  <div class="chart-card animate-in delay-1">
    <div class="chart-card-header">
      <h3>Incentivo al subregistro: ΔP vs. Δh (conservadora − baseline)</h3>
      <span class="chart-type-tag">Interactivo</span>
    </div>
    <div class="chart-canvas rendering md" id="chart-delta-incentivos"><div class="chart-skeleton"></div></div>
  </div>

  <div class="read-callout" style="margin-top:1.25rem">
    <strong>Cómo leer la gráfica de dispersión (trampa ACT):</strong> cada punto representa un estudiante simulado en el trimestre 8, posicionado según su ACT trimestral (eje x) y su P total (eje y). Los puntos conservadores se concentran en la esquina de ACT alto / P bajo: demuestran que ACT elevado no es señal de progreso sino de subregistro. Los puntos agresivos muestran la relación inversa&#8202;—&#8202;ACT moderado pero P alto&#8202;—&#8202;que corresponde a trayectorias de egreso más rápidas. La nube de puntos bloqueados (círculos vacíos) se sitúa sistemáticamente debajo de la nube desbloqueada, anticipando el análisis de la penalización topológica.
  </div>

  <div class="chart-card animate-in delay-1">
    <div class="chart-card-header">
      <h3>La trampa ACT: índice P(T=8) vs. tiempo esperado de egreso</h3>
      <span class="chart-type-tag">Interactivo</span>
    </div>
    <div class="chart-canvas rendering md" id="chart-trampa-scatter"><div class="chart-skeleton"></div></div>
  </div>

  <div class="howto-card animate-in delay-1">
    <h3>El mecanismo de la trampa en tres pasos</h3>
    <div class="howto-step">
      <div class="howto-icon">1</div>
      <div class="howto-text"><strong>El incentivo de corto plazo.</strong> En los primeros trimestres, el estudiante observa que inscribir 24 créditos fáciles y aprobarlos todos eleva su P por encima de compañeros con carga mayor. La señal del índice es coherente con lo que el estudiante experimenta: parece avanzar más rápido.</div>
    </div>
    <div class="howto-step">
      <div class="howto-icon">2</div>
      <div class="howto-text"><strong>El cruce invisible.</strong> Alrededor del trimestre 3, el menor ritmo de acumulación de créditos comienza a superar la ventaja del ACT alto. El cruce ocurre gradualmente y el estudiante no lo percibe: sigue con carga reducida porque la señal del índice aún es favorable o neutral. El daño se acumula de forma subperceptible.</div>
    </div>
    <div class="howto-step">
      <div class="howto-icon">3</div>
      <div class="howto-text"><strong>El costo diferido.</strong> A partir del trimestre 5, la estrategia conservadora produce un P inferior y una distancia al egreso cuatro trimestres mayor. El estudiante que siguió la señal del índice durante los primeros trimestres llega tarde a este diagnóstico. La trampa no es accidental: es una consecuencia directa de que el mayor peso del índice recompensa la fracción aprobada/inscrita, no los créditos aprobados en valor absoluto.</div>
    </div>
  </div>
</section>

<!-- ── Sección 3: El índice h ── -->
<section class="section animate-in" id="indice-h">
  <div class="section-header">
    <h2>El índice h&#8202;—&#8202;neutralidad informativa y fin del subregistro</h2>
    <span class="section-count">h = (η<sub>t</sub> + ⟨η⟩<sub>t</sub>) / 2 · sin calificaciones · sin incentivo al subregistro</span>
  </div>
  <p class="section-desc">El <em>índice h</em> se construye a partir del principio de <strong>máxima entropía de Shannon</strong> para la ponderación entre el desempeño instantáneo y el histórico. Define η<sub>t</sub>&#8202;=&#8202;C<sub>A,t&#8722;1</sub>/C<sub>max</sub>, donde <em>C<sub>max</sub>&#8202;=&#8202;45 cr/trim</em> es un máximo fijo e invariante, y ⟨η⟩<sub>t</sub> es la media acumulada de ese cociente hasta el trimestre <em>t</em>. El índice h es la semisuma de ambas cantidades. La diferencia clave respecto al índice <em>P</em> reside en el denominador: C<sub>max</sub>&#8202;=&#8202;45 es un techo fijo, no la carga inscrita en el trimestre en curso. En consecuencia, la <strong>única forma de elevar h es aprobar más créditos en valor absoluto</strong>&#8202;—&#8202;inscribir menos créditos, aunque se aprueben todos, siempre reduce h porque el numerador crece más despacio que lo que el denominador fijo admite.</p>
  <p class="section-desc">Una segunda propiedad relevante es que h no incorpora calificaciones: opera exclusivamente sobre conteos de créditos aprobados. Esto elimina la componente de heterogeneidad docente que contamina al índice <em>P</em>: la variabilidad del coeficiente de variación de eficacia docente entre grupos de la misma UEA es de <strong>CV&#8202;=&#8202;0.21</strong> en promedio, y en el 34% de las UEA de alto riesgo supera 0.30. Un índice sensible a calificaciones introduce un ruido de origen institucional&#8202;—&#8202;qué tan exigente es el profesor asignado&#8202;—&#8202;que el estudiante no puede controlar. El índice h es ciego a ese ruido por diseño.</p>
  <p class="section-desc">La <strong>coherencia ordinal</strong> del índice h es monótona: en todos los trimestres, la estrategia conservadora produce el h más bajo, seguida de la línea base, la h&#8209;óptima (38 cr/trim en orden topológico) y la agresiva. Este orden se mantiene sin inversiones desde el trimestre 1 hasta el final del horizonte simulado, lo que significa que h nunca confunde al estudiante sobre qué estrategia es mejor&#8202;—&#8202;una propiedad que el índice <em>P</em> viola sistemáticamente en los primeros tres trimestres.</p>

  <div class="read-callout">
    <strong>Cómo leer la gráfica de coherencia (h):</strong> el diagrama de dispersión enfrenta, para cada estrategia y trimestre, el valor del índice h (eje y) contra el tiempo esperado de egreso E[T&#8202;|&#8202;tit] (eje x). Una alineación perfectamente monótona&#8202;—&#8202;h más alto implica siempre egreso más rápido&#8202;—&#8202;indicaría coherencia ordinal completa. Compare esta nube con la dispersión análoga del índice <em>P</em>: en el caso de P aparecen inversiones (puntos conservadores por encima de puntos baseline en los trimestres tempranos) que no existen en la nube de h. La ausencia de inversiones en h no es trivial: es la propiedad que convierte al índice en una señal confiable para el estudiante y para el sistema de alerta temprana.
  </div>

  <div class="chart-card animate-in delay-1">
    <div class="chart-card-header">
      <h3>Coherencia del índice h con el egreso — E[T|tit] por estrategia</h3>
      <span class="chart-type-tag">Interactivo</span>
    </div>
    <div class="chart-canvas rendering md" id="chart-h-coherencia"><div class="chart-skeleton"></div></div>
  </div>

  <div class="howto-card animate-in delay-1">
    <h3>Qué diferencia la coherencia de h de la dispersión de P</h3>
    <div class="howto-step">
      <div class="howto-icon">1</div>
      <div class="howto-text"><strong>Busque las inversiones de orden.</strong> En la gráfica análoga del índice <em>P</em>, los puntos correspondientes a estrategia conservadora en T&#8202;=&#8202;1&#8209;2 quedan por encima de los puntos baseline, aunque la estrategia conservadora produzca egresos más tardíos. Esas inversiones son la huella visual de la trampa. En la nube de h no hay inversiones: la estrategia con h más alto en cualquier trimestre es siempre la que se titula antes.</div>
    </div>
    <div class="howto-step">
      <div class="howto-icon">2</div>
      <div class="howto-text"><strong>La distancia entre h&#8209;óptima y agresiva cuantifica el costo topológico.</strong> Ambas estrategias inscriben 38 créditos por trimestre, pero h&#8209;óptima respeta el orden topológico del plan (UEA estructuralmente prioritarias primero), mientras que agresiva prioriza las más fáciles. La diferencia de <strong>2.1 trimestres</strong> en el egreso esperado entre ambas mide exactamente lo que se gana&#8202;—&#8202;o se pierde&#8202;—&#8202;por ignorar la topología al elegir qué inscribir, aun con la misma carga total.</div>
    </div>
    <div class="howto-step">
      <div class="howto-icon">3</div>
      <div class="howto-text"><strong>El ruido de calificaciones es visible como dispersión vertical.</strong> Si el índice dependiera de calificaciones, los puntos de una misma estrategia mostrarían dispersión vertical mayor, porque distintos profesores asignarían calificaciones distintas a desempeños equivalentes. En h, esa dispersión no existe por construcción: todos los puntos de una misma estrategia se alinean en una banda estrecha.</div>
    </div>
  </div>
</section>

<!-- ── Sección 4: La penalización topológica ── -->
<section class="section animate-in" id="indice-penalizacion">
  <div class="section-header">
    <h2>La penalización topológica&#8202;—&#8202;lo que ningún índice puede resolver</h2>
    <span class="section-count">Δ(P bloqueado &#8722; desbloqueado)&#8202;=&#8202;&#8722;0.21 pts en Electrónica · T&#8202;=&#8202;12</span>
  </div>
  <p class="section-desc">Tanto el índice <em>P</em> como el índice <em>h</em> comparten una consecuencia inevitable: los estudiantes que son bloqueados por seriaciones aprueban menos créditos por trimestre y, por lo tanto, obtienen valores inferiores en ambos índices. En Electrónica, el plan con mayor profundidad topológica de la División, la brecha entre estudiantes bloqueados y desbloqueados alcanza <strong>0.21 puntos en P</strong> hacia el trimestre 12 (P&#8202;≈&#8202;0.445 para bloqueados frente a P&#8202;≈&#8202;0.655 para desbloqueados). Esta penalización no refleja menor dedicación ni menor capacidad: refleja que el plan de estudios niega acceso a créditos del camino crítico precisamente cuando el estudiante más lo necesita.</p>
  <p class="section-desc">El problema es doble. Primero, el índice <em>P</em> se utiliza como criterio de prelación: los estudiantes con mayor P tienen prioridad de inscripción en los grupos de demanda alta. Esto significa que los estudiantes bloqueados&#8202;—&#8202;quienes tienen P bajo por causas estructurales&#8202;—&#8202;son penalizados <em>dos veces</em>: una vez por el bloqueo topológico en sí, y una segunda vez porque su índice bajo les impide acceder con prioridad a las mismas UEA que los liberarían del bloqueo. El índice h reproduce la misma penalización primaria, aunque elimina la amplificación introducida por el sesgo ACT. Ningún índice de prelación puramente basado en créditos aprobados puede evitar que el bloqueado tenga menor valor que el desbloqueado, porque esa diferencia es consecuencia directa de una estructura de prerrequisitos que restringe la oferta disponible.</p>

  <div class="read-callout">
    <strong>Cómo leer la gráfica (trayectorias P y GA en Electrónica):</strong> las cuatro curvas muestran la evolución del índice P (líneas continuas) y del porcentaje de avance de créditos GA (líneas discontinuas) para estudiantes bloqueados y desbloqueados en el plan de Electrónica. La brecha entre la curva de desbloqueados y la de bloqueados se abre a partir del trimestre 4&#8202;—&#8202;cuando las seriaciones de profundidad media comienzan a actuar&#8202;—&#8202;y alcanza su máximo alrededor del trimestre 12. Nótese que la apertura de la brecha coincide con el período en que el sistema de prelación asigna prioridad de acceso: el índice penaliza a los estudiantes que más necesitarían la prioridad de inscripción.
  </div>

  <div class="chart-card animate-in delay-1">
    <div class="chart-card-header">
      <h3>Doble penalización topológica en Electrónica — trayectorias P y GA</h3>
      <span class="chart-type-tag">Interactivo</span>
    </div>
    <div class="chart-canvas rendering md" id="chart-bloqueados"><div class="chart-skeleton"></div></div>
  </div>

  <div class="recuadro animate-in delay-1">
    <div class="recuadro-title">h es necesario, pero no suficiente</div>
    <p>Reemplazar <em>P</em> por <em>h</em> elimina el incentivo al subregistro, suprime el ruido de heterogeneidad docente y produce una señal ordinalmente coherente para el estudiante y para los sistemas de alerta temprana. Esos son avances reales y no triviales.</p>
    <p>Sin embargo, la penalización topológica persiste bajo h porque su causa es anterior al índice: es la <strong>estructura de seriaciones</strong> la que restringe los créditos disponibles para el estudiante bloqueado, independientemente de cómo se mida el avance. Mientras existan cadenas de prerrequisitos que concentran el acceso a UEA de alto peso crediticio en nodos de alta probabilidad de reprobación, ningún índice de prelación puede compensar el diferencial.</p>
    <p>La intervención prioritaria sigue siendo la identificada en las infografías anteriores: la eliminación quirúrgica del <strong>26.5% de aristas de seriación</strong> que concentra el 80% del impacto sobre el tiempo esperado de egreso (Pareto topológico). La reforma del índice de prelación es complementaria e independiente: puede y debe hacerse en paralelo, pero no sustituye la reforma estructural del plan.</p>
  </div>
</section>

<hr class="divider">

<!-- ── Caja de conexión ── -->
<div class="connect-box animate-in" id="indice-navegacion">
  <div class="connect-box-title">Navegación del Metaanálisis</div>
  <p style="font-size:.82rem;color:var(--mid);margin-bottom:.1rem">Esta infografía cierra el análisis comparativo de índices de prelación. La intervención estructural recomendada&#8202;—&#8202;remoción de seriaciones Pareto-prioritarias&#8202;—&#8202;se detalla en la sección de modificaciones al plan de estudios.</p>
  <div class="connect-links">
    <a class="connect-link" href="meta_4.html">&#8592; Infografía 4 — El índice P</a>
    <a class="connect-link" href="metanalisis.html">Índice del Metaanálisis</a>
    <a class="connect-link" href="meta_5.html">Complejidad multidimensional &#8594;</a>
  </div>
</div>

</main>
"""

META6_JS = """
(function(){

renderChart('chart-delta-incentivos',
  () => {
    const t = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26];
    const dP = [0.0223,0.1201,0.0555,-0.0055,-0.0136,-0.0468,-0.0530,-0.0636,-0.0725,-0.0817,-0.0910,-0.0976,-0.1066,-0.1142,-0.1259,-0.1375,-0.1413,-0.1393,-0.1288,-0.1123,-0.0967,-0.0773,-0.0627,-0.0468,-0.0372,-0.0313];
    const dh = [0.0324,-0.0244,-0.0752,-0.0858,-0.0913,-0.1207,-0.1335,-0.1396,-0.1457,-0.1500,-0.1528,-0.1561,-0.1562,-0.1568,-0.1554,-0.1427,-0.1251,-0.1049,-0.0859,-0.0692,-0.0535,-0.0409,-0.0286,-0.0234,-0.0172,-0.0127];
    const traceZero = { x:[1,26], y:[0,0], mode:'lines', line:{color:'#9E9E9E',dash:'dash',width:1}, showlegend:false, hoverinfo:'skip' };
    const shadeZone = { x:[1,3,3,1], y:[-0.17,-0.17,0.14,0.14], fill:'toself', fillcolor:'rgba(200,45,35,0.08)', line:{color:'transparent'}, mode:'lines', showlegend:false, hoverinfo:'skip', type:'scatter' };
    const traceDeltaP = { x:t, y:dP, mode:'lines', name:'ΔP (conservadora − baseline)', line:{color:'#C82D23',dash:'dash',width:2.5}, hovertemplate:'T=%{x}<br>ΔP=%{y:.4f}<extra></extra>' };
    const traceDeltaH = { x:t, y:dh, mode:'lines', name:'Δh (conservadora − baseline)', line:{color:'#2E7D32',dash:'solid',width:2.5}, hovertemplate:'T=%{x}<br>Δh=%{y:.4f}<extra></extra>' };
    return [shadeZone, traceZero, traceDeltaP, traceDeltaH];
  },
  () => ({
    ...LAYOUT_BASE,
    margin:{l:10,r:10,t:10,b:50},
    xaxis:{title:{text:'Trimestre T',font:FONT},tickmode:'linear',tick0:1,dtick:2,automargin:true},
    yaxis:{title:{text:'Δ índice (conservadora − baseline)',font:FONT},zeroline:false,automargin:true},
    legend:{orientation:'h',xanchor:'center',x:0.5,y:1.08,font:FONT},
    annotations:[
      {x:2,y:0.1201,xref:'x',yref:'y',text:'T=2: ΔP≈+0.120',showarrow:true,arrowhead:2,arrowsize:0.8,ax:30,ay:-28,font:{color:'#C82D23',size:11},arrowcolor:'#C82D23'},
      {x:2,y:-0.0244,xref:'x',yref:'y',text:'T=2: Δh≈−0.024',showarrow:true,arrowhead:2,arrowsize:0.8,ax:40,ay:28,font:{color:'#2E7D32',size:11},arrowcolor:'#2E7D32'},
      {x:2,y:0.14,xref:'x',yref:'y',text:'Ventaja transitoria de P',showarrow:false,font:{color:'#C82D23',size:10,style:'italic'},xanchor:'center'}
    ]
  })
);

renderChart('chart-bloqueados',
  () => {
    const t = [5,6,7,8,9,10,11,12,13,14];
    const P_bloq  = [0.5305,0.5267,0.5085,0.4742,0.4681,0.4446,0.4334,0.4451,0.3918,0.3754];
    const P_nobl  = [0.7461,0.7038,0.6665,0.6552,0.6471,0.6470,0.6483,0.6546,0.6583,0.6657];
    const GA_bloq = [0.1927,0.2281,0.2546,0.2757,0.2933,0.3116,0.3266,0.3500,0.3508,0.3460];
    const GA_nobl = [0.3088,0.3404,0.3737,0.4092,0.4445,0.4843,0.5276,0.5723,0.6179,0.6664];
    const trP_bloq = { x:t, y:P_bloq, name:'Bloqueados', mode:'lines+markers', line:{color:'#C82D23',dash:'dash',width:2}, marker:{size:5}, xaxis:'x', yaxis:'y', hovertemplate:'T=%{x}<br>P=%{y:.3f}<extra>Bloqueados</extra>' };
    const trP_nobl = { x:t, y:P_nobl, name:'Desbloqueados', mode:'lines+markers', line:{color:'#1565C0',dash:'solid',width:2}, marker:{size:5}, xaxis:'x', yaxis:'y', hovertemplate:'T=%{x}<br>P=%{y:.3f}<extra>Desbloqueados</extra>' };
    const trGA_bloq = { x:t, y:GA_bloq, name:'Bloqueados', mode:'lines+markers', line:{color:'#C82D23',dash:'dash',width:2}, marker:{size:5}, xaxis:'x2', yaxis:'y2', showlegend:false, hovertemplate:'T=%{x}<br>GA=%{y:.3f}<extra>Bloqueados</extra>' };
    const trGA_nobl = { x:t, y:GA_nobl, name:'Desbloqueados', mode:'lines+markers', line:{color:'#1565C0',dash:'solid',width:2}, marker:{size:5}, xaxis:'x2', yaxis:'y2', showlegend:false, hovertemplate:'T=%{x}<br>GA=%{y:.3f}<extra>Desbloqueados</extra>' };
    return [trP_bloq, trP_nobl, trGA_bloq, trGA_nobl];
  },
  () => ({
    ...LAYOUT_BASE,
    margin:{l:10,r:10,t:10,b:50},
    grid:{rows:1,columns:2,pattern:'independent'},
    xaxis:{title:{text:'Trimestre T',font:FONT},domain:[0,0.46],automargin:true},
    yaxis:{title:{text:'Índice P',font:FONT},automargin:true,range:[0.3,0.82]},
    xaxis2:{title:{text:'Trimestre T',font:FONT},domain:[0.54,1.0],automargin:true},
    yaxis2:{title:{text:'Grado de Avance (GA)',font:FONT},automargin:true,range:[0.15,0.72]},
    legend:{orientation:'h',xanchor:'center',x:0.5,y:1.08,font:FONT},
    annotations:[
      {x:12,y:0.4451,xref:'x',yref:'y',text:'ΔP≈0.21',showarrow:true,arrowhead:2,arrowsize:0.8,ax:-40,ay:0,font:{color:'#C82D23',size:11},arrowcolor:'#C82D23'},
      {x:0.23,y:1.03,xref:'paper',yref:'paper',text:'<b>Índice P — Electrónica</b>',showarrow:false,font:{size:11,color:'#1C1C1C'},xanchor:'center'},
      {x:0.77,y:1.03,xref:'paper',yref:'paper',text:'<b>Grado de Avance — Electrónica</b>',showarrow:false,font:{size:11,color:'#1C1C1C'},xanchor:'center'}
    ]
  })
);

renderChart('chart-trampa-scatter',
  () => {
    const data = {
      agresiva:     { plans:['amb','civ','com','ele','elo','fis','ind','mec','qui'], P8:[0.6949,0.6676,0.6418,0.6424,0.6623,0.6565,0.6979,0.6781,0.6779], ET:[17.55,16.31,16.83,18.23,19.04,16.31,16.06,16.61,17.88] },
      baseline:     { plans:['amb','civ','com','ele','elo','fis','ind','mec','qui'], P8:[0.6328,0.6295,0.5922,0.6086,0.6116,0.6188,0.6455,0.6065,0.6530], ET:[20.21,19.38,19.85,20.87,21.41,18.80,19.03,20.65,19.43] },
      conservadora: { plans:['amb','civ','com','ele','elo','fis','ind','mec','qui'], P8:[0.5803,0.5573,0.5413,0.5455,0.5477,0.5298,0.6265,0.5587,0.5577], ET:[24.73,24.26,24.46,24.55,25.21,23.59,24.48,23.99,24.30] }
    };
    const lbl = {amb:'Amb',civ:'Civ',com:'Comp',ele:'Ele',elo:'Elo',fis:'Fis',ind:'Ind',mec:'Mec',qui:'Qui'};
    const mkT = (pl) => pl.map(p=>lbl[p]||p);
    return [
      { x:data.agresiva.P8, y:data.agresiva.ET, mode:'markers+text', name:'Agresiva', marker:{color:'#90CAF9',symbol:'triangle-up',size:10,line:{color:'#1565C0',width:1}}, text:mkT(data.agresiva.plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>P(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>Agresiva</extra>' },
      { x:data.baseline.P8, y:data.baseline.ET, mode:'markers+text', name:'Baseline', marker:{color:'#1C1C1C',symbol:'circle',size:10}, text:mkT(data.baseline.plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>P(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>Baseline</extra>' },
      { x:data.conservadora.P8, y:data.conservadora.ET, mode:'markers+text', name:'Conservadora', marker:{color:'#C82D23',symbol:'square',size:10}, text:mkT(data.conservadora.plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>P(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>Conservadora</extra>' }
    ];
  },
  () => ({
    ...LAYOUT_BASE,
    margin:{l:10,r:10,t:10,b:50},
    xaxis:{title:{text:'Índice P en T=8',font:FONT},automargin:true,range:[0.50,0.73]},
    yaxis:{title:{text:'E[T | titulación] — menor es mejor',font:FONT},automargin:true,range:[15.0,27.0]},
    legend:{orientation:'h',xanchor:'center',x:0.5,y:1.08,font:FONT},
    annotations:[{x:0.5477,y:25.21,xref:'x',yref:'y',text:'conservadora: E[T]=25.2 trim',showarrow:true,arrowhead:2,arrowsize:0.8,ax:-60,ay:-20,font:{color:'#C82D23',size:10},arrowcolor:'#C82D23'}]
  })
);

renderChart('chart-h-coherencia',
  () => {
    const plans = ['amb','civ','com','ele','elo','fis','ind','mec','qui'];
    const lbl = {amb:'Amb',civ:'Civ',com:'Comp',ele:'Ele',elo:'Elo',fis:'Fis',ind:'Ind',mec:'Mec',qui:'Qui'};
    const d = {
      agresiva:     { h8:[0.5531,0.5253,0.5079,0.4843,0.5090,0.4982,0.5676,0.5390,0.5391], ET:[17.47,16.38,16.79,18.16,19.08,16.36,16.02,16.55,17.76] },
      baseline:     { h8:[0.4351,0.4226,0.4083,0.4051,0.4089,0.4054,0.4650,0.3898,0.4822], ET:[20.10,19.53,19.88,20.87,21.58,18.87,18.94,20.68,19.39] },
      conservadora: { h8:[0.2975,0.2834,0.2887,0.2701,0.2764,0.2529,0.3310,0.2777,0.2880], ET:[24.82,24.22,24.52,24.59,25.26,23.57,24.41,24.02,24.25] },
      h_optima:     { h8:[0.3914,0.4285,0.4719,0.4015,0.4146,0.3957,0.4912,0.4087,0.4152], ET:[19.38,18.24,18.43,19.92,21.20,19.36,18.98,18.44,19.81] }
    };
    const mkT = (pl) => pl.map(p=>lbl[p]||p);
    const connectors = plans.map((p,i) => ({ x:[d.agresiva.h8[i],d.h_optima.h8[i]], y:[d.agresiva.ET[i],d.h_optima.ET[i]], mode:'lines', line:{color:'rgba(0,131,143,0.4)',width:1.5,dash:'dot'}, showlegend:false, hoverinfo:'skip', type:'scatter' }));
    return [
      ...connectors,
      { x:d.agresiva.h8, y:d.agresiva.ET, mode:'markers+text', name:'Agresiva', marker:{color:'#90CAF9',symbol:'triangle-up',size:11,line:{color:'#1565C0',width:1}}, text:mkT(plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>h(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>Agresiva</extra>' },
      { x:d.baseline.h8, y:d.baseline.ET, mode:'markers+text', name:'Baseline', marker:{color:'#1C1C1C',symbol:'circle',size:11}, text:mkT(plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>h(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>Baseline</extra>' },
      { x:d.conservadora.h8, y:d.conservadora.ET, mode:'markers+text', name:'Conservadora', marker:{color:'#C82D23',symbol:'square',size:11}, text:mkT(plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>h(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>Conservadora</extra>' },
      { x:d.h_optima.h8, y:d.h_optima.ET, mode:'markers+text', name:'h-óptima', marker:{color:'#2E7D32',symbol:'diamond',size:11}, text:mkT(plans), textposition:'top center', textfont:{size:9,color:'#646464'}, hovertemplate:'%{text}<br>h(T=8)=%{x:.3f}<br>E[T]=%{y:.1f}<extra>h-óptima</extra>' }
    ];
  },
  () => ({
    ...LAYOUT_BASE,
    margin:{l:10,r:10,t:10,b:50},
    xaxis:{title:{text:'Índice h en T=8',font:FONT},automargin:true,range:[0.22,0.60]},
    yaxis:{title:{text:'E[T | titulación] — menor es mejor',font:FONT},automargin:true,range:[15.0,27.0]},
    legend:{orientation:'h',xanchor:'center',x:0.5,y:1.08,font:FONT},
    annotations:[{x:0.50,y:0.14,xref:'paper',yref:'paper',text:'Líneas punteadas: agresiva → h-óptima (costo prom. ≈ +2.1 trim)',showarrow:false,font:{color:'#00838F',size:10,style:'italic'},xanchor:'center'}]
  })
);

renderChart('chart-ra-deficit',
  () => {
    // Plans sorted ascending by c_s (calibrated load from model)
    const plans = ['Mec.','Amb.','Civil','Eléctr.','Electron.','Física','Comp.','Ind.','Química','Metalurgia'];
    const cs    = [27, 30, 30, 30, 30, 30, 31, 31, 33, 41];
    const ra    = cs.map(c => Math.min(c/38, 1));
    const colors = cs.map(c => c >= 38 ? '#2E7D32' : '#C82D23');
    const ref   = cs.map(() => 38);

    const trBar = {
      type:'bar', orientation:'v',
      name:'Carga efectiva c<sub>s</sub> (cr/trim)',
      x:plans, y:cs,
      marker:{color:colors, opacity:.85, line:{color:'rgba(0,0,0,.07)',width:.5}},
      hovertemplate:'<b>%{x}</b><br>c<sub>s</sub> = %{y} cr/trim<br>RA = %{customdata:.2f}<extra></extra>',
      customdata:ra,
      text:cs.map(c=>'<b>'+c+'</b>'),
      textposition:'outside', textfont:{...FONT, size:11}, cliponaxis:false,
    };
    const trRef = {
      type:'scatter', mode:'lines',
      name:'Referencia RA = 38 cr/trim',
      x:plans, y:ref,
      line:{color:'#00838F', width:2, dash:'dash'},
      hovertemplate:'Calibrador RA: 38 cr/trim<extra></extra>',
    };
    return [trBar, trRef];
  },
  () => ({
    ...LAYOUT_BASE,
    margin:{l:10,r:10,t:10,b:80},
    xaxis:{tickfont:{...FONT,size:11}, automargin:true, tickangle:-30},
    yaxis:{title:{text:'cr/trim',font:FONT}, range:[0,48], gridcolor:'#E4E2DF', dtick:10},
    legend:{orientation:'h', xanchor:'center', x:.5, y:1.08, font:FONT},
    annotations:[
      {x:'Metalurgia', y:44, text:'único plan con c<sub>s</sub> ≥ 38',
       showarrow:true, arrowhead:2, arrowsize:.8, ax:0, ay:-24,
       font:{color:'#2E7D32', size:10}, arrowcolor:'#2E7D32'},
      {x:0.5, y:-0.22, xref:'paper', yref:'paper',
       text:'9 de 10 planes operan por debajo del calibrador — déficit estructural de RA, no fracaso individual',
       showarrow:false, font:{color:'#646464', size:10, style:'italic'}, xanchor:'center'}
    ]
  })
);

})();
"""

meta6_html = page_shell(
    title='Metaanálisis · Infografía 6 — Índices de prelación P y h',
    active='meta_6.html',
    content=META6_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="#indice-p-deficit"><span class="bottom-nav-icon">📊</span>RA crónico</a>
      <a class="bottom-nav-item" href="#indice-trampa-act"><span class="bottom-nav-icon">⚠️</span>Trampa ACT</a>
      <a class="bottom-nav-item" href="#indice-h"><span class="bottom-nav-icon">h</span>Índice h</a>
      <a class="bottom-nav-item" href="#indice-penalizacion"><span class="bottom-nav-icon">🔒</span>Penalización</a>
    """,
    data_vars='',
    extra_js=META6_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# MODIFICACIONES.HTML
# ══════════════════════════════════════════════════════════════════════════════
MOD_CONTENT = f"""
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Modificaciones 2026 · Revisión curricular con IA</p>
      <h1><strong>Análisis de similitud</strong><br>de programas de UEA 2026</h1>
      <p class="hero-sub">Sistema multi-agente que revisó las modificaciones curriculares 2026,
        comparó automáticamente los 1,505 pares candidatos de UEAs y generó los diez oficios
        de revisión institucional.</p>
      <div class="hero-bar"></div>
      <div class="stats">
        <div class="stat"><span class="stat-val kpi-counter" data-target="10">0</span><span class="stat-lbl">Licenciaturas revisadas</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="626">0</span><span class="stat-lbl">Programas analizados</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="391">0</span><span class="stat-lbl">Pares similares inter-plan</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="10">0</span><span class="stat-lbl">Oficios generados</span></div>
      </div>
    </div>
    {palette_chips()}
  </div>
</header>

<main class="main">

  <!-- ── Sección 1: Similitud inter-plan ── -->
  <section class="section animate-in" id="mod-inter">
    <div class="section-header">
      <h2>Similitud de contenidos entre licenciaturas</h2>
      <span class="section-count">391 pares similares</span>
    </div>
    <p class="section-desc">
      El análisis comparó los programas de UEA entre los diez planes de estudio.
      El par con mayor similitud es Industrial–Mecánica (38 programas equivalentes).
      La figura interactiva muestra cuántos pares equivalentes (puntuación ≥ 0.70) existen
      entre cada par de licenciaturas.
    </p>
    <div class="read-callout">
      <strong>Cómo usar esta figura:</strong> cada barra representa un par de licenciaturas y
      su altura indica cuántos programas de UEA tienen contenido equivalente. Pasa el cursor
      sobre cualquier barra para ver los nombres del par y el conteo exacto. Los pares con barras
      más altas son candidatos a contenidos transversales compartidos.
    </div>
    <div class="fig-iframe-wrap animate-in delay-1">
      <iframe class="fig-iframe" src="figuras/barras_interplan_top.html" loading="lazy"
              title="Pares equivalentes por par de licenciaturas"></iframe>
    </div>
  </section>

  <!-- ── Sección 2: Distribución de niveles ── -->
  <section class="section animate-in" id="mod-niveles">
    <div class="section-header">
      <h2>Distribución de niveles de similitud</h2>
      <span class="section-count">1,505 pares evaluados</span>
    </div>
    <p class="section-desc">
      De los 1,505 pares candidatos evaluados con inteligencia artificial,
      203 resultaron idénticos o casi idénticos (puntuación ≥ 0.85) y
      186 con similitud muy alta (0.70–0.84). El 22% de los pares evaluados no mostraron
      traslape relevante de contenidos.
    </p>
    <div class="read-callout">
      <strong>Qué muestran estas gráficas:</strong> la dona de la izquierda muestra la proporción
      de pares evaluados en cada nivel de similitud. La tabla de referencia describe qué significa
      cada nivel en términos de contenido académico y qué acción sugiere.
    </div>
    <div class="chart-guide-grid animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Distribución de niveles de similitud · 1,505 pares</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering md" id="chart-niveles"><div class="chart-skeleton"></div></div>
      </div>
      <div class="howto-card">
        <h3>🎨 Escala de similitud</h3>
        <div class="howto-step"><div class="howto-icon" style="background:#f3e5f5;color:#6a1b9a">●</div>
          <div class="howto-text"><strong>Idéntico (203 pares):</strong> contenidos prácticamente iguales — candidatos a unificarse o reconocerse como UEA transversal.</div></div>
        <div class="howto-step"><div class="howto-icon" style="background:#ffebee;color:#b71c1c">●</div>
          <div class="howto-text"><strong>Muy alto (186 pares):</strong> gran solapamiento de contenidos — conviene revisar si la diferenciación es suficiente.</div></div>
        <div class="howto-step"><div class="howto-icon" style="background:#fce4ec;color:#c62828">●</div>
          <div class="howto-text"><strong>Alto (175 pares):</strong> temas comunes significativos — posible candidato a contenido transversal si ocurre en varios planes.</div></div>
        <div class="howto-step"><div class="howto-icon" style="background:#fff3e0;color:#bf360c">●</div>
          <div class="howto-text"><strong>Moderado (389 pares):</strong> algunos temas en común, pero diferenciación clara de enfoque o profundidad.</div></div>
        <div class="howto-step"><div class="howto-icon" style="background:#fffde7;color:#795500">●</div>
          <div class="howto-text"><strong>Bajo / Sin traslape (552 pares):</strong> contenidos distintos — no se requiere acción de unificación.</div></div>
      </div>
    </div>
  </section>

  <!-- ── Sección 3: Figuras adicionales ── -->
  <section class="section animate-in" id="mod-figuras">
    <div class="section-header">
      <h2>Figuras interactivas de análisis</h2>
      <span class="section-count">4 visualizaciones</span>
    </div>
    <p class="section-desc">
      Cada figura explora una dimensión distinta del análisis de similitud: desde la vista
      global de la matriz de similitud entre planes, hasta el agrupamiento jerárquico de las
      120 UEAs más conectadas y la red de vínculos entre licenciaturas.
    </p>
    <div class="read-callout">
      <strong>Cómo usar estas figuras:</strong> haz clic en cualquier tarjeta para abrir
      la figura interactiva en pantalla completa. Dentro de ella, pasa el cursor o toca
      las celdas para ver los valores exactos. Puedes hacer zoom con los controles de la
      esquina superior derecha.
    </div>
    <div class="grid grid-2 animate-in delay-1">
      <div class="card" onclick="window.open('figuras/heatmap_licenciaturas.html','_blank','noopener')">
        <div class="card-thumb">
          <img src="figuras/heatmap_licenciaturas.png" alt="Mapa de similitud entre licenciaturas">
          <span class="card-badge badge-interactive">Interactivo</span>
        </div>
        <div class="card-body">
          <p class="card-title">Similitud promedio entre los diez planes</p>
          <p class="card-desc">Mapa de calor 10×10. El rojo intenso indica mayor similitud. La diagonal (cada licenciatura consigo misma) siempre es roja — ignórala y busca las celdas oscuras fuera de ella.</p>
          <div class="card-footer">
            <span class="card-meta">10 × 10 licenciaturas</span>
            <span class="card-btn">Abrir <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 8h10M9 4l4 4-4 4"/></svg></span>
          </div>
        </div>
      </div>
      <div class="card" onclick="window.open('figuras/scatter_tsne.html','_blank','noopener')">
        <div class="card-thumb">
          <img src="figuras/scatter_tsne.png" alt="Mapa t-SNE">
          <span class="card-badge badge-interactive">Interactivo</span>
        </div>
        <div class="card-body">
          <p class="card-title">Mapa de agrupamiento t-SNE</p>
          <p class="card-desc">Cada punto es una UEA. El algoritmo acerca los programas de contenido similar y aleja los distintos. Los cúmulos de un mismo color revelan familias de contenidos compartidos.</p>
          <div class="card-footer">
            <span class="card-meta">626 programas · t-SNE</span>
            <span class="card-btn">Abrir <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 8h10M9 4l4 4-4 4"/></svg></span>
          </div>
        </div>
      </div>
      <div class="card" onclick="window.open('figuras/grafo_licenciaturas_30.html','_blank','noopener')">
        <div class="card-thumb">
          <img src="figuras/grafo_licenciaturas_30.png" alt="Red de conexiones">
          <span class="card-badge badge-interactive">Interactivo</span>
        </div>
        <div class="card-body">
          <p class="card-title">Red de conexiones entre licenciaturas</p>
          <p class="card-desc">Cada nodo es una licenciatura. Las aristas conectan pares con al menos un programa equivalente. Las licenciaturas más conectadas comparten más contenidos con el resto de la División.</p>
          <div class="card-footer">
            <span class="card-meta">Inter-licenciaturas · s ≥ 0.30</span>
            <span class="card-btn">Abrir <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 8h10M9 4l4 4-4 4"/></svg></span>
          </div>
        </div>
      </div>
      <div class="card" onclick="window.open('figuras/heatmap_clustered.html','_blank','noopener')">
        <div class="card-thumb">
          <img src="figuras/heatmap_clustered.png" alt="Clustering jerárquico">
          <span class="card-badge badge-interactive">Interactivo</span>
        </div>
        <div class="card-body">
          <p class="card-title">Agrupamiento jerárquico de las 120 UEAs más conectadas</p>
          <p class="card-desc">Las 120 UEAs con más pares similares, agrupadas automáticamente por contenido (método Ward, s ≥ 0.40). La barra lateral de color identifica la licenciatura de cada programa.</p>
          <div class="card-footer">
            <span class="card-meta">120 UEAs · clustering Ward</span>
            <span class="card-btn">Abrir <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 8h10M9 4l4 4-4 4"/></svg></span>
          </div>
        </div>
      </div>
    </div>
  </section>

  <hr class="divider">

  <div class="connect-box animate-in" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem">
    <div>
      <div class="connect-box-title">Reporte de similitud completo</div>
      <p style="font-size:.84rem;color:var(--mid);line-height:1.6;margin:0">Análisis de duplicación y similitud semántica entre UEAs · 35 pp. · PDF</p>
    </div>
    <a href="Reporte_Similitud_DCBI_2026.pdf" download
       style="display:inline-flex;align-items:center;gap:.5rem;background:var(--red);color:#fff;
              text-decoration:none;padding:.7rem 1.4rem;border-radius:8px;font-size:.85rem;
              font-weight:700;white-space:nowrap;transition:opacity .15s"
       onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      ↓ Descargar PDF
    </a>
  </div>

  <div class="connect-box animate-in" style="margin-top:1rem">
    <div class="connect-box-title">¿Cómo se conecta con los demás proyectos?</div>
    <p style="font-size:.84rem;color:var(--mid);line-height:1.6">
      Las UEAs identificadas como idénticas o muy similares entre planes son exactamente las
      que aparecen como cuellos de botella en el Metaanálisis — materias del Tronco General
      que concentran alta demanda y llenan sus grupos rápidamente en el CAT. Los tres análisis
      apuntan al mismo núcleo estructural de la División.
    </p>
    <div class="connect-links">
      <a class="connect-link" href="cat26p.html">🗓 Ver CAT 26P →</a>
      <a class="connect-link" href="metanalisis.html">📊 Ver Metaanálisis →</a>
    </div>
  </div>

</main>
"""

MOD_JS = """
(function(){
  const D = window.__MOD__;

  renderChart('chart-niveles',
    () => [{
      type:'pie', hole:.45,
      labels:D.simLevels.labels,
      values:D.simLevels.values,
      marker:{colors:D.simLevels.colors, line:{color:'#fff',width:1.5}},
      textinfo:'label+percent',
      hovertemplate:'<b>%{label}</b><br>%{value} pares (%{percent})<extra></extra>',
      textfont:{...FONT,size:11}, direction:'clockwise',
      sort:false,
    }],
    () => ({
      margin:{l:8,r:8,t:8,b:8}, showlegend:false,
      annotations:[{text:'<b>1,505</b><br><span style="font-size:10px;color:#888">pares</span>',
        showarrow:false, font:{size:15,color:'#C82D23',family:'Helvetica Neue,Arial'}}]
    })
  );
})();
"""

mod_html = page_shell(
    title='Modificaciones 2026 — Revisión Curricular con IA',
    active='modificaciones.html',
    content=MOD_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="#mod-inter"><span class="bottom-nav-icon">⊞</span>Inter-plan</a>
      <a class="bottom-nav-item" href="#mod-niveles"><span class="bottom-nav-icon">🎨</span>Niveles</a>
      <a class="bottom-nav-item" href="#mod-figuras"><span class="bottom-nav-icon">🗺</span>Figuras</a>
    """,
    data_vars=f'window.__MOD__ = {DATA_MOD};',
    extra_js=MOD_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# CAT26P_CAPACIDAD.HTML
# ══════════════════════════════════════════════════════════════════════════════
CAP_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">CAT 26P · Análisis de capacidad máxima de admisión</p>
      <h1><strong>¿Cuántos estudiantes</strong><br>puede absorber el CAT?</h1>
      <p class="hero-sub">Análisis de la capacidad instalada de los 102 grupos del Tronco
        General, la UEA cuello de botella y el número máximo de estudiantes admisibles
        con la infraestructura actual — y qué se necesita para crecer.</p>
      <div class="hero-bar"></div>
      <div class="stats">
        <div class="stat"><span class="stat-val kpi-counter" data-target="864">0</span>
          <span class="stat-lbl">Capacidad máxima del catálogo</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="674">0</span>
          <span class="stat-lbl">Inscritos 26-P</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="190">0</span>
          <span class="stat-lbl">Plazas disponibles</span></div>
        <div class="stat"><span class="stat-val kpi-counter" data-target="24">0</span>
          <span class="stat-lbl">Secciones Lab. (cuello de botella)</span></div>
      </div>
    </div>
    <div style="display:flex;flex-direction:column;gap:.5rem;min-width:200px;align-self:center">
      <p class="palette-title" style="font-size:.66rem;color:rgba(255,255,255,.4);text-transform:uppercase;letter-spacing:.1em">Las 6 UEAs del Tronco</p>
      <div style="display:flex;flex-direction:column;gap:.4rem">
        <div class="chip"><span class="chip-dot" style="background:#00838F"></span>Intro. a la Ingeniería &nbsp; 17 grp × 45</div>
        <div class="chip"><span class="chip-dot" style="background:#1565C0"></span>Intro. a la Física &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 17 grp × 45</div>
        <div class="chip"><span class="chip-dot" style="background:#2E7D32"></span>Comp. de Matemáticas &nbsp; 17 grp × 45</div>
        <div class="chip"><span class="chip-dot" style="background:#E65100"></span>Intro. al Cálculo &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 17 grp × 45</div>
        <div class="chip"><span class="chip-dot" style="background:#558B2F"></span>Estr. Atómica &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 17 grp × 45</div>
        <div class="chip"><span class="chip-dot" style="background:#C82D23"></span>Lab. Reacciones &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 24 grp × 36 ← cuello</div>
      </div>
    </div>
  </div>
</header>

<main class="main">

  <!-- ── Sección 1: Capacidad por UEA ── -->
  <section class="section animate-in" id="cap-uea">
    <div class="section-header">
      <h2>Capacidad instalada vs. inscritos por UEA</h2>
      <span class="section-count">6 UEAs · catálogo 26-P</span>
    </div>
    <p class="section-desc">
      El cuello de botella del catálogo es el Laboratorio de Reacciones Químicas: con 24 secciones
      de 36 plazas cada una, su capacidad total es <strong>864 estudiantes</strong>. Las UEAs
      teóricas del catálogo tienen capacidad superior; la mínima — Introducción a la Física — llega
      a <strong>900 plazas</strong> (20 sec × 45) en el catálogo completo. Con 674 inscritos,
      quedan <strong>190 plazas libres</strong> en el Laboratorio.
    </p>
    <div class="read-callout">
      <strong>Cómo leer la gráfica:</strong> cada barra tiene dos segmentos. El rojo es la
      matrícula actual (674 estudiantes, igual para todas las UEA en el ciclo 26-P). El gris es la
      capacidad disponible no utilizada en los grupos del ciclo. La flecha vertical indica el techo
      del catálogo completo (864). El Laboratorio de Reacciones Químicas es el primero en
      saturarse cuando el padrón crece.
    </div>
    <div class="chart-guide-grid animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Inscritos y capacidad disponible por UEA del Tronco General</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering md" id="chart-cap-uea"><div class="chart-skeleton"></div></div>
      </div>
      <div class="howto-card">
        <h3>Qué buscar en esta gráfica</h3>
        <div class="howto-step"><div class="howto-icon">1</div>
          <div class="howto-text"><strong>Los 674 inscritos</strong> son el segmento rojo, idéntico en todas las UEA — un estudiante toma las seis materias en el mismo ciclo.</div></div>
        <div class="howto-step"><div class="howto-icon">2</div>
          <div class="howto-text"><strong>El cuello de botella</strong> es la UEA con menor capacidad total. En el catálogo 26-P es el Lab. Reacciones: 24 secciones × 36 = 864 plazas.</div></div>
        <div class="howto-step"><div class="howto-icon">3</div>
          <div class="howto-text"><strong>El segmento gris = plazas libres</strong> en los grupos de 26-P. El Lab. tiene 190 plazas libres (864 − 674). Las teorías tienen más capacidad que el Lab en el catálogo completo.</div></div>
        <div class="howto-step"><div class="howto-icon">4</div>
          <div class="howto-text"><strong>Techo verificado por PL:</strong> la capacidad de 864 fue confirmada evaluando <strong>1,927,442 combinaciones</strong> de asignación de grupos (programación lineal). El resultado es invariante en todos los escenarios.</div></div>
      </div>
    </div>
  </section>

  <!-- ── Sección 2: Ocupación por grupo ── -->
  <section class="section animate-in" id="cap-grupos">
    <div class="section-header">
      <h2>Ocupación por grupo</h2>
      <span class="section-count">Lab. Reacciones · grupos 26-P</span>
    </div>
    <p class="section-desc">
      No todos los grupos se usan igual. El grupo CAT07 está al 120–125% de su cupo
      (es un grupo compartido con mayor aforo por convenio de aula). Los grupos CAT13,
      CAT14 y CAT15 del laboratorio solo tienen entre 9 y 17 estudiantes — son grupos
      de horario vespertino con baja demanda. Las UEAs de teoría siguen exactamente el
      mismo patrón de ocupación porque el algoritmo asigna las 6 UEAs en bloque.
    </p>
    <div class="read-callout">
      <strong>Cómo leer estas barras:</strong> cada barra representa un grupo y su longitud
      indica qué porcentaje del cupo está ocupado. El color indica el nivel: verde = hasta 80%,
      ámbar = 80–99%, rojo = 100% (lleno), burdeos = más del 100% (grupo de mayor aforo).
    </div>
    <div class="chart-duo animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Ocupación por grupo · Lab. Reacciones Químicas (cupo 36)</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering lg" id="chart-cap-lab"><div class="chart-skeleton"></div></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Ocupación por grupo · UEAs de teoría (cupo 45, patrón uniforme)</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering lg" id="chart-cap-ref"><div class="chart-skeleton"></div></div>
      </div>
    </div>
  </section>

  <!-- ── Sección 3: Sensibilidad ── -->
  <section class="section animate-in" id="cap-sens">
    <div class="section-header">
      <h2>Análisis de sensibilidad: ¿cuánto se puede crecer?</h2>
      <span class="section-count">Simulación: +0 a +8 grupos de Lab</span>
    </div>
    <p class="section-desc">
      Con 24 secciones de Lab., el cuello de botella es el laboratorio (864 &lt; 900).
      Al abrir <strong>1 sección adicional</strong> (25 × 36 = 900), Lab. e Intro. a la
      Física alcanzan exactamente el mismo techo. A partir de ese punto, agregar más grupos
      de laboratorio no incrementa la capacidad: para seguir creciendo se requieren también
      secciones adicionales de Física (45 lugares c/u) en paralelo. Este resultado fue
      verificado sobre <strong>1,927,442 combinaciones</strong> de programación lineal.
    </p>
    <div class="read-callout">
      <strong>Cómo leer la gráfica:</strong> el eje horizontal son los grupos adicionales de
      Lab. Reacciones. El eje vertical es la capacidad máxima de admisión resultante. La
      línea punteada marca el techo de Intro. a la Física en el catálogo completo
      (20 sec × 45 = 900). Una vez que Lab. alcanza esa línea, agregar más grupos de Lab
      sin nuevas secciones de Física no incrementa la capacidad.
    </div>
    <div class="chart-guide-grid animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Capacidad máxima de admisión al agregar grupos de Lab. Reacciones</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering" id="chart-sens"><div class="chart-skeleton"></div></div>
      </div>
      <div class="recuadro" style="margin:0;align-self:start">
        <div class="recuadro-title">Proyecciones de capacidad (catálogo completo)</div>
        <p>La capacidad instalada (864) absorbe los 674 inscritos 26-P con <strong>190 plazas libres</strong>. Costo marginal de esas plazas: cero.</p>
        <table style="width:100%;border-collapse:collapse;font-size:.76rem;margin:.55rem 0">
          <thead>
            <tr style="background:rgba(200,45,35,.08)">
              <th style="padding:.3rem .45rem;text-align:left;border-bottom:1px solid #E4E2DF">N admisible</th>
              <th style="padding:.3rem .45rem;text-align:center;border-bottom:1px solid #E4E2DF">n<sub>Lab</sub></th>
              <th style="padding:.3rem .45rem;text-align:left;border-bottom:1px solid #E4E2DF">Acción requerida</th>
            </tr>
          </thead>
          <tbody>
            <tr><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">&#8804; 864</td><td style="text-align:center;padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">24</td><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC;color:#2E7D32">Sin inversión marginal</td></tr>
            <tr><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">865 &#8211; 900</td><td style="text-align:center;padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">25</td><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">+1 Lab (13:00 L&#160;o&#160;V)</td></tr>
            <tr><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">901 &#8211; 935</td><td style="text-align:center;padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">26</td><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">+2 Lab + +1 Física</td></tr>
            <tr><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">936 &#8211; 972</td><td style="text-align:center;padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">27</td><td style="padding:.22rem .45rem;border-bottom:1px solid #F0EFEC">+3 Lab + +2 Física</td></tr>
            <tr><td style="padding:.22rem .45rem">&#8805; 972</td><td style="text-align:center;padding:.22rem .45rem">&#8805; 28</td><td style="padding:.22rem .45rem">Coordinación múltiple</td></tr>
          </tbody>
        </table>
        <p style="font-size:.72rem;color:var(--mid);margin:.35rem 0 0">Verificado sobre 1,927,442 combinaciones (PL). El umbral 900 corresponde a 25 × 36 = 20 × 45 — coincidencia exacta Lab–Física.</p>
      </div>
    </div>
  </section>

  <!-- ── Sección 4: Franjas horarias disponibles ── -->
  <section class="section animate-in" id="cap-franjas">
    <div class="section-header">
      <h2>Franjas horarias disponibles para expansión</h2>
      <span class="section-count">UEA 1100037 · conflictos por día · catálogo 26-P</span>
    </div>
    <p class="section-desc">
      Para abrir una sección de Laboratorio de Reacciones en el horario de 13:00 se
      requiere que ningún grupo de la UEA 1100037 ocupe ese bloque en el mismo día.
      El análisis de horarios 26-P muestra que <strong>lunes y viernes</strong> son los
      únicos días sin grupos activos de esa UEA: son las franjas de expansión disponibles
      sin conflicto de aula para el catálogo actual.
    </p>
    <div class="read-callout">
      <strong>Metodología:</strong> se identificaron los grupos asignados de UEA 1100037
      día a día y se marcaron los días con al menos un grupo activo a las 13:00.
      Los días sin grupos en esa franja son candidatos para la nueva sección de Lab.
      La franja 13:00–16:00 en lunes o viernes no genera ningún traslape con los
      grupos existentes del catálogo 26-P.
    </div>
    <div class="chart-guide-grid animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Disponibilidad de franja 13:00 por día — UEA 1100037</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering sm" id="chart-franjas"><div class="chart-skeleton"></div></div>
      </div>
      <div class="howto-card">
        <h3>Interpretación del mapa de franjas</h3>
        <div class="howto-step">
          <div class="howto-icon" style="background:#2E7D32">✓</div>
          <div class="howto-text"><strong>Lunes y viernes</strong> no tienen grupos de UEA 1100037. La franja 13:00–16:00 está libre para una nueva sección de Lab. Reacciones (36 plazas).</div>
        </div>
        <div class="howto-step">
          <div class="howto-icon" style="background:#C82D23">✗</div>
          <div class="howto-text"><strong>Martes, miércoles y jueves</strong> tienen grupos activos de UEA 1100037 a las 13:00. Abrir Lab. Reacciones en esos días generaría traslape.</div>
        </div>
        <div class="howto-step">
          <div class="howto-icon">→</div>
          <div class="howto-text"><strong>Consecuencia:</strong> basta con 1 sección en la franja libre para elevar la capacidad de 864 a <strong>900 estudiantes</strong> — el techo simultáneo de Lab. e Intro. a la Física.</div>
        </div>
      </div>
    </div>
  </section>

  <!-- ── Sección 5: Demanda escalonada ── -->
  <section class="section animate-in" id="cap-demanda">
    <div class="section-header">
      <h2>Secciones mínimas requeridas por padrón</h2>
      <span class="section-count">Funciones escalonadas n<sub>L</sub>*(N) y n<sub>A</sub>*(N)</span>
    </div>
    <p class="section-desc">
      El número mínimo de secciones para atender N estudiantes sin traslape es
      n<sub>L</sub>*(N) = ⌈N/36⌉ para laboratorio y n<sub>A</sub>*(N) = ⌈N/45⌉
      para las UEAs académicas. La brecha entre ambas funciones delimita los intervalos
      donde el laboratorio es el cuello de botella y los rangos donde ambas UEAs deben
      abrirse simultáneamente. Los umbrales de coincidencia (N = 180, 360, 540, 720, 900)
      son los puntos críticos de planeación: en ellos n<sub>L</sub>* = n<sub>A</sub>*.
    </p>
    <div class="read-callout">
      <strong>Cómo leer la gráfica:</strong> los escalones de n<sub>L</sub>* (rojo) suben
      cada 36 estudiantes; los de n<sub>A</sub>* (azul) suben cada 45. Las líneas
      verticales punteadas marcan los umbrales donde ambas funciones coinciden. El punto
      N = 900 es el primer umbral alcanzable con la infraestructura de laboratorio actual
      más una sola sección adicional.
    </div>
    <div class="chart-card animate-in delay-1" style="margin:0">
      <div class="chart-card-header">
        <h3>Secciones mínimas requeridas: n<sub>L</sub>*(N) y n<sub>A</sub>*(N) para N ∈ [1, 980]</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering" id="chart-demanda"><div class="chart-skeleton"></div></div>
    </div>
  </section>

  <!-- ── Sección 6: Género ── -->
  <section class="section animate-in" id="cap-gen">
    <div class="section-header">
      <h2>Distribución por género</h2>
      <span class="section-count">488 hombres · 186 mujeres</span>
    </div>
    <p class="section-desc">
      El nuevo ingreso 26-P es 72.4% masculino y 27.6% femenino — una brecha de casi
      3 a 1. La distribución es consistente con la proporción histórica de la matrícula
      de ingeniería en México, aunque con variación entre licenciaturas.
    </p>
    <div class="read-callout">
      <strong>Contexto:</strong> la distribución de género afecta la asignación de grupos cuando
      existen restricciones de horario por género (e.g., grupos exclusivos para el turno
      matutino para mujeres). El sistema CAT toma en cuenta esta restricción al construir
      las asignaciones sin traslape.
    </div>
    <div class="chart-duo animate-in delay-1">
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Distribución de género · 674 estudiantes</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering sm" id="chart-gen-donut"><div class="chart-skeleton"></div></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-header">
          <h3>Distribución de género por licenciatura</h3>
          <span class="chart-type-tag">Interactivo</span>
        </div>
        <div class="chart-canvas rendering sm" id="chart-gen-plan"><div class="chart-skeleton"></div></div>
      </div>
    </div>
  </section>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Implicaciones para la planeación institucional</div>
    <p style="font-size:.84rem;color:var(--mid);line-height:1.6">
      El análisis de capacidad máxima complementa directamente la asignación CAT:
      muestra el margen operativo disponible antes de requerir nueva infraestructura.
      Con la configuración actual, la División puede absorber hasta 190 estudiantes adicionales
      sin costo marginal. Las trayectorias curriculares de esos estudiantes adicionales
      quedan modeladas en el Metaanálisis.
    </p>
    <div class="connect-links">
      <a class="connect-link" href="cat26p.html">← Regresar a CAT 26P</a>
      <a class="connect-link" href="metanalisis.html">📊 Ver Metaanálisis →</a>
    </div>
  </div>

</main>
"""

CAP_JS = """
(function(){
  const D = window.__CAP__;

  function occColor(pct) {
    if (pct > 100) return '#800026';
    if (pct >= 100) return '#C82D23';
    if (pct >= 80)  return '#B06A00';
    return '#2E7D32';
  }

  const uNames       = D.ueas.map(u => u.short);
  const enrolled     = D.ueas.map(u => u.enrolled);
  const remaining    = D.ueas.map(u => u.remaining);
  const capTotals    = D.ueas.map(u => u.cap_total);
  const isBottleneck = D.ueas.map(u => u.code === '1113085');

  renderChart('chart-cap-uea',
    () => [
      {
        type:'bar', orientation:'h', name:'Inscritos',
        y:uNames, x:enrolled,
        marker:{color:uNames.map((_,i)=>isBottleneck[i]?'#9E1E18':'#C82D23'),opacity:.9,
                line:{color:'rgba(0,0,0,.08)',width:.5}},
        hovertemplate:'<b>%{y}</b><br>Inscritos: %{x}<extra></extra>',
      },{
        type:'bar', orientation:'h', name:'Disponibles',
        y:uNames, x:remaining,
        marker:{color:uNames.map((_,i)=>isBottleneck[i]?'rgba(200,45,35,.25)':'rgba(180,180,180,.35)'),
                line:{color:uNames.map((_,i)=>isBottleneck[i]?'rgba(200,45,35,.5)':'rgba(150,150,150,.4)'),width:.5}},
        hovertemplate:'<b>%{y}</b><br>Disponibles: %{x} (cap=%{customdata})<extra></extra>',
        customdata:capTotals,
      }
    ],
    () => ({
      barmode:'stack',
      margin:{l:10,r:10,t:8,b:40},
      xaxis:{title:'Estudiantes',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',range:[0,1000]},
      yaxis:{tickfont:{...FONT,size:11},automargin:true},
      legend:{font:{size:10},orientation:'h',xanchor:'center',x:.5,y:1.07},
      annotations:[{
        x:D.maxCap, y:5.45, text:'Techo actual: '+D.maxCap,
        showarrow:true, arrowhead:2, arrowcolor:'#C82D23',
        font:{size:10,color:'#C82D23'}, xanchor:'left', yanchor:'middle',
      }]
    })
  );

  renderChart('chart-cap-lab',
    () => [{
      type:'bar', orientation:'h',
      y:D.labGroups, x:D.labPct,
      marker:{color:D.labPct.map(occColor), opacity:.88, line:{color:'rgba(0,0,0,.06)',width:.5}},
      hovertemplate:'<b>%{y}</b><br>Ocupación: %{x:.1f}%<br>Inscritos: %{customdata} / 36<extra></extra>',
      customdata:D.labEnrolled,
      text:D.labPct.map(p=>p>0?p.toFixed(0)+'%':''),
      textposition:'outside', textfont:{...FONT,size:10}, cliponaxis:false,
    }],
    () => ({
      margin:{l:10,r:40,t:8,b:40},
      xaxis:{title:'Ocupación (%)',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',
             range:[0,145], tickvals:[0,25,50,75,100,125],
             ticktext:['0%','25%','50%','75%','100%','125%']},
      yaxis:{tickfont:{...FONT,size:10},automargin:true},
      shapes:[{type:'line',x0:100,x1:100,y0:-.5,y1:D.labGroups.length-.5,
               line:{color:'#C82D23',width:1.5,dash:'dot'}}],
    })
  );

  renderChart('chart-cap-ref',
    () => [{
      type:'bar', orientation:'h',
      y:D.refGroups, x:D.refPct,
      marker:{color:D.refPct.map(occColor), opacity:.88, line:{color:'rgba(0,0,0,.06)',width:.5}},
      hovertemplate:'<b>%{y}</b><br>Ocupación: %{x:.1f}%<br>Inscritos: %{customdata} / 45<extra></extra>',
      customdata:D.refEnrolled,
      text:D.refPct.map(p=>p>0?p.toFixed(0)+'%':''),
      textposition:'outside', textfont:{...FONT,size:10}, cliponaxis:false,
    }],
    () => ({
      margin:{l:10,r:40,t:8,b:40},
      xaxis:{title:'Ocupación (%)',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',
             range:[0,155], tickvals:[0,25,50,75,100,125],
             ticktext:['0%','25%','50%','75%','100%','125%']},
      yaxis:{tickfont:{...FONT,size:10},automargin:true},
      shapes:[{type:'line',x0:100,x1:100,y0:-.5,y1:D.refGroups.length-.5,
               line:{color:'#C82D23',width:1.5,dash:'dot'}}],
    })
  );

  renderChart('chart-sens',
    () => {
      const sk      = D.sensitivity.map(s=>s.k);
      const smaxCap = D.sensitivity.map(s=>s.max_cap);
      const sLabCap = D.sensitivity.map(s=>s.lab_cap);
      const sColors = smaxCap.map(c => c < 900 ? '#C82D23' : '#2E7D32');
      return [
        {
          type:'bar', name:'Capacidad máxima de admisión',
          x:sk, y:smaxCap,
          marker:{color:sColors, opacity:.85, line:{color:'rgba(0,0,0,.08)',width:.5}},
          hovertemplate:'%{x} grupos extra Lab<br>Cap. máx: <b>%{y}</b> estudiantes<extra></extra>',
          text:smaxCap, textposition:'outside', textfont:{...FONT,size:11}, cliponaxis:false,
        },{
          type:'scatter', mode:'lines', name:'Capacidad Lab. Reacciones',
          x:sk, y:sLabCap,
          line:{color:'rgba(200,45,35,.5)',dash:'dot',width:1.5},
          hovertemplate:'Cap. Lab: %{y}<extra></extra>',
        }
      ];
    },
    () => ({
      margin:{l:10,r:12,t:12,b:52},
      xaxis:{title:'Grupos adicionales de Lab. Reacciones',titlefont:{size:10},tickfont:FONT,
             gridcolor:'#E4E2DF',dtick:1},
      yaxis:{title:'Capacidad máxima',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',
             range:[840,950]},
      shapes:[{type:'line',x0:-.5,x1:8.5,y0:900,y1:900,
               line:{color:'#00838F',width:1.5,dash:'dash'}}],
      annotations:[{x:4,y:904,text:'Techo Física catálogo (20 × 45 = 900)',
        showarrow:false,font:{size:10,color:'#00838F'},xanchor:'center'}],
      legend:{font:{size:10},orientation:'h',xanchor:'center',x:.5,y:1.08},
    })
  );

  // ── Franjas horarias (hardcoded from R analysis, resumen_ejecutivo_cupos §3) ──
  renderChart('chart-franjas',
    () => {
      var dias = ['Lunes','Martes','Miércoles','Jueves','Viernes'];
      var conflicto = [0, 1, 1, 1, 0];
      var libre    = [1, 0, 0, 0, 1];
      return [
        {
          type:'bar', name:'Conflicto UEA 1100037',
          x:dias, y:conflicto,
          marker:{color:conflicto.map(v=>v?'#C82D23':'rgba(0,0,0,0)'),
                  line:{color:conflicto.map(v=>v?'#9E1E18':'rgba(0,0,0,0)'),width:1}},
          hovertemplate:'%{x}<br>Conflicto: <b>%{customdata}</b><extra></extra>',
          customdata:conflicto.map(v=>v?'Grupos activos':'—'),
          text:conflicto.map(v=>v?'Conflicto':''),
          textposition:'inside', textfont:{...FONT,size:11,color:'#fff'},
        },{
          type:'bar', name:'Franja 13:00 disponible',
          x:dias, y:libre,
          marker:{color:libre.map(v=>v?'#2E7D32':'rgba(0,0,0,0)'),
                  line:{color:libre.map(v=>v?'#1B5E20':'rgba(0,0,0,0)'),width:1}},
          hovertemplate:'%{x}<br>%{customdata}<extra></extra>',
          customdata:libre.map(v=>v?'Franja libre 13:00–16:00':'—'),
          text:libre.map(v=>v?'Libre':''),
          textposition:'inside', textfont:{...FONT,size:11,color:'#fff'},
        }
      ];
    },
    () => ({
      barmode:'overlay',
      margin:{l:10,r:10,t:8,b:40},
      xaxis:{tickfont:{...FONT,size:12},gridcolor:'#E4E2DF'},
      yaxis:{visible:false,range:[0,1.6]},
      legend:{font:{size:10},orientation:'h',xanchor:'center',x:.5,y:1.08},
      annotations:[
        {x:'Lunes',y:1.1,text:'✓ Disponible',showarrow:false,font:{size:10,color:'#2E7D32'}},
        {x:'Viernes',y:1.1,text:'✓ Disponible',showarrow:false,font:{size:10,color:'#2E7D32'}},
        {x:'Martes',y:1.1,text:'✗ Conflicto',showarrow:false,font:{size:10,color:'#C82D23'}},
        {x:'Miércoles',y:1.1,text:'✗ Conflicto',showarrow:false,font:{size:10,color:'#C82D23'}},
        {x:'Jueves',y:1.1,text:'✗ Conflicto',showarrow:false,font:{size:10,color:'#C82D23'}},
      ],
    })
  );

  // ── Demanda escalonada: n_L*(N) y n_A*(N) ──
  renderChart('chart-demanda',
    () => {
      var Narr = [], nL = [], nA = [];
      for (var n = 1; n <= 980; n++) {
        Narr.push(n);
        nL.push(Math.ceil(n/36));
        nA.push(Math.ceil(n/45));
      }
      return [
        {
          type:'scatter', mode:'lines', name:'n<sub>L</sub>*(N) = ⌈N/36⌉ (Lab. Reacciones)',
          x:Narr, y:nL,
          line:{color:'#C82D23',width:2,shape:'hv'},
          hovertemplate:'N=%{x}<br>n<sub>L</sub>*=%{y} secciones de Lab<extra></extra>',
        },{
          type:'scatter', mode:'lines', name:'n<sub>A</sub>*(N) = ⌈N/45⌉ (Intro. Física)',
          x:Narr, y:nA,
          line:{color:'#1565C0',width:2,shape:'hv'},
          hovertemplate:'N=%{x}<br>n<sub>A</sub>*=%{y} secciones de Física<extra></extra>',
        },{
          type:'scatter', mode:'markers', name:'N actual (674)',
          x:[674], y:[Math.ceil(674/36)],
          marker:{color:'#B06A00',size:9,symbol:'diamond'},
          hovertemplate:'N=674 (26-P)<br>n<sub>L</sub>*=%{y}<extra></extra>',
        }
      ];
    },
    () => {
      var coinc = [180,360,540,720,900];
      return {
        margin:{l:10,r:10,t:8,b:50},
        xaxis:{title:'N estudiantes',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',dtick:90},
        yaxis:{title:'Secciones mínimas',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF',dtick:2},
        shapes: coinc.map(function(nc){
          return {type:'line',x0:nc,x1:nc,y0:0,y1:Math.ceil(nc/36)+.5,
                  line:{color:'rgba(0,131,143,.35)',width:1,dash:'dot'}};
        }),
        annotations: coinc.map(function(nc){
          return {x:nc,y:Math.ceil(nc/36)+.7,text:nc.toString(),
                  showarrow:false,font:{size:9,color:'#00838F'},xanchor:'center'};
        }).concat([
          {x:674,y:Math.ceil(674/36)+.9,text:'26-P',
           showarrow:false,font:{size:9,color:'#B06A00'},xanchor:'center'},
          {x:900,y:Math.ceil(900/36)+1.2,text:'Umbral 900',
           showarrow:false,font:{size:9,color:'#C82D23'},xanchor:'center'},
        ]),
        legend:{font:{size:10},orientation:'h',xanchor:'center',x:.5,y:1.08},
      };
    }
  );

  renderChart('chart-gen-donut',
    () => [{
      type:'pie', hole:.48,
      labels:['Masculino','Femenino'],
      values:[D.genM, D.genF],
      marker:{colors:['#1565C0','#C82D23'], line:{color:'#fff',width:2}},
      textinfo:'label+percent',
      hovertemplate:'<b>%{label}</b><br>%{value} estudiantes (%{percent})<extra></extra>',
      textfont:{...FONT,size:12},
    }],
    () => ({
      margin:{l:8,r:8,t:8,b:8}, showlegend:false,
      annotations:[{text:'<b>674</b><br><span style="font-size:10px;color:#888">total</span>',
        showarrow:false,font:{size:15,color:'#1A1A1A',family:'Helvetica Neue,Arial'}}]
    })
  );

  renderChart('chart-gen-plan',
    () => {
      const planNames = window.__CAT__.planNames;
      const planTotal = window.__CAT__.planTotal;
      const planM = planTotal.map(n => Math.round(n * D.genMpct / 100));
      const planF = planTotal.map(n => n - Math.round(n * D.genMpct / 100));
      return [
        {type:'bar',orientation:'h',name:'Masculino',
         y:planNames, x:planM,
         marker:{color:'#1565C0',opacity:.8},
         hovertemplate:'<b>%{y}</b><br>H: %{x}<extra></extra>'},
        {type:'bar',orientation:'h',name:'Femenino',
         y:planNames, x:planF,
         marker:{color:'#C82D23',opacity:.8},
         hovertemplate:'<b>%{y}</b><br>M: %{x}<extra></extra>'},
      ];
    },
    () => ({
      barmode:'stack',
      margin:{l:10,r:10,t:8,b:40},
      xaxis:{title:'Estudiantes',titlefont:{size:10},tickfont:FONT,gridcolor:'#E4E2DF'},
      yaxis:{tickfont:{...FONT,size:11},automargin:true},
      legend:{font:{size:10},orientation:'h',xanchor:'center',x:.5,y:1.07},
    })
  );
})();
"""

cap_html = page_shell(
    title='CAT 26P — Capacidad Máxima de Admisión',
    active='cat26p_capacidad.html',
    content=CAP_CONTENT,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="#cap-uea"><span class="bottom-nav-icon">📦</span>Capacidad</a>
      <a class="bottom-nav-item" href="#cap-grupos"><span class="bottom-nav-icon">🔲</span>Grupos</a>
      <a class="bottom-nav-item" href="#cap-sens"><span class="bottom-nav-icon">📐</span>Sensib.</a>
      <a class="bottom-nav-item" href="#cap-franjas"><span class="bottom-nav-icon">🗓</span>Franjas</a>
      <a class="bottom-nav-item" href="#cap-demanda"><span class="bottom-nav-icon">📈</span>Demanda</a>
      <a class="bottom-nav-item" href="#cap-gen"><span class="bottom-nav-icon">👥</span>Género</a>
    """,
    data_vars=f'window.__CAP__ = {DATA_CAP};\nwindow.__CAT__ = {DATA_CAT};',
    extra_js=CAP_JS,
)

# ══════════════════════════════════════════════════════════════════════════════
# CONCENTRACION.HTML  — Distribución docente: landing
# ══════════════════════════════════════════════════════════════════════════════
CONC_CONTENT = f"""
<style>{INDEX_CSS}
  .meta-nodes {{ display:grid; grid-template-columns:1fr 1fr; gap:.7rem }}
  .meta-node {{ background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.12);
               border-radius:12px; padding:1rem 1.1rem; cursor:pointer;
               transition:background var(--ease),border-color var(--ease);
               text-decoration:none; display:block }}
  .meta-node:hover {{ background:rgba(255,255,255,.09); border-color:rgba(200,45,35,.5) }}
  .meta-node-tag {{ font-size:.6rem; font-weight:700; letter-spacing:.12em;
                   text-transform:uppercase; color:var(--red); margin-bottom:.3rem }}
  .meta-node-title {{ font-size:.95rem; font-weight:700; color:#fff; margin-bottom:.25rem }}
  .meta-node-desc {{ font-size:.75rem; color:rgba(255,255,255,.4); line-height:1.45 }}
  .meta-node-sub {{ margin-top:.55rem; padding-top:.5rem;
                   border-top:1px solid rgba(255,255,255,.1);
                   font-size:.72rem; color:rgba(200,45,35,.8); font-weight:600 }}
  .meta-node-wide {{ grid-column:1/-1 }}
  @media(max-width:960px) {{ .meta-nodes {{ grid-template-columns:1fr }} }}
</style>
<main>
  <div class="landing-hero">
    <div class="landing-hero-inner">
      <div>
        <p class="landing-tag animate-in">UAM Azcapotzalco · División de Ciencias Básicas e Ingeniería</p>
        <h1 class="landing-h1 animate-in delay-1">
          Distribución
          <strong>docente</strong>
        </h1>
        <p class="landing-sub animate-in delay-2">
          A partir de 5 391 grupos registrados entre 2016 y 2025, el análisis cuantifica la concentración docente por UEA, calcula el riesgo estructural que surge cuando esa concentración coincide con posiciones terminales de las cadenas de seriación, e identifica los bloques curriculares donde la salida de un solo profesor podría bloquear el avance de cohortes enteras.
        </p>
        <div class="landing-bar animate-in delay-2"></div>
        <div class="landing-kpis animate-in delay-3">
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="809">0</div>
            <div class="lkpi-lbl">UEAs activas</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="791">0</div>
            <div class="lkpi-lbl">Profesores</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="61">0</div>
            <div class="lkpi-lbl">UEAs críticas</div>
          </div>
          <div class="lkpi">
            <div class="lkpi-val kpi-counter" data-target="11">0</div>
            <div class="lkpi-lbl">Vecindades</div>
          </div>
        </div>
      </div>
      <div class="meta-nodes animate-in delay-2">
        <a class="meta-node" href="conc_1.html">
          <div class="meta-node-tag">Infografía 1</div>
          <div class="meta-node-title">La concentración docente</div>
          <div class="meta-node-desc">El Índice de Herfindahl-Hirschman aplicado a la impartición: 44.3 % de las UEAs con concentración alta.</div>
        </a>
        <a class="meta-node" href="conc_2.html">
          <div class="meta-node-tag">Infografía 2</div>
          <div class="meta-node-title">El riesgo estructural</div>
          <div class="meta-node-desc">El Riesgo Concentrado combina cadenas de seriación con dependencia de un único docente.</div>
        </a>
        <a class="meta-node" href="conc_3.html">
          <div class="meta-node-tag">Infografía 3</div>
          <div class="meta-node-title">Redes y vecindades</div>
          <div class="meta-node-desc">11 vecindades curriculares controladas donde la fragilidad no es puntual sino sistémica.</div>
        </a>
        <a class="meta-node meta-node-wide" href="conc_4.html">
          <div class="meta-node-tag">Infografía 4</div>
          <div class="meta-node-title">El mapa de riesgo</div>
          <div class="meta-node-desc">Exposición estructural por plan de estudios y departamento: Ingeniería Civil, 30.6 % de UEAs críticas.</div>
        </a>
      </div>
    </div>
    <div class="scroll-hint">▾ desplázate para explorar</div>
  </div>
  <div style="padding:3rem 2rem;max-width:1440px;margin:0 auto" id="content">
    <div class="connect-box animate-in">
      <div class="connect-box-title">¿Cómo se conecta con los demás proyectos?</div>
      <p style="font-size:.84rem;color:var(--mid);line-height:1.6">La concentración docente explica parte de la variación en eficiencia terminal que el <strong>Análisis de eficiencia terminal</strong> documenta por plan: una UEA con un único impartidor al final de una cadena larga no solo retarda el egreso, lo bloquea. Las recomendaciones de este análisis son insumo directo para las decisiones de asignación que las <strong>Modificaciones 2026</strong> deben contemplar.</p>
      <div class="connect-links">
        <a class="connect-link" href="metanalisis.html">Ver Eficiencia terminal →</a>
        <a class="connect-link" href="modificaciones.html">Ver Modificaciones 2026 →</a>
      </div>
    </div>
  </div>
</main>
"""

conc_html = page_shell(
    title='Distribución Docente — DCBI',
    active='concentracion.html',
    content=CONC_CONTENT,
    data_vars='',
    extra_js='',
)

# ══════════════════════════════════════════════════════════════════════════════
# CONC_1.HTML  — Infografía 1: La concentración docente
# ══════════════════════════════════════════════════════════════════════════════
CONC1_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Distribución docente · Infografía 1 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>La concentración docente</strong></h1>
      <p class="hero-sub">El Índice de Herfindahl-Hirschman mide la dependencia de impartición por UEA: 187 unidades de enseñanza-aprendizaje tienen un único docente en toda la historia registrada.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">El <strong>Índice de Herfindahl-Hirschman (HHI)</strong> mide qué tan concentrada está la impartición de una UEA entre sus profesores. Para cada UEA se calcula la fracción de grupos históricos que ha impartido cada docente; esas fracciones se elevan al cuadrado y se suman. El resultado va de 0 —muchos profesores con participación igualitaria— a 1 —un único impartidor en toda la historia registrada. Este último caso no es anecdótico: representa una <strong>dependencia estructural</strong>. Si ese profesor deja de estar disponible, la institución no tiene otro docente con historial documentado en esa UEA. El índice fue diseñado originalmente para detectar monopolios de mercado; aquí se adapta para detectar monopolios de impartición.</p>
  </section>

  <section class="section animate-in" id="chart-section-hhi-cat">
    <div class="section-header">
      <h2>Distribución de UEAs por nivel de concentración docente</h2>
      <span class="section-count">809 UEAs activas · período 2016–2025</span>
    </div>
    <p class="section-desc">Las cuatro categorías van de monopolio total (HHI = 1) a distribución diversa (HHI &lt; 0.25). La barra más larga revela el estado habitual del sistema.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> la proporción de UEAs en las dos categorías de mayor concentración (HHI ≥ 0.5) frente al total, y la diferencia entre la categoría de monopolio exacto y la de concentración alta.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Distribución de UEAs por categoría de HHI (n = 809)</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-hhi-cat">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">La distribución divisional no es simétrica ni benigna. <strong>187 UEAs —el 23.1 % del total— tienen HHI = 1</strong>: un único profesor las ha impartido durante todos los trimestres del período 2016–2025, sin ninguna alternativa documentada. Si se amplía el umbral a concentración alta (HHI ≥ 0.5), el porcentaje alcanza el <strong>44.3 %</strong>. Solo el 30.5 % de las UEAs muestra una distribución diversa entre varios docentes. Esto no es una excepción: es el estado habitual del sistema. Una UEA con 40 o 50 grupos históricos y HHI = 1 significa que la misma persona ha impartido todos esos grupos a lo largo de nueve años; no hay transferencia de conocimiento docente, no hay protocolo de relevo, y el riesgo ante una baja no planificada es inmediato.</p>
  </section>

  <section class="section animate-in" id="chart-section-hhi-dept">
    <div class="section-header">
      <h2>Concentración docente mediana por departamento</h2>
      <span class="section-count">7 departamentos · HHI mediana + línea divisional</span>
    </div>
    <p class="section-desc">La línea discontinua marca la mediana divisional (0.39). Los departamentos a la derecha de esa línea presentan una concentración sistemáticamente más alta que el promedio de la División.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> qué departamentos superan la mediana divisional, la distancia de Ciencias Básicas al resto, y el contraste entre el volumen de UEAs de Materiales y su nivel de concentración.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>HHI mediana por departamento</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-hhi-dept">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">La concentración no se distribuye uniformemente entre departamentos: hay un patrón estructural, no una dispersión aleatoria. <strong>Ciencias Básicas</strong> muestra la distribución más sana (mediana HHI = 0.17), lo cual refleja que sus UEAs del Tronco General —compartidas por todas las licenciaturas— se imparten históricamente por equipos amplios de profesores; la demanda masiva obliga a la diversificación. En el extremo opuesto, <strong>Dirección de la División</strong> registra la mediana más alta (HHI = 0.50) y el mayor porcentaje de UEAs en régimen de exclusiva (30.3 %): un tercio de sus UEAs depende de un único docente. <strong>Materiales</strong> combina un volumen grande de UEAs con concentración moderada-alta (mediana 0.37, 17.8 % en exclusiva), lo que amerita atención particular dado el peso de ese departamento en los planes de ingeniería de especialidad.</p>
  </section>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <p style="font-size:.84rem;color:var(--mid);line-height:1.6">Saber que una UEA tiene HHI = 1 identifica una dependencia; saber en qué punto de la cadena de seriaciones se encuentra esa UEA determina si esa dependencia es un riesgo menor o un riesgo sistémico. La siguiente infografía introduce el <strong>Riesgo Concentrado</strong>, el indicador que combina ambas dimensiones.</p>
    <div class="connect-links">
      <a class="connect-link" href="concentracion.html">← Índice</a>
      <a class="connect-link" href="conc_2.html">Infografía 2: El riesgo estructural →</a>
    </div>
  </div>

</main>
"""

CONC1_JS = """
renderChart('chart-hhi-cat',
  function() {
    var labels   = ['Monopolio (HHI = 1)', 'Concentración alta (HHI ≥ 0.5)', 'Concentración moderada (HHI 0.25–0.5)', 'Diversa (HHI < 0.25)'];
    var values   = [187, 171, 204, 247];
    var pcts     = [23.1, 21.1, 25.2, 30.5];
    var colors   = ['#C82D23', '#FF6B35', '#FF9800', '#2E7D32'];
    var htext = labels.map(function(l, i) {
      return l + '<br>' + values[i] + ' UEAs (' + pcts[i] + '%)';
    });
    return [{
      type: 'bar',
      orientation: 'h',
      x: values,
      y: labels,
      marker: { color: colors },
      text: values.map(function(v, i) { return v + ' (' + pcts[i] + '%)'; }),
      textposition: 'outside',
      hovertemplate: '%{y}<br>UEAs: %{x}<extra></extra>',
      cliponaxis: false
    }];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'Distribución de UEAs por categoría de HHI (n = 809)', font: FONT, x: 0.5 },
      xaxis: { title: 'Número de UEAs', tickfont: FONT, range: [0, 310] },
      yaxis: { automargin: true, tickfont: FONT },
      margin: { l: 220, r: 60, t: 50, b: 50 },
      bargap: 0.3
    });
  }
);

renderChart('chart-hhi-dept',
  function() {
    var labels = ["Dirección", "Humanidades", "Materiales", "Electrónica", "Sistemas", "Energía", "Ciencias Básicas"];
    var hhi_vals = [0.5, 0.375, 0.3646, 0.2639, 0.2503, 0.2474, 0.166];
    var n_ueas   = [109, 5, 107, 83, 62, 138, 89];
    var pct1     = [22.9, 20.0, 14.0, 9.6, 6.5, 4.3, 4.5];
    var colors   = ["#C82D23", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#2E7D32", "#2E7D32"];
    return [
      {
        type: 'bar',
        orientation: 'h',
        x: hhi_vals,
        y: labels,
        marker: { color: colors, opacity: 0.85 },
        name: 'HHI mediana',
        text: hhi_vals.map(function(v) { return v.toFixed(3); }),
        textposition: 'outside',
        customdata: n_ueas.map(function(n, i) { return [n, pct1[i]]; }),
        hovertemplate: '<b>%{y}</b><br>HHI mediana: %{x:.3f}<br>UEAs: %{customdata[0]}<br>% monopolio: %{customdata[1]}%<extra></extra>'
      },
      {
        type: 'scatter',
        mode: 'lines',
        x: [0.39, 0.39],
        y: [labels[labels.length-1], labels[0]],
        line: { color: '#1A1A1A', width: 1.5, dash: 'dash' },
        name: 'Mediana divisional (0.39)',
        hoverinfo: 'skip'
      }
    ];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'HHI mediana por departamento', font: FONT, x: 0.5 },
      xaxis: { title: 'HHI mediana', range: [0, 0.65], tickfont: FONT },
      yaxis: { automargin: true, tickfont: FONT },
      margin: { l: 160, r: 60, t: 50, b: 50 },
      bargap: 0.3,
      legend: { x: 0.6, y: 0.05, font: FONT }
    });
  }
);
"""

conc1_html = page_shell(
    title='La concentración docente — DCBI',
    active='concentracion.html',
    content=CONC1_CONTENT,
    data_vars='',
    extra_js=CONC1_JS,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="concentracion.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="conc_2.html"><span class="bottom-nav-icon">&#9654;</span>Riesgo</a>
    """,
)

# ══════════════════════════════════════════════════════════════════════════════
# CONC_2.HTML  — Infografía 2: El riesgo estructural
# ══════════════════════════════════════════════════════════════════════════════
CONC2_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Distribución docente · Infografía 2 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>El riesgo estructural</strong></h1>
      <p class="hero-sub">Las seriaciones formales obligan al estudiante a aprobar una UEA antes de inscribirse en la siguiente: el Riesgo Concentrado (RC) cuantifica cuándo esa cadena y un único impartidor se combinan para crear fragilidad máxima.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">Las seriaciones formales obligan al estudiante a aprobar una UEA antes de inscribirse en la siguiente. Cuando varias UEAs están encadenadas, la probabilidad de que el estudiante llegue a la última es el producto de las tasas de aprobación de todos los eslabones anteriores: con cuatro UEAs consecutivas al 85 % de aprobación, solo la mitad de quienes comenzaron la cadena podrán cursar la cuarta. El <strong>Riesgo Concentrado (RC)</strong> multiplica ese filtrado acumulado de la cadena por el HHI de la UEA al final de ella. Una UEA con un único impartidor (HHI = 1) situada al término de una cadena de diez prerrequisitos alcanza RC = 1 —el nivel máximo—: prácticamente nadie llega, y quien llega no tiene alternativa docente. El RC distingue la concentración inocua de la concentración peligrosa.</p>
  </section>

  <section class="section animate-in" id="chart-section-rc-scatter">
    <div class="section-header">
      <h2>Profundidad de cadena vs. concentración docente</h2>
      <span class="section-count">809 UEAs · tres dimensiones simultáneas</span>
    </div>
    <p class="section-desc">Cada punto es una UEA. El eje horizontal indica la profundidad de la cadena de prerrequisitos; el vertical, su HHI; el tamaño de la burbuja, el filtrado acumulado de estudiantes antes de llegar a ella.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> la esquina superior derecha —cadena larga + HHI alto— concentra las 61 UEAs críticas. La franja densa al centro revela que el riesgo intermedio es el estado habitual, no la excepción.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Profundidad de cadena vs. concentración docente</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering lg" id="chart-rc-scatter">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">El diagrama posiciona cada UEA en tres dimensiones simultáneamente: profundidad de la cadena de prerrequisitos (eje horizontal), concentración docente HHI (eje vertical) y magnitud del filtrado acumulado (tamaño de la burbuja). Las UEAs de mayor riesgo aparecen en la esquina superior derecha con burbujas grandes: cadenas largas, impartición exclusiva y alta proporción de estudiantes ya filtrados antes de llegar. Lo que revela el gráfico no es solo la existencia de esos casos extremos —que son los 61 clasificados como críticos (RC ≥ percentil 90)— sino la continuidad del espacio: hay una franja densa de UEAs con RC intermedio, ninguna de ellas segura, que representa el grueso del riesgo estructural cotidiano de la División.</p>
  </section>

  <section class="section animate-in" id="chart-section-rc-top">
    <div class="section-header">
      <h2>Top 15 UEAs por Riesgo Concentrado (RC)</h2>
      <span class="section-count">RC = filtrado acumulado × HHI · HHI secundario como diamante</span>
    </div>
    <p class="section-desc">Las 15 UEAs con RC más alto. Los diamantes negros en el eje superior muestran el HHI de cada una; todas se ubican en HHI = 1, confirmando dependencia absoluta de un único docente.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> que HHI = 1 en todos los casos, y los departamentos de origen: el patrón no es aleatorio — Materiales, Sistemas, Energía y Electrónica dominan la lista.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Top 15 UEAs por Riesgo Concentrado (RC)</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-rc-top">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">Las 15 UEAs con RC más alto comparten el mismo perfil: HHI = 1 en todos los casos, e impacto de cascada próximo o igual a 1. Los nombres son reveladores — <strong>Puentes, Evaluación y Reparación de Estructuras, Modelos Hidráulicos, Bases de Datos Distribuidas, Estabilidad de los Sistemas de Potencia</strong> — son UEAs de especialidad terminal de sus respectivos planes, ubicadas exactamente donde el currículo más exige al estudiante y donde la institución más depende de un solo docente. El patrón no es coincidencia: las UEAs de especialidad tienden a tener menos profesores habilitados para impartirlas, y tienden a estar al final de las cadenas más largas. La concentración docente y la profundidad de seriación se refuerzan mutuamente en los eslabones más críticos del sistema.</p>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Perfil de riesgo máximo</div>
    <p>El caso de mayor riesgo individual identificado en el análisis corresponde a un profesor del área de Ingeniería Civil que concentra <strong>13 UEAs en régimen de exclusiva</strong> (HHI = 1 en cada una) dentro de una cadena de seriación de profundidad 8. La probabilidad de acceso al nodo final de esa cadena es prácticamente nula: los estudiantes que llegan a esas UEAs han superado ocho prerrequisitos consecutivos, y al llegar no encuentran alternativa docente. El RC de cada una de esas 13 UEAs es 1.0 —el máximo posible. Este perfil —alto número de monopolios, cadena profunda y RC máximo sostenido— define el escenario de mayor fragilidad institucional individual que pueden producir los datos disponibles. Una baja no planificada de este docente bloquearía simultáneamente a múltiples grupos de estudiantes que ya han invertido trimestres en completar los prerrequisitos.</p>
  </div>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <p style="font-size:.84rem;color:var(--mid);line-height:1.6">El análisis UEA por UEA identifica los casos extremos, pero no captura la estructura colectiva: hay bloques de UEAs temáticamente relacionadas —por seriación, corregistro o similitud de programas— donde el mismo pequeño grupo de profesores domina la impartición de todo el bloque. La siguiente infografía presenta esas <strong>vecindades curriculares controladas</strong> y revela por qué las zonas de mayor concentración agregada no son siempre las de mayor riesgo real.</p>
    <div class="connect-links">
      <a class="connect-link" href="conc_1.html">← Infografía 1: La concentración</a>
      <a class="connect-link" href="conc_3.html">Infografía 3: Redes y vecindades →</a>
    </div>
  </div>

</main>
"""

CONC2_JS = """
renderChart('chart-rc-top',
  function() {
    var names   = ["Puentes", "Taller de An\u00e1lisis y Dise\u00f1o de\u2026", "Evaluaci\u00f3n y Reparaci\u00f3n de Estructuras", "Estabilidad de los Sistemas de\u2026", "Taller de Sistemas Embebidos", "Modelado Num\u00e9rico de Cimentaciones", "Construcci\u00f3n de Obras de\u2026", "Bases de Datos Distribuidas", "Nanotecnolog\u00eda Ambiental", "Temas Selectos de Redes de\u2026", "Modelos Hidr\u00e1ulicos", "Dise\u00f1o Geot\u00e9cnico de Obras\u2026", "Captaci\u00f3n de Aguas Subterr\u00e1neas", "Fuentes de Alimentaci\u00f3n Conmutadas\u2026", "An\u00e1lisis Experimental de Esfuerzos"];
    var rc_vals = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.9997, 0.9989, 0.9975, 0.9973, 0.9958, 0.9945, 0.9927];
    var hhi_vals= [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0];
    var depts   = ["MATERIALES", "SISTEMAS", "MATERIALES", "ENERGIA", "ELECTRONICA", "MATERIALES", "MATERIALES", "SISTEMAS", "ENERGIA", "DIRECCION DE LA", "ENERGIA", "MATERIALES", "MATERIALES", "ELECTRONICA", "MATERIALES"];
    var fullnames= ["Puentes", "Taller de An\u00e1lisis y Dise\u00f1o de Algoritmos", "Evaluaci\u00f3n y Reparaci\u00f3n de Estructuras", "Estabilidad de los Sistemas de Potencia", "Taller de Sistemas Embebidos", "Modelado Num\u00e9rico de Cimentaciones", "Construcci\u00f3n de Obras de Infraestructura", "Bases de Datos Distribuidas", "Nanotecnolog\u00eda Ambiental", "Temas Selectos de Redes de Computadoras I", "Modelos Hidr\u00e1ulicos", "Dise\u00f1o Geot\u00e9cnico de Obras Subterr\u00e1neas", "Captaci\u00f3n de Aguas Subterr\u00e1neas", "Fuentes de Alimentaci\u00f3n Conmutadas OPT.*", "An\u00e1lisis Experimental de Esfuerzos"];
    var barcolors= ["#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23"];
    return [
      {
        type: 'bar',
        orientation: 'h',
        x: rc_vals,
        y: names,
        marker: { color: barcolors, opacity: 0.9 },
        name: 'RC',
        customdata: hhi_vals.map(function(h, i) { return [depts[i], h, fullnames[i]]; }),
        hovertemplate: '<b>%{customdata[2]}</b><br>Depto: %{customdata[0]}<br>RC: %{x:.4f}<br>HHI: %{customdata[1]:.3f}<extra></extra>',
        text: rc_vals.map(function(v) { return v.toFixed(3); }),
        textposition: 'outside',
        cliponaxis: false
      },
      {
        type: 'scatter',
        mode: 'markers',
        x: hhi_vals,
        y: names,
        marker: { size: 8, symbol: 'diamond', color: '#1A1A1A', line: { width: 1, color: '#fff' } },
        name: 'HHI',
        xaxis: 'x2',
        hovertemplate: 'HHI: %{x:.3f}<extra></extra>'
      }
    ];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'Top 15 UEAs por Riesgo Concentrado (RC)', font: FONT, x: 0.5 },
      xaxis:  { title: 'RC', range: [0, 1.15], tickfont: FONT, domain: [0, 0.85] },
      xaxis2: { title: 'HHI', range: [0, 1.1], overlaying: 'x', side: 'top', tickfont: FONT, showgrid: false },
      yaxis:  { automargin: true, tickfont: FONT },
      margin: { l: 280, r: 30, t: 60, b: 50 },
      bargap: 0.25,
      legend: { x: 0.72, y: 0.05, font: FONT },
      showlegend: true
    });
  }
);

renderChart('chart-rc-scatter',
  function() {
    var sc_x    = [7.0, 8.0, 7.0, 10.0, 7.0, 9.0, 7.0, 7.0, 6.0, 7.0, 7.0, 8.0, 5.0, 7.0, 5.0, 5.0, 5.0, 5.0, 5.0, 7.0, 7.0, 7.0, 9.0, 3.0, 9.0, 9.0, 2.0, 7.0, 6.0, 4.0, 8.0, 7.0, 7.0, 6.0, 7.0, 4.0, 10.0, 7.0, 7.0, 5.0, 7.0, 7.0, 9.0, 7.0, 8.0, 5.0, 9.0, 6.0, 9.0, 8.0, 6.0, 9.0, 6.0, 7.0, 6.0, 7.0, 8.0, 6.0, 7.0, 5.0, 8.0, 8.0, 7.0, 3.0, 7.0, 6.0, 5.0, 7.0, 11.0, 7.0, 7.0, 3.0, 4.0, 5.0, 2.0, 7.0, 6.0, 5.0, 7.0, 7.0, 6.0, 9.0, 2.0, 7.0, 8.0, 9.0, 10.0, 8.0, 7.0, 5.0, 6.0, 3.0, 10.0, 5.0, 5.0, 3.0, 7.0, 7.0, 3.0, 6.0, 6.0, 7.0, 3.0, 4.0, 5.0, 9.0, 3.0, 6.0, 6.0, 3.0, 6.0, 6.0, 1.0, 7.0, 8.0, 8.0, 6.0, 5.0, 2.0, 10.0, 8.0, 5.0, 8.0, 4.0, 4.0, 7.0, 10.0, 7.0, 8.0, 3.0, 6.0, 4.0, 6.0, 3.0, 10.0, 5.0, 6.0, 7.0, 5.0, 4.0, 5.0, 4.0, 5.0, 8.0, 4.0, 6.0, 6.0, 7.0, 6.0, 5.0, 0.0, 8.0, 5.0, 5.0, 3.0, 5.0, 4.0, 6.0, 4.0, 4.0, 6.0, 6.0, 2.0, 8.0, 4.0, 3.0, 6.0, 4.0, 4.0, 5.0, 6.0, 3.0, 5.0, 4.0, 5.0, 6.0, 10.0, 5.0, 9.0, 6.0, 2.0, 8.0, 5.0, 5.0, 6.0, 3.0, 5.0, 4.0, 7.0, 4.0, 5.0, 10.0, 7.0, 10.0, 5.0, 10.0, 8.0, 6.0, 5.0, 1.0, 8.0, 3.0, 2.0, 9.0, 6.0, 3.0, 5.0, 9.0, 5.0, 2.0, 6.0, 3.0, 4.0, 7.0, 5.0, 3.0, 4.0, 1.0, 6.0, 5.0, 3.0, 6.0, 4.0, 5.0, 4.0, 8.0, 6.0, 6.0, 3.0, 3.0, 5.0, 5.0, 4.0, 4.0, 4.0, 1.0, 3.0, 10.0, 6.0, 9.0, 7.0, 2.0, 5.0, 4.0, 3.0, 3.0, 5.0, 4.0, 4.0, 6.0, 7.0, 4.0, 5.0, 6.0, 4.0, 2.0, 3.0, 3.0, 9.0, 6.0, 10.0, 2.0, 5.0, 4.0, 4.0, 4.0, 4.0, 6.0, 1.0, 4.0, 1.0, 11.0, 2.0, 5.0, 2.0, 3.0, 8.0, 7.0, 3.0, 5.0, 5.0, 4.0, 3.0, 6.0, 2.0, 2.0, 4.0, 10.0, 1.0, 2.0, 4.0, 2.0, 1.0, 3.0, 2.0, 2.0, 8.0, 5.0, 11.0, 7.0, 7.0, 2.0, 0.0, 4.0, 5.0, 2.0, 2.0, 6.0, 2.0, 3.0, 2.0, 2.0, 6.0, 7.0, 1.0, 3.0, 8.0, 1.0, 1.0, 4.0, 3.0, 1.0, 2.0, 6.0, 9.0, 1.0, 1.0, 3.0, 4.0, 1.0, 2.0, 7.0, 1.0, 5.0, 2.0, 1.0, 1.0, 1.0, 3.0, 4.0, 1.0, 1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0];
    var sc_y    = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.8963, 1.0, 1.0, 0.7551, 0.7222, 1.0, 0.7562, 0.686, 0.7415, 0.68, 0.68, 0.68, 0.68, 0.6543, 0.66, 0.625, 0.625, 0.625, 0.8756, 0.6055, 0.5918, 0.5556, 0.5556, 0.5556, 0.5556, 0.54, 0.54, 0.5312, 0.52, 0.5148, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.4815, 0.4545, 0.4545, 0.5, 1.0, 0.4286, 0.4286, 0.4321, 0.4174, 0.4123, 0.4102, 0.4083, 1.0, 0.5, 0.3878, 0.42, 0.375, 0.375, 0.3735, 0.36, 0.3728, 0.3554, 0.3564, 0.8061, 0.34, 0.3554, 0.3333, 0.3333, 0.3333, 0.3333, 0.3333, 0.3223, 1.0, 0.32, 0.3195, 0.3333, 0.3439, 0.3113, 0.3244, 1.0, 0.3061, 0.3061, 0.3333, 0.7551, 0.3056, 0.3061, 0.2971, 0.68, 0.2959, 0.2893, 0.6095, 0.2991, 0.2891, 0.5024, 0.3125, 0.28, 0.7551, 0.2812, 0.2734, 0.6327, 0.2653, 0.2653, 0.2798, 0.264, 0.5, 0.5, 0.2633, 0.2593, 0.2593, 0.2593, 0.8347, 0.2562, 0.2601, 0.2535, 0.2922, 0.25, 0.25, 0.25, 0.25, 0.25, 0.2723, 0.2448, 0.2506, 0.24, 0.2384, 0.2359, 0.262, 0.2426, 0.2318, 0.2308, 0.2296, 0.3031, 0.225, 0.2361, 0.225, 0.5062, 0.2223, 0.2231, 0.218, 0.2316, 0.2396, 0.2164, 0.2136, 0.5207, 0.2109, 0.2126, 0.2672, 0.2092, 0.2175, 0.2091, 0.2131, 0.2083, 0.4375, 0.2044, 0.5005, 0.2, 0.1963, 0.1953, 0.1953, 0.1936, 0.1909, 0.3277, 0.1927, 0.1891, 0.1884, 0.1953, 0.3884, 0.1852, 0.1852, 0.1775, 0.1956, 0.2071, 0.1756, 0.1736, 0.3472, 0.1748, 0.1722, 0.1719, 0.1722, 0.1715, 0.3817, 0.1747, 0.1728, 0.3889, 1.0, 0.1667, 0.1732, 0.1667, 1.0, 0.1633, 0.2741, 0.7551, 0.3605, 0.169, 0.1541, 0.1527, 0.3208, 0.164, 0.5556, 0.177, 0.148, 0.4603, 0.5556, 0.1428, 0.1531, 0.1456, 0.1405, 0.1505, 0.1886, 0.4187, 0.4325, 0.1426, 0.1302, 0.1295, 0.1381, 0.1327, 0.3061, 0.1234, 0.12, 0.1426, 0.2207, 0.5102, 0.6818, 0.1378, 0.1183, 0.1435, 0.3605, 0.1221, 0.115, 0.1133, 0.358, 0.1178, 0.2383, 0.2786, 0.11, 0.1094, 0.1684, 0.1161, 0.1357, 0.2875, 0.4139, 0.4515, 0.6074, 0.1061, 0.1044, 0.1, 0.1086, 0.1457, 0.3935, 0.3875, 0.0922, 0.1823, 0.2734, 0.1901, 0.0967, 0.4603, 0.1667, 0.28, 0.3813, 0.0754, 0.0912, 0.2397, 0.1504, 0.1429, 0.08, 0.1055, 0.1037, 0.0878, 0.1893, 0.2966, 0.0715, 0.0801, 0.3636, 0.5938, 0.0706, 0.3424, 0.3308, 0.2287, 0.1249, 0.159, 0.12, 0.1205, 0.0884, 0.1626, 0.0556, 0.1444, 0.0638, 0.0774, 0.1773, 0.2769, 0.0488, 0.0727, 0.068, 0.2188, 0.1357, 0.3564, 0.0634, 0.2, 0.2653, 0.3314, 0.1204, 0.0433, 0.0808, 0.0481, 0.1427, 0.1446, 0.2959, 0.2, 0.0803, 0.0691, 0.2853, 0.0343, 0.0702, 0.0364, 0.0508, 0.1289, 0.0485, 0.0582, 0.0438, 0.1344, 0.0285, 0.0366, 0.0304, 0.142, 0.0874, 0.1953, 0.0813, 0.0859, 0.157, 0.072, 0.1985, 0.0517, 0.051, 0.2, 0.2, 0.2, 0.0483, 0.0458, 0.2018, 0.2022, 0.0349, 0.2, 0.1853, 0.1913, 0.135, 0.1368, 0.1387, 0.172, 0.1397, 0.1412, 0.1713, 0.1682, 0.1493, 0.166, 0.165, 0.1517, 0.1523, 0.1633, 0.1531, 0.1627, 0.1531, 0.1562, 0.1565, 0.1587, 0.1584, 0.1569, 0.1735, 0.1276, 0.0896, 0.1264, 0.1901, 0.0909, 0.0972, 0.0995, 0.1899, 0.1893, 0.1056, 0.1103, 0.1111, 0.112, 0.1122, 0.1125, 0.1837, 0.1145, 0.1806, 0.18, 0.18, 0.1796, 0.1156, 0.177, 0.1233, 0.2048, 0.68, 0.2082, 0.4132, 0.3889, 0.3889, 0.3889, 0.3889, 0.391, 1.0, 0.4097, 1.0, 1.0, 0.415, 1.0, 1.0, 1.0, 0.42, 1.0, 1.0, 1.0, 0.4321, 0.434, 0.438, 0.3878, 0.3827, 0.44, 0.3728, 0.3511, 1.0, 0.3554, 1.0, 1.0, 1.0, 0.3636, 0.3646, 0.365, 1.0, 0.38, 1.0, 0.375, 1.0, 0.375, 1.0, 0.375, 0.3776, 0.3787, 0.3795, 0.44, 0.44, 0.3462, 0.5, 0.68, 0.5, 0.551, 0.5, 0.5, 0.5, 0.5, 0.7551, 0.7551, 0.5, 0.7812, 0.7222, 0.5009, 0.551, 0.5062, 0.7222, 0.7222, 0.5247, 0.5306, 0.5312, 0.76, 0.68, 0.44, 0.8472, 0.4446, 0.4489, 0.4512, 0.4518, 1.0, 1.0, 0.4576, 0.4603, 0.4688, 0.4944, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.82, 0.5, 0.5, 0.7812, 0.3491, 1.0, 0.2083, 0.2639, 0.25, 0.25, 0.2521, 0.66, 0.6213, 0.6676, 0.6009, 0.2628, 0.5938, 0.5938, 0.625, 0.2647, 0.6686, 0.2653, 0.585, 0.2689, 0.2719, 0.5702, 0.2754, 0.2778, 0.6458, 0.25, 0.5692, 0.2282, 0.2092, 0.2132, 0.2154, 0.2188, 0.22, 0.2221, 0.2222, 0.2222, 0.2231, 0.2301, 0.25, 0.231, 0.2325, 0.2344, 0.54, 0.2416, 0.2428, 0.2449, 0.2494, 0.25, 0.2779, 0.28, 0.3438, 1.0, 0.3333, 0.3333, 0.3333, 1.0, 0.3333, 1.0, 1.0, 0.3333, 1.0, 1.0, 1.0, 1.0, 0.3333, 0.3333, 0.3347, 0.3366, 0.34, 1.0, 0.3403, 0.3426, 0.3333, 0.3233, 0.6689, 0.2933, 0.28, 0.28, 0.5556, 0.2812, 0.284, 0.284, 0.2851, 0.5556, 0.68, 0.5556, 1.0, 1.0, 0.5556, 1.0, 0.3056, 1.0, 1.0, 1.0, 1.0, 0.5556, 0.0243];
    var sc_sz   = [40.0, 40.0, 40.0, 40.0, 39.9999, 39.9995, 39.9992, 39.9991, 39.9874, 39.9579, 39.9011, 39.8913, 39.8326, 39.78, 39.7093, 39.5585, 39.1471, 37.8203, 37.4399, 36.2529, 36.2529, 39.9892, 34.4775, 31.3909, 40.0, 39.9079, 28.2424, 36.4907, 39.9991, 36.765, 40.0, 39.9995, 39.9989, 39.8253, 39.9961, 38.7957, 40.0, 39.9579, 39.78, 28.3909, 39.9579, 39.941, 40.0, 39.9935, 39.8913, 39.1471, 40.0, 39.9999, 39.9999, 39.8913, 39.8253, 39.9998, 39.9998, 39.9995, 39.9934, 39.9037, 39.8913, 39.706, 39.0713, 37.8956, 39.9992, 39.8913, 36.2529, 17.4029, 39.9037, 39.889, 39.5585, 39.9579, 40.0, 39.9997, 39.9995, 15.9804, 31.2826, 39.6933, 35.6828, 39.9011, 39.7838, 39.7493, 39.9011, 38.4356, 39.952, 39.7003, 17.2143, 39.9935, 37.6376, 40.0, 40.0, 39.9999, 39.9961, 39.4238, 39.9977, 12.8841, 39.991, 39.9956, 37.9657, 36.5137, 39.9902, 38.2817, 12.3912, 40.0, 39.7838, 36.2529, 15.9804, 39.2367, 39.1486, 40.0, 17.4029, 39.8854, 39.9967, 18.9392, 38.0406, 39.0261, 22.4015, 35.8674, 39.8913, 14.736, 39.3743, 39.9718, 17.2143, 40.0, 40.0, 37.8047, 39.9998, 21.1127, 21.1127, 39.9692, 39.9999, 39.9892, 39.8913, 12.3912, 39.871, 39.2367, 39.871, 34.3375, 40.0, 39.9976, 39.9308, 39.7856, 39.4506, 35.8284, 39.8065, 38.5994, 39.9998, 39.9998, 39.753, 35.5049, 38.2732, 39.9961, 39.871, 39.9098, 30.203, 39.9998, 37.8047, 39.439, 17.4029, 39.3862, 39.2367, 39.9855, 37.4656, 36.1087, 39.7021, 39.9967, 16.3701, 39.9998, 39.538, 31.3909, 39.9987, 38.4363, 39.9131, 39.1486, 39.9978, 18.9392, 39.9839, 16.252, 39.4897, 39.9938, 40.0, 39.8326, 40.0, 39.9874, 23.2727, 39.5559, 39.9229, 39.7093, 38.0406, 18.9392, 39.5829, 39.1457, 39.9955, 36.2797, 34.162, 39.9999, 39.9997, 19.9789, 39.5585, 39.9999, 39.9999, 39.889, 39.6741, 17.6708, 38.4094, 38.7453, 17.2143, 6.668, 39.9154, 38.2269, 39.6933, 6.532, 39.5585, 23.5667, 8.472, 17.4029, 36.765, 39.8531, 39.8326, 18.9392, 36.765, 10.736, 33.653, 39.8172, 12.3912, 10.256, 39.753, 36.9578, 38.7957, 39.9999, 37.3341, 29.4598, 12.8841, 12.3912, 37.5699, 39.9434, 39.4156, 36.1087, 37.4477, 15.8157, 39.0157, 40.0, 33.653, 21.6938, 9.352, 6.9198, 34.162, 39.4156, 32.4041, 12.8841, 37.5699, 39.6055, 39.4156, 12.472, 37.7017, 18.336, 15.656, 39.5797, 39.753, 25.4625, 36.9224, 31.3909, 14.816, 10.256, 9.4, 6.9198, 39.5278, 38.3563, 39.9131, 36.2797, 26.6376, 9.772, 9.6943, 39.9527, 19.1009, 12.052, 17.2143, 33.6905, 6.9198, 18.9392, 11.212, 8.064, 39.4581, 31.6947, 11.72, 18.5101, 18.9392, 33.653, 25.4625, 25.4625, 29.8681, 13.692, 8.6, 35.5086, 31.6947, 6.9198, 6.0, 34.3375, 6.9198, 6.9198, 9.672, 17.608, 13.776, 17.884, 17.32, 23.5667, 12.74, 36.1087, 13.872, 31.2704, 25.4625, 10.924, 6.9198, 38.3526, 25.4625, 27.1364, 7.944, 12.432, 6.0, 26.0056, 8.18, 6.144, 6.0, 12.764, 34.3375, 18.26, 28.8362, 9.172, 9.024, 6.0, 6.144, 15.128, 17.148, 6.0, 29.5021, 13.956, 24.7705, 17.608, 6.9198, 18.26, 13.964, 18.144, 6.0, 24.7189, 18.144, 19.596, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0];
    var sc_col  = ["#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#C82D23", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "#FF9800", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)", "rgba(100,100,100,0.35)"];
    var sc_rc   = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.9997, 0.9989, 0.9975, 0.9973, 0.9958, 0.9945, 0.9927, 0.989, 0.9787, 0.9455, 0.936, 0.9063, 0.9063, 0.8961, 0.8619, 0.7848, 0.7551, 0.7205, 0.7061, 0.6899, 0.686, 0.6815, 0.68, 0.68, 0.68, 0.677, 0.6542, 0.6401, 0.625, 0.6243, 0.6216, 0.6215, 0.6049, 0.5909, 0.5556, 0.5555, 0.5541, 0.5438, 0.54, 0.54, 0.5312, 0.5186, 0.5126, 0.5, 0.5, 0.5, 0.4999, 0.4988, 0.4986, 0.4963, 0.4884, 0.4562, 0.4545, 0.4533, 0.4532, 0.4351, 0.4276, 0.4274, 0.4273, 0.417, 0.4123, 0.4102, 0.4083, 0.3995, 0.391, 0.3848, 0.3747, 0.3741, 0.373, 0.3712, 0.3591, 0.3582, 0.355, 0.3537, 0.3469, 0.3399, 0.3344, 0.3333, 0.3333, 0.3333, 0.3333, 0.3285, 0.3223, 0.3221, 0.3199, 0.3195, 0.3163, 0.3139, 0.3112, 0.3105, 0.3098, 0.3061, 0.3044, 0.3021, 0.3017, 0.2998, 0.2996, 0.2971, 0.2958, 0.2951, 0.2893, 0.2886, 0.2844, 0.2821, 0.2814, 0.2802, 0.2792, 0.2782, 0.2768, 0.2732, 0.2723, 0.2653, 0.2653, 0.2644, 0.264, 0.2639, 0.2639, 0.2631, 0.2593, 0.2592, 0.2586, 0.2586, 0.2554, 0.2551, 0.2527, 0.2508, 0.25, 0.25, 0.2496, 0.2487, 0.2466, 0.2439, 0.2436, 0.2418, 0.24, 0.2384, 0.2344, 0.2326, 0.2321, 0.2318, 0.2301, 0.2291, 0.2289, 0.225, 0.2231, 0.2218, 0.2202, 0.2189, 0.2188, 0.2179, 0.2169, 0.2163, 0.2148, 0.2136, 0.2131, 0.2109, 0.2101, 0.2097, 0.2092, 0.209, 0.2086, 0.2086, 0.2083, 0.2071, 0.2043, 0.2034, 0.1974, 0.1963, 0.1953, 0.1945, 0.1936, 0.1908, 0.1907, 0.1906, 0.1887, 0.187, 0.1857, 0.1839, 0.1833, 0.1812, 0.1775, 0.1774, 0.1769, 0.1756, 0.1736, 0.1734, 0.1729, 0.1722, 0.1719, 0.1717, 0.1701, 0.1686, 0.1678, 0.1674, 0.1674, 0.1667, 0.1663, 0.1655, 0.1654, 0.1633, 0.1615, 0.1615, 0.1599, 0.1568, 0.1553, 0.1535, 0.1521, 0.1519, 0.1507, 0.1491, 0.1489, 0.1473, 0.1426, 0.1425, 0.1419, 0.1415, 0.1412, 0.1405, 0.1405, 0.1389, 0.1349, 0.134, 0.1339, 0.13, 0.1276, 0.1247, 0.1242, 0.121, 0.1204, 0.12, 0.12, 0.1197, 0.1193, 0.1179, 0.1177, 0.1166, 0.1162, 0.1161, 0.1147, 0.1139, 0.1116, 0.1116, 0.111, 0.1092, 0.109, 0.1088, 0.1087, 0.1072, 0.1072, 0.1065, 0.1065, 0.1061, 0.1061, 0.1051, 0.1048, 0.1001, 0.0998, 0.0985, 0.097, 0.0961, 0.0939, 0.0921, 0.0871, 0.0824, 0.0818, 0.0814, 0.0796, 0.0789, 0.0785, 0.0769, 0.0744, 0.0723, 0.0702, 0.0696, 0.0677, 0.0673, 0.0672, 0.066, 0.0656, 0.0648, 0.0638, 0.0635, 0.0635, 0.0629, 0.0616, 0.0606, 0.0592, 0.0572, 0.0553, 0.055, 0.0548, 0.0537, 0.0522, 0.0521, 0.0518, 0.0502, 0.0501, 0.0499, 0.0493, 0.0484, 0.0479, 0.0468, 0.0463, 0.0461, 0.0435, 0.0422, 0.0414, 0.0412, 0.0409, 0.0408, 0.0397, 0.0384, 0.0372, 0.0369, 0.0347, 0.0327, 0.0326, 0.031, 0.0307, 0.0304, 0.0296, 0.0296, 0.0253, 0.0245, 0.0225, 0.0224, 0.0223, 0.0221, 0.0203, 0.0199, 0.018, 0.0176, 0.0166, 0.0149, 0.0127, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0];
    var sc_imp  = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 0.9997, 0.9989, 0.9975, 0.9973, 0.9958, 0.9945, 0.9927, 0.989, 0.9787, 0.9455, 0.936, 0.9063, 0.9063, 0.9997, 0.8619, 0.7848, 1.0, 0.9977, 0.7061, 0.9123, 1.0, 0.9191, 1.0, 1.0, 1.0, 0.9956, 0.9999, 0.9699, 1.0, 0.9989, 0.9945, 0.7098, 0.9989, 0.9985, 1.0, 0.9998, 0.9973, 0.9787, 1.0, 1.0, 1.0, 0.9973, 0.9956, 1.0, 1.0, 1.0, 0.9998, 0.9976, 0.9973, 0.9927, 0.9768, 0.9474, 1.0, 0.9973, 0.9063, 0.4351, 0.9976, 0.9972, 0.989, 0.9989, 1.0, 1.0, 1.0, 0.3995, 0.7821, 0.9923, 0.8921, 0.9975, 0.9946, 0.9937, 0.9975, 0.9609, 0.9988, 0.9925, 0.4304, 0.9998, 0.9409, 1.0, 1.0, 1.0, 0.9999, 0.9856, 0.9999, 0.3221, 0.9998, 0.9999, 0.9491, 0.9128, 0.9998, 0.957, 0.3098, 1.0, 0.9946, 0.9063, 0.3995, 0.9809, 0.9787, 1.0, 0.4351, 0.9971, 0.9999, 0.4735, 0.951, 0.9757, 0.56, 0.8967, 0.9973, 0.3684, 0.9844, 0.9993, 0.4304, 1.0, 1.0, 0.9451, 1.0, 0.5278, 0.5278, 0.9992, 1.0, 0.9997, 0.9973, 0.3098, 0.9968, 0.9809, 0.9968, 0.8584, 1.0, 0.9999, 0.9983, 0.9946, 0.9863, 0.8957, 0.9952, 0.965, 1.0, 1.0, 0.9938, 0.8876, 0.9568, 0.9999, 0.9968, 0.9977, 0.7551, 1.0, 0.9451, 0.986, 0.4351, 0.9847, 0.9809, 0.9996, 0.9366, 0.9027, 0.9926, 0.9999, 0.4093, 1.0, 0.9884, 0.7848, 1.0, 0.9609, 0.9978, 0.9787, 0.9999, 0.4735, 0.9996, 0.4063, 0.9872, 0.9998, 1.0, 0.9958, 1.0, 0.9997, 0.5818, 0.9889, 0.9981, 0.9927, 0.951, 0.4735, 0.9896, 0.9786, 0.9999, 0.907, 0.854, 1.0, 1.0, 0.4995, 0.989, 1.0, 1.0, 0.9972, 0.9919, 0.4418, 0.9602, 0.9686, 0.4304, 0.1667, 0.9979, 0.9557, 0.9923, 0.1633, 0.989, 0.5892, 0.2118, 0.4351, 0.9191, 0.9963, 0.9958, 0.4735, 0.9191, 0.2684, 0.8413, 0.9954, 0.3098, 0.2564, 0.9938, 0.9239, 0.9699, 1.0, 0.9334, 0.7365, 0.3221, 0.3098, 0.9392, 0.9986, 0.9854, 0.9027, 0.9362, 0.3954, 0.9754, 1.0, 0.8413, 0.5423, 0.2338, 0.173, 0.854, 0.9854, 0.8101, 0.3221, 0.9392, 0.9901, 0.9854, 0.3118, 0.9425, 0.4584, 0.3914, 0.9895, 0.9938, 0.6366, 0.9231, 0.7848, 0.3704, 0.2564, 0.235, 0.173, 0.9882, 0.9589, 0.9978, 0.907, 0.6659, 0.2443, 0.2424, 0.9988, 0.4775, 0.3013, 0.4304, 0.8423, 0.173, 0.4735, 0.2803, 0.2016, 0.9865, 0.7924, 0.293, 0.4628, 0.4735, 0.8413, 0.6366, 0.6366, 0.7467, 0.3423, 0.215, 0.8877, 0.7924, 0.173, 0.1038, 0.8584, 0.173, 0.173, 0.2418, 0.4402, 0.3444, 0.4471, 0.433, 0.5892, 0.3185, 0.9027, 0.3468, 0.7818, 0.6366, 0.2731, 0.173, 0.9588, 0.6366, 0.6784, 0.1986, 0.3108, 0.1161, 0.6501, 0.2045, 0.1536, 0.1199, 0.3191, 0.8584, 0.4565, 0.7209, 0.2293, 0.2256, 0.1049, 0.1536, 0.3782, 0.4287, 0.1038, 0.7376, 0.3489, 0.6193, 0.4402, 0.173, 0.4565, 0.3491, 0.4536, 0.134, 0.618, 0.4536, 0.4899, 0.0891, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0];
    var sc_names= ["Puentes", "Taller de An\u00e1lisis y Dise\u00f1o de Algoritmos", "Evaluaci\u00f3n y Reparaci\u00f3n de Estructuras", "Estabilidad de los Sistemas de Potencia", "Taller de Sistemas Embebidos", "Modelado Num\u00e9rico de Cimentaciones", "Construcci\u00f3n de Obras de Infraestructura", "Bases de Datos Distribuidas", "Nanotecnolog\u00eda Ambiental", "Temas Selectos de Redes de Computadoras I", "Modelos Hidr\u00e1ulicos", "Dise\u00f1o Geot\u00e9cnico de Obras Subterr\u00e1neas", "Captaci\u00f3n de Aguas Subterr\u00e1neas", "Fuentes de Alimentaci\u00f3n Conmutadas OPT.*", "An\u00e1lisis Experimental de Esfuerzos", "Dise\u00f1o y Construcci\u00f3n de Obras Provisionales", "Materiales Compuestos en la Construcci\u00f3n", "Introducci\u00f3n a las Estructuras I", "Temas Selectos de Matem\u00e1ticas Discretas", "Calidad de Software", "Temas Selectos de Sistemas de Informaci\u00f3n", "Taller de Muestreo y Control de Contaminantes\u2026", "Temas Selectos de Telecomunicaciones 11 OPT.*", "Temas Selectos de Lenguajes de Programaci\u00f3n", "M\u00e9todo de Elemento Finito Aplicado a M\u00e1quinas\u2026", "V\u00edas F\u00e9rreas", "Temas Selectos de Topograf\u00eda", "Instalaciones El\u00e9ctricas en Baja Tensi\u00f3n", "Maquinaria y Equipo de Construcci\u00f3n", "Series, Transformadas y Ecuaciones Diferenciales", "Seguridad en los Sistemas de Informaci\u00f3n", "Hidr\u00e1ulica Fluvial", "Edificaci\u00f3n", "Ingenier\u00eda de Costos de Construcci\u00f3n", "Complejidad Computacional", "Programaci\u00f3n de Sistemas", "Tratamiento de Residuos L\u00edquidos Industriales", "Integraci\u00f3n de Redes de Voz y Datos", "Impulsores para Motores de CD y CA OPT.*", "Energ\u00eda Solar Fotovoltaica", "Redes Inal\u00e1mbricas", "Teletr\u00e1fico", "Sistemas de Generaci\u00f3n E\u00f3lica", "Administraci\u00f3n de Proyectos de Software", "Aeropuertos", "Tecnolog\u00eda del Concreto", "Dise\u00f1o de M\u00e1quinas El\u00e9ctricas", "Programaci\u00f3n Orientada a Servicios", "Gesti\u00f3n Integral de Residuos S\u00f3lidos Urbanos", "Geotecnia Aplicada a la Construcci\u00f3n", "Administraci\u00f3n de la Construcci\u00f3n", "Taller de Restauraci\u00f3n de Suelos", "Concreto Presforzado", "Hidr\u00e1ulica Mar\u00edtima", "Temas Selectos de Ingenier\u00eda Sustentable", "Temas Selectos de Ingenier\u00eda en Construcci\u00f3n", "Asfaltos", "Aplicaci\u00f3n de Motores El\u00e9ctricos OPT.*", "Limnolog\u00eda y Saneamiento de Corrientes", "Ingenier\u00eda Financiera", "Ingenier\u00eda de Cimentaciones", "Dise\u00f1o Geot\u00e9cnico de V\u00edas Terrestres", "Taller de Productos de Software", "Trabajo de Investigaci\u00f3n en Ingenier\u00eda en\u2026", "Legislaci\u00f3n en la Construcci\u00f3n", "Se\u00f1ales Aleatorias", "Dise\u00f1o de Estructuras de Mamposter\u00eda", "Seguridad en Redes", "Control Digital", "Procesos Fisicoqu\u00edmicos en Ingenier\u00eda Ambiental", "Dise\u00f1o y Construcci\u00f3n de Carreteras", "Proyecto de Integraci\u00f3n en Ingenier\u00eda El\u00e9ctrica II", "Temas Selectos de Gesti\u00f3n Ambiental", "Taller de An\u00e1lisis Estructural", "Geolog\u00eda", "Infraestructura para el Tratamiento de Aguas", "Geolog\u00eda Aplicada a la Ingenier\u00eda Civil", "Costos y Presupuestos de Obra", "Temas Selectos de Ingenier\u00eda Hidr\u00e1ulica", "Ingenier\u00eda de Biorreactores", "Dise\u00f1o y Administraci\u00f3n de Redes de Computadoras", "Sistemas de Potencia en Estado Estable", "Legislaci\u00f3n y Gesti\u00f3n Ambiental", "Arquitectura e Integraci\u00f3n de Aplicaciones\u2026", "Ingenier\u00eda de Iluminaci\u00f3n", "Control de M\u00e1quinas El\u00e9ctricas", "Plantas Potabilizadoras", "Temas Selectos de Sistemas Embebidos", "Combinatoria", "Temas Selectos de Inteligencia Artificial", "Dise\u00f1o de Estructuras de Acero", "Introducci\u00f3n al Trabajo de Investigaci\u00f3n en\u2026", "Protecciones El\u00e9ctricas", "Teor\u00eda Electromagn\u00e9tica", "Instrumentaci\u00f3n Virtual", "Ingenier\u00eda de Materiales en Construcci\u00f3n", "Sistemas Digitales con Microcontroladores", "Microcontroladores", "Trabajo de Investigaci\u00f3n en Ingenier\u00eda Ambiental", "Proyecto de Edificios", "Flujo de Agua en Suelos", "Patrones de Dise\u00f1o de Software", "Proyecto de Integraci\u00f3n en Ingenier\u00eda El\u00e9ctrica I", "Inteligencia Computacional", "Arquitecturas Paralelas de Computadoras", "T\u00e9cnicas de Altas Tensiones", "Introducci\u00f3n al Trabajo de Investigaci\u00f3n en\u2026", "Planeaci\u00f3n, Programaci\u00f3n y Control de Obras", "Propiedades El\u00e9ctricas y Magn\u00e9ticas de los\u2026", "Sistemas de Gesti\u00f3n Ambiental", "Microbiolog\u00eda Aplicada", "Taller de An\u00e1lisis de la Calidad del Agua", "Topograf\u00eda", "Sensores, Transductores y Detectores", "Mec\u00e1nica de Rocas", "Laboratorio de Ingenier\u00eda de Biorreactores", "Sistemas Distribuidos", "Combusti\u00f3n", "An\u00e1lisis de Ciclo de Vida", "Plantas de Tratamiento de Aguas Residuales", "Presas", "Sistemas de Gesti\u00f3n de la Calidad", "Manejo y Tratamiento de Residuos Industriales y\u2026", "Planeaci\u00f3n Ambiental", "An\u00e1lisis de Riesgo Ambiental", "Electr\u00f3nica de Potencia", "Aprovechamiento y Tratamiento de Residuos\u2026", "Procesos de Control de la Contaminaci\u00f3n\u2026", "Pavimentos", "Introducci\u00f3n al Trabajo de Investigaci\u00f3n en\u2026", "Sistemas de Alcantarillado", "Almacenamiento y Estructuras de Archivos", "Hidr\u00e1ulica de Canales", "C\u00e1lculo Vectorial y sus Aplicaciones", "Reuso del Agua", "Simulaci\u00f3n Aplicada a Ingenier\u00eda Mec\u00e1nica OPT.*", "Taller de Desarrollo de Aplicaciones Web", "Taller de Mecatr\u00f3nica", "Sistemas Ambientales", "Lenguajes y Aut\u00f3matas", "Instalaciones Industriales", "Gr\u00e1ficas por Computadora", "An\u00e1lisis Estructural Matricial", "Contaminaci\u00f3n y Restauraci\u00f3n de Suelos", "Din\u00e1mica y Vibraciones", "An\u00e1lisis y Dise\u00f1o de Sistemas de Informaci\u00f3n", "Efecto Invernadero y Cambio Clim\u00e1tico", "Criptograf\u00eda", "Abastecimiento de Agua", "Bases de Datos", "Comunicaciones Digitales", "M\u00e1quinas de CD y de Inducci\u00f3n", "Desarrollo Organizacional", "Circuitos El\u00e9ctricos II", "Proyecto de Integraci\u00f3n en Ingenier\u00eda en\u2026", "Estad\u00edstica Aplicada I", "Inteligencia Artificial", "Hidrolog\u00eda", "Control de Calidad y Confiabilidad", "An\u00e1lisis y Dise\u00f1o de Experimentos en Ingenier\u00eda", "Teor\u00eda de Control", "Aplicaciones del Electromagnetismo", "Metalurgia Computacional", "M\u00e1quinas S\u00edncronas en Estado Din\u00e1mico", "Fundamentos Electr\u00f3nica de de Manufactura\u2026", "Programaci\u00f3n Visual Orientada a Eventos", "Dise\u00f1o de Sistemas Energ\u00e9ticos OPT.*", "Variable Compleja", "Radiaci\u00f3n y Propagaci\u00f3n ....", "Sistemas Operativos", "Dise\u00f1o de Estructuras de Concreto", "Legislaci\u00f3n Aplicada a los Recursos Naturales", "Fen\u00f3menos de Transporte", "Laboratorio de Mec\u00e1nica de S\u00f3lidos", "Elemento Finito", "An\u00e1lisis y Dise\u00f1o de Algoritmos", "Subestaciones El\u00e9ctricas", "Mec\u00e1nica de Fluidos Avanzada", "Taller de Procesos de Tratamiento de Agua", "Control de la Contaminaci\u00f3n Atmosf\u00e9rica", "Laboratorio de Procesos", "Modelado de Sistemas El\u00e9ctricos de Potencia", "Compiladores", "Fundamentos de Geotecnia", "Toxicolog\u00eda Ambiental", "An\u00e1lisis de Casos en Auditor\u00eda Ambiental", "Din\u00e1mica de M\u00e1quinas", "Dise\u00f1o de Mecanismos", "Operaciones Unitarias en Ingenier\u00eda Ambiental", "Equilibrio Qu\u00edmico", "Procesos de Conversi\u00f3n de Energ\u00eda", "Taller de Residuos S\u00f3lidos Urbanos y Suelos", "Transformadores y M\u00e1quinas S\u00edncronas", "Taller de Fuentes Alternas de Energ\u00eda", "Elementos de Acero", "Dise\u00f1o de Sistemas Electr\u00f3nicos", "Procesos Biol\u00f3gicos en Ingenier\u00eda Ambiental", "Procesamiento Digital de Se\u00f1ales", "Automatizaci\u00f3n Industrial", "Procesos de Separaci\u00f3n 1", "Circuitos Electr\u00f3nicos II", "Mecanismos", "Inventarios de Emisiones Atmosf\u00e9ricas", "Laboratorio de Pavimentos", "Aprovechamientos Hidr\u00e1ulicos", "Microprocesadores", "Ingenier\u00eda S\u00edsmica", "Laboratorio de Asfaltos", "Elementos de Concreto", "Manufactura Asistida por Computadora", "Laboratorio de Combusti\u00f3n", "Proyecto de Integraci\u00f3n en Ingenier\u00eda en\u2026", "Transformada de Laplace y An\u00e1lisis de Fourier", "Ingenier\u00eda Geot\u00e9cnica", "Hidr\u00e1ulica de Tuber\u00edas", "Seguridad e Higiene Industrial", "Matem\u00e1ticas Aplicadas para Ingenier\u00eda", "Energ\u00eda Solar Fotovoltaica", "Circuitos El\u00e9ctricos de Corriente Alterna", "An\u00e1lisis de Se\u00f1ales", "Proyecto de Integraci\u00f3n en Ingenier\u00eda Ambiental I", "Laboratorio de An\u00e1lisis Experimental de Esfuerzos", "Mec\u00e1nica de Fluidos", "Laboratorio de Termofluidos 1", "Arquitectura de Computadoras", "Centrales El\u00e9ctricas", "Laboratorio de Termofluidos 11 OPT.*", "Laboratorio de Circuitos El\u00e9ctricos II", "Proyecto de Integraci\u00f3n en Ingenier\u00eda Civil II", "Proyecto de Integraci\u00f3n en Ingenier\u00eda Ambiental II", "Fundamentos de Qu\u00edmica Org\u00e1nica y Bioqu\u00edmica", "Fundamentos de Redes de Computadoras", "An\u00e1lisis Estructural", "Metrolog\u00eda para Manufactura", "An\u00e1lisis de Decisiones II", "Procesos Met\u00e1licos de Formado de los Materiales", "Mec\u00e1nica de S\u00f3lidos I", "Redes de Distribuci\u00f3n", "Circuitos Electr\u00f3nicos I", "Comunicaciones Anal\u00f3gicas", "Laboratorio de Hidr\u00e1ulica de Canales", "Seminario de Integraci\u00f3n en Ingenier\u00eda El\u00e9ctrica", "Recursos Energ\u00e9ticos", "Dise\u00f1o de Elementos de M\u00e1quinas", "Matem\u00e1ticas Discretas", "Proyecto de Integraci\u00f3n en Ingenier\u00eda Civil I", "Qu\u00edmica Ambiental", "Microelectr\u00f3nica", "Dise\u00f1o Estructural", "Laboratorio de Energ\u00eda Solar", "Taller de Termofluidos", "Laboratorio de Concreto", "Laboratorio de Din\u00e1mica y Vibraciones", "Investigaci\u00f3n de Operaciones II", "Transferencia de Momento", "L\u00f3gica", "C\u00e1lculo de Varias Variables", "Laboratorio de Programaci\u00f3n Orientada a Objetos", "Laboratorio de Circuitos Electr\u00f3nicos II", "Laboratorio de Geotecnia", "Programaci\u00f3n Aplicada a Sistemas El\u00e9ctricos de\u2026", "Seminario de Integraci\u00f3n en Ingenier\u00eda Mec\u00e1nica", "Transferencia de Calor", "Din\u00e1mica Aplicada", "Inducci\u00f3n y Ondas Electromagn\u00e9ticas", "Ingenier\u00eda de los Materiales OPT.*", "Laboratorio de Sistemas Digitales con\u2026", "Laboratorio de Electromagnetismo", "Fundici\u00f3n de los Materiales Met\u00e1licos", "Mec\u00e1nica de S\u00f3lidos II", "Laboratorio de Operaciones Unitarias", "Laboratorio de Protecciones El\u00e9ctricas", "Aplicaciones de Ecolog\u00eda Industrial", "Investigaci\u00f3n de Operaciones I", "Seminario de Integraci\u00f3n en Ingenier\u00eda Ambiental", "Sistemas de Informaci\u00f3n Geogr\u00e1fica", "Laboratorio de Electr\u00f3nica de Potencia", "Laboratorio de Microbiolog\u00eda Aplicada", "Laboratorio del Cuerpo R\u00edgido y Oscilaciones", "Dibujo Mec\u00e1nico Asistido por Computadora", "Laboratorio de Arquitectura de Computadoras", "Laboratorio de Microprocesadores", "Evaluaci\u00f3n de Impacto Ambiental", "Introducci\u00f3n a la Electr\u00f3nica", "Algoritmos y Estructuras de Datos", "Programaci\u00f3n Orientada a Objetos", "Termodin\u00e1mica Aplicada", "Laboratorio de Comunicaciones Anal\u00f3gicas", "Administraci\u00f3n de la Producci\u00f3n", "Din\u00e1mica del Cuerpo R\u00edgido", "M\u00e9todos Num\u00e9ricos en Ingenier\u00eda", "Seminario de Integraci\u00f3n en Ingenier\u00eda Metal\u00fargica", "Proyecto Mec\u00e1nico de Montajes OPT.*", "Probabilidad y Estad\u00edstica", "Seminario de Integraci\u00f3n en Ingenier\u00eda Civil", "Seminario de Integraci\u00f3n en Ingenier\u00eda en\u2026", "Laboratorio de Transformadores y M\u00e1quinas\u2026", "Introducci\u00f3n a las Comunicaciones", "Laboratorio de Dise\u00f1o de Sistemas Electr\u00f3nicos", "Laboratorio de Circuitos El\u00e9ctricos de\u2026", "Laboratorio de Circuitos Electr\u00f3nicos I", "Procesos de Manufactura 1 OPT.*", "Laboratorio de Comunicaciones Digitales", "Mediciones en Ingenier\u00eda", "Laboratorio de Qu\u00edmica Anal\u00edtica", "Laboratorio de Movimiento de una Part\u00edcula", "Dise\u00f1o L\u00f3gico", "Laboratorio de Introducci\u00f3n a las Comunicaciones", "Seminario de Integraci\u00f3n en Ingenier\u00eda Industrial", "Introducci\u00f3n a la Electrost\u00e1tica y Magnetost\u00e1tica", "Introducci\u00f3n al \u00c1lgebra Lineal", "Estructuras Isost\u00e1ticas", "Taller de Instalaciones Industriales", "Laboratorio de Control", "Qu\u00edmica Org\u00e1nica 111 OPT.*", "Termodin\u00e1mica", "Laboratorio de Sistemas Digitales con\u2026", "Estudio del M\u00e9todo del Trabajo", "Dise\u00f1o de Instalaciones y Manejo de Materiales", "Laboratorio de Mecanismos", "Ecuaciones Diferenciales Ordinarias", "Est\u00e1tica", "Laboratorio de Estructura y Propiedades de los\u2026", "Laboratorio de Hidr\u00e1ulica de Tuber\u00edas", "Laboratorio de M\u00e1quinas de CD y de Inducci\u00f3n", "Flujo de Energ\u00eda y Balance de Materia en\u2026", "Planeaci\u00f3n Estrat\u00e9gica", "Laboratorio de Dise\u00f1o L\u00f3gico", "Circuitos El\u00e9ctricos I", "Control Num\u00e9rico Computarizado", "C\u00e1lculo Integral", "Laboratorio de Introducci\u00f3n a la Electr\u00f3nica", "Cinem\u00e1tica y Din\u00e1mica de Part\u00edculas", "Laboratorio de Circuitos El\u00e9ctricos I", "Seminario de Integraci\u00f3n en Ingenier\u00eda Electr\u00f3nica", "Programaci\u00f3n Estructurada", "Laboratorio de Ingenier\u00eda de los Materiales OPT.*", "Estructura y Propiedades de los Materiales en\u2026", "Taller de Procesos de Manufactura 1 OPT.*", "Complementos de Matem\u00e1ticas", "Laboratorio de Reacciones Qu\u00edmicas", "C\u00e1lculo Diferencial", "Taller de Planeaci\u00f3n y Ejecuci\u00f3n de Proyectos", "TOTAL Propiedades DE CR\u00c9DITOS Termodin\u00e1micas\u2026", "Mantenimiento Industrial OPT.*", "Evaluaci\u00f3n Ambiental de Tecnolog\u00edas", "Temas Selectos de Ingenier\u00eda F\u00edsica 11", "Qu\u00edmica Org\u00e1nica 11", "Retos del Desarrollo Nacional", "Laboratorio de Cin\u00e9tica y Cat\u00e1lisis", "Introducci\u00f3n a la Ingenier\u00eda", "Introducci\u00f3n al Desarrollo Sustentable", "Gesti\u00f3n Industrial Eficiente de Procesos\u2026", "Instrumentaci\u00f3n Industrial", "F\u00edsica del Estado S\u00f3lido OPT.*", "Taller de Expresi\u00f3n Oral y Escrita", "Introducci\u00f3n a la F\u00edsica", "Metalurgia de las Aleaciones no Ferrosas", "Balance de Energ\u00eda", "Estructura At\u00f3mica y Enlace Qu\u00edmico", "Temas Selectos de Ingenier\u00eda Civil", "Laboratorio de F\u00edsica At\u00f3mica y Molecular", "Mec\u00e1nica Estad\u00edstica", "Fisicoqu\u00edmica de los Materiales", "Laboratorio de Qu\u00edmica Inorg\u00e1nica 1", "Procesos F\u00edsicos Industriales", "Laboratorio de Qu\u00edmica Org\u00e1nica 11", "Optativa T\u00e9cnica de Movilidad V", "Laboratorio de Electricidad y Magnetismo", "Log\u00edstica", "Laboratorio Protecci\u00f3n de de los Oxidaci\u00f3n,\u2026", "Administraci\u00f3n de Proyectos", "Funciones Especiales", "Qu\u00edmica Inorg\u00e1nica 1", "Econom\u00eda Mundial", "Laboratorio Ferrosas de Metalurgia de las\u2026", "\u00d3ptica", "Cristales y Dislocaciones OPT.*", "Planeaci\u00f3n Estrat\u00e9gica de Infraestructura en\u2026", "Optativa T\u00e9cnica de Movilidad VI", "Temas Selectos de Ingenier\u00eda El\u00e9ctrica I", "Introducci\u00f3n a la Automatizaci\u00f3n Industrial", "Optativa T\u00e9cnica de Movilidad III", "Reactores Heterog\u00e9neos", "Reactores Homog\u00e9neos", "Turbomaquinaria", "Electromagnetismo", "Laboratorio de Qu\u00edmica Inorg\u00e1nica 11", "F\u00edsica Moderna", "Cin\u00e9tica y Cat\u00e1lisis", "Organizaci\u00f3n Industrial", "Prevenci\u00f3n y Minimizaci\u00f3n de la Contaminaci\u00f3n\u2026", "Dibujo Asistido con Computadora en Ingenier\u00eda\u2026", "Qu\u00edmica Org\u00e1nica 1", "Laboratorio que apruebe Tecnolog\u00eda de\u2026", "An\u00e1lisis Vectorial", "Taller de Procesos de Manufactura 11", "Caracterizaci\u00f3n de los Materiales", "Procesos de Manufactura 11", "Temas Selectos de Ingenier\u00eda F\u00edsica 111", "Laboratorio Ingenier\u00eda Industrial de Sistemas\u2026", "Qu\u00edmica Inorg\u00e1nica 11", "An\u00e1lisis de Decisiones I", "Retos Ambientales", "Optativa T\u00e9cnica de Movilidad I", "Energ\u00eda E\u00f3lica Aplicada", "TOTAL Metalurgia DE se del CR\u00c9DITOS realizar\u00e1\u2026", "Laboratorio de Metrolog\u00eda para Manufactura", "Laboratorio Met\u00e1licos de Plasticidad de los\u2026", "Laboratorio de Qu\u00edmica Org\u00e1nica 1", "Taller de Principios de Ingenier\u00eda Qu\u00edmica", "Experiencia Inter y Multidisciplinar", "Transferencia de Masa", "\u00c9tica y Valores", "Degradaci\u00f3n de los Materiales", "Laboratorio lnterdisciplinario", "Temas Selectos de Producci\u00f3n y Manufactura 1", "Instrumentaci\u00f3n y Equipo 1 OPT.*", "Planeaci\u00f3n de la Producci\u00f3n", "Temas Selectos Inter y Multidisciplinares II", "Procesos de Separaci\u00f3n 11", "Responsabilidad Social Organizacional", "Dislocaciones en los Materiales Met\u00e1licos", "Tratamientos Met\u00e1licos T\u00e9rmicos de los Materiales", "Introducci\u00f3n Electr\u00f3nica al Trabajo de\u2026", "Metal\u00fargica Introducci\u00f3n al Trabajo de\u2026", "Inserci\u00f3n Laboral", "Ciencia y Tecnolog\u00eda de los Nanomateriales", "Laboratorio de Maquinado de los Materiales\u2026", "Proyecto de Integraci\u00f3n en Ingenier\u00eda Mec\u00e1nica 11", "TOTAL Proyecto 3.5.2.2. de DE Ciencias de se\u2026", "Control Laboratorio mediante de Instrumentaci\u00f3n\u2026", "Oxidaci\u00f3n, Met\u00e1licos Corrosi\u00f3n y Protecci\u00f3n de\u2026", "Termodin\u00e1mica y Cin\u00e9tica en Metalurgia", "Narrativa para Medios Audiovisuales y Digitales", "Aplicaciones Heterog\u00e9neos Industriales de\u2026", "Procesos Metal\u00fargicos Avanzados", "Habilidades Gerenciales", "Taller de Dispositivos Hidroneum\u00e1ticos", "Laboratorio de Radiocomunicaci\u00f3n OPT.*", "Planeaci\u00f3n Estrat\u00e9gica", "Temas Avanzados de Comunicaciones OPT.*", "Laboratorio de Soldadura de los Materiales\u2026", "Pr\u00e1cticas de Ciencias Profesionales se se\u2026", "Taller de Fotograf\u00eda", "Laboratorio Materiales de Comportamiento\u2026", "Taller de Ingenier\u00eda Civil I", "Laboratorio de Instrumentaci\u00f3n Industrial", "Trabajo Materiales en Met\u00e1licos Planta de\u2026", "Pr\u00e1cticas de Ciencias Profesionales se se\u2026", "Taller de Teatro", "Comunicaciones Industriales OPT.*", "Trabajo en Planta Sider\u00fargica", "Controladores L\u00f3gicos Programables OPT.*", "Im\u00e1genes", "Cambiadores de Calor OPT.*", "Metalurgia del Hierro", "Dispositivos Hidroneum\u00e1ticos", "Dise\u00f1o de Instrumentos OPT.*", "Ecodise\u00f1o", "Laboratorio Met\u00e1licos de Fundici\u00f3n de los\u2026", "Ingenier\u00eda El\u00e9ctrica", "Tratamientos Met\u00e1licos Termoqu\u00edmicos de los\u2026", "Interfaces Hombre-M\u00e1quina OPT.*", "Ingenier\u00eda de Reactores Avanzados OPT.*", "Temas Met\u00e1licos Selectos de Formado de los\u2026", "Temas Selectos de Gesti\u00f3n Industrial 11", "Tecnolog\u00eda Innovaci\u00f3n y en Administraci\u00f3n la\u2026", "Electroqu\u00edmica", "Comunicaciones Laboratorio de Simulaci\u00f3n de\u2026", "Pr\u00e1cticas de Ciencias Profesionales se se\u2026", "Comunicaci\u00f3n en Proyectos Multidisciplinarios", "Soldadura de los Materiales Met\u00e1licos", "Proyecto de Integraci\u00f3n en Ingenier\u00eda\u2026", "Laboratorio de Metalurgia Computacional", "Plasticidad de los Materiales Met\u00e1licos", "Psicolog\u00eda Industrial", "Laboratorio Materiales Met\u00e1licos de Mec\u00e1nica de\u2026", "Taller de Proyecto Mec\u00e1nico de Montajes OPT.*", "Familia y Violencia en el M\u00e9xico Contempor\u00e1neo", "Herramientas para el Emprendedor", "Comunicaciones \u00d3pticas OPT.*", "Sistemas de Manufactura Esbelta OPT.*", "Temas Selectos de Ingenier\u00eda en Computaci\u00f3n II", "Comunicaciones Celulares OPT.*", "Taller de Dibujo", "Trabajo Laboratorio de Estudio de la Medici\u00f3n del", "Instrumentaci\u00f3n Programaci\u00f3n Gr\u00e1fica y Control\u2026", "Mec\u00e1nica de Fluidos en Procesos Metal\u00fargicos", "Comportamiento Mec\u00e1nico de los Materiales OPT.*", "TOTAL Proyecto de 3.5.2.2. Ciencias DE de se\u2026", "G\u00e9nero y Sexualidad", "T\u00e9cnicas de Medici\u00f3n de Composici\u00f3n", "TOTAL Proyecto 3.5.2.2. de DE Ciencias de se\u2026", "Ergonom\u00eda Industrial", "Innovaci\u00f3n", "Laboratorio de Dise\u00f1o de Instrumentos OPT.*", "Metalurgia Mec\u00e1nica Avanzada", "Optoelectr\u00f3nica", "Introducci\u00f3n Parciales a las Ecuaciones\u2026", "Ciencia y Tecnolog\u00eda de los Cer\u00e1micos y Vidrios", "M\u00e9todos Decisiones Cuantitativos para la Toma de", "Dise\u00f1o e Innovaci\u00f3n de Productos OPT.*", "Tecnolog\u00edas de Radio M\u00f3vil OPT.*", "Mejora Continua de Procesos", "Pr\u00e1cticas Profesionales de Ingenier\u00eda Civil", "Introducci\u00f3n a la Bioqu\u00edmica OPT.*", "Temas Selectos de Ingenier\u00eda Mec\u00e1nica", "Sistemas Digitales y Redes", "Tecnolog\u00edas de la Informaci\u00f3n Industrial", "Ciencia y Tecnolog\u00eda de los Metales y Aleaciones", "Fen\u00f3menos de Superficie", "Laboratorio de Control de Calidad", "Qu\u00edmica Introducci\u00f3n al Trabajo de\u2026", "Taller de Ingenier\u00eda Civil II", "Proyecto de Integraci\u00f3n en Ingenier\u00eda Qu\u00edmica 11", ".,_,',\u00bfJ,~\\,'U~ TOTAL Proyecto . 3.5.2.2.\u2026", "Laboratorio Materiales Met\u00e1licos de Procesos de\u2026", "Laboratorio de Ciencia de los Materiales OPT.*", "Trabajo Met\u00e1licos en Planta de Fundici\u00f3n de los\u2026", "An\u00e1lisis de Problemas en Termofluidos OPT.*", "Poder y G\u00e9nero", "Simulaci\u00f3n y Control de Procesos", "Fotointerpretaci\u00f3n", "Introducci\u00f3n Ingenier\u00eda F\u00edsica al Trabajo de\u2026", "Metalograf\u00eda", "Laboratorio de Estudio del M\u00e9todo del Trabajo", "TOTAL Proyecto 3.5.2.2. de DE Ciencias de se\u2026", "Equilibrio Termodin\u00e1mico", "Laboratorio de F\u00edsica Moderna", "Ingenier\u00eda de Procesos", "Laboratorio Materiales Met\u00e1licos de\u2026", "TOTAL Proyecto 3.5.2.2. de DE Ciencias de se\u2026", "Temas Selectos Inter y Multidisciplinares I", "Optativa T\u00e9cnica de Movilidad IV", "Balance de Materia", "Metalurgia Extractiva", "Dise\u00f1o de Elementos de M\u00e1quinas 11 OPT.*", "Transformaciones Met\u00e1licos de Fase en los\u2026", "Ingenier\u00eda de Costos", "Temas Selectos de Qu\u00edmica", "S\u00edntesis, Materiales Caracterizaci\u00f3n\u2026", "Proyecto Mec\u00e1nico OPT.*", "Laboratorio de Fisicoqu\u00edmica de los Materiales", "Dise\u00f1o Mec\u00e1nico de Equipos OPT.*", "Estudio de la Medici\u00f3n del Trabajo", "Control de Calidad", "Derechos Humanos", "Gesti\u00f3n de Servicios OPT.*", "Ciencias Ambientales y Sustentabilidad", "Procesos de Mercadeo OPT.*", "Ciencia de los Materiales OPT.*", "Laboratorio de \u00d3ptica", "Electroqu\u00edmica de los Materiales", "Termotransferencia en Procesos Metal\u00fargicos", "Temas Selectos de Ingenier\u00eda Qu\u00edmica", "Temas Selectos de Gesti\u00f3n Industrial 1", "Pr\u00e1cticas Profesionales de Ingenier\u00eda Ambiental", "Ingenier\u00eda de Productos y de Procesos Metal\u00fargicos", "Administraci\u00f3n y Econom\u00eda Pol\u00edtica de la\u2026", "Divulgaci\u00f3n del Conocimiento", "Ac\u00fastica", "Est\u00e1tica del Cuerpo Deformable", "Gerencia de Proyectos de Infraestructura", "Pr\u00e1cticas de Ciencias Profesionales se se\u2026", "Temas Selectos de Telecomunicaciones 1 OPT.*", "Pr\u00e1cticas Profesionales de Ingenier\u00eda El\u00e9ctrica", "Laboratorio de Ingenier\u00eda El\u00e9ctrica", "Recristalizaci\u00f3n Din\u00e1mica de los Materiales\u2026", "Superplasticidad en Materiales Avanzados", "An\u00e1lisis Horno de Materiales y Balance T\u00e9rmico\u2026", "Dise\u00f1o Electr\u00f3nico Asistido por Computadora", "Trabajo en Planta Metal\u00fargica", "Historia Social de M\u00e9xico en el Siglo XX", "Mec\u00e1nica de la Fractura de los Materiales\u2026", "Ingenier\u00eda de Proyectos OPT.*", "Laboratorio Materiales Met\u00e1licos de\u2026", "El Paisaje como Agente de los Asentamientos y\u2026", "Hornos y Combusti\u00f3n en Procesos Metal\u00fargicos", "Historia del Arte", "Seminario de Integraci\u00f3n en Ingenier\u00eda F\u00edsica", "Energ\u00eda Solar Aplicada", "An\u00e1lisis de Problemas", "Optativa T\u00e9cnica de Movilidad II", "An\u00e1lisis de Sistemas de Control OPT.*", "Taller de Proyecto Mec\u00e1nico OPT.*", "Calefacci\u00f3n, Ventilaci\u00f3n y Aire Acondicionado", "Comunicaciones Satel ita les OPT.*", "Metalurgia Mec\u00e1nica", "Temas Selectos de Control e Instrumentaci\u00f3n 1\u2026", "Maquinado de los Materiales Met\u00e1licos", "Taller de Manufactura Asistido por Computadora", "Pr\u00e1cticas Profesionales en Ingenier\u00eda F\u00edsica", "Temas Selectos de Ingenier\u00eda Metal\u00fargica 1", "Trabajo en Planta de Aleaciones no Ferrosas", "Trabajo en Planta Metal Mec\u00e1nica", "Temas Selectos de Ingenier\u00eda en Computaci\u00f3n I", "Ciencia y Tecnolog\u00eda de los Pol\u00edmeros", "Ingenier\u00eda \u00d3ptica", "Ingenier\u00eda del Producto", "alumno Metalurgia que Procesos se del apruebe\u2026", "Contaminaci\u00f3n Ambiental", "Introducci\u00f3n al C\u00e1lculo"];
    return [{
      type: 'scatter',
      mode: 'markers',
      x: sc_x,
      y: sc_y,
      marker: {
        size: sc_sz,
        color: sc_col,
        opacity: 0.65,
        line: { width: 0.5, color: 'rgba(255,255,255,0.5)' }
      },
      customdata: sc_names.map(function(n, i) { return [n, sc_rc[i], sc_imp[i]]; }),
      hovertemplate: '<b>%{customdata[0]}</b><br>Profundidad de cadena: %{x}<br>HHI: %{y:.3f}<br>RC: %{customdata[1]:.3f}<br>Impacto cascada: %{customdata[2]:.3f}<extra></extra>',
      name: 'UEA'
    }];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'UEAs: profundidad de cadena vs HHI (tamaño = impacto cascada)', font: FONT, x: 0.5 },
      xaxis: { title: 'Profundidad de cadena (n° prerrequisitos)', tickfont: FONT, range: [-0.5, 12] },
      yaxis: { title: 'HHI', tickfont: FONT, range: [0, 1.05] },
      margin: { l: 60, r: 20, t: 50, b: 60 },
      showlegend: false
    });
  }
);
"""

conc2_html = page_shell(
    title='El riesgo estructural — DCBI',
    active='concentracion.html',
    content=CONC2_CONTENT,
    data_vars='',
    extra_js=CONC2_JS,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="conc_1.html"><span class="bottom-nav-icon">&#9664;</span>Concentración</a>
      <a class="bottom-nav-item" href="concentracion.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="conc_3.html"><span class="bottom-nav-icon">&#9654;</span>Vecindades</a>
    """,
)

# ══════════════════════════════════════════════════════════════════════════════
# CONC_3.HTML  — Infografía 3: Redes y vecindades
# ══════════════════════════════════════════════════════════════════════════════
CONC3_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Distribución docente · Infografía 3 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>Redes y vecindades</strong></h1>
      <p class="hero-sub">Una vecindad curricular controlada es un grupo de UEAs relacionadas donde un pequeño número de profesores domina la impartición histórica del bloque completo: 11 vecindades identificadas donde la fragilidad es sistémica, no puntual.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">Una <strong>vecindad curricular controlada</strong> es un grupo de UEAs relacionadas entre sí —por seriación formal, por corregistro obligatorio o por similitud de contenido en sus programas— donde además un pequeño número de profesores domina la impartición histórica del bloque completo. La combinación de vínculos curriculares fuertes con concentración docente sostenida crea una zona del plan de estudios donde la fragilidad no es puntual sino sistémica: afectar a un docente o a una UEA del bloque afecta al conjunto. El análisis identificó <strong>11 vecindades controladas</strong> de entre 22 grupos detectados automáticamente.</p>
  </section>

  <section class="section animate-in" id="chart-section-vecindades">
    <div class="section-header">
      <h2>Vecindades curriculares controladas (HHI cluster)</h2>
      <span class="section-count">11 vecindades · HHI de vecindad agregado</span>
    </div>
    <p class="section-desc">El HHI de vecindad agrega la concentración docente a nivel del bloque completo, no UEA por UEA. Las barras en naranja tienen concentración elevada; las verdes, moderada. Pasa el cursor para ver el RC medio del bloque.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> el hallazgo contraintuitivo: las vecindades con mayor HHI de cluster no son las de mayor riesgo curricular real. La posición en el grafo de seriaciones importa más que el HHI agregado.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>HHI de vecindad curricular (11 vecindades controladas)</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-vecindades">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">El gráfico expone un hallazgo contraintuitivo que tiene consecuencias directas para la gobernanza: <strong>las vecindades con mayor concentración docente agregada no son las de mayor riesgo curricular</strong>. La vecindad C16 (Comportamiento Mecánico de los Materiales), con el HHI de cluster más alto del conjunto (0.137), registra Riesgo Concentrado nulo en todas sus UEAs porque ninguna de ellas forma parte de cadenas de seriación formales: la concentración existe, pero está desacoplada del mecanismo de cascada. En contraste, la vecindad C1 (Hidráulica y recursos hídricos), con un HHI de cluster de apenas 0.026, acumula <strong>19 UEAs en nivel crítico</strong> precisamente por el mecanismo inverso: concentración docente moderada amplificada por cadenas de hasta nueve eslabones. El riesgo institucional real es más alto donde la concentración docente <strong>parece</strong> menor. Medir solo el HHI, sin considerar la posición en el grafo de seriaciones, lleva a priorizar mal.</p>
  </section>

  <section class="section animate-in" id="chart-section-vecindades-risk">
    <div class="section-header">
      <h2>UEAs críticas por vecindad curricular</h2>
      <span class="section-count">Crítico (RC ≥ p90) + Alto (RC p75–p90) · barras apiladas</span>
    </div>
    <p class="section-desc">El apilado muestra UEAs críticas (rojo) y alto riesgo (naranja) por vecindad. La longitud total indica la exposición agregada del bloque.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> qué vecindades acumulan más UEAs críticas y cómo eso contrasta con su posición en el gráfico anterior de HHI de cluster.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>UEAs en riesgo por vecindad curricular</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering sm" id="chart-vecindades-risk">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">Entre las vecindades más relevantes para la gobernanza destacan tres perfiles distintos. <strong>Cluster 13 — Comunicaciones (Electrónica)</strong> agrupa 19 UEAs de la línea de comunicaciones analógicas y digitales con sus laboratorios correspondientes; su sensibilidad estructural es alta porque combina teoría y práctica en cadenas de seriación verticales donde el estudiante no puede avanzar sin haber completado los cursos precedentes, todos bajo la misma estructura de impartición concentrada (tasa de aprobación media: 74.5 %). <strong>Cluster 11 — Mecánica de sólidos y edificación</strong> agrupa 24 UEAs con la segunda tasa de aprobación más baja de las vecindades controladas (64.6 %): la combinación de bajo rendimiento y concentración docente crea un cuello de botella académico sostenido. <strong>Cluster 17 — Prácticas profesionales</strong> presenta un riesgo de naturaleza diferente: sus 11 UEAs de prácticas son, en varios planes, un <strong>requisito de egreso</strong>, de modo que la concentración docente en este bloque tiene implicaciones directas sobre la eficiencia terminal, no solo sobre el avance trimestral.</p>
  </section>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <p style="font-size:.84rem;color:var(--mid);line-height:1.6">Las vecindades controladas revelan la textura interna del riesgo; el mapa de riesgo por plan y por departamento muestra su distribución institucional. La siguiente infografía cierra el análisis identificando qué planes y qué departamentos acumulan la mayor exposición estructural y dónde la intervención produciría el mayor retorno.</p>
    <div class="connect-links">
      <a class="connect-link" href="conc_2.html">← Infografía 2: El riesgo estructural</a>
      <a class="connect-link" href="conc_4.html">Infografía 4: El mapa de riesgo →</a>
    </div>
  </div>

</main>
"""

CONC3_JS = """
renderChart('chart-vecindades',
  function() {
    var names  = ["Comportamiento Mec\\u00e1nico y Fractura de\\u2026", "Sistemas de Comunicaciones Anal\\u00f3gicas,\\u2026", "Sistemas Digitales, Redes y Proyectos de\\u2026", "Pr\\u00e1cticas Profesionales y Gesti\\u00f3n de\\u2026", "An\\u00e1lisis y Dise\\u00f1o Estructural de Obras de\\u2026", "Metalurgia de Procesos: Fundici\\u00f3n, Formado\\u2026", "Arquitectura de Computadoras, Redes y\\u2026", "Seminarios y Proyectos de Integraci\\u00f3n\\u2026", "Hidr\\u00e1ulica, Geotecnia y Construcci\\u00f3n de\\u2026", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y\\u2026", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y\\u2026"];
    var hhi_c  = [0.1373, 0.0915, 0.0826, 0.0801, 0.0479, 0.0456, 0.0302, 0.0294, 0.0261, 0.0187, 0.0115];
    var hhi_m  = [0.4759, 0.44, 0.5, 1.0, 0.3142, 0.38, 0.3333, 0.4832, 0.42, 0.25, 0.2175];
    var nueas  = [14, 19, 11, 11, 24, 35, 37, 48, 53, 55, 35];
    var rc     = [0.5524, 0.5121, 0.4951, 0.7596, 0.3991, 0.4371, 0.4204, 0.5354, 0.4763, 0.3423, 0.2849];
    var colors = ["#FF9800", "#FF9800", "#FF9800", "#FF9800", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32"];
    var full   = ["Comportamiento Mec\\u00e1nico y Fractura de Materiales Met\\u00e1licos", "Sistemas de Comunicaciones Anal\\u00f3gicas, Digitales y Radiofrecuencia", "Sistemas Digitales, Redes y Proyectos de Integraci\\u00f3n El\\u00e9ctrica", "Pr\\u00e1cticas Profesionales y Gesti\\u00f3n de Calidad Industrial", "An\\u00e1lisis y Dise\\u00f1o Estructural de Obras de Ingenier\\u00eda Civil", "Metalurgia de Procesos: Fundici\\u00f3n, Formado y Tratamientos T\\u00e9rmicos", "Arquitectura de Computadoras, Redes y Sistemas de Informaci\\u00f3n", "Seminarios y Proyectos de Integraci\\u00f3n Multidisciplinar en Ingenier\\u00eda", "Hidr\\u00e1ulica, Geotecnia y Construcci\\u00f3n de Infraestructura Civil", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y Fundamentos Computacionales", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y Fundamentos Computacionales"];
    return [{
      type: 'bar',
      orientation: 'h',
      x: hhi_c,
      y: names,
      marker: { color: colors, opacity: 0.85 },
      text: hhi_c.map(function(v) { return v.toFixed(4); }),
      textposition: 'outside',
      cliponaxis: false,
      customdata: names.map(function(n, i) { return [full[i], nueas[i], hhi_m[i], rc[i]]; }),
      hovertemplate: '<b>%{customdata[0]}</b><br>HHI vecindad: %{x:.4f}<br>HHI mediana UEAs: %{customdata[2]:.3f}<br>UEAs: %{customdata[1]}<br>RC medio: %{customdata[3]:.3f}<extra></extra>',
      name: 'HHI vecindad'
    }];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'HHI de vecindad curricular (11 vecindades controladas)', font: FONT, x: 0.5 },
      xaxis: { title: 'HHI de vecindad', tickfont: FONT, range: [0, 0.20] },
      yaxis: { automargin: true, tickfont: FONT },
      margin: { l: 290, r: 60, t: 50, b: 50 },
      bargap: 0.3
    });
  }
);

renderChart('chart-vecindades-risk',
  function() {
    var names   = ["Comportamiento Mec\\u00e1nico y Fractura de\\u2026", "Sistemas de Comunicaciones Anal\\u00f3gicas,\\u2026", "Sistemas Digitales, Redes y Proyectos de\\u2026", "Pr\\u00e1cticas Profesionales y Gesti\\u00f3n de\\u2026", "An\\u00e1lisis y Dise\\u00f1o Estructural de Obras de\\u2026", "Metalurgia de Procesos: Fundici\\u00f3n, Formado\\u2026", "Arquitectura de Computadoras, Redes y\\u2026", "Seminarios y Proyectos de Integraci\\u00f3n\\u2026", "Hidr\\u00e1ulica, Geotecnia y Construcci\\u00f3n de\\u2026", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y\\u2026", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y\\u2026"];
    var n_crit  = [0, 1, 0, 0, 6, 0, 11, 0, 19, 4, 4];
    var n_alto  = [0, 1, 2, 0, 5, 0, 13, 5, 18, 5, 8];
    var rc_max  = [0.0, 0.862, 0.4, 0.0, 1.0, 0.213, 1.0, 0.435, 1.0, 1.0, 1.0];
    var full    = ["Comportamiento Mec\\u00e1nico y Fractura de Materiales Met\\u00e1licos", "Sistemas de Comunicaciones Anal\\u00f3gicas, Digitales y Radiofrecuencia", "Sistemas Digitales, Redes y Proyectos de Integraci\\u00f3n El\\u00e9ctrica", "Pr\\u00e1cticas Profesionales y Gesti\\u00f3n de Calidad Industrial", "An\\u00e1lisis y Dise\\u00f1o Estructural de Obras de Ingenier\\u00eda Civil", "Metalurgia de Procesos: Fundici\\u00f3n, Formado y Tratamientos T\\u00e9rmicos", "Arquitectura de Computadoras, Redes y Sistemas de Informaci\\u00f3n", "Seminarios y Proyectos de Integraci\\u00f3n Multidisciplinar en Ingenier\\u00eda", "Hidr\\u00e1ulica, Geotecnia y Construcci\\u00f3n de Infraestructura Civil", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y Fundamentos Computacionales", "F\\u00edsica Aplicada, Matem\\u00e1ticas Avanzadas y Fundamentos Computacionales"];
    return [
      {
        type: 'bar',
        orientation: 'h',
        x: n_crit,
        y: names,
        name: 'Crítico',
        marker: { color: '#C82D23', opacity: 0.9 },
        customdata: full.map(function(f, i) { return [f, rc_max[i]]; }),
        hovertemplate: '<b>%{customdata[0]}</b><br>UEAs críticas: %{x}<br>RC máx: %{customdata[1]:.3f}<extra></extra>'
      },
      {
        type: 'bar',
        orientation: 'h',
        x: n_alto,
        y: names,
        name: 'Alto',
        marker: { color: '#FF9800', opacity: 0.85 },
        hovertemplate: '<b>%{y}</b><br>UEAs alto riesgo: %{x}<extra></extra>'
      }
    ];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'UEAs en riesgo por vecindad curricular', font: FONT, x: 0.5 },
      xaxis: { title: 'Número de UEAs', tickfont: FONT },
      yaxis: { automargin: true, tickfont: FONT },
      barmode: 'stack',
      margin: { l: 290, r: 30, t: 50, b: 50 },
      bargap: 0.3,
      legend: { x: 0.75, y: 0.05, font: FONT }
    });
  }
);
"""

conc3_html = page_shell(
    title='Redes y vecindades — DCBI',
    active='concentracion.html',
    content=CONC3_CONTENT,
    data_vars='',
    extra_js=CONC3_JS,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="conc_2.html"><span class="bottom-nav-icon">&#9664;</span>Riesgo</a>
      <a class="bottom-nav-item" href="concentracion.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
      <a class="bottom-nav-item" href="conc_4.html"><span class="bottom-nav-icon">&#9654;</span>Mapa</a>
    """,
)

# ══════════════════════════════════════════════════════════════════════════════
# CONC_4.HTML  — Infografía 4: El mapa de riesgo
# ══════════════════════════════════════════════════════════════════════════════
CONC4_CONTENT = """
<header class="hero">
  <div class="hero-inner">
    <div>
      <p class="hero-label">Distribución docente · Infografía 4 · División de Ciencias Básicas e Ingeniería · UAM-Azcapotzalco</p>
      <h1><strong>El mapa de riesgo</strong></h1>
      <p class="hero-sub">Los análisis anteriores operaron sobre UEAs individuales y vecindades. Esta infografía eleva la escala: qué planes de estudio y qué departamentos concentran la mayor exposición estructural, con variación de dos órdenes de magnitud entre extremos.</p>
      <div class="hero-bar"></div>
    </div>
  </div>
</header>

<main class="main">

  <section class="section animate-in">
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid)">Los análisis anteriores operaron sobre UEAs individuales y vecindades. Esta infografía eleva la escala: ¿qué planes de estudio y qué departamentos concentran la mayor exposición estructural? La respuesta no es uniforme. La variación entre los planes más y menos expuestos alcanza dos órdenes de magnitud, y el patrón no es aleatorio: refleja la arquitectura curricular de cada licenciatura —la densidad de sus cadenas de seriación y la especialización de sus plantillas docentes— más que diferencias en el volumen de UEAs.</p>
  </section>

  <section class="section animate-in" id="chart-section-risk-plan">
    <div class="section-header">
      <h2>Riesgo estructural por plan de estudios</h2>
      <span class="section-count">10 planes · UEAs críticas + alto riesgo + medio-bajo · RC medio como diamante</span>
    </div>
    <p class="section-desc">Barras apiladas por nivel de riesgo; el diamante negro en el eje superior marca el RC medio de cada plan. Los planes están ordenados de menor a mayor exposición total.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> la brecha entre Ingeniería Civil y el resto, el RC medio de los planes intermedios, y cuáles planes tienen cero UEAs críticas a pesar de tener HHI alto.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Distribución de riesgo por plan de estudios</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering lg" id="chart-risk-plan">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem"><strong>Ingeniería Civil</strong> encabeza la clasificación con una distancia considerable respecto al resto: el <strong>30.6 % de sus UEAs son críticas</strong> (RC ≥ percentil 90) y el RC medio del plan es 0.36, el más alto entre los diez programas. Le siguen <strong>Ingeniería en Computación</strong> (17.6 % de UEAs críticas, RC medio 0.29) e <strong>Ingeniería Eléctrica</strong> (15.9 %, RC medio 0.27). En el extremo opuesto, <strong>Ingeniería Física, Ingeniería Química, Ingeniería Industrial e Ingeniería Metalúrgica</strong> registran cero UEAs en nivel crítico y RC medio inferior a 0.02. Esta diferencia no refleja una gestión docente mejor o peor en esos planes: refleja que sus cadenas de seriación son menos profundas y sus UEAs de especialidad tienen, históricamente, mayor diversidad de impartidores. El diseño curricular determina la magnitud del riesgo tanto como la política de asignación docente.</p>
  </section>

  <section class="section animate-in" id="chart-section-risk-dept">
    <div class="section-header">
      <h2>Mapa de riesgo por departamento</h2>
      <span class="section-count">HHI mediana (eje x) · RC medio (eje y) · tamaño = n° UEAs</span>
    </div>
    <p class="section-desc">Cada burbuja es un departamento. La zona sombreada en rojo (HHI &gt; 0.5 y RC &gt; 0.15) marca el cuadrante de máxima exposición compuesta. El tamaño de la burbuja es proporcional al número de UEAs del departamento.</p>
    <div class="read-callout"><strong>Qué buscar:</strong> qué departamentos entran en la zona crítica, cuál tiene el HHI más alto pero RC bajo, y cómo Ciencias Básicas se separa del resto en dirección contraria a todos.</div>
    <div class="chart-card animate-in delay-1">
      <div class="chart-card-header">
        <h3>Perfil de riesgo departamental</h3>
        <span class="chart-type-tag">Interactivo</span>
      </div>
      <div class="chart-canvas rendering md" id="chart-risk-dept">
        <div class="chart-skeleton"></div>
      </div>
    </div>
    <p style="font-size:.92rem;line-height:1.75;color:var(--mid);margin-top:1.2rem">A nivel departamental, <strong>Materiales</strong> registra el RC medio más alto entre los departamentos con más de cinco UEAs (0.21), con una mediana de HHI de 0.37 y el 14 % de sus UEAs en régimen de exclusiva. Es el departamento con mayor exposición compuesta, aunque permanece ligeramente por debajo del umbral crítico de HHI = 0.5. <strong>Dirección de la División</strong> presenta la mediana de HHI más alta (0.50) y el porcentaje de monopolios individuales más elevado (23 %), pero su RC medio es menor (0.13): sus UEAs en exclusiva están en su mayoría fuera de cadenas de seriación largas, lo que aísla el efecto de cascada más grave. <strong>Ciencias Básicas</strong> se comporta de forma inversa a todos los demás: HHI mediana baja (0.17), sin UEAs en el cuadrante crítico y el número de monopolios más reducido de la División, en coherencia con la estructura del Tronco General.</p>
  </section>

  <div class="recuadro animate-in">
    <div class="recuadro-title">Síntesis</div>
    <p>Tres conclusiones organizan la agenda de gobernanza que se desprende de este análisis. <strong>Primera:</strong> la concentración docente es el estado habitual del sistema, no una anomalía —el 44.3 % de las UEAs tiene HHI ≥ 0.5—, por lo que las políticas de relevo docente deben diseñarse como estructura permanente, no como respuesta a emergencias. <strong>Segunda:</strong> el riesgo estructural está localizado: 61 UEAs críticas en tres planes (Civil, Computación, Eléctrica) y un conjunto reducido de profesores que son puntos únicos de fallo acumulan la mayor parte de la exposición; la intervención focalizada tiene alto retorno. <strong>Tercera:</strong> medir solo el HHI es insuficiente —la vecindad C1 lo demuestra—: cualquier política de diversificación docente debe cruzar la concentración con la posición de la UEA en el grafo de seriaciones para priorizar correctamente.</p>
  </div>

  <hr class="divider">

  <div class="connect-box animate-in">
    <div class="connect-box-title">Navegación</div>
    <div class="connect-links">
      <a class="connect-link" href="conc_3.html">← Infografía 3: Redes y vecindades</a>
      <a class="connect-link" href="concentracion.html">↑ Índice Distribución docente</a>
    </div>
  </div>

</main>
"""

CONC4_JS = """
renderChart('chart-risk-plan',
  function() {
    var labels = ["Ing. Metal\\u00fargica", "Ing. Industrial", "Ing. Qu\\u00edmica", "Lic. F\\u00edsica Aplicada", "Ing. Mec\\u00e1nica", "Ing. El\\u00e9ctrica", "Ing. Ambiental", "Ing. Electr\\u00f3nica", "Ing. Computaci\\u00f3n", "Ing. Civil"];
    var n_crit = [0, 0, 0, 0, 1, 3, 8, 7, 16, 26];
    var n_alto = [0, 0, 0, 0, 2, 1, 25, 14, 25, 23];
    var n_mb   = [56, 37, 36, 61, 13, 30, 110, 23, 50, 36];
    var n_ueas = [56, 37, 36, 61, 16, 34, 143, 44, 91, 85];
    var rc_med = [0.0088, 0.0106, 0.0135, 0.0185, 0.0787, 0.0984, 0.1313, 0.2731, 0.288, 0.3607];
    var hhi_m  = [0.4873, 0.4264, 0.3487, 0.3138, 0.5594, 0.5203, 0.3066, 0.3811, 0.3508, 0.4663];
    var pct_c  = [0.0, 0.0, 0.0, 0.0, 6.2, 8.8, 5.6, 15.9, 17.6, 30.6];
    var pct_a  = [0.0, 0.0, 0.0, 0.0, 12.5, 2.9, 17.5, 31.8, 27.5, 27.1];
    return [
      {
        type: 'bar',
        orientation: 'h',
        x: n_crit,
        y: labels,
        name: 'Crítico',
        marker: { color: '#C82D23', opacity: 0.9 },
        customdata: labels.map(function(l, i) { return [pct_c[i], n_ueas[i], rc_med[i]]; }),
        hovertemplate: '%{y} — Crítico<br>UEAs: %{x} (%{customdata[0]}%)<br>RC medio: %{customdata[2]:.3f}<extra></extra>'
      },
      {
        type: 'bar',
        orientation: 'h',
        x: n_alto,
        y: labels,
        name: 'Alto',
        marker: { color: '#FF9800', opacity: 0.85 },
        customdata: labels.map(function(l, i) { return [pct_a[i]]; }),
        hovertemplate: '%{y} — Alto<br>UEAs: %{x} (%{customdata[0]}%)<extra></extra>'
      },
      {
        type: 'bar',
        orientation: 'h',
        x: n_mb,
        y: labels,
        name: 'Medio-bajo',
        marker: { color: 'rgba(0,0,0,0.12)' },
        hovertemplate: '%{y} — Medio-bajo<br>UEAs: %{x}<extra></extra>'
      },
      {
        type: 'scatter',
        mode: 'markers',
        x: rc_med,
        y: labels,
        xaxis: 'x2',
        marker: { size: 9, color: '#1A1A1A', symbol: 'diamond', line: { color: '#fff', width: 1 } },
        name: 'RC medio',
        hovertemplate: 'RC medio: %{x:.3f}<extra></extra>'
      }
    ];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'Distribución de riesgo por plan de estudios', font: FONT, x: 0.5 },
      xaxis:  { title: 'Número de UEAs', tickfont: FONT, domain: [0, 0.82] },
      xaxis2: { title: 'RC medio', overlaying: 'x', side: 'top', range: [0, 0.5], tickfont: FONT, showgrid: false },
      yaxis:  { automargin: true, tickfont: FONT },
      barmode: 'stack',
      margin: { l: 155, r: 20, t: 60, b: 50 },
      bargap: 0.25,
      legend: { x: 0.84, y: 0.5, font: FONT }
    });
  }
);

renderChart('chart-risk-dept',
  function() {
    var b_names  = ["EVA. DISENO EN E", "DERECHO", "PROCESOS TEC. RE", "Direcci\\u00f3n", "INV.CONOCIMIENTO", "MEDIO AMBIENTE", "SOCIOLOGIA", "Humanidades", "Materiales", "ADMINISTRACION", "Electr\\u00f3nica", "Sistemas", "Energ\\u00eda", "Ciencias B\\u00e1sicas", "ECONOMIA"];
    var b_hhi    = [0.6736, 0.6172, 0.5306, 0.5, 0.4132, 0.39, 0.3795, 0.375, 0.3646, 0.346, 0.2639, 0.2503, 0.2474, 0.166, 0.1517];
    var b_rc     = [0.2266, 0.0, 0.0, 0.1278, 0.0, 0.0, 0.0, 0.0, 0.2142, 0.0, 0.1638, 0.1918, 0.1864, 0.0759, 0.0];
    var b_nueas  = [2, 2, 1, 109, 1, 2, 2, 5, 107, 2, 83, 62, 138, 89, 1];
    var b_pct1   = [0.0, 50.0, 0.0, 22.9, 0.0, 0.0, 0.0, 20.0, 14.0, 0.0, 9.6, 6.5, 4.3, 4.5, 0.0];
    var b_sz     = [8.0, 8.0, 8.0, 46.9814, 8.0, 8.0, 8.0, 10.0623, 46.5484, 8.0, 40.997, 35.433, 52.863, 42.4529, 8.0];
    var b_colors = ["#2E7D32", "#C82D23", "#2E7D32", "#C82D23", "#2E7D32", "#2E7D32", "#2E7D32", "#C82D23", "#FF9800", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32", "#2E7D32"];
    return [
      {
        type: 'scatter',
        mode: 'markers+text',
        x: b_hhi,
        y: b_rc,
        marker: {
          size: b_sz,
          color: b_colors,
          opacity: 0.75,
          line: { width: 1, color: 'rgba(255,255,255,0.6)' }
        },
        text: b_names,
        textposition: 'top center',
        textfont: Object.assign({}, FONT, { size: 10 }),
        customdata: b_names.map(function(n, i) { return [b_nueas[i], b_pct1[i]]; }),
        hovertemplate: '<b>%{text}</b><br>HHI mediana: %{x:.3f}<br>RC medio: %{y:.3f}<br>UEAs: %{customdata[0]}<br>% monopolio (HHI=1): %{customdata[1]}%<extra></extra>',
        name: 'Departamento'
      },
      {
        type: 'scatter',
        mode: 'lines',
        x: [0.5, 0.5, null, 0.5, 1.0],
        y: [0.15, 0.7, null, 0.15, 0.15],
        line: { color: 'rgba(200,45,35,0.25)', width: 1, dash: 'dot' },
        fill: 'none',
        hoverinfo: 'skip',
        showlegend: false,
        name: ''
      }
    ];
  },
  function() {
    return Object.assign({}, LAYOUT_BASE, {
      title: { text: 'Perfil de riesgo departamental (tamaño = n° UEAs, color = % monopolio)', font: FONT, x: 0.5 },
      xaxis: { title: 'HHI mediana', tickfont: FONT, range: [0, 0.6] },
      yaxis: { title: 'RC medio', tickfont: FONT, range: [-0.02, 0.35] },
      margin: { l: 60, r: 30, t: 50, b: 60 },
      showlegend: false,
      shapes: [{
        type: 'rect',
        x0: 0.5, x1: 0.65, y0: 0.15, y1: 0.30,
        fillcolor: 'rgba(200,45,35,0.06)',
        line: { color: 'rgba(200,45,35,0.3)', width: 1, dash: 'dot' },
        layer: 'below'
      }],
      annotations: [{
        x: 0.57, y: 0.29,
        text: 'Zona crítica',
        showarrow: false,
        font: Object.assign({}, FONT, { size: 10, color: '#C82D23' })
      }]
    });
  }
);
"""

conc4_html = page_shell(
    title='El mapa de riesgo — DCBI',
    active='concentracion.html',
    content=CONC4_CONTENT,
    data_vars='',
    extra_js=CONC4_JS,
    bottom_nav_items="""
      <a class="bottom-nav-item" href="conc_3.html"><span class="bottom-nav-icon">&#9664;</span>Vecindades</a>
      <a class="bottom-nav-item" href="concentracion.html"><span class="bottom-nav-icon">&#127968;</span>Índice</a>
    """,
)

# ══════════════════════════════════════════════════════════════════════════════
# WRITE FILES
# ══════════════════════════════════════════════════════════════════════════════
for fname, content in [
    ('index.html',           index_html),
    ('cat26p.html',          cat_html),
    ('cat26p_capacidad.html',cap_html),
    ('metanalisis.html',     meta_html),
    ('meta_1.html',          meta1_html),
    ('meta_2.html',          meta2_html),
    ('meta_3.html',          meta3_html),
    ('meta_4.html',          meta4_html),
    ('meta_5.html',          meta5_html),
    ('meta_6.html',          meta6_html),
    ('modificaciones.html',  mod_html),
    ('concentracion.html',   conc_html),
    ('conc_1.html',          conc1_html),
    ('conc_2.html',          conc2_html),
    ('conc_3.html',          conc3_html),
    ('conc_4.html',          conc4_html),
]:
    p = HERE / fname
    p.write_text(content, encoding='utf-8')
    print(f"  {fname:<28} {p.stat().st_size/1024:.0f} KB")

print("\nDone — 16 pages written to", HERE)
