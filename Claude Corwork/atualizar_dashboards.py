"""
Atualizar Dashboards de Inventário
====================================
Lê automaticamente os arquivos Excel do OneDrive e regera os 4 dashboards HTML.

Uso:
  - Duplo-clique em "Atualizar Dashboards.bat"
  - Ou: python atualizar_dashboards.py

Requisitos:
  pip install pandas openpyxl numpy

Arquivos esperados em C:\\Users\\Cleuvo\\OneDrive\\ (ou subpastas):
  - Inventário Análise de Erros*.xlsx
  - Inventário Analítico Portal*.xlsx
"""
import os, sys, glob, json, math, unicodedata
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("ERRO: Instale as dependências com:  pip install pandas openpyxl numpy")
    input("\nPressione Enter para sair...")
    sys.exit(1)

# ── CONFIGURAÇÃO ─────────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).parent.resolve()
ONEDRIVE    = Path(r"C:\Users\Cleuvo\OneDrive")
OUT_DIR     = SCRIPT_DIR  # Salvar HTMLs na mesma pasta do script

def make_pid(plano):
    """Gera um ID ASCII seguro a partir do nome do plano (sem acentos, sem espaços)."""
    nfkd = unicodedata.normalize('NFKD', plano)
    ascii_str = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return ascii_str.replace(' ', '_').replace('/', '_')

PLAN_MAP = {
    'INVENTÁRIO GERAL VCP CLIMATIZADO': 'Climatizado',
    'INVENTARIO ODONTO VCP': 'Odonto',
    'Inventario geral armazém 72 vcp': 'Armazém 72',
    'INVENTÁRIO GERAL VCP CLIMATIZADO 70.02.007 RUA 1 A 12 NÍVEL 7': 'Climatizado RUA7/Nív7',
}
PLAN_ORDER = ['Climatizado','Odonto','Armazém 72','Climatizado RUA7/Nív7']
RANK = {'A':1,'B':2,'C':3}

# ── LOCALIZAR ARQUIVOS ───────────────────────────────────────────────────────
def find_file(pattern_name):
    """Busca um arquivo por padrão glob no OneDrive e subpastas."""
    patterns = [
        ONEDRIVE / f"{pattern_name}*.xlsx",
        ONEDRIVE / "**" / f"{pattern_name}*.xlsx",
    ]
    for pat in patterns:
        found = sorted(glob.glob(str(pat), recursive=True))
        if found:
            # Pega o mais recente por nome (data no nome do arquivo)
            return found[-1]
    return None

# ── HELPERS ───────────────────────────────────────────────────────────────────
class NpEnc(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, np.integer): return int(o)
        if isinstance(o, np.floating): return None if math.isnan(o) else float(o)
        if isinstance(o, np.ndarray): return o.tolist()
        return super().default(o)

def s(v, d=0):
    if v is None or (isinstance(v, float) and math.isnan(v)): return d
    return v

def parse_addr(addr):
    p = str(addr).split('.')
    return {k: (p[i] if i < len(p) else '') for i, k in enumerate(['arm','area','rua','predio','nivel'])}

def fmt_n(v):
    try: return f'{int(v):,}'
    except: return str(v)

def fmt_m(m):
    if not m or m == 0: return '—'
    m = float(m)
    if m < 1:  return f'{m*60:.0f}s'
    if m < 60: return f'{m:.1f}m'
    return f'{m/60:.1f}h'

def col_for(m):
    if not m or m == 0: return '#6b7280'
    m = float(m)
    if m < 1:  return '#22c55e'
    if m < 2:  return '#84cc16'
    if m < 5:  return '#eab308'
    if m < 15: return '#f97316'
    return '#ef4444'

CNT_STYLE = {'A':('#1e3a5f','#60a5fa'),'B':('#14532d','#4ade80'),'C':('#4c1d95','#c084fc')}
TIPO_CLS = {
    'Divergência pequena (<10%)':'bt-warn','Sobra significativa':'bt-err',
    'Falta significativa':'bt-err2','Posição vazia/contagem=0':'bt-gray',
    'Posição vazia correta':'bt-ok'
}
BIN_LBL = ['<1min','1-2min','2-5min','5-15min','15-30min','>30min']
HIST_COLORS = ['#22c55e','#84cc16','#eab308','#f97316','#ef4444','#991b1b']

# ── ANÁLISE DE DADOS ─────────────────────────────────────────────────────────
def build_plan(plano, ep, pp):
    grp_keys = ['cd_posicao','nr_produto','cd_lote']

    # Primeira contagem (A) → saldo original do sistema
    first = (ep.sort_values(['cd_posicao','nr_produto','cd_lote','rank'])
               .groupby(grp_keys, dropna=False).first().reset_index()
               .rename(columns={'qt_saldoinicial':'saldo_sis','cd_contagem':'first_cnt'}))

    # Última contagem (C>B>A) → resultado final contado
    last = (ep.sort_values(['cd_posicao','nr_produto','cd_lote','rank'], ascending=[True,True,True,False])
              .groupby(grp_keys, dropna=False).first().reset_index()
              .rename(columns={'qt_contagem':'cnt_final','cd_contagem':'last_cnt',
                               'ds_statuscontagem':'status_final'}))

    base_cols_f = grp_keys + ['saldo_sis','first_cnt','ds_produto']
    base_cols_l = grp_keys + ['cnt_final','last_cnt','status_final',
                               'nr_produtocontagem','ds_produtocontagem','cd_lotecontagem',
                               'COD X FIS','LOTE X FIS','ds_usuario']
    pos = first[base_cols_f].merge(last[base_cols_l], on=grp_keys, how='outer')
    pos['saldo_sis'] = pos['saldo_sis'].fillna(0)
    pos['cnt_final'] = pos['cnt_final'].fillna(0)
    pos['net_div']   = (pos['cnt_final'] - pos['saldo_sis']).round(2)
    pos['is_fin_s']  = pos['status_final'] == 'FINALIZADO S/ DIVERGÊNCIA'
    pos['is_fin_c']  = pos['status_final'] == 'FINALIZADO C/ DIVERGÊNCIA'
    pos['is_pend']   = pos['status_final'] == 'PENDENTE'
    pos['is_em_c']   = pos['status_final'] == 'EM CONTAGEM'

    # Inversões por posição distinta
    last_by_pos = (ep.sort_values(['cd_posicao','rank'], ascending=[True,False])
                     .groupby('cd_posicao').first().reset_index())
    inv_cod_pos = last_by_pos[last_by_pos['COD X FIS']=='ERRADO'].copy()
    inv_lot_pos = last_by_pos[(last_by_pos['LOTE X FIS']=='ERRADO')&(last_by_pos['COD X FIS']=='CERTO')].copy()

    first_by_pos = (ep.sort_values(['cd_posicao','rank'])
                      .groupby('cd_posicao').first().reset_index()
                      .rename(columns={'qt_saldoinicial':'saldo_sis_pos'}))

    def make_inv_recs(df, include_prod_cnt=True):
        df2 = df.merge(first_by_pos[['cd_posicao','saldo_sis_pos']], on='cd_posicao', how='left')
        recs = []
        for _, r in df2.iterrows():
            rec = {
                'cd_posicao': str(r.get('cd_posicao','—')),
                'nr_produto':  s(r.get('nr_produto'),'—'),
                'ds_produto':  str(r.get('ds_produto','—') or '—'),
                'cd_lote':     str(r.get('cd_lote','—') or '—'),
                'qt_saldoinicial': float(s(r.get('saldo_sis_pos'),0)),
                'qt_contagem':     float(s(r.get('qt_contagem'),0)),
            }
            if include_prod_cnt:
                rec['nr_produtocontagem'] = s(r.get('nr_produtocontagem'),'—')
                rec['ds_produtocontagem'] = str(r.get('ds_produtocontagem','—') or '—')
                rec['cd_lotecontagem']    = str(r.get('cd_lotecontagem','—') or '—')
            else:
                rec['cd_lotecontagem'] = str(r.get('cd_lotecontagem','—') or '—')
            recs.append(rec)
        return recs

    inv_cod_recs = make_inv_recs(inv_cod_pos, True)
    inv_lot_recs = make_inv_recs(inv_lot_pos, False)

    # Validação de lotes
    sg = pos.groupby(['nr_produto','cd_lote'], dropna=False).agg(
        n_pos=('cd_posicao','count'), n_ok=('is_fin_s','sum'),
        ds_produto=('ds_produto','first'),
        sum_saldo=('saldo_sis','sum'), sum_cnt=('cnt_final','sum'),
    ).reset_index()
    sg['all_ok']  = sg['n_pos'] == sg['n_ok']
    sg['net_div'] = (sg['sum_cnt'] - sg['sum_saldo']).round(2)
    validated = sg[sg['all_ok'] & (sg['net_div'].abs() < 0.01)].copy()

    def classify(row):
        sal=float(row['sum_saldo']); cnt=float(row['sum_cnt']); div=float(row['net_div'])
        if cnt == 0 and sal > 0:  return 'Posição vazia/contagem=0'
        if sal == 0 and cnt == 0: return 'Posição vazia correta'
        pct = abs(div/sal)*100 if sal > 0 else 999
        ratio = cnt/sal if sal > 0 else 0
        rnd = round(ratio)
        if rnd > 1 and abs(ratio-rnd) < 0.05 and rnd in [2,3,4,5,6,10,12,25,50,60,100]:
            return f'Erro de UM (fator x{rnd:.0f})'
        if pct < 10: return 'Divergência pequena (<10%)'
        if div > 0:  return 'Sobra significativa'
        return 'Falta significativa'

    lot_alert = sg[sg['all_ok'] & (sg['net_div'].abs() >= 0.01)].copy()
    lot_alert['tipo']    = lot_alert.apply(classify, axis=1)
    lot_alert['pct_div'] = lot_alert.apply(
        lambda r: round(float(r['net_div'])/float(r['sum_saldo'])*100,1) if float(r['sum_saldo'])>0 else 0, axis=1)
    pos_list = pos.groupby(['nr_produto','cd_lote'])['cd_posicao'].apply(lambda x: ', '.join(sorted(x.unique()))).reset_index()
    pos_list.columns = ['nr_produto','cd_lote','posicoes']
    lot_alert = lot_alert.merge(pos_list, on=['nr_produto','cd_lote'], how='left')

    lot_alert_recs = sorted([{
        'nr_produto': str(s(r.get('nr_produto'),'—')),
        'ds_produto': str(r.get('ds_produto','—') or '—'),
        'cd_lote':    str(r.get('cd_lote','—') or '—'),
        'n_pos': int(r.get('n_pos',0)),
        'sum_saldo':     float(r['sum_saldo']),
        'sum_contagem':  float(r['sum_cnt']),
        'net_div': float(r['net_div']),
        'pct_div': float(r['pct_div']),
        'tipo':    str(r['tipo']),
        'posicoes': str(r.get('posicoes',''))[:150],
    } for _, r in lot_alert.iterrows()], key=lambda x: abs(x['net_div']), reverse=True)

    sku_ok_recs = [{
        'nr_produto': str(s(r.get('nr_produto'),'—')),
        'ds_produto': str(r.get('ds_produto','—') or '—'),
        'cd_lote':    str(r.get('cd_lote','—') or '—'),
        'tot': int(r.get('n_pos',1)),
        'sum_saldo':    float(r['sum_saldo']),
        'sum_contagem': float(r['sum_cnt']),
    } for _, r in validated.sort_values('nr_produto').iterrows()]

    # Posições pendentes por nível de contagem (contagem distinta)
    def err_cnt_stats(cnt_level):
        sub = ep[ep['cd_contagem'] == cnt_level]
        return {
            'pend_total': int(sub[sub['ds_statuscontagem']=='PENDENTE']['cd_posicao'].nunique()),
            'em_cnt':     int(sub[sub['ds_statuscontagem']=='EM CONTAGEM']['cd_posicao'].nunique()),
            'fin_s':      int(sub[sub['ds_statuscontagem']=='FINALIZADO S/ DIVERGÊNCIA']['cd_posicao'].nunique()),
            'fin_c':      int(sub[sub['ds_statuscontagem']=='FINALIZADO C/ DIVERGÊNCIA']['cd_posicao'].nunique()),
            'total':      int(sub['cd_posicao'].nunique()),
        }

    err_cnt_A = err_cnt_stats('A')
    err_cnt_B = err_cnt_stats('B')
    err_cnt_C = err_cnt_stats('C')

    def port_cnt_stats(cnt_letter):
        if len(pp) == 0: return {'pend':0,'conc':0,'total':0}
        pend = pp[pp['statusatual'].str.startswith(f'{cnt_letter} - Em Andamento', na=False)]['cd_posicao'].nunique()
        conc = pp[pp['statusatual'].str.startswith(f'{cnt_letter} - Concluído', na=False)]['cd_posicao'].nunique()
        return {'pend':int(pend),'conc':int(conc),'total':int(pend+conc)}

    port_a = port_cnt_stats('A'); port_b = port_cnt_stats('B'); port_c = port_cnt_stats('C')

    # Pareto
    pos_div = pos[pos['net_div'].abs() > 0].copy()
    pos_div['abs_div'] = pos_div['net_div'].abs()
    pareto = []
    for _, r in pos_div.sort_values('abs_div', ascending=False).head(20).iterrows():
        pareto.append({
            'cd_posicao': str(r.get('cd_posicao','—')),
            'ds_produto': str(r.get('ds_produto','—') or '—'),
            'cd_lote':    str(r.get('cd_lote','—') or '—'),
            'saldo':    float(r['saldo_sis']),
            'contagem': float(r['cnt_final']),
            'net_div':  float(r['net_div']),
            'status':   str(r.get('status_final','—') or '—'),
        })

    # Contadores por operador
    ev = ep[ep['dur_min'].between(0,480)].copy()
    ctd = {}
    for c in ['A','B','C']:
        sub = ep[(ep['cd_contagem']==c) & ep['dur_min'].between(0,480)]
        grp = sub.groupby('ds_usuario').agg(
            n=('cd_posicao','count'), dur=('dur_min','sum'), avg=('dur_min','mean')
        ).reset_index().sort_values('n', ascending=False).head(15)
        ctd[c] = [{'user':str(r['ds_usuario']),'n':int(r['n']),'dur':round(float(r['dur']),1),'avg':round(float(r['avg']),2)} for _,r in grp.iterrows()]

    # Tempo por dimensão de endereço
    def time_by_dim(df, dim, cnt_col='cd_contagem', dur_col='dur_min'):
        recs = []
        for (dv,cnt), g in df.groupby([dim,cnt_col]):
            recs.append({'dim_val':dv,'cnt':cnt,'media':round(g[dur_col].mean(),2),
                         'mediana':round(g[dur_col].median(),2),'n':int(len(g))})
        return recs

    time_err = {d: time_by_dim(ev,d) for d in ['area','rua','predio','nivel']}

    time_port = {}
    for dim in ['area','rua','predio','nivel']:
        recs = []
        for c in ['a','b','c']:
            dur_col = f'dur_{c}_min'; ini_col = f'dt_contagem{c}_ini'
            sub = pp[pp[dur_col].between(0,480)].drop_duplicates(subset=['cd_posicao',ini_col]) if len(pp)>0 else pd.DataFrame()
            for dv, g in (sub.groupby(dim) if len(sub)>0 else []):
                recs.append({'dim_val':dv,'cnt':c.upper(),'media':round(g[dur_col].mean(),2),
                             'mediana':round(g[dur_col].median(),2),'n':int(len(g))})
        time_port[dim] = recs

    # Estimativa de tempo para pendentes
    def pending_time_estimate_err():
        estimates = []
        for cnt in ['A','B','C']:
            sub_done = ev[ev['cd_contagem']==cnt]
            sub_pend = ep[(ep['cd_contagem']==cnt) & (ep['ds_statuscontagem'].isin(['PENDENTE','EM CONTAGEM']))]
            for rua, rg in sub_pend.groupby('rua'):
                n_pend = int(rg['cd_posicao'].nunique())
                med = sub_done[sub_done['rua']==rua]['dur_min'].median()
                if pd.isna(med): med = sub_done['dur_min'].median()
                if pd.isna(med): med = 2.0
                estimates.append({'rua':rua,'cnt':cnt,'n_pend':n_pend,
                                   'med_min':round(float(med),2),
                                   'est_min':round(float(med)*n_pend,1),
                                   'est_h':round(float(med)*n_pend/60,2)})
        return estimates

    def pending_time_estimate_port():
        estimates = []
        for c in ['a','b','c']:
            dur_col = f'dur_{c}_min'; ini_col = f'dt_contagem{c}_ini'
            sub_done = pp[pp[dur_col].between(0,480)].drop_duplicates(subset=['cd_posicao',ini_col]) if len(pp)>0 else pd.DataFrame()
            pend_status = f'{c.upper()} - Em Andamento'
            sub_pend = pp[pp['statusatual'].str.startswith(pend_status, na=False)] if len(pp)>0 else pd.DataFrame()
            for rua, rg in (sub_pend.groupby('rua') if len(sub_pend)>0 else []):
                n_pend = int(rg['cd_posicao'].nunique())
                med = sub_done[sub_done['rua']==rua][dur_col].median() if len(sub_done)>0 else np.nan
                if pd.isna(med) and len(sub_done)>0: med = sub_done[dur_col].median()
                if pd.isna(med): med = 2.0
                estimates.append({'rua':rua,'cnt':c.upper(),'n_pend':n_pend,
                                   'med_min':round(float(med),2),
                                   'est_min':round(float(med)*n_pend,1),
                                   'est_h':round(float(med)*n_pend/60,2)})
        return estimates

    pend_est_err  = pending_time_estimate_err()
    pend_est_port = pending_time_estimate_port()

    # Histogramas de tempo
    BINS = [0,1,2,5,15,30,480]
    hists_err = {}
    for c in ['A','B','C']:
        sub = ev[ev['cd_contagem']==c]['dur_min']
        if len(sub)>0:
            cuts = pd.cut(sub, bins=BINS, labels=BIN_LBL, right=False)
            hists_err[c] = cuts.value_counts().reindex(BIN_LBL).fillna(0).astype(int).tolist()

    hists_port = {}
    for c in ['a','b','c']:
        dur_col = f'dur_{c}_min'; ini_col = f'dt_contagem{c}_ini'
        sub = pp[pp[dur_col].between(0,480)].drop_duplicates(subset=['cd_posicao',ini_col]) if len(pp)>0 else pd.DataFrame()
        if len(sub)>0:
            cuts = pd.cut(sub[dur_col], bins=BINS, labels=BIN_LBL, right=False)
            hists_port[c.upper()] = cuts.value_counts().reindex(BIN_LBL).fillna(0).astype(int).tolist()

    # Estatísticas gerais de tempo
    overall_port = []
    for c in ['a','b','c']:
        dur_col = f'dur_{c}_min'; ini_col = f'dt_contagem{c}_ini'
        sub = pp[pp[dur_col].between(0,480)].drop_duplicates(subset=['cd_posicao',ini_col]) if len(pp)>0 else pd.DataFrame()
        if len(sub)>0:
            overall_port.append({'cnt':c.upper(),'media':round(float(sub[dur_col].mean()),2),
                                  'mediana':round(float(sub[dur_col].median()),2),
                                  'n':int(len(sub)),'total_h':round(float(sub[dur_col].sum()/60),1),
                                  'p25':round(float(sub[dur_col].quantile(.25)),2),
                                  'p75':round(float(sub[dur_col].quantile(.75)),2)})

    overall_err = []
    for c in ['A','B','C']:
        sub = ev[ev['cd_contagem']==c]
        if len(sub)>0:
            overall_err.append({'cnt':c,'media':round(float(sub['dur_min'].mean()),2),
                                 'mediana':round(float(sub['dur_min'].median()),2),
                                 'n':int(len(sub)),'total_h':round(float(sub['dur_min'].sum()/60),1),
                                 'p25':round(float(sub['dur_min'].quantile(.25)),2),
                                 'p75':round(float(sub['dur_min'].quantile(.75)),2)})

    return {
        'nome': plano,
        'n_pos_total': int(ep['cd_posicao'].nunique()),
        'n_fin_s': int(pos['is_fin_s'].sum()), 'n_fin_c': int(pos['is_fin_c'].sum()),
        'n_pend': int(pos['is_pend'].sum()),   'n_em_c': int(pos['is_em_c'].sum()),
        'n_inv_cod': len(inv_cod_recs), 'n_inv_lot': len(inv_lot_recs),
        'inv_cod': inv_cod_recs[:30], 'inv_lot': inv_lot_recs[:50],
        'n_sku_ok': int(validated['nr_produto'].nunique()), 'n_lot_ok': int(len(validated)),
        'n_lot_alert': len(lot_alert_recs),
        'n_sku_total': int(ep['nr_produto'].nunique()),
        'n_lot_total': int(ep[['nr_produto','cd_lote']].drop_duplicates().shape[0]),
        'sku_ok': sku_ok_recs, 'lot_alert': lot_alert_recs,
        'err_cnt_A': err_cnt_A, 'err_cnt_B': err_cnt_B, 'err_cnt_C': err_cnt_C,
        'port_pend_a': port_a['pend'], 'port_conc_a': port_a['conc'],
        'port_pend_b': port_b['pend'], 'port_conc_b': port_b['conc'],
        'port_pend_c': port_c['pend'], 'port_conc_c': port_c['conc'],
        'port_tot_pos': int(pp['cd_posicao'].nunique()) if len(pp)>0 else 0,
        'pareto': pareto, 'ctd_A': ctd.get('A',[]), 'ctd_B': ctd.get('B',[]), 'ctd_C': ctd.get('C',[]),
        'time_err': time_err, 'time_port': time_port,
        'hists_err': hists_err, 'hists_port': hists_port,
        'overall_err': overall_err, 'overall_port': overall_port,
        'pend_est_err': pend_est_err, 'pend_est_port': pend_est_port,
    }

# ── HTML BUILDERS ────────────────────────────────────────────────────────────
BASE_CSS = '''
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}
.hdr{padding:16px 26px;border-bottom:1px solid #1e293b;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.hdr-left h1{font-size:19px;font-weight:700;color:#fff}
.hdr-left p{color:#94a3b8;font-size:11px;margin-top:2px}
.hdr-badge{font-size:11px;background:#14532d;color:#4ade80;padding:4px 12px;border-radius:20px;border:1px solid #22c55e}
.tabs{display:flex;gap:5px;padding:10px 20px;background:#0a1628;border-bottom:1px solid #1e293b;flex-wrap:wrap}
.tab-btn{padding:6px 16px;border-radius:6px;border:1px solid #334155;background:#1e293b;color:#94a3b8;cursor:pointer;font-size:12px;font-weight:500;transition:all .2s}
.tab-btn:hover,.tab-btn.active{background:#1d4ed8;border-color:#3b82f6;color:#fff}
.content{padding:16px 20px;max-width:1350px;margin:0 auto}
.card{background:#1e293b;border-radius:10px;padding:14px 16px;margin-bottom:12px;border:1px solid #334155}
.card h3{font-size:13px;font-weight:600;color:#cbd5e1;margin-bottom:9px}
.kpi-row{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
.kpi{flex:1;min-width:100px;background:#0f172a;border-radius:7px;padding:10px 12px;border-top:3px solid #334155;text-align:center}
.kpi label{font-size:9px;color:#64748b;text-transform:uppercase;letter-spacing:.06em;display:block;margin-bottom:3px}
.kpi .val{font-size:20px;font-weight:700}
.scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#0f172a;color:#64748b;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;padding:6px 9px;border-bottom:2px solid #334155;white-space:nowrap}
td{padding:5px 9px;border-bottom:1px solid #182030;white-space:nowrap}
tr:hover td{background:#172033}
.badge-tag{display:inline-block;padding:2px 6px;border-radius:10px;font-size:10px;font-weight:700}
.bt-warn{background:#78350f33;color:#fbbf24;border:1px solid #78350f}
.bt-err{background:#7f1d1d33;color:#f87171;border:1px solid #7f1d1d}
.bt-err2{background:#4c1d9533;color:#c084fc;border:1px solid #4c1d95}
.bt-gray{background:#1e293b;color:#94a3b8;border:1px solid #334155}
.bt-ok{background:#14532d33;color:#4ade80;border:1px solid #14532d}
.bt-purp{background:#5b21b633;color:#c084fc;border:1px solid #5b21b6}
.progress-bar{height:8px;background:#0f172a;border-radius:4px;overflow:hidden;margin:3px 0}
.progress-fill{height:100%;border-radius:4px}
.addr-box{background:#0f172a;border:1px solid #334155;border-radius:7px;padding:8px 12px;margin-bottom:12px;display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:11px;color:#94a3b8}
.addr-tag{background:#1e293b;border-radius:3px;padding:2px 7px}
.addr-tag span{color:#60a5fa;font-weight:700}
.legend-row{display:flex;flex-wrap:wrap;gap:10px;font-size:11px;margin-bottom:12px}
.ldot{width:9px;height:9px;border-radius:50%;display:inline-block;margin-right:3px}
details summary{cursor:pointer;font-size:11px;color:#6b7280;user-select:none;margin-top:7px}
details summary:hover{color:#e2e8f0}
.corr-note{background:#14532d22;border:1px solid #22c55e44;border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:12px;color:#94a3b8;line-height:1.6}
'''

JS_PLAN = '''
var _planPrefix='plan';
function showPlan(pid){
  document.querySelectorAll('[data-plan="'+_planPrefix+'"]').forEach(e=>e.style.display='none');
  document.querySelectorAll('[data-tab="'+_planPrefix+'"]').forEach(e=>e.classList.remove('active'));
  var b=document.getElementById(pid); if(b)b.style.display='block';
  var t=document.getElementById('tab-'+_planPrefix+'-'+pid); if(t)t.classList.add('active');
}
window.addEventListener('DOMContentLoaded',function(){
  var first=document.querySelector('[data-plan="'+_planPrefix+'"]');
  if(first)first.style.display='block';
  var ft=document.querySelector('[data-tab="'+_planPrefix+'"]');
  if(ft)ft.classList.add('active');
});
'''

def build_kpi_row(D, source):
    def kpi(lbl,val,col,bc,id_):
        return f'<div class="kpi" style="border-top-color:{bc}" id="kpi_{id_}"><label>{lbl}</label><div class="val" style="color:{col}">{val}</div></div>'
    if source == 'err':
        items = [
            ('Fin.S/Div',fmt_n(D['n_fin_s']),'#4ade80','#22c55e','fin_s'),
            ('Fin.C/Div',fmt_n(D['n_fin_c']),'#f87171','#ef4444','fin_c'),
            ('Pendente',fmt_n(D['n_pend']),'#fbbf24','#f59e0b','pend'),
            ('Em Contagem',fmt_n(D['n_em_c']),'#60a5fa','#3b82f6','em_c'),
            ('Lotes Válidos',fmt_n(D['n_lot_ok']),'#4ade80','#22c55e','lot_ok'),
            ('SKUs Válidos',fmt_n(D['n_sku_ok']),'#60a5fa','#3b82f6','sku_ok'),
            ('Alertas Qtd',fmt_n(D['n_lot_alert']),'#f97316','#ea580c','lot_alert'),
            ('Inv.Código',fmt_n(D['n_inv_cod']),'#f87171','#ef4444','inv_cod'),
        ]
    else:
        items = [
            ('Fin.S/Div',fmt_n(D['n_fin_s']),'#4ade80','#22c55e','fin_s'),
            ('Fin.C/Div',fmt_n(D['n_fin_c']),'#f87171','#ef4444','fin_c'),
            ('Pendente',fmt_n(D['n_pend']),'#fbbf24','#f59e0b','pend'),
            ('Pend.A',fmt_n(D['port_pend_a']),'#60a5fa','#1e3a5f','pend_a'),
            ('Pend.B',fmt_n(D['port_pend_b']),'#4ade80','#14532d','pend_b'),
            ('Pend.C',fmt_n(D['port_pend_c']),'#c084fc','#4c1d95','pend_c'),
            ('Lotes Válidos',fmt_n(D['n_lot_ok']),'#4ade80','#22c55e','lot_ok'),
            ('Alertas Qtd',fmt_n(D['n_lot_alert']),'#f97316','#ea580c','lot_alert'),
        ]
    return f'<div class="kpi-row" id="kpi_main">{"".join(kpi(*i) for i in items)}</div>'

def pend_blk_err(cA, cB, cC):
    def blk(c, lbl, bg, fg):
        tot=c['total']; fin=c['fin_s']+c['fin_c']; pend=c['pend_total']; em=c['em_cnt']
        pct = round(fin/tot*100,1) if tot else 0
        return f'''<div style="flex:1;min-width:170px;background:{bg};border:1px solid {fg}33;border-radius:8px;padding:11px 13px">
  <div style="font-size:10px;color:{fg};font-weight:700;letter-spacing:.06em;margin-bottom:5px">{lbl}</div>
  <div style="font-size:9px;color:#6b7280;margin-bottom:6px">{fmt_n(tot)} posições distintas</div>
  <div class="progress-bar"><div class="progress-fill" style="width:{pct}%;background:{fg}"></div></div>
  <div style="font-size:9px;color:#9ca3af;margin-top:3px">{pct}% concluídas</div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-top:7px;font-size:10px">
    <div><div style="color:#6b7280">Fin.S/Div</div><div style="color:#4ade80;font-weight:700">{fmt_n(fin)}</div></div>
    <div><div style="color:#6b7280">Pendente</div><div style="color:#f87171;font-weight:700">{fmt_n(pend)}</div></div>
    <div><div style="color:#6b7280">Em Contagem</div><div style="color:#fbbf24;font-weight:700">{fmt_n(em)}</div></div>
    <div><div style="color:#6b7280">Total</div><div style="color:{fg};font-weight:700">{fmt_n(tot)}</div></div>
  </div>
</div>'''
    return (f'<div class="card"><h3>📊 Posições Pendentes por Contagem — Contagem Distinta</h3>'
            f'<div style="display:flex;flex-wrap:wrap;gap:10px">'
            f'{blk(cA,"🔵 Contagem A","#1e3a5f","#60a5fa")}'
            f'{blk(cB,"🟡 Contagem B","#14532d","#4ade80")}'
            f'{blk(cC,"🔴 Contagem C","#4c1d95","#c084fc")}'
            f'</div></div>')

def pend_blk_port(D):
    def blk(cnt, pend, conc, bg, fg):
        tot = pend+conc; pct = round(conc/tot*100,1) if tot else 0
        return f'''<div style="flex:1;min-width:170px;background:{bg};border:1px solid {fg}33;border-radius:8px;padding:11px 13px">
  <div style="font-size:10px;color:{fg};font-weight:700;letter-spacing:.06em;margin-bottom:4px">Contagem {cnt}</div>
  <div class="progress-bar"><div class="progress-fill" style="width:{pct}%;background:{fg}"></div></div>
  <div style="font-size:9px;color:#9ca3af;margin-top:3px">{pct}% concluídas</div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-top:7px;font-size:10px">
    <div><div style="color:#6b7280">Concluído</div><div style="color:#4ade80;font-weight:700">{fmt_n(conc)}</div></div>
    <div><div style="color:#6b7280">Em Andamento</div><div style="color:#f87171;font-weight:700">{fmt_n(pend)}</div></div>
  </div>
</div>'''
    return (f'<div class="card"><h3>📊 Posições por Contagem — Contagem Distinta (Portal)</h3>'
            f'<div style="display:flex;flex-wrap:wrap;gap:10px">'
            f'{blk("A",D["port_pend_a"],D["port_conc_a"],"#1e3a5f","#60a5fa")}'
            f'{blk("B",D["port_pend_b"],D["port_conc_b"],"#14532d","#4ade80")}'
            f'{blk("C",D["port_pend_c"],D["port_conc_c"],"#4c1d95","#c084fc")}'
            f'</div></div>')

def alert_table(recs):
    if not recs: return '<p style="color:#555;font-size:12px">Sem alertas de quantidade.</p>'
    rows = ''
    for r in recs[:100]:
        d  = str(r.get('ds_produto','') or '—')[:26]
        nd = float(r.get('net_div',0)); nc = f'+{nd:,.0f}' if nd>0 else f'{nd:,.0f}'
        nc_col = '#f87171' if nd<0 else '#4ade80'
        pct = float(r.get('pct_div',0))
        tipo = str(r.get('tipo','')); cls = TIPO_CLS.get(tipo,'bt-warn')
        if 'Erro de UM' in tipo: cls = 'bt-purp'
        ss = float(r.get('sum_saldo',0) or 0); sc = float(r.get('sum_contagem',0) or 0)
        pos = str(r.get('posicoes',''))[:45]
        rows += (f'<tr><td style="color:#94a3b8;font-size:10px">{r.get("nr_produto","")}</td>'
                 f'<td title="{r.get("ds_produto","")}" style="max-width:130px;overflow:hidden;text-overflow:ellipsis">{d}</td>'
                 f'<td style="color:#f87171">{r.get("cd_lote","")}</td>'
                 f'<td style="text-align:right">{ss:,.0f}</td><td style="text-align:right">{sc:,.0f}</td>'
                 f'<td style="color:{nc_col};font-weight:700;text-align:right">{nc}</td>'
                 f'<td style="color:#6b7280;text-align:right">{pct:+.1f}%</td>'
                 f'<td><span class="badge-tag {cls}" style="font-size:9px">{tipo}</span></td>'
                 f'<td style="font-size:10px;color:#475569" title="{r.get("posicoes","")}">{pos}</td></tr>')
    extra = f'<p style="font-size:10px;color:#6b7280;margin-top:5px">Mostrando 100 de {len(recs)}</p>' if len(recs)>100 else ''
    return (f'<div class="scroll"><table><thead><tr>'
            f'<th>SKU</th><th>Produto</th><th>Lote</th><th>Qt.Sist.(A)</th><th>Qt.Cont.(Últ.)</th>'
            f'<th>Div.Líq.</th><th>%</th><th>Tipo</th><th>Posição(ões)</th>'
            f'</tr></thead><tbody>{rows}</tbody></table></div>{extra}')

def sku_ok_table(recs):
    if not recs: return '<p style="color:#555;font-size:12px">Sem lotes validados.</p>'
    rows = ''
    for r in recs[:200]:
        d = str(r.get('ds_produto','') or '—')[:30]
        rows += (f'<tr><td style="color:#94a3b8;font-size:10px">{r.get("nr_produto","")}</td>'
                 f'<td title="{r.get("ds_produto","")}">{d}</td>'
                 f'<td style="color:#4ade80">{r.get("cd_lote","")}</td>'
                 f'<td style="text-align:center;color:#64748b">{r.get("tot",1)}</td>'
                 f'<td style="text-align:right;color:#4ade80">{float(r.get("sum_saldo",0)):,.0f}</td>'
                 f'<td style="text-align:right;color:#60a5fa">{float(r.get("sum_contagem",0)):,.0f}</td></tr>')
    extra = f'<p style="font-size:10px;color:#6b7280;margin-top:5px">Mostrando 200 de {len(recs)}</p>' if len(recs)>200 else ''
    return (f'<div class="scroll"><table><thead><tr>'
            f'<th>SKU</th><th>Produto</th><th>Lote</th><th>Pos.</th><th>Qt.Sist.(A)</th><th>Qt.Cont.(Últ.)</th>'
            f'</tr></thead><tbody>{rows}</tbody></table></div>{extra}')

def inv_table(recs, mode='cod'):
    if not recs: return '<p style="color:#555;font-size:12px">Sem inversões.</p>'
    rows = ''
    for r in recs:
        ss = float(r.get('qt_saldoinicial',0) or 0); sc = float(r.get('qt_contagem',0) or 0)
        d = str(r.get('ds_produto','') or '—')[:20]
        if mode == 'cod':
            dc = str(r.get('ds_produtocontagem','') or '—')[:20]
            rows += (f'<tr><td style="color:#60a5fa;font-size:10px">{r.get("cd_posicao","")}</td>'
                     f'<td title="{r.get("ds_produto","")}">{d}</td>'
                     f'<td style="color:#f87171">{str(r.get("cd_lote",""))}</td>'
                     f'<td style="color:#fbbf24;font-size:10px">{str(r.get("nr_produtocontagem",""))}</td>'
                     f'<td title="{r.get("ds_produtocontagem","")}" style="color:#fbbf24">{dc}</td>'
                     f'<td style="color:#fbbf24">{str(r.get("cd_lotecontagem",""))}</td>'
                     f'<td style="text-align:right">{ss:,.0f}</td>'
                     f'<td style="text-align:right;color:#f97316">{sc:,.0f}</td></tr>')
        else:
            rows += (f'<tr><td style="color:#60a5fa;font-size:10px">{r.get("cd_posicao","")}</td>'
                     f'<td>{d}</td><td style="color:#f87171">{str(r.get("cd_lote",""))}</td>'
                     f'<td style="color:#fbbf24">{str(r.get("cd_lotecontagem",""))}</td>'
                     f'<td style="text-align:right">{ss:,.0f}</td>'
                     f'<td style="text-align:right;color:#f97316">{sc:,.0f}</td></tr>')
    hdr = ('<th>End.</th><th>Prod.Esp.</th><th>Lote Esp.</th><th>SKU Cont.</th><th>Prod.Cont.</th><th>Lote Cont.</th><th>Qt.Sist.(A)</th><th>Qt.Cont.(Últ.)</th>'
           if mode=='cod' else
           '<th>End.</th><th>Produto</th><th>Lote Esp.</th><th>Lote Cont.</th><th>Qt.Sist.(A)</th><th>Qt.Cont.(Últ.)</th>')
    return f'<div class="scroll"><table><thead><tr>{hdr}</tr></thead><tbody>{rows}</tbody></table></div>'

def dim_time_table(recs, lbl):
    if not recs: return ''
    pivot = {}
    for r in recs:
        dv = r['dim_val']
        if dv not in pivot: pivot[dv] = {}
        pivot[dv][r['cnt']] = r
    rows = ''
    for dv in sorted(pivot.keys()):
        row = f'<td style="font-weight:600;color:#e2e8f0">{lbl} {dv}</td>'
        for cnt in ['A','B','C']:
            _,fg = CNT_STYLE.get(cnt,('#1e293b','#e2e8f0'))
            if cnt not in pivot[dv]:
                row += '<td colspan="3" style="color:#374151;text-align:center">—</td>'
            else:
                r = pivot[dv][cnt]; med = r.get('mediana',0) or 0; col = col_for(med)
                row += f'<td style="color:{col};font-weight:700">{fmt_m(r.get("media",0))}</td><td style="color:{col}">{fmt_m(med)}</td><td style="color:#6b7280;font-size:10px">{fmt_n(r.get("n",0))}</td>'
        rows += f'<tr>{row}</tr>'
    return (f'<div class="scroll"><table>'
            f'<thead><tr><th rowspan="2">{lbl}</th>'
            f'<th colspan="3" style="background:#1e3a5f;color:#60a5fa">Cnt A</th>'
            f'<th colspan="3" style="background:#14532d;color:#4ade80">Cnt B</th>'
            f'<th colspan="3" style="background:#4c1d95;color:#c084fc">Cnt C</th></tr>'
            f'<tr><th>Média</th><th>Mediana</th><th>N</th><th>Média</th><th>Mediana</th><th>N</th><th>Média</th><th>Mediana</th><th>N</th></tr></thead>'
            f'<tbody>{rows}</tbody></table></div>')

def heatmap_rua(recs):
    pivot = {}
    for r in recs:
        dv = r['dim_val']
        if dv not in pivot: pivot[dv] = {}
        pivot[dv][r['cnt']] = r
    if not pivot: return ''
    hdr = '<div style="display:grid;grid-template-columns:70px repeat(3,88px);gap:3px;margin-bottom:5px"><div></div>'
    for c,(_, fg) in [('A',CNT_STYLE['A']),('B',CNT_STYLE['B']),('C',CNT_STYLE['C'])]:
        hdr += f'<div style="text-align:center;font-size:10px;color:{fg};font-weight:700">Cnt {c}</div>'
    hdr += '</div>'
    cells = ''
    for rua in sorted(pivot.keys()):
        row = f'<div style="font-size:9px;color:#9ca3af;text-align:center;padding:2px">Rua<br><strong style="color:#e2e8f0;font-size:11px">{rua}</strong></div>'
        for cnt in ['A','B','C']:
            _,fg = CNT_STYLE[cnt]
            if cnt not in pivot[rua]:
                row += f'<div style="background:#0f172a;border-radius:3px;text-align:center;padding:3px;font-size:10px;color:#374151">—</div>'
            else:
                r = pivot[rua][cnt]; med = r.get('mediana',0) or 0
                col = col_for(med); bg = col+'22'
                row += f'<div style="background:{bg};border:1px solid {col}44;border-radius:3px;text-align:center;padding:3px"><div style="font-size:12px;font-weight:700;color:{col}">{fmt_m(med)}</div><div style="font-size:9px;color:#6b7280">n={fmt_n(r.get("n",0))}</div></div>'
        cells += f'<div style="display:grid;grid-template-columns:70px repeat(3,88px);gap:3px;margin-bottom:3px">{row}</div>'
    return f'<div style="overflow-x:auto"><div style="min-width:350px">{hdr}{cells}</div></div>'

def hist_bars(hists):
    out = ''
    for cnt in ['A','B','C']:
        vals = hists.get(cnt,[])
        if not vals: continue
        total = sum(vals)
        if not total: continue
        _,fg = CNT_STYLE.get(cnt,('#1e293b','#e2e8f0'))
        bars = ''
        for i,(lbl,val) in enumerate(zip(BIN_LBL,vals)):
            pct = val/total*100
            bars += (f'<div style="display:flex;align-items:center;gap:5px;margin:2px 0">'
                     f'<div style="width:58px;font-size:9px;color:#9ca3af;text-align:right;flex-shrink:0">{lbl}</div>'
                     f'<div style="flex:1;background:#0f172a;border-radius:2px;height:12px">'
                     f'<div style="width:{pct:.1f}%;background:{HIST_COLORS[i]};height:100%;border-radius:2px"></div></div>'
                     f'<div style="width:65px;font-size:9px;color:#d1d5db;flex-shrink:0">{fmt_n(val)} ({pct:.0f}%)</div></div>')
        out += f'<div style="flex:1;min-width:190px"><div style="font-size:10px;color:{fg};font-weight:700;margin-bottom:4px">Cnt {cnt} · {fmt_n(total)}</div>{bars}</div>'
    return f'<div style="display:flex;flex-wrap:wrap;gap:18px">{out}</div>'

def kpi_cnt_block(overall):
    cards = ''
    for o in overall:
        cnt = o['cnt']; bg,fg = CNT_STYLE.get(cnt,('#1e293b','#e2e8f0'))
        med = o.get('mediana',0) or 0; col = col_for(med)
        cards += (f'<div style="flex:1;min-width:170px;background:{bg};border:1px solid {fg}33;border-radius:8px;padding:12px 14px">'
                  f'<div style="font-size:10px;color:{fg};font-weight:700;letter-spacing:.06em">CONTAGEM {cnt}</div>'
                  f'<div style="font-size:26px;font-weight:700;color:{col};margin:4px 0">{fmt_m(med)}</div>'
                  f'<div style="font-size:9px;color:#6b7280">mediana por posição</div>'
                  f'<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px;margin-top:7px;font-size:10px">'
                  f'<div style="text-align:center"><div style="color:#6b7280">Média</div><div style="color:{fg};font-weight:600">{fmt_m(o.get("media",0))}</div></div>'
                  f'<div style="text-align:center;border-left:1px solid #334155;border-right:1px solid #334155"><div style="color:#6b7280">P25–P75</div><div style="color:{fg};font-weight:600">{fmt_m(o.get("p25",0))}–{fmt_m(o.get("p75",0))}</div></div>'
                  f'<div style="text-align:center"><div style="color:#6b7280">Posições</div><div style="color:{fg};font-weight:600">{fmt_n(o.get("n",0))}</div></div>'
                  f'</div><div style="text-align:center;font-size:9px;color:#6b7280;margin-top:4px">Total: {o.get("total_h",0):.1f}h</div></div>')
    return f'<div style="display:flex;flex-wrap:wrap;gap:10px">{cards}</div>'

def pending_est_table(estimates, title):
    if not estimates: return ''
    grand = sum(r['est_h'] for r in estimates)
    rows = ''.join(
        f'<tr><td><span style="color:{CNT_STYLE.get(r["cnt"],("#1e293b","#e2e8f0"))[1]};font-weight:700">{r["cnt"]}</span></td>'
        f'<td>Rua {r["rua"]}</td><td style="text-align:right;color:#f87171">{fmt_n(r["n_pend"])}</td>'
        f'<td style="text-align:right;color:#eab308">{fmt_m(r["med_min"])}</td>'
        f'<td style="text-align:right;font-weight:700;color:#60a5fa">{fmt_m(r["est_min"])}</td>'
        f'<td style="text-align:right;color:#94a3b8">{r["est_h"]:.1f}h</td></tr>'
        for r in sorted(estimates, key=lambda x:(x['cnt'],x['rua']))
    )
    return (f'<div class="card" style="border-left:4px solid #f97316">'
            f'<h3>⏳ {title} — Total estimado: <span style="color:#f97316">{grand:.1f}h</span></h3>'
            f'<p style="font-size:10px;color:#6b7280;margin-bottom:8px">Mediana por rua × posições pendentes distintas. Fallback para mediana geral se rua sem histórico.</p>'
            f'<div class="scroll"><table><thead><tr><th>Cnt</th><th>Rua</th><th>Pos.Pendentes</th><th>Mediana/Pos</th><th>Tempo Est.</th><th>Horas</th></tr></thead>'
            f'<tbody>{rows}</tbody></table></div></div>')

def pareto_section(D):
    if not D.get('pareto'): return ''
    rows = ''.join(
        f'<tr><td style="color:#60a5fa;font-size:10px">{r["cd_posicao"]}</td>'
        f'<td title="{r["ds_produto"]}" style="max-width:130px;overflow:hidden;text-overflow:ellipsis">{str(r["ds_produto"])[:22]}</td>'
        f'<td style="color:#f87171">{r["cd_lote"]}</td>'
        f'<td style="text-align:right">{float(r["saldo"]):,.0f}</td>'
        f'<td style="text-align:right">{float(r["contagem"]):,.0f}</td>'
        f'<td style="color:{"#f87171" if float(r["net_div"])<0 else "#fbbf24"};font-weight:700;text-align:right">{float(r["net_div"]):+,.0f}</td>'
        f'<td style="font-size:9px;color:#6b7280">{str(r["status"])[:14]}</td></tr>'
        for r in D['pareto']
    )
    return (f'<div class="card"><h3>📈 Top 20 Posições por Divergência</h3>'
            f'<div class="scroll"><table><thead><tr><th>Posição</th><th>Produto</th><th>Lote</th>'
            f'<th>Qt.Sist.(A)</th><th>Qt.Cont.(Últ.)</th><th>Div.Líq.</th><th>Status</th>'
            f'</tr></thead><tbody>{rows}</tbody></table></div></div>')

def ctd_section(D):
    blocks = ''
    for cnt in ['A','B','C']:
        recs = D.get(f'ctd_{cnt}',[])
        if not recs: continue
        _,fg = CNT_STYLE[cnt]
        rows = ''.join(
            f'<tr><td style="color:#e2e8f0;max-width:130px;overflow:hidden;text-overflow:ellipsis">{r["user"][:22]}</td>'
            f'<td style="text-align:right;color:#60a5fa">{fmt_n(r["n"])}</td>'
            f'<td style="text-align:right;color:#eab308">{r["dur"]:.0f}m</td>'
            f'<td style="text-align:right;color:#4ade80">{r["avg"]:.1f}m</td></tr>'
            for r in recs
        )
        blocks += (f'<div style="flex:1;min-width:200px">'
                   f'<div style="font-size:10px;color:{fg};font-weight:700;margin-bottom:5px">Contagem {cnt}</div>'
                   f'<div class="scroll"><table><thead><tr><th>Contador</th><th>Pos.</th><th>Total</th><th>Média</th></tr></thead>'
                   f'<tbody>{rows}</tbody></table></div></div>')
    if not blocks: return ''
    return f'<div class="card"><h3>👥 Ranking de Contadores</h3><div style="display:flex;flex-wrap:wrap;gap:14px">{blocks}</div></div>'

ADDR_BOX = '''<div class="addr-box">
<strong>Endereço:</strong>
<div class="addr-tag">Ex: <span>70.01.007.002.03</span></div>
<div class="addr-tag">Arm: <span>70</span></div>
<div class="addr-tag">Área: <span>01</span></div>
<div class="addr-tag">Rua: <span>007</span></div>
<div class="addr-tag">Prédio: <span>002</span></div>
<div class="addr-tag">Nível: <span>03</span></div>
</div>'''

LEGEND = '''<div class="legend-row">
<strong style="color:#e2e8f0">Tempo:</strong>
<span><span class="ldot" style="background:#22c55e"></span>&lt;1min</span>
<span><span class="ldot" style="background:#84cc16"></span>1–2m</span>
<span><span class="ldot" style="background:#eab308"></span>2–5m</span>
<span><span class="ldot" style="background:#f97316"></span>5–15m</span>
<span><span class="ldot" style="background:#ef4444"></span>&gt;15m</span>
</div>'''

def analysis_plan(plano, D, is_first, source):
    pid = make_pid(plano)
    display = 'block' if is_first else 'none'
    kpis = build_kpi_row(D, source)
    pend = (pend_blk_err(D['err_cnt_A'],D['err_cnt_B'],D['err_cnt_C'])
            if source=='err' else pend_blk_port(D))

    tipo_dist = {}
    for a in D['lot_alert']: t=a['tipo']; tipo_dist[t]=tipo_dist.get(t,0)+1
    tipo_html = ''.join(
        f'<span class="badge-tag {"bt-purp" if "Erro de UM" in t else TIPO_CLS.get(t,"bt-warn")}" style="margin:2px;font-size:9px">{t}: {n}</span>'
        for t,n in sorted(tipo_dist.items(), key=lambda x:-x[1])
    )

    alert_sec = ''
    if D['lot_alert']:
        alert_sec = (f'<div class="card" style="border-left:4px solid #ef4444">'
                     f'<h3>⚠️ FINALIZADO S/DIV com Divergência de Quantidade — {fmt_n(len(D["lot_alert"]))} lotes</h3>'
                     f'<div style="font-size:10px;color:#94a3b8;margin-bottom:7px">Qt.Sist.(A) = saldo original Contagem A · Qt.Cont.(Últ.) = última contagem por posição</div>'
                     f'<div style="margin-bottom:7px">{tipo_html}</div>'
                     f'{alert_table(D["lot_alert"])}</div>')

    valid_sec = (f'<div class="card" style="border-left:4px solid #22c55e">'
                 f'<h3>✅ Lotes Validados — {fmt_n(D["n_lot_ok"])} lotes · {fmt_n(D["n_sku_ok"])} SKUs</h3>'
                 f'<p style="font-size:10px;color:#6b7280;margin-bottom:7px">Todas as posições com FINALIZADO S/DIV e divergência líquida = 0 (saldo A = contagem última).</p>'
                 f'{sku_ok_table(D["sku_ok"])}</div>')

    inv_cod_s = ''
    if D['inv_cod']:
        inv_cod_s = (f'<div class="card"><h3>🔀 Inversões de Código — {fmt_n(D["n_inv_cod"])} posições distintas</h3>'
                     f'<p style="font-size:10px;color:#6b7280;margin-bottom:7px">Posições com produto encontrado ≠ produto esperado (última contagem). Qt.Sist.(A) = saldo original.</p>'
                     f'{inv_table(D["inv_cod"][:30],"cod")}'
                     f'{"<p style="+chr(39)+"font-size:10px;color:#6b7280;margin-top:5px"+chr(39)+">Mostrando 30 de "+str(D["n_inv_cod"])+"</p>" if D["n_inv_cod"]>30 else ""}'
                     f'</div>')

    inv_lot_s = ''
    if D['inv_lot']:
        inv_lot_s = (f'<div class="card"><h3>🏷️ Inversões de Lote — {fmt_n(D["n_inv_lot"])} posições</h3>'
                     f'{inv_table(D["inv_lot"][:50],"lot")}</div>')

    corr_note = '''<div class="corr-note">
✅ <strong style="color:#4ade80">Correção aplicada:</strong>
<strong style="color:#4ade80">Qt. Sist.(A)</strong> = <code>qt_saldoinicial</code> da Contagem A (saldo original do sistema) ·
<strong style="color:#60a5fa">Qt. Cont.(Últ.)</strong> = <code>qt_contagem</code> da última contagem por posição (C&gt;B&gt;A)
</div>'''

    return (f'<div id="{pid}" data-plan="plan" style="display:{display}">'
            f'<h2 style="font-size:17px;font-weight:700;color:#fff;margin-bottom:10px">📦 {plano}</h2>'
            f'{corr_note}{kpis}{pend}{alert_sec}{valid_sec}{inv_cod_s}{inv_lot_s}'
            f'{pareto_section(D)}{ctd_section(D)}</div>')

def time_plan(plano, D, is_first, source):
    pid = make_pid(plano)
    display = 'block' if is_first else 'none'
    overall = D['overall_err'] if source=='err' else D['overall_port']
    hists   = D['hists_err']  if source=='err' else D['hists_port']
    tdims   = D['time_err']   if source=='err' else D['time_port']
    pend_est = D['pend_est_err'] if source=='err' else D.get('pend_est_port',[])
    kpi_t = kpi_cnt_block(overall) if overall else '<p style="color:#555">Sem dados de tempo.</p>'
    h_block = hist_bars(hists)
    rua_recs = [r for r in tdims.get('rua',[]) if r.get('dim_val')]
    heat = heatmap_rua(rua_recs)
    dim_tables = ''
    for dim,lbl in [('area','Área'),('predio','Prédio'),('nivel','Nível')]:
        recs = tdims.get(dim,[])
        if recs: dim_tables += f'<div class="card"><h3>⏱ Tempo por {lbl}</h3>{dim_time_table(recs,lbl)}</div>'
    est_block = pending_est_table(pend_est,'Estimativa de Tempo para Posições Pendentes')
    return (f'<div id="{pid}" data-plan="plan" style="display:{display}">'
            f'<h2 style="font-size:17px;font-weight:700;color:#fff;margin-bottom:10px">📦 {plano}</h2>'
            f'<div class="card"><h3>🎯 Tempo de Contagem por Tipo</h3>'
            f'<p style="font-size:10px;color:#6b7280;margin-bottom:9px">Mediana = valor central (robusto a outliers). Fonte: {"dt_inicio/dt_conclusao (Erros)" if source=="err" else "dt_contagem*_ini/fim (Portal)"}</p>'
            f'{kpi_t}</div>'
            f'<div class="card"><h3>📊 Distribuição de Tempo por Faixa</h3>{h_block}</div>'
            f'{est_block}'
            f'<div class="card"><h3>🗺 Mapa de Tempo por Rua (mediana)</h3>{heat}'
            f'<details><summary>Ver tabela completa por Rua</summary><div style="margin-top:8px">{dim_time_table(rua_recs,"Rua")}</div></details></div>'
            f'{dim_tables}</div>')

def update_js_analysis(source):
    return f'''
window.dashboardUpdate = function(DATA, ts) {{
  const PLAN_ORDER = ['Climatizado','Odonto','Armazém 72','Climatizado RUA7/Nív7'];
  const PLAN_IDS   = {{'Climatizado':'Climatizado','Odonto':'Odonto','Armazém 72':'Armazem_72','Climatizado RUA7/Nív7':'Climatizado_RUA7_Niv7'}};
  PLAN_ORDER.forEach(function(plano) {{
    if (!DATA[plano]) return;
    const D = DATA[plano];
    const pid = PLAN_IDS[plano] || plano.replace(/ /g,'_').replace(/\//g,'_');
    const el = document.getElementById(pid);
    if (!el) return;
    const upd = (id, val) => {{ const e=el.querySelector('#kpi_'+id+' .val'); if(e) e.textContent=val.toLocaleString('pt-BR'); }};
    upd('fin_s', D.n_fin_s||0); upd('fin_c', D.n_fin_c||0);
    upd('pend',  D.n_pend||0);  upd('em_c',  D.n_em_c||0);
    upd('lot_ok',D.n_lot_ok||0); upd('sku_ok',D.n_sku_ok||0);
    upd('lot_alert',D.n_lot_alert||0); upd('inv_cod',D.n_inv_cod||0);
    upd('pend_a',D.port_pend_a||0); upd('pend_b',D.port_pend_b||0); upd('pend_c',D.port_pend_c||0);
    showToast('✅ ' + plano + ' atualizado!');
  }});
  const tsEl=document.getElementById('last-updated');
  if(tsEl) tsEl.textContent='Atualizado: '+ts;
}};
function showToast(msg){{
  const t=document.createElement('div');
  t.style.cssText='position:fixed;top:20px;right:20px;background:#14532d;color:#4ade80;padding:10px 18px;border-radius:8px;font-size:13px;font-weight:700;z-index:99999;border:1px solid #22c55e';
  t.textContent=msg; document.body.appendChild(t); setTimeout(()=>t.remove(),3000);
}}
'''

def write_dashboard(title, hdr_grad, source, plan_fn, filename, ALL, ENGINE_JS, MODAL_JS):
    tabs = ''; blocks = ''; first = True
    for plano in PLAN_ORDER:
        if plano not in ALL: continue
        pid = make_pid(plano)
        tabs += f'<button class="tab-btn" data-tab="plan" id="tab-plan-{pid}" onclick="showPlan(\'{pid}\')">{plano}</button>\n'
        blocks += plan_fn(plano, ALL[plano], first, source)
        first = False

    ts = datetime.now().strftime('%d/%m/%Y %H:%M')
    html = f'''<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
<style>{BASE_CSS}.hdr{{background:linear-gradient(135deg,{hdr_grad})}}</style>
</head><body>
<div class="hdr">
  <div class="hdr-left">
    <h1>{title}</h1>
    <p id="last-updated">Atualizado: {ts} · OneDrive: {ONEDRIVE}</p>
  </div>
  <div class="hdr-badge">📂 OneDrive Conectado</div>
</div>
<div class="tabs">{tabs}</div>
<div class="content">
{ADDR_BOX}
{""+LEGEND if "Tempo" in title else ""}
{blocks}
</div>
<script>
{JS_PLAN}
{ENGINE_JS}
{MODAL_JS}
{update_js_analysis(source)}
</script>
</body></html>'''

    path = OUT_DIR / filename
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  ✅ {filename} ({len(html):,} chars)")

# ── ANÁLISE SKU-LOTE COM BASE DE ESTOQUE ────────────────────────────────────
def _build_stock_analysis(err_df, stock_path, out_dir):
    """Gera dashboard de cruzamento SKU-Lote usando a base de estoque PMC.
    Versão aprimorada com: validade, lotes vencidos, fechamento SKU-Lote, filtro data-tipo.
    """
    import html as htmlmod
    stock = pd.read_excel(stock_path, sheet_name='Plan1')
    stock = stock[stock['Cod.'].notna() & stock['Lote'].notna()].copy()
    stock['sku'] = stock['Cod.'].astype(str).str.replace('.0','',regex=False)
    stock['lote'] = stock['Lote'].astype(str).str.strip()
    stock['desc'] = stock['Descrição'].astype(str)

    # ── Validade map ──
    stock['validade_dt'] = pd.to_datetime(stock.get('Validade'), errors='coerce')
    val_map = {}  # (sku, lote) → datetime
    lote_val_map = {}  # lote → datetime (fallback)
    for _, sr in stock[stock['validade_dt'].notna()].iterrows():
        key = (sr['sku'], sr['lote'])
        val_map[key] = sr['validade_dt']
        lote_val_map[sr['lote']] = sr['validade_dt']
    today = pd.Timestamp.now().normalize()

    sku_lotes_map = stock.groupby('sku')['lote'].apply(set).to_dict()
    lote_skus_map = stock.groupby('lote')['sku'].apply(set).to_dict()
    sku_desc_map  = stock.drop_duplicates('sku').set_index('sku')['desc'].to_dict()

    # ── Saldo por (sku, lote) ──
    saldo_map = stock.groupby(['sku','lote'])['Qtd'].sum().to_dict()

    div = err_df[(err_df['ds_statuscontagem']=='FINALIZADO C/ DIVERGÊNCIA') & (err_df['cd_contagem']=='A')].copy()
    div['sku_sis']  = div['cd_produtocliente'].astype(str).str.replace('.0','',regex=False)
    div['lote_sis'] = div['cd_lote'].astype(str).str.strip()
    div['sku_cnt']  = div['cd_produtocontagemcliente'].astype(str).str.replace('.0','',regex=False)
    div['lote_cnt'] = div['cd_lotecontagem'].astype(str).str.strip()

    TIPO_LABELS = {
        'sku_confirmado':           '🔄 SKU Contado Confirmado no Estoque',
        'lote_sku_errado':          '🚨 Lote Pertence a Outro SKU',
        'prod_div_mesmo':           '⚠️ Produto Divergente (Mesmo SKU)',
        'prod_div_nao_encontrado':  '❓ Produto Divergente (Lote Não Encontrado)',
        'lote_div_ok':              '✅ Lote Divergente (Correto no Estoque)',
        'lote_outro_sku':           '🔀 Lote Divergente → Outro SKU',
        'lote_nao_encontrado':      '🏷️ Lote Não Encontrado no Estoque',
        'lote_div_nid':             '🏷️ Lote Divergente (Não Identificado)',
        'qtd_div':                  '📊 Quantidade Divergente',
        'vazia':                    '📦 Posição Vazia',
        'outro':                    '❔ Outro',
    }
    TIPO_COLORS = {
        'sku_confirmado':'#1565c0','lote_sku_errado':'#c62828','prod_div_mesmo':'#e65100',
        'prod_div_nao_encontrado':'#6a1b9a','lote_div_ok':'#00695c','lote_outro_sku':'#ad1457',
        'lote_nao_encontrado':'#37474f','lote_div_nid':'#546e7a','qtd_div':'#e65100',
        'vazia':'#4527a0','outro':'#546e7a',
    }

    def classify(r):
        ss = r['sku_sis']; ls = r['lote_sis']
        sc = r['sku_cnt'] if r['sku_cnt'] not in ['nan','None'] else ''
        lc = r['lote_cnt'] if r['lote_cnt'] not in ['nan','None'] else ''
        erro = str(r.get('ds_erro',''))
        lr = lc if lc else ls
        if ss in ['nan','None'] and sc in ['nan','None','']:
            return 'vazia','','','Verificar fisicamente se posição está vazia'
        if 'Produto divergente' in erro:
            if lr in lote_skus_map:
                poss = lote_skus_map[lr]
                if sc and sc in poss:
                    return 'sku_confirmado',sc,sku_desc_map.get(sc,'?'),f'Lote {lr} → SKU {sc} confirmado no estoque.'
                other = sorted(poss - {ss})
                if other:
                    return 'lote_sku_errado',', '.join(other),sku_desc_map.get(other[0],'?'),f'Lote {lr} pertence ao(s) SKU(s) {", ".join(other)}. Reabrir no SKU correto.'
                return 'prod_div_mesmo','','','Produto divergente, lote do mesmo SKU.'
            return 'prod_div_nao_encontrado','','',f'Lote {lr} não encontrado no estoque.'
        if 'Lote do produto divergente' in erro:
            if lc and lc in lote_skus_map:
                if ss in lote_skus_map[lc]:
                    return 'lote_div_ok','','',f'Lote {lc} existe para este SKU. Ajustar lote.'
                other = sorted(lote_skus_map[lc] - {ss})
                if other:
                    return 'lote_outro_sku',', '.join(other),sku_desc_map.get(other[0],'?'),f'Lote {lc} pertence ao SKU {", ".join(other)}.'
                return 'lote_div_nid','','','Lote divergente não identificado.'
            return 'lote_nao_encontrado','','',f'Lote {lc} não encontrado. Verificar digitação.'
        if 'Quantidade divergente' in erro:
            return 'qtd_div','','','Verificar movimentações e recontar.'
        if 'Posição Vazia' in erro:
            return 'vazia','','','Verificar fisicamente. Se vazia, registrar 0.'
        return 'outro','','','Verificar manualmente.'

    div[['tipo','sku_correto','desc_correto','acao']] = div.apply(lambda r: pd.Series(classify(r)), axis=1)

    # ── Validade para cada linha ──
    def get_validade(r):
        sku = r['sku_sis']; lote = r['lote_cnt'] if r['lote_cnt'] not in ['nan','None'] else r['lote_sis']
        v = val_map.get((sku, lote)) or lote_val_map.get(lote)
        return v
    div['validade_dt'] = div.apply(get_validade, axis=1)
    div['vencido'] = div['validade_dt'].apply(lambda v: v is not None and pd.notna(v) and v < today)

    esc = lambda s: htmlmod.escape(str(s)) if pd.notna(s) else ''
    summary = div['tipo'].value_counts()
    inversions = div[div['tipo'].isin(['sku_confirmado','lote_sku_errado','lote_outro_sku'])].sort_values('cd_posicao')
    by_plan = div.groupby(['plano','tipo']).size().unstack(fill_value=0)
    n_vencidos_div = int(div['vencido'].sum())
    n_vencidos_inv = int(inversions['vencido'].sum())
    tipo_order = ['sku_confirmado','lote_sku_errado','lote_outro_sku','prod_div_nao_encontrado','prod_div_mesmo','lote_div_ok','lote_nao_encontrado','qtd_div','vazia','outro']

    def fmt_val(v):
        if v is None or pd.isna(v):
            return ''
        return v.strftime('%d/%m/%Y')

    def val_cell(r):
        v = r['validade_dt']
        if v is None or pd.isna(v):
            return '<td>—</td>'
        vs = fmt_val(v)
        if r['vencido']:
            return f'<td class="vencido">{esc(vs)}</td>'
        return f'<td>{esc(vs)}</td>'

    # ── Summary cards ──
    cards_html = ''
    for tipo, count in summary.items():
        color = TIPO_COLORS.get(tipo,'#546e7a'); label = TIPO_LABELS.get(tipo,tipo); pct = count/len(div)*100
        cards_html += f'<div class="summary-card" style="border-left:4px solid {color}"><div class="sc-label">{label}</div><div class="sc-row"><span class="sc-count" style="color:{color}">{count}</span><span class="sc-pct">{pct:.1f}%</span><div class="sc-bar-bg"><div class="sc-bar" style="width:{pct}%;background:{color}"></div></div></div></div>'
    # Vencidos card
    if n_vencidos_div > 0:
        pct_v = n_vencidos_div / len(div) * 100
        cards_html += f'<div class="summary-card" style="border-left:4px solid #b71c1c"><div class="sc-label">⏰ Lotes Vencidos</div><div class="sc-row"><span class="sc-count" style="color:#b71c1c">{n_vencidos_div}</span><span class="sc-pct">{pct_v:.1f}%</span><div class="sc-bar-bg"><div class="sc-bar" style="width:{pct_v}%;background:#b71c1c"></div></div></div></div>'

    # ── Inversions rows (with validade) ──
    inv_rows = ''
    for _, r in inversions.iterrows():
        color = TIPO_COLORS.get(r['tipo'],'#546e7a'); label = TIPO_LABELS.get(r['tipo'],'')
        lr = r['lote_cnt'] if r['lote_cnt'] not in ['nan','None'] else r['lote_sis']
        vc = val_cell(r)
        inv_rows += f'<tr><td class="td-pos">{esc(r["cd_posicao"])}</td><td><span class="badge" style="background:{color}22;color:{color};border:1px solid {color}44">{esc(label)}</span></td><td class="td-sku">{esc(r["sku_sis"])}</td><td class="td-desc">{esc(str(r.get("ds_produto",""))[:50])}</td><td class="td-sku" style="color:#c62828">{esc(r["sku_cnt"])}</td><td>{esc(lr)}</td>{vc}<td class="td-sku" style="color:#2e7d32">{esc(r["sku_correto"])}</td><td class="td-desc">{esc(str(r["desc_correto"])[:50])}</td><td class="td-obs">{esc(r["acao"])}</td></tr>'

    # ── All items rows (with data-tipo attribute and validade) ──
    all_rows = ''
    for _, r in div.sort_values('cd_posicao').iterrows():
        color = TIPO_COLORS.get(r['tipo'],'#546e7a'); label_short = TIPO_LABELS.get(r['tipo'],'?')[:30]
        lr = r['lote_cnt'] if r['lote_cnt'] not in ['nan','None'] else r['lote_sis']
        vc = val_cell(r)
        all_rows += f'<tr data-tipo="{esc(r["tipo"])}"><td class="td-pos">{esc(r["cd_posicao"])}</td><td><span class="tipo-badge" style="background:{color}">{esc(label_short)}</span></td><td class="td-sku">{esc(r["sku_sis"])}</td><td class="td-desc">{esc(str(r.get("ds_produto",""))[:40])}</td><td>{esc(lr)}</td>{vc}<td class="td-sku" style="color:#2e7d32">{esc(r["sku_correto"])}</td><td class="td-obs">{esc(r["acao"][:80])}</td></tr>'

    # ── Plan breakdown ──
    plan_table = '<table class="plan-table"><thead><tr><th>Plano</th>'
    for t in tipo_order:
        if t in summary.index:
            icon = TIPO_LABELS.get(t,'').split(' ')[0]
            plan_table += f'<th title="{esc(TIPO_LABELS.get(t,""))}">{icon}</th>'
    plan_table += '<th>Total</th></tr></thead><tbody>'
    for plano in ['Climatizado','Odonto','Armazém 72','Clim. RUA7/Nív7']:
        if plano in by_plan.index:
            row = by_plan.loc[plano]; rt = 0
            plan_table += f'<tr><td><strong>{esc(plano)}</strong></td>'
            for t in tipo_order:
                if t in summary.index:
                    v = row.get(t,0); rt += v; c = TIPO_COLORS.get(t,'#546e7a')
                    plan_table += f'<td style="color:{c};font-weight:700">{v if v else "-"}</td>'
            plan_table += f'<td style="font-weight:800">{rt}</td></tr>'
    plan_table += '</tbody></table>'

    # ── Type filter options ──
    opts_html = '<option value="">Todos os tipos</option>'
    for t in tipo_order:
        if t in summary.index:
            opts_html += f'<option value="{esc(t)}">{esc(TIPO_LABELS.get(t,""))} ({summary[t]})</option>'

    # ══════════════════════════════════════════════════════════════════════════
    # ── ABA 4: FECHAMENTO SKU-LOTE ──
    # ══════════════════════════════════════════════════════════════════════════
    # Uses ALL inventory data (not just divergent), taking last count per position
    rank_dict = {'A':1,'B':2,'C':3}
    all_inv = err_df.copy()
    all_inv['sku']  = all_inv['cd_produtocliente'].astype(str).str.replace('.0','',regex=False)
    all_inv['lote'] = all_inv['cd_lote'].astype(str).str.strip()
    all_inv['plano_f'] = all_inv['ds_descricao'].map(PLAN_MAP)
    all_inv['rank_n'] = all_inv['cd_contagem'].map(rank_dict)
    all_inv['qt_contagem_f'] = pd.to_numeric(all_inv.get('qt_contagem'), errors='coerce').fillna(0)
    all_inv['status'] = all_inv['ds_statuscontagem'].astype(str)

    # Last count per position (highest rank = C>B>A → max rank_n)
    idx_last = all_inv.groupby('cd_posicao')['rank_n'].idxmax()
    last = all_inv.loc[idx_last].copy()

    # Build closing aggregation
    closing_rows_data = []
    grp = last.groupby(['sku','lote','plano_f'])
    for (sku, lote, plano), g in grp:
        if sku in ['nan','None'] or lote in ['nan','None']:
            continue
        n_ok = int((g['status']=='FINALIZADO S/ DIVERGÊNCIA').sum())
        n_div = int((g['status']=='FINALIZADO C/ DIVERGÊNCIA').sum())
        n_pend = int((~g['status'].isin(['FINALIZADO S/ DIVERGÊNCIA','FINALIZADO C/ DIVERGÊNCIA'])).sum())
        qt_cnt = float(g['qt_contagem_f'].sum())
        qt_saldo = float(saldo_map.get((sku, lote), 0))
        diff = qt_cnt - qt_saldo
        # Validade
        v = val_map.get((sku, lote)) or lote_val_map.get(lote)
        venc = v is not None and pd.notna(v) and v < today
        desc = sku_desc_map.get(sku, str(g.iloc[0].get('ds_produto','')))[:50]
        closing_rows_data.append({
            'sku':sku,'lote':lote,'plano':plano or '','desc':desc,
            'n_ok':n_ok,'n_div':n_div,'n_pend':n_pend,
            'qt_saldo':qt_saldo,'qt_cnt':qt_cnt,'diff':diff,
            'validade':v,'vencido':venc
        })

    cdf = pd.DataFrame(closing_rows_data) if closing_rows_data else pd.DataFrame()
    n_closing = len(cdf)
    n_closing_venc = int(cdf['vencido'].sum()) if len(cdf) > 0 else 0
    total_saldo = float(cdf['qt_saldo'].sum()) if len(cdf) > 0 else 0
    total_cnt = float(cdf['qt_cnt'].sum()) if len(cdf) > 0 else 0
    total_pend = int(cdf['n_pend'].sum()) if len(cdf) > 0 else 0
    total_ok = int(cdf['n_ok'].sum()) if len(cdf) > 0 else 0
    total_div = int(cdf['n_div'].sum()) if len(cdf) > 0 else 0

    # Plano filter options for closing tab
    planos_closing = sorted(cdf['plano'].unique()) if len(cdf) > 0 else []
    opts_plano = '<option value="">Todos os planos</option>'
    for p in planos_closing:
        if p:
            opts_plano += f'<option value="{esc(p)}">{esc(p)}</option>'

    # Build closing table rows
    closing_rows = ''
    if len(cdf) > 0:
        for _, cr in cdf.sort_values(['sku','lote']).iterrows():
            vs = fmt_val(cr['validade']) if cr['validade'] is not None and pd.notna(cr['validade']) else '—'
            vcls = ' class="vencido"' if cr['vencido'] else ''
            diff_cls = ''
            if cr['diff'] > 0:
                diff_cls = ' style="color:#c62828;font-weight:700"'
            elif cr['diff'] < 0:
                diff_cls = ' style="color:#e65100;font-weight:700"'
            pend_cls = ' style="color:#b71c1c;font-weight:700"' if cr['n_pend'] > 0 else ''
            closing_rows += f'<tr data-plano="{esc(cr["plano"])}" data-pend="{cr["n_pend"]}"><td class="td-sku">{esc(cr["sku"])}</td><td class="td-desc">{esc(cr["desc"])}</td><td>{esc(cr["lote"])}</td><td{vcls}>{esc(vs)}</td><td>{esc(cr["plano"])}</td><td>{cr["n_ok"]}</td><td style="color:#c62828">{cr["n_div"]}</td><td{pend_cls}>{cr["n_pend"]}</td><td style="text-align:right">{cr["qt_saldo"]:,.0f}</td><td style="text-align:right">{cr["qt_cnt"]:,.0f}</td><td{diff_cls} style="text-align:right">{cr["diff"]:+,.0f}</td></tr>'

    ts_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    html_out = f'''<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Análise SKU-Lote — Cruzamento com Estoque PMC</title>
<style>
:root{{--bg:#f4f6f8;--card:#fff;--border:#e0e4e8;--text:#1a2332;--muted:#5a6a7e}}
*{{box-sizing:border-box;margin:0;padding:0}}body{{font-family:'Segoe UI',sans-serif;background:var(--bg);color:var(--text)}}
.page-header{{background:linear-gradient(135deg,#0d2137 0%,#1e3a5c 50%,#0d3320 100%);color:white;padding:28px 36px;display:flex;align-items:center;gap:18px}}
.header-icon{{font-size:44px}}.header-title h1{{font-size:22px;font-weight:700}}.header-title p{{font-size:12px;opacity:.7;margin-top:3px}}
.header-badges{{margin-left:auto;display:flex;gap:10px;flex-wrap:wrap}}.hbadge{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);border-radius:16px;padding:5px 12px;font-size:11px;font-weight:600}}
.nav{{background:#1a3a5c;display:flex;gap:0;border-bottom:2px solid #0d2137;flex-wrap:wrap}}.nav-btn{{background:none;border:none;color:rgba(255,255,255,.6);padding:11px 20px;font-size:12px;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s}}.nav-btn:hover,.nav-btn.active{{color:white;border-bottom-color:#64b5f6;background:rgba(255,255,255,.06)}}
.container{{max-width:1500px;margin:0 auto;padding:24px 20px}}.section{{display:none}}.section.active{{display:block}}
.summary-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px;margin-bottom:24px}}
.summary-card{{background:var(--card);border-radius:8px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}.sc-label{{font-size:13px;font-weight:600;color:var(--text);margin-bottom:6px}}.sc-row{{display:flex;align-items:center;gap:10px}}.sc-count{{font-size:22px;font-weight:800;min-width:42px}}.sc-pct{{font-size:11px;color:var(--muted);min-width:38px}}.sc-bar-bg{{flex:1;height:6px;background:#eef0f3;border-radius:3px;overflow:hidden}}.sc-bar{{height:100%;border-radius:3px;min-width:2px}}
.plan-table{{width:100%;border-collapse:collapse;background:var(--card);border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:24px}}.plan-table th{{background:#1a3a5c;color:white;padding:10px;font-size:11px;text-align:center}}.plan-table td{{padding:9px 10px;border-bottom:1px solid var(--border);text-align:center;font-size:13px}}.plan-table tr:hover td{{background:#f8f9fb}}
.badge{{display:inline-block;padding:3px 8px;border-radius:4px;font-size:10px;font-weight:600;white-space:nowrap}}.tipo-badge{{display:inline-block;padding:3px 7px;border-radius:4px;font-size:10px;font-weight:600;color:white;white-space:nowrap}}
.table-wrap{{overflow-x:auto;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.items-table{{width:100%;border-collapse:collapse;background:var(--card);font-size:12px}}.items-table th{{background:#1a3a5c;color:white;padding:9px 10px;font-size:10px;text-align:left;white-space:nowrap;position:sticky;top:0;z-index:10}}.items-table td{{padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:top}}.items-table tr:hover td{{background:#f5f7fa}}
.td-pos{{font-family:monospace;font-size:11px;color:var(--muted);white-space:nowrap}}.td-sku{{font-family:monospace;font-size:11px;font-weight:700;color:#1565c0}}.td-desc{{max-width:200px;color:#2d3748;font-size:11px}}.td-obs{{max-width:250px;color:#374151;font-style:italic;font-size:11px}}
.filter-bar{{display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap}}.filter-input{{padding:7px 12px;border:1px solid var(--border);border-radius:5px;font-size:12px;flex:1;min-width:180px;outline:none}}.filter-input:focus{{border-color:#1565c0;box-shadow:0 0 0 2px rgba(21,101,192,.15)}}.filter-select{{padding:7px 10px;border:1px solid var(--border);border-radius:5px;font-size:12px;outline:none}}.result-count{{font-size:11px;color:var(--muted);white-space:nowrap}}
.insight-box{{background:#fffbeb;border:1px solid #f59e0b44;border-radius:8px;padding:16px 18px;margin-bottom:20px}}.insight-box h3{{font-size:14px;color:#92400e;margin-bottom:6px}}.insight-box p{{font-size:13px;color:#78350f;line-height:1.6}}
.vencido{{color:#b71c1c!important;font-weight:700;background:#fef2f2}}
.chk-label{{font-size:12px;color:var(--muted);cursor:pointer;display:flex;align-items:center;gap:4px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px}}.kpi-card{{background:var(--card);border-radius:8px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.06);text-align:center}}.kpi-val{{font-size:24px;font-weight:800}}.kpi-lbl{{font-size:11px;color:var(--muted);margin-top:4px}}
@media print{{.nav{{display:none}}.section{{display:block!important}}}}
</style></head><body>
<div class="page-header"><div class="header-icon">🔬</div><div class="header-title"><h1>Análise SKU-Lote — Cruzamento com Base de Estoque</h1><p>Fonte: Posição de Estoque PMC + Inventário Análise de Erros · Atualizado: {ts_str} · {len(div)} posições divergentes analisadas</p></div><div class="header-badges"><span class="hbadge">🔄 {len(inversions)} Inversões</span><span class="hbadge">📊 {len(div)} Divergências</span><span class="hbadge">⏰ {n_vencidos_div} Vencidos</span><span class="hbadge">📦 {len(stock)} Linhas Estoque</span></div></div>
<nav class="nav">
<button class="nav-btn active" onclick="showSec('resumo')">📊 Resumo</button>
<button class="nav-btn" onclick="showSec('inversoes')">🔄 Inversões ({len(inversions)})</button>
<button class="nav-btn" onclick="showSec('todos')">📋 Todas Divergências ({len(div)})</button>
<button class="nav-btn" onclick="showSec('fechamento')">📦 Fechamento SKU-Lote ({n_closing})</button>
</nav>

<!-- TAB 1: RESUMO -->
<div class="section active" id="sec-resumo"><div class="container">
<div class="insight-box"><h3>🔍 Resultado Principal</h3><p>Foram identificadas <strong>{len(inversions)} inversões de SKU-Lote</strong> onde o lote contado pertence a um SKU diferente. Dessas, <strong>{summary.get('sku_confirmado',0)}</strong> tiveram o SKU contado confirmado na base de estoque, e <strong>{summary.get('lote_sku_errado',0)}</strong> tinham o lote alocado em outro SKU. Foram encontrados <strong>{n_vencidos_div} lotes vencidos</strong> entre as divergências.</p></div>
<div class="summary-grid">{cards_html}</div>
<h3 style="font-size:14px;margin-bottom:12px;color:var(--muted)">Distribuição por Plano</h3>{plan_table}
</div></div>

<!-- TAB 2: INVERSOES -->
<div class="section" id="sec-inversoes"><div class="container">
<div class="filter-bar">
<span style="font-size:12px;color:var(--muted)">🔍</span>
<input class="filter-input" id="flt-inv" type="text" placeholder="Posição, SKU, lote..." oninput="filterInv()">
<label class="chk-label"><input type="checkbox" id="chk-venc-inv" onchange="filterInv()"> Só vencidos</label>
<span class="result-count" id="cnt-inv">{len(inversions)} itens</span>
</div>
<div class="table-wrap"><table class="items-table"><thead><tr><th>Posição</th><th>Tipo</th><th>SKU Sistema</th><th>Descrição</th><th>SKU Contado</th><th>Lote</th><th>Validade</th><th>SKU Correto</th><th>Desc. Estoque</th><th>Ação</th></tr></thead><tbody id="tb-inv">{inv_rows}</tbody></table></div>
</div></div>

<!-- TAB 3: TODAS DIVERGENCIAS -->
<div class="section" id="sec-todos"><div class="container">
<div class="filter-bar">
<span style="font-size:12px;color:var(--muted)">🔍</span>
<input class="filter-input" id="flt-all" type="text" placeholder="Posição, SKU, lote, tipo..." oninput="filterAll()">
<select class="filter-select" id="sel-tipo" onchange="filterAll()">{opts_html}</select>
<label class="chk-label"><input type="checkbox" id="chk-venc-all" onchange="filterAll()"> Só vencidos</label>
<span class="result-count" id="cnt-all">{len(div)} itens</span>
</div>
<div class="table-wrap"><table class="items-table"><thead><tr><th>Posição</th><th>Classificação</th><th>SKU Sistema</th><th>Descrição</th><th>Lote</th><th>Validade</th><th>SKU Correto</th><th>Ação</th></tr></thead><tbody id="tb-all">{all_rows}</tbody></table></div>
</div></div>

<!-- TAB 4: FECHAMENTO -->
<div class="section" id="sec-fechamento"><div class="container">
<div class="kpi-grid">
<div class="kpi-card"><div class="kpi-val" style="color:#1565c0">{n_closing:,}</div><div class="kpi-lbl">SKU-Lote-Plano</div></div>
<div class="kpi-card"><div class="kpi-val" style="color:#2e7d32">{total_ok:,}</div><div class="kpi-lbl">Posições OK</div></div>
<div class="kpi-card"><div class="kpi-val" style="color:#c62828">{total_div:,}</div><div class="kpi-lbl">Posições Divergentes</div></div>
<div class="kpi-card"><div class="kpi-val" style="color:#e65100">{total_pend:,}</div><div class="kpi-lbl">Posições Pendentes</div></div>
<div class="kpi-card"><div class="kpi-val">{total_saldo:,.0f}</div><div class="kpi-lbl">Qt Saldo Total</div></div>
<div class="kpi-card"><div class="kpi-val">{total_cnt:,.0f}</div><div class="kpi-lbl">Qt Contagem Total</div></div>
</div>
<div class="filter-bar">
<span style="font-size:12px;color:var(--muted)">🔍</span>
<input class="filter-input" id="flt-fech" type="text" placeholder="SKU, lote, descrição..." oninput="filterFech()">
<select class="filter-select" id="sel-plano-fech" onchange="filterFech()">{opts_plano}</select>
<label class="chk-label"><input type="checkbox" id="chk-pend-fech" onchange="filterFech()"> Só pendentes</label>
<label class="chk-label"><input type="checkbox" id="chk-venc-fech" onchange="filterFech()"> Só vencidos</label>
<span class="result-count" id="cnt-fech">{n_closing} itens</span>
</div>
<div class="table-wrap"><table class="items-table"><thead><tr><th>SKU</th><th>Descrição</th><th>Lote</th><th>Validade</th><th>Plano</th><th>OK</th><th>Diverg.</th><th>Pend.</th><th style="text-align:right">Qt Saldo</th><th style="text-align:right">Qt Contagem</th><th style="text-align:right">Diferença</th></tr></thead><tbody id="tb-fech">{closing_rows}</tbody></table></div>
</div></div>

<script>
function showSec(id){{document.querySelectorAll('.section').forEach(function(s){{s.classList.remove('active')}});document.querySelectorAll('.nav-btn').forEach(function(b){{b.classList.remove('active')}});document.getElementById('sec-'+id).classList.add('active');event.currentTarget.classList.add('active')}}

function filterInv(){{
  var q=document.getElementById('flt-inv').value.toLowerCase();
  var onlyVenc=document.getElementById('chk-venc-inv').checked;
  var rows=document.querySelectorAll('#tb-inv tr');
  var vis=0;
  rows.forEach(function(r){{
    var txt=r.textContent.toLowerCase();
    var show=(!q||txt.includes(q));
    if(onlyVenc&&show){{show=r.querySelector('.vencido')!==null}}
    r.style.display=show?'':'none';
    if(show)vis++;
  }});
  document.getElementById('cnt-inv').textContent=vis+' itens';
}}

function filterAll(){{
  var q=document.getElementById('flt-all').value.toLowerCase();
  var tipo=document.getElementById('sel-tipo').value;
  var onlyVenc=document.getElementById('chk-venc-all').checked;
  var rows=document.querySelectorAll('#tb-all tr');
  var vis=0;
  rows.forEach(function(r){{
    var txt=r.textContent.toLowerCase();
    var show=(!q||txt.includes(q));
    if(tipo&&show){{show=(r.getAttribute('data-tipo')===tipo)}}
    if(onlyVenc&&show){{show=r.querySelector('.vencido')!==null}}
    r.style.display=show?'':'none';
    if(show)vis++;
  }});
  document.getElementById('cnt-all').textContent=vis+' itens';
}}

function filterFech(){{
  var q=document.getElementById('flt-fech').value.toLowerCase();
  var plano=document.getElementById('sel-plano-fech').value;
  var onlyPend=document.getElementById('chk-pend-fech').checked;
  var onlyVenc=document.getElementById('chk-venc-fech').checked;
  var rows=document.querySelectorAll('#tb-fech tr');
  var vis=0;
  rows.forEach(function(r){{
    var txt=r.textContent.toLowerCase();
    var show=(!q||txt.includes(q));
    if(plano&&show){{show=(r.getAttribute('data-plano')===plano)}}
    if(onlyPend&&show){{show=parseInt(r.getAttribute('data-pend')||'0')>0}}
    if(onlyVenc&&show){{show=r.querySelector('.vencido')!==null}}
    r.style.display=show?'':'none';
    if(show)vis++;
  }});
  document.getElementById('cnt-fech').textContent=vis+' itens';
}}
</script></body></html>'''

    path = out_dir / 'Análise SKU-Lote - Cruzamento Estoque.html'
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html_out)
    print(f"  📄 Análise SKU-Lote ({len(html_out):,} chars) — 4 abas: Resumo, Inversões, Divergências, Fechamento")


# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  ATUALIZAR DASHBOARDS DE INVENTÁRIO")
    print("=" * 60)
    print()

    # Carregar JS (engine + modal) da mesma pasta do script
    engine_path = SCRIPT_DIR / 'analysis_engine.js'
    modal_path  = SCRIPT_DIR / 'update_modal.js'

    if not engine_path.exists():
        print(f"ERRO: Arquivo não encontrado: {engine_path}")
        print("Certifique-se de que 'analysis_engine.js' está na mesma pasta que este script.")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    if not modal_path.exists():
        print(f"ERRO: Arquivo não encontrado: {modal_path}")
        print("Certifique-se de que 'update_modal.js' está na mesma pasta que este script.")
        input("\nPressione Enter para sair...")
        sys.exit(1)

    ENGINE_JS = engine_path.read_text(encoding='utf-8')
    MODAL_JS  = modal_path.read_text(encoding='utf-8')

    # Localizar arquivos Excel
    print(f"🔍 Buscando arquivos Excel em: {ONEDRIVE}")
    err_path   = find_file("Inventário Análise de Erros")
    port_path  = find_file("Inventário Analítico Portal")
    stock_path = find_file("POSIÇÃO DE ESTOQUE TODOS OS GRUPOS PMC")

    if not err_path:
        print(f"\nERRO: Arquivo 'Inventário Análise de Erros*.xlsx' não encontrado em {ONEDRIVE}")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    if not port_path:
        print(f"\nERRO: Arquivo 'Inventário Analítico Portal*.xlsx' não encontrado em {ONEDRIVE}")
        input("\nPressione Enter para sair...")
        sys.exit(1)

    print(f"  📄 Erros:  {Path(err_path).name}")
    print(f"  📄 Portal: {Path(port_path).name}")
    if stock_path:
        print(f"  📄 Estoque: {Path(stock_path).name}")
    else:
        print(f"  ⚠️  Estoque PMC não encontrado — análise SKU-Lote será ignorada")
    print()

    # Carregar Excel
    print("📊 Carregando planilhas...")
    err  = pd.read_excel(err_path,  sheet_name='ANALISE INVENTARIO')
    port = pd.read_excel(port_path, sheet_name='Consulta1')
    print(f"  Erros:  {len(err):,} linhas · Portal: {len(port):,} linhas")

    err['plano']  = err['ds_descricao'].map(PLAN_MAP)
    port['plano'] = port['ds_descricao'].map(PLAN_MAP)
    err['rank']   = err['cd_contagem'].map(RANK)

    # Parsear endereços
    addr_e = err['cd_posicao'].apply(parse_addr).apply(pd.Series)
    err    = pd.concat([err.reset_index(drop=True), addr_e], axis=1)
    addr_p = port['cd_posicao'].apply(parse_addr).apply(pd.Series)
    port   = pd.concat([port.reset_index(drop=True), addr_p], axis=1)

    # Colunas de tempo
    for c in ['a','b','c']:
        port[f'dt_contagem{c}_ini'] = pd.to_datetime(port[f'dt_contagem{c}_ini'], errors='coerce')
        port[f'dt_contagem{c}_fim'] = pd.to_datetime(port[f'dt_contagem{c}_fim'], errors='coerce')
        port[f'dur_{c}_min'] = (port[f'dt_contagem{c}_fim']-port[f'dt_contagem{c}_ini']).dt.total_seconds()/60

    err['dt_inicio']    = pd.to_datetime(err['dt_inicio'], errors='coerce')
    err['dt_conclusao'] = pd.to_datetime(err['dt_conclusao'], errors='coerce')
    err['dur_min']      = (err['dt_conclusao']-err['dt_inicio']).dt.total_seconds()/60

    # Analisar por plano
    print()
    print("🔄 Processando planos...")
    ALL = {}
    for plano in PLAN_ORDER:
        ep = err[err['plano']==plano].copy()
        pp = port[port['plano']==plano].copy()
        if len(ep) == 0:
            print(f"  ⚠️  {plano}: sem dados no Erros, pulando.")
            continue
        print(f"\n  📦 {plano} ({len(ep):,} linhas Erros / {len(pp):,} Portal)")
        ALL[plano] = build_plan(plano, ep, pp)
        d = ALL[plano]
        print(f"     Pos:{d['n_pos_total']} FinS:{d['n_fin_s']} FinC:{d['n_fin_c']} Pend:{d['n_pend']}")
        print(f"     LotesOK:{d['n_lot_ok']} Alertas:{d['n_lot_alert']} InvCod:{d['n_inv_cod']}")

    # Gerar dashboards
    print()
    print("🖥️  Gerando dashboards HTML...")

    DASHBOARDS = [
        ('📋 Análise de Erros — Corrigido',  '#0f172a 0%,#1a2a4a 100%', 'err',  analysis_plan, 'dashboard_erros_corrigido.html'),
        ('📋 Portal Analítico — Corrigido',  '#0f172a 0%,#0d2a1a 100%', 'port', analysis_plan, 'dashboard_portal_corrigido.html'),
        ('⏱ Análise de Erros — Tempo',       '#0f172a 0%,#1a2a4a 100%', 'err',  time_plan,     'dashboard_tempo_erros_v2.html'),
        ('⏱ Portal Analítico — Tempo',       '#0f172a 0%,#0d2a1a 100%', 'port', time_plan,     'dashboard_tempo_portal_v2.html'),
    ]
    for title, grad, source, fn, fname in DASHBOARDS:
        write_dashboard(title, grad, source, fn, fname, ALL, ENGINE_JS, MODAL_JS)

    # ── ANÁLISE SKU-LOTE COM BASE DE ESTOQUE ──────────────────────────────
    if stock_path:
        print()
        print("🔬 Gerando análise SKU-Lote com base de estoque...")
        try:
            _build_stock_analysis(err, stock_path, OUT_DIR)
            print("  ✅ Análise SKU-Lote - Cruzamento Estoque.html gerado!")
        except Exception as exc:
            print(f"  ⚠️  Erro na análise SKU-Lote: {exc}")

    ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    print()
    print(f"✅ Todos os dashboards atualizados com sucesso! [{ts}]")
    print(f"📁 Pasta: {OUT_DIR}")
    print()
    input("Pressione Enter para fechar...")

if __name__ == '__main__':
    main()
