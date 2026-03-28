// ============================================================
//  ANALYSIS ENGINE — full inventory analysis in JavaScript
//  Uses SheetJS (xlsx) — loaded from CDN
// ============================================================

window.InventarioEngine = (function() {
  "use strict";

  const PLAN_MAP = {
    'INVENTÁRIO GERAL VCP CLIMATIZADO': 'Climatizado',
    'INVENTARIO ODONTO VCP': 'Odonto',
    'Inventario geral armazém 72 vcp': 'Armazém 72',
    'INVENTÁRIO GERAL VCP CLIMATIZADO 70.02.007 RUA 1 A 12 NÍVEL 7': 'Climatizado RUA7/Nív7',
  };
  const PLAN_ORDER = ['Climatizado','Odonto','Armazém 72','Climatizado RUA7/Nív7'];
  const RANK = {A:1,B:2,C:3};

  function parseAddr(s) {
    const p = String(s||'').split('.');
    return { arm:p[0]||'', area:p[1]||'', rua:p[2]||'', predio:p[3]||'', nivel:p[4]||'' };
  }

  function toNum(v) { return (v==null||v===''||isNaN(v)) ? 0 : Number(v); }
  function safeStr(v) { return (v==null||v===undefined) ? '—' : String(v); }

  // Parse Excel date serial to JS Date
  function excelDate(v) {
    if (!v) return null;
    if (v instanceof Date) return v;
    if (typeof v === 'string') return new Date(v);
    if (typeof v === 'number') {
      // Excel serial: days since 1899-12-30
      return new Date((v - 25569) * 86400 * 1000);
    }
    return null;
  }

  function durMin(ini, fim) {
    const a = excelDate(ini), b = excelDate(fim);
    if (!a || !b) return null;
    const d = (b - a) / 60000;
    return (d >= 0 && d <= 480) ? d : null;
  }

  function median(arr) {
    if (!arr.length) return null;
    const s = [...arr].sort((a,b)=>a-b);
    const m = Math.floor(s.length/2);
    return s.length%2===0 ? (s[m-1]+s[m])/2 : s[m];
  }

  function mean(arr) { return arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : null; }

  function classifyAlert(sumSaldo, sumCnt) {
    const div = sumCnt - sumSaldo;
    if (sumCnt === 0 && sumSaldo > 0) return 'Posição vazia/contagem=0';
    if (sumSaldo === 0 && sumCnt === 0) return 'Posição vazia correta';
    const pct = sumSaldo > 0 ? Math.abs(div/sumSaldo)*100 : 999;
    const ratio = sumSaldo > 0 ? sumCnt/sumSaldo : 0;
    const rnd = Math.round(ratio);
    if (rnd > 1 && Math.abs(ratio-rnd) < 0.05 && [2,3,4,5,6,10,12,25,50,60,100].includes(rnd))
      return `Erro de UM (fator x${rnd})`;
    if (pct < 10) return 'Divergência pequena (<10%)';
    if (div > 0) return 'Sobra significativa';
    return 'Falta significativa';
  }

  // ── PROCESS ERROS SHEET ──────────────────────────────────
  function processErros(rows) {
    // Attach plan, rank, address
    rows.forEach(r => {
      r._plano = PLAN_MAP[r.ds_descricao] || '(outro)';
      r._rank  = RANK[r.cd_contagem] || 0;
      const addr = parseAddr(r.cd_posicao);
      Object.assign(r, addr);
      r._dur = durMin(r.dt_inicio, r.dt_conclusao);
    });

    const result = {};
    for (const plano of PLAN_ORDER) {
      const ep = rows.filter(r => r._plano === plano);
      if (!ep.length) continue;

      // Group by position+SKU+lot
      const groups = {};
      ep.forEach(r => {
        const key = `${r.cd_posicao}||${r.nr_produto}||${r.cd_lote}`;
        if (!groups[key]) groups[key] = [];
        groups[key].push(r);
      });

      // For each group: first count (min rank) = saldo_sis; last count (max rank) = cnt_final
      const pos = Object.values(groups).map(g => {
        const sorted = [...g].sort((a,b) => a._rank - b._rank);
        const first = sorted[0];
        const last  = sorted[sorted.length-1];
        return {
          cd_posicao: first.cd_posicao,
          nr_produto: first.nr_produto,
          cd_lote:    first.cd_lote,
          ds_produto: first.ds_produto,
          rua: first.rua, area: first.area, predio: first.predio, nivel: first.nivel,
          saldo_sis:    toNum(first.qt_saldoinicial),   // ← Count A saldo
          cnt_final:    toNum(last.qt_contagem),         // ← Last count qty
          first_cnt:    first.cd_contagem,
          last_cnt:     last.cd_contagem,
          status_final: last.ds_statuscontagem,
          cod_xfis:     last['COD X FIS'],
          lot_xfis:     last['LOTE X FIS'],
          nr_produtocontagem: last.nr_produtocontagem,
          ds_produtocontagem: last.ds_produtocontagem,
          cd_lotecontagem:    last.cd_lotecontagem,
        };
      });

      pos.forEach(p => {
        p.net_div    = Math.round((p.cnt_final - p.saldo_sis)*100)/100;
        p.is_fin_s   = p.status_final === 'FINALIZADO S/ DIVERGÊNCIA';
        p.is_fin_c   = p.status_final === 'FINALIZADO C/ DIVERGÊNCIA';
        p.is_pend    = p.status_final === 'PENDENTE';
        p.is_em_c    = p.status_final === 'EM CONTAGEM';
        p.is_cod_err = p.cod_xfis === 'ERRADO';
        p.is_lot_err = p.lot_xfis === 'ERRADO';
      });

      // Inversions (distinct positions, last count still ERRADO)
      const lastByPos = {};
      ep.forEach(r => {
        const key = r.cd_posicao;
        if (!lastByPos[key] || r._rank > lastByPos[key]._rank) lastByPos[key] = r;
      });
      const firstByPos = {};
      ep.forEach(r => {
        const key = r.cd_posicao;
        if (!firstByPos[key] || r._rank < firstByPos[key]._rank) firstByPos[key] = r;
      });

      const inv_cod = Object.values(lastByPos)
        .filter(r => r['COD X FIS'] === 'ERRADO')
        .map(r => ({
          cd_posicao: r.cd_posicao, ds_produto: safeStr(r.ds_produto),
          cd_lote: safeStr(r.cd_lote), nr_produtocontagem: r.nr_produtocontagem,
          ds_produtocontagem: safeStr(r.ds_produtocontagem), cd_lotecontagem: safeStr(r.cd_lotecontagem),
          qt_saldoinicial: toNum((firstByPos[r.cd_posicao]||r).qt_saldoinicial),
          qt_contagem: toNum(r.qt_contagem),
        }));

      const inv_lot = Object.values(lastByPos)
        .filter(r => r['LOTE X FIS'] === 'ERRADO' && r['COD X FIS'] === 'CERTO')
        .map(r => ({
          cd_posicao: r.cd_posicao, ds_produto: safeStr(r.ds_produto),
          cd_lote: safeStr(r.cd_lote), cd_lotecontagem: safeStr(r.cd_lotecontagem),
          qt_saldoinicial: toNum((firstByPos[r.cd_posicao]||r).qt_saldoinicial),
          qt_contagem: toNum(r.qt_contagem),
        }));

      // SKU+lot aggregation
      const sgMap = {};
      pos.forEach(p => {
        const key = `${p.nr_produto}||${p.cd_lote}`;
        if (!sgMap[key]) sgMap[key] = { nr_produto: p.nr_produto, cd_lote: p.cd_lote,
          ds_produto: p.ds_produto, n_pos:0, n_ok:0, sum_saldo:0, sum_cnt:0, posicoes:[] };
        const g = sgMap[key];
        g.n_pos++; g.n_ok += p.is_fin_s?1:0;
        g.sum_saldo += p.saldo_sis; g.sum_cnt += p.cnt_final;
        g.posicoes.push(p.cd_posicao);
      });

      const validated=[], lot_alert=[];
      Object.values(sgMap).forEach(g => {
        g.all_ok  = g.n_pos === g.n_ok;
        g.net_div = Math.round((g.sum_cnt - g.sum_saldo)*100)/100;
        if (g.all_ok && Math.abs(g.net_div)<0.01) {
          validated.push({...g, sum_contagem: g.sum_cnt});
        } else if (g.all_ok && Math.abs(g.net_div)>=0.01) {
          const tipo = classifyAlert(g.sum_saldo, g.sum_cnt);
          const pct  = g.sum_saldo>0 ? Math.round(g.net_div/g.sum_saldo*1000)/10 : 0;
          lot_alert.push({...g, sum_contagem: g.sum_cnt, tipo, pct_div: pct,
            posicoes: g.posicoes.slice(0,5).join(', ')});
        }
      });
      lot_alert.sort((a,b) => Math.abs(b.net_div)-Math.abs(a.net_div));

      // Pending counts (distinct positions per count level)
      const cntStats = cnt => {
        const sub = ep.filter(r => r.cd_contagem === cnt);
        const uniq = field => new Set(sub.filter(r=>r.ds_statuscontagem===field).map(r=>r.cd_posicao)).size;
        return {
          pend_total: uniq('PENDENTE'), em_cnt: uniq('EM CONTAGEM'),
          fin_s: uniq('FINALIZADO S/ DIVERGÊNCIA'), fin_c: uniq('FINALIZADO C/ DIVERGÊNCIA'),
          total: new Set(sub.map(r=>r.cd_posicao)).size
        };
      };

      // Time analysis
      const validDur = ep.filter(r=>r._dur!=null);
      const timeDim = (dim, cnt) => {
        const sub = validDur.filter(r=>r.cd_contagem===cnt);
        const byDim = {};
        sub.forEach(r=>{ const dv=r[dim]; if(!byDim[dv])byDim[dv]=[]; byDim[dv].push(r._dur); });
        return Object.entries(byDim).map(([dv,vals])=>({
          dim_val:dv, cnt, media:mean(vals), mediana:median(vals), n:vals.length
        }));
      };
      const time_err = {};
      ['area','rua','predio','nivel'].forEach(dim=>{
        time_err[dim]=[...timeDim(dim,'A'),...timeDim(dim,'B'),...timeDim(dim,'C')];
      });

      // Pending time estimate (distinct positions)
      const pend_est_err = [];
      const globalMed = median(validDur.map(r=>r._dur)) || 2;
      ['A','B','C'].forEach(cnt=>{
        const done = validDur.filter(r=>r.cd_contagem===cnt);
        const overallMed = median(done.map(r=>r._dur)) || globalMed;
        const pendRows = ep.filter(r=>r.cd_contagem===cnt&&['PENDENTE','EM CONTAGEM'].includes(r.ds_statuscontagem));
        const byRua = {};
        pendRows.forEach(r=>{ if(!byRua[r.rua])byRua[r.rua]=new Set(); byRua[r.rua].add(r.cd_posicao); });
        Object.entries(byRua).forEach(([rua,posSet])=>{
          const n = posSet.size; // DISTINCT positions
          const ruaDone = done.filter(r=>r.rua===rua).map(r=>r._dur);
          const med = (ruaDone.length ? median(ruaDone) : null) || overallMed;
          pend_est_err.push({rua, cnt, n_pend:n, med_min:Math.round(med*100)/100,
            est_min:Math.round(med*n*10)/10, est_h:Math.round(med*n/60*100)/100});
        });
      });

      // Overall time stats
      const overall_err = ['A','B','C'].map(cnt=>{
        const sub = validDur.filter(r=>r.cd_contagem===cnt).map(r=>r._dur);
        if(!sub.length) return null;
        const sorted=[...sub].sort((a,b)=>a-b);
        return {cnt, media:Math.round(mean(sub)*100)/100, mediana:Math.round(median(sub)*100)/100,
          p25:Math.round(sorted[Math.floor(sorted.length*.25)]*100)/100,
          p75:Math.round(sorted[Math.floor(sorted.length*.75)]*100)/100,
          n:sub.length, total_h:Math.round(sub.reduce((a,b)=>a+b,0)/60*10)/10};
      }).filter(Boolean);

      // Histogram
      const BINS=[0,1,2,5,15,30,480], LBLS=['<1min','1-2min','2-5min','5-15min','15-30min','>30min'];
      const hists_err={};
      ['A','B','C'].forEach(cnt=>{
        const sub=validDur.filter(r=>r.cd_contagem===cnt).map(r=>r._dur);
        if(!sub.length) return;
        const counts=new Array(LBLS.length).fill(0);
        sub.forEach(v=>{ for(let i=0;i<BINS.length-1;i++){if(v>=BINS[i]&&v<BINS[i+1]){counts[i]++;break;}} });
        hists_err[cnt]=counts;
      });

      // Pareto
      const pareto = [...pos].filter(p=>Math.abs(p.net_div)>0)
        .sort((a,b)=>Math.abs(b.net_div)-Math.abs(a.net_div)).slice(0,20)
        .map(p=>({cd_posicao:p.cd_posicao, ds_produto:safeStr(p.ds_produto),
          cd_lote:safeStr(p.cd_lote), saldo:p.saldo_sis, contagem:p.cnt_final,
          net_div:p.net_div, status:safeStr(p.status_final)}));

      // Counters
      const ctd={};
      ['A','B','C'].forEach(cnt=>{
        const sub=validDur.filter(r=>r.cd_contagem===cnt&&r.ds_usuario);
        const byUser={};
        sub.forEach(r=>{ if(!byUser[r.ds_usuario])byUser[r.ds_usuario]={n:0,dur:0};
          byUser[r.ds_usuario].n++; byUser[r.ds_usuario].dur+=r._dur; });
        ctd[cnt]=Object.entries(byUser).map(([u,d])=>({user:u,n:d.n,dur:Math.round(d.dur),avg:Math.round(d.dur/d.n*100)/100}))
          .sort((a,b)=>b.n-a.n).slice(0,15);
      });

      result[plano]={
        nome:plano,
        n_pos_total: new Set(ep.map(r=>r.cd_posicao)).size,
        n_fin_s: pos.filter(p=>p.is_fin_s).length,
        n_fin_c: pos.filter(p=>p.is_fin_c).length,
        n_pend:  pos.filter(p=>p.is_pend).length,
        n_em_c:  pos.filter(p=>p.is_em_c).length,
        n_inv_cod: inv_cod.length, n_inv_lot: inv_lot.length,
        inv_cod: inv_cod.slice(0,30), inv_lot: inv_lot.slice(0,50),
        n_sku_ok: new Set(validated.map(v=>v.nr_produto)).size,
        n_lot_ok: validated.length, n_lot_alert: lot_alert.length,
        n_sku_total: new Set(ep.map(r=>r.nr_produto)).size,
        n_lot_total: new Set(ep.map(r=>`${r.nr_produto}||${r.cd_lote}`)).size,
        sku_ok: validated, lot_alert,
        err_cnt_A: cntStats('A'), err_cnt_B: cntStats('B'), err_cnt_C: cntStats('C'),
        pareto, ctd_A:ctd.A||[], ctd_B:ctd.B||[], ctd_C:ctd.C||[],
        time_err, overall_err, hists_err, pend_est_err,
      };
    }
    return result;
  }

  // ── PROCESS PORTAL SHEET ────────────────────────────────
  function processPortal(rows, errosResult) {
    rows.forEach(r => {
      r._plano = PLAN_MAP[r.ds_descricao] || '(outro)';
      const addr = parseAddr(r.cd_posicao);
      Object.assign(r, addr);
      r.dur_a = durMin(r.dt_contagema_ini, r.dt_contagema_fim);
      r.dur_b = durMin(r.dt_contagemb_ini, r.dt_contagemb_fim);
      r.dur_c = durMin(r.dt_contagemc_ini, r.dt_contagemc_fim);
    });

    for (const plano of PLAN_ORDER) {
      const pp = rows.filter(r=>r._plano===plano);
      if (!pp.length || !errosResult[plano]) continue;
      const D = errosResult[plano];

      const portCntStats = letter => {
        const pend = new Set(pp.filter(r=>(r.statusatual||'').startsWith(`${letter} - Em Andamento`)).map(r=>r.cd_posicao)).size;
        const conc = new Set(pp.filter(r=>(r.statusatual||'').startsWith(`${letter} - Concluído`)).map(r=>r.cd_posicao)).size;
        return {pend,conc,total:pend+conc};
      };
      const pA=portCntStats('A'), pB=portCntStats('B'), pC=portCntStats('C');

      D.port_pend_a=pA.pend; D.port_conc_a=pA.conc;
      D.port_pend_b=pB.pend; D.port_conc_b=pB.conc;
      D.port_pend_c=pC.pend; D.port_conc_c=pC.conc;
      D.port_tot_pos=new Set(pp.map(r=>r.cd_posicao)).size;

      // Time by dimension (portal)
      const time_port={};
      ['area','rua','predio','nivel'].forEach(dim=>{
        time_port[dim]=[];
        [{c:'a',f:'dur_a'},{c:'b',f:'dur_b'},{c:'c',f:'dur_c'}].forEach(({c,f})=>{
          const sub=pp.filter(r=>r[f]!=null);
          const byDim={};
          sub.forEach(r=>{ const dv=r[dim]; if(!byDim[dv])byDim[dv]=[]; byDim[dv].push(r[f]); });
          Object.entries(byDim).forEach(([dv,vals])=>{
            time_port[dim].push({dim_val:dv,cnt:c.toUpperCase(),media:mean(vals),mediana:median(vals),n:vals.length});
          });
        });
      });
      D.time_port=time_port;

      const BINS=[0,1,2,5,15,30,480],LBLS=['<1min','1-2min','2-5min','5-15min','15-30min','>30min'];
      const hists_port={};
      [{c:'A',f:'dur_a'},{c:'B',f:'dur_b'},{c:'C',f:'dur_c'}].forEach(({c,f})=>{
        const sub=pp.map(r=>r[f]).filter(v=>v!=null);
        if(!sub.length) return;
        const counts=new Array(LBLS.length).fill(0);
        sub.forEach(v=>{for(let i=0;i<BINS.length-1;i++){if(v>=BINS[i]&&v<BINS[i+1]){counts[i]++;break;}}});
        hists_port[c]=counts;
      });
      D.hists_port=hists_port;

      const overall_port=['a','b','c'].map(c=>{
        const sub=pp.map(r=>r[`dur_${c}`]).filter(v=>v!=null);
        if(!sub.length) return null;
        const sorted=[...sub].sort((a,b)=>a-b);
        return {cnt:c.toUpperCase(),media:Math.round(mean(sub)*100)/100,
          mediana:Math.round(median(sub)*100)/100,
          p25:Math.round(sorted[Math.floor(sorted.length*.25)]*100)/100,
          p75:Math.round(sorted[Math.floor(sorted.length*.75)]*100)/100,
          n:sub.length,total_h:Math.round(sub.reduce((a,b)=>a+b,0)/60*10)/10};
      }).filter(Boolean);
      D.overall_port=overall_port;

      // Pending time estimate portal
      const globalMed=median(pp.map(r=>r.dur_a).filter(v=>v!=null))||2;
      const pend_est_port=[];
      [{c:'A',f:'dur_a',stat:'A - Em Andamento'},{c:'B',f:'dur_b',stat:'B - Em Andamento'},{c:'C',f:'dur_c',stat:'C - Em Andamento'}].forEach(({c,f,stat})=>{
        const done=pp.filter(r=>r[f]!=null);
        const overallMed=median(done.map(r=>r[f]))||globalMed;
        const pendRows=pp.filter(r=>(r.statusatual||'').startsWith(stat));
        const byRua={};
        pendRows.forEach(r=>{ if(!byRua[r.rua])byRua[r.rua]=new Set(); byRua[r.rua].add(r.cd_posicao); });
        Object.entries(byRua).forEach(([rua,posSet])=>{
          const n=posSet.size;
          const ruaMed=median(done.filter(r=>r.rua===rua).map(r=>r[f]))||overallMed;
          pend_est_port.push({rua,cnt:c,n_pend:n,med_min:Math.round(ruaMed*100)/100,
            est_min:Math.round(ruaMed*n*10)/10,est_h:Math.round(ruaMed*n/60*100)/100});
        });
      });
      D.pend_est_port=pend_est_port;
    }
    return errosResult;
  }

  // ── MAIN ENTRY POINT ──────────────────────────────────────
  async function analyzeFiles(errosFile, portalFile, onProgress) {
    onProgress && onProgress('Lendo Análise de Erros...', 10);
    const errosWb  = XLSX.read(await errosFile.arrayBuffer(), {type:'array', cellDates:true});
    onProgress && onProgress('Lendo Analítico Portal...', 30);
    const portalWb = XLSX.read(await portalFile.arrayBuffer(), {type:'array', cellDates:true});

    onProgress && onProgress('Processando Análise de Erros...', 50);
    const errosSheet  = errosWb.Sheets['ANALISE INVENTARIO'];
    const portalSheet = portalWb.Sheets['Consulta1'];

    if (!errosSheet) throw new Error('Aba "ANALISE INVENTARIO" não encontrada no arquivo de Erros!');
    if (!portalSheet) throw new Error('Aba "Consulta1" não encontrada no arquivo Portal!');

    const errosRows  = XLSX.utils.sheet_to_json(errosSheet,  {defval:null, raw:false, dateNF:'yyyy-mm-dd hh:mm:ss'});
    onProgress && onProgress('Processando Portal...', 65);
    const portalRows = XLSX.utils.sheet_to_json(portalSheet, {defval:null, raw:false, dateNF:'yyyy-mm-dd hh:mm:ss'});

    onProgress && onProgress('Calculando métricas...', 80);
    const errosResult  = processErros(errosRows);
    const combined     = processPortal(portalRows, errosResult);

    onProgress && onProgress('Concluído!', 100);
    return {data: combined, timestamp: new Date().toLocaleString('pt-BR')};
  }

  return { analyzeFiles, PLAN_ORDER };
})();
