// ============================================================
//  UPDATE MODAL — shared across all 4 dashboards
//  Injects the UI and wires up the file-picker
// ============================================================

(function() {
  "use strict";

  const MODAL_CSS = `
#upd-fab{position:fixed;bottom:24px;right:24px;z-index:9000;background:#1d4ed8;color:#fff;
  border:none;border-radius:50px;padding:12px 22px;font-size:14px;font-weight:700;
  cursor:pointer;box-shadow:0 4px 20px #0008;transition:all .2s;display:flex;align-items:center;gap:8px}
#upd-fab:hover{background:#2563eb;transform:translateY(-2px)}
#upd-overlay{display:none;position:fixed;inset:0;background:#000a;z-index:9001;align-items:center;justify-content:center}
#upd-overlay.open{display:flex}
#upd-modal{background:#1e293b;border:1px solid #334155;border-radius:14px;width:90%;max-width:560px;
  padding:28px 32px;box-shadow:0 8px 40px #0006}
#upd-modal h2{font-size:18px;font-weight:700;color:#fff;margin-bottom:6px}
#upd-modal p{font-size:13px;color:#94a3b8;margin-bottom:18px;line-height:1.6}
.upd-file-zone{background:#0f172a;border:2px dashed #334155;border-radius:8px;padding:18px;
  margin-bottom:14px;text-align:center;cursor:pointer;transition:border-color .2s}
.upd-file-zone:hover,.upd-file-zone.has-file{border-color:#3b82f6}
.upd-file-zone input{display:none}
.upd-file-zone label{cursor:pointer;font-size:13px;color:#94a3b8;display:block}
.upd-file-zone .file-name{color:#60a5fa;font-weight:600;font-size:13px;margin-top:6px;word-break:break-all}
#upd-progress-wrap{display:none;margin:14px 0}
#upd-progress-bar{height:8px;background:#0f172a;border-radius:4px;overflow:hidden;margin-bottom:6px}
#upd-progress-fill{height:100%;background:#3b82f6;width:0;border-radius:4px;transition:width .3s}
#upd-progress-msg{font-size:12px;color:#94a3b8}
#upd-error{color:#f87171;font-size:13px;margin-top:8px;display:none}
.upd-btn-row{display:flex;gap:10px;margin-top:18px}
.upd-btn{flex:1;padding:10px;border-radius:7px;border:none;font-size:14px;font-weight:700;cursor:pointer}
.upd-btn-primary{background:#1d4ed8;color:#fff} .upd-btn-primary:hover{background:#2563eb}
.upd-btn-secondary{background:#334155;color:#e2e8f0} .upd-btn-secondary:hover{background:#475569}
.upd-btn:disabled{opacity:.4;cursor:not-allowed}
#upd-ts{font-size:11px;color:#4ade80;margin-left:8px;font-weight:400}
.upd-hint{background:#172033;border-radius:6px;padding:10px 14px;font-size:11px;color:#64748b;margin-bottom:14px;line-height:1.7}
.upd-hint strong{color:#94a3b8}
`;

  const MODAL_HTML = `
<style>${MODAL_CSS}</style>
<button id="upd-fab" onclick="updOpen()">🔄 Atualizar<span id="upd-ts"></span></button>
<div id="upd-overlay">
  <div id="upd-modal">
    <h2>🔄 Atualizar Indicadores</h2>
    <p>Selecione os dois arquivos Excel de base para recalcular todos os indicadores do dashboard com os dados mais recentes.</p>
    <div class="upd-hint">
      <strong>Caminho padrão:</strong> <code>C:\\Users\\Cleuvo\\OneDrive\\</code><br>
      📄 <strong>Análise de Erros:</strong> Inventário Análise de Erros *.xlsx<br>
      📄 <strong>Portal Analítico:</strong> Inventário Analítico Portal *.xlsx
    </div>

    <div class="upd-file-zone" id="zone-err" onclick="document.getElementById('inp-err').click()">
      <input type="file" id="inp-err" accept=".xlsx" onchange="updFileSet('err',this)">
      <label for="inp-err">📂 <strong>Análise de Erros</strong> — clique para selecionar</label>
      <div class="file-name" id="name-err"></div>
    </div>

    <div class="upd-file-zone" id="zone-port" onclick="document.getElementById('inp-port').click()">
      <input type="file" id="inp-port" accept=".xlsx" onchange="updFileSet('port',this)">
      <label for="inp-port">📂 <strong>Portal Analítico</strong> — clique para selecionar</label>
      <div class="file-name" id="name-port"></div>
    </div>

    <div id="upd-progress-wrap">
      <div id="upd-progress-bar"><div id="upd-progress-fill"></div></div>
      <div id="upd-progress-msg">Aguardando...</div>
    </div>
    <div id="upd-error"></div>

    <div class="upd-btn-row">
      <button class="upd-btn upd-btn-secondary" onclick="updClose()">Cancelar</button>
      <button class="upd-btn upd-btn-primary" id="upd-run-btn" onclick="updRun()" disabled>▶ Calcular e Atualizar</button>
    </div>
  </div>
</div>`;

  // Inject on DOM ready
  document.addEventListener('DOMContentLoaded', () => {
    document.body.insertAdjacentHTML('beforeend', MODAL_HTML);
    // Restore last update timestamp
    const ts = localStorage.getItem('upd-last-ts');
    if (ts) document.getElementById('upd-ts').textContent = ' · ' + ts;
  });

  window.updOpen  = () => document.getElementById('upd-overlay').classList.add('open');
  window.updClose = () => {
    document.getElementById('upd-overlay').classList.remove('open');
    document.getElementById('upd-error').style.display='none';
    document.getElementById('upd-progress-wrap').style.display='none';
    document.getElementById('upd-progress-fill').style.width='0';
  };

  let _files = {err:null, port:null};
  window.updFileSet = (key, input) => {
    _files[key] = input.files[0] || null;
    const nameEl = document.getElementById(`name-${key}`);
    const zoneEl = document.getElementById(`zone-${key}`);
    if (_files[key]) {
      nameEl.textContent = '✓ ' + _files[key].name;
      zoneEl.classList.add('has-file');
    } else {
      nameEl.textContent = '';
      zoneEl.classList.remove('has-file');
    }
    document.getElementById('upd-run-btn').disabled = !(_files.err && _files.port);
  };

  function setProgress(msg, pct) {
    document.getElementById('upd-progress-wrap').style.display='block';
    document.getElementById('upd-progress-fill').style.width = pct+'%';
    document.getElementById('upd-progress-msg').textContent = msg;
  }

  window.updRun = async () => {
    const btn = document.getElementById('upd-run-btn');
    btn.disabled = true;
    document.getElementById('upd-error').style.display='none';
    try {
      // Load SheetJS if not already loaded
      if (typeof XLSX === 'undefined') {
        setProgress('Carregando SheetJS...', 5);
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
      }
      const result = await InventarioEngine.analyzeFiles(_files.err, _files.port, setProgress);
      setProgress('Atualizando dashboard...', 95);
      // Call dashboard-specific update function
      if (typeof window.dashboardUpdate === 'function') {
        window.dashboardUpdate(result.data, result.timestamp);
      }
      const ts = new Date().toLocaleString('pt-BR');
      localStorage.setItem('upd-last-ts', ts);
      document.getElementById('upd-ts').textContent = ' · ' + ts;
      setProgress('✅ Atualizado com sucesso!', 100);
      setTimeout(updClose, 1500);
    } catch(e) {
      const errEl = document.getElementById('upd-error');
      errEl.style.display='block';
      errEl.textContent = '❌ Erro: ' + e.message;
    } finally {
      btn.disabled = false;
    }
  };

  function loadScript(src) {
    return new Promise((res,rej)=>{
      const s=document.createElement('script');
      s.src=src; s.onload=res; s.onerror=rej;
      document.head.appendChild(s);
    });
  }
})();
