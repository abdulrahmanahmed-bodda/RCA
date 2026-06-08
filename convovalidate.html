<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ConvoValidate — RCA Platform</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#ffffff;--bg2:#f8f8f7;--bg3:#f2f2f0;
  --text:#1a1a1a;--text2:#555;--text3:#999;
  --border:rgba(0,0,0,0.12);--border2:rgba(0,0,0,0.2);
  --radius:8px;--radius-lg:12px;
  --green:#639922;--green-bg:#EAF3DE;--green-text:#27500A;
  --red:#E24B4A;--red-bg:#FCEBEB;--red-text:#791F1F;
  --amber:#EF9F27;--amber-bg:#FAEEDA;--amber-text:#633806;
  --blue:#185FA5;--blue-bg:#E6F1FB;--blue-text:#0C447C;
  --gray-bg:#F1EFE8;--gray-text:#444441;
}
@media(prefers-color-scheme:dark){
  :root{
    --bg:#1e1e1c;--bg2:#282826;--bg3:#303030;
    --text:#e8e8e6;--text2:#aaa;--text3:#666;
    --border:rgba(255,255,255,0.1);--border2:rgba(255,255,255,0.2);
  }
}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg3);color:var(--text);font-size:14px;height:100vh;overflow:hidden}
.app{display:flex;flex-direction:column;height:100vh}
.topbar{background:var(--bg);border-bottom:1px solid var(--border);padding:0 20px;display:flex;align-items:center;gap:12px;height:52px;flex-shrink:0}
.topbar-title{font-weight:600;font-size:15px;display:flex;align-items:center;gap:8px}
.topbar-sub{font-size:11px;color:var(--text3)}
.badge{font-size:11px;padding:2px 8px;border-radius:20px;font-weight:500;background:var(--green-bg);color:var(--green-text)}
.topbar-actions{margin-left:auto;display:flex;gap:8px;align-items:center}
.btn{display:inline-flex;align-items:center;gap:5px;padding:6px 13px;border-radius:var(--radius);border:1px solid var(--border2);background:var(--bg);color:var(--text);font-size:13px;cursor:pointer;font-family:inherit;transition:background .15s}
.btn:hover{background:var(--bg2)}
.btn-primary{background:var(--green);border-color:var(--green);color:#fff}
.btn-primary:hover{background:#3B6D11}
.btn-danger{background:var(--red);border-color:var(--red);color:#fff}
.btn-danger:hover{background:#A32D2D}
.btn-sm{padding:4px 10px;font-size:12px}
.main{display:flex;flex:1;overflow:hidden}
.sidebar{width:210px;background:var(--bg);border-right:1px solid var(--border);padding:10px;flex-shrink:0;overflow-y:auto;display:flex;flex-direction:column;gap:2px}
.sidebar-label{font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;padding:10px 8px 4px}
.nav-item{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:var(--radius);cursor:pointer;font-size:13px;color:var(--text2);transition:background .12s}
.nav-item:hover{background:var(--bg2)}
.nav-item.active{background:var(--green-bg);color:var(--green-text);font-weight:500}
.nav-item i{font-size:15px;flex-shrink:0}
.nbadge{margin-left:auto;font-size:11px;padding:1px 6px;border-radius:20px;font-weight:500}
.nb-green{background:var(--green-bg);color:var(--green-text)}
.nb-red{background:var(--red-bg);color:var(--red-text)}
.nb-amber{background:var(--amber-bg);color:var(--amber-text)}
.nb-gray{background:var(--gray-bg);color:var(--gray-text)}
.divider{height:1px;background:var(--border);margin:6px 4px}
.content{flex:1;overflow:hidden;display:flex;flex-direction:column}
.toolbar{background:var(--bg);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;flex-shrink:0}
.stats-row{display:flex;gap:10px;padding:10px 14px;background:var(--bg2);border-bottom:1px solid var(--border);flex-shrink:0;overflow-x:auto}
.stat{background:var(--bg);border:1px solid var(--border);border-radius:var(--radius);padding:8px 14px;min-width:100px;flex-shrink:0}
.stat-lbl{font-size:11px;color:var(--text3);margin-bottom:3px}
.stat-n{font-size:20px;font-weight:600}
.c-red{color:#A32D2D}.c-green{color:#3B6D11}.c-amber{color:#854F0B}
.table-wrap{flex:1;overflow:auto;padding:14px}
table{width:100%;border-collapse:collapse;background:var(--bg);border:1px solid var(--border);border-radius:var(--radius-lg);overflow:hidden;font-size:12px;min-width:700px}
thead{background:var(--bg2)}
th{padding:9px 10px;text-align:left;font-weight:600;font-size:11px;color:var(--text2);border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--bg2)}
tr.sel td{background:rgba(99,153,34,0.07)}
.pill{display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:20px;font-size:11px;font-weight:500;white-space:nowrap}
.p-pos{background:var(--green-bg);color:var(--green-text)}
.p-neg{background:var(--red-bg);color:var(--red-text)}
.p-neu{background:var(--gray-bg);color:var(--gray-text)}
.p-mix{background:var(--amber-bg);color:var(--amber-text)}
.p-unsat{background:var(--red-bg);color:var(--red-text)}
.p-sat{background:var(--green-bg);color:var(--green-text)}
.p-pend{background:var(--amber-bg);color:var(--amber-text)}
.cell-txt{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:block;max-width:200px}
.act-btns{display:flex;gap:3px}
.ib{width:26px;height:26px;border-radius:var(--radius);border:1px solid var(--border);background:transparent;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;color:var(--text2);font-size:14px}
.ib:hover{background:var(--bg2)}
.empty{text-align:center;padding:60px 20px;color:var(--text3)}
.empty i{font-size:36px;margin-bottom:10px;display:block}
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:200}
.modal{background:var(--bg);border-radius:var(--radius-lg);border:1px solid var(--border);width:520px;max-width:96vw;max-height:90vh;display:flex;flex-direction:column;box-shadow:0 8px 32px rgba(0,0,0,.2)}
.modal-lg{width:620px}
.mhead{padding:16px 18px 12px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.mtitle{font-weight:600;font-size:14px;display:flex;align-items:center;gap:8px}
.mbody{padding:18px;overflow-y:auto;flex:1}
.mfoot{padding:12px 18px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px;flex-shrink:0}
.fg{margin-bottom:14px}
.flabel{font-size:11px;font-weight:600;color:var(--text2);margin-bottom:5px;display:block}
input[type=text],input[type=url],select,textarea{width:100%;padding:7px 10px;border-radius:var(--radius);border:1px solid var(--border2);background:var(--bg);color:var(--text);font-family:inherit;font-size:13px}
textarea{resize:vertical;min-height:72px}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--green);box-shadow:0 0 0 2px rgba(99,153,34,.15)}
.drop-zone{border:2px dashed var(--border2);border-radius:var(--radius-lg);padding:28px;text-align:center;cursor:pointer;color:var(--text2);transition:all .15s}
.drop-zone:hover,.drop-zone.drag-over{border-color:var(--green);background:rgba(99,153,34,.05)}
.col-chip{display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border-radius:20px;background:var(--bg2);border:1px solid var(--border);font-size:12px;margin:2px}
.col-chip .rm{border:none;background:none;cursor:pointer;color:var(--text3);font-size:14px;padding:0 1px;line-height:1}
.col-chip .rm:hover{color:#A32D2D}
.n8n-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.dot-on{background:var(--green)}.dot-off{background:#888}
.trig-opt{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:var(--radius);border:1px solid var(--border);margin-bottom:8px;cursor:pointer;transition:all .15s}
.trig-opt:hover{background:var(--bg2)}
.trig-opt.on{border-color:var(--green);background:rgba(99,153,34,.05)}
.wf-card{background:var(--bg);border:1px solid var(--border);border-radius:var(--radius-lg);padding:14px;margin-bottom:12px}
.rca-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}
.rca-card{background:var(--bg);border:1px solid var(--border);border-radius:var(--radius-lg);padding:14px}
.prog{height:5px;background:var(--bg2);border-radius:3px;overflow:hidden;margin-top:3px}
.prog-fill{height:100%;border-radius:3px}
.inline-select{font-size:11px;padding:2px 6px;border-radius:4px;border:1px solid var(--border);background:var(--bg);color:var(--text);font-family:inherit;cursor:pointer;max-width:130px}
.inline-select:focus{outline:none;border-color:var(--green)}
pre{white-space:pre-wrap;word-break:break-all}
</style>
</head>
<body>
<div class="app">

<div class="topbar">
  <div>
    <div class="topbar-title"><i class="ti ti-analyze"></i> ConvoValidate <span class="badge">RCA</span></div>
    <div class="topbar-sub">Unsatisfactory Root Cause Analysis</div>
  </div>
  <div class="topbar-actions">
    <div style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2)">
      <div class="n8n-dot dot-off" id="top-dot"></div>
      <span id="top-n8n-lbl">n8n off</span>
    </div>
    <button class="btn btn-sm" onclick="openM('import')"><i class="ti ti-upload"></i> Import</button>
    <button class="btn btn-sm btn-primary" onclick="openM('cols')"><i class="ti ti-table-column"></i> Columns</button>
    <button class="btn btn-sm" onclick="openM('n8n')"><i class="ti ti-topology-star"></i> n8n</button>
    <button class="btn btn-sm" onclick="openM('add')"><i class="ti ti-plus"></i> Add row</button>
  </div>
</div>

<div class="main">
<div class="sidebar">
  <div class="sidebar-label">Views</div>
  <div class="nav-item active" id="nav-table" onclick="switchView('table',this)"><i class="ti ti-table"></i> Conversations</div>
  <div class="nav-item" id="nav-rca" onclick="switchView('rca',this)"><i class="ti ti-chart-pie"></i> RCA Dashboard</div>
  <div class="nav-item" id="nav-wf" onclick="switchView('wf',this)"><i class="ti ti-topology-star"></i> Workflows</div>
  <div class="divider"></div>
  <div class="sidebar-label">Status</div>
  <div class="nav-item" onclick="setStatusFilter('',this)"><i class="ti ti-filter"></i> All <span class="nbadge nb-gray" id="nb-all">0</span></div>
  <div class="nav-item" onclick="setStatusFilter('Unsatisfactory',this)"><i class="ti ti-circle-x" style="color:var(--red)"></i> Unsatisfactory <span class="nbadge nb-red" id="nb-unsat">0</span></div>
  <div class="nav-item" onclick="setStatusFilter('Satisfactory',this)"><i class="ti ti-circle-check" style="color:var(--green)"></i> Satisfactory <span class="nbadge nb-green" id="nb-sat">0</span></div>
  <div class="nav-item" onclick="setStatusFilter('Pending',this)"><i class="ti ti-clock" style="color:var(--amber)"></i> Pending <span class="nbadge nb-amber" id="nb-pend">0</span></div>
  <div class="divider"></div>
  <div class="sidebar-label">Sentiment</div>
  <div class="nav-item" onclick="setSentFilter('Positive',this)"><i class="ti ti-mood-smile" style="color:var(--green)"></i> Positive <span class="nbadge nb-green" id="nb-pos">0</span></div>
  <div class="nav-item" onclick="setSentFilter('Negative',this)"><i class="ti ti-mood-sad" style="color:var(--red)"></i> Negative <span class="nbadge nb-red" id="nb-neg">0</span></div>
  <div class="nav-item" onclick="setSentFilter('Neutral',this)"><i class="ti ti-mood-empty" style="color:#888"></i> Neutral <span class="nbadge nb-gray" id="nb-neu">0</span></div>
  <div class="nav-item" onclick="setSentFilter('Mixed',this)"><i class="ti ti-mood-wink" style="color:var(--amber)"></i> Mixed <span class="nbadge nb-amber" id="nb-mix">0</span></div>
  <div class="divider"></div>
  <button class="btn btn-sm" style="width:100%;justify-content:center" onclick="exportCSV()"><i class="ti ti-download"></i> Export CSV</button>
</div>

<div class="content">
  <div id="view-table" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
    <div class="toolbar">
      <input type="text" id="q" placeholder="Search…" style="width:200px" oninput="render()">
      <button class="btn btn-sm" onclick="selAll()"><i class="ti ti-checkbox"></i> All</button>
      <button class="btn btn-sm" onclick="selNone()"><i class="ti ti-square"></i> None</button>
      <span id="sel-info" style="font-size:12px;color:var(--text3)"></span>
      <button class="btn btn-sm btn-danger" id="del-btn" style="display:none" onclick="delSel()"><i class="ti ti-trash"></i> Delete</button>
    </div>
    <div class="stats-row">
      <div class="stat"><div class="stat-lbl">Total</div><div class="stat-n" id="st-total">0</div></div>
      <div class="stat"><div class="stat-lbl">Unsatisfactory</div><div class="stat-n c-red" id="st-unsat">0</div></div>
      <div class="stat"><div class="stat-lbl">Satisfactory</div><div class="stat-n c-green" id="st-sat">0</div></div>
      <div class="stat"><div class="stat-lbl">Pending</div><div class="stat-n c-amber" id="st-pend">0</div></div>
      <div class="stat"><div class="stat-lbl">Negative</div><div class="stat-n c-red" id="st-neg">0</div></div>
      <div class="stat"><div class="stat-lbl">Positive</div><div class="stat-n c-green" id="st-pos">0</div></div>
    </div>
    <div class="table-wrap">
      <table id="tbl" style="display:none">
        <thead id="thead"></thead>
        <tbody id="tbody"></tbody>
      </table>
      <div id="empty" class="empty">
        <i class="ti ti-file-import"></i>
        <div style="font-size:14px;font-weight:600;margin-bottom:6px">No conversations yet</div>
        <div style="margin-bottom:14px">Import a CSV / Excel file or add rows manually</div>
        <div style="display:flex;gap:8px;justify-content:center">
          <button class="btn btn-primary" onclick="openM('import')"><i class="ti ti-upload"></i> Import file</button>
          <button class="btn" onclick="openM('add')"><i class="ti ti-plus"></i> Add row</button>
        </div>
      </div>
    </div>
  </div>

  <div id="view-rca" style="display:none;flex:1;overflow:auto;padding:16px">
    <div style="font-size:15px;font-weight:600;margin-bottom:14px">Root Cause Analysis Dashboard</div>
    <div id="rca-content"></div>
  </div>

  <div id="view-wf" style="display:none;flex:1;overflow:auto;padding:16px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">
      <div style="font-size:15px;font-weight:600">n8n Workflow Integrations</div>
      <button class="btn btn-sm btn-primary" onclick="openM('n8n')"><i class="ti ti-settings"></i> Configure</button>
    </div>
    <div id="wf-content"></div>
  </div>
</div>
</div>
</div>

<!-- IMPORT MODAL -->
<div id="m-import" class="overlay" style="display:none">
<div class="modal">
  <div class="mhead"><div class="mtitle"><i class="ti ti-upload"></i> Import File</div><button class="ib" onclick="closeM('import')"><i class="ti ti-x"></i></button></div>
  <div class="mbody">
    <div class="drop-zone" id="dz" onclick="document.getElementById('fi').click()" ondragover="event.preventDefault();this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="onDrop(event)">
      <i class="ti ti-table-import" style="font-size:32px;color:var(--green);display:block;margin-bottom:8px"></i>
      <div style="font-weight:600;margin-bottom:4px">Drop CSV or Excel here</div>
      <div style="font-size:12px">.csv · .xlsx · .xls</div>
      <input type="file" id="fi" accept=".csv,.xlsx,.xls" style="display:none" onchange="onFile(this.files[0])">
    </div>
    <div id="imp-preview" style="display:none;margin-top:14px">
      <div style="font-size:11px;font-weight:600;margin-bottom:6px">Preview (first 5 rows)</div>
      <div id="prev-tbl" style="overflow:auto;max-height:180px;border:1px solid var(--border);border-radius:var(--radius)"></div>
      <div style="font-size:11px;color:var(--text3);margin-top:6px" id="imp-info"></div>
    </div>
    <div style="margin-top:14px;display:grid;grid-template-columns:1fr 1fr;gap:10px">
      <div><label class="flabel">Conversation text column</label><select id="map-text"><option value="">— select —</option></select></div>
      <div><label class="flabel">Status column (optional)</label><select id="map-status"><option value="">— auto —</option></select></div>
    </div>
    <div style="margin-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:10px">
      <div><label class="flabel">Sentiment column (optional)</label><select id="map-sent"><option value="">— none —</option></select></div>
      <div><label class="flabel">Root cause column (optional)</label><select id="map-rca"><option value="">— none —</option></select></div>
    </div>
  </div>
  <div class="mfoot">
    <button class="btn" onclick="closeM('import')">Cancel</button>
    <button class="btn btn-primary" id="imp-btn" disabled onclick="doImport()"><i class="ti ti-check"></i> Import</button>
  </div>
</div>
</div>

<!-- COLUMNS MODAL -->
<div id="m-cols" class="overlay" style="display:none">
<div class="modal">
  <div class="mhead"><div class="mtitle"><i class="ti ti-table-column"></i> Manage Columns</div><button class="ib" onclick="closeM('cols')"><i class="ti ti-x"></i></button></div>
  <div class="mbody">
    <div style="font-size:12px;font-weight:600;margin-bottom:10px">Active columns</div>
    <div id="col-chips" style="display:flex;flex-wrap:wrap;margin-bottom:16px"></div>
    <div style="height:1px;background:var(--border);margin-bottom:14px"></div>
    <div style="font-size:12px;font-weight:600;margin-bottom:10px">Add new column</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">
      <div><label class="flabel">Column name</label><input type="text" id="nc-name" placeholder="e.g. Agent ID…"></div>
      <div><label class="flabel">Type</label>
        <select id="nc-type" onchange="onColTypeChange()">
          <option value="text">Text</option>
          <option value="select">Dropdown</option>
          <option value="status">Status</option>
          <option value="date">Date</option>
          <option value="number">Number</option>
        </select>
      </div>
    </div>
    <div id="nc-opts-wrap" style="display:none;margin-bottom:10px"><label class="flabel">Options (comma-separated)</label><input type="text" id="nc-opts" placeholder="Option A, Option B, Option C"></div>
    <button class="btn btn-primary" onclick="addCol()" style="width:100%;justify-content:center"><i class="ti ti-plus"></i> Add column</button>
  </div>
  <div class="mfoot"><button class="btn btn-primary" onclick="closeM('cols')">Done</button></div>
</div>
</div>

<!-- ADD ROW MODAL -->
<div id="m-add" class="overlay" style="display:none">
<div class="modal">
  <div class="mhead"><div class="mtitle"><i class="ti ti-plus"></i> Add Conversation</div><button class="ib" onclick="closeM('add')"><i class="ti ti-x"></i></button></div>
  <div class="mbody" id="add-body"></div>
  <div class="mfoot">
    <button class="btn" onclick="closeM('add')">Cancel</button>
    <button class="btn btn-primary" onclick="saveAdd()"><i class="ti ti-check"></i> Save</button>
  </div>
</div>
</div>

<!-- DETAIL MODAL -->
<div id="m-detail" class="overlay" style="display:none">
<div class="modal modal-lg">
  <div class="mhead"><div class="mtitle"><i class="ti ti-message-circle"></i> Conversation Detail</div><button class="ib" onclick="closeM('detail')"><i class="ti ti-x"></i></button></div>
  <div class="mbody" id="det-body"></div>
  <div class="mfoot">
    <button class="btn" onclick="closeM('detail')">Close</button>
    <button class="btn btn-primary" onclick="saveDet()"><i class="ti ti-device-floppy"></i> Save changes</button>
  </div>
</div>
</div>

<!-- n8n MODAL -->
<div id="m-n8n" class="overlay" style="display:none">
<div class="modal">
  <div class="mhead"><div class="mtitle"><i class="ti ti-topology-star"></i> n8n Configuration</div><button class="ib" onclick="closeM('n8n')"><i class="ti ti-x"></i></button></div>
  <div class="mbody">
    <div class="fg"><label class="flabel">Webhook base URL</label><input type="url" id="n8n-url" placeholder="https://your-n8n.example.com/webhook/"></div>
    <button class="btn btn-sm" onclick="testN8n()" style="margin-bottom:10px"><i class="ti ti-plug"></i> Test connection</button>
    <div id="n8n-test" style="font-size:12px;margin-bottom:12px"></div>
    <div style="height:1px;background:var(--border);margin-bottom:12px"></div>
    <div style="font-size:12px;font-weight:600;margin-bottom:10px">Triggers</div>
    <div id="trig-list"></div>
    <div style="height:1px;background:var(--border);margin:12px 0"></div>
    <div style="font-size:12px;font-weight:600;margin-bottom:6px">Event log</div>
    <div id="n8n-log" style="background:var(--bg2);border-radius:var(--radius);padding:8px;font-size:11px;font-family:monospace;max-height:100px;overflow-y:auto;color:var(--text2)">No events yet.</div>
  </div>
  <div class="mfoot">
    <button class="btn" onclick="closeM('n8n')">Cancel</button>
    <button class="btn btn-primary" onclick="saveN8n()"><i class="ti ti-device-floppy"></i> Save</button>
  </div>
</div>
</div>

<script>
let rows=[];
let cols=[
  {id:'id',label:'ID',type:'text',builtin:true},
  {id:'conv',label:'Conversation',type:'text',builtin:true},
  {id:'status',label:'Status',type:'status',builtin:true},
  {id:'sent',label:'Sentiment',type:'sent',builtin:true},
  {id:'rca',label:'Root Cause',type:'text',builtin:true},
  {id:'rcaCat',label:'RCA Category',type:'select',builtin:true,opts:['Agent Behavior','Wrong Information','Long Wait Time','Unresolved Issue','Process Failure','Technical Issue','Other']},
  {id:'notes',label:'Notes',type:'text',builtin:true},
  {id:'date',label:'Date',type:'date',builtin:true},
];
let sel=new Set();
let statusFilt='';
let sentFilt='';
let curDetIdx=-1;
let impData=[],impHeaders=[];
let n8nCfg={url:'',triggers:{import:true,status:true,unsatisfactory:true,export:false}};
let n8nLog=[];

function seed(){
  rows=[
    {id:'CONV-001',conv:'Customer called about billing discrepancy. Agent was unable to resolve and customer had to call back three times.',status:'Unsatisfactory',sent:'Negative',rca:'Repeat contact, unresolved billing',rcaCat:'Unresolved Issue',notes:'',date:'2025-01-10'},
    {id:'CONV-002',conv:'Agent helped customer reset password. Issue resolved on first contact. Customer satisfied.',status:'Satisfactory',sent:'Positive',rca:'',rcaCat:'',notes:'',date:'2025-01-10'},
    {id:'CONV-003',conv:'Customer waited 25 minutes on hold. Issue not resolved. Expressed frustration.',status:'Unsatisfactory',sent:'Negative',rca:'Excessive wait, no first-contact resolution',rcaCat:'Long Wait Time',notes:'Escalation needed',date:'2025-01-11'},
    {id:'CONV-004',conv:'Refund processed quickly and professionally. Customer happy with outcome.',status:'Satisfactory',sent:'Positive',rca:'',rcaCat:'',notes:'',date:'2025-01-11'},
    {id:'CONV-005',conv:'Agent gave incorrect product info. Customer had to call back after purchase.',status:'Unsatisfactory',sent:'Negative',rca:'Incorrect product information provided',rcaCat:'Wrong Information',notes:'Training gap',date:'2025-01-12'},
  ];
}

function getFiltered(){
  let r=rows;
  const q=(document.getElementById('q').value||'').toLowerCase();
  if(q) r=r.filter(row=>cols.some(c=>String(row[c.id]||'').toLowerCase().includes(q)));
  if(statusFilt) r=r.filter(row=>row.status===statusFilt);
  if(sentFilt) r=r.filter(row=>row.sent===sentFilt);
  return r;
}

function render(){
  const fr=getFiltered();
  const empty=document.getElementById('empty');
  const tbl=document.getElementById('tbl');
  updateStats(); updateSidebar();
  if(!rows.length){empty.style.display='block';tbl.style.display='none';return;}
  empty.style.display='none';tbl.style.display='';
  document.getElementById('thead').innerHTML='<tr><th style="width:28px"><input type="checkbox" id="chk-all" onchange="toggleAll(this.checked)"></th>'+cols.map(c=>`<th>${c.label}</th>`).join('')+'<th>Actions</th></tr>';
  document.getElementById('tbody').innerHTML=fr.map(row=>{
    const i=rows.indexOf(row);
    return `<tr class="${sel.has(i)?'sel':''}">
      <td><input type="checkbox" ${sel.has(i)?'checked':''} onchange="toggleSel(${i},this.checked)"></td>
      ${cols.map(c=>renderCell(row,c,i)).join('')}
      <td><div class="act-btns">
        <button class="ib" title="Edit" onclick="openDet(${i})"><i class="ti ti-pencil"></i></button>
        <button class="ib" title="Delete" onclick="delRow(${i})"><i class="ti ti-trash"></i></button>
      </div></td>
    </tr>`;
  }).join('');
  const allChecked=fr.length>0&&fr.every(r=>sel.has(rows.indexOf(r)));
  const cka=document.getElementById('chk-all');
  if(cka) cka.checked=allChecked;
  updateSelUI();
}

function renderCell(row,col,idx){
  const v=row[col.id]||'';
  if(col.type==='status'){
    return `<td><select class="inline-select" onchange="updateField(${idx},'status',this.value)">
      ${['Pending','Satisfactory','Unsatisfactory'].map(s=>`<option ${v===s?'selected':''}>${s}</option>`).join('')}
    </select></td>`;
  }
  if(col.type==='sent'){
    return `<td><select class="inline-select" onchange="updateField(${idx},'sent',this.value)">
      ${['','Positive','Negative','Neutral','Mixed'].map(s=>`<option value="${s}" ${v===s?'selected':''}>${s||'—'}</option>`).join('')}
    </select></td>`;
  }
  if(col.type==='select'&&col.opts){
    return `<td><select class="inline-select" onchange="updateField(${idx},'${col.id}',this.value)">
      <option value="">—</option>${col.opts.map(o=>`<option ${v===o?'selected':''}>${o}</option>`).join('')}
    </select></td>`;
  }
  if(col.id==='conv'){
    return `<td><span class="cell-txt" style="max-width:200px" title="${v.replace(/"/g,'&quot;')}">${v.slice(0,60)}${v.length>60?'…':''}</span></td>`;
  }
  return `<td><span class="cell-txt">${v}</span></td>`;
}

function updateField(idx,field,val){
  rows[idx][field]=val;
  if(n8nCfg.url&&n8nCfg.triggers.status&&field==='status') fireN8n('status_change',{id:rows[idx].id,field,newValue:val});
  if(n8nCfg.url&&n8nCfg.triggers.unsatisfactory&&field==='status'&&val==='Unsatisfactory') fireN8n('unsatisfactory',rows[idx]);
  updateStats(); updateSidebar();
}

function updateStats(){
  const t=rows.length;
  document.getElementById('st-total').textContent=t;
  document.getElementById('st-unsat').textContent=rows.filter(r=>r.status==='Unsatisfactory').length;
  document.getElementById('st-sat').textContent=rows.filter(r=>r.status==='Satisfactory').length;
  document.getElementById('st-pend').textContent=rows.filter(r=>!r.status||r.status==='Pending').length;
  document.getElementById('st-neg').textContent=rows.filter(r=>r.sent==='Negative').length;
  document.getElementById('st-pos').textContent=rows.filter(r=>r.sent==='Positive').length;
}

function updateSidebar(){
  document.getElementById('nb-all').textContent=rows.length;
  document.getElementById('nb-unsat').textContent=rows.filter(r=>r.status==='Unsatisfactory').length;
  document.getElementById('nb-sat').textContent=rows.filter(r=>r.status==='Satisfactory').length;
  document.getElementById('nb-pend').textContent=rows.filter(r=>!r.status||r.status==='Pending').length;
  document.getElementById('nb-pos').textContent=rows.filter(r=>r.sent==='Positive').length;
  document.getElementById('nb-neg').textContent=rows.filter(r=>r.sent==='Negative').length;
  document.getElementById('nb-neu').textContent=rows.filter(r=>r.sent==='Neutral').length;
  document.getElementById('nb-mix').textContent=rows.filter(r=>r.sent==='Mixed').length;
}

function toggleAll(checked){getFiltered().forEach(r=>{const i=rows.indexOf(r);checked?sel.add(i):sel.delete(i);});render();}
function toggleSel(i,checked){checked?sel.add(i):sel.delete(i);updateSelUI();}
function selAll(){rows.forEach((_,i)=>sel.add(i));render();}
function selNone(){sel.clear();render();}
function updateSelUI(){
  const n=sel.size;
  document.getElementById('sel-info').textContent=n?`${n} selected`:'';
  document.getElementById('del-btn').style.display=n?'':'none';
}
function delSel(){[...sel].sort((a,b)=>b-a).forEach(i=>rows.splice(i,1));sel.clear();render();}
function delRow(i){rows.splice(i,1);sel.delete(i);render();}
function setStatusFilter(v){statusFilt=v;sentFilt='';render();}
function setSentFilter(v){sentFilt=sentFilt===v?'':v;statusFilt='';render();}

function onDrop(e){e.preventDefault();document.getElementById('dz').classList.remove('drag-over');onFile(e.dataTransfer.files[0]);}
function onFile(f){
  if(!f) return;
  const ext=f.name.split('.').pop().toLowerCase();
  if(ext==='csv'){Papa.parse(f,{header:true,complete:r=>processImp(r.data,r.meta.fields)});}
  else{
    const reader=new FileReader();
    reader.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const data=XLSX.utils.sheet_to_json(ws,{defval:''});
      processImp(data,Object.keys(data[0]||{}));
    };
    reader.readAsBinaryString(f);
  }
}
function processImp(data,headers){
  impData=data;impHeaders=headers;
  document.getElementById('prev-tbl').innerHTML=`<table style="font-size:11px;width:100%;border-collapse:collapse">
    <tr>${headers.map(h=>`<th style="padding:3px 6px;text-align:left;border:1px solid var(--border);background:var(--bg2)">${h}</th>`).join('')}</tr>
    ${data.slice(0,5).map(r=>`<tr>${headers.map(h=>`<td style="padding:2px 6px;border:1px solid var(--border)">${String(r[h]||'').slice(0,25)}</td>`).join('')}</tr>`).join('')}
  </table>`;
  document.getElementById('imp-preview').style.display='block';
  document.getElementById('imp-info').textContent=`${data.length} rows · ${headers.length} columns`;
  ['map-text','map-status','map-sent','map-rca'].forEach(id=>{
    document.getElementById(id).innerHTML=(id==='map-text'?'<option value="">— select —</option>':'<option value="">— none —</option>')+headers.map(h=>`<option value="${h}">${h}</option>`).join('');
  });
  const g=(patterns,selId)=>{const h=headers.find(h=>patterns.some(p=>new RegExp(p,'i').test(h)));if(h)document.getElementById(selId).value=h;};
  g(['conv','text','message','transcript','comment','body'],'map-text');
  g(['status','result','outcome'],'map-status');
  g(['sentiment','sent','feeling','emotion'],'map-sent');
  g(['root.cause','rca','reason','cause'],'map-rca');
  document.getElementById('imp-btn').disabled=false;
}
function doImport(){
  const tCol=document.getElementById('map-text').value;
  const sCol=document.getElementById('map-status').value;
  const stCol=document.getElementById('map-sent').value;
  const rCol=document.getElementById('map-rca').value;
  impHeaders.forEach(h=>{if(!cols.find(c=>c.id===h))cols.push({id:h,label:h,type:'text'});});
  const newRows=impData.map((r,i)=>{
    const row={id:`IMP-${String(rows.length+i+1).padStart(3,'0')}`,conv:'',status:'Pending',sent:'',rca:'',rcaCat:'',notes:'',date:new Date().toISOString().slice(0,10)};
    impHeaders.forEach(h=>{row[h]=r[h];});
    if(tCol) row.conv=r[tCol]||'';
    if(sCol) row.status=r[sCol]||'Pending';
    if(stCol) row.sent=r[stCol]||'';
    if(rCol) row.rca=r[rCol]||'';
    return row;
  });
  rows=[...rows,...newRows];
  if(n8nCfg.url&&n8nCfg.triggers.import) fireN8n('import',{count:newRows.length,timestamp:new Date().toISOString()});
  closeM('import');render();renderColChips();
}

function onColTypeChange(){document.getElementById('nc-opts-wrap').style.display=document.getElementById('nc-type').value==='select'?'block':'none';}
function addCol(){
  const name=document.getElementById('nc-name').value.trim();
  const type=document.getElementById('nc-type').value;
  if(!name) return;
  const col={id:'c_'+Date.now(),label:name,type};
  if(type==='select') col.opts=document.getElementById('nc-opts').value.split(',').map(s=>s.trim()).filter(Boolean);
  cols.push(col);
  document.getElementById('nc-name').value='';
  document.getElementById('nc-opts').value='';
  renderColChips();render();
}
function removeCol(id){cols=cols.filter(c=>c.id!==id);renderColChips();render();}
function renderColChips(){
  document.getElementById('col-chips').innerHTML=cols.map(c=>`<span class="col-chip">${c.label}${!c.builtin?`<button class="rm" onclick="removeCol('${c.id}')">×</button>`:''}</span>`).join('');
}

function renderRowForm(row){
  return cols.filter(c=>c.id!=='id').map(c=>{
    const v=row[c.id]||'';
    if(c.type==='status') return `<div class="fg"><label class="flabel">${c.label}</label><select id="af-${c.id}">${['Pending','Satisfactory','Unsatisfactory'].map(s=>`<option ${v===s?'selected':''}>${s}</option>`).join('')}</select></div>`;
    if(c.type==='sent') return `<div class="fg"><label class="flabel">${c.label}</label><select id="af-${c.id}">${['','Positive','Negative','Neutral','Mixed'].map(s=>`<option value="${s}" ${v===s?'selected':''}>${s||'—'}</option>`).join('')}</select></div>`;
    if(c.type==='select'&&c.opts) return `<div class="fg"><label class="flabel">${c.label}</label><select id="af-${c.id}"><option value="">—</option>${c.opts.map(o=>`<option ${v===o?'selected':''}>${o}</option>`).join('')}</select></div>`;
    if(c.id==='conv'||c.id==='notes') return `<div class="fg"><label class="flabel">${c.label}</label><textarea id="af-${c.id}" rows="${c.id==='conv'?4:2}">${v}</textarea></div>`;
    return `<div class="fg"><label class="flabel">${c.label}</label><input type="text" id="af-${c.id}" value="${v}"></div>`;
  }).join('');
}
function saveAdd(){
  const row={id:'ROW-'+String(rows.length+1).padStart(3,'0')};
  cols.filter(c=>c.id!=='id').forEach(c=>{const el=document.getElementById('af-'+c.id);row[c.id]=el?el.value:'';});
  if(!row.date) row.date=new Date().toISOString().slice(0,10);
  rows.push(row);closeM('add');render();
}

function openDet(i){
  curDetIdx=i;
  document.getElementById('det-body').innerHTML=
    `<div style="background:var(--bg2);border-radius:var(--radius);padding:12px;margin-bottom:14px;font-size:13px;line-height:1.7">${rows[i].conv||'No text'}</div>`+
    `<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">`+
    cols.filter(c=>!['id','conv'].includes(c.id)).map(c=>{
      const v=rows[i][c.id]||'';
      if(c.type==='status') return `<div class="fg"><label class="flabel">${c.label}</label><select id="df-${c.id}">${['Pending','Satisfactory','Unsatisfactory'].map(s=>`<option ${v===s?'selected':''}>${s}</option>`).join('')}</select></div>`;
      if(c.type==='sent') return `<div class="fg"><label class="flabel">${c.label}</label><select id="df-${c.id}">${['','Positive','Negative','Neutral','Mixed'].map(s=>`<option value="${s}" ${v===s?'selected':''}>${s||'—'}</option>`).join('')}</select></div>`;
      if(c.type==='select'&&c.opts) return `<div class="fg"><label class="flabel">${c.label}</label><select id="df-${c.id}"><option value="">—</option>${c.opts.map(o=>`<option ${v===o?'selected':''}>${o}</option>`).join('')}</select></div>`;
      if(c.id==='notes') return `<div class="fg" style="grid-column:span 2"><label class="flabel">${c.label}</label><textarea id="df-${c.id}">${v}</textarea></div>`;
      if(c.id==='rca') return `<div class="fg" style="grid-column:span 2"><label class="flabel">${c.label}</label><input type="text" id="df-${c.id}" value="${v}"></div>`;
      return `<div class="fg"><label class="flabel">${c.label}</label><input type="text" id="df-${c.id}" value="${v}"></div>`;
    }).join('')+`</div>`;
  openM('detail');
}
function saveDet(){
  cols.filter(c=>!['id','conv'].includes(c.id)).forEach(c=>{const el=document.getElementById('df-'+c.id);if(el) rows[curDetIdx][c.id]=el.value;});
  if(n8nCfg.url&&n8nCfg.triggers.status) fireN8n('status_change',rows[curDetIdx]);
  if(n8nCfg.url&&n8nCfg.triggers.unsatisfactory&&rows[curDetIdx].status==='Unsatisfactory') fireN8n('unsatisfactory',rows[curDetIdx]);
  closeM('detail');render();
}

function renderRCA(){
  const el=document.getElementById('rca-content');
  const t=rows.length;
  if(!t){el.innerHTML='<div class="empty"><i class="ti ti-chart-pie"></i><div>Import data to see RCA insights</div></div>';return;}
  const u=rows.filter(r=>r.status==='Unsatisfactory');
  const uPct=Math.round(u.length/t*100);
  const cats={};
  u.forEach(r=>{const c=r.rcaCat||'Uncategorized';cats[c]=(cats[c]||0)+1;});
  const sents={Positive:0,Negative:0,Neutral:0,Mixed:0};
  const sentAnalyzed=rows.filter(r=>r.sent).length;
  rows.forEach(r=>{if(r.sent) sents[r.sent]=(sents[r.sent]||0)+1;});
  el.innerHTML=`
    <div class="rca-grid">
      <div class="rca-card" style="grid-column:span 2;display:flex;gap:28px;align-items:center">
        <div><div style="font-size:32px;font-weight:700;color:#A32D2D">${uPct}%</div><div style="font-size:12px;color:var(--text3)">Unsatisfactory rate</div></div>
        <div><div style="font-size:32px;font-weight:700">${t}</div><div style="font-size:12px;color:var(--text3)">Total conversations</div></div>
        <div><div style="font-size:32px;font-weight:700;color:#A32D2D">${u.length}</div><div style="font-size:12px;color:var(--text3)">Unsatisfactory</div></div>
        <div><div style="font-size:32px;font-weight:700;color:#3B6D11">${rows.filter(r=>r.status==='Satisfactory').length}</div><div style="font-size:12px;color:var(--text3)">Satisfactory</div></div>
      </div>
      <div class="rca-card">
        <div style="font-size:12px;font-weight:600;margin-bottom:10px">Root Cause Breakdown</div>
        ${Object.entries(cats).sort((a,b)=>b[1]-a[1]).map(([c,n])=>`
          <div style="margin-bottom:7px">
            <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px"><span>${c}</span><span style="font-weight:600">${n} (${u.length?Math.round(n/u.length*100):0}%)</span></div>
            <div class="prog"><div class="prog-fill" style="width:${u.length?Math.round(n/u.length*100):0}%;background:#E24B4A"></div></div>
          </div>`).join('')||'<div style="font-size:12px;color:var(--text3)">No RCA categories assigned yet</div>'}
      </div>
      <div class="rca-card">
        <div style="font-size:12px;font-weight:600;margin-bottom:10px">Sentiment Distribution</div>
        ${sentAnalyzed?Object.entries(sents).filter(([,n])=>n).map(([s,n])=>`
          <div style="margin-bottom:7px">
            <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px"><span>${s}</span><span style="font-weight:600">${n} (${Math.round(n/sentAnalyzed*100)}%)</span></div>
            <div class="prog"><div class="prog-fill" style="width:${Math.round(n/sentAnalyzed*100)}%;background:${s==='Positive'?'#639922':s==='Negative'?'#E24B4A':s==='Mixed'?'#EF9F27':'#888'}"></div></div>
          </div>`).join('')
        :'<div style="font-size:12px;color:var(--text3)">Set sentiment on rows to see distribution</div>'}
      </div>
    </div>
    <div class="rca-card">
      <div style="font-size:12px;font-weight:600;margin-bottom:10px">Unsatisfactory Conversations</div>
      ${u.length?`<table style="width:100%;border-collapse:collapse;font-size:12px">
        <tr style="background:var(--bg2)">
          <th style="padding:6px 8px;text-align:left;font-weight:600;border-bottom:1px solid var(--border)">ID</th>
          <th style="padding:6px 8px;text-align:left;font-weight:600;border-bottom:1px solid var(--border)">Sentiment</th>
          <th style="padding:6px 8px;text-align:left;font-weight:600;border-bottom:1px solid var(--border)">RCA Category</th>
          <th style="padding:6px 8px;text-align:left;font-weight:600;border-bottom:1px solid var(--border)">Root Cause</th>
        </tr>
        ${u.map(r=>`<tr>
          <td style="padding:5px 8px;border-bottom:1px solid var(--border)">${r.id}</td>
          <td style="padding:5px 8px;border-bottom:1px solid var(--border)"><span class="pill ${r.sent==='Negative'?'p-neg':r.sent==='Positive'?'p-pos':'p-neu'}">${r.sent||'—'}</span></td>
          <td style="padding:5px 8px;border-bottom:1px solid var(--border)">${r.rcaCat||'—'}</td>
          <td style="padding:5px 8px;border-bottom:1px solid var(--border)">${(r.rca||'—').slice(0,60)}</td>
        </tr>`).join('')}
      </table>`:'<div style="font-size:12px;color:var(--text3)">No unsatisfactory conversations</div>'}
    </div>`;
}

const trigDefs=[
  {key:'import',icon:'ti-upload',color:'#639922',label:'On data import',desc:'Fires when rows are imported from a file'},
  {key:'status',icon:'ti-flag',color:'#EF9F27',label:'On status change',desc:'Fires when a row status is updated'},
  {key:'unsatisfactory',icon:'ti-alert-triangle',color:'#E24B4A',label:'On unsatisfactory',desc:'Fires only when status = Unsatisfactory'},
  {key:'export',icon:'ti-download',color:'#185FA5',label:'On export',desc:'Fires when data is exported to CSV'},
];
function renderTrigList(){
  document.getElementById('trig-list').innerHTML=trigDefs.map(t=>`
    <div class="trig-opt ${n8nCfg.triggers[t.key]?'on':''}" id="trig-${t.key}" onclick="toggleTrig('${t.key}')">
      <i class="ti ${t.icon}" style="font-size:18px;color:${t.color};flex-shrink:0"></i>
      <div style="flex:1"><div style="font-weight:600;font-size:13px">${t.label}</div><div style="font-size:11px;color:var(--text3)">${t.desc}</div></div>
      <i class="ti ${n8nCfg.triggers[t.key]?'ti-circle-check':'ti-circle'}" style="font-size:17px;color:${n8nCfg.triggers[t.key]?'#639922':'var(--text3)'}"></i>
    </div>`).join('');
}
function toggleTrig(key){n8nCfg.triggers[key]=!n8nCfg.triggers[key];renderTrigList();}

function renderWFContent(){
  const on=!!n8nCfg.url;
  document.getElementById('wf-content').innerHTML=`
    <div class="wf-card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
        <div style="display:flex;align-items:center;gap:8px;font-weight:600"><i class="ti ti-topology-star" style="color:var(--green)"></i> n8n Integration</div>
        <span class="pill ${on?'p-pos':'p-neu'}">${on?'Active':'Inactive'}</span>
      </div>
      <div style="font-size:12px;color:var(--text2);margin-bottom:10px">${on?'Connected: '+n8nCfg.url:'Not configured — click Configure above.'}</div>
      <div style="display:flex;gap:6px;flex-wrap:wrap">${trigDefs.map(t=>`<span class="pill ${n8nCfg.triggers[t.key]?'p-pos':'p-neu'}">${t.label}</span>`).join('')}</div>
    </div>
    <div class="wf-card" style="border-style:dashed">
      <div style="font-size:12px;font-weight:600;margin-bottom:8px;color:var(--text2)">Example webhook payload</div>
      <pre style="font-size:11px;font-family:monospace;color:var(--text2)">${JSON.stringify({event:"unsatisfactory",payload:{id:"CONV-003",status:"Unsatisfactory",sent:"Negative",rca:"Excessive wait time",rcaCat:"Long Wait Time"},source:"ConvoValidate",timestamp:new Date().toISOString()},null,2)}</pre>
    </div>
    <div class="wf-card">
      <div style="font-size:12px;font-weight:600;margin-bottom:8px">Webhook endpoints</div>
      ${trigDefs.map(t=>`<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:12px">
        <i class="ti ${t.icon}" style="font-size:14px;color:${t.color}"></i>
        <code style="font-size:11px;color:var(--text2)">${n8nCfg.url||'https://your-n8n/webhook/'}${t.key}</code>
      </div>`).join('')}
    </div>`;
}

function testN8n(){
  const url=document.getElementById('n8n-url').value;
  const el=document.getElementById('n8n-test');
  if(!url){el.innerHTML='<span style="color:#A32D2D">Enter a URL first.</span>';return;}
  el.innerHTML='<span style="color:#185FA5">Testing…</span>';
  fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({test:true,source:'ConvoValidate'})})
    .then(r=>{el.innerHTML=r.ok?'<span style="color:#639922">✓ Connected!</span>':`<span style="color:#EF9F27">Got ${r.status}</span>`;})
    .catch(()=>{el.innerHTML='<span style="color:#EF9F27">Could not reach URL — check your n8n webhook settings</span>';});
}
function saveN8n(){
  n8nCfg.url=document.getElementById('n8n-url').value.trim();
  const on=!!n8nCfg.url;
  document.getElementById('top-dot').className='n8n-dot '+(on?'dot-on':'dot-off');
  document.getElementById('top-n8n-lbl').textContent=on?'n8n on':'n8n off';
  closeM('n8n');renderWFContent();
}
function fireN8n(event,payload){
  if(!n8nCfg.url) return;
  const ts=new Date().toLocaleTimeString();
  n8nLog.unshift(`[${ts}] ${event}: ${JSON.stringify(payload).slice(0,70)}…`);
  const logEl=document.getElementById('n8n-log');
  if(logEl) logEl.textContent=n8nLog.slice(0,15).join('\n');
  fetch(n8nCfg.url+event,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({event,payload,source:'ConvoValidate',timestamp:new Date().toISOString()})}).catch(()=>{});
}

function exportCSV(){
  const hdr=cols.map(c=>c.label).join(',');
  const body=rows.map(r=>cols.map(c=>JSON.stringify(r[c.id]||'')).join(',')).join('\n');
  const blob=new Blob([hdr+'\n'+body],{type:'text/csv'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='rca_export.csv';a.click();
  if(n8nCfg.url&&n8nCfg.triggers.export) fireN8n('export',{count:rows.length,timestamp:new Date().toISOString()});
}

function switchView(v,el){
  ['table','rca','wf'].forEach(id=>{
    document.getElementById('view-'+id).style.display='none';
    document.getElementById('nav-'+id).classList.remove('active');
  });
  document.getElementById('view-'+v).style.display=v==='table'?'flex':'block';
  (el||document.getElementById('nav-'+v)).classList.add('active');
  if(v==='rca') renderRCA();
  if(v==='wf') renderWFContent();
}

function openM(id){
  if(id==='add'){document.getElementById('add-body').innerHTML=renderRowForm({});}
  if(id==='cols') renderColChips();
  if(id==='n8n'){document.getElementById('n8n-url').value=n8nCfg.url;renderTrigList();}
  document.getElementById('m-'+id).style.display='flex';
}
function closeM(id){document.getElementById('m-'+id).style.display='none';}
window.addEventListener('click',e=>{if(e.target.classList.contains('overlay'))closeM(e.target.id.replace('m-',''));});

seed();render();renderColChips();
</script>
</body>
</html>
