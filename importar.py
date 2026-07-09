"""
SICAMEZ — Módulo de Importación Universal de Pólizas
Servidor local en http://localhost:5001
Acepta: PDF, imagen (JPG/PNG), Excel (.xlsx), CSV
Extrae datos con Google Vision OCR o pandas
Guarda en C:\SICAMEZ\polizas.csv y clientes_cumpleanos.csv

Agente: Francisco Miguel Amezcua Izquierdo
"""

import os, json, csv, base64, re, webbrowser, threading
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from flask import Flask, request, jsonify, render_template_string

# Cargar .env local si existe (en la nube se usan variables de entorno)
_env_local = Path(r"C:\SICAMEZ\.env")
if _env_local.exists():
    load_dotenv(str(_env_local))
else:
    load_dotenv()

GOOGLE_KEY   = os.getenv("GOOGLE_VISION_API_KEY", "")
TWILIO_SID   = os.getenv("TWILIO_ACCOUNT_SID", "")
TWILIO_TOKEN = os.getenv("TWILIO_AUTH_TOKEN", "")
FROM_WA      = "whatsapp:+14155238886"

# Directorio base: C:\SICAMEZ en local, /data/sicamez en la nube (Railway con volumen)
_base_local = Path(r"C:\SICAMEZ")
_base_cloud = Path(os.getenv("SICAMEZ_DATA_DIR", "/data/sicamez"))
BASE_DIR = _base_local if _base_local.exists() else _base_cloud
BASE_DIR.mkdir(parents=True, exist_ok=True)

POLIZAS_CSV  = BASE_DIR / "polizas.csv"
CUMPLE_CSV   = BASE_DIR / "clientes_cumpleanos.csv"
UPLOAD_DIR   = BASE_DIR / "importaciones"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB máximo

# ── HTML del panel de importación ─────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SICAMEZ — Importar pólizas</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f0;color:#1a1a1a;min-height:100vh}
.header{background:#1a1a2e;color:white;padding:14px 24px;display:flex;align-items:center;gap:12px}
.header h1{font-size:17px;font-weight:500}
.header span{font-size:12px;opacity:.6;margin-left:auto}
.main{max-width:900px;margin:0 auto;padding:24px 16px}
.card{background:white;border-radius:12px;border:1px solid #e8e8e0;padding:20px;margin-bottom:16px}
.sec{font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin-bottom:12px}
.zone{border:2px dashed #d0d0c8;border-radius:10px;padding:36px 20px;text-align:center;cursor:pointer;transition:all .2s}
.zone:hover,.zone.over{border-color:#4a90d9;background:#f0f7ff}
.zone-icon{font-size:40px;margin-bottom:10px}
.zone h2{font-size:16px;font-weight:500;margin-bottom:6px}
.zone p{font-size:13px;color:#666;margin-bottom:14px}
.pills{display:flex;flex-wrap:wrap;gap:6px;justify-content:center;margin-bottom:16px}
.pill{font-size:11px;padding:3px 10px;border-radius:999px;background:#f0f0ea;border:1px solid #ddd;color:#555}
.btn{padding:9px 20px;border-radius:8px;border:1px solid #ccc;background:white;font-size:13px;cursor:pointer;transition:all .15s}
.btn:hover{background:#f5f5f0}
.btn-primary{background:#1a1a2e;color:white;border-color:#1a1a2e}
.btn-primary:hover{background:#2a2a4e}
.btn-success{background:#2e7d32;color:white;border-color:#2e7d32}
.progress-wrap{display:none;margin-top:16px}
.progress-bar{height:6px;background:#e8e8e0;border-radius:3px;overflow:hidden;margin-bottom:8px}
.progress-fill{height:100%;background:#4a90d9;border-radius:3px;transition:width .4s;width:0%}
.log{font-size:12px;color:#666;line-height:1.8}
.results{display:none}
.res-header{display:flex;align-items:center;gap:10px;margin-bottom:14px}
.badge{font-size:11px;padding:3px 10px;border-radius:999px;font-weight:600}
.badge-ok{background:#e8f5e9;color:#2e7d32}
.badge-warn{background:#fff3e0;color:#e65100}
.badge-ramo{background:#e3f2fd;color:#1565c0}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.04em;padding:6px 8px;border-bottom:1px solid #e8e8e0}
td{padding:7px 8px;border-bottom:1px solid #f0f0ea;vertical-align:middle}
td input{border:none;background:transparent;font-size:13px;width:100%;color:#1a1a1a;font-weight:500;padding:0}
td input:focus{outline:none;border-bottom:1px solid #4a90d9}
.conf-ok{font-size:10px;padding:2px 7px;border-radius:999px;background:#e8f5e9;color:#2e7d32;font-weight:600;white-space:nowrap}
.conf-warn{font-size:10px;padding:2px 7px;border-radius:999px;background:#fff3e0;color:#e65100;font-weight:600;white-space:nowrap}
.actions{display:flex;gap:8px;justify-content:flex-end;margin-top:16px;flex-wrap:wrap}
.batch-row{display:grid;grid-template-columns:2fr 1fr 1fr 1fr 80px;gap:8px;padding:7px 8px;border-bottom:1px solid #f0f0ea;font-size:13px;align-items:center}
.batch-head{font-size:11px;font-weight:600;color:#888;text-transform:uppercase}
.dot{width:8px;height:8px;border-radius:50%}
.dot-ok{background:#4caf50}.dot-warn{background:#ff9800}
.toast{position:fixed;bottom:20px;right:20px;background:#1a1a2e;color:white;padding:12px 20px;border-radius:10px;font-size:13px;display:none;z-index:999}
</style>
</head>
<body>
<div class="header">
  <div>📂</div>
  <h1>SICAMEZ — Importación de pólizas</h1>
  <span>Servidor local · todo en tu equipo</span>
</div>
<div class="main">

<div class="card">
  <p class="sec">Sube el archivo</p>
  <div class="zone" id="zone" onclick="document.getElementById('file-input').click()"
       ondragover="ev(event,'over')" ondragleave="ev(event,'')" ondrop="drop(event)">
    <div class="zone-icon">📄</div>
    <h2>Arrastra aquí o haz clic para seleccionar</h2>
    <p>Una póliza suelta <strong>o</strong> una relación con todas las pólizas</p>
    <div class="pills">
      <span class="pill">PDF</span><span class="pill">JPG / PNG</span>
      <span class="pill">Excel .xlsx</span><span class="pill">CSV</span>
      <span class="pill">Cualquier compañía</span><span class="pill">Cualquier ramo</span>
    </div>
    <button class="btn">Seleccionar archivo</button>
  </div>
  <input type="file" id="file-input" style="display:none" accept=".pdf,.jpg,.jpeg,.png,.xlsx,.xls,.csv"
         onchange="upload(this.files[0])">

  <div class="progress-wrap" id="progress-wrap">
    <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
    <div class="log" id="log"></div>
  </div>
</div>

<div class="card results" id="results-single">
  <div class="res-header">
    <span>✅</span>
    <strong id="res-titulo" style="font-size:15px"></strong>
    <span class="badge badge-ramo" id="res-ramo"></span>
    <span class="badge badge-ok" id="res-confianza"></span>
  </div>
  <table>
    <thead><tr><th>Campo</th><th>Valor extraído</th><th></th></tr></thead>
    <tbody id="tabla-campos"></tbody>
  </table>
  <div class="actions">
    <button class="btn" onclick="resetUI()">Subir otro</button>
    <button class="btn btn-success" id="btn-guardar-single" onclick="guardar('single')">Guardar en SICAMEZ</button>
    <button class="btn" id="btn-wa-single" style="display:none;background:#25D366;color:white;border-color:#25D366" onclick="enviarWA('single')">📱 Enviar bienvenida WhatsApp</button>
  </div>
</div>

<div class="card results" id="results-batch">
  <div class="res-header">
    <span>✅</span>
    <strong id="batch-titulo" style="font-size:15px"></strong>
  </div>
  <div style="font-size:12px;color:#666;margin-bottom:10px;" id="hojas-info"></div>
  <div class="batch-row batch-head">
    <span>Nombre</span><span>Compañía · Ramo</span><span>Póliza</span><span>Prima</span><span>Hoja · Estado</span>
  </div>
  <div id="batch-lista"></div>
  <div class="actions">
    <button class="btn" onclick="resetUI()">Subir otro</button>
    <button class="btn btn-success" id="btn-guardar-batch" onclick="guardar('batch')">Importar todo a SICAMEZ</button>
    <button class="btn" id="btn-wa-batch" style="display:none;background:#25D366;color:white;border-color:#25D366" onclick="enviarWA('batch')">📱 Enviar bienvenidas WhatsApp</button>
  </div>
</div>

</div>
<div class="toast" id="toast"></div>

<script>
let resultData = null;

function ev(e, cls) {
  e.preventDefault();
  document.getElementById('zone').className = 'zone ' + cls;
}

function drop(e) {
  e.preventDefault();
  document.getElementById('zone').className = 'zone';
  const f = e.dataTransfer.files[0];
  if (f) upload(f);
}

function setProgress(pct, msg) {
  document.getElementById('progress-fill').style.width = pct + '%';
  const log = document.getElementById('log');
  const line = document.createElement('div');
  line.textContent = '✓ ' + msg;
  log.appendChild(line);
}

async function upload(file) {
  if (!file) return;
  document.getElementById('progress-wrap').style.display = 'block';
  document.getElementById('results-single').style.display = 'none';
  document.getElementById('results-batch').style.display = 'none';
  document.getElementById('log').innerHTML = '';

  const isExcel = file.name.match(/\.(xlsx|xls|csv)$/i);
  setProgress(10, 'Leyendo ' + file.name + '...');

  const form = new FormData();
  form.append('file', file);

  setTimeout(() => setProgress(30, isExcel ? 'Procesando hoja de cálculo...' : 'Enviando a OCR Google Vision...'), 400);
  setTimeout(() => setProgress(55, 'Extrayendo campos...'), 900);
  setTimeout(() => setProgress(75, 'Identificando compañía y ramo...'), 1400);

  try {
    const res = await fetch('/upload', { method: 'POST', body: form });
    const data = await res.json();
    setProgress(100, 'Listo — ' + (data.tipo === 'batch' ? data.polizas.length + ' pólizas' : data.campos.length + ' campos') + ' extraídos');
    resultData = data;
    setTimeout(() => mostrar(data), 400);
  } catch(e) {
    setProgress(100, 'Error: ' + e.message);
  }
}

function mostrar(data) {
  if (data.tipo === 'batch') {
    document.getElementById('batch-titulo').textContent =
      data.total + ' pólizas en ' + (data.hojas ? data.hojas.length : 1) + ' hoja(s) — ' + data.archivo;

    if (data.hojas && data.hojas.length > 1) {
      const resumen = data.hojas.map(h => `<strong>${h.hoja}</strong>: ${h.filas} pólizas`).join('  ·  ');
      document.getElementById('hojas-info').innerHTML = resumen;
    }

    const lista = document.getElementById('batch-lista');
    lista.innerHTML = '';
    data.polizas.forEach(p => {
      const row = document.createElement('div');
      row.className = 'batch-row';
      const ok = p.ok !== false;
      row.innerHTML = `<span><strong>${p.nombre||'—'}</strong></span>
        <span style="color:#555">${p.compania||'—'} ${p.ramo ? '· '+p.ramo : ''}</span>
        <span style="color:#888;font-size:12px">${p.poliza||'—'}</span>
        <span style="color:#555">${p.prima||'—'}</span>
        <span style="font-size:12px">${p.hoja ? p.hoja+' · ' : ''}<span class="dot ${ok?'dot-ok':'dot-warn'}" style="display:inline-block;vertical-align:middle"></span> ${ok?'OK':'Revisar'}</span>`;
      lista.appendChild(row);
    });
    document.getElementById('results-batch').style.display = 'block';
  } else {
    document.getElementById('res-titulo').textContent = data.compania + ' · ' + data.numero_poliza;
    document.getElementById('res-ramo').textContent = data.ramo;
    document.getElementById('res-confianza').textContent = 'Confianza: ' + data.confianza + '%';
    const tbody = document.getElementById('tabla-campos');
    tbody.innerHTML = '';
    data.campos.forEach(f => {
      const tr = document.createElement('tr');
      const ok = f.c === 'high';
      tr.innerHTML = `<td style="color:#666;white-space:nowrap">${f.l}</td>
        <td><input value="${f.v||''}" data-key="${f.k||f.l}"></td>
        <td><span class="${ok?'conf-ok':'conf-warn'}">${ok?'Correcto':'Revisar'}</span></td>`;
      tbody.appendChild(tr);
    });
    document.getElementById('results-single').style.display = 'block';
  }
}

async function guardar(tipo) {
  const res = await fetch('/guardar', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ tipo, data: resultData })
  });
  const r = await res.json();
  showToast(r.mensaje || 'Guardado correctamente');
  // Mostrar botón WhatsApp si hay teléfono disponible
  if (r.wa_disponible) {
    if (tipo === 'single') {
      document.getElementById('btn-guardar-single').style.display = 'none';
      document.getElementById('btn-wa-single').style.display = '';
    } else {
      document.getElementById('btn-guardar-batch').style.display = 'none';
      document.getElementById('btn-wa-batch').style.display = '';
    }
  }
}

async function enviarWA(tipo) {
  const btn = document.getElementById(tipo === 'single' ? 'btn-wa-single' : 'btn-wa-batch');
  btn.textContent = '⏳ Enviando...';
  btn.disabled = true;
  try {
    const res = await fetch('/whatsapp', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ tipo, data: resultData })
    });
    const r = await res.json();
    showToast(r.mensaje || '✅ Mensajes enviados');
    btn.textContent = '✅ Enviado';
  } catch(e) {
    showToast('❌ Error al enviar: ' + e.message);
    btn.textContent = '❌ Error';
  }
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.style.display = 'block';
  setTimeout(() => t.style.display = 'none', 3500);
}

function resetUI() {
  document.getElementById('progress-wrap').style.display = 'none';
  document.getElementById('results-single').style.display = 'none';
  document.getElementById('results-batch').style.display = 'none';
  document.getElementById('log').innerHTML = '';
  document.getElementById('file-input').value = '';
}
</script>
</body>
</html>"""

# ── Procesadores ──────────────────────────────────────────────────────────────

def ocr_con_vision(ruta):
    """Usa Google Cloud Vision para extraer texto de imagen o PDF."""
    import requests as req

    # Limitar tamaño: no procesar archivos > 8MB con OCR
    tam = os.path.getsize(ruta)
    if tam > 8 * 1024 * 1024:
        return f"[Archivo grande ({tam//1024//1024}MB). Usa Excel/CSV para importación masiva.]"

    try:
        with open(ruta, "rb") as f:
            contenido = base64.b64encode(f.read()).decode()

        ext = Path(ruta).suffix.lower()
        payload = {
            "requests": [{
                "image": {"content": contenido},
                "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
                "imageContext": {"languageHints": ["es"]}
            }]
        }
        url = f"https://vision.googleapis.com/v1/images:annotate?key={GOOGLE_KEY}"
        resp = req.post(url, json=payload, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        return data["responses"][0]["fullTextAnnotation"]["text"]
    except Exception as e:
        return f"[OCR no disponible: {e}]"


def parsear_texto(texto):
    """Extrae campos comunes de texto OCR de cualquier póliza."""
    def buscar(patrones, texto):
        for p in patrones:
            m = re.search(p, texto, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return ""

    poliza  = buscar([r'[Pp]óliza[:\s#No.]*([A-Z0-9\-]+)', r'No\.?\s*([A-Z0-9\-]{5,20})'], texto)
    nombre  = buscar([r'[Tt]itular[:\s]+([A-ZÁÉÍÓÚÑ ]{5,60})', r'[Aa]segurado[:\s]+([A-ZÁÉÍÓÚÑ ]{5,60})', r'[Cc]ontratante[:\s]+([A-ZÁÉÍÓÚÑ ]{5,60})'], texto)
    fecnac  = buscar([r'[Ff]echa?\s+[Nn]ac[a-z.]*[:\s]+(\d{2}[/\-]\d{2}[/\-]\d{4})', r'F\.\s*Nac[.:\s]+(\d{2}[/\-]\d{2}[/\-]\d{4})'], texto)
    rfc     = buscar([r'RFC[:\s]+([A-Z]{3,4}\d{6}[A-Z0-9]{3})'], texto)
    vigini  = buscar([r'[Vv]igencia[:\s]+(\d{2}[/\-]\d{2}[/\-]\d{4})'], texto)
    vigfin  = buscar([r'al\s+(\d{2}[/\-]\d{2}[/\-]\d{4})'], texto)
    prima   = buscar([r'[Pp]rima\s+[Nn]eta[:\s]+\$?([\d,]+\.?\d*)'], texto)
    sa      = buscar([r'[Ss]uma\s+[Aa]segurada[:\s]+\$?([\d,]+\.?\d*)'], texto)
    tel     = buscar([r'[Tt]el[éeéfono]*[:\s]+(\d[\d\s\-]{8,12}\d)'], texto)

    compania = ""
    for c in ["Atlas","GNP","AXA","Qualitas","Mapfre","Chubb","Metlife","Allianz","HDI","Sura","BBVA","Insignia"]:
        if c.lower() in texto.lower():
            compania = c; break

    ramo = ""
    mapa = {"GMM":"GMM","[Gg]astos [Mm]édicos":"GMM","[Aa]utomóvil":"Auto","[Aa]uto":"Auto",
            "[Vv]ida":"Vida","[Hh]ogar":"Hogar","[Ee]mpresarial":"Empresarial",
            "[Tt]ransporte":"Transporte","[Vv]iajes":"Viajes"}
    for pat, val in mapa.items():
        if re.search(pat, texto):
            ramo = val; break

    campos = [
        {"k":"numero_poliza",    "l":"Número de póliza",   "v":poliza, "c":"high" if poliza else "low"},
        {"k":"nombre",           "l":"Nombre del titular", "v":nombre, "c":"high" if nombre else "low"},
        {"k":"fecha_nacimiento", "l":"Fecha de nacimiento","v":fecnac, "c":"high" if fecnac else "low"},
        {"k":"rfc",              "l":"RFC",                "v":rfc,    "c":"high" if rfc else "low"},
        {"k":"telefono",         "l":"Teléfono",           "v":tel,    "c":"med" if tel else "low"},
        {"k":"vigencia_inicio",  "l":"Vigencia inicio",    "v":vigini, "c":"high" if vigini else "low"},
        {"k":"vigencia_fin",     "l":"Vigencia fin",       "v":vigfin, "c":"high" if vigfin else "low"},
        {"k":"suma_asegurada",   "l":"Suma asegurada",     "v":"$"+sa if sa else "", "c":"high" if sa else "low"},
        {"k":"prima_neta",       "l":"Prima neta",         "v":"$"+prima if prima else "", "c":"high" if prima else "low"},
    ]
    ok = sum(1 for c in campos if c["c"] == "high")
    confianza = round(ok / len(campos) * 100)

    return {
        "tipo": "single",
        "compania": compania or "Detectar",
        "ramo": ramo or "Detectar",
        "numero_poliza": poliza or "—",
        "confianza": confianza,
        "campos": campos,
        "texto_ocr": texto[:800]
    }


def normalizar_col(c):
    """Normaliza nombre de columna para comparación."""
    return str(c).strip().lower()\
        .replace(" ","_").replace("é","e").replace("ó","o")\
        .replace("í","i").replace("á","a").replace("ú","u").replace("ñ","n")


COL_MAP = {
    "nombre":          ["nombre","name","asegurado","contratante","titular","cliente","nombre_completo","aseguradotitular","cliente_nombre","nombre_cliente","razon_social"],
    "compania":        ["compania","compañia","aseguradora","empresa","company","cia","cia.","cía","companía","compañía"],
    "ramo":            ["ramo","producto","tipo","product","cobertura","tipo_seguro","ramo_seguro","ramo_principal","subramo"],
    "prima":           ["prima","prima_neta","prima_anual","importe","monto","precio","costo","prima_total","monto_prima"],
    "poliza":          ["poliza","poliza_num","no_poliza","numero","num_poliza","policy","no.poliza","#poliza","numero_poliza","documento","no_doc","certificado"],
    "fecha_nacimiento":["fecha_nac","f_nac","nacimiento","birthday","fec_nac","fecha_nacimiento","fnacimiento","fecha_de_nacimiento","fecha_nac."],
    "telefono":        ["telefono","tel","phone","celular","movil","cel","telefono_celular","tel_celular"],
    "vigencia_inicio": ["vigencia","vigencia_inicio","inicio","inicio_vigencia","fecha_inicio","vig_inicio","inicio_de_vigencia","fecha_inicio_vigencia"],
    "vigencia_fin":    ["vigencia_fin","fin","fin_vigencia","fecha_fin","vencimiento","expiracion","vig_fin","fin_de_vigencia"],
    "suma_asegurada":  ["suma_asegurada","suma","sa","monto_asegurado","valor_asegurado","suma_aseg"],
}


def renombrar_columnas(df):
    rename = {}
    for target, opciones in COL_MAP.items():
        for col in df.columns:
            if normalizar_col(col) in opciones and target not in rename.values():
                rename[col] = target
                break
    return df.rename(columns=rename)


def detectar_fila_header(df_raw):
    """
    Busca en las primeras 15 filas cuál es el encabezado real.
    Devuelve el índice de fila con más columnas reconocibles.
    """
    todos_campos = set()
    for vals in COL_MAP.values():
        todos_campos.update(vals)

    mejor_fila = 0
    mejor_score = 0
    for i in range(min(15, len(df_raw))):
        row = df_raw.iloc[i]
        score = sum(1 for v in row if normalizar_col(str(v)) in todos_campos)
        if score > mejor_score:
            mejor_score = score
            mejor_fila = i

    return mejor_fila if mejor_score >= 1 else 0


def es_fila_seccion(row):
    """
    Detecta si una fila es un separador de sección (ej. 'ENERO — 191 registros')
    y no un registro de cliente.
    """
    vals = [v for v in row if str(v) not in ("nan", "None", "", "NaT")]
    if len(vals) <= 2:
        return True  # fila casi vacía → sección o separador
    # Si el primer valor tiene "registros" o "total" y los demás son NaN
    primer = str(vals[0]).lower()
    if any(p in primer for p in ["registros", "total", "subtotal", "resumen", "detalle", "cobranza", "pendiente"]):
        return True
    return False


def df_a_polizas(df, hoja=""):
    """Convierte un DataFrame a lista de pólizas, filtrando filas de sección."""
    import pandas as pd
    df = renombrar_columnas(df)
    df = df.dropna(how="all")
    polizas = []
    campos_interes = ["nombre","compania","ramo","prima","poliza","fecha_nacimiento",
                      "telefono","vigencia_inicio","vigencia_fin","suma_asegurada"]
    for _, row in df.iterrows():
        # Saltar filas que son separadores de sección
        if es_fila_seccion(row):
            continue
        p = {}
        for k in campos_interes:
            val = str(row[k]).strip() if k in df.columns and str(row.get(k,"")) not in ("nan","None","","NaT") else ""
            p[k] = val
        p["hoja"] = hoja
        p["ok"] = bool(p.get("nombre") and len(p.get("nombre","")) > 2)
        if p["ok"] or any(p.get(k) for k in ["poliza","compania","ramo"]):
            polizas.append(p)
    return polizas


def leer_hoja_inteligente(xls, nombre_hoja):
    """
    Lee una hoja de Excel detectando automáticamente la fila de encabezado.
    Funciona con archivos de cualquier tamaño.
    """
    import pandas as pd
    # Primero leer sin header para detectar dónde están los encabezados
    df_raw = pd.read_excel(xls, sheet_name=nombre_hoja, header=None)
    if df_raw.shape[1] < 2 or df_raw.shape[0] < 2:
        return None

    fila_header = detectar_fila_header(df_raw)

    # Leer desde la fila de header (sin límite de filas)
    df = pd.read_excel(xls, sheet_name=nombre_hoja, header=fila_header)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def procesar_excel(ruta):
    """
    Lee un Excel con una o varias hojas — cualquier tamaño, cualquier estructura.
    Detecta automáticamente dónde están los encabezados y filtra separadores.
    """
    try:
        import pandas as pd
        ext = Path(ruta).suffix.lower()

        if ext == ".csv":
            # Intentar diferentes encodings
            for enc in ("utf-8-sig", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(ruta, encoding=enc)
                    break
                except Exception:
                    continue
            polizas = df_a_polizas(df)
            hojas_info = [{"hoja": "CSV", "filas": len(polizas)}]
        else:
            xls = pd.ExcelFile(ruta)
            hojas_info = []
            polizas = []

            for nombre_hoja in xls.sheet_names:
                try:
                    df = leer_hoja_inteligente(xls, nombre_hoja)
                    if df is None or df.shape[0] < 1:
                        continue
                    ps = df_a_polizas(df, hoja=nombre_hoja)
                    if ps:
                        polizas.extend(ps)
                        hojas_info.append({"hoja": nombre_hoja, "filas": len(ps)})
                except Exception:
                    continue

        return {
            "tipo":    "batch",
            "archivo": Path(ruta).name,
            "polizas": polizas,
            "hojas":   hojas_info,
            "total":   len(polizas),
        }
    except Exception as e:
        return {"tipo": "error", "error": str(e)}


# ── Rutas Flask ───────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/upload", methods=["POST"])
def upload():
    archivo = request.files.get("file")
    if not archivo:
        return jsonify({"error":"Sin archivo"}), 400

    ruta = UPLOAD_DIR / archivo.filename
    archivo.save(str(ruta))
    ext = ruta.suffix.lower()

    if ext in (".xlsx", ".xls", ".csv"):
        resultado = procesar_excel(str(ruta))
    else:
        texto = ocr_con_vision(str(ruta))
        resultado = parsear_texto(texto)
        resultado["texto_ocr"] = texto[:600]

    return jsonify(resultado)


@app.route("/guardar", methods=["POST"])
def guardar():
    body = request.get_json()
    tipo = body.get("tipo")
    data = body.get("data", {})
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")

    guardados = 0

    if tipo == "single":
        campos = {c["k"]: c["v"] for c in data.get("campos", [])}
        fila = {
            "fecha_importacion": ahora,
            "compania":          data.get("compania",""),
            "ramo":              data.get("ramo",""),
            "numero_poliza":     data.get("numero_poliza",""),
            "nombre":            campos.get("nombre",""),
            "fecha_nacimiento":  campos.get("fecha_nacimiento",""),
            "rfc":               campos.get("rfc",""),
            "telefono":          campos.get("telefono",""),
            "prima_neta":        campos.get("prima_neta",""),
            "suma_asegurada":    campos.get("suma_asegurada",""),
            "vigencia_inicio":   campos.get("vigencia_inicio",""),
            "vigencia_fin":      campos.get("vigencia_fin",""),
        }
        _append_csv(POLIZAS_CSV, fila)
        if fila.get("nombre") and fila.get("fecha_nacimiento"):
            _append_cumple(fila)
        guardados = 1

    elif tipo == "batch":
        for p in data.get("polizas", []):
            fila = {"fecha_importacion":ahora, **{k:p.get(k,"") for k in
                    ["nombre","compania","ramo","poliza","prima","fecha_nacimiento","telefono"]}}
            _append_csv(POLIZAS_CSV, fila)
            if p.get("nombre") and p.get("fecha_nacimiento"):
                _append_cumple(p)
            guardados += 1

    # Verificar si hay teléfonos para WhatsApp
    wa_disponible = False
    if tipo == "single":
        campos = {c["k"]: c["v"] for c in data.get("campos", [])}
        wa_disponible = bool(campos.get("telefono", "").strip())
    elif tipo == "batch":
        wa_disponible = any(p.get("telefono", "").strip() for p in data.get("polizas", []))

    return jsonify({
        "mensaje": f"✅ {guardados} póliza(s) guardada(s) en SICAMEZ",
        "wa_disponible": wa_disponible and bool(TWILIO_SID)
    })


def _append_csv(ruta, fila):
    existe = ruta.exists()
    with open(ruta, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(fila.keys()))
        if not existe:
            w.writeheader()
        w.writerow(fila)


def _append_cumple(p):
    existe = CUMPLE_CSV.exists()
    fila = {
        "nombre":            p.get("nombre",""),
        "telefono":          p.get("telefono",""),
        "fecha_nacimiento":  p.get("fecha_nacimiento",""),
        "compania":          p.get("compania",""),
        "ramo":              p.get("ramo",""),
        "tipo":              "titular",
        "nombre_titular":    "",
        "telefono_titular":  "",
    }
    with open(CUMPLE_CSV, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(fila.keys()))
        if not existe:
            w.writeheader()
        w.writerow(fila)


@app.route("/whatsapp", methods=["POST"])
def enviar_whatsapp():
    """Envía mensaje de bienvenida por WhatsApp a los clientes recién importados."""
    if not TWILIO_SID or not TWILIO_TOKEN:
        return jsonify({"mensaje": "❌ Credenciales Twilio no configuradas en .env"})

    try:
        from twilio.rest import Client
        client = Client(TWILIO_SID, TWILIO_TOKEN)
    except Exception as e:
        return jsonify({"mensaje": f"❌ Error conectando Twilio: {e}"})

    body = request.get_json()
    tipo = body.get("tipo")
    data = body.get("data", {})
    enviados = 0
    errores = 0

    def formatear_tel(t):
        tel = str(t).strip().replace(" ","").replace("-","").replace("(","").replace(")","")
        if tel.startswith("+"): tel = tel[1:]
        if not tel.startswith("52") and not (tel.startswith("1") and len(tel) == 11):
            tel = "52" + tel
        return f"whatsapp:+{tel}"

    def enviar_uno(nombre, telefono, compania="", ramo=""):
        if not telefono or not telefono.strip():
            return False
        primer = nombre.split()[0].capitalize() if nombre else "Cliente"
        det = ""
        if compania and ramo:
            det = f"\n📋 Póliza {ramo} — {compania}"
        elif compania:
            det = f"\n📋 {compania}"
        msg = (
            f"🎉 ¡Bienvenido(a), {primer}!\n\n"
            f"Nos da mucho gusto tenerte como cliente de Seguros Francisco.{det}\n\n"
            f"Cualquier duda o trámite, aquí estamos para atenderte.\n\n"
            f"— Seguros Francisco\nFrancisco Amezcua · Tu agente de seguros"
        )
        try:
            client.messages.create(from_=FROM_WA, to=formatear_tel(telefono), body=msg)
            return True
        except Exception:
            return False

    if tipo == "single":
        campos = {c["k"]: c["v"] for c in data.get("campos", [])}
        nombre   = campos.get("nombre", "")
        telefono = campos.get("telefono", "")
        compania = data.get("compania", "")
        ramo     = data.get("ramo", "")
        if enviar_uno(nombre, telefono, compania, ramo):
            enviados += 1
        else:
            errores += 1

    elif tipo == "batch":
        for p in data.get("polizas", []):
            if p.get("telefono"):
                if enviar_uno(p.get("nombre",""), p.get("telefono",""), p.get("compania",""), p.get("ramo","")):
                    enviados += 1
                else:
                    errores += 1

    if errores == 0:
        msg = f"✅ {enviados} bienvenida(s) enviada(s) por WhatsApp"
    else:
        msg = f"✅ {enviados} enviados · ⚠️ {errores} sin teléfono o con error"

    return jsonify({"mensaje": msg, "enviados": enviados, "errores": errores})


# ── Arranque ──────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    def abrir():
        import time; time.sleep(1.2)
        webbrowser.open("http://localhost:5001")
    threading.Thread(target=abrir, daemon=True).start()
    print("\n  SICAMEZ Importación corriendo en http://localhost:5001")
    print("  Ctrl+C para detener\n")
    app.run(host="0.0.0.0", port=5001, debug=False, threaded=True)
