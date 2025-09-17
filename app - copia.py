# -*- coding: utf-8 -*-
import io, os, re, unicodedata
from datetime import datetime, timedelta
from collections import defaultdict
import pandas as pd
import streamlit as st

# ---------- Config UI ----------
st.set_page_config(
    page_title="Generador de Asientos Producción",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed"
)
st.markdown(
    "<style>.stButton > button{background:#16a34a;color:#fff;border:0;border-radius:10px;padding:.8rem 1.2rem;font-weight:600} .stButton > button:hover{filter:brightness(.95)}</style>",
    unsafe_allow_html=True
)
st.title("Generador de Asientos Producción")
st.caption("Sube tu Excel/CSV, presiona **Ejecutar** y descarga el TXT tabulado listo para el sistema contable.")

# ---------- Utils ----------
REQUIRED_COLUMNS = [
    'GL_Account','GL_Month','GL_Year','GL_Group',
    'TransactionDate','DebitAmount','CreditAmount','JobNumber'
]
DATE_PATTERNS = [
    '%d/%m/%Y','%m/%d/%Y','%Y-%m-%d','%Y/%m/%d',
    '%d-%m-%Y','%m-%d-%Y','%Y-%m-%d %H:%M:%S','%d/%m/%Y %H:%M:%S','%m/%d/%Y %H:%M:%S'
]
MONTHS_ES = {
    1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
    7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'
}

def strip_accents(t):
    if t is None: return ''
    return ''.join(ch for ch in unicodedata.normalize('NFD', str(t)) if not unicodedata.combining(ch))

def parse_date(v):
    """Devuelve DD/MM/AAAA"""
    if v is None or str(v).strip()=='':
        raise ValueError("Fecha vacía")
    s = str(v).strip().replace('T',' ').split('.')[0]
    # Excel serial
    try:
        num = float(s); base = datetime(1899,12,30)
        dt = base + timedelta(days=num)
        return f"{dt.day}/{dt.month}/{dt.year}"  # DD/MM/AAAA
    except Exception:
        pass
    for pat in DATE_PATTERNS:
        try:
            dt = datetime.strptime(s, pat)
            return f"{dt.day}/{dt.month}/{dt.year}"  # DD/MM/AAAA
        except Exception:
            continue
    if len(s)==8 and s.isdigit():
        y,m,d = s[0:4], s[4:6], s[6:8]
        return f"{int(d)}/{int(m)}/{int(y)}"  # DD/MM/AAAA
    raise ValueError(f"No se reconoce formato de fecha: {s}")

def fmt_amount(x):
    s = str(x).strip()
    if s=='' or s.upper()=='NA':
        val = 0.0
    else:
        s = s.replace(' ','').replace(',','')
        val = float(s)
    return f"{val:.2f}"

def ensure_headers(df):
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return missing

def month_name_es(m):
    try:
        m_int = int(float(m))
    except Exception:
        raise ValueError(f"GL_Month inválido: {m}")
    if m_int not in MONTHS_ES:
        raise ValueError(f"GL_Month fuera de rango (1-12): {m}")
    return MONTHS_ES[m_int], str(m_int)

def normalize_row(row):
    gl_account = str(row['GL_Account']).strip()
    mes_nombre, gl_month = month_name_es(row['GL_Month'])
    gl_year = str(int(float(row['GL_Year']))) if str(row['GL_Year']).strip()!='' else ''
    gl_group = strip_accents(str(row.get('GL_Group','')).strip())
    trx_date = parse_date(row['TransactionDate'])
    debit = fmt_amount(row['DebitAmount'])
    credit = fmt_amount(row['CreditAmount'])
    job = str(row['JobNumber']).strip()
    gl_note = strip_accents(f"Provisión {mes_nombre}")
    gl_reference = strip_accents(f"Provisión producción {mes_nombre} {gl_year}")
    return {
        'GL_Account': gl_account, 'GL_Note': gl_note,
        'GL_Month': gl_month, 'GL_Year': gl_year, 'GL_Group': gl_group,
        'TransactionDate': trx_date, 'GL_Reference': gl_reference,
        'DebitAmount': debit, 'CreditAmount': credit, 'JobNumber': job
    }

def add_auto_offsets(rows, offset_account='1300102.5', agg='total',
                     offset_note="Provision (auto)", offset_ref="Provision produccion (auto)"):
    """
    Genera contrapartidas:
      - Si agg == 'total': UNA contrapartida por el total. Nota/Referencia se AUTOGENERAN por GL_Month/GL_Year.
      - Si agg == 'none': espejo por cada crédito 8xxxx.
      - Si agg == 'by_ref' / 'by_job': agrupa por referencia o por JobNumber, respetando Nota/Referencia originales.
    """
    base_rows = [normalize_row(r) for r in rows]
    credits = [r for r in base_rows if r['GL_Account'].strip().startswith('8') and float(r['CreditAmount'])>0.0]
    added = []

    if agg == 'none':
        for r in credits:
            added.append({
                **r, 'GL_Account': offset_account,
                'DebitAmount': r['CreditAmount'], 'CreditAmount': '0.00'
            })

    elif agg in ('total','by_ref','by_job'):
        buckets = defaultdict(list)
        if agg == 'total':
            keyfn = lambda r: 'TOTAL'
        elif agg == 'by_ref':
            keyfn = lambda r: r['GL_Reference']
        else:  # by_job
            keyfn = lambda r: r['JobNumber']

        for r in credits: buckets[keyfn(r)].append(r)

        for _, group in buckets.items():
            total = sum(float(r['CreditAmount']) for r in group)
            g0 = group[0]

            if agg == 'total':
                mes_nombre, _ = month_name_es(g0['GL_Month'])
                anno = g0['GL_Year']
                gl_note = strip_accents(f"Provisión {mes_nombre}")
                gl_ref  = strip_accents(f"Provisión producción {mes_nombre} {anno}")
            else:
                gl_note = g0['GL_Note']
                gl_ref  = g0['GL_Reference']

            added.append({
                **g0,
                'GL_Account': offset_account,
                'GL_Note': gl_note,
                'GL_Reference': gl_ref,
                'DebitAmount': f"{total:.2f}",
                'CreditAmount': '0.00'
            })
    else:
        raise ValueError("Valor inválido para 'agg'")

    return base_rows + added

def totals(rows):
    d = sum(float(r['DebitAmount']) for r in rows)
    c = sum(float(r['CreditAmount']) for r in rows)
    return d, c, round(d-c,2)

# ---------- Agencia -> Cuenta (base fija en repo) ----------
def read_any_csv(file_bytes):
    for enc in ('utf-8-sig','cp1252','latin1'):
        try:
            text = file_bytes.decode(enc, errors='strict')
            return pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep=None, engine='python')
        except Exception:
            continue
    return None

def strip_accents_local(s):
    if s is None: return ""
    return "".join(ch for ch in unicodedata.normalize("NFD", str(s)) if not unicodedata.combining(ch))

def normalize_cols(cols):
    out=[]
    for c in cols:
        c0=strip_accents_local(str(c)).lower().strip()
        c0=re.sub(r'[^a-z0-9]+','_', c0)
        out.append(c0)
    return out

def build_agency_map(df):
    df=df.copy(); df.columns=normalize_cols(df.columns)
    ag_cols=[c for c in df.columns if 'agencia' in c or 'cliente' in c or c in ('agencia','agente')]
    ct_cols=[c for c in df.columns if 'cuenta' in c or 'account' in c]
    if not ag_cols or not ct_cols: return {}
    agencia_col=ag_cols[0]; cuenta_col=ct_cols[0]
    mapping={}
    for _, r in df.iterrows():
        ag=str(r[agencia_col]).strip(); ct=str(r[cuenta_col]).strip()
        if ag and ct and ct.lower()!='nan': mapping[ag]=ct
    return mapping

def load_fixed_agency_map():
    for fname in ("agencias.xlsx","agencias.csv"):
        if os.path.isfile(fname):
            try:
                if fname.endswith(".xlsx"):
                    df=pd.read_excel(fname, dtype=str, engine='openpyxl')
                else:
                    with open(fname,"rb") as f: df=read_any_csv(f.read())
                m=build_agency_map(df)
                if m: return m
            except Exception: pass
    return {}

fixed_agency_map = load_fixed_agency_map()

# ---------- Selector de AGENCIA visible (arriba) ----------
selected_agency=None
offset_account_from_agency=None
use_agency_account=False
agency_error=False

if fixed_agency_map:
    cols_ag = st.columns([2,2,2])
    with cols_ag[0]:
        selected_agency = st.selectbox("Agencia", sorted(fixed_agency_map.keys()))
    with cols_ag[1]:
        use_agency_account = st.checkbox("Usar cuenta según agencia", value=True)

    if use_agency_account and selected_agency:
        if selected_agency not in fixed_agency_map or str(fixed_agency_map.get(selected_agency, "")).strip()=="":
            st.error(f"La agencia **{selected_agency}** no tiene una cuenta asignada en la base. Usa la **cuenta manual** en Opciones avanzadas.")
            use_agency_account = False
            agency_error = True
        else:
            offset_account_from_agency = fixed_agency_map[selected_agency]
            with cols_ag[2]:
                st.text_input("Cuenta (auto por agencia)", value=offset_account_from_agency, disabled=True)

# ---------- Uploader y botón ----------
file = st.file_uploader("Sube tu archivo de entrada (CSV o Excel .xlsx)", type=['csv','xlsx'])
run = st.button("▶ Ejecutar y generar salida.txt", disabled=agency_error)

# ---------- Mapeo forzado (utilidades) ----------
MONTHS_BY_NAME = {
    'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
    'julio':7,'agosto':8,'septiembre':9,'setiembre':9,'octubre':10,'noviembre':11,'diciembre':12
}

def to_number_safe(val):
    s = str(val).strip()
    if s == '' or s.upper() == 'NA':
        return 0.0
    s = s.replace(' ', '')
    if s.count(',')>0 and s.count('.')>0:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(',', '')
        else:
            s = s.replace('.', '').replace(',', '.')
    else:
        if s.count(',')==1 and s.count('.')==0:
            s = s.replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def apply_forced_mapping(df):
    """
    Mapeo forzado:
      GL_Account <- Código
      GL_Month   <- Mes   (acepta nombre o número)
      GL_Year    <- derivado de Fecha
      GL_Group   <- vacío
      TransactionDate <- Fecha (salida DD/MM/AAAA)
      DebitAmount <- 0.00
      CreditAmount <- Venta
      JobNumber <- Trabajo
    """
    def norm(c): return strip_accents_local(str(c)).lower().strip()
    cols = list(df.columns)
    ncols = {norm(c): c for c in cols}
    required = ['codigo','mes','fecha','venta','trabajo']
    missing = [r for r in required if r not in ncols]
    if missing:
        raise ValueError(
            "Faltan columnas en tu archivo: " + ", ".join(missing) +
            ". Debe incluir: Código, Mes, Fecha, Venta, Trabajo."
        )

    out = pd.DataFrame()
    out['GL_Account'] = df[ncols['codigo']].astype(str).str.strip()

    raw_m = df[ncols['mes']].astype(str).str.strip()
    def conv_m(x):
        if x.isdigit():
            return x
        nx = strip_accents_local(x).lower()
        return str(MONTHS_BY_NAME.get(nx, ''))
    out['GL_Month'] = raw_m.apply(conv_m)

    # Año derivado de Fecha
    raw_f = df[ncols['fecha']]
    def year_from_date(v):
        try:
            s = parse_date(v)           # DD/MM/AAAA
            return s.split('/')[-1]     # AAAA
        except:
            return ''
    out['GL_Year'] = raw_f.apply(year_from_date)

    out['GL_Group'] = ''  # vacío
    out['TransactionDate'] = raw_f.apply(parse_date)  # DD/MM/AAAA
    out['DebitAmount'] = '0.00'
    out['CreditAmount'] = df[ncols['venta']].apply(lambda v: f"{to_number_safe(v):.2f}")
    out['JobNumber'] = df[ncols['trabajo']].astype(str).str.strip()

    mapping_resumen = {
        "GL_Account": ncols['codigo'],
        "GL_Month":   ncols['mes'],
        "GL_Year":    "Derivado de " + ncols['fecha'],
        "GL_Group":   "(vacío)",
        "TransactionDate": ncols['fecha'] + " → DD/MM/AAAA",
        "DebitAmount": "0.00 (fijo)",
        "CreditAmount": ncols['venta'],
        "JobNumber":   ncols['trabajo'],
    }
    return out, mapping_resumen

# ---------- Opciones avanzadas ----------
with st.expander("Opciones avanzadas (contrapartida y texto)", expanded=False):
    agg = st.selectbox("Tipo de contrapartida", options=['total','none','by_ref','by_job'], index=0)
    auto_offset = st.checkbox("Generar contrapartida automática", value=True)
    manual_help = "Se usará la cuenta por agencia si está activado arriba." if use_agency_account else "Se usará esta cuenta."
    offset_account_manual = st.text_input("Cuenta de contrapartida (manual)", value="1300102.5", help=manual_help)
    st.text_input("Nota (si 'total')", value="(Se autogenera: Provisión <Mes>)", disabled=True)
    st.text_input("Referencia (si 'total')", value="(Se autogenera: Provisión producción <Mes> <Año>)", disabled=True)

st.markdown("---")
st.caption("Formato estándar: GL_Account, GL_Month, GL_Year, GL_Group, TransactionDate, DebitAmount, CreditAmount, JobNumber. "
           "Si subes un archivo con columnas **Código, Mes, Fecha, Venta, Trabajo** se aplica mapeo forzado automáticamente.")

# ---------- Ejecutar ----------
if run:
    if not file:
        st.error("Sube un archivo primero."); st.stop()
    try:
        # Leer archivo de entrada (csv/xlsx)
        if file.name.lower().endswith('.csv'):
            raw = file.getvalue(); df = None
            for enc in ('utf-8-sig','cp1252','latin1'):
                try:
                    text = raw.decode(enc, errors='strict')
                    df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep=None, engine='python')
                    break
                except Exception:
                    continue
            if df is None:
                st.error("No pude leer el CSV. Guarda como 'CSV UTF-8' o sube un .xlsx."); st.stop()
        else:
            df = pd.read_excel(file, dtype=str, engine='openpyxl')

        # Si el archivo ya es estándar, úsalo tal cual; si no, aplica mapeo forzado.
        missing_std = ensure_headers(df)
        if len(missing_std) == 0:
            df_std = df
            mapping_resumen = None
        else:
            df_std, mapping_resumen = apply_forced_mapping(df)

        # Tarjeta de confirmación del mapeo forzado (si aplica)
        if mapping_resumen:
            st.success("Se aplicó mapeo forzado para **Código/Mes/Fecha/Venta/Trabajo**.")
            st.markdown(
                """
**Mapeo aplicado:**

- **GL_Account** ← `{GL_Account}`
- **GL_Month** ← `{GL_Month}` (acepta número o nombre)
- **GL_Year** ← {GL_Year}
- **GL_Group** ← {GL_Group}
- **TransactionDate** ← `{TransactionDate}`
- **DebitAmount** ← {DebitAmount}
- **CreditAmount** ← `{CreditAmount}`
- **JobNumber** ← `{JobNumber}`
                """.format(**mapping_resumen)
            )
            # Compatibilidad: toggle o checkbox según versión de Streamlit
            if hasattr(st, "toggle"):
                show_preview = st.toggle("Mostrar vista previa del estándar aplicado", value=False)
            else:
                show_preview = st.checkbox("Mostrar vista previa del estándar aplicado", value=False)
            if show_preview:
                st.dataframe(df_std.head(10), use_container_width=True)

        # A partir de aquí, df_std tiene columnas estándar requeridas
        rows = df_std.to_dict(orient='records')

        # Cuenta de contrapartida efectiva (agencia o manual)
        effective_offset = offset_account_from_agency if (use_agency_account and offset_account_from_agency) else offset_account_manual

        # Generar salida
        if auto_offset:
            out_rows = add_auto_offsets(rows, offset_account=effective_offset, agg=agg)
        else:
            out_rows = [normalize_row(r) for r in rows]

        d, c, diff = totals(out_rows)

        # Construir TXT
        out_buffer = io.StringIO()
        for r in out_rows:
            fields = [r['GL_Account'], r['GL_Note'], r['GL_Month'], r['GL_Year'], r['GL_Group'],
                      r['TransactionDate'], r['GL_Reference'], r['DebitAmount'], r['CreditAmount'], r['JobNumber']]
            out_buffer.write('\t'.join(fields) + '\n')
        data = out_buffer.getvalue().encode('utf-8')

        st.success("Archivo generado.")
        st.write(f"**Débitos:** {d:.2f}  |  **Créditos:** {c:.2f}  |  **Diferencia (D-C):** {diff:.2f}")
        if abs(diff)<=0.01: st.write("✅ Asiento **CUADRA**.")
        else: st.warning("⚠️  Asiento **NO cuadra**. Revisa montos/agrupación.")

        st.download_button("⬇ Descargar salida.txt", data=data, file_name="salida.txt", mime="text/plain")

    except Exception as e:
        st.exception(e)
