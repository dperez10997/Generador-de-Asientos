# -*- coding: utf-8 -*-
import io
import csv
import unicodedata
from datetime import datetime, timedelta
from collections import defaultdict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Generador de Asientos", layout="wide")
st.title("Generador de Asientos (TXT tabulado)")
st.caption("Sube tu Excel/CSV, configura la contrapartida y descarga el archivo listo para subir al sistema contable.")

REQUIRED_COLUMNS = [
    'GL_Account','GL_Month','GL_Year','GL_Group',
    'TransactionDate','DebitAmount','CreditAmount','JobNumber'
]

DATE_PATTERNS = [
    '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%Y/%m/%d',
    '%d-%m-%Y', '%m-%d-%Y', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S',
    '%m/%d/%Y %H:%M:%S'
]

MONTHS_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

def strip_accents(text):
    if text is None:
        return ''
    nf = unicodedata.normalize('NFD', str(text))
    return ''.join(ch for ch in nf if not unicodedata.combining(ch))

def parse_date(value):
    if value is None or str(value).strip() == '':
        raise ValueError("Fecha vacía")
    s = str(value).strip()
    s = s.replace('T', ' ').split('.')[0]
    # Excel serial
    try:
        num = float(s)
        base = datetime(1899, 12, 30)
        dt = base + timedelta(days=num)
        return f"{dt.month}/{dt.day}/{dt.year}"
    except Exception:
        pass
    for pat in DATE_PATTERNS:
        try:
            dt = datetime.strptime(s, pat)
            return f"{dt.month}/{dt.day}/{dt.year}"
        except Exception:
            continue
    if len(s) == 8 and s.isdigit():
        y, m, d = s[0:4], s[4:6], s[6:8]
        return f"{int(m)}/{int(d)}/{int(y)}"
    raise ValueError(f"No se reconoce formato de fecha: {s}")

def fmt_amount(x):
    s = str(x).strip()
    if s == '' or s.upper() == 'NA':
        val = 0.0
    else:
        s = s.replace(' ', '').replace(',', '')
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
    month_text, gl_month = month_name_es(row['GL_Month'])
    gl_year = str(int(float(row['GL_Year']))) if str(row['GL_Year']).strip() != '' else ''
    gl_group = strip_accents(str(row.get('GL_Group', '')).strip())
    trx_date = parse_date(row['TransactionDate'])
    debit = fmt_amount(row['DebitAmount'])
    credit = fmt_amount(row['CreditAmount'])
    job = str(row['JobNumber']).strip()

    gl_note_raw = f"Provisión {month_text}"
    gl_reference_raw = f"Provisión producción {month_text} {gl_year}"
    gl_note = strip_accents(gl_note_raw)
    gl_reference = strip_accents(gl_reference_raw)

    return {
        'GL_Account': gl_account, 'GL_Note': gl_note,
        'GL_Month': gl_month, 'GL_Year': gl_year, 'GL_Group': gl_group,
        'TransactionDate': trx_date, 'GL_Reference': gl_reference,
        'DebitAmount': debit, 'CreditAmount': credit,
        'JobNumber': job
    }

def add_auto_offsets(rows, offset_account='1300102.5', agg='none', offset_note=None, offset_ref=None):
    base_rows = [normalize_row(r) for r in rows]
    first = base_rows[0] if base_rows else None
    default_note = (first['GL_Note'] + ' (auto)') if first else 'Auto offset'
    default_ref = first['GL_Reference'] if first else 'Auto offset'
    offset_note = strip_accents(offset_note) if offset_note else default_note
    offset_ref = strip_accents(offset_ref) if offset_ref else default_ref

    credits = [r for r in base_rows if r['GL_Account'].strip().startswith('8') and float(r['CreditAmount']) > 0.0]

    added = []
    if agg == 'none':
        for r in credits:
            added.append({
                **r,
                'GL_Account': offset_account,
                'GL_Note': r['GL_Note'],
                'GL_Reference': r['GL_Reference'],
                'DebitAmount': r['CreditAmount'],
                'CreditAmount': '0.00'
            })
    elif agg in ('total', 'by_ref', 'by_job'):
        buckets = defaultdict(list)
        keyfn = (lambda r: 'TOTAL') if agg=='total' else ((lambda r: r['GL_Reference']) if agg=='by_ref' else (lambda r: r['JobNumber']))
        for r in credits:
            buckets[keyfn(r)].append(r)
        for key, group in buckets.items():
            total = sum(float(r['CreditAmount']) for r in group)
            g0 = group[0]
            added.append({
                **g0,
                'GL_Account': offset_account,
                'GL_Note': offset_note if agg=='total' else g0['GL_Note'],
                'GL_Reference': offset_ref if agg=='total' else g0['GL_Reference'],
                'DebitAmount': f"{total:.2f}",
                'CreditAmount': '0.00',
                'JobNumber': g0['JobNumber']
            })
    else:
        raise ValueError("Valor inválido para --agg")

    return base_rows + added

def totals(rows):
    d = sum(float(r['DebitAmount']) for r in rows)
    c = sum(float(r['CreditAmount']) for r in rows)
    return d, c, round(d - c, 2)

# --- Sidebar: opciones ---
st.sidebar.header("Opciones")
auto_offset = st.sidebar.checkbox("Generar contrapartida automática", value=True)
agg = st.sidebar.selectbox("Tipo de contrapartida", options=['total','none','by_ref','by_job'], index=0)
offset_account = st.sidebar.text_input("Cuenta de contrapartida", value="1300102.5")
offset_note = st.sidebar.text_input("Nota (si 'total')", value="Provision (auto)")
offset_ref = st.sidebar.text_input("Referencia (si 'total')", value="Provision produccion (auto)")

file = st.file_uploader("Sube tu archivo de entrada (CSV o Excel .xlsx)", type=['csv','xlsx'])
run = st.button("▶ Ejecutar y generar salida.txt")

if run:
    if not file:
        st.error("Sube un archivo primero.")
    else:
        try:
            # === LECTOR ROBUSTO DE CSV/EXCEL ===
            if file.name.lower().endswith('.csv'):
                raw = file.getvalue()  # bytes del CSV
                df = None
                for enc in ('utf-8-sig', 'cp1252', 'latin1'):
                    try:
                        text = raw.decode(enc, errors='strict')
                        df = pd.read_csv(
                            io.StringIO(text),
                            dtype=str,
                            keep_default_na=False,
                            sep=None,        # autodetecta , ; o \t
                            engine='python'  # necesario para sep=None
                        )
                        break
                    except Exception:
                        continue
                if df is None:
                    st.error("No pude leer el CSV. Guarda como 'CSV UTF-8 (delimitado por comas)' o sube un .xlsx.")
                    st.stop()
            else:
                df = pd.read_excel(file, dtype=str, engine='openpyxl')

            missing = ensure_headers(df)
            if missing:
                st.error(f"Faltan columnas requeridas: {', '.join(missing)}")
            else:
                rows = df.to_dict(orient='records')
                if auto_offset:
                    out_rows = add_auto_offsets(rows, offset_account=offset_account, agg=agg, offset_note=offset_note, offset_ref=offset_ref)
                else:
                    out_rows = [normalize_row(r) for r in rows]

                d, c, diff = totals(out_rows)

                # Construir TXT en memoria
                out_buffer = io.StringIO()
                for r in out_rows:
                    fields = [r['GL_Account'], r['GL_Note'], r['GL_Month'], r['GL_Year'], r['GL_Group'],
                              r['TransactionDate'], r['GL_Reference'], r['DebitAmount'], r['CreditAmount'],
                              r['JobNumber']]
                    out_buffer.write('\t'.join(fields) + '\n')
                data = out_buffer.getvalue().encode('utf-8')

                st.success("Archivo generado.")
                st.write(f"**Débitos:** {d:.2f}  |  **Créditos:** {c:.2f}  |  **Diferencia (D-C):** {diff:.2f}")
                if abs(diff) <= 0.01:
                    st.write("✅ Asiento **CUADRA**.")
                else:
                    st.warning("⚠️  Asiento **NO cuadra**. Revisa montos/agrupación.")

                st.download_button("⬇ Descargar salida.txt", data=data, file_name="salida.txt", mime="text/plain")
        except Exception as e:
            st.exception(e)

st.markdown("---")
st.markdown("**Formato esperado del archivo de entrada (mínimo):**")
st.code("GL_Account, GL_Month, GL_Year, GL_Group, TransactionDate, DebitAmount, CreditAmount, JobNumber", language="text")
st.caption("GL_Note y GL_Reference se generan automáticamente sin tildes.")
