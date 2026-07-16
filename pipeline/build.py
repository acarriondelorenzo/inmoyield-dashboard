import csv, json, sys, datetime, os
from collections import defaultdict, Counter

# (expected label as it appears in row 4 of the sheet export, internal key)
# Matching is done by NAME, not position, so the parser survives columns being
# reordered or new columns being inserted in the source sheet.
EXPECTED_COLUMNS = [
    ("Proyectos", "Proyectos"), ("Fase", "Fase"), ("SL", "SL"),
    ("Provincia", "Provincia"), ("Municipio", "Municipio"),
    ("Tipo Activo", "TipoActivo"), ("Tipología", "Tipologia"), ("Canal", "Canal"),
    ("Partner", "Partner"), ("Formato de inversión", "Formato"), ("Contrato", "Contrato"),
    ("Inv. Total. Est.", "InvTotalEst"), ("Invertido", "Invertido"),
    ("FC Inicial", "FCInicial"), ("FC Final", "FCFinal"),
    ("B° Realizado", "BRealizado"), ("Est. B°", "EstB"),
    ("Fecha inicio", "FechaInicio"), ("Fecha Fin", "FechaFin"),
    ("Est. Fin0", "EstFin0"), ("Retraso", "Retraso"), ("Est. Fin1", "EstFin1"),
    ("% Ejecución", "PctEjecucion"), ("Est. TIR", "EstTIR"), ("Est. ROI", "EstROI"),
    ("Meses", "Meses"), ("TIR", "TIR"), ("ROI", "ROI"), ("Link", "Link"),
    ("Estado", "Estado"), ("Comentarios", "Comentarios"),
]
# columns without which we cannot reliably build the active-investments table;
# if any of these can't be located by name in the actual header row, abort
# rather than risk silently misaligned data.
CRITICAL_KEYS = {"Proyectos", "Fase", "Partner", "Invertido", "FechaInicio", "FechaFin",
                  "EstFin0", "EstFin1", "Retraso", "PctEjecucion", "EstTIR", "Estado", "Comentarios"}

def _norm(s):
    return ' '.join((s or '').split()).strip()

def resolve_columns(header_row):
    """Map each expected internal key to the column index found in the actual
    header row, matched by name. Returns (index_by_key, missing_keys, unknown_columns)."""
    actual = [_norm(c) for c in header_row]
    label_to_idx = {}
    for i, cell in enumerate(actual):
        if cell:
            label_to_idx[cell] = i
    index_by_key = {}
    missing = []
    for label, key in EXPECTED_COLUMNS:
        if label in label_to_idx:
            index_by_key[key] = label_to_idx[label]
        else:
            missing.append((label, key))
    known_labels = {label for label, _ in EXPECTED_COLUMNS}
    unknown_columns = [c for c in actual if c and c not in known_labels]
    return index_by_key, missing, unknown_columns

def row_get(cells, index_by_key, key, default=''):
    idx = index_by_key.get(key)
    if idx is None or idx >= len(cells):
        return default
    return cells[idx]

def clean_num(s):
    if s is None: return None
    s = s.replace('\\-', '-').replace('€', '').strip()
    if s == '': return None
    neg = s.startswith('-')
    s = s.lstrip('-').strip()
    s = s.replace('.', '').replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except: return None

def clean_pct(s):
    if s is None: return None
    s = s.replace('\\-', '-').replace('%', '').strip()
    if s == '': return None
    s = s.replace('.', '').replace(',', '.')
    try: return float(s)
    except: return None

def clean_date(s):
    if not s: return None
    s = s.strip().replace('.', '/')
    parts = s.split('/')
    if len(parts) == 3:
        d, m, y = parts
        try: return f"{y}-{int(m):02d}-{int(d):02d}"
        except: return None
    return None

class SchemaError(Exception):
    def __init__(self, missing, unknown):
        self.missing = missing
        self.unknown = unknown
        labels = ", ".join(f"'{label}'" for label, key in missing)
        msg = f"La estructura de la hoja parece haber cambiado: no se encontraron las columnas críticas {labels}."
        if unknown:
            msg += f" Columnas nuevas/no reconocidas presentes en la hoja: {', '.join(unknown)}."
        super().__init__(msg)

def parse_csv(csv_path):
    """Parse the raw CSV export of the 'BBDD INVERSIONES' tab into a list of
    active-investment dicts (Estado == 'Invertido'), with duplicate/anomaly flags.
    Columns are located by NAME (header row), not fixed position, so the parser
    keeps working if columns get reordered or new ones are inserted upstream."""
    with open(csv_path, encoding='utf-8') as f:
        lines = f.readlines()
    # first 3 lines are metadata (title row, date row, blank row), 4th is the real header
    header_line = lines[3] if len(lines) > 3 else ''
    header_row = next(csv.reader([header_line]), [])
    index_by_key, missing, unknown_columns = resolve_columns(header_row)

    missing_critical = [(label, key) for label, key in missing if key in CRITICAL_KEYS]
    if missing_critical:
        raise SchemaError(missing_critical, unknown_columns)

    data_lines = lines[4:]
    reader = csv.reader(data_lines)
    all_rows = list(reader)

    min_cols = max(index_by_key.values()) + 1 if index_by_key else 0
    total_rows = len([r for r in all_rows if len(r) >= min_cols and row_get(r, index_by_key, 'Proyectos').strip()])

    def g(cells, key):
        return row_get(cells, index_by_key, key)

    parsed = []
    for cells in all_rows:
        if len(cells) < min_cols: continue
        if not g(cells, 'Proyectos').strip(): continue
        if g(cells, 'Estado').strip() != 'Invertido': continue
        parsed.append({
            "proyecto": g(cells, 'Proyectos'), "fase": g(cells, 'Fase'), "sl": g(cells, 'SL'),
            "provincia": g(cells, 'Provincia') or None, "municipio": g(cells, 'Municipio') or None,
            "tipoActivo": g(cells, 'TipoActivo'), "tipologia": g(cells, 'Tipologia'), "canal": g(cells, 'Canal'),
            "partner": g(cells, 'Partner'), "formato": g(cells, 'Formato'), "contrato": g(cells, 'Contrato'),
            "importe": clean_num(g(cells, 'Invertido')),
            "fechaInicio": clean_date(g(cells, 'FechaInicio')),
            "fechaFin": clean_date(g(cells, 'FechaFin')),
            "estFin0": g(cells, 'EstFin0').strip() or None, "estFin1": g(cells, 'EstFin1').strip() or None,
            "retraso": clean_num(g(cells, 'Retraso')),
            "pctEjecucion": clean_pct(g(cells, 'PctEjecucion')),
            "estTIR": clean_pct(g(cells, 'EstTIR')), "estROI": clean_pct(g(cells, 'EstROI')),
            "meses": g(cells, 'Meses').strip() or None,
            "comentarios": g(cells, 'Comentarios').strip(),
            "flags": []
        })

    # generic duplicate detection: same proyecto name appearing more than once with
    # identical importe/fechaInicio/comentarios across "Fase" -> likely a copy/paste
    # duplicate in the sheet rather than a genuine second tranche. Keep first, flag it,
    # drop the rest (do NOT silently double-count capital).
    groups = defaultdict(list)
    for i, r in enumerate(parsed):
        groups[r['proyecto']].append(i)
    drop_idx = set()
    for proyecto, idxs in groups.items():
        if len(idxs) < 2: continue
        keyf = lambda i: (parsed[i]['importe'], parsed[i]['fechaInicio'], parsed[i]['comentarios'])
        if len(set(keyf(i) for i in idxs)) == 1:
            keep = idxs[0]
            parsed[keep]['flags'].append(
                f"Existe(n) {len(idxs)-1} fila(s) adicional(es) idéntica(s) de '{proyecto}' en otra Fase "
                f"con importe/fecha/comentario idénticos -> posible duplicado, contado una sola vez. Verificar con el equipo.")
            drop_idx.update(idxs[1:])
    out = [r for i, r in enumerate(parsed) if i not in drop_idx]

    # anomaly: FC Final / Fecha Fin registrada pese a seguir "Invertido"
    for cells in all_rows:
        if len(cells) < min_cols: continue
        if g(cells, 'Estado').strip() == 'Invertido' and g(cells, 'FechaFin').strip():
            for r in out:
                if r['proyecto'] == g(cells, 'Proyectos') and r['fase'] == g(cells, 'Fase'):
                    r['flags'].append("Tiene FC Final / Fecha Fin registrada pese a Estado=Invertido -> verificar si sigue activa.")

    return out, total_rows, missing, unknown_columns


def validate(new_rows, total_rows_seen, baseline_rows):
    """Safeguards: refuse to publish if the fetch looks broken/truncated or the
    portfolio has swung by an implausible amount vs. yesterday. Returns (ok, reasons)."""
    reasons = []
    if total_rows_seen < 60:
        reasons.append(f"La hoja fuente solo devolvió {total_rows_seen} filas totales (se esperaban ~80). Posible lectura truncada.")
    if len(new_rows) < 15:
        reasons.append(f"Solo se detectaron {len(new_rows)} operaciones activas (Estado=Invertido). Cifra sospechosamente baja.")
    new_total = sum(r['importe'] or 0 for r in new_rows)
    if baseline_rows:
        old_total = sum(r['importe'] or 0 for r in baseline_rows)
        if old_total > 0:
            delta_pct = abs(new_total - old_total) / old_total
            if delta_pct > 0.25:
                reasons.append(f"El capital invertido total cambia un {delta_pct*100:.0f}% respecto a ayer ({old_total:,.0f}€ -> {new_total:,.0f}€), por encima del umbral de seguridad (25%).")
    return (len(reasons) == 0), reasons


def diff_changelog(old_rows, new_rows):
    def key(r): return (r['proyecto'], r['fase'])
    old_map = {key(r): r for r in old_rows}
    new_map = {key(r): r for r in new_rows}
    changes = []
    for k, r in new_map.items():
        if k not in old_map:
            changes.append({"tipo": "nuevo_detectado", "proyecto": r['proyecto'], "fase": r['fase'],
                             "partner": r['partner'],
                             "detalle": f"Aparece en cartera activa por {r['importe']:,.0f} €".replace(',', '.')})
    for k, r in old_map.items():
        if k not in new_map:
            changes.append({"tipo": "ya_no_activa", "proyecto": r['proyecto'], "fase": r['fase'],
                             "partner": r['partner'],
                             "detalle": f"Ya no aparece como 'Invertido' (antes {r['importe']:,.0f} €)".replace(',', '.')})
    for k in old_map.keys() & new_map.keys():
        o, n = old_map[k], new_map[k]
        if o['importe'] != n['importe']:
            changes.append({"tipo": "importe", "proyecto": n['proyecto'], "fase": n['fase'], "partner": n['partner'],
                             "detalle": f"Importe invertido: {o['importe']:,.0f} € → {n['importe']:,.0f} €".replace(',', '.')})
        if (o['comentarios'] or '').strip() != (n['comentarios'] or '').strip() and n['comentarios']:
            changes.append({"tipo": "comentario", "proyecto": n['proyecto'], "fase": n['fase'], "partner": n['partner'],
                             "detalle": f"Nuevo comentario: \"{n['comentarios']}\""})
        if o['estTIR'] != n['estTIR']:
            changes.append({"tipo": "tir", "proyecto": n['proyecto'], "fase": n['fase'], "partner": n['partner'],
                             "detalle": f"TIR estimada: {o['estTIR']}% → {n['estTIR']}%"})
    return changes


def update_history(history, changes, today_str):
    """Roll the changelog into a 7-day and 30-day rolling window."""
    if 'log' not in history:
        history['log'] = []  # list of {"fecha": "YYYY-MM-DD", "cambios": [...]}
    if changes:
        history['log'].append({"fecha": today_str, "cambios": changes})
    today = datetime.date.fromisoformat(today_str)
    week_start = today - datetime.timedelta(days=7)
    month_start = today - datetime.timedelta(days=30)
    log = history['log']
    week_changes = [c for entry in log if datetime.date.fromisoformat(entry['fecha']) > week_start for c in entry['cambios']]
    month_changes = [c for entry in log if datetime.date.fromisoformat(entry['fecha']) > month_start for c in entry['cambios']]
    # trim log to last 35 days to keep the file small
    history['log'] = [entry for entry in log if datetime.date.fromisoformat(entry['fecha']) > month_start]
    result = {
        "semana": {"desde": week_start.isoformat(), "hasta": today_str, "cambios": week_changes},
        "mes": {"desde": month_start.isoformat(), "hasta": today_str, "cambios": month_changes}
    }
    return result, history


def write_snapshot(snapshots_dir, date_str, data_rows, treasury, keep_days=35):
    """Write today's dated snapshot (data + treasury) and refresh the index of
    available dates, pruning anything older than `keep_days`. Snapshots live at
    the repo root (served as static files by Vercel) so the dashboard can fetch
    a past day's JSON directly, e.g. /snapshots/2026-07-16.json."""
    os.makedirs(snapshots_dir, exist_ok=True)
    snap = {"fecha": date_str, "data": data_rows, "treasury": treasury}
    with open(os.path.join(snapshots_dir, f"{date_str}.json"), 'w', encoding='utf-8') as f:
        json.dump(snap, f, ensure_ascii=False)

    cutoff = datetime.date.today() - datetime.timedelta(days=keep_days)
    kept = []
    for fname in os.listdir(snapshots_dir):
        if not fname.endswith('.json') or fname == 'index.json':
            continue
        d = fname[:-5]
        try:
            dd = datetime.date.fromisoformat(d)
        except ValueError:
            continue
        if dd < cutoff:
            os.remove(os.path.join(snapshots_dir, fname))
        else:
            kept.append(d)
    kept.sort()
    with open(os.path.join(snapshots_dir, 'index.json'), 'w', encoding='utf-8') as f:
        json.dump(kept, f, ensure_ascii=False)
    return kept


def render_dashboard(template_path, out_path, data_rows, treasury, history_public, source_date_str):
    with open(template_path, encoding='utf-8') as f:
        html = f.read()
    html = html.replace('__DATA_JSON__', json.dumps(data_rows, ensure_ascii=False))
    html = html.replace('__TREASURY_JSON__', json.dumps(treasury, ensure_ascii=False))
    html = html.replace('__HISTORY_JSON__', json.dumps(history_public, ensure_ascii=False))
    html = html.replace('__SOURCE_DATE__', source_date_str)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


if __name__ == '__main__':
    # Usage: build.py <live_csv> <template_html> <baseline_json> <history_json> <treasury_json> <out_html> <out_baseline> <out_history> <snapshots_dir>
    (live_csv, template_html, baseline_json, history_json, treasury_json,
     out_html, out_baseline, out_history) = sys.argv[1:9]
    snapshots_dir = sys.argv[9] if len(sys.argv) > 9 else None

    today_str = datetime.date.today().isoformat()

    try:
        new_rows, total_rows_seen, missing_noncritical, unknown_columns = parse_csv(live_csv)
    except SchemaError as e:
        print("VALIDATION_FAILED")
        print(" -", str(e))
        sys.exit(1)

    if missing_noncritical:
        labels = ", ".join(f"'{label}'" for label, key in missing_noncritical)
        print(f"AVISO (no crítico): no se encontraron estas columnas esperadas, se han dejado en blanco: {labels}")
    if unknown_columns:
        print(f"AVISO: columnas presentes en la hoja que no se reconocen (puede ser una columna nueva añadida en origen): {', '.join(unknown_columns)}")

    try:
        baseline_rows = json.load(open(baseline_json, encoding='utf-8'))
    except FileNotFoundError:
        baseline_rows = []

    ok, reasons = validate(new_rows, total_rows_seen, baseline_rows)
    if not ok:
        print("VALIDATION_FAILED")
        for r in reasons:
            print(" -", r)
        sys.exit(1)

    changes = diff_changelog(baseline_rows, new_rows)

    try:
        history = json.load(open(history_json, encoding='utf-8'))
    except FileNotFoundError:
        history = {}
    history_public, history_full = update_history(history, changes, today_str)

    try:
        treasury = json.load(open(treasury_json, encoding='utf-8'))
    except FileNotFoundError:
        treasury = {}

    source_date_str = datetime.date.today().strftime('%d/%m/%Y')
    render_dashboard(template_html, out_html, new_rows, treasury, history_public, source_date_str)
    json.dump(new_rows, open(out_baseline, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    json.dump(history_full, open(out_history, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    snap_count = None
    if snapshots_dir:
        kept = write_snapshot(snapshots_dir, today_str, new_rows, treasury)
        snap_count = len(kept)

    total = sum(r['importe'] or 0 for r in new_rows)
    print("OK", "rows_total_sheet=", total_rows_seen, "activas=", len(new_rows), "total=", round(total, 2),
          "cambios=", len(changes), "snapshots_guardados=", snap_count)
