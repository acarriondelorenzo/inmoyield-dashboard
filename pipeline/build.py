import csv, json, sys, datetime
from collections import defaultdict, Counter

HEADER = ["Proyectos","Fase","SL","Provincia","Municipio","TipoActivo","Tipologia","Canal","Partner",
          "Formato","Contrato","InvTotalEst","Invertido","FCInicial","FCFinal","BRealizado","EstB",
          "FechaInicio","FechaFin","EstFin0","Retraso","EstFin1","PctEjecucion","EstTIR","EstROI",
          "Meses","TIR","ROI","Link","Estado","Comentarios"]

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

def parse_csv(csv_path):
    """Parse the raw CSV export of the 'BBDD INVERSIONES' tab into a list of
    active-investment dicts (Estado == 'Invertido'), with duplicate/anomaly flags."""
    with open(csv_path, encoding='utf-8') as f:
        lines = f.readlines()
    # first 3 lines are metadata (title row, date row, blank row), 4th is the real header
    data_lines = lines[4:]
    reader = csv.reader(data_lines)
    all_rows = list(reader)

    total_rows = len([r for r in all_rows if len(r) >= 31 and r[0].strip()])

    parsed = []
    for cells in all_rows:
        if len(cells) < 31: continue
        d = dict(zip(HEADER, cells))
        if not d.get('Proyectos', '').strip(): continue
        if d.get('Estado', '').strip() != 'Invertido': continue
        parsed.append({
            "proyecto": d['Proyectos'], "fase": d['Fase'], "sl": d['SL'],
            "provincia": d['Provincia'] or None, "municipio": d['Municipio'] or None,
            "tipoActivo": d['TipoActivo'], "tipologia": d['Tipologia'], "canal": d['Canal'],
            "partner": d['Partner'], "formato": d['Formato'], "contrato": d['Contrato'],
            "importe": clean_num(d['Invertido']),
            "fechaInicio": clean_date(d['FechaInicio']),
            "fechaFin": clean_date(d.get('FechaFin')),
            "estFin0": d['EstFin0'].strip() or None, "estFin1": d['EstFin1'].strip() or None,
            "retraso": clean_num(d['Retraso']),
            "pctEjecucion": clean_pct(d['PctEjecucion']),
            "estTIR": clean_pct(d['EstTIR']), "estROI": clean_pct(d['EstROI']),
            "meses": d['Meses'].strip() or None,
            "comentarios": d['Comentarios'].strip(),
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
        if len(cells) < 31: continue
        d = dict(zip(HEADER, cells))
        if d.get('Estado', '').strip() == 'Invertido' and d.get('FechaFin', '').strip():
            for r in out:
                if r['proyecto'] == d['Proyectos'] and r['fase'] == d['Fase']:
                    r['flags'].append("Tiene FC Final / Fecha Fin registrada pese a Estado=Invertido -> verificar si sigue activa.")

    return out, total_rows


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


def render_dashboard(template_path, out_path, data_rows, treasury, history_public):
    with open(template_path, encoding='utf-8') as f:
        html = f.read()
    html = html.replace('__DATA_JSON__', json.dumps(data_rows, ensure_ascii=False))
    html = html.replace('__TREASURY_JSON__', json.dumps(treasury, ensure_ascii=False))
    html = html.replace('__HISTORY_JSON__', json.dumps(history_public, ensure_ascii=False))
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


if __name__ == '__main__':
    # Usage: build.py <live_csv> <template_html> <baseline_json> <history_json> <treasury_json> <out_html> <out_baseline> <out_history>
    (live_csv, template_html, baseline_json, history_json, treasury_json,
     out_html, out_baseline, out_history) = sys.argv[1:9]

    today_str = datetime.date.today().isoformat()

    new_rows, total_rows_seen = parse_csv(live_csv)

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

    render_dashboard(template_html, out_html, new_rows, treasury, history_public)
    json.dump(new_rows, open(out_baseline, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    json.dump(history_full, open(out_history, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    total = sum(r['importe'] or 0 for r in new_rows)
    print("OK", "rows_total_sheet=", total_rows_seen, "activas=", len(new_rows), "total=", round(total, 2), "cambios=", len(changes))
