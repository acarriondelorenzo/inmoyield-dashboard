import csv, json, re, sys, datetime

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

DATE_RE = re.compile(r'^\d{1,2}-\d{1,2}$')
MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto",
         "Septiembre","Octubre","Noviembre","Diciembre"]

def find_label_value(grid, label):
    """Find `label` anywhere in the grid and return the next non-empty cell in the same row."""
    for row in grid:
        for i, cell in enumerate(row):
            if cell.strip() == label:
                for j in range(i + 1, len(row)):
                    if row[j].strip():
                        return clean_num(row[j])
    return None

def find_all_label_values(grid, label):
    """Find every occurrence of `label` and return the list of next non-empty cell values, in row order."""
    out = []
    for row in grid:
        for i, cell in enumerate(row):
            if cell.strip() == label:
                for j in range(i + 1, len(row)):
                    if row[j].strip():
                        out.append(clean_num(row[j]))
                        break
    return out

def parse_treasury(csv_path):
    with open(csv_path, encoding='utf-8') as f:
        grid = list(csv.reader(f))

    caja = find_label_value(grid, 'Caja Disponible')
    fr = find_label_value(grid, 'Estimación FR (pendiente)')
    liq30 = find_label_value(grid, 'Liq. Est. 30 días')
    eic = find_label_value(grid, 'EIC')
    bafi = find_label_value(grid, 'Bafi')

    # monthly "margen de maniobra" figures, in the order they appear (row-major)
    margen_vals = find_all_label_values(grid, 'M de Maniobra')

    # month names present anywhere in the grid, in the order they appear, used to
    # line up against margen_vals (skip the first month found if there's a mismatch
    # in counts -> conservative: only pair what we can, don't guess).
    month_names_found = []
    for row in grid:
        for cell in row:
            c = cell.strip()
            if c in MESES and c not in month_names_found:
                month_names_found.append(c)

    margen_mensual = []
    if margen_vals:
        # heuristic: the months line up starting from the 2nd month found
        # (first month is usually already mid-way through and excluded from the
        # "M de Maniobra" forward-looking figures in this sheet's convention)
        candidate_months = month_names_found[1:1 + len(margen_vals)]
        if len(candidate_months) == len(margen_vals):
            margen_mensual = [{"mes": m, "valor": v} for m, v in zip(candidate_months, margen_vals)]

    mes_deficit = None
    if margen_mensual:
        peor = min(margen_mensual, key=lambda m: m['valor'])
        if peor['valor'] < 0:
            mes_deficit = peor['mes']

    # concrete upcoming payments: any row with a DD-MM date anchor in an early
    # column, followed by a label and a value
    pagos = []
    for row in grid:
        for i, cell in enumerate(row):
            if DATE_RE.match(cell.strip()):
                rest = [c for c in row[i + 1:] if c.strip()]
                if len(rest) >= 2:
                    val = clean_num(rest[1])
                    if val is not None:
                        pagos.append({"fecha": cell.strip(), "concepto": rest[0].strip(), "importe": val})
                break

    obligaciones_fijas = (abs(eic) if eic is not None else 0) + (abs(bafi) if bafi is not None else 0)

    result = {
        "cajaDisponible": caja,
        "frPendiente": fr,
        "liqEst30d": liq30,
        "obligacionesFijas": obligaciones_fijas,
        "margenMensual": margen_mensual,
        "mesDeficit": mes_deficit,
        "pagos": pagos,
        "fechaReferencia": datetime.date.today().isoformat(),
    }
    return result

def validate(t):
    reasons = []
    if t.get('cajaDisponible') is None:
        reasons.append("No se pudo localizar 'Caja Disponible' en la pestaña.")
    if t.get('liqEst30d') is None:
        reasons.append("No se pudo localizar 'Liq. Est. 30 días' en la pestaña.")
    if not t.get('margenMensual'):
        reasons.append("No se pudo construir el margen de maniobra mensual (formato de la pestaña puede haber cambiado).")
    if t.get('mesDeficit') is None and t.get('margenMensual'):
        # not necessarily an error (portfolio might have no deficit month at all),
        # but flag it for visibility since it's unusual
        pass
    if not t.get('pagos'):
        reasons.append("No se detectaron pagos concretos con fecha.")
    return (len(reasons) == 0), reasons

if __name__ == '__main__':
    csv_path, out_path = sys.argv[1], sys.argv[2]
    t = parse_treasury(csv_path)
    ok, reasons = validate(t)
    if not ok:
        print("TREASURY_VALIDATION_FAILED")
        for r in reasons:
            print(" -", r)
        sys.exit(1)
    json.dump(t, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print("OK", json.dumps(t, ensure_ascii=False))
