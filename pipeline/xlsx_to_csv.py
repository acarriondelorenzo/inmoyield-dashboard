"""Convierte una pestaña de un .xlsx (exportado del Google Sheet vía conector de Drive)
en un CSV con los valores *tal como se muestran* en Google Sheets en locale es-ES
(miles con punto, decimales con coma, "€", "%", fechas según formato de celda).

Así build.py / treasury_parse.py reciben exactamente el mismo formato que daba la
exportación CSV de Drive o el endpoint gviz, sin necesidad de navegador ni de
que el ordenador de Antonio esté encendido.

Uso: python3 xlsx_to_csv.py libro.xlsx "Nombre pestaña" salida.csv
"""
import csv, datetime, re, sys, warnings
warnings.filterwarnings("ignore")
import openpyxl


def es_num(v, dec):
    s = f"{abs(v):,.{dec}f}"  # 1,234,567.89
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if v < 0 and round(abs(v), dec) != 0 else "") + s


def fmt_date(d, nf):
    f = nf.lower().replace("\\", "")
    if f.startswith("yyyy-mm"):
        return d.strftime("%Y-%m")
    if f.startswith("yyyy-m"):
        return f"{d.year}-{d.month}"
    if f.startswith("dd-mm-yyyy"):
        return d.strftime("%d-%m-%Y")
    if f.startswith("dd/mm/yyyy"):
        return d.strftime("%d/%m/%Y")
    if f.startswith("d/mm/yyyy"):
        return f"{d.day}/{d.month:02d}/{d.year}"
    if f.startswith("d/m/yyyy"):
        return f"{d.day}/{d.month}/{d.year}"
    if f.startswith("dd-mm"):
        return d.strftime("%d-%m")
    return d.strftime("%d/%m/%Y")


def fmt(v, nf):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (datetime.datetime, datetime.date)):
        return fmt_date(v, nf or "")
    if isinstance(v, (int, float)):
        nf = nf or "General"
        if "%" in nf:
            m = re.search(r"0\.(0+)%", nf)
            return es_num(v * 100, len(m.group(1)) if m else 0) + "%"
        m = re.search(r"0\.(0+)", nf)
        dec = len(m.group(1)) if m else 0
        if nf == "General":
            if float(v).is_integer():
                return str(int(v))
            return str(v).replace(".", ",")
        s = es_num(v, dec)
        if "€" in nf:
            s += " €" if "\\ " in nf or " [$€" in nf or ' "€"' in nf else "€"
        return s
    return str(v)


def main(xlsx, sheet, out):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb[sheet]
    rows = []
    for r in ws.iter_rows():
        rows.append([fmt(c.value, c.number_format) for c in r])
    while rows and not any(x.strip() for x in rows[-1]):
        rows.pop()
    width = max((max((i + 1 for i, x in enumerate(r) if x != ""), default=0) for r in rows), default=0)
    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        for r in rows:
            w.writerow((r + [""] * width)[:width])
    print(f"OK {len(rows)} filas x {width} columnas -> {out}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
