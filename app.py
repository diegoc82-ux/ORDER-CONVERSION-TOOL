import io
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
import pdfplumber

app = Flask(__name__)
BASE = Path(__file__).parent


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/products.json')
def products():
    return send_file(BASE / 'products.json', mimetype='application/json')


@app.route('/imporcarsa_prices.json')
def imporcarsa_prices():
    return send_file(BASE / 'imporcarsa_prices.json', mimetype='application/json')


@app.route('/8a_prices.json')
def a8_prices():
    return send_file(BASE / '8a_prices.json', mimetype='application/json')


@app.route('/competitive_edge_prices.json')
def competitive_edge_prices():
    return send_file(BASE / 'competitive_edge_prices.json', mimetype='application/json')


@app.route('/beicruz_prices.json')
def beicruz_prices():
    return send_file(BASE / 'beicruz_prices.json', mimetype='application/json')


@app.route('/garner_prices.json')
def garner_prices():
    return send_file(BASE / 'garner_prices.json', mimetype='application/json')


@app.route('/pagsa_prices.json')
def pagsa_prices():
    return send_file(BASE / 'pagsa_prices.json', mimetype='application/json')


@app.route('/trebol_prices.json')
def trebol_prices():
    return send_file(BASE / 'trebol_prices.json', mimetype='application/json')


@app.route('/maxilub_prices.json')
def maxilub_prices():
    return send_file(BASE / 'maxilub_prices.json', mimetype='application/json')


@app.route('/servistar_prices.json')
def servistar_prices():
    return send_file(BASE / 'servistar_prices.json', mimetype='application/json')


@app.route('/frenoseguro_prices.json')
def frenoseguro_prices():
    return send_file(BASE / 'frenoseguro_prices.json', mimetype='application/json')


@app.route('/ukr_prices.json')
def ukr_prices():
    return send_file(BASE / 'ukr_prices.json', mimetype='application/json')


@app.route('/prosupply_prices.json')
def prosupply_prices():
    return send_file(BASE / 'prosupply_prices.json', mimetype='application/json')


@app.route('/abc_prices.json')
def abc_prices():
    return send_file(BASE / 'abc_prices.json', mimetype='application/json')


@app.route('/daher_prices.json')
def daher_prices():
    return send_file(BASE / 'daher_prices.json', mimetype='application/json')


@app.route('/asaray_prices.json')
def asaray_prices():
    return send_file(BASE / 'asaray_prices.json', mimetype='application/json')


@app.route('/dianca_prices.json')
def dianca_prices():
    return send_file(BASE / 'dianca_prices.json', mimetype='application/json')


@app.route('/costs.json')
def costs():
    return send_file(BASE / 'costs.json', mimetype='application/json')


@app.route('/houston_prices.json')
def houston_prices():
    return send_file(BASE / 'houston_prices.json', mimetype='application/json')


@app.route('/miami_prices.json')
def miami_prices():
    return send_file(BASE / 'miami_prices.json', mimetype='application/json')


@app.route('/api/extract-dispatch', methods=['POST'])
def extract_dispatch():
    import re
    from collections import defaultdict
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file uploaded'}), 400
    try:
        items, dispatch_num, po_number, order_date = [], '', '', ''
        with pdfplumber.open(io.BytesIO(f.read())) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                # Extract dispatch number ("Dispatch Note XXXX")
                for i, w in enumerate(words):
                    if w['text'] == 'Note' and i > 0 and words[i-1]['text'] == 'Dispatch':
                        if i+1 < len(words):
                            dispatch_num = words[i+1]['text']
                # Extract PO number and order date from right-side header
                for w in words:
                    if w['text'] == 'Number:' and w['x0'] > 500:
                        val = next((v['text'] for v in words
                                    if v['x0'] > 650 and abs(v['top'] - w['top']) < 5), '')
                        po_number = val
                    if w['text'] == 'order:' and w['x0'] > 500:
                        val = next((v['text'] for v in words
                                    if v['x0'] > 650 and abs(v['top'] - w['top']) < 5), '')
                        order_date = val
                # Group words by rounded top value (2px tolerance for sub-pixel differences)
                rows = defaultdict(list)
                for w in words:
                    rows[round(w['top'] / 2) * 2].append(w)
                # Find header row and detect column positions.
                # Supports two formats:
                #   Standard U1P:   CODE … SPEC … UNIT QTY …
                #   FRENOSEGURO:    UL Number … PACKAGE SIZE … UNIT QTY …
                header_top = None
                spec_lo = spec_hi = qty_lo = qty_hi = None
                for top, rw in sorted(rows.items()):
                    has_code = any(w['text'] == 'CODE' for w in rw)
                    has_spec = any(w['text'] == 'SPEC' for w in rw)
                    has_number = any(w['text'] == 'Number' for w in rw)
                    has_package = any(w['text'] == 'PACKAGE' for w in rw)
                    if (has_code and has_spec) or (has_number and has_package):
                        header_top = top
                        # Anchor for spec column: SPEC (standard) or PACKAGE (FRENOSEGURO)
                        spec_anchor = (next((w for w in rw if w['text'] == 'SPEC'), None)
                                       or next((w for w in rw if w['text'] == 'PACKAGE'), None))
                        unit_w = next((w for w in rw if w['text'] == 'UNIT'), None)
                        qty_candidates = sorted(
                            [w for w in rw if w['text'] == 'QTY' and unit_w and w['x0'] > unit_w['x1']],
                            key=lambda x: x['x0']
                        )
                        unit_qty_w = qty_candidates[0] if qty_candidates else None
                        if spec_anchor and unit_w:
                            spec_lo = spec_anchor['x0'] - 15
                            spec_hi = unit_w['x0'] - 3
                            qty_lo  = unit_w['x0'] - 5
                            qty_hi  = (unit_qty_w['x1'] if unit_qty_w else unit_w['x1']) + 35
                        break
                if header_top is None or spec_lo is None:
                    continue
                # Extract line items from rows below the header
                for top, row_words in sorted(rows.items()):
                    if top <= header_top:
                        continue
                    code_words = [w for w in row_words
                                  if re.match(r'^(UL|TR)\d+$', w['text'], re.I) and w['x0'] < 130]
                    if not code_words:
                        continue
                    code = code_words[0]['text'].upper()
                    spec_words = sorted([w for w in row_words if spec_lo <= w['x0'] <= spec_hi],
                                        key=lambda x: x['x0'])
                    presentation = ' '.join(w['text'] for w in spec_words)
                    if not presentation:
                        continue
                    qty_words = [w for w in row_words if qty_lo <= w['x0'] <= qty_hi]
                    if not qty_words:
                        continue
                    try:
                        qty = int(float(qty_words[0]['text'].replace(',', '')))
                    except (ValueError, TypeError):
                        continue
                    if qty <= 0:
                        continue
                    items.append({'ul_code': code, 'presentation': presentation, 'qty': qty})
        return jsonify({'items': items, 'dispatchNum': dispatch_num,
                        'poNumber': po_number, 'orderDate': order_date})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/extract-garner', methods=['POST'])
def extract_garner():
    import re
    from collections import defaultdict
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file uploaded'}), 400
    try:
        # Column x-midpoint ranges → presentation
        COL_PRES = [
            (690, 730, 'BOX (6G)'),    # 6/1-Gal Case
            (630, 690, 'PAIL (5G)'),   # 5G Pail
            (570, 630, 'DRUM (55G)'),  # 55G Drum
            (525, 570, 'TOTE (250G)'), # Tote (coolant bulk)
            (460, 525, 'BULK (COOL)'), # Bulk per gallon
        ]
        items, po_number, order_date = [], '', ''
        with pdfplumber.open(io.BytesIO(f.read())) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                # Extract PO number and date
                for i, w in enumerate(words):
                    if re.match(r'^PO#$', w['text'], re.I):
                        # PO# and its number are separate tokens — grab the next word
                        if i + 1 < len(words) and re.match(r'^\d+$', words[i + 1]['text']):
                            po_number = 'PO#' + words[i + 1]['text']
                        else:
                            po_number = w['text']
                    if re.match(r'^\d{1,2}-\d{1,2}-\d{2,4}$', w['text']):
                        order_date = w['text']
                # Group words into rows by top position (10pt bands)
                rows = defaultdict(list)
                for w in words:
                    rows[round(w['top'] / 10) * 10].append(w)
                for row_words in rows.values():
                    ul = next((w for w in row_words if re.match(r'^UL\d+$', w['text'])), None)
                    if not ul:
                        continue
                    for w in row_words:
                        try:
                            qty = float(w['text'].replace(',', ''))
                        except (ValueError, TypeError):
                            continue
                        if qty <= 0:
                            continue
                        x_mid = (w['x0'] + w['x1']) / 2
                        pres = next((p for lo, hi, p in COL_PRES if lo <= x_mid <= hi), None)
                        if pres:
                            items.append({'ul_code': ul['text'], 'presentation': pres, 'qty': int(qty)})
        return jsonify({'items': items, 'poNumber': po_number, 'orderDate': order_date})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/extract-ukr', methods=['POST'])
def extract_ukr():
    import re
    from collections import defaultdict
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file uploaded'}), 400
    try:
        items, po_number, order_date = [], '', ''
        with pdfplumber.open(io.BytesIO(f.read())) as pdf:
            for page in pdf.pages:
                words = page.extract_words()
                # Extract PO number: "PURCHASE ORDER # 74"
                for i, w in enumerate(words):
                    if w['text'] == '#' and i >= 2 and words[i-2]['text'] == 'PURCHASE':
                        if i + 1 < len(words) and words[i+1]['text'].isdigit():
                            po_number = 'PO_' + words[i+1]['text'].zfill(5)
                    if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', w['text']):
                        order_date = w['text']
                # Group words by row
                rows = defaultdict(list)
                for w in words:
                    rows[round(w['top'] / 3) * 3].append(w)
                # Find header row (has "SKU" and "Qty" columns)
                header_top = sku_lo = sku_hi = qty_lo = qty_hi = None
                for top, rw in sorted(rows.items()):
                    texts = [w['text'] for w in rw]
                    if 'SKU' in texts and 'Qty' in texts:
                        header_top = top
                        sku_w = next((w for w in rw if w['text'] == 'SKU'), None)
                        qty_w = next((w for w in rw if w['text'] == 'Qty'), None)
                        if sku_w and qty_w:
                            sku_lo = sku_w['x0'] - 5
                            sku_hi = sku_w['x1'] + 130
                            qty_lo = qty_w['x0'] - 10
                            qty_hi = qty_w['x1'] + 35
                        break
                if header_top is None or sku_lo is None:
                    continue
                for top, row_words in sorted(rows.items()):
                    if top <= header_top:
                        continue
                    sku_cands = [w for w in row_words
                                 if sku_lo <= w['x0'] <= sku_hi
                                 and re.match(r'^[A-Z]{2}\w{4,}$', w['text'], re.I)]
                    if not sku_cands:
                        continue
                    sku = sku_cands[0]['text'].upper()
                    qty_cands = [w for w in row_words
                                 if qty_lo <= w['x0'] <= qty_hi
                                 and re.match(r'^\d+$', w['text'])]
                    if not qty_cands:
                        continue
                    qty = int(qty_cands[0]['text'])
                    if qty <= 0:
                        continue
                    items.append({'sku': sku, 'qty': qty})
        return jsonify({'items': items, 'poNumber': po_number, 'orderDate': order_date})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/extract-pdf', methods=['POST'])
def extract_pdf():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file uploaded'}), 400
    try:
        with pdfplumber.open(io.BytesIO(f.read())) as pdf:
            text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        return jsonify({'text': text})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/counter', methods=['GET'])
def get_counter():
    import json
    path = BASE / 'dispatch_counter.json'
    return jsonify(json.loads(path.read_text()) if path.exists() else {'last': 3086})


@app.route('/api/counter', methods=['POST'])
def set_counter():
    import json
    path = BASE / 'dispatch_counter.json'
    path.write_text(json.dumps(request.get_json()))
    return jsonify({'ok': True})


@app.route('/api/generate-sli', methods=['POST'])
def generate_sli():
    import re
    import openpyxl
    from datetime import datetime, date
    import io as io_mod

    data        = request.get_json()
    consignee   = data.get('consignee', {})
    agent       = data.get('forwarding_agent', {})
    ref         = data.get('reference', '')
    date_str    = data.get('date', date.today().isoformat())
    items       = data.get('items', [])

    try:
        sli_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    except Exception:
        sli_date = date.today()

    GALS = {
        'BOX (12Q)':      3.0,   'BOX (3/5QTS)':   3.75,
        'BOX (4G)':       4.0,   'BOX (6G)':        6.0,
        'BOX (1 X 2.5G)': 2.5,   'BOX (2 X 2.5G)':  5.0,
        'DRUM (55G)':    55.0,   'PAIL (5G)':        5.0,
        'TOTE (265G)':  265.0,   'TOTE (250G)':    250.0,
        'TOTE (330G)':  330.0,   'BULK (OIL)':       1.0,
        'BULK (COOL)':    1.0,   'JERRYCAN (20L)':   5.283,
        'CASE 10/1':      2.5,
    }

    def cat(it):
        c = it.get('ul_code', '')
        if c == 'UL990': return 'def'
        if re.match(r'^UL9\d{2}$', c): return 'coolant'
        return 'lube'

    groups = {'coolant': [], 'lube': [], 'def': []}
    for it in items:
        groups[cat(it)].append(it)

    def calc_group(grp):
        total_gal = sum(it['qty'] * GALS.get(it['presentation'], 0) for it in grp)
        bbl       = total_gal / 42.0
        weight_kg = sum(it.get('weight_lbs', 0) for it in grp) / 2.20462
        value_usd = sum(it.get('value_usd', 0) for it in grp)
        boxes = sum(it['qty'] for it in grp if 'BOX' in it.get('presentation', ''))
        drums = sum(it['qty'] for it in grp if 'DRUM' in it.get('presentation', ''))
        pails = sum(it['qty'] for it in grp if 'PAIL' in it.get('presentation', ''))
        totes = sum(it['qty'] for it in grp if 'TOTE' in it.get('presentation', ''))
        bulk  = sum(it['qty'] for it in grp if 'BULK' in it.get('presentation', ''))
        jerry = sum(it['qty'] for it in grp if 'JERRYCAN' in it.get('presentation', ''))
        parts = []
        if boxes: parts.append(f"{boxes} cases")
        if drums: parts.append(f"{drums} drums")
        if pails: parts.append(f"{pails} pails")
        if totes: parts.append(f"{totes} totes")
        if bulk:  parts.append(f"{bulk} gallons (bulk)")
        if jerry: parts.append(f"{jerry} jerrycans")
        return round(bbl, 2), round(weight_kg, 3), round(value_usd, 2), (' and '.join(parts) or '0 units')

    product_lines = []
    if groups['coolant']:
        bbl, wkg, val, desc = calc_group(groups['coolant'])
        product_lines.append(('D', '3820000000', f'{bbl:.2f} BBL', desc, wkg, 'EAR99', 'N', 'N/A', val, 'N/A'))
    if groups['lube']:
        bbl, wkg, val, desc = calc_group(groups['lube'])
        product_lines.append(('D', '2710193020', f'{bbl:.2f} BBL', desc, wkg, 'EAR99', 'N', 'N/A', val, 'N/A'))
    if groups['def']:
        bbl, wkg, val, desc = calc_group(groups['def'])
        product_lines.append(('D', '3102100000', f'{bbl:.2f} BBL', desc, wkg, 'EAR99', 'N', 'N/A', val, 'N/A'))

    total_kg = sum(it.get('weight_lbs', 0) for it in items) / 2.20462

    # ── Load the formatted template and fill in variable fields only ──────────
    template_path = BASE / 'sli_template.xlsx'
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    def w(cell_ref, value):
        """Write a value to a cell (top-left of merged range) without touching formatting."""
        ws[cell_ref] = value

    # ── USPPI — always Ultrachem LLC (boxes 1, 2, 6) ─────────────────────────
    w('A3', 'Ultrachem LLC')
    w('A5', '1444 Northwest 82nd Ave.')
    w('A6', 'Doral, FL 33126')
    w('C8', '82-3520413')          # USPPI EIN (IRS) No

    # ── Freight Location — always U1Dynamics (boxes 3, 4) ────────────────────
    w('E3', 'U1Dynamics Manufacturing')
    w('E5', '4468 Genoa-Red Bluff Rd,')
    w('E6', 'Pasadena, TX. 77505')

    # ── Forwarding Agent (column J, rows 3-6) ────────────────────────────────
    w('J3', agent.get('name', ''))
    w('J4', agent.get('address1', ''))
    w('J5', agent.get('address2', ''))
    w('J6', agent.get('address3', ''))

    # ── Checkbox / select-one fields (static defaults) ───────────────────────
    # Box 7 — Related Party Indicator
    w('J8', '\u2610  Related')
    w('L8', '\u2611  Non-Related')
    # Box 9 — Routed Export Transaction
    w('J9', '\u2610  Yes')
    w('L9', '\u2611  No')
    # Box 11 — Ultimate Consignee Type
    w('E11', '\u2610  Direct Consumer')
    w('E12', '\u2610  Government Entity')
    w('E13', '\u2611  Reseller')
    w('E14', '\u2610  Other/Unknown')
    # Box 15 — Hazardous Material
    w('D18', '\u2610  Yes          \u2611  No')
    # Box 19 — TIB / Carnet
    w('L17', '\u2610  Yes')
    w('L18', '\u2611  No')

    # ── Reference # ──────────────────────────────────────────────────────────
    w('C9', ref)

    # ── Consignee (rows 11-14) ────────────────────────────────────────────────
    tax_id = consignee.get('tax_id', '')
    w('A11', consignee.get('name', ''))
    if tax_id:
        w('A12', tax_id)
        w('A13', consignee.get('address1', ''))
        w('A14', consignee.get('address2', ''))
    else:
        w('A12', consignee.get('address1', ''))
        w('A13', consignee.get('address2', ''))
        w('A14', '')

    # ── Country of Ultimate Destination ──────────────────────────────────────
    w('D17', consignee.get('country', '').upper())

    # ── Gross Weight (kg) ────────────────────────────────────────────────────
    w('B20', round(total_kg, 3))

    # ── Product Lines (rows 22-26) ───────────────────────────────────────────
    # Clear all 5 data rows first
    for r in range(22, 27):
        for col in ['A', 'B', 'D', 'E', 'F', 'G', 'H', 'I', 'L', 'M']:
            ws[f'{col}{r}'] = None

    for i, (df, hts, bbl_str, units_desc, wkg, eccn, sme, lic, val, licval) in enumerate(product_lines):
        r = 22 + i
        w(f'A{r}', df)
        w(f'B{r}', hts)    # B:C merged
        w(f'D{r}', bbl_str)
        w(f'E{r}', units_desc)
        w(f'F{r}', wkg)
        w(f'G{r}', eccn)
        w(f'H{r}', sme)
        w(f'I{r}', lic)    # I:K merged
        w(f'L{r}', val)
        w(f'M{r}', licval)

    # ── Date ─────────────────────────────────────────────────────────────────
    ws['M34'] = sli_date
    ws['M34'].number_format = 'MM/DD/YYYY'

    out = io_mod.BytesIO()
    wb.save(out)
    out.seek(0)
    cn = consignee.get('name', 'EXPORT').replace('/', '-').replace('\\', '-')
    filename = f"SLI - {cn} - {ref}.xlsx"
    return send_file(out, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    if not (BASE / 'products.json').exists():
        print("WARNING: products.json not found. Run build_catalog.py first.")
    print("Starting U1P Order Conversion Tool -> http://localhost:5000")
    app.run(debug=False, port=5000)
