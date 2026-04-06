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
                                  if re.match(r'^(UL|TR)\d+$', w['text'], re.I) and w['x0'] < 90]
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


if __name__ == '__main__':
    if not (BASE / 'products.json').exists():
        print("WARNING: products.json not found. Run build_catalog.py first.")
    print("Starting U1P Order Conversion Tool -> http://localhost:5000")
    app.run(debug=False, port=5000)
