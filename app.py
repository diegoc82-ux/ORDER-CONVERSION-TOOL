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


@app.route('/api/generate-sli', methods=['POST'])
def generate_sli():
    import re
    import openpyxl
    from openpyxl.styles import Font, Alignment
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

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'SLI'

    col_widths = {'A':30,'B':24,'C':14,'D':18,'E':24,'F':16,'G':16,'H':9,'I':24,'J':24,'K':14,'L':22,'M':22}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    def s(row, col, val, bold=False, sz=10):
        c = ws.cell(row=row, column=col, value=val)
        c.font = Font(name='Arial', bold=bold, size=sz)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        return c

    s(1,  1, "                                          SHIPPER'S LETTER OF INSTRUCTIONS (SLI)            ", bold=True, sz=11)
    s(2,  1, '1. USPPI Name:', bold=True)
    s(2,  5, '3. Freight Location Company Name:', bold=True)
    s(2, 10, '5. Forwarding Agent:   ', bold=True)
    s(3,  1, 'Ultrachem LLC ')
    s(3,  5, 'U1Dynamics Manufacturing')
    s(3, 10, agent.get('name', ''))
    s(4,  1, '2. USPPI Address Including Zip Code:', bold=True)
    s(4,  5, '4. Freight Location Address (if not box #2):', bold=True)
    s(5,  1, '1444 Northwest 82 nd Ave.  ')
    s(5,  5, '4468 Genoa-Red Bluff Rd,')
    s(5, 10, agent.get('address1', ''))
    s(6,  1, 'Doral, FL 33126')
    s(6,  5, 'Pasadena, TX. 77505')
    s(6, 10, agent.get('address2', ''))
    s(7, 10, agent.get('address3', ''))
    s(8,  1, '6.  USPPI EIN (IRS) No:                                  ', bold=True)
    s(8,  3, '82-3520413')
    s(8,  5, '7.   Related Party Indicator (select one):', bold=True)
    s(9,  1, '8.  USPPI Reference#: ', bold=True)
    s(9,  3, ref)
    s(9,  5, '9.   Routed Export Transaction (select one):', bold=True)
    s(10, 1, '10.  Ultimate Consignee Name & Address: ', bold=True)
    s(10, 5, '11. Ultimate Consignee Type (select one):', bold=True)
    s(10,10, '12. Intermediate Consignee Name & Address:', bold=True)
    s(11, 1, consignee.get('name', ''))
    tax_id = consignee.get('tax_id', '')
    if tax_id:
        s(12, 1, tax_id)
        s(13, 1, consignee.get('address1', ''))
        s(14, 1, consignee.get('address2', ''))
    else:
        s(12, 1, consignee.get('address1', ''))
        s(13, 1, consignee.get('address2', ''))
    s(16, 1, '13. State of Origin: ', bold=True)
    s(16, 4, 'TEXAS ')
    s(16, 7, '16. In-Bond Code:', bold=True)
    s(16,12, '19. TIB / Carnet?', bold=True)
    s(17, 1, '14. Country of Ultimate Destination:', bold=True)
    s(17, 4, consignee.get('country', '').upper())
    s(17, 7, '17. Entry Number:', bold=True)
    s(18, 1, '15. Hazardous Material: ', bold=True)
    s(18, 7, '18. FTZ Identifier:', bold=True)
    s(19, 1, 'INSTRUCTIONS TO FORWARDER:                                                                                                                                                                              ')
    s(20, 1, '20. Gross Weight (kilos)', bold=True)
    s(20, 2, round(total_kg, 3))
    s(20, 3, '21. SOLAS  Certification', bold=True)
    s(20, 4, 'By checking the Box 21 certification, I am certifying that the full shipment weight shown in box 20 is the Certified Gross Weight which may be added to the container tare weight and used as the Verified Gross Mass (VGM) under the Method 2 of the SOLAS VGM regulation which becomes effective July 1, 2016.')
    s(21, 1, '22.          Domestic  or     Foreign (D/F)', bold=True)
    s(21, 2, '23.                                    Schedule B / HTS Number and Commercial  Commodity Description                                    For Vehicles: VIN/Year, Make, Model and Vehicle Title Number are required', bold=True)
    s(21, 4, '24.             Quantity in Schedule B / HTS Units ', bold=True)
    s(21, 5, '25.                       DDTC Quantity and DDTC Unit of Measure', bold=True)
    s(21, 6, '26.              Shipping Weight             (in Kilos)', bold=True)
    s(21, 7, '27.                   ECCN, EAR99 or USML Category No.  ', bold=True)
    s(21, 8, '28 . S    M    E  (Y/ N)', bold=True)
    s(21, 9, '29.                                            Export License No., License Exception Symbol,                         DDTC Exemption No.,           DDTC ACM No.                              or NLR     ', bold=True)
    s(21,12, '30.              Value at the Port of Export                (US Dollars)', bold=True)
    s(21,13, '31.                License Value by item (if applicable)              (US Dollars)', bold=True)

    for i, (df, hts, bbl_str, units_desc, wkg, eccn, sme, lic, val, licval) in enumerate(product_lines):
        r = 22 + i
        s(r, 1, df); s(r, 2, hts); s(r, 4, bbl_str); s(r, 5, units_desc)
        s(r, 6, wkg); s(r, 7, eccn); s(r, 8, sme); s(r, 9, lic)
        s(r,12, val); s(r,13, licval)

    s(27, 1, '32. DDTC Applicant Registration Number: ', bold=True)
    s(27, 7, '33. Eligible Party Certification:', bold=True)
    s(28, 1, '34.')
    s(28, 2, 'Check here if there are any remaining non-licensable Schedule B / HTS Numbers that are valued $2500.00 or less and that do not otherwise require AES filing.')
    s(29, 1, '35.')
    s(29, 2, 'Check here if the USPPI authorizes the above named forwarder to act as its true and lawful agent for purposes of preparing and filing the Electronic Export Information ("EEI") in accordance with the laws and regulations of the United States. ')
    s(30, 1, '35 a')
    s(30, 2, 'Shipper grants carrier consent to screen cargo as may be required by the Transportation Security Administration.')
    s(31, 1, '36.  I certify that the statements made and all information contained herein are true and correct. I understand that civil and criminal penalties, including forfeiture and sale, may be imposed for making false and fraudulent statements herein., failing to provide the requested information or for violation of U.S. laws on exportation (13 U.S.C. Sec . 305:  22 U.S.C. Sec. 401, 18 U.S.C. Sec 1001, 50 U.S.C. app. 2410).  ')
    s(32, 1, '37. USPPI E-mail Address: ', bold=True)
    s(32, 4, 'dcastro@ultra1plus.com')
    s(32, 7, '38. USPPI Telephone No.: ', bold=True)
    s(32,11, '305-988-7624')
    s(33, 1, '39.  Printed Name of Duly authorized officer or employee: ', bold=True)
    s(33, 7, 'Diego Castro')
    s(34, 1, '40.  Signature: ', bold=True)
    s(34, 7, '41. Title: ', bold=True)
    s(34, 8, 'COO')
    s(34,12, '42. Date: ', bold=True)
    dc = ws.cell(row=34, column=13, value=sli_date)
    dc.number_format = 'MM/DD/YYYY'
    dc.font = Font(name='Arial', size=10)
    dc.alignment = Alignment(wrap_text=True, vertical='top')
    s(35, 1, '43.          Check here to validate Electronic Signature.  Electronic signatures must be typed in all capital letters in Box 39 in order to be valid. ')

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
