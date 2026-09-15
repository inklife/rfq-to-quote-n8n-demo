"""A local, deterministic PDF RFQ demonstration. No network or LLM calls."""
from __future__ import annotations

import base64
import hashlib
import io
import json
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pdfplumber
import xlsxwriter

CENT = Decimal('0.01')
MAX_PDF = 2_000_000
MAX_ROWS = 100
HEADERS = ['Line', 'Product code', 'Description', 'Quantity', 'Unit']


def normal(value):
    if not isinstance(value, str):
        return ''
    return ' '.join(unicodedata.normalize('NFKC', value).split()).casefold()


def decimal_value(value, *, allow_zero=False):
    if not isinstance(value, str) or not re.fullmatch(r'[0-9]{1,8}(?:\.[0-9]{1,6})?', value):
        raise ValueError('Use an unsigned decimal string, at most 8 integer and 6 decimal digits')
    try:
        number = Decimal(value)
    except InvalidOperation as exc:
        raise ValueError('Invalid decimal') from exc
    if number < 0 or (not allow_zero and number == 0):
        raise ValueError('Value must be positive')
    return number


def extract_pdf(pdf):
    if not isinstance(pdf, bytes) or len(pdf) > MAX_PDF or not pdf.startswith(b'%PDF-'):
        raise ValueError('Expected a PDF of at most 2 MB')
    records = []
    source_issues = []
    seen_lines = set()
    scanned_rows = 0
    try:
        with pdfplumber.open(io.BytesIO(pdf)) as document:
            if not 1 <= len(document.pages) <= 5:
                raise ValueError('Demo accepts 1-5 pages')
            for page_number, page in enumerate(document.pages, 1):
                tables = page.extract_tables()
                if len(tables) != 1:
                    raise ValueError(f'Page {page_number}: expected exactly one bordered RFQ table')
                table = tables[0]
                if not table or [normal(x) for x in table[0]] != [normal(x) for x in HEADERS]:
                    raise ValueError(f'Page {page_number}: unsupported table header/layout')
                for table_row, cells in enumerate(table[1:], 2):
                    scanned_rows += 1
                    if scanned_rows > MAX_ROWS:
                        raise ValueError(f'Demo accepts at most {MAX_ROWS} RFQ rows including exceptions')
                    source = f'PDF p{page_number}, table row {table_row}'
                    if len(cells) != 5:
                        source_issues.append({'source': source, 'reason': 'Malformed table row', 'cells': cells})
                        continue
                    line, code, description, quantity, unit = [str(x or '').strip() for x in cells]
                    if not line.isdigit() or line in seen_lines:
                        source_issues.append({'source': source, 'reason': 'Missing/duplicate line number', 'cells': cells})
                        continue
                    seen_lines.add(line)
                    records.append(dict(line=line, code=code, description=description,
                                        quantity=quantity, unit=unit, source=source))
                    if len(records) > MAX_ROWS:
                        raise ValueError(f'Demo accepts at most {MAX_ROWS} RFQ rows')
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError('PDF could not be read; no quote was produced') from exc
    if not records:
        raise ValueError('No item rows found; no quote was produced')
    return records, source_issues


def prepare_quote(records, catalogue, price_list, source_issues=None):
    if not isinstance(records, list) or not 1 <= len(records) <= MAX_ROWS:
        raise ValueError('Expected 1-100 RFQ rows')
    if not isinstance(catalogue, list) or not 1 <= len(catalogue) <= 1000:
        raise ValueError('Expected 1-1000 catalogue rows')
    if not isinstance(price_list, dict) or not isinstance(price_list.get('items'), list):
        raise ValueError('Expected a versioned price list')
    if not isinstance(price_list.get('version'), str) or not price_list['version'].strip():
        raise ValueError('Price-list version is required')
    if price_list.get('currency') != 'AED':
        raise ValueError('This demonstration is explicitly limited to AED')
    if not 1 <= len(price_list['items']) <= 1000:
        raise ValueError('Expected 1-1000 price-list rows')
    by_code, by_price = defaultdict(list), defaultdict(list)
    for item in catalogue:
        if not isinstance(item, dict):
            raise ValueError('Catalogue rows must be objects')
        by_code[normal(item.get('code'))].append(item)
    for item in price_list['items']:
        if not isinstance(item, dict):
            raise ValueError('Price-list rows must be objects')
        by_price[normal(item.get('code'))].append(item)
    rows, subtotal = [], Decimal(0)
    for index, record in enumerate(records, 1):
        if not isinstance(record, dict):
            raise ValueError('RFQ rows must be objects')
        row = {key: record.get(key, '') for key in ['line', 'code', 'description', 'quantity', 'unit', 'source']}
        row.update(status='REVIEW', reasons=[], price=None, amount=None, description_ar='')
        if any(not isinstance(row[key], str) for key in ['code', 'description', 'quantity', 'unit']):
            row['reasons'].append('RFQ fields must be strings')
        try:
            quantity = decimal_value(record.get('quantity'))
        except ValueError:
            quantity = None
            row['reasons'].append('Invalid or non-positive quantity')
        matches = by_code.get(normal(record.get('code')), []) if normal(record.get('code')) else []
        if len(matches) != 1:
            row['reasons'].append('Unknown code' if not matches else 'Duplicate catalogue code')
        else:
            item = matches[0]
            row['description_ar'] = item.get('description_ar', '')
            if not normal(record.get('unit')) or normal(record.get('unit')) != normal(item.get('unit')):
                row['reasons'].append('Unit missing or does not match catalogue')
            if not normal(record.get('description')) or normal(record.get('description')) != normal(item.get('description')):
                row['reasons'].append('Description differs from catalogue; human approval required')
            prices = by_price.get(normal(item.get('code')), [])
            if len(prices) != 1:
                row['reasons'].append('Missing approved price' if not prices else 'Duplicate price-list code')
            else:
                price_record = prices[0]
                if normal(price_record.get('unit')) != normal(item.get('unit')):
                    row['reasons'].append('Price-list unit does not match catalogue')
                try:
                    price = decimal_value(price_record.get('price'))
                except ValueError:
                    price = None
                    row['reasons'].append('Invalid or non-positive approved price')
                if not row['reasons']:
                    amount = (quantity * price).quantize(CENT, rounding=ROUND_HALF_UP)
                    if amount >= Decimal('1000000000'):
                        row['reasons'].append('Amount exceeds the demonstration limit')
                    else:
                        row.update(status='MATCHED', price=str(price), amount=f'{amount:.2f}')
                        subtotal += amount
        rows.append(row)
    issues = source_issues or []
    review_count = sum(row['status'] == 'REVIEW' for row in rows) + len(issues)
    return dict(status='DRAFT', currency='AED', price_version=price_list['version'], rows=rows,
                source_issues=issues, review_count=review_count, matched_subtotal=f'{subtotal:.2f}',
                complete_total=None if review_count else f'{subtotal:.2f}',
                policy='Exact normalized code + description + unit; no substitutions. Line ROUND_HALF_UP to AED 0.01. No tax, discounts or currency conversion.')


def workbook_bytes(quote):
    stream = io.BytesIO()
    book = xlsxwriter.Workbook(stream, {'in_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False})
    book.set_properties({'title': 'RFQ Review Draft', 'comments': 'Synthetic demonstration; regenerated snapshot, not a price source.'})
    base = {'font_name': 'Arial', 'font_size': 10, 'valign': 'vcenter'}
    fmt = lambda props: book.add_format({**base, **props})
    title = fmt({'bold': True, 'font_size': 22, 'font_color': '#153D4C'})
    note = fmt({'font_color': '#50616A', 'text_wrap': True})
    head = fmt({'bold': True, 'bg_color': '#153D4C', 'font_color': '#FFFFFF', 'text_wrap': True})
    text = fmt({'text_wrap': True, 'bottom': 1, 'bottom_color': '#E3E8EA'})
    review = fmt({'text_wrap': True, 'bg_color': '#FFF2D6', 'font_color': '#754500', 'bottom': 1, 'bottom_color': '#E3E8EA'})
    number = fmt({'num_format': '0.######', 'bottom': 1, 'bottom_color': '#E3E8EA'})
    money = fmt({'num_format': '#,##0.00', 'bottom': 1, 'bottom_color': '#E3E8EA'})
    label = fmt({'bold': True, 'font_color': '#153D4C'})
    for sheet_name, rtl in [('Quote EN', False), ('عرض السعر AR', True)]:
        sheet = book.add_worksheet(sheet_name)
        if rtl:
            sheet.right_to_left()
        sheet.hide_gridlines(2)
        sheet.set_landscape()
        sheet.set_paper(9)
        sheet.fit_to_pages(1, 0)
        sheet.set_margins(0.25, 0.25, 0.4, 0.4)
        sheet.set_column('A:A', 7)
        sheet.set_column('B:B', 16)
        sheet.set_column('C:C', 30)
        sheet.set_column('D:D', 12)
        sheet.set_column('E:E', 9)
        sheet.set_column('F:G', 14)
        sheet.set_column('H:H', 15)
        sheet.set_column('I:I', 44)
        sheet.set_column('J:J', 25)
        sheet.set_row(0, 38)
        sheet.merge_range('A1:J1', 'مسودة عرض السعر · AED' if rtl else 'QUOTATION / REVIEW DRAFT', title)
        sheet.merge_range('A2:J2', 'SYNTHETIC DEMONSTRATION · All rows require final human approval · Price version: ' + quote['price_version'], note)
        sheet.set_row(1, 30)
        sheet.merge_range('A3:J3', 'الأسعار من القائمة المعتمدة فقط؛ المراجعة البشرية مطلوبة.' if rtl else 'Approved prices only. A review row withholds the complete quotation total. Regenerate after input changes.', note)
        headers = ['#', 'رمز المنتج', 'الوصف', 'الكمية', 'الوحدة', 'السعر AED', 'المبلغ AED', 'الحالة', 'سبب المراجعة', 'المصدر'] if rtl else ['#', 'Product code', 'Description', 'Quantity', 'Unit', 'Price AED', 'Amount AED', 'Status', 'Review reason', 'Source']
        sheet.write_row(4, 0, headers, head)
        sheet.set_row(4, 32)
        for r, row in enumerate(quote['rows'], 5):
            is_review = row['status'] == 'REVIEW'
            style = review if is_review else text
            sheet.set_row(r, 46)
            for c, val in [(0, row['line']), (1, row['code']), (2, (row['description_ar'] or row['description']) if rtl and not is_review else row['description']), (4, row['unit']), (7, row['status']), (8, '; '.join(row['reasons'])), (9, row['source'])]:
                sheet.write_string(r, c, str(val), style)
            try:
                sheet.write_number(r, 3, float(decimal_value(row['quantity'])), number)
            except ValueError:
                sheet.write_string(r, 3, str(row['quantity']), style)
            if not is_review:
                sheet.write_number(r, 5, float(Decimal(row['price'])), number)
                sheet.write_formula(r, 6, f'=ROUND(D{r+1}*F{r+1},2)', money, float(Decimal(row['amount'])))
        end = 5 + len(quote['rows'])
        sheet.autofilter(4, 0, end-1, 9)
        sheet.freeze_panes(5, 2)
        sheet.merge_range(end+1, 3, end+1, 5, 'المجموع المطابق' if rtl else 'Matched subtotal', label)
        sheet.write_formula(end+1, 6, f'=SUM(G6:G{end})', money, float(Decimal(quote['matched_subtotal'])))
        sheet.merge_range(end+2, 3, end+2, 5, 'عناصر للمراجعة' if rtl else 'Review items', label)
        sheet.write_formula(end+2, 6, f'=COUNTIF(H6:H{end},"REVIEW")+COUNTA(Exceptions!A6:A105)', number, quote['review_count'])
        sheet.merge_range(end+3, 3, end+3, 5, 'المجموع الكامل' if rtl else 'Complete total', label)
        sheet.write_formula(end+3, 6, f'=IF(G{end+3}>0,"",G{end+2})', money, '' if quote['complete_total'] is None else float(Decimal(quote['complete_total'])))
        sheet.merge_range(end+5, 0, end+5, 9, 'DRAFT ONLY / مسودة فقط · No tax, discount, pack conversion or automatic substitution. Review reasons and source references stay in English.', note)
        sheet.set_row(end+5, 30)
        sheet.print_area(0, 0, end+5, 9)
        sheet.repeat_rows(0, 4)
    ex = book.add_worksheet('Exceptions')
    ex.hide_gridlines(2)
    ex.set_column('A:A', 28)
    ex.set_column('B:B', 45)
    ex.set_column('C:C', 70)
    ex.merge_range('A1:C1', 'PDF EXTRACTION EXCEPTIONS', title)
    ex.set_row(0, 36)
    ex.merge_range('A2:C2', 'Non-item or malformed rows block a complete total. These are never silently dropped.', note)
    ex.set_row(1, 30)
    ex.write_row(4, 0, ['Source', 'Reason', 'Raw table cells'], head)
    ex.set_row(4, 28)
    for r, issue in enumerate(quote['source_issues'], 5):
        ex.write_string(r, 0, issue['source'], review)
        ex.write_string(r, 1, issue['reason'], review)
        ex.write_string(r, 2, json.dumps(issue['cells'], ensure_ascii=False), review)
        ex.set_row(r, 44)
    if not quote['source_issues']:
        ex.merge_range('A3:C3', 'No extraction exceptions in this sample. Item-level review flags remain on the quotation sheets.', note)
    ex.freeze_panes(5, 0)
    ex.set_landscape()
    ex.fit_to_pages(1, 0)
    book.close()
    return stream.getvalue()


def process(payload):
    if not isinstance(payload, dict):
        raise ValueError('Expected a JSON object')
    try:
        pdf = base64.b64decode(payload['pdf_base64'], validate=True)
    except (KeyError, ValueError, TypeError) as exc:
        raise ValueError('pdf_base64 must contain valid base64') from exc
    records, issues = extract_pdf(pdf)
    quote = prepare_quote(records, payload.get('catalogue'), payload.get('price_list'), issues)
    quote['pdf_sha256'] = hashlib.sha256(pdf).hexdigest()
    quote['xlsx_base64'] = base64.b64encode(workbook_bytes(quote)).decode('ascii')
    return quote


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.respond(200 if self.path == '/health' else 404, {'status': 'ok' if self.path == '/health' else 'not found'})

    def do_POST(self):
        if self.path != '/quote':
            return self.respond(404, {'error': 'not found'})
        try:
            length = int(self.headers.get('Content-Length', 0))
            if not 0 < length <= 4_000_000:
                raise ValueError('Body must be between 1 byte and 4 MB')
            result = process(json.loads(self.rfile.read(length)))
            self.respond(200, result)
        except (ValueError, TypeError, KeyError) as exc:
            self.respond(400, {'error': str(exc)})

    def respond(self, code, data):
        raw = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def log_message(self, *_):
        pass


if __name__ == '__main__':
    print('Synthetic demo helper: http://127.0.0.1:8795', flush=True)
    ThreadingHTTPServer(('127.0.0.1', 8795), Handler).serve_forever()
