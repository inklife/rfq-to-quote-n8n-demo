import base64
import hashlib
import importlib.metadata
import json
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from rfq import process

ROOT = Path(__file__).parent
ROOT.joinpath('sample').mkdir(exist_ok=True)
pdf_file = ROOT / 'sample' / 'rfq-input.pdf'
styles = getSampleStyleSheet()
lines = [
 ['1', 'SW-A01', 'Wall switch', '2', 'EA'],
 ['2', 'CL-B02', 'Cable clip', '0.145', 'EA'],
 ['3', 'CB-C03', 'Cable metre', '1.005', 'M'],
 ['4', 'SW-A01', 'Main breaker', '1', 'EA'],
 ['5', 'FZ-D04', 'Fuse holder', '3', 'EA'],
 ['6', 'NO-SUCH', 'Unknown item', '1', 'EA'],
 ['7', 'SW-A01', 'Wall switch', '1', 'BOX'],
 ['8', 'CL-B02', 'Cable clip', '0', 'EA'],
]
doc = SimpleDocTemplate(str(pdf_file), pagesize=A4, rightMargin=35, leftMargin=35,
                        topMargin=35, bottomMargin=35, title='Synthetic RFQ input', invariant=1)
table = Table([['Line', 'Product code', 'Description', 'Quantity', 'Unit']] + lines,
              colWidths=[35, 95, 225, 65, 45], rowHeights=[30] + [30]*len(lines))
table.setStyle(TableStyle([
 ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
 ('FONTSIZE', (0,0), (-1,-1), 10),
 ('BACKGROUND', (0,0),(-1,0),colors.HexColor('#153D4C')),
 ('TEXTCOLOR',(0,0),(-1,0),colors.white),
 ('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#899DA6')),
 ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
 ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F3F6F7')]),
]))
doc.build([Paragraph('RFQ / ELECTRICAL PARTS',styles['Title']),
           Paragraph('Synthetic input for workflow verification · 15 September 2026',styles['Normal']),
           Spacer(1,24),table,Spacer(1,20),
           Paragraph('This is a demonstration of one fixed, bordered text-PDF layout. Product data are invented. No customer information is used.',styles['Normal'])])
catalogue = [
 {'code':'SW-A01','description':'Wall switch','description_ar':'مفتاح حائط','unit':'EA'},
 {'code':'CL-B02','description':'Cable clip','description_ar':'مشبك كابل','unit':'EA'},
 {'code':'CB-C03','description':'Cable metre','description_ar':'كابل بالمتر','unit':'M'},
 {'code':'FZ-D04','description':'Fuse holder','description_ar':'حامل مصهر','unit':'EA'},
]
price_list = {'version':'SYNTHETIC-AED-2026-09-15','currency':'AED','items':[
 {'code':'SW-A01','unit':'EA','price':'12.50'},
 {'code':'CL-B02','unit':'EA','price':'1.00'},
 {'code':'CB-C03','unit':'M','price':'1.00'},
]}
payload = {'pdf_base64':base64.b64encode(pdf_file.read_bytes()).decode(), 'catalogue':catalogue,'price_list':price_list}
ROOT.joinpath('sample','request.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
quote = process(payload)
ROOT.joinpath('sample','quote-draft.xlsx').write_bytes(base64.b64decode(quote.pop('xlsx_base64')))
ROOT.joinpath('sample','result.json').write_text(json.dumps(quote,ensure_ascii=False,indent=2),encoding='utf-8')
fixture_code = 'return [{ json: ' + json.dumps(payload,ensure_ascii=False) + ' }];'
workflow = {
 'name':'RFQ proof - PDF to AED Excel draft', 'active':False, 'settings':{'executionOrder':'v1'},
 'nodes':[
  {'id':'start','name':'Run synthetic example','type':'n8n-nodes-base.manualTrigger','typeVersion':1,'position':[0,0],'parameters':{}},
  {'id':'input','name':'Synthetic PDF and approved data','type':'n8n-nodes-base.code','typeVersion':2,'position':[250,0],'parameters':{'jsCode':fixture_code}},
  {'id':'quote','name':'Extract match and render','type':'n8n-nodes-base.httpRequest','typeVersion':4.2,'position':[500,0],'parameters':{
   'method':'POST','url':'http://127.0.0.1:8795/quote','sendBody':True,'specifyBody':'json','jsonBody':'={{ $json }}','options':{'timeout':20000}}},
  {'id':'output','name':'Excel file and review summary','type':'n8n-nodes-base.code','typeVersion':2,'position':[750,0],'parameters':{'jsCode':"const {xlsx_base64, ...summary} = $input.first().json; return [{json: summary, binary: {data: {data: xlsx_base64, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName: 'quote-draft.xlsx'}}}];"}},
 ],
 'connections':{
  'Run synthetic example':{'main':[[{'node':'Synthetic PDF and approved data','type':'main','index':0}]]},
  'Synthetic PDF and approved data':{'main':[[{'node':'Extract match and render','type':'main','index':0}]]},
  'Extract match and render':{'main':[[{'node':'Excel file and review summary','type':'main','index':0}]]},
 }
}
ROOT.joinpath('workflow.json').write_text(json.dumps(workflow,ensure_ascii=False,indent=2),encoding='utf-8')
ROOT.joinpath('requirements.txt').write_text('\n'.join(f'{p}=={importlib.metadata.version(p)}' for p in ['pdfplumber','xlsxwriter','reportlab'])+'\n',encoding='utf-8')
print(json.dumps({'rows':len(quote['rows']),'review':quote['review_count'],'subtotal':quote['matched_subtotal'],'complete_total':quote['complete_total'],'pdf_sha256':quote['pdf_sha256']},indent=2))
