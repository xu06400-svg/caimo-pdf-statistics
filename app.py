import os,re,sqlite3,io
from datetime import datetime
from pathlib import Path
from flask import Flask,request,jsonify,render_template,send_file
import fitz
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter

BASE=Path(__file__).resolve().parent
app=Flask(__name__)
DB=Path(os.environ.get('DB_PATH',str(BASE/'data.db')))
FIELDS=['date','document','line_no','material_code','item_name','specification','quantity','unit','production_no','sales_order','source_file']

def db():
 c=sqlite3.connect(DB);c.row_factory=sqlite3.Row
 c.execute('CREATE TABLE IF NOT EXISTS materials(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, document TEXT, line_no INTEGER, material_code TEXT, item_name TEXT, specification TEXT, quantity REAL, unit TEXT, production_no TEXT, sales_order TEXT, source_file TEXT)')
 c.execute('CREATE INDEX IF NOT EXISTS idx_date ON materials(date)')
 c.execute('CREATE INDEX IF NOT EXISTS idx_order ON materials(sales_order)')
 c.commit();return c

def compact(v):return re.sub(r'\s+','',v or '')
def extract(file_data,filename):
 doc=fitz.open(stream=file_data,filetype='pdf');out=[];warnings=[]
 for pi,page in enumerate(doc):
  blocks=page.get_text('blocks');words=page.get_text('words');txt=page.get_text()
  no=re.search(r'JT\d+Tra\d+',txt)
  dt=re.search(r'20\d{2}-\d{2}-\d{2}',txt)
  document=no.group() if no else '';date=dt.group() if dt else ''
  # Find row starts in the leftmost column. PDF text is re-flowed; preserve x/y layout.
  starts=[]
  for w in words:
   x0,y0,x1,y1,s,*_=w
   if x0<52 and y0>195 and re.fullmatch(r'\d{1,3}',s) and int(s)%10==0:
    if not any(abs(y0-y)<6 for y,_ in starts):starts.append((y0,int(s)))
  starts.sort()
  for idx,(y,line_no) in enumerate(starts):
   next_y=starts[idx+1][0] if idx+1<len(starts) else page.rect.height-52
   lo=y-16;hi=min(y+46,next_y-17) if idx+1<len(starts) else min(y+48,page.rect.height-48)
   if hi<lo:hi=next_y-3
   def cell(xlo,xhi):
    hits=[]
    for b in blocks:
     if b[6]!=0:continue
     bx0,by0,bx1,by1,t,*_=b
     mid=(bx0+bx1)/2
     if bx0>=xlo-2 and bx0<xhi and by1>=lo and by0<=hi:
      # exclude row above if touching boundary
      hits.append((by0,bx0,t))
    return '\n'.join(z for _,_,z in sorted(hits))
   code=compact(cell(46,94));name=compact(cell(94,136))
   rawquantity=cell(210,280)
   # PDF splits decimal places across lines, e.g. "2,842.\n470400".
   numeric=re.sub(r'[^0-9.,]','',rawquantity).replace(',','')
   qty=None
   try:
    if numeric and numeric.count('.')<=1:qty=float(numeric)
   except ValueError:pass
   pro=compact(cell(336,376));sale=compact(cell(374,418))
   # reconstruct split identifiers and numeric fractions without losing leading zeros
   material='R00107-'+code[-5:] if code and re.search(r'\d{5}$',code) else code
   if not material or not name or qty is None:
    warnings.append(f'{filename} 第{pi+1}页 行号{line_no}：字段不完整，请核对')
    continue
   out.append(dict(date=date,document=document,line_no=line_no,material_code=material,item_name=name,specification=compact(cell(135,210)),quantity=qty,unit='米',production_no=pro,sales_order=sale,source_file=filename))
 return out,warnings

@app.route('/')
def home():return render_template('index.html')
@app.route('/api/upload',methods=['POST'])
def upload():
 fs=request.files.getlist('files');added=0;warnings=[];files=[]
 c=db()
 for f in fs:
  if not f.filename.lower().endswith('.pdf'):
   warnings.append(f'{f.filename}：只支持 PDF');continue
  try:
   rows,w=extract(f.read(),os.path.basename(f.filename));warnings+=w
   if not rows:warnings.append(f'{f.filename}：未识别到领料明细');continue
   # same filename re-upload replaces prior records to prevent double counting
   c.execute('DELETE FROM materials WHERE source_file=?',(os.path.basename(f.filename),))
   c.executemany('INSERT INTO materials('+','.join(FIELDS)+') VALUES ('+','.join('?' for _ in FIELDS)+')',[[r[k] for k in FIELDS] for r in rows]);c.commit()
   added+=len(rows);files.append(f.filename)
  except Exception as ex:warnings.append(f'{f.filename}：解析失败 {str(ex)}')
 c.close();return jsonify(added=added,files=files,warnings=warnings)

def query_rows():
 f=request.args;terms=[];args=[]
 for field,param,op in [('date','start','>='),('date','end','<=')]:
  if f.get(param):terms.append(f'{field}{op}?');args.append(f[param])
 for field,param in [('sales_order','order'),('material_code','material'),('document','document')]:
  if f.get(param):terms.append(f'{field} LIKE ?');args.append('%'+f[param].strip()+'%')
 q='SELECT * FROM materials'+(' WHERE '+' AND '.join(terms) if terms else '')+' ORDER BY date DESC,document,line_no,id';c=db();rows=[dict(r) for r in c.execute(q,args)];c.close();return rows
@app.route('/api/rows')
def rows():
 r=query_rows();return jsonify(rows=r,count=len(r),total=round(sum(x['quantity'] for x in r),6),documents=len({x['document'] for x in r}),orders=len({x['sales_order'] for x in r if x['sales_order']}))
@app.route('/api/export')
def export():
 rows=query_rows();wb=Workbook();ws=wb.active;ws.title='PDF领料明细';heads=['单据日期','调拨单号','行号','料号','品名','规格','申请数量','单位','生产订单号','销售订单号','来源PDF'];ws.append(heads)
 for r in rows:ws.append([r[k] for k in FIELDS]);ws.cell(ws.max_row,7).number_format='#,##0.000000'
 for cell in ws[1]:cell.fill=PatternFill('solid',fgColor='194E74');cell.font=Font(color='FFFFFF',bold=True);cell.alignment=Alignment(horizontal='center')
 for i,width in enumerate([15,26,9,20,34,22,18,10,24,22,42],1):ws.column_dimensions[get_column_letter(i)].width=width
 ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
 ws2=wb.create_sheet('按订单汇总');ws2.append(['销售订单号','领料笔数','领料合计（米）'])
 totals={}
 for r in rows:
  key=r['sales_order'] or '(未填写订单号)';v=totals.setdefault(key,[0,0]);v[0]+=1;v[1]+=r['quantity']
 for k,v in sorted(totals.items()):ws2.append([k,v[0],round(v[1],6)])
 for c in ws2[1]:c.fill=PatternFill('solid',fgColor='194E74');c.font=Font(bold=True,color='FFFFFF')
 for col,w in zip('ABC',[26,18,25]):ws2.column_dimensions[col].width=w
 ws2.freeze_panes='A2';ws2.auto_filter.ref=ws2.dimensions
 b=io.BytesIO();wb.save(b);b.seek(0)
 return send_file(b,as_attachment=True,download_name='彩膜PDF领料统计.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/api/files')
def files():
 c=db();a=[dict(r) for r in c.execute('SELECT source_file,COUNT(*) AS count FROM materials GROUP BY source_file ORDER BY source_file')];c.close();return jsonify(files=a)
@app.route('/api/files/delete',methods=['POST'])
def delete_file():
 name=(request.json or {}).get('name');c=db();c.execute('DELETE FROM materials WHERE source_file=?',(name,));c.commit();c.close();return jsonify(ok=True)
if __name__=='__main__':
 print('浏览器打开 http://127.0.0.1:5050')
 app.run(host='127.0.0.1',port=5050,debug=False)
