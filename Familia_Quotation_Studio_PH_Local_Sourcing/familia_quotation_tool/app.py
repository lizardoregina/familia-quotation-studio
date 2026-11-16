import csv, io, os, re, sqlite3, datetime, copy, secrets, hmac, json, time
from pathlib import Path
from urllib.parse import urlparse
import requests
from flask import Flask, request, jsonify, render_template, send_file, session, redirect, url_for
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange
from werkzeug.utils import secure_filename

ROOT=Path(__file__).resolve().parent
DB=Path(os.environ.get('FAMILIA_DB', str(ROOT/'familia.db')))
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=8*1024*1024
app.secret_key=os.environ.get('SECRET_KEY', 'local-development-only-key')
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax', SESSION_COOKIE_SECURE=os.environ.get('RENDER','').lower()=='true' or os.environ.get('HTTPS_ONLY')=='1')
SUPABASE_URL=os.environ.get('SUPABASE_URL','').rstrip('/')
SUPABASE_SECRET_KEY=os.environ.get('SUPABASE_SECRET_KEY','')
CLOUD=bool(SUPABASE_URL and SUPABASE_SECRET_KEY)
if bool(SUPABASE_URL)!=bool(SUPABASE_SECRET_KEY):
    raise RuntimeError('Set SUPABASE_URL and SUPABASE_SECRET_KEY together.')
if os.environ.get('RENDER') and not CLOUD:
    raise RuntimeError('Cloud deployment requires SUPABASE_URL and SUPABASE_SECRET_KEY; local SQLite is not persistent on free Render.')
if os.environ.get('RENDER') and (not os.environ.get('APP_PASSWORD') or not os.environ.get('SECRET_KEY')):
    raise RuntimeError('Cloud deployment requires APP_PASSWORD and SECRET_KEY.')

def cloud_request(method, table, *, params=None, payload=None, prefer='return=representation'):
    res=requests.request(method,f'{SUPABASE_URL}/rest/v1/{table}',params=params,json=payload,
        headers={'apikey':SUPABASE_SECRET_KEY,
                 'Content-Type':'application/json','Prefer':prefer},timeout=20)
    if not res.ok: raise RuntimeError(f'Supabase storage error: HTTP {res.status_code}. Check SQL table setup and API credentials.')
    return res.json() if res.content else []

@app.before_request
def require_login():
    if request.path in ('/login','/health') or request.path.startswith('/static/'):
        return None
    if os.environ.get('APP_PASSWORD') and not session.get('logged_in'):
        if request.path.startswith('/api/'):
            return jsonify(error='Session expired. Please sign in again.'),401
        return redirect(url_for('login'))
    if request.method in ('POST','PUT','PATCH','DELETE') and request.headers.get('Origin'):
        from urllib.parse import urlsplit
        origin=urlsplit(request.headers['Origin'])
        if origin.netloc != request.host:
            return jsonify(error='Cross-site request blocked.'),403

@app.get('/health')
def health(): return jsonify(status='ok',storage='supabase' if CLOUD else 'local')

@app.route('/login',methods=['GET','POST'])
def login():
    if not os.environ.get('APP_PASSWORD'):
        return redirect('/')
    error=False
    if request.method=='POST':
        supplied=request.form.get('password','')
        if hmac.compare_digest(supplied,os.environ['APP_PASSWORD']):
            session.clear();session['logged_in']=True
            return redirect('/')
        error=True
        time.sleep(0.5)
    return render_template('login.html',error=error)

@app.post('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


def connect():
    con=sqlite3.connect(DB)
    con.row_factory=sqlite3.Row
    return con

def init_db():
    DB.parent.mkdir(parents=True,exist_ok=True)
    with connect() as c:
        c.execute('CREATE TABLE IF NOT EXISTS products (id INTEGER PRIMARY KEY AUTOINCREMENT, sku TEXT, name TEXT NOT NULL, uom TEXT DEFAULT "pc", cost REAL NOT NULL, supplier TEXT DEFAULT "", source_url TEXT DEFAULT "")')
if not CLOUD:
    init_db()

def number(x, default=None):
    try:
        if x is None or str(x).strip()=='': return default
        v=float(str(x).replace(',','').replace('₱','').strip())
        if not -1e12<v<1e12: return default
        return v
    except (ValueError,TypeError): return default

def valid_product(x):
    name=str(x.get('name') or x.get('product') or x.get('description') or '').strip()
    cost=number(x.get('cost', x.get('price')))
    if not name or cost is None or cost<0: raise ValueError('Product name and a valid nonnegative supplier cost are required.')
    return (str(x.get('sku','')).strip(),name,str(x.get('uom') or 'pc').strip(),cost,str(x.get('supplier') or '').strip(),str(x.get('source_url') or '').strip())

@app.get('/')
def index(): return render_template('index.html')

@app.get('/api/products')
def products():
    if CLOUD:
        rows=cloud_request('GET','products',params={'select':'*','order':'name.asc'})
    else:
        with connect() as c: rows=[dict(row) for row in c.execute('SELECT * FROM products ORDER BY name COLLATE NOCASE')]
    return jsonify(rows)

@app.post('/api/products')
def add_product():
    try: p=valid_product(request.get_json(force=True))
    except ValueError as e: return jsonify(error=str(e)),400
    if CLOUD:
        rows=cloud_request('POST','products',payload=dict(zip(('sku','name','uom','cost','supplier','source_url'),p)))
        product_id=rows[0]['id']
    else:
        with connect() as c:
            cur=c.execute('INSERT INTO products (sku,name,uom,cost,supplier,source_url) VALUES (?,?,?,?,?,?)',p)
            product_id=cur.lastrowid
    return jsonify(id=product_id),201

@app.delete('/api/products/<int:product_id>')
def delete_product(product_id):
    if CLOUD:
        cloud_request('DELETE','products',params={'id':f'eq.{product_id}'},prefer='return=minimal')
    else:
        with connect() as c: c.execute('DELETE FROM products WHERE id=?',(product_id,))
    return jsonify(ok=True)

@app.post('/api/import')
def import_products():
    f=request.files.get('file')
    if not f: return jsonify(error='Select an Excel (.xlsx) or CSV file.'),400
    ext=Path(secure_filename(f.filename or '')).suffix.lower()
    if ext not in ('.csv','.xlsx'): return jsonify(error='Only .csv and .xlsx are supported.'),400
    try:
        if ext=='.csv':
            contents=f.read().decode('utf-8-sig')
            records=list(csv.DictReader(io.StringIO(contents)))
        else:
            wb=load_workbook(io.BytesIO(f.read()),read_only=True,data_only=True)
            it=wb.active.iter_rows(values_only=True)
            headers=[str(v or '').strip() for v in next(it)]
            records=[dict(zip(headers,row)) for row in it if any(v is not None for v in row)]
        aliases={'product name':'name','product':'name','description':'name','item':'name','unit cost':'cost','supplier price':'cost','cost price':'cost','price':'cost','unit':'uom','unit of measure':'uom','vendor':'supplier','link':'source_url','url':'source_url'}
        good=[]; bad=0
        for rec in records:
            mapped={aliases.get(str(k).lower().strip(),str(k).lower().strip()):v for k,v in rec.items() if k is not None}
            try: good.append(valid_product(mapped))
            except ValueError: bad+=1
        if CLOUD:
            for offset in range(0,len(good),200):
                batch=[dict(zip(('sku','name','uom','cost','supplier','source_url'),row)) for row in good[offset:offset+200]]
                cloud_request('POST','products',payload=batch,prefer='return=minimal')
        else:
            with connect() as c:
                c.executemany('INSERT INTO products (sku,name,uom,cost,supplier,source_url) VALUES (?,?,?,?,?,?)',good)
        return jsonify(imported=len(good),skipped=bad)
    except Exception as e:
        return jsonify(error='Could not read file. Use a header row: name, cost, uom, sku, supplier, source_url.'),400

# Philippine storefront indicators are not proof that stock ships from the Philippines.
# Marketplace sellers may still dispatch overseas; staff must verify ship-from origin.
PH_MARKET_DOMAINS=('shopee.ph','lazada.com.ph','zalora.com.ph','tiktok.com','shop.tiktok.com')
PH_LOCAL_STORE_DOMAINS=('acehardware.ph','wilcon.com.ph','nationalbookstore.com','abenson.com','allhome.com.ph','truevalue.com.ph','smappliance.com','octagon.com.ph','datablitz.com.ph','hardwaremaxx.com.ph')
PH_MARKET_NAMES=('shopee','lazada','tiktok shop','zalora')

def philippine_listing(item):
    from urllib.parse import urlparse
    price_text=str(item.get('price') or '').strip()
    alternative=item.get('alternative_price') or {}
    # Never mislabel foreign-currency prices as pesos.
    if price_text and not ('₱' in price_text or 'php' in price_text.lower()):
        return None
    if not price_text and str(alternative.get('currency') or '').upper() != 'PHP':
        return None
    seller=str(item.get('source') or '').strip()
    link=str(item.get('link') or item.get('product_link') or '').strip()
    domain=urlparse(link).hostname or ''
    domain=domain.lower().removeprefix('www.')
    is_ph_domain=domain.endswith('.ph')
    known_market=any(m in seller.lower() for m in PH_MARKET_NAMES) or any(domain==d or domain.endswith('.'+d) for d in PH_MARKET_DOMAINS)
    known_store=any(domain==d or domain.endswith('.'+d) for d in PH_LOCAL_STORE_DOMAINS)
    # Google Shopping links can point to Google rather than a seller's website;
    # accept explicitly named PH marketplaces with a caveat, not generic Google listings.
    if not (is_ph_domain or known_market or known_store):
        return None
    price=number(item.get('extracted_price'))
    if price is None:
        match=re.search(r'[\d,]+(?:\.\d+)?',price_text)
        price=number(match.group()) if match else None
    if price is None or price<0 or not link.startswith('https://'):
        return None
    return {'name':item.get('title') or '', 'cost':price,'supplier':seller or 'PH storefront',
            'source_url':link,'uom':'pc','currency':'PHP',
            'delivery_note':str(item.get('delivery') or ''),
            'local_status':'PH marketplace/store listing — seller ship-from location NOT verified'}

@app.get('/api/search')
def online_search():
    q=(request.args.get('q') or '').strip()
    if len(q)<2: return jsonify(error='Enter a product name.'),400
    api_key=os.environ.get('SERPAPI_KEY','').strip()
    if not api_key: return jsonify(error='Online sourcing is not connected. Set SERPAPI_KEY on the server to enable live Google Shopping results.',needs_setup=True),503
    try:
        response=requests.get('https://serpapi.com/search.json',params={
            'engine':'google_shopping','q':q+' Philippines local supplier',
            'gl':'ph','hl':'en','location':'Metro Manila, Philippines',
            'api_key':api_key,'num':60},timeout=22)
        response.raise_for_status()
        data=response.json()
        if data.get('error'): return jsonify(error=data['error']),502
        offers=[]
        seen=set()
        for item in data.get('shopping_results',[]):
            offer=philippine_listing(item)
            if offer is None: continue
            unique=(offer['name'].lower(),offer['supplier'].lower(),round(offer['cost'],2))
            if unique in seen: continue
            seen.add(unique)
            offers.append(offer)
        offers.sort(key=lambda x:x['cost'])
        return jsonify(results=offers[:3],searched=q,notice=(
            'Showing up to three lowest-priced eligible peso listings from Philippine-facing storefronts returned by Google Shopping, NOT all local suppliers. '
            'Local stock and shipping origin are UNVERIFIED; confirm seller location, stock, exact model/specifications, warranty, VAT, shipping and total landed cost before adding to a quotation. '
            'A Philippine marketplace listing can still ship from overseas.'))
    except requests.RequestException:
        return jsonify(error='The shopping service could not be reached. Please retry.'),502

@app.get('/api/quotations')
def quotation_history():
    if not CLOUD:
        return jsonify([])
    rows=cloud_request('GET','quotations',params={'select':'id,created_at,company,contact,total,item_count','order':'created_at.desc','limit':'100'})
    return jsonify(rows)

@app.post('/api/quotation.xlsx')
def quotation():
    data=request.get_json(force=True)
    items=data.get('items',[])
    if not isinstance(items,list) or not items or len(items)>100: return jsonify(error='Add 1–100 valid items before generating a quotation.'),400
    checked=[]
    for item in items:
        qty=number(item.get('qty')); cost=number(item.get('cost')); markup=number(item.get('markup'),0); vat=number(item.get('vat'),12)
        if not str(item.get('name') or '').strip() or qty is None or qty<=0 or cost is None or cost<0 or markup is None or markup<0 or vat is None or vat<0 or vat>100:
            return jsonify(error='Each item requires a name, positive quantity, valid cost, markup, and VAT.'),400
        price=round(cost*(1+markup/100)*(1+vat/100)+1e-9,2)
        checked.append((item,qty,price))
    wb=load_workbook(ROOT/'assets'/'quotation_template.xlsx')
    ws=wb.active
    extra=max(0,len(checked)-2)
    if extra:
        old=list(ws.merged_cells.ranges)
        for merged in old: ws.unmerge_cells(str(merged))
        ws.insert_rows(20,extra)
        for m in old:
            a=CellRange(str(m))
            if a.min_row>=20: a.shift(row_shift=extra)
            elif a.max_row>=20: a.max_row+=extra
            ws.merge_cells(str(a))
    # Release the template's first product description merge before filling item rows.
    if 'C18:E18' in [str(m) for m in ws.merged_cells.ranges]: ws.unmerge_cells('C18:E18')
    # Use the original template's first line formatting for each additional line
    for idx,(item,qty,price) in enumerate(checked):
        r=18+idx
        if idx>=2:
            for col in range(1,8):
                origin=ws.cell(18,col); target=ws.cell(r,col)
                if origin.has_style: target._style=copy.copy(origin._style)
                target.alignment=copy.copy(origin.alignment)
            ws.row_dimensions[r].height=ws.row_dimensions[18].height
        for col in (3,4,5):
            ws.cell(r,col).value=None
        # merge description across C:E, same as template's first item
        key=f'C{r}:E{r}'
        if key not in [str(x) for x in ws.merged_cells.ranges]: ws.merge_cells(key)
        ws.cell(r,1).value=qty
        ws.cell(r,2).value=str(item.get('uom') or 'pc')[:20]
        ws.cell(r,3).value=str(item['name'])[:300]
        ws.cell(r,3).alignment=copy.copy(ws.cell(18,3).alignment)
        ws.cell(r,6).value=price
        ws.cell(r,7).value=f'=A{r}*F{r}'
        ws.cell(r,6).number_format='#,##0.00'
        ws.cell(r,7).number_format='#,##0.00'
        ws.row_dimensions[r].height=max(ws.row_dimensions[r].height or 15,28)
    if len(checked)==1:
        for c in (1,2,3,5,6,7): ws.cell(19,c).value=None
    nothing_row=20+extra; total_row=21+extra
    ws[f'C{nothing_row}']='***Nothing Follows***'
    ws[f'G{total_row}']=f'=SUM(G18:G{17+len(checked)})'
    ws[f'G{total_row}'].number_format='"PHP "#,##0.00'
    ws['A6']=datetime.date.today()
    ws['A8']=str(data.get('contact') or '')[:160]
    ws['A9']=str(data.get('position') or '')[:160]
    ws['A11']=str(data.get('company') or '')[:160]
    ws['A12']=str(data.get('address') or '')[:300]
    ws['A14']=f"Dear {str(data.get('salutation') or 'Sir/Madam').strip()[:90]},"
    ws[f'A{25+extra}']='AVAILABILITY: '+str(data.get('availability') or 'ORDER BASIS!')[:200]
    ws[f'A{26+extra}']='WARRANTY: '+str(data.get('warranty') or 'N/A')[:200]
    ws[f'A{27+extra}']='LEADTIME: '+str(data.get('leadtime') or '1 - 2 WEEKS')[:200]
    ws[f'A{28+extra}']='DELIVERY CHARGE: '+str(data.get('delivery') or 'Free delivery around Metro Manila only.')[:200]
    ws[f'A{29+extra}']='TERMS OF PAYMENT: '+str(data.get('payment') or '30 DAYS')[:200]
    ws[f'A{30+extra}']='REMARKS: '+str(data.get('remarks') or 'Item price and stock availability subject to change without prior notice.')[:300]
    ws[f'F{35+extra}']=str(data.get('contact') or '')[:160]
    ws[f'F{36+extra}']=str(data.get('position') or '')[:160]
    ws[f'A{35+extra}']=str(data.get('salesperson') or 'Mr. Charles Dorado')[:160]
    ws[f'A{36+extra}']=str(data.get('sales_title') or 'Account Executive')[:160]
    ws.print_options.horizontalCentered=True
    ws.print_area=f'A1:G{43+extra}'
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.fitToWidth=1
    ws.page_setup.fitToHeight=1
    # Supplier costs, margins and links remain outside the customer quotation.
    output=io.BytesIO(); wb.save(output); output.seek(0)
    if CLOUD:
        try:
            cloud_request('POST','quotations',payload={
                'company':str(data.get('company') or '')[:160],
                'contact':str(data.get('contact') or '')[:160],
                'total':round(sum(q*p for _,q,p in checked),2),
                'item_count':len(checked),
                'details':data
            },prefer='return=minimal')
        except Exception:
            app.logger.exception('Quotation generated but failed to save history')
    filename='Familia_Quotation_'+datetime.date.today().isoformat()+'.xlsx'
    return send_file(output,as_attachment=True,download_name=filename,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__=='__main__':
    app.run(host='0.0.0.0',port=int(os.environ.get('PORT','5000')),debug=False)
