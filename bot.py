import os, re, json, sqlite3, shutil, threading
from datetime import datetime
from pathlib import Path
from copy import copy
from flask import Flask
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN')
BASE_DIR = Path(__file__).resolve().parent
EXCEL_NAME = os.environ.get('EXCEL_FILE', 'SISTEMA_PEDIDOS_SERGIO_V6.xlsx')
EXCEL_FILE = Path(EXCEL_NAME)
if not EXCEL_FILE.is_absolute():
    EXCEL_FILE = BASE_DIR / EXCEL_FILE
DB_NAME = os.environ.get('DB_FILE', 'bot_v6.sqlite3')
DB_FILE = Path(DB_NAME)
if not DB_FILE.is_absolute():
    DB_FILE = BASE_DIR / DB_FILE
OUT_DIR = Path(os.environ.get('OUT_DIR', 'pedidos_finales'))
if not OUT_DIR.is_absolute(): OUT_DIR = BASE_DIR / OUT_DIR
app = Flask(__name__)

# V6: catalog-driven order assistant. No fake visual recognition is claimed.
# Photos are attached to the item; matching is catalog + natural-language parsing.

ALIASES = {
    'barca': 'Barcelona', 'barça': 'Barcelona', 'fcb': 'Barcelona', 'fc barcelona': 'Barcelona',
    'real madrid': 'Real Madrid', 'rma': 'Real Madrid',
    'local': 'Local', 'home': 'Local', 'titular': 'Local',
    'visitante': 'Visitante', 'away': 'Visitante', 'segunda': 'Visitante',
    'player edition': 'Player Edition', 'player': 'Player Edition', 'jugador': 'Player Edition', 'jugadores': 'Player Edition',
    'fan edition': 'Fan Edition', 'fan': 'Fan Edition', 'aficionado': 'Fan Edition',
    'retro': 'Retro', 'retro jersey': 'Retro', 'clasica': 'Retro', 'clásica': 'Retro',
    'liga': 'Liga', 'laliga': 'Liga', 'la liga': 'Liga',
    'corta': 'Corta', 'manga corta': 'Corta', 'larga': 'Larga', 'manga larga': 'Larga',
}
SIZE_ORDER = ['XS','S','M','L','XL','XXL','3XL','4XL','5XL']

# Roster shortcuts used only for the personalization UI. The product/ID still
# comes from CATALOGO; these names do not create products that are absent there.
PLAYER_ROSTERS = {
    'Real Madrid': [
        ('1','Courtois'),('5','Bellingham'),('7','Vini Jr.'),('8','Valverde'),
        ('10','Mbappé'),('11','Rodrygo'),('14','Tchouameni'),('15','Arda Güler'),
        ('6','Camavinga'),('22','Rüdiger'),('4','Huijsen')
    ],
    'Barcelona': [
        ('1','Joan García'),('3','Balde'),('5','Pau Cubarsí'),('6','Gavi'),
        ('7','Fermín López'),('8','Pedri'),('9','Gabriel Jesus'),('10','Lamine Yamal'),
        ('11','Raphinha'),('20','Dani Olmo'),('21','Frenkie de Jong')
    ],
}

MODEL_LABELS = {
    'Local': '🏠 Local / Casa',
    'Visitante': '✈️ Visitante',
    'Tercera': '3️⃣ Tercera',
}

# ---------- persistence ----------
def db():
    c = sqlite3.connect(str(DB_FILE), check_same_thread=False)
    c.row_factory = sqlite3.Row
    c.execute('CREATE TABLE IF NOT EXISTS orders (chat_id INTEGER PRIMARY KEY, order_id INTEGER, status TEXT, data TEXT)')
    c.execute('CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)')
    if c.execute("SELECT 1 FROM meta WHERE k='next_order'").fetchone() is None:
        c.execute("INSERT INTO meta(k,v) VALUES('next_order','1')")
    c.commit(); return c

def load_state(chat_id):
    c=db(); r=c.execute('SELECT * FROM orders WHERE chat_id=?',(chat_id,)).fetchone(); c.close()
    return json.loads(r['data']) if r else None

def save_state(chat_id,state):
    c=db(); c.execute('INSERT OR REPLACE INTO orders(chat_id,order_id,status,data) VALUES(?,?,?,?)',
                      (chat_id,state['id'],state['status'],json.dumps(state,ensure_ascii=False))); c.commit(); c.close()

def next_id():
    c=db(); r=c.execute("SELECT v FROM meta WHERE k='next_order'").fetchone(); n=int(r['v'])
    c.execute("UPDATE meta SET v=? WHERE k='next_order'",(str(n+1),)); c.commit(); c.close(); return n

# ---------- excel/catalog ----------
def norm(s):
    s='' if s is None else str(s).strip().lower()
    return s.translate(str.maketrans('áéíóúüñ','aeiouun'))

def catalog_rows():
    if not EXCEL_FILE.exists(): return []
    wb=load_workbook(str(EXCEL_FILE),read_only=True,data_only=True)
    if 'CATALOGO' not in wb.sheetnames: wb.close(); return []
    ws=wb['CATALOGO']; rows=list(ws.iter_rows(values_only=True)); wb.close()
    if not rows: return []
    headers=[str(x).strip() if x is not None else '' for x in rows[0]]
    return [dict(zip(headers,r)) for r in rows[1:] if any(x is not None for x in r)]

def active(r):
    v=norm(r.get('ACTIVO'))
    return v in ('','true','1','si','yes','activo')

def unique_values(field, rows=None):
    rows = catalog_rows() if rows is None else rows
    out=[]
    for r in rows:
        v=r.get(field)
        if v not in (None,'') and str(v) not in out: out.append(str(v))
    return out

def parse_sizes(raw):
    if raw in (None,''): return []
    t=norm(raw).upper().replace('–','-').replace('—','-').replace(' ','')
    found=[]
    # Expand ranges such as S-3XL so the UI shows S, M, L, XL, XXL, 3XL.
    parts=re.split(r'[-/,;|]+',t)
    for part in parts:
        if part in SIZE_ORDER and part not in found: found.append(part)
    for i in range(len(parts)-1):
        a,b=parts[i],parts[i+1]
        if a in SIZE_ORDER and b in SIZE_ORDER:
            ia,ib=SIZE_ORDER.index(a),SIZE_ORDER.index(b)
            if ia<=ib:
                for x in SIZE_ORDER[ia:ib+1]:
                    if x not in found: found.append(x)
    # Also detect sizes embedded in text.
    for s in SIZE_ORDER:
        if re.search(rf'(?<![A-Z0-9]){re.escape(s)}(?![A-Z0-9])',t) and s not in found:
            found.append(s)
    return sorted(found,key=lambda x: SIZE_ORDER.index(x) if x in SIZE_ORDER else 99)

def parse_number(x):
    if x in (None,''): return None
    m=re.search(r'(?<!\d)(\d{1,2})(?!\d)',str(x))
    return m.group(1) if m else None

def canonical_aliases():
    # catalog names become first-class vocabulary; this is what makes the parser catalog-driven.
    out=dict(ALIASES)
    for field in ('EQUIPO','TIPO_EDICION','MODELO','PARCHE','CATEGORIA','MANGA','TEMPORADA','JUGADOR','ID_PRODUCTO'):
        for v in unique_values(field):
            if v: out[norm(v)] = str(v)
    return out

def extract_name_number(text):
    raw=text or ''
    name=None; number=None
    m=re.search(r'(?:nombre|name)\s*[:=\-]?\s*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ .\'-]{1,35}?)(?=\s+(?:numero|número|num|#|talla|size|parche|patch|cantidad|x)\b|\s*$)',raw,re.I)
    if m: name=m.group(1).strip(' .-')
    m=re.search(r'(?:numero|número|num|#)\s*[:=\-]?\s*(\d{1,2})\b',raw,re.I)
    if m: number=m.group(1)
    # Common compact form: Yamal 10 / Mbappe 9 / Sergio #10.
    if not name:
        for r in catalog_rows():
            player=r.get('JUGADOR')
            if player and norm(player) in norm(raw):
                name=str(player); break
    if not number and name:
        # Use catalog official number only when the detected name exactly matches a catalog player.
        for r in catalog_rows():
            if r.get('JUGADOR') and norm(r.get('JUGADOR'))==norm(name) and r.get('NUMERO') not in (None,''):
                number=parse_number(r.get('NUMERO')); break
    return name, number

def parse(text):
    t=norm(text)
    d={'size':None,'team':None,'edition':None,'model':None,'patch':None,'name':None,'number':None,'category':None,'quantity':1,'season':None,'sleeve':None}
    # sizes: longest first, word boundaries
    for s in sorted(SIZE_ORDER,key=len,reverse=True):
        if re.search(rf'(?<![a-z0-9]){norm(s)}(?![a-z0-9])',t): d['size']=s; break
    aliases=canonical_aliases()
    # Exact-ish aliases first, longest first.
    for a,b in sorted(aliases.items(),key=lambda kv:len(kv[0]),reverse=True):
        if not a: continue
        if re.search(rf'(?<![a-z0-9]){re.escape(a)}(?![a-z0-9])',t):
            if b in unique_values('EQUIPO') or b in ('Barcelona','Real Madrid'): d['team']=b
            elif b in unique_values('TIPO_EDICION') or norm(b) in ('player edition','fan edition','retro'): d['edition']=b
            elif b in unique_values('MODELO') or b in ('Local','Visitante'): d['model']=b
            elif b in unique_values('PARCHE') or norm(b)=='liga': d['patch']=b
            elif b in unique_values('CATEGORIA'): d['category']=b
            elif b in unique_values('TEMPORADA'): d['season']=b
            elif b in unique_values('MANGA') or b in ('Corta','Larga'): d['sleeve']=b
    # Team fuzzy fallbacks for natural phrases.
    for r in catalog_rows():
        team=r.get('EQUIPO')
        if team and norm(team) in t: d['team']=str(team)
    name,number=extract_name_number(text)
    d['name']=name; d['number']=number
    m=re.search(r'(?:cantidad|qty|unidades|uds?|x)\s*[:=]?\s*(\d+)',t)
    if m: d['quantity']=max(1,min(50,int(m.group(1))))
    return d

def row_matches_context(r,d):
    if not active(r): return False
    for field,key in [('EQUIPO','team'),('TIPO_EDICION','edition'),('MODELO','model'),('PARCHE','patch'),('CATEGORIA','category'),('TEMPORADA','season'),('MANGA','sleeve')]:
        if d.get(key) and norm(r.get(field))!=norm(d[key]): return False
    return True

def score_row(r,d,text=''):
    if not active(r): return -1
    s=0
    def add(field,key,pts):
        nonlocal s
        if d.get(key) and norm(r.get(field))==norm(d[key]): s+=pts
    add('EQUIPO','team',45); add('TIPO_EDICION','edition',30); add('MODELO','model',20); add('PARCHE','patch',12)
    add('CATEGORIA','category',10); add('TEMPORADA','season',8); add('MANGA','sleeve',5)
    if d.get('size') and d['size'] in parse_sizes(r.get('TALLAS')): s+=8
    if d.get('name'):
        if norm(d['name'])==norm(r.get('JUGADOR')): s+=28
        elif norm(d['name']) in norm(r.get('JUGADOR')): s+=16
    if d.get('number') and parse_number(r.get('NUMERO'))==d['number']: s+=18
    rid=norm(r.get('ID_PRODUCTO'))
    if rid and rid in norm(text): s+=80
    # small bonus for words from supplier description/model
    blob=norm(' '.join(str(r.get(k) or '') for k in ('MODELO','DESCRIPCION_PROVEEDOR','NOTAS')))
    for w in set(re.findall(r'[a-z0-9]+',norm(text))):
        if len(w)>=4 and w in blob: s+=1
    return s

def candidates(text,limit=5):
    d=parse(text); scored=sorted(((score_row(r,d,text),r) for r in catalog_rows()),key=lambda x:x[0],reverse=True)
    return [(s,r) for s,r in scored if s>0][:limit]

def desc(r):
    return f"{r.get('EQUIPO','')} {r.get('TEMPORADA','')} — {r.get('TIPO_EDICION','')} — {r.get('MODELO','')} — {r.get('JUGADOR') or 'sin jugador'} {r.get('NUMERO') or ''} — {r.get('PARCHE') or 'sin parche'}"

# ---------- excel helpers ----------
def ensure_headers(ws,headers):
    existing={str(c.value).strip():i for i,c in enumerate(ws[1],1) if c.value}
    for h in headers:
        if h not in existing:
            col=ws.max_column+1; ws.cell(1,col,h); existing[h]=col
    return existing

def find_order_rows(ws,order_id):
    heads=[str(c.value).strip() if c.value else '' for c in ws[1]]; idx={h:i+1 for i,h in enumerate(heads) if h}
    pid=idx.get('PEDIDO_ID');
    if not pid: return []
    return [r for r in range(2,ws.max_row+1) if str(ws.cell(r,pid).value)==str(order_id)]

def append_order_row(state,item):
    wb=load_workbook(str(EXCEL_FILE)); ws=wb['PEDIDOS']
    col=ensure_headers(ws,['PEDIDO_ID','FECHA','CLIENTE','ESTADO','ID_PRODUCTO','MODELO','DESCRIPCION','TALLA','CANTIDAD','NOMBRE','NUMERO','PARCHE','IMAGEN_REFERENCIA','NOTAS','ORIGEN','CONFIANZA','TIPO_EDICION','NOMBRE_SOLICITADO','NUMERO_SOLICITADO','FOTO_TELEGRAM','ITEM_ID'])
    # Stable item ID lets us edit/delete without leaving stale rows.
    item_id=item.get('item_id') or f"{state['id']}-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    item['item_id']=item_id
    r=ws.max_row+1; c=item.get('catalog') or {}; d=item.get('parsed') or {}
    values={'PEDIDO_ID':state['id'],'FECHA':datetime.now(),'CLIENTE':item.get('cliente','Telegram'),'ESTADO':'ABIERTO',
            'ID_PRODUCTO':c.get('ID_PRODUCTO'),'MODELO':c.get('MODELO'),'DESCRIPCION':desc(c),'TALLA':d.get('size',''),
            'CANTIDAD':item.get('quantity',d.get('quantity',1)),'NOMBRE':c.get('JUGADOR') or '','NUMERO':c.get('NUMERO') or '',
            'PARCHE':d.get('patch') or c.get('PARCHE') or '','IMAGEN_REFERENCIA':c.get('IMAGEN') or '',
            'NOTAS':item.get('caption') or item.get('text') or '','ORIGEN':'Telegram','CONFIANZA':item.get('confidence',0),
            'TIPO_EDICION':c.get('TIPO_EDICION') or d.get('edition') or '',
            'NOMBRE_SOLICITADO':item.get('requested_name') or '','NUMERO_SOLICITADO':item.get('requested_number') or '',
            'FOTO_TELEGRAM':item.get('photo_file_id') or '','ITEM_ID':item_id}
    for k,v in values.items(): ws.cell(r,col[k],v)
    wb.save(str(EXCEL_FILE)); wb.close()

def delete_order_item_row(order_id,item_id):
    wb=load_workbook(str(EXCEL_FILE)); ws=wb['PEDIDOS']
    heads=[str(c.value).strip() if c.value else '' for c in ws[1]]; idx={h:i+1 for i,h in enumerate(heads) if h}
    cid=idx.get('ITEM_ID'); cpid=idx.get('PEDIDO_ID')
    if not cid: wb.close(); return
    for r in range(ws.max_row,1,-1):
        if str(ws.cell(r,cid).value)==str(item_id) and (not cpid or str(ws.cell(r,cpid).value)==str(order_id)):
            ws.delete_rows(r,1); break
    wb.save(str(EXCEL_FILE)); wb.close()

def set_order_status(order_id,status):
    wb=load_workbook(str(EXCEL_FILE)); ws=wb['PEDIDOS']; heads=[str(c.value).strip() if c.value else '' for c in ws[1]]; idx={h:i+1 for i,h in enumerate(heads) if h}
    pid=idx.get('PEDIDO_ID'); ce=idx.get('ESTADO')
    if pid and ce:
        for r in range(2,ws.max_row+1):
            if str(ws.cell(r,pid).value)==str(order_id): ws.cell(r,ce,status)
    wb.save(str(EXCEL_FILE)); wb.close()

def rebuild_output(order_id):
    OUT_DIR.mkdir(parents=True,exist_ok=True); out=OUT_DIR/f'PEDIDO_PROVEEDOR_{order_id:04d}.xlsx'; shutil.copy2(str(EXCEL_FILE),str(out))
    wb=load_workbook(str(out)); wp=wb['PEDIDO_PROVEEDOR']; wo=wb['PEDIDOS']
    # Preserve the supplier sheet template: clear only values in rows after the header row.
    for row in range(2,wp.max_row+1):
        for col in range(1,wp.max_column+1): wp.cell(row,col).value=None
    # Remove existing images from old template rows so the final file only carries current order images.
    try: wp._images=[]
    except Exception: pass
    heads=[str(c.value).strip() if c.value else '' for c in wo[1]]; idx={h:i+1 for i,h in enumerate(heads) if h}
    rows=[]
    for rr in range(2,wo.max_row+1):
        if str(wo.cell(rr,idx.get('PEDIDO_ID',1)).value)==str(order_id) and str(wo.cell(rr,idx.get('ESTADO',4)).value) in ('ABIERTO','CERRADO'):
            rows.append({h:wo.cell(rr,i).value for h,i in idx.items()})
    outrow=2
    for item in rows:
        qty=int(item.get('CANTIDAD') or 1)
        for _ in range(max(1,qty)):
            # Existing V1 supplier template is primarily MODELO / DESCRIPCION / TALLA.
            if wp.max_column >= 1: wp.cell(outrow,1,item.get('MODELO') or '')
            if wp.max_column >= 2: wp.cell(outrow,2,(item.get('DESCRIPCION') or '').replace('—','-'))
            if wp.max_column >= 3: wp.cell(outrow,3,item.get('TALLA') or '')
            if wp.max_column >= 4: wp.cell(outrow,4,item.get('ID_PRODUCTO') or '')
            if wp.max_column >= 5: wp.cell(outrow,5,item.get('NOMBRE_SOLICITADO') or '')
            if wp.max_column >= 6: wp.cell(outrow,6,item.get('NUMERO_SOLICITADO') or '')
            if wp.max_column >= 7: wp.cell(outrow,7,item.get('PARCHE') or '')
            wp.row_dimensions[outrow].height=95
            img=item.get('IMAGEN_REFERENCIA')
            if img and os.path.exists(str(img)):
                try:
                    x=XLImage(str(img)); x.width=85; x.height=85; wp.add_image(x,f'H{outrow}')
                except Exception: pass
            outrow+=1
    wb.save(str(out)); wb.close(); return str(out)

# ---------- UI ----------
def cb(text,data): return InlineKeyboardButton(text,callback_data=data)

def candidate_label(r):
    rid=r.get('ID_PRODUCTO') or 'SIN-ID'; team=r.get('EQUIPO') or ''; edition=r.get('TIPO_EDICION') or ''; model=r.get('MODELO') or ''
    return f"{rid} — {team} — {edition} — {model}"[:64]

def candidate_keyboard(rs):
    return InlineKeyboardMarkup([[cb(candidate_label(r),f'cand:{i}')] for i,(_,r) in enumerate(rs)])

def filtered_rows(d):
    rows=[r for r in catalog_rows() if active(r)]
    for field,key in [('EQUIPO','team'),('CATEGORIA','category'),('TEMPORADA','season'),('MODELO','model'),('TIPO_EDICION','edition'),('PARCHE','patch'),('MANGA','sleeve')]:
        if d.get(key): rows=[r for r in rows if norm(r.get(field))==norm(d[key])]
    return rows

def all_context_options(field,d):
    rows=filtered_rows(d)
    vals=[]
    for r in rows:
        v=r.get(field)
        if v not in (None,'') and str(v) not in vals: vals.append(str(v))
    return vals

def size_keyboard(d):
    sizes=[]
    for r in filtered_rows(d):
        for s in parse_sizes(r.get('TALLAS')):
            if s not in sizes: sizes.append(s)
    sizes=sorted(sizes,key=lambda x: SIZE_ORDER.index(x) if x in SIZE_ORDER else 99)
    return InlineKeyboardMarkup([[cb(s,f'set:size:{s}') for s in sizes[i:i+4]] for i in range(0,len(sizes),4)]) if sizes else None

def smart_keyboard(item):
    d=item['parsed']; rows=[]
    # Edition must be explicit: Player and Fan are separate supplier variants.
    if not d.get('edition'):
        opts=all_context_options('TIPO_EDICION',d)
        # Always expose the two main choices; if a product is not yet in the
        # catalog, selection will be rejected rather than inventing an ID.
        ed=[]
        for x in ('Player Edition','Fan Edition'):
            if x in opts or x in ('Player Edition','Fan Edition'):
                ed.append(cb('👕 Player' if x=='Player Edition' else '👕 Fan',f'set:edition:{x}'))
        if ed: rows.append(ed)

    # Shirt version: only actual catalog models are offered for the current context.
    if not d.get('model'):
        opts=all_context_options('MODELO',d)
        # The supplier distinguishes Home/Local, Away/Visitante and Third/Tercera.
        # Show the standard three when the catalog has those variants; selecting one
        # never invents a product if the corresponding catalog row is absent.
        standard=[]
        for x in ('Local','Visitante','Tercera'):
            if x in opts or x in ('Local','Visitante','Tercera'):
                standard.append(cb(MODEL_LABELS.get(x, x),f'set:model:{x}'))
        if standard: rows.append(standard)

    if not d.get('size'):
        sk=size_keyboard(d)
        if sk: rows.extend(sk.inline_keyboard)

    # Player Edition: choose a known player by shirt number, or choose custom.
    if d.get('edition') and norm(d.get('edition'))=='player edition' and not item.get('player_selected') and not item.get('custom_answered'):
        team=d.get('team') or (item.get('catalog') or {}).get('EQUIPO')
        roster=PLAYER_ROSTERS.get(team,[])
        if roster:
            rows.append([cb(f'{num} — {name}',f'player:{num}:{name}') for num,name in roster[:3]])
            for i in range(3,len(roster),3):
                rows.append([cb(f'{num} — {name}',f'player:{num}:{name}') for num,name in roster[i:i+3]])
        rows.append([cb('✏️ Personalizado','custom:yes'),cb('Sin nombre','custom:no')])

    if not d.get('patch'):
        opts=all_context_options('PARCHE',d)
        if opts: rows.append([cb(x,f'set:patch:{x}') for x in opts[:4]])

    # Quantity starts at ONE. The +/- controls never silently turn a photo into 2 units.
    qty=int(item.get('quantity') or 1)
    qty=max(1,min(50,qty))
    rows.append([cb('➖', 'qty:-1'),cb(f'Cantidad: {qty}','qty:noop'),cb('➕','qty:+1')])
    rows.append([cb('✅ Confirmar artículo','confirm'),cb('❌ Cancelar artículo','cancel')])
    return InlineKeyboardMarkup(rows)

def item_summary(item):
    c=item.get('catalog') or {}; d=item.get('parsed') or {}
    return (f"📋 {c.get('ID_PRODUCTO','Sin ID')} — {c.get('EQUIPO','')} — {c.get('TIPO_EDICION','')} — {c.get('MODELO','')}\n"
            f"📏 Talla: {d.get('size') or '—'}\n"
            f"👤 Nombre solicitado: {item.get('requested_name') or '—'}\n"
            f"🔢 Número solicitado: {item.get('requested_number') or '—'}\n"
            f"🏷️ Parche: {d.get('patch') or c.get('PARCHE') or '—'}\n"
            f"🔢 Cantidad: {item.get('quantity',d.get('quantity',1))}")

def kb_confirm(item):
    return InlineKeyboardMarkup([[cb('✅ Confirmar pedido','confirm'),cb('✏️ Editar','edit')],[cb('❌ Cancelar artículo','cancel')]])

# ---------- telegram ----------
async def start(update,context):
    await update.message.reply_text('🤖 Bot V6 listo.\n\nEscribe «nuevo pedido» y luego envía fotos o descripciones. El bot usa el CATALOGO para interpretar equipos, ediciones, modelos, tallas y personalizaciones.')

async def new_order(update,context):
    chat=update.effective_chat.id; old=load_state(chat)
    if old and old.get('status')=='ABIERTO':
        await update.message.reply_text(f'⚠️ Ya tienes el pedido #{old["id"]:04d} abierto. Usa «cierro pedido» o «reabrir pedido».'); return
    state={'id':next_id(),'status':'ABIERTO','items':[],'pending':None}; save_state(chat,state)
    await update.message.reply_text(f'🟢 Pedido #{state["id"]:04d} abierto.\nEnvía la primera foto o escribe los datos.')

async def close_order(update,context):
    chat=update.effective_chat.id; state=load_state(chat)
    if not state: await update.message.reply_text('No hay pedido abierto.'); return
    if state.get('pending'): await update.message.reply_text('⚠️ Tienes un artículo pendiente. Confírmalo o cancélalo antes de cerrar.'); return
    if not state['items']: await update.message.reply_text('⚠️ El pedido está vacío.'); return
    set_order_status(state['id'],'CERRADO'); state['status']='CERRADO'; save_state(chat,state)
    try: path=rebuild_output(state['id'])
    except Exception as e:
        await update.message.reply_text(f'⚠️ Cerré el pedido, pero no pude generar el Excel: {e}'); return
    await update.message.reply_text(f'🔒 Pedido #{state["id"]:04d} cerrado.\n📦 Artículos: {len(state["items"])}\n📄 Enviando Excel final…')
    with open(path,'rb') as f: await update.message.reply_document(document=f,filename=os.path.basename(path),caption=f'📄 PEDIDO_PROVEEDOR_{state["id"]:04d}.xlsx — pedido cerrado. Usa «reabrir pedido» si necesitas corregirlo.')

async def reopen(update,context):
    chat=update.effective_chat.id; state=load_state(chat)
    if not state: await update.message.reply_text('No hay pedido para reabrir.'); return
    state['status']='ABIERTO'; state['pending']=None; save_state(chat,state)
    try: set_order_status(state['id'],'ABIERTO')
    except Exception: pass
    await update.message.reply_text(f'🔓 Pedido #{state["id"]:04d} reabierto. Puedes corregir, eliminar o agregar artículos.')

async def show_order(update,context):
    state=load_state(update.effective_chat.id)
    if not state: await update.message.reply_text('No hay pedido.'); return
    lines=[f'📦 Pedido #{state["id"]:04d} — {state["status"]}','']
    for i,x in enumerate(state['items'],1):
        lines.append(f"{i}. {x.get('catalog',{}).get('ID_PRODUCTO','?')} — {x.get('catalog',{}).get('EQUIPO','')} — {x.get('catalog',{}).get('TIPO_EDICION','')} | Talla {x.get('parsed',{}).get('size') or '?'} | {x.get('requested_name') or 'sin nombre'} {x.get('requested_number') or ''}")
    if state.get('pending'): lines.append('\n⏳ Artículo pendiente de confirmación.')
    await update.message.reply_text('\n'.join(lines))

async def handle_photo(update,context):
    state=load_state(update.effective_chat.id)
    if not state or state['status']!='ABIERTO': await update.message.reply_text('🔒 No hay pedido abierto. Usa «nuevo pedido».'); return
    if state.get('pending'):
        await update.message.reply_text('⚠️ Primero termina el artículo anterior.'); return
    caption=update.message.caption or ''
    p={'photo_file_id':update.message.photo[-1].file_id,'caption':caption,'text':caption,'parsed':parse(caption),'catalog':None,'quantity':1}
    rs=candidates(caption)
    p['candidates']=[r for _,r in rs]; state['pending']=p; save_state(update.effective_chat.id,state)
    if rs:
        await update.message.reply_text('📸 Foto recibida. Identificación del catálogo:',reply_markup=candidate_keyboard(rs))
    else:
        await update.message.reply_text('📸 Foto recibida. No pude determinar el producto solo con el texto de la foto. Escribe equipo/modelo/edición/talla y te muestro las opciones del CATALOGO.',reply_markup=smart_keyboard(p))

async def ask_custom_name(q,state,p):
    p['awaiting']='custom_name'; state['pending']=p; save_state(q.message.chat.id,state)
    await q.message.reply_text('✏️ Escribe el nombre y número. Ejemplos: «Sergio 10», «nombre Sergio número 10» o «Yamal 10».')

async def process_text_for_pending(update,text):
    state=load_state(update.effective_chat.id); p=state.get('pending') if state else None
    if not state or state.get('status')!='ABIERTO' or not p: return False
    if p.get('awaiting')=='custom_name':
        name,number=extract_name_number(text)
        if not name:
            # In free-form custom mode, first word(s) before a number is the name.
            m=re.search(r'^\s*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ .\'-]{1,35}?)(?:\s*#?\s*(\d{1,2}))?\s*$',text)
            if m: name=m.group(1).strip(); number=number or m.group(2)
        p['requested_name']=name; p['requested_number']=number; p['custom_answered']=True; p['awaiting']=None
        save_state(update.effective_chat.id,state)
        await update.message.reply_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return True
    # Merge text into the same pending item. Do not create a duplicate.
    p['text']=(p.get('text','')+' '+text).strip(); p['parsed']=parse(p['text'])
    if p.get('catalog'):
        d=p['parsed']
        # Preserve explicit manual selections.
        if p.get('manual_size'): d['size']=p['manual_size']
        if p.get('manual_edition'): d['edition']=p['manual_edition']
        if p.get('manual_patch'): d['patch']=p['manual_patch']
        if p.get('manual_model'): d['model']=p['manual_model']
        if d.get('name') and not p.get('requested_name'):
            # A catalog player is official; anything else must be explicitly treated as personalization.
            official=norm(p['catalog'].get('JUGADOR'))
            if norm(d['name'])!=official:
                p['awaiting']='custom_name'; save_state(update.effective_chat.id,state)
                await update.message.reply_text(f'⚠️ «{d["name"]}» no figura como jugador oficial de {p["catalog"].get("ID_PRODUCTO")}. ¿Es personalización? Escribe «sí» o indica el nombre/número.')
                return True
            p['requested_name']=d['name']; p['requested_number']=d.get('number')
        if not d.get('size'):
            await update.message.reply_text('📏 Falta la talla. Elige una:',reply_markup=smart_keyboard(p)); return True
        if norm(p['catalog'].get('TIPO_EDICION'))=='player edition' and not p.get('custom_answered') and not p.get('requested_name'):
            await update.message.reply_text(item_summary(p)+'\n\n¿Quieres poner nombre y número?',reply_markup=smart_keyboard(p)); return True
        await update.message.reply_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return True
    rs=candidates(p.get('text','')); p['candidates']=[r for _,r in rs]; save_state(update.effective_chat.id,state)
    if rs: await update.message.reply_text('🔎 Identificación actualizada:',reply_markup=candidate_keyboard(rs))
    else: await update.message.reply_text('No encontré una coincidencia clara. Usa los botones para completar los datos:',reply_markup=smart_keyboard(p))
    return True

async def text_item(update,context,text):
    state=load_state(update.effective_chat.id)
    if not state or state['status']!='ABIERTO': await update.message.reply_text('🔒 No hay pedido abierto. Usa «nuevo pedido».'); return
    if state.get('pending'):
        await process_text_for_pending(update,text); return
    p={'photo_file_id':None,'caption':'','text':text,'parsed':parse(text),'catalog':None,'quantity':1}
    rs=candidates(text); p['candidates']=[r for _,r in rs]; state['pending']=p; save_state(update.effective_chat.id,state)
    if rs: await update.message.reply_text('🔎 Identificación encontrada en el CATALOGO:',reply_markup=candidate_keyboard(rs))
    else: await update.message.reply_text('⚠️ No encontré coincidencia suficiente. Completa los datos:',reply_markup=smart_keyboard(p))

async def handle_text(update,context):
    raw=update.message.text or ''; t=norm(raw)
    if t=='nuevo pedido': return await new_order(update,context)
    if t in ('cierro pedido','cierra pedido','cerrar pedido','cierra el pedido','cerrar el pedido','finalizar pedido','finaliza pedido','terminar pedido','termina pedido'): return await close_order(update,context)
    if t in ('reabrir pedido','reabrir'): return await reopen(update,context)
    if t in ('muéstrame el pedido','muestrame el pedido','mostrar pedido','ver pedido'): return await show_order(update,context)
    state=load_state(update.effective_chat.id)
    if t in ('corrige la ultima','corrige la última'):
        if state and state['items']:
            item=state['items'].pop();
            try: delete_order_item_row(state['id'],item.get('item_id'))
            except Exception: pass
            state['pending']=item; item['awaiting']=None; save_state(update.effective_chat.id,state)
            await update.message.reply_text('✏️ Último artículo devuelto a edición. Revisa y confirma nuevamente.',reply_markup=smart_keyboard(item))
        else: await update.message.reply_text('No hay artículo que corregir.')
        return
    if t in ('elimina la ultima','elimina la última'):
        if state and state['items']:
            item=state['items'].pop()
            try: delete_order_item_row(state['id'],item.get('item_id'))
            except Exception: pass
            save_state(update.effective_chat.id,state); await update.message.reply_text('🗑️ Último artículo eliminado del pedido.')
        else: await update.message.reply_text('No hay artículo que eliminar.')
        return
    if state and state.get('pending') and state['pending'].get('awaiting')=='custom_confirm':
        if t in ('si','sí','yes'):
            state['pending']['awaiting']='custom_name'; save_state(update.effective_chat.id,state)
            await update.message.reply_text('Perfecto. Escribe el nombre y número, por ejemplo: «Sergio 10».'); return
        if t in ('no','no gracias'):
            state['pending']['awaiting']=None; state['pending']['custom_answered']=True; save_state(update.effective_chat.id,state)
            await update.message.reply_text(item_summary(state['pending'])+'\n\nConfirma el artículo:',reply_markup=kb_confirm(state['pending'])); return
    if state and state.get('pending') and state['pending'].get('awaiting')=='custom_name':
        if t in ('si','sí','yes'):
            await update.message.reply_text('Perfecto. Escribe el nombre y número, por ejemplo: «Sergio 10».'); return
    await text_item(update,context,raw)

async def callbacks(update,context):
    q=update.callback_query; await q.answer(); state=load_state(q.message.chat.id)
    if not state or not state.get('pending'): return
    p=state['pending']; data=q.data
    if data.startswith('cand:'):
        i=int(data.split(':')[1]); rs=p.get('candidates',[])
        if i<0 or i>=len(rs): return
        p['catalog']=rs[i]; p['confidence']=100; d=p['parsed']; c=p['catalog']
        # Candidate selection can be followed by missing fields; never assume S/3XL only.
        if not d.get('size'):
            await q.edit_message_text(f'✅ Seleccionado: {candidate_label(c)}\n\n📏 Ahora elige la talla disponible:',reply_markup=smart_keyboard(p)); state['pending']=p; save_state(q.message.chat.id,state); return
        if norm(c.get('TIPO_EDICION'))=='player edition' and not p.get('player_selected') and not p.get('custom_answered'):
            await q.edit_message_text(item_summary(p)+'\n\n👕 Player Edition: elige el jugador por número o selecciona Personalizado.',reply_markup=smart_keyboard(p)); state['pending']=p; save_state(q.message.chat.id,state); return
        # Explicit name/number or detected catalog player.
        if d.get('name'):
            if c.get('JUGADOR') and norm(d['name'])==norm(c.get('JUGADOR')):
                p['requested_name']=d['name']; p['requested_number']=d.get('number') or parse_number(c.get('NUMERO')); p['custom_answered']=True
            elif re.search(r'\b(?:nombre|name)\b',p.get('text',''),re.I):
                p['requested_name']=d['name']; p['requested_number']=d.get('number'); p['custom_answered']=True
            else:
                p['awaiting']='custom_confirm'; state['pending']=p; save_state(q.message.chat.id,state)
                await q.edit_message_text(f'⚠️ «{d["name"]}» no figura como jugador oficial de este producto.\n\n¿Quieres usarlo como personalización?', reply_markup=InlineKeyboardMarkup([[cb('✅ Sí, personalizar','custom:yes'),cb('❌ No','custom:no')]]))
                return
        if norm(c.get('TIPO_EDICION'))=='player edition' and not p.get('requested_name') and not p.get('custom_answered'):
            await q.edit_message_text(item_summary(p)+'\n\n¿Quieres poner nombre y número?',reply_markup=smart_keyboard(p)); state['pending']=p; save_state(q.message.chat.id,state); return
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); state['pending']=p; save_state(q.message.chat.id,state); return
    if data.startswith('player:'):
        _,num,name=data.split(':',2)
        p['player_selected']=True; p['requested_name']=name; p['requested_number']=num; p['custom_answered']=True
        p['parsed']['name']=name; p['parsed']['number']=num
        state['pending']=p; save_state(q.message.chat.id,state)
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
    if data=='custom:yes':
        await ask_custom_name(q,state,p); return
    if data=='custom:no':
        p['custom_answered']=True; p['requested_name']=None; p['requested_number']=None; state['pending']=p; save_state(q.message.chat.id,state)
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
    if data.startswith('qty:'):
        action=data.split(':',1)[1]
        if action=='noop': return
        qv=max(1,min(50,int(p.get('quantity') or 1)+(1 if action=='+1' else -1)))
        p['quantity']=qv; p['parsed']['quantity']=qv; state['pending']=p; save_state(q.message.chat.id,state)
        await q.edit_message_reply_markup(reply_markup=smart_keyboard(p)); return
    if data.startswith('set:'):
        _,field,val=data.split(':',2)
        if field=='size': p['parsed']['size']=val; p['manual_size']=val
        elif field=='edition': p['parsed']['edition']=val; p['manual_edition']=val
        elif field=='patch': p['parsed']['patch']=val; p['manual_patch']=val
        elif field=='model': p['parsed']['model']=val; p['manual_model']=val
        elif field=='qty': p['quantity']=max(1,min(50,int(val))); p['parsed']['quantity']=p['quantity']
        # Re-rank after every manual choice; if one clear product remains, select it automatically.
        search_text=p.get('text','')+' '+val
        rs=candidates(search_text)
        # Keep candidates constrained by all manually selected fields.
        rs=[(score_row(r,p['parsed'],search_text),r) for r in catalog_rows()]
        rs=sorted([(s,r) for s,r in rs if s>0],key=lambda x:x[0],reverse=True)[:5]
        p['candidates']=[r for _,r in rs]
        if not p['candidates'] and (p['parsed'].get('edition') or p['parsed'].get('model')):
            state['pending']=p; save_state(q.message.chat.id,state)
            await q.edit_message_text('⚠️ Esa combinación todavía no existe en el CATALOGO. No voy a inventar un ID de proveedor. Puedes elegir otra edición/modelo o agregar ese producto al Excel.')
            return
        if not p.get('catalog') and len(p['candidates'])==1:
            p['catalog']=p['candidates'][0]; p['confidence']=95
        elif p.get('catalog') and any(p.get(k) and norm(p['catalog'].get(f))!=norm(p[k]) for f,k in [('EQUIPO','team'),('TIPO_EDICION','edition'),('MODELO','model')]):
            p['catalog']=None
        state['pending']=p; save_state(q.message.chat.id,state)
        if p.get('catalog'):
            if not p['parsed'].get('size'):
                await q.edit_message_reply_markup(reply_markup=smart_keyboard(p)); return
            if norm(p['catalog'].get('TIPO_EDICION'))=='player edition' and not p.get('player_selected') and not p.get('custom_answered'):
                await q.edit_message_text(item_summary(p)+'\n\n👕 Player Edition: elige el jugador por número o selecciona Personalizado.',reply_markup=smart_keyboard(p)); return
            await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
        await q.edit_message_reply_markup(reply_markup=smart_keyboard(p)); return
    if data=='confirm':
        if not p.get('catalog'):
            await q.message.reply_text('⚠️ Primero selecciona un producto del CATALOGO.'); return
        if not p.get('parsed',{}).get('size'):
            await q.message.reply_text('⚠️ Falta la talla. Selecciona una antes de confirmar.',reply_markup=smart_keyboard(p)); return
        if norm(p['catalog'].get('TIPO_EDICION'))=='player edition' and not p.get('custom_answered') and not p.get('requested_name'):
            await q.message.reply_text('⚠️ Indica si quieres nombre/número antes de confirmar.',reply_markup=smart_keyboard(p)); return
        p['parsed']=parse(p.get('text',''))
        # restore manual selections after reparsing
        for k,mk in [('size','manual_size'),('edition','manual_edition'),('patch','manual_patch'),('model','manual_model')]:
            if p.get(mk): p['parsed'][k]=p[mk]
        if p.get('requested_name') is None and p.get('custom_answered') is False: p['requested_name']=None
        p['quantity']=p.get('quantity',p['parsed'].get('quantity',1))
        if not p.get('item_id'): p['item_id']=f"{state['id']}-{len(state['items'])+1}-{datetime.now().strftime('%H%M%S%f')}"
        try: append_order_row(state,p)
        except Exception as e: await q.message.reply_text(f'⚠️ No pude guardar el artículo en Excel: {e}'); return
        state['items'].append(p); state['pending']=None; save_state(q.message.chat.id,state)
        await q.edit_message_text(f'✅ Artículo agregado al pedido #{state["id"]:04d}.\n📦 Artículos: {len(state["items"])}\n\n¿Qué quieres hacer ahora?', reply_markup=InlineKeyboardMarkup([[cb('📋 Revisar pedido','review'),cb('➡️ Siguiente artículo','next')]]))
        return
    if data=='review':
        lines=[f'📦 Pedido #{state["id"]:04d} — {state["status"]}','']
        for i,x in enumerate(state['items'],1):
            c=x.get('catalog',{}); d=x.get('parsed',{})
            lines.append(f'{i}. {c.get("ID_PRODUCTO","?")} — {c.get("EQUIPO","")} — {c.get("TIPO_EDICION","")} — {c.get("MODELO","")} | {d.get("size") or "?"} | {x.get("requested_name") or "sin nombre"} {x.get("requested_number") or ""} | x{x.get("quantity",1)}')
        lines.append('\nCuando termines escribe «cierro pedido».')
        await q.message.reply_text('\n'.join(lines)); return
    if data=='next':
        await q.edit_message_text('➡️ Listo. Envía la foto del siguiente artículo.')
        return
    if data=='cancel':
        state['pending']=None; save_state(q.message.chat.id,state); await q.edit_message_text('❌ Artículo cancelado. Puedes enviar otra foto.'); return
    if data=='edit':
        await q.message.reply_text('✏️ Escribe la corrección. Ej.: «talla XL, nombre Sergio, número 10, parche Liga».'); return

async def ping(update,context): await update.message.reply_text('V6 activo.')

def flask_thread(): app.run(host='0.0.0.0',port=int(os.environ.get('PORT','10000')))

@app.get('/')
def health():
    exists=EXCEL_FILE.exists(); n=len(catalog_rows()) if exists else 0
    return f'Bot V6 activo - Excel: {"OK" if exists else "FALTA"} - productos: {n}',200 if exists else 503

def main():
    if not TOKEN: raise RuntimeError('Falta TELEGRAM_BOT_TOKEN')
    if not EXCEL_FILE.exists(): raise RuntimeError(f'No existe el Excel del catálogo: {EXCEL_FILE}')
    if not catalog_rows(): raise RuntimeError('El CATALOGO está vacío o no se pudo leer.')
    threading.Thread(target=flask_thread,daemon=True).start()
    application=Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler('start',start)); application.add_handler(CommandHandler('nuevo',new_order)); application.add_handler(CommandHandler('cerrar',close_order));
    application.add_handler(CallbackQueryHandler(callbacks)); application.add_handler(MessageHandler(filters.PHOTO,handle_photo)); application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,handle_text));
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__=='__main__': main()
