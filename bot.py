import os, re, json, sqlite3, shutil, threading, urllib.request, urllib.error, base64, tempfile, zipfile
from datetime import datetime
from pathlib import Path
from copy import copy
from flask import Flask
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN')
OPENAI_API_KEY = os.environ.get('OPENAI_API_KEY')
VISION_MODEL = os.environ.get('VISION_MODEL', 'gpt-5.6-luna')
BASE_DIR = Path(__file__).resolve().parent
EXCEL_NAME = os.environ.get('EXCEL_FILE', 'SISTEMA_PEDIDOS_SERGIO_V9.xlsx')
EXCEL_FILE = Path(EXCEL_NAME)
if not EXCEL_FILE.is_absolute():
    EXCEL_FILE = BASE_DIR / EXCEL_FILE
DB_NAME = os.environ.get('DB_FILE', 'bot_v9.sqlite3')
DB_FILE = Path(DB_NAME)
if not DB_FILE.is_absolute():
    DB_FILE = BASE_DIR / DB_FILE
OUT_DIR = Path(os.environ.get('OUT_DIR', 'pedidos_finales'))
if not OUT_DIR.is_absolute(): OUT_DIR = BASE_DIR / OUT_DIR
app = Flask(__name__)

# V9.3: multi-user isolation. Each customer's state is keyed by Telegram user_id (not a single/global account).
# A photo is analyzed by a vision model when OPENAI_API_KEY is configured; the
# model identifies the club/selection and kit attributes, then the catalog decides
# the supplier ID. Player mode validates official roster numbers; custom mode does not.

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
    'Real Madrid': [('1','Courtois'),('13','Lunin'),('2','Asencio'),('3','E. Militão'),('4','Huijsen'),('12','Trent'),('16','Konaté'),('17','Cucurella'),('18','Á. Carreras'),('22','Rüdiger'),('23','F. Mendy'),('24','Dumfries'),('5','Bellingham'),('6','Camavinga'),('8','Valverde'),('14','Tchouameni'),('15','Arda Güler'),('20','Bernardo'),('27','Thiago'),('7','Vini Jr.'),('9','Endrick'),('10','Mbappé'),('11','Rodrygo'),('19','C. Espí'),('21','Brahim'),('25','Yan Diomande')],
    'Barcelona': [('1','Joan García'),('2','João Cancelo'),('3','Alejandro Balde'),('4','Brian Fariñas'),('5','Pau Cubarsí'),('6','Gavi'),('7','Fermín López'),('8','Pedri'),('9','Gabriel Jesus'),('10','Lamine Yamal'),('11','Raphinha'),('12','Xavi Espart'),('13','Szczęsny'),('14','Adeyemi'),('15','Andreas Christensen'),('16','Rodrigo'),('17','Anthony Gordon'),('18','Gerard Martín'),('19','Roony Bardghji'),('20','Dani Olmo'),('21','Frenkie de Jong'),('22','Marc Bernal'),('23','Jules Kounde'),('24','Eric Garcia'),('25','Dominik Livaković'),('27','Jesse Bisiwu'),('29','Hamza Abdelkarim')],
    'Inter Miami': [('10','Lionel Messi')],
    'Panamá': [('1','Luis Mejía'),('2','César Blackman'),('3','José Córdoba'),('4','Fidel Escobar'),('5','Edgardo Fariña'),('6','Cristian Martínez'),('7','José Luis Rodríguez'),('8','Adalberto Carrasquilla'),('9','Tomás Rodríguez'),('10','Ismael Díaz'),('11','Yoel Bárcenas'),('12','César Samudio'),('13','Jiovany Ramos'),('14','Carlos Harvey'),('15','Eric Davis'),('16','Andrés Andrade'),('17','José Fajardo'),('18','Cecilio Waterman'),('19','Alberto Quintero'),('20','Aníbal Godoy'),('21','César Yanis'),('22','Orlando Mosquera'),('23','Amir Murillo'),('24','Azarías Londoño'),('25','Roderick Miller'),('26','Jorge Gutiérrez')],
}

PATCH_OPTIONS = ['Liga','Champions']

PLAYER_ALIASES = {
    'yamal': 'Lamine Yamal',
    'lamin yamal': 'Lamine Yamal',
    'lamine yamal': 'Lamine Yamal',
    'messi': 'Lionel Messi',
    'mbappe': 'Mbappé',
    'mbappé': 'Mbappé',
    'vini': 'Vini Jr.',
    'vinicius': 'Vini Jr.',
    'vinicius jr': 'Vini Jr.',
    'pedri': 'Pedri',
    'raphinha': 'Raphinha',
    'fermin': 'Fermín López',
    'fermín': 'Fermín López',
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

def load_state(user_id):
    c=db(); r=c.execute('SELECT * FROM orders WHERE chat_id=?',(int(user_id),)).fetchone(); c.close()
    return json.loads(r['data']) if r else None

def save_state(user_id,state):
    c=db(); c.execute('INSERT OR REPLACE INTO orders(chat_id,order_id,status,data) VALUES(?,?,?,?)',
                      (int(user_id),state['id'],state['status'],json.dumps(state,ensure_ascii=False))); c.commit(); c.close()

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

def roster_for(team):
    if not team: return []
    if not EXCEL_FILE.exists(): return PLAYER_ROSTERS.get(team, [])
    try:
        wb=load_workbook(str(EXCEL_FILE),read_only=True,data_only=True)
        if 'JUGADORES' not in wb.sheetnames:
            wb.close(); return PLAYER_ROSTERS.get(team, [])
        ws=wb['JUGADORES']; rows=list(ws.iter_rows(values_only=True)); wb.close()
        if not rows: return PLAYER_ROSTERS.get(team, [])
        heads=[str(x).strip() if x is not None else '' for x in rows[0]]
        idx={h:i for i,h in enumerate(heads)}
        out=[]
        for rr in rows[1:]:
            t=rr[idx.get('EQUIPO',0)] if idx.get('EQUIPO') is not None and idx.get('EQUIPO')<len(rr) else None
            num=rr[idx.get('NUMERO',3)] if idx.get('NUMERO',3)<len(rr) else None
            name=rr[idx.get('JUGADOR',4)] if idx.get('JUGADOR',4)<len(rr) else None
            active_v=rr[idx.get('ACTIVO',6)] if idx.get('ACTIVO',6)<len(rr) else 'SI'
            if t and name and norm(t)==norm(team) and norm(active_v) not in ('no','0','false'):
                out.append((str(num),str(name)))
        return out or PLAYER_ROSTERS.get(team, [])
    except Exception:
        return PLAYER_ROSTERS.get(team, [])

def roster_match(team, query):
    q=norm(query)
    if not team or not q: return None
    rr=roster_for(team)
    # Common football-name aliases / surname-only input.
    alias=PLAYER_ALIASES.get(q)
    if alias:
        q=norm(alias)
    if q.isdigit():
        for num,name in rr:
            if num==q: return (num,name)
    # Exact canonical name first.
    for num,name in rr:
        if norm(name)==q: return (num,name)
    # Surname / token match: 'Yamal' must resolve to 'Lamine Yamal'.
    q_tokens=set(re.findall(r'[a-z0-9]+',q))
    hits=[]
    for num,name in rr:
        nt=set(re.findall(r'[a-z0-9]+',norm(name)))
        if q in norm(name) or q_tokens.intersection(nt):
            hits.append((num,name))
    if len(hits)==1: return hits[0]
    # Small typo tolerance, e.g. 'Lamin Yamal'.
    import difflib
    best=None; best_score=0.0
    for num,name in rr:
        score=difflib.SequenceMatcher(None,q,norm(name)).ratio()
        if score>best_score:
            best_score=score; best=(num,name)
    if best and best_score>=0.72: return best
    return None

def vision_catalog_context(limit=120):
    rows=[r for r in catalog_rows() if active(r)]
    out=[]
    for r in rows[:limit]:
        out.append({
            'id':r.get('ID_PRODUCTO'),'team':r.get('EQUIPO'),'season':r.get('TEMPORADA'),
            'edition':r.get('TIPO_EDICION'),'model':r.get('MODELO'),'category':r.get('CATEGORIA'),
            'sleeve':r.get('MANGA'),'patch':r.get('PARCHE'),'sizes':r.get('TALLAS'),
            'player':r.get('JUGADOR'),'number':r.get('NUMERO')})
    return out

def _extract_json(text):
    if not text: return {}
    text=text.strip()
    try: return json.loads(text)
    except Exception: pass
    m=re.search(r'\{.*\}',text,re.S)
    if m:
        try: return json.loads(m.group(0))
        except Exception: pass
    return {}

def vision_identify(image_path, caption=''):
    if not OPENAI_API_KEY: return None, 'OPENAI_API_KEY no configurada'
    raw=Path(image_path).read_bytes(); b64=base64.b64encode(raw).decode('ascii')
    ext=Path(image_path).suffix.lower()
    mime='image/jpeg' if ext in ('.jpg','.jpeg') else 'image/png' if ext=='.png' else 'image/webp'
    prompt={
      'task':'Identify the football shirt in the photo for a Telegram order bot. Return ONLY JSON.',
      'rules':[
        'Identify the club or national team from visible crest, colors, sponsor, badge and design.',
        'Identify home/local, away/visitante, or third/tercera only when visually supported; otherwise null.',
        'Identify Player Edition vs Fan Edition only when visually supported; otherwise null.',
        'NEVER identify or infer a patch from the image. The patch is always a customer order choice and must remain null unless the customer explicitly selects or writes it.',
        'Do not invent a supplier ID. The catalog will choose the ID after your visual result.',
        'If the image is not a football shirt, team=null and confidence=0.',
        'confidence is 0-100 and means visual confidence, not certainty.'
      ],
      'catalog':vision_catalog_context(),
      'caption':caption,
      'schema':{'team':'string|null','season':'string|null','model':'Local|Visitante|Tercera|null','edition':'Player Edition|Fan Edition|Retro|null','category':'string|null','patch':'null','confidence':'number','notes':'string'}
    }
    body={
      'model':VISION_MODEL,
      'input':[{'role':'user','content':[{'type':'input_text','text':json.dumps(prompt,ensure_ascii=False)}, {'type':'input_image','image_url':f'data:{mime};base64,{b64}'}]}],
      'max_output_tokens':500
    }
    req=urllib.request.Request('https://api.openai.com/v1/responses',data=json.dumps(body).encode('utf-8'),headers={'Authorization':f'Bearer {OPENAI_API_KEY}','Content-Type':'application/json'})
    try:
        with urllib.request.urlopen(req,timeout=45) as resp: data=json.loads(resp.read().decode('utf-8'))
        texts=[]
        for item in data.get('output',[]):
            for c in item.get('content',[]):
                if c.get('type')=='output_text': texts.append(c.get('text',''))
        obj=_extract_json('\n'.join(texts))
        return obj, None
    except Exception as e:
        return None, str(e)

def canonical_aliases():
    # catalog names become first-class vocabulary; this is what makes the parser catalog-driven.
    out=dict(ALIASES)
    for field in ('EQUIPO','TIPO_EDICION','MODELO','PARCHE','CATEGORIA','MANGA','TEMPORADA','JUGADOR','ID_PRODUCTO'):
        for v in unique_values(field):
            if v: out[norm(v)] = str(v)
    return out

def extract_name_number(text, team=None):
    raw=text or ''
    name=None; number=None
    m=re.search(r'(?:numero|número|num|#)\s*[:=\-]?\s*(\d{1,2})\b', raw, re.I)
    if m: number=m.group(1)
    rr=roster_for(team) if team else []
    low=norm(raw)
    # Canonical names, surnames and common aliases all resolve to the official roster.
    search_names=[]
    for k,v in PLAYER_ALIASES.items(): search_names.append((k,v))
    for num,pname in rr: search_names.append((norm(pname),pname))
    for alias,canonical in sorted(search_names,key=lambda x:len(x[0]),reverse=True):
        if alias and re.search(rf'(?<![a-z0-9]){re.escape(alias)}(?![a-z0-9])',low):
            hit=roster_match(team,canonical)
            if hit:
                name=hit[1]
                if number is None: number=hit[0]
                break
    if not name:
        # Explicit 'nombre X' syntax.
        m=re.search(r"(?:nombre|name)\s*[:=\-]?\s*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ .'-]{1,35}?)(?=\s+(?:numero|número|num|#|talla|size|parche|patch|cantidad|x)\b|\s*$)",raw,re.I)
        if m: name=m.group(1).strip(' .-')
    if not name and team:
        # Free-form text: find a unique roster player anywhere in the sentence.
        hits=[]
        for num,pname in rr:
            tokens=set(re.findall(r'[a-z0-9]+',norm(pname)))
            if norm(pname) in low or (len(tokens)>0 and any(re.search(rf'(?<![a-z0-9]){re.escape(tok)}(?![a-z0-9])',low) for tok in tokens if len(tok)>=4)):
                hits.append((num,pname))
        if len(hits)==1:
            name=hits[0][1]
            if number is None: number=hits[0][0]
    if not name:
        m=re.match(r"^\s*([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ .'-]{1,35}?)(?:\s*#?\s*(\d{1,2}))?\s*$",raw)
        if m:
            candidate=m.group(1).strip(); number=number or m.group(2)
            hit=roster_match(team,candidate)
            name=hit[1] if hit else candidate
    if name and team:
        hit=roster_match(team,name)
        if hit:
            name=hit[1]
            if number is None: number=hit[0]
    return name, number

def parse(text, team=None):
    t=norm(text)
    d={'size':None,'team':None,'edition':None,'model':None,'patch':None,'name':None,'number':None,'category':None,'quantity':1,'season':None,'sleeve':None}
    rows=catalog_rows()
    # Sizes
    for size in sorted(SIZE_ORDER,key=len,reverse=True):
        if re.search(rf'(?<![a-z0-9]){re.escape(norm(size))}(?![a-z0-9])',t):
            d['size']=size; break
    # Catalog-driven vocabulary + fixed aliases.
    aliases=canonical_aliases()
    for a,b in sorted(aliases.items(),key=lambda kv:len(kv[0]),reverse=True):
        if not a: continue
        if re.search(rf'(?<![a-z0-9]){re.escape(a)}(?![a-z0-9])',t):
            if b in unique_values('EQUIPO') or norm(b) in ('barcelona','real madrid','inter miami','panama'):
                d['team']=b
            elif b in unique_values('TIPO_EDICION') or norm(b) in ('player edition','fan edition','retro'):
                d['edition']=b
            elif b in unique_values('MODELO') or norm(b) in ('local','visitante','tercera'):
                d['model']=b
            elif b in unique_values('PARCHE') or norm(b) in ('liga','champions'):
                d['patch']=b
            elif b in unique_values('CATEGORIA'): d['category']=b
            elif b in unique_values('TEMPORADA'): d['season']=b
            elif b in unique_values('MANGA') or norm(b) in ('corta','larga'): d['sleeve']=b
    # Fallback team detection from all catalog rows.
    for r in rows:
        team_name=r.get('EQUIPO')
        if team_name and norm(team_name) in t:
            d['team']=str(team_name); break
    name,number=extract_name_number(text,d.get('team') or team)
    d['name']=name; d['number']=number
    m=re.search(r'(?:cantidad|qty|unidades|uds?|x)\s*[:=]?\s*(\d+)',t)
    if m: d['quantity']=max(1,min(50,int(m.group(1))))
    return d


def row_matches_context(r,d):
    if not active(r): return False
    # PATCH IS AN ORDER CUSTOMIZATION, NOT A SKU IDENTITY FIELD.
    # A catalog row may describe the supplier's default patch (e.g. Liga),
    # while the customer can request another patch (e.g. Champions).
    for field,key in [('EQUIPO','team'),('TIPO_EDICION','edition'),('MODELO','model'),('CATEGORIA','category'),('TEMPORADA','season'),('MANGA','sleeve')]:
        if d.get(key) and r.get(field) not in (None,'') and norm(r.get(field))!=norm(d[key]): return False
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

def candidates(text, d=None, limit=5):
    d=d or parse(text)
    scored=[]
    for r in catalog_rows():
        if not active(r): continue
        score=score_row(r,d,text)
        # Hard exclusions only for explicit visual/manual facts.
        bad=False
        for field,key in [('EQUIPO','team'),('MODELO','model'),('TIPO_EDICION','edition'),('TEMPORADA','season')]:
            if d.get(key) and r.get(field) not in (None,'') and norm(r.get(field))!=norm(d[key]): bad=True
        if not bad and score>0: scored.append((score,r))
    return sorted(scored,key=lambda x:x[0],reverse=True)[:limit]


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
            'PARCHE':d.get('patch') or c.get('PARCHE') or '','IMAGEN_REFERENCIA':item.get('telegram_photo_path') or c.get('IMAGEN') or '',
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
    for field,key in [('EQUIPO','team'),('CATEGORIA','category'),('TEMPORADA','season'),('MODELO','model'),('TIPO_EDICION','edition'),('MANGA','sleeve')]:
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
        ed=[]
        for x in opts:
            ed.append(cb('👕 Player' if norm(x)=='player edition' else ('👕 Fan' if norm(x)=='fan edition' else f'👕 {x}'),f'set:edition:{x}'))
        if ed: rows.append(ed)

    # Only expose model variants that actually exist in the current catalog context.
    # This prevents the UI from offering Visitante/Tercera when those supplier SKUs
    # have not yet been loaded into CATALOGO.
    if not d.get('model'):
        opts=all_context_options('MODELO',d)
        model_buttons=[cb(MODEL_LABELS.get(x, x),f'set:model:{x}') for x in opts]
        if model_buttons: rows.append(model_buttons[:3])

    if not d.get('size'):
        sk=size_keyboard(d)
        if sk: rows.extend(sk.inline_keyboard)

    # Player Edition: choose a known player by shirt number, or choose custom.
    if d.get('edition') and norm(d.get('edition'))=='player edition' and not item.get('player_selected') and not item.get('custom_answered'):
        team=d.get('team') or (item.get('catalog') or {}).get('EQUIPO')
        roster=roster_for(team)
        if roster:
            rows.append([cb(f'{num} — {name}',f'player:{num}:{name}') for num,name in roster[:3]])
            for i in range(3,len(roster),3):
                rows.append([cb(f'{num} — {name}',f'player:{num}:{name}') for num,name in roster[i:i+3]])
        rows.append([cb('✏️ Personalizado','custom:yes'),cb('Sin nombre','custom:no')])

    if not d.get('patch'):
        opts=all_context_options('PARCHE',d)
        # Customer patch is an order attribute; expose the standard supplier options even
        # when the current catalog row is not split into separate patch IDs.
        patch_opts=[]
        for x in PATCH_OPTIONS + opts:
            if x and x not in patch_opts: patch_opts.append(x)
        if patch_opts: rows.append([cb(x,f'set:patch:{x}') for x in patch_opts[:4]])

    # One photo = one unit by default. Quantity controls are intentionally hidden;
    # quantity is only changed when the customer explicitly writes it.

    # Never expose confirmation while a required field is still missing.
    # For Player Edition, the player decision (official/custom/sin nombre)
    # must also be completed before confirmation.
    ready = bool(d.get('size')) and bool(d.get('patch'))
    if norm(d.get('edition')) == 'player edition':
        ready = ready and bool(item.get('player_selected') or item.get('custom_answered'))
    rows.append([cb('✏️ Escribir datos manualmente','manual:write')])
    if ready:
        rows.append([cb('✅ Confirmar artículo','confirm'),cb('❌ Cancelar artículo','cancel')])
    else:
        rows.append([cb('❌ Cancelar artículo','cancel')])
    return InlineKeyboardMarkup(rows)

def item_summary(item):
    c=item.get('catalog') or {}; d=item.get('parsed') or {}
    return (f"📋 {c.get('ID_PRODUCTO','Sin ID')} — {c.get('EQUIPO','')} — {c.get('TIPO_EDICION','')} — {c.get('MODELO','')}\n"
            f"📏 Talla: {d.get('size') or '—'}\n"
            f"👤 Nombre solicitado: {item.get('requested_name') or '—'}\n"
            f"🔢 Número solicitado: {item.get('requested_number') or '—'}\n"
            f"🏷️ Parche: {d.get('patch') or '—'}\n"
            f"🔢 Cantidad: {item.get('quantity',d.get('quantity',1))}")

def kb_confirm(item):
    return InlineKeyboardMarkup([[cb('✅ Confirmar artículo','confirm'),cb('✏️ Editar','edit')],[cb('❌ Cancelar artículo','cancel')]])

def choose_catalog_from_vision(parsed, vision_conf=0):
    """Select a catalog row from visual facts without inventing IDs."""
    rs=candidates('', parsed, limit=10)
    if not rs: return None, []
    # If vision explicitly identifies team/model/edition/season, respect them.
    # When the catalog has one active row for the identified team, it is safe to
    # use that row as the supplier mapping while clearly retaining the visual result.
    if parsed.get('team'):
        team_rows=[(score_row(r,parsed,''),r) for r in catalog_rows()
                   if active(r) and norm(r.get('EQUIPO'))==norm(parsed['team'])]
        team_rows=sorted(team_rows,key=lambda x:x[0],reverse=True)
        if team_rows:
            # Reject an explicit contradiction in a catalog dimension.
            for field,key in [('MODELO','model'),('TIPO_EDICION','edition'),('TEMPORADA','season')]:
                if parsed.get(key) and team_rows[0][1].get(field) not in (None,'') and norm(team_rows[0][1].get(field))!=norm(parsed[key]):
                    return None,[r for _,r in team_rows[:5]]
            if len(team_rows)==1: return team_rows[0][1], [team_rows[0][1]]
            if vision_conf>=88 and team_rows[0][0]-team_rows[1][0]>=15: return team_rows[0][1],[r for _,r in team_rows[:5]]
    return None,[r for _,r in rs]

# ---------- telegram ----------
async def start(update,context):
    u=update.effective_user
    await update.message.reply_text(f"🤖 Hola {u.first_name or ''}. Este bot está habilitado para múltiples clientes. Tu pedido se guarda separado por tu cuenta de Telegram.\n\nEscribe «nuevo pedido» para comenzar y luego envía fotos o descripciones.")

async def new_order(update,context):
    user_id=update.effective_user.id; old=load_state(user_id)
    if old and old.get('status')=='ABIERTO':
        await update.message.reply_text(f'⚠️ Ya tienes el pedido #{old["id"]:04d} abierto. Usa «cierro pedido» o «reabrir pedido».'); return
    state={'id':next_id(),'status':'ABIERTO','items':[],'pending':None,'customer':{'user_id':user_id,'username':getattr(update.effective_user,'username',None),'name':getattr(update.effective_user,'full_name',None)}}; save_state(user_id,state)
    await update.message.reply_text(f'🟢 Pedido #{state["id"]:04d} abierto.\nEnvía la primera foto o escribe los datos.')

async def close_order(update,context):
    user_id=update.effective_user.id; state=load_state(user_id)
    if not state: await update.message.reply_text('No hay pedido abierto.'); return
    if state.get('status')!='ABIERTO': await update.message.reply_text(f'🔒 El pedido #{state["id"]:04d} ya está cerrado.'); return
    if state.get('pending'):
        await update.message.reply_text('⚠️ Tienes un artículo pendiente. Confírmalo o cancélalo antes de cerrar.'); return
    if not state.get('items'):
        await update.message.reply_text('⚠️ El pedido está vacío.'); return
    try:
        path=rebuild_output(state['id'])
        if not path or not os.path.exists(path): raise RuntimeError('No se generó el Excel final.')
        # Only after the workbook exists successfully do we mark the order closed.
        set_order_status(state['id'],'CERRADO')
        state['status']='CERRADO'; save_state(user_id,state)
        await update.message.reply_text(f'🔒 Pedido #{state["id"]:04d} cerrado.\n📦 Artículos: {len(state["items"])}\n📄 Enviando Excel final…')
        with open(path,'rb') as f: await update.message.reply_document(document=f,filename=os.path.basename(path),caption=f'📄 PEDIDO_PROVEEDOR_{state["id"]:04d}.xlsx — pedido cerrado.')
    except Exception as e:
        # Never leave the order half-closed.
        try:
            state['status']='ABIERTO'; save_state(user_id,state); set_order_status(state['id'],'ABIERTO')
        except Exception: pass
        await update.message.reply_text(f'⚠️ No pude cerrar el pedido porque falló la generación/envío del Excel. El pedido sigue ABIERTO.\n\nDetalle: {e}')


async def reopen(update,context):
    user_id=update.effective_user.id; state=load_state(user_id)
    if not state: await update.message.reply_text('No hay pedido para reabrir.'); return
    state['status']='ABIERTO'; state['pending']=None; save_state(user_id,state)
    try: set_order_status(state['id'],'ABIERTO')
    except Exception: pass
    await update.message.reply_text(f'🔓 Pedido #{state["id"]:04d} reabierto. Puedes corregir, eliminar o agregar artículos.')

async def show_order(update,context):
    state=load_state(update.effective_user.id)
    if not state: await update.message.reply_text('No hay pedido.'); return
    lines=[f'📦 Pedido #{state["id"]:04d} — {state["status"]}','']
    for i,x in enumerate(state['items'],1):
        lines.append(f"{i}. {x.get('catalog',{}).get('ID_PRODUCTO','?')} — {x.get('catalog',{}).get('EQUIPO','')} — {x.get('catalog',{}).get('TIPO_EDICION','')} | Talla {x.get('parsed',{}).get('size') or '?'} | {x.get('requested_name') or 'sin nombre'} {x.get('requested_number') or ''}")
    if state.get('pending'): lines.append('\n⏳ Artículo pendiente de confirmación.')
    await update.message.reply_text('\n'.join(lines))

async def handle_photo(update,context):
    user_id=update.effective_user.id; state=load_state(user_id)
    if not state or state['status']!='ABIERTO':
        await update.message.reply_text('🔒 No hay pedido abierto. Usa «nuevo pedido».'); return
    if state.get('pending'):
        await update.message.reply_text('⚠️ Primero termina el artículo anterior.'); return
    caption=update.message.caption or ''
    photo=update.message.photo[-1]
    p={'photo_file_id':photo.file_id,'caption':caption,'text':caption,'parsed':parse(caption),
       'catalog':None,'quantity':1,'confidence':0,'vision':None,'candidates':[]}
    # Persist the Telegram image so the final supplier workbook can include the reference photo.
    photo_dir=OUT_DIR/'telegram_photos'/str(state['id']); photo_dir.mkdir(parents=True,exist_ok=True)
    photo_path=photo_dir/f'{len(state.get("items",[]))+1}_{photo.file_unique_id}.jpg'
    try:
        tgfile=await context.bot.get_file(photo.file_id)
        await tgfile.download_to_drive(custom_path=str(photo_path))
        p['telegram_photo_path']=str(photo_path)
    except Exception as e:
        p['telegram_photo_error']=str(e)
    # REAL VISION CALL: the downloaded image is sent as a base64 input_image to /v1/responses.
    vis=None; err=None
    if p.get('telegram_photo_path'):
        vis,err=vision_identify(p['telegram_photo_path'],caption)
        # Retry once with the exact same downloaded image. This prevents transient
        # vision/API responses from making identical photos behave differently.
        if vis is None and OPENAI_API_KEY:
            vis2,err2=vision_identify(p['telegram_photo_path'],caption)
            if vis2 is not None:
                vis,err=vis2,None
            elif err2:
                err=f'{err}; reintento: {err2}'
    if vis:
        p['vision']=vis
        for k in ('team','season','model','edition','category','sleeve'):
            if vis.get(k): p['parsed'][k]=vis[k]
        # Patch is NEVER populated from the photo. It is an explicit customer choice.
        p['parsed']['patch'] = p.get('parsed', {}).get('patch') if p.get('parsed', {}).get('patch') and caption else None
        try: p['confidence']=max(0,min(100,float(vis.get('confidence') or 0)))
        except Exception: p['confidence']=0
        # Caption is authoritative for explicit order fields (size/name/number/patch).
        capd=parse(caption,p['parsed'].get('team'))
        for k in ('size','name','number','quantity'):
            if capd.get(k) not in (None,''): p['parsed'][k]=capd[k]
        selected,rs=choose_catalog_from_vision(p['parsed'],p['confidence'])
        p['catalog']=selected; p['candidates']=rs
    else:
        p['candidates']=[r for _,r in candidates(caption,p['parsed'])]
    state['pending']=p; save_state(user_id,state)

    if vis is None:
        extra=''
        if not OPENAI_API_KEY: extra=' Falta OPENAI_API_KEY en Render.'
        elif err: extra=f' Error de visión: {err}'
        await update.message.reply_text('📸 No pude obtener la identificación visual de esta foto.'+extra+'\n\nPuedes escribir la descripción y completar los datos.',reply_markup=smart_keyboard(p)); return

    # Always tell the customer what the vision model actually saw.
    vteam=vis.get('team') or 'No determinado'; vmodel=vis.get('model') or 'No determinado'; ved=vis.get('edition') or 'No determinada'
    vconf=int(round(p['confidence']))
    vision_msg=(f'📸 **Imagen analizada por IA**\n\n⚽ Equipo: **{vteam}**\n'
                f'👕 Modelo: **{vmodel}**\n🎽 Edición: **{ved}**\n'
                f'🎯 Confianza visual: **{vconf}%**\n')
    if vis.get('notes'): vision_msg+=f'📝 {vis.get("notes")}\n'
    if p.get('catalog'):
        c=p['catalog']; vision_msg+=(f'\n🆔 Producto del catálogo: **{c.get("ID_PRODUCTO")}**\n'
                                     f'📦 {c.get("EQUIPO")} · {c.get("TEMPORADA")} · {c.get("TIPO_EDICION")} · {c.get("MODELO")}\n')
        if (not p['parsed'].get('size') or not p['parsed'].get('patch') or (norm(c.get('TIPO_EDICION'))=='player edition' and not p.get('requested_name') and not p.get('custom_answered'))):
            await update.message.reply_text(vision_msg+'\nCompletemos los datos del pedido:',reply_markup=smart_keyboard(p),parse_mode='Markdown')
        else:
            await update.message.reply_text(vision_msg+'\n'+item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p),parse_mode='Markdown')
    else:
        await update.message.reply_text(vision_msg+'\n⚠️ La imagen sí fue identificada, pero no hay una coincidencia inequívoca en CATALOGO. No voy a inventar un ID.\n\nPuedes completar los datos o elegir una coincidencia.',reply_markup=smart_keyboard(p),parse_mode='Markdown')


async def ask_custom_name(q,state,p):
    p['awaiting']='custom_name'; state['pending']=p; save_state(q.from_user.id,state)
    await q.message.reply_text('✏️ Escribe el nombre y número. Ejemplos: «Sergio 10», «nombre Sergio número 10» o «Yamal 10».')

def apply_manual_text(p, state, text):
    """Deterministically merge a manual description into the SAME pending item.
    Manual input is authoritative for explicit order fields and must never be
    interpreted as a new article or require a callback button.
    """
    old = p.get('parsed') or {}
    combined = ((p.get('text') or '') + ' ' + (text or '')).strip()
    # If this is the dedicated manual form, parse the complete text in one pass.
    parsed = parse(combined, old.get('team') or (p.get('catalog') or {}).get('EQUIPO'))
    for k in ('team','season','category','sleeve'):
        if old.get(k): parsed[k] = old[k]
    # Explicit manual values are authoritative over previous vision values.
    for mk,k in [('manual_size','size'),('manual_edition','edition'),('manual_patch','patch'),('manual_model','model')]:
        if p.get(mk): parsed[k] = p[mk]
    # Re-parse explicit manual values so "talla M", "parche Liga", etc. are retained.
    explicit = parse(text, parsed.get('team') or (p.get('catalog') or {}).get('EQUIPO'))
    for k in ('team','season','edition','model','patch','size','name','number','category','sleeve'):
        if explicit.get(k) not in (None,''):
            parsed[k] = explicit[k]
    p['text'] = combined
    p['parsed'] = parsed
    p['quantity'] = 1

    # Resolve player aliases/surnames against the official roster.
    if parsed.get('name'):
        team = parsed.get('team') or (p.get('catalog') or {}).get('EQUIPO')
        hit = roster_match(team, parsed['name'])
        if hit:
            p['requested_name'] = hit[1]
            p['requested_number'] = parsed.get('number') or hit[0]
            p['player_selected'] = True
            p['custom_answered'] = True
            parsed['name'] = hit[1]
            parsed['number'] = p['requested_number']

    # Match catalog using identity fields only; patch/size/name are order details.
    rs = candidates(combined, parsed)
    p['candidates'] = [r for _,r in rs]
    if p.get('catalog'):
        c=p['catalog']
        for field,key in [('EQUIPO','team'),('MODELO','model'),('TIPO_EDICION','edition'),('TEMPORADA','season')]:
            if parsed.get(key) and c.get(field) not in (None,'') and norm(c.get(field)) != norm(parsed[key]):
                p['catalog']=None
                break
    if not p.get('catalog') and rs:
        # Prefer a unique top result; with one active SKU for a team this is deterministic.
        top_score, top_row = rs[0]
        second_score = rs[1][0] if len(rs)>1 else -1
        if len(rs)==1 or top_score > second_score:
            p['catalog']=top_row
            p['confidence']=max(float(p.get('confidence') or 0),95)
    return p

async def process_text_for_pending(update,text):
    state=load_state(update.effective_user.id); p=state.get('pending') if state else None
    if not state or state.get('status')!='ABIERTO' or not p: return False
    if p.get('awaiting') in ('custom_name','custom_confirm','player_mismatch','manual_data'):
        if p.get('awaiting')=='manual_data':
            # Dedicated manual path: resolve the SAME pending item immediately.
            p['awaiting']=None
            p['manual_data_entered']=True
            apply_manual_text(p, state, text)
            # A manual entry must never silently advance to the next article.
            if not p.get('catalog'):
                state['pending']=p; save_state(update.effective_user.id,state)
                await update.message.reply_text('⚠️ No pude encontrar una combinación válida en el CATALOGO.\n\nEscribe, por ejemplo: «Barcelona, Yamal 10, talla M, parche Liga».')
                return True
            if not p.get('parsed',{}).get('size'):
                state['pending']=p; save_state(update.effective_user.id,state)
                await update.message.reply_text(item_summary(p)+'\n\n📏 Falta la talla. Elige una:',reply_markup=smart_keyboard(p)); return True
            if not p.get('parsed',{}).get('patch'):
                state['pending']=p; save_state(update.effective_user.id,state)
                await update.message.reply_text(item_summary(p)+'\n\n🏷️ Elige el parche:',reply_markup=smart_keyboard(p)); return True
            if norm(p['catalog'].get('TIPO_EDICION'))=='player edition' and not p.get('custom_answered') and not p.get('requested_name'):
                state['pending']=p; save_state(update.effective_user.id,state)
                await update.message.reply_text(item_summary(p)+'\n\n👕 Elige el jugador o Personalizado:',reply_markup=smart_keyboard(p)); return True
            state['pending']=p; save_state(update.effective_user.id,state)
            await update.message.reply_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return True

        if p.get('awaiting')=='custom_confirm':
            t=norm(text)
            if t in ('si','sí','yes'):
                p['awaiting']='custom_name'; save_state(update.effective_user.id,state)
                await update.message.reply_text('Perfecto. Escribe el nombre y número, por ejemplo: «Sergio 10».'); return True
            if t in ('no','no gracias'):
                p['awaiting']=None; p['custom_answered']=True; save_state(update.effective_user.id,state)
                await update.message.reply_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return True
        if p.get('awaiting')=='custom_name':
            name,number=extract_name_number(text,p.get('parsed',{}).get('team') or (p.get('catalog') or {}).get('EQUIPO'))
            if not name:
                await update.message.reply_text('No pude leer el nombre. Ejemplo: «Sergio 10».'); return True
            p['requested_name']=name; p['requested_number']=number; p['custom_answered']=True; p['awaiting']=None
            save_state(update.effective_user.id,state)
            await update.message.reply_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return True
        if p.get('awaiting')=='player_mismatch':
            await update.message.reply_text('Selecciona el jugador oficial corregido o usa Personalizado.',reply_markup=InlineKeyboardMarkup([[cb(f'⚽ {p.get("mismatch_name")} #{p.get("mismatch_official_number")}', 'playerfix'), cb('✏️ Personalizado','custom:yes')]])); return True

    # IMPORTANT: corrections are merged into the SAME pending item; no new item is created.
    old=p.get('parsed') or {}
    merged_text=(p.get('text') or '')+' '+text
    p['text']=merged_text.strip()
    parsed=parse(p['text'], old.get('team') or (p.get('catalog') or {}).get('EQUIPO'))
    # Vision/manual facts survive reparsing.
    for k in ('team','season','category','sleeve'):
        if old.get(k): parsed[k]=old[k]
    for mk,k in [('manual_size','size'),('manual_edition','edition'),('manual_patch','patch'),('manual_model','model')]:
        if p.get(mk): parsed[k]=p[mk]
    p['parsed']=parsed
    if parsed.get('quantity'): p['quantity']=parsed['quantity']

    # If the user explicitly changed team/model/edition, rematch catalog.
    if p.get('catalog'):
        c=p['catalog']
        contradiction=False
        for field,key in [('EQUIPO','team'),('MODELO','model'),('TIPO_EDICION','edition')]:
            if parsed.get(key) and c.get(field) and norm(c.get(field))!=norm(parsed[key]): contradiction=True
        if contradiction: p['catalog']=None
    if not p.get('catalog'):
        rs=candidates(p['text'],parsed)
        p['candidates']=[r for _,r in rs]
        if len(rs)==1: p['catalog']=rs[0][1]; p['confidence']=max(float(p.get('confidence') or 0),95)

    # Player name/number validation.
    if parsed.get('name'):
        team=parsed.get('team') or (p.get('catalog') or {}).get('EQUIPO')
        hit=roster_match(team,parsed['name'])
        if hit and parsed.get('number') and parsed['number']!=hit[0]:
            p['awaiting']='player_mismatch'; p['mismatch_name']=hit[1]; p['mismatch_official_number']=hit[0]; p['mismatch_requested_number']=parsed['number']
            save_state(update.effective_user.id,state)
            await update.message.reply_text(f'⚠️ {hit[1]} corresponde al número {hit[0]}, no al {parsed["number"]}.',reply_markup=InlineKeyboardMarkup([[cb(f'⚽ {hit[1]} #{hit[0]}','playerfix'),cb('✏️ Personalizado','custom:yes')]])); return True
        if hit:
            p['requested_name']=hit[1]; p['requested_number']=parsed.get('number') or hit[0]; p['player_selected']=True; p['custom_answered']=True
        elif norm((p.get('catalog') or {}).get('TIPO_EDICION'))=='player edition':
            p['awaiting']='custom_confirm'; save_state(update.effective_user.id,state)
            await update.message.reply_text(f'⚠️ «{parsed["name"]}» no figura en la lista oficial. ¿Quieres usarlo como personalización?',reply_markup=InlineKeyboardMarkup([[cb('✅ Sí, personalizar','custom:yes'),cb('❌ No','custom:no')]])); return True

    # Explicit patch/name/size corrections are reflected immediately.
    if parsed.get('size') and p.get('catalog'):
        await update.message.reply_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); save_state(update.effective_user.id,state); return True
    save_state(update.effective_user.id,state)
    await update.message.reply_text('✏️ Corrección aplicada al mismo artículo:\n\n'+item_summary(p)+'\n\n'+('Selecciona lo que falte:' if not p.get('catalog') or not parsed.get('size') else 'Confirma el artículo:'),reply_markup=(smart_keyboard(p) if (not p.get('catalog') or not parsed.get('size')) else kb_confirm(p)))
    return True


async def text_item(update,context,text):
    state=load_state(update.effective_user.id)
    if not state or state['status']!='ABIERTO': await update.message.reply_text('🔒 No hay pedido abierto. Usa «nuevo pedido».'); return
    if state.get('pending'):
        await process_text_for_pending(update,text); return
    p={'photo_file_id':None,'caption':'','text':text,'parsed':parse(text),'catalog':None,'quantity':1}
    rs=candidates(text); p['candidates']=[r for _,r in rs]; state['pending']=p; save_state(update.effective_user.id,state)
    if rs: await update.message.reply_text('🔎 Identificación encontrada en el CATALOGO:',reply_markup=candidate_keyboard(rs))
    else: await update.message.reply_text('⚠️ No encontré coincidencia suficiente. Completa los datos:',reply_markup=smart_keyboard(p))

async def handle_text(update,context):
    raw=update.message.text or ''; t=norm(raw)
    if t=='nuevo pedido': return await new_order(update,context)
    if re.match(r'^(cierro|cierra|cerrar|finalizar|finaliza|terminar|termina)( el)? pedido$', t): return await close_order(update,context)
    if t in ('reabrir pedido','reabrir'): return await reopen(update,context)
    if t in ('muestrame el pedido','mostrar pedido','ver pedido','pedido'): return await show_order(update,context)
    state=load_state(update.effective_user.id)
    if t in ('corrige la ultima',):
        if state and state['items']:
            item=state['items'].pop();
            try: delete_order_item_row(state['id'],item.get('item_id'))
            except Exception: pass
            state['pending']=item; item['awaiting']=None; save_state(update.effective_user.id,state)
            await update.message.reply_text('✏️ Último artículo devuelto a edición. Revisa y confirma nuevamente.',reply_markup=smart_keyboard(item))
        else: await update.message.reply_text('No hay artículo que corregir.')
        return
    if t in ('elimina la ultima',):
        if state and state['items']:
            item=state['items'].pop()
            try: delete_order_item_row(state['id'],item.get('item_id'))
            except Exception: pass
            save_state(update.effective_user.id,state); await update.message.reply_text('🗑️ Último artículo eliminado del pedido.')
        else: await update.message.reply_text('No hay artículo que eliminar.')
        return
    if state and state.get('pending') and state['pending'].get('awaiting')=='custom_confirm':
        if t in ('si','sí','yes'):
            state['pending']['awaiting']='custom_name'; save_state(update.effective_user.id,state)
            await update.message.reply_text('Perfecto. Escribe el nombre y número, por ejemplo: «Sergio 10».'); return
        if t in ('no','no gracias'):
            state['pending']['awaiting']=None; state['pending']['custom_answered']=True; save_state(update.effective_user.id,state)
            await update.message.reply_text(item_summary(state['pending'])+'\n\nConfirma el artículo:',reply_markup=kb_confirm(state['pending'])); return
    if state and state.get('pending') and state['pending'].get('awaiting')=='custom_name':
        if t in ('si','sí','yes'):
            await update.message.reply_text('Perfecto. Escribe el nombre y número, por ejemplo: «Sergio 10».'); return
    await text_item(update,context,raw)

async def callbacks(update,context):
    q=update.callback_query; await q.answer(); state=load_state(q.from_user.id)
    if not state or not state.get('pending'): return
    p=state['pending']; data=q.data
    if data.startswith('cand:'):
        i=int(data.split(':')[1]); rs=p.get('candidates',[])
        if i<0 or i>=len(rs): return
        p['catalog']=rs[i]; p['confidence']=100; d=p['parsed']; c=p['catalog']
        # Candidate selection can be followed by missing fields; never assume S/3XL only.
        if not d.get('size'):
            await q.edit_message_text(f'✅ Seleccionado: {candidate_label(c)}\n\n📏 Ahora elige la talla disponible:',reply_markup=smart_keyboard(p)); state['pending']=p; save_state(q.from_user.id,state); return
        if norm(c.get('TIPO_EDICION'))=='player edition' and not p.get('player_selected') and not p.get('custom_answered'):
            await q.edit_message_text(item_summary(p)+'\n\n👕 Player Edition: elige el jugador por número o selecciona Personalizado.',reply_markup=smart_keyboard(p)); state['pending']=p; save_state(q.from_user.id,state); return
        # Explicit name/number or detected catalog player.
        if d.get('name'):
            if c.get('JUGADOR') and norm(d['name'])==norm(c.get('JUGADOR')):
                p['requested_name']=d['name']; p['requested_number']=d.get('number') or parse_number(c.get('NUMERO')); p['custom_answered']=True
            elif re.search(r'\b(?:nombre|name)\b',p.get('text',''),re.I):
                p['requested_name']=d['name']; p['requested_number']=d.get('number'); p['custom_answered']=True
            else:
                p['awaiting']='custom_confirm'; state['pending']=p; save_state(q.from_user.id,state)
                await q.edit_message_text(f'⚠️ «{d["name"]}» no figura como jugador oficial de este producto.\n\n¿Quieres usarlo como personalización?', reply_markup=InlineKeyboardMarkup([[cb('✅ Sí, personalizar','custom:yes'),cb('❌ No','custom:no')]]))
                return
        if norm(c.get('TIPO_EDICION'))=='player edition' and not p.get('requested_name') and not p.get('custom_answered'):
            await q.edit_message_text(item_summary(p)+'\n\n¿Quieres poner nombre y número?',reply_markup=smart_keyboard(p)); state['pending']=p; save_state(q.from_user.id,state); return
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); state['pending']=p; save_state(q.from_user.id,state); return
    if data=='playerfix':
        name=p.get('mismatch_name'); num=p.get('mismatch_official_number')
        p['player_selected']=True; p['requested_name']=name; p['requested_number']=num; p['custom_answered']=True; p['awaiting']=None
        save_state(q.from_user.id,state)
        if not p.get('parsed',{}).get('size'):
            await q.edit_message_text(item_summary(p)+'\n\n📏 Ahora elige la talla:',reply_markup=smart_keyboard(p)); return
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
    if data.startswith('player:'):
        _,num,name=data.split(':',2)
        p['player_selected']=True; p['requested_name']=name; p['requested_number']=num; p['custom_answered']=True
        p['parsed']['name']=name; p['parsed']['number']=num
        state['pending']=p; save_state(q.from_user.id,state)
        if not p.get('parsed',{}).get('size'):
            await q.edit_message_text(item_summary(p)+'\n\n📏 Ahora elige la talla:',reply_markup=smart_keyboard(p)); return
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
    if data=='custom:yes':
        await ask_custom_name(q,state,p); return
    if data=='custom:no':
        p['custom_answered']=True; p['requested_name']=None; p['requested_number']=None; state['pending']=p; save_state(q.from_user.id,state)
        if not p.get('parsed',{}).get('size'):
            await q.edit_message_text(item_summary(p)+'\n\n📏 Ahora elige la talla:',reply_markup=smart_keyboard(p)); return
        await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
    if data.startswith('qty:'):
        action=data.split(':',1)[1]
        if action=='noop': return
        qv=max(1,min(50,int(p.get('quantity') or 1)+(1 if action=='+1' else -1)))
        p['quantity']=qv; p['parsed']['quantity']=qv; state['pending']=p; save_state(q.from_user.id,state)
        await q.edit_message_reply_markup(reply_markup=smart_keyboard(p)); return
    if data.startswith('set:'):
        _,field,val=data.split(':',2)
        if field=='size': p['parsed']['size']=val; p['manual_size']=val
        elif field=='edition': p['parsed']['edition']=val; p['manual_edition']=val
        elif field=='patch': p['parsed']['patch']=val; p['manual_patch']=val
        elif field=='model': p['parsed']['model']=val; p['manual_model']=val
        elif field=='qty': p['quantity']=max(1,min(50,int(val))); p['parsed']['quantity']=p['quantity']
        # Re-rank after every manual choice. PATCH does not identify the supplier SKU,
        # so changing Liga <-> Champions must never invalidate an already selected product.
        search_text=p.get('text','')+' '+val
        identity_d={k:v for k,v in p['parsed'].items() if k not in ('patch','size','name','number','quantity')}
        rs=[]
        for r in catalog_rows():
            if not active(r): continue
            bad=False
            for field,key in [('EQUIPO','team'),('MODELO','model'),('TIPO_EDICION','edition'),('TEMPORADA','season')]:
                if identity_d.get(key) and r.get(field) not in (None,'') and norm(r.get(field))!=norm(identity_d[key]):
                    bad=True; break
            if not bad:
                rs.append((score_row(r,p['parsed'],search_text),r))
        rs=sorted([(s,r) for s,r in rs if s>0],key=lambda x:x[0],reverse=True)[:5]
        p['candidates']=[r for _,r in rs]
        if not p.get('catalog') and len(p['candidates'])==1:
            p['catalog']=p['candidates'][0]; p['confidence']=95
        elif p.get('catalog') and any(p.get(k) and k!='patch' and norm(p['catalog'].get(f))!=norm(p[k]) for f,k in [('EQUIPO','team'),('TIPO_EDICION','edition'),('MODELO','model')]):
            p['catalog']=None
        # If an edition/model was explicitly selected and there truly is no SKU, stop here.
        if not p.get('catalog') and not p['candidates'] and (p['parsed'].get('edition') or p['parsed'].get('model')):
            state['pending']=p; save_state(q.from_user.id,state)
            await q.edit_message_text('⚠️ Esa combinación de equipo/edición/modelo todavía no existe en el CATALOGO. No voy a inventar un ID de proveedor. Elige una variante disponible o agrega ese producto al Excel.',reply_markup=smart_keyboard(p))
            return
        state['pending']=p; save_state(q.from_user.id,state)
        if p.get('catalog'):
            if not p['parsed'].get('size'):
                await q.edit_message_reply_markup(reply_markup=smart_keyboard(p)); return
            if norm(p['catalog'].get('TIPO_EDICION'))=='player edition' and not p.get('player_selected') and not p.get('custom_answered'):
                await q.edit_message_text(item_summary(p)+'\n\n👕 Player Edition: elige el jugador por número o selecciona Personalizado.',reply_markup=smart_keyboard(p)); return
            await q.edit_message_text(item_summary(p)+'\n\nConfirma el artículo:',reply_markup=kb_confirm(p)); return
        await q.edit_message_reply_markup(reply_markup=smart_keyboard(p)); return
    if data=='manual:write':
        p['awaiting']='manual_data'; state['pending']=p; save_state(q.from_user.id,state)
        await q.message.reply_text('✏️ Escribe los datos manualmente. Ejemplo: «Barcelona, Yamal 10, talla L, parche Liga».\n\nPuedes escribirlos en cualquier orden; el bot actualizará este mismo artículo.')
        return
    if data=='confirm':
        if not p.get('catalog'):
            await q.message.reply_text('⚠️ Primero selecciona un producto del CATALOGO.'); return
        if not p.get('parsed',{}).get('size'):
            await q.message.reply_text('⚠️ Falta la talla. Selecciona una antes de confirmar.',reply_markup=smart_keyboard(p)); return
        if norm(p['catalog'].get('TIPO_EDICION'))=='player edition' and not p.get('custom_answered') and not p.get('requested_name'):
            await q.message.reply_text('⚠️ Indica si quieres nombre/número antes de confirmar.',reply_markup=smart_keyboard(p)); return
        reparsed=parse(p.get('text',''))
        # Restore manual and vision selections after reparsing.
        for k,mk in [('size','manual_size'),('edition','manual_edition'),('patch','manual_patch'),('model','manual_model')]:
            if p.get(mk): reparsed[k]=p[mk]
        for k in ('team','season','category','sleeve'):
            if p.get('parsed',{}).get(k): reparsed[k]=p['parsed'][k]
        p['parsed']=reparsed
        if p.get('requested_name') is None and p.get('custom_answered') is False: p['requested_name']=None
        p['quantity']=p.get('quantity',p['parsed'].get('quantity',1))
        if not p.get('item_id'): p['item_id']=f"{state['id']}-{len(state['items'])+1}-{datetime.now().strftime('%H%M%S%f')}"
        try: append_order_row(state,p)
        except Exception as e: await q.message.reply_text(f'⚠️ No pude guardar el artículo en Excel: {e}'); return
        state['items'].append(p); state['pending']=None; save_state(q.from_user.id,state)
        await q.edit_message_text(f'✅ Artículo agregado al pedido #{state["id"]:04d}.\n📦 Artículos: {len(state["items"])}\n\nPuedes enviar otra foto directamente o elegir:', reply_markup=InlineKeyboardMarkup([[cb('📋 Revisar pedido','review'),cb('➡️ Siguiente artículo','next')],[cb('🔒 Cerrar pedido','close_now')]]))
        return
    if data=='new_order_btn':
        # Start a fresh order only after the previous one is closed.
        ns={'id':next_id(),'status':'ABIERTO','items':[],'pending':None,'customer':{'user_id':q.from_user.id,'username':getattr(q.from_user,'username',None),'name':getattr(q.from_user,'full_name',None)}}; save_state(q.from_user.id,ns)
        await q.message.reply_text(f'🟢 Pedido #{ns["id"]:04d} abierto. Envía la primera foto.')
        return
    if data=='close_now':
        if state.get('status')!='ABIERTO':
            await q.message.reply_text(f'🔒 El pedido #{state["id"]:04d} ya está cerrado.'); return
        if state.get('pending'):
            await q.message.reply_text('⚠️ Primero termina el artículo anterior.'); return
        if not state.get('items'):
            await q.message.reply_text('⚠️ El pedido está vacío.'); return
        try:
            path=rebuild_output(state['id'])
            if not path or not os.path.exists(path): raise RuntimeError('No se generó el Excel final.')
            set_order_status(state['id'],'CERRADO')
            state['status']='CERRADO'; save_state(q.from_user.id,state)
            await q.message.reply_text(f'🔒 Pedido #{state["id"]:04d} cerrado.\n📦 Artículos: {len(state["items"])}\n📄 Enviando Excel final…')
            with open(path,'rb') as f: await q.message.reply_document(document=f,filename=os.path.basename(path),caption=f'📄 PEDIDO_PROVEEDOR_{state["id"]:04d}.xlsx — pedido cerrado.')
        except Exception as e:
            try:
                state['status']='ABIERTO'; save_state(q.from_user.id,state); set_order_status(state['id'],'ABIERTO')
            except Exception: pass
            await q.message.reply_text(f'⚠️ No pude cerrar el pedido. Sigue ABIERTO.\n\nDetalle: {e}')
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
        state['pending']=None; save_state(q.from_user.id,state); await q.edit_message_text('❌ Artículo cancelado. Puedes enviar otra foto.'); return
    if data=='edit':
        await q.message.reply_text('✏️ Escribe la corrección. Ej.: «talla XL, nombre Sergio, número 10, parche Liga».'); return

async def ping(update,context): await update.message.reply_text(f'V9 activo. Vision: {"OK" if OPENAI_API_KEY else "FALTA OPENAI_API_KEY"} ({VISION_MODEL}). Excel: {"OK" if EXCEL_FILE.exists() else "FALTA"}.')

def flask_thread(): app.run(host='0.0.0.0',port=int(os.environ.get('PORT','10000')))

@app.get('/')
def health():
    exists=EXCEL_FILE.exists(); n=len(catalog_rows()) if exists else 0
    return f'Bot V9 activo - Excel: {"OK" if exists else "FALTA"} - productos: {n} - vision: {"OK" if OPENAI_API_KEY else "FALTA OPENAI_API_KEY"}',200 if exists else 503

def main():
    if not TOKEN: raise RuntimeError('Falta TELEGRAM_BOT_TOKEN')
    if not EXCEL_FILE.exists(): raise RuntimeError(f'No existe el Excel del catálogo: {EXCEL_FILE}')
    if not zipfile.is_zipfile(EXCEL_FILE):
        raise RuntimeError(f'El archivo Excel no es un XLSX válido (no es un ZIP): {EXCEL_FILE}. Reemplázalo por el SISTEMA_PEDIDOS_SERGIO_V8_1.xlsx corregido.')
    try:
        wb_check=load_workbook(str(EXCEL_FILE), read_only=True, data_only=True)
        required={'CATALOGO','PEDIDOS','MATCH_LOG','JUGADORES'}
        missing=required-set(wb_check.sheetnames)
        wb_check.close()
    except Exception as e:
        raise RuntimeError(f'No se pudo abrir el Excel {EXCEL_FILE}: {e}')
    if missing: raise RuntimeError(f'Faltan hojas obligatorias en el Excel: {sorted(missing)}')
    if not catalog_rows(): raise RuntimeError('El CATALOGO está vacío o no se pudo leer.')
    print(f'V9.3 MULTIUSER listo | Excel={EXCEL_FILE.name} | productos={len(catalog_rows())} | Vision={VISION_MODEL} | API_KEY={"OK" if OPENAI_API_KEY else "FALTA"} | estado=por_user_id', flush=True)
    threading.Thread(target=flask_thread,daemon=True).start()
    application=Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler('start',start)); application.add_handler(CommandHandler('nuevo',new_order)); application.add_handler(CommandHandler('cerrar',close_order));
    application.add_handler(CallbackQueryHandler(callbacks)); application.add_handler(MessageHandler(filters.PHOTO,handle_photo)); application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,handle_text));
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__=='__main__': main()
