import os, re, json, sqlite3, shutil, threading, tempfile
from datetime import datetime
from pathlib import Path
from copy import copy
from flask import Flask
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters

TOKEN = os.environ.get('TELEGRAM_BOT_TOKEN')
EXCEL_FILE = os.environ.get('EXCEL_FILE', 'SISTEMA_PEDIDOS_SERGIO_V1_V3.xlsx')
DB_FILE = os.environ.get('DB_FILE', 'bot_v5.sqlite3')
OUT_DIR = Path(os.environ.get('OUT_DIR', 'pedidos_finales'))
app = Flask(__name__)

# V5: catalog-driven Telegram order assistant.
# Notes: Telegram photos are stored by file_id. This build does not claim to perform
# computer-vision matching against Yupoo; matching is based on catalog + text/caption.

ALIASES = {
    'barca':'barcelona','barça':'barcelona','fc barcelona':'barcelona',
    'real':'real madrid','rma':'real madrid',
    'visitante':'visitante','away':'visitante','local':'local','home':'local',
    'jugador':'player edition','player':'player edition','player edition':'player edition',
    'fan':'fan edition','fan edition':'fan edition',
    'retro':'retro','clasica':'retro','clásica':'retro',
    'liga':'liga','laliga':'liga','la liga':'liga'
}

# ---------- persistence ----------
def db():
    c = sqlite3.connect(DB_FILE, check_same_thread=False)
    c.row_factory = sqlite3.Row
    c.execute('CREATE TABLE IF NOT EXISTS orders (chat_id INTEGER PRIMARY KEY, order_id INTEGER, status TEXT, data TEXT)')
    c.execute('CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)')
    if c.execute("SELECT 1 FROM meta WHERE k='next_order'").fetchone() is None:
        c.execute("INSERT INTO meta(k,v) VALUES('next_order','1')")
    c.commit(); return c

def load_state(chat_id):
    c=db(); r=c.execute('SELECT * FROM orders WHERE chat_id=?',(chat_id,)).fetchone(); c.close()
    return json.loads(r['data']) if r else None

def save_state(chat_id, state):
    c=db(); c.execute('INSERT OR REPLACE INTO orders(chat_id,order_id,status,data) VALUES(?,?,?,?)',(chat_id,state['id'],state['status'],json.dumps(state,ensure_ascii=False))); c.commit(); c.close()

def next_id():
    c=db(); r=c.execute("SELECT v FROM meta WHERE k='next_order'").fetchone(); n=int(r['v']); c.execute("UPDATE meta SET v=? WHERE k='next_order'",(str(n+1),)); c.commit(); c.close(); return n

# ---------- excel/catalog ----------
def norm(s):
    s='' if s is None else str(s).strip().lower()
    return s.translate(str.maketrans('áéíóúüñ','aeiouun'))

def catalog_rows():
    if not os.path.exists(EXCEL_FILE): return []
    wb=load_workbook(EXCEL_FILE,read_only=True,data_only=True)
    if 'CATALOGO' not in wb.sheetnames: wb.close(); return []
    ws=wb['CATALOGO']; rows=list(ws.iter_rows(values_only=True)); wb.close()
    if not rows: return []
    headers=[str(x).strip() if x is not None else '' for x in rows[0]]
    return [dict(zip(headers,r)) for r in rows[1:] if any(x is not None for x in r)]

def active(r):
    v=norm(r.get('ACTIVO'))
    return v in ('','true','1','si','yes','activo')

def vals(field):
    out=[]
    for r in catalog_rows():
        v=r.get(field)
        if v not in (None,'') and str(v) not in out: out.append(str(v))
    return out

def parse(text):
    t=norm(text)
    data={'size':None,'team':None,'edition':None,'model':None,'patch':None,'name':None,'number':None,'category':None,'quantity':1}
    for s in ('3xl','xxl','xl','l','m','s'):
        if re.search(rf'\b{s}\b',t): data['size']=s.upper(); break
    for a,b in ALIASES.items():
        if a in t:
            if b in ('barcelona','real madrid'): data['team']=b.title() if b=='real madrid' else 'Barcelona'
            elif b in ('fan edition','player edition','retro'): data['edition']=b
            elif b in ('local','visitante'): data['model']=b
            elif b=='liga': data['patch']='Liga'
    m=re.search(r'(?:nombre|name)\s*[:=]?\s*([\wÀ-ÿ.\'-]{2,30})(?=\s+(?:numero|número|num|#|parche|talla|size)\b|\s+\d{1,2}\b|$)',text,re.I)
    if m: data['name']=m.group(1).strip()
    m=re.search(r'(?:numero|número|num|#)\s*[:=]?\s*(\d{1,2})',text,re.I)
    if m: data['number']=m.group(1)
    elif data['name'] and data['name'].lower() in ('yamal','lamine yamal'): data['number']='10'
    m=re.search(r'(?:cantidad|qty|x)\s*[:=]?\s*(\d+)',t)
    if m: data['quantity']=max(1,int(m.group(1)))
    for c in vals('CATEGORIA'):
        if norm(c) in t: data['category']=c; break
    return data

def score_row(r,d):
    if not active(r): return -1
    s=0
    def eq(field,key,pts):
        nonlocal s
        if d.get(key) and norm(r.get(field))==norm(d[key]): s+=pts
    eq('EQUIPO','team',40); eq('TIPO_EDICION','edition',25); eq('MODELO','model',15); eq('PARCHE','patch',10); eq('CATEGORIA','category',10)
    if d.get('size') and d['size'] in norm(r.get('TALLAS')): s+=5
    if d.get('name') and norm(d['name']) in norm(r.get('JUGADOR')): s+=20
    if d.get('number') and str(r.get('NUMERO') or '')==d['number']: s+=15
    # exact product-id mention
    if r.get('ID_PRODUCTO') and norm(r['ID_PRODUCTO']) in norm(str(d)): s+=50
    return s

def candidates(text,limit=5):
    d=parse(text); scored=sorted([(score_row(r,d),r) for r in catalog_rows()],key=lambda x:x[0],reverse=True)
    return [(s,r) for s,r in scored if s>0][:limit]

def desc(r):
    return f"{r.get('EQUIPO','')} {r.get('TEMPORADA','')} — {r.get('TIPO_EDICION','')} — {r.get('MODELO','')} — {r.get('JUGADOR') or 'sin jugador'} {r.get('NUMERO') or ''} — {r.get('PARCHE') or 'sin parche'}"

def ensure_headers(ws, headers):
    existing={str(c.value).strip():i for i,c in enumerate(ws[1],1) if c.value}
    for h in headers:
        if h not in existing:
            col=ws.max_column+1; ws.cell(1,col,h); existing[h]=col
    return existing

def append_order_row(state,item):
    wb=load_workbook(EXCEL_FILE); ws=wb['PEDIDOS']
    col=ensure_headers(ws,['PEDIDO_ID','FECHA','CLIENTE','ESTADO','ID_PRODUCTO','MODELO','DESCRIPCION','TALLA','CANTIDAD','NOMBRE','NUMERO','PARCHE','IMAGEN_REFERENCIA','NOTAS','ORIGEN','CONFIANZA','TIPO_EDICION','NOMBRE_SOLICITADO','NUMERO_SOLICITADO','FOTO_TELEGRAM'])
    r=ws.max_row+1; c=item.get('catalog') or {}; d=item.get('parsed') or {}
    values={'PEDIDO_ID':state['id'],'FECHA':datetime.now(),'CLIENTE':item.get('cliente','Telegram'),'ESTADO':'ABIERTO','ID_PRODUCTO':c.get('ID_PRODUCTO'),'MODELO':c.get('MODELO'),'DESCRIPCION':desc(c),'TALLA':d.get('size',''),'CANTIDAD':item.get('quantity',1),'NOMBRE':c.get('JUGADOR') or '','NUMERO':c.get('NUMERO') or '','PARCHE':c.get('PARCHE') or '','IMAGEN_REFERENCIA':c.get('IMAGEN') or '','NOTAS':item.get('caption',''),'ORIGEN':'Telegram','CONFIANZA':item.get('confidence',0),'TIPO_EDICION':c.get('TIPO_EDICION') or d.get('edition') or '','NOMBRE_SOLICITADO':item.get('requested_name') or d.get('name') or '','NUMERO_SOLICITADO':item.get('requested_number') or d.get('number') or '','FOTO_TELEGRAM':item.get('photo_file_id') or ''}
    for k,v in values.items(): ws.cell(r,col[k],v)
    wb.save(EXCEL_FILE); wb.close()

def rebuild_output(order_id):
    OUT_DIR.mkdir(exist_ok=True); out=OUT_DIR/f'PEDIDO_PROVEEDOR_{order_id:04d}.xlsx'; shutil.copy2(EXCEL_FILE,out)
    wb=load_workbook(out); wp=wb['PEDIDO_PROVEEDOR']; wo=wb['PEDIDOS']
    # clear output data while retaining first-row template/header and formatting
    for row in range(2,wp.max_row+1):
        for col in range(1,wp.max_column+1): wp.cell(row,col).value=None
    wp._images=[]
    heads=[str(c.value).strip() if c.value else '' for c in wo[1]]; idx={h:i+1 for i,h in enumerate(heads) if h}
    rows=[]
    for rr in range(2,wo.max_row+1):
        if str(wo.cell(rr,idx.get('PEDIDO_ID',1)).value)==str(order_id) and str(wo.cell(rr,idx.get('ESTADO',4)).value) in ('ABIERTO','CERRADO'):
            rows.append({h:wo.cell(rr,i).value for h,i in idx.items()})
    outrow=2
    for item in rows:
        qty=int(item.get('CANTIDAD') or 1)
        for _ in range(max(1,qty)):
            wp.cell(outrow,1,item.get('MODELO') or '')
            wp.cell(outrow,2,(item.get('DESCRIPCION') or '').replace('—','-'))
            wp.cell(outrow,3,item.get('TALLA') or '')
            # supplier-friendly extra data in unused columns
            wp.cell(outrow,4,item.get('ID_PRODUCTO') or '')
            wp.cell(outrow,5,item.get('NOMBRE_SOLICITADO') or '')
            wp.cell(outrow,6,item.get('NUMERO_SOLICITADO') or '')
            wp.cell(outrow,7,item.get('PARCHE') or '')
            wp.row_dimensions[outrow].height=95
            img=item.get('IMAGEN_REFERENCIA')
            if img and os.path.exists(str(img)):
                try:
                    x=XLImage(str(img)); x.width=85; x.height=85; wp.add_image(x,f'H{outrow}')
                except Exception: pass
            outrow+=1
    wb.save(out); wb.close(); return str(out)

# ---------- telegram UI ----------
def kb_confirm(item):
    return InlineKeyboardMarkup([[InlineKeyboardButton('✅ Confirmar pedido',callback_data='confirm')],[InlineKeyboardButton('✏️ Editar datos',callback_data='edit'),InlineKeyboardButton('❌ Cancelar artículo',callback_data='cancel')]])

def button_options(state,item):
    d=item['parsed']; rows=[]
    # dynamic catalog options filtered by current team/category where possible
    cats=[]
    for r in catalog_rows():
        if d.get('team') and norm(r.get('EQUIPO'))!=norm(d['team']): continue
        if d.get('category') and norm(r.get('CATEGORIA'))!=norm(d['category']): continue
        if r.get('TIPO_EDICION') and str(r['TIPO_EDICION']) not in cats: cats.append(str(r['TIPO_EDICION']))
    if cats and not d.get('edition'):
        rows.append([InlineKeyboardButton(x,callback_data='set:edition:'+x) for x in cats[:3]])
    if d.get('edition')=='player edition' and not item.get('requested_name'):
        rows.append([InlineKeyboardButton('📝 Nombre personalizado',callback_data='ask:name')])
    if d.get('size') is None:
        sizes=[]
        for r in catalog_rows():
            if d.get('team') and norm(r.get('EQUIPO'))!=norm(d['team']): continue
            for x in str(r.get('TALLAS') or '').split('-'):
                if x and x not in sizes: sizes.append(x)
        if sizes: rows.append([InlineKeyboardButton(x,callback_data='set:size:'+x) for x in sizes[:6]])
    if d.get('patch') is None:
        patches=vals('PARCHE')
        if patches: rows.append([InlineKeyboardButton(x,callback_data='set:patch:'+x) for x in patches[:4]])
    rows.append([InlineKeyboardButton('1',callback_data='set:qty:1'),InlineKeyboardButton('2',callback_data='set:qty:2'),InlineKeyboardButton('3',callback_data='set:qty:3')])
    return InlineKeyboardMarkup(rows + [[InlineKeyboardButton('✅ Confirmar pedido',callback_data='confirm')],[InlineKeyboardButton('❌ Cancelar artículo',callback_data='cancel')]])

async def start(update,context):
    await update.message.reply_text('🤖 Bot V5 listo.\n\nUsa «Nuevo pedido» o escribe: nuevo pedido.\nEnvía una foto para iniciar automáticamente el siguiente artículo.')

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
    state['status']='CERRADO'; save_state(chat,state)
    path=rebuild_output(state['id'])
    await update.message.reply_text(f'🔒 Pedido #{state["id"]:04d} cerrado.\n📦 Artículos: {len(state["items"])}\n📄 Generando Excel final…')
    with open(path,'rb') as f: await update.message.reply_document(document=f,filename=os.path.basename(path),caption=f'📄 PEDIDO_PROVEEDOR_{state["id"]:04d}.xlsx — revisa y dime si hay que reabrirlo.')

async def reopen(update,context):
    chat=update.effective_chat.id; state=load_state(chat)
    if not state: await update.message.reply_text('No hay pedido para reabrir.'); return
    state['status']='ABIERTO'; state['pending']=None; save_state(chat,state)
    # mark existing rows open again
    try:
        wb=load_workbook(EXCEL_FILE); ws=wb['PEDIDOS']; heads=[str(c.value).strip() if c.value else '' for c in ws[1]]; ci=heads.index('PEDIDO_ID')+1; ce=heads.index('ESTADO')+1
        for r in range(2,ws.max_row+1):
            if str(ws.cell(r,ci).value)==str(state['id']): ws.cell(r,ce,'ABIERTO')
        wb.save(EXCEL_FILE); wb.close()
    except Exception: pass
    await update.message.reply_text(f'🔓 Pedido #{state["id"]:04d} reabierto. Puedes enviar otra foto o usar correcciones.')

async def show_order(update,context):
    state=load_state(update.effective_chat.id)
    if not state: await update.message.reply_text('No hay pedido.'); return
    lines=[f'📦 Pedido #{state["id"]:04d} — {state["status"]}', '']
    for i,x in enumerate(state['items'],1): lines.append(f'{i}. {desc(x["catalog"])} | Talla {x["parsed"].get("size") or "?"} | Nombre {x.get("requested_name") or x["catalog"].get("JUGADOR") or "—"} | # {x.get("requested_number") or x["catalog"].get("NUMERO") or "—"}')
    if state.get('pending'): lines.append('\n⏳ Artículo pendiente de confirmación.')
    await update.message.reply_text('\n'.join(lines))

async def handle_photo(update,context):
    state=load_state(update.effective_chat.id)
    if not state or state['status']!='ABIERTO': await update.message.reply_text('🔒 No hay pedido abierto. Usa «nuevo pedido».'); return
    # New photo ALWAYS starts the next article in the current order.
    if state.get('pending'):
        await update.message.reply_text('⚠️ Primero confirma o cancela el artículo anterior.')
        return
    caption=update.message.caption or ''
    state['pending']={'photo_file_id':update.message.photo[-1].file_id,'caption':caption,'text':caption,'parsed':parse(caption),'catalog':None}
    state['pending']['candidates']=[r for _,r in candidates(caption)]
    save_state(update.effective_chat.id,state)
    if state['pending']['candidates']:
        await update.message.reply_text('📸 Foto recibida. Encontré coincidencias en tu catálogo. Selecciona la correcta:',reply_markup=candidate_keyboard(state['pending']['candidates']))
    else:
        await update.message.reply_text('📸 Foto recibida. No tengo una coincidencia textual suficiente. Escribe debajo los datos del artículo y te muestro las opciones.',reply_markup=button_options(state,state['pending']))


def candidate_keyboard(rs):
    buttons=[]
    for i,r in enumerate(rs): buttons.append([InlineKeyboardButton(f'{r.get("EQUIPO")} · {r.get("TIPO_EDICION")} · {r.get("MODELO")} · {r.get("ID_PRODUCTO")}',callback_data=f'cand:{i}')])
    return InlineKeyboardMarkup(buttons)

async def text_item(update,context,text):
    state=load_state(update.effective_chat.id)
    if not state or state['status']!='ABIERTO': await update.message.reply_text('🔒 No hay pedido abierto. Usa «nuevo pedido».'); return
    if state.get('pending'):
        # augment pending item instead of creating a duplicate
        p=state['pending']; p['text']=(p.get('text','')+' '+text).strip(); p['caption']=p.get('caption','') or text; p['parsed']=parse(p['text']); rs=[r for _,r in candidates(p['text'])]; p['candidates']=rs; save_state(update.effective_chat.id,state)
        if rs: await update.message.reply_text('🔎 Actualicé los datos. Elige el modelo correcto:',reply_markup=candidate_keyboard(rs))
        else: await update.message.reply_text('No encontré coincidencia. Ajusta los datos o escribe más detalles.',reply_markup=button_options(state,p))
        return
    if not state['items'] and norm(text) in ('nuevo pedido','nuevo','abrir pedido'): await new_order(update,context); return
    p={'photo_file_id':None,'caption':'','text':text,'parsed':parse(text),'catalog':None}; rs=[r for _,r in candidates(text)]; p['candidates']=rs; state['pending']=p; save_state(update.effective_chat.id,state)
    if rs: await update.message.reply_text('🔎 Selecciona el producto del catálogo:',reply_markup=candidate_keyboard(rs))
    else: await update.message.reply_text('⚠️ No encontré coincidencia en el catálogo. Revisa el nombre/modelo o agrega ese producto al CATALOGO.',reply_markup=button_options(state,p))

async def handle_text(update,context):
    t=norm(update.message.text); raw=update.message.text
    if t=='nuevo pedido': return await new_order(update,context)
    if t in ('cierro pedido','cerrar pedido'): return await close_order(update,context)
    if t in ('reabrir pedido','reabrir'): return await reopen(update,context)
    if t in ('muéstrame el pedido','muestrame el pedido','mostrar pedido','ver pedido'): return await show_order(update,context)
    state=load_state(update.effective_chat.id)
    if t=='corrige la ultima' or t=='corrige la última':
        if state and state['items']: state['pending']=state['items'].pop(); save_state(update.effective_chat.id,state); await update.message.reply_text('✏️ Último artículo devuelto a edición.',reply_markup=button_options(state,state['pending']))
        else: await update.message.reply_text('No hay artículo que corregir.')
        return
    if t=='elimina la ultima' or t=='elimina la última':
        if state and state['items']:
            state['items'].pop(); save_state(update.effective_chat.id,state); await update.message.reply_text('🗑️ Último artículo eliminado del pedido.')
        else: await update.message.reply_text('No hay artículo que eliminar.')
        return
    await text_item(update,context,raw)

async def callbacks(update,context):
    q=update.callback_query; await q.answer(); state=load_state(q.message.chat.id)
    if not state or not state.get('pending'): return
    p=state['pending']; data=q.data
    if data.startswith('cand:'):
        i=int(data.split(':')[1]); rs=p.get('candidates',[])
        if i>=len(rs): return
        p['catalog']=rs[i]; p['confidence']=100
        d=p['parsed']; c=p['catalog']
        # Official player validation: if requested name is an official-player-like request but catalog does not contain it, don't silently label it official.
        requested=d.get('name'); official=c.get('JUGADOR')
        explicit_custom = bool(re.search(r'\b(?:nombre|name)\s*[:=]', p.get('text',''), re.I))
        if explicit_custom:
            p['requested_name']=requested; p['requested_number']=d.get('number')
        elif requested and official and norm(requested)!=norm(official):
            await q.edit_message_text(f'⚠️ El catálogo no valida «{requested}» como jugador oficial de este producto.\n\nSi quieres que sea personalización, escribe: «nombre {requested} número {d.get("number") or "?"}». Si debe ser jugador oficial, selecciona otro producto.')
            state['pending']=p; save_state(q.message.chat.id,state); return
        elif requested and not official:
            await q.edit_message_text(f'⚠️ «{requested}» no está registrado como jugador oficial en este producto.\n\nSi es personalización, escribe «nombre {requested}» y confirma nuevamente.')
            state['pending']=p; save_state(q.message.chat.id,state); return
        await q.edit_message_text(f'📋 Producto: {desc(c)}\nTalla: {d.get("size") or "—"}\nNombre solicitado: {p.get("requested_name") or "—"}\nNúmero solicitado: {p.get("requested_number") or "—"}\nParche: {d.get("patch") or c.get("PARCHE") or "—"}\n\nConfirma para agregarlo al pedido.',reply_markup=kb_confirm(p)); state['pending']=p; save_state(q.message.chat.id,state); return
    if data=='confirm':
        p['parsed']=parse(p.get('text','')); state['items'].append(p); state['pending']=None; save_state(q.message.chat.id,state)
        try: append_order_row(state,p)
        except Exception as e: await q.message.reply_text(f'⚠️ No pude guardar el artículo en Excel: {e}'); return
        await q.edit_message_text(f'✅ Artículo agregado al pedido #{state["id"]:04d}.\n📦 Total confirmado: {len(state["items"])}\n📸 Puedes enviar otra foto para el siguiente artículo.')
    elif data=='cancel': state['pending']=None; save_state(q.message.chat.id,state); await q.edit_message_text('❌ Artículo cancelado. Puedes enviar otra foto.')
    elif data=='edit': await q.message.reply_text('Escribe la corrección (ej.: «talla XL, nombre Sergio, número 10, parche Liga»).')
    elif data.startswith('set:'):
        _,field,val=data.split(':',2)
        if field=='size': p['parsed']['size']=val
        elif field=='edition': p['parsed']['edition']=val
        elif field=='patch': p['parsed']['patch']=val
        elif field=='qty': p['quantity']=int(val)
        p['candidates']=[r for _,r in candidates(p.get('text','')+' '+val)]
        state['pending']=p; save_state(q.message.chat.id,state); await q.message.edit_reply_markup(reply_markup=button_options(state,p))

async def ping(update,context): await update.message.reply_text('V5 activo.')

def flask_thread(): app.run(host='0.0.0.0',port=int(os.environ.get('PORT','10000')))

@app.get('/')
def health(): return 'Bot V5 activo',200

def main():
    if not TOKEN: raise RuntimeError('Falta TELEGRAM_BOT_TOKEN')
    threading.Thread(target=flask_thread,daemon=True).start()
    application=Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler('start',start)); application.add_handler(CommandHandler('nuevo',new_order)); application.add_handler(CommandHandler('cerrar',close_order)); application.add_handler(CallbackQueryHandler(callbacks)); application.add_handler(MessageHandler(filters.PHOTO,handle_photo)); application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,handle_text)); application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__=='__main__': main()
