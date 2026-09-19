import os
import threading
import re
from datetime import datetime
from flask import Flask
from openpyxl import load_workbook
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
EXCEL_FILE = "SISTEMA_PEDIDOS_SERGIO_V1_V3.xlsx"
app = Flask(__name__)

orders = {}
next_order = 1

@app.get("/")
def health():
    return "Bot activo - V3", 200

def get_open_order(chat_id):
    return orders.get(chat_id)

def norm(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    replacements = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    return "".join(replacements.get(c,c) for c in s)

def catalog_rows():
    if not os.path.exists(EXCEL_FILE):
        return []
    wb = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    if "CATALOGO" not in wb.sheetnames:
        wb.close(); return []
    ws = wb["CATALOGO"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return []
    headers = [str(x).strip() if x is not None else "" for x in rows[0]]
    return [dict(zip(headers, r)) for r in rows[1:] if any(x is not None for x in r)]

def parse_request(text):
    t = norm(text)
    size = None
    for s in ("3xl","xxl","xl","l","m","s"):
        if re.search(rf"\b{re.escape(s)}\b", t):
            size = s.upper(); break
    team = None
    if "barcelona" in t or "barca" in t or "barça" in t:
        team = "Barcelona"
    elif "real madrid" in t or re.search(r"\brma\b", t):
        team = "Real Madrid"
    player = None
    if "yamal" in t or "lamine" in t:
        player = "Lamine Yamal"
    number = None
    m = re.search(r"(?:numero|número|num|#)\s*[:.]?\s*(\d{1,2})", t)
    if m: number = m.group(1)
    elif player == "Lamine Yamal" and re.search(r"\b10\b", t): number = "10"
    patch = "Liga" if "liga" in t or "laliga" in t else None
    edition = "Player Edition" if any(x in t for x in ("player","jugador","player edition")) else None
    return team, size, player, number, patch, edition

def find_candidates(text, limit=3):
    team, size, player, number, patch, edition = parse_request(text)
    scored = []
    for r in catalog_rows():
        if str(r.get("ACTIVO", "")).lower() not in ("", "true", "1", "si", "sí", "yes"):
            continue
        score = 0
        if team and norm(r.get("EQUIPO")) == norm(team): score += 50
        if player and norm(r.get("JUGADOR")) == norm(player): score += 20
        if number and str(r.get("NUMERO") or "") == number: score += 15
        if patch and norm(r.get("PARCHE")) == norm(patch): score += 10
        if edition and norm(r.get("TIPO_EDICION")) == norm(edition): score += 10
        if size and size in norm(r.get("TALLAS")): score += 5
        if score: scored.append((score, r))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:limit]

def candidate_text(r):
    return (f"{r.get('ID_PRODUCTO')} — {r.get('EQUIPO')} {r.get('TEMPORADA')} — "
            f"{r.get('TIPO_EDICION')} {r.get('MODELO')} — "
            f"{r.get('JUGADOR') or 'sin jugador'} {r.get('NUMERO') or ''} — "
            f"parche {r.get('PARCHE') or 'sin parche'}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hola. Soy tu bot de pedidos V3.\n\n"
        "🆕 nuevo pedido\n"
        "📸 Envía una foto y luego sus datos\n"
        "📦 muestra el pedido\n"
        "✏️ corrige la última\n"
        "🗑️ elimina la última\n"
        "🔒 cierro pedido"
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    order = get_open_order(chat_id)
    if not order or order["status"] != "ABIERTO":
        await update.message.reply_text("⚠️ No hay un pedido abierto. Escribe «nuevo pedido» primero.")
        return
    photo = update.message.photo[-1]
    context.user_data["pending_photo"] = {"file_id": photo.file_id, "caption": update.message.caption or ""}
    caption = (update.message.caption or "").strip()
    if caption:
        await update.message.reply_text(f"📸 Foto recibida.\n📝 {caption}\n\nAhora envíame talla, nombre, número y parche.")
    else:
        await update.message.reply_text("📸 Foto recibida.\n\nAhora envíame los datos, por ejemplo: «Barcelona talla L, Yamal 10, parche Liga». ")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global next_order
    if not update.message or not update.message.text: return
    text = update.message.text.strip(); low = norm(text); chat_id = update.effective_chat.id
    if low in ("nuevo pedido", "/nuevo"):
        order_id = f"{next_order:04d}"; next_order += 1
        orders[chat_id] = {"id": order_id, "status": "ABIERTO", "items": []}
        context.user_data.clear()
        await update.message.reply_text(f"🟢 Pedido #{order_id} abierto."); return
    if low in ("cierro pedido", "/cerrar"):
        order = get_open_order(chat_id)
        if not order: await update.message.reply_text("⚠️ No hay un pedido abierto."); return
        order["status"] = "CERRADO"
        await update.message.reply_text(f"🔴 Pedido #{order['id']} cerrado.\nTotal de artículos: {len(order['items'])}"); return
    if low in ("muestra el pedido", "muéstrame el pedido", "muestrame el pedido", "/pedido"):
        order = get_open_order(chat_id)
        if not order: await update.message.reply_text("⚠️ No hay un pedido activo."); return
        if not order["items"]: await update.message.reply_text(f"📦 Pedido #{order['id']} está vacío."); return
        lines=[f"📦 Pedido #{order['id']} — {order['status']}"]
        for i,item in enumerate(order["items"],1):
            c=item.get("catalog")
            ident=f" | {c.get('ID_PRODUCTO')}" if c else " | pendiente"
            lines.append(f"{i}. {'📸 ' if item.get('photo') else ''}{item['text']}{ident}")
        await update.message.reply_text("\n".join(lines)); return
    order = get_open_order(chat_id)
    if not order or order["status"] != "ABIERTO":
        await update.message.reply_text("⚠️ No hay un pedido abierto. Escribe «nuevo pedido»."); return
    if low in ("corrige la última", "corrige la ultima", "editar última", "editar la última"):
        if order["items"]: order["items"].pop()
        await update.message.reply_text("✏️ Último artículo eliminado. Envíame de nuevo los datos."); return
    if low in ("elimina la última", "elimina la ultima", "borrar última", "borrar la última"):
        if order["items"]: order["items"].pop()
        await update.message.reply_text("🗑️ Último artículo eliminado."); return
    # Confirmación de candidatos pendientes: 1, 2 o 3
    pending_candidates = context.user_data.get("pending_candidates")
    pending_item_index = context.user_data.get("pending_item_index")
    if pending_candidates and low in ("1", "2", "3"):
        selected_index = int(low) - 1
        if selected_index < len(pending_candidates):
            if pending_item_index is not None and 0 <= pending_item_index < len(order["items"]):
                item = order["items"][pending_item_index]
            else:
                item = order["items"][-1]
            candidate = pending_candidates[selected_index][1]
            item["catalog"] = candidate
            item["status"] = "IDENTIFICADO"
            context.user_data.pop("pending_candidates", None)
            context.user_data.pop("pending_item_index", None)
            await update.message.reply_text(
                "✅ Producto confirmado: {}\n\n👕 {}\n📏 Datos: {}\n📦 Agregado al pedido #{}".format(
                    candidate.get("ID_PRODUCTO"), candidate_text(candidate), item["text"], order["id"]
                )
            )
            return
        await update.message.reply_text("⚠️ Esa opción no existe. Responde con 1, 2 o 3.")
        return

    # Permite confirmar también escribiendo directamente el ID del producto.
    if pending_candidates and re.fullmatch(r"[A-Za-z0-9_-]+", text):
        wanted = norm(text)
        for _, candidate in pending_candidates:
            if norm(candidate.get("ID_PRODUCTO")) == wanted:
                if pending_item_index is not None and 0 <= pending_item_index < len(order["items"]):
                    item = order["items"][pending_item_index]
                else:
                    item = order["items"][-1]
                item["catalog"] = candidate
                item["status"] = "IDENTIFICADO"
                context.user_data.pop("pending_candidates", None)
                context.user_data.pop("pending_item_index", None)
                await update.message.reply_text(
                    "✅ Producto confirmado: {}\n\n👕 {}\n📏 Datos: {}\n📦 Agregado al pedido #{}".format(
                        candidate.get("ID_PRODUCTO"), candidate_text(candidate), item["text"], order["id"]
                    )
                )
                return

    pending_photo = context.user_data.pop("pending_photo", None)
    candidates = find_candidates(text)
    item = {"text": text, "photo": pending_photo is not None, "photo_file_id": pending_photo["file_id"] if pending_photo else None, "photo_caption": pending_photo["caption"] if pending_photo else "", "status":"PENDIENTE_IDENTIFICACION", "catalog": None}
    if len(candidates) == 1 and candidates[0][0] >= 60:
        item["catalog"] = candidates[0][1]; item["status"] = "IDENTIFICADO"
        order["items"].append(item)
        await update.message.reply_text("✅ Identificado y agregado al pedido.\n\n📦 Pedido #{}\n🏷️ {}\n📏 Datos: {}".format(order["id"], candidate_text(item["catalog"]), text)); return
    order["items"].append(item)
    if candidates:
        msg = "🔎 Encontré estas coincidencias:\n\n" + "\n".join(f"{i+1}. {candidate_text(r)} (score {s})" for i,(s,r) in enumerate(candidates))
        msg += "\n\nResponde con 1, 2 o 3 para confirmar."
        context.user_data["pending_candidates"] = candidates
        context.user_data["pending_item_index"] = len(order["items"]) - 1
        await update.message.reply_text(msg)
    else:
        await update.message.reply_text("⚠️ No encontré una coincidencia en el catálogo. El artículo quedó pendiente de identificación.")

def run_flask():
    port=int(os.environ.get("PORT","10000")); app.run(host="0.0.0.0", port=port)

def main():
    if not TOKEN: raise RuntimeError("Falta la variable TELEGRAM_BOT_TOKEN")
    threading.Thread(target=run_flask, daemon=True).start()
    application=Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot V3.1 iniciado."); application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__": main()
