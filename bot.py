import os
import threading
from flask import Flask
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
app = Flask(__name__)

@app.get("/")
def health():
    return "Bot activo - V2", 200

orders = {}
next_order = 1

def get_open_order(chat_id):
    return orders.get(chat_id)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hola. Soy tu bot de pedidos V2.\n\n"
        "🆕 nuevo pedido — abre un pedido\n"
        "📸 Envía una foto — la guardo para identificarla\n"
        "👕 Después envía talla, nombre, número y parche\n"
        "📦 muestra el pedido — muestra lo registrado\n"
        "🔒 cierro pedido — cierra el pedido"
    )

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    order = get_open_order(chat_id)
    if not order or order["status"] != "ABIERTO":
        await update.message.reply_text(
            "⚠️ No hay un pedido abierto. Escribe «nuevo pedido» primero."
        )
        return

    photo = update.message.photo[-1]
    context.user_data["pending_photo"] = {
        "file_id": photo.file_id,
        "caption": update.message.caption or ""
    }
    caption = (update.message.caption or "").strip()

    if caption:
        await update.message.reply_text(
            "📸 Foto recibida y asociada al pedido.\n\n"
            f"📝 Texto de la foto: {caption}\n\n"
            "Ahora envíame los datos que falten, por ejemplo: «Talla L, Yamal 10, parche Liga»."
        )
    else:
        await update.message.reply_text(
            "📸 Foto recibida y asociada al pedido.\n\n"
            "Ahora envíame los datos, por ejemplo:\n"
            "«Talla L, Yamal 10, parche Liga»"
        )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global next_order
    if not update.message or not update.message.text:
        return

    text = update.message.text.strip()
    low = text.lower()
    chat_id = update.effective_chat.id

    if low in ("nuevo pedido", "/nuevo"):
        order_id = f"{next_order:04d}"
        next_order += 1
        orders[chat_id] = {"id": order_id, "status": "ABIERTO", "items": []}
        context.user_data.pop("pending_photo", None)
        await update.message.reply_text(f"🟢 Pedido #{order_id} abierto.")
        return

    if low in ("cierro pedido", "/cerrar"):
        order = get_open_order(chat_id)
        if not order:
            await update.message.reply_text("⚠️ No hay un pedido abierto.")
            return
        order["status"] = "CERRADO"
        await update.message.reply_text(
            f"🔴 Pedido #{order['id']} cerrado.\n"
            f"Total de artículos: {len(order['items'])}"
        )
        return

    if low in ("muestra el pedido", "muéstrame el pedido", "muestrame el pedido", "/pedido"):
        order = get_open_order(chat_id)
        if not order:
            await update.message.reply_text("⚠️ No hay un pedido activo.")
            return
        if not order["items"]:
            await update.message.reply_text(f"📦 Pedido #{order['id']} está vacío.")
            return
        lines = [f"📦 Pedido #{order['id']} — {order['status']}"]
        for i, item in enumerate(order["items"], 1):
            if isinstance(item, dict):
                mark = "📸 " if item.get("photo") else ""
                lines.append(f"{i}. {mark}{item['text']}")
            else:
                lines.append(f"{i}. {item}")
        await update.message.reply_text("\n".join(lines))
        return

    order = get_open_order(chat_id)
    if not order or order["status"] != "ABIERTO":
        await update.message.reply_text(
            "⚠️ No hay un pedido abierto. Escribe «nuevo pedido»."
        )
        return

    pending_photo = context.user_data.pop("pending_photo", None)
    item = {
        "text": text,
        "photo": pending_photo is not None,
        "photo_file_id": pending_photo["file_id"] if pending_photo else None,
        "photo_caption": pending_photo["caption"] if pending_photo else "",
        "status": "PENDIENTE_IDENTIFICACION"
    }
    order["items"].append(item)

    if pending_photo:
        await update.message.reply_text(
            "✅ Pedido registrado con foto.\n\n"
            f"📦 Pedido #{order['id']}\n"
            f"👕 Datos: {text}\n"
            "📸 Foto: guardada\n"
            "🔎 Estado: pendiente de identificación\n\n"
            "La siguiente etapa conectará la foto con el catálogo maestro."
        )
    else:
        await update.message.reply_text(f"✅ Agregado al pedido #{order['id']}:\n{text}")

def run_flask():
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

def main():
    if not TOKEN:
        raise RuntimeError("Falta la variable TELEGRAM_BOT_TOKEN")
    threading.Thread(target=run_flask, daemon=True).start()
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot V2 iniciado.")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
