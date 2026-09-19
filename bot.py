import os
import threading
from flask import Flask
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")

app = Flask(__name__)

@app.get("/")
def health():
    return "Bot activo", 200

orders = {}
next_order = 1

def get_open_order(chat_id):
    return orders.get(chat_id)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hola. Soy tu bot de pedidos.\n\n"
        "Usa «nuevo pedido» para comenzar.\n"
        "Luego envíame las camisetas en texto.\n"
        "Usa «cierro pedido» para cerrar.\n"
        "Usa «muestra el pedido» para verlo."
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
            lines.append(f"{i}. {item}")
        await update.message.reply_text("\n".join(lines))
        return

    order = get_open_order(chat_id)
    if not order or order["status"] != "ABIERTO":
        await update.message.reply_text(
            "⚠️ No hay un pedido abierto.\n"
            "Escribe «nuevo pedido» para comenzar uno."
        )
        return

    order["items"].append(text)
    await update.message.reply_text(
        f"✅ Agregado al pedido #{order['id']}:\n{text}"
    )

def run_flask():
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

def main():
    if not TOKEN:
        raise RuntimeError("Falta la variable TELEGRAM_BOT_TOKEN")

    threading.Thread(target=run_flask, daemon=True).start()

    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("Bot iniciado.")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
