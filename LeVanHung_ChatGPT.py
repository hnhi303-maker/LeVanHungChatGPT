import os
import logging
from openai import OpenAI
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, filters
import openpyxl
from docx import Document
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

# Bật tính năng ghi log
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Khởi tạo OpenAI Client với API Key từ biến môi trường
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# --- TẠO WEB SERVER NHỎ ĐỂ RENDER KHÔNG BỊ TIMEOUT (Gói Free Web Service) ---
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot is alive and running!")

def run_http_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(('0.0.0.0', port), HealthCheckHandler)
    server.serve_forever()
# ------------------------------------------------------------------------

# 1. Lệnh /excel: Tự động tạo file Excel tổng hợp khối lượng
async def excel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    await context.bot.send_message(chat_id=chat_id, text="Đang xử lý tạo bảng tính Excel tổng hợp khối lượng cho anh, chờ chút nhé...")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TongHopKhoiLuong"
    ws['A1'] = "STT"
    ws['B1'] = "Hạng mục công việc"
    ws['C1'] = "Đơn vị"
    ws['D1'] = "Khối lượng"
    
    ws['A2'] = 1
    ws['B2'] = "Đắp cát san nền bãi vật liệu"
    ws['C2'] = "m3"
    ws['D2'] = 1500.5
    
    file_path = "BangTinh_KhoiLuong.xlsx"
    wb.save(file_path)
    
    await update.message.reply_document(
        document=open(file_path, 'rb'), 
        caption="Dạ, đây là file Excel tổng hợp khối lượng của anh!"
    )
    os.remove(file_path)

# 2. Lệnh /word: Soạn thảo văn bản và báo cáo Word
async def word_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    await context.bot.send_message(chat_id=chat_id, text="Đang soạn thảo văn bản và báo cáo Word cho anh...")
    
    doc = Document()
    doc.add_heading('BÁO CÁO CÔNG TÍCH & TRIỂN KHAI KỸ THUẬT', 0)
    doc.add_paragraph('Báo cáo chi tiết các hạng mục thi công, giám sát hạ tầng và kỹ thuật xây dựng.')
    
    file_path = "BaoCao_KyThuat.docx"
    doc.save(file_path)
    
    await update.message.reply_document(
        document=open(file_path, 'rb'), 
        caption="Dạ, đây là file Word báo cáo của anh!"
    )
    os.remove(file_path)

# 3. Lệnh /help: Hướng dẫn sử dụng
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "🤖 **Trợ lý AI Kỹ thuật Xây dựng & Quản lý**\n\n"
        "Các lệnh hỗ trợ:\n"
        "• /chat [nội dung] - Trò chuyện và tư vấn kỹ thuật AI\n"
        "• /excel - Tự động tạo file Excel tính toán khối lượng\n"
        "• /word - Soạn thảo văn bản báo cáo kỹ thuật Word\n"
        "• /help - Hiển thị bảng hướng dẫn này"
    )
    await update.message.reply_text(help_text, parse_mode="Markdown")

# 4. Lệnh /chat hoặc phản hồi tin nhắn văn bản thông thường bằng AI
async def chat_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = " ".join(context.args) if context.args else update.message.text
    if not user_text or user_text.startswith('/chat'):
        user_text = "Xin chào, hãy giới thiệu về các chức năng hỗ trợ kỹ thuật xây dựng của bạn."

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Bạn là một trợ lý AI kỹ thuật xây dựng và văn phòng cực kỳ chuyên nghiệp, thành thạo tiếng Việt, chuyên gia xử lý số liệu, tính toán và lập văn bản."},
                {"role": "user", "content": user_text}
            ],
            temperature=0.3
        )
        reply_text = response.choices[0].message.content
        await update.message.reply_text(reply_text)
    except Exception as e:
        await update.message.reply_text(f"Đã xảy ra lỗi kết nối AI: {str(e)}")

if __name__ == '__main__':
    TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not TOKEN:
        raise ValueError("Chưa thiết lập biến môi trường TELEGRAM_BOT_TOKEN!")
        
    # Khởi chạy HTTP Server giả lập trên luồng phụ để qua mặt Health Check của Render
    threading.Thread(target=run_http_server, daemon=True).start()
    
    app = ApplicationBuilder().token(TOKEN).build()
    
    # Đăng ký các Handler cho từng lệnh cụ thể
    app.add_handler(CommandHandler("excel", excel_command))
    app.add_handler(CommandHandler("word", word_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("chat", chat_command))
    
    # Lắng nghe các tin nhắn văn bản thông thường khác để chat với AI
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, chat_command))
    
    print("Bot ChatGPT toàn năng đang chạy...")
    app.run_polling()