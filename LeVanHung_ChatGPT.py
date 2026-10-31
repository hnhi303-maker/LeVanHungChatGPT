import os
import logging
from openai import OpenAI
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, MessageHandler, filters
import openpyxl
from docx import Document

# Bật tính năng ghi log để theo dõi lỗi nếu có
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

# Khởi tạo OpenAI Client với API Key bảo mật từ biến môi trường
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# Hàm xử lý toàn năng cho mọi tin nhắn và yêu cầu tạo file
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    chat_id = update.message.chat_id

    try:
        # 1. TÍNH NĂNG TỰ ĐỘNG TẠO FILE EXCEL (.xlsx)
        if "excel" in user_text.lower() or "bảng tính" in user_text.lower() or "khối lượng" in user_text.lower():
            await context.bot.send_message(chat_id=chat_id, text="Đang xử lý tạo bảng tính Excel cho anh, chờ chút nhé...")
            
            # Tạo file Excel mẫu bằng openpyxl (Bạn có thể tinh chỉnh cấu trúc bảng ở đây)
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
            
            # Gửi file trực tiếp về Telegram
            await update.message.reply_document(
                document=open(file_path, 'rb'), 
                caption="Dạ, đây là file Excel tổng hợp khối lượng của anh!"
            )
            os.remove(file_path) # Xóa file tạm trên server
            return

        # 2. TÍNH NĂNG TỰ ĐỘNG TẠO FILE WORD (.docx)
        if "word" in user_text.lower() or "văn bản" in user_text.lower() or "báo cáo" in user_text.lower():
            await context.bot.send_message(chat_id=chat_id, text="Đang soạn thảo văn bản Word cho anh...")
            
            doc = Document()
            doc.add_heading('BÁO CÁO CÔNG TÍCH & TRIỂN KHAI KỸ THUẬT', 0)
            doc.add_paragraph(f'Nội dung yêu cầu từ hệ thống: {user_text}')
            doc.add_paragraph('Báo cáo chi tiết các hạng mục thi công và giám sát công trình.')
            
            file_path = "BaoCao_KyThuat.docx"
            doc.save(file_path)
            
            await update.message.reply_document(
                document=open(file_path, 'rb'), 
                caption="Dạ, đây là file Word báo cáo của anh!"
            )
            os.remove(file_path)
            return

        # 3. TÍNH NĂNG CHAT AI THÔNG MINH (Dùng mô hình GPT-4o-mini hoặc GPT-4o)
        response = client.chat.completions.create(
            model="gpt-4o-mini", # Model thông minh, nhanh và tiết kiệm chi phí
            messages=[
                {"role": "system", "content": "Bạn là một trợ lý AI kỹ thuật xây dựng và văn phòng cực kỳ chuyên nghiệp, thành thạo tiếng Việt, chuyên gia xử lý số liệu, tính toán và lập văn bản."},
                {"role": "user", "content": user_text}
            ],
            temperature=0.3
        )
        
        reply_text = response.choices[0].message.content
        await update.message.reply_text(reply_text)

    except Exception as e:
        await update.message.reply_text(f"Đã xảy ra lỗi khi xử lý yêu cầu: {str(e)}")

if __name__ == '__main__':
    TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not TOKEN:
        raise ValueError("Chưa thiết lập biến môi trường TELEGRAM_BOT_TOKEN!")
        
    app = ApplicationBuilder().token(TOKEN).build()
    
    # Lắng nghe tin nhắn văn bản thông thường (Tuân thủ Group Privacy)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    print("Bot ChatGPT toàn năng đang chạy...")
    app.run_polling()