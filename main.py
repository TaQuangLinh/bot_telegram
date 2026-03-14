import logging
import os
import re
from telethon import TelegramClient, events, types
from utils_scraper import scrape_mercari_data, create_excel_with_image
from func_cop_link import upload_to_sheets

API_ID = 35396043
API_HASH = 'b2d225f192c903a4c5485bcbead411f1'
BOT_TOKEN = '8552090761:AAEqGJHo2a7DP56H-WmJU8I2yIkLLFss98o'
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1q3b9YSAuKgJAnsfvI-vNgaaFQRDLHYmBrwd8RNGRB7I/edit?gid=0#gid=0' # test tool telegram
# GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1-A-4S3h6GwFEIPKAzueFvAq9O330sl2gL9t7TwpJOHk/edit?gid=0#gid=0' # TrungCheck


ALLOWED_GROUP_NAME = ["cop link, tên & cop tracking", "test tools"]

client = TelegramClient('bot_session', API_ID, API_HASH).start(bot_token=BOT_TOKEN)
logging.basicConfig(level=logging.INFO)
user_states = {}


# ... (Các phần cấu hình giữ nguyên) ...

@client.on(events.NewMessage)
async def handler(event):
    if not event.is_group: return
    chat = await event.get_chat()
    if chat.title not in ALLOWED_GROUP_NAME: return

    chat_id = event.chat_id
    text = event.raw_text.strip() if event.raw_text else ""

    # BẮT ĐẦU PHIÊN
    if text.lower() == 'cop link':
        if chat_id in user_states:
            await event.respond(
                "**Hệ thống đang trong một phiên làm việc!**\nVui lòng hoàn thành phiên hoặc gõ `kết thúc` để đóng phiên hiện tại.")
            return

        user_states[chat_id] = {'status': 'COLLECTING', 'data': []}
        await event.respond(
            "**Bắt đầu phiên cop link.**\n- Hãy gửi theo dạng: `[Link] [Tên khách]`\n- Gửi xong hết gõ: `hết rồi`\n- Hủy phiên, làm lại gõ: `kết thúc`")
        return

    # KẾT THÚC/HỦY PHIÊN
    if text.lower() == 'kết thúc':
        if chat_id in user_states:
            del user_states[chat_id]
            await event.respond("**Đã kết thúc phiên làm việc hiện tại.**")
        return

    if chat_id not in user_states: return
    state = user_states[chat_id]

    # THU THẬP LINK
    if state['status'] == 'COLLECTING':
        if text.lower() == 'hết rồi':
            if not state['data']:
                await event.respond("❌ Chưa có dữ liệu.");
                return

            await event.respond("Đang xử lý dữ liệu bạn gửi...")
            results = await scrape_mercari_data(state['data'])
            output_file = f"ketqua_{chat_id}.xlsx"
            create_excel_with_image(results, output_file)

            user_states[chat_id]['status'] = 'WAITING_CONFIRM'
            user_states[chat_id]['file'] = output_file
            user_states[chat_id]['raw_results'] = results  # Lưu lại để so sánh

            await client.send_file(chat_id, output_file, caption=(
                "🔍 Hãy kiểm tra file Excel.\n- Click chuột phải vào file và chọn **Save As...** để tải file kiểm tra\n- Nếu đúng: gõ `đã chính xác`\n- Nếu sai: sửa file và gửi lại"
            ), force_document=True)
            return

        # CHỈ NHẬN TIN NHẮN BẮT ĐẦU BẰNG HTTP
        if text.lower().startswith('http'):
            parts = re.split(r'\s+', text, maxsplit=1)
            if len(parts) >= 2:
                state['data'].append({'url': parts[0].strip(), 'customer': parts[1].strip()})
        else:
            # Bỏ qua tin nhắn linh tinh, không phản hồi để tránh loãng group
            pass

    elif state['status'] == 'WAITING_CONFIRM':
        # Trường hợp người dùng gửi file .xlsx đã sửa
        if event.document and event.document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            # Lưu file người dùng gửi lên đè vào file cũ
            path = await event.download_media(file=f"edited_{chat_id}.xlsx")
            user_states[chat_id]['file'] = path  # Cập nhật đường dẫn file mới nhất
            await event.respond(
                "**Đã nhận file chỉnh sửa.**\nBây giờ hãy gõ `đã chỉnh sửa` để bot cop lên Google Sheet.")
            return

        if text.lower() in ['đã chính xác', 'đã chỉnh sửa']:
            await event.respond("**Đang đẩy dữ liệu lên Google Sheet...**")

            # Luôn lấy file từ state (có thể là file gốc ketqua_... hoặc file đã sửa edited_...)
            excel_file = state.get('file')

            if not excel_file or not os.path.exists(excel_file):
                await event.respond(
                    "❌ Không tìm thấy file dữ liệu. Vui lòng gửi lại file hoặc gõ `cop link` để làm lại.")
                return

            try:
                sheet_name, report = upload_to_sheets(excel_file, GOOGLE_SHEET_URL)
                final_msg = f"✅ **Đã hoàn tất!**\nĐã cop lên bảng: `{sheet_name}`\n\n**BÁO CÁO CHI TIẾT:**\n{report}"
                await event.respond(final_msg)
            except Exception as e:
                await event.respond(f"❌ Lỗi khi ghi Sheet: {e}")

            # Dọn dẹp tất cả file tạm liên quan đến phiên này
            if os.path.exists(f"ketqua_{chat_id}.xlsx"): os.remove(f"ketqua_{chat_id}.xlsx")
            if os.path.exists(f"edited_{chat_id}.xlsx"): os.remove(f"edited_{chat_id}.xlsx")
            del user_states[chat_id]



print(f"Bot đang chạy trong nhóm: {ALLOWED_GROUP_NAME}")
client.run_until_disconnected()