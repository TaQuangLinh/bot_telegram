import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pandas as pd
import re


def clean_price_to_int(price_val):
    """Xử lý triệt để mọi định dạng số từ Excel để tránh lỗi nhân 10"""
    if pd.isna(price_val): return None
    try:
        # 1. Chuyển về chuỗi và làm sạch khoảng trắng
        val_str = str(price_val).strip()

        # 2. Nếu là chuỗi rỗng hoặc N/A thì bỏ qua
        if val_str.lower() in ['nan', 'n/a', '']: return None

        # 3. Loại bỏ ký hiệu tiền tệ và dấu phẩy (dấu phân cách hàng nghìn)
        # Ví dụ: "¥4,500.0" -> "4500.0"
        val_str = val_str.replace('¥', '').replace(',', '')

        # 4. CHÌA KHÓA: Chuyển về float trước.
        # Float sẽ hiểu "4500.0" là 4500.0, còn "4500" là 4500.0
        num_float = float(val_str)

        # 5. Ép về số nguyên (int) để loại bỏ hoàn toàn phần thập phân
        return int(num_float)
    except Exception as e:
        print(f"Lỗi chuyển đổi giá ({price_val}): {e}")
        return None


def upload_to_sheets(file_path, sheet_url):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(sheet_url)
    sheet = spreadsheet.sheet1

    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]

    now = datetime.now().strftime("%Hh%M' ngày %d/%m/%Y")
    data_to_insert = []
    success_count = 0
    fail_details = []

    for index, row in df.iterrows():
        name_sp = str(row.get('Tên sản phẩm', '')).strip()
        cust = str(row.get('Tên khách', '')).strip()
        link = str(row.get('Link', '')).strip()
        raw_price = row.get('Giá tiền (JPY)', 'N/A')

        # CHUYỂN VỀ SỐ NGUYÊN TUYỆT ĐỐI
        numeric_price = clean_price_to_int(raw_price)

        is_valid = all([
            name_sp not in ["nan", "N/A", "Không tìm thấy / Lỗi", "LINK KHÔNG HỢP LỆ", ""],
            cust not in ["nan", "N/A", ""],
            numeric_price is not None
        ])

        if is_valid:
            note = ""
            if success_count == 0:
                note = f"bot bắt đầu điền link lúc {now}"

            line = ["", datetime.now().strftime("%d/%m/%Y"), cust.lower(), "", link, numeric_price, note]
            data_to_insert.append(line)
            success_count += 1
        else:
            fail_details.append(f"- Dòng {index + 2}: {cust} ({link[:100]})")

    if data_to_insert:
        data_to_insert[-1][6] = f"bot kết thúc điền link lúc {now}"
        # USER_ENTERED giúp Google Sheets tự định dạng ¥ và cho phép hàm SUM hoạt động
        sheet.append_rows(data_to_insert, value_input_option='USER_ENTERED')

    report = f"Tổng cộng: {len(df)} dòng\n✅ Thành công: {success_count} dòng\n❌ Thất bại: {len(df) - success_count} dòng"
    if fail_details:
        report += "\n\n**Chi tiết lỗi:**\n" + "\n".join(fail_details[:5])

    return spreadsheet.title, report