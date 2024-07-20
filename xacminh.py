import requests
from bs4 import BeautifulSoup
import pandas as pd
from unidecode import unidecode

def parse_date(date_str):
    date_formats = ['%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y']
    for fmt in date_formats:
        try:
            return pd.to_datetime(date_str, format=fmt)
        except ValueError:
            continue
    raise ValueError(f"Định dạng ngày không hợp lệ: {date_str}")

def create_username_and_password(ho_va_ten, ngay_thang_nam_sinh):
    ho_va_ten = unidecode(ho_va_ten)
    ten = ho_va_ten.split()[-1].lower()
    
    if isinstance(ngay_thang_nam_sinh, str):
        ngay_thang_nam_sinh = parse_date(ngay_thang_nam_sinh)
    ddmmyy = ngay_thang_nam_sinh.strftime('%d%m%y')
    username = f"{ten}{ddmmyy}"
    password = f"{username}1"
    
    return username, password

def generate_next_username_and_password(username, attempt):
    if attempt == 0:
        return username, f"{username}1"
    next_username = f"{username}{chr(96 + attempt)}"
    next_password = f"{next_username}1"
    return next_username, next_password

def extract_data(ho_va_ten, ngay_thang_nam_sinh, ngay_thi, max_attempts=27):
    base_username, base_password = create_username_and_password(ho_va_ten, ngay_thang_nam_sinh)
    
    for attempt in range(max_attempts):
        current_username, current_password = generate_next_username_and_password(base_username, attempt)
        url_authenticate = "https://api.certiport.com/authentication/Authentication/authenticate"
        payload = {
            "username": current_username,
            "password": current_password
        }

        try:
            response = requests.post(url_authenticate, json=payload)

            if response.status_code == 200:
                response_data = response.json()
                user_display_name = response_data.get('UserDisplayName')
                print(f"Tìm thấy tài khoản:{user_display_name}")
                session_id = response_data.get('PortalUserSessionID')
                name, dob = user_display_name.rsplit(' ', 1)
                dob_api = pd.to_datetime(dob, format='%d%b%Y')
                ho_va_ten_no_diacritics = unidecode(ho_va_ten).lower().strip().replace(" ", "")
                user_display_name_no_diacritics = unidecode(name).lower().strip().replace(" ", "")

                if ho_va_ten_no_diacritics == user_display_name_no_diacritics and ngay_thang_nam_sinh == dob_api:
                    if session_id:
                        return process_login(session_id, current_username, current_password, ho_va_ten, ngay_thi)
                    else:
                        print("Không tìm thấy PortalUserSessionID trong kết quả trả về")
                    break
            else:
                print(f"Lỗi {response.status_code}. {current_username} - {current_password}")
        except requests.RequestException as e:
            print(f"Yêu cầu HTTP thất bại: {e}")
    else:
        print("Không tìm thấy username phù hợp sau khi thử lại nhiều lần.")

    return None, format_error(None, None, ho_va_ten, "Không tìm thấy username phù hợp")

def process_login(session_id, current_username, current_password, ho_va_ten, ngay_thi):
    url_login_redirect = f"https://www.certiport.com/Portal/SSL/LoginRedirect.aspx?sessionId={session_id}"
    response_redirect = requests.get(url_login_redirect)
    
    if response_redirect.status_code == 200:
        print(f"Đăng nhập thành công: {current_username}")
        html_content = response_redirect.text
        soup = BeautifulSoup(html_content, 'html.parser')
        span = soup.find('span', string="IC3 Digital Literacy Certification")
        lang_eng = soup.find('td', string="Exam")
        
        if span:
            target_table = span.find_next('table')
            if target_table:
                print("Đã tìm thấy kết quả thi")
                results = extract_results(target_table, lang_eng, ngay_thi)
                if results:
                    return format_results(results, current_username, current_password, ho_va_ten)
                else:
                    print("Không có ngày thi trùng khớp")
                    return None, format_error(current_username, current_password, ho_va_ten, "Không có ngày thi trùng khớp")
            else:
                print("Tài khoản cần dò thủ công")
                return None, format_error(current_username, current_password, ho_va_ten, "Tài khoản cần dò thủ công")
        else:
            print("Không tìm thấy kết quả")
            return None, format_error(current_username, current_password, ho_va_ten, "Không tìm thấy kết quả")
    else:
        print(f"Lỗi {response_redirect.status_code}: {response_redirect.text}")
        return None, format_error(current_username, current_password, ho_va_ten, f"Lỗi {response_redirect.status_code}: {response_redirect.text}")

def extract_results(target_table, lang_eng, ngay_thi):
    results = []
    rows = target_table.find_all('tr')
    for row in rows:
        cells = row.find_all(['td', 'th'])
        cell_texts = [cell.get_text(strip=True) for cell in cells]
        if len(cell_texts) > 1:
            results.append(cell_texts)
    
    exam_dates = []
    for row in results[1:]:
        if lang_eng:
            exam_dates.append(pd.to_datetime(row[1], format='%m/%d/%Y'))
        else:
            exam_dates.append(parse_date(row[1]))
    
    if any(date == ngay_thi for date in exam_dates):
        return results[1:]
    return None

def format_results(results, current_username, current_password, ho_va_ten):
    output = {
        "Họ và Tên": ho_va_ten,
        "Username": current_username,
        "Mật khẩu": current_password,
        "Kết quả thi": results,
        "Ghi chú": ""
    }
    
    df = pd.DataFrame(output["Kết quả thi"], columns=["Bài thi", "Ngày", "Điểm", "Trạng thái", "ID Certiport", "Exam Group ID"])
    df = df.drop(columns=["ID Certiport", "Exam Group ID"])
    df.insert(0, "Họ và Tên", output["Họ và Tên"])
    df.insert(1, "Username", output["Username"])
    df.insert(2, "Mật khẩu", output["Mật khẩu"])
    df["Ghi chú"] = output["Ghi chú"]
    
    return df, None

def format_error(current_username, current_password, ho_va_ten, ghi_chu):
    error_info = {
        "Họ và Tên": ho_va_ten,
        "Username": current_username if current_username else "",
        "Mật khẩu": current_password if current_password else "",
        "Bài thi": "",
        "Ngày": "",
        "Điểm": "",
        "Trạng thái": "",
        "Ghi chú": ghi_chu
    }
    return pd.DataFrame([error_info])

def process_row(row):
    ho_va_ten = row.get('ho ten')
    ngay_thang_nam_sinh = row.get('ngay sinh')
    ngay_thi = row.get('ngay thi')
    
    if pd.isna(ho_va_ten) or pd.isna(ngay_thang_nam_sinh) or pd.isna(ngay_thi):
        print(f"Dữ liệu không hợp lệ: {ho_va_ten}, {ngay_thang_nam_sinh}, {ngay_thi}")
        return format_error(None, None, ho_va_ten, "Dữ liệu không hợp lệ")
    
    try:
        ngay_thang_nam_sinh = parse_date(ngay_thang_nam_sinh)
        ngay_thi = parse_date(ngay_thi)
        
        result_df, error_info = extract_data(ho_va_ten, ngay_thang_nam_sinh, ngay_thi)
        if result_df is not None:
            return result_df
        elif error_info is not None:
            return error_info
    except ValueError as ve:
        print(f"Lỗi định dạng dữ liệu: {ve}")
        return format_error(None, None, ho_va_ten, f"Lỗi định dạng dữ liệu: {ve}")
    except requests.RequestException as re:
        print(f"Lỗi yêu cầu HTTP: {re}")
        return format_error(None, None, ho_va_ten, f"Lỗi yêu cầu HTTP: {re}")
    except Exception as e:
        print(f"Lỗi không xác định: {e}")
        return format_error(None, None, ho_va_ten, f"Lỗi không xác định: {e}")

def process_excel(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Không thể đọc tệp Excel: {e}")
        return
    
    required_columns = {'ho ten', 'ngay sinh', 'ngay thi'}
    if not required_columns.issubset(df.columns):
        print(f"Tệp Excel phải chứa các cột sau: {', '.join(required_columns)}")
        return
    
    results_df = pd.DataFrame()

    for _, row in df.iterrows():
        result_df = process_row(row)
        results_df = pd.concat([results_df, result_df], ignore_index=True)
    
    with pd.ExcelWriter(output_file) as writer:
        results_df.to_excel(writer, sheet_name='Kết quả thi', index=False)
        print(f"Kết quả đã được lưu vào tệp '{output_file}'.")

process_excel('input.xlsx', 'ket_qua_thi.xlsx')
