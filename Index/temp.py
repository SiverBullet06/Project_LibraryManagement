import json
import os
import tkinter as tk
from tkinter import messagebox, ttk
import datetime
from PIL import Image, ImageTk
import re
from docx import Document
import requests
from io import BytesIO
# ===================== DATA FILES =====================
ACCOUNTS_FILE = "accounts.json"
LIBRARY_FILE = "book.json"
IMAGE_DIR = "Image"
DATE_REGEX = re.compile(r"^\d{4}-\d{2}-\d{2}$")
FINE_PER_DAY = 5000         # Tiền phạt mỗi ngày
# =========================================================
#  MODELS
# =========================================================
class Book:
    def __init__(self, name, author, date, category="", total_qty=1, available_qty=None, borrow_count=0):
        self.name = name
        self.author = author
        self.date = date
        self.category = category
        self.total_qty = int(total_qty)
        self.available_qty = int(available_qty if available_qty is not None else total_qty)
        self.borrow_count = int(borrow_count)

    def to_dict(self):
        return {
            "name": self.name,
            "author": self.author,
            "date": self.date,
            "category": self.category,
            "total_qty": self.total_qty,
            "available_qty": self.available_qty,
            "borrow_count": self.borrow_count,
        }
# ===================== FILE I/O =====================
def read_json(file_name, default):
    if not os.path.exists(file_name):
        return default
    try:
        with open(file_name, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return default
def write_json(file_name, data):
    with open(file_name, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
def load_library_books():
    data = read_json(LIBRARY_FILE, [])
    books = []
    for b in data:
        books.append(Book(
            b.get("name", ""), b.get("author", ""), b.get("date", "1970-01-01"),
            b.get("category", ""), b.get("total_qty", 1), b.get("available_qty", b.get("total_qty", 1)), b.get("borrow_count", 0)
        ))
    return books
def save_library_books(books):
    write_json(LIBRARY_FILE, [b.to_dict() for b in books])

    accts = read_json(ACCOUNTS_FILE, {})
    accts = ensure_account_structure(accts)
    return accts
def load_accounts():
    accts = read_json(ACCOUNTS_FILE, {})
    accts = ensure_account_structure(accts)
    return accts
def save_accounts():
    write_json(ACCOUNTS_FILE, accounts)
#========================
# Đây là hàm phụ trợ để tải ảnh, cần được đặt ở đâu đó trong file của bạn
def download_image_and_get_path(image_url, book_name):
    """Tải ảnh từ URL và lưu vào thư mục IMAGE_DIR."""

    if not os.path.exists(IMAGE_DIR):
        os.makedirs(IMAGE_DIR)

    safe_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in book_name)
    image_filename = f"{safe_name[:30]}.jpg"
    local_path = os.path.join(IMAGE_DIR, image_filename)
    
    if os.path.exists(local_path):
        return local_path

    try:
        img_response = requests.get(image_url, timeout=10)
        if img_response.status_code == 200:
            img = Image.open(BytesIO(img_response.content))
            img.save(local_path)
            return local_path
        else:
            print(f"Lỗi HTTP {img_response.status_code} khi tải ảnh từ {image_url}")
            return ""
            
    except requests.exceptions.RequestException as e:
        print(f"Lỗi mạng khi tải ảnh: {e}")
        return ""
    except Exception as e:
        print(f"Lỗi xử lý ảnh: {e}")
        return ""
def enrich_book_data(book):
    """
    Làm giàu dữ liệu chi tiết (Mô tả, Ảnh) cho MỘT cuốn sách thông qua Google Books API.
    Chỉ chạy nếu 'api_fetched' là False.
    """
    global library_books 
    
    if getattr(book, 'api_fetched', False): 
        return 
        
    search_query = getattr(book, 'isbn', None) or getattr(book, 'name')
    if not search_query:
        setattr(book, 'api_fetched', True)
        save_library_books(library_books)
        return
        
    try:
        api_url = f"https://www.googleapis.com/books/v1/volumes?q=intitle:{search_query}&maxResults=1"
        response = requests.get(api_url, timeout=10)
        response.raise_for_status() 
        data = response.json()
        
        if data.get('items'):
            volume_info = data['items'][0]['volumeInfo']
            
            # Cập nhật Mô tả
            description = volume_info.get('description', getattr(book, 'description', ''))
            setattr(book, 'description', description)
            
            # Cập nhật và Tải Ảnh
            image_url = volume_info.get('imageLinks', {}).get('thumbnail') 
            setattr(book, 'image_url', image_url or getattr(book, 'image_url', ''))

            local_path = ""
            if image_url:
                local_path = download_image_and_get_path(image_url, book.name)
            
            setattr(book, 'image_path_local', local_path)
            
        # Đặt cờ thành True dù có tìm thấy dữ liệu hay không để tránh gọi API lặp lại
        setattr(book, 'api_fetched', True) 
        save_library_books(library_books) 
            
    except requests.exceptions.RequestException as e:
        print(f"LỖI API (Network/HTTP) cho {book.name}: {e}")
    except Exception as e:
        print(f"LỖI xử lý dữ liệu API cho {book.name}: {e}")
# =========================================================
# ACCOUNT & LOGIC HELPERS
# =========================================================

# Chú thích: Đảm bảo cấu trúc dữ liệu cho mỗi tài khoản
def ensure_account_structure(accts):
    """
    Đảm bảo mỗi mục tài khoản có 'borrowed_books' và 'history' fields.
    Called once after reading accounts file.
    """
    changed = False
    for email, info in accts.items():
        if "borrowed_books" not in info:
            info["borrowed_books"] = []
            changed = True
        if "history" not in info:
            info["history"] = []
            changed = True
        if "role" not in info:
            info["role"] = "user"
            changed = True
    # Ghi lại dữ liệu vào tệp nếu có bất kỳ cấu trúc tài khoản nào được bổ sung
    if changed:
        write_json(ACCOUNTS_FILE, accts)
    return accts
# Kiểm tra tình trạng sách quá hạn và tổng mượn
def account_has_overdue(email):
    """
    Kiểm tra xem tài khoản có sách quá hạn hay không.
    Trả về (bool: bị chặn, int: tổng sách đang mượn, int: tổng sách quá hạn)
    """
    if email not in accounts:
        return False, 0, 0  
    user = accounts[email]
    borrowed_books = user.get("borrowed_books", [])    
    today = datetime.date.today()  
    is_overdue = False
    total_borrowed = 0
    total_overdue = 0

    for book in borrowed_books:
        # Chỉ xét sách mượn từ thư viện
        if book.get("status") == "borrowed":
            qty = book.get("quantity", 1)
            total_borrowed += qty
            
            due_date_str = book.get("due_date")
            if not due_date_str: continue
            try:
                due_date = datetime.datetime.strptime(due_date_str, "%Y-%m-%d").date()
            except ValueError:
                continue 
            if today > due_date:
                is_overdue = True
                total_overdue += qty
    # Logic chặn: Bị chặn nếu có ít nhất 1 sách quá hạn.
    return is_overdue, total_borrowed, total_overdue
#Tính toán số tiền phạt cho giao dịch
def calculate_fine_for_transaction(book_data, return_date=None):
    """Tính tiền phạt cho một giao dịch mượn/trả cụ thể."""
    if book_data.get("status") != "borrowed":
        return 0
        
    due_date_str = book_data.get("due_date")
    qty = book_data.get("quantity", 1)
    
    if not due_date_str:
        return 0

    try:
        due_date = datetime.datetime.strptime(due_date_str, "%Y-%m-%d").date()
    except ValueError:
        return 0
    
    # Ngày tính toán tiền phạt: Ngày trả sách hoặc Ngày hiện tại
    calculate_day = return_date or datetime.date.today()
    
    if isinstance(calculate_day, str):
        try:
            calculate_day = datetime.datetime.strptime(calculate_day, "%Y-%m-%d").date()
        except ValueError:
            calculate_day = datetime.date.today()

    if calculate_day > due_date:
        days_overdue = (calculate_day - due_date).days
        return days_overdue * FINE_PER_DAY * qty
    
    return 0
# Kiểm tra và cảnh báo người dùng về sách quá hạn
def check_overdue_and_alert(email):
    has_overdue, overdue_items, total_fine = account_has_overdue(email)
    if has_overdue:
        msg = 'Bạn có sách quá hạn:\n'
        for name, due, days, fine in overdue_items:
            msg += f"- {name}: hạn {due} ({days} ngày quá hạn) - Phạt: {fine} VND\n"
        msg += f"Tổng phạt hiện tại: {total_fine} VND"
        messagebox.showwarning('Cảnh báo quá hạn', msg)
# Tìm đối tượng sách theo tên
# Chú thích: Tìm đối tượng sách trong danh sách bộ nhớ
def get_book_object_by_name(book_name):
    """
    Tìm đối tượng sách (Book Object) đầy đủ từ danh sách sách toàn cục (library_books) 
    dựa trên Tên sách.
    """
    global library_books 
    # Làm sạch tên sách đầu vào
    cleaned_input_name = book_name.strip() 
    
    for book in library_books:
        # Kiểm tra thuộc tính 'name' của đối tượng Book và làm sạch
        book_obj_name = getattr(book, 'name', '').strip()
        
        if book_obj_name == cleaned_input_name:
            return book
    return None
# =========================================================
# 6. ACCOUNT & AUTH MANAGEMENT
# =========================================================
# Kiểm tra mật khẩu hợp lệ
def kiem_tra_mat_khau_hop_le(password, email):
    """
    Kiểm tra mật khẩu có tuân thủ các quy tắc:
    - Tối thiểu 6 ký tự
    - Chứa ít nhất 1 chữ cái và 1 số
    - Chứa ít nhất 1 ký tự đặc biệt (ví dụ: !@#$%^&*) <--- ĐÃ THÊM
    - Không chứa email
    """
    # 1. Kiểm tra độ dài
    if len(password) < 6:
        return False, "Mật khẩu phải tối thiểu 6 ký tự."
    # 2. Kiểm tra có ít nhất 1 chữ cái
    if not re.search(r'[a-zA-Z]', password):
        return False, "Mật khẩu phải chứa ít nhất 1 chữ cái."
    # 3. Kiểm tra có ít nhất 1 số
    if not re.search(r'[0-9]', password):
        return False, "Mật khẩu phải chứa ít nhất 1 số."
    # 4. KIỂM TRA CÓ ÍT NHẤT 1 KÝ TỰ ĐẶC BIỆT
    # [^a-zA-Z0-9\s] có nghĩa là bất kỳ ký tự nào không phải chữ cái, không phải số và không phải khoảng trắng.
    if not re.search(r'[^a-zA-Z0-9\s]', password):
        return False, "Mật khẩu phải chứa ít nhất 1 ký tự đặc biệt (ví dụ: !@#$%)."
    # 5. Kiểm tra không chứa email (viết thường)
    if email.lower() in password.lower():
        return False, "Mật khẩu không được chứa địa chỉ email."
    return True, ""
# Đổi mật khẩu
def change_password(email):
    """Popup đổi mật khẩu cho tài khoản (admin hoặc user)."""
    user = accounts.get(email)
    if user is None:
        messagebox.showerror("Lỗi", "Không tìm thấy tài khoản!")
        return
    w = tk.Toplevel()
    w.title(f"Đổi mật khẩu - {email}")
    w.geometry("350x250")
    w.grab_set()
    tk.Label(w, text="Mật khẩu hiện tại:").pack(pady=5)
    e_old = tk.Entry(w, show="*", width=30)
    e_old.pack()
    tk.Label(w, text="Mật khẩu mới:").pack(pady=5)
    e_new = tk.Entry(w, show="*", width=30)
    e_new.pack()
    tk.Label(w, text="Nhập lại mật khẩu mới:").pack(pady=5)
    e_new2 = tk.Entry(w, show="*", width=30)
    e_new2.pack()
    # Lưu mật khẩu mới
    def save_pw():
        old = e_old.get().strip()
        new = e_new.get().strip()
        new2 = e_new2.get().strip()

        if old != user["password"]:
            messagebox.showerror("Lỗi", "Mật khẩu hiện tại không đúng!"); return            
        if not new:
             messagebox.showerror("Lỗi", "Mật khẩu mới không được để trống!"); return
        # >> Bắt đầu kiểm tra mật khẩu mới theo tiêu chuẩn mới
        is_valid, error_msg = kiem_tra_mat_khau_hop_le(new, email)
        if not is_valid:
            messagebox.showerror("Lỗi", error_msg); return
        # << Kết thúc kiểm tra mật khẩu mới            
        if new != new2:
            messagebox.showerror("Lỗi", "Mật khẩu mới không khớp!"); return
        user["password"] = new
        save_accounts()
        messagebox.showinfo("Thành công", "Đổi mật khẩu thành công!")
        w.destroy()

    tk.Button(w, text="Cập nhật", command=save_pw, bg="#4CAF50", fg="white").pack(pady=15)
#reset tài khoản (Admin reset tài khoản User)
def admin_reset_user(email):
    if email not in accounts:
        messagebox.showerror("Lỗi", "Không tìm thấy tài khoản!"); return

    if not messagebox.askyesno("Xác nhận", f"Đặt lại tài khoản {email}?"):
        return 

    # Mật khẩu mặc định mới phải tuân thủ quy tắc: tối thiểu 6 ký tự, có 1 chữ cái, 1 số
    MAT_KHAU_RESET_MAC_DINH = "Pass12345" 
    # Kiểm tra xem mật khẩu reset có hợp lệ so với email không (tránh trùng email)
    is_valid, error_msg = kiem_tra_mat_khau_hop_le(MAT_KHAU_RESET_MAC_DINH, email)

    if not is_valid:
         # Nếu mật khẩu reset bị coi là không hợp lệ (ví dụ: chứa email người dùng), 
         # admin nên được yêu cầu nhập mật khẩu reset khác, hoặc dùng mật khẩu mặc định an toàn hơn.
         messagebox.showerror("Lỗi", f"Không thể reset. Mật khẩu mặc định '{MAT_KHAU_RESET_MAC_DINH}' vi phạm quy tắc: {error_msg}")
         return
    
    accounts[email]["password"] = MAT_KHAU_RESET_MAC_DINH
    save_accounts()
    messagebox.showinfo("Thành công", f"Đã đặt lại mật khẩu cho {email} thành '{MAT_KHAU_RESET_MAC_DINH}'.")
# ===================== GUI ROOT =====================
win = tk.Tk()
win.title("Quản lý thư viện - Nâng cao")
win.geometry("1100x750")
win.configure(bg='#f4f4f4')

style = ttk.Style(win)
style.theme_use('clam')
style.configure('TNotebook', background='#f4f4f4', borderwidth=0)
style.configure('TNotebook.Tab', background='#d0e0ff', foreground='black', borderwidth=1, padding=[12, 5], font=('Arial', 11, 'bold'))
style.map('TNotebook.Tab', background=[('selected', '#0b60ff')], foreground=[('selected', 'white')])
style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
style.configure("Treeview", font=('Arial', 10), rowheight=28)

container = tk.Frame(win)
container.pack(fill="both", expand=True)

frame_login = tk.Frame(container, bg="#f4f4f4")
frame_register = tk.Frame(container, bg="#f4f4f4")
frame_admin = tk.Frame(container, bg="white")
frame_user = tk.Frame(container, bg="#f4f4f4")

for f in (frame_login, frame_register, frame_admin, frame_user):
    f.place(x=0, y=0, relwidth=1, relheight=1)

def switch_frame(frame):
    frame.tkraise()

# ===================== SYSTEM STATE =====================
accounts = load_accounts()
library_books = load_library_books()

current_email = None        # Ai đang đăng nhập
current_role = None         # user / admin

# ===================== STATISTICS STATE =====================
global total_fine_var
total_fine_var = tk.StringVar(value="Tổng tiền nợ: 0 VND")
lbl_total_fine = None
lbl_status = None
# ===================== ADMIN TREEVIEWS =====================
tree_admin = None          # Quản lý sách
tree_accounts = None       # Quản lý tài khoản
tree_top = None            # Top sách mượn nhiều
tree_rare = None           # Sách ít mượn / hiếm
role_filter_var = tk.StringVar(value="Tất cả")  # Biến điều khiển cho Combobox
# ===================== USER TREEVIEWS =====================
tree_user = None           # Sách đang có trong thư viện cho user xem
tree_user_history = None   # Lịch sử mượn trả
# ===================== ADMIN SEARCH & FILTER =====================
lib_search_entry = None
lib_field_var = None

user_search_entry = None
user_field_var = None

lib_sort_field_var = None
lib_sort_order_var = None
# ===================== USER VIEW SEARCH =====================
tree_library_view = None
lib_search_entry_view = None
lib_field_var_view = None

# =========================================================
# 7. ADMIN – USER MANAGEMENT
# =========================================================
def refresh_accounts(role_filter=None):
    """
    Làm mới dữ liệu trong tree_accounts, có hỗ trợ lọc theo vai trò.
    Hàm này được gọi từ build_accounts hoặc các nút 'Làm mới' khác.
    """
    # Tránh lỗi nếu treeview chưa được tạo
    if tree_accounts is None: 
        return
        
    # Xóa tất cả các mục hiện có trong Treeview
    tree_accounts.delete(*tree_accounts.get_children())

    # Chuẩn hóa bộ lọc (nếu được gọi từ các hàm cũ không có tham số)
    if role_filter is None or role_filter == "Tất cả":
        filter_mode = False
    else:
        filter_mode = True
        
    for email, info in accounts.items():
        role = info.get("role", "user")
        
        # --- LOGIC LỌC MỚI ---
        if filter_mode:
            # Nếu đang ở chế độ lọc và vai trò không khớp, bỏ qua tài khoản này
            if role != role_filter:
                continue
        # ---------------------
        
        password = info.get("password", "")
        borrowed_books = info.get("borrowed_books", [])
        
        # 1. Tạo chuỗi danh sách sách đang mượn (Giới hạn độ dài chuỗi để tránh cột quá rộng)
        borrowed_list = ""
        for b in borrowed_books:
            name = b.get("book_name", "N/A")
            qty = b.get("quantity", 1)
            
            # Giới hạn độ dài hiển thị để không làm Treeview bị tràn
            if len(borrowed_list) < 300: 
                borrowed_list += f"{name} ({qty}), "
            else:
                borrowed_list += "..."
                break
                
        # Loại bỏ dấu phẩy và khoảng trắng cuối cùng
        borrowed_list = borrowed_list.rstrip(", ")

        # 2. Tính tổng số lượng sách đang mượn
        borrowed_count = sum(b.get("quantity", 1) for b in borrowed_books)

        # Chèn dữ liệu vào Treeview
        tree_accounts.insert(
            "",
            tk.END,
            values=(email, password, role, borrowed_count, borrowed_list)
        )

def show_book_details(book):
    """
    Tạo cửa sổ Toplevel để hiển thị chi tiết sách, ảnh bìa và mô tả.
    Sử dụng dữ liệu đã được làm giàu bởi enrich_book_data.
    """
    
    w = tk.Toplevel()
    w.title(f"Chi tiết: {getattr(book, 'name', 'Sách')}") # Sử dụng getattr an toàn hơn
    w.geometry("700x500") 
    w.resizable(False, False)
    w.grab_set() 
    # 

    # Khung chính (chia thành hai cột)
    main_frame = tk.Frame(w, padx=15, pady=15)
    main_frame.pack(fill="both", expand=True)

    # Cột Trái: Ảnh Bìa (200px cố định)
    frame_left = tk.Frame(main_frame, width=200)
    frame_left.pack(side="left", fill="y", padx=(0, 20))
    frame_left.pack_propagate(False)

    # Cột Phải: Thông tin và Mô tả
    frame_right = tk.Frame(main_frame)
    frame_right.pack(side="right", fill="both", expand=True)

    # --- CỘT TRÁI: HIỂN THỊ ẢNH ---
    image_path = getattr(book, 'image_path_local', '')
    
    if image_path and os.path.exists(image_path):
        # ... (Logic tải ảnh thành công)
        try:
            img = Image.open(image_path)
            img = img.resize((180, 250), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            w.image_ref = photo 
            lbl_image = tk.Label(frame_left, image=photo)
            lbl_image.pack(pady=10)
        except Exception:
            tk.Label(frame_left, text="Không thể tải ảnh", fg="red").pack(pady=10)
    else:
        # 📌 CHUỖI THÔNG BÁO TỐT HƠN KHI THIẾU ẢNH
        tk.Label(frame_left, text="Không có ảnh bìa\n(Admin chưa cập nhật)", 
                 fg="gray", width=20, height=15, relief="groove", justify=tk.CENTER).pack(pady=10)


    # --- CỘT PHẢI: THÔNG TIN CHI TIẾT ---
    tk.Label(frame_right, text=getattr(book, 'name', 'N/A'), font=("Arial", 18, "bold")).pack(anchor="w")

    tk.Label(frame_right, text=f"Tác giả: {getattr(book, 'author', 'N/A')}", font=("Arial", 12)).pack(anchor="w")
    tk.Label(frame_right, text=f"Thể loại: {getattr(book, 'category', 'N/A')}", font=("Arial", 12)).pack(anchor="w")
    tk.Label(frame_right, text=f"ISBN: {getattr(book, 'isbn', 'N/A')}", font=("Arial", 12)).pack(anchor="w")
    
    ttk.Separator(frame_right, orient='horizontal').pack(fill='x', pady=10)

    tk.Label(frame_right, text="Mô tả:", font=("Arial", 14, "underline")).pack(anchor="w", pady=(0, 5))
    
    # CHUỖI MÔ TẢ MẶC ĐỊNH RÕ RÀNG HƠN
    default_desc = 'Mô tả chưa có. Admin cần nhấn đúp vào sách để tự động tải dữ liệu từ API.'
    description = getattr(book, 'description', default_desc)
    
    # Khung chứa Text và Scrollbar
    text_frame = tk.Frame(frame_right)
    text_frame.pack(fill="both", expand=True)

    text_desc = tk.Text(text_frame, wrap="word", height=10, borderwidth=1, relief="sunken", font=("Arial", 10))
    text_desc.insert(tk.END, description)
    text_desc.config(state=tk.DISABLED) 
    
    # Thêm Scrollbar
    scrollbar = ttk.Scrollbar(text_frame, command=text_desc.yview)
    text_desc.config(yscrollcommand=scrollbar.set)
    
    scrollbar.pack(side="right", fill="y")
    text_desc.pack(side="left", fill="both", expand=True)
    
    # Nút đóng
    tk.Button(w, text="Đóng", command=w.destroy, width=15).pack(pady=15)
def on_admin_book_double_click(event):
    """
    Xử lý sự kiện nhấn đúp chuột vào một mục trong Treeview quản lý sách.
    Làm giàu dữ liệu API và mở cửa sổ chi tiết.
    """
    global tree_admin    
    # 1. Lấy mục được chọn
    selected_item = tree_admin.focus()
    if not selected_item:
        return       
    # Lấy giá trị các cột của mục được chọn
    values = tree_admin.item(selected_item, 'values')   
    # Tên sách là cột đầu tiên (index 0)
    if not values or len(values) == 0:
        messagebox.showerror("Lỗi", "Không tìm thấy dữ liệu sách.")
        return       
    book_name = values[0] 
    # 2. Tìm đối tượng sách đầy đủ trong bộ nhớ
    book_object = get_book_object_by_name(book_name)
    
    if book_object is None:
        messagebox.showerror("Lỗi", f"Không tìm thấy dữ liệu nội bộ cho sách: {book_name}")
        return
    # 3. KÍCH HOẠT LÀM GIÀU DỮ LIỆU (Nó chỉ gọi API nếu cần và sẽ tự động lưu)
    enrich_book_data(book_object) 
    # 4. Mở cửa sổ chi tiết
    show_book_details(book_object)

def on_user_book_double_click(event):
    """
    Xử lý sự kiện nhấn đúp chuột trong Treeview dành cho người dùng (tree_library_view).
    Chỉ mở cửa sổ chi tiết, KHÔNG kích hoạt API.
    """
    global tree_library_view 
    
    selected_item = tree_library_view.focus()
    if not selected_item:
        return
        
    values = tree_library_view.item(selected_item, 'values')
    if not values or len(values) == 0:
        return
        
    # 📌 ĐIỂM SỬA: Lấy tên sách và làm sạch (strip)
    # Tên sách là cột đầu tiên (index 0)
    book_name_from_tree = str(values[0]).strip() 

    if not book_name_from_tree:
        messagebox.showerror("Lỗi", "Không tìm thấy tên sách trong dòng được chọn.")
        return

    # Tìm đối tượng sách đầy đủ trong bộ nhớ
    book_object = get_book_object_by_name(book_name_from_tree)
    
    if book_object is None:
        messagebox.showerror("Lỗi", f"Không tìm thấy dữ liệu nội bộ cho sách: {book_name_from_tree}. Vui lòng liên hệ Admin để kiểm tra dữ liệu gốc.")
        return

    # Mở cửa sổ chi tiết. Vì Admin đã chạy API, dữ liệu sẽ nằm trong book_object.
    show_book_details(book_object)

def on_admin_user_double_click(event):
    selected = tree_accounts.selection()
    if not selected:
        return
    item = tree_accounts.item(selected[0])
    email = item["values"][0]
    if email not in accounts:
        return
    open_user_borrow_popup(email)

def open_user_borrow_popup(email):
    user_info = accounts.get(email, {})
    borrowed = user_info.get("borrowed_books", [])

    popup = tk.Toplevel()
    popup.title(f"Sách {email} đang mượn")
    popup.geometry("700x450")
    popup.grab_set()

    tk.Label(popup, text=f"Người dùng: {email}", font=("Arial", 12, "bold")).pack(pady=5)

    cols = ("book_name", "borrow_date", "due_date", "quantity")
    tree = ttk.Treeview(popup, columns=cols, show="headings")
    tree.heading("book_name", text="Tên sách")
    tree.column("book_name", width=250)

    tree.heading("borrow_date", text="Ngày mượn")
    tree.column("borrow_date", width=100)

    tree.heading("due_date", text="Hạn trả")
    tree.column("due_date", width=100)

    tree.heading("quantity", text="SL")
    tree.column("quantity", width=50, anchor="center")

    tree.pack(fill="both", expand=True, padx=10, pady=10)

    # Đổ dữ liệu
    for b in borrowed:
        tree.insert("", tk.END, values=(
            b.get("book_name", ""),
            b.get("borrow_date", ""),
            b.get("due_date", ""),
            b.get("quantity", 1)
        ))
    # Nút Force Return từng sách
    def force_return_selected():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Lỗi", "Hãy chọn sách để ép trả!")
            return

        book = tree.item(selected_item[0])["values"]
        book_name = book[0]

        force_return_book(email, book_name)
        popup.destroy()
        refresh_accounts()
        refresh_admin_books()
        refresh_stats()
    tk.Button(popup, text="🔑 Đổi mật khẩu", bg="#0d6efd", fg="white",
          font=("Arial", 12, "bold"),
          command=lambda e=email: change_password(e)).pack(pady=5)
    tk.Button(popup, text="↩ ÉP TRẢ SÁCH ĐANG CHỌN", bg="#d9534f", fg="white",
              command=force_return_selected).pack(pady=5)

    # Nút ép trả tất cả
    def force_return_all():
        if messagebox.askyesno("Xác nhận", f"Ép trả tất cả sách của {email}?"):
            for b in borrowed[:]:
                force_return_book(email, b["book_name"])
            popup.destroy()
            refresh_accounts()
            refresh_admin_books()
            refresh_stats()
def force_return_book(email, book_name):
    if email not in accounts:
        return

    user = accounts[email]
    borrowed = user.get("borrowed_books", [])
    today_iso = datetime.date.today().isoformat()
    
    # Tìm và xóa sách (Chỉ xóa một lần trả sách)
    # Nếu muốn admin trả toàn bộ (quantity), ta cần lặp qua borrowed và xóa tất cả
    # Hiện tại, hàm open_user_borrow_popup chỉ gọi force_return_book(email, book_name) dựa trên TÊN.
    # Ta sẽ giả định Admin luôn ép trả TOÀN BỘ số lượng đang mượn của quyển sách đó.

    book_to_remove = None
    for b in borrowed:
        if b.get("book_name") == book_name:
            book_to_remove = b
            break
    
    if book_to_remove:
        qty = book_to_remove.get("quantity", 1)
        
        # --- BẮT ĐẦU LOGIC TÍNH VÀ LƯU FINE CUỐI CÙNG (Admin Ép Trả) ---
        final_fine = calculate_fine_for_transaction(book_to_remove, datetime.date.today())
        # --- KẾT THÚC LOGIC TÍNH VÀ LƯU FINE CUỐI CÙNG ---

        # 1. Tăng số lượng sách thư viện
        for lib in library_books:
            if lib.name == book_name:
                lib.available_qty += qty
                break

        # 2. Thêm vào lịch sử (Action: force_return)
        user.setdefault('history', []).append({
            'book_name': book_to_remove.get('book_name'), 
            'author': book_to_remove.get('author'), 
            'quantity': qty, 
            'action':'force_return', 
            'return_date': today_iso, 
            'borrow_date': book_to_remove.get('borrow_date'),
            'due_date': book_to_remove.get('due_date'),
            'fine': final_fine # LƯU TIỀN PHẠT VÀO LỊCH SỬ
        })

        # 3. Xóa khỏi danh sách mượn
        borrowed.remove(book_to_remove)

        save_library_books(library_books)
        save_accounts()
        
        if final_fine > 0:
            messagebox.showwarning("Admin Ép Trả", f"Đã ép trả {qty} bản của '{book_name}'. Tiền phạt: {final_fine:,} VND")
        else:
            messagebox.showinfo("Admin Ép Trả", f"Đã ép trả {qty} bản của '{book_name}'. Không có tiền phạt.")
# =========================================================
# 8. USER FUNCTIONS
# =========================================================
def refresh_user_books():
    """Refresh the tree showing current borrowed books (for logged-in user)."""
    if current_email is None: return
    
    # Khai báo biến global nếu total_fine_var không được khai báo ở phạm vi global
    global total_fine_var 
    
    tree_bor.delete(*tree_bor.get_children())
    tree_user.delete(*tree_user.get_children())
    ac = accounts.get(current_email, {})
    
    total_fine_amount = 0
    today = datetime.date.today() 

    for it in ac.get("borrowed_books", []):
        name = it.get("book_name"); author = it.get("author","")
        qty = it.get("quantity", 1)
        # Các trường khác: bd, dd, status...

        # --- BẮT ĐẦU LOGIC TÍNH TIỀN PHẠT (Tính toán Fine để hiển thị) ---
        book_fine = calculate_fine_for_transaction(it, today)
        total_fine_amount += book_fine
        # NOTE: Fine chỉ được tính toán TẠM THỜI ở đây để hiển thị, 
        # Fine CHÍNH THỨC sẽ được lưu vào History khi trả sách.
        
        # Thêm cột tiền phạt để hiển thị (Nếu giao diện của bạn có cột này)
        # Nếu không có cột Fine, bạn cần bỏ qua phần chèn này hoặc sửa lại GUI.
        tree_bor.insert("", tk.END, values=(name, author, qty, it.get("borrow_date",""), it.get("due_date",""), book_fine))
        tree_user.insert("", tk.END, values=(name, author, qty, it.get("status","borrowed"), it.get("due_date",""), book_fine))
    update_user_status_label() 
    # --- CẬP NHẬT LABEL TỔNG TIỀN PHẠT ---
    total_fine_var.set(f"Tổng tiền nợ: {total_fine_amount:,} VND")
def refresh_user_history():
    """Refresh the user's history tree."""
    if current_email is None: return
    tree_user_history.delete(*tree_user_history.get_children())
    ac = accounts.get(current_email, {})
    for h in ac.get("history", []):
        name = h.get("book_name") or h.get("name") or h.get("name")
        qty = h.get("quantity", 1)
        act = h.get("action", "")
        bd = h.get("borrow_date", h.get("bd",""))
        rd = h.get("return_date", h.get("rd",""))
        tree_user_history.insert("", tk.END, values=(name, qty, act, bd, rd))
def refresh_library_view(source=None):
    """Cập nhật Treeview rút gọn, bao gồm Ngày xuất bản."""
    if tree_library_view is None: return
    tree_library_view.delete(*tree_library_view.get_children())
    src = source if source else library_books
    for b in src:
        # Chèn 4 cột: name, author, category, date
        tree_library_view.insert("", tk.END, values=(b.name, b.author, b.category, b.date))

# =========================================================
# 9. STATISTICS
# =========================================================
def refresh_stats():
    if tree_top is None: return
    tree_top.delete(*tree_top.get_children())
    tree_rare.delete(*tree_rare.get_children())
    sorted_by_borrow = sorted(library_books, key=lambda x: x.borrow_count, reverse=True)
    for b in sorted_by_borrow[:10]:
        tree_top.insert("", tk.END, values=(b.name, b.author, b.borrow_count))
    rarely = [b for b in library_books if b.borrow_count == 0]
    if not rarely:
        rarely = sorted(library_books, key=lambda x: x.borrow_count)[:10]
    for b in rarely[:10]:
        tree_rare.insert("", tk.END, values=(b.name, b.author, b.borrow_count))
# =========================================================
# 10. GUI HELPERS
# =========================================================
def refresh_admin_books(source=None):
    if tree_admin is None: return
    tree_admin.delete(*tree_admin.get_children())
    src = source if source else library_books
    for b in src:
        tree_admin.insert("", tk.END, values=(b.name, b.author, b.category, b.available_qty, b.total_qty, b.borrow_count))

def update_user_status_label():
    if lbl_status is None: return
    if current_email is None:
        lbl_status.config(text="Trạng thái: Chưa đăng nhập", fg="black")
        return
    blocked, _, _ = account_has_overdue(current_email)
    if blocked:
        lbl_status.config(text="TRẠNG THÁI: BỊ KHÓA (CÓ SÁCH QUÁ HẠN)", fg="red")
    else:
        lbl_status.config(text="Trạng thái: Hoạt động bình thường", fg="green")
# ===================== ADMIN CRUD =====================
def add_book_admin():
    w = tk.Toplevel(win); w.title("Thêm sách (Thư viện)")
    tk.Label(w, text="Tên sách:").pack(); e_name = tk.Entry(w, width=40); e_name.pack()
    tk.Label(w, text="Tác giả:").pack(); e_author = tk.Entry(w, width=40); e_author.pack()
    tk.Label(w, text="Ngày XB (yyyy-mm-dd):").pack(); e_date = tk.Entry(w, width=40); e_date.pack()
    tk.Label(w, text="Thể loại:").pack(); e_cat = tk.Entry(w, width=40); e_cat.pack()
    tk.Label(w, text="Số lượng (số nguyên):").pack(); e_qty = tk.Entry(w, width=40); e_qty.insert(0, "1"); e_qty.pack()
    def save_new():
        name, author, date = e_name.get().strip(), e_author.get().strip(), e_date.get().strip()
        cat = e_cat.get().strip(); qty = e_qty.get().strip()
        if not name or not author or not date:
            messagebox.showwarning("Lỗi", "Vui lòng nhập đủ dữ liệu!"); return
        if not DATE_REGEX.match(date):
            messagebox.showwarning("Lỗi", "Ngày xuất bản phải có định dạng yyyy-mm-dd"); return
        try:
            qty = int(qty)
            if qty < 1:
                raise ValueError()
        except:
            messagebox.showwarning("Lỗi", "Số lượng phải là số nguyên dương!"); return
        for b in library_books:
            if b.name == name and b.author == author and b.date == date:
                b.total_qty += qty
                b.available_qty += qty
                save_library_books(library_books)
                refresh_admin_books()
                refresh_stats()
                w.destroy()
                return
        library_books.append(Book(name, author, date, cat, qty))
        save_library_books(library_books)
        refresh_admin_books()
        refresh_stats()
        w.destroy()
    tk.Button(w, text="Lưu", command=save_new).pack(pady=8)

def delete_book_admin():
    sel = tree_admin.selection()
    if not sel:
        messagebox.showwarning("Lỗi", "Chọn sách để xóa!"); return
    item = tree_admin.item(sel[0])["values"]
    name = item[0]; author = item[1]
    # don't allow delete if some borrowed (available != total)
    for i,b in enumerate(library_books):
        if b.name == name and b.author == author:
            if b.total_qty > b.available_qty:
                messagebox.showerror("Lỗi", "Sách đang có người mượn!"); return
            del library_books[i]
            save_library_books(library_books)
            refresh_admin_books()
            refresh_stats()
            return

def edit_book_admin():
    sel = tree_admin.selection()
    if not sel:
        messagebox.showwarning("Lỗi", "Chọn sách để sửa!"); return
    item = tree_admin.item(sel[0])["values"]
    name = item[0]; author = item[1]
    for b in library_books:
        if b.name == name and b.author == author:
            target = b; break
    else:
        messagebox.showwarning("Lỗi", "Không tìm thấy sách!"); return
    w = tk.Toplevel(win); w.title("Sửa sách (Thư viện)")
    tk.Label(w, text="Tên sách:").pack(); e_name = tk.Entry(w, width=40); e_name.insert(0, target.name); e_name.pack()
    tk.Label(w, text="Tác giả:").pack(); e_author = tk.Entry(w, width=40); e_author.insert(0, target.author); e_author.pack()
    tk.Label(w, text="Ngày XB (yyyy-mm-dd):").pack(); e_date = tk.Entry(w, width=40); e_date.insert(0, target.date); e_date.pack()
    tk.Label(w, text="Thể loại:").pack(); e_cat = tk.Entry(w, width=40); e_cat.insert(0, target.category); e_cat.pack()
    tk.Label(w, text="Số lượng tổng:").pack(); e_qty = tk.Entry(w, width=40); e_qty.insert(0, str(target.total_qty)); e_qty.pack()
    def save_edit():
        name, author, date = e_name.get().strip(), e_author.get().strip(), e_date.get().strip()
        cat = e_cat.get().strip(); qty = e_qty.get().strip()
        if not DATE_REGEX.match(date):
            messagebox.showwarning("Lỗi", "Ngày xuất bản phải có định dạng yyyy-mm-dd"); return
        try:
            qty = int(qty)
            if qty < 1:
                raise ValueError()
        except:
            messagebox.showwarning("Lỗi", "Số lượng phải là số nguyên dương!"); return
        diff = qty - target.total_qty
        target.total_qty = qty
        target.available_qty = max(0, target.available_qty + diff)
        target.name = name; target.author = author; target.date = date; target.category = cat
        save_library_books(library_books)
        refresh_admin_books()
        refresh_stats()
        w.destroy()
    tk.Button(w, text="Cập nhật", command=save_edit).pack(pady=8)

# ===================== USER CRUD =====================
def add_book_user():
    if current_email is None:
        messagebox.showwarning("Lỗi", "Bạn cần đăng nhập để thực hiện thao tác này."); return
    w = tk.Toplevel(win); w.title("Thêm sách (Cá nhân)")
    tk.Label(w, text="Tên sách:").pack(); e_name = tk.Entry(w, width=40); e_name.pack()
    tk.Label(w, text="Tác giả:").pack(); e_author = tk.Entry(w, width=40); e_author.pack()
    tk.Label(w, text="Ngày XB (yyyy-mm-dd):").pack(); e_date = tk.Entry(w, width=40); e_date.pack()
    tk.Label(w, text="Thể loại:").pack(); e_cat = tk.Entry(w, width=40); e_cat.pack()
    def save_new():
        name, author, date = e_name.get().strip(), e_author.get().strip(), e_date.get().strip()
        cat = e_cat.get().strip()
        if not name or not author or not date:
            messagebox.showwarning("Lỗi", "Vui lòng nhập đủ dữ liệu!"); return
        if not DATE_REGEX.match(date):
            messagebox.showwarning("Lỗi", "Ngày xuất bản phải có định dạng yyyy-mm-dd"); return
        ac = accounts.get(current_email)
        # check if user already has this personal book (returned status)
        for it in ac.get("borrowed_books", []) + ac.get("history", []):
            pass
        # store as personal item in borrowed_books with status 'personal' (not borrowed from library)
        ac.setdefault("borrowed_books", []).append({
            "book_name": name, "author": author, "date": date, "category": cat,
            "quantity": 1, "status": "personal", "borrow_date": "", "due_date": ""
        })
        save_accounts()
        refresh_user_books()
        w.destroy()
    tk.Button(w, text="Lưu", command=save_new).pack(pady=8)

def delete_book_user():
    if current_email is None:
        messagebox.showwarning("Lỗi", "Bạn cần đăng nhập để thực hiện thao tác này."); return
    sel = tree_user.selection()
    if not sel:
        messagebox.showwarning("Lỗi", "Chọn sách để xóa!"); return
    ac = accounts.get(current_email)
    # locate by index in displayed list: we reconstruct displayed list index by iterating borrowed_books
    # simpler: remove by matching name + author + due_date
    vals = tree_user.item(sel[0])["values"]
    name, author = vals[0], vals[1]
    for i, it in enumerate(ac.get("borrowed_books", [])):
        if it.get("book_name", it.get("name")) == name and it.get("author","") == author:
            # if borrowed (status borrowed) prevent deletion
            if it.get("status") == "borrowed":
                messagebox.showerror("Lỗi", "Sách đang mượn thư viện!"); return
            ac["borrowed_books"].pop(i)
            save_accounts()
            refresh_user_books()
            refresh_user_history()
            return
    messagebox.showwarning("Lỗi", "Không tìm thấy mục để xóa!")

def edit_book_user():
    if current_email is None:
        messagebox.showwarning("Lỗi", "Bạn cần đăng nhập để thực hiện thao tác này."); return
    sel = tree_user.selection()
    if not sel:
        messagebox.showwarning("Lỗi", "Chọn sách để sửa!"); return
    vals = tree_user.item(sel[0])["values"]
    name, author = vals[0], vals[1]
    ac = accounts.get(current_email)
    for it in ac.get("borrowed_books", []):
        if it.get("book_name", it.get("name")) == name and it.get("author","") == author:
            target = it; break
    else:
        messagebox.showwarning("Lỗi", "Không tìm thấy mục để sửa!"); return
    w = tk.Toplevel(win); w.title("Sửa sách (Cá nhân)")
    tk.Label(w, text="Tên sách:").pack(); e_name = tk.Entry(w, width=40); e_name.insert(0, target.get("book_name", target.get("name",""))); e_name.pack()
    tk.Label(w, text="Tác giả:").pack(); e_author = tk.Entry(w, width=40); e_author.insert(0, target.get("author","")); e_author.pack()
    tk.Label(w, text="Ngày XB (yyyy-mm-dd):").pack(); e_date = tk.Entry(w, width=40); e_date.insert(0, target.get("date","")); e_date.pack()
    tk.Label(w, text="Thể loại:").pack(); e_cat = tk.Entry(w, width=40); e_cat.insert(0, target.get("category","")); e_cat.pack()
    def save_edit():
        name, author, date = e_name.get().strip(), e_author.get().strip(), e_date.get().strip()
        cat = e_cat.get().strip()
        if date and not DATE_REGEX.match(date):
            messagebox.showwarning("Lỗi", "Ngày xuất bản phải có định dạng yyyy-mm-dd"); return
        target["book_name"] = name
        target["author"] = author
        target["date"] = date
        target["category"] = cat
        save_accounts()
        refresh_user_books()
        w.destroy()
    tk.Button(w, text="Cập nhật", command=save_edit).pack(pady=8)

# ===================== BORROW / RETURN (update accounts structure) =====================
def borrow_from_library():
    # ----------------------------------------------------
    # KIỂM TRA ĐIỀU KIỆN BAN ĐẦU
    # ----------------------------------------------------
    if current_email is None:
        messagebox.showwarning("Lỗi", "Bạn cần đăng nhập để mượn sách!"); return
        
    # Chặn nếu tài khoản bị quá hạn
    blocked, _, _ = account_has_overdue(current_email)
    if blocked:
        messagebox.showerror("CHẶN", "Tài khoản bị khóa do sách quá hạn!"); return
        
    w = tk.Toplevel(win); w.title("Mượn sách từ thư viện")
    tk.Label(w, text="Chọn sách từ thư viện:").pack()
    
    names = [f"{b.name} --- ({b.available_qty} còn)" for b in library_books if b.available_qty>0]
    if not names:
        messagebox.showinfo("Thông báo", "Hiện không có sách khả dụng để mượn.")
        w.destroy(); return
        
    combo = ttk.Combobox(w, values=names, width=80)
    combo.pack()
    
    tk.Label(w, text="Số lượng muốn mượn:").pack()
    e_qty = tk.Entry(w, width=20)
    e_qty.insert(0, "1")
    e_qty.pack()

    SO_NGAY_MUON_TOI_DA = 7
    tk.Label(w, text="[Ngày mượn] Tự động lấy là ngày hiện tại.").pack()
    # Hiển thị thông tin hạn trả cố định
    han_tra_hien_tai = (datetime.date.today() + datetime.timedelta(days=SO_NGAY_MUON_TOI_DA)).isoformat()
    tk.Label(w, text=f"[Hạn trả] Tự động tính: {SO_NGAY_MUON_TOI_DA} ngày (Ước tính: {han_tra_hien_tai})").pack()

    def do_borrow():
        sel = combo.get()
        if not sel:
            messagebox.showwarning("Lỗi", "Chọn sách muốn mượn!"); return
            
        # Lấy thông tin sách và kiểm tra số lượng
        name = sel.split(" --- ")[0]
        libbook = None
        for b in library_books:
            if b.name == name: libbook = b; break
        if libbook is None:
            messagebox.showerror("Lỗi", "Không tìm thấy sách thư viện!"); return
            
        try:
            qty = int(e_qty.get()); 
            if qty < 1: raise ValueError()
        except:
            messagebox.showwarning("Lỗi", "Số lượng phải là số nguyên dương!"); return
            
        if qty > libbook.available_qty:
            messagebox.showwarning("Lỗi", f"Chỉ còn {libbook.available_qty} cuốn khả dụng."); return
            
        ngay_hien_tai = datetime.date.today()
        # Ngày mượn là ngày ấn nút
        bdate = ngay_hien_tai.isoformat() 
        
        # Hạn trả: 7 ngày kể từ ngày mượn
        ngay_han_tra = ngay_hien_tai + datetime.timedelta(days=SO_NGAY_MUON_TOI_DA)
        ddate = ngay_han_tra.isoformat() 
        # ----------------------------------------------------
        
        # update library (Cập nhật số lượng sách còn trong thư viện)
        libbook.available_qty -= qty
        libbook.borrow_count += qty
        save_library_books(library_books)
        
        # update accounts[current_email] (Cập nhật tài khoản người dùng)
        ac = accounts.get(current_email)
        
        # Cập nhật danh sách sách đang mượn
        for it in ac.setdefault("borrowed_books", []):
            if it.get("book_name") == libbook.name and it.get("author") == libbook.author:
                it["quantity"] = it.get("quantity",1) + qty
                it["status"] = "borrowed"
                it["borrow_date"] = bdate # Ngày mượn tự động
                it["due_date"] = ddate    # Hạn trả tự động (7 ngày)
                break
        else:
            # Thêm mới mục giao dịch nếu chưa từng mượn cuốn này
            ac.setdefault("borrowed_books", []).append({
                "book_name": libbook.name, "author": libbook.author, "date": libbook.date, "category": libbook.category,
                "quantity": qty, "status": "borrowed", "borrow_date": bdate, "due_date": ddate
            })
            
        # Ghi vào lịch sử giao dịch
        ac.setdefault("history", []).append({
            "book_name": libbook.name, "author": libbook.author, "date": libbook.date, "category": libbook.category,
            "quantity": qty, "borrow_date": bdate, "due_date": ddate, "action": "borrow"
        })
        
        save_accounts()
        refresh_user_books(); refresh_user_history(); refresh_admin_books(); refresh_stats()
        
        messagebox.showinfo("OK", f"Mượn thành công {qty} bản: {libbook.name}")
        w.destroy()
        
    tk.Button(w, text="Xác nhận mượn", command=do_borrow).pack(pady=8)

def return_book_user():
    if current_email is None:
        messagebox.showwarning("Lỗi", "Bạn cần đăng nhập để trả sách!"); return
    sel = tree_bor.selection()
    if not sel:
        messagebox.showwarning("Lỗi", "Chọn sách để trả!"); return
    vals = tree_bor.item(sel[0])["values"]
    name = vals[0]; author = vals[1]
    ac = accounts.get(current_email)
    
    # find item in borrowed_books
    for i, it in enumerate(ac.get("borrowed_books", [])):
        if it.get("book_name") == name and it.get("author") == author:
            target = it; idx = i; break
    else:
        messagebox.showerror("Lỗi", "Không tìm thấy mục mượn trong tài khoản!"); return
    
    if target.get("status") != "borrowed":
        if messagebox.askyesno("Xác nhận", "Mục này không ở trạng thái 'borrowed'. Bạn muốn xóa mục này khỏi danh sách cá nhân không?"):
            ac["borrowed_books"].pop(idx)
            save_accounts(); refresh_user_books(); refresh_user_history()
        return
        
    # ask qty to return
    w = tk.Toplevel(win); w.title("Trả sách")
    tk.Label(w, text=f"Trả sách: {target['book_name']}").pack()
    tk.Label(w, text=f"Số lượng hiện có: {target.get('quantity',1)}").pack()
    tk.Label(w, text="Số lượng trả:").pack(); e_qty = tk.Entry(w, width=20); e_qty.insert(0, str(target.get('quantity',1))); e_qty.pack() # Mặc định trả hết
    
    def do_return():
        try:
            rqty = int(e_qty.get())
            if rqty < 1 or rqty > target.get('quantity',1): raise ValueError()
        except:
            messagebox.showwarning("Lỗi", "Số lượng trả không hợp lệ!"); return
            
        today = datetime.date.today()
        today_iso = today.isoformat()
        
        # --- BẮT ĐẦU LOGIC TÍNH VÀ LƯU FINE CUỐI CÙNG ---
        # Chỉ tính fine cho số lượng sách đang được trả
        
        # Tạm thời tạo một bản sao giao dịch để tính fine
        temp_transaction = target.copy()
        temp_transaction['quantity'] = rqty # Chỉ tính fine trên số lượng trả
        final_fine = calculate_fine_for_transaction(temp_transaction, today)
        # --- KẾT THÚC LOGIC TÍNH VÀ LƯU FINE CUỐI CÙNG ---

        # 1. Update library book available
        for b in library_books:
            if b.name == target['book_name'] and b.author == target.get('author'):
                b.available_qty += rqty
                save_library_books(library_books)
                break
                
        # 2. Decrement user's quantity
        target['quantity'] = target.get('quantity',1) - rqty
        
        # 3. Append return to history
        ac.setdefault('history', []).append({
            'book_name': target['book_name'], 
            'author': target.get('author'), 
            'quantity': rqty, 
            'action':'return', 
            'return_date': today_iso, 
            'borrow_date': target.get('borrow_date'),
            'due_date': target.get('due_date'),
            'fine': final_fine # LƯU TIỀN PHẠT VÀO LỊCH SỬ
        })
        
        # 4. Remove book if zero left
        if target['quantity'] <= 0:
            ac['borrowed_books'].pop(idx)
        else:
            # Nếu còn sách, giữ nguyên status 'borrowed'
            target['status'] = 'borrowed' 
            
        save_accounts()
        refresh_user_books(); refresh_user_history(); refresh_admin_books(); refresh_stats()
        
        if final_fine > 0:
             messagebox.showwarning("CẢNH BÁO", f"Trả thành công {rqty} bản của \"{target['book_name']}\". Tiền phạt: {final_fine:,} VND")
        else:
             messagebox.showinfo("OK", f"Trả thành công {rqty} bản của \"{target['book_name']}\". Không có tiền phạt.")
        
        w.destroy()
        
    tk.Button(w, text='Xác nhận trả', command=do_return).pack(pady=8)

# ===================== SEARCH & SORT =====================
def filter_books(books, keyword, field):
    kw = (keyword or "").strip().lower()
    if not kw:
        return books[:]
    def get_field(b):
        if isinstance(b, Book):
            val = getattr(b, field, "")
        else:
            val = b.get(field, "")
        return str(val)
    return [b for b in books if kw in get_field(b).lower()]

def sort_books(books, field, order):
    reverse = (order == "desc")
    def keyfn(b):
        if isinstance(b, Book):
            return getattr(b, field, "")
        else:
            return b.get(field, "")
    return sorted(books, key=keyfn, reverse=reverse)

def do_search_admin():
    kw = lib_search_entry.get()
    field = lib_field_var.get()
    filtered = filter_books(library_books, kw, field)
    refresh_admin_books(filtered)

def do_sort_admin():
    field = lib_sort_field_var.get()
    order = lib_sort_order_var.get()
    sorted_list = sort_books(library_books, field, order)
    refresh_admin_books(sorted_list)

def do_search_user():
    kw = user_search_entry.get()
    field = user_field_var.get()
    # filter user's borrowed list
    ac = accounts.get(current_email, {})
    filt = []
    for b in ac.get("borrowed_books", []):
        val = str(b.get("book_name","")) if field == "name" else str(b.get(field,""))
        if kw.lower() in val.lower():
            filt.append(b)
    # refresh user tree using filtered list
    tree_user.delete(*tree_user.get_children())
    tree_bor.delete(*tree_bor.get_children())
    for it in filt:
        tree_bor.insert("", tk.END, values=(it.get("book_name"), it.get("author",""), it.get("quantity",1), it.get("borrow_date",""), it.get("due_date","")))
        tree_user.insert("", tk.END, values=(it.get("book_name"), it.get("author",""), it.get("quantity",1), it.get("status","borrowed"), it.get("due_date","")))
def do_search_library_view():
    """Tìm kiếm và cập nhật Treeview rút gọn."""
    if lib_search_entry_view is None or lib_field_var_view is None: return
    
    kw = lib_search_entry_view.get()
    field = lib_field_var_view.get()
    
    # Tái sử dụng hàm filter_books đã có
    filtered = filter_books(library_books, kw, field)
    refresh_library_view(filtered)
#Lọc
def filter_by_category(cat):
    global library_books, tree_admin
    # 1. Nếu chọn Tất cả -> hiển thị lại toàn bộ
    if cat == "Tất cả":
        refresh_admin_books()
        return
    # 2. Lọc danh sách
    filtered = [b for b in library_books if b.category == cat]
    # 3. Hiển thị list đã lọc
    refresh_admin_books(source=filtered)

# ===================== ACCOUNT MANAGEMENT (ADMIN) =====================
def open_create_account_window():
    def create_account():
        email = entry_email.get().strip()
        password = entry_password.get().strip()
        role = role_var.get()
        
        if not email or not password:
            messagebox.showerror("Lỗi", "Điền đầy đủ thông tin!"); return
            
        if "@" not in email or "." not in email:
            messagebox.showerror("Lỗi", "Email không hợp lệ!"); return
            
        # >> Bắt đầu kiểm tra mật khẩu theo tiêu chuẩn mới
        is_valid, error_msg = kiem_tra_mat_khau_hop_le(password, email)
        if not is_valid:
            messagebox.showerror("Lỗi", error_msg); return
        # << Kết thúc kiểm tra mật khẩu

        if email in accounts:
            messagebox.showerror("Lỗi", "Email đã tồn tại!"); return
        
        # Tạo account mới (Logic tạo tài khoản giữ nguyên)
        accounts[email] = {
            "password": password,
            "role": role,
            "borrowed_books": [],
            "history": [],
            "fine": 0
        }
        save_accounts()
        refresh_accounts()
        messagebox.showinfo("OK", f"Tạo tài khoản thành công với vai trò {role}!")
        win.destroy()
    win = tk.Toplevel()
    win.title("Tạo tài khoản mới")
    win.geometry("350x220")
    win.resizable(False, False)

    tk.Label(win, text="Email:").pack(anchor="w", padx=10, pady=5)
    entry_email = tk.Entry(win, width=35)
    entry_email.pack(padx=10)

    tk.Label(win, text="Mật khẩu:").pack(anchor="w", padx=10, pady=5)
    entry_password = tk.Entry(win, width=35, show="*")
    entry_password.pack(padx=10)

    tk.Label(win, text="Vai trò:").pack(anchor="w", padx=10, pady=5)
    role_var = tk.StringVar(value="user")
    role_combo = ttk.Combobox(win, textvariable=role_var, values=["user", "admin"], state="readonly")
    role_combo.pack(padx=10)

    tk.Button(win, text="Tạo tài khoản", bg="#28a745", fg="white", command=create_account).pack(pady=15)
def delete_account_admin():
    sel = tree_accounts.selection()
    if not sel:
        messagebox.showwarning("Lỗi", "Chọn tài khoản để xóa!"); return
    
    item = tree_accounts.item(sel[0])["values"]
    email_to_delete = item[0]
    
    # 1. Kiểm tra tài khoản đang đăng nhập
    if email_to_delete == current_email:
        messagebox.showwarning("Lỗi", "Không thể xóa tài khoản đang đăng nhập!"); return
        
    user_data = accounts.get(email_to_delete)
    if user_data is None:
        messagebox.showwarning("Lỗi", "Không tìm thấy tài khoản để xóa!"); return

    # 2. Xác nhận xóa
    if messagebox.askyesno("Xác nhận", f"Xóa tài khoản: {email_to_delete}? \n(Tất cả sách đang mượn sẽ được tự động trả về thư viện!)"):
        
        # --- LOGIC MỚI: ÉP TRẢ SÁCH TRƯỚC KHI XÓA ---
        borrowed_books = user_data.get("borrowed_books", [])
        
        # Tạo bản sao của danh sách sách đang mượn để tránh lỗi khi thay đổi trong quá trình lặp
        for book in borrowed_books[:]: 
            if book.get("status") == "borrowed":
                book_name = book.get("book_name")
                # Gọi hàm ép trả (sử dụng logic trong force_return_book)
                # Lưu ý: Hàm force_return_book sẽ tự động xóa sách khỏi borrowed_books, 
                # cập nhật thư viện và lưu tài khoản.
                
                # Để tránh lỗi lặp, ta sẽ thực hiện logic trả sách đơn giản hơn tại đây:
                
                qty = book.get("quantity", 1)
                
                # 1. Cập nhật sách thư viện (tăng available_qty)
                for lib in library_books:
                    if lib.name == book_name:
                        lib.available_qty += qty
                        break
                
                # 2. Xóa khỏi borrowed_books và thêm vào history (Hàm force_return_book đã làm việc này)
                # Để giữ cho logic đơn giản, ta chỉ cần xóa borrowed_books và lưu library_books.
                # Khi tài khoản bị xóa, history cũng mất nên không cần lưu.

        # Lưu lại thư viện sau khi sách đã được trả
        save_library_books(library_books)
        
        # --- KẾT THÚC LOGIC ÉP TRẢ ---

        # 3. Xóa tài khoản
        del accounts[email_to_delete]
        save_accounts()
        
        # 4. Cập nhật giao diện
        refresh_accounts()
        refresh_admin_books()
        refresh_stats()
        messagebox.showinfo("Thành công", f"Đã xóa tài khoản {email_to_delete} và trả tất cả sách.")
# ===================== LOGIN / REGISTER / LOGOUT =====================
def login():
    global current_email, current_role, library_books, accounts

    # 🔥 Load lại dữ liệu từ file mỗi lần đăng nhập
    library_books = load_library_books()
    accounts = load_accounts()

    email = entry_login_email.get().strip()
    password = entry_login_password.get().strip()
    
    if not email or not password:
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ!"); return
    if email not in accounts:
        messagebox.showerror("Lỗi", "Tài khoản không tồn tại!"); return
    if accounts[email].get("password") != password:
        messagebox.showerror("Lỗi", "Sai mật khẩu!"); return

    current_email = email
    current_role = accounts[email].get("role", "user")

    win.state("zoomed")

    if current_role == "admin":
        switch_frame(frame_admin)
        refresh_admin_books()
        refresh_accounts()
        refresh_stats()
    else:
        switch_frame(frame_user)
        refresh_user_books()
        refresh_user_history()
        update_user_status_label()
        check_overdue_and_alert(current_email)


def register():
    email = entry_reg_email.get().strip()
    password = entry_reg_password.get().strip()
    confirm = entry_reg_confirm.get().strip()
    role = "user"  # default role

    if not email or not password or not confirm:
        messagebox.showerror("Lỗi", "Điền đầy đủ thông tin!"); return
        
    # simple validations
    if "@" not in email or "." not in email:
        messagebox.showerror("Lỗi", "Email không hợp lệ!"); return
        
    # >> Bắt đầu kiểm tra mật khẩu theo tiêu chuẩn mới
    is_valid, error_msg = kiem_tra_mat_khau_hop_le(password, email)
    if not is_valid:
        messagebox.showerror("Lỗi", error_msg); return
    # << Kết thúc kiểm tra mật khẩu
    if password != confirm:
        messagebox.showerror("Lỗi", "Xác nhận mật khẩu không khớp!"); return
    if email in accounts:
        messagebox.showerror("Lỗi", "Email đã tồn tại!"); return
    accounts[email] = {"password": password, "role": "user", "borrowed_books": [], "history": []}
    save_accounts()
    messagebox.showinfo("OK", f"Đăng ký thành công với vai trò user!")
    switch_frame(frame_login)

def logout():
    # hide admin/user frames (they will be lower level frames so just switch)
    switch_frame(frame_login)
    # after logout, refresh admin/account views so admin list updated when re-login
    refresh_admin_books()
    refresh_accounts()
# ===================== LOGIN FRAME =====================
login_center = tk.Frame(frame_login, bg="white", bd=2, relief="groove", padx=40, pady=40)
login_center.place(relx=0.5, rely=0.5, anchor="center")
tk.Label(login_center, text="ĐĂNG NHẬP THƯ VIỆN", font=("Arial", 20, "bold"), fg="#0b60ff", bg="white").pack(pady=10)
tk.Label(login_center, text="Email:", bg="white").pack(anchor="w")
entry_login_email = tk.Entry(login_center, width=35, font=("Arial", 11))
entry_login_email.pack(pady=5)
tk.Label(login_center, text="Mật khẩu:", bg="white").pack(anchor="w")
entry_login_password = tk.Entry(login_center, width=35, font=("Arial", 11), show="*")
entry_login_password.pack(pady=5)
# ===================== LOGIN FRAME =====================
login_center = tk.Frame(frame_login, bg="white", bd=2, relief="groove", padx=40, pady=40)
login_center.place(relx=0.5, rely=0.5, anchor="center")

tk.Label(login_center, text="ĐĂNG NHẬP THƯ VIỆN",
         font=("Arial", 20, "bold"), fg="#0b60ff", bg="white").pack(pady=10)

tk.Label(login_center, text="Email:", bg="white").pack(anchor="w")
entry_login_email = tk.Entry(login_center, width=35, font=("Arial", 11))
entry_login_email.pack(pady=5)

tk.Label(login_center, text="Mật khẩu:", bg="white").pack(anchor="w")
entry_login_password = tk.Entry(login_center, width=35, font=("Arial", 11), show="*")
entry_login_password.pack(pady=5)

tk.Button(login_center, text="Đăng nhập", bg="#ffa500", fg="white",
          font=("Arial", 11, "bold"), width=30, command=login).pack(pady=15)

tk.Button(login_center, text="Chưa có tài khoản? Đăng ký ngay",
          bg="white", fg="#0077cc", bd=0,
          command=lambda: switch_frame(frame_register)).pack(pady=5)
# ===================== REGISTER FRAME =====================
reg_center = tk.Frame(frame_register, bg="white", bd=2, relief="groove", padx=40, pady=40)
reg_center.place(relx=0.5, rely=0.5, anchor="center")

tk.Label(reg_center, text="ĐĂNG KÝ TÀI KHOẢN",
         font=("Arial", 20, "bold"), bg="white", fg="#0b60ff").pack(pady=10)
# Email
tk.Label(reg_center, text="Email:", bg="white").pack(anchor="w")
entry_reg_email = tk.Entry(reg_center, width=35)
entry_reg_email.pack(pady=5)
# Password + eye button
tk.Label(reg_center, text="Mật khẩu:", bg="white").pack(anchor="w")
frame_password = tk.Frame(reg_center, bg="white"); frame_password.pack(pady=5)

entry_reg_password = tk.Entry(frame_password, width=28, show="*")
entry_reg_password.pack(side="left")

show_password = tk.BooleanVar(value=False)
def toggle_password():
    entry_reg_password.config(show="" if not show_password.get() else "*")
    show_password.set(not show_password.get())
btn_eye = tk.Button(frame_password, text="👁️", command=toggle_password, bd=0, bg="white")
btn_eye.pack(side="left", padx=2)

# Confirm password
tk.Label(reg_center, text="Nhập lại mật khẩu:", bg="white").pack(anchor="w")
frame_confirm = tk.Frame(reg_center, bg="white"); frame_confirm.pack(pady=5)

entry_reg_confirm = tk.Entry(frame_confirm, width=28, show="*")
entry_reg_confirm.pack(side="left")

show_confirm = tk.BooleanVar(value=False)
def toggle_confirm():
    entry_reg_confirm.config(show="" if not show_confirm.get() else "*")
    show_confirm.set(not show_confirm.get())
btn_eye_confirm = tk.Button(frame_confirm, text="👁️", command=toggle_confirm, bd=0, bg="white")
btn_eye_confirm.pack(side="left", padx=2)

# Register button
tk.Button(reg_center, text="Đăng ký", bg="#28a745", fg="white",
          font=("Arial", 11, "bold"), width=30,
          command=register).pack(pady=15)

tk.Button(reg_center, text="Quay lại đăng nhập",
          bg="white", fg="#555", bd=0,
          command=lambda: switch_frame(frame_login)).pack()

# ===================== ADMIN FRAME UI =====================
admin_header_frame = tk.Frame(frame_admin, bg="white", height=100); admin_header_frame.pack(fill="x", pady=0)
# Try to load logo if exists
def load_logo_for_header(frame, path="Logo.png", size=(80,80)):
    if not os.path.exists(path): return None
    try:
        img = Image.open(path)
        img = img.resize(size, Image.Resampling.LANCZOS)
        logo = ImageTk.PhotoImage(img)
        lbl = tk.Label(frame, image=logo, bg='white')
        lbl.image = logo
        return lbl
    except:
        return None

logo_img_admin = load_logo_for_header(admin_header_frame)
if logo_img_admin: logo_img_admin.pack(side="left", padx=20, pady=10)
title_frame = tk.Frame(admin_header_frame, bg="white"); title_frame.pack(side="left", padx=10)
tk.Label(title_frame, text="ĐẠI HỌC CÔNG THƯƠNG TP. HỒ CHÍ MINH", font=("Arial", 20, "bold"), fg="blue", bg="white").pack(anchor="w")
tk.Label(title_frame, text="TRUNG TÂM THÔNG TIN - THƯ VIỆN", font=("Arial", 14, "italic"), fg="red", bg="white").pack(anchor="w")

# Sidebar menu for admin (collapsible)
sidebar_open = False
menu_frame = tk.Frame(frame_admin, bg="#f0f0f0", width=200)

def toggle_menu():
    global sidebar_open
    if sidebar_open:
        menu_frame.place_forget()
    else:
        # Menu nằm bên PHẢI
        menu_frame.place(relx=1.0, y=60, anchor="ne", relheight=1, width=200)
        menu_frame.tkraise()
    sidebar_open = not sidebar_open


hamburger_btn = tk.Label(admin_header_frame, text="≡", font=("Arial", 26, "bold"), bg="white", cursor="hand2")
hamburger_btn.pack(side="right", padx=20)
hamburger_btn.bind("<Button-1>", lambda e: toggle_menu())

# Notebook for admin
admin_notebook = ttk.Notebook(frame_admin); admin_notebook.pack(fill="both", expand=True, padx=20, pady=20)
tab_Intro = tk.Frame(admin_notebook, bg="white"); admin_notebook.add(tab_Intro, text=" Trang chủ ")
tab_lib = tk.Frame(admin_notebook, bg="white"); admin_notebook.add(tab_lib, text=" Sách thư viện ")
tab_acc = tk.Frame(admin_notebook, bg="white"); admin_notebook.add(tab_acc, text=" Tài khoản ")
tab_stats = tk.Frame(admin_notebook, bg="white"); admin_notebook.add(tab_stats, text=" Thống kê ")
def show_page_admin(name):
    toggle_menu()
    if name == "Intro": admin_notebook.select(tab_Intro)
    elif name == "library": admin_notebook.select(tab_lib)
    elif name == "accounts": admin_notebook.select(tab_acc)
    elif name == "stats": admin_notebook.select(tab_stats)

def admin_do_logout():
    # admin logout -> reuse generic logout
    logout()
menu_items = [("📚  Trang chủ", lambda: show_page_admin("Intro")),
              ("📖  Sách thư viện", lambda: show_page_admin("library")),
              ("👤  Tài khoản", lambda: show_page_admin("accounts")),
              ("📊  Thống kê", lambda: show_page_admin("stats")),
             ("🔒  Đổi mật khẩu", lambda: change_password(current_email)),
                ("🚪  Đăng xuất", admin_do_logout)]
for txt, cmd in menu_items:
    tk.Button(menu_frame, text=txt, font=("Arial", 12), anchor="w", relief="flat", bg="#e8e8e8", command=cmd).pack(fill="x", pady=2, padx=5)

# Build Intro page
def build_intro(tab):
    content_frame = tk.Frame(tab, bg="white"); content_frame.pack(fill="both", expand=True)
    sub_sidebar = tk.Frame(content_frame, width=150, bg="#f8f8f8"); sub_sidebar.pack(side="left", fill="y")
    main = tk.Frame(content_frame, bg="white"); main.pack(side="right", fill="both", expand=True)
    content_text = tk.Text(main, wrap="word", font=("Arial", 11))
    content_text.pack(fill="both", expand=True, padx=10, pady=10)
    def display_doc(filename):
        content_text.delete("1.0", tk.END)
        if not os.path.exists(filename):
            content_text.insert(tk.END, f"⚠️ Không tìm thấy file: {filename}")
            return
        try:
            doc = Document(filename)
            for para in doc.paragraphs:
                content_text.insert(tk.END, para.text + "\n\n")
        except Exception as e:
            content_text.insert(tk.END, f"Lỗi: {e}")
    docs = {"Giới thiệu": "gioi_thieu.docx",
             "Hướng dẫn": "huong_dan.docx", 
             "Quy định": "quydinh.docx"}
    for name, fpath in docs.items(): tk.Button(sub_sidebar, text=name, font=("Arial", 12),command=lambda fp=fpath: display_doc(fp), relief="flat", bg="#ddd").pack(fill="x", pady=1)
    display_doc(docs["Giới thiệu"])
build_intro(tab_Intro)

# Build library page (admin)
def build_library(tab):
    global lib_search_entry, lib_field_var, tree_admin, lib_sort_field_var, lib_sort_order_var
    ctrl = tk.Frame(tab, bg="white"); ctrl.pack(fill="x", pady=10)
    tk.Label(ctrl, text="Tìm kiếm:", bg="white").pack(side="left", padx=10)
    lib_search_entry = tk.Entry(ctrl, width=30); lib_search_entry.pack(side="left", padx=5)
    lib_field_var = tk.StringVar(value="name")
    ttk.Combobox(ctrl, textvariable=lib_field_var, values=["name", "author", "category"], width=10, state="readonly").pack(side="left")
    tk.Button(ctrl, text="Tìm", command=lambda: do_search_admin()).pack(side="left", padx=5)
    tk.Button(ctrl, text="Làm mới", command=refresh_admin_books).pack(side="left")
    # ======= COMBOBOX PHÂN LOẠI =======
    tk.Label(ctrl, text="Phân loại:", bg="white").pack(side="left", padx=5)

    category_filter_var = tk.StringVar(value="Tất cả")

    category_list = [
        "Tất cả", "Lập trình", "Kỹ thuật", "Toán", "Hệ điều hành", "CSDL", "AI", "Phân tích",
        "Kiến trúc", "Mạng", "Điện tử", "Cloud", "Data", "Web", "Iot", "Robot", "Thiết kế",
        "Game", "Kinh tế", "Marketing", "Quản trị", "Kỹ năng"
    ]

    cat_box = ttk.Combobox(
        ctrl,
        textvariable=category_filter_var,
        values=category_list,
        width=15,
        state="readonly"
    )
    cat_box.pack(side="left", padx=5)

    def on_category_change(event=None):
        filter_by_category(category_filter_var.get())

    cat_box.bind("<<ComboboxSelected>>", on_category_change)

    cols = ("name", "author", "cat", "avail", "total", "bor")
    tree_admin = ttk.Treeview(tab, columns=cols, show="headings")
    for c, t, w in zip(cols, ["Tên sách", "Tác giả", "Thể loại", "Còn", "Tổng", "Đã mượn"], [250, 150, 100, 60, 60, 60]):
        tree_admin.heading(c, text=t); tree_admin.column(c, width=w, anchor='center')
    tree_admin.column("name", anchor="w")
    tree_admin.pack(fill="both", expand=True, padx=10)
    tree_admin.bind("<Double-1>", on_admin_book_double_click)
    btns = tk.Frame(tab, bg="white"); btns.pack(pady=10)
    tk.Button(btns, text="➕ Thêm", bg="#28a745", fg="white", width=12, command=add_book_admin).pack(side="left", padx=10)
    tk.Button(btns, text="✏️ Sửa", bg="#ffc107", width=12, command=edit_book_admin).pack(side="left", padx=10)
    tk.Button(btns, text="❌ Xóa", bg="#dc3545", fg="white", width=12, command=delete_book_admin).pack(side="left", padx=10)
build_library(tab_lib)

# Build accounts page (admin) - shows accounts with borrowed count summary
def build_accounts(tab):
    global tree_accounts
    global role_filter_var

    ctrl = tk.Frame(tab, bg="white")
    ctrl.pack(fill="x", pady=10,padx=10)

    tk.Button(ctrl, text="🔄 Làm mới DS", command=refresh_accounts).pack(side="left")
    
    # Label cho Combobox
    tk.Label(ctrl, text="Lọc theo vai trò:").pack(side="left", padx=(15, 5)) 

    combo_filter_roles = ttk.Combobox(
        ctrl, 
        textvariable=role_filter_var, 
        values=("Tất cả", "admin", "user"),
        state="readonly", # Ngăn người dùng nhập giá trị tùy ý
        width=10
    )
    combo_filter_roles.pack(side="left")

    # Gắn sự kiện: Khi giá trị Combobox thay đổi, gọi hàm làm mới
    combo_filter_roles.bind("<<ComboboxSelected>>", lambda event: refresh_accounts(role_filter_var.get()))
    # 👉 Thêm cột mới: borrowed_list
    tree_accounts = ttk.Treeview(
        tab,
        columns=("email", "password","role", "borrow_count", "borrowed_list"),
        show="headings"
    )

    tree_accounts.heading("email", text="Email / User")
    tree_accounts.column("email", width=200)

    tree_accounts.heading("password", text="Mật khẩu")
    tree_accounts.column("password", width=150)

    tree_accounts.heading("role", text="Vai trò")
    tree_accounts.column("role", width=80, anchor="center")

    tree_accounts.heading("borrow_count", text="Đang mượn")
    tree_accounts.column("borrow_count", width=80, anchor="center")

    # 👉 NEW COLUMN: List of books being borrowed
    tree_accounts.heading("borrowed_list", text="Sách đang mượn")
    tree_accounts.column("borrowed_list", width=350)

    tree_accounts.pack(fill="both", expand=True, padx=10)
    tree_accounts.bind("<Double-1>", on_admin_user_double_click)
    tk.Button(
        tab, 
        text="➕ Tạo tài khoản",
          bg="#0d6efd", 
          font=("Arial", 15, "bold"),
          fg="white",
          width=15,
            command=open_create_account_window
            ).pack(side="left", padx=5)
    tk.Button(
        tab,
        text="🗑 Xóa tài khoản",
        bg="#dc3545",
        font=("Arial", 15, "bold"),
        fg="white",
        width=15,
        command=delete_account_admin
    ).pack(side="left", padx=10)
build_accounts(tab_acc)

# Build stats page (admin)
def build_stats(tab):
    global tree_top, tree_rare
    tk.Label(tab, text="🔥 Top sách được mượn nhiều nhất", bg="white", font=("Arial", 12, "bold"), fg="#d35400").pack(anchor="w", padx=10, pady=(15,5))
    tree_top = ttk.Treeview(tab, columns=("name", "author", "bor"), show="headings", height=6)
    tree_top.heading("name", text="Tên sách"); tree_top.column("name", width=400)
    tree_top.heading("author", text="Tác giả"); tree_top.column("author", width=250)
    tree_top.heading("bor", text="Số lần mượn"); tree_top.column("bor", anchor="center")
    tree_top.pack(fill="x", padx=10)
    tk.Label(tab, text="❄️ Sách ít được mượn (Gợi ý thanh lý)", bg="white", font=("Arial", 12, "bold"), fg="#2980b9").pack(anchor="w", padx=10, pady=(15,5))
    tree_rare = ttk.Treeview(tab, columns=("name", "author", "bor"), show="headings", height=6)
    tree_rare.heading("name", text="Tên sách"); tree_rare.column("name", width=400)
    tree_rare.heading("author", text="Tác giả"); tree_rare.column("author", width=250)
    tree_rare.heading("bor", text="Số lần mượn"); tree_rare.column("bor", anchor="center")
    tree_rare.pack(fill="x", padx=10)
    tk.Button(tab, text="Cập nhật thống kê", command=refresh_stats).pack(pady=10)
build_stats(tab_stats)

# ===================== USER FRAME =====================
user_head = tk.Frame(frame_user, bg="white", height=60); user_head.pack(fill="x")

tk.Label(user_head, text="CỔNG SINH VIÊN", font=("Arial", 16, "bold"), bg="white", fg="#0b60ff").pack(side="left", padx=20, pady=15)
lbl_status = tk.Label(user_head, text="Trạng thái: ...", font=("Arial", 11, "bold"), bg="white"); lbl_status.pack(side="left", padx=30)
tk.Button(user_head, text="Đăng xuất", bg="#ff4d4d", fg="white", command=logout).pack(side="right", padx=10)
btn_change_pass = tk.Button(
    user_head,
    text="Đổi mật khẩu",
    bg="#4CAF50",
    fg="white",
    padx=10,
    pady=5,
    command=lambda: change_password(current_email)
)
btn_change_pass.pack(side="right",pady=10)
# user controls & trees
user_nb = ttk.Notebook(frame_user); user_nb.pack(fill="both", expand=True, padx=20, pady=20)
# Tab: Giới thiệu/Quy định
tab_user_intro = tk.Frame(user_nb, bg="white"); 
user_nb.add(tab_user_intro, text=" 🏠 Giới thiệu/Quy định ")
build_intro(tab_user_intro)
# Tab: Xem sách thư viện
tab_lib_view = tk.Frame(user_nb, bg="white");
user_nb.add(tab_lib_view, text=" 📚 Sách Thư viện ")
# GỌI HÀM BUILD MỚI: Chỉ cần gọi hàm này một lần duy nhất tại đây.
def build_library_view(tab):
    # Khai báo các biến global cần thiết cho tính năng tìm kiếm
    global lib_search_entry_view, lib_field_var_view, tree_library_view 

    # 1. Khung tìm kiếm (Giữ nguyên)
    ctrl = tk.Frame(tab, bg="white"); 
    ctrl.pack(fill="x", pady=10)
    
    tk.Label(ctrl, text="Tìm kiếm:", bg="white").pack(side="left", padx=10)
    
    lib_search_entry_view = tk.Entry(ctrl, width=30); 
    lib_search_entry_view.pack(side="left", padx=5)
    
    lib_field_var_view = tk.StringVar(value="name")
    # Thêm "date" vào danh sách tìm kiếm
    ttk.Combobox(ctrl, textvariable=lib_field_var_view, values=["name", "author", "category", "date"], width=10, state="readonly").pack(side="left")
    
    tk.Button(ctrl, text="Tìm", command=lambda: do_search_library_view()).pack(side="left", padx=5)
    tk.Button(ctrl, text="Làm mới", command=lambda: refresh_library_view()).pack(side="left")
    
    cols = ("name", "author", "cat", "date") 
    
    tree_library_view = ttk.Treeview(tab, columns=cols, show="headings")
    
    # Định nghĩa tiêu đề và kích thước
    # Thêm cột 'Ngày XB' (width=100)
    for c, t, w in zip(cols, ["Tên sách", "Tác giả", "Thể loại", "Ngày XB"], [300, 180, 150, 100]):
        tree_library_view.heading(c, text=t)
        tree_library_view.column(c, width=w, anchor='center' if c == "date" else 'w')
    
    tree_library_view.column("name", anchor="w")
    tree_library_view.pack(fill="both", expand=True, padx=10)
    tree_library_view.bind("<Double-1>", on_user_book_double_click)
    refresh_library_view()
build_library_view(tab_lib_view)
# Tab: Sách đang mượn
tab_bor = tk.Frame(user_nb, bg="white"); user_nb.add(tab_bor, text=" 📖 Sách đang mượn ")
cols_b = ("name", "author", "qty", "bd", "dd")
tree_bor = ttk.Treeview(tab_bor, columns=cols_b, show="headings")
for c, t, w in zip(cols_b, ["Tên sách", "Tác giả", "SL", "Ngày mượn", "Hạn trả"], [250, 150, 60, 110, 110]):
    tree_bor.heading(c, text=t); tree_bor.column(c, width=w, anchor="center")
tree_bor.pack(fill="both", expand=True, padx=10, pady=10)
btn_bor = tk.Frame(tab_bor, bg="white"); btn_bor.pack(pady=10)
tk.Button(btn_bor, text="➕ Mượn sách (Thư viện)", bg="#28a745", fg="white", font=("Arial", 10, "bold"), width=25, height=2, command=lambda: borrow_from_library()).pack(side="left", padx=20)
tk.Button(btn_bor, text="↩️ Trả sách đã chọn", bg="#ffc107", font=("Arial", 10, "bold"), width=25, height=2, command=lambda: return_book_user()).pack(side="left", padx=20)

# Tab: Quản lý cá nhân (danh sách cá nhân + lịch sử)
tab_per = tk.Frame(user_nb, bg="white"); user_nb.add(tab_per, text=" 📂 Quản lý cá nhân ")
pc = tk.Frame(tab_per, bg="white"); pc.pack(fill="x", pady=10, padx=10)
user_search_entry = tk.Entry(pc, width=30); user_search_entry.pack(side="left")
tk.Button(pc, text="Tìm tên", command=lambda: do_search_user()).pack(side="left", padx=5)
tk.Button(pc, text="Làm mới", command=lambda: refresh_user_books()).pack(side="left")
cols_p = ("name", "author", "qty", "st", "dd")
tree_user = ttk.Treeview(tab_per, columns=cols_p, show="headings")
for c, t in zip(cols_p, ["Tên sách", "Tác giả", "SL", "Trạng thái", "Hạn trả"]):
    tree_user.heading(c, text=t); tree_user.column(c, anchor="center")
tree_user.pack(fill="both", expand=True, padx=10)
pb = tk.Frame(tab_per, bg="white"); pb.pack(pady=10)
# Label tổng tiền nợ
lbl_total_fine = tk.Label(
    tab_per, 
    textvariable=total_fine_var,  # Dùng textvariable để giá trị tự động cập nhật
    font=("Arial", 14, "bold"), 
    fg="#dc3545", 
    bg="white"
)
# Đặt Label này ở vị trí bạn muốn, ví dụ: ngay dưới tab_per
lbl_total_fine.pack(pady=5)
tk.Button(pb, text="Thêm sách ngoài", command=lambda: add_book_user()).pack(side="left", padx=10)
tk.Button(pb, text="Xóa sách", bg="#dc3545", fg="white", command=lambda: delete_book_user()).pack(side="left", padx=10)

# Tab: Lịch sử (user)
tab_his = tk.Frame(user_nb, bg="white"); user_nb.add(tab_his, text=" 🕒 Lịch sử ")
tree_user_history = ttk.Treeview(tab_his, columns=("name", "qty", "act", "bd", "rd"), show="headings")
for c, t in zip(["name", "qty", "act", "bd", "rd"], ["Tên sách", "SL", "Hành động", "Ngày mượn", "Ngày trả"]):
    tree_user_history.heading(c, text=t); tree_user_history.column(c, anchor="center")
tree_user_history.pack(fill="both", expand=True, padx=10, pady=10)

# ===================== START =====================
refresh_admin_books()
refresh_accounts()
refresh_stats()

switch_frame(frame_login)
win.mainloop()
