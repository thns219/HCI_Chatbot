"""
SUPEO chatbot - bản chạy độc lập (không có cell đo/đánh giá), tạo từ supeo_chatbot_v5.ipynb.
Cần cùng thư mục (hoặc cùng Kaggle Dataset): supeo_app.py, supeo_api.py, supeo_data_complete.xlsx
Chạy trong Kaggle (giữ được Secret GEMINI_API_KEY):
    import sys; sys.path.insert(0, "<thư mục chứa file>")
    exec(open("<thư mục chứa file>/supeo_app.py").read())
Chạy local:  pip install "gradio>=5,<6" fastapi "uvicorn[standard]" requests pandas openpyxl
             SUPEO_SHARE=0 python supeo_app.py     (Windows: xem hướng dẫn kèm theo) -> mở http://127.0.0.1:7860
"""
import os, sys
_HERE = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
LAUNCH = os.environ.get("SUPEO_NO_LAUNCH") != "1"      # SUPEO_NO_LAUNCH=1: chỉ nạp bot, không mở giao diện



# ==================== API + tìm Excel ====================
import os, glob, time, threading, requests, uvicorn

# --- 1) Tìm file dữ liệu Excel ---------------------------------------------------
def ensure_xlsx(name="supeo_data_complete.xlsx"):
    cands = [os.environ.get("SUPEO_XLSX"), name, f"/kaggle/working/{name}", f"/content/{name}"]
    for pat in (f"/kaggle/input/**/{name}", f"/content/**/{name}"):
        cands += glob.glob(pat, recursive=True)
    for c in cands:
        if c and os.path.exists(c):
            return os.path.abspath(c)
    try:                                           # Google Colab: cho upload tay
        from google.colab import files
        print("Chọn file supeo_data_complete.xlsx để upload...")
        files.upload()
        if os.path.exists(name):
            return os.path.abspath(name)
    except ImportError:
        pass
    raise FileNotFoundError(
        "Không thấy supeo_data_complete.xlsx. Trên Kaggle: Add Data -> Upload -> chọn file Excel, rồi chạy lại cell này.")

os.environ["SUPEO_XLSX"] = ensure_xlsx()
# Ngày KHÔNG cố định ở đây: các cell đo (9c, 9d, 9e) tự đặt SUPEO_TODAY=2026-10-28 bằng fixed_today() rồi trả lại giá trị cũ.
print("Dữ liệu:", os.environ["SUPEO_XLSX"])

# --- 2) Chạy FastAPI trong thread nền ---------------------------------------------
API_PORT = 8008
API_URL = f"http://127.0.0.1:{API_PORT}"

def start_api():
    try:
        if requests.get(API_URL + "/health", timeout=1).ok:
            return "đã chạy sẵn"
    except requests.RequestException:
        pass
    import importlib, supeo_api
    importlib.reload(supeo_api)
    server = uvicorn.Server(uvicorn.Config(supeo_api.app, host="127.0.0.1", port=API_PORT, log_level="warning"))
    threading.Thread(target=server.run, daemon=True).start()
    for _ in range(50):
        try:
            if requests.get(API_URL + "/health", timeout=1).ok:
                return "vừa khởi động"
        except requests.RequestException:
            time.sleep(0.2)
    raise RuntimeError("API không khởi động được")

print("API:", start_api())
print(requests.get(API_URL + "/health").json())
print("Swagger UI (khi chạy máy local):", API_URL + "/docs")


# ==================== Client + lexicon ====================
import re, json, time, unicodedata, requests
from urllib.parse import quote

API_URL = globals().get("API_URL", "http://127.0.0.1:8008")
API_LOG = []        # nhật ký các lần gọi API của lượt hội thoại hiện tại (hiện ở panel debug)
USE_LLM = False     # cell LLM bên dưới sẽ bật lên True nếu tải được mô hình

class ApiDown(Exception):
    pass

def strip_acc(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")

def vnd(n):
    return f"{int(n):,}".replace(",", ".") + "đ"

def api_get(path, **params):
    """GET tới Mock API. Trả (status_code, json). Mất kết nối -> ApiDown."""
    params = {k: v for k, v in params.items() if v is not None and v != ""}
    log = {"GET": path, "params": params}
    try:
        r = requests.get(API_URL + path, params=params, timeout=8)
        log["status"] = r.status_code
        data = r.json()
    except Exception as e:
        log["status"], log["error"] = 0, str(e)[:100]
        API_LOG.append(log)
        raise ApiDown(str(e))
    API_LOG.append(log)
    return r.status_code, data

# --- Nạp "từ vựng" từ chính API (không hardcode danh mục/dịp/mã) ----------------
CATALOG, CATEGORIES, OCCASIONS, VOUCHER_CODES = [], [], [], []
DISTINCT, COLOR_LEX, UNIQUE_TOK = {}, {}, {}

CAT_WORDS = {"ao", "chan", "vay", "midi", "blazer", "so", "mi", "quan", "ong", "rong", "dam"}
STOP = {"co", "la", "cho", "va", "voi", "tay", "dang", "mau", "bo", "can", "cao"}
MATERIAL_LEX = {"linen": "linen", "satin": "satin", "tuyet mua": "tuyết mưa", "kaki": "kaki", "cotton": "cotton", "voan": "voan",
                "tweed": "tweed", "dui": "đũi", "thun": "thun", "chiffon": "chiffon", "poplin": "poplin", "oxford": "oxford", "lua": "lụa"}

def load_lexicons():
    global CATALOG, CATEGORIES, OCCASIONS, VOUCHER_CODES, DISTINCT, COLOR_LEX, UNIQUE_TOK
    _, d = api_get("/v1/products", limit=200)
    CATALOG, CATEGORIES = d["items"], d["danh_muc"]
    _, d = api_get("/v1/lookbook");       OCCASIONS = d["dip_mac"]
    _, d = api_get("/v1/vouchers/codes"); VOUCHER_CODES = d["codes"]
    DISTINCT = {}
    for p in CATALOG:                                  # từ "đặc trưng" của tên từng sản phẩm
        toks = set(re.findall(r"[a-z0-9]+", strip_acc(p["ten"]).lower()))
        DISTINCT[p["id"]] = {t for t in toks - CAT_WORDS - STOP if len(t) > 1}
    cnt = {}                                           # từ chỉ xuất hiện ở ĐÚNG 1 sản phẩm (culottes, babydoll...) -> nhận diện món
    for pid, toks in DISTINCT.items():
        for tk in toks:
            cnt.setdefault(tk, []).append(pid)
    UNIQUE_TOK = {tk: ids[0] for tk, ids in cnt.items() if len(ids) == 1}
    cols = {strip_acc(c).lower(): c.lower() for p in CATALOG for c in p["mau"]}       # {không dấu: có dấu}
    COLOR_LEX = dict(sorted(cols.items(), key=lambda kv: -len(kv[0])))
    API_LOG.clear()
    return {"san_pham": len(CATALOG), "danh_muc": CATEGORIES, "dip": OCCASIONS, "ma_voucher": VOUCHER_CODES}

print(load_lexicons())


# ==================== Schema + DST ====================
# ============ BƯỚC 1: KHAI BÁO SCHEMA (SGP-TOD) ============
INTENTS = ["product_info", "size_advice", "lookbook", "order_check", "voucher", "chitchat"]
SCHEMA = {                       # 6 slot cốt lõi + 3 slot phụ trợ
    "loai_do":   "Áo Blazer | Chân váy Midi | Áo sơ mi | Quần ống rộng | Đầm",
    "dip_mac":   "dịp mặc (Hẹn hò, Trà chiều, Công sở, Tiệc tối, Dạo phố, Gặp đối tác...)",
    "chieu_cao": "cm (int)", "can_nang": "kg (int)",
    "ma_don":    "#SPxxxx", "ma_voucher": "mã giảm giá khách nhập",
    "ten_sp": "tên sản phẩm cụ thể (khớp danh mục)", "tu_khoa": "màu / chất liệu", "cart_total": "giá trị giỏ hàng (VNĐ)",
}
NEEDED = {                       # slot bắt buộc theo intent (điều kiện đủ ở hàm missing)
    "product_info": ["loai_do"], "size_advice": ["chieu_cao", "can_nang"],
    "lookbook": ["dip_mac"], "order_check": ["ma_don"], "voucher": [], "chitchat": [],
}
FOLLOW_KEYS = {"product_info": ["loai_do", "ten_sp", "tu_khoa"], "size_advice": ["chieu_cao", "can_nang"],
               "lookbook": ["dip_mac"], "order_check": ["ma_don"]}

def missing(intent, S):
    if intent == "product_info":
        return [] if any(k in S for k in ("loai_do", "ten_sp", "tu_khoa")) else ["loai_do"]
    return [k for k in NEEDED.get(intent, []) if k not in S]

# Từ khóa nhận diện (đã bỏ dấu). Thứ tự = độ ưu tiên: cụ thể trước, chung sau.
CATS = {
    "Đầm": [r"dam"],
    "Áo Blazer": [r"blazer", r"vest"],
    "Áo sơ mi": [r"so mi"],
    "Chân váy Midi": [r"chan vay", r"vay midi", r"midi"],
    "Quần ống rộng": [r"quan(?! trong| ao)", r"culottes"],
}
DIPS = {
    "Hẹn hò thơ mộng": [r"hen ho tho mong", r"tho mong"],
    "Hẹn hò": [r"hen ho"],
    "Gặp đối tác": [r"doi tac"],
    "Công sở": [r"cong so", r"di lam"],
    "Casual Friday": [r"casual"],
    "Du lịch hè": [r"du lich"],
    "Tiệc cocktail": [r"cocktail"],
    "Dự tiệc ngoài trời": [r"ngoai troi", r"garden party"],
    "Tiệc tối": [r"tiec toi", r"du tiec", r"tiec"],
    "Trà chiều": [r"tra chieu"],
    "Dạo phố mùa thu": [r"dao pho mua thu", r"dao pho thu"],   # nhóm chung nhất đặt CUỐI ('cuối tuần đi tiệc' phải ra Tiệc, không phải Dạo phố)
    "Dạo phố": [r"dao pho", r"di choi", r"cuoi tuan"],
}

NEG = re.compile(r"\bchu\s+(?:khong|dung)\b|\bkhong\s+(?:phai|muon|can)\b")

def positive_part(text):
    """Bỏ vế phủ định: 'váy chứ không phải quần ống rộng' -> 'váy'; 'không phải quần, mà là váy' -> 'váy'."""
    nfc = unicodedata.normalize("NFC", text).lower()
    t = strip_acc(nfc)
    m = NEG.search(t) if len(t) == len(nfc) else None
    if not m:
        return nfc
    before = nfc[:m.start()].strip(" ,.;")
    if re.search(r"\w", strip_acc(before)):
        return before
    m2 = re.search(r"\b(?:ma|la)\s+(.*)$", t[m.end():])
    return nfc[m.end():][m2.start(1):] if m2 else nfc

def is_generic_vay(ptxt):
    """'váy' chung (đầm + chân váy). Cần dấu để không nhầm với 'vậy'; không dấu thì phải đi sau mua/xem/tìm..."""
    return bool(re.search(r"\bváy\b", ptxt) or re.search(r"\b(?:mua|xem|tim|muon|can|cho)\s+vay\b", ptxt))

ORD = {"nhat": 1, "dau tien": 1, "dau": 1, "hai": 2, "ba": 3, "bon": 4, "nam": 5, "sau": 6, "bay": 7, "tam": 8, "chin": 9, "muoi": 10}
ORD_RE = r"\b(?:mau|cai|so|thu)\s+(?:thu\s+)?(nhat|dau tien|dau|hai|ba|bon|nam|sau|bay|tam|chin|muoi|\d{1,2})\b"

def pick_item(text, lst):
    """Khách chọn mẫu bằng mã (P03) hoặc số thứ tự trong danh sách bot vừa gửi ('mẫu thứ 2', 'cái đầu tiên')."""
    t = strip_acc(text).lower()
    m = re.search(r"\bp(\d{2})\b", t)
    if m:
        p = next((x for x in CATALOG if x["id"] == "P" + m.group(1)), None)
        if p:
            return p["ten"]
    if lst:
        m = re.search(ORD_RE, t)
        if m:
            k = ORD.get(m.group(1)) or (int(m.group(1)) if m.group(1).isdigit() else None)
            if k and 1 <= k <= len(lst):
                return lst[k - 1]
    return None

def match_key(text, table):
    t = strip_acc(text).lower()
    for key, pats in table.items():
        if any(re.search(r"\b" + p + r"\b", t) for p in pats):
            return key
    return None

def match_product(text, cat):
    """Khớp tên sản phẩm trong danh mục (tải từ API). Ưu tiên tên đầy đủ, sau đó từ đặc trưng trong cùng loại."""
    t = strip_acc(text).lower()
    words = set(re.findall(r"[a-z0-9]+", t))
    best = None
    for p in CATALOG:
        if strip_acc(p["ten"]).lower() in t and (best is None or len(p["ten"]) > len(best["ten"])):
            best = p
    if best:
        return best["ten"]
    uniq = [p for p in CATALOG if (not cat or p["loai"] == cat)
            and any(len(tk) >= 7 and tk in words and UNIQUE_TOK.get(tk) == p["id"] for tk in DISTINCT.get(p["id"], ()))]
    if len(uniq) == 1:                                       # vd "culottes", "babydoll", "croptop"
        return uniq[0]["ten"]
    if not cat:
        return None
    scored = []
    for p in CATALOG:
        d = DISTINCT.get(p["id"])
        if p["loai"] == cat and d:
            hit = len(d & words)
            if hit:
                scored.append((hit / len(d), p["ten"]))
    scored.sort(reverse=True)
    if scored and scored[0][0] >= 0.6 and (len(scored) == 1 or scored[0][0] > scored[1][0]):
        return scored[0][1]
    return None

def match_keywords(text):
    t = strip_acc(text).lower()
    kws = [vi for k, vi in MATERIAL_LEX.items() if re.search(r"\b" + k + r"\b", t)]
    kws += [vi for k, vi in COLOR_LEX.items() if re.search(r"\bmau\s+" + re.escape(k) + r"\b", t)][:1]
    return " ".join(kws) or None

def parse_money(text):
    t = strip_acc(text).lower()
    val = None
    m = re.search(r"\b(\d+)\s?(?:tr|trieu)\s?(\d{1,3})?\b", t)
    if m:
        val = int(m.group(1)) * 1_000_000
        if m.group(2):
            frac = m.group(2)
            val += int(frac) * {1: 100_000, 2: 10_000, 3: 1_000}[len(frac)]
    if val is None:
        m = re.search(r"\b(\d+(?:[.,]\d+)?)\s?(?:k|nghin|ngan)\b", t)
        if m:
            val = int(float(m.group(1).replace(",", ".")) * 1000)
    if val is None:
        m = re.search(r"\b(\d{1,3}(?:[.,]\d{3})+)\b", t) or re.search(r"\b(\d{6,8})\b", t)
        if m:
            val = int(re.sub(r"[.,]", "", m.group(1)))
    return val if val and 50_000 <= val <= 100_000_000 else None

def regex_slots(text):
    """Regex bóc mã đơn, chiều cao, cân nặng, mã voucher (ưu tiên hơn LLM vì chính xác tuyệt đối)."""
    t = strip_acc(text).lower()
    out = {}
    m = re.search(r"#?\bsp\s?-?(\d{3,})\b", t)
    if m:
        out["ma_don"] = "#SP" + m.group(1)
    m = re.search(r"\b1\s?m\s?(\d{1,2})\b", t)                               # 1m65, 1m6
    if m:
        n = m.group(1)
        out["chieu_cao"] = 100 + (int(n) * 10 if len(n) == 1 else int(n))
    else:
        m = re.search(r"\b(1[4-9]\d|20\d)\s?(?:cm|xang|phan)\b", t) or re.search(r"cao\s*(?:khoang\s*)?(1[4-9]\d|20\d)\b", t)
        if m:
            out["chieu_cao"] = int(m.group(1))
    m = re.search(r"\b(\d{2,3})\s?(?:kg|ki|kilo|ky)\b", t) or re.search(r"\bnang\s*(?:khoang\s*)?(\d{2,3})\b", t)
    if m:
        out["can_nang"] = int(m.group(1))
    for code in VOUCHER_CODES:                                                # mã đã biết (không phân biệt hoa thường)
        if re.search(r"\b" + code.lower() + r"\b", t):
            out["ma_voucher"] = code
            break
    else:                                                                     # mã lạ viết HOA: SUMMER2026...
        m = re.search(r"\b(?!SP\d)[A-Z][A-Z0-9]{5,}\b", text)
        if m and m.group(0) != "SUPEO":
            out["ma_voucher"] = m.group(0)
    return out

def field_asked(text):
    t = strip_acc(text).lower()
    if re.search(r"\bsize\b", t) and re.search(r"(nhung|may|co) size|size nao|size gi co|du size", t):
        return "size"
    if re.search(r"chat lieu|chat vai|vai gi|vai nao|lam bang", t):
        return "chat_lieu"
    nfc = unicodedata.normalize("NFC", text).lower()
    # "màu" (có dấu) khác "mẫu": bỏ dấu thì "mẫu thứ 2 giá bao nhiêu" bị hiểu nhầm là hỏi màu
    if re.search(r"\bmàu\b", nfc) if strip_acc(nfc) != nfc else re.search(r"\bmau\b", t):
        return "mau"
    if re.search(r"con hang|het hang|ton kho|con khong|con mau", t):
        return "ton_kho"
    if re.search(r"\bgia\b|bao nhieu|bao nhieu tien|nhieu tien", t):
        return "gia"
    if re.search(r"hinh anh|xem hinh|\b(?:xem|gui|coi|thay)\s+(?:duoc\s+)?(?:cai\s+|nhung\s+)?anh\b|\banh\s+(?:san pham|mau|that|thuc te)\b", t):
        return "anh"                                              # khách muốn xem ảnh sản phẩm
    return None

PERSONAL_SIZE = (r"\b(?:minh|em|toi|chi|con)\s+(?:mac|lay|chon|dung)\b|\bnen\s+(?:mac|lay|chon)\b|\bmac size\b"
                 r"|\bsize cua (?:minh|em|toi)\b|\bhop size\b")

def rule_intent(text, new, state_slots):
    t = strip_acc(text).lower()
    S = {**state_slots, **new}
    if "ma_don" in new or re.search(r"don hang|van don|check don|kiem tra don|tra cuu don|theo doi don|giao toi dau|ship toi dau|don cua minh", t):
        return "order_check"
    if "ma_voucher" in new or "cart_total" in new or re.search(r"voucher|ma giam|giam gia|khuyen mai|uu dai|gio hang|ap ma|ma nao|coupon|freeship", t):
        return "voucher"
    asks_size_list = (field_asked(text) == "size" and "chieu_cao" not in new and "can_nang" not in new
                      and not re.search(PERSONAL_SIZE, t))
    if not asks_size_list and ("chieu_cao" in new or "can_nang" in new or re.search(r"\bsize\b|mac vua|co nguoi|nang bao nhieu|cao bao nhieu", t)):
        return "size_advice"
    if "dip_mac" in new or re.search(r"phoi|lookbook|goi y (?:bo|set|do|outfit)|set do|outfit", t):
        return "lookbook"
    if any(k in new for k in ("loai_do", "ten_sp", "tu_khoa")) or field_asked(text) or re.search(r"mau (?:nao|gi) |san pham|mat hang", t):
        return "product_info"
    return "chitchat"

# ============ BƯỚC 2: ĐIỀN KHUNG (DST) bằng LLM (tùy chọn) + Regex + luật ============
DST_SYS = """Bạn là bộ trích xuất thông tin (DST) cho chatbot thời trang SUPEO.
Trả về DUY NHẤT một JSON hợp lệ, không giải thích:
{"intent": <product_info|size_advice|lookbook|order_check|voucher|chitchat>,
 "slots": {"loai_do": <Áo Blazer|Chân váy Midi|Áo sơ mi|Quần ống rộng|Đầm|null>,
           "dip_mac": <tên dịp mặc hoặc null>, "chieu_cao": <số cm hoặc null>, "can_nang": <số kg hoặc null>,
           "ma_don": <#SPxxxx hoặc null>}}
Chỉ lấy thông tin trong tin nhắn mới nhất; thiếu thì để null.
Dùng "Lịch sử gần nhất" và "Đang nói về" để hiểu các từ như 'cái này', 'mẫu đó', 'đơn đó', nhưng chỉ điền slot khi tin nhắn mới có nhắc hoặc ám chỉ rõ."""

def dst_input(text, last_bot, cur_intent, hist="", focus=""):
    return (f"Lịch sử gần nhất:\n{hist or 'không'}\nĐang nói về: {focus or 'chưa có'}\n"
            f"Bot vừa hỏi: {last_bot or 'không'}\nIntent đang xử lý: {cur_intent}\nTin nhắn khách: {text}")

_E = '"loai_do":%s,"dip_mac":%s,"chieu_cao":%s,"can_nang":%s,"ma_don":%s'
FEWSHOT = [      # ví dụ TỰ VIẾT, cố ý KHÔNG trùng câu nào trong dialogue_benchmarks / bộ held-out (xem cell kiểm tra rò rỉ)
    ("Chị cao 1m72 thì lấy size nào nhỉ?", "none", "chitchat",
     '{"intent":"size_advice","slots":{' + _E % ("null", "null", "172", "null", "null") + '}}'),
    ("Tra giúp em đơn SP4455", "none", "chitchat",
     '{"intent":"order_check","slots":{' + _E % ("null", "null", "null", "null", '"#SP4455"') + '}}'),
    ("60 cân", "Nàng cho SUPEO xin cân nặng (kg) nhé!", "size_advice",
     '{"intent":"size_advice","slots":{' + _E % ("null", "null", "null", "60", "null") + '}}'),
    ("Mình muốn xem blazer cho đi làm", "none", "chitchat",
     '{"intent":"product_info","slots":{' + _E % ('"Áo Blazer"', '"Công sở"', "null", "null", "null") + '}}'),
    ("Có freeship không bạn", "none", "chitchat",
     '{"intent":"voucher","slots":{' + _E % ("null", "null", "null", "null", "null") + '}}'),
    ("Hello bạn", "none", "chitchat",
     '{"intent":"chitchat","slots":{' + _E % ("null", "null", "null", "null", "null") + '}}'),
]

def llm_dst(text, state):
    msgs = [{"role": "system", "content": DST_SYS}]
    for u, lb, ci, a in FEWSHOT:
        msgs += [{"role": "user", "content": dst_input(u, lb, ci)}, {"role": "assistant", "content": a}]
    msgs.append({"role": "user", "content": dst_input(text, state["last_bot"], state["intent"],
                                                      recent_history(state), focus_text(state))})
    try:
        raw = llm(msgs, 120)
        return json.loads(re.search(r"\{.*\}", raw, re.S).group(0))
    except Exception:
        return None

FOLLOW_ORDER = r"\b(?:giao|khi nao|bao gio|toi noi|den noi|van chuyen|van don|trang thai|hoan|huy|ship|don nay|don do|don kia)\b"
DEICTIC_ACC = re.compile(r"\b(?:áo|đầm|váy|quần|blazer|mẫu|sản phẩm|món|cái|set|bộ)\s+(?:này|đó|kia|ấy|vừa rồi|trên)\b")
DEICTIC_NOACC = re.compile(r"\b(?:ao|dam|vay|quan|blazer|san pham|mon|cai)\s+(?:nay|kia|ay)\b|\bmau\s+(?:nay|kia)\b|\bcai do\b")

def resolve_reference(text, ptxt, cat, state):
    """Hiểu "áo này", "mẫu đó", hoặc câu hỏi tiếp (giá? còn hàng không?) mà không nêu món -> món đang nói tới (bộ nhớ focus)."""
    f = state.get("focus", {})
    if not f.get("ten_sp"):
        return None
    nfc = unicodedata.normalize("NFC", text).lower()
    na = strip_acc(nfc)
    if DEICTIC_ACC.search(nfc) or DEICTIC_NOACC.search(na):
        if cat and cat != "Váy" and cat != f.get("loai_do"):
            return None                                      # "Áo này" nhưng món đang nói là đầm -> không ép
        return f["ten_sp"]
    if field_asked(text) and not cat and not state["slots"].get("ten_sp") \
            and not re.search(r"voucher|ma giam|gio hang|don hang|ma don|\bsp\d", na):
        return f["ten_sp"]
    return None

def recent_history(state, n=3):
    h = state.get("history", [])[-2 * n:]
    return "\n".join(("Khách: " if r == "user" else "SUPEO: ") + str(m)[:160].replace("\n", " ") for r, m in h)

def focus_text(state):
    f, lab = state.get("focus", {}), {"ten_sp": "sản phẩm", "loai_do": "nhóm đồ", "dip_mac": "dịp", "ma_don": "đơn"}
    return "; ".join(f"{lab[k]}: {f[k]}" for k in lab if f.get(k))

def new_state():
    return {"intent": "chitchat", "slots": {}, "history": [], "last_bot": "", "last_reply": "", "last_text": "",
            "focus": {}, "trace": []}      # focus = "bộ nhớ dài": món/đơn/dịp đang nói tới, KHÔNG bị xóa khi đổi chủ đề

KEEP_ALWAYS = {"chieu_cao", "can_nang"}           # số đo cơ thể dùng lại xuyên suốt phiên

def update_state(state, text):
    state["last_text"] = text
    state["slots"].pop("truong", None)
    t = strip_acc(text).lower()
    llm_intent, llm_slots = None, {}
    if USE_LLM:
        d = llm_dst(text, state)
        if d:
            llm_intent, llm_slots = d.get("intent"), d.get("slots") or {}
    new = {}
    ptxt = positive_part(text)                                    # chỉ lấy vế khẳng định để nhận loại đồ / dịp / sản phẩm
    cat = match_key(ptxt, CATS) or (llm_slots.get("loai_do") if llm_slots.get("loai_do") in CATEGORIES else None)
    if not cat and is_generic_vay(ptxt):
        cat = "Váy"                                               # khách nói "váy" chung -> tìm cả đầm và chân váy
    dip = match_key(ptxt, DIPS) or (llm_slots.get("dip_mac") if llm_slots.get("dip_mac") in OCCASIONS else None)
    if cat: new["loai_do"] = cat
    if dip: new["dip_mac"] = dip
    focus = state.setdefault("focus", {})
    prod = pick_item(text, state["slots"].get("_list") or focus.get("_list")) or match_product(ptxt, cat)
    if not prod:
        prod = resolve_reference(text, ptxt, cat, state)          # "áo này", "mẫu đó", "còn hàng không?"
    if prod:
        new["ten_sp"] = prod
        new["loai_do"] = next(p["loai"] for p in CATALOG if p["ten"] == prod)
    kw = None if prod else match_keywords(ptxt)                   # đã chỉ đích danh món thì không thêm từ khóa lấy từ chính tên món
    if kw: new["tu_khoa"] = kw
    new.update(regex_slots(text))                                 # regex ưu tiên hơn LLM cho số và mã
    if "ma_don" not in new and focus.get("ma_don") and re.search(FOLLOW_ORDER, t) \
            and (state["intent"] == "order_check" or re.search(r"\bdon\s+(?:nay|do|kia)\b", t)):
        new["ma_don"] = focus["ma_don"]                           # "thế khi nào giao tới?" -> vẫn là đơn vừa tra
    for k in ("chieu_cao", "can_nang"):
        if k not in new and isinstance(llm_slots.get(k), (int, float)):
            new[k] = int(llm_slots[k])
    if state["intent"] == "size_advice" and state["last_bot"]:    # khách trả lời trơn "52" hoặc "165 52"
        for n in [int(x) for x in re.findall(r"\b\d{2,3}\b", t)]:
            if 130 <= n <= 210 and "chieu_cao" not in new and "chieu_cao" not in state["slots"]:
                new["chieu_cao"] = n
            elif 30 <= n <= 120 and "can_nang" not in new and "can_nang" not in state["slots"]:
                new["can_nang"] = n
    if re.search(r"gio hang|voucher|ma giam|giam gia|uu dai|ap ma|ma nao|khuyen mai", t) or state["intent"] == "voucher" \
            or "ma_voucher" in new:
        cart = parse_money(text)
        if cart and "ma_don" not in new: new["cart_total"] = cart
    truong = field_asked(text)
    # ---- intent ----
    rule = rule_intent(text, new, state["slots"])
    intent = llm_intent if llm_intent in INTENTS else rule
    if intent == "chitchat" and rule != "chitchat":
        intent = rule
    if "ma_don" in new:
        intent = "order_check"
    if intent == "chitchat" and state["last_bot"] and state["intent"] in FOLLOW_KEYS \
            and any(k in new for k in FOLLOW_KEYS[state["intent"]]):
        intent = state["intent"]                                  # đang trả lời câu hỏi lại của bot
    if intent != state["intent"]:                                 # đổi chủ đề -> bỏ slot cũ không còn liên quan
        keep = KEEP_ALWAYS | ({"ten_sp", "loai_do"} if intent == "size_advice" else set())
        state["slots"] = {k: v for k, v in state["slots"].items() if k in keep}
    if "ten_sp" in new or "loai_do" in new:                       # khách đổi sang món/loại khác
        for k in ("ten_sp", "tu_khoa"):
            if k not in new: state["slots"].pop(k, None)
    state["intent"] = intent
    state["slots"].update(new)
    if truong and intent == "product_info": state["slots"]["truong"] = truong
    return state

print("DST sẵn sàng")


# ==================== Policy + NLG ====================
# ============ BƯỚC 3: POLICY & ACTION - thiếu slot thì hỏi lại, đủ slot thì GỌI API ============
ASK = {
    "chieu_cao": "Dạ nàng cho SUPEO xin thêm chiều cao (cm) nhé!",
    "can_nang": "Dạ nàng chia sẻ thêm cân nặng (kg) để shop tư vấn chuẩn xác nha!",
    "loai_do": "Dạ nàng đang quan tâm nhóm đồ nào ạ? SUPEO có: {cats}.",
    "dip_mac": "Dạ nàng định mặc đi dịp nào ạ? SUPEO có sẵn set cho: {dips}.",
    "ma_don": "Dạ nàng nhập giúp SUPEO mã đơn hàng đúng định dạng #SPxxxx (ví dụ #SP2908) nhé!",
}
HELP = ("Dạ SUPEO có thể giúp nàng: tra cứu sản phẩm, tư vấn size, gợi ý set đồ theo dịp, kiểm tra đơn hàng và tìm mã giảm giá. "
        "Nàng thử nhắn như: 'Có áo blazer linen màu be không?', 'Mình cao 1m62 nặng 52kg mặc size gì?', "
        "'Check giúp đơn #SP2908' hoặc 'Giỏ hàng 750k dùng mã nào?' nhé!")
# Chính sách shop KHÔNG nằm trong Excel. Dưới đây là câu TỰ VIẾT, không chép từ đáp án mẫu của dialogue_benchmarks
# (chép sẽ gây rò rỉ khi đo). Hãy THAY bằng thông tin thật của shop.
FAQ_SHIPPING = ("Dạ về hình thức và thời gian giao hàng, SUPEO chưa có thông tin chính xác trong hệ thống nên chưa dám báo nàng ạ. "
                "Nàng nhắn fanpage/hotline của shop để được báo cụ thể nhé! Nếu nàng đã đặt hàng, gửi SUPEO mã đơn #SPxxxx để SUPEO kiểm tra tình trạng giúp nàng.")
FAQ_THANKS = "Dạ cảm ơn nàng đã tin tưởng SUPEO ạ! Nàng cần hỗ trợ thêm gì cứ nhắn SUPEO nhé."

def ask_text(slot):
    return ASK[slot].format(cats=", ".join(CATEGORIES), dips=", ".join(OCCASIONS))

def norm(s): return strip_acc(s).lower().strip()

# ---------- Chức năng 1: product_info ----------
def lc(x):
    x = str(x)
    return x[:1].lower() + x[1:]

def stock_text(p):
    if p["tinh_trang"] == "Sắp hết hàng":
        return f"sắp hết hàng (chỉ còn {p['ton_kho']} sản phẩm)"
    if p["tinh_trang"] == "Hết hàng":
        return "tạm hết hàng"
    return "đang còn hàng"

def product_card(p, truong=None):
    mau, size = ", ".join(p["mau"]), ", ".join(p["size"])
    ten = f"{p['ten']} ({p['id']})"
    opener = {"size": f"Dạ {ten} hiện có size {size} nè nàng.",
              "chat_lieu": f"Dạ {ten} được may bằng {lc(p['chat_lieu'])} ạ.",
              "mau": f"Dạ {ten} có các màu {mau} nàng nhé.",
              "gia": f"Dạ {ten} có giá niêm yết {vnd(p['gia'])} ạ.",
              "ton_kho": f"Dạ {ten} hiện {stock_text(p)} ạ.",
              "anh": f"Dạ SUPEO chưa gửi được ảnh trong khung chat, nàng tham khảo thông tin {ten} bên dưới nhé."}.get(truong, f"Dạ SUPEO gửi nàng thông tin {ten} nè.")
    parts = {"gia": f"giá niêm yết {vnd(p['gia'])}", "chat_lieu": f"chất liệu {lc(p['chat_lieu'])}",
             "phom": f"phom {str(p['phom_dang']).lower()}", "mau": f"màu {mau}", "size": f"size {size}",
             "ton_kho": f"hiện {stock_text(p)}"}
    detail = ", ".join(v for k, v in parts.items() if k != truong)       # không lặp lại điều đã nói ở câu mở đầu
    detail = detail[0].upper() + detail[1:] + "."
    mota = str(p["mo_ta"]).strip().rstrip(".") + "."
    return (f"{opener}\n{detail}\n{mota}\n"
            "Nàng cho SUPEO xin chiều cao, cân nặng để tư vấn size, hoặc nói dịp mặc để SUPEO gợi ý set phối đồ nhé!")

def act_product(S):
    if S.get("ten_sp"):
        _, d = api_get("/v1/products", q=S["ten_sp"])
        items = [x for x in d["items"] if norm(x["ten"]) == norm(S["ten_sp"])] or d["items"]
    elif S.get("loai_do") == "Váy":                                  # "váy" chung = Đầm + Chân váy Midi
        items = []
        for c in ("Đầm", "Chân váy Midi"):
            items += api_get("/v1/products", loai=c, q=S.get("tu_khoa"))[1]["items"]
    else:
        _, d = api_get("/v1/products", loai=S.get("loai_do"), q=S.get("tu_khoa"))
        items = d["items"]
    if not items:
        return "inform", ("Dạ hiện SUPEO chưa có sản phẩm nào khớp với yêu cầu này ạ. "
                          f"Nàng xem thử các nhóm: {', '.join(CATEGORIES)} hoặc đổi màu/chất liệu khác, SUPEO tìm lại ngay nhé!")
    if len(items) == 1:
        p = items[0]
        S["ten_sp"], S["loai_do"] = p["ten"], p["loai"]               # nhớ món vừa gửi để hiểu "áo này", "còn hàng không?"
        card = product_card(p, S.get("truong"))
        h, w = S.get("chieu_cao"), S.get("can_nang")
        if h and w and 140 <= h <= 200 and 30 <= w <= 120:            # đã biết số đo của khách -> dùng luôn, không hỏi lại
            _, r = api_get("/v1/size-advice", chieu_cao=h, can_nang=w)
            sz = r["de_xuat"]["size"]
            tail = (f"mẫu này có size {sz} nè nàng!" if sz in p["size"]
                    else f"nhưng mẫu này chỉ có size {', '.join(p['size'])}, nàng cân nhắc giúp SUPEO nhé.")
            card = card.rsplit("\n", 1)[0] + f"\nTheo số đo nàng đã chia sẻ ({h}cm, {w}kg), nàng hợp Size {sz}, " + tail
        return "inform", card
    label = ""
    if S.get("loai_do"):
        label = " váy (đầm và chân váy)" if S["loai_do"] == "Váy" else f" {S['loai_do']}"
    if S.get("tu_khoa"):
        label += f" ({S['tu_khoa']})"
    shown = items[:10]
    S["_list"] = [p["ten"] for p in shown]                            # nhớ danh sách để khách chọn 'mẫu thứ 2'
    head = (f"Dạ SUPEO chưa gửi được ảnh trong khung chat, nàng tham khảo các mẫu{label} dưới đây nhé:" if S.get("truong") == "anh"
            else f"Dạ SUPEO đang có {len(items)} mẫu{label} nàng tham khảo nè:")
    lines = [f"{i}. {p['ten']} ({p['id']}) - {vnd(p['gia'])}; màu {', '.join(p['mau'])}; {p['tinh_trang'].lower()}"
             for i, p in enumerate(shown, 1)]
    more = f"\n(còn {len(items) - 10} mẫu khác nữa ạ)" if len(items) > 10 else ""
    return "inform", (head + "\n" + "\n".join(lines) + more +
                      "\nNàng thích mẫu nào cứ nhắn tên hoặc số thứ tự (ví dụ 'mẫu thứ 2'), SUPEO gửi chi tiết ngay nhé! "
                      "Nàng cũng có thể cho SUPEO biết chiều cao, cân nặng để chọn đúng size.")

# ---------- Chức năng 2: size_advice (Max-Boundary nằm ở API) ----------
def act_size(S):
    h, w = S["chieu_cao"], S["can_nang"]
    if not (140 <= h <= 200 and 30 <= w <= 120):
        S.pop("chieu_cao", None); S.pop("can_nang", None)
        return "ask", "Dạ chiều cao và cân nặng nàng nhập chưa hợp lý, nàng gửi lại giúp SUPEO (ví dụ: 1m65, 52kg) nhé!"
    _, r = api_get("/v1/size-advice", chieu_cao=h, can_nang=w)
    de = r["de_xuat"]; kg = r["theo_can_nang"]["khoang"]; cm = r["theo_chieu_cao"]["khoang"]
    if r["lech"]:
        txt = (f"Dạ với chiều cao {h}cm nàng đang ở size {r['theo_chieu_cao']['size']} ({cm[0]}-{cm[1]}cm), "
               f"còn cân nặng {w}kg lại thuộc size {r['theo_can_nang']['size']} ({kg[0]}-{kg[1]}kg). "
               f"Khi hai số đo lệch nhau, SUPEO ưu tiên size lớn hơn để nàng mặc thoải mái, nên SUPEO khuyên nàng chọn Size {de['size']} nha!")
    else:
        txt = (f"Dạ với chiều cao {h}cm và cân nặng {w}kg, nàng mặc chuẩn Size {de['size']} "
               f"({de['chieu_cao'][0]}-{de['chieu_cao'][1]}cm, {de['can_nang'][0]}-{de['can_nang'][1]}kg) là vừa vặn nhất ạ!")
    txt += f"\n({de['loi_khuyen']})"
    if r["ngoai_dai"]["chieu_cao"] or r["ngoai_dai"]["can_nang"]:
        txt += "\nSố đo của nàng nằm ngoài dải size chuẩn, nàng nhắn SUPEO để được tư vấn kỹ hơn nhé."
    if S.get("ten_sp"):                                           # kiểm tra mẫu đó có size này không
        _, d = api_get("/v1/products", q=S["ten_sp"])
        p = next((x for x in d["items"] if norm(x["ten"]) == norm(S["ten_sp"])), None)
        if p and de["size"] not in p["size"]:
            txt += f"\nLưu ý: {p['ten']} ({p['id']}) hiện chỉ có size {', '.join(p['size'])}, nàng cân nhắc giúp SUPEO nhé."
        elif p:
            txt += f"\n{p['ten']} ({p['id']}) có size {de['size']} và {stock_text(p)} nè nàng."
    elif S.get("loai_do") and S["loai_do"] != "Váy":              # gợi ý mẫu còn size phù hợp trong nhóm khách đang quan tâm
        _, d = api_get("/v1/products", loai=S["loai_do"])
        ok = [x for x in d["items"] if de["size"] in x["size"] and x["ton_kho"] > 0][:3]
        if ok:
            txt += (f"\nCác mẫu {S['loai_do']} còn size {de['size']} nàng có thể thử: "
                    + "; ".join(f"{x['ten']} ({x['id']}) {vnd(x['gia'])}" for x in ok) + ".")
    else:
        txt += "\nNàng nói tên mẫu hoặc nhóm đồ đang quan tâm, SUPEO kiểm tra giúp mẫu đó còn size này không nhé!"
    return "inform", txt

# ---------- Chức năng 3: lookbook ----------
def act_lookbook(S):
    st, d = api_get(f"/v1/lookbook/{quote(S['dip_mac'], safe='')}")
    if st == 404:
        S.pop("dip_mac", None)
        return "ask", ask_text("dip_mac")
    lines = []
    for p in d["items"]:
        note = "" if p["tinh_trang"] == "Còn hàng" else f" ({p['tinh_trang'].lower()})"
        lines.append(f"- {p['ten']} ({p['id']}): {vnd(p['gia'])}{note}")
    return "inform", (f"Dạ cho dịp {d['dip_mac']}, SUPEO gợi ý nàng set '{d['ten_bo']}' phong cách {d['phong_cach']} nè:\n"
                      + "\n".join(lines) + f"\nTổng cả set: {vnd(d['tong_gia_tien'])}.\n{d['ghi_chu']}\n"
                      "Nàng cho SUPEO xin chiều cao, cân nặng để chọn size từng món, hoặc nhắn tên món nàng thích để SUPEO gửi chi tiết nhé!")

# ---------- Chức năng 4: order_check ----------
def act_order(S):
    st, d = api_get(f"/v1/orders/{quote(S['ma_don'], safe='')}/tracking")
    if st == 422:
        S.pop("ma_don", None)
        return "ask", ask_text("ma_don")
    if st == 404:
        code = S["ma_don"]
        return "inform", (f"Dạ SUPEO đã kiểm tra nhưng chưa tìm thấy mã đơn {code} trên hệ thống. Nàng vui lòng kiểm tra lại "
                          "tin nhắn xác nhận đơn hoặc gửi số điện thoại đặt hàng để shop hỗ trợ tra cứu nha!")
    sp = "; ".join(d["san_pham"])
    tt = d["trang_thai"]
    txt = f"Dạ SUPEO đã kiểm tra đơn hàng {d['ma_don']} ({sp}) của nàng rồi nè. Đơn đặt ngày {d['ngay_dat']} và hiện đang ở trạng thái '{tt}'."
    if tt == "Đã hủy":
        txt = f"Dạ đơn hàng {d['ma_don']} ({sp}) đặt ngày {d['ngay_dat']} đã ở trạng thái 'Đã hủy'. Nếu cần đặt lại, nàng nhắn SUPEO hỗ trợ tạo đơn mới nha!"
    elif d["da_ban_giao_vc"]:
        txt += f" Đơn qua {d['dvvc']} (mã vận đơn: {d['ma_van_don']})."
    elif d["dvvc"] and d["dvvc"] != "Chưa bàn giao":
        txt += f" Đơn sẽ giao qua {d['dvvc']}, mã vận đơn sẽ có khi bàn giao."
    else:
        txt += " Đơn chưa bàn giao cho đơn vị vận chuyển."
    if tt == "Đã giao thành công":
        txt += " Chúc nàng mặc đồ SUPEO thật xinh nhé!"
    elif tt.startswith("Giao thất bại"):
        txt += f" Đơn đang được hoàn về shop, nàng nhắn SUPEO để được hỗ trợ giao lại nhé (dự kiến cũ: {d['du_kien']})."
    elif d["du_kien"]:
        txt += f" Dự kiến giao/bàn giao đến nàng ngày {d['du_kien']}."
    return "inform", txt + f" Tổng giá trị đơn: {vnd(d['tong_tien'])}. Nàng cần SUPEO hỗ trợ thêm gì cho đơn này không ạ?"

# ---------- Chức năng 5: voucher ----------
def voucher_line(v, cart=None):
    s = f"- {v['ma']}: {v['mo_ta']}"
    if v["doi_tuong"] and v["doi_tuong"] != "Mọi khách hàng":
        s += f" (Đối tượng: {v['doi_tuong']})"
    s += f" - HSD {v['han_dung']}"
    if v["con_lai_ngay"] is not None and v["con_lai_ngay"] <= 7:
        s += f" (sắp hết hạn, còn {v['con_lai_ngay']} ngày)"
    if cart is not None and v.get("uoc_tinh_giam"):
        s += f" => ước tính giảm {vnd(v['uoc_tinh_giam'])}"
    return s

def act_voucher(S):
    cart = S.get("cart_total")
    if S.get("ma_voucher"):
        _, d = api_get(f"/v1/vouchers/{quote(S['ma_voucher'], safe='')}", cart_total=cart)
        tt = d["trang_thai"]
        if tt == "khong_kha_dung":
            return "inform", (f"Dạ SUPEO chưa thấy mã {d['ma']} trong danh sách ưu đãi đang mở. Nàng kiểm tra lại mã, "
                              "hoặc hỏi SUPEO 'có mã giảm giá nào' để xem các mã đang áp dụng nhé!")
        if tt == "het_han":
            return "inform", f"Dạ mã {d['ma']} đã hết hạn từ {d['han_dung']} ạ. Nàng hỏi SUPEO 'có mã giảm giá nào' để xem mã còn hạn nhé!"
        if tt == "chua_du_dieu_kien":
            return "inform", (f"Dạ mã {d['ma']} ({d['mo_ta']}) cần đơn từ {vnd(d['don_toi_thieu'])}. Giỏ hàng nàng đang {vnd(cart)}, "
                              f"nàng thêm {vnd(d['con_thieu'])} nữa là dùng được ạ!")
        txt = f"Dạ mã {d['ma']} dùng được nha: {d['mo_ta']}. HSD {d['han_dung']}."
        if cart is not None: txt += f" Với giỏ {vnd(cart)}, ước tính giảm {vnd(d['uoc_tinh_giam'])}."
        return "inform", txt + f" Lưu ý: {d['dieu_khoan']}"
    _, d = api_get("/v1/vouchers/available", cart_total=cart)
    if not d["vouchers"]:
        return "inform", "Dạ hiện SUPEO chưa có mã ưu đãi nào đang mở, nàng theo dõi shop để cập nhật sớm nhất nhé!"
    if cart is None:
        return "inform", (f"Dạ hiện SUPEO đang có {d['count']} mã ưu đãi:\n" + "\n".join(voucher_line(v) for v in d["vouchers"])
                          + "\nNàng cho SUPEO biết giá trị giỏ hàng để chọn mã tốt nhất nhé!")
    ok = [v for v in d["vouchers"] if v["du_dieu_kien"]]
    if not ok:
        up = d["goi_y_upsell"]
        return "inform", (f"Dạ với giỏ hàng {vnd(cart)} nàng chưa đủ điều kiện mã nào. " +
                          (f"Nàng thêm {vnd(up['con_thieu'])} nữa để dùng {up['ma']} ({up['mo_ta']}) nhé!" if up else ""))
    best = next(v for v in ok if v["ma"] == d["tot_nhat"])
    txt = (f"Dạ với giỏ hàng {vnd(cart)}, nàng áp {best['ma']} là lợi nhất, giảm khoảng {vnd(best['uoc_tinh_giam'])}.\n"
           "Các mã nàng đủ điều kiện:\n" + "\n".join(voucher_line(v, cart) for v in ok))
    up = d["goi_y_upsell"]
    if up:
        txt += (f"\nGợi ý: nàng chọn thêm khoảng {vnd(up['con_thieu'])} để đơn đạt {vnd(up['don_toi_thieu'])} "
                f"và dùng thêm {up['ma']} ({up['mo_ta']}).")
    return "inform", txt

# ---------- Chitchat / FAQ ----------
def act_chitchat(text):
    t = strip_acc(text).lower()
    if re.search(r"cam on|thanks|thank", t): return "chitchat", FAQ_THANKS
    if re.search(r"tam biet|bye|hen gap lai|thoi nhe", t):
        return "chitchat", "Dạ cảm ơn nàng đã ghé SUPEO, chúc nàng một ngày thật xinh và hẹn gặp lại nàng nhé!"
    if re.search(r"hoa toc|giao nhanh|phi ship|phi van chuyen|giao hang|bao lau|thoi gian giao|giao toan quoc", t): return "chitchat", FAQ_SHIPPING
    if re.search(r"doi tra|bao hanh|thanh toan|\bcod\b|hotline|so dien thoai|dia chi|cua hang o dau|ban ship cod", t):
        return "chitchat", ("Dạ phần này SUPEO chưa có thông tin chính xác trong hệ thống nên không dám trả lời sai với nàng. "
                            "Nàng nhắn trực tiếp cho shop (fanpage/hotline) để được hỗ trợ nhanh nhất nhé! "
                            "Còn về sản phẩm, size, set đồ, đơn hàng hay mã giảm giá thì nàng cứ hỏi SUPEO nha.")
    if re.search(r"\b(?:xin chao|chao|hello|hi|alo|shop oi|ban oi)\b", t):
        return "chitchat", "Dạ SUPEO chào nàng ạ! " + HELP[3:]
    return "chitchat", HELP

def policy(state):
    intent, S = state["intent"], state["slots"]
    need = missing(intent, S)
    if intent == "size_advice" and len(need) == 2:
        return "ask", "Dạ nàng cho SUPEO xin chiều cao (cm) và cân nặng (kg) để tư vấn size chuẩn nhất nhé!"
    if need == ["loai_do"] and S.get("truong") == "anh":
        return "ask", "Dạ SUPEO chưa gửi được ảnh trong khung chat ạ. Nàng nói tên mẫu hoặc nhóm đồ (" + ", ".join(CATEGORIES) + ") để SUPEO gửi thông tin chi tiết nhé!"
    if need:
        return "ask", ask_text(need[0])
    if intent == "product_info": return act_product(S)
    if intent == "size_advice": return act_size(S)
    if intent == "lookbook": return act_lookbook(S)
    if intent == "order_check": return act_order(S)
    if intent == "voucher": return act_voucher(S)
    return act_chitchat(state["last_text"])

# ============ NLG: diễn đạt tự nhiên, nhưng KHÔNG được đổi số liệu ============
NLG_SYS = ("Bạn là tư vấn viên thời trang SUPEO, xưng 'SUPEO', gọi khách là 'nàng', lịch sự, thân thiện. "
           "Có thể xem 'Hội thoại gần đây' để câu văn liền mạch, nhưng TUYỆT ĐỐI không lấy số liệu từ đó. "
           "Viết lại thông tin sau cho tự nhiên, ngắn gọn. GIỮ NGUYÊN mọi con số, giá tiền, mã, ngày tháng, tên sản phẩm. "
           "TUYỆT ĐỐI không thêm thông tin không có trong dữ liệu. Giữ nguyên mã sản phẩm dạng (P01) và số thứ tự trong danh sách. Giữ định dạng gạch đầu dòng nếu có.")

def facts_ok(facts, out):
    toks = re.findall(r"[#\w.]*\d[\w.]*", facts)
    return all(tk.rstrip(".") in out for tk in toks)

def nlg(facts, user_text, state=None):
    if not USE_LLM:
        return facts
    try:
        out = llm([{"role": "system", "content": NLG_SYS},
                   {"role": "user", "content": (f"Hội thoại gần đây:\n{recent_history(state)}\n\n" if state and state.get("history") else "")
                    + f"Câu khách: {user_text}\nDữ liệu thật từ hệ thống:\n{facts}"}], 350).strip()
        return out if out and facts_ok(facts, out) else facts      # chốt chặn chống bịa số liệu
    except Exception:
        return facts

FOCUS_KEYS = ("ten_sp", "loai_do", "ma_don", "dip_mac", "_list", "cart_total")

def update_focus(state):
    """Ghi nhớ dài hạn: món / đơn / dịp / danh sách đang nói tới. Khác với slots, focus không bị xóa khi đổi chủ đề."""
    S, f = state["slots"], state.setdefault("focus", {})
    if state["intent"] == "product_info" and "ten_sp" not in S:
        f.pop("ten_sp", None)                                    # đang xem danh sách nhiều mẫu -> chưa có món cụ thể
    for k in FOCUS_KEYS:
        if S.get(k) is not None:
            f[k] = S[k]

def bot_reply(text, state):
    API_LOG.clear()
    try:
        update_state(state, text)
        kind, out = policy(state)
    except ApiDown:
        kind, out = "error", "Dạ hệ thống tra cứu của SUPEO đang bận, nàng thử lại sau ít phút nhé!"
    reply = nlg(out, text, state) if kind == "inform" else out
    state["history"] = (state["history"] + [("user", text), ("bot", reply)])[-40:]     # giữ 20 lượt gần nhất
    state["last_bot"] = reply if kind == "ask" else ""
    state["last_reply"] = reply
    update_focus(state)
    state["trace"].append({"user": text, "intent": state["intent"], "slots": dict(state["slots"]),
                           "action": kind, "api_calls": list(API_LOG)})
    return reply

print("Policy & NLG sẵn sàng")


# ==================== Bản vá v5 ====================
# ============ BẢN VÁ v5: nhận món / màu / size / voucher / nhớ đơn / che tên món ============
# Đặt cell này NGAY SAU cell "Policy + NLG" (cell 13) và TRƯỚC cell bật LLM. Cell chỉ ĐỊNH NGHĨA LẠI hàm
# (không dùng thay chuỗi nên không thể "không khớp"). Không có luật riêng cho câu test nào.
import re, math, unicodedata

def _nfc(s):
    return unicodedata.normalize("NFC", str(s)).lower()

def _has_acc(s):
    return strip_acc(s) != s

# ---------------------------------------------------------------- 1) Chỉ mục nhận món (idf) ----
GENERIC_NAME_TOK = {"mua", "he", "thu", "nhe", "cap", "dep", "dang"}        # từ trong tên món nhưng không phân biệt
COMMON_VI = {"tien", "bang", "no", "nhieu", "gia", "vua", "xinh", "hop", "mac", "lam", "ban", "cho", "hoi", "ve", "di",
             "mua", "do", "la", "nay", "kia", "nao", "gi", "khong", "co", "con", "het"}   # từ hay gặp khi gõ KHÔNG dấu

NAME_TOK, NAME_TOK_ACC, NON_ID = {}, {}, set()

def build_match_index():
    """Gọi lại sau load_lexicons(). NON_ID = từ chung/loại/màu/dịp/chất liệu: không dùng để nhận diện món."""
    global NAME_TOK, NAME_TOK_ACC, NON_ID
    toks = lambda s: set(re.findall(r"[a-z0-9]+", strip_acc(s).lower()))
    color_tok = set().union(*[toks(k) for k in COLOR_LEX]) if COLOR_LEX else set()
    dip_tok = set().union(*[toks(p) for pats in DIPS.values() for p in pats])
    mat_tok = set().union(*[toks(k) for k in MATERIAL_LEX])
    NON_ID = set(CAT_WORDS) | set(STOP) | color_tok | dip_tok | mat_tok | GENERIC_NAME_TOK
    NAME_TOK = {p["id"]: {t for t in toks(p["ten"]) - NON_ID if len(t) > 1} for p in CATALOG}
    NAME_TOK_ACC = {p["id"]: {t for t in re.findall(r"\w+", _nfc(p["ten"]))
                              if len(t) > 1 and strip_acc(t) not in NON_ID} for p in CATALOG}

def idf_pick(words, group, tokmap, min_score=1.0, margin=1.0, banned=frozenset()):
    """Chọn món theo độ HIẾM của từ khóa trong nhóm: từ chỉ có ở 1 món nặng hơn từ có ở nhiều món.
    Chỉ trả món khi điểm >= min_score VÀ hơn món thứ hai ít nhất `margin`; không đủ chắc -> None (để trả danh sách)."""
    n = len(group)
    if n < 2:
        return None
    df = {}
    for p in group:
        for t in tokmap[p["id"]]:
            df[t] = df.get(t, 0) + 1
    scored = []
    for p in group:
        hits = {t for t in tokmap[p["id"]] & words if t not in banned}
        if not hits:
            continue
        if len(hits) == 1 and len(next(iter(hits))) < 3:                  # 1 từ 2 ký tự ('nơ', 'ly'...) chưa đủ chắc
            continue
        scored.append((sum(math.log(n / df[t]) for t in hits), p["ten"]))
    scored.sort(reverse=True)
    if not scored or scored[0][0] < min_score:
        return None
    if len(scored) > 1 and scored[0][0] - scored[1][0] < margin:
        return None
    return scored[0][1]

def mask_product_names(text):
    """Che TÊN MÓN ĐẦY ĐỦ khỏi câu (trả về chuỗi không dấu) để chữ trong tên ('công sở', 'dự tiệc')
    không bị hiểu là dịp mặc / loại đồ khi khách chỉ nhắc tên món."""
    t = strip_acc(text).lower()
    for p in sorted(CATALOG, key=lambda x: -len(x["ten"])):
        t = t.replace(strip_acc(p["ten"]).lower(), " sanpham ")
    return t

def match_product(text, cat):
    """Tên đầy đủ -> chọn; không có thì chọn theo idf trong cùng nhóm (nếu biết loại) hoặc toàn danh mục."""
    t = strip_acc(text).lower()
    best = None
    for p in CATALOG:
        if strip_acc(p["ten"]).lower() in t and (best is None or len(p["ten"]) > len(best["ten"])):
            best = p
    if best:
        return best["ten"]
    if cat == "Váy":
        group = [p for p in CATALOG if p["loai"] in ("Đầm", "Chân váy Midi")]
    elif cat in CATEGORIES:
        group = [p for p in CATALOG if p["loai"] == cat]
    else:
        group = CATALOG
    nfc = _nfc(text)
    if _has_acc(nfc):                                   # gõ có dấu -> so sánh CÓ DẤU ('tiên' khác 'tiền', 'nơ' khác 'nó')
        return idf_pick(set(re.findall(r"\w+", nfc)), group, NAME_TOK_ACC)
    return idf_pick(set(re.findall(r"[a-z0-9]+", t)), group, NAME_TOK, banned=COMMON_VI)

# ---------------------------------------------------------------- 2) Nhận màu không cần chữ "màu" ----
SHORT_AMBIG = {"be", "do"}                                                  # 'be' dễ nhầm -> chỉ nhận khi đứng ngay sau tên loại đồ
NOUN_BEFORE = r"(?:blazer|vest|ao|dam|vay|quan|so mi|midi|culottes)"

def match_color(text):
    nfc = _nfc(text)
    t = strip_acc(nfc)
    acc = t != nfc
    for k, vi in COLOR_LEX.items():                                         # COLOR_LEX đã sắp theo độ dài giảm dần
        if re.search(r"\bmau\s+" + re.escape(k) + r"\b", t):                # có chữ "màu"
            return vi
    for k, vi in COLOR_LEX.items():
        if acc and k not in SHORT_AMBIG:
            if re.search(r"(?<!\w)" + re.escape(vi) + r"(?!\w)", nfc):      # có dấu: 'đen' khác 'đến'
                return vi
        elif not acc and (len(k) >= 6 or " " in k) and k not in SHORT_AMBIG:
            if re.search(r"\b" + re.escape(k) + r"\b", t):                  # không dấu: chỉ nhận màu dài ('champagne')
                return vi
        if re.search(r"\b" + NOUN_BEFORE + r"\s+" + re.escape(k) + r"\b", t):   # 'đầm be', 'blazer kem'
            return vi
    return None

def match_keywords(text):
    t = strip_acc(text).lower()
    kws = [vi for k, vi in MATERIAL_LEX.items() if re.search(r"\b" + k + r"\b", t)]
    c = match_color(text)
    if c:
        kws.append(c)
    return " ".join(kws) or None

# ---------------------------------------------------------------- 3) Dịp mặc (không nhận từ trong tên món) ----
def detect_dip(ptxt, prod_from_text=None):
    t = mask_product_names(ptxt)
    if prod_from_text:                                  # khách gõ một phần tên ('suông công sở'): bỏ cụm dịp nằm trong tên món
        pn = strip_acc(prod_from_text).lower()
        for pats in DIPS.values():
            for p in pats:
                if re.search(r"\b" + p + r"\b", pn) and len(re.findall(r"\b" + p + r"\b", t)) == 1:
                    t = re.sub(r"\b" + p + r"\b", " ", t)
    return match_key(t, DIPS)

# ---------------------------------------------------------------- 4) Size: hỏi hàng vs xin tư vấn ----
SIZE_CODE = r"(?:xxl|xl|xs|s|m|l)"
PERSONAL_SIZE = (r"\b(?:minh|em|toi|chi|con)\s+(?:mac|lay|chon|dung)\b|\bnen\s+(?:mac|lay|chon)\b|\bmac size\b"
                 r"|\bsize cua (?:minh|em|toi)\b|\bhop size\b|\bsize nao\s+(?:hop|vua|phu hop|dung|nen)\b"
                 r"|\b(?:hop|vua|phu hop)\s+voi\s+(?:minh|em|toi|chi)\b|\bsize\s+(?:gi|nao)\s+(?:cho|voi)\s+(?:minh|em|toi|chi)\b")

def asked_size_code(text):
    """'còn size M không' -> 'M' (hỏi hàng). 'mặc size M được không' / 'size M có vừa không' là xin tư vấn -> None."""
    t = strip_acc(text).lower()
    m = re.search(r"\bsize\s+(" + SIZE_CODE + r")\b", t)
    if not m or re.search(PERSONAL_SIZE, t) or re.search(r"\b(?:vua|hop|mac|duoc|thoai mai|nen)\b", t):
        return None
    return m.group(1).upper() if re.search(r"\b(?:con|het|co|ban|du)\b", t) else None

def field_asked(text):
    t = strip_acc(text).lower()
    if asked_size_code(text):
        return "size"
    if re.search(r"\bsize\b", t) and re.search(r"(nhung|may|co) size|size nao|size gi co|du size", t):
        return "size"
    if re.search(r"chat lieu|chat vai|vai gi|vai nao|lam bang", t):
        return "chat_lieu"
    nfc = _nfc(text)
    if re.search(r"\bmàu\b", nfc) if strip_acc(nfc) != nfc else re.search(r"\bmau\b", t):
        return "mau"
    if re.search(r"con hang|het hang|ton kho|con khong|con mau", t):
        return "ton_kho"
    if re.search(r"\bgia\b|bao nhieu|bao nhieu tien|nhieu tien", t):
        return "gia"
    if re.search(r"hinh anh|xem hinh|\b(?:xem|gui|coi|thay)\s+(?:duoc\s+)?(?:cai\s+|nhung\s+)?anh\b|\banh\s+(?:san pham|mau|that|thuc te)\b", t):
        return "anh"
    return None

# ---------------------------------------------------------------- 5) Voucher + intent ----
VOUCHER_RE = (r"\b(?:voucher|ma giam|giam gia|khuyen mai|uu dai|gio hang|ap ma|ma nao|ma gi|coupon|freeship|"
              r"giam bao nhieu|duoc giam|giam duoc|dung ma|nhap ma)\b")

def rule_intent(text, new, state_slots):
    t = strip_acc(text).lower()
    if "ma_don" in new or re.search(r"don hang|van don|check don|kiem tra don|tra cuu don|theo doi don|giao toi dau|ship toi dau|don cua minh", t):
        return "order_check"
    if "ma_voucher" in new or "cart_total" in new or re.search(VOUCHER_RE, t):
        return "voucher"
    asks_size_list = (field_asked(text) == "size" and "chieu_cao" not in new and "can_nang" not in new
                      and not re.search(PERSONAL_SIZE, t))
    if not asks_size_list and ("chieu_cao" in new or "can_nang" in new or re.search(r"\bsize\b|mac vua|co nguoi|nang bao nhieu|cao bao nhieu", t)):
        return "size_advice"
    if "dip_mac" in new or re.search(r"phoi|lookbook|goi y (?:bo|set|do|outfit)|set do|outfit", t):
        return "lookbook"
    if any(k in new for k in ("loai_do", "ten_sp", "tu_khoa")) or field_asked(text) or re.search(r"mau (?:nao|gi) |san pham|mat hang", t):
        return "product_info"
    return "chitchat"

# ---------------------------------------------------------------- 6) Hỏi tiếp về đơn / tham chiếu ----
FOLLOW_ORDER = (r"\b(?:giao|khi nao|bao gio|bao lau|toi noi|den noi|toi chua|den chua|toi dau|den dau|dang o dau|nhan duoc|"
                r"van chuyen|van don|trang thai|tinh trang|du kien|hoan|huy|ship|don nay|don do|don kia|don day)\b")
DEICTIC_ACC = re.compile(r"\b(?:áo|đầm|váy|quần|blazer|mẫu|sản phẩm|món|cái|set|bộ)\s+(?:này|đó|đấy|kia|ấy|vừa rồi|trên)\b|\bnó\b")
DEICTIC_NOACC = re.compile(r"\b(?:ao|dam|vay|quan|blazer|san pham|mon|cai)\s+(?:nay|kia|ay|do|day)\b|\bmau\s+(?:nay|kia|do|day)\b|\bcai do\b")


# ---------------------------------------------------------------- 6b) resolve_reference: so loại với CHÍNH món đang nhớ ----
def resolve_reference(text, ptxt, cat, state):
    """'áo này', 'nó', 'mẫu đó', 'còn hàng không?' -> món đang nói tới. Loại đồ của món lấy từ danh mục (không từ focus['loai_do'],
    vì giá trị đó có thể bị LLM sai ghi đè)."""
    f = state.get("focus", {})
    if not f.get("ten_sp"):
        return None
    p = next((x for x in CATALOG if x["ten"] == f["ten_sp"]), None)
    nfc = _nfc(text)
    na = strip_acc(nfc)
    if DEICTIC_ACC.search(nfc) or DEICTIC_NOACC.search(na):
        if cat and cat != "Váy" and p and cat != p["loai"]:
            return None                                      # "Áo này" nhưng món đang nói là đầm -> không ép
        return f["ten_sp"]
    if field_asked(text) and not cat and not state["slots"].get("ten_sp") \
            and not re.search(r"voucher|ma giam|gio hang|don hang|ma don|\bsp\d", na):
        return f["ten_sp"]
    return None


# ---------------------------------------------------------------- 6c) "70 cân" = 70 kg (cách nói thường ngày) ----
try:
    _regex_slots_v4
except NameError:
    _regex_slots_v4 = regex_slots            # chỉ bọc 1 lần dù chạy lại cell

def regex_slots(text):
    out = _regex_slots_v4(text)
    if "can_nang" not in out:
        m = re.search(r"\b(\d{2,3})\s?can\b", strip_acc(text).lower())     # số đứng TRƯỚC 'cân'; 'cần 70' không khớp
        if m and 30 <= int(m.group(1)) <= 120:
            out["can_nang"] = int(m.group(1))
    return out

# ---------------------------------------------------------------- 7) update_state (bản đầy đủ, thay bản cũ) ----
def update_state(state, text):
    state["last_text"] = text
    state["slots"].pop("truong", None)
    state["slots"].pop("_size", None)
    t = strip_acc(text).lower()
    llm_intent, llm_slots = None, {}
    if USE_LLM:
        d = llm_dst(text, state)
        if d:
            llm_intent, llm_slots = d.get("intent"), d.get("slots") or {}
    new = {}
    ptxt = positive_part(text)
    cat_rule = match_key(ptxt, CATS) or ("Váy" if is_generic_vay(ptxt) else None)      # loại đồ do LUẬT nhận từ chính câu khách
    cat = cat_rule or (llm_slots.get("loai_do") if llm_slots.get("loai_do") in CATEGORIES else None)
    focus = state.setdefault("focus", {})
    if cat and not cat_rule and field_asked(text) and (state["slots"].get("ten_sp") or focus.get("ten_sp")):
        cat = None                                                # câu hỏi tiếp về món đang nói: không nghe loại đồ do LLM bịa ra
    prod = pick_item(text, state["slots"].get("_list") or focus.get("_list"))
    prod_text = prod or match_product(ptxt, cat)                  # món nhận từ chính câu khách
    prod = prod_text or resolve_reference(text, ptxt, cat_rule, state)   # "áo này", "nó", "mẫu đó", "còn hàng không?" (LLM sai loại đồ cũng không làm mất ngữ cảnh)
    if prod and not cat_rule:
        cat = None                                                # loại đồ lấy từ món đã xác định, không lấy từ LLM
    dip = detect_dip(ptxt, prod_text) or (llm_slots.get("dip_mac") if llm_slots.get("dip_mac") in OCCASIONS else None)
    if cat: new["loai_do"] = cat
    if dip: new["dip_mac"] = dip
    if prod:
        new["ten_sp"] = prod
        new["loai_do"] = next(p["loai"] for p in CATALOG if p["ten"] == prod)
    kw = None if prod else match_keywords(ptxt)
    if kw: new["tu_khoa"] = kw
    new.update(regex_slots(text))
    if "ma_don" not in new and focus.get("ma_don") and re.search(FOLLOW_ORDER, t) \
            and (state["intent"] == "order_check" or re.search(r"\bdon\s+(?:nay|do|kia|day)\b", t)):
        new["ma_don"] = focus["ma_don"]                           # "bao lâu nữa tới?", "đơn đó giao chưa?" -> vẫn là đơn vừa tra
    for k in ("chieu_cao", "can_nang"):
        if k not in new and isinstance(llm_slots.get(k), (int, float)):
            new[k] = int(llm_slots[k])
    if state["intent"] == "size_advice" and state["last_bot"]:
        for n in [int(x) for x in re.findall(r"\b\d{2,3}\b", t)]:
            if 130 <= n <= 210 and "chieu_cao" not in new and "chieu_cao" not in state["slots"]:
                new["chieu_cao"] = n
            elif 30 <= n <= 120 and "can_nang" not in new and "can_nang" not in state["slots"]:
                new["can_nang"] = n
    if re.search(VOUCHER_RE, t) or state["intent"] == "voucher" or "ma_voucher" in new:
        cart = parse_money(text)
        if cart and "ma_don" not in new: new["cart_total"] = cart
    truong = field_asked(text)
    sz = asked_size_code(text)
    if sz: new["_size"] = sz
    rule = rule_intent(text, new, state["slots"])
    intent = llm_intent if llm_intent in INTENTS else rule
    if intent == "chitchat" and rule != "chitchat":
        intent = rule
    if "ma_don" in new:
        intent = "order_check"
    if intent == "chitchat" and state["last_bot"] and state["intent"] in FOLLOW_KEYS \
            and any(k in new for k in FOLLOW_KEYS[state["intent"]]):
        intent = state["intent"]
    if intent != state["intent"]:
        keep = KEEP_ALWAYS | ({"ten_sp", "loai_do"} if intent == "size_advice" else set())
        state["slots"] = {k: v for k, v in state["slots"].items() if k in keep}
    if "ten_sp" in new or "loai_do" in new:
        for k in ("ten_sp", "tu_khoa"):
            if k not in new: state["slots"].pop(k, None)
    state["intent"] = intent
    state["slots"].update(new)
    if truong and intent == "product_info": state["slots"]["truong"] = truong
    return state

# ---------------------------------------------------------------- 8) Hỏi "còn size M không" -> trả lời đúng size đó ----
try:
    _act_product_v4
except NameError:
    _act_product_v4 = act_product                                   # chỉ bọc 1 lần dù chạy lại cell

def act_product(S):
    kind, out = _act_product_v4(S)
    sz = S.get("_size")
    if sz and kind == "inform" and S.get("ten_sp"):
        p = next((x for x in CATALOG if x["ten"] == S["ten_sp"]), None)
        if p and out.startswith(f"Dạ {p['ten']}") and "\n" in out:
            head = (f"Dạ {p['ten']} ({p['id']}) hiện có size {sz} nè nàng." if sz in p["size"]
                    else f"Dạ {p['ten']} ({p['id']}) hiện không có size {sz} ạ, mẫu này có size {', '.join(p['size'])}.")
            out = head + "\n" + out.split("\n", 1)[1]
    return kind, out

build_match_index()
print("Bản vá v5 sẵn sàng:", len(NAME_TOK), "món được lập chỉ mục")


# ==================== Bản vá v6 ====================
# ============ BẢN VÁ v6: không hỏi lại số đo | phối đồ theo món đang xem | giảm độ trễ khi dùng Gemini ============
# Đặt NGAY SAU cell "Bản vá v5". Cell chỉ định nghĩa lại / bọc hàm, chạy lại nhiều lần vẫn an toàn.
import os, re, copy, time

# ---------------------------------------------------------------- 1) Số đo đã biết: dùng lại, không hỏi lại ----
_SIZE_CACHE = {}

def known_measures(S):
    h, w = S.get("chieu_cao"), S.get("can_nang")
    if isinstance(h, (int, float)) and isinstance(w, (int, float)) and 140 <= h <= 200 and 30 <= w <= 120:
        return int(h), int(w)
    return None

def my_size(S):
    """Size đề xuất theo số đo khách đã nói (gọi API size-advice, có nhớ kết quả)."""
    m = known_measures(S)
    if not m:
        return None
    if m not in _SIZE_CACHE:
        _, r = api_get("/v1/size-advice", chieu_cao=m[0], can_nang=m[1])
        _SIZE_CACHE[m] = r["de_xuat"]["size"]
    return _SIZE_CACHE[m]

def _prod_by_id(pid):
    return next((p for p in CATALOG if p["id"] == pid), None)

def _size_tag(p, sz):
    if not sz or not p:
        return ""
    return f" - có size {sz}" if sz in p["size"] else f" - không có size {sz} (có {', '.join(p['size'])})"

def act_lookbook(S):
    st, d = api_get(f"/v1/lookbook/{quote(S['dip_mac'], safe='')}")
    if st == 404:
        S.pop("dip_mac", None)
        return "ask", ask_text("dip_mac")
    sz = my_size(S)
    lines, thieu = [], 0
    for p in d["items"]:
        note = "" if p["tinh_trang"] == "Còn hàng" else f" ({p['tinh_trang'].lower()})"
        cp = _prod_by_id(p["id"])
        tag = _size_tag(cp, sz)
        thieu += bool(sz and cp and sz not in cp["size"])
        lines.append(f"- {p['ten']} ({p['id']}): {vnd(p['gia'])}{note}{tag}")
    head = f"Dạ cho dịp {d['dip_mac']}, SUPEO gợi ý nàng set '{d['ten_bo']}' phong cách {d['phong_cach']} nè:\n"
    body = "\n".join(lines) + f"\nTổng cả set: {vnd(d['tong_gia_tien'])}.\n{d['ghi_chu']}\n"
    if sz:
        h, w = known_measures(S)
        foot = (f"Theo số đo nàng đã chia sẻ ({h}cm, {w}kg), nàng hợp Size {sz}. "
                + ("Cả set đều có size này, nàng chốt được luôn nhé! " if not thieu else
                   "Món nào chưa có size này SUPEO đã ghi chú ở trên để nàng cân nhắc nhé! ")
                + "Nàng muốn xem chi tiết món nào cứ nhắn tên món, hoặc hỏi SUPEO phối thêm áo/quần với món nàng thích nha!")
    else:
        foot = ("Nàng cho SUPEO xin chiều cao, cân nặng để chọn size từng món, "
                "hoặc nhắn tên món nàng thích để SUPEO gửi chi tiết nhé!")
    return "inform", head + body + foot

# ---------------------------------------------------------------- 2) Danh sách nhiều mẫu: ghi chú size theo số đo đã biết ----
_LIST_ASK = "Nàng cũng có thể cho SUPEO biết chiều cao, cân nặng để chọn đúng size."
try:
    _act_product_v5
except NameError:
    _act_product_v5 = act_product              # act_product ở đây đã gồm bản bọc "còn size M không" của v5

def act_product(S):
    kind, out = _act_product_v5(S)
    sz = my_size(S)
    if sz and kind == "inform" and _LIST_ASK in out:
        h, w = known_measures(S)
        new_lines = []
        for ln in out.split("\n"):
            m = re.match(r"^\d+\. .*\((P\d+)\)", ln)
            new_lines.append(ln + _size_tag(_prod_by_id(m.group(1)), sz) if m else ln)
        out = "\n".join(new_lines).replace(
            _LIST_ASK, f"Theo số đo nàng đã chia sẻ ({h}cm, {w}kg), nàng hợp Size {sz}; SUPEO đã ghi chú size ở từng mẫu rồi nè.")
    return kind, out

# ---------------------------------------------------------------- 3) "phối với áo nào?" -> gợi ý theo món đang nói tới ----
PAIR_RE = (r"\bphoi\b.*\b(?:voi|cung|chung)\b|\b(?:voi|cung)\s+(?:ao|quan|vay|dam|blazer|chan vay|so mi)?\s*nao\b"
           r"|\bmac\s+(?:gi\s+)?(?:voi|cung|kem)\b|\bket hop\b|\b(?:di|hop|dung)\s+voi\s+(?:ao|quan|vay|dam|blazer)\b")
COMPLEMENT = {"Đầm": ["Áo Blazer"], "Áo Blazer": ["Đầm", "Áo sơ mi", "Quần ống rộng", "Chân váy Midi"],
              "Áo sơ mi": ["Quần ống rộng", "Chân váy Midi", "Áo Blazer"],
              "Quần ống rộng": ["Áo sơ mi", "Áo Blazer"], "Chân váy Midi": ["Áo sơ mi", "Áo Blazer"]}

def _is_pairing(text):
    return bool(re.search(PAIR_RE, strip_acc(text).lower()))

def _lookbook_sets_with(pid):
    out = []
    for dip in OCCASIONS:
        st, d = api_get(f"/v1/lookbook/{quote(dip, safe='')}")
        if st == 200 and any(i["id"] == pid for i in d["items"]):
            out.append(d)
    return out

def act_pairing(state):
    S = state["slots"]
    anchor = _prod_by_id(next((p["id"] for p in CATALOG if p["ten"] == S.get("ten_sp")), None))
    if not anchor:
        return "ask", ask_text("dip_mac")
    want = match_key(positive_part(state["last_text"]), CATS)            # khách nói rõ muốn phối với nhóm nào ('phối với áo sơ mi')
    if want == anchor["loai"] and not re.search(r"\bkhac\b|\bnua\b", strip_acc(state["last_text"]).lower()):
        want = None
    sz = my_size(S)
    picks, seen = [], {anchor["id"]}
    for d in _lookbook_sets_with(anchor["id"]):                            # 1) món đã được SUPEO phối sẵn cùng món này
        for it in d["items"]:
            cp = _prod_by_id(it["id"])
            if it["id"] in seen or not cp or (want and cp["loai"] != want):
                continue
            seen.add(it["id"])
            picks.append((cp, f" - SUPEO đã phối sẵn trong set '{d['ten_bo']}' (dịp {d['dip_mac']})"))
    cats = [want] if want else COMPLEMENT.get(anchor["loai"], [])
    for c in cats:                                                          # 2) bổ sung món còn hàng ở nhóm có thể phối
        cand = [p for p in CATALOG if p["loai"] == c and p["id"] not in seen and p["ton_kho"] > 0
                and (not sz or sz in p["size"])]
        for cp in cand[:2 if len(picks) >= 2 else 3]:
            seen.add(cp["id"])
            picks.append((cp, ""))
    picks = picks[:5]
    h = f"Dạ để phối với {anchor['ten']} ({anchor['id']}), SUPEO gợi ý nàng:\n"
    if not picks:
        return "inform", h + "Hiện SUPEO chưa có món phù hợp để phối kèm, nàng nói dịp mặc để SUPEO gợi ý set khác nhé!"
    lines = [f"- {cp['ten']} ({cp['id']}): {vnd(cp['gia'])}{tag}{_size_tag(cp, sz)}" for cp, tag in picks]
    if sz:
        m = known_measures(S)
        foot = f"\nTheo số đo nàng đã chia sẻ ({m[0]}cm, {m[1]}kg), nàng hợp Size {sz}, SUPEO ưu tiên món có size này nhé!"
    else:
        foot = "\nNàng cho SUPEO xin chiều cao, cân nặng để chọn size từng món nhé!"
    return "inform", h + "\n".join(lines) + foot + "\nNàng thích món nào cứ nhắn tên món, SUPEO gửi chi tiết ngay."

def resolve_reference(text, ptxt, cat, state):
    """'áo này', 'nó', 'mẫu đó', 'còn hàng không?', 'phối với áo nào?' -> món đang nói tới (loại đồ so với chính món đang nhớ)."""
    f = state.get("focus", {})
    if not f.get("ten_sp"):
        return None
    p = next((x for x in CATALOG if x["ten"] == f["ten_sp"]), None)
    nfc = _nfc(text)
    na = strip_acc(nfc)
    if re.search(PAIR_RE, na) and not re.search(r"voucher|ma giam|gio hang|don hang|ma don|\bsp\d", na):
        return f["ten_sp"]                                                  # câu hỏi phối đồ luôn xoay quanh món đang xem
    if DEICTIC_ACC.search(nfc) or DEICTIC_NOACC.search(na):
        if cat and cat != "Váy" and p and cat != p["loai"]:
            return None
        return f["ten_sp"]
    if field_asked(text) and not cat and not state["slots"].get("ten_sp") \
            and not re.search(r"voucher|ma giam|gio hang|don hang|ma don|\bsp\d", na):
        return f["ten_sp"]
    return None

_rule_intent_v5 = rule_intent

def rule_intent(text, new, state_slots):
    if "ten_sp" in new and _is_pairing(text) and "dip_mac" not in new:
        return "lookbook"
    return _rule_intent_v5(text, new, state_slots)

_policy_v5 = policy

def policy(state):
    S = state["slots"]
    if state["intent"] == "lookbook" and not S.get("dip_mac") and S.get("ten_sp") and _is_pairing(state["last_text"]):
        return act_pairing(state)
    return _policy_v5(state)

# ---------------------------------------------------------------- 4) Tốc độ khi bật Gemini ----
# (a) DST bằng Gemini chỉ gọi khi luật không hiểu câu (mặc định 'fallback'); 'always' = như cũ; 'off' = không gọi.
# (b) NLG (viết lại câu trả lời bằng Gemini) mặc định TẮT: mẫu câu đã tự nhiên, bớt 1 lượt gọi mạng, không nguy cơ bịa số.
LLM_DST_MODE = os.environ.get("SUPEO_LLM_DST", "fallback")                  # fallback | always | off
LLM_NLG = os.environ.get("SUPEO_NLG", "0") == "1"                           # 1 = bật viết lại bằng Gemini
TIMING = os.environ.get("SUPEO_DEBUG_TIMING", "0") == "1"                   # 1 = in thời gian từng lượt ra console

_FAQ_RE = (r"cam on|thanks|thank|tam biet|\bbye\b|hen gap lai|thoi nhe|hoa toc|giao nhanh|phi ship|phi van chuyen|giao hang|"
           r"bao lau|thoi gian giao|giao toan quoc|doi tra|bao hanh|thanh toan|\bcod\b|hotline|so dien thoai|dia chi|"
           r"cua hang o dau|\b(?:xin chao|chao|hello|hi|alo|shop oi|ban oi)\b")

try:
    _update_state_v5
except NameError:
    _update_state_v5 = update_state

def update_state(state, text):
    global USE_LLM
    if not USE_LLM or LLM_DST_MODE == "always":
        if LLM_DST_MODE == "off" and USE_LLM:
            prev, USE_LLM = USE_LLM, False
            try:
                return _update_state_v5(state, text)
            finally:
                USE_LLM = prev
        return _update_state_v5(state, text)
    if LLM_DST_MODE == "off":
        return _update_state_v5(state, text)
    snap = copy.deepcopy(state)
    prev, USE_LLM = USE_LLM, False                                           # thử bằng luật trước (vài mili-giây)
    try:
        _update_state_v5(state, text)
    finally:
        USE_LLM = prev
    if state["intent"] == "chitchat" and not re.search(_FAQ_RE, strip_acc(text).lower()):
        state.clear(); state.update(snap)                                    # luật không hiểu -> nhờ Gemini phân tích
        return _update_state_v5(state, text)
    return state

_nlg_v5 = nlg

def nlg(facts, user_text, state=None):
    return _nlg_v5(facts, user_text, state) if LLM_NLG else facts

_bot_reply_v5 = bot_reply

def bot_reply(text, state):
    t0 = time.time()
    n0 = _G["calls"] if "_G" in globals() and "calls" in _G else 0
    reply = _bot_reply_v5(text, state)
    if TIMING:
        n1 = _G["calls"] if "_G" in globals() and "calls" in _G else 0
        print(f"[timing] {1000 * (time.time() - t0):.0f} ms | gọi Gemini {n1 - n0} lần | intent={state['intent']} | {text[:40]!r}")
    return reply

print(f"Bản vá v6 sẵn sàng (DST Gemini: {LLM_DST_MODE}, NLG Gemini: {'bật' if LLM_NLG else 'tắt'})")


# ==================== Bật LLM (Gemini) ====================
# Bật LLM cho DST + NLG bằng Gemini API (Google AI Studio). Không có key / lỗi kết nối thì chatbot tự chạy bằng luật + regex.
ENABLE_LLM = True
LLM_BACKEND = "gemini"               
GEMINI_MODEL = "gemini-3.5-flash-lite"   # đã kiểm tra: dùng được với tài khoản này
GEMINI_FALLBACKS = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-flash-lite-latest", "gemini-3.5-flash", "gemini-3-flash-preview"]   # chỉ gồm model đã trả HTTP 200; gemini-2.5-* đã bị Google ngừng cấp cho tài khoản mới
MODEL_ID = "Qwen/Qwen2.5-3B-Instruct" 

# Key: nên lưu ở Kaggle -> Add-ons -> Secrets (Label: GEMINI_API_KEY, bật Attach). Chỉ dán thẳng vào đây khi chạy cá nhân, KHÔNG chia sẻ notebook kèm key.
GEMINI_API_KEY = ""

import os, time, json, requests
USE_LLM = False

if ENABLE_LLM and LLM_BACKEND == "gemini":
    GBASE = "https://generativelanguage.googleapis.com/v1beta"
    _G = {"model": GEMINI_MODEL, "key": "", "fails": 0, "block_until": 0.0, "repicked": False, "tried": set(), "calls": 0, "think": {}}
    _SESSION = requests.Session()          # giữ kết nối (keep-alive): bỏ bước bắt tay TLS mỗi lượt gọi

    def _get_key():
        if GEMINI_API_KEY.strip():
            return GEMINI_API_KEY.strip()
        try:
            from kaggle_secrets import UserSecretsClient
            k = UserSecretsClient().get_secret("GEMINI_API_KEY")
            if k: return k.strip()
        except Exception:
            pass
        return os.environ.get("GEMINI_API_KEY", "").strip()

    def _pick_model(key):
        """Liệt kê model của tài khoản, chọn 1 model flash hỗ trợ generateContent."""
        r = requests.get(f"{GBASE}/models", headers={"x-goog-api-key": key}, params={"pageSize": 200}, timeout=20)
        r.raise_for_status()
        names = [m["name"].split("/", 1)[1] for m in r.json().get("models", [])
                 if "generateContent" in m.get("supportedGenerationMethods", [])]
        bad = ("lite", "image", "tts", "audio", "live", "embedding", "vision", "robotics", "computer")
        cand = [n for n in names if "flash" in n and not any(b in n for b in bad)]
        pool = [n for n in cand if "preview" not in n and "exp" not in n] or cand
        if not pool:
            raise RuntimeError("Không tìm thấy model flash nào hỗ trợ generateContent")
        return sorted(pool)[-1]

    def _next_model(cur):
        """Model kế tiếp chưa thử (để đổi sang khi model hiện tại hết hạn mức hoặc không tồn tại)."""
        _G["tried"].add(cur)
        return next((m for m in [GEMINI_MODEL] + GEMINI_FALLBACKS if m not in _G["tried"]), None)

    def gemini_llm(messages, max_new_tokens=200, timeout=25):
        """Cùng giao diện với llm() cũ: messages = [{role: system|user|assistant, content}], trả về chuỗi văn bản."""
        if time.time() < _G["block_until"]:
            raise RuntimeError("Gemini tạm nghỉ sau nhiều lỗi liên tiếp")
        sys_txt = "\n".join(m["content"] for m in messages if m["role"] == "system")
        contents = []
        for m in messages:
            if m["role"] == "system":
                continue
            role = "model" if m["role"] == "assistant" else "user"
            if contents and contents[-1]["role"] == role:
                contents[-1]["parts"][0]["text"] += "\n" + m["content"]
            else:
                contents.append({"role": role, "parts": [{"text": m["content"]}]})
        last_err = None
        for attempt in range(len(GEMINI_FALLBACKS) + 3):
            model = _G["model"]
            gen = {"temperature": 0, "maxOutputTokens": max_new_tokens + 512}
            tc = _THINK_OPTS[_G["think"].get(model, 0)]                 # giảm/tắt "suy nghĩ" (nguyên nhân chính khiến Gemini 3.x trả lời chậm)
            if tc is not None:
                gen["thinkingConfig"] = tc
            body = {"contents": contents, "generationConfig": gen}
            if sys_txt:
                body["system_instruction"] = {"parts": [{"text": sys_txt}]}
            try:
                _G["calls"] += 1
                r = _SESSION.post(f"{GBASE}/models/{model}:generateContent",
                                  headers={"x-goog-api-key": _G["key"], "Content-Type": "application/json"}, json=body, timeout=timeout)
                if r.status_code == 400 and "hink" in r.text and _G["think"].get(model, 0) < len(_THINK_OPTS) - 1:
                    _G["think"][model] = _G["think"].get(model, 0) + 1   # model không nhận mức suy nghĩ này -> thử mức nhẹ hơn / bỏ cấu hình
                    continue
                if r.status_code in (404, 429):
                    last_err = f"HTTP {r.status_code} (model {model}): {r.text[:300]}"
                    nxt = _next_model(model)
                    if nxt:
                        _G["model"] = nxt                              # model này hết hạn mức / không có -> thử model khác
                        continue
                    if r.status_code == 404 and not _G["repicked"]:
                        _G["repicked"] = True
                        _G["model"] = _pick_model(_G["key"])
                        continue
                    if r.status_code == 429:
                        time.sleep(2 * (attempt + 1))
                        continue
                    break
                if r.status_code in (500, 502, 503, 504):
                    last_err = f"HTTP {r.status_code}"
                    time.sleep(2 * (attempt + 1))
                    continue
                if r.status_code >= 400:
                    raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
                parts = r.json()["candidates"][0]["content"]["parts"]
                _G["fails"] = 0
                _G["tried"] = set()
                return "".join(p.get("text", "") for p in parts)
            except Exception as e:
                last_err = str(e)
                if "HTTP 4" in last_err and "HTTP 429" not in last_err:
                    break                                               # lỗi 4xx (sai key, sai tham số): không thử lại
        _G["fails"] += 1
        if _G["fails"] >= 3:
            _G["block_until"] = time.time() + 60                        # lỗi liên tiếp: nghỉ 60 giây, bot dùng luật + regex
        raise RuntimeError(last_err or "Gemini lỗi")

    # Mức "suy nghĩ" thử lần lượt (nhẹ nhất trước); model nào trả 400 vì không nhận thì tự chuyển sang mức kế tiếp / bỏ cấu hình.
    _THINK_OPTS = [{"thinkingLevel": "minimal"}, {"thinkingLevel": "low"}, {"thinkingBudget": 0}, None]
    _G["key"] = _get_key()
    if not _G["key"]:
        print("Chưa có GEMINI_API_KEY -> chạy chế độ luật + regex.\n"
              "Thêm key: Add-ons -> Secrets -> Add secret (Label: GEMINI_API_KEY), bật Attach, rồi chạy lại cell này.")
    else:
        try:
            llm = gemini_llm
            out = llm([{"role": "system", "content": "Chỉ trả lời đúng một từ."}, {"role": "user", "content": "Hãy nói: OK"}], 10)
            USE_LLM = True
            print(f"Đã kết nối Gemini ({_G['model']}). Phản hồi thử: {out.strip()[:30]!r}")
        except Exception as e:
            print(f"Không kết nối được Gemini -> chạy chế độ luật + regex.\nChi tiết: {e}\n"
                  "Kiểm tra: Internet đã bật chưa, key có đúng không, tên model còn dùng được không.")

elif ENABLE_LLM and LLM_BACKEND == "qwen":
    import torch
    from transformers import AutoTokenizer, AutoModelForCausalLM
    tok = AutoTokenizer.from_pretrained(MODEL_ID)
    model = AutoModelForCausalLM.from_pretrained(MODEL_ID, torch_dtype=torch.float16, device_map="auto")

    def llm(messages, max_new_tokens=200):
        prompt = tok.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
        inputs = tok(prompt, return_tensors="pt").to(model.device)
        with torch.no_grad():
            out = model.generate(**inputs, max_new_tokens=max_new_tokens, do_sample=False)
        return tok.decode(out[0][inputs["input_ids"].shape[1]:], skip_special_tokens=True)

    USE_LLM = True
    print("Đã tải", MODEL_ID)
else:
    print("Chạy chế độ luật + regex (USE_LLM = False)")


# ==================== Giao diện Gradio ====================
import gradio as gr
import os, re

GREETING = ("Xin chào! Mình là stylist AI của SUPEO. Mình có thể gợi ý phối đồ theo dịp, "
            "tư vấn size, cập nhật đơn hàng và mã giảm giá cho nàng.")
EXAMPLES = ["Quần ống suông cạp chun sau có màu gì vậy?", "Mình cao 1m69, 58kg nên lấy size gì ạ?",
            "Đi dự tiệc ngoài trời nên phối đồ thế nào?", "Cho mình tra tình trạng đơn #SP3220 với",
            "Đơn 1tr thì dùng mã giảm giá nào được?", "Mình đang tìm váy"]

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500&family=Be+Vietnam+Pro:wght@400;500;600&display=swap');
body, .gradio-container { background:#e2e2e2 !important; font-family:'Be Vietnam Pro',system-ui,sans-serif !important; }
footer { display:none !important; }
#app { max-width:480px; margin:12px auto !important; background:#fff; border-radius:28px; padding:0 !important;
       overflow:hidden; gap:0 !important; border:none !important; box-shadow:none !important; }
#app > .block, #app .form { border:none !important; box-shadow:none !important; background:transparent !important; }
.hd { padding:16px 22px; border-bottom:1px solid #ececec; }
.hd .logo { font-family:'Playfair Display',Georgia,serif; font-size:28px; letter-spacing:.14em; line-height:1.1; color:#111; }
.hd .sub { font-size:11px; letter-spacing:.2em; color:#777; margin-top:2px; }
#exbox { padding:10px 14px 4px !important; background:transparent !important; border:none !important; gap:0 !important; }
#exlabel { font-size:12px; color:#777; margin:0 0 6px; }
#chips { flex-wrap:wrap !important; gap:6px !important; }
#chips button { flex:0 0 auto !important; min-width:0 !important; width:auto !important; background:#fff !important; color:#111 !important;
       border:1px solid #ddd !important; border-radius:16px !important; padding:6px 12px !important; font-size:13px !important;
       font-weight:400 !important; box-shadow:none !important; }
#chips button:hover { border-color:#111 !important; }
#chat, #chat .bubble-wrap { background:#fff !important; border:none !important; }
#chat .message-row { padding:0 10px; }
#chat .message { font-size:15px !important; line-height:1.55; border-radius:20px !important; box-shadow:none !important; }
#chat .bot, #chat [data-testid="bot"] { background:#f4f4f4 !important; border:none !important; color:#111 !important; }
#chat .user, #chat [data-testid="user"] { background:#000 !important; border:none !important; color:#fff !important;
       font-weight:600; border-radius:20px 20px 4px 20px !important; }
#chat .user *, #chat [data-testid="user"] * { color:#fff !important; }
#chat .message ul, #chat .message ol { margin:4px 0 4px 18px; }
#chat .message-buttons, #chat .avatar-container { display:none !important; }
#inrow { padding:6px 14px 0; gap:8px !important; flex-wrap:nowrap !important; align-items:center; }
#box textarea, #box input { border:1px solid #ddd !important; border-radius:22px !important; padding:11px 16px !important;
       font-size:15px !important; box-shadow:none !important; background:#fff !important; }
#box { border:none !important; background:transparent !important; box-shadow:none !important; }
#send { background:#000 !important; color:#fff !important; border:none !important; border-radius:22px !important;
        font-weight:600; min-width:72px !important; max-width:90px; }
#new { background:#fff !important; color:#111 !important; border:1px solid #ddd !important; border-radius:22px !important;
       margin:8px 14px 14px; width:calc(100% - 28px); }
"""

def render(hist):
    """State giữ danh sách tin nhắn dạng đơn giản; trả về nguyên danh sách để hiển thị."""
    return list(hist)

def respond(msg, state):
    msg = (msg or "").strip()
    if callable(state):                                  # phòng khi trang chưa kịp khởi tạo state
        state = make_state()
    hist = state.setdefault("hist", [{"role": "assistant", "content": GREETING}])
    if msg:
        reply = bot_reply(msg, state)
        hist += [{"role": "user", "content": msg}, {"role": "assistant", "content": reply}]
    show_examples = not any(m["role"] == "user" for m in hist)
    return "", render(hist), state, gr.update(visible=show_examples)

def reset():
    st = new_state()
    st["hist"] = [{"role": "assistant", "content": GREETING}]
    return "", render(st["hist"]), st, gr.update(visible=True)

def make_state():
    st = new_state()
    st["hist"] = [{"role": "assistant", "content": GREETING}]
    return st

theme = gr.themes.Base(font=["Be Vietnam Pro", "system-ui", "sans-serif"])

with gr.Blocks(title="SUPEO - AI Stylist", css=CSS, theme=theme) as demo:
    state = gr.State(make_state)
    with gr.Column(elem_id="app"):
        gr.HTML('<div class="hd"><div class="logo">SUPEO</div><div class="sub">AI STYLIST</div></div>')
        with gr.Column(elem_id="exbox") as exbox:                       # câu hỏi mẫu: nằm trên đầu, biến mất sau câu hỏi đầu tiên
            gr.HTML('<div id="exlabel">Câu hỏi mẫu</div>')
            with gr.Row(elem_id="chips"):
                chip_btns = [gr.Button(e, size="sm") for e in EXAMPLES]
        chat = gr.Chatbot(value=render([{"role": "assistant", "content": GREETING}]), type="messages", height=460,
                          show_label=False, container=False, elem_id="chat")
        with gr.Row(elem_id="inrow"):
            box = gr.Textbox(placeholder="Nhập câu hỏi rồi nhấn Enter...", show_label=False, container=False,
                             lines=1, max_lines=1, scale=6, elem_id="box")
            send = gr.Button("Gửi", variant="primary", scale=1, elem_id="send")
        clear = gr.Button("Cuộc trò chuyện mới", elem_id="new")
    outs = [box, chat, state, exbox]
    box.submit(respond, [box, state], outs, show_progress="hidden")
    send.click(respond, [box, state], outs, show_progress="hidden")
    for b, e in zip(chip_btns, EXAMPLES):
        b.click(lambda st, e=e: respond(e, st), [state], outs, api_name=False, show_progress="hidden")
    clear.click(reset, None, outs, show_progress="hidden")

def cloudflare_link(port):
    """Dự phòng khi Gradio không tạo được link share: mở tunnel Cloudflare tới cổng local (cần bật Internet trên Kaggle)."""
    import subprocess, time, urllib.request, stat
    exe = "/kaggle/working/cloudflared" if os.path.isdir("/kaggle/working") else "./cloudflared"
    if not os.path.exists(exe):
        urllib.request.urlretrieve("https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64", exe)
        os.chmod(exe, os.stat(exe).st_mode | stat.S_IEXEC)
    log = open("cloudflared.log", "w")
    subprocess.Popen([exe, "tunnel", "--no-autoupdate", "--url", f"http://127.0.0.1:{port}"], stdout=log, stderr=log)
    for _ in range(60):
        time.sleep(1)
        m = re.search(r"https://[a-z0-9-]+\.trycloudflare\.com", open("cloudflared.log").read())
        if m:
            return m.group(0)
    return None

LAUNCH = globals().get("LAUNCH", True)
if LAUNCH:
    SHARE = os.environ.get("SUPEO_SHARE", "1") != "0"                    # SUPEO_SHARE=0: chỉ chạy local (http://127.0.0.1:7860), không tạo link công khai
    _, local_url, share_url = demo.launch(share=SHARE, debug=False)      # share=True: link công khai tạm thời. Giữ phiên Kaggle mở khi demo.
    if SHARE and not share_url:
        try:
            link = cloudflare_link(int(local_url.rsplit(":", 1)[1].strip("/")))
        except Exception as e:
            link = None
            print("Không tạo được tunnel dự phòng:", e)
        print("\n>>> LINK CHATBOT:", link or "chưa tạo được link - kiểm tra Internet đã bật chưa (Settings -> Internet on)")

try:
    audit_leak()          # kiểm tra lại rò rỉ với EXAMPLES của giao diện
except NameError:
    pass
