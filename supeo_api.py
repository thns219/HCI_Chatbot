"""
SUPEO Mock API (FastAPI) - đọc dữ liệu từ supeo_data_complete.xlsx

Chạy độc lập:   SUPEO_XLSX=supeo_data_complete.xlsx uvicorn supeo_api:app --port 8008
Tài liệu API:   http://127.0.0.1:8008/docs

Endpoint:
  GET /v1/products?loai=&q=                 Tra cứu sản phẩm (loại, từ khóa màu/chất liệu/tên)
  GET /v1/products/{id}                     Chi tiết 1 sản phẩm
  GET /v1/size-advice?chieu_cao=&can_nang=  Tư vấn size (Max-Boundary)
  GET /v1/lookbook                          Danh sách các dịp
  GET /v1/lookbook/{occasion}               Set đồ theo dịp
  GET /v1/orders/{order_id}/tracking        Tra cứu đơn hàng (#SPxxxx)
  GET /v1/vouchers/available?cart_total=    Mã còn hạn, hien_thi_chatbot = TRUE
  GET /v1/vouchers/codes                    Danh sách mã (chỉ chuỗi mã, để chatbot nhận diện)
  GET /v1/vouchers/{ma}?cart_total=         Kiểm tra 1 mã khách nhập

Biến môi trường:
  SUPEO_XLSX   đường dẫn file Excel (mặc định tự tìm)
  SUPEO_TODAY  ngày giả lập, định dạng YYYY-MM-DD hoặc dd/mm/yyyy (mặc định: hôm nay)
"""
import glob
import os
import re
import unicodedata
from datetime import date, datetime
from typing import Optional

import pandas as pd
from fastapi import FastAPI, HTTPException, Query

XLSX_NAME = "supeo_data_complete.xlsx"


# ----------------------------------------------------------------------------
# Tiện ích
# ----------------------------------------------------------------------------
def strip_acc(s) -> str:
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def norm(s) -> str:
    return strip_acc(s).lower().strip()


def find_xlsx() -> str:
    cands = [os.getenv("SUPEO_XLSX"), XLSX_NAME, os.path.join("/kaggle/working", XLSX_NAME),
             os.path.join("/content", XLSX_NAME), os.path.join("/mnt/user-data/uploads", XLSX_NAME)]
    for pat in ("/kaggle/input/**/" + XLSX_NAME, "/content/**/" + XLSX_NAME):
        cands += glob.glob(pat, recursive=True)
    for c in cands:
        if c and os.path.exists(c):
            return c
    raise FileNotFoundError(
        f"Không tìm thấy {XLSX_NAME}. Hãy upload file (Kaggle: Add Data -> Upload) "
        f"hoặc đặt biến môi trường SUPEO_XLSX.")


def parse_date(x) -> Optional[date]:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    s = str(x).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def fmt_date(d: Optional[date]) -> Optional[str]:
    return d.strftime("%d/%m/%Y") if d else None


def to_int(x, default=0) -> int:
    try:
        return int(float(x))
    except (TypeError, ValueError):
        return default


def to_bool(x) -> bool:
    return str(x).strip().upper() in ("TRUE", "1", "YES", "CÓ")


def get_today(today: Optional[str] = None) -> date:
    s = today or os.getenv("SUPEO_TODAY")
    if s:
        d = parse_date(s)
        if d:
            return d
    return date.today()


def read_sheet(path: str, sheet: str, first_col: str) -> list:
    """Sheet có dòng tiêu đề + dòng trống phía trên: tự dò dòng header theo tên cột đầu."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    hdr = None
    for i in range(len(raw)):
        if str(raw.iloc[i, 0]).strip() == first_col:
            hdr = i
            break
    if hdr is None:
        raise ValueError(f"Sheet '{sheet}': không thấy cột '{first_col}'")
    cols = [str(c).strip() for c in raw.iloc[hdr]]
    df = raw.iloc[hdr + 1:].copy()
    df.columns = cols
    df = df[df[first_col].notna()]
    rows = []
    for rec in df.to_dict("records"):
        rows.append({k: (None if (v is None or (not isinstance(v, str) and pd.isna(v))) else v)
                     for k, v in rec.items()})
    return rows


def tinh_trang(ton: int) -> str:
    if ton <= 0:
        return "Hết hàng"
    if ton < 20:
        return "Sắp hết hàng"
    return "Còn hàng"


# ----------------------------------------------------------------------------
# Database layer (đọc 1 lần khi khởi động; POST /v1/reload để nạp lại)
# ----------------------------------------------------------------------------
class DB:
    def __init__(self, path: str):
        self.path = path
        self.load()

    def load(self):
        p = self.path
        self.products = []
        for r in read_sheet(p, "products", "id"):
            ton = to_int(r.get("ton_kho"))
            self.products.append({
                "id": str(r["id"]).strip(), "ten": r["ten"], "loai": r["loai"],
                "gia": to_int(r["gia"]),
                "mau": [x.strip() for x in str(r.get("mau") or "").split(",") if x.strip()],
                "chat_lieu": r.get("chat_lieu"),
                "size": [x.strip() for x in str(r.get("size") or "").split(",") if x.strip()],
                "phom_dang": r.get("phom_dang"), "ton_kho": ton, "tinh_trang": tinh_trang(ton),
                "mo_ta": r.get("mo_ta"),
            })
        self.by_id = {x["id"]: x for x in self.products}
        self.categories = list(dict.fromkeys(x["loai"] for x in self.products))

        self.sizes = []
        for r in read_sheet(p, "size_chart", "size"):
            self.sizes.append({
                "size": str(r["size"]).strip(),
                "chieu_cao": [to_int(r["chieu_cao_min"]), to_int(r["chieu_cao_max"])],
                "can_nang": [to_int(r["can_nang_min"]), to_int(r["can_nang_max"])],
                "vong_nguc": to_int(r.get("vong_nguc")), "vong_eo": to_int(r.get("vong_eo")),
                "vong_mong": to_int(r.get("vong_mong")), "loi_khuyen": r.get("loi_khuyen_tu_van"),
            })

        self.lookbook = []
        for r in read_sheet(p, "lookbook", "dip_mac"):
            self.lookbook.append({
                "dip_mac": str(r["dip_mac"]).strip(), "ten_bo": r.get("ten_bo"),
                "product_ids": [x.strip() for x in str(r.get("product_ids") or "").split(",") if x.strip()],
                "tong_gia_excel": to_int(r.get("tong_gia_tien")),
                "phong_cach": r.get("phong_cach"), "ghi_chu": r.get("ghi_chu_tu_van"),
            })

        self.orders = {}
        for r in read_sheet(p, "orders", "ma_don"):
            ma = str(r["ma_don"]).strip().upper()
            self.orders[ma] = {
                "ma_don": ma, "trang_thai": r.get("trang_thai"),
                "ngay_dat": fmt_date(parse_date(r.get("ngay_dat"))),
                "du_kien": fmt_date(parse_date(r.get("du_kien"))),
                "dvvc": r.get("dvvc"), "ma_van_don": r.get("ma_van_don"),
                "san_pham": [x.strip() for x in str(r.get("san_pham") or "").split(";") if x.strip()],
                "tong_tien": to_int(r.get("tong_tien")),
            }

        self.vouchers = []
        for r in read_sheet(p, "vouchers", "ma"):
            mo_ta = str(r.get("mo_ta") or "")
            m = re.search(r"(\d+)\s*%", mo_ta)
            if m:
                loai_giam = "phan_tram"
            elif "vận chuyển" in mo_ta.lower():
                loai_giam = "van_chuyen"
            else:
                loai_giam = "tien_mat"
            self.vouchers.append({
                "ma": str(r["ma"]).strip(), "mo_ta": mo_ta, "giam_gia": to_int(r.get("giam_gia")),
                "don_toi_thieu": to_int(r.get("don_toi_thieu")), "han": parse_date(r.get("han_dung")),
                "doi_tuong": r.get("doi_tuong"), "hien_thi": to_bool(r.get("hien_thi_chatbot")),
                "dieu_khoan": r.get("dieu_khoan"), "loai_giam": loai_giam,
                "pct": int(m.group(1)) if m else None,
            })


db: Optional[DB] = None


def get_db() -> DB:
    global db
    if db is None:
        db = DB(find_xlsx())
    return db


app = FastAPI(title="SUPEO Mock API", version="1.0",
              description="Mock API/Database layer cho chatbot SUPEO (SGP-TOD). Nguồn: supeo_data_complete.xlsx")


# ----------------------------------------------------------------------------
# Health
# ----------------------------------------------------------------------------
@app.get("/")
@app.get("/health")
def health():
    d = get_db()
    return {"status": "ok", "file": os.path.basename(d.path), "today": fmt_date(get_today()),
            "so_luong": {"san_pham": len(d.products), "lookbook": len(d.lookbook),
                         "don_hang": len(d.orders), "voucher": len(d.vouchers)}}


@app.post("/v1/reload")
def reload_db():
    get_db().load()
    return {"status": "reloaded"}


# ----------------------------------------------------------------------------
# 1. Products
# ----------------------------------------------------------------------------
def _match_loai(user_loai: str, cats: list) -> Optional[str]:
    n = norm(user_loai)
    for c in cats:
        if norm(c) == n:
            return c
    for c in cats:
        if n and (n in norm(c) or norm(c) in n):
            return c
    return None


@app.get("/v1/products")
def products(loai: Optional[str] = None, q: Optional[str] = None,
             con_hang: Optional[bool] = None, gia_max: Optional[int] = None,
             limit: int = Query(50, ge=1, le=200)):
    d = get_db()
    items = d.products
    loai_hd = None
    if loai:
        loai_hd = _match_loai(loai, d.categories)
        items = [x for x in items if x["loai"] == loai_hd] if loai_hd else []
    if q:
        toks = [t for t in re.findall(r"[a-z0-9]+", norm(q))]
        def hit(x):
            hay = norm(" ".join([x["ten"], ",".join(x["mau"]), str(x["chat_lieu"]),
                                 str(x["phom_dang"]), str(x["mo_ta"])]))
            return all(re.search(r"\b" + re.escape(t) + r"\b", hay) for t in toks)
        items = [x for x in items if hit(x)]
    if con_hang:
        items = [x for x in items if x["ton_kho"] > 0]
    if gia_max:
        items = [x for x in items if x["gia"] <= gia_max]
    items = items[:limit]
    return {"count": len(items), "loai": loai_hd, "q": q, "danh_muc": d.categories, "items": items}


@app.get("/v1/products/{pid}")
def product_detail(pid: str):
    x = get_db().by_id.get(pid.strip().upper())
    if not x:
        raise HTTPException(404, detail={"loi": "khong_tim_thay_san_pham", "id": pid})
    return x


# ----------------------------------------------------------------------------
# 2. Size (Max-Boundary)
# ----------------------------------------------------------------------------
def _band(value: float, sizes: list, key: str) -> int:
    """Chỉ số dải đầu tiên có value <= max. Thấp hơn min -> dải đầu; vượt max cuối -> dải cuối."""
    for i, s in enumerate(sizes):
        if value <= s[key][1]:
            return i
    return len(sizes) - 1


@app.get("/v1/size-advice")
def size_advice(chieu_cao: float = Query(..., ge=140, le=200), can_nang: float = Query(..., ge=30, le=120)):
    sizes = get_db().sizes
    ih, iw = _band(chieu_cao, sizes, "chieu_cao"), _band(can_nang, sizes, "can_nang")
    chosen = sizes[max(ih, iw)]                      # Max-Boundary: lệch dải -> chọn size lớn hơn
    return {
        "chieu_cao": chieu_cao, "can_nang": can_nang,
        "theo_chieu_cao": {"size": sizes[ih]["size"], "khoang": sizes[ih]["chieu_cao"]},
        "theo_can_nang": {"size": sizes[iw]["size"], "khoang": sizes[iw]["can_nang"]},
        "lech": ih != iw,
        "de_xuat": {"size": chosen["size"], "chieu_cao": chosen["chieu_cao"], "can_nang": chosen["can_nang"],
                    "vong_nguc": chosen["vong_nguc"], "vong_eo": chosen["vong_eo"],
                    "vong_mong": chosen["vong_mong"], "loi_khuyen": chosen["loi_khuyen"]},
        "ngoai_dai": {"chieu_cao": chieu_cao < sizes[0]["chieu_cao"][0] or chieu_cao > sizes[-1]["chieu_cao"][1],
                      "can_nang": can_nang < sizes[0]["can_nang"][0] or can_nang > sizes[-1]["can_nang"][1]},
        "thu_tu_size": [s["size"] for s in sizes],
    }


# ----------------------------------------------------------------------------
# 3. Lookbook
# ----------------------------------------------------------------------------
@app.get("/v1/lookbook")
def lookbook_list():
    return {"dip_mac": [x["dip_mac"] for x in get_db().lookbook]}


@app.get("/v1/lookbook/{occasion}")
def lookbook(occasion: str):
    d = get_db()
    n = norm(occasion)
    found = [x for x in d.lookbook if norm(x["dip_mac"]) == n]
    if not found:
        found = [x for x in d.lookbook if n and (n in norm(x["dip_mac"]) or norm(x["dip_mac"]) in n)]
    if len(found) != 1:
        raise HTTPException(404, detail={"loi": "khong_tim_thay_dip", "dip": occasion,
                                         "goi_y": [x["dip_mac"] for x in d.lookbook]})
    lb = found[0]
    items = [d.by_id[i] for i in lb["product_ids"] if i in d.by_id]
    tong = sum(x["gia"] for x in items)
    return {"dip_mac": lb["dip_mac"], "ten_bo": lb["ten_bo"], "phong_cach": lb["phong_cach"],
            "ghi_chu": lb["ghi_chu"], "items": items, "tong_gia_tien": tong,
            "tong_gia_excel": lb["tong_gia_excel"], "khop_gia": tong == lb["tong_gia_excel"]}


# ----------------------------------------------------------------------------
# 4. Orders
# ----------------------------------------------------------------------------
@app.get("/v1/orders/{order_id}/tracking")
def order_tracking(order_id: str):
    ma = re.sub(r"[\s\-]", "", order_id).upper()
    if not re.fullmatch(r"#?SP\d{3,}", ma):
        raise HTTPException(422, detail={"loi": "sai_dinh_dang", "ma_don": order_id,
                                         "dinh_dang": "#SPxxxx (ví dụ #SP2908)"})
    if not ma.startswith("#"):
        ma = "#" + ma
    o = get_db().orders.get(ma)
    if not o:
        raise HTTPException(404, detail={"loi": "khong_tim_thay_don", "ma_don": ma})
    return {**o, "da_ban_giao_vc": bool(o["ma_van_don"])}


# ----------------------------------------------------------------------------
# 5. Vouchers
# ----------------------------------------------------------------------------
def _estimate(v: dict, amount: int) -> int:
    if v["pct"]:
        return min(round(amount * v["pct"] / 100), v["giam_gia"])
    return v["giam_gia"]


def _view(v: dict, today: date, cart: Optional[int]) -> dict:
    out = {"ma": v["ma"], "mo_ta": v["mo_ta"], "giam_gia": v["giam_gia"], "don_toi_thieu": v["don_toi_thieu"],
           "han_dung": fmt_date(v["han"]), "con_lai_ngay": (v["han"] - today).days if v["han"] else None,
           "doi_tuong": v["doi_tuong"], "dieu_khoan": v["dieu_khoan"], "loai_giam": v["loai_giam"]}
    if cart is not None:
        du = cart >= v["don_toi_thieu"]
        out.update(du_dieu_kien=du, con_thieu=0 if du else v["don_toi_thieu"] - cart,
                   uoc_tinh_giam=_estimate(v, cart) if du else None)
    return out


@app.get("/v1/vouchers/available")
def vouchers_available(cart_total: Optional[int] = Query(None, ge=0), today: Optional[str] = None):
    t = get_today(today)
    vs = [v for v in get_db().vouchers if v["hien_thi"] and v["han"] and v["han"] >= t]
    views = [_view(v, t, cart_total) for v in vs]
    best, upsell = None, None
    if cart_total is None:
        views.sort(key=lambda x: x["don_toi_thieu"])
    else:
        ok = [x for x in views if x["du_dieu_kien"]]
        no = [x for x in views if not x["du_dieu_kien"]]
        ok.sort(key=lambda x: -x["uoc_tinh_giam"])
        no.sort(key=lambda x: x["con_thieu"])
        views = ok + no
        if ok:
            best = ok[0]["ma"]
        if no:
            nxt = no[0]
            src = next(v for v in vs if v["ma"] == nxt["ma"])
            upsell = {"ma": nxt["ma"], "con_thieu": nxt["con_thieu"], "don_toi_thieu": nxt["don_toi_thieu"],
                      "uoc_tinh_giam_neu_du": _estimate(src, nxt["don_toi_thieu"]), "mo_ta": nxt["mo_ta"]}
    return {"today": fmt_date(t), "cart_total": cart_total, "count": len(views),
            "tot_nhat": best, "goi_y_upsell": upsell, "vouchers": views}


@app.get("/v1/vouchers/codes")
def voucher_codes():
    return {"codes": [v["ma"] for v in get_db().vouchers]}


@app.get("/v1/vouchers/{ma}")
def voucher_check(ma: str, cart_total: Optional[int] = Query(None, ge=0), today: Optional[str] = None):
    t = get_today(today)
    v = next((x for x in get_db().vouchers if x["ma"].upper() == ma.strip().upper()), None)
    # Mã nội bộ / cá nhân hóa (hien_thi_chatbot = FALSE) không được tiết lộ ra chatbot
    if v is None or not v["hien_thi"]:
        return {"ma": ma.strip().upper(), "trang_thai": "khong_kha_dung"}
    view = _view(v, t, cart_total)
    if v["han"] and v["han"] < t:
        return {**view, "trang_thai": "het_han"}
    if cart_total is not None and not view["du_dieu_kien"]:
        return {**view, "trang_thai": "chua_du_dieu_kien"}
    return {**view, "trang_thai": "hop_le"}
