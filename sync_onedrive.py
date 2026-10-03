#!/usr/bin/env python3
"""
Sync master data Kalkulator Promo Apple  ->  data.json

Alur (sama dengan Dashtp3staff):
    file master (.xlsx)  --download-->  parsing 6 sheet  -->  data.json  -->  index.html

Sumber file diambil dari environment variable ONEDRIVE_URL (diisi dari GitHub Secret
ONEDRIVE_DIRECT_URL). Link yang didukung:
    - direct link .xlsx (OneDrive / SharePoint / hosting lain)
    - link share OneDrive (1drv.ms / onedrive.live.com)  -> otomatis diubah jadi link unduh
    - link Google Sheets (docs.google.com/spreadsheets/d/...)  -> otomatis di-export ke xlsx

Pemakaian lokal (tanpa download):
    python sync_onedrive.py --file "Calculator Promo.xlsx"

Port dari Code.gs (Apps Script) dengan perbaikan:
    - header sheet dibaca tanpa peduli huruf besar/kecil & spasi ganda ("Promotion   Price")
    - baris judul model = kolom A terisi, kolom B & C kosong (tahan sel nyasar)
    - label RAM/Storage Mac (16G/512GB, 24GB/1T, 16GB/256-IND, dst.)
    - bank di sheet Promo yang belum ada di CARD_MATRIX ditambahkan otomatis
    - tiap promo diberi end_date (dibaca dari kolom Periode) supaya promo lewat tidak ditawarkan
"""
import argparse
import base64
import calendar
import hashlib
import io
import json
import os
import re
import sys
from datetime import date, datetime, timedelta, timezone

import requests
from openpyxl import load_workbook

OUT_FILE = "data.json"
WIB = timezone(timedelta(hours=7))

SHEET_NAMES = {
    "price": "Price List",
    "promo": "Promo Berjalan",
    "bnpl": "BNPL",
    "provider": "Provider",
    "qoala": "Qoala Protection",
    "trade": "Trade in",
}

# Tenor cicilan 0% kartu kredit (tidak ada di sheet -> dikelola di sini).
CARD_MATRIX = [
    ("BCA", [3, 6, 12, 18, 24]),
    ("BRI", [3, 6, 12, 18, 24]),
    ("BNI", [3, 6, 12, 18, 24]),
    ("CIMB", [3, 6, 12, 24]),
    ("Bank BSI", [3, 6, 12, 24]),
    ("Panin", [3, 6, 12]),
    ("Mandiri", [3, 6, 12, 18, 24]),
    ("Permata", [3, 6, 12, 18, 24]),
    ("DBS", [3, 6, 12, 18, 24]),
    ("HSBC", [3, 6, 12, 18, 24]),
    ("Danamon", [3, 6, 12, 18, 24]),
    ("UOB", [3, 6, 12, 18, 24]),
    ("Maybank", [3, 6, 12, 24]),
    ("Jenius (BTPN)", [3, 6, 12]),
    ("KB Bank", [3, 6, 12, 18, 24]),
    ("BRI Samsung Card", [3, 6, 12, 18, 24]),
    ("BRI Debit", [3, 6, 12]),
    ("OCBC", [3, 6, 12]),
]
DEFAULT_NEW_CARD_TENORS = [3, 6, 12]


class SheetError(Exception):
    """Struktur sheet tidak sesuai harapan -> data.json lama dipertahankan."""


# --------------------------------------------------------------------------- helpers
def text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():  # sel angka (mis. model "5") jangan jadi "5.0"
        return str(int(v))
    return str(v).strip()


def norm_ws(v):
    return re.sub(r"\s+", " ", text(v)).strip()


def hdr(v):
    return norm_ws(v).lower()


def cell(row, i):
    return row[i] if 0 <= i < len(row) else None


def num(v):
    """Angka dari sel. Nilai numerik dipakai apa adanya; teks dibaca format US/Indonesia."""
    if v is None or v == "" or isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        n = v
    else:
        s = re.sub(r"[^0-9.,\-]", "", str(v))
        if not s or s == "-":
            return 0
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", s):  # 16.999.000 / 16.999.000,50
            s = s.replace(".", "").replace(",", ".")
        else:  # 16,999,000 / 16,999,000.50
            s = s.replace(",", "")
        try:
            n = float(s)
        except ValueError:
            return 0
    if n != n or n in (float("inf"), float("-inf")):
        return 0
    n = round(n, 4)  # buang noise floating point (599400.0000000001)
    return int(n) if float(n).is_integer() else n


def normalize_bank(s):
    """Sama persis dengan normalizeBank() di index.html."""
    s = text(s).lower()
    s = re.sub(r"bank\s*", "", s)
    s = re.sub(r"\s+cc$", "", s)
    s = re.sub(r"\s+card$", "", s)
    return re.sub(r"\s+", "", s).strip()


# --------------------------------------------------------------------------- download
def direct_url(url):
    u = url.strip()
    m = re.match(r"https://docs\.google\.com/spreadsheets/d/([\w-]+)", u)
    if m:
        return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
    if re.match(r"https?://(1drv\.ms|onedrive\.live\.com)/", u):
        b64 = base64.urlsafe_b64encode(u.encode()).decode().rstrip("=")
        return f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"
    if "sharepoint.com" in u and "download=1" not in u:
        return u + ("&" if "?" in u else "?") + "download=1"
    return u


def download(url):
    r = requests.get(
        direct_url(url),
        timeout=90,
        allow_redirects=True,
        headers={"User-Agent": "Mozilla/5.0 (kalkulator-promo-sync)"},
    )
    r.raise_for_status()
    data = r.content
    if data[:2] != b"PK":
        raise SheetError(
            "File yang terunduh bukan .xlsx (biasanya halaman login / izin akses). "
            "Pastikan link dibagikan 'siapa saja yang memiliki link' dan formatnya .xlsx."
        )
    return data


def read_sheets(raw):
    if raw[:2] != b"PK":
        raise SheetError("File bukan .xlsx. Simpan ulang file master sebagai .xlsx (Excel Workbook).")
    wb = load_workbook(io.BytesIO(raw), data_only=True)
    return {ws.title.strip(): [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


# --------------------------------------------------------------------------- parsers
def extract_memory_label(description):
    s = norm_ws(description)
    # Mac: spesifikasi dipisah "/" ; RAM & storage ada di dua token berurutan.
    if "/" in s:
        parts = [p.strip() for p in re.sub(r"-IND$", "", s.upper()).split("/")]
        for i in range(len(parts) - 2, -1, -1):
            ram = re.fullmatch(r"(\d+)\s*(?:GB|G)?", parts[i])
            sto = re.fullmatch(r"(\d+)\s*(GB|TB|T)?(?:-IND)?", parts[i + 1])
            if ram and sto:
                n = int(sto.group(1))
                if sto.group(2):
                    unit = "TB" if sto.group(2) == "T" else sto.group(2)
                else:
                    unit = "GB" if n >= 64 else "TB"
                return f"{ram.group(1)}GB / {n}{unit}"
        if not re.search(r"\d+\s*(GB|TB)", s, re.I):
            return s  # tanpa RAM/storage (mis. Studio Display): pakai deskripsi penuh
    m = re.search(r"(\d+\s*GB)\s*[/%x×]\s*(\d+\s*(?:GB|TB))", s, re.I)
    if m:
        return re.sub(r"\s+", "", m.group(1)) + " / " + re.sub(r"\s+", "", m.group(2))
    m = re.search(r"(\d+\s*(?:GB|TB))(?!.*\d+\s*(?:GB|TB))", s, re.I)
    return re.sub(r"\s+", "", m.group(1)) if m else "Standard"


def tab_for(category, description, group):
    if category == "Apple Device - iPhone":
        return "iPhone"
    if category == "Apple Device - iPad":
        return "iPad"
    if "APPLE WATCH" in description.upper() or group.lower().startswith("apple watch"):
        return "Apple Watch"
    return "Mac"


def parse_price_list(rows):
    hi = None
    for r in range(min(len(rows), 15)):
        h = [hdr(v) for v in rows[r]]
        if (
            "sap article" in h and "sapdescription" in h and "category" in h
            and "normal price" in h and any(x.startswith("promotion") for x in h)
        ):
            hi = r
            break
    if hi is None:
        raise SheetError(
            "Header sheet 'Price List' tidak ditemukan. Harus ada kolom: SAP Article, SAPDescription, "
            "Category, Normal Price, Promotion Price (di 15 baris pertama)."
        )
    h = [hdr(v) for v in rows[hi]]
    ia, ib, ic = h.index("sap article"), h.index("sapdescription"), h.index("category")
    inormal = h.index("normal price")
    ipromo = next(i for i, x in enumerate(h) if x.startswith("promotion"))
    ichange = next((i for i, x in enumerate(h) if x.startswith("penurunan")), -1)
    iremarks = h.index("cash price remarks") if "cash price remarks" in h else -1

    catalog = {"iPhone": [], "iPad": [], "Apple Watch": [], "Mac": []}
    current = ""
    for row in rows[hi + 1:]:
        a, b, c = text(cell(row, ia)), text(cell(row, ib)), text(cell(row, ic))
        if a and not b and not c:  # baris judul model
            current = a
            continue
        if not (a and b and c) or not re.match(r"^Apple Device - ", c, re.I):
            continue
        normal = num(cell(row, inormal))
        promo = num(cell(row, ipromo))
        group = norm_ws(current or b)
        catalog[tab_for(c, b, group)].append({
            "article": a,
            "description": b,
            "category": c,
            "group": group,
            "normal_price": normal,
            "promo_price": promo or normal,
            "change": num(cell(row, ichange)) if ichange >= 0 else 0,
            "remarks": text(cell(row, iremarks)) if iremarks >= 0 else "",
            "memory_label": extract_memory_label(b),
        })
    return catalog


def parse_catalog(rows):
    """Sheet Provider & Qoala Protection."""
    if not rows:
        return []
    hi = 0
    for r in range(min(5, len(rows))):
        h = [hdr(v) for v in rows[r]]
        if "sap article" in h and ("sap description" in h or "sapdescription" in h):
            hi = r
            break
    h = [hdr(v) for v in rows[hi]]

    def find(pred):
        return next((i for i, x in enumerate(h) if pred(x)), -1)

    ib = find(lambda x: x == "brand")
    ia = find(lambda x: x == "sap article")
    idesc = find(lambda x: x in ("sap description", "sapdescription"))
    iref = find(lambda x: x.startswith("reference"))
    icur = find(lambda x: x.startswith("current"))
    irep = find(lambda x: "repricing" in x)
    if ia < 0:
        raise SheetError("Kolom 'SAP Article' tidak ditemukan di sheet Provider/Qoala.")
    if irep < 0:
        raise SheetError("Kolom 'Repricing Cash' tidak ditemukan di sheet Provider/Qoala.")

    out = []
    for row in rows[hi + 1:]:
        article = text(cell(row, ia))
        if not article:
            continue
        out.append({
            "brand": text(cell(row, ib)) if ib >= 0 else "",
            "article": article,
            "description": text(cell(row, idesc)) if idesc >= 0 else article,
            "reference": num(cell(row, iref)) if iref >= 0 else 0,
            "current": num(cell(row, icur)) if icur >= 0 else 0,
            "repricing": num(cell(row, irep)),
        })
    return out


def parse_qoala(rows):
    qoala = parse_catalog(rows)
    for q in qoala:
        m = re.search(r"Max\s*([0-9.]+)", q["description"], re.I)
        q["max_device"] = int(m.group(1).replace(".", "")) if m else 0
        m = re.search(r"(\d+)\s*bulan", q["description"], re.I)
        tenor = int(m.group(1)) if m else None
        if not tenor:
            m = re.search(r"\((\d+)-", q["article"])
            tenor = int(m.group(1)) if m else None
        q["qoala_tenor"] = tenor
    return qoala


def parse_bnpl(rows):
    if len(rows) < 2:
        return []
    h = [hdr(v) for v in rows[0]]
    iname = h.index("financing") if "financing" in h else 0
    ibunga = next((i for i, x in enumerate(h) if x.startswith("bunga")), 4)
    out = []
    for row in rows[1:]:
        name = text(cell(row, iname))
        if not name:
            continue
        tenors = [int(n) for n in (num(cell(row, c)) for c in range(iname + 1, ibunga)) if n]
        out.append({"name": name, "tenors": tenors, "interest": text(cell(row, ibunga))})
    return out


def parse_trade_in(rows):
    if len(rows) < 3:
        return []
    hi = None
    for r in range(min(5, len(rows))):
        h = [hdr(v) for v in rows[r]]
        if "brand" in h and "model name" in h:
            hi = r
            break
    if hi is None:
        raise SheetError("Header sheet 'Trade in' tidak ditemukan (butuh kolom BRAND dan MODEL NAME).")
    h = [hdr(v) for v in rows[hi]]
    ibrand, imodel = h.index("brand"), h.index("model name")
    gcols = {g: h.index("grade " + g.lower()) for g in "SABCD" if ("grade " + g.lower()) in h}
    out = []
    for row in rows[hi + 1:]:
        brand, model = text(cell(row, ibrand)), text(cell(row, imodel))
        if not brand or not model:
            continue
        grades = {g: num(cell(row, i)) for g, i in gcols.items() if cell(row, i) not in (None, "")}
        if grades:
            out.append({"brand": brand, "model": model, "grades": grades})
    return out


MONTHS = {
    "januari": 1, "jan": 1, "februari": 2, "feb": 2, "maret": 3, "mar": 3, "april": 4, "apr": 4,
    "mei": 5, "juni": 6, "jun": 6, "juli": 7, "jul": 7, "agustus": 8, "agu": 8, "agt": 8, "ags": 8,
    "september": 9, "sep": 9, "sept": 9, "oktober": 10, "okt": 10, "november": 11, "nov": 11,
    "desember": 12, "des": 12,
}


def period_end(period, today):
    """Tanggal berakhir dari teks Periode (mis. '24 April - 13 September', 'Hingga 31 Desember'). None bila tidak terbaca."""
    words = [w for w in re.findall(r"[a-z]+", text(period).lower()) if w in MONTHS]
    if not words:
        return None
    start_m, end_m = MONTHS[words[0]], MONTHS[words[-1]]
    days = re.findall(r"(\d{1,2})\s*([a-z]+)", text(period).lower())
    days = [(int(d), MONTHS[w]) for d, w in days if w in MONTHS]
    year = today.year
    if start_m > end_m and today.month >= start_m:  # melintasi pergantian tahun (Des -> Jan)
        year += 1
    if days:
        day, month = days[-1]
    else:
        day, month = 31, end_m
    day = min(day, calendar.monthrange(year, month)[1])
    try:
        return date(year, month, day)
    except ValueError:
        return None


def parse_promos(rows, today):
    hi = None
    need = ["promo", "periode", "bank", "scheme", "minimal amount", "maximal amount"]
    for r in range(min(len(rows), 6)):
        h = [hdr(v) for v in rows[r]]
        if all(n in h for n in need):
            hi = r
            break
    if hi is None:
        raise SheetError(
            "Header sheet 'Promo Berjalan' tidak ditemukan. Butuh kolom: Promo, Periode, Bank, Scheme, "
            "Minimal Amount, Maximal Amount, Discount/Cashback."
        )
    h = [hdr(v) for v in rows[hi]]
    ip, ipe, ib, isch = h.index("promo"), h.index("periode"), h.index("bank"), h.index("scheme")
    imin, imax = h.index("minimal amount"), h.index("maximal amount")
    idisc = next((i for i, x in enumerate(h) if x.startswith("discount")), -1)
    if idisc < 0:
        raise SheetError("Kolom 'Discount/Cashback' tidak ditemukan di sheet Promo Berjalan.")
    out, carry = [], ""
    for row in rows[hi + 1:]:
        promo, period = text(cell(row, ip)), text(cell(row, ipe)) or carry
        bank, scheme = text(cell(row, ib)), text(cell(row, isch))
        if text(cell(row, ipe)):
            carry = text(cell(row, ipe))
        if not promo or not bank or not scheme:
            continue
        end = period_end(period, today)
        out.append({
            "promo": promo, "period": period, "bank": bank, "scheme": scheme,
            "min": num(cell(row, imin)), "max": num(cell(row, imax)), "discount": num(cell(row, idisc)),
            "end_date": end.isoformat() if end else None,
        })
    return out


def build_card_options(promos, warnings):
    cards = [{"name": n, "tenors": t} for n, t in CARD_MATRIX]
    known = {normalize_bank(c["name"]) for c in cards}
    for p in promos:
        k = normalize_bank(p["bank"])
        if k not in known:
            known.add(k)
            cards.append({"name": p["bank"], "tenors": list(DEFAULT_NEW_CARD_TENORS)})
            warnings.append(
                f"Bank '{p['bank']}' ada di sheet Promo tapi belum ada di CARD_MATRIX -> ditambahkan otomatis "
                f"dengan tenor {DEFAULT_NEW_CARD_TENORS}. Cek tenornya di sync_onedrive.py."
            )
    return cards


# --------------------------------------------------------------------------- build
def find_label_collisions(catalog):
    """Varian yang harganya beda tapi label RAM/Storage sama -> hanya satu yang bisa dipilih di kalkulator."""
    out = []
    for tab, items in catalog.items():
        seen = {}
        for p in items:
            conn = ""
            if tab == "iPad":
                d = p["description"].lower()
                conn = "C" if re.search(r"wi-?fi\s*\+?\s*cell|cellular", d) else "W"
            seen.setdefault((p["group"], conn, p["memory_label"]), set()).add(p["promo_price"])
        for (group, conn, label), prices in seen.items():
            if len(prices) > 1:
                out.append(f"[{tab}] {group} / {label}{' / ' + conn if conn else ''}: harga {sorted(prices)}")
    return out


def build(sheets, today, source_name, warnings):
    for key, name in SHEET_NAMES.items():
        if name not in sheets:
            raise SheetError(f"Sheet '{name}' tidak ditemukan. Sheet yang ada: {', '.join(sheets)}")
    catalog = parse_price_list(sheets[SHEET_NAMES["price"]])
    promos = parse_promos(sheets[SHEET_NAMES["promo"]], today)
    data = {
        "source_file": source_name,
        "catalog": catalog,
        "bnpl": parse_bnpl(sheets[SHEET_NAMES["bnpl"]]),
        "providers": parse_catalog(sheets[SHEET_NAMES["provider"]]),
        "qoala": parse_qoala(sheets[SHEET_NAMES["qoala"]]),
        "trade_in": parse_trade_in(sheets[SHEET_NAMES["trade"]]),
        "card_options": build_card_options(promos, warnings),
        "promos": promos,
    }
    # Pengaman: jangan publish data kosong/rusak.
    for tab, items in catalog.items():
        if not items:
            raise SheetError(f"Katalog '{tab}' kosong setelah parsing. Cek kolom Category di sheet Price List.")
    for key in ("bnpl", "providers", "qoala", "trade_in", "promos"):
        if not data[key]:
            raise SheetError(f"Data '{key}' kosong setelah parsing.")
    warnings.extend(find_label_collisions(catalog))
    expired = [p for p in promos if p["end_date"] and p["end_date"] < today.isoformat()]
    if expired:
        warnings.append(
            f"{len(expired)} dari {len(promos)} promo sudah lewat periodenya (disembunyikan otomatis di kalkulator); "
            f"sebaiknya baris-nya dihapus dari sheet Promo Berjalan."
        )
    unreadable = [p for p in promos if not p["end_date"]]
    if unreadable:
        warnings.append(f"{len(unreadable)} promo dengan Periode yang tidak terbaca tanggalnya (tetap ditampilkan).")
    return data


def fingerprint(data):
    blob = json.dumps(data, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()[:16]


def total_products(data):
    return sum(len(v) for v in data.get("catalog", {}).values())


def main():
    ap = argparse.ArgumentParser(description="Sync master Excel -> data.json")
    ap.add_argument("--file", help="pakai file .xlsx lokal (tanpa download)")
    ap.add_argument("--out", default=OUT_FILE)
    ap.add_argument("--source-name", default=os.environ.get("SOURCE_NAME", "Calculator Promo"))
    ap.add_argument("--force", action="store_true", help="tulis ulang walau data tidak berubah")
    args = ap.parse_args()

    now = datetime.now(WIB)
    warnings = []
    try:
        if args.file:
            raw = open(args.file, "rb").read()
        else:
            url = os.environ.get("ONEDRIVE_URL") or os.environ.get("SOURCE_URL")
            if not url:
                raise SheetError("ONEDRIVE_URL kosong. Isi GitHub Secret ONEDRIVE_DIRECT_URL (Settings > Secrets > Actions).")
            raw = download(url)
        data = build(read_sheets(raw), now.date(), args.source_name, warnings)
    except SheetError as e:
        print(f"::error::{e}")
        return 1
    except requests.RequestException as e:
        print(f"::error::Gagal mengunduh file master: {e}")
        return 1

    version = fingerprint(data)
    old = None
    if os.path.exists(args.out):
        try:
            old = json.load(open(args.out, encoding="utf-8"))
        except (OSError, ValueError):
            old = None

    for w in warnings:
        print(f"::warning::{w}")

    if old and total_products(old) and total_products(data) < total_products(old) * 0.5:
        print(f"::error::Jumlah produk turun drastis ({total_products(old)} -> {total_products(data)}). "
              "Dicurigai sheet terpotong/rusak; data.json lama dipertahankan.")
        return 1

    cat = {k: len(v) for k, v in data["catalog"].items()}
    print(f"Terbaca: katalog {cat} | qoala {len(data['qoala'])} | provider {len(data['providers'])} | "
          f"trade-in {len(data['trade_in'])} | bnpl {len(data['bnpl'])} | kartu {len(data['card_options'])} | promo {len(data['promos'])}")

    if old and old.get("version") == version and not args.force:
        print("Tidak ada perubahan data. data.json tidak diubah.")
        return 0

    data["version"] = version
    data["updated_at"] = now.isoformat(timespec="seconds")
    tmp = args.out + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, args.out)
    print(f"data.json diperbarui (versi {version}, {os.path.getsize(args.out) / 1024:.0f} KB).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
