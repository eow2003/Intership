import sys
import pandas as pd
from pathlib import Path
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

INPUT  = Path(r"D:\Internship\Accounts_2026_06_08_2.xlsx")
OUTPUT = Path(r"D:\Internship\Accounts_2026_06_08_sapnhap_UTF8.csv")
REPORT = OUTPUT.with_suffix(".report.txt")

# ═══════════════════════════════════════════════════════════════════════════
# Mapping: Billing State (Old) → Billing State New
# ═══════════════════════════════════════════════════════════════════════════
MAPPING = {
    # ── Sáp nhập: tên tỉnh THAY ĐỔI ─────────────────────────────────────────
    "Tỉnh Hà Nam":              "Tỉnh Ninh Bình",           # Hà Nam + Nam Định + Ninh Bình
    "Tỉnh Nam Định":            "Tỉnh Ninh Bình",
    "Tỉnh Quảng Nam":           "Thành phố Đà Nẵng",        # Quảng Nam → TP Đà Nẵng
    "Tỉnh Vĩnh Phúc":           "Tỉnh Phú Thọ",             # Vĩnh Phúc + Hòa Bình → Phú Thọ
    "Tỉnh Hoà Bình":            "Tỉnh Phú Thọ",
    "Tỉnh Hòa Bình":            "Tỉnh Phú Thọ",
    "Tỉnh Thừa Thiên Huế":      "Thành phố Huế",             # Thừa Thiên Huế → TP Huế
    "Tỉnh Thừa Thiên - Huế":    "Thành phố Huế",
    "Thừa Thiên Huế":           "Thành phố Huế",
    "Thành Phố Đồng Hới":       "Tỉnh Quảng Trị",           # Đồng Hới (tỉnh lỵ Q.Bình) → Quảng Trị
    "Tỉnh Quảng Bình":          "Tỉnh Quảng Trị",           # Quảng Bình → Quảng Trị
    "Tỉnh Đắk Nông":            "Tỉnh Lâm Đồng",            # Đắk Nông + Bình Thuận → Lâm Đồng
    "Tỉnh Bình Thuận":          "Tỉnh Lâm Đồng",
    "Tỉnh Phú Yên":             "Tỉnh Đắk Lắk",             # Phú Yên → Đắk Lắk
    "Tỉnh Kon Tum":             "Tỉnh Quảng Ngãi",          # Kon Tum → Quảng Ngãi
    "Tỉnh Kom Tum":             "Tỉnh Quảng Ngãi",          # typo
    "Tỉnh Bình Định":           "Tỉnh Gia Lai",             # Bình Định → Gia Lai
    "Tỉnh Ninh Thuận":          "Tỉnh Khánh Hòa",           # Ninh Thuận → Khánh Hòa
    "Tỉnh Hà Giang":            "Tỉnh Tuyên Quang",         # Hà Giang → Tuyên Quang
    "Tỉnh Yên Bái":             "Tỉnh Lào Cai",             # Yên Bái → Lào Cai
    "Tỉnh Bắc Kạn":             "Tỉnh Thái Nguyên",         # Bắc Kạn → Thái Nguyên
    "Tỉnh Bắc Giang":           "Tỉnh Bắc Ninh",            # Bắc Giang → Bắc Ninh
    "Tỉnh Thái Bình":           "Tỉnh Hưng Yên",            # Thái Bình → Hưng Yên
    "Tỉnh Hải Dương":           "Thành phố Hải Phòng",      # Hải Dương → TP Hải Phòng
    "Tỉnh Bình Dương":          "Thành phố Hồ Chí Minh",   # Bình Dương + BR-VT → TP HCM
    "Tỉnh Bà Rịa - Vũng Tàu":  "Thành phố Hồ Chí Minh",
    "Tỉnh Bình Phước":          "Tỉnh Đồng Nai",            # Bình Phước → Đồng Nai
    "Tỉnh Long An":             "Tỉnh Tây Ninh",            # Long An → Tây Ninh
    "Tỉnh Sóc Trăng":           "Thành phố Cần Thơ",        # Sóc Trăng + Hậu Giang → Cần Thơ
    "Tỉnh Hậu Giang":           "Thành phố Cần Thơ",
    "Tỉnh Bến Tre":             "Tỉnh Vĩnh Long",           # Bến Tre + Trà Vinh → Vĩnh Long
    "Tỉnh Trà Vinh":            "Tỉnh Vĩnh Long",
    "Tỉnh Tiền Giang":          "Tỉnh Đồng Tháp",           # Tiền Giang → Đồng Tháp
    "Tỉnh Bạc Liêu":            "Tỉnh Cà Mau",              # Bạc Liêu → Cà Mau
    "Tỉnh Kiên Giang":          "Tỉnh An Giang",            # Kiên Giang → An Giang

    # ── Chuẩn hoá viết tắt / sai chính tả ───────────────────────────────────
    "TP Đà Nẵng":               "Thành phố Đà Nẵng",
    "TP Cần Thơ":               "Thành phố Cần Thơ",
    "TP Hồ Chí Minh":           "Thành phố Hồ Chí Minh",
    "TP Hải Phòng":             "Thành phố Hải Phòng",
    "Thành phố Đà Nẳng":        "Thành phố Đà Nẵng",        # typo ẳ→ẵ
    "Tỉnh Khánh Hoà":           "Tỉnh Khánh Hòa",
    "Tỉnh Thanh Hoá":           "Tỉnh Thanh Hóa",
    "Tỉnh Hà Tỉnh":             "Tỉnh Hà Tĩnh",             # typo Tỉnh→Tĩnh
    "Khánh Hòa":                "Tỉnh Khánh Hòa",           # thiếu prefix
    "Thanh Hóa":                "Tỉnh Thanh Hóa",

    # ── Giữ nguyên (tỉnh/TP không thay đổi sau sáp nhập) ─────────────────────
    "Thành phố Hà Nội":         "Thành phố Hà Nội",
    "Thành phố Hồ Chí Minh":    "Thành phố Hồ Chí Minh",
    "Thành phố Hải Phòng":      "Thành phố Hải Phòng",
    "Thành phố Đà Nẵng":        "Thành phố Đà Nẵng",
    "Thành phố Cần Thơ":        "Thành phố Cần Thơ",
    "Thành phố Huế":            "Thành phố Huế",
    "Tỉnh An Giang":            "Tỉnh An Giang",
    "Tỉnh Bắc Ninh":            "Tỉnh Bắc Ninh",
    "Tỉnh Cao Bằng":            "Tỉnh Cao Bằng",
    "Tỉnh Cà Mau":              "Tỉnh Cà Mau",
    "Tỉnh Điện Biên":           "Tỉnh Điện Biên",
    "Tỉnh Đắk Lắk":             "Tỉnh Đắk Lắk",
    "Tỉnh Đồng Nai":            "Tỉnh Đồng Nai",
    "Tỉnh Đồng Tháp":           "Tỉnh Đồng Tháp",
    "Tỉnh Gia Lai":             "Tỉnh Gia Lai",
    "Tỉnh Hà Tĩnh":             "Tỉnh Hà Tĩnh",
    "Tỉnh Hưng Yên":            "Tỉnh Hưng Yên",
    "Tỉnh Khánh Hòa":           "Tỉnh Khánh Hòa",
    "Tỉnh Lai Châu":            "Tỉnh Lai Châu",
    "Tỉnh Lạng Sơn":            "Tỉnh Lạng Sơn",
    "Tỉnh Lào Cai":             "Tỉnh Lào Cai",
    "Tỉnh Lâm Đồng":            "Tỉnh Lâm Đồng",
    "Tỉnh Nghệ An":             "Tỉnh Nghệ An",
    "Tỉnh Ninh Bình":           "Tỉnh Ninh Bình",
    "Tỉnh Phú Thọ":             "Tỉnh Phú Thọ",
    "Tỉnh Quảng Ngãi":          "Tỉnh Quảng Ngãi",
    "Tỉnh Quảng Ninh":          "Tỉnh Quảng Ninh",
    "Tỉnh Quảng Trị":           "Tỉnh Quảng Trị",
    "Tỉnh Sơn La":              "Tỉnh Sơn La",
    "Tỉnh Thanh Hóa":           "Tỉnh Thanh Hóa",
    "Tỉnh Thái Nguyên":         "Tỉnh Thái Nguyên",
    "Tỉnh Tuyên Quang":         "Tỉnh Tuyên Quang",
    "Tỉnh Tây Ninh":            "Tỉnh Tây Ninh",
    "Tỉnh Vĩnh Long":           "Tỉnh Vĩnh Long",

    # ── Ngoài Việt Nam ────────────────────────────────────────────────────────
    "Ngoài Việt Nam":           "Ngoài Việt Nam",
    "Germany":                  "Germany",
    "Baden-Württemberg":        "Baden-Württemberg",
    "Phnom Penh":               "Phnom Penh",

    # ── Giá trị ghép — giữ nguyên ─────────────────────────────────────────────
    "Tỉnh Bắc Ninh hoặc Tỉnh Lạng Sơn":            "Tỉnh Bắc Ninh hoặc Tỉnh Lạng Sơn",
    "Thành phố Hồ Chí Minh / Tỉnh Đắk Lắk":        "Thành phố Hồ Chí Minh / Tỉnh Đắk Lắk",
    "Thành phố Hồ Chí Minh / Tỉnh Nghệ An":         "Thành phố Hồ Chí Minh / Tỉnh Nghệ An",
    "Thành phố Hà Nội / Chưa xác định":             "Thành phố Hà Nội / Chưa xác định",
}

# ═══════════════════════════════════════════════════════════════════════════
# Mapping: Billing State New → Khu vực / Miền
# ═══════════════════════════════════════════════════════════════════════════
REGION_MAPPING = {
    # 1. Trung du và miền núi phía Bắc
    "Tỉnh Cao Bằng":        "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Lạng Sơn":        "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Phú Thọ":         "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Thái Nguyên":     "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Tuyên Quang":     "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Lào Cai":         "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Lai Châu":        "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Sơn La":          "Vùng trung du và miền núi phía Bắc",
    "Tỉnh Điện Biên":       "Vùng trung du và miền núi phía Bắc",

    # 2. Đồng bằng sông Hồng
    "Thành phố Hà Nội":     "Vùng đồng bằng sông Hồng",
    "Thành phố Hải Phòng":  "Vùng đồng bằng sông Hồng",
    "Tỉnh Hưng Yên":        "Vùng đồng bằng sông Hồng",
    "Tỉnh Bắc Ninh":        "Vùng đồng bằng sông Hồng",
    "Tỉnh Ninh Bình":       "Vùng đồng bằng sông Hồng",
    "Tỉnh Quảng Ninh":      "Vùng đồng bằng sông Hồng",

    # 3. Bắc Trung Bộ
    "Tỉnh Thanh Hóa":       "Vùng Bắc Trung Bộ",
    "Tỉnh Nghệ An":         "Vùng Bắc Trung Bộ",
    "Tỉnh Hà Tĩnh":         "Vùng Bắc Trung Bộ",
    "Tỉnh Quảng Trị":       "Vùng Bắc Trung Bộ",
    "Thành phố Huế":        "Vùng Bắc Trung Bộ",

    # 4. Duyên hải Nam Trung Bộ và Tây Nguyên
    "Thành phố Đà Nẵng":    "Vùng duyên hải Nam Trung Bộ và Tây Nguyên",
    "Tỉnh Quảng Ngãi":      "Vùng duyên hải Nam Trung Bộ và Tây Nguyên",
    "Tỉnh Gia Lai":         "Vùng duyên hải Nam Trung Bộ và Tây Nguyên",
    "Tỉnh Đắk Lắk":         "Vùng duyên hải Nam Trung Bộ và Tây Nguyên",
    "Tỉnh Khánh Hòa":       "Vùng duyên hải Nam Trung Bộ và Tây Nguyên",
    "Tỉnh Lâm Đồng":        "Vùng duyên hải Nam Trung Bộ và Tây Nguyên",

    # 5. Đông Nam Bộ
    "Thành phố Hồ Chí Minh": "Vùng Đông Nam Bộ",
    "Tỉnh Đồng Nai":         "Vùng Đông Nam Bộ",
    "Tỉnh Tây Ninh":         "Vùng Đông Nam Bộ",

    # 6. Đồng bằng sông Cửu Long
    "Thành phố Cần Thơ":    "Vùng đồng bằng sông Cửu Long",
    "Tỉnh Vĩnh Long":       "Vùng đồng bằng sông Cửu Long",
    "Tỉnh An Giang":        "Vùng đồng bằng sông Cửu Long",
    "Tỉnh Đồng Tháp":       "Vùng đồng bằng sông Cửu Long",
    "Tỉnh Cà Mau":          "Vùng đồng bằng sông Cửu Long",
}

# ═══════════════════════════════════════════════════════════════════════════
# Đọc & xử lý
# ═══════════════════════════════════════════════════════════════════════════
print(f"Đọc file: {INPUT}")
df = pd.read_excel(INPUT, dtype=str)
df = df.fillna("")

src_col = "Billing State (Old)"
new_col = "Billing State New"

if src_col not in df.columns:
    raise ValueError(f"Không tìm thấy cột '{src_col}'. Các cột: {list(df.columns)}")

changed   = Counter()
unmatched = Counter()

def apply_map(val):
    v = val.strip()
    if not v:
        return ""
    if v in MAPPING:
        new = MAPPING[v]
        changed[(v, new)] += 1
        return new
    else:
        unmatched[v] += 1
        return v   # giữ nguyên nếu không có trong mapping

df[new_col] = df[src_col].apply(apply_map)

# ── Áp dụng Region Mapping vào cột "Khu vực / Miền" ────────────────────────
region_col    = "Khu vực / Miền"
region_changed = Counter()
region_no_map  = Counter()

def apply_region(row):
    province = row[new_col].strip()
    if province in REGION_MAPPING:
        new_region = REGION_MAPPING[province]
        old_region = row[region_col].strip()
        if old_region != new_region:
            region_changed[(old_region, new_region)] += 1
        return new_region
    else:
        region_no_map[province] += 1
        return row[region_col]   # giữ nguyên nếu tỉnh nằm ngoài mapping

df[region_col] = df.apply(apply_region, axis=1)

# ── Sắp xếp cột: Billing State New ngay sau Billing State (Old) ─────────────
cols = list(df.columns)
old_pos = cols.index(src_col)
cols.remove(new_col)
cols.insert(old_pos + 1, new_col)
df = df[cols]

# ═══════════════════════════════════════════════════════════════════════════
# Ghi CSV UTF-8 (có BOM cho Excel)
# ═══════════════════════════════════════════════════════════════════════════
df.to_csv(OUTPUT, index=False, encoding="utf-8-sig")
print(f"✓ Đã lưu: {OUTPUT}")

# ═══════════════════════════════════════════════════════════════════════════
# Report
# ═══════════════════════════════════════════════════════════════════════════
with open(REPORT, "w", encoding="utf-8") as f:
    f.write("KẾT QUẢ XỬ LÝ TỈNH/TP + KHU VỰC\n")
    f.write("=" * 80 + "\n\n")
    f.write(f"Input : {INPUT}\n")
    f.write(f"Output: {OUTPUT}\n")
    f.write(f"Tổng  : {len(df)} dòng\n\n")

    f.write("1. BILLING STATE (Old → New):\n")
    f.write(f"{'SL':>5}  {'CŨ':<42} → MỚI\n")
    f.write("-" * 80 + "\n")
    for (old, new), cnt in changed.most_common():
        marker = " ← SÁP NHẬP" if old != new else ""
        f.write(f"{cnt:5}  {old:<42} → {new}{marker}\n")

    f.write("\n2. KHU VỰC / MIỀN (đã cập nhật):\n")
    f.write(f"{'SL':>5}  {'CŨ':<35} → MỚI\n")
    f.write("-" * 80 + "\n")
    for (old, new), cnt in region_changed.most_common():
        f.write(f"{cnt:5}  {old:<35} → {new}\n")

    if region_no_map:
        f.write("\nTỉnh/TP không có trong REGION_MAPPING (giữ nguyên Khu vực):\n")
        for val, cnt in region_no_map.most_common():
            f.write(f"{cnt:5}  {val}\n")

    if unmatched:
        f.write("\nBilling State (Old) không có trong MAPPING (giữ nguyên):\n")
        for val, cnt in unmatched.most_common():
            f.write(f"{cnt:5}  {val}\n")

# ── In tóm tắt ra màn hình ───────────────────────────────────────────────────
total_prov_changed  = sum(c for (o, n), c in changed.items() if o != n)
total_prov_same     = sum(c for (o, n), c in changed.items() if o == n)
total_region_changed = sum(region_changed.values())

print()
print("=" * 70)
print(f"{'[1] BILLING STATE: Old → New':^70}")
print("=" * 70)
print(f"  {'BILLING STATE (OLD)':<38}  →  BILLING STATE NEW  [SL]")
print("-" * 70)
for (old, new), cnt in sorted(
        {k: v for k, v in changed.items() if k[0] != k[1]}.items(),
        key=lambda x: -x[1]):
    print(f"  {old:<38}  →  {new}  [{cnt}]")
print("-" * 70)
print(f"  Đổi tên (sáp nhập)  : {total_prov_changed:>5} records")
print(f"  Giữ nguyên          : {total_prov_same:>5} records")

print()
print("=" * 70)
print(f"{'[2] KHU VỰC / MIỀN: đã cập nhật':^70}")
print("=" * 70)
print(f"  {'KHU VỰC CŨ':<30}  →  KHU VỰC MỚI  [SL]")
print("-" * 70)
for (old, new), cnt in sorted(region_changed.items(), key=lambda x: -x[1]):
    print(f"  {old:<30}  →  {new}  [{cnt}]")
print("-" * 70)
print(f"  Khu vực được cập nhật: {total_region_changed:>5} records")
if region_no_map:
    print(f"  Tỉnh chưa có region  : {sum(region_no_map.values()):>5} records (xem report)")
print("=" * 70)
print(f"\n  Report chi tiết: {REPORT}\n")
