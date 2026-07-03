"""
calculator.py — Module xử lý PDF danh sách sản phẩm
==================================================
Đối chiếu Picking List PDF với master_data.xlsx để tính số lượng thực tế.
Mỗi Seller SKU trong PDF được tra cứu trong master_data, sau đó nhân số lượng.
"""
import os
import re
from collections import defaultdict
from datetime import datetime

from fpdf import FPDF
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# MASTER DATA
# ============================================================

def load_master_data(master_path: str) -> list[dict]:
    """
    Đọc file master_data.xlsx (hoặc mã combo.xlsx).
    Tự động nhận diện cột theo tên header, không phụ thuộc vị trí.
    Trả về list các dòng, mỗi dòng là dict:
      {seller_sku, sku, qty, qty_sold, promo_qty}
    """
    from openpyxl import load_workbook

    wb = load_workbook(master_path, data_only=True)
    ws = wb.active

    # ── Đọc header để xác định vị trí các cột ──
    headers = {}
    for ci, cell in enumerate(ws[1], 1):
        val = str(cell.value).lower().strip() if cell.value else ''
        headers[ci] = val

    def find_col(keywords: list[str]) -> int:
        """Tìm cột theo từ khóa (không phân biệt hoa thường)."""
        for ci, h in headers.items():
            for kw in keywords:
                if kw in h:
                    return ci
        return 0

    col_seller_sku = find_col(['combo', 'seller sku', 'seller_sku', 'mã combo', 'ma combo']) or 1
    col_sku = find_col(['sku']) or 2
    col_qty = find_col(['sl', 'qty', 'số lượng', 'so luong']) or 3
    col_qty_sold = find_col(['sl bán', 'sl ban', 'qty sold', 'qty_sold', 'bán', 'ban']) or 4
    col_promo = find_col(['sl km', 'promo qty', 'promo_qty', 'km']) or 5
    col_unit = find_col(['đơn vị tính', 'don vi tinh', 'đơn vị', 'don vi', 'dvt', 'unit']) or 0

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col_seller_sku - 1] is None:
            continue
        seller_sku = str(row[col_seller_sku - 1]).strip()
        sku = str(row[col_sku - 1]).strip() if row[col_sku - 1] is not None else seller_sku
        qty = int(row[col_qty - 1]) if row[col_qty - 1] is not None else 0
        qty_sold = int(row[col_qty_sold - 1]) if row[col_qty_sold - 1] is not None else 0
        promo_qty = int(row[col_promo - 1]) if row[col_promo - 1] is not None else 0
        unit = str(row[col_unit - 1]).strip() if col_unit and row[col_unit - 1] is not None else ''
        rows.append({
            "seller_sku": seller_sku,
            "sku": sku,
            "qty": qty,
            "qty_sold": qty_sold,
            "promo_qty": promo_qty,
            "unit": unit,
        })
    wb.close()
    return rows


def load_retail_data(retail_path: str) -> dict[str, dict]:
    """
    Doc file san pham ban le.xlsx.
    Cot: SKU | Don vi tinh | SL | SL ban | SL KM
    Khong co cot COMBO -> Seller SKU chinh la SKU.
    Tra ve dict: {sku: {seller_sku, sku, unit, qty, qty_sold, promo_qty}}
    """
    from openpyxl import load_workbook

    wb = load_workbook(retail_path, data_only=True)
    ws = wb.active

    headers = {}
    for ci, cell in enumerate(ws[1], 1):
        val = str(cell.value).lower().strip() if cell.value else ''
        headers[ci] = val

    def find_col(keywords):
        for ci, h in headers.items():
            for kw in keywords:
                if kw in h:
                    return ci
        return 0

    col_sku = find_col(['sku', 'mã sản phẩm', 'ma san pham', 'mã sản phẩm']) or 1
    col_name = find_col(['tên sản phẩm', 'ten san pham', 'tên sp', 'product name']) or 0
    col_unit = find_col(['đơn vị tính', 'don vi tinh', 'đơn vị', 'don vi', 'dvt', 'unit']) or 0
    col_qty = find_col(['sl', 'qty', 'số lượng', 'so luong']) or 3
    col_sold = find_col(['sl bán', 'sl ban', 'qty sold', 'bán', 'ban']) or 5
    col_promo = find_col(['sl km', 'promo qty', 'khuyến mại', 'khuyen mai', 'km']) or 6

    retail = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        sku = str(row[col_sku - 1]).strip() if row[col_sku - 1] is not None else ''
        if not sku:
            continue
        retail[sku] = {
            "seller_sku": sku,
            "sku": sku,
            "product_name": str(row[col_name - 1]).strip() if col_name and row[col_name - 1] is not None else '',
            "unit": str(row[col_unit - 1]).strip() if col_unit and row[col_unit - 1] is not None else '',
            "qty": int(row[col_qty - 1]) if row[col_qty - 1] is not None else 1,
            "qty_sold": int(row[col_sold - 1]) if row[col_sold - 1] is not None else 1,
            "promo_qty": int(row[col_promo - 1]) if row[col_promo - 1] is not None else 0,
        }
    wb.close()
    print(f"   Retail data: {len(retail)} SKUs")
    return retail


def build_sku_index(master_data: list[dict]) -> tuple[set[str], dict[str, str], dict[str, list[str]]]:
    """
    Từ master_data trả về:
      - master_skus: set tất cả Seller SKU
      - prefix_map: map từ SKU bị ngắt dòng → full Seller SKU (chỉ khi unique)
        VD: "CER01-" → "CER01-BOG17-2"
      - ambiguous_map: map từ prefix ambiguous → list các full SKU có thể
        VD: "ANTQ-3-" → ["ANTQ-3-6MET01", "ANTQ-3-TRG01"]
    """
    master_skus = set(r["seller_sku"] for r in master_data)

    prefix_map: dict[str, str] = {}
    ambiguous_map: dict[str, list[str]] = {}
    for sku in master_skus:
        if '-' in sku:
            parts = sku.split('-')
            for i in range(1, len(parts)):
                prefix = '-'.join(parts[:i]) + '-'
                if prefix not in prefix_map and prefix not in ambiguous_map:
                    prefix_map[prefix] = sku
                elif prefix in prefix_map:
                    # Trở thành ambiguous
                    ambiguous_map[prefix] = [prefix_map[prefix], sku]
                    del prefix_map[prefix]
                else:
                    ambiguous_map[prefix].append(sku)

    return master_skus, prefix_map, ambiguous_map


# ============================================================
# PDF EXTRACTION
# ============================================================

def _parse_pdf_header(full_text: str) -> dict:
    """
    Trich xuat thong tin header tu Picking List PDF.
    VD: Order quantity: 3 Product quantity: 3 Item quantity: 3
    Tra ve: {order_qty, product_qty, item_qty, print_time}
    """
    info = {}
    m = re.search(r'Order quantity:\s*(\d+)', full_text)
    if m:
        info['order_qty'] = int(m.group(1))
    m = re.search(r'Product quantity:\s*(\d+)', full_text)
    if m:
        info['product_qty'] = int(m.group(1))
    m = re.search(r'Item quantity:\s*(\d+)', full_text)
    if m:
        info['item_qty'] = int(m.group(1))
    m = re.search(r'Print time:\s*(.+)', full_text)
    if m:
        info['print_time'] = m.group(1).strip()
    return info


def extract_order_counts(
    pdf_path: str,
    master_skus: set[str],
    prefix_map: dict[str, str],
    ambiguous_map: dict[str, list[str]],
    retail_lookup: dict[str, dict] | None = None,
) -> tuple[dict[str, int], dict]:
    """
    Trich xuat so don hang cho moi Seller SKU tu PDF.
    Dung regex tim pattern: SellerSKU + Qty + OrderID (15+ chu so).

    Thu tu uu tien:
      1. Exact match trong master_skus (combo)
      2. Prefix match unique
      3. Prefix ambiguous -> doan tu context
      4. Tim trong retail_lookup (san pham don le)
      5. Khong tim thay -> bo qua + canh bao

    Tra ve: ({seller_sku: tong_so_qty}, {order_qty, product_qty, item_qty, print_time})
    """
    with pdfplumber.open(pdf_path) as pdf:
        texts = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texts.append(t)
    full_text = "\n".join(texts)

    # Gộp các dòng để xử lý SKU ngắt dòng (thay \n = space)
    flat_text = full_text.replace('\n', ' ')

    # Pattern: Mã SKU (chứa ít nhất 1 chữ cái, có thể bắt đầu bằng số, có thể kết thúc bằng - nếu bị ngắt)
    #           + Qty (số đơn hàng) + OrderID (15+ chữ số)
    pattern = r'\b((?=[A-Za-z0-9]*[A-Za-z])[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*(?:-)?)\s+(\d+)\s+(\d{15,})'

    counts: dict[str, int] = defaultdict(int)

    for match in re.finditer(pattern, flat_text):
        candidate = match.group(1)
        qty = int(match.group(2))

        if candidate in master_skus:
            # Exact match: Seller SKU xuat hien day du trong PDF
            counts[candidate] += qty
        elif candidate.endswith('-') and candidate in prefix_map:
            # Partial match unique: SKU bi ngat dong, tra prefix_map
            full_sku = prefix_map[candidate]
            counts[full_sku] += qty
        elif candidate.endswith('-') and candidate in ambiguous_map:
            # Prefix ambiguous: thu doan tu context sau Order ID
            options = ambiguous_map[candidate]
            after_text = flat_text[match.end():match.end()+120]
            best_match = None
            for opt in options:
                suffix = opt[len(candidate):]
                if suffix and suffix in after_text:
                    best_match = opt
                    break
            if best_match:
                counts[best_match] += qty
        elif candidate.endswith('-'):
            # Prefix khong co trong map nao -> bo qua
            pass
        elif retail_lookup and candidate in retail_lookup:
            # Tim thay trong san pham ban le
            counts[candidate] += qty
        else:
            # SKU la - khong co trong combo lan retail
            print(f"   ⚠ SKU la: {candidate} (x{qty}) - khong co trong ca 2 file")

    header_info = _parse_pdf_header(full_text)
    return dict(counts), header_info


# ============================================================
# CALCULATION
# ============================================================

def calculate_results(
    master_data: list[dict],
    order_counts: dict[str, int],
    retail_lookup: dict[str, dict] | None = None,
) -> list[dict]:
    """
    Doi chieu master_data voi order_counts tu PDF.
    - Seller SKU co trong master_data -> combo (theo dinh luong)
    - Seller SKU co trong retail_lookup -> san pham ban le
    - Khong tim thay -> canh bao + bo qua
    - Seller SKU không có trong master_data → sản phẩm đơn lẻ (1 đơn = 1 sp)
    """
    results = []
    matched_skus = set()
    for row in master_data:
        seller_sku = row["seller_sku"]
        if seller_sku in order_counts:
            matched_skus.add(seller_sku)
            mult = order_counts[seller_sku]
            # Tra cứu tên sản phẩm từ file retail (nếu SKU có trong đó)
            product_name = ''
            if retail_lookup:
                ri = retail_lookup.get(row["sku"])
                if ri:
                    product_name = ri.get("product_name", '')
            results.append({
                "seller_sku": seller_sku,
                "sku": row["sku"],
                "qty": row["qty"] * mult,
                "qty_sold": row["qty_sold"] * mult,
                "promo_qty": row["promo_qty"] * mult,
                "unit": row["unit"],
                "product_name": product_name,
            })

    # Sản phẩm bán lẻ: có trong PDF nhưng không có trong master_data combo
    for seller_sku, count in order_counts.items():
        if seller_sku not in matched_skus and retail_lookup and seller_sku in retail_lookup:
            r = retail_lookup[seller_sku]
            results.append({
                "seller_sku": seller_sku,
                "sku": r["sku"],
                "qty": r["qty"] * count,
                "qty_sold": r["qty_sold"] * count,
                "promo_qty": r["promo_qty"] * count,
                "unit": r["unit"],
                "product_name": r.get("product_name", ''),
            })
            matched_skus.add(seller_sku)

    # Cảnh báo SKU không tìm thấy ở đâu
    for seller_sku in order_counts:
        if seller_sku not in matched_skus:
            print(f"   ⚠ SKU không xác định: {seller_sku} (x{order_counts[seller_sku]}) - bỏ qua")

    return results



def generate_grouped_excel(results: list[dict], output_path: str, carrier: str = '', source_label: str = '', order_count: int = 0) -> str:
    """
    Tạo file Excel gộp theo SKU (không hiện Seller SKU).
    Format: SKU | Đơn vị tính | Qty | Qty Sold | Promo Qty
    Có dòng tiêu đề in đậm ở đầu để nhận diện khi in giấy.
    """
    # Gộp theo SKU
    grouped = {}
    for r in results:
        sku = r["sku"]
        if sku not in grouped:
            grouped[sku] = {"qty": 0, "qty_sold": 0, "promo_qty": 0, "unit": r.get("unit", ""), "product_name": r.get("product_name", "")}
        grouped[sku]["qty"] += r["qty"]
        grouped[sku]["qty_sold"] += r["qty_sold"]
        grouped[sku]["promo_qty"] += r["promo_qty"]
        # Giữ product_name đầu tiên khác rỗng
        if not grouped[sku]["product_name"] and r.get("product_name", ""):
            grouped[sku]["product_name"] = r["product_name"]

    grouped_list = [{"sku": k, **v} for k, v in grouped.items()]
    grouped_list.sort(key=lambda x: x["sku"])

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # ── Styles ──
    title_font = Font(name="Arial", size=14, bold=True, color="1F4E79")
    title_align = Alignment(horizontal="center", vertical="center")
    hdr_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    data_font = Font(name="Arial", size=10)
    total_font = Font(name="Arial", size=10, bold=True)

    # ── Title row (dòng nhận diện khi in giấy) ──
    ncols = 7  # STT, SKU, Tên SP, ĐVT, SL, SL bán, SL KM
    if carrier and source_label:
        title_text = f'{carrier} — {source_label}'
    else:
        title_text = carrier or source_label or 'Báo cáo gộp SKU'
    if order_count > 0:
        title_text += f' — ({order_count} đơn)'
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title_text)
    c.font = title_font
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28  # đủ cao để hiển thị rõ

    # ── Column headers (row 2) ──
    headers = ["STT", "SKU", "Tên sản phẩm", "Đơn vị tính", "SL", "SL bán", "SL KM"]
    # Độ rộng cột tối ưu: SKU=12, Tên SP=36 wrap, cột số=7 → tổng ~86 vừa A4 ngang
    col_widths = [5, 12, 36, 12, 7, 7, 7]
    col_aligns = ['C', 'C', 'L', 'C', 'R', 'R', 'R']  # center / left / right

    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = thin_border

    # ── Data (bắt đầu từ row 3) ──
    for ri, r in enumerate(grouped_list):
        rn = ri + 3
        vals = [ri + 1, r["sku"], r.get("product_name", ""), r.get("unit", ""), r["qty"], r["qty_sold"], r["promo_qty"]]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=rn, column=ci, value=v)
            c.font = data_font
            c.border = thin_border
            # Căn lề theo cột + wrap text cho cột Tên sản phẩm (cột 3)
            ha = col_aligns[ci - 1] if ci <= len(col_aligns) else 'L'
            c.alignment = Alignment(horizontal={'C':'center','L':'left','R':'right'}.get(ha, 'left'),
                                    vertical='center',
                                    wrap_text=(ci == 3))

    # ── Total row ──
    tr = len(grouped_list) + 3
    tong_qty = sum(r["qty"] for r in grouped_list)
    tong_sold = sum(r["qty_sold"] for r in grouped_list)
    tong_promo = sum(r["promo_qty"] for r in grouped_list)

    ws.cell(row=tr, column=1, value="Tổng").font = total_font
    ws.cell(row=tr, column=1).border = thin_border
    ws.cell(row=tr, column=2).border = thin_border  # SKU trống
    ws.cell(row=tr, column=3).border = thin_border  # Tên sản phẩm trống
    ws.cell(row=tr, column=4).border = thin_border  # Đơn vị tính trống

    for ci, val in [(5, tong_qty), (6, tong_sold), (7, tong_promo)]:
        c = ws.cell(row=tr, column=ci, value=val)
        c.font = total_font
        c.border = thin_border

    # ── Column widths ──
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── Print setup: vừa trang in, tránh mất cột ──
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # tự động số dòng
    ws.page_setup.paperSize = 9  # A4

    wb.save(output_path)
    print(f"   📊 Đã tạo Excel gộp: {os.path.basename(output_path)}")
    return output_path


# ============================================================
# PDF GENERATION (gộp theo SKU)
# ============================================================

def generate_grouped_pdf(results: list[dict], output_path: str, carrier: str = '', source_label: str = '', order_count: int = 0) -> str:
    """Tạo file PDF báo cáo gộp theo SKU."""
    # Gộp theo SKU
    grouped = {}
    for r in results:
        sku = r["sku"]
        if sku not in grouped:
            grouped[sku] = {"qty": 0, "qty_sold": 0, "promo_qty": 0, "unit": r.get("unit", ""), "product_name": r.get("product_name", "")}
        grouped[sku]["qty"] += r["qty"]
        grouped[sku]["qty_sold"] += r["qty_sold"]
        grouped[sku]["promo_qty"] += r["promo_qty"]
        if not grouped[sku]["product_name"] and r.get("product_name", ""):
            grouped[sku]["product_name"] = r["product_name"]

    grouped_list = [{"sku": k, **v} for k, v in grouped.items()]
    grouped_list.sort(key=lambda x: x["sku"])

    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_font('Arial', '', r'C:\Windows\Fonts\Arial.ttf', uni=True)
    pdf.add_font('Arial', 'B', r'C:\Windows\Fonts\Arialbd.ttf', uni=True)
    pdf.add_page()

    # Tiêu đề
    if carrier and source_label:
        title = f'{carrier} — {source_label}'
    else:
        title = carrier or source_label or 'Bao cao gop SKU'
    if order_count > 0:
        title += f' — ({order_count} don)'
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 10, title, align='C')
    pdf.ln(12)

    # Bảng — tận dụng tối đa chiều ngang A4 (297mm)
    col_w = [10, 28, 155, 24, 16, 18, 18]  # STT, SKU, Ten SP, DVT, SL, SL ban, SL KM
    headers = ['STT', 'SKU', 'Ten SP', 'DVT', 'SL', 'SL ban', 'SL KM']
    pdf.set_font('Arial', 'B', 8)
    pdf.set_fill_color(47, 84, 150)
    pdf.set_text_color(255, 255, 255)
    for i, (h, w) in enumerate(zip(headers, col_w)):
        pdf.cell(w, 10, h, border=1, fill=True, align='C')
    pdf.ln()

    # Data
    pdf.set_font('Arial', '', 9)
    pdf.set_text_color(0, 0, 0)
    tong_qty = tong_sold = tong_promo = 0
    for i, r in enumerate(grouped_list, 1):
        vals = [str(i), r['sku'], r.get('product_name', ''), r.get('unit', ''),
                str(r['qty']), str(r['qty_sold']), str(r['promo_qty'])]
        aligns = ['C', 'L', 'L', 'C', 'R', 'R', 'R']
        for v, w, a in zip(vals, col_w, aligns):
            pdf.cell(w, 9, v, border=1, align=a)
        pdf.ln()
        tong_qty += r['qty']; tong_sold += r['qty_sold']; tong_promo += r['promo_qty']

    # Total
    pdf.set_font('Arial', 'B', 9)
    total_vals = ['', 'Tong', '', '', str(tong_qty), str(tong_sold), str(tong_promo)]
    total_aligns = ['C', 'L', 'C', 'C', 'R', 'R', 'R']
    for v, w, a in zip(total_vals, col_w, total_aligns):
        pdf.cell(w, 10, v, border=1, align=a)

    pdf.output(output_path)
    print(f'   📄 Đã tạo PDF: {output_path}')
    return output_path



# ============================================================
# MAIN ENTRY POINT
# ============================================================

def process_all(
    pdf_files: list[str],
    output_dir: str,
    master_path: str | None = None,
    retail_path: str | None = None,
    carrier: str = '',
) -> list[dict]:
    """
    Xu ly toan bo pipeline cho nhieu file PDF cung 1 carrier.
    TAT CA PDF duoc gop chung vao 1 file bao cao duy nhat.
    """
    if master_path is None:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        master_path = os.path.join(base, "master_data.xlsx")

    if not os.path.exists(master_path):
        raise ValueError(f"Khong tim thay file master_data tai: {master_path}")

    # Load master data (combo)
    print(f"Load master data: {master_path}")
    master_data = load_master_data(master_path)
    master_skus, prefix_map, ambiguous_map = build_sku_index(master_data)
    print(f"   {len(master_data)} dong, {len(master_skus)} Seller SKU"
          + (f", {len(prefix_map)} prefix map" if prefix_map else ""))

    # Load retail data
    retail_lookup = None
    if retail_path and os.path.exists(retail_path):
        retail_lookup = load_retail_data(retail_path)
    elif retail_path:
        print(f"   ⚠ Khong tim thay file retail: {retail_path}")

    # Gom order_counts tu TAT CA PDF
    from collections import defaultdict
    merged_order_counts: dict[str, int] = defaultdict(int)
    total_order_qty = 0   # Số đơn hàng thực tế (Order quantity từ header)

    for pdf_path in pdf_files:
        print(f"\nXu ly: {os.path.basename(pdf_path)}")
        order_counts, header_info = extract_order_counts(pdf_path, master_skus, prefix_map, ambiguous_map, retail_lookup)
        if not order_counts:
            print(f"   ⚠ Khong tim thay Seller SKU nao trong PDF!")
            continue
        for sku, count in order_counts.items():
            merged_order_counts[sku] += count
        order_qty = header_info.get('order_qty', 0)
        total_order_qty += order_qty
        print(f"   Tim thay {len(order_counts)} Seller SKU, {order_qty} don hang (header), {sum(order_counts.values())} mat hang")

    if not merged_order_counts:
        raise ValueError("Không trích xuất được dữ liệu từ bất kỳ PDF nào. "
                         "Kiểm tra file đầu vào và master_data.")

    # Tinh toan tu merged order counts
    results = calculate_results(master_data, dict(merged_order_counts), retail_lookup)
    print(f"\n📊 KET QUA CHUNG: {len(results)} dong | "
          f"Qty={sum(r['qty'] for r in results)} | "
          f"Sold={sum(r['qty_sold'] for r in results)} | "
          f"Promo={sum(r['promo_qty'] for r in results)}")

    # Sinh 1 file PDF duy nhat
    now = datetime.now()
    carrier_safe = carrier.replace(' ', '_').replace('&', 'n') if carrier else ''
    prefix = f"Bao_cao_gop_SKU_{carrier_safe}_" if carrier_safe else "Bao_cao_gop_SKU_"
    pdf_path = os.path.join(output_dir, f"{prefix}{now.strftime('%m-%d_%H-%M-%S')}.pdf")

    if carrier:
        source_label = f"{len(pdf_files)} Picking list — {now.strftime('%d/%m %H:%M')}"
    else:
        source_label = f"{len(pdf_files)} Picking list — {now.strftime('%d/%m %H:%M')}"

    generate_grouped_pdf(results, pdf_path, carrier=carrier, source_label=source_label, order_count=total_order_qty)

    return [{
        "base_name": f"Combined {carrier or 'all'}",
        "rows": len(results),
        "tong_qty": sum(r["qty"] for r in results),
        "tong_sold": sum(r["qty_sold"] for r in results),
        "tong_promo": sum(r["promo_qty"] for r in results),
        "files": {"pdf_report": pdf_path},
    }]


# ============================================================
# CLI — Test nhanh
# ============================================================
if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python calculator.py <pdf_file> [pdf_file2...]")
        sys.exit(1)

    pdfs = sys.argv[1:]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
    os.makedirs(out_dir, exist_ok=True)

    master = os.path.join(os.path.dirname(os.path.abspath(__file__)), "master_data.xlsx")

    try:
        results = process_all(pdfs, out_dir, master)
        print("\n" + "=" * 60)
        print("✅ HOÀN THÀNH!")
        for r in results:
            print(f"\n📦 {r['base_name']}:")
            print(f"   Dòng: {r['rows']} | Qty: {r['tong_qty']} | "
                  f"Sold: {r['tong_sold']} | Promo: {r['tong_promo']}")
            for k, v in r["files"].items():
                print(f"   📎 {k}: {os.path.basename(v)}")
    except ValueError as e:
        print(f"\n❌ Lỗi: {e}")
        sys.exit(1)
